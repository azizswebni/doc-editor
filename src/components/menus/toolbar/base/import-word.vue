<template>
  <menus-button
    ico="word"
    :text="t('base.importWord.text')"
    huge
    @menu-click="importWord"
  />
</template>

<script setup lang="ts">
import mammoth from 'mammoth'

const container = inject('container')
const editor = inject('editor')
const options = inject('options')
const $options = options.value.importWord

const importWord = () => {
  const { open, onChange } = useFileDialog({
    accept: '.docx',
    reset: true,
    multiple: false,
  })
  // Open file dialog
  open()
  // Insert file
  onChange(async (files: FileList | null) => {
    const [file] = Array.from(files ?? [])
    if (!file) {
      return
    }
    if (file.size > $options.maxSize) {
      useMessage('error', {
        attach: container,
        content: t('base.importWord.limitSize', {
          limitSize: $options.maxSize / (1024 * 1024),
        }),
      })
      return
    }
    const message = await useMessage('loading', {
      attach: container,
      content: t('base.importWord.converting'),
    })

    // Use custom import method
    if ($options?.useCustomMethod) {
      const result = await $options.onCustomImportMethod?.(file)
      message.close()
      try {
        if (result?.messages?.type === 'error') {
          useMessage('error', {
            attach: container,
            content: `${t('base.importWord.convertError')} (${result.messages.message})`,
          })
          return
        }
        if (result?.value) {
          editor.value?.commands.setContent(result.value)
        } else {
          useMessage('error', {
            attach: container,
            content: t('base.importWord.importError'),
          })
        }
      } catch {
        useMessage('error', {
          attach: container,
          content: t('base.importWord.importError'),
        })
      }
      return
    }

    // Use Mammoth import
    const arrayBuffer = await file.arrayBuffer()

    // Enhanced configuration with text styling support
    const customOptions = {
      // Convert document properties including headers and footers
      includeDefaultStyleMap: true,
      includeEmbeddedStyleMap: true,
      ignoreEmptyParagraphs: true,
      idPrefix: 'docGen',
      styleHook(el, pr) {
        // Handle paragraph-level styles (alignment, indent, line-height)
        if (el.name === 'w:p') {
          const result = <any>[]

          const alignEl = el.firstOrEmpty('w:pPr').first('w:jc')
          if (alignEl) {
            const alignVal = alignEl.attributes['w:val']
            result.push(`text-align:${alignVal}`)
          }

          const indentEl = el.firstOrEmpty('w:pPr').first('w:ind')
          if (indentEl) {
            const indentVal = indentEl.attributes['w:firstLineChars']
            if (indentVal) {
              result.push(
                `text-indent:${parseFloat(indentVal / 100).toFixed(0)}em`,
              )
            }
          }

          const lineHeightEl = el.firstOrEmpty('w:pPr').first('w:spacing')
          if (lineHeightEl) {
            const lineHeightVal = lineHeightEl.attributes['w:line']
            const lineTypeVal = lineHeightEl.attributes['w:lineRule']
            if (lineHeightVal && lineTypeVal === 'auto') {
              const lineHeight = parseFloat(lineHeightVal / 240).toFixed(2)
              result.push(`line-height:${lineHeight}`)
            } else if (lineHeightVal && lineTypeVal === 'exact') {
              const lineHeight = parseFloat(lineHeightVal / 20).toFixed(2)
              result.push(`line-height:${lineHeight}pt`)
            }
          }

          return result.join(';') || ''
        }
        // Handle text-run-level styles (color, font-size, bold) - these should be on inline elements
        else if (el.name === 'w:r') {
          const result = <any>[]
          const rPr = el.firstOrEmpty('w:rPr')

          const fontSizeEl = rPr.first('w:sz')
          if (fontSizeEl) {
            const fontSizeVal = fontSizeEl.attributes['w:val']
            result.push(`font-size:${fontSizeVal / 2}pt`)
          }

          const colorEl = rPr.first('w:color')
          if (colorEl) {
            const colorVal = colorEl.attributes['w:val']
            result.push(`color:#${colorVal}`)
          }

          const boldEl = rPr.first('w:b')
          if (boldEl) {
            const boldVal = boldEl.attributes['w:val']
            if (boldVal !== '0') {
              result.push(`font-weight:bold`)
            }
          }

          return result.join(';') || ''
        } else {
          return ''
        }
      },
      styleMap: [
        // Headings
        "p[style-name='Heading 1'] => h1:fresh",
        "p[style-name='Heading 2'] => h2:fresh",
        "p[style-name='Heading 3'] => h3:fresh",
        "p[style-name='Heading 4'] => h4:fresh",
        "p[style-name='Heading 5'] => h5:fresh",
        "p[style-name='Heading 6'] => h6:fresh",
        "p[style-name='heading 1'] => h1:fresh",
        "p[style-name='heading 2'] => h2:fresh",
        "p[style-name='heading 3'] => h3:fresh",
        "p[style-name='heading 4'] => h4:fresh",
        "p[style-name='heading 5'] => h5:fresh",
        "p[style-name='heading 6'] => h6:fresh",

        // Blockquotes
        "p[style-name='Quote'] => blockquote.blockquote > p:fresh",
        "p[style-name='quote'] => blockquote.blockquote > p:fresh",
        "p[style-name='blockquote'] => blockquote.blockquote > p:fresh",
        "p[style-name='BlockQuote'] => blockquote.blockquote > p:fresh",

        // Code blocks
        "p[style-name='Code Block'] => pre.preCode > code:fresh",
        "p[style-name='code block'] => pre.preCode > code:fresh",
        "p[style-name='Code'] => pre.preCode > code:fresh",
        "p[style-name='code'] => pre.preCode > code:fresh",

        // Text formatting
        'b => strong',
        'i => em',
        'u => u',
        'strike => s',

        // Keep inline code
        'code => code',
      ],
    }

    try {
      const { messages, value } = await mammoth.convertToHtml(
        { arrayBuffer },
        {
          ...customOptions,
          ...$options.options,
        },
      )

      if (messages.some((msg: any) => msg.type === 'error')) {
        const errorMsg = messages.find((msg: any) => msg.type === 'error')
        useMessage('error', {
          attach: container,
          content: `${t('base.importWord.convertError')} ${errorMsg && errorMsg.message}`,
        })
        message.close()
        return
      }
      editor.value?.commands.setContent(value)
      message.close()
    } catch (error) {
      message.close()
      useMessage('error', {
        attach: container,
        content: t('base.importWord.importError'),
      })
      console.error('Import error:', error)
    }
  })
}
</script>
