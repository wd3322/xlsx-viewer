/*
* package: xlsx-viewer
* e-mail: diquick@qq.com
* author: wd3322
*/

import Dayjs from 'dayjs'
import ExcelJS from 'exceljs'

interface xlsxOptions {
  initialSheetIndex?: number,
  frameRenderSize?: number,
  onLoad?: Function,
  onRender?: Function,
  onSwitch?: Function
}

interface SheetItem {
  id: number,
  name: string,
  columns: any[],
  rows: any[],
  merges: any[],
  worksheet: any,
  rendered: boolean
}

interface ViewerParams {
  arrayBuffer?: ArrayBuffer,
  sheetList: SheetItem[],
  currentSheetId?: number
}

interface ViewerElements {
  xlsxElement?: HTMLElement,
  containerElement?: HTMLElement,
  tipElement?: HTMLElement,
  sheetElement?: HTMLElement,
  tableElement?: HTMLElement
}

interface ViewerMethods {
  loadXlsxDataWorkbook: Function,
  createXlsxContainerElement: Function,
  createTableContainerElement: Function,
  createTableContentElement: Function,
}

interface UtilMethods {
  blobOrFileToArrayBuffer: Function,
  parseARGB: Function
}

function getPrototype(value: any): string {
  return Object.prototype.toString.call(value).replace(/^\[object (\S+)\]$/, '$1').toLowerCase()
}

async function renderXlsx(
  xlsxData: Blob | File | ArrayBuffer,
  xlsxElement: HTMLElement,
  xlsxOptions: xlsxOptions = {}
): Promise<void> {
  const {
    initialSheetIndex = 0,
    frameRenderSize = 500,
    onLoad = () => {},
    onRender = () => {},
    onSwitch = () => {}
  }: xlsxOptions = xlsxOptions
  if (!['blob', 'file', 'arraybuffer'].includes(getPrototype(xlsxData))) {
    throw new Error(`renderXlsx ${xlsxData} is not a file`)
  }
  if (getPrototype(xlsxElement).indexOf('element') === -1) {
    throw new Error(`renderXlsx ${xlsxElement} is not a element`)
  }
  if (getPrototype(xlsxOptions) !== 'object') {
    throw new Error(`renderXlsx ${xlsxOptions} is not a object`)
  }
  if (getPrototype(initialSheetIndex) !== 'number') {
    throw new Error('renderXlsx \'initialSheetIndex\' is not a number')
  }
  if (getPrototype(frameRenderSize) !== 'number') {
    throw new Error('renderXlsx \'frameRenderSize\' is not a number')
  }
  if (getPrototype(onLoad).indexOf('function') !== -1) {
    throw new Error('renderXlsx \'onLoad\' is not a function')
  }
  if (getPrototype(onRender).indexOf('function') !== -1) {
    throw new Error('renderXlsx \'onRender\' is not a function')
  }
  if (getPrototype(onSwitch).indexOf('function') !== -1) {
    throw new Error('renderXlsx \'onSwitch\' is not a function')
  }
  // viewer params init
  const viewerParams: ViewerParams = {
    arrayBuffer: undefined,
    sheetList: [],
    currentSheetId: undefined
  }
  // viewer elements init
  const viewerElements: ViewerElements = {
    xlsxElement: undefined,
    containerElement: undefined,
    tipElement: undefined,
    sheetElement: undefined,
    tableElement: undefined
  }
  // viewer methods init
  const viewerMethods: ViewerMethods = {
    async loadXlsxDataWorkbook(): Promise<void> {
      return new Promise((resolve: Function) => {
        try {
          // load workbook
          (new ExcelJS.Workbook().xlsx.load((viewerParams.arrayBuffer as ArrayBuffer))).then((workbook: any) => {
            workbook.eachSheet((worksheet: any, sheetId: number) => {
              const sheetItem: SheetItem = {
                id: sheetId,
                name: worksheet.name,
                columns: [],
                rows: [],
                merges: [],
                worksheet,
                rendered: false
              }
              // set sheet column
              for (let i = 0; i < worksheet.columnCount; i++) {
                const column: any = worksheet.getColumn(i + 1)
                sheetItem.columns.push(column)
              }
              // set sheet row
              for (let i = 0; i < worksheet.rowCount; i++) {
                const row: any = worksheet.getRow(i + 1)
                // set sheet row cell merges
                for (let j = 0; j < row.cellCount; j++) {
                  const cell: any = row.getCell(j + 1)
                  if (cell.isMerged) {
                    const targetAddress: any = sheetItem.merges.find((item: any) => item.address === cell.master._address)
                    if (targetAddress) {
                      targetAddress.cells.push(cell)
                    } else {
                      sheetItem.merges.push({
                        address: cell._address,
                        master: cell,
                        cells: [cell]
                      })
                    }
                  }
                }
                sheetItem.rows.push(row)
              }
              viewerParams.sheetList.push(sheetItem)
            })
            if (viewerElements.tipElement && getPrototype(viewerElements.tipElement).indexOf('element') !== -1) {
              viewerElements.tipElement.style.display = 'none'
              onLoad(viewerParams.sheetList)
            }
            resolve()
          })
        } catch (err) {
          if (viewerElements.tipElement && getPrototype(viewerElements.tipElement).indexOf('element') !== -1) {
            viewerElements.tipElement.innerText = `Load error: ${err}`
          }
          console.error('[xlsx-viewer] load error: ', err)
        }
      })
    },
    createXlsxContainerElement(): void {
      const xlsxViewerContainerElement: HTMLElement = document.createElement('div')
      const xlsxViewerSheetElement: HTMLElement = document.createElement('div')
      const xlsxViewerTableElement: HTMLElement = document.createElement('div')
      const xlsxViewerTipElement: HTMLElement = document.createElement('div')
      const oldXlsxViewerContainerElement: HTMLElement | null = xlsxElement.querySelector('.xlsx-viewer-container')
      xlsxViewerContainerElement.classList.add('xlsx-viewer-container')
      xlsxViewerSheetElement.classList.add('xlsx-viewer-sheet')
      xlsxViewerTableElement.classList.add('xlsx-viewer-table')
      xlsxViewerTipElement.classList.add('xlsx-viewer-tip')
      xlsxViewerTipElement.innerText = 'Loading...'
      viewerElements.xlsxElement = xlsxElement
      viewerElements.sheetElement = xlsxViewerSheetElement
      viewerElements.tableElement = xlsxViewerTableElement
      viewerElements.tipElement = xlsxViewerTipElement
      viewerElements.containerElement = xlsxViewerContainerElement
      viewerElements.containerElement.appendChild(xlsxViewerSheetElement)
      viewerElements.containerElement.appendChild(xlsxViewerTableElement)
      viewerElements.containerElement.appendChild(xlsxViewerTipElement)
      if (oldXlsxViewerContainerElement) {
        viewerElements.xlsxElement.replaceChild(viewerElements.containerElement, oldXlsxViewerContainerElement)
      } else {
        viewerElements.xlsxElement.appendChild(viewerElements.containerElement)
      }
    },
    createTableContainerElement(): void {
      for (let i = 0; i < viewerParams.sheetList.length; i++) {
        const sheetItem: SheetItem = viewerParams.sheetList[i]
        const xlsxViewerSheetItemElement: HTMLElement = document.createElement('div')
        const xlsxViewerTableItemElement: HTMLElement = document.createElement('div')
        xlsxViewerSheetItemElement.innerText = sheetItem.name
        xlsxViewerSheetItemElement.classList.add('xlsx-viewer-sheet-content')
        xlsxViewerTableItemElement.classList.add('xlsx-viewer-table-content')
        xlsxViewerSheetItemElement.addEventListener('click', (e: Event) => {
          viewerElements.sheetElement?.querySelector('.xlsx-viewer-sheet-content.active')?.classList.remove('active')
          viewerElements.tableElement?.querySelector('.xlsx-viewer-table-content.active')?.classList.remove('active')
          xlsxViewerSheetItemElement.classList.add('active')
          xlsxViewerTableItemElement.classList.add('active')
          if (!sheetItem.rendered) {
            viewerMethods.createTableContentElement(sheetItem, xlsxViewerTableItemElement)
          }
          if (sheetItem.id !== viewerParams.currentSheetId) {
            viewerParams.currentSheetId = sheetItem.id
            onSwitch(sheetItem)
          }
        })
        if (i === (initialSheetIndex >= 0 && initialSheetIndex < viewerParams.sheetList.length ? initialSheetIndex : 0)) {
          viewerParams.currentSheetId = sheetItem.id
          xlsxViewerSheetItemElement.classList.add('active')
          xlsxViewerTableItemElement.classList.add('active')
          viewerMethods.createTableContentElement(sheetItem, xlsxViewerTableItemElement)
        }
        viewerElements.sheetElement?.appendChild(xlsxViewerSheetItemElement)
        viewerElements.tableElement?.appendChild(xlsxViewerTableItemElement)
      }
    },
    createTableContentElement(
      sheetItem: SheetItem,
      xlsxViewerTableItemElement: HTMLElement
    ): void {
      // set table element
      const tableElement: HTMLElement = document.createElement('table')
      const theadElement: HTMLElement = document.createElement('thead')
      const tbodyElement: HTMLElement = document.createElement('tbody')
      const tbodyTrElementArr: HTMLElement[] = []
      const appendTrElementToTbodyElement = (currentPage: number = 0) => {
        requestAnimationFrame(() => {
          for (let i = 0; i < frameRenderSize; i++) {
            const trElement: HTMLElement = tbodyTrElementArr[currentPage * frameRenderSize + i]
            if (trElement) {
              tbodyElement.appendChild(trElement)
            } else {
              break
            }
          }
          if (currentPage * frameRenderSize < tbodyTrElementArr.length) {
            appendTrElementToTbodyElement(currentPage + 1)
          } else {
            sheetItem.rendered = true
            onRender(sheetItem)
          }
        })
      }
      // set sheet columns element
      if (sheetItem.columns.length > 0) {
        const trElement: HTMLElement = document.createElement('tr')
        const firstThElement: HTMLElement = document.createElement('th')
        let tableWidth: number = 50
        firstThElement.style.width = '50px'
        trElement.appendChild(firstThElement)
        for (let i = 0; i < sheetItem.columns.length; i++) {
          const column: any = sheetItem.columns[i]
          const thElement: HTMLElement = document.createElement('th')
          const columnWidth: number = column.width > 0 ? column.width / 0.125 : 100
          tableWidth = tableWidth + columnWidth
          thElement.style.width = `${columnWidth}px`
          thElement.innerText = column.letter
          trElement.appendChild(thElement)
        }
        theadElement.appendChild(trElement)
        tableElement.style.width = `${tableWidth}px`
      }
      // set sheet rows element
      if (sheetItem.rows.length > 0) {
        for (let i = 0; i < sheetItem.rows.length; i++) {
          const row: any = sheetItem.rows[i]
          const trElement: HTMLElement = document.createElement('tr')
          const firstTdElement: HTMLElement = document.createElement('td')
          firstTdElement.innerText = (i + 1).toString()
          trElement.appendChild(firstTdElement)
          for (let j = 0; j < sheetItem.columns.length; j++) {
            const cell: any = row.getCell(j + 1)
            if (cell.isMerged && cell.master._address !== cell._address) {
              continue
            }
            const tdElement: HTMLElement = document.createElement('td')
            if (cell.isMerged && cell.master._address === cell._address) {
              const merge: any = sheetItem.merges.find(item => item.address === cell._address)
              if (merge) {
                const maxCol: number = Math.max.apply(Math, merge.cells.map((cell: any) => cell.col))
                const maxRow: number = Math.max.apply(Math, merge.cells.map((cell: any) => cell.row))
                const colSpan: number = maxCol - cell.col + 1
                const rowSpan: number = maxRow - cell.row + 1
                tdElement.setAttribute('colspan', colSpan.toString())
                tdElement.setAttribute('rowSpan', rowSpan.toString())
              }
            }
            // set row size
            if (row.height) {
              tdElement.style.height = `${row.height / 0.75}px`
            }
            // set cell alignment
            if (cell.style?.alignment) {
              const { horizontal, vertical } = cell.style.alignment
              tdElement.style.textAlign = horizontal
              tdElement.style.verticalAlign = vertical
            }
            // set cell background
            if (cell.style?.fill) {
              const { fgColor } = cell.style.fill
              tdElement.style.backgroundColor = fgColor?.argb ? (utilMethods.parseARGB(fgColor?.argb)?.color as string) : '#fff'
            }
            // set cell border
            if (cell.style?.border) {
              const { top, bottom, left, right } = cell.style.border
              tdElement.style.borderTop = top?.color?.argb ? '1px solid ' + (utilMethods.parseARGB(top?.color?.argb)?.color as string) : ''
              tdElement.style.borderBottom = bottom?.color?.argb ? '1px solid ' + (utilMethods.parseARGB(bottom?.color?.argb)?.color as string) : ''
              tdElement.style.borderLeft = left?.color?.argb ? '1px solid ' + (utilMethods.parseARGB(left?.color?.argb)?.color as string) : ''
              tdElement.style.borderRight = right?.color?.argb ? '1px solid ' + (utilMethods.parseARGB(right?.color?.argb)?.color as string) : ''
            }
            // set cell font
            if (cell.style?.font) {
              const { color, name, size, bold, italic, underline } = cell.style.font
              tdElement.style.color = color?.argb ? (utilMethods.parseARGB(color?.argb)?.color as string) : '#333'
              tdElement.style.fontFamily = name
              tdElement.style.fontSize = size ? `${size / 0.75}px` : '14px'
              tdElement.style.fontWeight = bold ? 'bold' : 'normal'
              tdElement.style.fontStyle = italic ? 'italic' : 'normal'
              tdElement.style.textDecoration = underline ? 'underline' : 'none'
            }
            // set cell value
            if (getPrototype(cell.value) === 'object') {
              const { richText, hyperlink } = cell.value
              if (richText && getPrototype(richText) === 'array') {
                for (const span of richText) {
                  const spanElement: HTMLElement = document.createElement('span')
                  if (span?.font) {
                    const { color, name, size, bold, italic, underline } = span.font
                    spanElement.style.color = color?.argb ? (utilMethods.parseARGB(color?.argb)?.color as string) : '#333'
                    spanElement.style.fontFamily = name
                    spanElement.style.fontSize = size ? `${size / 0.75}px` : '14px'
                    spanElement.style.fontWeight = bold ? 'bold' : 'normal'
                    spanElement.style.fontStyle = italic ? 'italic' : 'normal'
                    spanElement.style.textDecoration = underline ? 'underline' : 'none'
                  }
                  spanElement.innerText = span.text
                  tdElement.appendChild(spanElement)
                }
              } else if (getPrototype(hyperlink) === 'string') {
                const link = cell.value
                const linkElement: HTMLElement = document.createElement('a')
                linkElement.setAttribute('href', link.hyperlink)
                linkElement.setAttribute('target', '_blank')
                linkElement.style.color = '#2f63c1'
                linkElement.style.textDecoration = 'underline'
                linkElement.innerText = link.text
                tdElement.appendChild(linkElement)
              }
            } else if (getPrototype(cell.value) === 'date') {
              tdElement.innerText = Dayjs(cell.value).format('YYYY-MM-DD HH:mm:ss')
            } else {
              tdElement.innerText = cell.value
            }
            trElement.appendChild(tdElement)
          }
          tbodyTrElementArr.push(trElement)
        }
        appendTrElementToTbodyElement()
      }
      tableElement.appendChild(theadElement)
      tableElement.appendChild(tbodyElement)
      xlsxViewerTableItemElement.appendChild(tableElement)
    }
  }
  // util methods init
  const utilMethods: UtilMethods = {
    blobOrFileToArrayBuffer(blob: Blob | File): Promise<ArrayBuffer> {
      return new Promise(resolve => {
        const fileReader: FileReader = new FileReader()
        fileReader.onload = (e: any) => {
          resolve(e.target.result)
        }
        fileReader.readAsArrayBuffer(blob)
      })
    },
    parseARGB(argb: string): {
      argb: { a: number, r: number, g: number, b: number},
      color: string
    } | undefined {
      if (getPrototype(argb) !== 'string' || argb.length !== 8) {
        return undefined
      }
      let result: any
      const color: string[] = []
      for (let i = 0; i < 4; i++) {
        color.push(argb.substr(i * 2, 2))
      }
      const [a, r, g, b] = color.map((v) => parseInt(v, 16))
      result = {
        argb: { a, r, g, b },
        color: `rgba(${r}, ${g}, ${b}, ${a / 255})`
      }
      return result
    }
  }
  // check browser compatibility
  viewerMethods.createXlsxContainerElement()
  if (
    (window.navigator.userAgent.indexOf('MSIE') !== -1 || 'ActiveXObject' in window) &&
    (viewerElements.tipElement && getPrototype(viewerElements.tipElement).indexOf('element') !== -1)
  ) {
    viewerElements.tipElement.innerText = 'Browser incompatibility.'
    return
  }
  // load xlsx data
  if (['blob', 'file'].includes(getPrototype(xlsxData))) {
    viewerParams.arrayBuffer = await utilMethods.blobOrFileToArrayBuffer(xlsxData)
  } else if (getPrototype(xlsxData) === 'arraybuffer') {
    viewerParams.arrayBuffer = (xlsxData as ArrayBuffer)
  }
  await viewerMethods.loadXlsxDataWorkbook()
  viewerMethods.createTableContainerElement()
}

export default { renderXlsx }
