# Xlsx Viewer

![table](https://wd3322.gitee.io/to-vue3/img/xlsx-viewer/table.png)

## Install

```
npm install xlsx-viewer
```

## Import

```javascript
import xlsxViewer from 'xlsx-viewer'
import 'xlsx-viewer/src/style.css'
```

## Render

default

```javascript
xlsxViewer.renderXlsx(data, document.querySelector('div'))
```

append options

```javascript
xlsxViewer.renderXlsx(data, document.querySelector('div'), {
  initialSheetIndex: 0,
  frameRenderSize: 500,
  onLoad(sheets) {
    console.log('onLoad', sheets)
  },
  onRender(sheet) {
    console.log('onRender', sheet)
  },
  onSwitch(sheet) {
    console.log('onSwitch', sheet)
  }
})
```


| Prop        | Prop Type  | Type                    | Required |
| :-------    | :-------   | :-------                | :------  |
| data        | Attribute  | ArrayBuffer, Blob, File | True     |
| element     | Attribute  | HTMLElement             | True     |
| opitons     | Attribute  | Object                  | False    |
