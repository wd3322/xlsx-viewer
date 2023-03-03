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

xlsxViewer.renderXlsx(data, document.querySelector('div'))
```

| Prop        | Prop Type  | Type                    | Required |
| :-------    | :-------   | :-------                | :------  |
| data        | Attribute  | ArrayBuffer, Blob, File | True     |
| element     | Attribute  | HTMLElement             | True     |
