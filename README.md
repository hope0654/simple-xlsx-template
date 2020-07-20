# simple-xlsx-template

在浏览器中根据xlsx模板生产xlsx并下载

## 功能

- 保留字段原来样式
- 修改文本字段
- 修改数值字段
- 联动更新相关图表

## 不支持

- 只支持修改字段，不支持新增
- 不支持公式字段自动修改

## 安装

```sh
npm i simple-xlsx-template
```

## 使用

```javascript
import patch_xlsx from 'simple-xlsx-template'
import xlsx_url from './simple.xlsx'

const data = {
    sheet1: {
        A1: '满意',
        B1: 45,
    },
}

const download_filename = 'demo.xlsx'

patch_xlsx(xlsx_url, data, download_filename)
```

## 证书

MIT
