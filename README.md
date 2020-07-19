# simple-xlsx-template
generate a new xlsx from a xlsx template


## 安装

```sh
npm i simple-xlsx-template
```

## 使用

```javascript
import patch_xlsx from 'simple-xlsx-template'
import xlsx_url from '../lib/v2.xlsx'

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