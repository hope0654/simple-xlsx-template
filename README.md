# simple-xlsx-template

在浏览器中根据xlsx模板生产xlsx并下载

### 约定

- 每个sheet用**sheetN**表示，比如**sheet1**, **sheet2**...

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
import patch_template from 'simple-xlsx-template'
import xlsx_url from './simple.xlsx'

const xlsx_data = {
    sheet1: {
        A1: '满意',
        B1: 45,
    },
}

const download_filename = 'demo.xlsx'

patch_template(xlsx_url, xlsx_data, download_filename)
```

## 支持pptx模板

### 约定

- 每个slide用**slideN**表示，比如**slide1**, **slide2**...
- 每个image用**imageN**表示，比如**image1**, **image2**...
- 每个chart用**chartN**表示，比如**chart1**, **chart2**...

### 功能

- 替换文本
- 替换图片
- 修改图表

### 不支持

- 每个slide中的具体的image所对应的表示**imageN**是不确定的，需要尝试, 只有一个的话就用**image1**
- 每个slide中的具体的chart所对应的表示**chartN**是不确定的，需要尝试, 只有一个的话就用**chart1**

### 使用pptx

```javascript
import patch_template from 'simple-xlsx-template'
import pptx_url from './simple.pptx'
import image_url from './image2.png'

const pptx_data = {
    slide1: {
    	// 替换文本
        text: {
            'hello world': '你好世界',
        },
        // 替换图片
        image: {
            image1: image_url,
        },
        // 修改图表
        chart: {
            chart1: {
                sheet1: {
                    A2: '类型x',
                    B1: '系列y',
                    B2: 4.3,
                },
            },
        },
    },

}

const download_filename = 'demo.pptx'

patch_template(pptx_url, pptx_data, download_filename, 'pptx')
```

## 如何开发

1. 首选选定一个基础的版本，比如v1.xlsx或者v1.pptx
2. 用编辑器修改部分数据，并另存为第二个版本，比如v2.xlsx或者v2.pptx
3. 执行命令 `./diffxml.sh v1.xlsx v2.xlsx` 或者 `./diffxml.sh v1.pptx v2.pptx`
3. 根据diffxml的结果，编写代码实现修改
4. 打包到前端项目测试结果

## 证书

MIT
