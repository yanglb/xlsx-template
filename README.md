# template-go
一款用于从 Excel 模板中渲染数据的工具，支持Windows、Linux、MacOS。

| 模板   |      数据      |  输出 |
|-------|---------------|-------|
| Hello ${name} | {name: "Alice"} | Hello Alice |
| ${home\|link:content} | {home: "https://yanglb.com", content: "Home Page"} | [Home Page](https://yanglb.com) |
| ${avatar\|image:wf=2cm} | {avatar: "https://avatars.githubusercontent.com/u/6257395?s=40&v=4"} | ![](https://avatars.githubusercontent.com/u/6257395?s=40&v=4) |
> 更多内容请参考下文

## 安装
```sh
dotnet add package TemplateGO

# 支持 Windows、Linux、MacOS
```

## 使用
```c#
var templateFile = "{template file path}";
var jsonString = "{json data}";
TemplateGO.Render(templateFile, jsonString, "out.xlsx");
```

## 文档
### 模板语法
${ **property** [| **processor** [: **options** ]]}

### property 属性路径
指示 TemplateGO 如何从JSON数据中获取要插入的内容。

如下所示:
```js
// data:
{
  "name": "Alice",
  "score": [{
    "course": "physics",
    "score": 85
  }],
  "school": {
    "name": "北京大学",
    "address": "北京市海淀区"
  }
}
// property:
// name => Alice
// score[0].score => 85
// school.name => "北京大学"
// notExists => undefined
```

### processor 处理单元
指定数据处理单元，目前支持以下几种:
| 处理单元  |  作用  |  说明 |
|:------|:--------------|:------|
| value (default) | 文字替换 | **默认** 使用property属性值替换标记块。<br /> **Excel:** 如果单元格仅包含该标记块，则根据JSON数据类型自动设置单元格数据类型。|
| link | 超链接 | 将property所指内容视为链接地址，向文档添加超链接。 |
| image | 图片 | 将property所指内容视为图片并插入到标记块所在位置。<br />支持本地文件、base64编码图片以及网络图片。 |
| qr | 二维码 | 将property所指内容转为二维码并插入到标记块所在位置。|
| table | 表格 | 向文档添加列表数据，property所指属性必需为数组。 |

### value 文字替换
#### 语法
> \${**property**} 或 ${property|value}

其中value为默认处理单元，可省略。

#### 示例
```json
{"name": "Alice", "age": 18, "school": {}, "graduated": false }
```
| 标记块 | 输出 | 说明 |
| ----  | ---- | ---- |
| ${name} | Alice | 单元格类型为 String |
| ${age} | 18 | Number | 单元格类型为 Number |
| Age: ${age} | Age: 18 | 单元格类型为 String |
| ${graduated} | FALSE | 单元格类型为 Boolean |
| Is Graduated: ${graduated} | IS Graduated: FALSE | 单元格类型为 String |
| ${school} | [object Object] | TemplateGO 会将 JSON 中 Object或Array 转为 [object Object] 字符串。|
| ${notExists} |  | JSON 中不存在的属性将不输出任何内容。 |

**Excel:** 如果单元格仅包含该标记块，则单元格数据类型将根据JSON数据设置。

### link 超链接
#### 语法
> ${property|link[:content=link content property]}
* property 作为链接地址，如 https://yanglb.com
* content 作为链接显示内容，如 首页

#### 选项
| 名称 | 类型 | 可选 | 说明
| -- | -- | -- | -- |
| **content** | String | 是 | 链接内容，如用英文引号("")包裹则视为字符串直接插入，否则将其当作属性key从JSON数据中获取。 |

#### 示例
```json
{"title": "Home Page", "url": "https://yanglb.com" }
```
| 标记块 | 输出 | 说明 |
| ----  | ---- | ---- |
| ${url\|link:title} | [Home Page](https://yanglb.com) | 输出超链接 |
| ${url\|link:"Test Page"} | [Test Page](https://yanglb.com) | 输出超链接 |
| My Home Page${url\|link} | [My Home Page](https://yanglb.com) | **Excel 特有用法** Excel将超链接作用到单元格之上，因此整个单元格可点击。 |

## License
Copyright (c) yanglb.com. All rights reserved.

Licensed under the MIT license.
