# Codex
![.NET Core](https://github.com/FredCof/Codex/workflows/.NET%20Core/badge.svg?branch=master)

An open source library for manipulating Word documents

Release Test

## developing......

开发参考：[Microsoft MSDN](https://docs.microsoft.com/en-us/office/open-xml/open-xml-sdk?view=openxml-2.8.1)

开发参考：[Xceed](https://github.com/xceedsoftware/DocX)

开发参考：[Microsoft dotNet](https://docs.microsoft.com/zh-cn/dotnet/csharp/programming-guide/concepts/linq/)

## 计划开发的文档元素
- [ ] 段落
- [ ] 图片
- [ ] 文档信息
- [ ] 表格
- [ ] 书签
- [ ] 超链接
- [ ] 表格
- [ ] 页眉页脚
- [ ] 文本框
- [ ] 尾注
- [ ] 脚注
- [ ] 子元素属性
- [ ] 目录
- [ ] 分栏
- [ ] 元素的复制粘贴
- [ ] 文字方向
- [ ] 图表（数据图表格）
- [ ] 页面大小
- [ ] 文档保护
- [ ] 水印
- [ ] 艺术字
- [ ] 形状
- [ ] 公式
- [ ] 批注

# 计划功能

## 转存PDF
Word格式文档转存PDF

## Markdown支持
支持基础Markdown转Word
- [ ] 标题
- [ ] 表格
- [ ] 注释
- [ ] 段落
- [ ] 图片
- [ ] 列表元素
- [ ] 代码
- [ ] 超链接
- [ ] 加粗、斜体

## LaTeX支持
支持简单的LaTeX公式转Word公式

## 类似CSS选择器的文档操作方式
计划元素类型：
- table
- tr
- tc
- para
- img
- formula
- hyperlink
- menu
- textarea
- endnote
- footnote
- watermark
- pagebreak

伪元素：
- ::first-line
- ::first-word
- ::first-character

伪类：
- :nth-child()
- :not()
- :first-type
- :first-child

选择方式：
- \+
- space
- \>
- ,
- ~

## 文档格式统一
例如：论文格式调整
使用类似CSS调整文档。
将类似CSS的文件文件类型设置为.docss。
读取文件调整格式。

# 大佬博客
[C#压缩成zip](https://blog.csdn.net/weixin_37744986/article/details/82711713)
[C#压缩](https://blog.csdn.net/zhulongxi/article/details/51819431)
[C# IO.packaging](https://www.cnblogs.com/zhaozhan/archive/2012/05/28/2520701.html)
[stackflow](https://stackoverflow.com/questions/6386113/using-system-io-packaging-to-generate-a-zip-file)