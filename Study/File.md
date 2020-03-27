# Open XML包结构
> Open XML文件存储在 ZIP 存档中以方便打包和压缩。可以使用 ZIP 查看器来查看任何Open XML文件的结构。一个Open XML 文档由多个文档部件构成。
>
> 这些部件之间的关系自己存储在文档部件中。**ZIP 格式支持随机访问每个部件**。例如，应用程序可以将一张幻灯片从一个演示文稿中移到另一个演示文稿中，而无需分析幻灯片内容。同样地，应用程序可以删除字处理文档中的所有注释，而不用分析文档的任何内容。

包中包含一个 [Content_Types].xml 部件，通过该部件可以确定包中所有文档部件的内容类型。 

以 .rels 扩展名结尾的关系部件中包含一组源包或部件的显式关系。

## [WordprocessingML](https://docs.microsoft.com/zh-cn/office/open-xml/working-with-wordprocessingml-documents)

- 主文档（必须）
- 词汇表文档
- 页眉和页脚
- 注释
- 文本框
- 脚注和尾注

#### 文件结构

```html
.docx
├── _rels
│  └── .rels
├── *customXml
│  ├── _rels
│  │  └── item*.xml.rels
│  ├── item*.xml
│  └── itemProps*.xml
├── docProps
│  ├── app.xml
│  ├── core.xml
│  └── *custom.xml
├── word
│  ├── _rels
│  │  └── document.xml.rels
│  ├── theme
│  │  └── theme1.xml
│  ├── *media
│  │  └── ...
│  ├── *embeddings
│  │  └── oleObject*.bin
│  ├── document.xml
│  ├── *endnotes.xml
│  ├── fontTable.xml
│  ├── *footer1.xml
│  ├── *fontnotes.xml
│  ├── *numbering.xml
│  ├── settings.xml
│  ├── style.xml
│  └── webSettings.xml
└── [Content_Types].xml
```

#### 标签属性

```html
document.xml
w:document		// 根节点
└── w:body		// 文档节点
    ├── w:p		// 段落
    │   ├── w:pPr		// 段落属性
    │   │   ├── w:pStyle
    │   │   ├── w:ind
    │   │   ├── w:shd
    │   │   ├── w:keepNext
    │   │   ├── w:keepLines
    │   │   ├── w:spacing
    │   │   ├── w:outlineLvl
    │   │   ├── w:widowControl
    │   │   ├── w:numPr
    │   │   │   ├── w:ilvl
    │   │   │   └── w:numId
    │   │   ├── w:tabs
    │   │   │   └── w:tab
    │   │   ├── w:snapToGrid
    │   │   ├── w:pBdr
    │   │   │   ├── w:top
    │   │   │   ├── w:left
    │   │   │   ├── w:bottom
    │   │   │   └── w:right
    │   │   └── w:jc
    │   ├── w:r	*						// 样式串，指明其中文本的显示样式
    │   │   ├── w:rPr
    │   │   │   ├── w:noProof
    │   │   │   ├── w:sz
    │   │   │   ├── w:color
    │   │   │   ├── w:szCs
    │   │   │   ├── w:lang
    │   │   │   ├── w:b
    │   │   │   ├── w:u
    │   │   │   ├── w:kern
    │   │   │   ├── w:vertAlign
    │   │   │   └── w:rFonts
    │   │   │       └── w:hint
    │   │   ├── w:seprarator
    │   │   ├── w:tab
    │   │   ├── w:drawing
    │   │   │   ├── w:sz
    │   │   │   ├── w:szCs
    │   │   │   ├── w:rFonts
    │   │   │   ├── w:inline
    │   │   │   │   └── 同wp:anchor
    │   │   │   └── wp:anchor
    │   │   │       ├── wp:simplePos
    │   │   │       ├── wp:positionH
    │   │   │       │   └── wp:posOffset
    │   │   │       ├── wp:positionV
    │   │   │       │   └── wp:posOffset
    │   │   │       ├── wp:extent
    │   │   │       ├── wp:effectExtent
    │   │   │       ├── wp:wrapSquare
    │   │   │       ├── wp:wrapNone
    │   │   │       ├── wp:docPr
    │   │   │       ├── wp:cNvGraphicFramePr
    │   │   │       │   └── wp:graphicFrameLocks
    │   │   │       └── a:graphic
    │   │   │           └── a:graphicData
    │   │   │               └── pic:pic
    │   │   │                   ├── pic:nvPicPr
    │   │   │                   │   ├── pic:cNvPr
    │   │   │                   │   └── pic:cNvPicPr
    │   │   │                   │       └── a:picLoacks
    │   │   │                   ├── pic:blipFill
    │   │   │                   │   ├── a:blip
    │   │   │                   │   └── a:stretch
    │   │   │                   │       └── a:fillRect
    │   │   │                   └── pic:spPr
    │   │   │                       ├── a:xfrm
    │   │   │                       │   ├── a:ext
    │   │   │                       │   └── a:ext
    │   │   │                       ├── a:prstGeom
    │   │   │                       │   └── a:avLst
    │   │   │                       ├── a:noFill
    │   │   │                       ├── a:solidFill
    │   │   │                       │   └── a:schemeClr
    │   │   │                       └── a:ln
    │   │   │                           ├── a:solidFill
    │   │   │                           │   └── a:prstClr
    │   │   │                           └── a:noFill
    │   │   ├── w:lastRenderedPageBread
    │   │   ├── w:continuationSeparator
    │   │   ├── mc:AlternateContent
    │   │   │   ├── mc:Choice
    │   │   │   │   └── w:drawing
    │   │   │   │       └── wp:anchor
    │   │   │   │         └── a:graphic
    │   │   │   │               └── a:graphicData
    │   │   │   │                   └── wsp:wsp
    │   │   │   │                       ├── wsp:cNvSpPr
    │   │   │   │                       ├── wsp:spPr
    │   │   │   │                       │   └── 同pic:spPr
    │   │   │   │                       ├── wsp:style
    │   │   │   │                       │   ├── a:lnRef
    │   │   │   │                       │   │   └── a:schemeClr
    │   │   │   │                       │   ├── a:fillRef
    │   │   │   │                       │   │   └── a:schemeClr
    │   │   │   │                       │   ├── a:effecrRef
    │   │   │   │                       │   │   └── a:schemeClr
    │   │   │   │                       │   ├── a:fontRef
    │   │   │   │                       │   │   └── a:schemeClr
    │   │   │   │                       ├── wsp:txbx
    │   │   │   │                       │   └── w:txbxContent
    │   │   │   │                       │       └── w:p
    │   │   │   │                       └── wsp:bodyPr
    │   │   │   │                           └── a:noAutofit
    │   │   │   └── mc:Fallback
    │   │   │       └── w:pict
    │   │   │           ├── v:shapetype
    │   │   │           │   ├── v:stroke
    │   │   │           │   └── v:path
    │   │   │           └── v:shape
    │   │   │               └── v:textbox
    │   │   │                   └── w:txbxContent
    │   │   └── w:t
    │   ├── w:hyperlink
    │   │   └── w:r
    │   ├── w:proofErr
    │   ├── w:bookmarkStart
    │   └── w:bookmarkEnd
    ├── w:tbl
    │   ├── w:tblPr
    │   │   ├── w:tblStyle
    │   │   ├── w:tblW
    │   │   ├── w:tblLayout
    │   │   ├── w:tblInd
    │   │   ├── w:tblCellMar
    │   │   │   ├── w:top
    │   │   │   ├── w:left
    │   │   │   ├── w:bottom
    │   │   │   └── w:right
    │   │   ├── w:tblBorders
    │   │   │   ├── w:top
    │   │   │   ├── w:left
    │   │   │   ├── w:bottom
    │   │   │   ├── w:right
    │   │   │   ├── w:insideH
    │   │   │   └── w:insideV
    │   │   └── w:tblLook
    │   ├── w:tblGrid
    │   │   └── w:gridCol *
    │   └── w:tr *
    │       └── w:tc *
    │           ├── w:tcPr
    │           │   └── w:tcW
    │           └── w:p *
    └── w:sectPr
        ├── w:hdr
        ├── w:ftr
        ├── w:footerReference
        ├── w:pgSz
        ├── w:pgMar
        ├── w:cols
        └── w:docGrid

endnotes.xml
w:endnotes
└── w:endnote
    └── w:p

fontTable.xml
w:fonts
└── w:font
    ├── w:altName
    ├── w:panosel
    ├── w:charset
    ├── w:family
    ├── w:pitch
    └── w:sig

footer1.xml
w:ftr
└── w:p

footnotes.xml
w:footnotes
└── w:footnote
    └── w:p

numbering.xml
w:numbering
├── w:num
│   └── w:abstractNumId
└── w:absractNum
    ├── w:nsid
    ├── w:multiLevelType
    ├── w:tmpl
    └── w:lvl *
        ├── w:start
        ├── w:numFmt
        ├── w:lvlText
        ├── w:lvlJc
        ├── w:pPr
        └── w:rPr

setting.xml
w:settings
├── w:zoom
├── w:embedSystemFonts
├── w:bordersDoNotSurroundHeader
├── w:bordersDoNotSurroundFooter
├── w:proofState
├── w:defaultTabStop
├── w:drawimgGridVerticalSpace
├── w:noPunctuationKerning
├── w:characterSpacingControl
├── w:hdrShapeDefaults
│   └── o:shapedefaults
│       └── v:fill
├── w:footnotePr
│   └── w:footnote
├── w:endnotePr
│   └── w:endnote
├── w:compat
│   ├── w:spaceForUL
│   ├── w:balanceSingleByteDoubleByteWidth
│   ├── w:doNotLeaveBackslashAlone
│   ├── w:doNotExpandShiftReturn
│   ├── w:adjustLineHeightInTable
│   ├── w:useFELayout
│   └── w:compatSetting
├── w:rsids
│   ├── w:rsidRoot
│   └── w:rsid
├── w:mathPr
│   ├── w:mathFont
│   ├── w:brkBin
│   ├── w:brkBinSub
│   ├── w:smallFrac
│   ├── w:dispDef
│   ├── w:lMargin
│   ├── w:rMargin
│   ├── w:defJc
│   ├── w:wrapIndent
│   ├── w:intLim
│   └── w:naryLim
├── w:themeFontLang
├── w:clrSchemeMapping
├── w:doNotIncludeSubdocsInStats
├── w:shapeDefaults
│   ├── o:shapedefaults
│   │   └── o:fill
│   └── o:shapelayout
│       └── o:idmap
├── w:decimalSymbol
├── w:listSeparator
├── w14:docId
└── w15:docId

styles.xml
w:styles
├── w:docDefaults
│   ├── w:rPrDefault
│   │   ├── w:lang
│   │   └── w:rFonts
│   └── w:pPrDefault
├── w:latentStyles
│   └── w:lsdException
└── w:style
    ├── w:name
    ├── w:baseOn
    ├── w:next
    ├── w:link
    ├── w:qFormat
    ├── w:unhideWhenUsed
    ├── w:uiPriority
    ├── w:semiHidden
    ├── w:pPr
    ├── w:rsid
    ├── w:tblPr
    └── w:rPr

webSettings.xml
w:webSettings
└── w:divs
    ├── w:div
    │   ├── w:bodyDiv
    │   ├── w:marLeft
    │   ├── w:marRight
    │   ├── w:marTop
    │   ├── w:marBottom
    │   └── w:divBdr
    │       ├── w:top
    │       ├── w:left
    │       ├── w:bottom
    │       └── w:right
    └── w:divsChild
        └── w:div
```

节点解释：

- w:p						段落
  - w14:paraId     段落唯一ID
  - w14:textId       文字唯一ID
  - w:rsidR            对应setting.xml中的w:rsids>w:rsid[w:val]
  - w:rsidRDefault
  - w:rsidP
- w:pStyle                  段落样式
  - w:val                 对应style.xml中的w:style[w:styleId]
- w:t                            文字
- w:rFonts
  - w:hint				 字体名
- w:style                      样式
  - w:name              对应w:latentStyles>w:lsdException[w:name]
  - w:baseOn          继承自
  - w:next                这样式的下一个样式，对应w:style[w:styleId]

## [PresentationML](https://docs.microsoft.com/zh-cn/office/open-xml/working-with-presentationml-documents)

- 幻灯片母版
- 备注母版
- 讲义母版
- 幻灯片版式
- 说明

## [SpreadsheetML](https://docs.microsoft.com/zh-cn/office/open-xml/working-with-spreadsheetml-documents)
- 工作簿部件（必须）
- 一张或多张工作表
- 图表
- 自定义XML