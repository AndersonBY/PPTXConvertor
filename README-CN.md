# PPTXCONVERTOR
PPT和PPTX文件批量转换工具，要求电脑中装有PowerPoint。

## 安装

```
pip install PPTXConvertor
```

## 示例

```Python
from PPTXConvertor import PPTXConvertor

convertor = PPTXConvertor('path/to/input/folder', 'path/to/output/folder')
convertor.convert('PDF')
convertor.convert('JPG')
```

## 支持的格式

根据 [PpSaveAsFileType](https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppsaveasfiletype)，可以支持以下格式。

convert()函数中必须填入以下格式的字符串，否则不支持。

- AddIn
- AnimatedGIF
- BMP
- Default
- EMF
- ExternalConverter
- GIF
- JPG
- MetaFile
- MP4
- OpenDocumentPresentation
- OpenXMLAddin
- OpenXMLPicturePresentation
- OpenXMLPresentation
- OpenXMLPresentationMacroEnabled
- OpenXMLShow
- OpenXMLShowMacroEnabled
- OpenXMLTemplate
- OpenXMLTemplateMacroEnabled
- OpenXMLTheme
- PDF
- PNG
- Presentation
- RTF
- Show
- StrictOpenXMLPresentation
- Template
- TIF
- WMV
- XMLPresentation
- XPS