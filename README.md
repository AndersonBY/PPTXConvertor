# PPTXCONVERTOR
Simple pptx file convertor. Powerpoint application is required.

## Installation

```
pip install pptxconvertor
```

## Example

```Python
from PPTXConvertor import PPTXConvertor

convertor = PPTXConvertor('path/to/input/folder', 'path/to/output/folder')
convertor.convert('PDF')
```

## Supported formats

[PpSaveAsFileType](https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppsaveasfiletype)

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