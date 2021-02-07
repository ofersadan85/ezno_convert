import enum
from typing import Union


@enum.unique
class PPT(enum.Enum):
    # Source: https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppsaveasfiletype
    AnimatedGIF = 40
    BMP = 19
    Default = 11
    EMF = 23
    External = 64000
    GIF = 16
    JPG = 17
    META = 15
    MP4 = 39
    OpenPresentation = 35
    PDF = 32
    PNG = 18
    Presentation = 1
    RTF = 6
    SHOW = 7
    Template = 5
    TIF = 21
    WMV = 37
    XPS = 33
    app = 'Powerpoint.Application'
    extensions = ('.ppt', '.pptx')


@enum.unique
class WORD(enum.Enum):
    # Source: https://docs.microsoft.com/en-us/office/vba/api/word.wdsaveformat
    DosText = 4
    DosTextLineBreaks = 5
    FilteredHTML = 10
    FlatXML = 19
    OpenDocumentText = 23
    HTML = 8
    RTF = 6
    Template = 1
    Text = 2
    TextLineBreaks = 3
    UnicodeText = 7
    WebArchive = 9
    XML = 11
    Document97 = 0
    DocumentDefault = 16
    PDF = 17
    XPS = 18
    app = 'Word.Application'
    extensions = ('.doc', '.docx')


@enum.unique
class XL(enum.Enum):
    # Source: https://docs.microsoft.com/en-us/office/vba/api/excel.xlfixedformattype
    # TODO: Implement "SaveAs" methods, see: https://docs.microsoft.com/en-us/office/vba/api/excel.workbook.saveas
    PDF = 0
    XPS = 1
    app = 'Excel.Application'
    extensions = ('.xls', '.xlsx')


enum_types = Union[PPT, WORD, XL]
