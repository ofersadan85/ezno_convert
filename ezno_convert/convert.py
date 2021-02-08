import enum
import logging
from dataclasses import dataclass
from os import PathLike
from pathlib import Path
from typing import Collection, Union, Optional

from comtypes import COMError
from comtypes.client import CreateObject

from ezno_convert.common import validate_paths, multi_glob
from ezno_convert.enums import PPT, WORD, XL, enum_types

logger = logging.getLogger('NativeOfficeConverter')


def word_convert(word_app, src: Path, dst: Path, target: WORD) -> Optional[Path]:
    doc = word_app.Documents.Open(str(src))
    try:
        doc.SaveAs(str(dst), FileFormat=target.value)
        return dst
    finally:
        doc.Close()


def ppt_convert(ppt_app, src: Path, dst: Path, target: PPT) -> Optional[Path]:
    doc = ppt_app.Presentations.Open(str(src))
    try:
        doc.SaveAs(str(dst), FileFormat=target.value)
        return dst
    finally:
        doc.Close()


def xl_convert(xl_app, src: Path, dst: Path, target: XL, sheets: Union[Collection, bool]) -> Optional[Path]:
    doc = xl_app.Workbooks.Open(str(src))
    try:
        if sheets:
            if sheets is True:  # Identical to True as opposed to evaluated as True - Meaning export all sheets
                sheets = [sh.Name for sh in doc.Sheets]
            for sheet in sheets:
                sheet_object = doc.Sheets(sheet)
                sheet_dst = dst.with_stem(f'{dst.stem}-{sheet_object.Name}')
                sheet_object.ExportAsFixedFormat(target.value, str(sheet_dst))
            return dst
        else:
            doc.ExportAsFixedFormat(target.value, str(dst))
            return dst
    finally:
        doc.Close()


def convert_one(
        src: PathLike,
        dst: Optional[PathLike] = None,
        app_object=None,
        target: Optional[enum_types] = None,
        date_fmt: Optional[str] = None,
        sheets: Union[Collection, bool] = False) -> Optional[Path]:
    src, dst = validate_paths(src, dst, date_fmt)

    app_opened_here = False
    match = False
    for app in (WORD, PPT, XL):
        if src.suffix in app.extensions.value:
            match = True
            if app_object is None:
                app_opened_here = True
                app_object = CreateObject(app.app.value)
            break
    if not match:
        raise ValueError(f'Unknown file extension {src.suffix} ({src})')

    target = target if target else app.PDF

    if app is WORD:
        result = word_convert(app_object, src, dst, target)
    elif app is PPT:
        result = ppt_convert(app_object, src, dst, target)
    elif app is XL:
        result = xl_convert(app_object, src, dst, target, sheets)
    else:
        raise RuntimeError(f'Function ran without an app defined: {app} ({app_object})')

    if app_opened_here:
        try:
            app_object.Quit()
        finally:
            pass
    return next(result.parent.glob(result.name + '*'))


@dataclass
class BatchConverter:
    src: Union[Collection[PathLike], PathLike]
    dst: Optional[PathLike] = None
    app: Optional[enum.EnumMeta] = None
    target: Optional[enum_types] = None
    recursive: bool = False
    date_fmt: Optional[str] = None
    sheets: Union[Collection, bool] = False

    def __post_init__(self):
        self.app_object = None

        if isinstance(self.src, (str, PathLike)):
            self.src = [self.src]

        if self.app is None:
            # Try to guess app if not supplied by the extension of the first src path
            guess_apps = [app for app in (WORD, PPT, XL) if Path(tuple(self.src)[0]).suffix in app.extensions.value]
            if guess_apps:
                self.app = guess_apps[0]
            else:
                raise RuntimeError('Could not guess correct app and none was provided')

        self.files = [Path(f) for f in self.src if Path(f).is_file()]
        extensions = [x if '*' in x else '*' + x for x in self.app.extensions.value]
        for d in (Path(f) for f in self.src if Path(f).is_dir()):
            self.files += multi_glob(d, extensions, recursive=self.recursive)

        self.files = {f for f in self.files if not f.name.startswith('~$')}

        if self.dst is not None and len(self.files) > 1:
            self.dst = Path(self.dst)
            if not Path(self.dst).is_dir():
                raise NotADirectoryError(f'Destination for batch conversion must be a folder (or empty) ({self.dst})')

    def __len__(self):
        return len(self.files)

    def __iter__(self):
        if self.files:
            self.app_object = CreateObject(self.app.app.value)
            for i, f in enumerate(self.files):
                try:
                    result = convert_one(f, self.dst, self.app_object, self.target, self.date_fmt, self.sheets)
                except (COMError, FileNotFoundError, NotADirectoryError, ValueError):
                    logger.exception(f'Failed to convert: {f}')
                    yield None
                else:
                    yield result
            self.app_object.Quit()

    def execute_all(self, output: bool = False) -> list[Optional[Path]]:
        all_results = []
        for result in self:
            all_results.append(result)
            if output:
                print(f'Success: {result}' if result else 'Conversion failed, see logs for details')
        return all_results


@dataclass
class WORDConverter(BatchConverter):
    """ Alias for BatchConverter(WORD, ...) """
    app: enum.EnumMeta = WORD


@dataclass
class PPTConverter(BatchConverter):
    """ Alias for BatchConverter(PPT, ...) """
    app: enum.EnumMeta = PPT


@dataclass
class XLConverter(BatchConverter):
    """ Alias for BatchConverter(XL, ...) """
    app: enum.EnumMeta = XL
