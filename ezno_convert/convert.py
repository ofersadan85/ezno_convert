import enum
import logging
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


def app_batch_convert(
        app: enum.EnumMeta,
        src: Union[Collection[PathLike], PathLike],
        dst: Optional[PathLike] = None,
        target: Optional[enum_types] = None,
        recursive: bool = False,
        date_fmt: Optional[str] = None,
        sheets: Union[Collection, bool] = False):

    if isinstance(src, (str, PathLike)):
        src = [src]

    files = [Path(f) for f in src if Path(f).is_file()]
    extensions = [x if '*' in x else '*' + x for x in app.extensions.value]
    for d in (Path(f) for f in src if Path(f).is_dir()):
        files += multi_glob(d, extensions, recursive=recursive)

    files = {f for f in files if not f.name.startswith('~$')}

    if dst is not None and len(files) > 1:
        dst = Path(dst)
        if not dst.is_dir():
            raise NotADirectoryError(f'Destination for batch conversion must be a directory (or empty) ({dst})')

    if files:
        total = len(files)
        app_object = CreateObject(app.app.value)
        for i, f in enumerate(files):
            try:
                result = convert_one(f, dst, app_object, target, date_fmt, sheets)
            except (COMError, FileNotFoundError, NotADirectoryError, ValueError):
                logger.exception(f'Failed to convert: {f}')
                yield total, i + 1, None
            else:
                yield total, i + 1, result
        app_object.Quit()
