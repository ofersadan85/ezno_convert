import sys
from itertools import chain
from os import PathLike
from datetime import datetime
from pathlib import Path
from typing import Optional, Collection

VERSION = '0.0.5b1'
DATE_FORMAT = '%Y%m%d-%H%M%S'
here = Path(sys.executable if getattr(sys, 'frozen', False) else __file__)


def validate_paths(src: PathLike, dst: Optional[PathLike] = None, date_fmt: Optional[str] = None) -> tuple[Path, Path]:
    src = Path(src)
    dst = Path(dst) if dst else src.parent
    timestamp = datetime.now().strftime(date_fmt) if date_fmt else ''
    if not src.is_file():
        raise FileNotFoundError(f'Failed to locate specified file {src}')
    if dst.is_dir():
        dst = dst / (src.stem + timestamp)
    elif not dst.parent.is_dir():
        raise NotADirectoryError(f'Failed to find destination directory {dst.parent}')
    return src.absolute(), dst.absolute()


def multi_glob(path: Path, patterns: Collection[str], recursive: bool = False) -> list[Path]:
    generators = [path.glob(f'**/{p}' if recursive else p) for p in patterns]
    return [f for f in chain(*generators)]
