import sys
from argparse import ArgumentParser, Namespace
from pathlib import Path
from typing import Optional, Sequence, Text

from ezno_convert.common import VERSION, DATE_FORMAT
from ezno_convert.convert import app_batch_convert as batch
from ezno_convert.enums import WORD, PPT, XL


class CommandLineInterface(ArgumentParser):
    def __init__(self):
        super().__init__(description=f'Native Office Converter {VERSION}')
        self.add_argument('PATH', nargs='+', type=Path, help='''
        Path(s) of files or folders to convert.
        If a folder is specified, include all files in it (to filter types, see folder options below)
        ''')
        self.add_argument('-o', '--output', type=Path, metavar='PATH', required=False, help='''
        Where to save output files.
        If input PATH is a folder (or multiple paths) this output PATH must also be a folder.
        Default is to save output files in the same folder as input files.
        If the output PATH isn't a folder (input is a file) and the proper extension is omitted,
        it will be added automatically.
        ''')
        self.add_argument('-c', '--converter', metavar='TYPE', default='PDF', help='''
        Type of conversion to preform. For available types see --list_types (Default: %(default)s)
        ''')
        self.add_argument('-t', '--no_timestamp', action='store_true', help='''
        Do not add timestamp to output filenames. Warning: this will overwrite existing output files.
        ''')
        self.add_argument('-d', '--dateformat', default=DATE_FORMAT, help='''
        Specify datetime format to add to filenames. Ignored if --no-timestamp specified. Default: %(default)s
        ''')
        # self.add_argument('-s', '--simulate', action='store_true',
        #                   help='List files and simulate conversions without actually converting anything')
        self.add_argument('-l', '--list_types', action='store_true', help='Print available conversion types and exit')
        self.add_argument('-v', '--version', action='version', version=f'%(prog)s {VERSION}')

        folder = self.add_argument_group(
            title='Folder Options',
            description='These options apply only to input paths that are folders, they are ignored otherwise'
        )
        folder.add_argument('-r', '--recursive', action='store_true', help='Search in sub-folders recursively')
        folder.add_argument('-w', '--word', action='store_true',
                            help=f'Convert all Word Documents {WORD.extensions.value}')
        folder.add_argument('-p', '--powerpoint', action='store_true',
                            help=f'Convert all PowerPoint Presentations {PPT.extensions.value}')
        folder.add_argument('-x', '--excel', action='store_true',
                            help=f'Convert all Excel Spreadsheets {XL.extensions.value}')
        folder.add_argument('-a', '--all', action='store_true', help=f'''
        Convert all possible files {WORD.extensions.value + PPT.extensions.value + XL.extensions.value}.
        Possible conversions depend on conversion type.
        Example 1: "-a -c PDF" and "-a -c XPS" will apply to all file types
        Example 2: "-a -c AnimatedGIF" will only apply to PowerPoint Presentations even without "-p"
        ''')

        xl = self.add_argument_group(
            title='Excel only options',
            description=f'These options apply only to excel files {XL.extensions.value}, they are ignored otherwise'
        )
        xl.add_argument('--split', action='store_true', help='Convert each worksheet separately')
        xl.add_argument('--sheet', nargs='+', help='''
        Specify names or indexes of specific sheets to convert, instead of converting the entire file. Implies --split
        ''')

    def parse_args(self, args: Optional[Sequence[Text]] = None) -> Namespace:
        args = super().parse_args(args)
        dirs = [p for p in args.PATH if p.is_dir()]

        for p in args.PATH:
            if not p.is_file() and not p.is_dir():
                print(f'Path not valid: {p}', file=sys.stderr)
        args.PATH = [p for p in args.PATH if p.is_file() or p.is_dir()]
        if not args.PATH:
            self.error('No paths were valid')

        if len(args.PATH) == 1 and args.PATH[0].is_file():
            valid_output = args.output is None or args.output.is_dir() or args.output.parent.is_dir()
        elif len(args.PATH) > 1 or dirs:
            valid_output = args.output is None or args.output.is_dir()
        else:
            valid_output = False

        if not valid_output:
            self.error(f'Output path invalid ({args.output})')

        if not any((args.word, args.powerpoint, args.excel)):
            args.all = True

        if args.no_timestamp:
            args.dateformat = None

        word = [p for p in args.PATH if p.suffix in WORD.extensions.value]
        args.word = word + dirs if (args.word or args.all) and hasattr(WORD, args.converter) else []
        excel = [p for p in args.PATH if p.suffix in XL.extensions.value]
        args.excel = excel + dirs if (args.excel or args.all) and hasattr(XL, args.converter) else []
        powerpoint = [p for p in args.PATH if p.suffix in PPT.extensions.value]
        args.powerpoint = powerpoint + dirs if (args.powerpoint or args.all) and hasattr(PPT, args.converter) else []

        if args.sheet:
            args.sheet = [int(sh) if sh.isdigit() else sh for sh in args.sheet]
        elif args.split:
            args.sheet = True

        return args

    def run_converters(self, args: Optional[Sequence[Text]] = None):
        opt = self.parse_args(args)
        kwargs = dict(dst=opt.output, recursive=opt.recursive, date_fmt=opt.dateformat)
        word_gen = batch(app=WORD, src=opt.word, target=getattr(WORD, opt.converter, None), **kwargs)
        pp_gen = batch(app=PPT, src=opt.powerpoint, target=getattr(PPT, opt.converter, None), **kwargs)
        xl_gen = batch(app=XL, src=opt.excel, target=getattr(XL, opt.converter, None), sheets=opt.sheet, **kwargs)

        # FIXME - use wrap execution in progressbar

        if opt.word:
            print('Converting Word Documents:')
            for total, i, result in word_gen:
                print(total, i, result)
        if opt.powerpoint:
            print('Converting PowerPoint Presentations:')
            for total, i, result in pp_gen:
                print(total, i, result)
        if opt.excel:
            print('Converting Excel Spreadsheets:')
            for total, i, result in xl_gen:
                print(total, i, result)


def main():
    CommandLineInterface().run_converters()


if __name__ == '__main__':
    main()
