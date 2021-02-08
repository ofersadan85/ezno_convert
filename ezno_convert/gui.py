import sys
from pathlib import Path
from typing import Collection

import wx

from ezno_convert.common import DATE_FORMAT, VERSION
from ezno_convert.convert import WORD, PPT, XL, BatchConverter, WORDConverter, PPTConverter, XLConverter

PDF = 'PDF'


class Progress(wx.ProgressDialog):
    def __init__(self, converters: Collection[BatchConverter]):
        super().__init__(
            title=f'Easy Native Office Convert v{VERSION}',
            message='Converting...',
            maximum=max(converters, key=len),
            style=wx.PD_CAN_ABORT | wx.PD_ELAPSED_TIME | wx.PD_REMAINING_TIME | wx.PD_APP_MODAL | wx.PD_AUTO_HIDE
        )
        self.converters = converters

    def run(self):
        self.Show()
        for con_i, converter in enumerate(self.converters):
            app_name = converter.app.app.value.split('.')[0]
            for i, result in enumerate(converter):
                if isinstance(result, Path):
                    self.Update(i, f'Running {app_name} converter... ({con_i}/{len(self.converters)})')
                    if self.WasCancelled():
                        self.Destroy()
                        return


class MainFrame(wx.Frame):
    def __init__(self):
        super().__init__(None, size=(500, 400), title=f'Easy Native Office Convert v{VERSION}')
        word_choices = [item for item in dir(WORD) if not item.startswith('_')]
        ppt_choices = [item for item in dir(PPT) if not item.startswith('_')]
        xl_choices = [item for item in dir(XL) if not item.startswith('_')]
        self.panel = wx.Panel(self)
        self.path = wx.TextCtrl(self.panel)
        self.path_select = wx.Button(self.panel, label='Browse...')
        self.recursive = wx.CheckBox(self.panel, label='Recursively check in sub-folders')
        self.save = wx.TextCtrl(self.panel, value=str(Path.home() / 'Documents'))
        self.save_select = wx.Button(self.panel, label='Save to...')
        self.use_location = wx.CheckBox(self.panel, label='Save output files in the same folder(s) as source files')
        self.word_check = wx.CheckBox(self.panel, label=f'Word documents {WORD.extensions.value}')
        self.ppt_check = wx.CheckBox(self.panel, label=f'Powerpoint Presentations {PPT.extensions.value}')
        self.xl_check = wx.CheckBox(self.panel, label=f'Excel Workbooks {XL.extensions.value}')
        self.word_fmt = wx.ComboBox(self.panel, style=wx.CB_READONLY, value=PDF, choices=word_choices)
        self.ppt_fmt = wx.ComboBox(self.panel, style=wx.CB_READONLY, value=PDF, choices=ppt_choices)
        self.xl_fmt = wx.ComboBox(self.panel, style=wx.CB_READONLY, value=PDF, choices=xl_choices)
        self.execute = wx.Button(self.panel, label='Start Converting...')
        grid = wx.GridBagSizer(10, 10)
        # grid.Add(wx.Bitmap(), pos=(0, 0), span=(1, 2), flag=wx.EXPAND)
        grid.Add(wx.StaticText(self.panel, label='Select files or folders:'), pos=(1, 1))
        grid.Add(self.path, pos=(2, 1), flag=wx.EXPAND)
        grid.Add(self.path_select, pos=(2, 2))
        grid.Add(self.recursive, pos=(3, 1))
        grid.Add(self.use_location, pos=(4, 1))
        grid.Add(wx.StaticText(self.panel, label='Select destination folder:'), pos=(5, 1))
        grid.Add(self.save, pos=(6, 1), flag=wx.EXPAND)
        grid.Add(self.save_select, pos=(6, 2))
        grid.Add(self.word_check, pos=(7, 1))
        grid.Add(self.word_fmt, pos=(7, 2), flag=wx.EXPAND)
        grid.Add(self.ppt_check, pos=(8, 1))
        grid.Add(self.ppt_fmt, pos=(8, 2), flag=wx.EXPAND)
        grid.Add(self.xl_check, pos=(9, 1))
        grid.Add(self.xl_fmt, pos=(9, 2), flag=wx.EXPAND)
        grid.Add(self.execute, pos=(10, 1), span=(1, 2), flag=wx.EXPAND)
        self.panel.SetSizer(grid)
        self.use_location.Bind(wx.EVT_CHECKBOX, self.use_location_handler)
        self.execute.Bind(wx.EVT_BUTTON, self.validate)
        # TODO: Set app icon
        # TODO: Enable selection of locations through file / directory dialogs
        # TODO: Add option to remove timestamp + warning about overwrites
        self.reset()

    def reset(self):
        self.use_location.SetValue(True)
        self.save.Disable()
        self.save_select.Disable()
        self.word_check.SetValue(True)
        self.ppt_check.SetValue(True)
        self.xl_check.SetValue(True)

    def use_location_handler(self, event: wx.Event):
        event.Skip()
        self.save.Enable(not self.use_location.IsChecked())
        self.save_select.Enable(not self.use_location.IsChecked())

    def validate(self, event: wx.Event):
        event.Skip()
        src = Path(self.path.GetValue())
        dst = None if self.use_location.GetValue() else Path(self.save.GetValue())
        kwargs = dict(src=src, dst=dst, recursive=self.recursive.GetValue(), date_fmt=DATE_FORMAT)

        if not src.exists() or (isinstance(dst, Path) and not dst.is_dir()):
            raise ValueError  # FIXME: Show error message to user

        converters = []
        if self.word_check.GetValue():
            converters.append(WORDConverter(target=getattr(WORD, self.word_fmt.GetValue(), None), **kwargs))
        if self.ppt_check.GetValue():
            converters.append(PPTConverter(target=getattr(PPT, self.ppt_fmt.GetValue(), None), **kwargs))
        if self.word_check.GetValue():
            converters.append(XLConverter(target=getattr(XL, self.xl_fmt.GetValue(), None), **kwargs))

        if converters:
            self.Destroy()
            progress = Progress(converters)
            progress.run()
        else:
            raise ValueError  # FIXME: Show error message to user


def main():
    app = wx.App()
    frame = MainFrame()

    try:
        frame.path.SetValue(sys.argv[1])
    except IndexError:
        pass

    frame.Show()
    frame.Center()
    app.MainLoop()


if __name__ == '__main__':
    main()
