import logging
import sys
from pathlib import Path
from typing import Collection, Optional

import wx

from ezno_convert.common import DATE_FORMAT, VERSION
from ezno_convert.convert import WORD, PPT, XL, BatchConverter, WORDConverter, PPTConverter, XLConverter

PDF = 'PDF'
logger = logging.getLogger('NativeOfficeConverter')


class WarningDialog(wx.MessageDialog):
    def __init__(self, parent, message):
        warning_style = wx.YES_NO | wx.CENTER | wx.ICON_WARNING
        super().__init__(parent=parent, message=message, caption='Warning!', style=warning_style)
        logger.warning(message)


class ErrorDialog(wx.MessageDialog):
    def __init__(self, parent, message):
        super().__init__(parent=parent, message=message, caption='Error!', style=wx.OK | wx.CENTER | wx.ICON_WARNING)
        logger.error(message)
        self.ShowModal()


class Progress(wx.ProgressDialog):
    def __init__(self, converters: Collection[BatchConverter]):
        super().__init__(
            title=f'Easy Native Office Convert v{VERSION}',
            message='Converting...',
            maximum=len(max(converters, key=len)),
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
        style = wx.MINIMIZE_BOX | wx.SYSTEM_MENU | wx.CAPTION | wx.CLOSE_BOX | wx.CLIP_CHILDREN
        super().__init__(None, size=(550, 500), title=f'Easy Native Office Convert v{VERSION}', style=style)
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
        self.date_fmt = wx.TextCtrl(self.panel, value=DATE_FORMAT)
        self.reset_btn = wx.Button(self.panel, label='Reset settings')
        self.execute_btn = wx.Button(self.panel, label='Start Converting...')
        img = wx.Image('images/ezno-banner-wide.png', type=wx.BITMAP_TYPE_PNG).Rescale(500, 100).ConvertToBitmap()
        img = wx.StaticBitmap(self.panel, bitmap=img)
        grid = wx.GridBagSizer(10, 10)
        grid.Add(img, pos=(0, 1), span=(1, 2), flag=wx.EXPAND)
        grid.Add(wx.StaticText(self.panel, label='Select a folder:'), pos=(1, 1))
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
        grid.Add(wx.StaticText(self.panel, label='Date/Time format on filenames:'), pos=(10, 1))
        grid.Add(self.date_fmt, pos=(10, 2), flag=wx.EXPAND)
        grid.Add(self.execute_btn, pos=(11, 1), flag=wx.EXPAND)
        grid.Add(self.reset_btn, pos=(11, 2), flag=wx.EXPAND)
        self.panel.SetSizer(grid)
        self.use_location.Bind(wx.EVT_CHECKBOX, self.use_location_handler)
        self.path_select.Bind(wx.EVT_BUTTON, self.select_path_location)
        self.save_select.Bind(wx.EVT_BUTTON, self.select_save_location)
        self.reset_btn.Bind(wx.EVT_BUTTON, self.reset)
        self.execute_btn.Bind(wx.EVT_BUTTON, self.validate)
        self.SetIcon(wx.Icon('images/ezno-icon.png'))
        self.reset()

    def reset(self, event: Optional[wx.Event] = None):
        if event:
            event.Skip()
        self.use_location.SetValue(True)
        self.save.Disable()
        self.save_select.Disable()
        self.word_check.SetValue(True)
        self.ppt_check.SetValue(True)
        self.xl_check.SetValue(True)
        self.date_fmt.SetValue(DATE_FORMAT)

    def select_path_location(self, event: wx.Event):
        event.Skip()
        self.select_folder(self.path, 'Select a source folder')

    def select_save_location(self, event: wx.Event):
        event.Skip()
        self.select_folder(self.save, 'Select a destination folder')

    def select_folder(self, text_ctrl: wx.TextCtrl, message: str):
        start = text_ctrl.GetValue()
        start = start if start and Path(start).is_dir() else str(Path.home())
        dir_dlg = wx.DirDialog(self, defaultPath=start, message=message)
        if dir_dlg.ShowModal() == wx.ID_OK:
            text_ctrl.SetValue(dir_dlg.GetPath())

    def use_location_handler(self, event: wx.Event):
        event.Skip()
        self.save.Enable(not self.use_location.IsChecked())
        self.save_select.Enable(not self.use_location.IsChecked())

    def validate(self, event: wx.Event):
        event.Skip()

        date_fmt = self.date_fmt.GetValue()
        if DATE_FORMAT != date_fmt:
            warning = 'Changing the timestamp format field could result in unintended overwrites of files!\n'
            warning += 'Are you sure you want to continue?'
            result = WarningDialog(self, warning).ShowModal()
            if result != wx.ID_YES:
                return

        src = self.path.GetValue()
        if src:
            src = Path(src)
        else:
            ErrorDialog(self, 'Source path cannot be empty')
            return
        if not src.is_dir() and not src.is_file():
            ErrorDialog(self, 'Source path provided is invalid')
            return

        dst = None if self.use_location.GetValue() else Path(self.save.GetValue())
        if isinstance(dst, Path) and not dst.is_dir():
            ErrorDialog(self, 'Destination path must be a valid directory')
            return

        kwargs = dict(src=src, dst=dst, recursive=self.recursive.GetValue(), date_fmt=date_fmt)

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
            ErrorDialog(self, 'Could not find any files to convert. Check your settings.')
            return


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
