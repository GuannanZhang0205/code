# -*- coding: utf-8 -*-

###########################################################################
# Python code generated with wxFormBuilder (version Oct 26 2018)
# http://www.wxformbuilder.org/
##
###########################################################################
__author__ = 'GuannanZhang'

import wx
import wx.xrc
import xlrd
import re
import time
from datetime import datetime
from xlrd import xldate_as_tuple

###########################################################################
# Class NewsMatchICBC
###########################################################################

wildcard = 'xls文件（*.xls）|*.xls'


class NewsMatchICBC (wx.Frame):

    def __init__(self, parent):
        wx.Frame.__init__(self, parent, id=wx.ID_ANY, title=u"舆情匹配工具    作者：张冠楠", pos=wx.DefaultPosition, size=wx.Size(
            820, 500), style=wx.DEFAULT_FRAME_STYLE | wx.TAB_TRAVERSAL)

        self.SetSizeHints(wx.DefaultSize, wx.DefaultSize)

        bSizer1 = wx.BoxSizer(wx.VERTICAL)

        bSizer2 = wx.BoxSizer(wx.HORIZONTAL)

        self.filePath1 = wx.TextCtrl(
            self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, 0)
        bSizer2.Add(self.filePath1, 1, wx.ALL, 5)

        self.openF1 = wx.Button(
            self, wx.ID_ANY, u"打开舆情文件", wx.DefaultPosition, wx.DefaultSize, 0)
        bSizer2.Add(self.openF1, 0, wx.ALL, 5)

        self.filePath2 = wx.TextCtrl(
            self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, 0)
        bSizer2.Add(self.filePath2, 1, wx.ALL, 5)

        self.openF2 = wx.Button(
            self, wx.ID_ANY, u"打开关键字文件", wx.DefaultPosition, wx.DefaultSize, 0)
        bSizer2.Add(self.openF2, 0, wx.ALL, 5)

        self.match = wx.Button(self, wx.ID_ANY, u"匹配",
                               wx.DefaultPosition, wx.DefaultSize, 0)
        bSizer2.Add(self.match, 0, wx.ALL, 5)

        bSizer1.Add(bSizer2, 0, wx.EXPAND, 5)

        self.showresult = wx.TextCtrl(
            self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, style=wx.TE_MULTILINE)
        bSizer1.Add(self.showresult, 1, wx.ALL | wx.EXPAND, 5)

        bSizer3 = wx.BoxSizer(wx.HORIZONTAL)

        self.m_staticText = wx.StaticText(
            self, wx.ID_ANY, u"处理进度：", wx.DefaultPosition, wx.DefaultSize, 0)
        self.m_staticText.Wrap(-1)

        bSizer3.Add(self.m_staticText, 0, wx.ALL, 5)

        self.m_gauge = wx.Gauge(self, wx.ID_ANY, 100, wx.DefaultPosition,
                                wx.DefaultSize, wx.GA_HORIZONTAL | wx.GA_TEXT)
        self.m_gauge.SetValue(0)
        bSizer3.Add(self.m_gauge, 0, wx.ALL, 5)

        bSizer1.Add(bSizer3, 0, wx.EXPAND, 5)

        self.infoText = wx.StaticText(
            self, wx.ID_ANY, u"提示：\n1、目前只支持*.xls格式的Excel表格，不支持*.xlsx！可以将xlsx文件另存为xls文件即可！\n2、需要手工将每日舆情信息汇总成一个含有多个sheet的汇总表，并手工构建一个户名关键字表格来进行。\n3、结果可以全选直接复制粘贴到一个新的Excel表格中方便查看。\n4、汇总表中各工作表不要改变格式，会匹配错误！\n5、文件较大时，运行会较慢，请耐心等待。。。\n6、欢迎查看readme", wx.DefaultPosition, wx.DefaultSize, 0)
        self.infoText.Wrap(-1)

        bSizer1.Add(self.infoText, 0, wx.ALL, 5)

        self.SetSizer(bSizer1)
        self.Layout()

        self.Centre(wx.BOTH)

        self.openF1.Bind(wx.EVT_BUTTON, self.yqopen)
        self.openF2.Bind(wx.EVT_BUTTON, self.popen)
        self.match.Bind(wx.EVT_BUTTON, self.yqmatch)

    def __del__(self):
        pass

    def yqopen(self, event):
        dlg = wx.FileDialog(self, u'选择舆情excle文件',
                            wildcard=wildcard, style=wx.DD_DEFAULT_STYLE)
        if dlg.ShowModal() == wx.ID_OK:
            self.filePath1.SetValue(dlg.GetPath())
            #self.wb1 = xlrd.open_workbook(dlg.GetPath())
        dlg.Destroy()

    def popen(self, event):
        dlg = wx.FileDialog(self, u'选择匹配excle文件',
                            wildcard=wildcard, style=wx.DD_DEFAULT_STYLE)
        if dlg.ShowModal() == wx.ID_OK:
            self.filePath2.SetValue(dlg.GetPath())
            #self.wb2 = xlrd.open_workbook(dlg.GetPath())
        dlg.Destroy()

    def yqmatch(self, event):
        try:
            yqbook = xlrd.open_workbook(self.filePath1.GetValue())
            pbook = xlrd.open_workbook(self.filePath2.GetValue())
            temptxt = ''
            yqworksheets = yqbook.sheet_names()
            pworksheets = pbook.sheet_names()
            self.showresult.SetValue(u'')
            self.showresult.write(u'结果如下：\nsheet\t户名\t关键字\t摘要\t发布日期\n')
            for yqworksheet_name in yqworksheets:
                yqworksheet = yqbook.sheet_by_name(yqworksheet_name)
                self.m_gauge.SetValue(int(yqworksheets.index(
                    yqworksheet_name)/len(yqworksheets)*100))
                yqnum_rows = yqworksheet.nrows
                yqnum_cols = yqworksheet.ncols
                for yqrown in range(yqnum_rows):
                    if yqrown > 0:  # 跳过舆情信息的表头
                        for yqcoln in range(yqnum_cols):
                            if yqcoln == 1 or yqcoln == 2 or yqcoln == 3:
                                yqcell = yqworksheet.cell_value(yqrown, yqcoln)
                                yqcellstr = str(yqcell)
                                for pworksheet_name in pworksheets:
                                    pworksheet = pbook.sheet_by_name(
                                        pworksheet_name)
                                    pnum_rows = pworksheet.nrows
                                    pnum_cols = pworksheet.ncols
                                    for prown in range(pnum_rows):
                                        for pcoln in range(pnum_cols):
                                            pcell = pworksheet.cell_value(
                                                prown, pcoln)
                                            pcellstr = str(pcell)
                                            if pcell == '':
                                                pass
                                            elif re.search(pcellstr, yqcellstr):
                                                showtxt = "%s\t%s\t%s\t%s\t%s\n" % (yqworksheet_name, str(pworksheet.cell_value(
                                                    prown, 0)), pcellstr, str(yqworksheet.cell_value(yqrown, 2)), datetime(*xldate_as_tuple(yqworksheet.cell_value(yqrown, 8), 0)))
                                                if self.showresult.GetValue().find(showtxt) == -1:
                                                    self.showresult.write(
                                                        showtxt)
                                                    temptxt = showtxt
                    else:
                        pass
            self.m_gauge.SetValue(100)
            dlg = wx.MessageDialog(
                None, u"已完成，久等了！结果仅供参考！", u"提示", wx.OK | wx.ICON_EXCLAMATION)
            if dlg.ShowModal() == wx.ID_OK:
                pass
            dlg.Destroy()
        except FileNotFoundError:
            dlg = wx.MessageDialog(
                None, u'表格未找到或打开错误，请选择有效的表格！', u'提示', wx.OK | wx.ICON_EXCLAMATION)
            if dlg.ShowModal() == wx.ID_OK:
                pass
            dlg.Destroy()
        except PermissionError:
            dlg = wx.MessageDialog(
                None, u'表格路径错误，请输入正确的表格路径！', u'提示', wx.OK | wx.ICON_EXCLAMATION)
            if dlg.ShowModal() == wx.ID_OK:
                pass
            dlg.Destroy()
        except xlrd.biffh.XLRDError:
            dlg = wx.MessageDialog(
                None, u'请选择有效的表格！', u'提示', wx.OK | wx.ICON_EXCLAMATION)
            if dlg.ShowModal() == wx.ID_OK:
                pass
            dlg.Destroy()
        except:
            dlg = wx.MessageDialog(
                None, u'未知错误！', u'提示', wx.OK | wx.ICON_EXCLAMATION)
            if dlg.ShowModal() == wx.ID_OK:
                pass
            dlg.Destroy()


if __name__ == '__main__':
    app = wx.App()  # 创建一个应用程序对象。每个wxPython程序必须有一个应用程序对象。

    frame = NewsMatchICBC(None)  # 创建一个NewsMatchICBC对象
    frame.Show()  # 调用该对象的 Show()方法以在屏幕上实际显示它
    # 进入主循环。主循环是一个无尽的循环。它捕获并发送应用程序生命周期中存在的所有事件。
    app.MainLoop()
