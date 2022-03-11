# -*- coding: utf-8 -*- 

###########################################################################
## Python code generated with wxFormBuilder (version Jun 17 2015)
## http://www.wxformbuilder.org/
##
## PLEASE DO "NOT" EDIT THIS FILE!
###########################################################################

import wx
import wx.xrc
import AutoUITest
import xlrd
###########################################################################
## Class frmMain
###########################################################################

class frmMain ( wx.Frame ):
	#初始化界面
	def __init__( self, parent, casesList=[] ):
		wx.Frame.__init__ ( self, parent, id = wx.ID_ANY, title = u"插件自动化测试 v1.1", pos = wx.DefaultPosition, size = wx.Size( 600,579 ), style = wx.DEFAULT_FRAME_STYLE|wx.TAB_TRAVERSAL )
		
		self.SetSizeHints( wx.DefaultSize, wx.DefaultSize )
		self.SetBackgroundColour( "White" )

		mainSizer = wx.BoxSizer( wx.VERTICAL )
		
		mainSizer.SetMinSize( wx.Size( 620,400 ) ) 
		self.lblBulletin = wx.StaticText( self, wx.ID_ANY, u"开始前请确认：\n  - 请以管理员身份运行\n  - 如果涉及键盘操作，请关闭输入法", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.lblBulletin.Wrap( -1 )
		self.lblBulletin.SetFont( wx.Font( 9, 70, 90, 92, False, "宋体" ) )
		self.lblBulletin.SetForegroundColour( "Red" )
		
		mainSizer.Add( self.lblBulletin, 0, wx.ALL, 5 )
		
		self.chkSelectAll = wx.CheckBox( self, wx.ID_ANY, u"全选", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.chkSelectAll.SetValue(True) 
		mainSizer.Add( self.chkSelectAll, 0, wx.TOP|wx.LEFT, 5 )
		
		lstTestCasesChoices = casesList
		self.lstTestCases = wx.CheckListBox( self, wx.ID_ANY, wx.DefaultPosition, wx.Size( -1,360 ), lstTestCasesChoices, wx.LB_MULTIPLE|wx.LB_NEEDED_SB|wx.VSCROLL )
		self.lstTestCases.SetCheckedItems(range(0,self.lstTestCases.GetCount()))
		mainSizer.Add( self.lstTestCases, 0, wx.TOP|wx.RIGHT|wx.LEFT|wx.EXPAND, 5 )
		
		self.btnGo = wx.Button( self, wx.ID_ANY, u"开始", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.btnGo.SetMinSize( wx.Size( -1,40 ) )
		
		mainSizer.Add( self.btnGo, 1, wx.ALL|wx.EXPAND, 5 )
		
		
		self.SetSizer( mainSizer )
		self.Layout()
		
		self.Centre( wx.BOTH )
		
		# Connect Events
		self.chkSelectAll.Bind( wx.EVT_CHECKBOX, self.chkSelectAll_OnCheckChanged )
		self.btnGo.Bind( wx.EVT_BUTTON, self.btnGo_OnClick )


	def __del__( self ):
		pass

	def chkSelectAll_OnCheckChanged( self, event ):
		if self.chkSelectAll.IsChecked():
			self.lstTestCases.SetChecked(range(0,self.lstTestCases.GetCount()))
		else:
			self.lstTestCases.SetChecked([])

	def btnGo_OnClick( self, event ):
		#冻结界面最小化
		self.Iconize()
		self.Enable = False

		#运行选中的测试用例
		checkedItems = self.lstTestCases.GetCheckedItems()
		checkedCases = [i+1 for i in checkedItems]
		AutoUITest.RunTests(checkedCases)

		#恢复界面
		self.Enable = True
		self.Restore()

def getCases(file = "test.xlsx"):
	workbook = xlrd.open_workbook(file)
	cases = [sheet.name for sheet in workbook.sheets()[1:]]
	return cases
 
if __name__=='__main__':
    app = wx.App()
    frmMain = frmMain(None,getCases())
    frmMain.Show()
    app.MainLoop()