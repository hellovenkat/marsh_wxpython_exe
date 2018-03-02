import os
import wx
import wx.lib.agw.multidirdialog as MDD
import wx.lib.scrolledpanel as scrolled
import xlrd
import xlsgrid as XG
import matplotlib
matplotlib.use('WXAgg')
from matplotlib.backends.backend_wxagg import FigureCanvasWxAgg as FigureCanvas
from matplotlib.backends.backend_wx import NavigationToolbar2Wx
from matplotlib.figure import Figure
wildcard = "Excel Workbook (*.xls)|*.xls|" \
            "All files (*.*)|*.*"
# Define the tab content as classes:
class TabOne(wx.Panel):

    def __init__(self, parent):
        wx.Panel.__init__(self, parent)
        self.InitUI()

    def InitUI(self):
        #panel = wx.Panel(self)
        #panel.SetBackgroundColour('#4f5049')
        '''t = wx.StaticText(self, -1, "This is the first tab", (20, 20))
        hbox1 = wx.BoxSizer(wx.HORIZONTAL)
        left = wx.Panel(self)
        left.SetBackgroundColour('cyan')
        #hbox1.Add(left, border=8)
        hbox1.Add(left, 3, wx.EXPAND | wx.ALL, 5)
        right = wx.Panel(self)
        right.SetBackgroundColour('red')
        #hbox1.Add(right,  border=8)
        hbox1.Add(right, 3, wx.EXPAND | wx.ALL, 5)
        self.SetSizer(hbox1)'''


        #panel = wx.Panel(self)
        self.SetBackgroundColour('#4f5049')

        vbox = wx.BoxSizer(wx.VERTICAL)
        hbox = wx.BoxSizer(wx.HORIZONTAL)

        #leftPan = wx.Panel(self)
        leftPan = wx.lib.scrolledpanel.ScrolledPanel(self)
        leftPan.SetupScrolling()
        # leftPan.SetBackgroundColour('cyan')
        vbox_leftPan = wx.BoxSizer(wx.VERTICAL)
        cb1 = wx.CheckBox(leftPan, label='Use a generic biom profile')
        cb2 = wx.CheckBox(leftPan, label='Add thin layer')
        cb3 = wx.CheckBox(leftPan, label='Calibrate to accretion rate')
        cb4 = wx.CheckBox(leftPan, label='for future development')
        '''vbox_leftPan.Add(cb1,0,wx.ALIGN_CENTER)
        vbox_leftPan.Add(cb2,0,wx.ALIGN_CENTER)

        vbox_leftPan.Add(cb3,0,wx.ALIGN_CENTER)
        vbox_leftPan.Add(cb4,0,wx.ALIGN_CENTER)'''
        leftPan.SetSizer(vbox_leftPan)
        vbox_leftPan.Add((-1, 3))
        vbox_leftPan.Add(cb1)
        vbox_leftPan.Add(cb2)
        vbox_leftPan.Add(cb3)
        vbox_leftPan.Add(cb4)
        vbox_leftPan.Add((-1, 10))
        runSim = wx.Button(leftPan, 1, 'Run Simulation')
        vbox_leftPan.Add(runSim)
        vbox_leftPan.Add((-1, 25))

        # self.t1 = wx.TextCtrl(panel)
        # vbox_leftPan.Add(self.t1, 1, wx.EXPAND | wx.ALIGN_LEFT | wx.ALL, 5)
        # gridPan = wx.Panel(leftPan)
        # gridPan.SetBackgroundColour('#ffffff')

        #############################
        lbl = wx.StaticText(leftPan, -1, style=wx.ALIGN_CENTER)
        txt = "                   Physical Inputs"
        font = wx.Font(10, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD)
        lbl.SetFont(font)
        lbl.SetLabel(txt)
        vbox_leftPan.Add(lbl)
        gs = wx.GridSizer(8, 3, 0, 0)
        gs.AddMany([(wx.StaticText(leftPan, style=wx.TE_RIGHT, label="Sea Level Forecast")),
                    (wx.TextCtrl(leftPan, style=wx.TE_RIGHT, value="80")),
                    (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="cm/100y")),
                    (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="Sea Level at Start")),
                    (wx.TextCtrl(leftPan, style=wx.TE_RIGHT, value="-1")),
                    (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="cm (NAVD)")),
                    (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="20th Cent Sea Level Rate")),
                    (wx.TextCtrl(leftPan, style=wx.TE_RIGHT, value="0.2")),
                    (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="cm/yr")),
                    (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="Mean Tidal Amplitude")),
                    (wx.TextCtrl(leftPan, style=wx.TE_RIGHT, value="0.2")),
                    (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="cm/yr")),
                    (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="Marsh Elevation @ t0")),
                    (wx.TextCtrl(leftPan, style=wx.TE_RIGHT, value="0.2")),
                    (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="cm/yr")),
                    (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="Suspended Min. Sed. Conc.")),
                    (wx.TextCtrl(leftPan, style=wx.TE_RIGHT, value="0.2")),
                    (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="cm/yr")),
                    (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="Suspended Org. Conc.")),
                    (wx.TextCtrl(leftPan, style=wx.TE_RIGHT, value="0.2")),
                    (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="cm/yr")),
                    (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="LT Accretion Rate")),
                    (wx.TextCtrl(leftPan, style=wx.TE_RIGHT, value="0.2")),
                    (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="cm/yr"))
                    ])

        vbox_leftPan.Add(gs, proportion=0, flag=wx.EXPAND)
        vbox_leftPan.Add((-1, 25))
        ############################
        lbl1 = wx.StaticText(leftPan, -1, style=wx.ALIGN_CENTER)
        txt1 = "                   Biological Inputs"
        font1 = wx.Font(10, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD)
        lbl1.SetFont(font1)
        lbl1.SetLabel(txt1)
        vbox_leftPan.Add(lbl1)
        gs1 = wx.GridSizer(10, 3, 0, 0)
        gs1.AddMany([(wx.StaticText(leftPan, style=wx.TE_RIGHT, label="max growth limit (rel MSL)")),
                     (wx.TextCtrl(leftPan, style=wx.TE_RIGHT, value="80")),
                     (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="cm/100y")),
                     (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="min growth limit (rel MSL)")),
                     (wx.TextCtrl(leftPan, style=wx.TE_RIGHT, value="-1")),
                     (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="cm (NAVD)")),
                     (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="opt growth elev (rel MSL)")),
                     (wx.TextCtrl(leftPan, style=wx.TE_RIGHT, value="0.2")),
                     (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="cm/yr")),
                     (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="max peak biomass")),
                     (wx.TextCtrl(leftPan, style=wx.TE_RIGHT, value="0.2")),
                     (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="cm/yr")),
                     (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="%OM below root zone")),
                     (wx.TextCtrl(leftPan, style=wx.TE_RIGHT, value="0.2")),
                     (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="cm/yr")),
                     (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="OM decay rate")),
                     (wx.TextCtrl(leftPan, style=wx.TE_RIGHT, value="0.2")),
                     (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="cm/yr")),
                     (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="BGBio to Shoot Ratio")),
                     (wx.TextCtrl(leftPan, style=wx.TE_RIGHT, value="0.2")),
                     (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="cm/yr")),
                     (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="BG turnover rate")),
                     (wx.TextCtrl(leftPan, style=wx.TE_RIGHT, value="0.2")),
                     (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="cm/yr")),
                     (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="Max Root Depth")),
                     (wx.TextCtrl(leftPan, style=wx.TE_RIGHT, value="0.2")),
                     (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="cm/yr")),
                     (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="Reserved")),
                     (wx.TextCtrl(leftPan, style=wx.TE_RIGHT, value="0.2")),
                     (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="cm/yr"))
                     ])

        vbox_leftPan.Add(gs1, proportion=0, flag=wx.EXPAND)
        vbox_leftPan.Add((-1, 25))
        ###########################
        lbl2 = wx.StaticText(leftPan, -1, style=wx.ALIGN_CENTER)
        txt2 = "                       Model"
        font2 = wx.Font(10, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD)
        lbl2.SetFont(font2)
        lbl2.SetLabel(txt2)
        vbox_leftPan.Add(lbl2)
        gs2 = wx.GridSizer(2, 3, 0, 0)
        gs2.AddMany([(wx.StaticText(leftPan, style=wx.TE_RIGHT, label="Max Capture Eff (q)")),
                     (wx.TextCtrl(leftPan, style=wx.TE_RIGHT, value="80")),
                     (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="tide")),
                     (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="Refrac. Fraction (kr)")),
                     (wx.TextCtrl(leftPan, style=wx.TE_RIGHT, value="-1")),
                     (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="g/g"))
                     ])

        vbox_leftPan.Add(gs2, proportion=0, flag=wx.EXPAND)
        vbox_leftPan.Add((-1, 25))
        ###############################
        lbl3 = wx.StaticText(leftPan, -1, style=wx.ALIGN_CENTER)
        txt3 = "                   Model Coefficeints"
        font3 = wx.Font(10, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD)
        lbl3.SetFont(font3)
        lbl3.SetLabel(txt3)
        vbox_leftPan.Add(lbl3)
        gs3 = wx.GridSizer(2, 3, 0, 0)
        gs3.AddMany([(wx.StaticText(leftPan, style=wx.TE_RIGHT, label="Max Capture Eff (q)")),
                     (wx.TextCtrl(leftPan, style=wx.TE_RIGHT, value="80")),
                     (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="tide")),
                     (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="Refrac. Fraction (kr)")),
                     (wx.TextCtrl(leftPan, style=wx.TE_RIGHT, value="-1")),
                     (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="g/g"))
                     ])

        vbox_leftPan.Add(gs3, proportion=0, flag=wx.EXPAND)
        vbox_leftPan.Add((-1, 25))
        # panel2 = wx.lib.scrolledpanel.ScrolledPanel(self, -1, size=(screenWidth, 400), pos=(0, 28),
        # style=wx.SIMPLE_BORDER)
        # leftPan.SetupScrolling()
        ###############################
        hbox.Add(leftPan, 1, wx.EXPAND | wx.ALL, 5)

        rupPan = wx.Panel(self)
        rupPan.SetBackgroundColour('#edeeff')
        vbox_rupPan = wx.BoxSizer(wx.VERTICAL)

        hbox_rupPan = wx.BoxSizer(wx.HORIZONTAL)
        # rupPan.SetSizer(hbox_rupPan)
        vbox_rupPan.Add((-1, 10))

        #rup_title = wx.StaticText(rupPan, style=wx.ALIGN_CENTRE, label = "North Inlet, SC")
        #vbox_rupPan.Add(rup_title, wx.LEFT|wx.RIGHT, 30)
        #ruptxt = "North Inlet, SC"+"\n"+"MEM-TLP 6.0"
        #rupfont = wx.Font(16, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD)
        #ruplbl.SetFont(rupfont)
        #ruplbl.SetLabel(ruptxt)
        vbox_rupPan.Add((-1, 30))
        figs = [Figure(figsize=(3, 1.5)) for _ in range(3)]
        axes = [fig.add_subplot(111) for fig in figs]
        canvases = [FigureCanvas(rupPan, -1, fig) for fig in figs]
        for canvas in canvases:
            hbox_rupPan.Add(canvas, 0, wx.LEFT|wx.RIGHT, 30)
            #        fig.set_yscale('log')# for fig in figs
        vbox_rupPan.Add(hbox_rupPan)
        vbox_rupPan.Add((-1, 50))
        hbox_rupPan1 = wx.BoxSizer(wx.HORIZONTAL)
        figs = [Figure(figsize=(3, 1.5)) for _ in range(3)]
        axes = [fig.add_subplot(111) for fig in figs]
        canvases = [FigureCanvas(rupPan, -1, fig) for fig in figs]
        for canvas in canvases:
                hbox_rupPan1.Add(canvas, 0, wx.LEFT|wx.RIGHT|wx.TOP, 30)
        vbox_rupPan.Add(hbox_rupPan1)
        vbox_rupPan.Add((-1, 40))

        rupPan.SetSizer(vbox_rupPan)

        qw = wx.StaticText(rupPan, style=wx.TE_CENTER,
                           label="                                                                             Copyright University of South Carolina 2010. All Rights Reserved, JT Morris 6-9-10")
        vbox_rupPan.Add(qw, 2, wx.EXPAND | wx.ALL, 0)
        # hbox_rdownPan.Add(text_rdownPan, 2, wx.EXPAND | wx.ALL, 0)









        # st1 = wx.StaticText(rupPan, label='North Inlet, SC')
        # st2 = wx.StaticText(rupPan, label='MEM-TLP 6.0')
        vbox.Add(rupPan, 2, wx.EXPAND | wx.ALL, 0)

        # rdownPan = wx.Panel(panel)
        # rdownPan.SetBackgroundColour('#eeeeee')
        hbox_rdownPan = wx.BoxSizer(wx.HORIZONTAL)
        # rupPan.SetSizer(vbox_rupPan)
        rbut_rdownPan = wx.Panel(self)
        # rbut_rdownPan.SetBackgroundColour('cyan')
        hbox_rdownPan.Add(rbut_rdownPan, 1, wx.EXPAND | wx.ALL, 0)

        vbox.Add(hbox_rdownPan, 1, wx.EXPAND | wx.ALL, 0)
        rbut_vbox = wx.BoxSizer(wx.VERTICAL)
        r1 = wx.RadioButton(rbut_rdownPan, label='Plum Island, MA')
        r2 = wx.RadioButton(rbut_rdownPan, label='North Inlet, SC')
        r3 = wx.RadioButton(rbut_rdownPan, label='Apalachicola, FL')
        r4 = wx.RadioButton(rbut_rdownPan, label='Grand Bay, MS')
        r5 = wx.RadioButton(rbut_rdownPan, label='Other Estuary')
        r1.Bind(wx.EVT_RADIOBUTTON, self.onRadioButton)
        # rupPan.SetSizer(vbox_rupPan)
        r2.Bind(wx.EVT_RADIOBUTTON, self.onRadioButton)
        r3.Bind(wx.EVT_RADIOBUTTON, self.onRadioButton)
        r4.Bind(wx.EVT_RADIOBUTTON, self.onRadioButton)
        r5.Bind(wx.EVT_RADIOBUTTON, self.onRadioButton)
        rbut_vbox.Add((-1, 10))
        rbut_vbox.Add(r1)
        rbut_vbox.Add((-1, 10))
        rbut_vbox.Add(r2)
        rbut_vbox.Add((-1, 10))
        rbut_vbox.Add(r3)
        rbut_vbox.Add((-1, 10))
        rbut_vbox.Add(r4)
        rbut_vbox.Add((-1, 10))
        rbut_vbox.Add(r5)
        rbut_rdownPan.SetSizer(rbut_vbox)

        text_rdownPan = wx.Panel(self)
        # text_rdownPan.SetBackgroundColour('red')
        #text_rdownPan.Add((-1, 10))
        wx.StaticText(text_rdownPan, style=wx.TE_LEFT, label="Metrics computed over the final 50 years of simulation")
        hbox_rdownPan.Add(text_rdownPan, 1, wx.EXPAND | wx.ALL, 0)

        browsePanel = wx.Panel(self)
        hbox_rdownPan.Add(browsePanel, 1, wx.EXPAND | wx.ALL, 0)
        browsePanel_vbox = wx.BoxSizer(wx.VERTICAL)
        AnotherFile = wx.StaticText(browsePanel, style=wx.TE_RIGHT, label="Choose another excel file")
        #textBox = wx.TextCtrl(browsePanel, style=wx.TE_CENTER,value="aaaaa")
        #what = textBox.GetValue()
        self.currentDirectory = os.getcwd()
        openFileDlgBtn = wx.Button(browsePanel, label="Browse")
        openFileDlgBtn.Bind(wx.EVT_BUTTON, self.onOpenFile)
        closeBtn = wx.Button(browsePanel, label="Change")
        closeBtn.Bind(wx.EVT_BUTTON, self.onClose)
        #button = wx.Button(browsePanel, id=wx.ID_ANY, label="Change")
        #button.Bind(wx.EVT_BUTTON, self.onButton, what)
        browsePanel_vbox.Add((-1, 10))
        browsePanel_vbox.Add(AnotherFile)
        browsePanel_vbox.Add((-1, 10))
        #browsePanel_vbox.Add(textBox)
        #browsePanel_vbox.Add((-1, 10))
        browsePanel_vbox.Add(openFileDlgBtn)
        browsePanel_vbox.Add((-1, 10))
        browsePanel_vbox.Add(closeBtn)
        browsePanel_vbox.Add((-1, 10))
        browsePanel.SetSizer(browsePanel_vbox)

        hbox.Add(vbox, 3, wx.EXPAND | wx.ALL, 5)
        self.SetSizer(hbox)

        #panel.EnableScrolling(True,True)
    def onOpenFile(self, event):
        """
        Create and show the Open FileDialog
        """
        dlg = wx.FileDialog(
            self, message="Choose a file",
            defaultDir=self.currentDirectory,
            defaultFile="",
            wildcard=wildcard,
            style=wx.FD_OPEN | wx.FD_MULTIPLE | wx.FD_CHANGE_DIR
            )
        if dlg.ShowModal() == wx.ID_OK:
            paths = dlg.GetPaths()
            print "You chose the following file(s):"
            for path in paths:
                print path
                MainFrame.filePath = path
        dlg.Destroy()
        #wx.Window.Destroy()
        #self.Destroy()
        #self.Close()
        #MainFrame().Show(False)

        #self.Update()
    def onClose(self, event):
        """"""
        #self.Close()
        frame = self.GetParent()
        print "hello"
        frame.Destroy()
        #wx.GetApp().Exit()
        #app = wx.App()
        MainFrame().Show()
        #app.MainLoop()
    def onRadioButton(self, e):
        cb_r = e.GetEventObject()
        #self.SetTitle(cb_r.GetLabel())
        # self.rupPan.ruplbl.SetLabel(cb_r.GetLabel())
        #print TabOne.filePath
    '''def onButton(self, event, qqq):
        """
        This method is fired when its corresponding button is pressed
        """
        #what = self.textBox.GetValue()
        print qqq
        print "Button pressed!"'''

class TabThree(wx.Panel):
    def __init__(self, parent):
        wx.Panel.__init__(self, parent)
        #t = wx.StaticText(self, -1, "This is the third tab", (20, 20))
        '''filename = "MEM_file.xls"
        book = xlrd.open_workbook(filename, formatting_info=1)
        sheetname = "Instructions"
        sheet = book.sheet_by_name(sheetname)
        rows, cols = sheet.nrows, sheet.ncols
        comments, texts = XG.ReadExcelCOM(filename, sheetname, rows, cols)

        xlsGrid = XG.XLSGrid(self)
        xlsGrid.PopulateGrid(book, sheet, texts, comments)

        sizer = wx.BoxSizer(wx.VERTICAL)
        sizer.Add(xlsGrid, 1, wx.EXPAND, 5)
        self.SetSizer(sizer)'''
        imageFile = "Instructions.jpg"
        #leftPan = wx.lib.scrolledpanel.ScrolledPanel(self)
        #leftPan.SetupScrolling()
        png = wx.Image(imageFile, wx.BITMAP_TYPE_ANY).ConvertToBitmap()
        wx.StaticBitmap(self, -1, png, (10, 5), (png.GetWidth(), png.GetHeight()))


class TabTwo(wx.Panel):
    def __init__(self, parent):
        wx.Panel.__init__(self, parent)
        #t = wx.StaticText(self, -1, "This is the third tab", (20, 20))
        filename = "C:\Users\VKOTHA\Downloads\Book2.xls"
        #filename = MainFrame.filePath
        book = xlrd.open_workbook(filename, formatting_info=1)
        #sheetname = "Numerical_Output"
        sheetname = "Data"
        sheet = book.sheet_by_name(sheetname)
        rows, cols = sheet.nrows, sheet.ncols
        comments, texts = XG.ReadExcelCOM(filename, sheetname, rows, cols)

        xlsGrid = XG.XLSGrid(self)
        xlsGrid.PopulateGrid(book, sheet, texts, comments)

        sizer = wx.BoxSizer(wx.VERTICAL)
        sizer.Add(xlsGrid, 1, wx.EXPAND, 5)
        self.SetSizer(sizer)


class TabFour(wx.Panel):
    def __init__(self, parent):
        wx.Panel.__init__(self, parent)
        #t = wx.StaticText(self, -1, "This is the third tab", (20, 20))
        #filename = "MEM_file.xls"
        filename = MainFrame.filePath
        book = xlrd.open_workbook(filename, formatting_info=1)
        sheetname = "Computations"
        sheet = book.sheet_by_name(sheetname)
        rows, cols = sheet.nrows, sheet.ncols
        comments, texts = XG.ReadExcelCOM(filename, sheetname, rows, cols)

        xlsGrid = XG.XLSGrid(self)
        xlsGrid.PopulateGrid(book, sheet, texts, comments)

        sizer = wx.BoxSizer(wx.VERTICAL)
        sizer.Add(xlsGrid, 1, wx.EXPAND, 5)
        self.SetSizer(sizer)

class TabFive(wx.Panel):
    def __init__(self, parent):
        wx.Panel.__init__(self, parent)
        #t = wx.StaticText(self, -1, "This is the third tab", (20, 20))
        #filename = "MEM_file.xls"
        filename = MainFrame.filePath
        book = xlrd.open_workbook(filename, formatting_info=1)
        sheetname = "rootdist"
        sheet = book.sheet_by_name(sheetname)
        rows, cols = sheet.nrows, sheet.ncols
        comments, texts = XG.ReadExcelCOM(filename, sheetname, rows, cols)

        xlsGrid = XG.XLSGrid(self)
        xlsGrid.PopulateGrid(book, sheet, texts, comments)

        sizer = wx.BoxSizer(wx.VERTICAL)
        sizer.Add(xlsGrid, 1, wx.EXPAND, 5)
        self.SetSizer(sizer)


class TabSix(wx.Panel):
    def __init__(self, parent):
        wx.Panel.__init__(self, parent)
        #t = wx.StaticText(self, -1, "This is the third tab", (20, 20))
        '''filename = "MEM_file.xls"
        book = xlrd.open_workbook(filename, formatting_info=1)
        sheetname = "Inundation Time"
        sheet = book.sheet_by_name(sheetname)
        rows, cols = sheet.nrows, sheet.ncols
        comments, texts = XG.ReadExcelCOM(filename, sheetname, rows, cols)

        xlsGrid = XG.XLSGrid(self)
        xlsGrid.PopulateGrid(book, sheet, texts, comments)

        sizer = wx.BoxSizer(wx.VERTICAL)
        sizer.Add(xlsGrid, 1, wx.EXPAND, 5)
        self.SetSizer(sizer)'''
        imageFile1 = "Inundation Time.PNG"
        #leftPan = wx.lib.scrolledpanel.ScrolledPanel(self)
        #leftPan.SetupScrolling()
        png1 = wx.Image(imageFile1, wx.BITMAP_TYPE_ANY).ConvertToBitmap()
        wx.StaticBitmap(self, -1, png1, (10, 5), (png1.GetWidth(), png1.GetHeight()))
class TabSeven(wx.Panel):
    def __init__(self, parent):
        wx.Panel.__init__(self, parent)
        #t = wx.StaticText(self, -1, "This is the third tab", (20, 20))
        #filename = "MEM_file.xls"
        filename = MainFrame.filePath
        book = xlrd.open_workbook(filename, formatting_info=1)
        sheetname = "Data"
        sheet = book.sheet_by_name(sheetname)
        rows, cols = sheet.nrows, sheet.ncols
        comments, texts = XG.ReadExcelCOM(filename, sheetname, rows, cols)

        xlsGrid = XG.XLSGrid(self)
        xlsGrid.PopulateGrid(book, sheet, texts, comments)

        sizer = wx.BoxSizer(wx.VERTICAL)
        sizer.Add(xlsGrid, 1, wx.EXPAND, 5)
        self.SetSizer(sizer)
class TabEight(wx.Panel):
    def __init__(self, parent):
        wx.Panel.__init__(self, parent)
        #t = wx.StaticText(self, -1, "This is the third tab", (20, 20))
        #filename = "MEM_file.xls"
        filename = MainFrame.filePath
        book = xlrd.open_workbook(filename, formatting_info=1)
        sheetname = "Sheet12"
        sheet = book.sheet_by_name(sheetname)
        rows, cols = sheet.nrows, sheet.ncols
        comments, texts = XG.ReadExcelCOM(filename, sheetname, rows, cols)

        xlsGrid = XG.XLSGrid(self)
        xlsGrid.PopulateGrid(book, sheet, texts, comments)

        sizer = wx.BoxSizer(wx.VERTICAL)
        sizer.Add(xlsGrid, 1, wx.EXPAND, 5)
        self.SetSizer(sizer)
class MainFrame(wx.Frame):
    filePath = "MEM_file.xls"
    def __init__(self):
        wx.Frame.__init__(self, None, size = wx.DefaultSize, title="MEM v6.0")
        # Create a panel and notebook (tabs holder)
        p = wx.Panel(self)
        nb = wx.Notebook(p)

        # Create the tab windows
        tab1 = TabOne(nb)
        tab2 = TabTwo(nb)
        '''tab3 = TabThree(nb)
        tab4 = TabFour(nb)
        tab5 = TabFive(nb)
        tab6 = TabSix(nb)
        tab7 = TabSeven(nb)
        tab8 = TabEight(nb)'''

        # Add the windows to tabs and name them.
        nb.AddPage(tab1, "IO Page")
        nb.AddPage(tab2, "Numerical Output")
        '''nb.AddPage(tab3, "Instructions")
        nb.AddPage(tab4, "Computations")
        nb.AddPage(tab5, "rootdist")
        nb.AddPage(tab6, "Inundation Time")
        nb.AddPage(tab7, "Data")
        nb.AddPage(tab8, "Sheet12")'''


        # Set noteboook in a sizer to create the layout
        sizer = wx.BoxSizer()
        sizer.Add(nb, 1, wx.EXPAND)
        p.SetSizer(sizer)



if __name__ == "__main__":
    app = wx.App()
    MainFrame().Show()
    app.MainLoop()
