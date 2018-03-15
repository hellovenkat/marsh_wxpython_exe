import os
import wx
import wx.lib.agw.multidirdialog as MDD
import wx.lib.scrolledpanel as scrolled
import xlrd
import xlsgrid as XG
import wx.grid as gridlib
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
        global cb1, cb2, cb3, cb4
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
        global phy_sea_level_forecast, phy_sea_level_start, phy_20th, phy_MTA, phy_Marsh_ele, phy_sus_minSed, phy_sus_org, phy_lt
        phy_sea_level_forecast = wx.TextCtrl(leftPan, style=wx.TE_RIGHT, value="80")
        phy_sea_level_start = wx.TextCtrl(leftPan, style=wx.TE_RIGHT, value="-1")
        phy_20th = wx.TextCtrl(leftPan, style=wx.TE_RIGHT, value="0.2")
        phy_MTA = wx.TextCtrl(leftPan, style=wx.TE_RIGHT, value="80")
        phy_Marsh_ele = wx.TextCtrl(leftPan, style=wx.TE_RIGHT, value="80")
        phy_sus_minSed = wx.TextCtrl(leftPan, style=wx.TE_RIGHT, value="80")
        phy_sus_org = wx.TextCtrl(leftPan, style=wx.TE_RIGHT, value="80")
        phy_lt = wx.TextCtrl(leftPan, style=wx.TE_RIGHT, value="80")
        gs.AddMany([(wx.StaticText(leftPan, style=wx.TE_RIGHT, label="Sea Level Forecast")),
                    phy_sea_level_forecast,
                    (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="cm/100y")),
                    (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="Sea Level at Start")),
                    phy_sea_level_start,
                    (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="cm (NAVD)")),
                    (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="20th Cent Sea Level Rate")),
                    phy_20th,
                    (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="cm/yr")),
                    (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="Mean Tidal Amplitude")),
                    phy_MTA,
                    (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="cm/ (MSL)")),
                    (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="Marsh Elevation @ t0")),
                    phy_Marsh_ele,
                    (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="cm/ (MSL)")),
                    (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="Suspended Min. Sed. Conc.")),
                    phy_sus_minSed,
                    (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="mg/l")),
                    (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="Suspended Org. Conc.")),
                    phy_sus_org,
                    (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="mg/l")),
                    (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="LT Accretion Rate")),
                    phy_lt,
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
        global bio_max_growth, bio_min_growth, bio_opt_growth, bio_max_peak, bio_OM_below_root, bio_OM_decay, bio_BGBio, bio_BG_turnover, bio_max_root_depth, bio_reserved
        bio_max_growth = wx.TextCtrl(leftPan, style=wx.TE_RIGHT, value="80")
        bio_min_growth = wx.TextCtrl(leftPan, style=wx.TE_RIGHT, value="-1")
        bio_opt_growth = wx.TextCtrl(leftPan, style=wx.TE_RIGHT, value="0.2")
        bio_max_peak = wx.TextCtrl(leftPan, style=wx.TE_RIGHT, value="80")
        bio_OM_below_root = wx.TextCtrl(leftPan, style=wx.TE_RIGHT, value="80")
        bio_OM_decay = wx.TextCtrl(leftPan, style=wx.TE_RIGHT, value="80")
        bio_BGBio = wx.TextCtrl(leftPan, style=wx.TE_RIGHT, value="80")
        bio_BG_turnover = wx.TextCtrl(leftPan, style=wx.TE_RIGHT, value="80")
        bio_max_root_depth = wx.TextCtrl(leftPan, style=wx.TE_RIGHT, value="80")
        bio_reserved = wx.TextCtrl(leftPan, style=wx.TE_RIGHT, value="80")
        gs1.AddMany([(wx.StaticText(leftPan, style=wx.TE_RIGHT, label="max growth limit (rel MSL)")),
                     #(wx.TextCtrl(leftPan, style=wx.TE_RIGHT, value="80")),
                     (bio_max_growth),
                     (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="cm")),
                     (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="min growth limit (rel MSL)")),
                     (bio_min_growth),
                     (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="cm")),
                     (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="opt growth elev (rel MSL)")),
                     (bio_opt_growth),
                     (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="cm")),
                     (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="max peak biomass")),
                     (bio_max_peak),
                     (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="g/m2")),
                     (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="%OM below root zone")),
                     (bio_OM_below_root),
                     (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="")),
                     (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="OM decay rate")),
                     (bio_OM_decay),
                     (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="1/year")),
                     (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="BGBio to Shoot Ratio")),
                     (bio_BGBio),
                     (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="g/g")),
                     (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="BG turnover rate")),
                     (bio_BG_turnover),
                     (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="1/year")),
                     (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="Max Root Depth")),
                     (bio_max_root_depth),
                     (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="cm")),
                     (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="Reserved")),
                     (bio_reserved),
                     (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="cm"))
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
        global model_max_capture,model_refrac
        model_max_capture = wx.TextCtrl(leftPan, style=wx.TE_RIGHT, value="80")
        model_refrac = wx.TextCtrl(leftPan, style=wx.TE_RIGHT, value="-1")
        gs2 = wx.GridSizer(2, 3, 0, 0)
        gs2.AddMany([(wx.StaticText(leftPan, style=wx.TE_RIGHT, label="Max Capture Eff (q)")),
                     (model_max_capture),
                     (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="tide")),
                     (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="Refrac. Fraction (kr)")),
                     (model_refrac),
                     (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="g/g"))
                     ])

        vbox_leftPan.Add(gs2, proportion=0, flag=wx.EXPAND)
        vbox_leftPan.Add((-1, 25))
        ###############################
        lbl3 = wx.StaticText(leftPan, -1, style=wx.ALIGN_CENTER)
        txt3 = "                   Episodic Storm Inputs or Thin Layer Placement"
        font3 = wx.Font(10, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD)
        lbl3.SetFont(font3)
        lbl3.SetLabel(txt3)
        vbox_leftPan.Add(lbl3)
        gs3 = wx.GridSizer(4, 3, 0, 0)
        gs3.AddMany([(wx.StaticText(leftPan, style=wx.TE_RIGHT, label="years from start")),
                     (wx.TextCtrl(leftPan, style=wx.TE_RIGHT, value="20")),
                     (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="years")),
                     (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="repeat interval")),
                     (wx.TextCtrl(leftPan, style=wx.TE_RIGHT, value="20")),
                     (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="years")),
                     (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="recovery time")),
                     (wx.TextCtrl(leftPan, style=wx.TE_RIGHT, value="10")),
                     (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="years")),
                     (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="add elevation")),
                     (wx.TextCtrl(leftPan, style=wx.TE_RIGHT, value="10")),
                     (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="cm"))
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
        r5 = wx.RadioButton(rbut_rdownPan, label='Coon Isl, SFB')
        r6 = wx.RadioButton(rbut_rdownPan, label='Other Estuary')
        r2.SetValue(True)
        r1.Bind(wx.EVT_RADIOBUTTON, self.onRadioButton)
        # rupPan.SetSizer(vbox_rupPan)
        r2.Bind(wx.EVT_RADIOBUTTON, self.onRadioButton)
        r3.Bind(wx.EVT_RADIOBUTTON, self.onRadioButton)
        r4.Bind(wx.EVT_RADIOBUTTON, self.onRadioButton)
        r5.Bind(wx.EVT_RADIOBUTTON, self.onRadioButton)
        r6.Bind(wx.EVT_RADIOBUTTON, self.onRadioButton)
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
        rbut_vbox.Add((-1, 10))
        rbut_vbox.Add(r6)
        rbut_rdownPan.SetSizer(rbut_vbox)
        text_rdownPan = wx.Panel(self)
        # text_rdownPan.SetBackgroundColour('red')
        #text_rdownPan.Add((-1, 10))
        wx.StaticText(text_rdownPan, style=wx.TE_LEFT, label=" Metrics computed over the final 50 years of simulation")
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
        #print "hello"
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
        #print cb_r.GetLabel()

        if cb_r.GetLabel() == 'North Inlet, SC':
            print 'North'
            #object1 = TabTwo("abcd")
            #sum = object1.rows
            #print texts[81]
            cb1.SetValue(False)
            cb2.SetValue(False)
            cb3.SetValue(False)
            cb4.SetValue(True)

            #k_4=k_2= k_3= k_5= k_6= k_7=k_8=k_9=[]
            #print data_texts
            w, h = 8, 14
            data_list = []
            #del data_list[:]
            myGrid.ClearGrid()
            #print data_list
            data_list = [["" for x in range(w)] for y in range(h)]
            #print data_list
            ind=ind_1=ind_2=ind_3=0
            #print data_texts[62]
            #for i in range(62, 75):
            for i in range(62, 76):
                #MSL - let 1996 be t0
                #data_list

                data_list[ind][0] = data_texts[i][5]
                data_list[ind][1] = data_texts[i][6]
                ind=ind+1
                #k_2.append(float(data_texts[i][5]))
                #k_3.append(float(data_texts[i][6]))
            for i in range(80, 89):
                #MSL - let 1996 be t0
                #print data_texts[i]
                data_list[ind_1][2] = str(float(data_texts[i][0]) - 1996)
                data_list[ind_1][3] = str(float(data_texts[i][1]) * 100)
                ind_1=ind_1+1
                #k_4.append(float(data_texts[i][0]) - 1996)
                #k_5.append(float(data_texts[i][1]) * 100)
                #print k_5 # convert to cm
            for i in range(17,31):
                data_list[ind_2][4] = data_texts[i][9]
                data_list[ind_2][5] = data_texts[i][10]
                ind_2 = ind_2 + 1
                #k_6.append(float(data_texts[i][9]))
                #k_7.append(float(data_texts[i][10]))
            for i in range(18,32):
                data_list[ind_3][6] = str(float(data_texts[i][14]) - 1996)
                data_list[ind_3][7] = str(float(data_texts[i][15]))
                ind_3=ind_3+1
                #k_8.append(float(data_texts[i][14])-1996)
                #k_9.append(float(data_texts[i][15]))
            #print data_list
            myGrid.SetCellValue(0, 0, "Hello")
            for i in range(0,len(data_list)):
                for j in range(0,len(data_list[i])):
                    myGrid.SetCellValue(i, j, data_list[i][j])
            phy_sea_level_forecast.SetLabel("30")
            phy_sea_level_start.SetLabel("-1")
            phy_20th.SetLabel("0.2")
            phy_MTA.SetLabel("70")
            phy_Marsh_ele.SetLabel("43")
            phy_sus_minSed.SetLabel("20")
            phy_sus_org.SetLabel("0")
            phy_lt.SetLabel("")
            bio_max_growth.SetLabel("110")
            bio_min_growth.SetLabel("-25")
            bio_opt_growth.SetLabel("35")
            bio_max_peak.SetLabel("1200")
            bio_OM_below_root.SetLabel("5")
            bio_OM_decay.SetLabel("-0.3")
            bio_BGBio.SetLabel("2")
            bio_BG_turnover.SetLabel("1")
            bio_max_root_depth.SetLabel("25")
            bio_reserved.SetLabel("")
            model_max_capture.SetLabel("1")
            model_refrac.SetLabel("0.1")
        if cb_r.GetLabel() == 'Grand Bay, MS':
            print 'Grand'
            cb1.SetValue(False)
            cb2.SetValue(False)
            cb3.SetValue(False)
            cb4.SetValue(False)
            #object1 = TabTwo("abcd")
            #sum = object1.rows
            #print texts[81]

            w, h = 8, 59
            data_list = []
            #del data_list[:]
            myGrid.ClearGrid()
            data_list = [["" for x in range(w)] for y in range(h)]
            #print data_list
            ind = ind_1 = ind_2 = ind_3 = 0
            #print data_texts[61]
            #print data_texts[61][5]
            #print data_texts[61][6]
            # for i in range(62, 75):
            for i in range(2, 61):
                # MSL - let 1996 be t0
                # data_list

                data_list[ind][0] = str(data_texts[i][5])
                data_list[ind][1] = str(data_texts[i][6])
                ind = ind + 1
                # k_2.append(float(data_texts[i][5]))
                # k_3.append(float(data_texts[i][6]))
            for i in range(41, 79):
                # MSL - let 1996 be t0
                #print data_texts[i]
                data_list[ind_1][2] = str(float(data_texts[i][0]) - 2013)
                data_list[ind_1][3] = str(float(data_texts[i][1]))
                ind_1 = ind_1 + 1
                # k_4.append(float(data_texts[i][0]) - 1996)
                # k_5.append(float(data_texts[i][1]) * 100)
                # print k_5 # convert to cm
            for i in range(2, 14):
                data_list[ind_2][4] = data_texts[i][9]
                data_list[ind_2][5] = data_texts[i][10]
                ind_2 = ind_2 + 1
                # k_6.append(float(data_texts[i][9]))
                # k_7.append(float(data_texts[i][10]))
            for i in range(2, 15):
                if data_texts[i][15] > -30 :
                    data_list[ind_3][6] = data_texts[i][14]
                #data_list[ind_3][6] = str(float(data_texts[i][14]) - 1996)
                data_list[ind_3][7] = data_texts[i][15]
                ind_3 = ind_3 + 1
                # k_8.append(float(data_texts[i][14])-1996)
                # k_9.append(float(data_texts[i][15]))'''
            #print data_list
            for i in range(0, len(data_list)):
                for j in range(0, len(data_list[i])):
                    myGrid.SetCellValue(i, j, data_list[i][j])
            phy_sea_level_forecast.SetLabel("100")
            phy_sea_level_start.SetLabel("9")
            phy_20th.SetLabel("0.25")
            phy_MTA.SetLabel("30")
            phy_Marsh_ele.SetLabel("14")
            phy_sus_minSed.SetLabel("15")
            phy_sus_org.SetLabel("0")
            phy_lt.SetLabel("")
            bio_max_growth.SetLabel("50")
            bio_min_growth.SetLabel("-30")
            bio_opt_growth.SetLabel("25")
            bio_max_peak.SetLabel("2400")
            bio_OM_below_root.SetLabel("8")
            bio_OM_decay.SetLabel("-0.4")
            bio_BGBio.SetLabel("2")
            bio_BG_turnover.SetLabel("0.8")
            bio_max_root_depth.SetLabel("30")
            bio_reserved.SetLabel("")
            model_max_capture.SetLabel("1")
            model_refrac.SetLabel("0.05")
            '''items=[['a','b','c']]
            file_name = "C:\Users\VKOTHA\Downloads\Temp.xls"
            #filename = MainFrame.filePath
            book = xlrd.open_workbook(file_name, formatting_info=1)
            sheetname = "Numerical_Output"
            # sheetname = "Data"
            sheet = book.sheet_by_name(sheetname)
            rows, cols = sheet.nrows, sheet.ncols
            print rows
            print cols
            comments, texts = XG.ReadExcelCOM(file_name, sheetname, rows, cols)
            xlsGrid = XG.XLSGrid(self)
            print book
            print sheet
            print texts
            print comments
            xlsGrid.PopulateGrid(book, sheet, items, comments)
            #print k_2'''

        if cb_r.GetLabel() == 'Plum Island, MA':
            print 'Plum'
            #object1 = TabTwo("abcd")
            #sum = object1.rows
            #print texts[81]
            cb1.SetValue(False)
            cb2.SetValue(False)
            cb3.SetValue(False)
            cb4.SetValue(False)

            w, h = 8, 59
            data_list = []
            #del data_list[:]
            myGrid.ClearGrid()
            data_list = [["" for x in range(w)] for y in range(h)]
            #print data_list
            ind = ind_1 = ind_2 = ind_3 = 0
            #print data_texts[61]
            #print data_texts[61][5]
            #print data_texts[61][6]
            # for i in range(62, 75):
            for i in range(2, 61):
                # MSL - let 1996 be t0
                # data_list

                data_list[ind][0] = str(data_texts[i][5])
                data_list[ind][1] = str(data_texts[i][6])
                ind = ind + 1
                # k_2.append(float(data_texts[i][5]))
                # k_3.append(float(data_texts[i][6]))
            for i in range(41, 79):
                # MSL - let 1996 be t0
                #print data_texts[i]
                data_list[ind_1][2] = str(float(data_texts[i][0]) - 2013)
                data_list[ind_1][3] = str(float(data_texts[i][1]) * 100)
                ind_1 = ind_1 + 1
                # k_4.append(float(data_texts[i][0]) - 1996)
                # k_5.append(float(data_texts[i][1]) * 100)
                # print k_5 # convert to cm
            for i in range(2, 14):
                data_list[ind_2][4] = data_texts[i][9]
                data_list[ind_2][5] = data_texts[i][10]
                ind_2 = ind_2 + 1
                # k_6.append(float(data_texts[i][9]))
                # k_7.append(float(data_texts[i][10]))
            for i in range(2, 15):
                if data_texts[i][15] > -30 :
                    data_list[ind_3][6] = data_texts[i][14]
                #data_list[ind_3][6] = str(float(data_texts[i][14]) - 1996)
                data_list[ind_3][7] = data_texts[i][15]
                ind_3 = ind_3 + 1
                # k_8.append(float(data_texts[i][14])-1996)
                # k_9.append(float(data_texts[i][15]))'''
            #print data_list
            for i in range(0, len(data_list)):
                for j in range(0, len(data_list[i])):
                    myGrid.SetCellValue(i, j, data_list[i][j])
            phy_sea_level_forecast.SetLabel("40")
            phy_sea_level_start.SetLabel("1.8")
            phy_20th.SetLabel("0.2")
            phy_MTA.SetLabel("160")
            phy_Marsh_ele.SetLabel("142.7")
            phy_sus_minSed.SetLabel("15")
            phy_sus_org.SetLabel("1")
            phy_lt.SetLabel("")
            bio_max_growth.SetLabel("195")
            bio_min_growth.SetLabel("0")
            bio_opt_growth.SetLabel("100")
            bio_max_peak.SetLabel("1400")
            bio_OM_below_root.SetLabel("18")
            bio_OM_decay.SetLabel("-0.2")
            bio_BGBio.SetLabel("2")
            bio_BG_turnover.SetLabel("1")
            bio_max_root_depth.SetLabel("25")
            bio_reserved.SetLabel("")
            model_max_capture.SetLabel("1")
            model_refrac.SetLabel("0.05")

        if cb_r.GetLabel() == 'Apalachicola, FL':
            print 'Apcola'
            # object1 = TabTwo("abcd")
            # sum = object1.rows
            # print texts[81]
            cb1.SetValue(False)
            cb2.SetValue(False)
            cb3.SetValue(False)
            cb4.SetValue(False)
            '''  
            k_4=k_2= k_3= k_5= k_6= k_7=k_8=k_9=[]
            for i in range(62, 75):
                #MSL - let 1996 be t0
                k_2.append(float(data_texts[i][5]))
                k_3.append(float(data_texts[i][6]))
            for i in range(80, 88):
                #MSL - let 1996 be t0
                print data_texts[i]
                k_4.append(float(data_texts[i][0]) - 1996)
                k_5.append(float(data_texts[i][1]) * 100)
                #print k_5 # convert to cm
            for i in range(17,30):
                k_6.append(float(data_texts[i][9]))
                k_7.append(float(data_texts[i][10]))
            for i in range(17,30):
                k_8.append(float(data_texts[i][14])-1996)
                k_9.append(float(data_texts[i][15]))
            print k_4
            '''
            phy_sea_level_forecast.SetLabel("100")
            phy_sea_level_start.SetLabel("11")
            phy_20th.SetLabel("0.2")
            phy_MTA.SetLabel("22")
            phy_Marsh_ele.SetLabel("24.2")
            phy_sus_minSed.SetLabel("20")
            phy_sus_org.SetLabel("0")
            phy_lt.SetLabel("")
            bio_max_growth.SetLabel("70")
            bio_min_growth.SetLabel("-10")
            bio_opt_growth.SetLabel("25")
            bio_max_peak.SetLabel("2400")
            bio_OM_below_root.SetLabel("25")
            bio_OM_decay.SetLabel("-0.3")
            bio_BGBio.SetLabel("2")
            bio_BG_turnover.SetLabel("0.8")
            bio_max_root_depth.SetLabel("30")
            bio_reserved.SetLabel("")
            #model_max_capture.SetLabel("1")
            model_refrac.SetLabel("0.05")
        if cb_r.GetLabel() == 'Coon Isl, SFB':
            print 'Coon'
            # object1 = TabTwo("abcd")
            # sum = object1.rows
            # print texts[81]
            cb1.SetValue(False)
            cb2.SetValue(False)
            cb3.SetValue(False)
            cb4.SetValue(False)
            '''  
            k_4=k_2= k_3= k_5= k_6= k_7=k_8=k_9=[]
            for i in range(62, 75):
                #MSL - let 1996 be t0
                k_2.append(float(data_texts[i][5]))
                k_3.append(float(data_texts[i][6]))
            for i in range(80, 88):
                #MSL - let 1996 be t0
                print data_texts[i]
                k_4.append(float(data_texts[i][0]) - 1996)
                k_5.append(float(data_texts[i][1]) * 100)
                #print k_5 # convert to cm
            for i in range(17,30):
                k_6.append(float(data_texts[i][9]))
                k_7.append(float(data_texts[i][10]))
            for i in range(17,30):
                k_8.append(float(data_texts[i][14])-1996)
                k_9.append(float(data_texts[i][15]))
            print k_4
            '''
            phy_sea_level_forecast.SetLabel("100")
            phy_sea_level_start.SetLabel("106")
            phy_20th.SetLabel("0.24")
            phy_MTA.SetLabel("85")
            phy_Marsh_ele.SetLabel("179")
            phy_sus_minSed.SetLabel("100")
            phy_sus_org.SetLabel("0")
            phy_lt.SetLabel("")
            bio_max_growth.SetLabel("89")
            bio_min_growth.SetLabel("-36")
            bio_opt_growth.SetLabel("64")
            bio_max_peak.SetLabel("1200")
            bio_OM_below_root.SetLabel("10")
            bio_OM_decay.SetLabel("-0.3")
            bio_BGBio.SetLabel("4")
            bio_BG_turnover.SetLabel("0.5")
            bio_max_root_depth.SetLabel("20")
            bio_reserved.SetLabel("")
            #model_max_capture.SetLabel("1")
            model_refrac.SetLabel("0.1")
        if cb_r.GetLabel() == 'Other Estuary':
            print 'Other'
            # object1 = TabTwo("abcd")
            # sum = object1.rows
            # print texts[81]
            '''cb1.SetValue(False)
            cb2.SetValue(False)
            cb3.SetValue(False)
            cb4.SetValue(False)'''
            '''  
            k_4=k_2= k_3= k_5= k_6= k_7=k_8=k_9=[]
            for i in range(62, 75):
                #MSL - let 1996 be t0
                k_2.append(float(data_texts[i][5]))
                k_3.append(float(data_texts[i][6]))
            for i in range(80, 88):
                #MSL - let 1996 be t0
                print data_texts[i]
                k_4.append(float(data_texts[i][0]) - 1996)
                k_5.append(float(data_texts[i][1]) * 100)
                #print k_5 # convert to cm
            for i in range(17,30):
                k_6.append(float(data_texts[i][9]))
                k_7.append(float(data_texts[i][10]))
            for i in range(17,30):
                k_8.append(float(data_texts[i][14])-1996)
                k_9.append(float(data_texts[i][15]))
            print k_4
            '''
            phy_sea_level_forecast.SetLabel("100")
            phy_sea_level_start.SetLabel("0")
            phy_20th.SetLabel("0.2")
            phy_MTA.SetLabel("70")
            phy_Marsh_ele.SetLabel("45")
            phy_sus_minSed.SetLabel("20")
            phy_sus_org.SetLabel("0")
            phy_lt.SetLabel("")
            if cb1.GetValue() == True:
                bio_max_growth.SetLabel("120")
                bio_min_growth.SetLabel("-30")
                bio_opt_growth.SetLabel("35")
            bio_max_peak.SetLabel("1200")
            bio_OM_below_root.SetLabel("5")
            bio_OM_decay.SetLabel("-0.4")
            bio_BGBio.SetLabel("4")
            bio_BG_turnover.SetLabel("0.5")
            bio_max_root_depth.SetLabel("30")
            bio_reserved.SetLabel("0.2")
            #model_max_capture.SetLabel("1")
            model_refrac.SetLabel("0.1")
            cb1.SetValue(True)
            cb2.SetValue(False)
            cb3.SetValue(False)
            cb4.SetValue(False)
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
        #panel = wx.Panel(self)
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
        imageFile = "C:\Users\VKOTHA\Downloads\Instructions.jpg"
        #leftPan = wx.lib.scrolledpanel.ScrolledPanel(self)
        #leftPan.SetupScrolling()
        png = wx.Image(imageFile, wx.BITMAP_TYPE_ANY).ConvertToBitmap()
        wx.StaticBitmap(self, -1, png, (10, 5), (png.GetWidth(), png.GetHeight()))



class TabTwo(wx.Panel):
    def __init__(self, parent):
        wx.Panel.__init__(self, parent)
        #t = wx.StaticText(self, -1, "This is the third tab", (20, 20))
        filename = "C:\Users\VKOTHA\Downloads\Temp.xls"
        #filename = MainFrame.filePath
        book = xlrd.open_workbook(filename, formatting_info=1)
        sheetname = "Numerical_Output"
        #sheetname = "Data"
        sheet = book.sheet_by_name(sheetname)
        rows, cols = sheet.nrows, sheet.ncols
        #global comments, texts
        comments, texts= XG.ReadExcelCOM(filename, sheetname, rows, cols)
        #global xlsGrid
        xlsGrid = XG.XLSGrid(self)
        xlsGrid.PopulateGrid(book, sheet, texts, comments)
        #print texts
        sizer = wx.BoxSizer(wx.VERTICAL)
        sizer.Add(xlsGrid, 1, wx.EXPAND, 5)
        self.SetSizer(sizer)


class TabFour(wx.Panel):
    def __init__(self, parent):
        wx.Panel.__init__(self, parent)
        #t = wx.StaticText(self, -1, "This is the third tab", (20, 20))
        filename = "C:\Users\VKOTHA\Downloads\Temp.xls"
        #filename = MainFrame.filePath
        book = xlrd.open_workbook(filename, formatting_info=1)
        sheetname = "Data"
        sheet = book.sheet_by_name(sheetname)
        rows, cols = sheet.nrows, sheet.ncols
        global data_texts
        comments, data_texts = XG.ReadExcelCOM(filename, sheetname, rows, cols)
        xlsGrid = XG.XLSGrid(self)
        xlsGrid.PopulateGrid(book, sheet, data_texts, comments)
        sizer = wx.BoxSizer(wx.VERTICAL)
        sizer.Add(xlsGrid, 1, wx.EXPAND, 5)
        self.SetSizer(sizer)

class TabFive(wx.Panel):
    def __init__(self, parent):
        wx.Panel.__init__(self, parent)
        # t = wx.StaticText(self, -1, "This is the third tab", (20, 20))
        '''filename = "C:\Users\VKOTHA\Downloads\Temp.xls"
        #filename = MainFrame.filePath
        book = xlrd.open_workbook(filename, formatting_info=1)
        sheetname = "IO_data"
        # sheetname = "Data"
        sheet = book.sheet_by_name(sheetname)
        rows, cols = sheet.nrows, sheet.ncols
        #print rows
        #print cols
        comments, texts = XG.ReadExcelCOM(filename, sheetname, rows, cols)


        aaGrid = XG.XLSGrid(self)
        #print book
        #print sheet
        #print texts
        #print comments
        aaGrid.PopulateGrid(book, sheet, texts, comments)
        sizer = wx.BoxSizer(wx.VERTICAL)
        sizer.Add(aaGrid, 1, wx.EXPAND, 5)
        self.SetSizer(sizer)'''
        global myGrid
        myGrid = gridlib.Grid(self)
        myGrid.CreateGrid(100, 10)

        sizer = wx.BoxSizer(wx.VERTICAL)
        sizer.Add(myGrid, 1, wx.EXPAND)

        self.SetSizer(sizer)


class TabSix(wx.Frame):
    """"""

    # ----------------------------------------------------------------------
    def __init__(self):
        """Constructor"""
        wx.Frame.__init__(self, parent=None, title="A Simple Grid")
        panel = wx.Panel(self)

        myGrid = gridlib.Grid(panel)
        myGrid.CreateGrid(12, 8)

        sizer = wx.BoxSizer(wx.VERTICAL)
        sizer.Add(myGrid, 1, wx.EXPAND)
        panel.SetSizer(sizer)

class MainFrame(wx.Frame):
    #filePath = "MEM_file.xls"
    #filePath = "test.xls"
    #filePath = "C://Users/VKOTHA/Downloads/Temp.xls"
    def __init__(self):
        wx.Frame.__init__(self, None, size = wx.DefaultSize, title="MEM v6.0")
        # Create a panel and notebook (tabs holder)
        p = wx.Panel(self)
        nb = wx.Notebook(p)

        # Create the tab windows
        tab1 = TabOne(nb)
        tab2 = TabTwo(nb)
        tab3 = TabThree(nb)
        tab4 = TabFour(nb)
        tab5 = TabFive(nb)
        #tab6 = TabSix(nb)
        #tab7 = TabSeven(nb)
        #tab8 = TabEight(nb)

        # Add the windows to tabs and name them.
        nb.AddPage(tab1, "IO Page")
        nb.AddPage(tab2, "Numerical Output")
        nb.AddPage(tab3, "Instructions")
        nb.AddPage(tab4, "Data")
        nb.AddPage(tab5, "IO_data")
        #nb.AddPage(tab6, "Inundation Time")
        #nb.AddPage(tab7, "Computations")
        #nb.AddPage(tab8, "Sheet12")


        # Set noteboook in a sizer to create the layout
        sizer = wx.BoxSizer()
        sizer.Add(nb, 1, wx.EXPAND)
        p.SetSizer(sizer)



if __name__ == "__main__":
    app = wx.App()
    MainFrame().Show()
    app.MainLoop()
