import os
import sys
import time, y_classes, XlsxHelper_GUI
import wx, threading


class Gui(wx.Frame):
    ver = "0.1 Dev GUI"

    def __init__(self, parent, id):
        wx.Frame.__init__(self, parent, id, title="Excel Work Helper", size=(600, 400))
        self.Center()
        # self.SetBackgroundColour(wx.BLUE)
        # panel=wx.Panel(self)
        menubar = wx.MenuBar()

        self.flag_finished = False
        self.flag_stopped = False

        menu_func = wx.Menu()
        menuitem_check = wx.MenuItem(text="Check", helpString="Check Config File", id=wx.ID_SPELL_CHECK)
        menuitem_repair = wx.MenuItem(text="Repair", helpString="Repair Config File", id=wx.ID_RESET)
        menuitem_load = wx.MenuItem(text="Load Config",helpString="Read Configs From Config File", id=wx.ID_OPEN)
        menuitem_setload = wx.MenuItem(text="Set Config",helpString="Set Config Path", id=wx.ID_SETUP)
        menuitem_setload.SetBitmap(wx.Bitmap("assest/settings.png"))
        menuitem_check.SetBitmap(wx.Bitmap("assest/Check.png"))
        menuitem_load.SetBitmap(wx.Bitmap("assest/Load.png"))
        menuitem_repair.SetBitmap(wx.Bitmap("assest/Repair.png"))
        menu_func.Append(menuitem_check)
        menu_func.Append(menuitem_load)
        menu_func.Append(menuitem_repair)
        menu_func.Append(menuitem_setload)
        menubar.Append(menu_func, "Functions(&F)")

        self.SetMenuBar(menubar)

        self.statusbar = self.CreateStatusBar(3)
        self.statusbar.SetStatusWidths([-1, -1, -1])
        self.statusbar.SetStatusText("Ready",0)
        self.gauge = wx.Gauge(self.statusbar, size=(140,20), pos=(440,4))

        choices = ["auto", "zone"]

        file_sizer = wx.BoxSizer(wx.HORIZONTAL)
        else_sizer = wx.BoxSizer(wx.HORIZONTAL)
        option_sizer = wx.BoxSizer(wx.HORIZONTAL)
        output_sizer = wx.BoxSizer(wx.HORIZONTAL)
        main_sizer = wx.BoxSizer(wx.VERTICAL)

        self.fileinl = wx.StaticText(self, label="FILEin:")
        self.fileinc = wx.TextCtrl(self,value="C:\\")
        self.filein_comfirm = wx.Button(self, label="Confirm")
        self.filein_browse = wx.Button(self, label="Browse")

        self.columninl = wx.StaticText(self, label="Column:")
        self.columninc = wx.TextCtrl(self, size=(50, 20), value="C")
        self.startlinl = wx.StaticText(self, label="StartLine:")
        self.startlinc = wx.TextCtrl(self, size=(70, 20), value="1")
        self.endlinl = wx.StaticText(self, label="EndLine:")
        self.endlinc = wx.TextCtrl(self, size=(70, 20))
        self.igrl = wx.StaticText(self,label="Ignore:")
        self.igrc = wx.TextCtrl(self, size=(70,20))
        self.modeinl = wx.StaticText(self, label="Mode:")
        self.modeinc = wx.Choice(self, choices=choices, style=wx.CB_SORT)
        self.modeinc.SetSelection(0)
        self.endlinc.Enable(False)
        self.igrc.SetValue("5")

        self.replace_mode = wx.CheckBox(self, label="Replace Mode")

        self.output = wx.ListCtrl(self, style=wx.LC_VRULES | wx.LC_HRULES | wx.LC_REPORT | wx.LIST_HITTEST_ABOVE,
                                  size=(580, 300))
        self.output.InsertColumn(0, "Row")
        self.output.InsertColumn(1, "Name")
        self.output.InsertColumn(2, "Event")
        self.output.InsertColumn(3, "Result")
        self.output.InsertColumn(4, "Operation")
        self.output.SetColumnWidth(0,50)
        self.output.SetColumnWidth(1, 80)
        self.output.SetColumnWidth(2, 200)
        self.output.SetColumnWidth(3, 60)
        self.output.SetColumnWidth(4, 80)
        self.output.SetBackgroundColour(wx.BLACK)
        self.output.SetTextColour(wx.WHITE)
        # self.output.Append([1,3,4,5])
        # self.output.SetItemTextColour(0,wx.BLUE)

        file_sizer.Add(self.fileinl, 1, wx.LEFT, 5)
        file_sizer.Add(self.fileinc, 9, wx.LEFT | wx.EXPAND, 5)
        file_sizer.Add(self.filein_comfirm, 2, wx.EXPAND | wx.LEFT, 5)
        file_sizer.Add(self.filein_browse, 2, wx.EXPAND | wx.LEFT, 5)

        else_sizer.Add(self.columninl, 0, wx.LEFT | wx.EXPAND | wx.ALIGN_LEFT, 5)
        else_sizer.Add(self.columninc, 0, wx.LEFT | wx.EXPAND, 5)
        else_sizer.Add(self.startlinl, 0, wx.LEFT | wx.EXPAND, 5)
        else_sizer.Add(self.startlinc, 0, wx.LEFT | wx.EXPAND, 5)
        else_sizer.Add(self.endlinl, 0, wx.LEFT | wx.EXPAND, 5)
        else_sizer.Add(self.endlinc, 0, wx.LEFT | wx.EXPAND, 5)
        else_sizer.Add(self.igrl, 0, wx.LEFT | wx.EXPAND,5)
        else_sizer.Add(self.igrc,0,wx.LEFT|wx.EXPAND,5)
        else_sizer.Add(self.modeinl, 0, wx.LEFT | wx.EXPAND, 5)
        else_sizer.Add(self.modeinc, 0, wx.LEFT | wx.EXPAND, 5)

        option_sizer.Add(self.replace_mode, 0, wx.ALIGN_LEFT | wx.LEFT | wx.EXPAND, 5)

        output_sizer.Add(self.output, 1, wx.ALL | wx.EXPAND | wx.CENTER, 5)

        main_sizer.Add(file_sizer, 0, wx.TOP | wx.EXPAND, 5)
        main_sizer.Add(else_sizer, 0, wx.TOP, 5)
        main_sizer.Add(option_sizer, 0, wx.TOP, 5)
        main_sizer.Add(output_sizer, 0, wx.TOP, 5)

        self.SetSizer(main_sizer)
        self.Fit()

        self.modeinc.Bind(wx.EVT_CHOICE, self.switch_choice)
        self.filein_browse.Bind(wx.EVT_BUTTON, self.file_browse)
        self.filein_comfirm.Bind(wx.EVT_BUTTON, self.OnCLick)
        menubar.Bind(wx.EVT_MENU, self.menu_handler)

    def Reset(self):
        self.output.ClearAll()
        self.gauge.SetValue(0)
        self.statusbar.SetStatusText("TimeCost:",0)

    def switch_choice(self, event):
        selection = self.modeinc.GetSelection()
        if selection == 0:
            self.endlinc.Enable(False)
            self.igrc.Enable(True)
        else:
            self.endlinc.Enable(True)
            self.igrc.Enable(False)

    def file_browse(self, event):
        dlg = wx.FileDialog(self, defaultDir="./", wildcard="Xlsx File(*.xlsx)|*.xlsx|All Files(*.*)|*.*")
        if dlg.ShowModal() == wx.ID_OK:
            self.fileinc.SetValue(dlg.GetPath())
        else:
            pass

    def OnCLick(self, event):

        path = self.fileinc.GetValue()
        if not os.path.exists(path) or not os.path.isfile(path) or os.path.splitext(path)[-1] != ".xlsx":
            wx.MessageDialog(None, "Please Choose Vaild File","Error",wx.ICON_ERROR).ShowModal()
            return 0
        self.filein_comfirm.Enable(False)
        self.filein_browse.Enable(False)

        #st = time.time()

        # t2 = threading.Thread(target=self.Timer(t,st,self.statusbar))


        t = threading.Thread(target=self.Main,args=())
        t.start()
        t2 = threading.Thread(target=self.Timer,args=(t,self.statusbar))
        t2.start()
    def Main(self):
        XlsxHelper_GUI.Load()
        try:
            y = y_classes.y_Contact(self.fileinc.GetValue(), self.columninc.GetValue(), self.startlinc.GetValue(), self.endlinc.GetValue(),
                                    "auto")
            XlsxHelper_GUI.Init(y)
            # XlsxHelper_GUI.Check()
            XlsxHelper_GUI.Open()
            #print("opened")
            print("'",self.endlinc.GetValue(),"'")
            if self.endlinc.GetValue().replace(" ","") == '':
                edl = XlsxHelper_GUI.Range(int(self.igrc.GetValue()))
            else:
                edl = int(self.endlinc.GetValue())
            print(int(edl))
            self.gauge.SetRange(edl-int(self.startlinc.GetValue()))
            # child.start()
            XlsxHelper_GUI.Search(output_main=self.output, output_stream=self.gauge, edl=edl)
            #print("fin")
            self.flag_finished = True
            self.filein_comfirm.Enable(True)
            self.filein_browse.Enable(True)

        except Exception as e:
            self.filein_comfirm.Enable(True)
            self.filein_browse.Enable(True)
            self.flag_stopped = True
            wx.MessageBox(str(e))
            os.system("taskkill -IM EXCEL.EXE -F -T")
            self.Reset()
            # sys.exit()

    def Timer(self,t,output: wx.StatusBar):
        if self.flag_stopped == False:
            st = time.time()
            while t.is_alive():
                ct = time.time()
                output.SetStatusText("TimeCost:%.2f"%(ct-st),0)
            if not self.flag_stopped:
                et = time.time()
                tc = et -st
                output.SetStatusText("TimeCost:%.2f"%(tc),0)
                wx.MessageDialog(None, "Finished", "Finished Time Cost: %.2f"%(tc)).ShowModal()
            else:
                self.Reset()
                sys.exit()
        else:
            self.Reset()
            tc = 0
            sys.exit()



    def menu_handler(self, event):
        if event.Id == wx.ID_SPELL_CHECK:
            XlsxHelper_GUI.Check()
        elif event.Id == wx.ID_RESET:
            XlsxHelper_GUI.Repair()
        elif event.Id == wx.ID_OPEN:
            XlsxHelper_GUI.Load()

    def Entrance(self):
        pass


print(__name__)
if __name__ == "__main__":
    app = wx.App()
    f = Gui(None, -1)
    f.Show()
    # print(app.GetAppName())
    app.MainLoop()
