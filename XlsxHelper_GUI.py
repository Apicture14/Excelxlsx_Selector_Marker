import os
import sys
import re
import time

import xlwings, wx,y_classes

ver = "0.1 Dev GUI"


workbook = None
worksheet = None
file = ""
column = ""
nc = "B"
start_line = 0
end_line = 0
mode = ""
target = ""
configf = "./config.txt"



app = xlwings.App(visible=False, add_book=False)
app.display_alerts = False
app.screen_updating = True
# bold = wx.Font(15,family=wx.DEFAULT,style=wx.NORMAL,weight=wx.BOLD)

def Init(yc):
    global column,start_line,end_line,mode,file

    column = yc.column
    start_line = int(yc.startline)
    mode = yc.mode
    file = yc.file
    if mode != "auto":
        end_line = int(yc.endline)

def Open():
    global app, workbook, worksheet
    if os.path.exists(file):
        # print("existed")
        workbook = xlwings.Book(file)
        # print("opeb")
        worksheet = workbook.sheets[0]
    else:
        if wx.MessageDialog(None, "File\"%s\"Not Found","Error",style=wx.ICON_ERROR).ShowModal() == wx.ID_OK:
            sys.exit()

def Range(ignore=5):
    global start_line,column
    row0 = start_line
    void = 0
    while void <= ignore:
        cell_xy = column + str(row0)
        # print(cell_xy)
        cell = worksheet.range(cell_xy)
        value = cell.value
        if not value:
            void += 1
        row0 += 1
        # print(cell_xy, value)
    return row0-1

def Load():
    global target
    if os.path.exists(configf):
        with open(configf,"r",encoding="utf-8") as f:
            target = str(f.readlines()[0])
            # wx.MessageDialog(None,"Loaded:\n%s"%target,"Info",wx.ICON_INFORMATION).ShowModal()
            f.close()
    else:
        wx.MessageDialog(None, "Not Found Config","Error",style=wx.ICON_ERROR).ShowModal()
        raise FileNotFoundError
        # sys.exit()

def Check():
    if os.path.exists(configf):
        with open(configf, "r", encoding="utf-8") as f:
            if str(f.readlines()[1]) != ver:
                wx.MessageDialog(None, "Config Not Fit","Warn",wx.ICON_WARNING).ShowModal()
            else:
                wx.MessageDialog(None, "No Problem Found","Info",wx.ICON_INFORMATION).ShowModal()
            f.close()
    else:
        wx.MessageDialog(None, "Not Found Config", "Error", style=wx.ICON_ERROR).ShowModal()

def Repair(default="优秀|满分|全对|搬书|发书|搬发|表扬|加分|比赛|乐于助人"):
    if wx.MessageDialog(None, "Will delete config.txt existed","Warn",
                        style=wx.ICON_WARNING|wx.YES_NO|wx.NO_DEFAULT).ShowModal() == wx.ID_YES:
        f = open(configf,"w+",encoding="utf-8")
        f.truncate()
        f.write(default+"\n%s"%ver)
        f.close()
    # print("Fin "+str(os.path.exists(configf)))

def Search(output_main: wx.ListCtrl, output_stream: wx.Gauge = None, output_infos: y_classes.y_Return = None,output_res: y_classes.y_Return = None, edl:int = 0):
    bold = wx.Font(8, family=wx.DECORATIVE, style=wx.NORMAL, weight=wx.BOLD)
    try:
        global start_line,end_line,mode,column,nc,target
        name = ""
        result = None
        operation = ""
        content = ""
        checked = 0
        changed = 0
        found = 0
        row = 0
        if mode == "auto":
            end_line = edl
        elif mode == "zone":
            end_line = edl + 1
        for i in range(start_line,end_line):
            checked+=1
            if output_stream is not None:
                output_stream.SetValue(checked)
            if output_infos is not None:
                output_infos.ret = [row,checked,found,changed]
            row = i
            cell_xy = column+str(row)
            cell = worksheet.range(cell_xy)
            namecellxy = nc+str(row)
            result = None
            #print(target,"   ",cell.value)
            #print("f: ",re.findall(target,str(cell.value)))
            if cell.value == None:
                result = None
                content = None
                operation = "Ignored"
            elif re.findall(target,str(cell.value)) != []:
                result = True
                content = cell.value
                found += 1
                if cell.font.color!=(255,0,0):
                    cell.font.color=(255,0,0)
                    changed += 1
                    operation = "Changed"

                else:
                    operation = "Kept"
                    # output.SetItemTextColour(0, wx.YELLOW)
            else:
                operation = "Passed"
                result = False
                content = cell.value
            namecell = worksheet.range(namecellxy)
            name = namecell.value
            element = [row, name, content,result,operation]
            # print(element)
            output_main.Append(element)
            output_main.EnsureVisible(row-1)
            if operation == "Changed":
                output_main.SetItemTextColour(checked-1, wx.GREEN)
            elif operation == "Kept":
                output_main.SetItemTextColour(checked-1, wx.YELLOW)
            elif operation == "Ignored":
                output_main.SetItemFont(checked-1,bold)
        if output_res is not None:
            output_res.ret = [checked,found,changed,time.time()]
        workbook.save()
        workbook.close()
        sys.exit()
    except Exception as e:
        raise e
