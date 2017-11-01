import Tkinter as tk
import tkFont
import ttk as ttk
import winsound
import subprocess
import os
from _winreg import *

class SoundMinor:
    def __init__(self, root):
        self.root = root
        self.includeTags = []
        self.requiredTags = []
        self.excludeTags = []
        self.currentSound = ""
        self.autoPlay = tk.BooleanVar()
        self.sfLocation = "C:\\Program Files (x86)\\Sony\\Sound Forge Pro 11.0\\Forge110.exe"
        self.CreateGUI()

    def CreateGUI(self):
        self.searchBox = ttk.Entry(self.root, width=20)
        self.searchBox.grid(column=0, row=0, padx=10, pady=10, sticky='nsew')
        self.searchBox.bind('<Return>', self.GetSearchParameters)
        self.searchInfoText = "Separate search tags with a space.\n"
        self.searchInfoText += "Use a \"+\" at the beginning of a search term to enter a required tag.\n"
        self.searchInfoText += "Use a \"-\" at the beginning of a search term to enter an exclude tag."
        self.searchInfo = ttk.Label(self.root, text=self.searchInfoText)
        self.searchInfo.grid(column=0, row=1, sticky='w', padx=10)
        self.searchButton = ttk.Button(self.root, text="Search")
        self.searchButton.grid(column=1, row=0, columnspan=2)
        self.searchButton.bind("<Button-1>", self.GetSearchParameters)
        self.autoPlayCheck = ttk.Checkbutton(self.root, text="AutoPlay", variable=self.autoPlay)
        self.autoPlayCheck.grid(column=3, row=0)
        self.CreateSoundList()
        self.sfButton = ttk.Button(text="Open in Sound Forge")
        self.sfButton.grid(column=0, row=4, sticky='w', padx=10)
        self.sfButton.bind('<Button-1>', self.OpenInSF)
        self.findSFButton = ttk.Button(text="Find Sound Forge")
        self.findSFButton.grid(column=1, row=4)
        self.findSFButton.bind('<Button-1>', self.FindSF)
        self.listInfoText = "Spacebar to play a sound (and stop if selected sound is already playing).\n"
        self.listInfoText += "Control + Space to stop playing sound (regardless of selection).\n"
        self.listInfoText += "Double click to open file location in explorer."
        self.listInfo = ttk.Label(self.root, text=self.listInfoText)
        self.listInfo.grid(column=0, row=5, sticky='w', padx=10)
        #self.themeChoice = tk.StringVar()
        #self.themeOptions = ['light', 'dark']
        #self.themeOptionMenu = ttk.OptionMenu(self.root, self.themeChoice, self.themeOptions[0], *self.themeOptions)
        #self.themeOptionMenu.grid(column=1, row=5)
        #self.themeChoice.trace(mode="w", callback=self.SetTheme)

    def GetSearchParameters(self, event):
        self.includeTags = []
        self.requiredTags = []
        self.excludeTags = []
        rawTags = self.searchBox.get().lower()
        print rawTags
        if (rawTags ==""):
            return 
        for tag in rawTags.split(" "):
            if ( tag[0] == "-"):
                self.excludeTags.append(tag[1:])
            elif ( tag[0] == "+" ):
                self.requiredTags.append(tag[1:])
            else:
                self.includeTags.append(tag)
        self.BuildFileList()

    def CreateSoundList(self):
        headers = ['Path', 'Filename', 'Length', 'Type', 'Tags']
        self.fileList = ttk.Treeview(columns=headers, show="headings")
        self.fileList.bind("<space>", self.PlaySound)
        self.root.bind_all("<Control-space>", self.StopSound)
        self.fileList.bind("<Double-1>", self.OpenPath)
        self.fileList.bind("<Button-1>", self.AutoPlaySound)
        vsb = ttk.Scrollbar(orient="vertical", command=self.fileList.yview)
        hsb = ttk.Scrollbar(orient="horizontal", command=self.fileList.xview)
        self.fileList.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.fileList.grid(column=0, row=2, padx=10, pady=10, sticky='nsew', in_=self.root, columnspan=4)
        self.root.rowconfigure(2, minsize=500)
        self.fileList.column("Path", width=80, minwidth=80)
        self.fileList.column("Filename", width=500, minwidth=500)
        self.fileList.column("Length", width=80, minwidth=80)
        self.fileList.column("Type", width=80, minwidth=80)
        self.fileList.column("Tags", width=80, minwidth=80)
        vsb.grid(column=4, row=2, sticky='nsw')
        hsb.grid(column=0, row=3, sticky='new', columnspan=4)
        for col in headers:
            self.fileList.heading(col, text=col.title(), command=lambda c=col: self.SortBy(self.fileList, c, 0))

    def BuildFileList(self):
        dataFile = open("soundlibrary.txt")
        dataLines = dataFile.readlines()
        self.dataList = []
        skip = False
        for line in dataLines:
            tempLine = line.replace("\n","").lower()
            tempLine = tempLine.split("|")
            tempName = tempLine[0].split("\\")[-1]
            tempPath = tempLine[0].replace(tempName,"")
            if ( len(self.requiredTags) > 0 ):
                for rTag in self.requiredTags:
                    if not ( rTag in tempLine[3] or rTag in tempLine[0] ):
                        skip = True
            if not (skip):
                if ( len(self.excludeTags) > 0 ):
                    for eTag in self.excludeTags:
                        if ( eTag in tempLine[3] or eTag in tempLine[0] ):
                            skip = True
            if not (skip):
                if ( len(self.includeTags) > 0 ):
                    for iTag in self.includeTags:
                        if ( iTag in tempLine[3] or iTag in tempLine[0] ):
                            self.dataList.append((tempPath, tempName, float(tempLine[1]), tempLine[2], tempLine[3]))
                            break
                else:
                    self.dataList.append((tempPath, tempName, float(tempLine[1]), tempLine[2], tempLine[3]))
            skip = False
        for item in self.fileList.get_children():
            self.fileList.delete(item)
        rowCount = 1
        for item in self.dataList:
            if (rowCount % 2 == 0):
                self.fileList.insert('', 'end', values=item, tags=('evenrow',))
            else:
                self.fileList.insert('', 'end', values=item, tags=('oddrow',))
            rowCount += 1
        #self.fileList.tag_configure('oddrow', background='lightgrey')
        #self.fileList.tag_configure('evenrow', background='darkgrey')

    def MakeNumeric(self, data):
        new_data = []
        for child, col in data:
            new_data.append((float(child), col))
        return new_data
        return data

    def SortBy(self, tree, col, descending):
        # grab values to sort
        data = [(tree.set(child, col), child) for child in tree.get_children('')]
        # if the data to be sorted is numeric change to float
        if (col == "Length"):
            data = self.MakeNumeric(data)
        data.sort(reverse=descending)
        for ix, item in enumerate(data):
            tree.move(item[1], '', ix)
        # switch the heading so it will sort in the opposite direction
        tree.heading(col, command=lambda col=col: self.SortBy(tree, col, int(not descending)))

    def AutoPlaySound(self, event):
        if (not self.autoPlay.get()):
            return
        item = self.fileList.identify('item', event.x, event.y)
        if ( item != "" ):
            soundName = self.fileList.item(item)["values"][0] + self.fileList.item(item)["values"][1]
            if (self.currentSound == soundName):
                self.StopSound(None)
            else:
                winsound.PlaySound(soundName, winsound.SND_ASYNC)
                self.currentSound = soundName

    def PlaySound(self, event):
        item = self.fileList.selection()[0]
        if ( item != "" ):
            soundName = self.fileList.item(item)["values"][0] + self.fileList.item(item)["values"][1]
            if (self.currentSound == soundName):
                self.StopSound(None)
            else:
                winsound.PlaySound(soundName, winsound.SND_ASYNC)
                self.currentSound = soundName

    def StopSound(self, event):
        winsound.PlaySound(None, winsound.SND_ASYNC)
        self.currentSound = ""

    def OpenPath(self, event):
        item = self.fileList.identify('item', event.x, event.y)
        soundName = self.fileList.item(item)["values"][0] + self.fileList.item(item)["values"][1]
        if ( item != "" ):
            subprocess.Popen(r'explorer /select, ' + soundName)

    def FindSF(self, event):
        """
        print self.sfLocation
        if (os.path.exists(self.sfLocation)):
            print "DONE"
        else:
            print "NOPE"
        """
        aReg = ConnectRegistry(None, HKEY_CURRENT_USER)
        aKey = OpenKey(aReg, r"Software\\Classes\\VirtualStore\\MACHINE\\SOFTWARE\\Wow6432Node\\Cakewalk Music Software\\Tools Menu\\Sound Forge Pro")
        val=QueryValueEx(aKey, "ExePath")
        self.sfLocation = val[0]
        print self.sfLocation

    def OpenInSF(self, event):
        try:
            item = self.fileList.selection()[0]
            if ( item != "" ):
                soundName = self.fileList.item(item)["values"][0] + self.fileList.item(item)["values"][1]
                cmdLine = "\"" + self.sfLocation + "\""
                cmdLine += " /Open:\"" + soundName + "\""
                subprocess.Popen(cmdLine)
        except:
            print "Can't open in SF for reasons."

    def SetTheme(self, varname, elementname, mode):
        #ttk.Style().configure("Treeview", background="#242424", foreground="white")
        #ttk.Style().configure("TButton", background=[('pressed', '!disabled', 'black'), ('active', 'white')], foreground=[('pressed', 'red'), ('active', 'blue')])
        tempColor = ""
        if (self.themeChoice.get() == "dark"):
            tempColor = "#242424"
        elif (self.themeChoice.get() == "light"):
            tempColor = "lightgrey"
        self.root.configure(background=tempColor)
        self.searchBox.configure(background=tempColor)
        self.searchInfo.configure(background=tempColor)
        self.listInfo.configure(background=tempColor)
        #self.searchButton.configure(background=tempColor)
        #self.autoPlayCheck.configure(background=tempColor)
        #self.sfButton.configure(background=tempColor)
        #self.findSFButton.configure(background=tempColor)
        #self.openSFButton.configure(background=tempColor)
        #self.themeOptionMenu.configure(background=tempColor)

root = tk.Tk()
root.wm_title("SoundMinor")
root.geometry("1040x700")
app = SoundMinor(root)
root.mainloop()