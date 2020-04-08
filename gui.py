from tkinter import Tk, Text, TOP, BOTH, X, N, LEFT, RAISED, RIGHT, messagebox, END
from tkinter.ttk import Frame, Label, Entry, Button, Checkbutton, Progressbar
import tkinter
import attribute

class Interface(Frame):

    def __init__(self):
        super().__init__()
        
        self.initUI()

    def initUI(self):

        self.master.title("Review")
        self.pack(fill=BOTH, expand=True)

        self.addSpace()

        self.inputFileName = self.addInputBox("Read CSV from file:", "input.csv")
        self.inputMasterEPN = self.addInputBox("Master EPN:", "ABC-123456-789")

        self.optGetFarnellCodes = tkinter.IntVar()
        self.optGetFarnellCodes.set(False)
        self.optGenerateSpecSheets = tkinter.IntVar()
        self.optGenerateSpecSheets.set(False)
        self.optGenerateEagleScript = tkinter.IntVar()
        self.optGenerateEagleScript.set(False)
        self.optGenerateBOM = tkinter.IntVar()
        self.optGenerateBOM.set(False)

        self.addSpace()

        self.checkButton1 = self.addCheckButton("Generate EPN Spec Sheets",self.optGenerateSpecSheets)
        self.checkButton2 = self.addCheckButton("Generate EAGLE Attributes", self.optGenerateEagleScript)
        self.checkButton3 = self.addCheckButton("Generate BOM.xlsx",self.optGenerateBOM)
        self.checkButton4 = self.addCheckButton("Get Farnell Codes", self.optGetFarnellCodes)

        self.addSpace()
        self.addSpace()

        self.text = self.addText()
        self.progressbar = self.addProgressBar()
        
        closeButton = Button(self, text="Cancel", command=quitProgram)
        closeButton.pack(side=RIGHT, padx=5, pady=5)
        okButton = Button(self, text="Start",command=self.executeStart)
        okButton.pack(side=RIGHT)

    def addSpace(v):
        space = Frame(v)
        space.pack(fill=X)
        spaceL = Label(space)
        spaceL.pack()

    def addInputBox(self, text, placeholder):
        frame = Frame(self)
        frame.pack(fill=X)
        label = Label(frame, text=text, width=23)
        label.pack(side=LEFT, padx=20, pady=5)
        inputBox = Entry(frame, validatecommand=self.onConfigChange, validate="focusout")
        inputBox.insert(END,placeholder)
        inputBox.pack(fill=X, padx=20, expand=True)
        return inputBox

    def addCheckButton(self, text, variable):
        frame = Frame(self)
        frame.pack(fill=X)
        label = Label(frame, text=text, width=23)
        label.pack(side=LEFT, padx=20, pady=5)
        button = Checkbutton(frame, variable=variable, onvalue=True, offvalue=False, command=self.onConfigChange)
        button.pack(fill=X, padx=20)
        return button

    def addProgressBar(self):
        p = Progressbar(self, orient="horizontal", length=750, mode="determinate")
        p.pack()
        p["value"]=0
        p["maximum"]=100
        return p

    def addText(self):
        t = Text(self,height=12, width=83, wrap="none", font=("Courier", 9))
        t.configure(state="disabled")
        t.pack()
        return t

    def printToConsole(v,text):
        v.text.configure(state="normal")
        v.text.insert(END, text+"\n")
        v.text.configure(state="disabled")
        v.text.see("end")
        v.text.update()

    def replaceConsole(v,text):
        v.text.configure(state="normal")
        v.text.delete('1.0',END)
        v.text.insert(END, text+"\n")
        v.text.configure(state="disabled")       

    def executeStart(v):
        if v.optGenerateBOM.get() or v.optGenerateSpecSheets.get() or v.optGenerateEagleScript.get():
            res = attribute.execute_main(v.inputFileName.get(), v.inputMasterEPN.get(), v.optGetFarnellCodes.get(), v.optGenerateSpecSheets.get(), v.optGenerateEagleScript.get(), v.optGenerateBOM.get(), v.progressbar, v.printToConsole )
            messagebox.showinfo("Info",res)
            v.progressbar["value"]=0
        else:
            messagebox.showinfo("Error","Please select job - at least one of 1)EAGLE attributes 2)BOM generation 3)EPN Spec Sheets")

    def onConfigChange(v):
        t=""
        if v.optGenerateSpecSheets.get(): t+="EPN Spec sheets will be saved to ./output/ folder.\n"
        if v.optGenerateEagleScript.get(): t+="Eagle script will be saved as 'eagle.txt'.\n"
        if v.optGenerateBOM.get(): t+="A Bill of Materials will be saved as '"+v.inputMasterEPN.get()+".xlsx'\n"
        if v.optGetFarnellCodes.get(): t+="Farnell API will be used to retrieve component codes.\n"

        v.replaceConsole(t)


def quitProgram():
    root.destroy()

def main():

    global root
    root = Tk()
    root.geometry("800x600+300+300")
    app = Interface()
    root.mainloop()

if __name__ == '__main__':
    main()