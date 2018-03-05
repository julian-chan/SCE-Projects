import tkinter
import time
from compress import zipFiles, progressBar
from tkinter import Radiobutton, Menu, Button, Label, PhotoImage, Toplevel, filedialog, StringVar, messagebox

# Button Functions -----------------------------------------
def chooseSourceDirectory():
    sourcePath = filedialog.askdirectory(initialdir="C:")
    if sourcePath != "":
        sourceEntry['text'] = sourcePath
        parameters["source"] = sourcePath
        sourceButton['bg'] = "green"
        checkReadyToCompress()

def chooseDataType():
    def storeDataType():
        current_selection = var.get()
        if current_selection != "None":
            label.config(text="Currently selected: " + current_selection)

            dataTypeEntry['text'] = current_selection
            parameters["Data Type"] = current_selection
            dataTypeButton['bg'] = "green"
            messagebox.showinfo(title="Selection Successful", message="Data Type saved.")
            checkReadyToCompress()
            top.destroy()
        else:
            dataTypeEntry['text'] = ""
            parameters["Data Type"] = ""
            dataTypeButton['bg'] = "red"
            checkReadyToCompress()
            messagebox.showerror(title="Invalid Selection", message="Please select a valid value. Valid values are:\nStreamData, 90minute, 9hour, 24hour, 90hour, 120hour, Monolith")

    top = Toplevel()
    top.iconbitmap("sce_icon.ico")
    top.resizable(0,0)
    top.geometry("400x200")
    top.title("Choose type/resolution of data:")
    var = StringVar()
    var.set("None")
    option1 = Radiobutton(top, text="StreamData (SCEP_...)", variable=var, value="StreamData", command=storeDataType)
    option1.pack(anchor="w")
    option2 = Radiobutton(top, text="90minute (SCEC_...)", variable=var, value="90minute", command=storeDataType)
    option2.pack(anchor="w")
    option3 = Radiobutton(top, text="9hour (SCED_...)", variable=var, value="9hour", command=storeDataType)
    option3.pack(anchor="w")
    option4 = Radiobutton(top, text="24hour (SCEK_...)", variable=var, value="24hour", command=storeDataType)
    option4.pack(anchor="w")
    option5 = Radiobutton(top, text="90hour (SCEM_...)", variable=var, value="90hour", command=storeDataType)
    option5.pack(anchor="w")
    option6 = Radiobutton(top, text="120hour (SCEW_...)", variable=var, value="120hour", command=storeDataType)
    option6.pack(anchor="w")
    option7 = Radiobutton(top, text="Monolith (SCET_...)", variable=var, value="Monolith", command=storeDataType)
    option7.pack(anchor="w")
    option8 = Radiobutton(top, text="Please Select One from Above", variable=var, value="None", command=storeDataType)
    option8.pack(anchor="w")

    label = Label(top)
    label.pack()
    top.mainloop()

def chooseDestDirectory():
    destPath = filedialog.askdirectory(initialdir="C:")
    if destPath != "":
        destEntry['text'] = destPath
        parameters["dest"] = destPath
        destButton['bg'] = "green"
        checkReadyToCompress()

def checkReadyToCompress():
    if sourceButton['bg'] == "green" and dataTypeButton['bg'] == "green" and destButton['bg'] == "green":
        compressButton['bg'] = "green"
    else:
        compressButton['bg'] = "red"

def resetParameters():
    sourceEntry['text'] = ""
    parameters["source"] = ""
    sourceButton['bg'] = "red"

    dataTypeEntry['text'] = ""
    parameters["data type"] = ""
    dataTypeButton['bg'] = "red"

    destEntry['text'] = ""
    parameters["dest"] = ""
    destButton['bg'] = "red"
    checkReadyToCompress()

def compress():
    if compressButton['bg'] == "red":
        messagebox.showerror(title="Compression Error", message="Please check that the Source Directory, Data Type/Resolution, and Destination Directory are valid selections (all 3 buttons should be green).\n\nA window will pop up after this message to display a valid configuration.")

        correct_selection = PhotoImage(file="correct_selection.png")

        top = Toplevel()
        top.iconbitmap("error.ico")
        top.resizable(0,0)
        top.geometry("450x150")
        top.title("Example of Valid Configuration")
        ref_pic = Label(top, image=correct_selection)
        ref_pic.grid(column=0, row=0)
        top.mainloop()
    else:
        srcPath = parameters["source"]
        dstPath = parameters["dest"]
        dataType = parameters["Data Type"]

        # Timing Execution
        start = time.time()

        zipFiles(srcPath, dstPath, dataType)

        # Timing Execution
        end = time.time()
        elapsed = end - start

        messagebox.showinfo(title="Status Update", message="Compression completed in:\r\n{0} seconds".format(elapsed))

        #Reset parameters
        resetParameters()

# ----------------------------------------------------------

# Menu Bar Functions -----------------------------------------
def displayAbout():
    logo = PhotoImage(file="sce_logo.png")

    top = Toplevel()
    top.iconbitmap("sce_icon.ico")
    top.resizable(0,0)
    top.geometry("255x260")
    top.title("About")
    logolabel = Label(top, image=logo)
    logolabel.grid(column=0, row=0)
    aboutlabel = Label(top, text="Created by:\r\nJulian Chan\r\nUndergraduate Summer Intern 2017\r\nPower Systems Technologies\r\nAdvanced Technology Group", justify="center", bg="lightblue", font="bold")
    versionlabel = Label(top, text="Version 1.0 (Release: July 20, 2017)", justify="center")
    aboutlabel.grid(column=0, row=1)
    versionlabel.grid(column=0, row=2)
    top.mainloop()

def displayFileStructure():
    logo = PhotoImage(file="filestructure.png")

    top = Toplevel()
    top.iconbitmap("sce_icon.ico")
    top.resizable(0,0)
    top.geometry("590x550")
    top.title("File Structure")
    logolabel = Label(top, image=logo)
    logolabel.grid(column=0, row=0)
    top.mainloop()
# ----------------------------------------------------------

root = tkinter.Tk()
# Set the window icon
root.iconbitmap("sce_icon.ico")
# Disable resizing
root.resizable(0,0)
# Set dimensions for home window
root.geometry("450x300")
root.title("Synchrophasor Data Compression")

# Maintain parameter dictionary for back end
parameters = {}

# Source Button
sourceButton = Button(root, text="Choose source directory", command=chooseSourceDirectory, width=20, wraplength=75, justify="center", bg="red")
sourceText = Label(root, text="\nSource Directory:\r\n")
sourceEntry = Label(root, text="", width=20, wraplength=75, justify="center")
sourceButton.grid(column=0, row=0)
sourceText.grid(column=0, row=1)
sourceEntry.grid(column=0, row=2)

# Data Type Button
dataTypeButton = Button(root, text="Choose type/resolution of data", command=chooseDataType, width=20, wraplength=85, justify="center", bg="red")
dataTypeText = Label(root, text="\nData Type/Resolution:\r\n")
dataTypeEntry = Label(root, text="", width=20, wraplength=75, justify="center")
dataTypeButton.grid(column=1, row=0)
dataTypeText.grid(column=1, row=1)
dataTypeEntry.grid(column=1, row=2)

# Destination Button
destButton = Button(root, text="Choose destination directory", command=chooseDestDirectory, width=20, wraplength=75, justify="center", bg="red")
destText = Label(root, text="\nDestination Directory:\r\n")
destEntry = Label(root, text="", width=20, wraplength=75, justify="center")
destButton.grid(column=2, row=0)
destText.grid(column=2, row=1)
destEntry.grid(column=2, row=2)

compressButton = Button(root, text="Compress!", command=compress, wid=20, wraplength=75, justify="center",bg="red")
compressButton.grid(column=1, row=4)

# Menu Options
menubar = Menu(root)
menubar.add_command(label="File Structure", command=displayFileStructure)
menubar.add_command(label="About Synchrophasor Data Compression...", command=displayAbout)
# helpmenu = Menu(menubar, tearoff=0)
# aboutmenu = Menu(menubar, tearoff=1)
# helpmenu.add_command(label="File Structure", command=displayFileStructure)
# aboutmenu.add_command(label="About Synchrophasor Data Compression...", command=displayAbout)
root.config(menu=menubar)
root.mainloop()
