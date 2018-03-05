import tkinter as tk
import tkinter.filedialog
import tkinter.messagebox
import time
import matplotlib.pyplot as plt
from plot_generator import generatePlots

class MainApplication:
    def __init__(self, master):
        self.master = master

        # Maximum number of bounds for each station
        self.num_bounds = 10
        # Initialize variables
        self.filepath = ""
        self.dirpath = ""
        self.station_name = ""
        self.bound_type = "None"
        self.bounds_entries = []
        self.VAR_bounds = []
        self.Volt_bounds = []

        # Remember the specific shade of gray for button reset
        self.original_color = self.master.cget("background")
        # Keep images of thumbs up/down for validating entry of information
        self.good = tk.PhotoImage(file="good.png")
        self.bad = tk.PhotoImage(file="bad.png")

        # Create the menubar
        self.createMenuBar()
        # Set the main window properties (e.g. dimensions, icon, title, etc.)
        self.setMainProperties()
        # Create main layout frames (and position them) to hold the widgets
        self.createMainFrames()
        # Create widget to lookup a file and save the path to that file
        self.createFileWidgets()
        # Create widget to allow for entry of a substation name and saves that name
        self.createNameWidgets()
        # Create widget to allow for selection of a bound type for the Voltage ranges
        self.createBoundTypeWidgets()
        # Create the widget to allow for entry of bounds given the bound type
        self.createBoundWidgets()
        # Create the widget to generate the plots of MW, MVAR, Voltage, and a temporal breakdown of each by month, day and hour
        self.createGenerateWidgets()

    def createMenuBar(self):
        # Menu Options
        self.menubar = tk.Menu(self.master)
        self.menubar.add_command(label="About", command=self.displayAbout)
        self.menubar.add_command(label="Help", command=self.displayHelp)
        self.master.config(menu=self.menubar)

    def setMainProperties(self):
        self.master.title("Volt VAR Summary Report Generator")
        self.master.iconbitmap("sce_icon.ico")
        self.master.geometry("{}x{}".format(810, 640))
        self.master.columnconfigure(0, weight=1)
        self.master.resizable(0, 0)

    def createMainFrames(self):
        # Main Application window frames
        self.fileframe = tk.Frame(self.master, width=740, height=75, borderwidth=2, relief="solid")
        self.nameframe = tk.Frame(self.master, width=740, height=75, borderwidth=2, relief="solid")
        self.boundtypeframe = tk.Frame(self.master, width=740, height=75, borderwidth=2, relief="solid")
        self.boundframe = tk.Frame(self.master, width=740, height=350, borderwidth=2, relief="solid")
        self.generateframe = tk.Frame(self.master, width=740, height=75, borderwidth=2, relief="solid")
        self.fileframecheck = tk.Frame(self.master, width=50, height=75, borderwidth=2, relief="solid")
        self.nameframecheck = tk.Frame(self.master, width=50, height=75, borderwidth=2, relief="solid")
        self.boundtypeframecheck = tk.Frame(self.master, width=50, height=75, borderwidth=2, relief="solid")
        self.boundframecheck = tk.Frame(self.master, width=50, height=350, borderwidth=2, relief="solid")
        self.generateframecheck = tk.Frame(self.master, width=50, height=75, borderwidth=2, relief="solid")

        self.fileframe.grid(row=0, column=0, sticky="nsew")
        self.nameframe.grid(row=1, column=0, sticky="nsew")
        self.boundtypeframe.grid(row=2, column=0, sticky="nsew")
        self.boundframe.grid(row=3, column=0, sticky="nsew")
        self.generateframe.grid(row=4, column=0, sticky="nsew")
        self.fileframecheck.grid(row=0, column=1, sticky="nsew")
        self.nameframecheck.grid(row=1, column=1, sticky="nsew")
        self.boundtypeframecheck.grid(row=2, column=1, sticky="nsew")
        self.boundframecheck.grid(row=3, column=1, sticky="nsew")
        self.generateframecheck.grid(row=4, column=1, sticky="nsew")

        # Widgets for the fileframecheck, nameframecheck, boundframecheck, generateframecheck
        self.filecheck = tk.Label(self.fileframecheck, image=self.bad)
        self.namecheck = tk.Label(self.nameframecheck, image=self.bad)
        self.boundtypecheck = tk.Label(self.boundtypeframecheck, image=self.bad)
        self.boundcheck = tk.Label(self.boundframecheck, image=self.bad)
        self.generatecheck = tk.Label(self.generateframecheck, image=self.bad)

        self.filecheck.grid(row=0, column=0, sticky="nsew")
        self.namecheck.grid(row=0, column=0, sticky="nsew")
        self.boundtypecheck.grid(row=0, column=0, sticky="nsew")
        self.boundcheck.grid(row=0, column=0, sticky="nsew")
        self.generatecheck.grid(row=0, column=0, sticky="nsew")

    def createFileWidgets(self):
        # Widgets for the fileframe
        self.instruction_label = tk.Label(self.fileframe,
                                          text="Please choose the data file you wish to generate a report for. Ensure that the file is in .xls or .xls* format.",
                                          justify="center")
        self.file_button = tk.Button(self.fileframe, text="Choose data file (.xls*)", command=self.locateFile,
                                     justify="center")
        self.file_label = tk.Label(self.fileframe, text="", justify="center")

        self.instruction_label.grid(row=0, column=0, columnspan=3, sticky="nw")
        self.file_button.grid(row=1, column=0, sticky="w")
        self.file_label.grid(row=2, column=0, sticky="w")

    def createNameWidgets(self):
        # Widgets for the nameframe
        self.name_instruction_label = tk.Label(self.nameframe, text="Please enter substation name:", justify="left")
        self.name_entry = tk.Entry(self.nameframe)
        self.name_entry_button = tk.Button(self.nameframe, text="OK", command=self.recordName)
        self.name_entry_label = tk.Label(self.nameframe, text="", justify="center")

        self.name_instruction_label.grid(row=0, column=0, sticky="nw")
        self.name_entry.grid(row=0, column=1, sticky="w")
        self.name_entry_button.grid(row=0, column=4, sticky="w")
        self.name_entry_label.grid(row=1, column=0, sticky="w")

    def createBoundTypeWidgets(self):
        # Widgets for boundtypeframe
        self.bound_type_instruction_label = tk.Label(self.boundtypeframe,
                                                      text="Please select the type of bounds for this substation:",
                                                      justify="left")
        self.bound_type_button = tk.Button(self.boundtypeframe, text="Select bounds type", command=self.boundsType)
        self.bound_type_label = tk.Label(self.boundtypeframe, text="", justify="center")

        self.bound_type_instruction_label.grid(row=0, column=0, columnspan=6, sticky="nw")
        self.bound_type_button.grid(row=1, column=0, sticky="w")
        self.bound_type_label.grid(row=2, column=0, sticky="w")

    def createBoundWidgets(self):
        # Widgets for boundframe
        self.bound_instruction_label1 = tk.Label(self.boundframe,
                                                 text="Please enter the SOB-17 ranges for this substation.\nIf the bounds type is 'all times' or 'range', fill in the 'Voltage' cells as if they were separate from the rest of the table.\nOtherwise, please fill in:\n       the 'MVAR' cells corresponding to the same row's 'MW' cells\n       the 'Voltage' cells corresponding to the same row's 'MVAR' cells",
                                                 justify="left")
        self.bound_instruction_label2 = tk.Label(self.boundframe,
                                                text="IMPORTANT: If there is no 'High MW' value because the SOB-17 says 'Above ___', please leave the cell blank.\nIMPORTANT: PLEASE MAKE SURE THE 'HIGH MW' OF THE PREVIOUS ROW IS EQUAL TO THE 'LOW MW' OF THE NEXT ROW!",
                                                foreground="red",
                                                justify="left")
        self.bound_entry_button = tk.Button(self.boundframe, text="OK", command=self.recordBounds)

        self.bound_instruction_label1.grid(row=0, column=0, columnspan=self.num_bounds+2, sticky="nw")
        self.bound_instruction_label2.grid(row=1, column=0, columnspan=self.num_bounds + 2, sticky="nw")
        self.bound_entry_button.grid(row=self.num_bounds+3, column=2, columnspan=2, sticky="nsew")

        self.bounds_entries = []
        self.VAR_bounds = []
        self.Volt_bounds = []

        for i in range(2, self.num_bounds+3):
            self.boundframe.columnconfigure(index=i, weight=1)
            self.boundframe.rowconfigure(index=i, weight=1)
            low_MW = tk.Entry(self.boundframe, state="disabled")
            high_MW = tk.Entry(self.boundframe, state="disabled")
            low_MVAR = tk.Entry(self.boundframe, state="disabled")
            high_MVAR = tk.Entry(self.boundframe, state="disabled")
            low_Voltage = tk.Entry(self.boundframe, state="disabled")
            high_Voltage = tk.Entry(self.boundframe, state="disabled")

            low_MW.grid(row=i, column=0, sticky="")
            high_MW.grid(row=i, column=1, sticky="")
            low_MVAR.grid(row=i, column=2, sticky="")
            high_MVAR.grid(row=i, column=3, sticky="")
            low_Voltage.grid(row=i, column=4, sticky="")
            high_Voltage.grid(row=i, column=5, sticky="")

    def createGenerateWidgets(self):
        self.generate_button = tk.Button(self.generateframe, text="Generate Summary Report", command=self.generateReport)
        self.generate_button.place(relx=0.5, rely=0.5, anchor="center")

    # Displays the about page when a user clicks on "About" in the menu bar
    def displayAbout(self):
        logo = tk.PhotoImage(file="sce_logo.png")

        top = tk.Toplevel()
        top.iconbitmap("sce_icon.ico")
        top.resizable(0, 0)
        top.geometry("255x260")
        top.title("About")
        logolabel = tk.Label(top, image=logo)
        logolabel.grid(column=0, row=0)
        aboutlabel = tk.Label(top,
                           text="Created by:\r\nJulian Chan\r\nUndergraduate Summer Intern 2017\r\nPower Systems Technologies\r\nAdvanced Technology Group",
                           justify="center", bg="lightblue", font="bold")
        versionlabel = tk.Label(top, text="Version 1.0 (Release: August 9, 2017)", justify="center")
        aboutlabel.grid(column=0, row=1)
        versionlabel.grid(column=0, row=2)
        top.mainloop()

    # Displays the help page when a user clicks on "Help" in the menu bar
    def displayHelp(self):
        logo = tk.PhotoImage(file="sce_logo.png")

        top = tk.Toplevel()
        top.iconbitmap("sce_icon.ico")
        top.resizable(0, 0)
        top.geometry("535x100")
        top.title("Help")
        logolabel = tk.Label(top, image=logo)
        logolabel.grid(column=0, row=0)
        helplabel = tk.Label(top,
                              text="Please see the 'Volt VAR Report Generator Manual' included with this program for more information.",
                              justify="center")
        helplabel.grid(column=0, row=1)
        top.mainloop()

    # ON-SELECT (called in boundsType()): alters the bounds grid for entry given the bound type is "all times"
    def allTimesGrid(self):
        self.bounds_entries = []
        low_MW_label = tk.Label(self.boundframe, text="Low MW")
        high_MW_label = tk.Label(self.boundframe, text="High MW")
        low_MVAR_label = tk.Label(self.boundframe, text="Low MVAR")
        high_MVAR_label = tk.Label(self.boundframe, text="High MVAR")
        dummy_label1 = tk.Label(self.boundframe, text="")
        dummy_label2 = tk.Label(self.boundframe, text="")
        dummy_label3 = tk.Label(self.boundframe, text="")
        Voltage_label = tk.Label(self.boundframe, text="Voltage")

        low_MW_label.grid(row=2, column=0, sticky="nsew")
        high_MW_label.grid(row=2, column=1, sticky="nsew")
        low_MVAR_label.grid(row=2, column=2, sticky="nsew")
        high_MVAR_label.grid(row=2, column=3, sticky="nsew")
        dummy_label1.grid(row=2, column=4, sticky="nsew")
        dummy_label2.grid(row=2, column=5, sticky="nsew")
        dummy_label3.grid(row=2, column=6, sticky="nsew")
        Voltage_label.grid(row=2, column=7, sticky="nsew")

        Voltage = tk.Entry(self.boundframe)
        Voltage.grid(row=3, column=7, sticky="nsew")

        for i in range(3, self.num_bounds + 3):
            low_MW = tk.Entry(self.boundframe)
            high_MW = tk.Entry(self.boundframe)
            low_MVAR = tk.Entry(self.boundframe)
            high_MVAR = tk.Entry(self.boundframe)
            dummy1 = tk.Entry(self.boundframe, state="disabled")
            dummy2 = tk.Entry(self.boundframe, state="disabled")
            dummy3 = tk.Entry(self.boundframe, state="disabled")

            if i > 3:
                dummy4 = tk.Entry(self.boundframe, state="disabled")
                dummy4.grid(row=i, column=7, sticky="nsew")

            low_MW.grid(row=i, column=0, sticky="nsew")
            high_MW.grid(row=i, column=1, sticky="nsew")
            low_MVAR.grid(row=i, column=2, sticky="nsew")
            high_MVAR.grid(row=i, column=3, sticky="nsew")
            dummy1.grid(row=i, column=4, sticky="nsew")
            dummy2.grid(row=i, column=5, sticky="nsew")
            dummy3.grid(row=i, column=6, sticky="nsew")


            self.bounds_entries.append(low_MW)
            self.bounds_entries.append(high_MW)
            self.bounds_entries.append(low_MVAR)
            self.bounds_entries.append(high_MVAR)
            self.bounds_entries.append(Voltage)

    # ON-SELECT (called in boundsType()): alters the bounds grid for entry given the bound type is "range"
    def rangeGrid(self):
        self.bounds_entries = []
        low_MW_label = tk.Label(self.boundframe, text="Low MW")
        high_MW_label = tk.Label(self.boundframe, text="High MW")
        low_MVAR_label = tk.Label(self.boundframe, text="Low MVAR")
        high_MVAR_label = tk.Label(self.boundframe, text="High MVAR")
        dummy_label1 = tk.Label(self.boundframe, text="")
        dummy_label2 = tk.Label(self.boundframe, text="")
        low_Voltage_label = tk.Label(self.boundframe, text="Low Voltage")
        high_Voltage_label = tk.Label(self.boundframe, text="High Voltage")

        low_MW_label.grid(row=2, column=0, sticky="nsew")
        high_MW_label.grid(row=2, column=1, sticky="nsew")
        low_MVAR_label.grid(row=2, column=2, sticky="nsew")
        high_MVAR_label.grid(row=2, column=3, sticky="nsew")
        dummy_label1.grid(row=2, column=4, sticky="nsew")
        dummy_label2.grid(row=2, column=5, sticky="nsew")
        low_Voltage_label.grid(row=2, column=6, sticky="nsew")
        high_Voltage_label.grid(row=2, column=7, sticky="nsew")

        low_Voltage = tk.Entry(self.boundframe)
        high_Voltage = tk.Entry(self.boundframe)
        low_Voltage.grid(row=3, column=6, sticky="nsew")
        high_Voltage.grid(row=3, column=7, sticky="nsew")

        for i in range(3, self.num_bounds + 3):
            low_MW = tk.Entry(self.boundframe)
            high_MW = tk.Entry(self.boundframe)
            low_MVAR = tk.Entry(self.boundframe)
            high_MVAR = tk.Entry(self.boundframe)
            dummy1 = tk.Entry(self.boundframe, state="disabled")
            dummy2 = tk.Entry(self.boundframe, state="disabled")

            if i > 3:
                dummy3 = tk.Entry(self.boundframe, state="disabled")
                dummy4 = tk.Entry(self.boundframe, state="disabled")
                dummy3.grid(row=i, column=6, sticky="nsew")
                dummy4.grid(row=i, column=7, sticky="nsew")

            low_MW.grid(row=i, column=0, sticky="nsew")
            high_MW.grid(row=i, column=1, sticky="nsew")
            low_MVAR.grid(row=i, column=2, sticky="nsew")
            high_MVAR.grid(row=i, column=3, sticky="nsew")
            dummy1.grid(row=i, column=4, sticky="nsew")
            dummy2.grid(row=i, column=5, sticky="nsew")

            self.bounds_entries.append(low_MW)
            self.bounds_entries.append(high_MW)
            self.bounds_entries.append(low_MVAR)
            self.bounds_entries.append(high_MVAR)
            self.bounds_entries.append(low_Voltage)
            self.bounds_entries.append(high_Voltage)

    # ON-SELECT (called in boundsType()): alters the bounds grid for entry given the bound type is "load dependent"
    def loadDependentGrid(self):
        self.bounds_entries = []
        low_MW_label1 = tk.Label(self.boundframe, text="Low MW")
        high_MW_label1 = tk.Label(self.boundframe, text="High MW")
        low_MVAR_label = tk.Label(self.boundframe, text="Low MVAR")
        high_MVAR_label = tk.Label(self.boundframe, text="High MVAR")
        dummy_label = tk.Label(self.boundframe, text="")
        low_MW_label2 = tk.Label(self.boundframe, text="Low MW")
        high_MW_label2 = tk.Label(self.boundframe, text="High MW")
        Voltage_label = tk.Label(self.boundframe, text="Voltage")

        low_MW_label1.grid(row=2, column=0, sticky="nsew")
        high_MW_label1.grid(row=2, column=1, sticky="nsew")
        low_MVAR_label.grid(row=2, column=2, sticky="nsew")
        high_MVAR_label.grid(row=2, column=3, sticky="nsew")
        dummy_label.grid(row=2, column=4, sticky="nsew")
        low_MW_label2.grid(row=2, column=5, sticky="nsew")
        high_MW_label2.grid(row=2, column=6, sticky="nsew")
        Voltage_label.grid(row=2, column=7, sticky="nsew")

        for i in range(3, self.num_bounds + 3):
            low_MW1 = tk.Entry(self.boundframe)
            high_MW1 = tk.Entry(self.boundframe)
            low_MVAR = tk.Entry(self.boundframe)
            high_MVAR = tk.Entry(self.boundframe)
            dummy = tk.Entry(self.boundframe, state="disabled")
            low_MW2 = tk.Entry(self.boundframe)
            high_MW2 = tk.Entry(self.boundframe)
            Voltage = tk.Entry(self.boundframe)

            low_MW1.grid(row=i, column=0, sticky="nsew")
            high_MW1.grid(row=i, column=1, sticky="nsew")
            low_MVAR.grid(row=i, column=2, sticky="nsew")
            high_MVAR.grid(row=i, column=3, sticky="nsew")
            dummy.grid(row=i, column=4, sticky="nsew")
            low_MW2.grid(row=i, column=5, sticky="nsew")
            high_MW2.grid(row=i, column=6, sticky="nsew")
            Voltage.grid(row=i, column=7, sticky="nsew")

            self.bounds_entries.append(low_MW1)
            self.bounds_entries.append(high_MW1)
            self.bounds_entries.append(low_MVAR)
            self.bounds_entries.append(high_MVAR)
            self.bounds_entries.append(low_MW2)
            self.bounds_entries.append(high_MW2)
            self.bounds_entries.append(Voltage)

    # ON-SELECT (called in boundsType()): alters the bounds grid for entry given the bound type is "load dependent range"
    def loadDependentRangeGrid(self):
        self.bounds_entries = []
        low_MW_label1 = tk.Label(self.boundframe, text="Low MW")
        high_MW_label1 = tk.Label(self.boundframe, text="High MW")
        low_MVAR_label = tk.Label(self.boundframe, text="Low MVAR")
        high_MVAR_label = tk.Label(self.boundframe, text="High MVAR")
        low_MW_label2 = tk.Label(self.boundframe, text="Low MW")
        high_MW_label2 = tk.Label(self.boundframe, text="High MW")
        low_Voltage_label = tk.Label(self.boundframe, text="Low Voltage")
        high_Voltage_label = tk.Label(self.boundframe, text="High Voltage")

        low_MW_label1.grid(row=2, column=0, sticky="nsew")
        high_MW_label1.grid(row=2, column=1, sticky="nsew")
        low_MVAR_label.grid(row=2, column=2, sticky="nsew")
        high_MVAR_label.grid(row=2, column=3, sticky="nsew")
        low_MW_label2.grid(row=2, column=4, sticky="nsew")
        high_MW_label2.grid(row=2, column=5, sticky="nsew")
        low_Voltage_label.grid(row=2, column=6, sticky="nsew")
        high_Voltage_label.grid(row=2, column=7, sticky="nsew")

        for i in range(3, self.num_bounds + 3):
            low_MW1 = tk.Entry(self.boundframe)
            high_MW1 = tk.Entry(self.boundframe)
            low_MVAR = tk.Entry(self.boundframe)
            high_MVAR = tk.Entry(self.boundframe)
            low_MW2 = tk.Entry(self.boundframe)
            high_MW2 = tk.Entry(self.boundframe)
            low_Voltage = tk.Entry(self.boundframe)
            high_Voltage = tk.Entry(self.boundframe)

            low_MW1.grid(row=i, column=0, sticky="nsew")
            high_MW1.grid(row=i, column=1, sticky="nsew")
            low_MVAR.grid(row=i, column=2, sticky="nsew")
            high_MVAR.grid(row=i, column=3, sticky="nsew")
            low_MW2.grid(row=i, column=4, sticky="nsew")
            high_MW2.grid(row=i, column=5, sticky="nsew")
            low_Voltage.grid(row=i, column=6, sticky="nsew")
            high_Voltage.grid(row=i, column=7, sticky="nsew")

            self.bounds_entries.append(low_MW1)
            self.bounds_entries.append(high_MW1)
            self.bounds_entries.append(low_MVAR)
            self.bounds_entries.append(high_MVAR)
            self.bounds_entries.append(low_MW2)
            self.bounds_entries.append(high_MW2)
            self.bounds_entries.append(low_Voltage)
            self.bounds_entries.append(high_Voltage)

    # ON-CLICK (called in createFileWidgets()): opens a dialog box for the user to select the data file
    def locateFile(self):
        self.filepath = tk.filedialog.askopenfilename()
        if self.filepath[-3:] != "xls" and self.filepath[-4:-1] != "xls":
            tk.messagebox.showerror(title="File Selection Error", message="Please choose a valid data file (must be in .xls or .xls* format.")
            return
        self.file_label["text"] = self.filepath
        if self.filepath != "":
            self.file_label["bg"] = "lightgreen"
            self.filecheck["image"] = self.good
        else:
            self.file_label["bg"] = self.original_color
            self.filecheck["image"] = self.bad
        self.checkReadyToGenerateReport()

    # ON-CLICK (called in createNameWidgets()): saves the name entered upon pressed button
    def recordName(self):
        self.station_name = self.name_entry.get()
        self.name_entry_label["text"] = self.station_name
        if self.station_name != "":
            self.name_entry_label["bg"] = "lightgreen"
            self.namecheck["image"] = self.good
            tk.messagebox.showinfo(title="Input Successful", message="Station name saved.")
        else:
            self.name_entry_label["bg"] = self.original_color
            self.namecheck["image"] = self.bad
        self.checkReadyToGenerateReport()

    # ON-CLICK (called in createBoundTypeWidgets()): opens a window for the user to select the bound type upon pressed button
    def boundsType(self):
        def setBounds():
            self.bound_type = var.get()
            self.bound_type_label["text"] = self.bound_type
            if self.bound_type != "None":
                self.bound_type_label["bg"] = "lightgreen"
                self.boundtypecheck["image"] = self.good
                tk.messagebox.showinfo(title="Input Successful", message="Bounds type saved.")
                if self.bound_type == "all times":
                    self.allTimesGrid()
                elif self.bound_type == "range":
                    self.rangeGrid()
                elif self.bound_type == "load dependent":
                    self.loadDependentGrid()
                elif self.bound_type == "load dependent range":
                    self.loadDependentRangeGrid()
            else:
                self.bound_type_label["text"] = ""
                self.bound_type_label["bg"] = self.original_color
                self.boundtypecheck["image"] = self.bad
                tk.messagebox.showerror(title="Invalid Selection",
                                     message="Please select a valid value. Valid values are:\nAll Times\nRange\nLoad Dependent\nLoad Dependent Range")
            top.destroy()
            self.checkReadyToGenerateReport()
        top = tk.Toplevel()
        top.iconbitmap("sce_icon.ico")
        top.resizable(0, 0)
        top.geometry("600x200")
        top.title("Choose bounds type:")
        var = tk.StringVar()
        var.set("None")
        option1 = tk.Radiobutton(top, text="All Times - Voltage should be kept at a constant value regardless of the MVAR value", variable=var, value="all times", command=setBounds)
        option1.pack(anchor="w")
        option2 = tk.Radiobutton(top, text="Range - Voltage should be kept within a range regardless of the MVAR value", variable=var, value="range", command=setBounds)
        option2.pack(anchor="w")
        option3 = tk.Radiobutton(top, text="Load Dependent - Voltage should be kept at a constant value depending on the MVAR value", variable=var, value="load dependent", command=setBounds)
        option3.pack(anchor="w")
        option4 = tk.Radiobutton(top, text="Load Dependent Range - Voltage should be kept within a range depending on the MVAR value", variable=var, value="load dependent range", command=setBounds)
        option4.pack(anchor="w")
        option5 = tk.Radiobutton(top, text="Please Select One from Above", variable=var, value="None", command=setBounds)
        option5.pack(anchor="w")
        top.mainloop()

    # ON-CLICK (called in createBoundWidgets()): saves the bounds entered in arrays upon pressed button
    def recordBounds(self):
        """
        VAR Bounds format:
            (List) list of tuples (low_MW, high_MW, low_MVAR, high_MVAR)

        Voltage Bounds format:
            "all times" - (Integer) reference voltage
            "range" - (Tuple) tuple of (low_Volt, high_Volt)
            "load dependent" - (List) list of tuples (low_MW, high_MW, reference voltage)
            "load dependent range" - (List) list of tuples (low_MW, high_MW, low_Volt, high_Volt)
        """
        self.VAR_bounds = []
        self.Volt_bounds = []

        if self.bound_type == "None":
            tk.messagebox.showerror(title="Input Error", message="Please ensure that a bound type has been selected.")
            return

        if all([entry.get() == "" for entry in self.bounds_entries]):
            tk.messagebox.showerror(title="Input Error", message="Please enter valid bounds found in the SOB-17.\nA report cannot be generated without criteria being specified.")
            return

        if self.bound_type == "all times":
            for i in range(len(self.bounds_entries) // 5):
                if self.bounds_entries[5 * i + 4].get() == "":
                    tk.messagebox.showerror(title="Input Error",
                                            message="Please ensure that the 'Voltage' cell is filled in!")
                    self.boundcheck["image"] = self.bad
                    return
                else:
                    Voltage = float(self.bounds_entries[5 * i + 4].get())
                self.Volt_bounds = [Voltage]

                if self.bounds_entries[5 * i].get() == "":
                    if self.bounds_entries[5 * i + 1].get() != "" or self.bounds_entries[5 * i + 2].get() != "" or self.bounds_entries[5 * i + 3].get() != "":
                        tk.messagebox.showerror(title="Input Error",
                                                message="Please ensure that the 'Low MW' cell is filled in!")
                        self.boundcheck["image"] = self.bad
                        return
                    else:
                        break
                else:
                    low_MW = float(self.bounds_entries[5 * i].get())

                high_MW = ""
                if self.bounds_entries[5 * i + 1].get() != "":
                    high_MW = float(self.bounds_entries[5 * i + 1].get())

                if self.bounds_entries[5 * i + 2].get() == "":
                    tk.messagebox.showerror(title="Input Error",
                                            message="Please ensure that the 'Low MVAR' cell is filled in!")
                    self.boundcheck["image"] = self.bad
                    return
                else:
                    low_MVAR = float(self.bounds_entries[5 * i + 2].get())

                if self.bounds_entries[5 * i + 3].get() == "":
                    tk.messagebox.showerror(title="Input Error",
                                            message="Please ensure that the 'High MVAR' cell is filled in!")
                    self.boundcheck["image"] = self.bad
                    return
                else:
                    high_MVAR = float(self.bounds_entries[5 * i + 3].get())

                if high_MW == "":
                    self.VAR_bounds.append((low_MW, low_MVAR, high_MVAR))
                else:
                    self.VAR_bounds.append((low_MW, high_MW, low_MVAR, high_MVAR))
        elif self.bound_type == "range":
            for i in range(len(self.bounds_entries) // 6):
                if self.bounds_entries[6 * i + 4].get() == "":
                    tk.messagebox.showerror(title="Input Error",
                                            message="Please ensure that the 'Low Voltage' cell is filled in!")
                    self.boundcheck["image"] = self.bad
                    return
                else:
                    low_Voltage = float(self.bounds_entries[6 * i + 4].get())

                if self.bounds_entries[6 * i + 5].get() == "":
                    tk.messagebox.showerror(title="Input Error",
                                            message="Please ensure that the 'High Voltage' cell is filled in!")
                    self.boundcheck["image"] = self.bad
                    return
                else:
                    high_Voltage = float(self.bounds_entries[6 * i + 5].get())
                self.Volt_bounds = (low_Voltage, high_Voltage)

                if self.bounds_entries[6 * i].get() == "":
                    if self.bounds_entries[6 * i + 1].get() != "" or self.bounds_entries[6 * i + 2].get() != "" or self.bounds_entries[6 * i + 3].get() != "":
                        tk.messagebox.showerror(title="Input Error",
                                                message="Please ensure that the 'Low MW' cell is filled in!")
                        self.boundcheck["image"] = self.bad
                        return
                    else:
                        break
                else:
                    low_MW = float(self.bounds_entries[6 * i].get())

                high_MW = ""
                if self.bounds_entries[6 * i + 1].get() != "":
                    high_MW = float(self.bounds_entries[6 * i + 1].get())

                if self.bounds_entries[6 * i + 2].get() == "":
                    tk.messagebox.showerror(title="Input Error",
                                            message="Please ensure that the 'Low MVAR' cell is filled in!")
                    self.boundcheck["image"] = self.bad
                    return
                else:
                    low_MVAR = float(self.bounds_entries[6 * i + 2].get())

                if self.bounds_entries[6 * i + 3].get() == "":
                    tk.messagebox.showerror(title="Input Error",
                                            message="Please ensure that the 'High MVAR' cell is filled in!")
                    self.boundcheck["image"] = self.bad
                    return
                else:
                    high_MVAR = float(self.bounds_entries[6 * i + 3].get())

                if high_MW == "":
                    self.VAR_bounds.append((low_MW, low_MVAR, high_MVAR))
                else:
                    self.VAR_bounds.append((low_MW, high_MW, low_MVAR, high_MVAR))
        elif self.bound_type == "load dependent":
            for i in range(len(self.bounds_entries) // 7):
                if self.bounds_entries[7 * i + 6].get() == "" and self.bounds_entries[7 * i + 4].get() != "":
                    tk.messagebox.showerror(title="Input Error",
                                            message="Please ensure that the 'Voltage' cell is filled in!")
                    self.boundcheck["image"] = self.bad
                    return
                else:
                    Voltage = self.bounds_entries[7 * i + 6].get()

                low_MW1 = ""
                if self.bounds_entries[7 * i].get() == "":
                    if self.bounds_entries[7 * i + 1].get() != "" or self.bounds_entries[7 * i + 2].get() != "" or self.bounds_entries[7 * i + 3].get() != "":
                        tk.messagebox.showerror(title="Input Error",
                                                message="Please ensure that the 'Low MW' cell is filled in!")
                        self.boundcheck["image"] = self.bad
                        return
                else:
                    low_MW1 = self.bounds_entries[7 * i].get()

                high_MW1 = ""
                if self.bounds_entries[7 * i + 1].get() != "":
                    high_MW1 = self.bounds_entries[7 * i + 1].get()

                if self.bounds_entries[7 * i + 2].get() == "" and self.bounds_entries[7 * i].get() != "":
                    tk.messagebox.showerror(title="Input Error",
                                            message="Please ensure that the 'Low MVAR' cell is filled in!")
                    self.boundcheck["image"] = self.bad
                    return
                else:
                    low_MVAR = self.bounds_entries[7 * i + 2].get()

                if self.bounds_entries[7 * i + 3].get() == "" and self.bounds_entries[7 * i].get() != "":
                    tk.messagebox.showerror(title="Input Error",
                                            message="Please ensure that the 'High MVAR' cell is filled in!")
                    self.boundcheck["image"] = self.bad
                    return
                else:
                    high_MVAR = self.bounds_entries[7 * i + 3].get()

                low_MW2 = ""
                if self.bounds_entries[7 * i + 4].get() == "":
                    if self.bounds_entries[7 * i + 5].get() != "" or self.bounds_entries[7 * i + 6].get() != "":
                        tk.messagebox.showerror(title="Input Error",
                                                message="Please ensure that the 'Low MW' cell is filled in!")
                        self.boundcheck["image"] = self.bad
                        return
                else:
                    low_MW2 = self.bounds_entries[7 * i + 4].get()

                high_MW2 = ""
                if self.bounds_entries[7 * i + 5].get() != "":
                    high_MW2 = self.bounds_entries[7 * i + 5].get()

                if low_MW1 != "" and high_MW1 == "":
                    self.VAR_bounds.append((float(low_MW1), float(low_MVAR), float(high_MVAR)))
                elif low_MW1 != "" and high_MW1 != "":
                    self.VAR_bounds.append((float(low_MW1), float(high_MW1), float(low_MVAR), float(high_MVAR)))

                if low_MW2 != "" and high_MW2 == "":
                    self.Volt_bounds.append((float(low_MW2), float(Voltage)))
                elif low_MW2 != "" and high_MW2 != "":
                    self.Volt_bounds.append((float(low_MW2), float(high_MW2), float(Voltage)))
        elif self.bound_type == "load dependent range":
            for i in range(len(self.bounds_entries) // 8):
                if self.bounds_entries[8 * i + 6].get() == "" and self.bounds_entries[8 * i + 4].get() != "":
                    tk.messagebox.showerror(title="Input Error",
                                            message="Please ensure that the 'Low Voltage' cell is filled in!")
                    self.boundcheck["image"] = self.bad
                    return
                else:
                    low_Voltage = self.bounds_entries[8 * i + 6].get()

                if self.bounds_entries[8 * i + 7].get() == "" and self.bounds_entries[8 * i + 4].get() != "":
                    tk.messagebox.showerror(title="Input Error",
                                            message="Please ensure that the 'High Voltage' cell is filled in!")
                    self.boundcheck["image"] = self.bad
                    return
                else:
                    high_Voltage = self.bounds_entries[8 * i + 7].get()

                if self.bounds_entries[8 * i].get() == "":
                    if self.bounds_entries[8 * i + 1].get() != "" or self.bounds_entries[8 * i + 2].get() != "" or self.bounds_entries[8 * i + 3].get() != "":
                        tk.messagebox.showerror(title="Input Error",
                                                message="Please ensure that the 'Low MW' cell is filled in!")
                        self.boundcheck["image"] = self.bad
                        return
                    else:
                        break
                else:
                    low_MW1 = self.bounds_entries[8 * i].get()

                high_MW1 = ""
                if self.bounds_entries[8 * i + 1].get() != "":
                    high_MW1 = self.bounds_entries[8 * i + 1].get()

                if self.bounds_entries[8 * i + 2].get() == "" and self.bounds_entries[8 * i] != "":
                    tk.messagebox.showerror(title="Input Error",
                                            message="Please ensure that the 'Low MVAR' cell is filled in!")
                    self.boundcheck["image"] = self.bad
                    return
                else:
                    low_MVAR = self.bounds_entries[8 * i + 2].get()

                if self.bounds_entries[8 * i + 3].get() == "" and self.bounds_entries[8 * i] != "":
                    tk.messagebox.showerror(title="Input Error",
                                            message="Please ensure that the 'High MVAR' cell is filled in!")
                    self.boundcheck["image"] = self.bad
                    return
                else:
                    high_MVAR = self.bounds_entries[8 * i + 3].get()

                low_MW2 = ""
                if self.bounds_entries[8 * i + 4].get() == "":
                    if self.bounds_entries[8 * i + 5].get() != "" or self.bounds_entries[8 * i + 6].get() != "" or self.bounds_entries[8 * i + 7].get() != "":
                        tk.messagebox.showerror(title="Input Error",
                                                message="Please ensure that the 'Low MW' cell is filled in!")
                        self.boundcheck["image"] = self.bad
                        return
                else:
                    low_MW2 = self.bounds_entries[8 * i + 4].get()

                high_MW2 = ""
                if self.bounds_entries[8 * i + 5].get() != "":
                    high_MW2 = self.bounds_entries[8 * i + 5].get()

                if low_MW1 != "" and high_MW1 == "":
                    self.VAR_bounds.append((float(low_MW1), float(low_MVAR), float(high_MVAR)))
                elif low_MW1 != "" and high_MW1 != "":
                    self.VAR_bounds.append((float(low_MW1), float(high_MW1), float(low_MVAR), float(high_MVAR)))

                if low_MW2 != "" and high_MW2 == "":
                    self.Volt_bounds.append((float(low_MW2), float(low_Voltage), float(high_Voltage)))
                elif low_MW2 != "" and high_MW2 != "":
                    self.Volt_bounds.append((float(low_MW2), float(high_MW2), float(low_Voltage), float(high_Voltage)))

        print("VAR Bounds: " + str(self.VAR_bounds))
        print("Voltage Bounds: " + str(self.Volt_bounds))
        self.boundcheck["image"] = self.good
        tk.messagebox.showinfo(title="Input Successful", message="Bounds saved.")
        self.checkReadyToGenerateReport()

    # Checks if a report is ready to be generated after each widget is filled in (changes the thumbs up icon also)
    def checkReadyToGenerateReport(self):
        if self.filepath != "" and self.station_name != "" and self.bound_type != "None" and (len(self.VAR_bounds) > 0 or len(self.Volt_bounds) > 0):
            self.generatecheck["image"] = self.good
        else:
            self.generatecheck["image"] = self.bad

    # ON-CLICK (called in createGenerateWidgets()): runs generatePlots() in plot_generator.py upon pressed button
    def generateReport(self):
        if self.filepath == "" or self.station_name == "" or self.bound_type == "None" and (len(self.VAR_bounds) > 0 or len(self.Volt_bounds) > 0):
            tk.messagebox.showerror(title="Input Error",
                                    message="Please ensure that all fields have been filled correctly.\nIf filled correctly, the right-hand column should all be thumbs up.")
            self.generatecheck["image"] = self.bad
        else:
            start = time.time()
            generatePlots(self.station_name, self.filepath, self.VAR_bounds, self.Volt_bounds, self.bound_type)
            end = time.time()
            tk.messagebox.showinfo(title="Report Generated", message="Summary Report for '" + self.station_name + "' successfully generated in: " + str(end - start) + " seconds!")
            plt.show()

# Application startup and execution
root = tk.Tk()
app = MainApplication(root)
root.mainloop()

