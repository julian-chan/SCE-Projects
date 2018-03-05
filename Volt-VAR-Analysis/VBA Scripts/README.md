The VBA script "Volt Analysis 500kV VBA Macro.bas" analyzed violations in voltage and VAR levels given the system load and generated a summary for the current state of that specific substation at the AA Bank (500kV) level.

The VBA script "VoltVAR Combined Analysis 66kV VBA Macro.bas" analyzed violations in voltage and VAR levels given the system load and generated a summary for the current state of that specific substation at the A Bank (66kV) level.

The VBA script "Summary VBA Macro.bas" gathered all the individual substation summaries and aggregated the results in a table that presented an overview of the state of the system at that voltage level (66 kV).

The VBA scripts "Reformat Data VBA Macro.bas" and "Copy Schedules VBA Macro.bas" converted the .csv output of the database into a readable format for analysis and visualization in Python.

The VBA script "Data Cleanup 500kV.bas" removes inherent outliers in the data by averaging the two adjacent non-zero values. It removes errors in data collection due to hardware error when monitoring substation conditions.

The Jupyter notebook provides code to read in and process the data into appropriate data structures, and plots them along with a breakdown of violations by month of the year, day of the week, and hour of the day.
