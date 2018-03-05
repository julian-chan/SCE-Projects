import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime
from matplotlib.backends.backend_pdf import PdfPages
import os

def MaxMonthlyViolation(dates, actual, high, low, include_non_violations=True):
    """
    This function finds the month where there was the greatest total amount of violation.

    INPUT:
        dates: array of dates and times aligned with the data arrays
        actual: actual MVAR load or kV data
        high: upper boundary on MVAR load or kV data range
        low: lower boundary on MVAR load or kV data range
        include_non_violations: boolean whether or not to include non-violations (as 0's) in data

    OUTPUT:
        list of means of violation in each month
    """
    # Initialize a dictionary, January 2016 = 0, March 2017 = 15
    violations = {}
    for relative_month in range(18):
        current_month = datetime(2016 + relative_month // 12, (relative_month % 12) + 1, 1).strftime("%m/%Y")
        violations[current_month] = []

    # Fill the respective arrays with violations
    for date_index in range(len(actual)):
        if len(high) > 0:
            max_value = high[date_index]
        else:
            max_value = 0

        if len(low) > 0:
            min_value = low[date_index]
        else:
            min_value = 0

        value = actual[date_index]

        current_month = dates[date_index].strftime("%m/%Y")

        # If the value is above the range, the violation is positive. If the value is below the range, the violation is negative.
        if value > max_value:
            violations[current_month].append(value - max_value)
        elif value < min_value:
            violations[current_month].append(value - min_value)
        else:
            if include_non_violations:
                violations[current_month].append(0)

    # Compute the medians of each array as a representative value of the dataset
    for key, arr in violations.items():
        if len(arr) != 0:
            violations[key] = np.median(arr)
        else:
            violations[key] = 0
    return violations


def MaxDailyViolation(dates, actual, high, low, include_non_violations=True):
    """
    This function finds the day where there was the greatest total amount of violation.

    INPUT:
        dates: array of dates and times aligned with the data arrays
        actual: actual MVAR load or kV data
        high: upper boundary on MVAR load or kV data range
        low: lower boundary on MVAR load or kV data range
        include_non_violations: boolean whether or not to include non-violations (as 0's) in data

    OUTPUT:
        list of means of violation in each day
    """
    # Initialize a dictionary; Monday = 0, Sunday = 6
    days_dict = {0: "Monday", 1: "Tuesday", 2: "Wednesday", 3: "Thursday", 4: "Friday", 5: "Saturday", 6: "Sunday"}
    violations = {}
    for day_num, day_name in days_dict.items():
        violations[day_name] = []

    # Fill the respective arrays with violations
    for date_index in range(len(actual)):
        if len(high) > 0:
            max_value = high[date_index]
        else:
            max_value = 0

        if len(low) > 0:
            min_value = low[date_index]
        else:
            min_value = 0

        value = actual[date_index]

        day_num = dates[date_index].weekday()
        day_name = days_dict[day_num]

        # If the value is above the range, the violation is positive. If the value is below the range, the violation is negative.
        if value > max_value:
            violations[day_name].append(value - max_value)
        elif value < min_value:
            violations[day_name].append(value - min_value)
        else:
            if include_non_violations:
                violations[day_name].append(0)

    # Compute the medians of each array as a representative value of the dataset
    for key, arr in violations.items():
        if len(arr) != 0:
            violations[key] = np.median(arr)
        else:
            violations[key] = 0
    return violations


def MaxHourlyViolation(dates, actual, high, low, include_non_violations=True):
    """
    This function finds the hour of the day where there was the greatest total amount of violation.

    INPUT:
        dates: array of dates and times aligned with the data arrays
        actual: actual MVAR load or kV data
        high: upper boundary on MVAR load or kV data range
        low: lower boundary on MVAR load or kV data range
        include_non_violations: boolean whether or not to include non-violations (as 0's) in data

    OUTPUT:
        list of means of violation in each hour
    """
    # Initialize a dictionary; 0 = 00:00, 23 = 23:00
    violations = {}
    for hour in range(24):
        violations[hour] = []

    # Fill the respective arrays with violations
    for date_index in range(len(actual)):
        if len(high) > 0:
            max_value = high[date_index]
        else:
            max_value = 0

        if len(low) > 0:
            min_value = low[date_index]
        else:
            min_value = 0

        value = actual[date_index]

        current_hour = dates[date_index].hour

        # If the value is above the range, the violation is positive. If the value is below the range, the violation is negative.
        if value > max_value:
            violations[current_hour].append(value - max_value)
        elif value < min_value:
            violations[current_hour].append(value - min_value)
        else:
            if include_non_violations:
                violations[current_hour].append(0)

    # Compute the medians of each array as a representative value of the dataset
    for key, arr in violations.items():
        if len(arr) != 0:
            violations[key] = np.median(arr)
        else:
            violations[key] = 0
    return violations


def plotMW(station_name, dates, MW_load, file):
    """
    Plots the MW load on the bank against time.

    INPUT:
        station_name: (String) name of substation to plot
        dates: (List) list of datetime objects incremented every minute
        MW_load: (List) list of MW-load values corresponding to the list of dates
        file: pdf file in which the plot will be saved
    """

    fig = plt.figure(figsize=(18, 5))
    ax = fig.add_subplot(1, 1, 1)
    ax.set_title(station_name + " MW")
    ax.set_xlabel("Date")
    ax.set_ylabel("MW")
    ax.grid(True)
    ax.scatter(dates, MW_load, 1, 'lightgreen', label='MW Load')
    ax.legend(loc='lower right')

    # If saving to PDF, uncomment both of the lines below; otherwise, leave them commented out
    # ax.set_rasterized(True)
    # file.savefig(bbox_inches='tight')

def plotMVAR(station_name, dates, MW_load, MVAR_load, boundaries, file):
    """
    Plots the MVAr load on the bank against time.

    INPUT:
        station_name: (String) name of substation to plot
        dates: (List) list of datetime objects incremented every minute
        MW_load: (List) list of MW-load values corresponding to the list of dates
        MVAR_load: (List) list of MVAR-load values corresponding to the list of dates
        boundaries: (List) list of tuples (low_MW, high_MW, low_MVAR, high_MVAR)
        file: pdf file in which the plot will be saved

    OUTPUT:
        lower: (List) lower boundary for the voltage (for use as parameter to plotBreakdown(...))
        upper: (List) upper boundary for the voltage (for use as parameter to plotBreakdown(...))
    """

    fig = plt.figure(figsize=(18, 5))
    ax = fig.add_subplot(1, 1, 1)
    ax.set_title(station_name + " MVAR")
    ax.set_xlabel("Date")
    ax.set_ylabel("MVAR")
    ax.grid(True)
    ax.scatter(dates, MVAR_load, 1, 'lightgreen', label='MVAR Load')

    # Create arrays to store the values of the boundaries at each point
    lower = []
    upper = []

    if len(boundaries) > 0:
        for i in range(len(MVAR_load)):
            load = MW_load[i]

            if load < 0:
                lower.append(0)
                upper.append(0)

            for bound in boundaries:
                low_MW = bound[0]
                high_MW = bound[1] if len(bound) == 4 else float('inf')
                low_MVAR = bound[2] if len(bound) == 4 else bound[1]
                high_MVAR = bound[3] if len(bound) == 4 else bound[2]

                if load >= low_MW and load < high_MW:
                    upper.append(high_MVAR)
                    lower.append(low_MVAR)
                    break

        plt.plot(dates, upper, 'r--', label='Upper Boundary')
        plt.plot(dates, lower, 'b--', label='Lower Boundary')

    plt.legend(loc='lower right')

    # If saving to PDF, uncomment both of the lines below; otherwise, leave them commented out
    # ax.set_rasterized(True)
    # file.savefig(bbox_inches='tight')

    return lower, upper

def plotVoltage(station_name, dates, MW_load, Voltage, boundaries, bound_type, file):
    """
    Plots the voltage across the bank against time.

    INPUT:
        station_name: (String) name of substation to plot
        dates: (List) list of datetime objects incremented every minute
        MW_load: (List) list of MW-load values corresponding to the list of dates
        Voltage: (List) list of Voltage values corresponding to the list of dates
        boundaries: Dependent on bound_type as follows
                "all times" - (Integer) reference voltage
                "range" - (Tuple) tuple of (low_Volt, high_Volt)
                "load dependent" - (List) list of tuples (low_MW, high_MW, reference voltage)
                "load dependent range" - (List) list of tuples (low_MW, high_MW, low_Volt, high_Volt)
        bound_type: (String) designation for the bound type - "all times", "range", "load dependent", "load dependent range"
        file: pdf file in which the plot will be saved

    OUTPUT:
        lower: (List) lower boundary for the voltage (for use as parameter to plotBreakdown(...))
        upper: (List) upper boundary for the voltage (for use as parameter to plotBreakdown(...))
    """

    fig = plt.figure(figsize=(18, 5))
    ax = fig.add_subplot(1, 1, 1)
    ax.set_title(station_name + " kV")
    ax.set_xlabel("Date")
    ax.set_ylabel("kV")
    ax.grid(True)
    ax.scatter(dates, Voltage, 1, 'lightgreen', label='Voltage')

    # Create arrays to store the values of the boundaries at each point
    lower = []
    upper = []

    if bound_type == "all times":
        const_voltage = boundaries[0]
        Ref_Voltage = [const_voltage for _ in range(len(Voltage))]
        lower = [0.98 * Ref_Voltage[i] for i in range(len(Voltage))]
        upper = [1.02 * Ref_Voltage[i] for i in range(len(Voltage))]
        plt.plot(dates, Ref_Voltage, 'k', label='Reference Voltage')
        plt.plot(dates, upper, 'r--', label='2% Upper Boundary')
        plt.plot(dates, lower, 'b--', label='2% Lower Boundary')
    elif bound_type == "range":
        low_voltage = boundaries[0]
        high_voltage = boundaries[1]
        lower = [low_voltage for _ in range(len(Voltage))]
        upper = [high_voltage for _ in range(len(Voltage))]
        plt.plot(dates, upper, 'r--', label='Upper SOB-17 Boundary')
        plt.plot(dates, lower, 'b--', label='Lower SOB-17 Boundary')
    elif bound_type == "load dependent":
        Ref_Voltage = []
        lower = []
        upper = []

        for i in range(len(Voltage)):
            load = MW_load[i]

            if load < 0:
                lower.append(0)
                upper.append(0)
                Ref_Voltage.append(0)

            for bound in boundaries:
                low_MW = bound[0]
                high_MW = bound[1] if len(bound) == 3 else float('inf')
                ref_volt = bound[2] if len(bound) == 3 else bound[1]

                if load >= low_MW and load < high_MW:
                    Ref_Voltage.append(ref_volt)
                    lower.append(0.98 * ref_volt)
                    upper.append(1.02 * ref_volt)
                    break
        plt.plot(dates, Ref_Voltage, 'k', label='Reference Voltage')
        plt.plot(dates, upper, 'r--', label='2% Upper Boundary')
        plt.plot(dates, lower, 'b--', label='2% Lower Boundary')
    elif bound_type == "load dependent range":
        lower = []
        upper = []

        for i in range(len(Voltage)):
            load = MW_load[i]

            if load < 0:
                lower.append(0)
                upper.append(0)

            for bound in boundaries:
                low_MW = bound[0]
                high_MW = bound[1] if len(bound) == 4 else float('inf')
                low_volt = bound[2] if len(bound) == 4 else bound[1]
                high_volt = bound[3] if len(bound) == 4 else bound[2]

                if load >= low_MW and load < high_MW:
                    lower.append(low_volt)
                    upper.append(high_volt)
                    break

        plt.plot(dates, upper, 'r--', label='Upper SOB-17 Boundary')
        plt.plot(dates, lower, 'b--', label='Lower SOB-17 Boundary')

    plt.legend(loc='upper right')

    # If saving to PDF, uncomment both of the lines below; otherwise, leave them commented out
    # ax.set_rasterized(True)
    # file.savefig(bbox_inches='tight')

    return lower, upper

def plotBreakdown(station_name, dates, MVAR_load, Voltage, lower_VAR, upper_VAR, lower_Voltage,
                  upper_Voltage, file):
    """
    Plots the breakdown of violations categorized by month of the year, day of the week, and hour of the day.

    Aggregates all violations into a list (no violation means a value of 0) and the median is computed for each
    item in the category.

    INPUT:
        station_name: (String) name of substation to plot
        dates: (List) list of datetime objects incremented every minute
        MVAR_load: (List) list of MVAR-load values corresponding to the list of dates
        Voltage: (List) list of Voltage values corresponding to the list of dates
        lower_VAR: (List) lower boundary for the VAR
        upper_VAR: (List) upper boundary for the VAR
        lower_Voltage: (List) lower boundary for the voltage
        upper_Voltage: (List) upper boundary for the voltage
        file: pdf file in which the bar graphs will be saved
    """
    monthly_VAR = MaxMonthlyViolation(dates, MVAR_load, upper_VAR, lower_VAR)
    daily_VAR = MaxDailyViolation(dates, MVAR_load, upper_VAR, lower_VAR)
    hourly_VAR = MaxHourlyViolation(dates, MVAR_load, upper_VAR, lower_VAR)

    monthly_Voltage = MaxMonthlyViolation(dates, Voltage, upper_Voltage, lower_Voltage)
    daily_Voltage = MaxDailyViolation(dates, Voltage, upper_Voltage, lower_Voltage)
    hourly_Voltage = MaxHourlyViolation(dates,Voltage, upper_Voltage, lower_Voltage)

    Xmonthly_VAR = MaxMonthlyViolation(dates, MVAR_load, upper_VAR, lower_VAR, False)
    Xdaily_VAR = MaxDailyViolation(dates, MVAR_load, upper_VAR, lower_VAR, False)
    Xhourly_VAR = MaxHourlyViolation(dates, MVAR_load, upper_VAR, lower_VAR, False)

    Xmonthly_Voltage = MaxMonthlyViolation(dates, Voltage, upper_Voltage, lower_Voltage, False)
    Xdaily_Voltage = MaxDailyViolation(dates, Voltage, upper_Voltage, lower_Voltage, False)
    Xhourly_Voltage = MaxHourlyViolation(dates, Voltage, upper_Voltage, lower_Voltage, False)

    fig = plt.figure()
    ax = fig.add_subplot(2, 1, 1)
    ax.bar(range(len(monthly_VAR)), monthly_VAR.values(), align="center")
    ax.set_xticks(range(len(monthly_VAR)))
    ax.set_xticklabels(monthly_VAR.keys(), rotation=45)
    ax.set_ylabel("MVAR")
    ax.set_title(station_name + " Monthly Out of Desired Operational Band (Median)", fontweight='bold')
    ax.set_rasterized(True)

    ax = fig.add_subplot(2, 1, 2)
    ax.bar(range(len(monthly_Voltage)), monthly_Voltage.values(), align="center")
    ax.set_xticks(range(len(monthly_Voltage)))
    ax.set_xticklabels(monthly_Voltage.keys(), rotation=45)
    ax.set_xlabel("Month of the Year")
    ax.set_ylabel("kV")
    ax.set_rasterized(True)

    plt.tight_layout()
    # file.savefig(bbox_inches='tight')

    fig = plt.figure()
    ax = fig.add_subplot(2, 1, 1)
    ax.bar(range(len(daily_VAR)), daily_VAR.values(), align="center")
    ax.set_xticks(range(len(daily_VAR)))
    ax.set_xticklabels(daily_VAR.keys(), rotation=45)
    ax.set_ylabel("MVAR")
    ax.set_title(station_name + " Daily Out of Desired Operational Band (Median)", fontweight='bold')
    ax.set_rasterized(True)

    ax = fig.add_subplot(2, 1, 2)
    ax.bar(range(len(daily_Voltage)), daily_Voltage.values(), align="center")
    ax.set_xticks(range(len(daily_Voltage)))
    ax.set_xticklabels(daily_Voltage.keys(), rotation=45)
    ax.set_xlabel("Day of the Week")
    ax.set_ylabel("kV")
    ax.set_rasterized(True)

    plt.tight_layout()
    # file.savefig(bbox_inches='tight')

    fig = plt.figure()
    ax = fig.add_subplot(2, 1, 1)
    ax.bar(range(len(hourly_VAR)), hourly_VAR.values(), align="center")
    ax.set_xticks(range(len(hourly_VAR)))
    ax.set_xticklabels(hourly_VAR.keys(), rotation=45)
    ax.set_ylabel("MVAR")
    ax.set_title(station_name + " Hourly Out of Desired Operational Band (Median)", fontweight='bold')
    ax.set_rasterized(True)

    ax = fig.add_subplot(2, 1, 2)
    ax.bar(range(len(hourly_Voltage)), hourly_Voltage.values(), align="center")
    ax.set_xticks(range(len(hourly_Voltage)))
    ax.set_xticklabels(hourly_Voltage.keys(), rotation=45)
    ax.set_xlabel("Hour of the Day")
    ax.set_ylabel("kV")
    ax.set_rasterized(True)

    plt.tight_layout()
    # file.savefig(bbox_inches='tight')

    fig = plt.figure()
    ax = fig.add_subplot(2, 1, 1)
    ax.bar(range(len(Xmonthly_VAR)), Xmonthly_VAR.values(), align="center")
    ax.set_xticks(range(len(Xmonthly_VAR)))
    ax.set_xticklabels(Xmonthly_VAR.keys(), rotation=45)
    ax.set_ylabel("MVAR")
    ax.set_title(station_name + " Monthly Out of Desired Operational Band (Median) (Excluding non-violations)", fontweight='bold')
    ax.set_rasterized(True)

    ax = fig.add_subplot(2, 1, 2)
    ax.bar(range(len(Xmonthly_Voltage)), Xmonthly_Voltage.values(), align="center")
    ax.set_xticks(range(len(Xmonthly_Voltage)))
    ax.set_xticklabels(Xmonthly_Voltage.keys(), rotation=45)
    ax.set_xlabel("Month of the Year")
    ax.set_ylabel("kV")
    ax.set_rasterized(True)

    plt.tight_layout()
    # file.savefig(bbox_inches='tight')

    fig = plt.figure()
    ax = fig.add_subplot(2, 1, 1)
    ax.bar(range(len(Xdaily_VAR)), Xdaily_VAR.values(), align="center")
    ax.set_xticks(range(len(Xdaily_VAR)))
    ax.set_xticklabels(Xdaily_VAR.keys(), rotation=45)
    ax.set_ylabel("MVAR")
    ax.set_title(station_name + " Daily Out of Desired Operational Band (Median) (Excluding non-violations)", fontweight='bold')
    ax.set_rasterized(True)

    ax = fig.add_subplot(2, 1, 2)
    ax.bar(range(len(Xdaily_Voltage)), Xdaily_Voltage.values(), align="center")
    ax.set_xticks(range(len(Xdaily_Voltage)))
    ax.set_xticklabels(Xdaily_Voltage.keys(), rotation=45)
    ax.set_xlabel("Day of the Week")
    ax.set_ylabel("kV")
    ax.set_rasterized(True)

    plt.tight_layout()
    # file.savefig(bbox_inches='tight')

    fig = plt.figure()
    ax = fig.add_subplot(2, 1, 1)
    ax.bar(range(len(Xhourly_VAR)), Xhourly_VAR.values(), align="center")
    ax.set_xticks(range(len(Xhourly_VAR)))
    ax.set_xticklabels(Xhourly_VAR.keys(), rotation=45)
    ax.set_ylabel("MVAR")
    ax.set_title(station_name + " Hourly Out of Desired Operational Band (Median) (Excluding non-violations)", fontweight='bold')
    ax.set_rasterized(True)

    ax = fig.add_subplot(2, 1, 2)
    ax.bar(range(len(Xhourly_Voltage)), Xhourly_Voltage.values(), align="center")
    ax.set_xticks(range(len(Xhourly_Voltage)))
    ax.set_xticklabels(Xhourly_Voltage.keys(), rotation=45)
    ax.set_xlabel("Hour of the Day")
    ax.set_ylabel("kV")
    ax.set_rasterized(True)

    plt.tight_layout()
    # file.savefig(bbox_inches='tight')

def generatePlots(station_name, filepath, VAR_bounds, Volt_bounds, bound_type):
    # Read in the data file
    table = pd.read_excel(filepath, sheetname=2)
    num_points = table.shape[0]

    # Store the dates and times in an array to index the plots
    datetimes = [pd.to_datetime(table.at[i, "Date"]) for i in range(num_points)]
    for i in range(num_points):
        datetimes[i] = datetimes[i].replace(hour=table.at[i, "Time"].hour, minute=table.at[i, "Time"].minute)

    # Store the MW, MVAR, and Voltage data in an array
    MW = [table.at[i, "MW"] for i in range(num_points)]
    MVAR = [table.at[i, "MVAR"] for i in range(num_points)]
    Voltage = [table.at[i, "kV"] for i in range(num_points)]

    # Create a pdf in which the plots will be saved
    # pdf = PdfPages(os.path.join(os.path.dirname(filepath), station_name + ".pdf"))
    pdf = None

    # Plotting functions
    plotMW(station_name, datetimes, MW, pdf)
    low_VAR, high_VAR = plotMVAR(station_name, datetimes, MW, MVAR, VAR_bounds, pdf)
    low_Volt, high_Volt = plotVoltage(station_name, datetimes, MW, Voltage, Volt_bounds, bound_type, pdf)
    plotBreakdown(station_name, datetimes, MVAR, Voltage, low_VAR, high_VAR, low_Volt, high_Volt, pdf)

    # pdf.close()