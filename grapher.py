import sys
import xlrd
import matplotlib.pyplot as plt


red = "\x1B[31m"
green = "\x1B[32m"
blue = "\x1B[34m"
yellow = "\x1B[33m"
purple = "\x1B[35m"
end = "\x1B[0m"


def main():
    # Error handling in case the user does not enter a workbook
    if len(sys.argv) != 2:
        print(red, end="")
        print("\nEnter the workbook as a command line argument:", end)
        print(green, end="")
        print("python3 grapher.py workbook.xlsx\n", end)
        return

    # assign the argument to the workbook
    workbook = sys.argv[1]

    # open first sheet of the workbook
    wb = xlrd.open_workbook(workbook)
    sheet = wb.sheet_by_index(0)

    xlabel = sheet.cell_value(1, 0)
    ylabel = sheet.cell_value(0, 1)

    print("\nUnits:")
    print("x:", end="")
    print(blue, xlabel, end)
    print("y:", end ="")
    print(blue, ylabel, end)

    # ask user for number of data sets
    sets = getSets()

    # error handling for invalid number of sets
    while sets < 1:
        print(red, end="")
        print("Please enter a positive number of data sets.", end)
        sets = getSets()

    print("")

    # boolean to specify scatter or line plot (ask user)
    scatter = scatterOrLine()

    print(green, end="")
    print("\nGraphing %d data sets...\n" % sets, end)

    # list of line names
    labels = []

    # collect line names through iteration of columns
    c = 1
    while c <= sets:
        labels.append(sheet.cell_value(1, c))
        c += 1

    # list of times (x axis)
    times = []

    # collect times from spreadsheet through iteration of rows
    row = 2
    while str(sheet.cell_value(row, 0)) != "":
        times.append(sheet.cell_value(row, 0))
        row += 1

    # specify number of total rows
    rows = row-2

    # create a dictionary of coordinates to plot
    data = {}
    for label in range(len(labels)):
        data[labels[label]] = []

    # iterate through data sets and fill dictionary
    col = 1
    for column in range(sets):
        row = 2
        for r in range(rows):
            if str(sheet.cell_value(row, col)) != "":
                data[labels[col-1]].append((times[row-2], sheet.cell_value(row, col)))
            row += 1
        col += 1

    # reorganize coordinates to make easier to plot
    num = 0
    for set in range(sets):
        x = []
        y = []
        i = 0
        for coord in range(len(data[labels[num]])):
            x.append(data[labels[num]][i][0])
            y.append(data[labels[num]][i][1])
            i += 1

        # scatter or line plot
        if scatter:
            plt.scatter(x, y, label=labels[num])
        else:
            plt.plot(x, y, label=labels[num])
        num += 1

    # axis labels, title, legend, & display graph
    plt.xlabel(xlabel)
    plt.ylabel(ylabel)
    plt.title("BG pressures (1e-7 torr) vs time (minutes)")
    plt.legend()
    plt.show()


def getSets():
    try:
        datasets = int(input("\nHow many data sets are there? "))
        return datasets
    except ValueError:
        print(red, end="")
        print("Please enter an integer.", end)
        getSets()


def scatterOrLine():
    scatter = input("Would you like a scatter plot (s) or line plot (l)? ")
    if scatter == "s":
        return True
    elif scatter == "l":
        return False
    else:
        print(red, end="")
        print("Please enter either s or l.", end)
        scatterOrLine()


if __name__ == "__main__":
    main()