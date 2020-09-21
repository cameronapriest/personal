# Python grapher for excel workbooks

I created this for my dad, who needed a more reliable way to graph
multiple data sets on one graph. Excel was not creating a graph
with line plots for all provided data sets.

The user will need to download two libraries: xlrd and matplotlib.
To do this, the user will need to use pip install or pip3 install,
depending on which version of python the user has installed:

    Python 3.x.x -> pip3
    Older than Python 3 -> pip

For example: 

    pip3 install xlrd
    pip3 install matplotlib
    
The user runs the program using the command:

    python3 grapher.py workbook.xlsx
    
The user then enters the number of data sets when prompted and
chooses whether a line or scatter plot is desired.
