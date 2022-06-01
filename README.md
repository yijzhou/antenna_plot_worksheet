# antenna_plot_worksheet
Excel worksheet to plot antenna parameters

ver: 0.1.0

Yijun Zhou
yijun.zhou@microsoft.com
yijun@icloud.com

This software uses Excel as the front end for data input and plot setting 
and plot the antenna parameters (S11, S21, Efficiency, Pattern) by Python scrips and xlwings

Useage:
1) make sure the *.xlsm file and *.py has the same name and put at the same directory
2) input S1P, CSV file name and path to the index spreadsheet
3) choose the spreadsheet you want to plot and select the data file source form the drop down list
4) adjust plot setting (grey area should be not modified)

Dependency:
Python
Matplotlib
Numpy
SKRF
xlwings


Installation:
1) Copy the *.xlsm and *.py files into the same directory
2) Install the dependency packages. The easiest way is to install visual studio code and install python from vscode.
3) Install xlwings:
	pip install xlwings
	xlwings addin install
4) In excel, change "Home" -> "Options" -> "Trust center" -> "Trust Center Settings..." -> "Macro Settings", turn on
	Enable VBA macros, Enable Excel 4.0 macros when VBA macros are enabled, Trust access to the VBA project object model.
5) Sometimes, I see windows security center is blocking the excel VBA. Need to manually load the xlwings.xlam.
	The file is located at C:\Users\***\AppData\Roaming\Microsoft\Excel\XLSTART\xlwings.xlam

Known issues:
1) xlwings does not support mac. So the UDFs (user defined functions) will not work in mac.
2) Pattern plot (3D polar) has issues of colormap.
3) Change subplot from 2 to 1 or vice versa will affect plot dimension.
	Workaround: delete the plot and replot by update the formula again.



