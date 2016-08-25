#Spreadsheet to LaTeX - import data in your LaTeX code automatically
v 0.1-201607
Copyright (C) 2016 Paolo Masulli

This program is free software; you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation; either version 3 of the License, or
(at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with this program; if not, write to the Free Software Foundation,
Inc., 51 Franklin Street, Fifth Floor, Boston, MA 02110-1301  USA


Usage:


Assume that you have your data contained in the file data.xslx and that you want to include those data in document.tex. Then prepare a LaTeX template where, instead of the data, you include tags of the form:

<data>SheetName:A2</data>

where SheetName is the name of the sheet in the Excel file and A2 are the coordinates of some cell. Note that you can specify a default sheet name in the Python script at line 8 and then just use tags of the form:

<data>B3</data>

to include the data in the cell B3 of the default sheet.
The name of the default sheet is specified in the Python script itself, for the time being.


Once you have the template, just run the script to generate the LaTeX file including the data from the spreadsheet:

$ python spreadsheet_to_latex.py data.xlsx document_template.txt

and then compile the LaTeX file normally:

$ pdflatex document_template.txt.tex