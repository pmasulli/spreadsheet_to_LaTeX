# -*- coding: utf-8 -*-

# 
# Spreadsheet to LaTeX - import data in your LaTeX code automatically
# v. 0.1-201607
# Copyright (C) 2016 Paolo Masulli
# 
# This program is free software; you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation; either version 3 of the License, or
# (at your option) any later version.
# 
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
# 
# You should have received a copy of the GNU General Public License
# along with this program; if not, write to the Free Software Foundation,
# Inc., 51 Franklin Street, Fifth Floor, Boston, MA 02110-1301  USA



from xlrd import open_workbook
import re
import sys
import os.path

default_sheet_name = "Sheet1" # insert here the default sheet name in the xlsx file

def cell_coords(cell_ref):
    pos = re.search("[0-9]", cell_ref).start()
    col = cell_ref[:pos].upper()
    if len(col) == 1:
        col = ord(col) - ord('A')
    elif len(col) == 2:
        col = 26 * (1 + ord(col[0]) - ord('A')) + ord(col[1]) - ord('A')
    row = int(cell_ref[pos:]) - 1
    return row, col
    

if len(sys.argv) < 3:
    print "Usage: ", sys.argv[0], "datafile.xlsx templatefile"
    exit(-1)
    

xlsx_file = sys.argv[1]
template_file = sys.argv[2]
latex_file = template_file + ".tex"

if not (os.path.isfile(xlsx_file) and os.path.isfile(template_file)):
    print "Cannot find the data file or the template file."
    exit(-1)


print "Writing output in %s." % latex_file,

tf = open(template_file, 'r')
lf = open(latex_file, 'w')

wb = open_workbook(xlsx_file)

p = re.compile("(<data>(.+?)</data>)")

for line in tf:
    line_rep = line
    for f in p.findall(line):
        sheet_name = None
        coords = f[1]
        if ":" in coords:
            sheet_name = coords[:coords.find(":")]
            coords = coords[(coords.find(":")+1):]
        row, col = cell_coords(coords)
        if not sheet_name:
            sheet_name = default_sheet_name
            
        sheet = wb.sheet_by_name(sheet_name)
        value =  str(sheet.cell(row, col).value).replace("_",r'\_')
        line_rep = line_rep.replace(f[0], value).encode('utf-8')
        print ".",
    lf.write(line_rep)
    

tf.close()
lf.close()

print
print "Done."