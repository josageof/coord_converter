# -*- coding: utf-8 -*-
"""
Created on Thu May 19 12:21:07 2022

@author: Josa -- josageof@gmail.com (support)
"""

# from IPython import get_ipython
# get_ipython().run_line_magic('reset', '-f')


import shutil
import psutil
import xlwings as xw
from osgeo import osr


input_esri_prj = "test_projection_file.prj"   #!!! input projection
output_epsg_code = 4326                       #!!! output projection


infile_data = 'test_excel_file_Mercator.xlsx'


sheet_order = 0
first_x_in = "D4"   ## Easthing   --> Long
first_y_in = "E4"   ## Northing   --> Lat


#%% definitions 
    
def reproj_xy_list(xl_in, yl_in, p_in, p_out):
    """
    Convert x,y coordenates lists to different projection.
    p.s: This works with OSR projections by GDAL.

    Args:
        xl_in (list): x coordenates input.
        yl_in (list): y coordenates input.
        p_in (projection): projection input.
        p_out (projection): projection output.

    Returns:
        xl_out (list): x coordenates output.
        yl_out (list): y coordenates output.

    """
    trans = osr.CoordinateTransformation(p_in, p_out)
    xl_out=[]; yl_out=[]
    for i in range(0, len(xl_in)):
        x, y, z = trans.TransformPoint(xl_in[i], yl_in[i])
        xl_out.append(x)
        yl_out.append(y)
    return xl_out, yl_out


def get_vlist_from_sheet(file, sheet_num, first_cell):
    """
    Get vertical list from Excel column.

    Args:
        file (string): input file.
        sheet_num (int): sheet number.
        first_cell (string): first cell (like excel).

    Returns:
        TYPE: python list.

    """
    app = xw.App(visible=False)
    book = xw.Book(file)
    sheet = book.sheets[sheet_num]
    last_row = sheet.range(first_cell).end('down').row
    last_col = sheet.range(first_cell).end('down').column
    vlist = sheet.range(first_cell, (last_row, last_col)).value
    book.close()
    app.quit()
    return vlist


def write_list_to_sheet_col(lst, file, sheet_num, first_cell):
    """
    Write vertical list to Excel column.

    Args:
        lst (list): python list.
        file (string): output file.
        sheet_num (int): sheet number.
        first_cell (string): first cell (like excel).

    """
    app = xw.App(visible=False)
    book = xw.Book(file)
    sheet = book.sheets[sheet_num]
    sheet.range(first_cell).options(transpose=True).value = lst
    book.save(file)
    book.close()
    app.quit()
    
      
def deg_to_dm(deg, pretty_print=None, ndp=4):
    """
    Convert from decimal degrees to degrees, minutes.
    """
    d = int(deg)
    m = (deg-d) * 60

    if deg < 0:
        d = -d
        m = -m

    if pretty_print:
        if pretty_print=='latitude':
            hemi = 'N' if deg>=0 else 'S'
        elif pretty_print=='longitude':
            hemi = 'E' if deg>=0 else 'W'
        else:
            hemi = '?'
        return '{d:d}° {m:.{ndp:d}f}′{hemi:1s}'.format(
                    d=abs(d), m=m, hemi=hemi, ndp=ndp)
    return d, m


def kill_excel():
    for proc in psutil.process_iter():
        if proc.name() == "EXCEL.EXE":
            proc.kill()



#%% setting projections by gdal modules

# read input projection from esri file format
with open(input_esri_prj, "r") as f:
    string = str(f.readlines()[0])
prj_in = osr.SpatialReference()
prj_in.ImportFromWkt(string)


# reading output projection using epsg code
prj_out = osr.SpatialReference()
prj_out.ImportFromEPSG(output_epsg_code)



#%% converting coordenates
  
## reading by xlswings (melhor)
# get cols form sheet_in to lists
x_in = get_vlist_from_sheet(infile_data, sheet_order, first_x_in)
y_in = get_vlist_from_sheet(infile_data, sheet_order, first_y_in)


## reproject coordentes
lat, long = reproj_xy_list(y_in, x_in, prj_in, prj_out)


# convert to degree, minutes
lat = [deg_to_dm(item, pretty_print='latitude') for item in lat]
long = [deg_to_dm(item, pretty_print='longitude') for item in long]


## writing by xlswings (melhor)
# make a output file copy
if "Mercator" in infile_data:
    outfile_data = infile_data.replace("Mercator", "LatLong")
else:
    outfile_data = infile_data.split(".")[0]+"_LatLong.xlsx"
shutil.copy(infile_data, outfile_data)

# write lists to excel file
write_list_to_sheet_col(lat, outfile_data, sheet_order, first_x_in)
write_list_to_sheet_col(long, outfile_data, sheet_order, first_y_in)


kill_excel()
