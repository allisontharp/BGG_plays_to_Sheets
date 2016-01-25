import gspread

global headers

def write_cell_colstring (colstring, row, data, sheet):

    headers = sheet.row_values(1)
    col = headers.index(colstring)+1
    sheet.update_cell(row,col,data)
    
def write_cels (lastrow,lastcol,data,sheet):
    lastcol = chr(55%26 - 1 + ord('a'))
    cellrange = 'a' + str(lastrow) + ':'
    
def col(colstring,headers):
        return headers.index(colstring)
