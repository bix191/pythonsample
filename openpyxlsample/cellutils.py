from openpyxl import utils

def getMergedCellValue(sheet,cell):
    cellidx=cell.coordinate
    for range in sheet.merged_cells.ranges:
        merged_cells = list(utils.rows_from_range(str(range)))
        for row in merged_cells:
            if cellidx in row:
                return sheet[merged_cells[0][0]].value
    return cell.value
