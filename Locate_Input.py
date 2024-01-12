import openpyxl
import pandas as pd


def Locate_Input(sheet, color_input, start_row, end_row, start_col, end_col,panel_name):
    """Converts a list of lists of colors into a pandas DataFrame"""
    # Read the cell test input colors into a numpy array
    # True if the cell is colored, False if not

    df=pd.DataFrame(columns=['PANEL', 'abscisse', 'ordinate', 'coord'])
    
    for row in sheet.iter_rows(min_row=start_row, max_row=end_row, min_col=start_col, max_col=end_col):
        for cell in row:       
            if cell is not None:
                if cell.fill.start_color.index == color_input:
                    new_row={'PANEL':panel_name,'abscisse': cell.column, 'ordinate': cell.row,'coord':cell.coordinate}
                    df.loc[len(df)] = new_row    
    return df
