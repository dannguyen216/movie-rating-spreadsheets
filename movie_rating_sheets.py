import openpyxl
import sys
from openpyxl.styles import Font, PatternFill, Alignment


# Reads an input file given by the user in a specific format and
# returns a list of movies along with their corresponding rating
# and release date.
def read_movie_input(input_file):
    movie_data_list = []
    with open(input_file, 'r') as file:
        for line in file.readlines():
            # Only append movie data if the line if the line has characters
            # other than whitespaces.
            if(line.strip()):
                movie_data = [s.strip() for s in line.split(';;;')]
                if(len(movie_data) == 3):
                    movie_data_list.append(movie_data)


    # print(movie_data_list)
    return movie_data_list


def write_column_names(worksheet):
    light_red_cell_fill = PatternFill(fill_type='solid', start_color='FF9999', end_color='FF9999')

    worksheet['A1'].font = Font(size=20, bold=True, underline='single')
    worksheet['A1'] = "Movie Title"
    worksheet['A1'].alignment = Alignment(horizontal='center')
    worksheet['A1'].fill = light_red_cell_fill
    worksheet.column_dimensions['A'].width = 50

    worksheet['B1'].font = Font(size=20, bold=True, underline='single')
    worksheet['B1'] = "Rating"
    worksheet['B1'].alignment = Alignment(horizontal='center')
    worksheet['B1'].fill = light_red_cell_fill
    worksheet.column_dimensions['B'].width = 25

    worksheet['C1'].font = Font(size=20, bold=True, underline='single')
    worksheet['C1'] = "Release Date"
    worksheet['C1'].alignment = Alignment(horizontal='center')
    worksheet['C1'].fill = light_red_cell_fill
    worksheet.column_dimensions['C'].width = 25
    
    return


def write_movie_data_to_spreadsheet(worksheet, movie_data_list):
    row_num = 2
    for movie_data in movie_data_list:
        movie_title_cell = 'A' + str(row_num)
    print(len(movie_data_list))

    return


def main():
    if len(sys.argv) != 2:
        print("Usage: python movie_rating_sheets.py input_file")
        return

    input_file = sys.argv[1]
    file_name = input_file.split('.')[0]

    movie_data_list = read_movie_input(input_file)

    try:
        print('Loading Movie_Ratings workbook...')
        workbook = openpyxl.load_workbook('Movie_Ratings.xlsx')
        if file_name in workbook.sheetnames:
            print('Spreadsheet of name \"%s\" already in workbook.' % file_name)
            print('Exiting...')
            return
        else:
            workbook.create_sheet(title=file_name)

    except FileNotFoundError:
        print('Movie ratings workbook not found in current directory...')
        print('Creating workbook file Movie_Ratings.xlsx...')
        workbook = openpyxl.Workbook()
        workbook.active.title = file_name
        
    worksheet = workbook.get_sheet_by_name(file_name)
    write_column_names(worksheet)
    print('Writing data to worksheet %s...' % worksheet.title)
    write_movie_data_to_spreadsheet(worksheet, movie_data_list)
    
    '''
    wb = openpyxl.load_workbook('example.xlsx')
    print(wb.get_sheet_names())
    sheet_one = wb.get_sheet_by_name('Sheet1')
    
    for row in sheet_one.iter_rows():
        for cell in row:
            if(cell.value):
                print(cell.value)
    
    for cell in list(sheet_one.columns)[1]:
        print(cell.value)

    print('max_column = %d' % sheet_one.max_column)
    print('max_row = %d' % sheet_one.max_row)
    '''

    print('Saving changes to workbook...')
    try:
        workbook.save('Movie_Ratings.xlsx')
        print('\nChanges saved successfully.')
    except OSError:
        print('\nFailed to save changes.')
        print('Workbook Movie_Ratings.xlsx is busy.')
        print('Confirm that the workbook is closed before running the script.')
        print('Exiting...')



if __name__ == '__main__':
    main()