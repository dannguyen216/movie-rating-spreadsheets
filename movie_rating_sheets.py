import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment


# Reads an input file given by the user in a specific format and
# returns a list of movies along with their corresponding rating
# and release date.
def read_movie_input(input_file):
    movie_data_list = []
    with open(input_file, 'r') as file:
        for line in file.readlines():
            movie_data = [s.strip() for s in line.split(';;;')]
            movie_data_list.append(movie_data)

    print(movie_data_list)
    return movie_data_list


def write_movie_data_to_spreadsheet(workbook, file_name, movie_data_list):
    # The name of the file will be the name of the worksheet
    '''
    file_name = input_file.split('.')[0]
    wb = openpyxl.Workbook()
    wb.create_sheet(title=file_name)
    movie_sheet = get_sheet_by_name(file_name)
    '''
    try:
        worksheet = workbook.get_sheet_by_name(file_name)

    except:
        workbook.active.title = file_name
        worksheet = workbook.get_sheet_by_name(file_name)
        worksheet['A1'].font = Font(size=20, bold=True, underline='single')
        worksheet['A1'] = "Movie Title"
        worksheet['A1'].alignment = Alignment(horizontal='center')
        worksheet.column_dimensions['A'].width = 50

        worksheet['B1'].font = Font(size=20, bold=True, underline='single')
        worksheet['B1'] = "Rating"
        worksheet['B1'].alignment = Alignment(horizontal='center')
        worksheet.column_dimensions['B'].width = 25

        worksheet['C1'].font = Font(size=20, bold=True, underline='single')
        worksheet['C1'] = "Release Date"
        worksheet['C1'].alignment = Alignment(horizontal='center')
        worksheet.column_dimensions['C'].width = 25


    return


def main():
    input_file = 'example_input.txt'
    file_name = input_file.split('.')[0]

    movie_data_list = read_movie_input(input_file)

    try:
        print('Loading Movie_Ratings workbook...')
        wb = openpyxl.load_workbook('Movie_Ratings.xlsx')
    except FileNotFoundError:
        print('Movie ratings workbook not found in current directory...')
        print('Creating workbook file Movie_Ratings.xlsx...')
        wb = openpyxl.Workbook()
        
        '''
        wb.active.title = file_name
        sheet = wb.get_sheet_by_name(file_name)
        sheet['A1'].font = Font(size=20, bold=True, underline='single')
        sheet['A1'] = "Movie Title"
        sheet['A1'].alignment = Alignment(horizontal='center')
        sheet.column_dimensions['A'].width = 50
        '''
    write_movie_data_to_spreadsheet(wb, file_name, movie_data_list)
    
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
    wb.save('Movie_Ratings.xlsx')



if __name__ == '__main__':
    main()