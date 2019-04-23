import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, colors
import sys


# Global variables for Worksheet styles
# Cell colors
LIGHT_RED_FILL = PatternFill(fill_type='solid',
                             start_color='FF9999',
                             end_color='FF9999')
LIGHT_BLUE_FILL = PatternFill(fill_type='solid',
                              start_color='CCCCFF',
                              end_color='CCCCFF')
RED_FILL = PatternFill(fill_type='solid',
                       start_color='FF6666',
                       end_color='FF6666')
ORANGE_FILL = PatternFill(fill_type='solid',
                          start_color='FFB266',
                          end_color='FFB266')
YELLOW_FILL = PatternFill(fill_type='solid',
                          start_color='FFFF66',
                          end_color='FFFF66')
GREEN_FILL = PatternFill(fill_type='solid',
                         start_color='66FF66',
                         end_color='66FF66')
BLACK_FILL = PatternFill(fill_type='solid',
                         start_color='000000',
                         end_color='000000')

# Cell value alignment
CENTER_ALIGN = Alignment(horizontal='center')

# Font Styles
COLUMN_TITLE_FONT = Font(size=20, bold=True, underline='single')
BOLD_FONT = Font(bold=True)
AVERAGE_FONT = Font(bold=True,color=colors.WHITE)

# Border Styles
CELL_BORDER = Border(left=Side(border_style='thick'),
                     right=Side(border_style='thick'),
                     top=Side(border_style='thick'),
                     bottom=Side(border_style='thick'))


# Function that reads an input file given by the user in a specific
# specific format and returns a list of movies along with their 
# corresponding rating and release date.
def read_movie_input(input_file):
    movie_data_list = []
    with open(input_file, 'r') as file:
        for line in file.readlines():
            # Only append movie data if the line if the line has characters
            # other than whitespaces.
            if line.strip():
                movie_data = [s.strip() for s in line.split(';;;')]
                if len(movie_data) == 3:
                    movie_data_list.append(movie_data)

    # print(movie_data_list)
    return movie_data_list


# Function used to write the titles for each column on the worksheet.
# The three columns are:
#    A.) Movie Title
#    B.) Rating (A review score for the movie from 1 to 5)
#    C.) Release Date of the movie
def write_column_names(worksheet):
    worksheet['A1'] = "Movie Title"
    worksheet['A1'].font = COLUMN_TITLE_FONT
    worksheet['A1'].alignment = CENTER_ALIGN
    worksheet['A1'].fill = LIGHT_RED_FILL
    worksheet['A1'].border = CELL_BORDER
    worksheet.column_dimensions['A'].width = 50

    worksheet['B1'] = "Rating"
    worksheet['B1'].font = COLUMN_TITLE_FONT
    worksheet['B1'].alignment = CENTER_ALIGN
    worksheet['B1'].fill = LIGHT_RED_FILL
    worksheet['B1'].border = CELL_BORDER
    worksheet.column_dimensions['B'].width = 25

    worksheet['C1'] = "Release Date"
    worksheet['C1'].font = COLUMN_TITLE_FONT
    worksheet['C1'].alignment = CENTER_ALIGN
    worksheet['C1'].fill = LIGHT_RED_FILL
    worksheet['C1'].border = CELL_BORDER
    worksheet.column_dimensions['C'].width = 25

    return


# Function that takes the movie data read from the input file
# and writes the data to the excel spreadsheet given as
# a parameter.
def write_movie_data_to_spreadsheet(worksheet, movie_data_list):
    row_num = 2
    for movie_data in movie_data_list:
        movie_title = movie_data[0]
        movie_title_cell = 'A' + str(row_num)
        worksheet[movie_title_cell] = movie_title
        worksheet[movie_title_cell].font = BOLD_FONT
        worksheet[movie_title_cell].alignment = CENTER_ALIGN
        worksheet[movie_title_cell].fill = LIGHT_BLUE_FILL
        worksheet[movie_title_cell].border = CELL_BORDER

        rating = movie_data[1]
        rating_cell = 'B' + str(row_num)
        worksheet[rating_cell] = '%s / 5' % rating
        worksheet[rating_cell].font = BOLD_FONT
        worksheet[rating_cell].alignment = CENTER_ALIGN
        worksheet[rating_cell].fill = get_rating_color(float(rating))
        worksheet[rating_cell].border = CELL_BORDER

        release_date = movie_data[2]
        date_cell = 'C' + str(row_num)
        worksheet[date_cell] = release_date
        worksheet[date_cell].font = BOLD_FONT
        worksheet[date_cell].alignment = CENTER_ALIGN
        worksheet[date_cell].fill = LIGHT_BLUE_FILL
        worksheet[date_cell].border = CELL_BORDER

        row_num += 1

    # print(len(movie_data_list))
    # print('num rows: %d' % row_num)
    return


# A helper function to the write_movie_data_to_spreadsheet
# function that determines the color of each cell in the
# ratings column, depending on what range the rating
# falls under. The ratings are as follows:
#    rating < 2: Red cell color (Bad movie)
#    2 < rating < 3: Orange cell color (Mediocre Movie)
#    rating = 3: Yellow cell color (Decent Movie)
#    rating > 3: Green cell color (Good Movie)
def get_rating_color(rating):
    if rating < 2:
        return RED_FILL
    elif rating < 3:
        return ORANGE_FILL
    elif rating == 3:
        return YELLOW_FILL
    else:
        return GREEN_FILL


def write_average_to_spreadsheet(worksheet):
    rating_value_string = 'LEFT(B2, SEARCH(\"/\", B2) - 1)'
    for cell in worksheet['B'][2:]:
        rating_value_string += ','
        cell_string = cell.column + str(cell.row)
        rating_value_string += 'LEFT({}, SEARCH(\"/\", {}) - 1)'.format(cell_string,
                                                                        cell_string)

    average_row = str(worksheet.max_row + 1)

    label_cell = 'A' + average_row
    worksheet[label_cell] = 'AVERAGE RATING'
    worksheet[label_cell].font = AVERAGE_FONT
    worksheet[label_cell].alignment = CENTER_ALIGN
    worksheet[label_cell].fill = BLACK_FILL

    average_cell = 'B' + average_row
    worksheet[average_cell] = '=ROUND(AVERAGE({}), 2)'.format(rating_value_string)
    worksheet[average_cell].font = AVERAGE_FONT
    worksheet[average_cell].alignment = CENTER_ALIGN
    worksheet[average_cell].fill = BLACK_FILL
    return


def main():
    if len(sys.argv) != 2:
        print("Usage: python movie_rating_sheets.py input_file")
        return

    input_file = sys.argv[1]
    file_name = input_file.split('.')[0]

    movie_data_list = read_movie_input(input_file)

    if len(movie_data_list) > 0:
        try:
            print('Loading Movie_Ratings workbook...')
            workbook = openpyxl.load_workbook('Movie_Ratings.xlsx')
            if file_name in workbook.sheetnames:
                print('Spreadsheet of name \"%s\" already in workbook.' %
                      file_name)
                print('Exiting...')
                return
            else:
                workbook.create_sheet(title=file_name)

        except FileNotFoundError:
            print('\nMovie ratings workbook not found in current directory.')
            print('Creating workbook file Movie_Ratings.xlsx...')
            workbook = openpyxl.Workbook()
            workbook.active.title = file_name

        worksheet = workbook.get_sheet_by_name(file_name)
        write_column_names(worksheet)
        print('\nWriting data to worksheet %s...' % worksheet.title)
        write_movie_data_to_spreadsheet(worksheet, movie_data_list)

        write_average_to_spreadsheet(worksheet)

        print('Saving changes to workbook...')
        try:
            workbook.save('Movie_Ratings.xlsx')
            print('\nChanges saved successfully.')
        except OSError:
            print('\nFailed to save changes.')
            print('Workbook Movie_Ratings.xlsx is busy.')
            print('Confirm that the workbook is closed before running the script.')
            print('Exiting...')

    else:
        print('No movie data obtained from input file %s' % input_file)
        print('Exiting...')


if __name__ == '__main__':
    main()
