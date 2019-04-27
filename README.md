# Movie Data-To-Spreadsheet Converter

## Summary
The Python program **movie_rating_sheets.py** takes an input file of an expected format as a command line argument and reads the file before writing its data to an Excel spreadsheet of the same name using the OpenPyXL module. All spreadsheets written by **movie_rating_sheets.py** will be inserted in a workbook file called **Movie_Ratings.xlsx**.

In addition to writing the information to a spreadsheet, the average rating score will be calculated and written to a cell under the rest of the data.

The file was worked on with Python 3.4.5 and will not work with Python 2 in its current state.

## Goals of the Project
After introducing myself to the OpenPyXL module, I wanted some experience implementing it in a program. I also wanted to keep track of the movies that I watch per year in a spreadsheet and keep track of what I thought of movies after I watched them to see if they ever change over time.

## Input File Format
For the input file, the program expects a movie title, a rating (numerical review score) out of a maximum of 5 and the release date of the film. Each line of the input file should generally look like the following:

```
Movie Title ;;; Rating ;;; Release Date
```

> NOTE: Having two sets of three semicolons (;;;) is important, as that is how the line is parsed. Any line without two sets of three semicolons will be skipped over and will not be written to the spreadsheet.

Two example input files, **example_input.txt** and **example_input2.txt** are included to give examples on how an expected input file could look like, including lines that will be read by the program and written to the Excel spreadsheet as well as lines that will be skipped over (ex. A blank line).

## Directions to Run
With Python 3 installed, the script can be run with the following command:
```
python movie_rating_sheets.py input_file
```
where the input_file is located in the same directory as the **movie_rating_sheets.py**. If both Python 2 and Python 3 are installed when running, you can ensure that Python 3 will be used with the following:
```
python3 movie_rating_sheets.py input_file
```
Again, **movie_rating_sheets.py** is intended to work with Python 3 and will result in errors when running with Python 2 (ex. FileNotFoundError does not exist in Python 2 among other formatting differences).

> NOTE: The OpenPyXL module is very important to the functions in **movie_rating_sheets.py** and needs to be installed before running the program.
>
> Take a look at the [OpenPyXL Documentation](https://openpyxl.readthedocs.io/en/stable/) for installation instructions and additional details.

If the program runs as intended, A workbook file named **Movie_Ratings.xlsx** should be created in the current directory with a worksheet of the same name as the input file. The resulting worksheet should have three columns containing movie titles, ratings and release dates in separate rows (and edited with varoius worksheet styles). Additionally, the average of all the movie ratings in the worksheet should be calculated and displayed in column 'B' below the written data from the input file.

## Error Handling
Some edge cases have been accounted for in situations that prevent the data in the input_file from being written to the Excel spreadsheet:

- The input file cannot be opened in the current directory
- A worksheet of the same name as the input file already exists in **Movie_Ratings.xlsx**
- **Movie_Ratings.xlsx** is busy (ex. the worksheet is currently open) and cannot be saved to.
- No relevant data is able to be obtained from the input file

In these cases, a message will notify the user that the data cannot be written to the spreadsheet, and the program will exit without saving any changes.

## Known Issues / Areas for Improvement
The program is currently very simple in its implementation, and assumes that the input file is of a specified format. Not much testing has been done in terms of different input values, which leads to many limitations in its current state:

- The current version of the project does not do much at all to check for valid input. Specifically, the program assumes that the Rating will be in the range from 0 to 5, and does not currently enforce that the score lies within that range.
- As of right now, the date column is not very useful, and the value is just copied and pasted from the input file. Ideally, I would like to see if I could sort the movies in the spreadsheet by release date and see if any more features can be added from there.
- Using an input file of the same name as an existing worksheet would not write anything to the worksheet and will instead exit. I would like a less rigid implementation, allowing users to add movie ratings to existing worksheets among other features.
- The columns contaning the data have a set width that should work for most reasonable input lengths. On the other hand, some movie titles can be long, and I would like to see if I could automatically change the length of a column with respect to the element with the maximum length.

I plan on working on fixing these limitations along with adding more features that come to mind as long, but at the very least, I wanted to have this basic version of the project posted.

## File Descriptions
1. **movie_rating_sheets.py**: The main program that contains the code written for the project. The code is written using Python 3 and will not work with Python 2 in its current state. Additional comments in the code help to explain my approach.

2. **Movie_Ratings.xlsx**: The workbook file that is created by **movie_rating_sheets.py** to contain the worksheets that are written.

3. **example_input.txt** and **example_input2.txt**: As stated before, these two text files serve as example inputs that input valid data that will be parsed by **movie_rating_sheets.py** and written to a spreadsheet. Specifically, **example_input2.txt** contains some examples of invalid lines that the program will not store any data from and instead skip over.

4. **useful_links.txt**: Any StackOverflow posts and links that I found useful when working on the project are pasted in this text file for reference, and the solutions and topics explained in those links were very helpful in terms of figuring out how to implement some of the features.

5. **README**: That's this file! This contains descriptions of the project from a user standpoint along with my approach to its implementation. It also contains directions to run the program and limitations that have not been dealt with yet.
