SAM.gov Webscrape 1.0
23 October 2025
=====================
Maintainer: Nick Sunshine <nsunshine@vtti.vt.edu>
PM Contact: Ammie Jo Carter <acarter@vtti.vt.edu>

Installation & Organization
---------------------------

Webscrape is distributed as a simple .ZIP file. Download and unzip into the directory of your choosing. Using a directory in the root drive, such as "C:\webscrape", can be convenient.

Once unzipped, the directory will contain the following:
-> webscrape.exe
-> cfg\
--> keep_cols.txt
--> keywords.txt
--> sam_url.txt

The file webscrape.exe is the main program executable. This executable is packaged in such a way that downloading any Python package dependencies should not be required.
The cfg directory contains a number of human-readable text files that control the execution of the program:
sam_url.txt: This file contains a web URL that is a direct link to the SAM.gov opportunities .CSV file. If the link stops working for some reason, it is straightforward to edit this text file to correct the URL.
keywords.txt: This file contains a list of initialisms, words, and phrases. These are used to match the .CSV file data set inside of the "Description" column. Any rows that have a match are kept and all others are discarded. Matches are for whole words/phrases. Matches are case-insensitive, except for initialisms, which are case-sensitive for all capitals.
keep_cols.txt: This file contains a list of the column names in the .CSV file that will be kept and all others are discarded.

After running the executable (manually or via Windows Task Scheduler [see instructions below]), three more directories will be created and populated with files:
-> logs\
-> input\
-> output\

The logs directory will store text files with the program output (e.g., "2025-10-23_07-51-59_webscrape.log"). These can be examined to ensure the program is executing correctly, finding matches, and filtering the data set.
The input directory stores the latest downloaded .CSV file from the SAM.gov website. As this is typically a large file (about 200 MB), only the latest version is kept.
The output directory stores the filtered data set as a Microsoft Excel file. One file is kept per day as filenames are timestamped by date, such as "2025-10-23_07-51-59_SAMOpportunities.xlsx".

Configuration with Windows Task Scheduler
-----------------------------------------

1. Make note of the installation path of webscrape.exe (e.g., C:\webscrape\webscrape.exe).
2. Open Windows Task Scheduler. Look for a panel on the right side titled "Actions".
3. Click "Create Task..."
4. On the "General" tab, name the task such as "Webscrape Daily". Default settings for "Security options" should be OK.
5. On the "Triggers" tab, click the "New" button. Set up the frequency of the task. For this example, choose "Daily". Select the start time for a few minutes from now, and recur every "1" day. Click the "OK" button.
6. On the "Actions" tab, click the "New" button. Leave the action as "Start a program". Browse for the program/script and select the webscrape application noted in step 1. Click "OK".
7. Optional conditions and settings can be set on the "Conditions" and "Settings" tabs, but defaults should be OK.
8. Click "OK" to create the task. Note that the new task will not show up in the "Active Tasks" window until it runs for the first time.
9. When the time for the task to run arrives, a Windows Command Prompt will pop up and the Webscrape program will execute. Allow time for the program to complete, as the download of the CSV file, filtering, and conversion to Excel can take a while.
10. After Webscrape is done executing, check the installation directory for the logs, input, and output directories and expected contents.

Results Excel File Orientation
------------------------------
Opening the results Excel file may initial pop up a warning message as follows, or similar:

"We found a problem with some content in '2025-10-23_07-51-59_SAMOpportunities.xlsx'. Do you want us to try to recover as much as we can? If you trust the source of this workbook, click Yes."

This occurs due to some of the data from the SAM.gov .CSV file being interpreted as a formula in Excel. Those data get removed. Click Yes to continue.
A window pops up detailing repairs made to the Excel file. An XML log file can be clicked to examine more information.

Two sheets are present in the Excel file: "Keyword Matches" and "Keyword Match Counts".

"Keyword Matches" shows the filtered data, where rows were kept where the "Description" field matched one or more items in keywords.txt and columns are kept according to the contents of keep_cols.txt. 
* Note that some false positives will occur, such as "ACC" where we mean "Adaptive Cruise Control" but means "Army Contracting Command" commonly from SAM.gov.

"Keyword Match Counts" shows a summary of the matching activity. The first column shows the keyword (initialism, word, or phrase) and the second column shows the number of matches that occured. This can be useful for verification, or to eliminate certain keywords that may be matching too often.
* For example, "compliance" is a commonly matching word and may not be useful on its own for filtering.