These two scripts will get the daily results for team member contributions in the World Community Grid.

The two scripts (WCG_Stats.ps1 and Results.ps1) and the text file (Daily_Results.txt) must reside in the same folder.

The first script, WCG_Stats.ps1 takes a manualy created text file (Daily_Results.txt), filters it and converts it to a .csv file.

For historical purposes, I also save the text file using the current date. The script doesn't use those files for anything, it's just a record.

The first line of the text file must contain the date the statistics are for and must be formatted like this:
Statistics Last Updated: 5/29/20 23:59:59 (UTC) [2 hour(s) ago]

That first line is actually copied off of the web page in the team member statistics. To get to the team member statistics, do the following:
  Click on My Contribution
  Click on My Team
  Scroll down to Team Member Details and Statistics
  Click on Points Generated (underneath Sorted By)
  Copy the line immediatley underneath "Team Member Details and Statistics: <team name>"
  Paste it as the first line in the Daily_Results.txt file.

The rest of the file contents is the actual member statistics. 

The easiest way is to set the records per page to display everyone in your group. My group has 98 people in it, so I set it to 100.
Highlight everyone on the list by placing the cursor in front of the first name, hold the left button down and drag it to the last record of the last name.
Be careful not to capture any other part of the web page. Just hightlight the records.
When all the records are highlighted, right-click on the highlighted area and select copy.
Paste that information into Daily_Results.txt starting at line 2.
Remove any extra blank lines at the end of the file. 

It won't looked lined up when pasted in as each item in the record is separated with tabs that don't always line up. Don't worry about that, the script will deal with it.

Save Daily_Results.txt
At this time, I also save it as the date I captured it (ie 2020-05-29.txt) for historical purposes.

Open up PowerShell_ISE.
Open the script WCG_Stats.ps1 and run it.
It will create a Results folder if one doesn't exist, and create a dated filename with an underscore preceeding the name (ie _2020-05-29.csv).

After it finishes, open up the script Results.ps1 and run it.
Results.ps1 will ask if you want to display people with no progress. If someone return zero results for the day, this is the option to display them.

Next, the script will ask you to open the newest file and presents a file-open dialog box. 
Select the newest file that has the underscore "_" preceeding the filename (ie _2020-05-29.csv).
Next, it will ask you to select the previous day. In this case the file would be _2020-05-28.csv.

When Results.ps1 finishes, it will open up notepad with formatted text that can be copied and pasted into the World Community Grid forums.
For simplicity, I just hit Ctrl-A, then Ctrl-C to select it all and copy it.

It looks like a lot, but the entire process takes less than a minute.
