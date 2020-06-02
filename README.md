# ExcelCard #

ExcelCard is an Excel racecard using VBA to query the Smartform database. This is intended as a simple example of pulling data into Excel and could be the starting point for your own customised race card. For example you might develop some database stored procedures that create statistics tables, for example, trainer, jockey and sire/dam strike rates, that you can query from VBA and populate your race cards with statistics that are tailored to your approach to betting.

The first sheet named "ExcelCard" has three buttons:

- UpdateDB - calls the SmartForm database update script
- Get Cards - creates today's racecards with one sheet per meeting
- Clear Cards - deletes the previously created racecards & sheets

## Database Connection ##

Before opening the spreadsheet for the first time you will need to do the following to enable Excel to talk to MySQL:

- In Excel VBA Editor (Alt-F11) you need to go Tools/References and check Microsoft Active X Data Objects 2.x library
- Install MySQL ODBC 5.3 (or later) Driver. See [dev.mysql.com/downloads/connector/odbc/5.3.html](dev.mysql.com/downloads/connector/odbc/3.51.html) or google "MySQL ODBC 5.3 Driver"

## Racecard Columns ##

The spreadsheet displays today's race cards with one sheet per meeting. The columns are as follows:

- A - Saddle cloth number (draw number)
- B - Runner name (country) with TT = owner name, yellow highlight if owner change since LTO with prior owner in TT
- C - Age / gender
- D - Official rating (lbs out of handicap)
- E - Weight (jockey allowance), yellow highlight if LTO race was of opposite handicap/stakes status
- F - Days since LTO with TT = LTO course & track type, yellow highlight if same course, red text if within last 2 weeks
- G - Trainer name, yellow highlight if trainer change since LTO with prior trainer in TT
- H - Jockey name
- I - Sire name
- J - Dam Name
- K - Foaling date (useful for 2yo races)

LTO = Last time out

TT = Tool Tip / mouse hover-over text.

The basis for calculating "lbs out of handicap" is explained here:

[https://answers.betwise.net/1173/working-out-if-a-runner-is-out-of-the-handicap](https://answers.betwise.net/1173/working-out-if-a-runner-is-out-of-the-handicap) 

Please feel free to ask questions by opening a GitHub Issue.