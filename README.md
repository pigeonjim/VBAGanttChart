# VBAGanttChart
Code to create a Gantt chart in VBA that provides the excel users with various functionality 

This started out as a favour to someone in work and has evolved. The reason I am putting it on gitthub? This version has more functionality then the one I did for work. I also dediced for better or for worse to OPP it too before posting it here. I had never tried OOP in VBA before and it is as limited as most things in VBA are. You cant pass variables to constructors (but there are work arounds) and it isnt fully OPP as inheritance isnt available.

What does it do?

It will allow end users to input a project start date and a project end date. The code then uses this to create a calendar on the same sheet which will cover from the project start date to the end of the month after the end date. Weekends are in grey cells and todays date is in a yellow cell.

Users can then enter tasks below this. For a task they need to enter a title, a start date and the number of working days to complete the task. There is event driven code that will
a. update a cell with the percentage of progress so far
b. a check box will be created. This will delete if the task details are deleted
c. a bar will be drawn on the calendar that covers the task dates but not including weekends. The bars will delete if the task details are removing.
d. The colour of the bar will indicate what stage the task is in.
  1. Light green = dates in the future
  2. Dark green = dates already passed
    The green bars can have both colours.
  3. Orange = the task is due to be completed today or in the next 2 working days
  4. Red = the task is overdue
  5. The end user should tick the check box to indicate that the task is complete. This will change the bar from orange or red to dark green
e. When the sheet is open the bars and calendar will refresh. The sheet will focus on todays date.
f. I am working on events for ticking and ticking the check box. This is not straight forward as the boxes are added and removed by the code using events. This means I cant use the click on event to run code. VBA requires that I know in advance how many boxes they are, where they are and what they are called before the sheet is used.
***as an interim solution there is a refresh button on the sheet***
g. A further addition over the non OOP version is code to populate and format the sheet ready for use. It adds the title, headers, etc. This will refresh and populate when the sheet is opened or refreshed.
