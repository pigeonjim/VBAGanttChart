# VBAGanttChart
Code to create a Gantt chart in VBA that provides the excel users with various functionality 

This started out as a favour to someone in work and has evolved. The reason I am putting it on gitthub? This version has more functionality then the one I did for work. I also dediced for better or for worse to OPP it too before posting it here. I had never tried OOP in VBA before and it is as limited as most things in VBA are. You cant pass variables to constructors (but there are work arounds) and it isnt fully OPP as inheritance isnt available. But it was a very worth while exercise as I have learned a lot and I enjoyed the process

What does it do?

It will allow end users to input a project start date and a project end date. Project start date should go into C5 and project end date into C4. The program will prompt the user if the project start date is prior to 01/01/21 and ask them to input another date. It will also prompt if the project end date is prior to the project start date.
The code then uses these dates to create a calendar on the same sheet which will cover from the project start date to the end of the month after the end date. Weekends are in grey cells and todays date is in a yellow cell. It has been coded to account for leap years.

Users can then enter tasks below this; from row 8 onwards. For a task they need to enter a start date and the number of working days to complete the task. Task start dates go into column 5 and the days to complete task go into column 6. Columns 7 onwards are used for the calander. There can be blank rows bewteen tasks.

There is event driven code that will
a. update a cell with the percentage of progress so far
b. a drop dowm will be created when there is a task start date and a day to complete task value. This will delete if both are removed
c. a bar will be drawn on the calendar that covers the task dates but not including weekends. The bars will delete if either task start date and a day to complete task value is removed.
d. The colour of the bar will indicate what stage the task is in.
  1. Light green = dates in the future
  2. Dark green = dates already passed
    The green bars can have both colours.
  3. Orange = the task is due to be completed today or in the next 2 working days
  4. Red = the task is overdue
  5. The end user should use the crop down to indicate if the task is complete. This will change the bar from orange or red to dark green
e. When the proect end date is removed the calander will refresh to only show the project start month
f. When the project start date is removed the calander will clear.
g. When the sheet is open the bars and calendar will refresh.

Other things to note -
The code name for the sheet needs to be changed to GanttSheet. The end user can all the call what they want
There is an CCalander object created when the sheet is open
A dictionary is created when the sheet is opened to hold the CTask objects. Each task has a CTask object that is inialised when when either a start date or a days to complete is input. When both the days to complete and the task start date are removed the oject is deleted
