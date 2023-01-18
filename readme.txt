I want to create a gui interface using tkinter that allows me to input my delivery driver
data into an excel sheet, I would start this processing creating a window with tkinter 
that allows me to add at date

* I should make in into a table
* The different values should be:
 date - MM/DD/YYYY
 amount - $
 hours - num
 total deliveries - num
 
 $ wage amount - my hourly rate * my hours
 $ per hour -  wage + tip amoutt
$ per deliver taken
 total made that day


# Have another table that would show:
tip total
total made from hours
total hours worked
average hourly income

total profit
days logged
average daily tip
average per delivery

# Add a button to udate excel to sheet.

my initial design is 
a button that lets me add data to the excel sheet
another button that saves the sheet
a button that displays can display all the totals

I want to create tables for the user inputted data
I want to have my program automatically extend the references of the table 
We will need two tables, one for userinputted data and statistics

I got the tables to work

* total table

Take the wage amount from the statistic table and return it where total wage is
Take the tip amount from the userinputted table and reutn it where total tip are
Take thedail total amount from the statistic table and return it where total income is

Create a button that updates totals

# Create a button that dispaly the most recent day logged ( just provide the date)
# Create a button the displays the highest tip
# add a button that allows me to open the excel sheet, and finally a button to quit


 I want to add a pop up for when i add an excel sheet or something to tell me that it has been updated


 I have ran into a problem though, i have just realized that my wage is actually 7 .98 per hour and not 6.98, so this will be a good test to see if something will break if i try updating it
 create a settings button, that create a pop up that has a button to a function that updates the wage amount column


 Ultimately I want to recreate the program in a object oriented manner using classes and object, right now I feel that my code is really ugly and very messy and I want to change that