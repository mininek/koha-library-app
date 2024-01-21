# Koha library app

## Description
Lets registered library users check the books they borrowed. Sends emails when less there are less than 3 days to the due date. Can check accounts of up to 20 users. You must have an account at http://koha.ekutuphane.gov.tr/


## Installation
Create new blank Google sheet. Change the locale to 'Turkiye' if you want the menus and emails to be in Turkish. Go to File -> Settings->General -> Locale -> Turkiye, save settings. Go to Extensions => Apps script. Rename project as 'Koha App'. Copy and paste all the code from Code.gs file in this repository. Save the code and close the google apps script porject. Reload the Google sheet. Select 'Authorize' from the Koha App menu. Select your user account, click advanced, proceed to unsafe and give permissions. Reload the Google sheet and select 'Authorize' again (this is a bug in Google Apps Script, the new menu should show up on the second time). The new menu should be visible now. 

## Usage
Select 'Add new user' to add as many as 20 users. After one or more users have been added you can select 'Check all users' to check the borrowed items of all added users. If there is less than 3 days to the earliest due date an email with the list of the books due will be sent. You can also add a trigger to check the accounts automatically. Go to Extensions -> Google Apps Script. On the left sidebar, select the clock icon (Triggers). Click blue 'add trigger' button on bottom right. Choose which function to run: check_all_users, Select event source: Time-driven, Select type of time based trigger: Day timer, Select time of day: 5am-6am. Click 'Save'. Add another trigger that will check for errors every two hours. Click blue 'add trigger' button on bottom right. Choose which function to run: recheck_if_error, Select event source: Time-driven, Select type of time based trigger: Hour timer, Select hour interval: Every 2 hours. Click 'Save'. These triggers will check the user accounts every day, update the Google sheet, send emails when books are due and will notify you if there are errors. The recheck function will continue checking every two hours if there are any errors.

## Support
You can contact me at: ani kamer (put the two words together) at the regular email service by Google :-)

## Roadmap
I could add a renewing function in the future


## License
Feel free to copy and use this code as you might see fit. No guarantees.

## Project status
I am using it daily for our family library accounts
