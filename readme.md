![sample splunk output](https://git.soma.salesforce.com/sfoo/casam/blob/master/app/splunk_aus22.png?raw=true)


# WD Google Spreadsheet 
Appscript to convert 2013 Excel with VBA micro to Google spreadsheet

## Steps to Access my report using Google Spreadsheet
# 1) Put up a request to owner to share the Spreadsheet with you
![sample splunk output](https://github.com/samxrc/wd_ss/blob/master/1_openemail.png?raw=true)

# 2) You'll receive email & click continue with ur account
![sample splunk output](https://github.com/samxrc/wd_ss/blob/master/2_continue.png?raw=true)

![sample splunk output](https://github.com/samxrc/wd_ss/blob/master/3_choose_ur_account.png?raw=true)

# 4) Trust & allow ,done 
![sample splunk output](https://github.com/samxrc/wd_ss/blob/master/4_advanced.png?raw=true)
![sample splunk output](https://github.com/samxrc/wd_ss/blob/master/5_trust.png?raw=true)
![sample splunk output](https://github.com/samxrc/wd_ss/blob/master/6_allow.png?raw=true)








2) clone template google spreadsheet ,git pull from wd_ss
3) import Appscript, done 


# ---Building steps for reference
~ clone sample report 

~ create equivalent VBA which is not supported on google spreadsheet using javascript. Build picklist for stock, put/call

~ build button to create new sheet 

~ summary page for us market 
 

# pending wish list  feature 
~ Auto update current price , from config page
~ auto update summary page when new ss created ,get info from config page
~ Button to email out weekly report , get info from config page





## -----reference 
google Api 
https://developers.google.com/apps-script/reference/calendar/calendar-app
https://developers.google.com/apps-script/reference/calendar/calendar-app#getCalendarById(String)
