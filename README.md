# DiningMenusApptegy
This is a repository for creating recurring ADA compliant menus that will dynamically update as information is filled for any Apptegy powered school district.

First, download the Dining Menus.xlsx file in this repository and upload it to Google Sheets. //In some cases, you may need to copy the formulas that will be listed in the XLSX file. It should parse out the remainder.

In the toolbar, choose Extensions > Apps Script, and paste the content of appscript.js. On run, this will add a toolbar to allow dynamic introduction of your menu, and allow you to choose or adjust the formulas from toolbar as you copy out content.

This is designed to take the color style/structure and append White/Black text based on the highest ADA compliance ratio (anything over 4.5).
