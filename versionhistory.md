---
layout: page
title: "Version History"
permalink: /versionhistory/
---

***Version 4.4.4:*** 6/19/2024
  - Helo's Package Data function no longer remove images out of their respective structure folders.

***Version 4.4.3:*** 5/1/2024
  - Updated Helo Extract Structures List

***Version 4.4.2:*** 4/30/2024
  - Changed Helo Block ID to D/T2499-0001
  - Fixed Helo Data Packaging bug
  - Minor improvements and bug fixes.

***Version 4.4.1:*** 4/22/2024
  - Added deletion of hidden desktop.ini and .DS_Store files
  - Updated Helo data processing.
  - Minor improvements and bug fixes.

***Version 4.4.0:*** 3/21/2024
  - Added "Time Sheet" to Quick Links
  - Combined Transmission and Distribution dropdowns into "Data Processing". The script now figures out which scope the user is processing using the name of the GIS extract or Filtered Extract.
  - Traveler sheet templates are now embedded in the app so user will not be asked to input them anymore when running "Complete Traveler" and "Merge Travelers". The script also decides which template to use.
  - Should the script encounter difficulty in identifying the correct template due to unconventional extract or traveler sheet names, it will prompt the user to double-check the accuracy of the entered sheets, and then ask the user which template to utilize.
  - Minor improvements and bug fixes.

***Version 4.3.5:*** 3/11/2024
  - Fixed a bug that occasionally caused the process of merging traveler data to be stuck, ensuring smoother operations.
  - Implemented automatic addition of "OH-" prefixes to structure IDs in merged traveler sheets when absent, enhancing data consistency.
  - Adjusted the table range in merged traveler sheets to automatically reflect the intended data set, eliminating the issue of bottom rows not being a part of the table.
  - Minor improvements and bug fixes.
  - Added yellow highlights and font color to internal refly structures ("Internal Refly Complete") and EZ poles that are in both scopes ("In Distro Package" or "EZ Pole on Trans Map") in FilteredExtract sheet
  - Added Internal Refly sheet to the Filtered_GIS_Extract file.
  - Watermark Prep now deletes hidden Mac files before processing images (this way, you don't have to filter extract first before watermarking).
  - Minor improvements and bug fixes.

***Version 4.3.4:*** 3/7/2024
  - Transmission's "Filter Extract" funtionality now asks for the distribution extract when it detects "In Distro Package" EZ Poles to accommodate for EZ poles that are in both scopes.
  - Minor improvements and bug fixes.

***Version 4.3.3:*** 3/5/2024
  - Added an issue summary pop-up after the Watermark Prep process is done.
  - When encountering an error during the watermark prep process, the problem image will be skipped and be added to issue summary instead of stopping the whole script.
  - Improved error handling and visuals.
  - Minor improvements and bug fixes.

***Version 4.3.2:*** 3/4/2024
  - Added a "Copy" functionality when user right-clicks on the textbox area
  - Added yellow highlights and font color to RTV structures in FilteredExtract sheet
  - Added RTV sheet to the Filtered_GIS_Extract file.
  - Minor improvements and bug fixes.

***Version 4.3.1:***
  - Daily delta for transmission and distribution are now combined ("Daily Delta" button).
  - Resolved an issue where list.remove(x) shows an error when resolving duplicate matches.
  - Font colors are added to the text box -- red font for issues, green for process completion (more in future updates).
  - Minor improvements and bug fixes.

***Version 4.3.0:***
  - Added "Pre-upload Validation" that checks for various things to make sure data is ready to be zipped and uploaded to OneDrive.
  - "Undo Package" now adapats to the OH- prefixes and names the structure folders without the OH-.
  - "Package Data" now checks if the images are named properly and if the directory has all the structures listed in the completed traveler sheet.
  - Fixed a bug where the dropdown list for Vendor Category in the traveler sheet is incorrect.
  - The completed traveler filename date is now based on the latest flight date on the traveler sheet.
  - Added a "Find" button (CTRL+F) to search for texts in the textbox area. (needs further improvements)
  - Tooltips are now fixed and reinstated.
  - Minor improvements and bug fixes.

***Version 4.2.0:***
  - Fixed a bug where Flight_Date column is not formatted correctly for transmission traveler sheet.
  - Fixed a bug where the datetaken format of an image is not recognized when doing Watermark Prep. It is now treated as an image with no date, so user will be asked to enter a date.
  - Added Team_Number and EZ_in_Distro to transmission traveler sheet.
  - Added a prefix "OH-" when renaming distribution images.
  - Added prefix "OH-" to SCE_STRUCT column on both traveler sheets.
  - Complete Traveler now creates a new file instead of replacing the traveler sheet template. New file is named "[qa_first_name]_[D or T]_C2_[filtered_extract_date].xlsx" and saved in the same directory as the filtered extract.
  - Merge Traveler now creates a new file instead of replacing the traveler sheet template. New file is named "[D or T]_C2_[date_today].xlsx" and saved in the same directory as the traveler sheet template.
  - Minor improvements and bug fixes.

***Version 4.1.1:***
  - Fixed Austin Tinnell's name. Sorry >_<.

***Version 4.1.0:***
  - Improved UI.
  - Added user "login" (no password... yet) and added QA name column to completed traveler sheet.
  - Added "Quick Links" to the menu options. It consists buttons that link to various web pages used daily by QA.
  - Integrated Data Conformance script (the one Ben uses) to "Daily Delta".
  - Added tabs to Filtered Extract for EZ Poles in Distro, EZ Poles in Trans, and AOC.
  - Added "OH-" before the structure number when renaming tranmission images.
  - Adjusted functionalities to conform with the new GIS extracts and traveler sheets.
  - Tooltips have been removed due to some issues (will be reinstated once fixed).
  - Minor improvements and bug fixes.

***Version 4.0.4:***
  - Minor improvements and bug fixes.

***Version 4.0.3:***
  - Minor improvements and bug fixes.

***Version 4.0.2:***
  - Minor improvements and bug fixes.

***Version 4.0.1:***
  - Minor improvements and bug fixes.

***Version 4.0.0:***
  - New and improved UI.
  - Added "Daily Delta" button to compare and validate data from the field to QA.
  - Minor improvements and bug fixes.

***Version 3.8.0:***
  - Added "Filter Helo" and "Package Helo" buttons to process helo structures.
  - Minor improvements and bug fixes.

***Version 3.7.0:***
  - Added an "Issue Tracker" button that directs user to the Issue Tracker Form.

***Version 3.6.7:***
  - Minor improvements and bug fixes.

***Version 3.6.6:***
  - "Complete Traveler" now takes empty traveler sheet templates for both distribution and transmission.
  - Minor improvements and bug fixes.

***Version 3.6.5:***
  - Enhanced handling of extract duplicate structures.
  - Minor improvements and bug fixes.

***Version 3.6.4:***
  - Minor improvements and bug fixes.

***Version 3.6.3:***
  - Minor improvements and bug fixes.

***Version 3.6.2:***
  - Minor improvements and bug fixes.

***Version 3.6.1:***
  - Minor improvements and bug fixes.

***Version 3.6.0:***
  - Enhanced handling efficiency for misnamed folders during "Filter Extract" operation.
  - Introduced a feature in "Filter Extract" to create a new tab for every team being processed, comparing folder names with GIS structure IDs from the respective flight dates.
  - Minor improvements and bug fixes.

***Version 3.5.6:***
  - Improved prompts and handling of incorrectly named folders when filtering extract.
  - Minor improvements and bug fixes.

***Version 3.5.5:***
  - Minor improvements and bug fixes.

***Version 3.5.4:***
  - Minor improvements and bug fixes.

***Version 3.5.3:***
  - Fixed a bug where the script thinks there are duplicate Structure IDs when there is none while filtering transmission extracts.

***Version 3.5.2:***
  - Updated "Filter Extract" to accomodate for duplicate Structure IDs in the extract.
  - Minor improvements and bug fixes.

***Version 3.5.1:***
  - "Merge Extract" now tries to fill out missing Mapped_Lat, Mapped_Lon, Structure_, and FLOC values, if possible. If there is no way to find those missing values, the script will flag a message showing whic h row these missing values are.

***Version 3.5.0:***
  - Deletion of unnecessary columns when filtering extracts are now back.
  - Added "Merge Extracts" button to combine the regular extract and new map extract into one.
  - Minor improvements and bug fixes to accomodate new map extract.
  

***Version 3.4.1:***
  - Temporarily removed the deletion of unnecessary column when filtering extract.
  - Minor improvement and bug fixes.

***Version 3.4.0:***
  - Changed "GIS vs UC vs TS" back to "GIS vs Upload Check" to compare only GIS and Upload Check data
  - Added "GIS vs Traveler" button that compares daily compiled traveler sheets to GIS in order to spot discrepancies before submitting to SCE.
  - Combined "Watermark (Sky)" and "Watermark (DJI)" into "Watermark Prep" button. It automatically detects which drone/camera was used to capture the images.
  - Added "Delta Report" button that finds discrepancies between the final Traveler Sheet and the GIS (and Upload Check). It is made to prep research on missing structures.

***Version 3.3.1:***
  - Fixed a bug where Filter and Rename for Distribution processes data as Transmission.
  - Fixed a bug where Filter Extract shows an error message when creating "Issues" sheet if there is no issue found. It now shows "No issue found. Extract is successfully filtered and processed." and skips creating the "Issues" sheet.
  - Minor improvements.

***Version 3.3.0:***
  - "[D or T]_Filtered_GIS_Extract_[date].xlsx" is now saved in the same directory as the original extract (not the structure folders).
  - Filter Extract now creates a new sheet called "Issues" that lists all issues found in the extract, which can be copied and pastied to Issue Tracker.
  - Minor improvements and bug fixes.

***Version 3.2.0:***
  - Changed "GIS vs Upload Check" to "GIS vs UC vs TS" which now compares GIS, Upload Check and Traveler Sheet.
  - Minor improvements and bug fixes.

***Version 3.1.1:***
  - Fixed a bug where version history is not showing when clicked.

***Version 3.1.0:***
  - Added "GIS vs Upload Check" button that compares structure IDs reported in GIS and Upload Check
  - Filter Extract now creates a new file instead of overwriting the GIS Extract. New file is named "[D or T]_Filtered_GIS_Extract_[date].xlsx" and is saved in the same directory as the structure folders.
  - Minor improvements and bug fixes.

***Version 3.0.1:***
  - The app is optimized for faster initial loading.
  - Minor improvements and bug fixes.

***Version 3.0.0:***
  - Main window now holds the "console". No more separate console window.
  - Added functionality to copy console messages and error messages.
  - Minor improvements and bug fixes.

***Version 2.1.2:***
  - Merge Directories now removes empty folders after merging.
  - Merge Directories and Merge Travelers widgets now stay on top of any open windows.

***Version 2.1.1:***
  - Added more print messages for easier troubleshooting.
  - For transmission, structure ID reference has been changed from "StructureN" column to "FLOC" column.

***Version 2.1.0:***
  - Added “Undo Package” button. It moves images into folders according to their structure ID numbers. Version 2.0.1:

***Version 2.0.1:***
  - Fixed a bug where update needs user to have Python in their system. Update should now work properly even without Python.

***Version 2.0.0:***
  - Bug fixes
  - Added Merge Travelers and Merge Directories.
  - EEAAO button, the button that does everything in one push, has been removed. It's too much XD.
  - The main window won't freeze anymore while running a script. This means you can now do multiple processes simultaneously.
  - Renaming scripts now correctly print the original names and the new names on the console.
  - Filter Extract now creates a new sheet that shows the distance of each image from the nadir. The farthest distance is highlighted. This will help us identify multiple structures that are mixed in one folder.
  - Filter Extract also checks for incorrect folder names, images with no GPS data, images with incorrect date, nadir shots with no "N". Use these information to update our issue tracker.
  - Package Data now pauses and prompts you to make sure images are named correctly, and traveler sheet is perfect, before it moves the images out of their respective structure folders (coz it can't be undone).
  - Complete Traveler now adds extra columns to show P1/P2/P3 Notes for distro, and P1 and Mile/Tower for trans for easy reference. Remember to delete these columns when you're done QAing.
  - Merge Directories counts the number of images in each Block ID folder and the total number of images in all of them combined.
  - Tooltips added and will appear when you hover over a button.
