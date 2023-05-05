---
layout: page
title: "Version History"
permalink: /versionhistory/
---

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
