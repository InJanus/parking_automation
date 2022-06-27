# Understanding Automation Script
> Brian Culberson | CompE 2022 | Contact [**here**](http://injanus.tech)

This is a quick guide to the script that I made to help with running this script in the future.

This script is made to go though a list of license plates, look them up in the Ohio state database, then link their info to T2. This was made for UC Parking and authorised by Matt Burden (96)

## Prerequisites.
---
- Python Version : 3.10.5
- Chromium Version (chromedriver.exe) : 102.0.5005.61 [Download](https://chromedriver.storage.googleapis.com/index.html?path=102.0.5005.61/)

If you ever run into an issue that says `ModuleNotFoundError: No module named 'example'`. Simply run `pip install example` to install the module. If you keep doing this then eventuly the errors will go away.

## Running the script
---

Up near the top of the code, you will see a constants section. This is where all of the login infromation goes as well as the local installation of the excel file to be read as well what sheet you are reading. **Make sure the format of the excel sheet is the same as it was in the past.**

Near the bottom of the code is the main loop. There is a starting index and an ending index. Change these to run diffrent sections of an excel sheet or just do the whole index range.

To officaly run the script, first close the excel file you want to write to. Type **`python script.py`**

This will open a chrome window and the program will start.

