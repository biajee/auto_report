# Excel-to-Powerpoint Auto Report

This is a small GUI python program to screenshot ranges from Excel and paste to Powerpoint. All codes created from Claude Code terminal 4.6.

### Claude resume

Resume this session with:
`$ claude --resume e4b4290a-30d9-4d90-bf56-3e57a3c97085`

### To start

 1. Install pip/python packages
 `$ pip install -r requirements.txt`
 
 2. Run the program
 `$ pythohn3 main.py`

### All prompt used

 - "Write a python code with gui frontend. Major task is to let user select spreadsheet and range to screenshot and record persistently for later use"
 - "try it out"
 - "Can you open the excel file and let user choose range on the excel?"
 - "try it out"
 - "For each screenshot, I also want to paste it onto a powerpoint slide. Let me open a powerpoint to paste to a page with specified position and size."
 - "Now, open powerpoint so that I can choose where to paste and how big will it be."
 - "Right now, user cannot select a different page on the powerpoint to define the paste area. Let's add that function so that user can choose which page to paste even after powerpoint is opened."
 - "User should be able to set or varify a few drop down boxes' values, let the excel spreadsheet calculate and copy the content in the range. Make sure only calculate if drop down boxes values are changed."
 - "Add a message box to display the current status one by one."
 - "try it out"


