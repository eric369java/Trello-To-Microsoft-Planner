# Trello-To-Microsoft-Planner
    Python script to take a JSON dump from a Trello board and format the results into a Excel file that Power Automate & Plans can integrate with. Need to build an additional Flow in Power Automate to complete migration. Doesn't require hosting. 

## Step 1: 
In Trello, `Show Menu > More > Print and Report > Export as JSON`. Copy the JSON dump into a JSON file. Save the file to the working directory. 

## Step 2: 
In the working directory make sure Python 3, `pandas`, and `xlsxwriter` are installed (`py -m pip install pandas` and `py -m pip install xlsxwriter`. May need to put `python`, `python3`, or `py3` instead of `py`). Create an empty Excel file called trello.xlsx in the working directory. Run `py -m trello [json] [table]`. Instead of `[json]` write the name of the json file saved previously. Instead of `[table]` type the name you want for the Excel sheet. For example: `py -m trello Informatics.json Informatics`.  

## Step 3:
In the Sharepoint for the Team, create a new List. Select Create from Excel file > Upload > choose trello.xlsx. 

## Step 4: 
In Power Automate, set up a Flow to read the Excel file and create tasks & buckets. The current flow looks like
```
Get items from Sharepoint Lists:
For each item: 
    If Created field of task is True:
        Create a new Bucket 
    List All Buckets in the Plan: 
    For each bucket: 
        If task belongs in the bucket:
            Create a new task in the bucket.
List all Tasks
For each task: 
    For each item: 
        If item Title matches task Title: 
            Update task details with description and checklists. 
```
If you have a copy of the existing one, it will need slight modification on its import/export location. Go to Edit flow > Get Items > Change fields Site Address and List Name to the right Sharepoint location and the right list (different for each department). Then in Apply to each 3 > Create bucket, Apply to each 3 > List Buckets, and List tasks, make sure the Group Id and Plan Id are pointing to the plan you want to export to. 

## Step 5:
Run the Flow in Power Automate. 

## Step 6:
Play a chess game. By the time you're done the Flow will finish running (it's really slow)