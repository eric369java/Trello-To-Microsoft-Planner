# Trello-To-Microsoft-Planner

## Step 1: 
In Trello, `Show Menu > More > Print and Report > Export as JSON`. Copy the JSON dump into a file called trello.json. Save in working directory. 

## Step 2: 
In the working directory make sure Python 3, `pandas`, and `xlsxwriter` are installed. Creatr an empty Excel file called trello.xlsx in the working directory. Run `py -m trello`. 

## Step 3:
In Sharepoint, create a list from trello.xlsx. 

## Step 4: 
In Power Automate, set up a Flow to read the Excel file and create tasks & buckets. The current flow looks like
```
Get Tasks from Sharepoint Lists:
For each item: 
    If Created field of task is True:
        Create a new Bucket 
    List All Buckets in the Plan: 
    For each bucket: 
        If task belongs in the bucket:
            Create a new task in the bucket.
```

## Step 5:
Run the Flow in Power Automate. 