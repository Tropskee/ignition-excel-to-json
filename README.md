# ignition-excel-to-json
Convert excel files to JSON - formatted to be digested by Ignition

## How to Use
Simply clink the shortcut (.lnk) to start the program.
You will need to browse for and select the .xlsx or .csv file you wish to convert to JSON.
Enter a name for the created JSON file.
Change the folder name if required (default is 'MM')

The program will run and the outputted format is as follow: 
{"name": "MM", 
 "tagtype": "Folder",
 "tags": [
    {
      -Excel_Content-
    }
  ]
 }

### Note
If for some reason the application is not launching, the original application file can be found in program_files -> dist -> ignition-excel-to-json
