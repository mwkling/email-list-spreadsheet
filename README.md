# email-list-spreadsheet
VBA code from a spreadsheet to manage filtering of email contacts when sending marketing campaigns.

## To Use

* Create a new Excel spreadsheet.
* Open the VBA console (Alt+F11)
* Create a new module and copy in the contents of main.vba
* Setup an area for settings on the main worksheet.  Add a named range for 'input_directory', 'out_file_name' and 'exclude_domains'.  Optionally add buttons to run the main 'do_all' macro and the 'reset_book' macro.  Look at 'example_sheet.png' for an example spreadsheet setup.
* Download your contacts and suppression reports into input directory
* Run do_all

For more information on what this is for, read the relevant [blog post](http://mikekling.com/excel-marketing-list-export-spreadsheet/).
