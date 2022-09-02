The program reads the Excel source file and, based on the data defined in the control panel, creates separate files and sends them to the appropriate people.
- **control panel.xlsx:**
   - **units_required** sheet - each row is dedicated for one email. 
      -   [Unit Name]: The list of Continents/Regions/Countries. You can place multiple units of the same kind (e.g two Regions) in one cell but they must be separated by a comma; 
      -   [Emails]: The list of recipients. Emails must be separated by a semicolon; 
      -   [Email Type]: The content of two types of emails is present in the 'emails' sheet; 
      -   [Operation Type]: _Full_ - creates the following sheets in the final file: _current data_, _TBMs_, _summary_. _Simple_ - creates the _summary_ sheet only.
   - **emails** sheet:
      -  There are two types of emails and their content can be edited. Each cell relates to one paragraph. Additional paragraphs can also be added.
      -  < last day of the previous month >  will not be displayed but replaced with the proper date.
   - **cycles, systems** sheets: read-only sheets.

- **source file:** - it contains three worksheets and must be placed in the source file/ subfolder.
 
- **Additional information:**
  - After starting the program, select whether you want the e-mails to be sent immediately or only displayed.
  - The program checks if the required files exist and if they contain the required sheets and columns.
  - The source file name does not matter, but there must be only one xlsx file in the source file/ folder.
  - If there are elements in the [Unit Name] column that are not in the source file, the name of such element will appear in the _elements not found.txt_ file.
  - All generated xlsx files are saved in the final files/ folder.
  - The first and last names in the source file are fictitious and any coincidence is accidental.
  - The program is designed for Windows because it uses MS Outlook to handle e-mails.
