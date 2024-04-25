# barcode-Manager
Barcode Manager Program

There is a GUI that allows you to decide if you want to start a fresh scanning sheet or work off of an old one.

Then from the GUI you can start scanning.

Once the file is created there is a sheet mangamnet window. Here you can decide to remove records from the sheet, or you can start scanning again in order to append more records to the sheet.

The purpose of this program is to allow for easy documentation of the items scanned into each box for item shipments. First a box is scanned, then each item following will be associated with said box. Until another box is scanned, then each item following with be associated with said box.

Now to end the scanning you need to input '/' which will end the scanning.

menu.py: the driver program for the GUI.

barcode.py: the driver program for the barcode scanning (keylogger).

processItems.py: adds a scanned item to a processed sheet.

buildInvoice.py: will be run when the user closes the program to turn the working processed item list into the final invoice format.

invoiceToProcessed.py: converts invoice format to back to the processed format (just a list of all the items data).


need to add:
1. Check if filename exists before it is sent to the sheet menu.
2. files not closing properly before being deleted when the quit process runs and the invoice is generated.
3. Check if scan is valid.
4. Pull boxfilename and itemfilename from a sheet and make them configurable from the application
5. make the font size comfigurable from the application
6. make the application window resulation configurable from the application
7. make the application start from the running of a single executable
8. store a last used sheet in a file to autofill in for the use an existing sheet field to default to
9. Add back buttons rather than having to quit to go back 
10. add label showing to press "/" to end the scanning process
11. add note to application window when item is added with no box

4/25/2024
1. Though attempting to create the version of the program that dirctly changed the amazon file I found a simplier way to create the invoice file. So I am going to take these file operations and make them work with the file self creation mode since the direct mode isnt possible as the amazon file has a password proteciton which was causing issues.
2. Copy init takes the amazon file and configures the working sheet so it has all of the modularity features that the amazon sheet has, and it labeled in the same way so when you are done the item matrix can be copied over to the amazon file.

3/14/2024
1. fixed build invoice to function with the new process file creation.
2. fixed exit response from barcode loop causing menu to crash from blank response.
3. Now the current box during the scan is shown on the sheet menu.
4. optimized the file operations


3/13/2024
1. Made it so the barcode scanner function now only scans one item them returns to the menu the item.
2. Now the processing of the scanned barcodes occur individually so the application can interact with each
item as their scanned to show feedback to the user about their scanning. (So it doesnt work in transactions alike the last interation)
3. Changed how the files are created and passed to the sheet menu so all scanning occurs after the files have been created. (This has
inadvertently broken the invoice generation function for now)
4. now if a item is scanned before a box is scanned it will be ignored.

3/12/2024
1. made some smaller bug fixes and worked on how im going to change the program structure (this version will be archived)

3/11/2024
1. Added a global text size field which controlls the size of all text on the application. 
2. Now the sheet menu allows for scrolling so you can always see all the elements pulled from the working sheet.
3. Clean up working files after the invoice is generated
4. Changed the source sheet to fit the required specifications

3/10/2024
1. invoiceToProcessed.py was created with the invoiceToProcessed() function to turn the invoice sheet into the processed sheet so it can be input by the user and converted to where the program can manipulate it.
2. Now invoice files can be input so the processessing files are no longer needed by the user.

3/8/2024
1. Finished the labeling for the generateInvoice() function.
2. Made the invoice closer to human readable. (Colors and proper formatting)
3. Now the invoice file is opened at the end of the user session.
4. Now the building of the invoice is run at the end of each user session from the working processed file. Where the menu calls the invoice generation function when the user presses the quit button.

3/7/2024
1. Cleaned up the build invoice command.
2. Now the invoice command fully labels the data, inserts the data, and fills in the item instance-box number association matrix.
3. raises the size of each cell in the sheet to show the full data in the cell for a better user experience.

3/6/2024
1. Started working on the buildInvoice functionality of the program (buildInvoice.py)
2. Now the basic information required for the invoice is parsed and passed to the generate invoice where the basic parts of the sheet are created, but its a work in progress.
3. now the generate invoice function is called after the processed items function is called. (which will need to be chagned to whenever the user session is completed)


2/24/2024
1. If a box is not scanned in then the items scanned without a related box will be dropped during processing.
2. Now you can from the opening menu start with an existing record file that is input by the user.
3. Now after you interact with a file and press quit the program will open the excel file.


2/23/2024
1. Finished the ability to delete items from the file from the SheetMenu Frame.
2. Created the ability to start scanning on an existing file and in the processing append the scanned data into the existing file.


2/22/2024
1. Created the GUI in order to make it more user friendly, and make all other features possible.
2. Created the menu to allow the user to select that they want to start a new sheet, and start scanning which then calls the scanning and processing functions.
3. Then I set up the basic funcationality for displaying the scanned items to the user to see if they want to modify the sheet once created.
4. Started working on giving the user the ability to delete rows in the sheet from the GUI menu. Though I wasnt able to get it fully functioning.


2/16/2024:
1. Created the basic functionality for scanning items. (reading the input from the scanner)
2. Created the ability to tell the difference form a box, item, and unregisted using data stored in the boxes sheet and the items sheet.
3. Create the ability to process the scanned items SKU and get the extra item information to display from the item sheet.
4. Store the scanned items into a sheet for later use.