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


need to add: (Pressing Matters)
1. find way to stop crashing from leaving window while in scan mode
    -> run in different process menu and scan?
2. Find another way to do the deletions now that there are not indexes 
3. input error handled required
4. made a bad sound for a misinput from the barcode scanner, filename misinput so on..
5. Made the instructions on the sheet menu easier to understand


need to add: (Luxaries)
1. make the font size comfigurable from the application
2. make the application window resulation configurable from the application
3. make the application start from the running of a single executable
4. store a last used sheet in a file to autofill in for the use an existing sheet field to default to
5. Add back buttons rather than having to quit to go back to the menu
6. add note to application window when item is added with no box


6/1/2024
1. Added a sound that plays when a box or item is scanned and processed since sometimes the scanner scans an item but doesnt properly input the barcode value. Which makes it hard to know if it saved the scanned data since the barcode scanner itself will make a noise no matter what.
2. The last input filename is now saved so you dont have to retype it in everytime you close the program.
3. Now the total expected item instances for each item object and the current value registered to boxes is shown in the menu for a better user experience.

5/30/2024
1. Extended the sleep period of the scanner.
2. Now screen menu update directly reads the values from the tables and shows the contents in the menu in a table format to make it understandable.

5/26/2024
1. Added new function to Sheet Menu to increment the Box number in the sheet.
2. Added the Box Number counter to the refresh function in Sheet Menu and a button to call the increment function.
3. Fixed the formatting and window size to fit all of the contents on the window at startup.

5/25/2024
1. Fixed the delete function for the Sheet menu so it gets the input in the correct format to inturrput it.
2. Added instructions to the Sheet Menu.
3. fixed the change sheet format to show all of the cell values fully when its created.

5/22/2024
1. Created the delteRecord file operation function to take in the list of items pulled from the file in the Sheet Menu update and the user input item to delete and makes the change in the file.
2. Fixed indexing error in the appendNewItem function.
3. Updated the GetFilename screen to differentiate between and amazon file and changed file. Then decide to use the copyInit function and pass the filename to the SheetMenu according to the input.
4. Modified the sheet menu page to utilize the file operations.
5.

5/21/2024
1. Added the boxed quantity formula to the created file so it dynamically counts the amount of item instances for each item record
2. Fixed the Sheet menu update function to Reads in the values from the current working sheet to display the current item instances in the menu
3. Added the file operation (appendNewItem) to update the file instance matrix in the working file when a new item is scanned when asscoaited with a box.

5/20/2024
1. finished file creation from origional amazon file.
2. fixed some problems with opening the amazon file from the menu.

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