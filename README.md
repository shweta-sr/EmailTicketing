EmailTicketing
==============

A combination of MS Word, MS Excel &amp; MS Outlook apps to create and send custom e-ticket PDFs.

Download all files to one folder.

First, take a look at the Word DOCM. This is the template based upon which the tickets are created. All the fields within the template that ought to be replaced with actual values are within # marks. And, of course, replace images and wording as applicable.

Next, look at the Excel file. It has just one tab and is the database of info needed to create your tickets. Cell A1 contains the path where you want the PDF tickets saved. 

To customise for your needs:
1. Change the Word doc to what you want your tickets to look like, keeping replaceable fields in #. You can feel free to change the names of the fields - as long as they *exactly match* the column names in Row 2 of the Excel doc.
2. In the Excel file:
 a. Change the path name.
 b. Keep names in column A, e-mail addresses in column B and the formula for file path and name in column T (although, if you wish, you can change the *names* of the columns).
 c. Clear away Indus' columns and place your own column names and data.
3. Go back to the Word document and run the macro MergeDocs.

This should create your tickets.

Now, to send your tickets.
1. Import AutoEmail.bas into your Outlook.
2. This macro is not very polished. You will need to read through and replace a lot of hard-coded stuff with your own. Fortunately, it is a very simple and short sub.
3. First run the macro with the Send line commented and the Save and Display lines uncommented. Once you are satisfied, you can use the Send line to actually get the e-mails out.
