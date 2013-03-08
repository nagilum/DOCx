# DOCx

DOCx is a PHP library designed to manipulate the content of .docx files.

Maintainer: Stian Hanger (pdnagilum@gmail.com)

.docx files (Microsoft Word 2007 and newer) and basically a zipped archive of
.xml files which holds different contents. If you unzip a .docx file you will
get a lot of files, among them word/document.xml, word/header.xml, and
word/footer.xml. For some reason Word adds multiple documents for each section.
So there might be 3 footer files, named: word/footer1.xml, word/footer2.xml,
and word/footer3.xml. This library takes that into account when manipulating
the variables.

Functions of the library:

* cleanTagVars - Cleans out the excessive '${' and '}' tags in the document.
* close - Resets the class for a new run.
* load - Loads a .docx file into memory and unzips it.
* save - Saves the temporary buffer to disk.
* setValue - Does a global search and replace with the given values.
* setValues - Does a global search and replace with an array of values.
* setValueDocument - Does a search and replace in the 'document' part of the file.
* setValueFooter - Does a search and replace in the 'footer' part of the file.
* setValueHeader - Does a search and replace in the 'header' part of the file.
