# THL_converter

This folder contains the documents necessary to convert Word documents marked-up using THDL styles into valid XML.

This folder contains the following 3 files:
- teiHeader.dat -- This is a data file that contains the mark up for the metadata. This is the file that contains the XML markup for the header into which the information from the metadata table is automatically inserted.
- converter.py -- This is the python program for the actual conversion. This should be run with the target document (docx) as a command-line argument. Multiple documents can be given separated by a space or the name of a folder in this directory that contains the documents. The program will output XML files with the same name as the original document(s) with a .xml extension.
- example.docx -- This is an example of a properly marked up Word document. Use it to create a sample output of the converter.

To convert a marked-up Word document (example.docx) to XML, enter this directory and run:
`$ python converter.py example.docx`
This command generates example.xml.

## Dependencies
- Python (>= 2.7)
- [lxml](http://lxml.de/installation.html) (>= 3.6.4)
- [python-docx](http://python-docx.readthedocs.io/en/latest/user/install.html) (>=0.8.6)