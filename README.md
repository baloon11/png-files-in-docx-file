#####Collecting png-files in the folder and adding  in a .docx file

This script collects all pictures (.png files) in the folder, transforms the size and adds a .docx file.

Each picture has a line with a comment.

Whole document has a title.

Usage
------
Put the script in folder containing .png files.

	python Merger.py test.docx 
Where: 
	test.docx -- it is name of  the new .docx file.
 
	You can specify any name of file
	(but with .docx at the end)


Pictures can have different size in the directory.

The size of images in the .docx file changes
to make it look nice and compact.

Requirements
------------

python 2.7

This script uses an additional  python-docx
install:
	pip install python-docx


documentation:
https://python-docx.readthedocs.org/en/latest/index.html

------------------------

A short video, how this script works:

https://www.youtube.com/watch?v=6nqHep8YgzY&feature=youtu.be

