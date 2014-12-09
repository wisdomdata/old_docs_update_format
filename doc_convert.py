
# program to convert various older formats (lotus, wordperfect) into
# last updated monday 8th Dec 2014
# Licence: use for personal / business use, please return improvements.

# import needed
import sys
import os
import magic
import shutil


# get print statement working

# template file for .wk1, this file needs to exist.
inital2xls = "tmp/AWGM100A.WK1"

# template file for wordperfect document, needs to exist.
intial2doc = "tmp/93ANREP"

_2xls = " --convert-to xls "
_2doc = " --convert-to doc "


ft1 =  magic.from_file(inital2xls)
print "ft1", str(ft1)

ft2 = magic.from_file(intial2doc)
print "ft2", str(ft2)

# location of original files
orig_loc = "Floppy_originals"

# where copy of files will be created - work will happen here.
out_loc = "floppy_convert"


shutil.copytree(orig_loc, out_loc)

typearray = []


for subdir, dirs, files in os.walk(out_loc):
    for file in files:
        filep = os.path.join(subdir, file)
	print filep
	#use magic to get the type of the file
	ftn = magic.from_file(filep)
	print ftn
	_conv = ""	
	
	# add type to array if not contained
	if ftn not in typearray:
		typearray.append(ftn)

	if ftn == ft1:
	    	_conv = _2xls
	elif ftn == ft2:
	    	_conv = _2doc
	if _conv != "":
		rstring = "libreoffice --headless "  + _conv + " " + filep + " --outdir " + "\"" + subdir + "\" "
		print rstring
		os.system(rstring)


		# print details of all types found (in case you need to add more)
for _str in typearray:
	print _str


