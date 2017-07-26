#! python3

# type w2p in to launch from Run.  Built this for Venator.  Automates the process of finding the resume in Downloads, converting to PDF, and saving to Temp


import sys
import os
import comtypes.client
import shutil
import time
import send2trash
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter   #process_pdf
from pdfminer.pdfpage import PDFPage
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
import re, pyperclip
from io import StringIO


def pdf_to_text(dl_filepath):
	'''
	reads pdf and returns the text
	'''
    # PDFMiner boilerplate
    rsrcmgr = PDFResourceManager()
    sio = StringIO()
    codec = 'utf-8'
    laparams = LAParams()
    device = TextConverter(rsrcmgr, sio, codec=codec, laparams=laparams)
    interpreter = PDFPageInterpreter(rsrcmgr, device)
    pdfname = dl_filepath

    # Extract text
    fp = open(pdfname, 'rb')
    for page in PDFPage.get_pages(fp):
        interpreter.process_page(page)
    fp.close()

    # Get text from StringIO
    text = sio.getvalue()

    # Cleanup
    device.close()
    sio.close()

    return text

def find_emails(text):
	'''
	sets regex and find email addresses
	'''
    # regex used to find email address
    emailRegex = re.compile(("([a-z0-9!#$%&'*+\/=?^_`{|}~-]+(?:\.[a-z0-9!#$%&'*+\/=?^_`"
                    "{|}~-]+)*(@|\sat\s)(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?(\.|"
                    "\sdot\s))+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?)"))

    emails = emailRegex.findall(text)
    for email in emails:
        print('Found email: ' + email[0])
        pyperclip.copy(email[0])
    if not emails:
        print('no email found..')



def find_digits(text):
	'''
	sets regex and find phone numbers
	'''

    #regex for phone number
    phoneRegex = re.compile(r'''(
        (\d{3}|\(\d{3}\))             # area code
        (\s|-|\.)?                    # separator
        (\d{3})                       # first 3 digits
        (\s|-|\.)?                    # separator
        (\d{4})                       # last 4 digits
        (\s*(ext|x|ext.)\s*\d{2,5})?  # extension
        )''', re.VERBOSE)

    digits = phoneRegex.findall(text)
    for digit in digits:
        print('Found number: ' + digit[0])
        pyperclip.copy(digit[0])
    if not digits:
        print('no phone number found..')



def convert_doc(filepath, filename):	
	'''
	convert doc to pdf and give option to delete doc version
	'''		
	wdFormatPDF = 17

	pdf_filepath = filepath.split('.')[0] + ".pdf" 	#create and return new downloads filepath for the pdf we're making'
	in_file = os.path.abspath(filepath)
	out_file = os.path.abspath(pdf_filepath)
	out_file = filepath.split('.')[0] + '.pdf' 	#create pdf filepath by grabbing .doc filepath up to the dot.

	word = comtypes.client.CreateObject('Word.Application')
	doc = word.Documents.Open(in_file)
	doc.SaveAs(out_file, FileFormat=wdFormatPDF)
	doc.Close()
	word.Quit()
	print('converted to pdf...')

	delete_doc(filepath, filename)				# give option to delete doc version

	return pdf_filepath


def get_doc_type(dl_filename, doc_types):
	'''
	DEPRECATED
	'''
	for doc_type in doc_types: 							#figure out which type of file it is.  Preference to .pdf
		if os.path.isfile('C:\\Users\\tcm21\\Downloads\\{}{}'.format(dl_filename, doc_type)):
			file_type = doc_type
			dl_filepath = 'C:\\Users\\tcm21\\Downloads\\{}{}'.format(dl_filename, doc_type)
			print("Found " + dl_filename + doc_type) 	#print filename w/ appropriate file type
			return (dl_filepath, file_type)
	return ('', '') 									#return null if no file found



def cls(): 			
	'''
	something I found that clears the screen
    '''
    os.system("cls")

def quitCheck(entry):
	if entry in ('quit', 'q', 'exit'):
		print('\nquitting program...\n')
		time.sleep(1)
		sys.exit()


def delete_doc(filepath, filename):
	'''
	ask if we want to delete the old .doc file since it's also in PDF form now
	'''
	delete_yn = input("Move the .doc to trash (y/n)?: ") 
	if delete_yn == 'y':
		send2trash.send2trash(filepath)
		print('\n\n' + filename + ' moved to Recycle Bin')
	else:
		print("ok we'll keep it")

def get_candidate_name():
	'''
	ask for candidate name to use for naming the new file
	'''
	while True:
		while True:
			candidate = input('Name of candidate?: ') 	# get name of candidate in order to use proper naming convention
			if (len(candidate) > 4) and (' ' in candidate):	#also using this loop to do a dirty validation
				break
			else:
				print('\nThat name looks unusual.  Try a better one.\n')

		temp_filename = candidate + " Resume"
		temp_filepath = 'C:\\Users\\tcm21\\Dropbox\\Temp\\{}.pdf'.format(temp_filename)

		if os.path.isfile(temp_filepath): 				# check if this guy already has a resume pdf in the Temp folder
			print(temp_filename + '.pdf already exists in Temp Folder')
			quit_yn = input('\n\nExit Program? (y/n): ')
			if quit_yn in ['y','yes']:
				sys.exit()
		else:
			return temp_filepath


def fileSearch():
	'''
	asks user for filename. can handle partial matching
	'''
	found_files = []
	filename_input = input('Enter filename: ')
	quitCheck(filename_input)
	for folderName, subfolders, filenames in os.walk('C:\\Users\\tcm21\\Downloads'):
		for filename in filenames:
			if filename_input.lower() in filename.lower():
				found_files.append(filename)
	if not found_files:										# if no matches are found, return null
		return
	elif len(found_files) == 1:
		choice = input('This one?: ' + found_files[0] + '?(y/n): ')
		if (choice == 'y') or (choice == 'yes'):
			return found_files[0]
	else:													#if there are multiple matches this will print a list and let you pick
		for option in found_files:
			print(str((found_files.index(option))+1) + '. ' + option)
		
		while True:
			try:	
				choice = input('\nwhich one?: ')
				print(found_files[int(choice)-1] + ' selected\n')
				return found_files[int(choice)-1]			#return the correct choice
			except IndexError:
				print('Not in range...')



def getFileType(filename):
	'''
	figures out what type of file we're working with
	'''
	file_type = filename.split('.')[-1] 	# check if the file exists and what type of file it is
	if file_type in ['pdf','doc','docx']:	#check if it's a filetype we can work with
		return file_type
	else:
		print('Not a good filetype. This wont work..\n\n')
		return



#doc_types = ['pdf', 'doc', 'docx'] #array of filetypes this resume might be
#file_type = '' # set this to null for now

while True:

	#prompts for filename and looks for it
	dl_filename = fileSearch() 				 
	
	#if found a file, determine the file_type and filepath in Downloads folder
	if dl_filename: 						
		file_type = getFileType(dl_filename)
		if file_type:
			dl_filepath = 'C:\\Users\\tcm21\\Downloads\\{}'.format(dl_filename)  	# make the filepath in dl folder for later use
			break
	else:					
		print('\nFile not found...')		#do this if no file found.  it'll let you know and clear the screen and start over
		time.sleep(1)		
		cls()								#this clears the screen

		
#	this function will ask for the candidates name and use that to create the final pdf location using the naming convention
temp_filepath = get_candidate_name() 		
if file_type != 'pdf':						# convert doc to pdf if not pdf already.  convert_doc also gives the option to delete the old word doc from DL folder
	dl_filepath = convert_doc(dl_filepath, dl_filename) 	



#	scrape text off resume PDF	
resume_text = pdf_to_text(dl_filepath)	

#	finds email and copies to clipboard
find_emails(resume_text)

#	finds phone # and copies to clipboard
find_digits(resume_text)

#	moves file from Downloads to Temp in Dropbox and prints message
shutil.move(dl_filepath, temp_filepath) 	
print('\n\n' + temp_filepath.split('\\')[-1] + ' created in Temp folder.\n\n')	#that temp_filepath part is just the final filename

