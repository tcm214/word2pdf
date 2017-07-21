#! python3

# type w2p in to launch from Run.  Built this for Venator.  Automates the process of finding the resume in Downloads, converting to PDF, and saving to Temp


import sys
import os
import comtypes.client
import shutil
import time
import send2trash


def convert_doc(filepath, filename):	#convert doc to pdf and give option to delete doc version
	wdFormatPDF = 17

	pdf_filepath = filepath.split('.')[0] + ".pdf" 	#create and return new downloads filepath for the pdf we're making'
	in_file = os.path.abspath(filepath)
	out_file = os.path.abspath(pdf_filepath)
	out_file = filepath.split('.')[0] + '.pdf' #create pdf filepath by grabbing .doc filepath up to the dot.

	word = comtypes.client.CreateObject('Word.Application')
	doc = word.Documents.Open(in_file)
	doc.SaveAs(out_file, FileFormat=wdFormatPDF)
	doc.Close()
	word.Quit()
	print('converted to pdf...')

	delete_doc(filepath, filename)	# give option to delete doc version

	return pdf_filepath


def get_doc_type(dl_filename, doc_types):

	for doc_type in doc_types: 		#figure out which type of file it is.  Preference to .pdf
		if os.path.isfile('C:\\Users\\tcm21\\Downloads\\{}{}'.format(dl_filename, doc_type)):
			file_type = doc_type
			dl_filepath = 'C:\\Users\\tcm21\\Downloads\\{}{}'.format(dl_filename, doc_type)
			print("Found " + dl_filename + doc_type) 	#print filename w/ appropriate file type
			return (dl_filepath, file_type)
	return ('', '') #return null if no file found



def cls(): #something I found that clears the screen
    os.system("cls")

def quitCheck(entry):
	if entry in ('quit', 'q', 'exit'):
		print('\nquitting program...\n')
		time.sleep(1)
		sys.exit()


def delete_doc(filepath, filename):
	delete_yn = input("Move the .doc to trash (y/n)?: ") #ask if we want to delete the old .doc file since it's also in PDF form now

	if delete_yn == 'y':
		send2trash.send2trash(filepath)
		print('\n\n' + filename + ' moved to Recycle Bin')
	else:
		print("ok we'll keep it")

def get_candidate_name():
	while True:
		while True:
			candidate = input('Name of candidate?: ') # get name of candidate in order to use proper naming convention
			if (len(candidate) > 4) and (' ' in candidate):	#also using this loop to do a dirty validation
				break
			else:
				print('\nThat name looks unusual.  Try a better one.\n')

		temp_filename = candidate + " Resume"
		temp_filepath = 'C:\\Users\\tcm21\\Dropbox\\Temp\\{}.pdf'.format(temp_filename)

		if os.path.isfile(temp_filepath): 	# check if this guy already has a resume pdf in the Temp folder
			print(temp_filename + '.pdf already exists in Temp Folder')
			quit_yn = input('\n\nExit Program? (y/n): ')
			if quit_yn in ['y','yes']:
				sys.exit()
		else:
			return temp_filepath


def fileSearch():
	found_files = []
	filename_input = input('Enter filename: ')
	quitCheck(filename_input)
	for folderName, subfolders, filenames in os.walk('C:\\Users\\tcm21\\Downloads'):
		for filename in filenames:
			if filename_input.lower() in filename.lower():
				found_files.append(filename)
	if not found_files:								# if no matches are found return null
		return
	elif len(found_files) == 1:
		choice = input('This one?: ' + found_files[0] + '?(y/n): ')
		if (choice == 'y') or (choice == 'yes'):
			#print('converted it!!')
			return found_files[0]
	else:															#if there are multiple matches this will print a list and let you pick
		for option in found_files:
			print(str((found_files.index(option))+1) + '. ' + option)
		
		while True:
			try:	
				choice = input('\nwhich one?: ')
				print(found_files[int(choice)-1] + ' selected\n')
				return found_files[int(choice)-1]					#return the correct choice
			except IndexError:
				print('Not in range...')



def getFileType(filename):
	file_type = filename.split('.')[-1] 	# check if the file exists and what type of file it is
	if file_type in ['pdf','doc','docx']:	#check if it's a filetype we can work with
		return file_type
	else:
		print('Not a good filetype. This wont work..\n\n')
		return



doc_types = ['pdf', 'doc', 'docx'] 		#array of filetypes this resume might be
file_type = '' 					# set this to null for now

while True:

	dl_filename = fileSearch() #input filename and this will find it.  
		
	if dl_filename: 	#if found a file, determine the file_type and filepath in Downloads folder
		file_type = getFileType(dl_filename)
		if file_type:
			dl_filepath = 'C:\\Users\\tcm21\\Downloads\\{}'.format(dl_filename)  	# make the filepath in dl folder for later use
			break
	else:					
		print('\nFile not found...')		#do this if no file found.  it'll let you know and clear the screen and start over
		time.sleep(1)		
		cls()								#this clears the screen

		
		
temp_filepath = get_candidate_name() #this function will ask for the candidates name and use that to create the final pdf location using the naming convention

if file_type != 'pdf':						# convert doc to pdf if not pdf already
	dl_filepath = convert_doc(dl_filepath, dl_filename) 	
	
	
shutil.move(dl_filepath, temp_filepath) 	#moves file from Downloads to Temp in Dropbox
print('\n\n' + temp_filepath.split('\\')[-1] + ' created in Temp folder.\n\n')	#that temp_filepath part is just the final filename

	






