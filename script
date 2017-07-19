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

	pdf_filepath = filepath.split('.')[0] + ".pdf" 	#create and return new downloads filepath for the pdf we just made

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

def quitCheck(filename):
	if (filename == 'quit') or (filename == 'exit') or (filename == 'q'): #q, quit, or exit can quit out here
		print('\nquitting program...\n')
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

		temp_filename = candidate + " Resume"
		temp_filepath = 'C:\\Users\\tcm21\\Dropbox\\Temp\\{}.pdf'.format(temp_filename)

		if os.path.isfile(temp_filepath): 	# check if this guy already has a resume pdf in the Temp folder
			print(temp_filename + '.pdf already exists...')
			quit_yn = input('\n\nExit Program? (y/n): ')
			if quit_yn == 'y':
				sys.exit()
		else:
			return temp_filepath


def fileSearch():
	found_files = []
	filename_input = input('Enter filename: ')
	for folderName, subfolders, filenames in os.walk('C:\\Users\\tcm21\\Downloads'):
		for filename in filenames:
			if filename_input in filename:
				found_files.append(filename)
	if not found_files:
		print('\n\nFile Not Found\n\n')
		return
	elif len(found_files) == 1:
		choice = input('Convert ' + found_files[0] + '?(y/n): ')
		if (choice == 'y') or (choice == 'yes'):
			#print('converted it!!')
			return found_files[0]
	else:
		for option in found_files:
			print(str((found_files.index(option))+1) + '. ' + option)
		
		while True:
			try:	
				choice = input('\nwhich one?: ')
				print(found_files[int(choice)-1] + ' selected\n')
				return found_files[int(choice)-1]
			except IndexError:
				print('Not in range...')




doc_types = ['pdf', 'doc', 'docx'] #array of filetypes this resume might be
file_type = '' # set this to null for now

while True:

	dl_filename = fileSearch() #input filename
	
	file_type = dl_filename.split('.')[-1] 	# check if the file exists and what type of file it is
	dl_filepath = 'C:\\Users\\tcm21\\Downloads\\{}'.format(dl_filename)
	
	if file_type in doc_types:	#if a file was found, there will be a filetype, otherwise say so, and loop back
		break
	else:
		print('\nFile not found...')
		time.sleep(1)
		cls()	#this clears the screen

 



temp_filepath = get_candidate_name() #this function will ask for the candidates name and use that to create the final pdf location using the naming convention





if file_type != '.pdf':						# convert doc to pdf if not pdf already
	dl_filepath = convert_doc(dl_filepath, dl_filename) 	
	
	
shutil.move(dl_filepath, temp_filepath)
print('\n\n' + temp_filepath.split('\\')[-1] + ' created in Temp folder.\n\n')	#that temp_filepath part is just the final filename

	






