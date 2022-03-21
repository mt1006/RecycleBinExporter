# Recycle Bin Exporter 1.0
# Author: https://github.com/mt1006

# py -m pip install pypiwin32
# py -m pip install winshell

import os
import errno
import winshell

PATH = 'recycle_bin'
INDEX_FILE_NAME = 'index.csv'
CREATE_INDEX_FILE = True

def copyRecycleBin(path = PATH, indexFileName = INDEX_FILE_NAME,
	createIndexFile = CREATE_INDEX_FILE, consoleOutput = True):
	def getNewName(name):
		retName = name
		if name in filesSet:
			i = 1
			fileName, fileExt = os.path.splitext(name)
			while True:
				newName = f'{fileName}_{i}{fileExt}'
				if newName not in filesSet:
					retName = newName
					break
				i += 1
		filesSet.add(retName)
		return retName

	finalPath = os.path.abspath(path)
	filesSet = set()
	try:
		os.mkdir(path)
	except OSError as error:
		print("Failed to copy recycle bin!")
		if error.errno == errno.EEXIST:
			print(f'Directory "{finalPath}" currently exists...')
		return False

	if createIndexFile:
		try:
			indexFile = open(indexFileName, 'x', encoding='utf-8')
		except OSError as error:
			print("Failed to copy recycle bin!")
			if error.errno == errno.EEXIST:
				print(f'Index file "{indexFileName}" currently exists...')
			return False
	os.chdir(path)
		
	elements = list(winshell.recycle_bin())
	if createIndexFile:
		indexFile.write(f'"Recycle bin files index";"Number of elements: {len(elements)}"\n\n')
		indexFile.write(f'"File name";"Original file name";"Created at";"Deleted at"\n')
	
	for i, element in enumerate(elements):
		curFilePath = element.filename()
		orgFileName = element.original_filename()
		fileName = getNewName(os.path.basename(orgFileName))
		try:
			cTime = str(element.getctime())
		except:
			cTime = '[unknown]'
		winshell.copy_file(curFilePath, fileName)
		if createIndexFile:
			indexFile.write(f'"{fileName}";"{orgFileName}";"{cTime}";"{element.recycle_date()}"\n')
		if consoleOutput:
			print(f'{i + 1}/{len(elements)}')
	
	if consoleOutput:
		print('Copying completed!')
	return True

if __name__ == '__main__':
	print(f'All your files in recycle bin will be copied to {os.path.abspath(PATH)}')
	cont = input('Do you want to continue? [Y/N]:')
	if cont.lower() != 'y':
		exit()
	copyRecycleBin()
	input()