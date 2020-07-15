#An excel class to insert data easily into MSExcel
import win32com.client as win32

class excel():
	def __init__(self, file_name=None):
		'''
		Initialize instances as Excel com clients
		'''
		self.x1App = win32.Dispatch('Excel.Application')
		if file_name:
			self.file_name = file_name
			self.x1Book = self.x1App.Workbooks.Open(file_name) #existing excel file, open it
		else:
			self.x1Book = self.x1App.Workbooks.Add() #add Workbooks
			self.file_name = ''
		

	def save(self,new_file_name= None ):
		'''
		Save File
		'''
		if new_file_name:
			self.file_name = new_file_name #rename file to new filename
			self.x1Book.SaveAs(new_file_name) #first time file, save as
		else:
			self.x1Book.Save() #existing file, save


	def close(self):
		'''
		Close App without saving- assuming that you saved if you intended to
		'''
		self.x1Book.Close(SaveChanges=0)
		del self.x1App

	#methods for setting and setting data in cells. Can specify sheetname or index, row, and column
	def getCell(self, sheet=self.Workbooks.Add().ActiveSheet, row, col):
		'''
		Get value of a single cell
		'''
		sht = self.x1Book.Worksheets(sheet)
		return sht.Cells(row,col).Value


	def setCell(self, sheet, row, col, value):
		'''
		Set value of a single cell
		'''
		sht = self.x1Book.Worksheets(sheet)
		sht.Cells(row,col).Value = value
		return sht.Cells(row,col)


	def getRange(self, sheet, row1, col1, row2, col2):
		'''
		Get Values within specified Range and return as a list of lists
		'''
		sht = self.x1Book.Worksheets(sheet)
		result =  list (sht.Range(sht.Cells(row1,col1), sht.Cells(row2,col2).Value))
		for i in result:
			i = list(i)
		return result


	def setRange(self, sheet, leftCol, topRow, data):
		'''
		Insert values in a 2d array starting at specified location
		Works out the size needed by itself
		'''
		rightCol = leftCol + len(data[0]-1) #computes the ending location
		sht = self.x1Book.Worksheets(sheet)
		sht.Range(
			sht.Cells(topRow, leftCol), 
			sht.Cells(bottomRow, rightCol)
		).Value = data


	def getContigousRange(self, sheet, row,col):
		'''
		Tracks down and across from top left cell until it encounters
		blank cells; returns element in non-blank range.
		Looks at first row and column; blanks at bottom or right are 
		Ok and returns None within the array
		'''
		sht = self.x1Book.Worksheets(sheet)
		bottom = row
		#find bottom row
		while sht.Cells(bottom + 1, col).Value not in [None, '']:
			bottom +=1 #move down to the next row
		right = col #find right column
		while sht.Cells(row, right+1).Value not in [None, '']:
			right+=1
		return sht.Range(sht.Cells(row, col), sht.Cells(bottom, right)).Value


	    
	def cleanStringsAndDates(self, matrix):
		'''
		Convert Unicode Strings to ordinary strings and COM dates
		to int. Cleans up on a column by column basis
		'''
		new_matrix = []
		for row in matrix:
			new_row = []
			for column in row:
				if type(column) is UnicodeType:
					new_row+=[str(column)]
				elif type(column) is TimeType:
					new_row+=[int(column)]
				else:
					new_row+=[cell]
		new_matrix.append(list(new_row))
		return new_matrix



		
		
		
		
