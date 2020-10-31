#at first import PyPDF2
#import xlwt
import xlwt as xl
import PyPDF2 as pdf 

# take the file name as input
# open pdf file using PdfFileReader()
file_name = input('Enter the file name: ')
a = pdf.PdfFileReader(file_name)

x = 0
data_list = [['Model','UPC'],]

for pageNo in range(a.getNumPages()):
#get a page data and extranct thsese using extractText()
	b = a.getPage(pageNo).extractText()
	#Splite this line to make a list.
	c = b.split()
	#initialize 2 variable and a list
	i = 0
	j = 0
	model_upc_list = []

	for word in c:
	# collect model no
		if i ==1:
			model_upc_list.append(word)
			i = 0
			x+=1

	# collect UPC id
		if j == 1:
			model_upc_list.append(word)
			#Enter model and upc in a master list.
			if len(model_upc_list)==2:
				data_list.append(model_upc_list)
			j = 0
			model_upc_list = []

	#check if the value match model:
		if word == 'Model:':
			i = 1

	# check if the value match UPC
		if word =='UPC:':
			j = 1


# ####now create Excel sheet using this master list
#firstly create a excel werkbok/file
wb = xl.Workbook()

# now create a worksheet in this workbook using add_sheet
ws = wb.add_sheet('test sheet')

#once sheet is creted, you can add value in cell 
# follow this way to add insert value
def insert_in_excel(lst):
	for lists in lst:
		for data in lists:
			ws.write(lst.index(lists),lists.index(data),data)
#call the inser_in_excel fuction
insert_in_excel(data_list)

# the last step, saving the work book
i=1
wb.save('newwob{}.xls'.format(i))
### sheet making is done

print('ther are {} models in this book'.format(x))