from pptx import Presentation
from docx import Document
# from docx import getdocumenttext

def parseOutline(doc_fn = 'sample.docx', verbose = False):

	occasion = ''
	theme = ''
	venue = ''
	date = ''
	author = 'Bp. Reuben Abante'
	text_verses = []
	text_book = ''

	doc = Document(doc_fn)
	# print(getdocumenttext(document))
	# print('sdlfasf')
	# print(len(doc.paragraphs))

	for para in doc.paragraphs:
		if 'Occasion' in para.text:
			occasion = para.text.split(':')[1].strip()
			# print(occasion)
		elif 'Theme' in para.text:
			theme = para.text.split(':')[1].strip()
		elif 'Venue' in para.text:
			venue = para.text.split(':')[1].strip()
		elif 'Date' in para.text:
			date = para.text.split(':')[1].strip()
		elif 'Title' in para.text:
			title = para.text.split(':')[1].strip()

		# print('para ', para.text)
	if verbose == True:
		print(occasion, theme, venue, date)

	# for para in doc.paragraphs:
	for i in range(len(doc.paragraphs)):
		if 'Text/s' in doc.paragraphs[i].text:
			i+=1
			while(len(doc.paragraphs[i].text)==0):
				i += 1
			text_book = doc.paragraphs[i].text.strip()
			i+=1	
			while(doc.paragraphs[i].text != 'KJV'):
				# print('text: ', doc.paragraphs[i].text)	
				# print('text: ', len(doc.paragraphs[i].text))	
				text_verses.append(doc.paragraphs[i].text.strip())
				if doc.paragraphs[i].text == 'KJV':
					break
				# elif len(doc.paragraphs[i].text) == 0:
					# print('throw')
				i += 1
	# print('text:')
	# print(text_book)
	# for text in text_verses:
	# 	print(text)

	return occasion, theme, venue, date, author, text_book, text_verses
	# for para in doc.paragraphs:
	# 	print('para', para.text)	
def main():
	parseOutline()

if __name__ == '__main__':
	main()