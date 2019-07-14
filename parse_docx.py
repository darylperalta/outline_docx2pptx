from pptx import Presentation
from docx import Document
# from docx import getdocumenttext

# bibleText = {
# 	"text" = ''
# 	"verses" = []
# }

bibleBooks = [
    "Genesis",
    "Gen",

    "Exodus",
    "Exo",

    "Leviticus",
    "Lev ",

    "Numbers",

    "Num ",

    "Deuteronomy",
    "Deut ",

    "Joshua",
    "Judges",
    "Ruth",
    "1 Samuel",
    "2 Samuel",
    "1 Sam",
    "2 Sam",
    "1 Kings",
    "2 Kings",
    "1 Chronicles",
    "2 Chronicles",
    "1 Chron ",
    "2 Chron ",
    "Ezra",
    "Nehemiah",
    "Esther",
    "Job",
    "Psalms",
    "Ps",
    "Proverbs",
    "Prov",
    "Ecclesiastes",
    "Song of Solomon",
    "Isaiah",
    "Jeremiah",
    "Lamentations",
    "Ezekiel",
    "Daniel",
    "Hosea",
    "Joel",
    "Amos",
    "Obadiah",
    "Jonah",
    "Micah",
    "Nahum",
    "Habakkuk",
    "Zephaniah",
    "Haggai",
    "Zechariah",
    "Malachi",
    "Matthew",
    "Matt",
    "Mark",
    "Luke",
    "John",
    "Acts",
    "Romans",
    "Rom",
    "1 Corinthians",
    "1 Cor",
    "2 Corinthians",
    "2 Cor",    
    "Galatians",
    "Gal",
    "Ephesians",
    "Eph",

    "Philippians",
    "Phil",
    "Colossians",
    "Col",

    "1 Thessalonians",
    "2 Thessalonians",
    "1 Thess",
    "2 Thess",
    "1 Timothy",
    "2 Timothy",
    "1 Tim",
    "2 Tim",
    "Titus",
    "Philemon",
    "Hebrews",
    "Heb",

    "James",
    "1 Peter",
    "2 Peter",
    "1 Pet",
    "2 Pet",
    "1 John",
    "2 John",
    "3 John",
    "Jude",
    "Revelation",
    "Rev"

]

def checkBible(text):
	for book in bibleBooks:
		if (text in book) or (book in text[0:15]):
			return True

	return False

def checkPoint(text):
	if len(text) == 0:
		return False
	elif text[0].isupper() and text[1] ==')':
		return True
	return False


def parseOutline(doc_fn = 'sample.docx', verbose = False):
	title = ''
	occasion = ''
	theme = ''
	venue = ''
	date = ''
	author = 'Bp. Reuben Abante'
	text_verses = []
	text_book = ''
	temp_book = ''
	temp_verses = []
	bibleReading = [] # {'book': 'verses'}
	message = []
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
		elif 'TITLE' in para.text:
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
				while(len(doc.paragraphs[i].text)==0):
					i += 1
				# print('text: ', doc.paragraphs[i].text)	
				# print('text: ', len(doc.paragraphs[i].text))	
				text_verses.append(doc.paragraphs[i].text.strip())
				if doc.paragraphs[i].text == 'KJV':
					break
				# elif len(doc.paragraphs[i].text) == 0:
					# print('throw')
				i += 1
	# for i in range(len(doc.paragraphs)):
	i = 0
	while(i < len(doc.paragraphs)):
		# print(len(doc.paragraphs))
		if 'Bible Reading' in doc.paragraphs[i].text:
			i+=1
			if i >= len(doc.paragraphs):
				break

			while(len(doc.paragraphs[i].text)==0):
				i += 1
				if i >= len(doc.paragraphs):
					break
			while(1):
				temp_book = doc.paragraphs[i].text.strip()
				# print('temp book', temp_book)
				i+=1
				if i >= len(doc.paragraphs):
					break	
				while((doc.paragraphs[i].text != 'KJV') and (doc.paragraphs[i].text != 'AMP') and (doc.paragraphs[i].text != 'BBE')):
					while(len(doc.paragraphs[i].text)==0):
						i += 1
						if i >= len(doc.paragraphs):
							break
					# print('text: ', doc.paragraphs[i].text)	
					# print('text: ', len(doc.paragraphs[i].text))	
					temp_verses.append(doc.paragraphs[i].text.strip())
					# if (doc.paragraphs[i].text != 'KJV') or (doc.paragraphs[i].text != 'AMP') and (doc.paragraphs[i].text != 'BBE'):
					# 	break
					i += 1
					if i >= len(doc.paragraphs):
						break
				bibleReading.append({'book':temp_book, 'verses':temp_verses})
				temp_verses = []
				i += 1
				if i >= len(doc.paragraphs):
					break
				while(len(doc.paragraphs[i].text)==0):
					i += 1
				if i >= len(doc.paragraphs):
					break
				if 'Song/s' in doc.paragraphs[i].text:
					break
		elif 'Song/s' in doc.paragraphs[i].text:
			break
		i+=1
		if i >= len(doc.paragraphs):
			break
		# print(doc.paragraphs[i].text)
	# print(i)
	# print()
	i = 0
	temp_verses = []
	while(i < len(doc.paragraphs)):
		# print(i)
		if 'INTRODUCTION' in doc.paragraphs[i].text:

			# print('yeah')
			while(len(doc.paragraphs[i].text)==0):
				# print('i', i)
				i += 1
				if i >= len(doc.paragraphs):
					break
			if i >= len(doc.paragraphs):
				break
			# print(doc.paragraphs[i].text)
			# print(doc.paragraphs[i+1].text)
			# print(doc.paragraphs[i+2].text)
			# print(doc.paragraphs[i+3].text)
			# print(doc.paragraphs[i+4].text)

			# break
			while(i < len(doc.paragraphs)):
				if (checkBible(doc.paragraphs[i].text.strip()) and (len(doc.paragraphs[i].text.strip()) != 0)):
					# print('true bible')
					# print(doc.paragraphs[i].text)
					# print(len(doc.paragraphs[i].text))
					# print('wjat')
					if ('(KJV)' in (doc.paragraphs[i].text)):
						temp_text = doc.paragraphs[i].text.strip()
						# print('do something')
						a = temp_text.split('(KJV)')
						if len(a[1]) != 0:
							temp_book = a[0]
							temp_verses.append(a[1])
						else:
							temp_book = a[0]
							i+=1
							while (len(doc.paragraphs[i].text)!=0):
								temp_verses.append(doc.paragraphs[i].text.strip())
								temp_verses = []
								i+=1
								if i >= len(doc.paragraphs):
									break

					else:
						temp_book = doc.paragraphs[i].text.strip()
						i+=1
						while((doc.paragraphs[i].text != 'KJV') and (doc.paragraphs[i].text != 'AMP') and (doc.paragraphs[i].text != 'BBE')):
							while(len(doc.paragraphs[i].text)==0):
								i += 1
								if i >= len(doc.paragraphs):
									break
							# print('text: ', doc.paragraphs[i].text)	
							# print('text: ', len(doc.paragraphs[i].text))	
							if i >= len(doc.paragraphs):
								break
							temp_verses.append(doc.paragraphs[i].text.strip())
							# if (doc.paragraphs[i].text != 'KJV') or (doc.paragraphs[i].text != 'AMP') and (doc.paragraphs[i].text != 'BBE'):
							# 	break
							i += 1
							if i >= len(doc.paragraphs):
								break

					message.append({'type': 'bible', 'book':temp_book, 'verses':temp_verses})
					temp_verses = []
				elif (((doc.paragraphs[i].style.name == 'List Paragraph') or (doc.paragraphs[i].style.name == 'List Numbers'))  and (len(doc.paragraphs[i].text.strip()) != 0)):
					# print('True point')
					point = doc.paragraphs[i].text.strip()
					# print('point', point, len(doc.paragraphs[i].text.strip()))
					message.append({'type': 'point', 'text': point})

				# else:
				# 	print(doc.paragraphs[i].text.strip())
				# 	print(doc.paragraphs[i].style.name)
				i+=1
		i+=1
	# print('message:::   !!! ')
	# print(message)
	# print(len(message))

	# print('bible reading', len(bibleReading))


	# print('text:')
	# print(text_book)
	# for text in text_verses:
	# 	print(text)

	return title, occasion, theme, venue, date, author, text_book, text_verses, bibleReading, message
	# for para in doc.paragraphs:
	# 	print('para', para.text)	
def main():
	parseOutline()
	# for i in range(5):
	# i = 0
	# while(i < 5):
	# 	print(i)
	# 	i+=1
	# 	print(i)
	# 	i+=1
	# 	if i>= 5:
	# 		break
	# 	print(i)
	# 	i+=1

	# some_text = []
	# verse1 = []
	# verse2 = []
	# verse1.append('verse1a')
	# verse1.append('verse1b')
	# verse2.append('verse2a')
	# verse2.append('verse2b')
	# verse1.append('verse1c')
	# # verse1 = 'verse'
	# some_text.append({'text':'text1','verses':verse1})
	# some_text.append({'text':'text2','verses':verse2})
	# # print(some_text)
	# for text in some_text:
 #  		print(text['text'], text['verses'])


if __name__ == '__main__':
	main()