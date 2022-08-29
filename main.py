from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.options import Options
from bs4 import BeautifulSoup
import requests
from lxml import etree
import re
import time
import docx
from docx.enum.section import WD_ORIENT
from docx.shared import Cm
from Chapter import Chapter
import csv
import concurrent.futures
import os
import shutil



subjectURLDict = {
	'Botany': 'https://www.neetprep.com/questions/53-Botany?courseId=8',
	'Chemistry': 'https://www.neetprep.com/questions/54-Chemistry?courseId=8',
	'Physics': 'https://www.neetprep.com/questions/55-Physics?courseId=8',
	'Zoology': 'https://www.neetprep.com/questions/56-Zoology?courseId=8'
}

class Question:
	def __init__(self, qNo, text, a, b, c, d, e, correct):
		self.qNo = qNo
		self.text = text
		self.a = a
		self.b = b
		self.c = c
		self.d = d
		self.e = e
		self.correct = correct
		
# Lay BeautifulSoup cua website
def getSoup(site):
	options = Options()
	options.headless = True
	driver = webdriver.Firefox(options=options)
	# driver = webdriver.Firefox()
	# driver.maximize_window()
	driver.get(site)
	time.sleep(5)
	soup = BeautifulSoup(driver.page_source,'lxml')
	driver.quit()
	return soup

# Lay BeautifulSoup cua website và click tìm câu trả lời
def getSoupWithAnswer(site):
	options = Options()
	options.headless = True
	driver = webdriver.Firefox(options=options)
	# driver = webdriver.Firefox()
	# driver.maximize_window()
	driver.get(site)
	time.sleep(5)

	# Click để tìm câu trả lời đúng (correct answer)
	answerList = driver.find_elements(By.CLASS_NAME, 'option-list')
	for answer in answerList:
		answer.find_element(By.TAG_NAME, 'input').click()

	soup = BeautifulSoup(driver.page_source,'lxml')
	driver.quit()
	return soup


def mathlmToWord(mathml_string):
	tree = etree.fromstring(mathml_string)
	xslt = etree.parse('MML2OMML.XSL')
	transform = etree.XSLT(xslt)
	new_dom = transform(tree)
	return new_dom.getroot()

def htmlTableToText(tableSoup):
	result = ''
	rows = tableSoup.find_all('tr')
	for row in rows:
		cells = row.find_all('td')
		for cell in cells:
			for element in cell.find_all(text=True): #lấy tất cả element con có text thay bằng text
				element.replace_with(element.text)
			result += cell.text.strip()
			result += '_'
		result = result[:-1]
		result += '~'
	result = result[:-1]
	rows2 = result.split('~')
	if len(rows2) > 1:
		columns1 = rows2[0].split('_')
		columns2 = rows2[1].split('_')
		if len(columns1) != len(columns2):
			result = result[result.index('~')+1:]
	# print(result)
	return result

def writeElementListToCell(_elementList, cell):
	elementList = _elementList
	text = elementList[0]
	# print(question.text[0])
	elementList.pop(0)

	if '<math>' in text or '<img>' in text or '<table>' in text:
		stringList = re.split(r'<math>|<img>|<table>',text) # cắt chuỗi theo các cụm <math>|<img>|<table>
		
		# string trước đó
		preString = '' 
		for string in stringList[:-1]:
			if '.png' in preString or '~' in preString: # nếu string trước đó là ảnh hay bảng thì xuống đoạn khác
				cell.add_paragraph()
			cell.paragraphs[-1].add_run(string) # duyệt các string sau khi cắt chuối text
			element = elementList[0]
			if 'math' in element:
				try:
					cell.paragraphs[-1]._element.append(mathlmToWord(element))
				except:
					cell.paragraphs[-1].add_run(element)
					print('Error formula: ' + element)
					pass
			elif '.png' in element:
				try:
					cell.add_paragraph().add_run().add_picture(element, width = Cm(6.5))
				except:
					cell.add_paragraph().add_run(element)
					print('Error image: ' + element)
					pass
			elif '~' in element:
				textToDocTable(element, cell)

			preString = elementList.pop(0)

		if '.png' in preString or '~' in preString:
			cell.add_paragraph().add_run(stringList[-1])
		else:
			cell.paragraphs[-1].add_run(stringList[-1])
	else:
		cell.text = text

def textToDocTable(text, cell):
	rowsStr = text.split('~')
	# print(str(len(rowsStr)))
	# print(str(len(rowsStr[0].split('_'))))
	table = cell.add_table(rows=len(rowsStr), cols=len(rowsStr[0].split('_')))
	i = j = 0
	for row in rowsStr:
		cellsStr = row.split('_')
		for string in cellsStr:
			# print(str(i) + str(j) + string)
			table.rows[i].cells[j].text = string
			j += 1
		j = 0
		i += 1

def renumberQuestionList(questionList):
	newNo = int(questionList[0].qNo)

	# lấy các câu hỏi đúng vào questionList2
	questionList2 = []
	for question in questionList:
		if question.text[0] == '' or question.a[0] == '' or question.b[0] == '' or question.c[0] == '' or question.d[0] == '':				
			# ghi log file
			with open('logfile.log', 'a') as log:
				log.write(question.qNo + '\n')
		else:
			# đổi số thứ tự câu hỏi
			question.qNo = str(newNo)
			questionList2.append(question)
			newNo += 1
	return questionList2


def writeDocFile(questionList, chapterName):
	doc = docx.Document()
	section = doc.sections[0]
	section.page_height = Cm(29.7)
	section.page_width = Cm(86)
	section.left_margin = Cm(2)
	section.right_margin = Cm(2)
	section.top_margin = Cm(2)
	section.bottom_margin = Cm(2)
	section.orientation = WD_ORIENT.LANDSCAPE
	# new_width, new_height = section.page_height, section.page_width
	# section.page_width = new_width
	# section.page_height = new_height
	table = doc.add_table(rows=1, cols=16)
	for i in range(16):
		if i == 0:
			table.columns[i].width = Cm(2)
		elif i == 1:
			table.columns[i].width = Cm(10)
		else:
			table.columns[i].width = Cm(5)
	row = table.rows[0].cells
	row[0].text = 'Q No'
	row[1].text = 'Qtext'
	row[2].text = 'A'
	row[3].text = 'B'
	row[4].text = 'C'
	row[5].text = 'D'
	row[6].text = 'E'
	row[7].text = 'Correct'
	row[8].text = 'Exp'
	row[9].text = 'Hint'
	row[10].text = 'Level'
	row[11].text = 'Mark'
	row[12].text = 'Time Duration'
	row[13].text = 'Chapter No'
	row[14].text = 'Topic No'
	row[15].text = 'Paragraph Text'

	for question in questionList:
		row = table.add_row().cells
		row[0].text = question.qNo
		print(question.qNo)

		# ghi question text vào cột 1
		writeElementListToCell(question.text, row[1])

		# ghi question a vào cột 2
		writeElementListToCell(question.a, row[2])

		# ghi question b vào cột 3
		writeElementListToCell(question.b, row[3])

		# ghi question c vào cột 4
		writeElementListToCell(question.c, row[4])

		# ghi question d vào cột 5
		writeElementListToCell(question.d, row[5])

		# ghi question e vào cột 6
		writeElementListToCell(question.e, row[6])

		row[7].text = question.correct
		row[8].text = ''
		row[9].text = ''
		row[10].text = ''
		row[11].text = ''
		row[12].text = ''
		row[13].text = ''
		row[14].text = ''
		row[15].text = ''
	
	doc.save(chapterName + '.docx')


def getQuestionsOnePage(pageURL, chapterName):
	print(pageURL)
	soup = getSoupWithAnswer(pageURL)
	# soup = getSoup(pageURL)

	questionList = []

	questionListHtml = soup.find_all('div', class_='question-body')
	for questionHtml in questionListHtml:
		q = Question('', [''], [''], [''], [''], [''], [''], '')

		# get Question No
		q.qNo = questionHtml.find_all('div', class_='question-tag')[0].text.replace('Q', '').replace(':', '').strip()
		print(q.qNo)

		imageCount = 1
		# get Question text and table
		textHtml = questionHtml.find_all('div', class_='question-text')[0].find_all('span')[0].find_all(recursive=False)
		for element in textHtml:

			questionTextList = []

			# lấy table
			if element.name == 'table':
				questionTextList.append(htmlTableToText(element))

			# lấy thẻ p
			elif element.name == 'p':
				# bỏ thẻ p có từ 2 &nbsp; trở lên
				if '\xa0\xa0' in str(element) or '\xa0 \xa0' in str(element):
					element.extract()
					break

				# Lấy tất cả element con trực tiếp của mỗi thẻ p (diagram, mathjax)
				childElements = element.find_all(recursive=False)

				# nếu có element con
				if len(childElements) > 0:
					for child in childElements:
						if child.name == 'span' and child.has_attr('class') and child['class'][0] == 'mjx-chtml': # con là span và class MathJax
							questionTextList.append(child['data-mathml'])		# nối chuỗi MathJax vào questionTextList
							child.replace_with('<math>')	# thay tag span thành <math>

						elif child.name == 'img': # con là img
							try:
								img_data = requests.get(child['src']).content
								with open('Images/' + chapterName + '_' + q.qNo + '_' + str(imageCount) + '.png', 'wb') as handler: # ghi file ảnh
									handler.write(img_data)
								questionTextList.append('Images/' + chapterName + '_' + q.qNo + '_' + str(imageCount) + '.png') # nối path của ảnh vào questionTextList
								child.replace_with('<img>')	# thay tag img thành <img>
								imageCount += 1
							except:
								child.extract()
								print('Error image: ' + q.qNo + '_' + chapterName)
								pass
								
						elif child.name == 'b' or child.name == 'i' or (child.name == 'span' and child.text is not None): 
							child.replace_with(child.text)

						else: #còn lại xóa đi
							child.extract()

				# sau cùng chèn p text vào đầu questionTextList
				if element.text != '':
					questionTextList.insert(0, element.text)

			# nếu không có text trong questionTextList
			if len(questionTextList) == 0:
				continue

			# xét questionText để phân loại là text hoặc các option
			try:
				if '1.' in questionTextList[0] and questionTextList[0].index('1.') == 0:
					questionTextList[0] = questionTextList[0][2:].strip()
					q.a = questionTextList
				elif '(1)' in questionTextList[0] and questionTextList[0].index('(1)') == 0:
					questionTextList[0] = questionTextList[0][3:].strip()
					q.a = questionTextList
				elif '2.' in questionTextList[0] and questionTextList[0].index('2.') == 0:
					questionTextList[0] = questionTextList[0][2:].strip()
					q.b = questionTextList
				elif '(2)' in questionTextList[0] and questionTextList[0].index('(2)') == 0:
					questionTextList[0] = questionTextList[0][3:].strip()
					q.b = questionTextList
				elif '3.' in questionTextList[0] and questionTextList[0].index('3.') == 0:
					questionTextList[0] = questionTextList[0][2:].strip()
					q.c = questionTextList
				elif '(3)' in questionTextList[0] and questionTextList[0].index('(3)') == 0:
					questionTextList[0] = questionTextList[0][3:].strip()
					q.c = questionTextList
				elif '4.' in questionTextList[0] and questionTextList[0].index('4.') == 0:
					questionTextList[0] = questionTextList[0][2:].strip()
					q.d = questionTextList
				elif '(4)' in questionTextList[0] and questionTextList[0].index('(4)') == 0:
					questionTextList[0] = questionTextList[0][3:].strip()
					q.d = questionTextList
				elif '5.' in questionTextList[0] and questionTextList[0].index('5.') == 0:
					questionTextList[0] = questionTextList[0][2:].strip()
					q.e = questionTextList
				elif '(5)' in questionTextList[0] and questionTextList[0].index('(5)') == 0:
					questionTextList[0] = questionTextList[0][3:].strip()
					q.e = questionTextList
				elif '|' in questionTextList[0] and '_' in questionTextList[0]: # nếu là bảng và qText không rỗng
					if q.text[0] != '':
						q.text[0] += '<table>'
						q.text.append(questionTextList[0])
						if '|2._' in questionTextList[0] or '|(2)_' in questionTextList[0] : # nếu là bảng trả lời
							q.a = ['1']
							q.b = ['2']
							q.c = ['3']
							q.d = ['4']
							q.e = ['5']
					else:
						continue
				else:
					q.text[0] += questionTextList[0]
					q.text.extend(questionTextList[1:])	# nối thêm questionTextList với các p textList mới
			except:
				# đặt q.text là list với chuỗi rỗng
				q.text = ['']

				print('Error qText: ' + q.qNo + '_' + chapterName)
				pass

		# get correct answer
		if len(questionHtml.find_all('div', class_='_2eaw _2kqr')) > 0:
			q.correct = questionHtml.find_all('div', class_='_2eaw _2kqr')[0].find_all('label')[0].text
		
		if '1' in q.correct:
			q.correct = 'A'
		elif '2' in q.correct:
			q.correct = 'B'
		elif '3' in q.correct:
			q.correct = 'C'
		elif '4' in q.correct:
			q.correct = 'D'
		elif '5' in q.correct:
			q.correct = 'E'

		questionList.append(q)

	return questionList

def getQuestionsOfChapter(chapter):

	questionList = []
	numberOfPages = int(int(chapter.numberOfQuestions)/10) + 1

	for page in range(numberOfPages):
		questionList.extend(getQuestionsOnePage(chapter.chapterURL + '&pageNo=' + str(page+1), chapter.chapterName))

	writeDocFile(renumberQuestionList(questionList), chapter.chapterName)

#----------------------------------------

if __name__ == '__main__':

	chapterList = []

	with open('Chapters.csv') as file:
		reader = csv.reader(file)
		for row in reader:
			chapter = Chapter(*row)
			if chapter.subjectName == 'Botany':
				print(chapter.chapterName)
				chapterList.append(chapter)


	# getQuestionsOfChapter(chapterList[29])

	# chapterList = [chapterList[11], chapterList[17], chapterList[20], chapterList[23]]

	with concurrent.futures.ThreadPoolExecutor() as executor:
		results = executor.map(getQuestionsOfChapter, chapterList)

	# for i in range(19):
	# 	questionList = getQuestionsOnePage('https://www.neetprep.com/questions/54-Chemistry/674-Chemistry-Everyday-Life?courseId=8&pageNo=' + str(i+1), 'DemoData')
	# 	writeDocFile(renumberQuestionList(questionList), 'DemoData')

	# questionList = getQuestionsOnePage('https://www.neetprep.com/questions/54-Chemistry/674-Chemistry-Everyday-Life?courseId=8&pageNo=5', 'DemoData')
	# writeDocFile(renumberQuestionList(questionList), 'DemoData')