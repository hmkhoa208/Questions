from selenium import webdriver
from selenium.webdriver.common.by import By
from bs4 import BeautifulSoup
import time
import concurrent.futures
import csv

subjectURLDict = {
	'Botany': 'https://www.neetprep.com/questions/53-Botany?courseId=8',
	'Chemistry': 'https://www.neetprep.com/questions/54-Chemistry?courseId=8',
	'Physics': 'https://www.neetprep.com/questions/55-Physics?courseId=8',
	'Zoology': 'https://www.neetprep.com/questions/56-Zoology?courseId=8'
}

class Chapter:
	def __init__(self, chapterName, chapterURL, numberOfQuestions, subjectName):
		self.chapterName = chapterName
		self.chapterURL = chapterURL
		self.numberOfQuestions = numberOfQuestions
		self.subjectName = subjectName

	def __repr__(self):
		return f'Chapter(name={self.chapterName}, numberOfQuestions={self.numberOfQuestions}, subject={self.subjectName})'

	def __iter__(self):
		return iter([self.chapterName, self.chapterURL, self.numberOfQuestions, self.subjectName])
		
# Lay BeautifulSoup cua website
def getSoup(site):
	driver = webdriver.Firefox()
	# driver.maximize_window()
	driver.get(site)
	time.sleep(5)
	soup = BeautifulSoup(driver.page_source,'lxml')
	driver.quit()
	return soup

def getChapterList(subject):
	chapterList = []
	soup = getSoup(subjectURLDict[subject])
	chapterHtmlList = soup.find_all('div', class_='alert alert-primary fade show mt-2 container')[0].find_all('div', class_='pl-4 row')[0].find_all('a')
	for aTag in chapterHtmlList:
		chapter = Chapter(aTag.text.strip(), 'https://www.neetprep.com' + aTag['href'], -1, subject)
		chapter.numberOfQuestions = getNumberOfQuestions(chapter.chapterURL)
		print(chapter)
		chapterList.append(chapter)

	with open('Chapters.csv', 'a') as file:
		writer = csv.writer(file)
		writer.writerows(chapterList)
	
	# return chapterList

# Lay BeautifulSoup và click tìm số lượng câu hỏi cua chapter
def getNumberOfQuestions(chapterURL):
	driver = webdriver.Firefox()
	# driver.maximize_window()
	driver.get(chapterURL)
	time.sleep(5)

	# Click Q No để lấy danh sách câu hỏi
	driver.find_element(By.XPATH, '//*[@id="question-number"]/a').click()

	questionList = driver.find_element(By.XPATH, '//*[@id="questionNoModal"]/div/div/div[2]/ul/div').find_elements(By.TAG_NAME, 'li')

	driver.quit()
	return len(questionList)

if __name__ == '__main__':

	# with concurrent.futures.ThreadPoolExecutor() as executor:
	# 	results = executor.map(getChapterList, subjectURLDict.keys())

	chapterList = []

	with open('Chapters.csv') as file:
		reader = csv.reader(file)
		for row in reader:
			chapterList.append(Chapter(*row))

	for chapter in chapterList:
		print(chapter)

