import os
from collections import Counter
from sklearn.naive_bayes import MultinomialNB
from sklearn.model_selection import train_test_split as tts
from sklearn.metrics import accuracy_score
import _pickle as c


def save(clf, name):
	with open(name, 'wb') as fp:
		c.dump(clf, fp)
	print ("saved")



def removeNonAlphabatic(arr):
	words = arr
	for i in range (len(words)):
		if not words[i].isalpha():
			words[i] = ""

	dictionary = Counter(words)
	del dictionary[""]
	return dictionary


def makeDictonary():

	dir = "enron1/emails/"
	files = os.listdir(dir)
	emails = [dir + email for email in files]

	words = []
	c = len(emails)

	for i,email in enumerate(emails):
		with open(email,encoding="iso8859_1") as f:
			text = f.read()
			words += text.split(" ")
			c -= 1
	print (words)
	return removeNonAlphabatic(words).most_common(3000)

#
# convert emails into feature vetor
# 

def makeDataset(dictionary):

	dir = "enron1/emails/"
	files = os.listdir(dir)
	emails = [dir + email for email in files]

	feature_set = []
	labels = []
	c = len(emails)

	for i,email in enumerate(emails):
		data = []
		with open(email,encoding="iso8859_1") as f:
			words = f.read().split(" ")
			for entry in dictionary:
				data.append(words.count(entry[0]))
			feature_set.append(data)
			if "ham" in email:
				labels.append(0)
			if "spam" in email:
				labels.append(1)
			c -= 1
	return feature_set, labels

d = makeDictonary()
features, labels = makeDataset(d)

x_train, x_test, y_train, y_test = tts(features, labels, test_size=0.2)

clf = MultinomialNB()
clf.fit(x_train, y_train)

preds = clf.predict(x_test)

print (accuracy_score(y_test,preds))

save (clf, "text-classifier.mdl")