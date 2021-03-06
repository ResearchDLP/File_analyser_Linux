import _pickle as c
import os
from sklearn import *
from collections import Counter

class 	Classifier:
	def load(clf_file):
		with open(clf_file, "rb") as fp:
			clf = c.load(fp)
			return clf

	def removeNonAlphabatic(self, arr):
		words = arr
		for i in range (len(words)):
			if not words[i].isalpha():
				words[i] = ""

		dictionary = Counter(words)
		del dictionary[""]
		return dictionary


	def makeDictonary(self):

		dir = "text/enron1/emails/"
		files = os.listdir(dir)
		emails = [dir + email for email in files]

		words = []
		c = len(emails)

		for i,email in enumerate(emails):
			with open(email, encoding="iso8859_1") as f:
				text = f.read()
				words += text.split(" ")
				c -= 1

		return self.removeNonAlphabatic(words).most_common(3000)

	clf = load("text/text-classifier.mdl")

	def classifer(self, a):
		clf = self.clf
		d = self.makeDictonary()
		features = []
		inp = a.split()
		for word in d:
			features.append(inp.count(word[0]))
		res = clf.predict([features])
		return (["Not Spam", "Spam!"][res[0]])




		
