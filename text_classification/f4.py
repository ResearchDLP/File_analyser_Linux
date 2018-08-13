import _pickle as c
import os
from sklearn import *
from collections import Counter

def load(clf_file):
	with open(clf_file, "rb") as fp:
		clf = c.load(fp)
		return clf

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
		with open(email, encoding="iso8859_1") as f:
			text = f.read()
			words += text.split(" ")
			c -= 1

	return removeNonAlphabatic(words).most_common(3000)

clf = load("text-classifier.mdl")
d = makeDictonary()

while True:
	features = []
	inp = input(">").split()
	if inp[0] == "exit":
		break
	for word in d:
		features.append(inp.count(word[0]))
	res = clf.predict([features])
	print (["Not Spam", "Spam!"][res[0]])