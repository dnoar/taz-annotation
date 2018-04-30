from __future__ import division

from maxent import MaxEnt
from rando import Rando
from random import shuffle, seed
import re
import time
import string

PRONOUN_LIST = set(['i','me','my','mine','you','your','yours','he','him','his','she','her','hers','we','us','our','ours','they','them','their','theirs'])

class Document():

    def __init__(self, data, label=None):
        self.data = data
        self.label = label
        self.feature_vector = []

    def features(self):
        """Set of words with punctuation and NLTK stopwords removed, and all lower-cased"""
        
        #get rid of punctuation
        data_sans_punct = ''.join(l for l in self.data.split('_')[1] if l not in string.punctuation)
        
        #lower-case and split the review
        word_list = data_sans_punct.lower().split()
        
        #add speaker
        speaker = self.data.split('_')[0]
        word_list.append('speaker={}'.format(speaker))
        
        #add universal bias feature
        word_list.append('bias')
        
        word_list = set(word_list)
        for pro in word_list.intersection(PRONOUN_LIST):
            word_list.add('pronoun={}'.format(pro))
        
        return set(word_list)
        
    def addFeature(self,new_feature):
        """Adds a feature index to this object's feature_vector"""
        self.feature_vector.append(new_feature)

class Corpus():
    
    def __init__(self, corpus_file, include_sublabels=False):
        self.documents = []
        for line in open(corpus_file,'r'):
            stripped_line = line.strip()
            if stripped_line == '':
                continue
            split_line = stripped_line.split('\t')
            
            #grab info
            label = split_line[0]
            if label in ('STAGE_DIRECTIONS','IN-CHARACTER_DIALOGUE'):
                continue
                
            sub_label = split_line[1]
            if include_sublabels:
                label += ':' + sub_label
            question = split_line[2]
            sentence = split_line[3]
            speaker = split_line[4]
            
            self.documents.append(Document("{}_{}".format(speaker,sentence), label))
        
        
    # Act as a mutable container for documents.
    def __len__(self): return len(self.documents)
    def __iter__(self): return iter(self.documents)
    def __getitem__(self, key): return self.documents[key]
    def __setitem__(self, key, value): self.documents[key] = value
    def __delitem__(self, key): del self.documents[key]


    
def accuracy(classifier, test):
    correct = [classifier.classify(x) == x.label for x in test]
    return float(sum(correct)) / len(correct)
'''
class MaxEntTest(TestCase):
    u"""Tests for the MaxEnt classifier."""

    
    def split_names_corpus(self, document_class=Name):
        """Split the names corpus into training, dev, and test sets"""
        names = NamesCorpus(document_class=document_class)
        self.assertEqual(len(names), 5001 + 2943) # see names/README
        seed(hash("names"))
        shuffle(names)
        return (names[:5000], names[5000:6000], names[6000:])

    def test_names_nltk(self):
        """Classify names using NLTK features"""
        train, dev, test = self.split_names_corpus()
        classifier = MaxEnt()
        classifier.train(train, dev)
        acc = accuracy(classifier, test)
        self.assertGreater(acc, 0.70)
    
    def split_review_corpus(self, document_class):
        """Split the yelp review corpus into training, dev, and test sets"""
        reviews = ReviewCorpus('yelp_reviews.json', document_class=document_class)
        seed(hash("reviews"))
        shuffle(reviews)
        return (reviews[:10000], reviews[10000:11000], reviews[11000:14000])

    def test_reviews_bag(self):
        """Classify sentiment using bag-of-words"""
        train, dev, test = self.split_review_corpus(BagOfWords)
        classifier = MaxEnt()
        classifier.train(train, dev)
        self.assertGreater(accuracy(classifier, test), 0.55)
'''   
if __name__ == '__main__':
    docs = Corpus('./gold_standard_all.txt',True)
    seed(time.time())
    shuffle(docs)
    
    first_80 = round(0.8 * len(docs))
    second_10 = first_80 + round(0.1 * len(docs))
    
    train, dev, test = (docs[:first_80],docs[first_80:second_10],docs[second_10:])
    
    classifier = MaxEnt()
    classifier.train(train,dev)
    print(accuracy(classifier,test))
    
