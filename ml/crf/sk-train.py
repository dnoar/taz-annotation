"""CRF training with scikit-learn API
"""

from itertools import chain

import nltk
import sklearn
import scipy.stats
import re
import string
from sklearn.metrics import make_scorer
from sklearn.cross_validation import cross_val_score
from sklearn.grid_search import RandomizedSearchCV
from nltk.corpus import wordnet as wn


import sklearn_crfsuite
from sklearn_crfsuite import scorers
from sklearn_crfsuite import metrics

#http://sklearn-crfsuite.readthedocs.io/en/latest/tutorial.html

PRONOUN_LIST = set(['i','me','my','mine','you','your','yours','he','him','his','she','her','hers','we','us','our','ours','they','them','their','theirs'])


#Location of training, dev, and test sets
TRAIN_SOURCE = './gold_standard_all_grp_train.txt'
TEST_SOURCE = './gold_standard_all_grp_test.txt'
EVALUATE_OUTPUT = True
OUTPUT_FILE = 'output_sk-train.model'

def get_tuples(filename,sublabel_task=False):
    """Turn a gold file into lists of tuples for CRF processing
    Inputs:
             filename: path to gold file
        sublabel_task: True if we want to predict the subtag as well, False otherwise
    Returns:
        docs: list of list of tuples
    """
    docs = list()

    #TODO: do this with less nesting
    with open(filename) as source:
        #The current document: a list of (token, pos, label) tuples
        current_doc = list()

        for line in source:
        
            #Once we reach the end of a document, add it to the list
            if line.strip() == '':
                docs.append(current_doc)
                current_doc = list()
                continue
                
            #Gather bits of each sentence
            main_tag,sub_tag,question,text,speaker = line.strip().split('\t')
            
            if sublabel_task:
                label = '{}_{}'.format(main_tag[:6],sub_tag)
            else:
                label = main_tag
            
            #Make a tuple of text, speaker, and labels
            #And add to current_doc
            current_doc.append((text,speaker,label))
                        
    #grab the last document if not donne already
    if len(current_doc) > 0:
        docs.append(current_doc)
    return docs

def sent2features(doc, i):
    """Convert sentence to features.
    Inputs:
        sent: list of (sent,speaker,label) tuples
        i: index of particular (sent,speaker,label) in the doc
    Returns:
        features: dict of features
    """
    sent_text = doc[i][0]
    speaker = doc[i][1]
    
    data_sans_punct = ''.join(l for l in sent_text if l not in string.punctuation)
        
    #lower-case and split the sentence
    word_list = data_sans_punct.lower().split()
    
    features = {
        'bias': 1.0,
        'speaker': speaker.lower(),
        'length': len(word_list)
    }
    
    for word_index in range(len(word_list)):
        features["word{}".format(word_index)] = word_list[word_index]
        
    pro_index = 0
    for pro in set(word_list).intersection(PRONOUN_LIST):
        features["pro{}".format(pro_index)] = pro
        pro_index += 1
        
    
    if i > 0:
        sent_text0 = doc[i-1][0]
        data_sans_punct0 = ''.join(l for l in sent_text0 if l not in string.punctuation)
        
        #lower-case and split the sentence
        word_list0 = data_sans_punct0.lower().split()
        '''
        pro_index = 0
        for pro in set(word_list0).intersection(PRONOUN_LIST):
            features["-1:pro{}".format(pro_index)] = pro
            pro_index += 1
        ''' 
        speaker0 = doc[i-1][1]
        features.update({
            '-1:speaker': speaker0.lower(),
            '-1:length': len(word_list0)
        })
    else:
        features['BOD'] = True

    if i < len(doc)-1:
        sent_text1 = doc[i+1][0]
        data_sans_punct1 = ''.join(l for l in sent_text1 if l not in string.punctuation)
        
        word_list1 = data_sans_punct1.lower().split()
        
        pro_index = 0
        for pro in set(word_list1).intersection(PRONOUN_LIST):
            features["+1:pro{}".format(pro_index)] = pro
            pro_index += 1
        
        speaker1 = doc[i+1][1]
        features.update({
            '+1:speaker': speaker1.lower(),
            '+1:length': len(word_list1)
        })
    else:
        features['EOD'] = True

    return features


def doc2features(doc):
    """Convert entire document to features
    Inputs:
        doc: list of (sent,speaker,label) tuples
    Returns:
        list of feature dicts
    """
    return [sent2features(doc, i) for i in range(len(doc))]

def doc2labels(doc):
    """Extract labels from document
    Inputs:
        doc: list of (sent,speaker,label) tuples
    Returns:
        list of labels
    """
    return [label for sent, speaker, label in doc]

def train(X_train, y_train):
    """Train CRF model
    Inputs:
        X_train: list of feature dicts for training set
        y_train: list of labels for training set
    Returns:
        model: trained CRF model
    """
    model = sklearn_crfsuite.CRF(
    algorithm='lbfgs',
    c1=0.1,
    c2=0.1,
    max_iterations=100,
    all_possible_transitions=True
    )
    model.fit(X_train, y_train)
    return model

def evaluate_model(crf, X_dev, y_dev, sub_task):
    """Evaluate the model
    Inputs:
        crf: trained CRF model
        X_dev: list of feature dicts for dev set
        y_dev: list of labels for dev set
    Returns:
        None (prints metrics)
    """

    #Get the labels we're evaluating
    labels = list(crf.classes_)

    #Ignore in-character dialogue and stage directions
    if sub_task:
        labels.remove('IN-CHA_')
        labels.remove('STAGE__')
    else:
        labels.remove('IN-CHARACTER_DIALOGUE')
        labels.remove('STAGE_DIRECTIONS')

    print("Predicting labels")
    y_pred = crf.predict(X_dev)

    #print(y_pred[:10]) #for debugging
    #print(y_dev[:10]) #for debugging

    print("Displaying accuracy")
    metrics.flat_f1_score(y_dev, y_pred,
                      average='weighted', labels=labels)

    sorted_labels = sorted(
        labels,
        key=lambda name: (name[1:], name[0])
    )
    print("Displaying detailed metrics")
    print(metrics.flat_classification_report(
        y_dev, y_pred, labels=sorted_labels, digits=3
    ))
    
    
def output_model(crf,x_dev):
    """Print predicted tags to file, one line per tag with a blank line in between sentences
    Inputs:
        crf: Trained CRF model
        x_dev: List of lists of feature dictionaries of test set
    """
    
    with open(OUTPUT_FILE,'w') as f:
        #Get the labels we're evaluating
        labels = list(crf.classes_)

        #Most labels are 'O', so we ignore,
        #otherwise our scores will seem higher than they actually are.
        labels.remove('O')
        
        y_pred = crf.predict(x_dev)
        
        for sentence in y_pred:
            for label in sentence:
                f.write(label)
                f.write('\n')
            f.write('\n')
                

if __name__ == "__main__":
    print("Converting gold files to sentence tuples...")
    
    sub_task = True
    
    train_docs = get_tuples(TRAIN_SOURCE,sub_task)
    test_docs = get_tuples(TEST_SOURCE,sub_task)

    #print(train_docs[2]) #for debugging
    #print(sent2features(train_docs[0])[0]) #for debugging

    print("Building training and dev sets...")
    X_train = [doc2features(d) for d in train_docs]
    y_train = [doc2labels(d) for d in train_docs]

    X_test = [doc2features(d) for d in test_docs]
    y_test = [doc2labels(d) for d in test_docs]

    print("Training model...")
    crf = train(X_train, y_train)

    if EVALUATE_OUTPUT:
        print("Evaluating model...")
        evaluate_model(crf, X_test, y_test,sub_task)
    else:
        output_model(crf,X_test)
        
