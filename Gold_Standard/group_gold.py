import xml.etree.ElementTree as ET
from collections import defaultdict
import re
import os
import sys

"""Takes all of the completed gold standard files in the completed
folder and puts them all in one big file. Each line is of the format:
MAIN_TAG    SUB_TAG   QUESTION    SENTENCE_TEXT

The border between documents is a blank newline.
"""

def find_speaker(sents,sentence,sent_index,top_level=True):
    sentence = re.sub(r'\[.*?\]','',sentence)
    potentialSpeaker = sentence[:sentence.find(':')].strip()
    if potentialSpeaker != '' and potentialSpeaker != sentence[:-1].strip():
        return (potentialSpeaker,top_level)
    
    previous_sentence_index = sent_index - 1
    return find_speaker(sents,sents[previous_sentence_index],previous_sentence_index,top_level=False)

with open('gold_standard_all.txt','w') as f:
    for dir,dirs,files in os.walk('./completed'):
        for file in files:
            my_xml = ET.parse(os.path.join(dir,file))
            root = my_xml.getroot()
            text = root[0]
            tags = root[1]
            
            sents = re.sub(r'\[s\d+\]\s*','',text.text).strip().split('\n')
            
            for child in tags:
                
                #ignore RENEGES for now (forever?)
                if child.tag == 'RENEGE':
                    continue
                
                #get the sentence
                sent_index = int(child.attrib['text'][1:])
                sentence = sents[sent_index]
                if child.tag == "STAGE_DIRECTIONS":
                    speaker = "stage"
                else:
                    speaker,top_level = find_speaker(sents,sentence,sent_index)
                    if top_level:
                        sentence = sentence[sentence.find(':')+2:]
                        
                
                #get the tag
                tag = child.tag
                
                #get the subtype
                try:
                    subtype = child.attrib['type']
                except KeyError:
                    subtype = ''
                    
                #get the question
                try:
                    question = child.attrib['question']
                except KeyError:
                    question = ''
                    
                f.write("{}\t{}\t{}\t{}\t{}\n".format(tag,subtype,question,sentence,speaker))
            
            f.write('\n')