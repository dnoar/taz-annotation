import xml.etree.ElementTree as ET
from collections import defaultdict
import re
import sys

try:
    file_name = sys.argv[1]
    
    f = ET.parse(file_name)
    root = f.getroot()
    text = root[0]      #everything between <TEXT> </TEXT>
    tags = root[1]      #everything between <TAGS> </TAGS>
    sents = text.text.strip().split('\n')
    finalSent = sents[-1]
    sentCount = int(re.search(r'\[s(\d+)\].*',finalSent).groups(1)[0]) + 1
    sent_tags = [0]*sentCount
    first_multi_tag = True
    for child in tags:
        if child.tag == 'RENEGE':
            continue
        sentence = int(child.attrib['text'][1:])
        sent_tags[sentence] += 1
        if sent_tags[sentence] > 1:
            if first_multi_tag:
                print("SENTENCES WITH MORE THAN ONE TAG")
                first_multi_tag = False
            print("s{}".format(sentence))

    first_missing_tag = True
    while 0 in sent_tags:
        if first_missing_tag:
            print("\nSENTENCES WITHOUT TAGS")
            first_missing_tag = False
        missing_sentence = sent_tags.index(0)
        print("s{}".format(missing_sentence))
        sent_tags[missing_sentence] = 1
    

except IndexError:
    print("To run this program, run \"python check_work.py <filename.xml>\"")
except FileNotFoundError:
    print("Could not find that file.\nPlease make sure the path and filename are correct and try again.")
