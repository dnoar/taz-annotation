from docx import Document
from nltk.tokenize import sent_tokenize
import re,os

ORIGINALS_DIR = 'tazscripts - culled'
OUTPUT_DIR = 'tazscripts - processed'
CENSOR = False

def textToRemove(text):
    if text == '':
        return True
    if re.match(r'{.*}',text):
        return True
    if re.fullmatch(r'\w+:',text):
        return True
    return False

def getSpeaker(paragraph,currentSpeaker):
    if paragraph.text[0] == '[' or re.fullmatch(r'\w+: \[[\w\s]+\]',paragraph.text.strip()):
        return 'stage'
    firstRun = paragraph.runs[0]
    if firstRun.text == '\t':
        firstRun = paragraph.runs[1]
    if firstRun.bold:
        potentialSpeaker = re.match(r'[A-Za-z]+',firstRun.text)
        if potentialSpeaker != None:
            return potentialSpeaker.group()
    
    return currentSpeaker
    
def censor(sentence):

    if not CENSOR:
        return sentence
        
    sentence = re.sub(r'fuck','fark',sentence,flags=re.I)
    sentence = re.sub(r'shit','crap',sentence,flags=re.I)
    
    #for damn, skip damnation
    sentence = re.sub(r'damnation','danmation',sentence,flags=re.I)
    sentence = re.sub(r'damn','dang',sentence,flags=re.I)
    sentence = re.sub(r'danmation','damnation',sentence,flags=re.I)
    
    sentence = re.sub(r'bitch','@&$#%',sentence,flags=re.I)
    
    return sentence

if __name__ == '__main__':
    
    filenames = list(os.walk(ORIGINALS_DIR))[0][2]
    
    for file in filenames:
        
        fileBare = file[:file.find('.docx')]
    
        with open(os.path.join(OUTPUT_DIR,fileBare + '.xml'),'w',encoding='utf8') as f:
            f.write('<?xml version="1.0" encoding="UTF-8" ?>\n')
            f.write('<Annotation_Schema_TAZ_v1.0>\n')
            f.write('<TEXT><![CDATA[\n')
            
            doc = Document(os.path.join(ORIGINALS_DIR,file))
            
            tagList = []
            s = 0
            speaker = ''
            namedSpeaker = ''
            index = 1
            
            for para in doc.paragraphs:
            
                #remove any whitespace, replace interrobangs with question marks
                text = re.sub(r'\?\!','?',re.sub(r'\!\?','?',para.text.strip()))
                if textToRemove(text):
                    continue
                
                speaker = getSpeaker(para,namedSpeaker).lower()
                if speaker != 'stage':
                    namedSpeaker = speaker
                
                sentences = sent_tokenize(text)
                
                for sent in sentences:
                    label = 's'+str(s)
                    span = (index+1,index+len(label)+1)
                    s += 1
                    
                    sentPrint = '[' + label + '] ' + censor(sent) + '\n'
                    
                    f.write(sentPrint)
                    index += len(sentPrint)
                    
                    if speaker == 'stage':
                        tagList.append(('STAGE_DIRECTIONS',str(span[0]),str(span[1]),label))
                    
                    elif speaker not in ('griffin','justin','travis','clint'):
                        tagList.append(('IN-CHARACTER_DIALOGUE',str(span[0]),str(span[1]),label))
                        
            
            f.write(']]></TEXT>\n')
            f.write('<TAGS>\n')
            
            stageID = 0
            dialogID = 0
            for type,spanBegin,spanEnd,label in tagList:
                if type == 'STAGE_DIRECTIONS':
                    id = 'S'+str(stageID)
                    stageID += 1
                else:
                    id = 'I'+str(dialogID)
                    dialogID += 1
                f.write('<' + type + ' id="' + id + '" spans="' + spanBegin + '~' + spanEnd + '" text="' + label + '" />\n')
            
            f.write('</TAGS>\n')
            f.write('</Annotation_Schema_TAZ_v1.0>')
