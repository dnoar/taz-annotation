
from openpyxl import Workbook
import xml.etree.ElementTree as ET
from collections import defaultdict
import re

class XML_comp():
    def __init__(self):
        self.sents = []
        self.sent_nums = []
        # main tags = 'NARRATION_AND_DESCRIPTION', 'ABOUT_THE_GAME', 'MECHANICS', 'NON-GAME_RELATED', 'NON-CONTENT'
        # main types is subtags from main tags
        # dec main = descision making or not
        # dec type is subtag for decision making
        # decision is both decision making and it's subtag
        dict_types = ['main tag', 'main type', 'question', 'decision', 'dec main', 'dec type', 'renege']
        self.annot1 = {}
        for name in dict_types:
            self.annot1[name] = defaultdict(lambda: " ")
        self.annot2 = defaultdict(lambda: defaultdict(lambda: " "))
        for name in dict_types:
            self.annot2[name] = defaultdict(lambda: " ")
        self.annot3 = defaultdict(lambda: defaultdict(lambda: " "))
        for name in dict_types:
            self.annot3[name] = defaultdict(lambda: " ")
        self.agree = {} #pariwise agreement counter
        for name in dict_types:
            self.agree[name] = 0
        self.wb = Workbook()
    def find_data1(self, file_name):
        f = ET.parse(file_name)
        root = f.getroot()
        text = root[0]
        tags = root[1]
        self.sents = text.text.split('\n')
        for child in tags:
            if child.tag == 'DECISION_MAKING':
                self.annot1['decision'][child.attrib['text']] = child.tag + ': ' + child.attrib['type']
                self.annot1['dec main'][child.attrib['text']] = child.tag
                self.annot1['dec type'][child.attrib['text']] = child.attrib['type']
            elif child.tag == 'RENEGE':
                self.annot1['renege'][child.attrib['fromText']] = child.attrib['toText']
            else:
                self.annot1['main tag'][child.attrib['text']] = child.tag
                try:
                    self.annot1['question'][child.attrib['text']] = child.attrib['question']
                    self.annot1['main type'][child.attrib['text']] = child.attrib['type']
                except:
                    pass
        for sent in self.annot1['main type']:
            if self.annot1['main type'][sent] == 'retcon':
                self.annot1['main type'][sent] += ": " + re.findall(r's.*$', self.annot1['renege'][sent])[0]
        self.sent_nums = sorted([int(x[1:]) for x in self.annot1['main tag'].keys()])
    def find_data2(self, file_name):
        f = ET.parse(file_name)
        root = f.getroot()
        text = root[0]
        tags = root[1]
        for child in tags:
            if child.tag == 'DECISION_MAKING':
                self.annot2['decision'][child.attrib['text']] = child.tag + ': ' + child.attrib['type']
                self.annot2['dec main'][child.attrib['text']] = child.tag
                self.annot2['dec type'][child.attrib['text']] = child.attrib['type']
            elif child.tag == 'RENEGE':
                self.annot2['renege'][child.attrib['fromText']] = child.attrib['toText']
            else:
                self.annot2['main tag'][child.attrib['text']] = child.tag
                try:
                    self.annot2['question'][child.attrib['text']] = child.attrib['question']
                    self.annot2['main type'][child.attrib['text']] = child.attrib['type']
                except:
                    pass
        for sent in self.annot2['main type']:
            if self.annot2['main type'][sent] == 'retcon':
                self.annot2['main type'][sent] += ": " + re.findall(r's.*$', self.annot2['renege'][sent])[0]
    def find_data3(self, file_name):
        f = ET.parse(file_name)
        root = f.getroot()
        text = root[0]
        tags = root[1]
        for child in tags:
            if child.tag == 'DECISION_MAKING':
                self.annot3['decision'][child.attrib['text']] = child.tag + ': ' + child.attrib['type']
                self.annot3['dec main'][child.attrib['text']] = child.tag
                self.annot3['dec type'][child.attrib['text']] = child.attrib['type']
            elif child.tag == 'RENEGE':
                self.annot3['renege'][child.attrib['fromText']] = child.attrib['toText']
            else:
                self.annot3['main tag'][child.attrib['text']] = child.tag
                try:
                    self.annot3['question'][child.attrib['text']] = child.attrib['question']
                    self.annot3['main type'][child.attrib['text']] = child.attrib['type']
                except:
                    pass
        for sent in self.annot3['main type']:
            if self.annot3['main type'][sent] == 'retcon':
                self.annot3['main type'][sent] += ": " + re.findall(r's.*$', self.annot3['renege'][sent])[0]
    def print_to_excel(self, file_name):
        ws = self.wb.active
        ws.title = "Data"
        ws.cell(row=1, column=4, value="annotator 1")
        ws.cell(row=1, column=10, value="annotator 2")
        ws.cell(row=2, column=1, value="sent")
        ws.cell(row=2, column=2, value="tag")
        ws.cell(row=2, column=3, value="type")
        ws.cell(row=2, column=4, value="decision")
        ws.cell(row=2, column=5, value="dec type")
        ws.cell(row=2, column=6, value="question")
        ws.cell(row=2, column=7, value="tag")
        ws.cell(row=2, column=8, value="type")
        ws.cell(row=2, column=9, value="decision")
        ws.cell(row=2, column=10, value="dec type")
        ws.cell(row=2, column=11, value="question")
        ws.cell(row=2, column=12, value="tag")
        ws.cell(row=2, column=13, value="type")
        ws.cell(row=2, column=14, value="decision")
        ws.cell(row=2, column=15, value="dec type")
        ws.cell(row=2, column=16, value="question")
        ws.cell(row=2, column=17, value="agree")
        ws.cell(row=2, column=18, value="sent")
        row = 3
        for num in self.sent_nums:
            sent = "s" + str(num)
            if self.annot1['main tag'][sent] != 'IN-CHARACTER_DIALOGUE' and self.annot1['main tag'][sent] != 'STAGE_DIRECTIONS':
                ws.cell(row=row, column=1, value=sent)
                ws.cell(row=row, column=2, value=self.annot1['main tag'][sent])
                ws.cell(row=row, column=3, value=self.annot1['main type'][sent])
                ws.cell(row=row, column=4, value=self.annot1['dec main'][sent])
                ws.cell(row=row, column=5, value=self.annot1['dec type'][sent])
                ws.cell(row=row, column=6, value=self.annot1['question'][sent])
                ws.cell(row=row, column=7, value=self.annot2['main tag'][sent])
                ws.cell(row=row, column=8, value=self.annot2['main type'][sent])
                ws.cell(row=row, column=9, value=self.annot2['dec main'][sent])
                ws.cell(row=row, column=10, value=self.annot2['dec type'][sent])
                ws.cell(row=row, column=11, value=self.annot2['question'][sent])
                ws.cell(row=row, column=12, value=self.annot3['main tag'][sent])
                ws.cell(row=row, column=13, value=self.annot3['main type'][sent])
                ws.cell(row=row, column=14, value=self.annot3['dec main'][sent])
                ws.cell(row=row, column=15, value=self.annot3['dec type'][sent])
                ws.cell(row=row, column=16, value=self.annot3['question'][sent])
                ws.cell(row=row, column=17, value=self.agree_check(sent))
                ws.cell(row=row, column=18, value=self.sents[num+1])
                self.color(row, sent, ws)
                row += 1
        self.wb.save(file_name)
    def agree_check(self, sent):
        if self.annot1['main tag'][sent] == self.annot2['main tag'][sent] == self.annot3['main tag'][sent] and \
        self.annot1['main type'][sent] == self.annot2['main type'][sent] == self.annot3['main type'][sent] and \
        self.annot1['question'][sent] == self.annot2['question'][sent] == self.annot3['question'][sent] and \
        self.annot1['decision'][sent] == self.annot2['decision'][sent] == self.annot3['decision'][sent]:
            return "yes"
        else:
            return "no"
    def color(self, row, sent, ws):
        self.color_helper('main tag', 2, row, sent, ws)
        self.color_helper('main type', 3, row, sent, ws)
        self.color_helper('dec main', 4, row, sent, ws)
        self.color_helper('dec type', 5, row, sent, ws)
        self.color_helper('question', 6, row, sent, ws)
        if self.agree_check(sent) == "yes":
            ws.cell(row=row, column=17).style = 'Good'
        else:
            ws.cell(row=row, column=17).style = 'Bad'
    def color_helper(self, tag, col, row, sent, ws):
        if self.annot1[tag][sent] == self.annot2[tag][sent] == self.annot3[tag][sent]:
            ws.cell(row=row, column=col).style = 'Good'
            ws.cell(row=row, column=col+5).style = 'Good'
            ws.cell(row=row, column=col+10).style = 'Good'
            self.agree[tag] += 3
        elif self.annot1[tag][sent] == self.annot2[tag][sent]:
            ws.cell(row=row, column=col).style = 'Neutral'
            ws.cell(row=row, column=col+5).style = 'Neutral'
            ws.cell(row=row, column=col+10).style = 'Bad'
            self.agree[tag] += 1
        elif self.annot1[tag][sent] == self.annot3[tag][sent]:
            ws.cell(row=row, column=col).style = 'Neutral'
            ws.cell(row=row, column=col+5).style = 'Bad'
            ws.cell(row=row, column=col+10).style = 'Neutral'
            self.agree[tag] += 1
        elif self.annot2[tag][sent] == self.annot3[tag][sent]:
            ws.cell(row=row, column=col).style = 'Bad'
            ws.cell(row=row, column=col+5).style = 'Neutral'
            ws.cell(row=row, column=col+10).style = 'Neutral'
            self.agree[tag] += 1
        else:
            ws.cell(row=row, column=col).style = 'Bad'
            ws.cell(row=row, column=col+5).style = 'Bad'
            ws.cell(row=row, column=col+10).style = 'Bad'
    def confusion_tags(self, file_name):
        matrix = defaultdict(int)
        total = 0
        for sent in self.annot1['main tag']:
            if self.annot1['main tag'][sent] != 'IN-CHARACTER_DIALOGUE' and self.annot1['main tag'][sent] != 'STAGE_DIRECTIONS':
                matrix[(self.annot1['main tag'][sent], self.annot2['main tag'][sent])] += 1
                total += 1
        ws = self.wb.create_sheet("tables")
        ws.cell(row=1, column=2, value='NARRATION_AND_DESCRIPTION')
        ws.cell(row=1, column=3, value='ABOUT_THE_GAME')
        ws.cell(row=1, column=4, value='MECHANICS')
        ws.cell(row=1, column=5, value='NON-GAME_RELATED')
        ws.cell(row=1, column=6, value='NON-CONTENT')
        ws.cell(row=2, column=1, value='NARRATION_AND_DESCRIPTION')
        ws.cell(row=3, column=1, value='ABOUT_THE_GAME')
        ws.cell(row=4, column=1, value='MECHANICS')
        ws.cell(row=5, column=1, value='NON-GAME_RELATED')
        ws.cell(row=6, column=1, value='NON-CONTENT')
        tags = ['NARRATION_AND_DESCRIPTION', 'ABOUT_THE_GAME', 'MECHANICS', 'NON-GAME_RELATED', 'NON-CONTENT']
        rows = 2
        cols = 2
        for tag1 in tags:
            for tag2 in tags:
                ws.cell(row=rows, column=cols, value=matrix[(tag1, tag2)])
                cols += 1
            rows += 1
            cols = 2
        ws.cell(row=12, column=1, value='NARRATION_AND_DESCRIPTION')
        ws.cell(row=13, column=1, value='ABOUT_THE_GAME')
        ws.cell(row=14, column=1, value='MECHANICS')
        ws.cell(row=15, column=1, value='NON-GAME_RELATED')
        ws.cell(row=16, column=1, value='NON-CONTENT')
        ws.cell(row=11, column=2, value='annot 1')
        rows = 12
        for tag in tags:
            ws.cell(row=rows, column=2, value='=COUNTIF(Data!B:B, "' + tag + '")')
            rows += 1
        ws.cell(row=11, column=3, value='annot 2')
        rows = 12
        for tag in tags:
            ws.cell(row=rows, column=3, value='=COUNTIF(Data!G:G, "' + tag + '")')
            rows += 1
        ws.cell(row=11, column=4, value='annot 3')
        rows = 12
        for tag in tags:
            ws.cell(row=rows, column=4, value='=COUNTIF(Data!L:L, "' + tag + '")')
            rows += 1
        ws.cell(row=11, column=5, value='agreement')
        rows = 12
        for tag in tags:
            ws.cell(row=rows, column=5, value='=sum(B'+str(rows)+':D'+str(rows)+')')
            rows += 1
        ws.cell(row=11, column=6, value='P(tag)')
        rows = 12
        for tag in tags:
            ws.cell(row=rows, column=6, value='=E'+str(rows)+'/'+str(3*total))
            rows += 1
        ws.cell(row=17, column=5, value='chance:')
        ws.cell(row=17, column=6, value='=F12^2+F13^2+F14^2+F15^2+F16^2')
        ws.cell(row=18, column=5, value='observed')
        ws.cell(row=18, column=6, value=self.agree['main tag']/(3*total))
        ws.cell(row=19, column=5, value='κ')
        ws.cell(row=19, column=6, value='=(F18-F17)/(1-F17)')
        self.wb.save(file_name)
        
        
if __name__ == '__main__':
    thingy = XML_comp()
    thingy.find_data2('21-annotated.xml')
    thingy.find_data1('21-D-Will.xml')
    thingy.find_data3('21-Jamie.xml')
    thingy.print_to_excel('data21-D.xlsx')
    thingy.confusion_tags('data21-D.xlsx')
    