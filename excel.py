
from openpyxl import Workbook
import xml.etree.ElementTree as ET
from collections import defaultdict
import re

class XML_comp():
    def __init__(self):
        self.sents = []
        self.sent_nums = []
        self.tag_dict1 = defaultdict(lambda:" ")
        self.q_dict1 = defaultdict(lambda:" ")
        self.type_dict1 = defaultdict(lambda:" ")
        self.dec_dict1 = defaultdict(lambda:" ")
        self.dec_main_dict1 = defaultdict(lambda:" ")
        self.dec_type_dict1 = defaultdict(lambda:" ")
        self.ren_dict1 = defaultdict(lambda:" ")
        self.tag_dict2 = defaultdict(lambda:" ")
        self.q_dict2 = defaultdict(lambda:" ")
        self.type_dict2 = defaultdict(lambda:" ")
        self.dec_dict2 = defaultdict(lambda:" ")
        self.dec_main_dict2 = defaultdict(lambda:" ")
        self.dec_type_dict2 = defaultdict(lambda:" ")
        self.ren_dict2 = defaultdict(lambda:" ")
        self.wb = Workbook()
    def find_data1(self, file_name):
        f = ET.parse(file_name)
        root = f.getroot()
        text = root[0]
        tags = root[1]
        self.sents = text.text.split('\n')
        for child in tags:
            if child.tag == 'DECISION_MAKING':
                self.dec_dict1[child.attrib['text']] = child.tag + ': ' + child.attrib['type']
                self.dec_main_dict1[child.attrib['text']] = child.tag
                self.dec_type_dict1[child.attrib['text']] = child.attrib['type']
            elif child.tag == 'RENEGE':
                self.ren_dict1[child.attrib['fromText']] = child.attrib['toText']
            else:
                self.tag_dict1[child.attrib['text']] = child.tag
                try:
                    self.q_dict1[child.attrib['text']] = child.attrib['question']
                    self.type_dict1[child.attrib['text']] = child.attrib['type']
                except:
                    pass
        for sent in self.type_dict1:
            if self.type_dict1[sent] == 'retcon':
                self.type_dict1[sent] += ": " + re.findall(r's.*$', self.ren_dict1[sent])[0]
        self.sent_nums = sorted([int(x[1:]) for x in self.tag_dict1.keys()])
    def find_data2(self, file_name):
        f = ET.parse(file_name)
        root = f.getroot()
        text = root[0]
        tags = root[1]
        for child in tags:
            if child.tag == 'DECISION_MAKING':
                self.dec_dict2[child.attrib['text']] = child.tag + ': ' + child.attrib['type']
                self.dec_main_dict2[child.attrib['text']] = child.tag
                self.dec_type_dict2[child.attrib['text']] = child.attrib['type']
            elif child.tag == 'RENEGE':
                self.ren_dict2[child.attrib['fromText']] = child.tag + ': ' + child.attrib['toText']
            else:
                self.tag_dict2[child.attrib['text']] = child.tag
                try:
                    self.q_dict2[child.attrib['text']] = child.attrib['question']
                    self.type_dict2[child.attrib['text']] = child.attrib['type']
                except:
                    pass
        for sent in self.type_dict2:
            if self.type_dict2[sent] == 'retcon':
                self.type_dict2[sent] += ": " + re.findall(r's.*$', self.ren_dict2[sent])[0]
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
        ws.cell(row=2, column=12, value="agree")
        ws.cell(row=2, column=13, value="sent")
        row = 3
        for num in self.sent_nums:
            sent = "s" + str(num)
            if self.tag_dict1[sent] != 'IN-CHARACTER_DIALOGUE' and self.tag_dict1[sent] != 'STAGE_DIRECTIONS':
                ws.cell(row=row, column=1, value=sent)
                ws.cell(row=row, column=2, value=self.tag_dict1[sent])
                ws.cell(row=row, column=3, value=self.type_dict1[sent])
                ws.cell(row=row, column=4, value=self.dec_main_dict1[sent])
                ws.cell(row=row, column=5, value=self.dec_type_dict1[sent])
                ws.cell(row=row, column=6, value=self.q_dict1[sent])
                ws.cell(row=row, column=7, value=self.tag_dict2[sent])
                ws.cell(row=row, column=8, value=self.type_dict2[sent])
                ws.cell(row=row, column=9, value=self.dec_main_dict2[sent])
                ws.cell(row=row, column=10, value=self.dec_type_dict2[sent])
                ws.cell(row=row, column=11, value=self.q_dict2[sent])
                ws.cell(row=row, column=12, value=self.agree(sent))
                ws.cell(row=row, column=13, value=self.sents[num+1])
                self.color(row, sent, ws)
                row += 1
        self.wb.save(file_name)
    def agree(self, sent):
        if self.tag_dict1[sent] == self.tag_dict2[sent] and \
        self.type_dict1[sent] == self.type_dict2[sent] and \
        self.q_dict1[sent] == self.q_dict2[sent] and \
        self.dec_dict1[sent] == self.dec_dict2[sent] and \
        self.ren_dict1[sent] == self.ren_dict2[sent]:
            return "yes"
        else:
            return "no"
    def color(self, row, sent, ws):
        if self.tag_dict1[sent] == self.tag_dict2[sent]:
            ws.cell(row=row, column=2).style = 'Good'
            ws.cell(row=row, column=7).style = 'Good'
        else:
            ws.cell(row=row, column=2).style = 'Bad'
            ws.cell(row=row, column=7).style = 'Bad'
        if self.type_dict1[sent] == self.type_dict2[sent]:
            ws.cell(row=row, column=3).style = 'Good'
            ws.cell(row=row, column=8).style = 'Good'
        else:
            ws.cell(row=row, column=3).style = 'Bad'
            ws.cell(row=row, column=8).style = 'Bad'
        if self.dec_main_dict1[sent] == self.dec_main_dict2[sent]:
            ws.cell(row=row, column=4).style = 'Good'
            ws.cell(row=row, column=9).style = 'Good'
        else:
            ws.cell(row=row, column=4).style = 'Bad'
            ws.cell(row=row, column=9).style = 'Bad'
        if self.dec_type_dict1[sent] == self.dec_type_dict2[sent]:
            ws.cell(row=row, column=5).style = 'Good'
            ws.cell(row=row, column=10).style = 'Good'
        else:
            ws.cell(row=row, column=5).style = 'Bad'
            ws.cell(row=row, column=10).style = 'Bad'
        if self.q_dict1[sent] == self.q_dict2[sent]:
            ws.cell(row=row, column=6).style = 'Good'
            ws.cell(row=row, column=11).style = 'Good'
        else:
            ws.cell(row=row, column=6).style = 'Bad'
            ws.cell(row=row, column=11).style = 'Bad'
        if self.agree(sent):
            ws.cell(row=row, column=12).style = 'Good'
        else:
            ws.cell(row=row, column=12).style = 'Bad'
            
        
if __name__ == '__main__':
    thingy = XML_comp()
    thingy.find_data1('will-20.xml')
    thingy.find_data2('20-annotated.xml')
    thingy.print_to_excel('data2.xlsx')
    