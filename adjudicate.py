"""
The Annotation Zone
    Jamie Brandon
    Dan Noar
    Irina Onoprienko
    Will Tietz

Script to adjudicate quickly
 - if two annotators made the same choice, that tag is automatically copied to GS
 - these tags are written the file <episode_number>.xml
 - sentences that were tagged differently by annotators are printed to a file "<episode_number>_to_check.txt"
 - Then, go into MAE and fix these tags yourself. Open both files, then go to File -> Start Adjudication.
 - copy the tags in <episode_number>.xml to the correct section in the GS file (between <TAGS> and </TAGS>)
 - use MAE and the to_check.txt to fix disagreed upon tags
 - once completed, move your files to gold_standard/completed (to avoid chance of overwriting the file)
 - delete the working files in gold_standard

 TODO:
  - weird error where to_check prints sentence numbers twice. WeIrD.

"""
import xml.etree.ElementTree as ET
import re

class adjudicate():
    def __init__(self, infile1, infile2):
        self.sents = []
        self.sent_nums = []

        # dictionaries map sentences to its tag
        self.annotA = self.read_file(infile1)
        self.annotB = self.read_file(infile2)
        self.GS = {}

    def compare_sents(self, error_file):
        # compare the annotators dictionaries
        # if they tagged it the same, copy to GS
        # else, write sentence number to file to be cleaned up
        with open(error_file, "w+") as f:
            for sent_num in self.sent_nums:
                try:
                    if self.annotA[sent_num].tag == self.annotB[sent_num].tag:
                        # if there is no attribute...
                        if self.annotA[sent_num].tag in {'MECHANICS', 'NON-GAME_RELATED', 'NON-CONTENT', 'STAGE_DIRECTIONS', 'IN-CHARACTER_DIALOGUE'}:
                            self.GS[sent_num] = self.annotA[sent_num]
                        else:
                            # if there is an attribute...
                            if self.annotA[sent_num].attrib['type'] == 'retcon' \
                                    or self.annotA[sent_num].attrib['type'] == 'retcon':
                                f.write(sent_num + " retcon: check renege tag\n")
                            else:
                                if self.annotA[sent_num].attrib['type'] == self.annotB[sent_num].attrib['type']:
                                    self.GS[sent_num] = self.annotA[sent_num]
                                else:
                                    f.write(sent_num + " type difference\n")
                    else:
                        # write it to a file so we know to clean this up
                        f.write(sent_num + "\n")

                except KeyError: # if one annotator didn't tag a sentence
                    f.write(sent_num + "\n")


    def write_agreed_tags_to_file(self, outfile):
        # write tags to GS file
        with open(outfile, "w+") as f:
            for k in self.GS:
                f.write(tag_to_string(self.GS[k]))

    def read_file(self, file):
        '''Reads in xml file'''

        # dict maps sentence numbers to their tag : attribute
        # dict['s14'] = "ABOUT_THE_GAME : comment"
        dict = {}

        f = ET.parse(file)
        root = f.getroot()
        text = root[0]      # everything between <TEXT> </TEXT>
        tags = root[1]      # everything between <TAGS> </TAGS>

        self.sents = text.text.split('\n')

        # get sentence numbers
        r = re.compile(r"(\[(s\d+)\])")
        for sent in self.sents:
            try:
                m = r.match(sent)
                self.sent_nums.append(m.group(2))
            except AttributeError: # if it finds an empty line
                pass

        for tag in tags:
            # tag.get("text") is the sentence number: s17
            # tag.tag is the name of the tag: "NARRATION_AND_DESCRIPTION"
            dict[tag.get("text")] = tag

        return dict

def tag_to_string(tag):
    '''Given a tag, output to file something like:
    < ABOUT_THE_GAME id = "A88" spans = "15579~15583" text = "s261" question = "NO" type = "comments" / >'''
    toRet = "<" + tag.tag + " "
    for a in tag.attrib:
        toRet += a + "=\"" + tag.attrib[a] + "\" "
    return toRet + "/>\n"


if __name__ == '__main__':
    # if you trust one annotators type judgements over another's, make the trustworthy one A
    annotA = "Kirsten"
    annotB = "Bingyang"
    episode_number = '35'
    a = adjudicate("Annotators/" + episode_number + "_" + annotA + ".xml",
                   "Annotators/" + episode_number + "_" + annotB + ".xml")
    a.compare_sents("gold_standard/" + episode_number + "_to_check.txt")
    a.write_agreed_tags_to_file("gold_standard/" + episode_number + ".xml")

    # f = ET.parse("Annotators/30_alex.xml")
    # root = f.getroot()
    # text = root[0]  # everything between <TEXT> </TEXT>
    # tags = root[1]
    # print(tag_to_string(tags[0]))
