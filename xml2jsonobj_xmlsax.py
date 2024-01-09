import xml.sax

class XMLHandler(xml.sax.ContentHandler):
    def __init__(self):
        self.CurrentData=""
        self.px=''
        self.py=''
        self.px=''
        self.text=''
        self.size=''
        self.color=''
    def startElement(self, tag, attributes):
        self.CurrentData=tag
        #if tag == " "
