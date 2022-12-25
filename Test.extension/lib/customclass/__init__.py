class MItem(object):
    def __init__(self,name):
        self.checked = False
        self.name = name

class RMitem(MItem):
    def __init__(self,name,elemid):
        MItem.__init__(self,name)
        self.elemid = elemid

class RFIItem(RMitem):
    def __init__(self,name,elemid,doc):
        RMitem.__init__(self,name,elemid)
        self.doc = doc
        self.elem = doc.GetElement(self.elemid)
        self.family = self.elem.Symbol.FamilyName