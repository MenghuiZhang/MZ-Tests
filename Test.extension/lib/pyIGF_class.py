# coding: utf8

class tempclass(object):
    def __init__(self):
        self.Name = ''
    @property
    def Name(self):
        return self._Name
    @Name.setter
    def Name(self,value):
        self._Name = value

class tempclass_id(tempclass):
    def __init__(self):
        self.select_id = 0
        tempclass.__init__(self)

    @property
    def select_id(self):
        return self._select_id
    @select_id.setter
    def select_id(self,value):
        self._select_id = value


class tempclass_checked(tempclass):
    def __init__(self):
        self.checked = False
        tempclass.__init__(self)

    @property
    def checked(self):
        return self._checked
    @checked.setter
    def checked(self,value):
        self._checked = value

class tempsystemtypclass(object):
    def __init__(self,systemtyp):
        self.Systemtyp = systemtyp
    @property
    def Systemtyp(self):
        return self._Systemtyp
    @Systemtyp.setter
    def Systemtyp(self,Value):
        self._Systemtyp = Value

class tempsystemclass(tempsystemtypclass):
    def __init__(self,systemtyp,systemname):
        self.Systemname = systemname
        tempsystemtypclass.__init__(self,systemtyp)
    @property
    def Systemname(self):
        return self._Systemname
    @Systemname.setter
    def Systemname(self,Value):
        self._Systemname = Value

# class AuswahlUI(forms.WPFWindow):
#     def __init__(self,windowfile = None,liste = None):
#         self.liste = liste
#         forms.WPFWindow.__init__(self, windowfile)
#         super().__init__()
