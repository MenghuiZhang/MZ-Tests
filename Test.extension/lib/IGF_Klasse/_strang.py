# coding: utf8

class AlleBauteile:
    def __init__(self,Liste,elem):
        self.elem = elem
        self.rohrliste = Liste
        self.liste = []
        self.liste1 = []
        self.rohr = []
        self.formteil = []
        self.t_stueck = None
        for el in self.rohrliste:
            self.liste.append(el)
            self.liste1.append(el)
            self.rohr.append(el)
        
        if self.elem.Category.Id.ToString() == '-2008049':self.formteil.append(self.elem.Id)       
        self.get_t_st(self.elem)
        self.get_t_st_1(self.elem)

    def get_t_st(self,elem):       
        elemid = elem.Id.ToString()
        self.liste.append(elemid)
        cate = elem.Category.Name

        if not cate in ['Rohr Systeme','Rohrdämmung']:
            conns = None
            try:conns = elem.MEPModel.ConnectorManager.Connectors
            except:
                try:conns = elem.ConnectorManager.Connectors
                except:pass
            if conns:                    
                for conn in conns:
                    if conn.IsConnected:
                        refs = conn.AllRefs
                        for ref in refs:
                            owner = ref.Owner
                            if owner.Id.ToString() not in self.liste:  
                                if owner.Category.Id.ToString() == '-2001140':
                                    return     
                                elif owner.Category.Id.ToString() == '-2008044':
                                    self.rohr.append(owner.Id.ToString())
                                elif owner.Category.Id.ToString() == '-2008049':
                                    self.formteil.append(owner.Id)               
                                self.get_t_st(owner)
    
    def get_t_st_1(self,elem):
        elemid = elem.Id.ToString()
        self.liste1.append(elemid)
        if self.t_stueck:return
        cate = elem.Category.Name

        if not cate in ['Rohr Systeme','Rohrdämmung']:
            conns = None
            try:conns = elem.MEPModel.ConnectorManager.Connectors
            except:
                try:conns = elem.ConnectorManager.Connectors
                except:pass
            if conns:
                if conns.Size > 2 and cate == 'Rohrformteile':
                    self.t_stueck = elem.Id
                    return
                    
                for conn in conns:
                    refs = conn.AllRefs
                    for ref in refs:
                        owner = ref.Owner
                        if owner.Id.ToString() not in self.liste1:  
                            if owner.Category.Id.ToString() == '-2008044':
                                if self.t_stueck == None:
                                    self.rohrliste.append(owner.Id.ToString())                         
                            self.get_t_st_1(owner)