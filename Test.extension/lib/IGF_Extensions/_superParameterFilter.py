from IGF_Klasse import DB

class Category_Parameter:
    def __init__(self,param,IsExemplar = True):
        self.param = param
        self.isExemplar = IsExemplar
        self.Id = self.param.Definition.Id
        self.name = self.param.Definition.Name
        self.storygetype = self.param.StorageType.ToString()
    
    @property
    def DisplayUnitType(self):
        if self.storygetype == 'Double':
            try:return self.param.DisplayUnitType
            except:
                try:return self.param.GetUnitTypeId()
                except:return
    
    def createEqualFilter(self,value):
        if self.storygetype != 'Double':
            return DB.ElementParameterFilter(DB.ParameterFilterRuleFactory.CreateEqualsRule(self.Id,value,True))
        try:
            value = float(value)
            value = DB.ConvertToInternalUnits(value,self.storygetype)
            return DB.ElementParameterFilter(DB.ParameterFilterRuleFactory.CreateEqualsRule(self.Id,value,0.1**6))
        except:
            return None
    
    def createBeginsWithFilter(self,value):
        if self.storygetype != 'Double':
            return DB.ElementParameterFilter(DB.ParameterFilterRuleFactory.CreateBeginsWithRule(self.Id,value,True))
        return None
    
    def createContainsFilter(self,value):
        if self.storygetype != 'Double':
            return DB.ElementParameterFilter(DB.ParameterFilterRuleFactory.CreateContainsRule(self.Id,value,True))
        return None
    
    def createEndsWithFilter(self,value):
        if self.storygetype != 'Double':
            return DB.ElementParameterFilter(DB.ParameterFilterRuleFactory.CreateEndsWithRule(self.Id,value,True))
        return None
    
    def createGreaterEqualFilter(self,value):
        if self.storygetype != 'Double':
            return DB.ElementParameterFilter(DB.ParameterFilterRuleFactory.CreateGreaterOrEqualRule(self.Id,value,True))
        try:
            value = float(value)
            value = DB.ConvertToInternalUnits(value,self.storygetype)
            return DB.ElementParameterFilter(DB.ParameterFilterRuleFactory.CreateGreaterOrEqualRule(self.Id,value,0.1**6))
        except:
            return None
        
    def createGreaterFilter(self,value):
        if self.storygetype != 'Double':
            return DB.ElementParameterFilter(DB.ParameterFilterRuleFactory.CreateGreaterRule(self.Id,value,True))
        try:
            value = float(value)
            value = DB.ConvertToInternalUnits(value,self.storygetype)
            return DB.ElementParameterFilter(DB.ParameterFilterRuleFactory.CreateGreaterRule(self.Id,value,0.1**6))
        except:
            return None
    
    def createNoValueFilter(self,value):
        return DB.ElementParameterFilter(DB.ParameterFilterRuleFactory.CreateHasNoValueParameterRule(self.Id))
        
    def createHasValueFilter(self,value):
        return DB.ElementParameterFilter(DB.ParameterFilterRuleFactory.CreateHasValueParameterRule(self.Id))
        
    def createLessEqualFilter(self,value):
        if self.storygetype != 'Double':
            return DB.ElementParameterFilter(DB.ParameterFilterRuleFactory.CreateLessOrEqualRule(self.Id,value,True))
        try:
            value = float(value)
            value = DB.ConvertToInternalUnits(value,self.storygetype)
            return DB.ElementParameterFilter(DB.ParameterFilterRuleFactory.CreateLessOrEqualRule(self.Id,value,0.1**6))
        except:
            return None
    
    def createLessFilter(self,value):
        if self.storygetype != 'Double':
            return DB.ElementParameterFilter(DB.ParameterFilterRuleFactory.CreateLessRule(self.Id,value,True))
        try:
            value = float(value)
            value = DB.ConvertToInternalUnits(value,self.storygetype)
            return DB.ElementParameterFilter(DB.ParameterFilterRuleFactory.CreateLessRule(self.Id,value,0.1**6))
        except:
            return None
    
    def createNotBeginsWithFilter(self,value):
        if self.storygetype != 'Double':
            return DB.ElementParameterFilter(DB.ParameterFilterRuleFactory.CreateNotBeginsWithRule(self.Id,value,True))
        return None
    
    def createNotContainsFilter(self,value):
        if self.storygetype != 'Double':
            return DB.ElementParameterFilter(DB.ParameterFilterRuleFactory.CreateNotContainsRule(self.Id,value,True))
        return None
        
    def createNotEndsWithFilter(self,value):
        if self.storygetype != 'Double':
            return DB.ElementParameterFilter(DB.ParameterFilterRuleFactory.CreateNotEndsWithRule(self.Id,value,True))
        return None
    
    def createNotEqualsFilter(self,value):
        if self.storygetype != 'Double':
            return DB.ElementParameterFilter(DB.ParameterFilterRuleFactory.CreateNotEqualsRule(self.Id,value,True))
        try:
            value = float(value)
            value = DB.ConvertToInternalUnits(value,self.storygetype)
            return DB.ElementParameterFilter(DB.ParameterFilterRuleFactory.CreateNotEqualsRule(self.Id,value,0.1**6))
        except:
            return None