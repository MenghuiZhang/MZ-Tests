# coding: utf8
from Autodesk.Revit.DB import IFailuresPreprocessor, FailureProcessingResult,FailureSeverity,FailureProcessingResult

class MyFailuresPreprocessor(IFailuresPreprocessor):
    def __init__(self):
        self.ErrorMessage = ""
        self.ErrorSeverity = ""
 
    def PreprocessFailures(self,failuresAccessor):
        failureMessages = failuresAccessor.GetFailureMessages()
        for failureMessageAccessor in failureMessages:
            Id = failureMessageAccessor.GetFailureDefinitionId()
            try:self.ErrorMessage = failureMessageAccessor.GetDescriptionText()
            except:self.ErrorMessage = "Unknown Error"
            try:
                failureSeverity = failureMessageAccessor.GetSeverity()
                self.ErrorSeverity = failureSeverity.ToString()
                if failureSeverity == FailureSeverity.Warning:failuresAccessor.DeleteWarning(failureMessageAccessor)
                else:return FailureProcessingResult.ProceedWithRollBack
            except:pass
        return FailureProcessingResult.Continue

