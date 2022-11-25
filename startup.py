# coding: utf8

import os
from datetime import datetime
from pyrevit import script,DB
import System
import clr

# guid2 = System.Guid("93382A45-89A9-4cfe-8B94-E0B0D9542D34")
# m_idError = DB.FailureDefinitionId(guid2)
# m_fdError = DB.FailureDefinition.CreateFailureDefinition(m_idError, DB.FailureSeverity.Error, "I am the error")
# m_fdError.AddResolutionType(DB.FailureResolutionType.DeleteElements, "DeleteElements", clr.GetClrType(DB.DeleteElements))
# m_fdError.SetDefaultResolutionType(DB.FailureResolutionType.DeleteElements)

output = script.get_output()

if datetime.today().strftime("%m/%d") == "04/01":
    output.print_image(r"R:\pyRevit\xx_Skripte\_Menghui\Test.extension\Test.gif")
    output.self_destruct(10)