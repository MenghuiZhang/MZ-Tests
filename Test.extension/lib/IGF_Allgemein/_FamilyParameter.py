
# coding: utf8
from Autodesk.Revit.DB import BuiltInParameterGroup
import clr
clr.AddReference("System.Collections")
from System.Collections.Generic import Dictionary


# Dict_DIN_Parameter = {
# 'DIN_m':[True,BuiltInParameterGroup.PG_GEOMETRY], 
# 'DIN_e':[True,BuiltInParameterGroup.PG_GEOMETRY], 
# 'Flange_b':[True,BuiltInParameterGroup.PG_GEOMETRY], 
# 'DIN_l':[True,BuiltInParameterGroup.PG_GEOMETRY], 
# 'DIN_g':[True,BuiltInParameterGroup.PG_GEOMETRY], 
# 'DIN_b':[True,BuiltInParameterGroup.PG_GEOMETRY], 
# 'Flange_a':[True,BuiltInParameterGroup.PG_GEOMETRY], 
# 'DIN_r':[True,BuiltInParameterGroup.PG_GEOMETRY], 
# 'DIN_a':[True,BuiltInParameterGroup.PG_GEOMETRY], 
# 'Flange_d':[True,BuiltInParameterGroup.PG_GEOMETRY], 
# 'DIN_c':[True,BuiltInParameterGroup.PG_GEOMETRY], 
# 'DIN_d':[True,BuiltInParameterGroup.PG_GEOMETRY], 
# 'DIN_f':[True,BuiltInParameterGroup.PG_GEOMETRY], 
# 'Flange_h':[True,BuiltInParameterGroup.PG_GEOMETRY], 
# 'Alpha':[True,BuiltInParameterGroup.PG_GEOMETRY], 
# 'Flange_c':[True,BuiltInParameterGroup.PG_GEOMETRY], 
# 'DIN_h':[True,BuiltInParameterGroup.PG_GEOMETRY], 
# 'Flange_g':[True,BuiltInParameterGroup.PG_GEOMETRY], 
# 'DIN_n':[True,BuiltInParameterGroup.PG_GEOMETRY]
# }

Dict_DIN_Parameter = Dictionary[str, list]()
Dict_DIN_Parameter.Add('Alpha',[True,BuiltInParameterGroup.PG_GEOMETRY])
Dict_DIN_Parameter.Add('DIN_a',[True,BuiltInParameterGroup.PG_GEOMETRY])
Dict_DIN_Parameter.Add('DIN_b',[True,BuiltInParameterGroup.PG_GEOMETRY])
Dict_DIN_Parameter.Add('DIN_c',[True,BuiltInParameterGroup.PG_GEOMETRY])
Dict_DIN_Parameter.Add('DIN_d',[True,BuiltInParameterGroup.PG_GEOMETRY])
Dict_DIN_Parameter.Add('DIN_e',[True,BuiltInParameterGroup.PG_GEOMETRY])
Dict_DIN_Parameter.Add('DIN_f',[True,BuiltInParameterGroup.PG_GEOMETRY])
Dict_DIN_Parameter.Add('DIN_g',[True,BuiltInParameterGroup.PG_GEOMETRY])
Dict_DIN_Parameter.Add('DIN_h',[True,BuiltInParameterGroup.PG_GEOMETRY])
Dict_DIN_Parameter.Add('DIN_l',[True,BuiltInParameterGroup.PG_GEOMETRY])
Dict_DIN_Parameter.Add('DIN_m',[True,BuiltInParameterGroup.PG_GEOMETRY])
Dict_DIN_Parameter.Add('DIN_n',[True,BuiltInParameterGroup.PG_GEOMETRY])
Dict_DIN_Parameter.Add('DIN_r',[True,BuiltInParameterGroup.PG_GEOMETRY])
Dict_DIN_Parameter.Add('Flange_a',[True,BuiltInParameterGroup.PG_GEOMETRY])
Dict_DIN_Parameter.Add('Flange_b',[True,BuiltInParameterGroup.PG_GEOMETRY])
Dict_DIN_Parameter.Add('Flange_c',[True,BuiltInParameterGroup.PG_GEOMETRY])
Dict_DIN_Parameter.Add('Flange_d',[True,BuiltInParameterGroup.PG_GEOMETRY])
Dict_DIN_Parameter.Add('Flange_g',[True,BuiltInParameterGroup.PG_GEOMETRY])
Dict_DIN_Parameter.Add('Flange_h',[True,BuiltInParameterGroup.PG_GEOMETRY])

# Dict_LIN_Parameter = {
# 'LIN_VE_DIM_TYP':[True,BuiltInParameterGroup.INVALID] ,
# 'LIN_VE_EKLIMAX_PARTCODE' :[True,BuiltInParameterGroup.INVALID],
# 'LIN_VE_EKLIMAX_EXPORTED' :[True,BuiltInParameterGroup.INVALID],
# 'LIN_VE_DIM_L':[True,BuiltInParameterGroup.INVALID] ,
# 'LIN_VE_DIM_A':[True,BuiltInParameterGroup.INVALID] ,
# 'LIN_VE_DIM_B':[True,BuiltInParameterGroup.INVALID] ,
# 'LIN_VE_DIM_C':[True,BuiltInParameterGroup.INVALID] ,
# 'LIN_VE_DIM_D':[True,BuiltInParameterGroup.INVALID] ,
# 'LIN_VE_DIM_E':[True,BuiltInParameterGroup.INVALID] ,
# 'LIN_VE_DIM_F':[True,BuiltInParameterGroup.INVALID] ,
# 'LIN_VE_DIM_H':[True,BuiltInParameterGroup.INVALID] ,
# 'LIN_VE_DIM_M':[True,BuiltInParameterGroup.INVALID] ,
# 'LIN_VE_DIM_N':[True,BuiltInParameterGroup.INVALID] ,
# 'LIN_VE_DIM_R':[True,BuiltInParameterGroup.INVALID] ,
# 'LIN_VE_DIM_W':[True,BuiltInParameterGroup.INVALID] ,
# 'LIN_VE_DIM_R1':[True,BuiltInParameterGroup.INVALID] ,
# 'LIN_VE_DIM_R2':[True,BuiltInParameterGroup.INVALID] ,
# 'LIN_VE_DIM_R3':[True,BuiltInParameterGroup.INVALID] ,
# 'LIN_VE_DIM_R4':[True,BuiltInParameterGroup.INVALID] ,
# 'LIN_VE_DIM_T':[True,BuiltInParameterGroup.INVALID] ,
# 'LIN_VE_DIM_U':[True,BuiltInParameterGroup.INVALID] ,
# 'LIN_VE_DIM_D1':[True,BuiltInParameterGroup.INVALID] ,
# 'LIN_VE_DIM_D2':[True,BuiltInParameterGroup.INVALID] ,
# 'LIN_VE_DIM_D3':[True,BuiltInParameterGroup.INVALID] ,
# 'LIN_VE_DIM_I':[True,BuiltInParameterGroup.INVALID] ,
# 'LIN_VE_DIM_L1':[True,BuiltInParameterGroup.INVALID] ,
# 'LIN_VE_DIM_L2':[True,BuiltInParameterGroup.INVALID] ,
# 'LIN_VE_DIM_L3':[True,BuiltInParameterGroup.INVALID] ,
# 'LIN_VE_DIM_L4':[True,BuiltInParameterGroup.INVALID] ,
# 'LIN_VE_DIM_X':[True,BuiltInParameterGroup.INVALID] ,
# 'LIN_VE_DIM_Y':[True,BuiltInParameterGroup.INVALID] ,
# 'LIN_VE_DIM_Y1':[True,BuiltInParameterGroup.INVALID] ,
# 'LIN_VE_DIM_X1':[True,BuiltInParameterGroup.INVALID] ,
# 'LIN_VE_DIM_X2':[True,BuiltInParameterGroup.INVALID] ,
# 'LIN_VE_DIM_Y2':[True,BuiltInParameterGroup.INVALID] ,
# 'LIN_VE_NUM_N':[True,BuiltInParameterGroup.INVALID],
# 'LIN_VE_ANG_W':[True,BuiltInParameterGroup.INVALID]
# }

Dict_LIN_Parameter = Dictionary[str, list]()
Dict_LIN_Parameter.Add('LIN_VE_DIM_TYP',[True,BuiltInParameterGroup.INVALID])
Dict_LIN_Parameter.Add('LIN_VE_EKLIMAX_PARTCODE',[True,BuiltInParameterGroup.INVALID])
Dict_LIN_Parameter.Add('LIN_VE_EKLIMAX_EXPORTED',[True,BuiltInParameterGroup.INVALID])
Dict_LIN_Parameter.Add('LIN_VE_DIM_A',[True,BuiltInParameterGroup.INVALID])
Dict_LIN_Parameter.Add('LIN_VE_DIM_B',[True,BuiltInParameterGroup.INVALID])
Dict_LIN_Parameter.Add('LIN_VE_DIM_C',[True,BuiltInParameterGroup.INVALID])
Dict_LIN_Parameter.Add('LIN_VE_DIM_D',[True,BuiltInParameterGroup.INVALID])
Dict_LIN_Parameter.Add('LIN_VE_DIM_E',[True,BuiltInParameterGroup.INVALID])
Dict_LIN_Parameter.Add('LIN_VE_DIM_F',[True,BuiltInParameterGroup.INVALID])
Dict_LIN_Parameter.Add('LIN_VE_DIM_H',[True,BuiltInParameterGroup.INVALID])
Dict_LIN_Parameter.Add('LIN_VE_DIM_I',[True,BuiltInParameterGroup.INVALID])
Dict_LIN_Parameter.Add('LIN_VE_DIM_L',[True,BuiltInParameterGroup.INVALID])
Dict_LIN_Parameter.Add('LIN_VE_DIM_M',[True,BuiltInParameterGroup.INVALID])
Dict_LIN_Parameter.Add('LIN_VE_DIM_N',[True,BuiltInParameterGroup.INVALID])
Dict_LIN_Parameter.Add('LIN_VE_DIM_R',[True,BuiltInParameterGroup.INVALID])
Dict_LIN_Parameter.Add('LIN_VE_DIM_T',[True,BuiltInParameterGroup.INVALID])
Dict_LIN_Parameter.Add('LIN_VE_DIM_U',[True,BuiltInParameterGroup.INVALID])
Dict_LIN_Parameter.Add('LIN_VE_DIM_W',[True,BuiltInParameterGroup.INVALID])
Dict_LIN_Parameter.Add('LIN_VE_DIM_X',[True,BuiltInParameterGroup.INVALID])
Dict_LIN_Parameter.Add('LIN_VE_DIM_Y',[True,BuiltInParameterGroup.INVALID])
Dict_LIN_Parameter.Add('LIN_VE_DIM_D1',[True,BuiltInParameterGroup.INVALID])
Dict_LIN_Parameter.Add('LIN_VE_DIM_D2',[True,BuiltInParameterGroup.INVALID])
Dict_LIN_Parameter.Add('LIN_VE_DIM_D3',[True,BuiltInParameterGroup.INVALID])
Dict_LIN_Parameter.Add('LIN_VE_DIM_L1',[True,BuiltInParameterGroup.INVALID])
Dict_LIN_Parameter.Add('LIN_VE_DIM_L2',[True,BuiltInParameterGroup.INVALID])
Dict_LIN_Parameter.Add('LIN_VE_DIM_L3',[True,BuiltInParameterGroup.INVALID])
Dict_LIN_Parameter.Add('LIN_VE_DIM_L4',[True,BuiltInParameterGroup.INVALID])
Dict_LIN_Parameter.Add('LIN_VE_DIM_R1',[True,BuiltInParameterGroup.INVALID])
Dict_LIN_Parameter.Add('LIN_VE_DIM_R2',[True,BuiltInParameterGroup.INVALID])
Dict_LIN_Parameter.Add('LIN_VE_DIM_R3',[True,BuiltInParameterGroup.INVALID])
Dict_LIN_Parameter.Add('LIN_VE_DIM_R4',[True,BuiltInParameterGroup.INVALID])
Dict_LIN_Parameter.Add('LIN_VE_DIM_X1',[True,BuiltInParameterGroup.INVALID])
Dict_LIN_Parameter.Add('LIN_VE_DIM_X2',[True,BuiltInParameterGroup.INVALID])
Dict_LIN_Parameter.Add('LIN_VE_DIM_Y1',[True,BuiltInParameterGroup.INVALID])
Dict_LIN_Parameter.Add('LIN_VE_DIM_Y2',[True,BuiltInParameterGroup.INVALID])
Dict_LIN_Parameter.Add('LIN_VE_NUM_N',[True,BuiltInParameterGroup.INVALID])
Dict_LIN_Parameter.Add('LIN_VE_ANG_W',[True,BuiltInParameterGroup.INVALID])

# Dict_MC_Eklimax_Parameter = {
# 'MC eKlimax A' :[True,BuiltInParameterGroup.INVALID] ,
# 'MC eKlimax Article' :[True,BuiltInParameterGroup.PG_DATA],
# 'MC eKlimax B'  :[True,BuiltInParameterGroup.INVALID],
# 'MC eKlimax C':[True,BuiltInParameterGroup.INVALID] ,
# 'MC eKlimax D':[True,BuiltInParameterGroup.INVALID] ,
# 'MC eKlimax E':[True,BuiltInParameterGroup.INVALID] ,
# 'MC eKlimax F':[True,BuiltInParameterGroup.INVALID] ,
# 'MC eKlimax G':[True,BuiltInParameterGroup.INVALID] ,
# 'MC eKlimax H':[True,BuiltInParameterGroup.INVALID] ,
# 'MC eKlimax L':[True,BuiltInParameterGroup.INVALID] ,
# 'MC eKlimax M':[True,BuiltInParameterGroup.INVALID] ,
# 'MC eKlimax N':[True,BuiltInParameterGroup.INVALID] ,
# 'MC eKlimax R':[True,BuiltInParameterGroup.INVALID] ,
# }

Dict_MC_Eklimax_Parameter = Dictionary[str, list]()
Dict_MC_Eklimax_Parameter.Add('MC eKlimax A',[True,BuiltInParameterGroup.INVALID])
Dict_MC_Eklimax_Parameter.Add('MC eKlimax B',[True,BuiltInParameterGroup.INVALID])
Dict_MC_Eklimax_Parameter.Add('MC eKlimax C',[True,BuiltInParameterGroup.INVALID])
Dict_MC_Eklimax_Parameter.Add('MC eKlimax D',[True,BuiltInParameterGroup.INVALID])
Dict_MC_Eklimax_Parameter.Add('MC eKlimax E',[True,BuiltInParameterGroup.INVALID])
Dict_MC_Eklimax_Parameter.Add('MC eKlimax F',[True,BuiltInParameterGroup.INVALID])
Dict_MC_Eklimax_Parameter.Add('MC eKlimax G',[True,BuiltInParameterGroup.INVALID])
Dict_MC_Eklimax_Parameter.Add('MC eKlimax H',[True,BuiltInParameterGroup.INVALID])
Dict_MC_Eklimax_Parameter.Add('MC eKlimax L',[True,BuiltInParameterGroup.INVALID])
Dict_MC_Eklimax_Parameter.Add('MC eKlimax M',[True,BuiltInParameterGroup.INVALID])
Dict_MC_Eklimax_Parameter.Add('MC eKlimax N',[True,BuiltInParameterGroup.INVALID])
Dict_MC_Eklimax_Parameter.Add('MC eKlimax R',[True,BuiltInParameterGroup.INVALID])
Dict_MC_Eklimax_Parameter.Add('MC eKlimax Article',[True,BuiltInParameterGroup.PG_DATA])


# Dict_MC_Parameter = {
# 'MC Length Instance':[True,BuiltInParameterGroup.PG_GEOMETRY] ,
# 'MC Height Instance':[True,BuiltInParameterGroup.PG_GEOMETRY] ,
# 'MC Width Instance':[True,BuiltInParameterGroup.PG_GEOMETRY] ,
# 'MC Diameter Instance':[True,BuiltInParameterGroup.PG_GEOMETRY] ,
# 'MC Water Flow Actual Cold':[False,BuiltInParameterGroup.PG_MECHANICAL] ,
# 'MC Water Flow Actual Hot':[False,BuiltInParameterGroup.PG_MECHANICAL] ,
# 'MC Water Pressure Drop Cold':[False,BuiltInParameterGroup.PG_MECHANICAL] ,
# 'MC Water Pressure Drop Hot':[False,BuiltInParameterGroup.PG_MECHANICAL] ,
# 'MC Discharge Unit':[False,BuiltInParameterGroup.PG_MECHANICAL] ,
# }

Dict_MC_Parameter = Dictionary[str, list]()
Dict_MC_Parameter.Add('MC Length Instance',[True,BuiltInParameterGroup.PG_GEOMETRY])
Dict_MC_Parameter.Add('MC Height Instance',[True,BuiltInParameterGroup.PG_GEOMETRY])
Dict_MC_Parameter.Add('MC Width Instance',[True,BuiltInParameterGroup.PG_GEOMETRY])
Dict_MC_Parameter.Add('MC Diameter Instance',[True,BuiltInParameterGroup.PG_GEOMETRY])
Dict_MC_Parameter.Add('MC Discharge Unit',[False,BuiltInParameterGroup.PG_MECHANICAL])
Dict_MC_Parameter.Add('MC Water Flow Actual Cold',[False,BuiltInParameterGroup.PG_MECHANICAL])
Dict_MC_Parameter.Add('MC Water Flow Actual Hot',[False,BuiltInParameterGroup.PG_MECHANICAL])
Dict_MC_Parameter.Add('MC Water Pressure Drop Cold',[False,BuiltInParameterGroup.PG_MECHANICAL])
Dict_MC_Parameter.Add('MC Water Pressure Drop Hot',[False,BuiltInParameterGroup.PG_MECHANICAL])
Dict_MC_Parameter.Add('MC Cooling Flow',[True,BuiltInParameterGroup.PG_MECHANICAL])
Dict_MC_Parameter.Add('MC Cooling Given Pressure Drop',[True,BuiltInParameterGroup.PG_MECHANICAL])
Dict_MC_Parameter.Add('MC Cooling Given Temperature Difference',[True,BuiltInParameterGroup.PG_MECHANICAL])
Dict_MC_Parameter.Add('MC Piping Flow',[True,BuiltInParameterGroup.PG_MECHANICAL])
Dict_MC_Parameter.Add('MC Piping Given Pressure Drop',[True,BuiltInParameterGroup.PG_MECHANICAL])
Dict_MC_Parameter.Add('MC Piping Given Temperature Difference',[True,BuiltInParameterGroup.PG_MECHANICAL])


# Dict_IGF_Parameter = {
# 'IGF_RLT_DIN_FORM':[True,BuiltInParameterGroup.PG_TEXT] ,
# 'IGF_X_Material':[True,BuiltInParameterGroup.PG_MATERIALS] ,
# 'IGF_X_Material_Text':[True,BuiltInParameterGroup.PG_MATERIALS] ,
# 'IGF_RLT_DIN_F1':[True,BuiltInParameterGroup.INVALID] ,
# 'IGF_RLT_DIN_F2':[True,BuiltInParameterGroup.INVALID] ,
# 'IGF_RLT_DIN_F3':[True,BuiltInParameterGroup.INVALID] ,
# 'IGF_RLT_DIN_F4':[True,BuiltInParameterGroup.INVALID] ,
# 'IGF_RLT_DIN_F5':[True,BuiltInParameterGroup.INVALID] ,
# 'IGF_RLT_DIN_TYP':[True,BuiltInParameterGroup.PG_TEXT],
# 'IGF_RLT_DIN_CODE':[True,BuiltInParameterGroup.PG_TEXT],
# 'IGF_X_Beschreibung':[False,BuiltInParameterGroup.PG_TEXT],
# 'IGF_X_Typ':[False,BuiltInParameterGroup.PG_TEXT],
# 'IGF_X_Größe':[False,BuiltInParameterGroup.PG_TEXT],
# 'IGF_X_Isolations_Zähler':[True,BuiltInParameterGroup.PG_IDENTITY_DATA],
# 'IGF_H_HK-Exponent':[False,BuiltInParameterGroup.PG_MECHANICAL],
# 'IGF_H_HK-Nennleistung':[False,BuiltInParameterGroup.PG_MECHANICAL],
# 'IGF_H_HK-Nennübertemperatur':[False,BuiltInParameterGroup.PG_MECHANICAL]
# }


Dict_IGF_Parameter = Dictionary[str, list]()
Dict_IGF_Parameter.Add('IGF_RLT_DIN_F1', [True, BuiltInParameterGroup.INVALID])
Dict_IGF_Parameter.Add('IGF_RLT_DIN_F2', [True, BuiltInParameterGroup.INVALID])
Dict_IGF_Parameter.Add('IGF_RLT_DIN_F3', [True, BuiltInParameterGroup.INVALID])
Dict_IGF_Parameter.Add('IGF_RLT_DIN_F4', [True, BuiltInParameterGroup.INVALID])
Dict_IGF_Parameter.Add('IGF_RLT_DIN_F5', [True, BuiltInParameterGroup.INVALID])
Dict_IGF_Parameter.Add('IGF_RLT_DIN_TYP', [True, BuiltInParameterGroup.PG_TEXT])
Dict_IGF_Parameter.Add('IGF_RLT_DIN_CODE', [True, BuiltInParameterGroup.PG_TEXT])
Dict_IGF_Parameter.Add('IGF_RLT_DIN_FORM', [True, BuiltInParameterGroup.PG_TEXT])
Dict_IGF_Parameter.Add('IGF_X_Material', [True, BuiltInParameterGroup.PG_MATERIALS])
Dict_IGF_Parameter.Add('IGF_X_Material_Text', [True, BuiltInParameterGroup.PG_MATERIALS])
Dict_IGF_Parameter.Add('IGF_X_Beschreibung', [False, BuiltInParameterGroup.PG_TEXT])
Dict_IGF_Parameter.Add('IGF_X_Typ', [False, BuiltInParameterGroup.PG_TEXT])
Dict_IGF_Parameter.Add('IGF_X_Größe', [False, BuiltInParameterGroup.PG_TEXT])
Dict_IGF_Parameter.Add('IGF_X_Isolations_Zähler', [True, BuiltInParameterGroup.PG_IDENTITY_DATA])
Dict_IGF_Parameter.Add('IGF_H_HK-Exponent', [False, BuiltInParameterGroup.PG_MECHANICAL])
Dict_IGF_Parameter.Add('IGF_H_HK-Nennleistung', [False, BuiltInParameterGroup.PG_MECHANICAL])
Dict_IGF_Parameter.Add('IGF_H_HK-Nennübertemperatur', [False, BuiltInParameterGroup.PG_MECHANICAL])

Dict_IGFGA_Parameter = Dictionary[str, list]()
Dict_IGFGA_Parameter.Add('IGF_GA_Typ', [False, BuiltInParameterGroup.PG_ELECTRICAL])
Dict_IGFGA_Parameter.Add('IGF_GA_Slave_von', [True, BuiltInParameterGroup.PG_ELECTRICAL])
Dict_IGFGA_Parameter.Add('IGF_GA_ID', [True, BuiltInParameterGroup.PG_ELECTRICAL])
Dict_IGFGA_Parameter.Add('IGF_GA_Kabellänge [m]', [True, BuiltInParameterGroup.PG_ELECTRICAL])
Dict_IGFGA_Parameter.Add('IGF_GA_Kabeltyp', [False, BuiltInParameterGroup.PG_ELECTRICAL])
Dict_IGFGA_Parameter.Add('IGF_GA_versorgt durch AN', [True, BuiltInParameterGroup.PG_ELECTRICAL])
Dict_IGFGA_Parameter.Add('IGF_GA_Spannung', [False, BuiltInParameterGroup.PG_ELECTRICAL])
Dict_IGFGA_Parameter.Add('IGF_GA_Leistung-AV', [False, BuiltInParameterGroup.PG_ELECTRICAL])
Dict_IGFGA_Parameter.Add('IGF_GA_Leistung-EN', [False, BuiltInParameterGroup.PG_ELECTRICAL])
Dict_IGFGA_Parameter.Add('IGF_GA_Leistung-SV', [False, BuiltInParameterGroup.PG_ELECTRICAL])
Dict_IGFGA_Parameter.Add('IGF_GA_Lieferant', [False, BuiltInParameterGroup.PG_ELECTRICAL])
Dict_IGFGA_Parameter.Add('IGF_GA_BUS_Datenpunkt_AA', [False, BuiltInParameterGroup.PG_ELECTRICAL])
Dict_IGFGA_Parameter.Add('IGF_GA_BUS_Datenpunkt_AE', [False, BuiltInParameterGroup.PG_ELECTRICAL])
Dict_IGFGA_Parameter.Add('IGF_GA_BUS_Datenpunkt_BA', [False, BuiltInParameterGroup.PG_ELECTRICAL])
Dict_IGFGA_Parameter.Add('IGF_GA_BUS_Datenpunkt_BE', [False, BuiltInParameterGroup.PG_ELECTRICAL])
Dict_IGFGA_Parameter.Add('IGF_GA_BUS_Datenpunkt_BZ', [False, BuiltInParameterGroup.PG_ELECTRICAL])
Dict_IGFGA_Parameter.Add('IGF_GA_Datenpunkt_AA', [False, BuiltInParameterGroup.PG_ELECTRICAL])
Dict_IGFGA_Parameter.Add('IGF_GA_Datenpunkt_AE', [False, BuiltInParameterGroup.PG_ELECTRICAL])
Dict_IGFGA_Parameter.Add('IGF_GA_Datenpunkt_BA', [False, BuiltInParameterGroup.PG_ELECTRICAL])
Dict_IGFGA_Parameter.Add('IGF_GA_Datenpunkt_BE', [False, BuiltInParameterGroup.PG_ELECTRICAL])
Dict_IGFGA_Parameter.Add('IGF_GA_Datenpunkt_BZ', [False, BuiltInParameterGroup.PG_ELECTRICAL])
Dict_IGFGA_Parameter.Add('IGF_GA_Schaltschrank Baugruppe FG', [False, BuiltInParameterGroup.PG_ELECTRICAL])
