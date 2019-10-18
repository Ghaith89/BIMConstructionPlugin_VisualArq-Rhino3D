#======================Get Rhino Referances======================#
import rhinoscriptsyntax as rs
import scriptcontext as sc
import Rhino as R

import clr
import sys


sys.path.append(r"C:\Program Files\Rhino 6\System") 
clr.AddReference ("VisualARQ.Script.dll")
sys.path.append(r"C:\Users\TishGhaith\Downloads\Microsoft.Office.Interop.Excel")
clr.AddReference ("Microsoft.Office.Interop.Excel.dll")

import VisualARQ.Script as va

sys.path.append(r"C:\Users\TishGhaith\Desktop")
from Microsoft.Office.Interop import Excel

#======================Get Rhino Referances========================#

#======================Get Rhino Objects===========================#
guids = rs.GetObjects('select objects')

#======================Get Rhino Objects===========================#
    
#======================Synchronizing To Excel======================#
ex = Excel.ApplicationClass()   
ex.Visible = True
ex.DisplayAlerts = False   
workbook = ex.Workbooks.Open('C:\Users\TishGhaith\Desktop\ZZ.xlsx')
ws = workbook.Worksheets[1]
#======================Synchronizing To Excel======================#


#MakeSure That the previouseStyle is deleted

for i in guids:
    ObjName = rs.ObjectName(i)
    if va.IsElement(i):
        rs.DeleteObject(i)
        ExistedStyles = va.GetGenericElementStyleId(ObjName)
        va.DeleteStyle(ExistedStyles)
    else:
        ExistedStyles = va.GetGenericElementStyleId(ObjName)
        va.DeleteStyle(ExistedStyles)
        
#Make sure that all styles are deleted
ListStyles = va.GetAllGenericElementStyleIds()

for n in ListStyles:
    va.DeleteStyle(n)


"""
M = rs.BlockNames(True)
if len(M)>1:
    for i in M:
        if i == ObjName:
            rs.DeleteBlock(i)
"""
va.IsElement


def redefigningBlocksOrigon(List_Blocks):
    CorrectedBlocks = []
    
    
    
    for strObject in List_Blocks:
        if rs.IsBlockInstance(strObject):
            BlockName = rs.BlockInstanceName(strObject)
            #Getting Block Parts
            BlockParts = rs.ExplodeBlockInstance(strObject)
            #Creating Block Base Point
            BlockRef = R.Geometry.Point3d(0,0,0)
            #Adding The Block Referance
            rs.AddBlock(BlockParts, BlockRef, BlockName, True)
            #Inserting The Block In its Place
            WinBlock = rs.InsertBlock( BlockName, BlockRef, (1,1,1), 0, (0,0,1) )
            #Modify Object Name
            rs.ObjectName(WinBlock, BlockName)
            CorrectedBlocks.append(WinBlock)
    return CorrectedBlocks
    

def CreatingElements(Block, RowNum,numProperties, PropRow):
    #GetObjectName
    ObjName = rs.BlockInstanceName(Block)
    rs.ObjectName(Block, ObjName)
    
    
    #va.DeleteStyle(styleId)
    S = numProperties
    ListPara = []
    
    #AddingParameters
    
    Ws = workbook.Worksheets[1]
    Properties = []
    co = 0
    valRowNum = RowNum
    for i in range(S):
        
        Parameter = Ws.Rows[PropRow].Value2[0,i]
        if Ws.Rows[valRowNum].Value2[0,i] != None :
            Value = Ws.Rows[valRowNum].Value2[0,i]
        else:
            Value = "Non"
        
        print (Value)
        print (Parameter)
        Properties.append(Parameter)
        priceId = va.AddObjectParameter(Block,Parameter, va.ParameterType.Text, "Specifications", "Name" )
        ListPara.append(priceId)
        
        va.SetParameterValue(priceId, Block, Value);
        
    return Block
    
#======================Cleaning Block List======================#
CorrectedBlocks = redefigningBlocksOrigon(guids)
#======================Cleaning Block List======================#


#============================Input==============================#
NumElements = 455
NumberOfProperties = 68
PropertiesRow = 9
ElemNaColNumber = 1
#============================Input==============================#

#======================Applying Methods=========================#
rowNum = 1

for geo in CorrectedBlocks:
    ObjName = rs.BlockInstanceName(geo)
    num = 1
    for l in range(NumElements+10):
        num+=1
        Name = ws.Rows[num].Value2[0,ElemNaColNumber]
        
        if str(Name)==ObjName:
            rowName = Name
            rowNum = num
    CreatingElements(geo, rowNum,NumberOfProperties, PropertiesRow)

#======================Applying Methods=========================#