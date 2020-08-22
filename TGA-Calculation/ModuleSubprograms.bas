Attribute VB_Name = "ModuleSubprograms"
Sub OpenTXT(l, o)
    Workbooks.OpenText l, origin:=932, startrow:=1, DataType:=xlDelimited, textqualifier:=xlDoubleQuote, consecutivedelimiter:=True, _
        Tab:=True, semicolon:=False, comma:=False, Space:=True, other:=False, _
        FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1), Array(7, 1)), _
        TrailingMinusNumbers:=True
    
    Set o = ActiveWorkbook.ActiveSheet
End Sub
