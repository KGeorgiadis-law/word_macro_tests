Attribute VB_Name = "NewMacros"
Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Macro1"
'
' Macro1 Macro
'
'
    Selection.EscapeKey
End Sub
Sub SelectFields()
Attribute SelectFields.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.SelectFields"
'
' SelectFields Macro
'
'
With Selection.Find
 .ClearFormatting
 .MatchWildcards = True
 .Text = "\<*\>"
 .Execute Forward:=True
 .Wrap = wdFindAsk
End With
End Sub
