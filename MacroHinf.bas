Attribute VB_Name = "NewMacros"

Sub Hinf2()
Attribute Hinf2.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Hinf2"
'
' Hinf2 Macro
'
'
    Selection.OMaths.Add Range:=Selection.Range
    Selection.TypeText Text:="H_"
    Selection.InsertSymbol CharacterNumber:=8734, Unicode:=True, Bias:=0
    Selection.OMaths.BuildUp
End Sub
Sub Hinf2_var()
Attribute Hinf2_var.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Hinf2_var"
'
' Hinf2_var Macro
'
'
    Selection.OMaths.Add Range:=Selection.Range
    Selection.InsertSymbol CharacterNumber:=8459, Unicode:=True, Bias:=0
    Selection.TypeText Text:="_"
    Selection.InsertSymbol CharacterNumber:=8734, Unicode:=True, Bias:=0
    Selection.OMaths.BuildUp
End Sub
