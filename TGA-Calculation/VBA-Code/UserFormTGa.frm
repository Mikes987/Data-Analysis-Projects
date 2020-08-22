VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormTGa 
   Caption         =   "TGA Calculation"
   ClientHeight    =   5190
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10140
   OleObjectBlob   =   "UserFormTGa.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "UserFormTGa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ButtonLoad_Click()
    Label.Caption = Application.GetOpenFilename("txt-Datei (*.txt), *.txt")
End Sub

Private Sub CommandButtonCancel_Click()
    Unload Me
End Sub

Private Sub CommandButtonStart_Click()
    ' Beginning
    ' First: Check if a file was chosen
    ' Second: Check if any organic ligand has been chosen
    
    ' Variables
    Dim temp As Double
    Dim s As String
    Dim o As Object
    
    If Label.Caption = "" Or Label.Caption = "Falsch" Then
        MsgBox "No file had been chosen, please choose a file"
        Exit Sub
    End If
    If OptionButtonOA.Value = False And OptionButtonSQ.Value = False Then
        MsgBox "Please choose an organic ligand"
        Exit Sub
    ElseIf OptionButtonOA.Value = True Then
        temp = 400
    Else
        temp = 450
    End If
    
    ' Save path
    s = Label.Caption
    
    ' Unload userform
    Unload Me
    
    ' Create a new File
    Call OpenTXT(s, o)
    
    ' Check if A1 is empty
    If o.Range("A1") = "" Then o.Columns("A:A").Delete
    
    ' Do Calculation
    Call Calc(o, temp)
End Sub
