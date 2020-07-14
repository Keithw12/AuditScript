VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} optionForm 
   Caption         =   "Script Options"
   ClientHeight    =   2565
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6945
   OleObjectBlob   =   "optionForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "optionForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cancelButton_Click()
    Unload optionForm
End Sub

Private Sub okButton_Click()
    If queryMemberRadio Then
        Debug.Print "Query Member chosen"
        Unload optionForm
        queryMemberForm.Show
    ElseIf auditUnitRadio Then
        Debug.Print "Audit Unit chosen"
        Unload optionForm
        auditUnitScript.MainMacro
    ElseIf auditAllUnitsRadio Then
        Debug.Print "Audit All Units chosen"
        Unload optionForm
        auditAllUnitsScript.MainMacro
    End If
End Sub
