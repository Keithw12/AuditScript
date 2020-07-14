VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} queryMemberForm 
   Caption         =   "Query Member"
   ClientHeight    =   2040
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7515
   OleObjectBlob   =   "queryMemberForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "queryMemberForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cancelButton_Click()
    Unload queryMemberForm
End Sub

Private Sub viewButton_Click()
    Call queryMemberScript.MainMacro(TextBox1.Value)
    Unload queryMemberForm
End Sub
