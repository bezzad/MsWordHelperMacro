VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "About"
   ClientHeight    =   1890
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   5835
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    ActiveDocument.FollowHyperlink Address:="file:E:\Install Packages\Word\Word_to_epub_instruction.pdf"
End Sub

Private Sub CommandButton2_Click()
    UserForm2.Hide
End Sub

Private Sub ListBox1_Click()

End Sub

Private Sub TextBox1_Change()

End Sub

