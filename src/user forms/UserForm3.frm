VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm3 
   Caption         =   "Replace"
   ClientHeight    =   1170
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   7635
   OleObjectBlob   =   "UserForm3.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    Selection.Text = ChrW(1728)
    
    Selection.Find.ClearFormatting
    With Selection.Find
        .Text = "[" & ChrW(1728) & ChrW(1730) & ChrW(1577) & "]"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchKashida = False
        .MatchDiacritics = False
        .MatchAlefHamza = False
        .MatchControl = False
        .MatchWildcards = True
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    ' Selection.Find.Execute
    
    If Selection.Find.Execute = False Then
        MsgBox ("Replacing process done!")
    End If
    
End Sub

Private Sub CommandButton2_Click()
    Selection.Text = ChrW(1577)
    Selection.Find.ClearFormatting
    With Selection.Find
        .Text = "[" & ChrW(1728) & ChrW(1730) & ChrW(1577) & "]"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchKashida = False
        .MatchDiacritics = False
        .MatchAlefHamza = False
        .MatchControl = False
        .MatchWildcards = True
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    'êê Selection.Find.Execute
    
    If Selection.Find.Execute = False Then
        MsgBox ("Replacing process done!")
    End If
End Sub

Private Sub CommandButton3_Click()
    Selection.Find.ClearFormatting
    With Selection.Find
        .Text = "[" & ChrW(1728) & ChrW(1730) & ChrW(1577) & "]"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchKashida = False
        .MatchDiacritics = False
        .MatchAlefHamza = False
        .MatchControl = False
        .MatchWildcards = True
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    ' Selection.Find.Execute
    
    If Selection.Find.Execute = False Then
        MsgBox ("Replacing process done!")
    End If
End Sub

Private Sub CommandButton4_Click()
    For Each msr In ActiveDocument.StoryRanges
        msr.Find.ClearFormatting
        With msr.Find
            .Text = "[" & ChrW(1728) & ChrW(1730) & ChrW(1577) & "]"
            .Replacement.Text = ChrW(1728)
            .Forward = True
            .Wrap = wdFindStop
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchKashida = False
            .MatchDiacritics = False
            .MatchAlefHamza = False
            .MatchControl = False
            .MatchWildcards = True
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        msr.Find.Execute Replace:=wdReplaceAll
    Next msr
End Sub

Private Sub CommandButton5_Click()
    Selection.Find.ClearFormatting
    With Selection.Find
        .Text = "[" & ChrW(1728) & ChrW(1730) & ChrW(1577) & "]"
        .Replacement.Text = ""
        .Forward = False
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchKashida = False
        .MatchDiacritics = False
        .MatchAlefHamza = False
        .MatchControl = False
        .MatchWildcards = True
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute
End Sub
