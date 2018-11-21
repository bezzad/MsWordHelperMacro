VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FindAndReplaceSymbolForm 
   Caption         =   "Find and Replace Symbol"
   ClientHeight    =   1695
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5640
   OleObjectBlob   =   "FindAndReplaceSymbolForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FindAndReplaceSymbolForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cIndex As Integer
Public CalcState As Long
Public EventState As Boolean
Public PageBreakState As Boolean

Private Sub UserForm_Initialize()
    cIndex = Selection.Start
    If cIndex >= ActiveDocument.Characters.Count - 1 And cIndex > 0 Then
        cIndex = 0
    End If
End Sub

Private Sub userform_terminate()
    Application.ScreenUpdating = True
    '    Dialogs(wdDialogInsertSymbol).Show
End Sub

Private Sub btnReplaceAll_Click()
    Dim selectedFont, replacementText, replacementFont As String
    Dim selectedCharNum, replaceCount, charNumBuffer  As Integer
    Dim ch As Range
    
    replacementText = txtReplacementSymbol.Value
    replacementFont = "Arial"
        
    ' --------- Select only one character, please
    If Len(Selection.Text) <> 1 Then
        MsgBox Prompt:="Please select only one character.", _
                Title:="Font and Symbol", _
                Buttons:=vbExclamation
        ActiveWindow.Selection.Collapse Direction:=wdCollapseStart
        Exit Sub
    End If
    
    ' --------- If selection is insertion point, extend one character
    If Selection.Start = Selection.End Then
        Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    End If
    
    ' --------- Check replacement char is not empty
    If Len(replacementText) = 0 Then
        MsgBox Prompt:="Please enter atlast one character for replace.", _
                Title:="Font and Symbol", _
                Buttons:=vbExclamation
        Exit Sub
    End If
      
    With Dialogs(wdDialogInsertSymbol)
        selectedFont = .font
        selectedCharNum = CInt(.charNum)
    End With
                   
    'Optimize Code
    Application.ScreenUpdating = False
    
    ' Checks each character in the document, highlighting
    ' if the character's font doesn't match the OK font
    For Each ch In ActiveDocument.Characters
        ch.Select
        If Len(Selection.Text) = 1 Then
            With Dialogs(wdDialogInsertSymbol)
                charNumBuffer = CInt(.charNum)
                If CInt(.charNum) = selectedCharNum And .font = selectedFont Then
                    Selection.InsertSymbol CharacterNumber:=AscW(replacementText), font:=replacementFont, Unicode:=True
                    replaceCount = replaceCount + 1
                End If
            End With
        End If
    Next
    
    'Optimize Code
    Application.ScreenUpdating = True
    
    MsgBox Prompt:=Str(replaceCount) & " chars replaced at all document.", _
                Title:="Font and Symbol", _
                Buttons:=vbInformation
End Sub

Private Sub btnNext_Click()
    Call DetectSymbols
End Sub

Private Sub btnRaplace_Click()
    Dim replacementText, replacementFont As String
    replacementText = txtReplacementSymbol.Value
    replacementFont = "Arial"
    
    ' --------- Select only one character, please
    If Len(Selection.Text) <> 1 Then
        MsgBox Prompt:="Please select only one character.", _
                Title:="Font and Symbol", _
                Buttons:=vbExclamation
        ActiveWindow.Selection.Collapse Direction:=wdCollapseStart
        Exit Sub
    End If
    
    ' --------- If selection is insertion point, extend one character
    If Selection.Start = Selection.End Then
        Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    End If
    
    ' --------- Check replacement char is not empty
    If Len(replacementText) = 0 Then
        MsgBox Prompt:="Please enter atlast one character for replace.", _
                Title:="Font and Symbol", _
                Buttons:=vbExclamation
        Exit Sub
    End If
    
    Selection.InsertSymbol CharacterNumber:=AscW(replacementText), font:=replacementFont, Unicode:=True
    Call DetectSymbols
End Sub

Private Sub DetectSymbols()
    Dim symFonts(7) As String
    Dim font As String
    Dim symFontsLen, minSymbolAsc, maxSymbolAsc, charNumBuffer  As Integer
    Dim dlg As Object
    Dim symF As Variant
    Dim ch As Range
        
    symFontsLen = UBound(symFonts)
    minSymbolAsc = AscW(ChrW(61472)) ' 61472 is unicode = -4064 is asc
    maxSymbolAsc = AscW(ChrW(61695)) ' 61695 is unicode = -3841 is asc
    symFonts(0) = UCase("(normal text)")
    symFonts(1) = UCase("Wingdings")
    symFonts(2) = UCase("Wingdings 1")
    symFonts(3) = UCase("Wingdings 2")
    symFonts(4) = UCase("Wingdings 3")
    symFonts(5) = UCase("Symbol")
    symFonts(6) = UCase("MT Symbol")
            
    'Optimize Code
    Application.ScreenUpdating = False
        
    Selection.WholeStory
    Options.DefaultHighlightColorIndex = wdNoHighlight
    Selection.Range.HighlightColorIndex = wdNoHighlight
    
    ' Checks each character in the document, highlighting
    ' if the character's font doesn't match the OK font
    For Each ch In ActiveDocument.Characters
        If ch.Start > cIndex Then
            cIndex = ch.Start
            ch.Select
            If Len(Selection.Text) = 1 Then
                With Dialogs(wdDialogInsertSymbol)
                    ' Selection.Find.Text = "[" & ChrW(61472) & "-" & ChrW(61695) & "]"
                    charNumBuffer = CInt(.charNum)
                    If charNumBuffer >= minSymbolAsc And charNumBuffer <= maxSymbolAsc Then
                        Call HighlightSelection
                        Exit Sub
                    Else
                        font = UCase(.font)
                        If font <> "(COMPLEX SCRIPT TEXT)" Then ' COMPLEX SCRIPT TEXT no symbol
                            For Each symF In symFonts
                                If font = symF Then
                                    If symF = symFonts(0) Then ' if font is "(normal text)":
                                       Set dlg = Dialogs(wdDialogFontSubstitution)
                                       ' use dlg.Display to show symbol hidden font
                                       If dlg.unavailablefont = "ZapfDingbats" Then
                                           Call HighlightSelection
                                           Exit Sub
                                       End If
                                    Else ' so, absolutlly char font of symbol
                                        Call HighlightSelection
                                        Exit Sub
                                    End If
                                End If
                            Next
                        End If
                    End If
                End With
            End If
        End If
    Next
    
    'Optimize Code
    Application.ScreenUpdating = True
    
    MsgBox Prompt:="Search completed.", _
                Title:="Font and Symbol", _
                Buttons:=vbInformation
    Unload Me
End Sub

Private Sub HighlightSelection()
    ' Highlight the selected range of text in red
    Selection.FormattedText.HighlightColorIndex = wdRed
    ActiveWindow.ScrollIntoView Selection.Range, True
    
    'Optimize Code
    Application.ScreenUpdating = True
End Sub
