'
' taaghche Macros
'
'

Dim table_counter As Integer


Sub optimize_enter()
       
    For Each msr In ActiveDocument.StoryRanges
        Application.DisplayStatusBar = True
        msr.Find.ClearFormatting
        msr.Find.Replacement.ClearFormatting
        
        '
        '
        'remove extra enter
        '
        '

        With msr.Find
            .Text = "(^13)@"
            .Replacement.Text = "^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchKashida = False
            .MatchDiacritics = False
            .MatchAlefHamza = False
            .MatchControl = False
            .MatchAllWordForms = False
            .MatchSoundsLike = False
            .MatchWildcards = True
        End With
        msr.Find.Execute Replace:=wdReplaceAll
        
        With msr.Find
            .Text = "(^l)@"
            .Replacement.Text = "^l"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchKashida = False
            .MatchDiacritics = False
            .MatchAlefHamza = False
            .MatchControl = False
            .MatchAllWordForms = False
            .MatchSoundsLike = False
            .MatchWildcards = True
        End With
        msr.Find.Execute Replace:=wdReplaceAll
        
    Next msr
End Sub


Sub convert_english_num_to_Persian()

    '
    '
    'convert English numbers to Persian
    '
    '
    
    Dim str_range_count As Integer
    Dim jobs As Integer
    Dim job_counter
    str_range_count = ActiveDocument.StoryRanges.Count
    jobs = 10 * str_range_count
    job_counter = 0
    
    For Each msr In ActiveDocument.StoryRanges
         Application.DisplayStatusBar = True
         msr.Find.ClearFormatting
         msr.Find.Replacement.ClearFormatting
        
         With msr.Find
             .Text = "0"
             .Replacement.Text = ChrW(1776)
             .Forward = True
             .Wrap = wdFindContinue
             .Format = False
             .MatchCase = False
             .MatchWholeWord = False
             .MatchKashida = False
             .MatchDiacritics = False
             .MatchAlefHamza = False
             .MatchControl = False
             .MatchWildcards = False
             .MatchSoundsLike = False
             .MatchAllWordForms = False
         End With
         msr.Find.Execute Replace:=wdReplaceAll
         job_counter = job_counter + 1
         Application.StatusBar = job_counter & " of " & jobs & " tasks is done."
         
         With msr.Find
             .Text = "1"
             .Replacement.Text = ChrW(1777)
             .Forward = True
             .Wrap = wdFindContinue
             .Format = False
             .MatchCase = False
             .MatchWholeWord = False
             .MatchKashida = False
             .MatchDiacritics = False
             .MatchAlefHamza = False
             .MatchControl = False
             .MatchWildcards = False
             .MatchSoundsLike = False
             .MatchAllWordForms = False
         End With
         msr.Find.Execute Replace:=wdReplaceAll
         job_counter = job_counter + 1
         Application.StatusBar = job_counter & " of " & jobs & " tasks is done."
        
        With msr.Find
             .Text = "2"
             .Replacement.Text = ChrW(1778)
             .Forward = True
             .Wrap = wdFindContinue
             .Format = False
             .MatchCase = False
             .MatchWholeWord = False
             .MatchKashida = False
             .MatchDiacritics = False
             .MatchAlefHamza = False
             .MatchControl = False
             .MatchWildcards = False
             .MatchSoundsLike = False
             .MatchAllWordForms = False
         End With
         msr.Find.Execute Replace:=wdReplaceAll
         job_counter = job_counter + 1
         Application.StatusBar = job_counter & " of " & jobs & " tasks is done."
        
         With msr.Find
             .Text = "3"
             .Replacement.Text = ChrW(1779)
             .Forward = True
             .Wrap = wdFindContinue
             .Format = False
             .MatchCase = False
             .MatchWholeWord = False
             .MatchKashida = False
             .MatchDiacritics = False
             .MatchAlefHamza = False
             .MatchControl = False
             .MatchWildcards = False
             .MatchSoundsLike = False
             .MatchAllWordForms = False
         End With
         msr.Find.Execute Replace:=wdReplaceAll
         job_counter = job_counter + 1
         Application.StatusBar = job_counter & " of " & jobs & " tasks is done."
        
        With msr.Find
             .Text = "4"
             .Replacement.Text = ChrW(1780)
             .Forward = True
             .Wrap = wdFindContinue
             .Format = False
             .MatchCase = False
             .MatchWholeWord = False
             .MatchKashida = False
             .MatchDiacritics = False
             .MatchAlefHamza = False
             .MatchControl = False
             .MatchWildcards = False
             .MatchSoundsLike = False
             .MatchAllWordForms = False
         End With
         msr.Find.Execute Replace:=wdReplaceAll
         job_counter = job_counter + 1
         Application.StatusBar = job_counter & " of " & jobs & " tasks is done."
        
        With msr.Find
             .Text = "5"
             .Replacement.Text = ChrW(1781)
             .Forward = True
             .Wrap = wdFindContinue
             .Format = False
             .MatchCase = False
             .MatchWholeWord = False
             .MatchKashida = False
             .MatchDiacritics = False
             .MatchAlefHamza = False
             .MatchControl = False
             .MatchWildcards = False
             .MatchSoundsLike = False
             .MatchAllWordForms = False
         End With
         msr.Find.Execute Replace:=wdReplaceAll
         job_counter = job_counter + 1
         Application.StatusBar = job_counter & " of " & jobs & " tasks is done."
        
        With msr.Find
             .Text = "6"
             .Replacement.Text = ChrW(1782)
             .Forward = True
             .Wrap = wdFindContinue
             .Format = False
             .MatchCase = False
             .MatchWholeWord = False
             .MatchKashida = False
             .MatchDiacritics = False
             .MatchAlefHamza = False
             .MatchControl = False
             .MatchWildcards = False
             .MatchSoundsLike = False
             .MatchAllWordForms = False
         End With
         msr.Find.Execute Replace:=wdReplaceAll
         job_counter = job_counter + 1
         Application.StatusBar = job_counter & " of " & jobs & " tasks is done."
         
         With msr.Find
             .Text = "7"
             .Replacement.Text = ChrW(1783)
             .Forward = True
             .Wrap = wdFindContinue
             .Format = False
             .MatchCase = False
             .MatchWholeWord = False
             .MatchKashida = False
             .MatchDiacritics = False
             .MatchAlefHamza = False
             .MatchControl = False
             .MatchWildcards = False
             .MatchSoundsLike = False
             .MatchAllWordForms = False
         End With
         msr.Find.Execute Replace:=wdReplaceAll
         job_counter = job_counter + 1
         Application.StatusBar = job_counter & " of " & jobs & " tasks is done."
         
         With msr.Find
             .Text = "8"
             .Replacement.Text = ChrW(1784)
             .Forward = True
             .Wrap = wdFindContinue
             .Format = False
             .MatchCase = False
             .MatchWholeWord = False
             .MatchKashida = False
             .MatchDiacritics = False
             .MatchAlefHamza = False
             .MatchControl = False
             .MatchWildcards = False
             .MatchSoundsLike = False
             .MatchAllWordForms = False
         End With
         msr.Find.Execute Replace:=wdReplaceAll
         job_counter = job_counter + 1
         Application.StatusBar = job_counter & " of " & jobs & " tasks is done."
        
        With msr.Find
             .Text = "9"
             .Replacement.Text = ChrW(1785)
             .Forward = True
             .Wrap = wdFindContinue
             .Format = False
             .MatchCase = False
             .MatchWholeWord = False
             .MatchKashida = False
             .MatchDiacritics = False
             .MatchAlefHamza = False
             .MatchControl = False
             .MatchWildcards = False
             .MatchSoundsLike = False
             .MatchAllWordForms = False
         End With
         msr.Find.Execute Replace:=wdReplaceAll
         job_counter = job_counter + 1
         Application.StatusBar = job_counter & " of " & jobs & " tasks is done."
    Next msr
    MsgBox ("Done")
    Application.DisplayStatusBar = False
End Sub


Sub convert_to_jpeg()
'
' convert_to_jpeg Macro
'
'
    Selection.Copy
    Selection.Delete Unit:=wdCharacter, Count:=1
    Selection.PasteSpecial Link:=False, DataType:=15, Placement:=wdInLine, _
        DisplayAsIcon:=False
End Sub
Sub reverse_word()
' last update: 2017.06.07
'
' reverse_word Macro
'
'
    'To call me use follow code:
    'Application.Run MacroName:="Project.NewMacros.reverse_word"
    
    
    ' If text is not already selected,
    ' use the Select method to select the text that is associated
    ' with a specific object and create a Selection object.
    ' The following instruction selects the first word in the active document.
    If Selection.Type = wdSelectionIP Then
        Selection.Words(1).Select
    End If
    
    ' It is possible for the user to select a region in a document
    ' that does not represent contiguous text
    ' (for example, when using the ALT key with the mouse).
    ' Because the behavior of such a selection can be unpredictable,
    ' you may want to include a step in your code that checks
    ' the Type property of a selection before performing any operations on it
    ' (Selection.Type = wdSelectionBlock). Similarly, selections that include
    ' table cells can also lead to unpredictable behavior. The Information property
    ' will tell you if a selection is inside a table (Selection.Information(wdWithinTable) = True).
    ' The following example determines if a selection is normal
    ' (for example, it is not a row or column in a table, it is not a vertical block of text);
    ' you could use it to test the current selection before performing any operations on it.
    If Selection.Type <> wdSelectionNormal Then
        MsgBox Prompt:="Not a valid selection! Exiting procedure..."
        Exit Sub
    End If

   

    ' Store position to check it after run Ltr
    Dim lastPos
    lastPos = Selection.Information(wdHorizontalPositionRelativeToTextBoundary)
    
    'If Selection.ParagraphFormat.ReadingOrder <> wdReadingOrderLtr Then
    ' Sets the reading order and alignment of the specified run to left-to-right.
    Selection.LtrRun
    
    ' If we are using an RTL language, even though
    ' the sequence of letters is the same order,
    ' the text should be displayed in reverse, as a RTL run.
    If lastPos = Selection.Information(wdHorizontalPositionRelativeToTextBoundary) Then
        Dim i As Long
        Dim OldString As Variant
        Dim RevText As Variant
        
        OldString = Split(Selection.Text, "[ ]")
        
        For i = 0 To UBound(OldString)
            OldString(i) = StrReverse(OldString(i))
        Next
        RevText = Join(OldString, " ")
        Selection.Text = RevText
    End If
End Sub
Sub replace_char()
'
' replace_char Macro
'
'
    For Each msr In ActiveDocument.StoryRanges
        msr.Find.ClearFormatting
        msr.Find.Replacement.ClearFormatting
        With msr.Find
            .Text = ChrW(1610)
            .Replacement.Text = ChrW(1740)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchKashida = False
            .MatchDiacritics = False
            .MatchAlefHamza = False
            .MatchControl = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        msr.Find.Execute Replace:=wdReplaceAll
    
        With msr.Find
            .Text = ChrW(1603)
            .Replacement.Text = ChrW(1705)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchKashida = False
            .MatchDiacritics = False
            .MatchAlefHamza = False
            .MatchControl = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        msr.Find.Execute Replace:=wdReplaceAll
    Next msr
    MsgBox ("Done")
End Sub

Sub show_hilight_dialog()
    UserForm1.Show
End Sub


Sub Help()
'
' Opens help file
'
'
    UserForm2.Show

End Sub

Sub Bold_selected()
'
' Bold_Selected Macro
'
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.font.Bold = True
    With Selection.Find
        .Text = Selection.Text
        .Replacement.Text = Selection.Text
        .Replacement.font.Bold = True
        .Replacement.font.BoldBi = True
        .Forward = False
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchKashida = False
        .MatchDiacritics = False
        .MatchAlefHamza = False
        .MatchControl = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll, Format:=True
    MsgBox ("Done")
End Sub

Sub find_first_large_paragraph()
    Dim prg_counter As Long
    Dim apara As Paragraph
    Dim lineText As String
    Dim maxCharsForPara As Integer
    Dim thBytes As Integer
    Dim firstParaSelected As Boolean
    
    thBytes = 26 * 1024
    maxCharsForPara = 13900
    prg_counter = 0
    
    For Each apara In ActiveDocument.Paragraphs
        If apara.Range.Characters.Count > maxCharsForPara Or _
            LenB(apara.Range.Text) >= thBytes Then
            prg_counter = prg_counter + 1
            If Not firstParaSelected Then
                apara.Range.Select
                firstParaSelected = True
            End If
        End If
    Next apara
    
    MsgBox ("There is " & prg_counter & " large paragraphs wich must be edited.")
    
End Sub
Sub initial_settings()
'
' initial_settings Macro
'

    '''''''removes any hyperlink in any storyrange
    Selection.WholeStory
    With Selection.font
        .Name = ""
        .Spacing = 0
        .Scaling = 100
        .Position = 0
        .NameBi = ""
    End With
    For Each msr In ActiveDocument.StoryRanges
        For Each oField In msr.Fields
            If oField.Type = wdFieldHyperlink Then
                oField.Unlink
            End If
        Next oField
    Next msr
    
    '''''''???????
    Selection.WholeStory
    If ActiveWindow.View.SplitSpecial <> wdPaneNone Then
        ActiveWindow.Panes(2).Close
    End If
    If ActiveWindow.ActivePane.View.Type <> wdPrintView Then
        ActiveWindow.ActivePane.View.Type = wdPrintView
    End If
    
    ''''''' make one column
    On Error GoTo ErrHandler

    With Selection.PageSetup.TextColumns
        .SetCount NumColumns:=1
        .EvenlySpaced = True
        .LineBetween = False
    End With

ErrHandler:
    '''''''replace automatic numbering to text and convert all list charachters to bulet
    ActiveDocument.Range.ListFormat.ConvertNumbersToText
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = ChrW(61623)
        .Replacement.Text = ChrW(9679)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchKashida = False
        .MatchDiacritics = False
        .MatchAlefHamza = False
        .MatchControl = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    With Selection.Find
        .Text = ChrW(61692)
        .Replacement.Text = ChrW(9679)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchKashida = False
        .MatchDiacritics = False
        .MatchAlefHamza = False
        .MatchControl = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    With Selection.Find
        .Text = ChrW(61656)
        .Replacement.Text = ChrW(9679)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchKashida = False
        .MatchDiacritics = False
        .MatchAlefHamza = False
        .MatchControl = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    With Selection.Find
        .Text = ChrW(61558)
        .Replacement.Text = ChrW(9679)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchKashida = False
        .MatchDiacritics = False
        .MatchAlefHamza = False
        .MatchControl = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    With Selection.Find
        .Text = ChrW(61607)
        .Replacement.Text = ChrW(9679)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchKashida = False
        .MatchDiacritics = False
        .MatchAlefHamza = False
        .MatchControl = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

   ' With Selection.Find
   '     .Text = "o"
   '     .Replacement.Text = ChrW(9675)
   '     .Forward = True
   '     .Wrap = wdFindContinue
   '     .Format = False
   '     .MatchCase = False
   '     .MatchWholeWord = False
   '     .MatchKashida = False
   '     .MatchDiacritics = False
   '     .MatchAlefHamza = False
   '     .MatchControl = False
   '     .MatchWildcards = False
   '     .MatchSoundsLike = False
   '     .MatchAllWordForms = False
   ' End With
   ' Selection.Find.Execute Replace:=wdReplaceAll
    
    '''''''Delete Kesh Dahande
    Delete_Kesh_Dahande

    '''''''optimize enter
    optimize_enter

    '''''''iterate all shapes to config them
    For Each myshape In ActiveDocument.Shapes
        '' wrap to bottom and top + align center
        myshape.WrapFormat.Type = wdWrapTopBottom
        myshape.Left = wdShapeCenter
        
        Next myshape
        
    '''''''iterate all shape inlines to config them
    For Each shpIn In ActiveDocument.InlineShapes
        Dim shape As Word.shape
        Set shape = shpIn.ConvertToShape
        
        '' wrap to bottom and top + align center
        shape.WrapFormat.Type = wdWrapTopBottom
        shape.Left = wdShapeCenter
        
        Next shpIn
        
    MsgBox ("Done")
End Sub

Sub Heh_hamza_noqteh()
'
' Heh_hamza_noqteh Macro
'
'
    On Error GoTo Done
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "[" & ChrW(1731) & ChrW(1577) & "]"
        .Replacement.Text = ChrW(1728)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchKashida = False
        .MatchDiacritics = False
        .MatchAlefHamza = False
        .MatchControl = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute
    
    Dim dlg As Dialog
    Set dlg = Application.Dialogs(wdDialogEditReplace)
    dlg.Show
    
Done:
If Err.Number = 5452 Then Exit Sub

End Sub
Sub nim_falsele()
'
' nim_falsele Macro
'
'
    For Each msr In ActiveDocument.StoryRanges
        msr.Find.ClearFormatting
        msr.Find.Replacement.ClearFormatting
        With msr.Find
            .Text = "[" & ChrW(61597) & ChrW(8192) & ChrW(8193) & ChrW(8194) & ChrW(8195) & ChrW(8196) & ChrW(8197) & ChrW(8198) & ChrW(8199) & ChrW(8200) & ChrW(8201) & ChrW(8202) & ChrW(8203) & ChrW(8204) & ChrW(8205) & ChrW(8206) & ChrW(8207) & "]{1,}"
            .Replacement.Text = "aaxx"
            .Replacement.font.Name = "Times New Roman"
            .Forward = True
            .Wrap = wdFindContinue
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
        msr.Find.Execute Replace:=wdReplaceAll, Format:=True
        
        With msr.Find
            .Text = "aaxx"
            .Replacement.Text = ChrW(8204)
            .Replacement.font.Name = "B Nazanin"
            .Forward = True
            .Wrap = wdFindContinue
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
        msr.Find.Execute Replace:=wdReplaceAll, Format:=True
        msr.Find.ClearFormatting
        msr.Find.Replacement.ClearFormatting
        With msr.Find
            .Text = "^31"
            .Replacement.Text = ChrW(8204)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = True
            .MatchCase = False
            .MatchWholeWord = False
            .MatchKashida = False
            .MatchDiacritics = False
            .MatchAlefHamza = False
            .MatchControl = False
            .MatchAllWordForms = False
            .MatchSoundsLike = False
            .MatchWildcards = True
        End With
        msr.Find.Execute Replace:=wdReplaceAll
    Next msr
End Sub
Sub Delete_Kesh_Dahande()
'
' Delete_Kesh_Dahande Macro
'
'
    For Each msr In ActiveDocument.StoryRanges
        msr.Find.ClearFormatting
        msr.Find.Replacement.ClearFormatting
        With msr.Find
             .Text = "([! 1234567890" & ChrW(1777) & ChrW(1778) & ChrW(1779) & ChrW(1780) & ChrW(1781) & ChrW(1782) & ChrW(1783) & ChrW(1784) & ChrW(1785) & ChrW(1776) & "])(" & ChrW(1600) & "{1,})([! 1234567890" & ChrW(1777) & ChrW(1778) & ChrW(1779) & ChrW(1780) & ChrW(1781) & ChrW(1782) & ChrW(1783) & ChrW(1784) & ChrW(1785) & ChrW(1776) & "])"
            .Replacement.Text = "\1\3"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchKashida = False
            .MatchDiacritics = False
            .MatchAlefHamza = False
            .MatchControl = False
            .MatchAllWordForms = False
            .MatchSoundsLike = False
            .MatchWildcards = True
            msr.Find.Execute Replace:=wdReplaceAll
        End With
    Next msr
End Sub
Sub Bold_Nama()
'
' Bold_nama Macro
'
'

For Each sentence In ActiveDocument.StoryRanges
    For Each w In sentence.Words
        With w.font
            If .Bold = True And .BoldBi = False Then
               .Bold = False
            End If
        
            If .Bold = False And .BoldBi = True Then
                .Bold = True
            End If
        End With
    Next
Next

MsgBox ("Done")
End Sub

Sub DetectSymbols()
    '
    ' Find and replace symbols
    '
    FindAndReplaceSymbolForm.Show 'vbModal 'modal for top most
End Sub