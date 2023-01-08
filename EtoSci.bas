Option Explicit

Sub ConvertToScientificNotation()
    'First of all trying to find sequences:  (digit)E+(any number of digits)
    '                                        (digit)E-(any number of digits)
    '                                        (digit)E(any number of digits)
    'using a regular expression. Case insensitive.
   
    Selection.Find.ClearFormatting
    With Selection.Find
        .Text = "[0-9][Ee][-+0-9]{1;}"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = True
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        
        Selection.Find.Execute
        
        While .Found
            'When found the sequence is selected
            'Reducing the selection by 1 left symbol therefore deselecting the last digit before symbol E
            Selection.MoveStart Unit:=wdCharacter, Count:=1
            
            'Deleting the symbol E. Case insensitive
            Selection.Text = Replace(Replace(Selection.Text, "E", ""), "e", "")
            
            'Converting the selection into float and back to string. Therefore deleting all heading zeros and symbol +
            Selection.Text = Val(Selection.Text)
            
            'Converting selected text into superscript
            Selection.Font.Superscript = wdToggle
            Selection.InsertBefore ("·10")
            
            'Collapsing selection and trying to find next sequence
            Selection.Collapse (wdCollapseEnd)                        
            .Execute
        Wend
    End With
End Sub

