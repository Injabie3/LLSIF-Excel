Attribute VB_Name = "Module3"
' EN Server - Score Match Macros

Sub SM_EN_Step1_Expert()
'
' Set Expert in Step 1 for EN Score Match Macro
'

'
    Dim EXPGain As Integer
    Dim BaseScore As Integer
    Dim LP As Integer
    
    EXPGain = Range("B10").Value
    LP = Range("B11").Value
    BaseScore = Range("B12").Value
        
    Range("B21").Value = "Expert"
    Range("C21").Value = EXPGain
    Range("B22").Value = "Expert"
    Range("C22").Value = BaseScore
    Range("E21").Value = LP
    
End Sub
Sub SM_EN_Step1_Hard()
'
' Set Hard in Step 1 for EN Score Match Macro
'

'
    Dim EXPGain As Integer
    Dim BaseScore As Integer
    Dim LP As Integer
    
    EXPGain = Range("C10").Value
    LP = Range("C11").Value
    BaseScore = Range("C12").Value
    
    Range("B21").Value = "Hard"
    Range("C21").Value = EXPGain
    Range("B22").Value = "Hard"
    Range("C22").Value = BaseScore
    Range("E21").Value = LP
    
End Sub
Sub SM_EN_Step1_Normal()
'
' Set Normal in Step 1 for EN Score Match Macro
'

'
    Dim EXPGain As Integer
    Dim BaseScore As Integer
    Dim LP As Integer
    
    EXPGain = Range("D10").Value
    LP = Range("D11").Value
    BaseScore = Range("D12").Value
    
    Range("B21").Value = "Normal"
    Range("C21").Value = EXPGain
    Range("B22").Value = "Normal"
    Range("C22").Value = BaseScore
    Range("E21").Value = LP
    
End Sub
Sub SM_EN_Step1_Easy()
'
' Set Easy in Step 1 for EN Score MatchMacro
'

'
    Dim EXPGain As Integer
    Dim BaseScore As Integer
    Dim LP As Integer
    
    EXPGain = Range("D10").Value
    LP = Range("E11").Value
    BaseScore = Range("E12").Value
    
    Range("B21").Value = "Easy"
    Range("C21").Value = EXPGain
    Range("B22").Value = "Easy"
    Range("C22").Value = BaseScore
    Range("E21").Value = LP
    
End Sub
Sub SM_EN_Step2_S()
'
' Set S score in Step 2 for EN Score Match Macro
'

    Dim ScoreRank As Double
    ScoreRank = Range("B16").Value
    
    Range("B23").Value = "S Rank"
    Range("C23").Value = ScoreRank

End Sub
Sub SM_EN_Step2_A()
'
' Set A score in Step 2 for EN Score Match Macro
'

    Dim ScoreRank As Double
    ScoreRank = Range("C16").Value
    
    Range("B23").Value = "A Rank"
    Range("C23").Value = ScoreRank

End Sub
Sub SM_EN_Step2_B()
'
' Set B score in Step 2 for EN Score Match Macro
'

    Dim ScoreRank As Double
    ScoreRank = Range("D16").Value
    
    Range("B23").Value = "B Rank"
    Range("C23").Value = ScoreRank

End Sub
Sub SM_EN_Step2_C()
'
' Set C score in Step 2 for EN Score Match Macro
'

    Dim ScoreRank As Double
    ScoreRank = Range("E16").Value
    
    Range("B23").Value = "C Rank"
    Range("C23").Value = ScoreRank

End Sub
Sub SM_EN_Step3_1st()
'
' Set 1st place in Step 3 for EN Score Match Macro
'

    Dim PlaceRank As Double
    PlaceRank = Range("B19").Value
    
    Range("B24").Value = "1st"
    Range("C24").Value = PlaceRank

End Sub
Sub SM_EN_Step3_2nd()
'
' Set 2nd place in Step 3 for EN Score Match Macro
'

    Dim PlaceRank As Double
    PlaceRank = Range("C19").Value
    
    Range("B24").Value = "2nd"
    Range("C24").Value = PlaceRank

End Sub
Sub SM_EN_Step3_3rd()
'
' Set 3rd place in Step 3 for EN Score Match Macro
'

    Dim PlaceRank As Double
    PlaceRank = Range("D19").Value
    
    Range("B24").Value = "3rd"
    Range("C24").Value = PlaceRank

End Sub
Sub SM_EN_Step3_4th()
'
' Set 4th place in Step 3 for EN Score Match Macro
'

    Dim PlaceRank As Double
    PlaceRank = Range("E19").Value
    
    Range("B24").Value = "4th"
    Range("C24").Value = PlaceRank

End Sub
Sub SM_EN_Add()
Attribute SM_EN_Add.VB_Description = "Adds the Score Match round points, along with EXP to their respective running totals on EN."
Attribute SM_EN_Add.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Adds Score Match round points and EXP to EN Macro
'

    Dim BasePoints As Integer
    Dim ScoreRank As Double
    Dim PlaceRank As Double
    Dim RoundPoints As Integer
    Dim EXPGain As Integer
    
    BasePoints = Range("C22").Value
    ScoreRank = Range("C23").Value
    PlaceRank = Range("C24").Value
    'the 0.00001 is to force VB to round up at .5
    RoundPoints = Round(BasePoints * ScoreRank * PlaceRank + 0.000001, 0)
    
    EXPGain = Range("C21").Value
    
    ' Add Round Points to the running total.
    Range("B28").Value = Range("B28").Value + RoundPoints
    Range("B5").Value = Range("B5").Value + EXPGain
    
    ' Add results to history
    SM_EN_HistoryAdd (RoundPoints)

End Sub
Sub SM_EN_Remove()
Attribute SM_EN_Remove.VB_Description = "Removes the Score Match round points, along with EXP to their respective running totals on EN."
Attribute SM_EN_Remove.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Removes Score Match round points & EXP from EN Macro
'

    Dim BasePoints As Integer
    Dim ScoreRank As Double
    Dim PlaceRank As Double
    Dim RoundPoints As Integer
    Dim EXPGain As Integer
    
    BasePoints = Range("C22").Value
    ScoreRank = Range("C23").Value
    PlaceRank = Range("C24").Value
    'the 0.00001 is to force VB to round up at .5
    RoundPoints = Round(BasePoints * ScoreRank * PlaceRank + 0.000001, 0)
    
    EXPGain = Range("C21").Value
    
    ' Remove Round Points from the running total.
    Range("B28").Value = Range("B28").Value - RoundPoints
    Range("B5").Value = Range("B5").Value - EXPGain
    
    ' Delete the last result from history.
    SM_EN_HistoryDel

End Sub
Sub SM_EN_HistoryAdd(RoundPoints)
Attribute SM_EN_HistoryAdd.VB_Description = "Adds the Score Match round information below the currently displayed information on EN."
Attribute SM_EN_HistoryAdd.VB_ProcData.VB_Invoke_Func = " \n14"
'
' SM_EN_HistoryAdd Macro
' Adds the current score match information to history on EN.
'

    Dim RowNumber As Integer
    RowNumber = Range("E28").Value
    
    ' Current Date/Time
    Range("A" + CStr(RowNumber)).Value = Now()
    
    ' Copy/Paste Difficulty
    Range("B" + CStr(RowNumber)).Value = Range("B22").Value
    
    ' Copy/Paste Score Rank
    Range("C" + CStr(RowNumber)).Value = Range("B23").Value
        
    ' Copy/Paste Placement
    Range("D" + CStr(RowNumber)).Value = Range("B24").Value
   
    ' Copy/Paste Round Points
    Range("E" + CStr(RowNumber)).Value = RoundPoints
    
    ' Add 1 to Rows
    Range("E28").Value = RowNumber + 1
    
End Sub

Sub SM_EN_HistoryDel()
Attribute SM_EN_HistoryDel.VB_Description = "Removes the most recent Score Match round history."
Attribute SM_EN_HistoryDel.VB_ProcData.VB_Invoke_Func = " \n14"
'
' SM_EN_HistoryDel Macro
' Deletes the last added SM history on EN
'

    Dim RowNumber As Integer
    RowNumber = Range("E28").Value - 1
    
    ' Current Date/Time
    Range("A" + CStr(RowNumber)).Value = ""
    
    ' Copy/Paste Difficulty
    Range("B" + CStr(RowNumber)).Value = ""
    
    ' Copy/Paste Score Rank
    Range("C" + CStr(RowNumber)).Value = ""
        
    ' Copy/Paste Placement
    Range("D" + CStr(RowNumber)).Value = ""
   
    ' Copy/Paste Round Points
    Range("E" + CStr(RowNumber)).Value = ""
    
    ' Put the reduced row number back in the cell.
    Range("E28").Value = RowNumber
    
End Sub
