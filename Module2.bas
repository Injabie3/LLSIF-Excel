Attribute VB_Name = "Module2"
' JP Server - Score Match Macros

Sub SM_JP_Step1_Expert()
'
' Set Expert in Step 1 for JP SM Macro
'

'
    Dim EXPGain As Integer
    Dim BaseScore As Integer
    Dim LP As Integer
    EXPGain = Range("H10").Value
    LP = Range("H11").Value
    BaseScore = Range("H12").Value
    
    Range("H21").Value = "Expert"
    Range("I21").Value = EXPGain
    Range("H22").Value = "Expert"
    Range("I22").Value = BaseScore
    Range("K21").Value = LP
    
End Sub
Sub SM_JP_Step1_Hard()
'
' Set Hard in Step 1 for JP SM Macro
'

'
    Dim EXPGain As Integer
    Dim BaseScore As Integer
    Dim LP As Integer
    EXPGain = Range("I10").Value
    LP = Range("I11").Value
    BaseScore = Range("I12").Value
    
    Range("H21").Value = "Hard"
    Range("I21").Value = EXPGain
    Range("H22").Value = "Hard"
    Range("I22").Value = BaseScore
    Range("K21").Value = LP
    
End Sub
Sub SM_JP_Step1_Normal()
'
' Set Normal in Step 1 for JP Score Match Macro
'

'
    Dim EXPGain As Integer
    Dim BaseScore As Integer
    Dim LP As Integer
    EXPGain = Range("J10").Value
    LP = Range("J11").Value
    BaseScore = Range("J12").Value
    
    Range("H21").Value = "Normal"
    Range("I21").Value = EXPGain
    Range("H22").Value = "Normal"
    Range("I22").Value = BaseScore
    Range("K21").Value = LP
    
End Sub
Sub SM_JP_Step1_Easy()
'
' Set Easy in Step 1 for JP Score Match Macro
'

'
    Dim EXPGain As Integer
    Dim BaseScore As Integer
    Dim LP As Integer
    EXPGain = Range("K10").Value
    LP = Range("K11").Value
    BaseScore = Range("K12").Value
    
    Range("H21").Value = "Easy"
    Range("I21").Value = EXPGain
    Range("H22").Value = "Easy"
    Range("I22").Value = BaseScore
    Range("K21").Value = LP
    
End Sub
Sub SM_JP_Step2_S()
'
' Set S score in Step 2 for JP Score Match Macro
'

    Dim ScoreRank As Double
    
    ScoreRank = Range("H16").Value
    
    Range("H23").Value = "S Rank"
    Range("I23").Value = ScoreRank

End Sub
Sub SM_JP_Step2_A()
'
' Set A score in Step 2 for JP Score Match Macro
'

    Dim ScoreRank As Double

    ScoreRank = Range("I16").Value
    
    Range("H23").Value = "A Rank"
    Range("I23").Value = ScoreRank

End Sub
Sub SM_JP_Step2_B()
'
' Set B score in Step 2 for JP Score Match Macro
'

    Dim ScoreRank As Double

    ScoreRank = Range("J16").Value
    
    Range("H23").Value = "B Rank"
    Range("I23").Value = ScoreRank

End Sub
Sub SM_JP_Step2_C()
'
' Set C score in Step 2 for JP Score Match Macro
'

    Dim ScoreRank As Double
    
    ScoreRank = Range("K16").Value
    
    Range("H23").Value = "C Rank"
    Range("I23").Value = ScoreRank

End Sub
Sub SM_JP_Step3_1st()
'
' Set 1st place in Step 3 for JP Score Match Macro
'

    Dim PlaceRank As Double

    PlaceRank = Range("H19").Value
    
    Range("H24").Value = "1st"
    Range("I24").Value = PlaceRank

End Sub
Sub SM_JP_Step3_2nd()
'
' Set 2nd place in Step 3 for JP Score Match Macro
'

    Dim PlaceRank As Double

    PlaceRank = Range("I19").Value
    
    Range("H24").Value = "2nd"
    Range("I24").Value = PlaceRank

End Sub
Sub SM_JP_Step3_3rd()
'
' Set 3rd place in Step 3 for JP Score Match Macro
'

    Dim PlaceRank As Double

    PlaceRank = Range("J19").Value
    
    Range("H24").Value = "3rd"
    Range("I24").Value = PlaceRank

End Sub
Sub SM_JP_Step3_4th()
'
' Set 4th place in Step 3 for JP Score Match Macro
'

    Dim PlaceRank As Double
    PlaceRank = Range("K19").Value
    
    Range("H24").Value = "4th"
    Range("I24").Value = PlaceRank

End Sub
Sub SM_JP_Add()
'
' Adds Score Match round points and EXP to JP Macro
'

    Dim BasePoints As Integer
    Dim ScoreRank As Double
    Dim PlaceRank As Double
    Dim RoundPoints As Integer
    Dim EXPGain As Integer
    
    
    BasePoints = Range("I22").Value
    ScoreRank = Range("I23").Value
    PlaceRank = Range("I24").Value
    'the 0.00001 is to force VB to round up at .5
    RoundPoints = Round(BasePoints * ScoreRank * PlaceRank + 0.000001, 0)
    
    EXPGain = Range("I21").Value
    
    ' Add Round Points to the running total.
    Range("H28").Value = Range("H28").Value + RoundPoints
    Range("H5").Value = Range("H5").Value + EXPGain
    
    ' Add results to history
    SM_JP_HistoryAdd (RoundPoints)

End Sub
Sub SM_JP_Remove()
'
' Removes Score Match round points & EXP from JP Macro
'

    Dim BasePoints As Integer
    Dim ScoreRank As Double
    Dim PlaceRank As Double
    Dim RoundPoints As Integer
    Dim EXPGain As Integer
    
    BasePoints = Range("I22").Value
    ScoreRank = Range("I23").Value
    PlaceRank = Range("I24").Value
    'the 0.00001 is to force VB to round up at .5
    RoundPoints = Round(BasePoints * ScoreRank * PlaceRank + 0.000001, 0)
    
    EXPGain = Range("I21").Value

    ' Remove Round Points from the running total.
    Range("H28").Value = Range("H28").Value - RoundPoints
    Range("H5").Value = Range("H5").Value - EXPGain
    
    ' Delete the last result from history.
    SM_JP_HistoryDel

End Sub
Sub SM_JP_HistoryAdd(RoundPoints)
'
' SM_JP_HistoryAdd Macro
' Adds the current score match information to history.
'

    Dim RowNumber As Integer
    RowNumber = Range("K28").Value
    
    ' Current Date/Time
    Range("G" + CStr(RowNumber)).Value = Now()
    
    ' Copy/Paste Difficulty
    Range("H" + CStr(RowNumber)).Value = Range("H22").Value
    
    ' Copy/Paste Score Rank
    Range("I" + CStr(RowNumber)).Value = Range("H23").Value
        
    ' Copy/Paste Placement
    Range("J" + CStr(RowNumber)).Value = Range("H24").Value
   
    ' Copy/Paste Round Points
    Range("K" + CStr(RowNumber)).Value = RoundPoints
    
    ' Add 1 to Rows
    Range("K28").Value = RowNumber + 1
    
End Sub

Sub SM_JP_HistoryDel()
'
' SM_JP_HistoryDel Macro
' Deletes the last added SM history on JP
'

    Dim RowNumber As Integer
    RowNumber = Range("K28").Value - 1
    
    ' Current Date/Time
    Range("G" + CStr(RowNumber)).Value = ""
    
    ' Copy/Paste Difficulty
    Range("H" + CStr(RowNumber)).Value = ""
    
    ' Copy/Paste Score Rank
    Range("I" + CStr(RowNumber)).Value = ""
        
    ' Copy/Paste Placement
    Range("J" + CStr(RowNumber)).Value = ""
   
    ' Copy/Paste Round Points
    Range("K" + CStr(RowNumber)).Value = ""
    
    ' Put the reduced row number back in the cell.
    Range("K28").Value = RowNumber
    
End Sub



