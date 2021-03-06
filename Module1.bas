Attribute VB_Name = "Module1"
' EN and JP Servers - Token Collection Macros
' Includes universal LCS (lovecas) macros

Sub AddLCS_EN()
Attribute AddLCS_EN.VB_Description = "Adds love gems to EN event (on current sheet)"
Attribute AddLCS_EN.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Add lovecas to EN Macro
'

'
    Range("E5").Select
    ActiveCell.FormulaR1C1 = ActiveCell.FormulaR1C1 + 1
End Sub
Sub TC_MinusToken_EN()
'
' Remove token from EN in the case of missing one during song Macro
'

'
    Range("B16").Select
    ActiveCell.FormulaR1C1 = ActiveCell.FormulaR1C1 - 1
    Range("B17").Select
    ActiveCell.FormulaR1C1 = ActiveCell.FormulaR1C1 - 1
End Sub
Sub TC_AddEXP_EN_Expert()
'
' Add EXP and Tokens to EN Macro
'

'
    Range("B5").Select
    ActiveCell.FormulaR1C1 = ActiveCell.FormulaR1C1 + 83
    Range("B16").Select
    ActiveCell.FormulaR1C1 = ActiveCell.FormulaR1C1 + 27
    Range("B17").Select
    ActiveCell.FormulaR1C1 = ActiveCell.FormulaR1C1 + 27
End Sub
Sub TC_AddEXP_EN_Hard()
'
' Add EXP and Tokens to EN Macro
'

'
    Range("B5").Select
    ActiveCell.FormulaR1C1 = ActiveCell.FormulaR1C1 + 46
    Range("B16").Select
    ActiveCell.FormulaR1C1 = ActiveCell.FormulaR1C1 + 16
    Range("B17").Select
    ActiveCell.FormulaR1C1 = ActiveCell.FormulaR1C1 + 16
End Sub
Sub TC_AddEXP_EN_Normal()
'
' Add EXP and Tokens to EN Macro
'

'
    Range("B5").Select
    ActiveCell.FormulaR1C1 = ActiveCell.FormulaR1C1 + 26
    Range("B16").Select
    ActiveCell.FormulaR1C1 = ActiveCell.FormulaR1C1 + 10
    Range("B17").Select
    ActiveCell.FormulaR1C1 = ActiveCell.FormulaR1C1 + 10
End Sub
Sub TC_AddEXP_EN_Easy()
'
' Add EXP and Tokens to EN Macro
'

'
    Range("B5").Select
    ActiveCell.FormulaR1C1 = ActiveCell.FormulaR1C1 + 12
    Range("B16").Select
    ActiveCell.FormulaR1C1 = ActiveCell.FormulaR1C1 + 5
    Range("B17").Select
    ActiveCell.FormulaR1C1 = ActiveCell.FormulaR1C1 + 5
End Sub
Sub TC_AddEXP_EN_Token_Expert()
'
' Add EXP and subtract tokens from EN Macro
'

'
    Dim TokensPerSong As Integer
    Dim EXPGain As Integer
    Dim PointsPerSong As Integer
    Range("B19").Select
    TokensPerSong = ActiveCell.FormulaR1C1
    Range("B10").Select
    EXPGain = ActiveCell.FormulaR1C1
    Range("B20").Select
    PointsPerSong = ActiveCell.FormulaR1C1

    Range("B5").Select
    ActiveCell.FormulaR1C1 = ActiveCell.FormulaR1C1 + EXPGain
    'Current Tokens
    Range("B16").Select
    ActiveCell.FormulaR1C1 = ActiveCell.FormulaR1C1 - TokensPerSong
    'Current Points
    Range("B17").Select
    ActiveCell.FormulaR1C1 = ActiveCell.FormulaR1C1 + PointsPerSong
End Sub
Sub TC_AddEXP_EN_Token_Hard()
'
' Add EXP and subtract tokens from EN Macro
'

'
    Dim TokensPerSong As Integer
    Dim EXPGain As Integer
    Dim PointsPerSong As Integer
    Range("C19").Select
    TokensPerSong = ActiveCell.FormulaR1C1
    Range("C10").Select
    EXPGain = ActiveCell.FormulaR1C1
    Range("C20").Select
    PointsPerSong = ActiveCell.FormulaR1C1
    
    Range("B5").Select
    ActiveCell.FormulaR1C1 = ActiveCell.FormulaR1C1 + EXPGain
    'Current Tokens
    Range("B16").Select
    ActiveCell.FormulaR1C1 = ActiveCell.FormulaR1C1 - TokensPerSong
    'Current Points
    Range("B17").Select
    ActiveCell.FormulaR1C1 = ActiveCell.FormulaR1C1 + PointsPerSong
End Sub
Sub TC_AddEXP_EN_Token_Normal()
'
' Add EXP and subtract tokens from EN Macro
'

'
    Dim TokensPerSong As Integer
    Dim EXPGain As Integer
    Dim PointsPerSong As Integer
    Range("D19").Select
    TokensPerSong = ActiveCell.FormulaR1C1
    Range("D10").Select
    EXPGain = ActiveCell.FormulaR1C1
    Range("D20").Select
    PointsPerSong = ActiveCell.FormulaR1C1
    
    Range("B5").Select
    ActiveCell.FormulaR1C1 = ActiveCell.FormulaR1C1 + EXPGain
    'Current Tokens
    Range("B16").Select
    ActiveCell.FormulaR1C1 = ActiveCell.FormulaR1C1 - TokensPerSong
    'Current Points
    Range("B17").Select
    ActiveCell.FormulaR1C1 = ActiveCell.FormulaR1C1 + PointsPerSong
End Sub
Sub TC_AddEXP_EN_Token_Easy()
'
' Add EXP and subtract tokens from EN Macro
'

'
    Dim TokensPerSong As Integer
    Dim EXPGain As Integer
    Dim PointsPerSong As Integer
    Range("E19").Select
    TokensPerSong = ActiveCell.FormulaR1C1
    Range("E10").Select
    EXPGain = ActiveCell.FormulaR1C1
    Range("E20").Select
    PointsPerSong = ActiveCell.FormulaR1C1
    
    Range("B5").Select
    ActiveCell.FormulaR1C1 = ActiveCell.FormulaR1C1 + EXPGain
    'Current Tokens
    Range("B16").Select
    ActiveCell.FormulaR1C1 = ActiveCell.FormulaR1C1 - TokensPerSong
    'Current Points
    Range("B17").Select
    ActiveCell.FormulaR1C1 = ActiveCell.FormulaR1C1 + PointsPerSong
End Sub
Sub TC_AddEXP_EN_Token_Hard_4x()
'
' Add EXP and subtract tokens from EN Macro
'

'
    Dim TokensPerSong As Integer
    Dim EXPGain As Integer
    Dim PointsPerSong As Integer
    Range("C19").Select
    TokensPerSong = ActiveCell.FormulaR1C1
    Range("C10").Select
    EXPGain = ActiveCell.FormulaR1C1
    Range("C20").Select
    PointsPerSong = ActiveCell.FormulaR1C1
    
    Range("B5").Select
    ActiveCell.FormulaR1C1 = ActiveCell.FormulaR1C1 + EXPGain
    'Current Tokens
    Range("B16").Select
    ActiveCell.FormulaR1C1 = ActiveCell.FormulaR1C1 - (4 * TokensPerSong)
    'Current Points
    Range("B17").Select
    ActiveCell.FormulaR1C1 = ActiveCell.FormulaR1C1 + (4 * PointsPerSong)
End Sub
Sub TC_AddEXP_EN_Token_Normal_4x()
'
' Add EXP and subtract tokens from EN Macro
'

'
    Dim TokensPerSong As Integer
    Dim EXPGain As Integer
    Dim PointsPerSong As Integer
    Range("D19").Select
    TokensPerSong = ActiveCell.FormulaR1C1
    Range("D10").Select
    EXPGain = ActiveCell.FormulaR1C1
    Range("D20").Select
    PointsPerSong = ActiveCell.FormulaR1C1
    
    Range("B5").Select
    ActiveCell.FormulaR1C1 = ActiveCell.FormulaR1C1 + EXPGain
    'Current Tokens
    Range("B16").Select
    ActiveCell.FormulaR1C1 = ActiveCell.FormulaR1C1 - (4 * TokensPerSong)
    'Current Points
    Range("B17").Select
    ActiveCell.FormulaR1C1 = ActiveCell.FormulaR1C1 + (4 * PointsPerSong)
End Sub
Sub TC_AddEXP_EN_Token_Easy_4x()
'
' Add EXP and subtract tokens from EN Macro
'

'
    Dim TokensPerSong As Integer
    Dim EXPGain As Integer
    Dim PointsPerSong As Integer
    Range("E19").Select
    TokensPerSong = ActiveCell.FormulaR1C1
    Range("E10").Select
    EXPGain = ActiveCell.FormulaR1C1
    Range("E20").Select
    PointsPerSong = ActiveCell.FormulaR1C1
    
    Range("B5").Select
    ActiveCell.FormulaR1C1 = ActiveCell.FormulaR1C1 + EXPGain
    'Current Tokens
    Range("B16").Select
    ActiveCell.FormulaR1C1 = ActiveCell.FormulaR1C1 - (4 * TokensPerSong)
    'Current Points
    Range("B17").Select
    ActiveCell.FormulaR1C1 = ActiveCell.FormulaR1C1 + (4 * PointsPerSong)
End Sub
Sub AddLCS_JP()
Attribute AddLCS_JP.VB_Description = "Add lovecas to JP event (on current sheet)"
Attribute AddLCS_JP.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Add lovecas to EN Macro
'

'
    Range("K5").Select
    ActiveCell.FormulaR1C1 = ActiveCell.FormulaR1C1 + 1
End Sub
Sub TC_MinusToken_JP()
'
' Remove token from JP in the case of missing one during song Macro
'

'
    Range("H16").Select
    ActiveCell.FormulaR1C1 = ActiveCell.FormulaR1C1 - 1
    Range("H17").Select
    ActiveCell.FormulaR1C1 = ActiveCell.FormulaR1C1 - 1
End Sub
Sub TC_AddEXP_JP_Expert()
'
' Add EXP and Tokens to JP Macro
'

'
    Range("H5").Select
    ActiveCell.FormulaR1C1 = ActiveCell.FormulaR1C1 + 83
    Range("H16").Select
    ActiveCell.FormulaR1C1 = ActiveCell.FormulaR1C1 + 27
    Range("H17").Select
    ActiveCell.FormulaR1C1 = ActiveCell.FormulaR1C1 + 27
End Sub
Sub TC_AddEXP_JP_Hard()
'
' Add EXP and Tokens to JP Macro
'

'
    Range("H5").Select
    ActiveCell.FormulaR1C1 = ActiveCell.FormulaR1C1 + 46
    Range("H16").Select
    ActiveCell.FormulaR1C1 = ActiveCell.FormulaR1C1 + 16
    Range("H17").Select
    ActiveCell.FormulaR1C1 = ActiveCell.FormulaR1C1 + 16
End Sub
Sub TC_AddEXP_JP_Normal()
'
' Add EXP and Tokens to JP Macro
'

'
    Range("H5").Select
    ActiveCell.FormulaR1C1 = ActiveCell.FormulaR1C1 + 26
    Range("H16").Select
    ActiveCell.FormulaR1C1 = ActiveCell.FormulaR1C1 + 10
    Range("H17").Select
    ActiveCell.FormulaR1C1 = ActiveCell.FormulaR1C1 + 10
End Sub
Sub TC_AddEXP_JP_Easy()
'
' Add EXP and Tokens to JP Macro
'

'
    Range("H5").Select
    ActiveCell.FormulaR1C1 = ActiveCell.FormulaR1C1 + 12
    Range("H16").Select
    ActiveCell.FormulaR1C1 = ActiveCell.FormulaR1C1 + 5
    Range("H17").Select
    ActiveCell.FormulaR1C1 = ActiveCell.FormulaR1C1 + 5
End Sub
Sub TC_AddEXP_JP_Token_Expert()
'
' Add EXP and subtract tokens from JP Macro
'

'
    Dim TokensPerSong As Integer
    Dim EXPGain As Integer
    Dim PointsPerSong As Integer
    
    Range("H19").Select
    TokensPerSong = ActiveCell.FormulaR1C1
    Range("H10").Select
    EXPGain = ActiveCell.FormulaR1C1
    Range("H20").Select
    PointsPerSong = ActiveCell.FormulaR1C1
    
    Range("H5").Select
    ActiveCell.FormulaR1C1 = ActiveCell.FormulaR1C1 + EXPGain
    'Current Tokens
    Range("H16").Select
    ActiveCell.FormulaR1C1 = ActiveCell.FormulaR1C1 - TokensPerSong
    'Current Points
    Range("H17").Select
    ActiveCell.FormulaR1C1 = ActiveCell.FormulaR1C1 + PointsPerSong
End Sub
Sub TC_AddEXP_JP_Token_Expert_4x()
'
' Add EXP and subtract tokens from JP Macro
'

'
    Dim TokensPerSong As Integer
    Dim EXPGain As Integer
    Dim PointsPerSong As Integer
    
    Range("H19").Select
    TokensPerSong = ActiveCell.FormulaR1C1
    Range("H10").Select
    EXPGain = ActiveCell.FormulaR1C1
    Range("H20").Select
    PointsPerSong = ActiveCell.FormulaR1C1
    
    Range("H5").Select
    ActiveCell.FormulaR1C1 = ActiveCell.FormulaR1C1 + EXPGain
    'Current Tokens
    Range("H16").Select
    ActiveCell.FormulaR1C1 = ActiveCell.FormulaR1C1 - (4 * TokensPerSong)
    'Current Points
    Range("H17").Select
    ActiveCell.FormulaR1C1 = ActiveCell.FormulaR1C1 + (4 * PointsPerSong)
End Sub
Sub TC_AddEXP_JP_Token_Hard()
'
' Add EXP and subtract tokens from EN Macro
'

'
    Dim TokensPerSong As Integer
    Dim EXPGain As Integer
    Dim PointsPerSong As Integer
    
    Range("I19").Select
    TokensPerSong = ActiveCell.FormulaR1C1
    Range("I10").Select
    EXPGain = ActiveCell.FormulaR1C1
    Range("I20").Select
    PointsPerSong = ActiveCell.FormulaR1C1
    
    Range("H5").Select
    ActiveCell.FormulaR1C1 = ActiveCell.FormulaR1C1 + EXPGain
    'Current Tokens
    Range("H16").Select
    ActiveCell.FormulaR1C1 = ActiveCell.FormulaR1C1 - TokensPerSong
    'Current Points
    Range("H17").Select
    ActiveCell.FormulaR1C1 = ActiveCell.FormulaR1C1 + PointsPerSong
End Sub
Sub TC_AddEXP_JP_Token_Hard_4x()
'
' Add EXP and subtract tokens from EN Macro
'

'
    Dim TokensPerSong As Integer
    Dim EXPGain As Integer
    Dim PointsPerSong As Integer
    
    Range("I19").Select
    TokensPerSong = ActiveCell.FormulaR1C1
    Range("I10").Select
    EXPGain = ActiveCell.FormulaR1C1
    Range("I20").Select
    PointsPerSong = ActiveCell.FormulaR1C1
    
    Range("H5").Select
    ActiveCell.FormulaR1C1 = ActiveCell.FormulaR1C1 + EXPGain
    'Current Tokens
    Range("H16").Select
    ActiveCell.FormulaR1C1 = ActiveCell.FormulaR1C1 - (4 * TokensPerSong)
    'Current Points
    Range("H17").Select
    ActiveCell.FormulaR1C1 = ActiveCell.FormulaR1C1 + (4 * PointsPerSong)
End Sub
Sub TC_AddEXP_JP_Token_Normal()
'
' Add EXP and subtract tokens from JP Macro
'

'
    Dim TokensPerSong As Integer
    Dim EXPGain As Integer
    Dim PointsPerSong As Integer
    
    Range("J19").Select
    TokensPerSong = ActiveCell.FormulaR1C1
    Range("J10").Select
    EXPGain = ActiveCell.FormulaR1C1
    Range("J20").Select
    PointsPerSong = ActiveCell.FormulaR1C1
    
    Range("H5").Select
    ActiveCell.FormulaR1C1 = ActiveCell.FormulaR1C1 + EXPGain
    'Current Tokens
    Range("H16").Select
    ActiveCell.FormulaR1C1 = ActiveCell.FormulaR1C1 - TokensPerSong
    'Current Points
    Range("H17").Select
    ActiveCell.FormulaR1C1 = ActiveCell.FormulaR1C1 + PointsPerSong
End Sub
Sub TC_AddEXP_JP_Token_Normal_4x()
'
' Add EXP and subtract tokens from JP Macro
'

'
    Dim TokensPerSong As Integer
    Dim EXPGain As Integer
    Dim PointsPerSong As Integer
    
    Range("J19").Select
    TokensPerSong = ActiveCell.FormulaR1C1
    Range("J10").Select
    EXPGain = ActiveCell.FormulaR1C1
    Range("J20").Select
    PointsPerSong = ActiveCell.FormulaR1C1
    
    Range("H5").Select
    ActiveCell.FormulaR1C1 = ActiveCell.FormulaR1C1 + EXPGain
    'Current Tokens
    Range("H16").Select
    ActiveCell.FormulaR1C1 = ActiveCell.FormulaR1C1 - (4 * TokensPerSong)
    'Current Points
    Range("H17").Select
    ActiveCell.FormulaR1C1 = ActiveCell.FormulaR1C1 + (4 * PointsPerSong)
End Sub
Sub TC_AddEXP_JP_Token_Easy()
'
' Add EXP and subtract tokens from JP Macro
'

'
    Dim TokensPerSong As Integer
    Dim EXPGain As Integer
    Dim PointsPerSong As Integer
    
    Range("K19").Select
    TokensPerSong = ActiveCell.FormulaR1C1
    Range("K10").Select
    EXPGain = ActiveCell.FormulaR1C1
    Range("K20").Select
    PointsPerSong = ActiveCell.FormulaR1C1
    
    Range("H5").Select
    ActiveCell.FormulaR1C1 = ActiveCell.FormulaR1C1 + EXPGain
    'Current Tokens
    Range("H16").Select
    ActiveCell.FormulaR1C1 = ActiveCell.FormulaR1C1 - TokensPerSong
    'Current Points
    Range("H17").Select
    ActiveCell.FormulaR1C1 = ActiveCell.FormulaR1C1 + PointsPerSong
End Sub
Sub TC_AddEXP_JP_Token_Easy_4x()
'
' Add EXP and subtract tokens from JP Macro
'

'
    Dim TokensPerSong As Integer
    Dim EXPGain As Integer
    Dim PointsPerSong As Integer
    
    Range("K19").Select
    TokensPerSong = ActiveCell.FormulaR1C1
    Range("K10").Select
    EXPGain = ActiveCell.FormulaR1C1
    Range("K20").Select
    PointsPerSong = ActiveCell.FormulaR1C1
    
    Range("H5").Select
    ActiveCell.FormulaR1C1 = ActiveCell.FormulaR1C1 + EXPGain
    'Current Tokens
    Range("H16").Select
    ActiveCell.FormulaR1C1 = ActiveCell.FormulaR1C1 - (4 * TokensPerSong)
    'Current Points
    Range("H17").Select
    ActiveCell.FormulaR1C1 = ActiveCell.FormulaR1C1 + (4 * PointsPerSong)
End Sub

