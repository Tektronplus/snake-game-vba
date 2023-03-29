Private Sub ComboBoxFieldColor_Change()
    Select Case ComboBoxFieldColor.Value
        Case "Verde"
            colorFieldIndex = 4
            comboBoxFieldColorListIndex = 0
        Case "Rosso"
            colorFieldIndex = 3
            comboBoxFieldColorListIndex = 1
        Case "Blu"
            colorFieldIndex = 5
            comboBoxFieldColorListIndex = 2
        Case "Nero"
            colorFieldIndex = 1
            comboBoxFieldColorListIndex = 3
    End Select
End Sub

Private Sub ComboBoxSnakeColor_Change()
    Select Case ComboBoxSnakeColor.Value
        Case "Magenta"
            colorSnakeIndex = 26
            comboBoxSnakeColorListIndex = 0
        Case "Ciano"
            colorSnakeIndex = 8
            comboBoxSnakeColorListIndex = 1
        Case "Arancione"
            colorSnakeIndex = 46
            comboBoxSnakeColorListIndex = 2
        Case "Bianco"
            colorSnakeIndex = 2
            comboBoxSnakeColorListIndex = 3
    End Select
End Sub

Private Sub ComboBoxGameSize_Change()
    Select Case ComboBoxGameSize.Value
        Case "Piccolo"
            gameFieldSize = "B2:P25"
            numberColByGameSize = 15
            comboBoxGameSizeListIndex = 0
        Case "Medio"
            numberColByGameSize = 25
            gameFieldSize = "B2:Z25"
            comboBoxGameSizeListIndex = 1
        Case "Grande"
            numberColByGameSize = 35
            gameFieldSize = "B2:AJ25"
            comboBoxGameSizeListIndex = 2
    End Select
End Sub

Private Sub UserForm_Initialize()
    
    With ComboBoxGameSize
        .AddItem "Piccolo"
        .AddItem "Medio"
        .AddItem "Grande"
    End With
    
    
    With ComboBoxFieldColor
        .AddItem "Verde"
        .AddItem "Rosso"
        .AddItem "Blu"
        .AddItem "Nero"
    End With
    
    With ComboBoxSnakeColor
        .AddItem "Magenta"
        .AddItem "Ciano"
        .AddItem "Arancione"
        .AddItem "Bianco"
    End With
    
    ComboBoxSnakeColor.ListIndex = comboBoxSnakeColorListIndex
    ComboBoxFieldColor.ListIndex = comboBoxFieldColorListIndex
    ComboBoxGameSize.ListIndex = comboBoxGameSizeListIndex
End Sub
