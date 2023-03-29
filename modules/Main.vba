
Public rinc As Integer, cinc As Integer

' Valori che vengono cambiati dalle impostazioni di gioco
Public colorFieldIndex As Integer, colorSnakeIndex As Integer, gameFieldSize As String, numberColByGameSize As Integer
' Valore per memmorizare l'index delle impostazioni scelte
Public comboBoxSnakeColorListIndex As Integer, comboBoxFieldColorListIndex As Integer, comboBoxGameSizeListIndex As Integer
' Valore del punteggio del giocatore
Public gameScore As Integer

Dim r() As Integer, c() As Integer

Sub StartGame()

    ' Funzioni che servono a resettare il colore e i bordi.
    ' Serve quando si passa da una dimensione del campo grande a una più piccola.
    Range("A1:AZ50").Interior.ColorIndex = 31
    Range("A1:AZ50").Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone
    
    
    ' Valori di default delle impostazioni di sistema.
    If colorFieldIndex = 0 Then
        colorFieldIndex = 4
    End If
    
    If colorSnakeIndex = 0 Then
        colorSnakeIndex = 7
    End If

    If gameFieldSize = "" Then
        gameFieldSize = "B2:P25"
    End If
    
    If numberColByGameSize = 0 Then
        numberColByGameSize = 15
    End If
    
    
    ' Colora un range celle in base alla dimensione del campo scelta.
    Range(gameFieldSize).Interior.ColorIndex = colorFieldIndex
    Range(gameFieldSize).Borders.LineStyle = xlContinuous
    
    ' Array per memmorizare le coordiante del corpo del serpente.
    ReDim r(2)
    ReDim c(2)
    
    ' Assegnare valori iniziali agli array del sepente.
    r(0) = 20: r(1) = 21: r(2) = 22
    c(0) = 10: c(1) = 10: c(2) = 10
    rinc = 0: cinc = 0
    
    ShowSnake
    bindKeys
    AddApple
    StartTimer

End Sub

Sub ShowSnake()
    ' Colorare le celle che corrispondo al corpo del serpente
    For i = UBound(r) To 0 Step -1
        Cells(r(i), c(i)).Interior.ColorIndex = colorSnakeIndex
    Next i
End Sub

Sub MoveSnake()
    If rinc <> 0 Or cinc <> 0 Then
        tail = UBound(r)
        Cells(r(tail), c(tail)).Interior.ColorIndex = colorFieldIndex
        
        For i = tail To 1 Step -1
            r(i) = r(i - 1)
            c(i) = c(i - 1)
        Next i
        
        r(0) = r(0) + rinc
        c(0) = c(0) + cinc
        
        ' Controllare se la "testa" del serpente coincide con la mela.
        If Cells(r(0), c(0)).Interior.Color = vbYellow Then
            gameScore = gameScore + 1
            apples = apples + 1
            ReDim Preserve r(UBound(r) + 1)
            ReDim Preserve c(UBound(c) + 1)
            r(UBound(r)) = r(UBound(r) - 1)
            c(UBound(c)) = c(UBound(c) - 1)
            AddApple
        ' If nel caso il serpente si scontrasse con se stesso o con il limite del campo.
        ElseIf Cells(r(0), c(0)).Interior.ColorIndex <> colorFieldIndex Then
            StopTimer
            GameOverWindow.Show
            Exit Sub
        End If
        
        ShowSnake
    End If
End Sub

Sub AddApple()
    ' Aggiunge mele nel campo randomicamente
    Randomize
        arow = (Rnd * 23) + 2
        
        acol = Int(Rnd() * numberColByGameSize) + 1
    Cells(arow, acol).Interior.Color = vbYellow
End Sub
