' Sub per collegare il tasto alla funzione di movimento
Sub bindKeys()
    Application.OnKey "{LEFT}", "moveLeft"
    Application.OnKey "{RIGHT}", "moveRight"
    Application.OnKey "{UP}", "moveUp"
    Application.OnKey "{DOWN}", "moveDown"
End Sub

Sub moveLeft()
    If cinc <> 1 Then
        cinc = -1
        rinc = 0
        MoveSnake
    End If
End Sub

Sub moveRight()
    If cinc <> -1 Then
        cinc = 1
        rinc = 0
        MoveSnake
    End If
End Sub

Sub moveUp()
    If rinc <> 1 Then
        cinc = 0
        rinc = -1
        MoveSnake
    End If
End Sub

Sub moveDown()
    If rinc <> -1 Then
        cinc = 0
        rinc = 1
        MoveSnake
    End If
End Sub

' Serve per disabbinare le funzioni con le frecce.
Sub freeKeys()
    Application.OnKey "{LEFT}"
    Application.OnKey "{RIGHT}"
    Application.OnKey "{UP}"
    Application.OnKey "{DOWN}"
End Sub
