Attribute VB_Name = "Module1"
Sub Player_1()
    Dim Column As Integer, RowIndex As Integer
    Dim X As Integer, Y As Integer
    Dim CountFour As Integer
    Dim Win As Boolean

    Do While True
        Column = InputBox("Your Turn Player 1!")
        If Column < 2 Or Column > 8 Then
            MsgBox ("Play inside the game [2 ; 8].")
        Else
            Exit Do
        End If
    Loop

    For RowIndex = 1 To 6
        If Cells(RowIndex, Column) = "" Then
            Cells(RowIndex, Column) = Range("L2")
            Cells(RowIndex, Column).Interior.Color = RGB(255, 0, 0)
            If RowIndex > 1 Then
                Cells(RowIndex - 1, Column) = ""
                Cells(RowIndex - 1, Column).Interior.Color = xlNone
            End If
        Else
            Exit For
        End If
        Application.Wait Now + TimeValue("00:00:01")
    Next

    X = RowIndex - 1
    Y = Column
    CountFour = 0
    Win = False
    Do While X >= 1 And Y >= 2 And X <= 6 And Y <= 8 And Cells(X, Y) = Range("L2") And Win = False
        CountFour = CountFour + 1
        X = X + 1
    Loop

    X = RowIndex - 1
    Y = Column
    Win = False
    Do While X > 1 And Y >= 2 And X <= 6 And Y <= 8 And Cells(X, Y) = Range("L2") And Win = False
        CountFour = CountFour + 1
        X = X - 1
    Loop

    If CountFour > 4 And Win = False Then
        MsgBox ("Player 1 Win!")
        Win = True
    End If

    X = RowIndex - 1
    Y = Column
    CountFour = 0
    Do While X >= 1 And Y >= 2 And X <= 6 And Y <= 8 And Cells(X, Y) = Range("L2") And Win = False
        CountFour = CountFour + 1
        Y = Y + 1
    Loop

    X = RowIndex - 1
    Y = Column
    Do While X >= 1 And Y >= 2 And X <= 6 And Y <= 8 And Cells(X, Y) = Range("L2") And Win = False
        CountFour = CountFour + 1
        Y = Y - 1
    Loop

    If CountFour > 4 And Win = False Then
        MsgBox ("Player 1 Win!")
        Win = True
    End If

    X = RowIndex - 1
    Y = Column
    CountFour = 0
    Do While X >= 1 And Y >= 2 And X <= 6 And Y <= 8 And Cells(X, Y) = Range("L2") And Win = False
        CountFour = CountFour + 1
        Y = Y + 1
        X = X + 1
    Loop

    X = RowIndex - 1
    Y = Column
    Do While X > 1 And Y >= 2 And X <= 6 And Y <= 8 And Cells(X, Y) = Range("L2") And Win = False
        CountFour = CountFour + 1
        Y = Y - 1
        X = X - 1
    Loop

    If CountFour > 4 And Win = False Then
        MsgBox ("Player 1 Win!")
        Win = True
    End If

    X = RowIndex - 1
    Y = Column
    CountFour = 0
    Do While X > 1 And Y >= 2 And X <= 6 And Y <= 8 And Cells(X, Y) = Range("L2") And Win = False
        CountFour = CountFour + 1
        Y = Y + 1
        X = X - 1
    Loop

    X = RowIndex - 1
    Y = Column
    Do While X >= 1 And Y >= 2 And X <= 6 And Y <= 8 And Cells(X, Y) = Range("L2") And Win = False
        CountFour = CountFour + 1
        Y = Y - 1
        X = X + 1
    Loop

    If CountFour > 4 And Win = False Then
        MsgBox ("Player 1 Win!")
        Win = True
    End If

    If Win = True Then Cells(3, 10).Value = Cells(3, 10).Value + 1
    If Cells(3, 10).Value = 3 Then MsgBox ("Player 1 is the winner!!!")
End Sub
Sub Player_2()
    Dim Column As Integer, RowIndex As Integer

    Do While True
        Column = InputBox("Your Turn Player 2!")
        If Column < 2 Or Column > 8 Then
            MsgBox ("Play inside the game [2 ; 8].")
        Else
            Exit Do
        End If
    Loop

    For RowIndex = 1 To 6
        If Cells(RowIndex, Column) = "" Then
            Cells(RowIndex, Column) = Range("L4")
            Cells(RowIndex, Column).Interior.Color = RGB(255, 255, 0)
            If RowIndex > 1 Then
                Cells(RowIndex - 1, Column) = ""
                Cells(RowIndex - 1, Column).Interior.Color = xlNone
            End If
        Else
            Exit For
        End If
        Application.Wait Now + TimeValue("00:00:01")
    Next

    X = RowIndex - 1
    Y = Column
    CountFour = 0
    Win = False
    Do While X >= 1 And Y >= 2 And X <= 6 And Y <= 8 And Cells(X, Y) = Range("L4") And Win = False
        CountFour = CountFour + 1
        X = X + 1
    Loop

    X = RowIndex - 1
    Y = Column
    Win = False
    Do While X > 1 And Y >= 2 And X <= 6 And Y <= 8 And Cells(X, Y) = Range("L4") And Win = False
        CountFour = CountFour + 1
        X = X - 1
    Loop

    If CountFour > 4 And Win = False Then
        MsgBox ("Player 2 Win!")
        Win = True
    End If

    X = RowIndex - 1
    Y = Column
    CountFour = 0
    Do While X >= 1 And Y >= 2 And X <= 6 And Y <= 8 And Cells(X, Y) = Range("L4") And Win = False
        CountFour = CountFour + 1
        Y = Y + 1
    Loop

    X = RowIndex - 1
    Y = Column
    Do While X >= 1 And Y >= 2 And X <= 6 And Y <= 8 And Cells(X, Y) = Range("L4") And Win = False
        CountFour = CountFour + 1
        Y = Y - 1
    Loop

    If CountFour > 4 And Win = False Then
        MsgBox ("Player 2 Win!")
        Win = True
    End If

    X = RowIndex - 1
    Y = Column
    CountFour = 0
    Do While X >= 1 And Y >= 2 And X <= 6 And Y <= 8 And Cells(X, Y) = Range("L4") And Win = False
        CountFour = CountFour + 1
        Y = Y + 1
        X = X + 1
    Loop

    X = RowIndex - 1
    Y = Column
    Do While X > 1 And Y >= 2 And X <= 6 And Y <= 8 And Cells(X, Y) = Range("L4") And Win = False
        CountFour = CountFour + 1
        Y = Y - 1
        X = X - 1
    Loop

    If CountFour > 4 And Win = False Then
        MsgBox ("Player 2 Win!")
        Win = True
    End If

    X = RowIndex - 1
    Y = Column
    CountFour = 0
    Do While X > 1 And Y >= 2 And X <= 6 And Y <= 8 And Cells(X, Y) = Range("L4") And Win = False
        CountFour = CountFour + 1
        Y = Y + 1
        X = X - 1
    Loop

    X = RowIndex - 1
    Y = Column
    Do While X >= 1 And Y >= 2 And X <= 6 And Y <= 8 And Cells(X, Y) = Range("L4") And Win = False
        CountFour = CountFour + 1
        Y = Y - 1
        X = X + 1
    Loop

    If CountFour > 4 And Win = False Then
        MsgBox ("Player 2 Win!")
        Win = True
    End If
    
    If Win = True Then Cells(5, 10).Value = Cells(5, 10).Value + 1
    If Cells(5, 10).Value = 3 Then MsgBox ("Player 2 is the winner!!!")
End Sub
