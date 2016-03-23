Private Sub CommandButton1_Click()
'заполняет элементы массива рандомными положительными и отрицательными значениями от -1000 до 1000'
For i = 1 To 20
Cells(1, i) = Int((2000 * Rnd) - 1000)
Next i
End Sub

Private Sub CommandButton2_Click()
'находит и выводит минимальное значение из всех нечетных элементов массива, которые делятся на 5'
Min = 1001
For i = 1 To 20
If Cells(1, i) Mod 2 <> 0 And Cells(1, i) Mod 5 = 0 And Cells(1, i) < Min Then
Min = Cells(1, i)
End If
Next i
MsgBox (Min)
End Sub

Private Sub CommandButton3_Click()
'закрывает форму'
UserForm1.Hide
End Sub