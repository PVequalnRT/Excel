Attribute VB_Name = "Module1"
Function Status(num)

'특정년도에서 현재 년도를 제하여 건물상태를 상, 중, 하로 나눠서 분류하는 함수

Dim val

val = 2018 - num

If val > 15 Then
    Status = "하"
ElseIf val > 5 Then
    Status = "중"
Else
    Status = "상"

End If
End Function

Sub 엑셀연습()
'엑셀 반복문 연습 HELLO WORLD x 1~10까지 반복출력

For i = 11 to 20
    Range("H"&i).Value = "Hello World! X " & i-10
Next 

End Sub

Sub 연습()

ActiveCell.Offset(5,0).Range("A1:A3").Select
ActiveCell.Range("A1:A3").Value = 5

End Sub


Sub color()
    x = 2
    Do While x > 1
        num1 = ActiveCell.offset(0,0).Value
        num2 = ActiveCell.Offset(0,1).Value

        if ActiveCell.offset(0,0).Value = "" Then
            MsgBox "채색이 완료되었습니다."
            Exit Sub

        ElseIf num1 > num2 Then
            ActiveCell.offset(0,0).Interior.color = RGB(255,0,0)
            ActiveCell.offset(1,0).Select

         ElseIf num1 < num2 Then
            ActiveCell.offset(0,1).Interior.color = RGB(255,0,0)
            ActiveCell.offset(1,0).Select

         Else
            ActiveCell.offset(0,0).Interior.color = RGB(255,0,0)
            ActiveCell.offset(0,1).Interior.color = RGB(255,0,0)
            ActiveCell.offset(1,0).Select

        End If

    Loop

End Sub