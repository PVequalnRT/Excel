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