Attribute VB_Name = "Module1"

Function Status(num)

'Ư���⵵���� ���� �⵵�� ���Ͽ� �ǹ����¸� ��, ��, �Ϸ� ������ �з��ϴ� �Լ�

Dim val

val = 2018 - num

If val > 15 Then
    Status = "��"
ElseIf val > 5 Then
    Status = "��"
Else
    Status = "��"

End If
End Function

Sub ��������()
'���� �ݺ��� ���� HELLO WORLD x 1~10���� �ݺ����

For i = 11 to 20
    Range("H"&i).Value = "Hello World! X " & i-10
Next 

End Sub

Sub ����()

ActiveCell.Offset(5,0).Range("A1:A3").Select
ActiveCell.Range("A1:A3").Value = 5

End Sub


Sub color()
    x = 2
    Do While x > 1
        num1 = ActiveCell.offset(0,0).Value
        num2 = ActiveCell.Offset(0,1).Value

        if ActiveCell.offset(0,0).Value = "" Then
            MsgBox "ä���� �Ϸ�Ǿ����ϴ�."
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

Sub Color2()   '�ǹ��� �⵵�� ���ϰ�, �⵵�� ���� �� �� �Ϸ� ������, ���� ���, ���� �Ͼ�, �ϴ� �ʷ����� ä���ϴ� ���α׷�

Dim cha As String
dim x As Byte
x = 2

Do While x > 1
    if ActiveCell.Offset(0,-1).Value = "" Then
     MsgBox "ä���� �Ϸ�Ǿ����ϴ�."
     Exit sub
    End If

    cha = Status(ActiveCell.Offset(0,-1).Value)

    If cha = "��" Then
     ActiveCell.Offset(0,0).Interior.color = RGB(0,255,0)
     ActiveCell.Offset(0,0).Value = cha
     Selection.Font.Bold = True

   ElseIf cha = "��" Then
     ActiveCell.Offset(0,0).Value = cha
     Selection.Font.Bold = True

    ElseIf cha "��" Then
     ActiveCell.Offset(0,0).Interior.color = RGB(255,255,0)
     ActiveCell.Offset(0,0).Value = cha
     Selection.Font.Bold = True

    End If

ActiveCell.Offset(0,0).Select

Loop
End sub