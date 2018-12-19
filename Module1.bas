Attribute VB_Name = "Module1"
Sub CombineString() '�� 2���� ���� ������ ���� ���� ���� ����ִ� ��ũ��

Dim x As Long
x = 2

Do
    
    If ActiveCell.Offset(0, -113).Value = "" Then
        MsgBox "�Ϸ�"
        Exit Sub
    End If
    
    

    Dim cha As String
    cha = ActiveCell.Offset(0, -113).Value & " " & ActiveCell.Offset(0, -112).Value
    ActiveCell.Offset(0, 0).Value = cha

    ActiveCell.Offset(1, 0).Select

Loop

End Sub


Function BuilingStatus(num)
'Ư���⵵���� ���� �⵵�� ���Ͽ� �ǹ����¸� ��, ��, �Ϸ� ������ �з��ϴ� �Լ�

Dim val As Byte


val = 2018 - num

If val > 15 Then
    BuilingStatus = "��"

ElseIf val > 5 Then
    BuilingStatus = "��"

Else
    BuilingStatus = "��"
End If
    
End Function



Sub CombineStringWithHyphen()

Dim cha As String
Dim t As Long

t = 1


Do
    If ActiveCell.Offset(0, 0).Value = "" Then
        MsgBox "�Ϸ�"
        Exit Sub
    End If
    
    If t = 1 Then
        cha = ActiveCell.Offset(0, 0).Value
        Range("d1").Value = "'" & cha
    
    Else
        cha = ActiveCell.Offset(0, 0).Value
    
        Range("d1").Value = Range("d1").Value & "-" & cha
    
    
    End If

    ActiveCell.Offset(1, 0).Select
    t = t + 1
Loop


End Sub

Sub ��������()
'���� �ݺ��� ���� Hello World x 1 ~ 10���� �ݺ��ؼ� ���

For i = 11 To 20
    Range("H" & i).Value = "Hello World! x " & i - 10

Next
    
End Sub

Sub ����()

ActiveCell.Offset(5, 0).Range("A1:A3").Select
ActiveCell.Range("A1:A3").Value = 5

End Sub


Sub CompareNumAndColoring() '�� ��ġ�� ���ؼ� ū ���� ��ĥ���ִ� ��ũ��

x = 2

Do While x > 1

    num1 = ActiveCell.Offset(0, 0).Value
    num2 = ActiveCell.Offset(0, 1).Value

    If ActiveCell.Offset(0, 0).Value = "" Then
        MsgBox "ä���� �Ϸ�Ǿ����ϴ�."
        Exit Sub
    
    ElseIf num1 > num2 Then
        ActiveCell.Offset(0, 0).Interior.color = RGB(255, 0, 0)
        ActiveCell.Offset(1, 0).Select

    ElseIf num1 < num2 Then
        ActiveCell.Offset(0, 1).Interior.color = RGB(255, 0, 0)
        ActiveCell.Offset(1, 0).Select
    

    Else
        ActiveCell.Offset(0, 0).Interior.color = RGB(255, 0, 0)
        ActiveCell.Offset(0, 1).Interior.color = RGB(255, 0, 0)
        ActiveCell.Offset(1, 0).Select

    End If


Loop
End Sub


Function BuilingStatus2(num)
'Ư���⵵���� ���� �⵵�� ���Ͽ� �ǹ����¸� ��, ��, �Ϸ� ������ �з��ϴ� �Լ�

Dim val

val = 2018 - num

If val > 15 Then
    ActiveCell.Interior.color = RGB(0, 255, 0)
    BuilingStatus2 = "��"
    

ElseIf val > 5 Then
    BuilingStatus2 = "��"

Else
    ActiveCell.Interior.color = RGB(255, 255, 0)
    BuilingStatus2 = "��"
    
    
End If
    
End Function


Sub CompareNumAndColoring2()

Dim cha As String
Dim x As Byte
x = 2

Do While x > 1
    
    If ActiveCell.Offset(0, -1).Value = "" Then
     MsgBox "ä���� �Ϸ�Ǿ����ϴ�."
     Exit Sub
    End If
    
    cha = BuilingStatus(ActiveCell.Offset(0, -1).Value)

    If cha = "��" Then
     ActiveCell.Offset(0, 0).Interior.color = RGB(0, 255, 0)
     ActiveCell.Offset(0, 0).Value = cha
     Selection.Font.Bold = True
    
    ElseIf cha = "��" Then
     ActiveCell.Offset(0, 0).Value = cha
     Selection.Font.Bold = True
    
    ElseIf cha = "��" Then
     ActiveCell.Offset(0, 0).Interior.color = RGB(255, 255, 0)
     ActiveCell.Offset(0, 0).Value = cha
     Selection.Font.Bold = True
   
    
    End If

ActiveCell.Offset(1, 0).Select

Loop

End Sub

Sub SumNum()

Dim num1, num2, t As Long
t = 1

Do
    If ActiveCell.Offset(0, -1).Value = "" Then
        Exit Sub
    End If
    
    If t = 1 Then
        ActiveCell.Offset(0, 0).Value = ActiveCell.Offset(0, -1).Value * 2
    Else
        num1 = ActiveCell.Offset(-1, 0).Value
        num2 = ActiveCell.Offset(0, -1).Value * 2
        
        ActiveCell.Offset(0, 0).Value = num1 + num2
    
    End If
    
    t = t + 1
    ActiveCell.Offset(1, 0).Select
Loop

End Sub

Sub ����()
Call SumNum
Range("b1").Select
Call CombineStringWithHyphen
End Sub
