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