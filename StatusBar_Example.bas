Attribute VB_Name = "Module1"
Sub tstProgressBar()

' ���� ������¸� ǥ���� ���� ���̰� �մϴ�.
' ������ ��Ÿ�� Lable2�� �ʺ�� 0���� ����մϴ�.
' �׸��� Lable3������ �� �۾����� ���� ó���ڷḦ ���ڷ� ��Ÿ���ϴ�.
UserForm1.Label2.Width = 0
UserForm1.Label3 = "0 / 1000"
UserForm1.Show

' ó���ؾ� �� �۾��� 1000���� ������ ��Ȳ���� �����մϴ�.
For i = 1 To 1000
    ' �Ʒ��� ó���ϴ� �۾��� ��ġ�մϴ�.
    ' �� ���������� �ܼ��� ������ �����Ͽ� �̸� ����մϴ�.
    For j = 1 To 1000
    ' �� �������� ó�� ������ �Ʒ� ������ ���� �����Ͽ� ���� �� �ֽ��ϴ�.
        For k = 1 To 1000
        Next k
    Next j
    
    ' �̾ ������¸� ǥ���� ���� �����մϴ�.
    ' ����� �ξ��� ���̺� �ʺ� 414�� ���⼭ ���˴ϴ�.
    UserForm1.Label2.Width = Int(i / 1000 * 414)
    UserForm1.Label3 = Trim(i) + " / 1000"
    UserForm1.Repaint
Next i

' ��� �۾��� ������ ���� ����ϴ�.
UserForm1.Hide
    
End Sub
