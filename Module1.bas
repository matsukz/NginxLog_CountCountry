Attribute VB_Name = "Module1"
Sub Refresh()
    
    '��Ƃ��s�����[�N�V�[�g���A�N�e�B�u�ɂ���
    Worksheets("Database").Select
    
    '���݂̃f�[�^�����J�E���g����
    Dim datacount As Long
    datacount = 0
    
    datacount = Worksheets("Database").ListObjects("NginxLog").ListRows.Count

    'MySQL�ɐڑ��ł��邩�e�X�g����
    Select Case request("http://192.168.11.15:5500/port?ip=100.96.0.1&port=3306&option=2")
        Case 0
            '�f�[�^�\�[�X���X�V����
            ActiveWorkbook.RefreshAll
        Case 1
            MsgBox "MySQL�T�[�o�[�ւ̐ڑ������Ɏ��s���܂����B" & vbLf & "���Ԃ��󂯂čĎ��s���Ă��������B"
            Worksheets("Dashboard").Select
            Exit Sub
        Case 2
            MsgBox "Flask����̉���������܂���ł���"
            Worksheets("Dashboard").Select
            Exit Sub
        Case 3
            MsgBox "�s���ȃG���[���������܂���"
            Worksheets("Dashboard").Select
            Exit Sub
    End Select
        
    
    '�ďƉ��̃f�[�^���J�E���g����
    Dim after_datacount As Long
    after_datacount = 0
    
    after_datacount = Worksheets("Database").ListObjects("NginxLog").ListRows.Count
    
    Debug.Print (datacount)
    Debug.Print (after_datacount)
    
    '�~�O���t�X�V
    Module2.Date_Country
    
    '�s�|�b�g�e�[�u���̍X�V
    
    'msg�̕\��
    Dim msg As String
    msg = "�X�V���������܂����B" & vbCrLf
    If after_datacount - datacount = 0 Then
        msg = msg & "�V�������R�[�h�͂���܂���"
    Else
        msg = msg & after_datacount - datacount & " ���ǉ�����܂���"
    End If
    
    Worksheets("Dashboard").Select
    
    MsgBox msg, vbInformation
    
End Sub

Function request(ByVal point As String) As Integer

    '�G���[�n���h�����O
    On Error GoTo errorHandler
    
    '�c�[���@�Q�Ɛݒ肩��"Microsoft XML v6.0"��L���ɂ��邱�ƁI�I�I
    Dim HttpReq As Object
    Set HttpReq = CreateObject("MSXML2.XMLHTTP")
    
    Dim response As Boolean: response = False
    
    '���N�G�X�g�쐬 ��O������True(�񓯊�) False(����)
    '�L���b�V���h�~�̈׃^�C���X�^���v������(Flask�ł͖���)
    Dim timestamp As String: timestamp = Format(Now, "yyyymmddhhmmss")
    HttpReq.Open "GET", point & "&nocache=" & timestamp, False
    HttpReq.send
    
    'sub�v���V�[�W���Ɍ��ʂ�n������
    response = HttpReq.responseText
    If response = True Then
        request = 0 '����ȏI���R�[�h
    ElseIf response = False Then
        request = 1 'Flask�ɂ͓��B������socket�ʐM�Ɏ��s����
    Else
        request = 3 '���̑��G���[
    End If
    
    GoTo cleanUP
    
errorHandler:
    '�����𗘗p����Ƃ���Flask�������Ă���Ƃ�
    MsgBox "Error!! " & Err.Description & "API�T�[�o�[���͐���ł����H"
    request = 3
    Set HttpReq = Nothing '�I�u�W�F�N�g���

cleanUP:
    Set HttpReq = Nothing '�I�u�W�F�N�g���
    
End Function
