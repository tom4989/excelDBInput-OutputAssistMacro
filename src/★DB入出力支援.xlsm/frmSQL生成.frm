VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSQL���� 
   Caption         =   "���݂̃V�[�g"
   ClientHeight    =   8388
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   12840
   OleObjectBlob   =   "frmSQL����.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "frmSQL����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public wb�N�����u�b�N As Workbook
Public wb�O����s���� As Workbook

Private txtSQL�쐬���V�[�g�� As String
Private txt�g�����U�N�V�������ʕ����� As String
Private txtSQL���s�o�b�`�t�@�C���p�X As String
Private txtSQL���s���O�t�@�C���p�X As String

Private obj�ݒ�l�V�[�g As cls�ݒ�l�V�[�g

' *********************************************************************************************************************
' * �@�\�@�F�t�H�[������������
' *********************************************************************************************************************
'
Private Sub UserForm_Initialize()

    Set wb�N�����u�b�N = ActiveWorkbook
    
    btnSQL����.SetFocus
    
End Sub

' *********************************************************************************************************************
' * �@�\�@�F�t�H�[������������
' *********************************************************************************************************************
'
Public Sub �ݒ�l���[�h(arg�ݒ�l�V�[�g As cls�ݒ�l�V�[�g)

    Set obj�ݒ�l�V�[�g = arg�ݒ�l�V�[�g
    
    Dim txt�ڑ���� As Variant
    
    If Not obj�ݒ�l�V�[�g Is Nothing Then
    For Each txt�ڑ���� In obj�ݒ�l�V�[�g.�ݒ�l���X�g.Item("�ڑ����")
    
        cmb�ڑ����.AddItem (txt�ڑ����)
        
    Next txt�ڑ����
        
    cmb�ڑ����.ListIndex = 0
    
    End If
    
End Sub

' *********************************************************************************************************************
' * �@�\�@�FSQL�����{�^���������̏���
' *********************************************************************************************************************
'
Private Sub btnSQL����_Click()

    '�G���[�g���b�v
    On Error GoTo ErrorCatch

    txb�X�e�[�^�X�o�[.Value = get�J�n���b�Z�[�W("SQL����")

    Dim obj�A�v���P�[�V�������� As New cls�A�v���P�[�V��������
    obj�A�v���P�[�V��������.�A�v���P�[�V��������ؑ� (False)

    wb�N�����u�b�N.Activate
    
    Dim obj�����f�[�^�V�[�g As cls�����f�[�^�V�[�g
    Set obj�����f�[�^�V�[�g = New cls�����f�[�^�V�[�g
    Call obj�����f�[�^�V�[�g.������(obj�ݒ�l�V�[�g, cmb�ڑ����.Text)
    
    Dim obj�ΏۃV�[�g As Worksheet
    Set obj�ΏۃV�[�g = ActiveSheet
    txtSQL�쐬���V�[�g�� = ActiveSheet.Name
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    txt�g�����U�N�V�������ʕ����� = FSO.getBaseName(wb�N�����u�b�N.Name) & "_" & txtSQL�쐬���V�[�g�� & "_" & getTimestamp
    
    txtSQL.Value = ""
    frmSQL����.Repaint
    
    ' SPOOL���̏o��
    If ckbSPOOL Then
        txtSQL.Value = txtSQL.Value & vbCrLf & _
            "SPOOL """ & txt�g�����U�N�V�������ʕ����� & ".sql,log" & vbCrLf & vbCrLf
    End If
    
    If True Then
    
        txtSQL.Value = vbCrLf & "-- �V�[�g���F" & obj�ΏۃV�[�g.Name & vbLf & _
            createSQL��(obj�����f�[�^�V�[�g, obj�ΏۃV�[�g)
    Else
    
        Dim strSQL As String
        
        For Each obj�ΏۃV�[�g In ActiveWorkbook.Sheets
        
            obj�ΏۃV�[�g.Activate
            
            strSQL = strSQL & vbLf & "-- �V�[�g���F" & obj�ΏۃV�[�g.Name
            strSQL = strSQL & createSQL��(obj�����f�[�^�V�[�g, obj�ΏۃV�[�g)
            
        Next
        
        txtSQL.Value = strSQL
    
    End If
    
    txb�X�e�[�^�X�o�[.Value = get�I�����b�Z�[�W("SQL����")
    
   ' ����I��
   GoTo Finally
    
ErrorCatch:

    txb�X�e�[�^�X�o�[.Value = get�I�����b�Z�[�W("SQL����")

Finally:

   '���s�O�̏�Ԃɖ߂�
   obj�A�v���P�[�V��������.�A�v���P�[�V��������ؑ� (True)
        
End Sub

Private Function createSQL��( _
    ByRef obj�����f�[�^�V�[�g As cls�����f�[�^�V�[�g, _
    ByRef obj�ΏۃV�[�g As Worksheet) As String
    
    If rdInsert Then
        createSQL�� = obj�����f�[�^�V�[�g.�ΏۃV�[�gSQL���쐬(obj�ΏۃV�[�g, SQL���.INSERT��, rdb�I���s�̂�)
        
    ElseIf rdUpdate Then
        createSQL�� = obj�����f�[�^�V�[�g.�ΏۃV�[�gSQL���쐬(obj�ΏۃV�[�g, SQL���.UPDATE��, rdb�I���s�̂�)
        
    ElseIf rdSELECT Then
        createSQL�� = obj�����f�[�^�V�[�g.�ΏۃV�[�gSQL���쐬(obj�ΏۃV�[�g, SQL���.SELECT��, rdb�I���s�̂�)
        
    ElseIf rdDelete Then
        createSQL�� = obj�����f�[�^�V�[�g.�ΏۃV�[�gSQL���쐬(obj�ΏۃV�[�g, SQL���.DELETE��, rdb�I���s�̂�)
        
    End If
    
End Function
        
Private Sub btn���R�[�h�擾_Click()
        
    Call ���R�[�h�擾�{�^��
        
    Unload frmSQL����
        
End Sub
        
Private Sub ���R�[�h�擾�{�^��()

    txb�X�e�[�^�X�o�[.Value = get�J�n���b�Z�[�W("���R�[�h�擾")
    
    '�G���[�g���b�v
    On Error GoTo ErrorCatch

    Dim obj�A�v���P�[�V�������� As New cls�A�v���P�[�V��������
    obj�A�v���P�[�V��������.�A�v���P�[�V��������ؑ� (False)
        
    ' �{����
        
    wb�N�����u�b�N.Activate
    
    Dim obj�����f�[�^�V�[�g As cls�����f�[�^�V�[�g
    Set obj�����f�[�^�V�[�g = New cls�����f�[�^�V�[�g
    Call obj�����f�[�^�V�[�g.������(obj�ݒ�l�V�[�g, cmb�ڑ����.Text)
    
    Set wb�O����s���� = obj�����f�[�^�V�[�g.get���R�[�h(Nothing)
    
    If Not (wb�O����s���� Is Nothing) Then
        Application.CutCopyMode = False
        
        wb�O����s����.Activate
        wb�O����s����.ActiveSheet.Range("A1").Select
        
        btn���R�[�h�ǉ��擾.Enabled = True
        
    End If
        
    txb�X�e�[�^�X�o�[.Value = get�I�����b�Z�[�W("���R�[�h�擾")
        
   ' ����I��
   GoTo Finally
    
ErrorCatch:

Finally:

    txb�X�e�[�^�X�o�[.Value = get�ُ펞���b�Z�[�W("���R�[�h�擾")

    '���s�O�̏�Ԃɖ߂�
    obj�A�v���P�[�V��������.�A�v���P�[�V��������ؑ� (True)

End Sub

Private Sub btn���ʂ��t�@�C���ɏo��_Click()

    mkdirIFNotExist obj�ݒ�l�V�[�g.�ݒ�l���X�g.Item("���ʃt�@�C���o�͐�")

    Dim txt�o�̓p�X As String
    txt�o�̓p�X = obj�ݒ�l�V�[�g.�ݒ�l���X�g.Item("���ʃt�@�C���o�͐�") & "\" & txtSQL�쐬���V�[�g�� & "_" & getTimestamp() & ".sql"

    Open txt�o�̓p�X For Output As #1

    Print #1, txtSQL.Value

    Close #1

    txb�X�e�[�^�X�o�[ = txt�o�̓p�X

End Sub

Private Sub btn�X�V�O���R�[�h�擾_Click()

    ���R�[�h�擾�{�^��
    btn�X�V�ヌ�R�[�h�擾.Enabled = True

End Sub

Private Sub btn�X�V�ヌ�R�[�h�擾_Click()

    txb�X�e�[�^�X�o�[.Value = get�J�n���b�Z�[�W("�X�V�ヌ�R�[�h�擾")
    
    '�G���[�g���b�v
    On Error GoTo ErrorCatch

    Dim obj�A�v���P�[�V�������� As New cls�A�v���P�[�V��������
    obj�A�v���P�[�V��������.�A�v���P�[�V��������ؑ� (False)
        
    ' �{����
    wb�N�����u�b�N.Activate
    
    Dim obj�����f�[�^�V�[�g As cls�����f�[�^�V�[�g
    Set obj�����f�[�^�V�[�g = New cls�����f�[�^�V�[�g
    
    Set wb�O����s���� = obj�����f�[�^�V�[�g.get���R�[�h(wb�O����s����)
    
    If Not (wb�O����s���� Is Nothing) Then
        Application.CutCopyMode = False
        
        wb�O����s����.Activate
        wb�O����s����.ActiveSheet.Range("A1").Select
        
        Call obj�����f�[�^�V�[�g.edit���s���ʍ���(wb�O����s����)
    
    End If

    txb�X�e�[�^�X�o�[.Value = get�I�����b�Z�[�W("�X�V�ヌ�R�[�h�擾")

    btn�X�V�ヌ�R�[�h�擾.Enabled = False

   ' ����I��
   GoTo Finally
    
ErrorCatch:

Finally:

    txb�X�e�[�^�X�o�[.Value = get�ُ펞���b�Z�[�W("���R�[�h�擾")

    '���s�O�̏�Ԃɖ߂�
    obj�A�v���P�[�V��������.�A�v���P�[�V��������ؑ� (True)

End Sub

