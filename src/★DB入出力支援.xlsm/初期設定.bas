Attribute VB_Name = "�����ݒ�"
Public globalWb�O����s���� As Workbook

Sub �O���[�v�\����\���̐؂�ւ�()

    Dim obj�����f�[�^�V�[�g As cls�����f�[�^�V�[�g
    Set obj�����f�[�^�V�[�g = New cls�����f�[�^�V�[�g
    
    obj�����f�[�^�V�[�g.�\����\���̐ؑ�
    
End Sub

Sub createSELECT��()

    frmSQL����.Show vbModeless
    
    Excel.Application.CutCopyMode = False
    
End Sub

Sub set���o����()

    Dim obj�ݒ�V�[�g As cls�ݒ�V�[�g
    Set obj�ݒ�V�[�g = New cls�ݒ�V�[�g
    
    obj�ݒ�V�[�g.set���o����

End Sub

Sub Auto_Open()

    Application.OnKey "{F9}", "createSELECT��"

End Sub
