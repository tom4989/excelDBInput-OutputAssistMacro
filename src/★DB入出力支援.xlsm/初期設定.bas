Attribute VB_Name = "�����ݒ�"
Public globalWb�O����s���� As Workbook

Private obj�ݒ�l�V�[�g As cls�ݒ�l�V�[�g

Sub �O���[�v�\����\���̐؂�ւ�()

    Dim obj�����f�[�^�V�[�g As cls�����f�[�^�V�[�g
    Set obj�����f�[�^�V�[�g = New cls�����f�[�^�V�[�g
    
    obj�����f�[�^�V�[�g.�\����\���̐ؑ�
    
End Sub

Sub createSELECT��()

    frmSQL����.Show vbModeless
    
    Call frmSQL����.�ݒ�l���[�h(obj�ݒ�l�V�[�g)
    
    Excel.Application.CutCopyMode = False
    
End Sub

Sub Auto_Open()

    Set obj�ݒ�l�V�[�g = New cls�ݒ�l�V�[�g
    obj�ݒ�l�V�[�g.���[�h

    Application.OnKey "{F9}", "createSELECT��"

End Sub
