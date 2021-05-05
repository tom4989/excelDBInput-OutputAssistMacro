Attribute VB_Name = "初期設定"
Public globalWb前回実行結果 As Workbook

Private obj設定値シート As cls設定値シート

Sub グループ表示非表示の切り替え()

    Dim obj試験データシート As cls試験データシート
    Set obj試験データシート = New cls試験データシート
    
    obj試験データシート.表示非表示の切替
    
End Sub

Sub createSELECT文()

    frmSQL生成.Show vbModeless
    
    Call frmSQL生成.設定値ロード(obj設定値シート)
    
    Excel.Application.CutCopyMode = False
    
End Sub

Sub Auto_Open()

    Set obj設定値シート = New cls設定値シート
    obj設定値シート.ロード

    Application.OnKey "{F9}", "createSELECT文"

End Sub
