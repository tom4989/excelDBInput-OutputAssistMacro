Attribute VB_Name = "初期設定"
Public globalWb前回実行結果 As Workbook

Sub グループ表示非表示の切り替え()

    Dim obj試験データシート As cls試験データシート
    Set obj試験データシート = New cls試験データシート
    
    obj試験データシート.表示非表示の切替
    
End Sub

Sub createSELECT文()

    frmSQL生成.Show vbModeless
    
    Excel.Application.CutCopyMode = False
    
End Sub

Sub set抽出条件()

    Dim obj設定シート As cls設定シート
    Set obj設定シート = New cls設定シート
    
    obj設定シート.set抽出条件

End Sub

Sub Auto_Open()

    Application.OnKey "{F9}", "createSELECT文"

End Sub
