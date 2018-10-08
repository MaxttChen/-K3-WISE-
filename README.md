# 插件开发：

#老单据/工业单据插件：
主要使用 m_BillTransfer 对象：

先写一个函数便以查询列序号的方法：
```VB
Public Function GetCtlIndexByFld(ByVal fldName As String, Optional ByVal isEntry As Boolean = False) As Long
Dim ctlIdx As Long
Dim i As Integer
Dim isFind As Boolean
Dim vValue As Variant
fldName = UCase(fldName)
isFind = False
With m_BillTransfer
If isEntry Then
    For i = LBound(.EntryCtl) To UBound(.EntryCtl)
    If UCase(.EntryCtl(i).FieldName) = fldName Then
    ctlIdx = .EntryCtl(i).FCtlOrder
    isFind = True
    Exit For
    End If
    Next i
Else
    For i = LBound(.HeadCtl) To UBound(.HeadCtl)
    If UCase(.HeadCtl(i).FieldName) = fldName Then
    ctlIdx = .HeadCtl(i).FCtlIndex
    isFind = True
    Exit For
    End If
    Next i
End If
End With
If isFind = True Then
GetCtlIndexByFld = ctlIdx
Else
GetCtlIndexByFld = 0
End If
End Function
```
以下示例需求：输入整数，采购申请单的所有数量乘以这个倍数，实现BOM关联：
```VB
Private Sub m_BillTransfer_UserMenuClick(ByVal Index As Long, ByVal Caption As String)
 
    'TODO: 请在此处添加代码响应事件 UserMenuClick
 
 
    Select Case Caption
    Case "增加套数"
        '此处添加处理 增加套数 菜单对象的 Click 事件
        Dim strIputVal As String
        Dim currentrow As Integer
        Dim bom_count As Integer
        Dim FAuxQty As Long
        Dim FQty As Long
        
        FAuxQty = GetCtlIndexByFld("FAuxQty", True)
        FQty = GetCtlIndexByFld("FQty", True)
        
        currentrow = m_BillTransfer.BillForm.get_MaxEntry

        If currentrow > 0 Then
            strIputVal = InputBox("输入生成的套数", "BOM套数")
            If strIputVal = "" Or Not IsNumeric(strIputVal) Then
                bom_count = 1
            Else
                bom_count = strIputVal
            End If
            For i = 1 To currentrow
                m_BillTransfer.SetGridText i, FQty, Val(m_BillTransfer.GetGridText(i, FQty)) * bom_count
                m_BillTransfer.SetGridText i, FAuxQty, Val(m_BillTransfer.GetGridText(i, FAuxQty)) * bom_count
            Next
        End If
        
        
    Case Else
    End Select

End Sub
```
