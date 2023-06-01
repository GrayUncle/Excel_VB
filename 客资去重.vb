Sub DeleteDuplicates()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim uniqueValues As Object
    Dim cellValue As Variant

    ' 设置工作表
    Set ws = ThisWorkbook.Worksheets("数据处理")

    ' 获取最后一行
    lastRow = ws.Cells(ws.Rows.Count, "J").End(xlUp).Row

    ' 创建字典对象用于存储唯一值
    Set uniqueValues = CreateObject("Scripting.Dictionary")

    ' 从最后一行向上遍历
    For i = lastRow To 2 Step -1
        cellValue = ws.Cells(i, "J").Value
        
        ' 检查字典中是否已存在该值
        If Not uniqueValues.Exists(cellValue) Then
            ' 如果不存在，则将该值添加到字典中
            uniqueValues.Add cellValue, 1
        Else
            ' 如果存在，则删除当前行的J列内容
            ws.Cells(i, "J").ClearContents
        End If
    Next i
End Sub
' 规则：从最上方开始删除重复，并保留唯一项，工作表数据列为J列
