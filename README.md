# VBA-first-seen-get
Public Sub ValidatePersonName()

'假定文件格式为
'用户名     客户经理   客户的提示  客户经理的提示
'假定数据从第2行、第1列开始

Dim xlApp As Excel.Application
Dim xlBook As Excel.Workbook
Dim sheet As Excel.Worksheet

Set xlApp = New Excel.Application
Set xlBook = ThisWorkbook
'set xlbook = xlapp.Workbooks.Open("");
Set sheet = xlBook.ActiveSheet

Dim first_column  As Long
first_column = 1

Dim row As Long
Dim column As Long
Dim row_cnt As Long
row_cnt = LastRow(sheet)

rows_loop:
For row = 2 To row_cnt Step 1
row_proc_start:
 Dim warning As String
 Dim count As Long
 
 Dim clt As String
 Dim clt_phone As String
 clt = Trim(sheet.Cells(row, first_column))
 clt_phone = Trim(sheet.Cells(row, first_column + 2))  '核查输入的客户，客户经理，以及手机号码
 If clt = "" Then
 GoTo row_proc_end
 End If
 warning = is_unique_name2(clt, clt_phone) '确定提示信息
 
 If "OK" <> warning Then
 sheet.Cells(row, first_column + 3) = "客户" + warning
 GoTo row_proc_end
 End If
 
 Dim cltMgr As String
 cltMgr = Trim(sheet.Cells(row, first_column + 1))
 If cltMgr = "" Then
 GoTo row_proc_end
 End If
 
 warning = is_valid_name(cltMgr)
 If "OK" <> warning Then
 sheet.Cells(row, first_column + 3) = "客户经理" + warning
 GoTo row_proc_end
 End If
 
 count = dao_module.count_cm(cltMgr)
 If 1 <> count Then
   If 0 = count Then
     sheet.Cells(row, first_column + 3) = "错误:此客户经理资格未确定"
   Else
     sheet.Cells(row, first_column + 3) = "错误:此客户经理有重名"
   End If
   GoTo row_proc_end
Else
    If cltMgr <> clt And 1 = dao_module.count_cm(clt) Then
     sheet.Cells(row, first_column + 3) = "错误:客户经理的客户经理只能是他本人"
     GoTo row_proc_end
    End If
 End If
 
 sheet.Cells(row, first_column + 3) = "OK"
 
row_proc_end:
Next

MsgBox ("客户/客户经理，检查完毕")

End Sub

'--------------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------------
Public Sub UpdateClientManager() '
'假定文件格式为
'用户名     客户经理   提示
'假定数据从第2行、第1列开始

Dim xlApp As Excel.Application
Dim xlBook As Excel.Workbook
Dim sheet As Excel.Worksheet

Set xlApp = New Excel.Application
Set xlBook = ThisWorkbook
'set xlbook = xlapp.Workbooks.Open("");
Set sheet = xlBook.ActiveSheet

Dim first_column  As Long
first_column = 1

Dim row As Long
Dim column As Long
Dim row_cnt As Long
row_cnt = LastRow(sheet)

rows_loop:
For row = 2 To row_cnt Step 1
row_proc_start:
 Dim warning As String
 Dim count As Long
 
 Dim clt As String
 Dim clt_phone As String
 clt = Trim(sheet.Cells(row, first_column))
 clt_phone = Trim(sheet.Cells(row, first_column + 2))
 If clt = "" Then
 GoTo row_proc_end
 End If
 warning = is_unique_name2(clt, clt_phone)
 
 If "OK" <> warning Then
 sheet.Cells(row, first_column + 3) = "错误:客户" + warning
 GoTo row_proc_end
 End If
 
 Dim cltMgr As String
 cltMgr = Trim(sheet.Cells(row, first_column + 1))
 If cltMgr = "" Then
 GoTo row_proc_end
 End If
 
 warning = is_valid_name(cltMgr)
 If "OK" <> warning Then
 sheet.Cells(row, first_column + 3) = "错误:客户经理" + warning
 GoTo row_proc_end
 End If
 
 count = dao_module.count_cm(cltMgr)
 If 1 <> count Then
   If 0 = count Then
     sheet.Cells(row, first_column + 3) = "错误:此客户经理资格未确定"
   Else
     sheet.Cells(row, first_column + 3) = "错误:此客户经理有重名"
   End If
   GoTo row_proc_end
Else
    If cltMgr <> clt And 1 = dao_module.count_cm(clt) Then
     sheet.Cells(row, first_column + 3) = "错误:客户经理的客户经理只能是他本人"
     GoTo row_proc_end
    End If
 End If
 
 Dim clt_sn As Long
 clt_sn = read_clt_sn(clt, clt_phone)
 Dim is_clt_cm_exist As Boolean
 is_clt_cm_exist = Is_exist_map2(clt_sn, cltMgr)
 
 If is_clt_cm_exist Then
   sheet.Cells(row, first_column + 3) = "未变"
 Else
   Call dao_module.insert_history2(clt_sn)
   Call dao_module.insert_cm2(clt_sn, cltMgr)
   sheet.Cells(row, first_column + 3) = "OK"
 End If
 
row_proc_end:
Next

MsgBox ("客户/客户经理，更新完毕")

End Sub

Private Function is_unique_name(name As String) As String
 Dim warning As String
 count = dao_module.read_clt_count(name)
 If 1 = count Then
  warning = "OK"
 Else
  If 2 = count Then
  warning = "存在重名"
 Else
  warning = "不存在此姓名"
  End If
 End If
 
 is_unique_name = warning
End Function

Public Function read_clt_sn(clt As String, phone As String) As Long
 
 Dim real_phone As String
 real_phone = Trim(phone)
 If real_phone <> "" Then
    read_clt_sn = dao_module.read_clt_sn2(clt, phone)
 Else
    read_clt_sn = dao_module.read_clt_sn(clt)
    End If
  
End Function

Private Function is_unique_name2(name As String, phone As String) As String
 Dim warning As String
 Dim real_phone As String
 real_phone = Trim(phone)  '根据用户信息和phone做出判断，显现不同warning
 If real_phone <> "" Then
    count = dao_module.read_clt_count2(name, phone)
 Else
    count = dao_module.read_clt_count(name)
 End If
 If 1 = count Then
  warning = "OK"
 Else
  If 2 = count Then
  warning = "存在重名"
 Else
  warning = "不存在此姓名"
  End If
 End If
 
 is_unique_name2 = warning
End Function

Private Function is_valid_name(name As String) As String
 Dim warning As String
 count = dao_module.read_clt_count(name)
 If count > 0 Then
  warning = "OK"
 Else
  warning = "不存在此姓名"
  End If
 
 is_valid_name = warning
End Function

Private Function is_valid_cm_name(name As String) As String
 Dim warning As String
 count = dao_module.read_clt_count(name)
 If count > 0 Then
  warning = "OK"
 Else
  warning = "不存在此姓名"
  End If
 End If
 
 is_unique_name = warning
End Function

Private Function is_unique_cm_name(name As String) As String
 Dim warning As String
 count = dao_module.read_clt_count(name)
 If 1 = count Then
  warning = "OK"
 Else
  If 2 = count Then
  warning = "存在重名"
 Else
  warning = "不存在此姓名"
  End If
 End If
 
 is_unique_name = warning
End Function

Private Function Is_exist_map(clt As String, cm As String) As Boolean
  Is_exist_map = dao_module.cm_map_count(clt, cm) = 1
End Function

Private Function Is_exist_map2(clt_sn As Long, cm As String) As Boolean
  Is_exist_map2 = (dao_module.cm_map_count2(clt_sn, cm) = 1)
End Function


