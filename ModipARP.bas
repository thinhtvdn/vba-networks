Attribute VB_Name = "ModipARP"
Option Explicit


' A- cot du lieu tho -> B - cot danh sach ip -> C-cot unique ip -> D-cot thong ke

Sub ipARP()
Sheets(nameipArp).Activate
Call chuanBiManHinh
Call PreData
Call getIpFromStringARP ' samples - ' Internet  192.168.90.1            -   502f.a853.b369  ARPA   Vlan90
Call createArrayUniqueIp
Call statisticsForIp
Call xemKetQua
End Sub

Sub statisticsForIp()
Dim t As String
Dim tmpFormular  As String
Dim addData As String
Dim tmpVlan As String
Range(addDataBegin).Offset(0, 2).Select
t = ActiveCell.Value
If Len(t) = 0 Then GoTo kt
addData = "$B$1:" & getaddLastCellColumn("B") ' vung du lieu can thong ke, dia chi tuyet doi
While t <> ""
    tmpFormular = "=COUNTIF(" & addData & "," & ActiveCell.Address & ")"      '=COUNTIF($C$6:$C$354,D6)"
    ActiveCell.Offset(0, 1).Value = tmpFormular
    ActiveCell.Offset(1, 0).Select: t = ActiveCell.Value
Wend
kt:
Range(addDataBegin).Select
End Sub

Function getaddLastCellColumnARP(tencot As String) ' lay dia chi o cuoi cung cua cot co du lieu
Dim t As Integer
t = Range(tencot & Rows.Count).End(xlUp).Row
getaddLastCellColumnARP = tencot & CStr(t)
End Function


Sub createArrayUniqueIp()
Dim arr1() As String  ' khong khai bao kich thuoc, bat dau tu 0
Dim sttCurrEle As Integer ' so thu tu cua phan tu hien tai
sttCurrEle = 0
Dim t As String
Dim tmpip  As String
Dim tmpVlan As String

Range(addDataBegin).Offset(0, 1).Select ' chon cot ip
t = ActiveCell.Value
If Len(t) = 0 Then GoTo kt
While t <> ""
    If (Not IsInArray(t, arr1)) Then ' neu chua co trong mang thi them vao
        ReDim Preserve arr1(sttCurrEle)
        arr1(sttCurrEle) = t
        sttCurrEle = sttCurrEle + 1 ' tang vi tri so thu tu ky tu trong mang
    End If
    ActiveCell.Offset(1, 0).Select: t = ActiveCell.Value
Wend
Dim item As Variant
Range(addDataBegin).Offset(0, 2).Select ' chon cot ip
For Each item In arr1
     ActiveCell.Value = item
     ActiveCell.Offset(1, 0).Select
Next item
kt:
Range(addDataBegin).Select
End Sub

'---------------------------------------------------------------------------
Sub chuanBiManHinh()
    Columns("B:G").Select
    Selection.Delete Shift:=xlToLeft
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Du lieu tho"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "danh sach ip"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "danh sach unique ip"
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "Thong ke arp"
    Range("D2").Select
End Sub

Sub getIpFromStringARP()
Dim t As String
Dim tmpip  As String
Dim tmpVlan As String
Range(addDataBegin).Select
t = ActiveCell.Value
If Len(t) = 0 Then GoTo kt
While t <> ""
    tmpip = getIpArp(t)
    ActiveCell.Offset(0, 1).Value = tmpip
    ActiveCell.Offset(1, 0).Select: t = ActiveCell.Value
Wend
kt:
Range(addDataBegin).Select
End Sub

Private Function getIpArp(s As String) ' tach lay chuoi ip tu day ky tu
Dim tmpip  As String ' Internet  192.168.90.1            -   502f.a853.b369  ARPA   Vlan90                                                                                      ting  Vlan20
Dim ibd As Byte
s = Trim(s)
ibd = InStr(1, s, " ") ' tim ky tu trang dau tien
tmpip = Mid(s, ibd, 18)
getIpArp = Trim(tmpip) ' 192.168.20.70
End Function

Sub xemKetQua()
Columns("B:D").Select
Columns("B:D").EntireColumn.AutoFit
Range(addDataBegin).Select
End Sub

Sub PreData()
Dim t As String
Dim tmpStr As String
Const sample1 = "Protocol"
Dim LResult As Integer

Range(addDataBegin).Select
t = ActiveCell.Value
If Len(t) = 0 Then GoTo kt
While t <> ""
    tmpStr = Left(t, 8)
    LResult = StrComp(tmpStr, sample1)
    If (LResult = 0) Then ' dung chuoi du lieu can xoa
         ActiveCell.EntireRow.Delete
      Else
        ActiveCell.Offset(1, 0).Select
    End If
   t = ActiveCell.Value
Wend
kt:
Range(addDataBegin).Select

End Sub
