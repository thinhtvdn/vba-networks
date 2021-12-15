Attribute VB_Name = "ModipBinding"
Option Explicit

Sub ipBinding()
' important : A6 - begin data to processing
' 192.168.30.42   01fc.aa14.6b54.3e       Nov 09 2020 11:40 PM    Automatic  Activ                                                                                        e     Vlan30
Sheets(nameipBinding).Activate
Call chuanBiManHinh
Call getIpVlanFromString
Call createArrayUniqueVlan
Call statisticsForVlan
Call addInfoVlan
Columns("B:J").EntireColumn.AutoFit ' chinh hien thi
End Sub

Sub addInfoVlan()
Dim t As String
Dim tmpNameVlan As String
Sheets(nameipBinding).Activate
Range(addDataBegin).Select
ActiveCell.Offset(0, 3).Select ' chon cot unique vlan
t = ActiveCell.Value
While t <> ""
    tmpNameVlan = getNameUnifiVlan(nameSheetUnifiAp, t)
    Sheets(nameipBinding).Activate
    ActiveCell.Offset(0, 2).Value = tmpNameVlan
    ActiveCell.Offset(1, 0).Select: t = ActiveCell.Value
Wend
End Sub

Function getNameUnifiVlan(nameSheet As String, tmpVlanFind As String)
Dim LResult As Integer
Dim t, tmpStr As String
tmpStr = ""
Sheets(nameSheet).Activate
Range(addDataInfoUnifiBegin).Select: t = ActiveCell.Value
If Len(t) = 0 Then GoTo kt
While t <> ""
    LResult = StrComp(tmpVlanFind, t)
    If (LResult = 0) Then ' dung chuoi du lieu can tim
        tmpStr = ActiveCell.Offset(0, -1).Value
        GoTo kt
    End If
    ActiveCell.Offset(1, 0).Select: t = ActiveCell.Value
Wend
kt:
   getNameUnifiVlan = tmpStr
End Function

Sub chuanBiManHinh()
    Columns("B:G").Select
    Selection.Delete Shift:=xlToLeft
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Du lieu tho"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Tat ca ip"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "Tat ca Vlan"
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "Unique Vlan"
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "Thong ke"
    Range("F1").Select
    ActiveCell.FormulaR1C1 = "Thong tin"
End Sub

Sub statisticsForVlan()
Dim t As String
Dim tmpFormular  As String
Dim addData As String
Dim tmpVlan As String
Range(addDataBegin).Select
ActiveCell.Offset(0, 3).Select
t = ActiveCell.Value
If Len(t) = 0 Then GoTo kt
addData = "$C$2:" & getaddLastCellColumn("C") ' vung du lieu can thong ke, dia chi tuyet doi
While t <> ""
    tmpFormular = "=COUNTIF(" & addData & "," & ActiveCell.Address & ")"      '=COUNTIF($C$6:$C$354,D6)"
    ActiveCell.Offset(0, 1).Value = tmpFormular
    ActiveCell.Offset(1, 0).Select: t = ActiveCell.Value
Wend
kt:
Range(addDataBegin).Select
End Sub


Function getaddLastCellColumn(tencot As String) ' lay dia chi o cuoi cung cua cot co du lieu
Dim t As Integer
t = Range(tencot & Rows.Count).End(xlUp).Row
getaddLastCellColumn = "$" & tencot & "$" & CStr(t)
End Function

Sub getIpVlanFromString()
Dim t As String
Dim tmpip  As String
Dim tmpVlan As String
Range(addDataBegin).Select
t = ActiveCell.Value
If Len(t) = 0 Then GoTo kt
While t <> ""
    tmpip = getIp(t)
    tmpVlan = getVLAN(tmpip)
    ActiveCell.Offset(0, 1).Value = tmpip
    ActiveCell.Offset(0, 2).Value = tmpVlan
    ActiveCell.Offset(1, 0).Select: t = ActiveCell.Value
Wend
kt:
Range(addDataBegin).Select
End Sub

Private Function getIp(s As String) ' tach lay chuoi ip tu day ky tu
Dim tmpip  As String ' 192.168.20.70   01f8.5971.1dda.30       Nov 09 2020 01:21 AM    Automatic  Selec                                                                                        ting  Vlan20
Dim i, vt As Byte
tmpip = Left(s, 15)
getIp = Trim(tmpip) ' 192.168.20.70
End Function

Private Function getVLAN(ip As String) ' tach lay chuoi vlan tu day ip
Dim tmp1, tmp2, tmp3, vlan As String
Dim dot1, dot2, dot3 As Byte
'  192.168.20.70
dot1 = InStr(1, ip, ".")  ' lay dau cham thu 1
tmp1 = Mid(ip, dot1 + 1, Len(ip) - dot1 + 1) ' 168.20.70
dot2 = InStr(1, tmp1, ".") 'lay dau cham thu 2
tmp2 = Mid(tmp1, dot2 + 1, Len(tmp1) - dot2 + 1) ' 20.70
dot3 = InStr(1, tmp2, ".") 'lay dau cham thu 3
vlan = Left(tmp2, dot3 - 1) ' 20
getVLAN = vlan
End Function

Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean ' kiem tra 1 chuoi co trong mang 1 chieu
On Error Resume Next ' neu bi loi thi thuc hien lenh tiep
Dim kq As Boolean
kq = False ' gia su chua co
kq = (UBound(Filter(arr, stringToBeFound)) > -1) ' neu mang empty, thi bi loi -> kq mac dinh la False
IsInArray = kq
End Function

Sub createArrayUniqueVlan()
Dim arr1() As String  ' khong khai bao kich thuoc, bat dau tu 0
Dim sttCurrEle As Integer ' so thu tu cua phan tu hien tai
sttCurrEle = 0
Dim t As String
Dim tmpip  As String
Dim tmpVlan As String
Range(addDataBegin).Select
ActiveCell.Offset(0, 2).Select
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
Range(addDataBegin).Select
ActiveCell.Offset(0, 3).Select
Dim item As Variant
For Each item In arr1
     ActiveCell.Value = item
     ActiveCell.Offset(1, 0).Select
Next item
kt:
Range(addDataBegin).Select
End Sub


