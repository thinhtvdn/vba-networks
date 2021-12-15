VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Dim cControl As CommandBarButton

Private Sub Workbook_AddinInstall()
On Error Resume Next 'Just in case

Application.CommandBars("Worksheet Menu Bar").Controls("VNT25-ipBinding").Delete
Application.CommandBars("Worksheet Menu Bar").Controls("VNT25-ipARP").Delete
Set cControl = Application.CommandBars("Worksheet Menu Bar").Controls.Add
With cControl
    .Caption = "VNT25-ipBinding"
    .Style = msoButtonIconAndCaptionBelow
    .OnAction = "ipBinding"
    .FaceId = 220
End With
Set cControl = Application.CommandBars("Worksheet Menu Bar").Controls.Add
With cControl
    .Caption = "VNT25-ipARP"
    .Style = msoButtonIconAndCaptionBelow
    .OnAction = "ipARP"
    .FaceId = 230
End With

On Error GoTo 0
End Sub


Private Sub Workbook_AddinUninstall()
On Error Resume Next 'In case it has already gone.
    Application.CommandBars("Worksheet Menu Bar").Controls("VNT25-ipBinding").Delete
    Application.CommandBars("Worksheet Menu Bar").Controls("VNT25-ipARP").Delete
   On Error GoTo 0
End Sub


