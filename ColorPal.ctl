VERSION 5.00
Begin VB.UserControl ColorPal 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4215
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3180
   FillStyle       =   0  'Solid
   MouseIcon       =   "ColorPal.ctx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   281
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   212
End
Attribute VB_Name = "ColorPal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim ColorList() As Long
Dim MaxCol As Integer
Dim TSize As Integer

Public Event Click()
Public Event DblClick()
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Public Event ColorSelected(cColor As Long)
Public Event ColorOver(cColor As Long)
Public Sub LoadPalette(Optional PalFile As String)
On Error Resume Next ' GoTo ErrLoad
Dim FF As Integer
Dim tStr As String
Dim n As Integer
Dim cQty As Integer
Dim Row As Integer
Dim Col As Integer

FF = FreeFile

If PalFile = "" Or Dir(PalFile) = "" Then PalFile = App.Path & "\Default.pal"

If Dir(PalFile) <> "" Then
Open PalFile For Input As #FF
Input #FF, tStr$ 'JASC-PAL
    If UCase(tStr) <> "JASC-PAL" Then
    Close #FF
    Exit Sub
    End If
Input #FF, tStr$ '0010
Input #FF, tStr$ '256 (color qty)
cQty = Int(tStr)
ReDim ColorList(Int(cQty))
n = 0
    While Not EOF(FF)
    Input #FF, tStr$
    ColorList(n) = RGB(Split(tStr, " ")(0), Split(tStr, " ")(1), Split(tStr, " ")(2))
    n = n + 1
    Wend
Close #FF
Col = 0
Row = 0
    For n = 0 To cQty - 1
    UserControl.Line (Col * TSize, Row * TSize)-(Col * TSize + TSize, Row * TSize + TSize), ColorList(n), BF
    Col = Col + 1
    If Col = MaxCol Then
    Col = 0
    Row = Row + 1
    End If
    Next n
UserControl.Width = UserControl.ScaleX((MaxCol * TSize) + 5, vbPixels, vbContainerSize)
UserControl.Height = UserControl.ScaleY((cQty / MaxCol * TSize) + TSize + 2, vbPixels, vbContainerSize)
End If
Exit Sub
ErrLoad:
Close #FF
End Sub

Public Property Get ColumnQty() As Integer
ColumnQty = MaxCol
End Property

Public Property Let ColumnQty(ByVal iColumnQty As Integer)
MaxCol = iColumnQty
LoadPalette
PropertyChanged "ColumnQty"
End Property

Public Property Get ThumbSize() As Integer
ThumbSize = TSize
End Property

Public Property Let ThumbSize(ByVal iThumbSize As Integer)
TSize = iThumbSize
LoadPalette
PropertyChanged "ThumbSize"
End Property

Private Sub UserControl_Click()
RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
RaiseEvent DblClick
End Sub


Private Sub UserControl_InitProperties()
TSize = 10
MaxCol = 12
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Dim tRow As Integer
Dim tCol As Integer
Dim tColor As Long
Dim tInd As Integer

If Button = 1 Then
tCol = X \ TSize
tRow = Y \ TSize
tInd = tRow * MaxCol + tCol
tColor = ColorList(tInd)
RaiseEvent ColorSelected(tColor)
End If

RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Dim tRow As Integer
Dim tCol As Integer
Dim tColor As Long
Dim tInd As Integer

tCol = X \ TSize
tRow = Y \ TSize
tInd = tRow * MaxCol + tCol
tColor = ColorList(tInd)
RaiseEvent ColorOver(tColor)
RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub


Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub


Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
With PropBag
    MaxCol = .ReadProperty("ColumnQty", 12)
    TSize = .ReadProperty("Thumbsize", 10)
End With
End Sub


Private Sub UserControl_Resize()
LoadPalette
End Sub



Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
     .WriteProperty "ColumnQty", MaxCol, 12
     .WriteProperty "Thumbsize", TSize, 10
End With
End Sub


