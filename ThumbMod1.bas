Attribute VB_Name = "ThumbMod1"
Public NewL%, NewH%, xx%
Public Pos%(29, 1), Info$()
Public Start%, Idx%
Public LL%, HH%

Public Enum T3dFill
T3dF0
T3dF1
End Enum

Public Enum Borderstyle
T3dRaiseRaise
T3dRaiseInset
T3dInsetRaise
T3dInsetInset
T3dNone
End Enum
'API for translating system colors to 'normal' colors
Private Declare Function TranslateColor Lib "olepro32.dll" Alias "OleTranslateColor" (ByVal clr As OLE_COLOR, ByVal palet As Long, Col As Long) As Long
  
  Private Const PLANES& = 14
  Private Const BITSPIXEL& = 12
  Private Declare Function GetDeviceCaps& Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long)
  Private Declare Function GetDC& Lib "user32" (ByVal hWnd As Long)
  Private Declare Function ReleaseDC& Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long)
   
  Private Function ColorDepth() As Integer
     Dim nPlanes As Integer, BitsPerPixel As Integer, dc As Long
     dc = GetDC(0)
     nPlanes = GetDeviceCaps(dc, PLANES)
     BitsPerPixel = GetDeviceCaps(dc, BITSPIXEL)
     ReleaseDC 0, dc
     ColorDepth = nPlanes * BitsPerPixel
  End Function

 
Public Function T3D(Obj0 As Object, Obj As Object, Bev%, Optional Style3D As Borderstyle, Optional T3dFilled As T3dFill)
Dim R%, G%, B%, R1%, G1%, B1%, R2%, G2%, B2%, R3%, G3%, B3%, R4%, G4%, B4%
Dim FC&, T3Dxx%, SM%
On Error Resume Next

'global things
SM = Obj0.ScaleMode 'save scalemode
Obj0.ScaleMode = 3 'pixel
Obj.Borderstyle = 0 'no border
If IsMissing(Style3D) Then Style3D = 0
If Style3D > 4 Then Style3D = 3

'get formcolor
FC = Obj0.BackColor
'in case formcolor = systemcolor --> call the function RealColor
FC = RealColor(FC)
' convert to RGB
R = FC And &HFF
G = Int((FC And &HFF00&) / 256)
B = Int((FC And &HFF0000) / 65536)
'-------------------
If Style3D = 0 Then 'RaiseRaise
    R1 = R + 64
    If R1 > 255 Then R1 = 255
    R2 = R - 64
    If R2 < 0 Then R2 = 0
    R3 = R1
    R4 = R2
    G1 = G + 64
    If G1 > 255 Then G1 = 255
    G2 = G - 64
    If G2 < 0 Then G2 = 0
    G3 = G1
    G4 = G2
    B1 = B + 64
    If B1 > 255 Then B1 = 255
    B2 = B - 64
    If B2 < 0 Then B2 = 0
    B3 = B1
    B4 = B2
End If
'-------------------
If Style3D = 1 Then 'RaiseInset
    R1 = R + 64
    If R1 > 255 Then R1 = 255
    R2 = R - 64
    If R2 < 0 Then R2 = 0
    R4 = R1
    R3 = R2
    G1 = G + 64
    If G1 > 255 Then G1 = 255
    G2 = G - 64
    If G2 < 0 Then G2 = 0
    G4 = G1
    G3 = G2
    B1 = B + 64
    If B1 > 255 Then B1 = 255
    B2 = B - 64
    If B2 < 0 Then B2 = 0
    B4 = B1
    B3 = B2
End If
If Style3D = 2 Then 'InsetRaise
    R2 = R + 64
    If R2 > 255 Then R2 = 255
    R1 = R - 64
    If R1 < 0 Then R1 = 0
    R4 = R1
    R3 = R2
    G2 = G + 64
    If G2 > 255 Then G2 = 255
    G1 = G - 64
    If G1 < 0 Then G1 = 0
    G4 = G1
    G3 = G2
    B2 = B + 64
    If B2 > 255 Then B2 = 255
    B1 = B - 64
    If B1 < 0 Then B1 = 0
    B4 = B1
    B3 = B2
End If
If Style3D = 3 Then 'InsetInset
    R2 = R + 64
    If R2 > 255 Then R2 = 255
    R1 = R - 64
    If R1 < 0 Then R1 = 0
    R3 = R1
    R4 = R2
    G2 = G + 64
    If G2 > 255 Then G2 = 255
    G1 = G - 64
    If G1 < 0 Then G1 = 0
    G3 = G1
    G4 = G2
    B2 = B + 64
    If B2 > 255 Then B2 = 255
    B1 = B - 64
    If B1 < 0 Then B1 = 0
    B3 = B1
    B4 = B2
End If
If Style3D = 4 Then 'No Border
R1 = R: R2 = R: R3 = R: R4 = R
G1 = G: G2 = G: G3 = G: G4 = G
B1 = B: B2 = B: B3 = B: B4 = B
End If
Bev = Bev + 1
T3Dxx = Bev 'just in case Filled = 1

'Outer
If IsMissing(T3dFilled) Or T3dFilled = 0 Then
    Obj0.Line (Obj.Left - Bev, Obj.Top - Bev)-(Obj.Left - Bev, Obj.Top + Obj.Height + Bev), RGB(R1, G1, B1)
    Obj0.Line (Obj.Left - Bev, Obj.Top - Bev)-(Obj.Left + Obj.Width + Bev, Obj.Top - Bev), RGB(R1, G1, B1)
    Obj0.Line (Obj.Left + Obj.Width + Bev, Obj.Top - Bev)-(Obj.Left + Obj.Width + Bev, Obj.Top + Obj.Height + Bev), RGB(R2, G2, B2)
    Obj0.Line (Obj.Left - Bev, Obj.Top + Obj.Height + Bev)-(Obj.Left + Obj.Width + Bev + 1, Obj.Top + Obj.Height + Bev), RGB(R2, G2, B2)
Else
For Bev = T3Dxx To 1 Step -1 'in case T3DF1 (filled)
    Obj0.Line (Obj.Left - Bev, Obj.Top - Bev)-(Obj.Left - Bev, Obj.Top + Obj.Height + Bev), RGB(R1, G1, B1)
    Obj0.Line (Obj.Left - Bev, Obj.Top - Bev)-(Obj.Left + Obj.Width + Bev, Obj.Top - Bev), RGB(R1, G1, B1)
    Obj0.Line (Obj.Left + Obj.Width + Bev, Obj.Top - Bev)-(Obj.Left + Obj.Width + Bev, Obj.Top + Obj.Height + Bev), RGB(R2, G2, B2)
    Obj0.Line (Obj.Left - Bev, Obj.Top + Obj.Height + Bev)-(Obj.Left + Obj.Width + Bev + 1, Obj.Top + Obj.Height + Bev), RGB(R2, G2, B2)
Next Bev
End If
'Inner
    Obj0.Line (Obj.Left - 1, Obj.Top - 1)-(Obj.Left - 1, Obj.Top + Obj.Height + 1), RGB(R3, G3, B3)
    Obj0.Line (Obj.Left - 1, Obj.Top - 1)-(Obj.Left + Obj.Width + 1, Obj.Top - 1), RGB(R3, G3, B3)
    Obj0.Line (Obj.Left + Obj.Width + 1, Obj.Top - 1)-(Obj.Left + Obj.Width + 1, Obj.Top + Obj.Height + 1), RGB(R4, G4, B4)
    Obj0.Line (Obj.Left - 1, Obj.Top + Obj.Height + 1)-(Obj.Left + Obj.Width + 2, Obj.Top + Obj.Height + 1), RGB(R4, G4, B4)

Obj0.ScaleMode = SM 'restore original scalemode
End Function
  
  ' if System Color then translate to 'normal color'
  ' else, do nothing
  Public Function RealColor(ByVal Color As OLE_COLOR) As Long
     Dim Col As Long
     Col = TranslateColor(Color, 0, RealColor)
  End Function

Public Sub Dimensions(L%, B%)
Dim T%
NewL = L: NewH = B
T = 1
Do While NewL > 70 Or NewH > 70
NewL = Int(L / T)
NewH = Int(B / T)
T = T + 1
Loop
End Sub


Public Sub SetThumbs()
With ThumbFrm
    For Nx = 0 To 29
        .Thumb(Nx).Move Pos(Nx, 0), Pos(Nx, 1), 70, 70
        .TLabel(Nx).Move Pos(Nx, 0), Pos(Nx, 1) + 70, 70, 13
        .TLabel(Nx).Caption = "g" & Format(Nx, "000") & ".jpg"
    .Picture1.Line (Pos(Nx, 0) - 2, Pos(Nx, 1) - 2)-(Pos(Nx, 0) + 71, Pos(Nx, 1) + 83), RGB(240, 240, 240), B
    .Picture1.Line (Pos(Nx, 0) - 2, Pos(Nx, 1) + 83)-(Pos(Nx, 0) + 72, Pos(Nx, 1) + 83), RGB(128, 128, 128)
    .Picture1.Line (Pos(Nx, 0) + 71, Pos(Nx, 1) - 2)-(Pos(Nx, 0) + 71, Pos(Nx, 1) + 83), RGB(128, 128, 128)
    Next Nx
End With
End Sub

Public Sub GetPositions()
Pos(0, 0) = 2: Pos(0, 1) = 2
Pos(1, 0) = 79: Pos(1, 1) = 2
Pos(2, 0) = 156: Pos(2, 1) = 2
Pos(3, 0) = 233: Pos(3, 1) = 2
Pos(4, 0) = 310: Pos(4, 1) = 2
Pos(5, 0) = 2: Pos(5, 1) = 90
Pos(6, 0) = 79: Pos(6, 1) = 90
Pos(7, 0) = 156: Pos(7, 1) = 90
Pos(8, 0) = 233: Pos(8, 1) = 90
Pos(9, 0) = 310: Pos(9, 1) = 90
Pos(10, 0) = 2: Pos(10, 1) = 178
Pos(11, 0) = 79: Pos(11, 1) = 178
Pos(12, 0) = 156: Pos(12, 1) = 178
Pos(13, 0) = 233: Pos(13, 1) = 178
Pos(14, 0) = 310: Pos(14, 1) = 178
Pos(15, 0) = 2: Pos(15, 1) = 266
Pos(16, 0) = 79: Pos(16, 1) = 266
Pos(17, 0) = 156: Pos(17, 1) = 266
Pos(18, 0) = 233: Pos(18, 1) = 266
Pos(19, 0) = 310: Pos(19, 1) = 266
Pos(20, 0) = 2: Pos(20, 1) = 354
Pos(21, 0) = 79: Pos(21, 1) = 354
Pos(22, 0) = 156: Pos(22, 1) = 354
Pos(23, 0) = 233: Pos(23, 1) = 354
Pos(24, 0) = 310: Pos(24, 1) = 354
Pos(25, 0) = 2: Pos(25, 1) = 442
Pos(26, 0) = 79: Pos(26, 1) = 442
Pos(27, 0) = 156: Pos(27, 1) = 442
Pos(28, 0) = 233: Pos(28, 1) = 442
Pos(29, 0) = 310: Pos(29, 1) = 442
End Sub
 
 Public Sub ClearThumbs()
 With ThumbFrm
 For xx = 0 To 29
 .Thumb(xx).Picture = LoadPicture("")
 .TLabel(xx).Caption = ""
 Next xx
 End With
 End Sub

Public Function GetNewDim(Pct%)
LL = Val(Info(Idx, 2)) * Pct / 100
HH = Val(Info(Idx, 3)) * Pct / 100
Thumbfrm2.Image1.ToolTipText = "Scale: " & Pct & "%"
If Pct = 100 Then
Thumbfrm2.Image1.ToolTipText = "Normal size"
End If
End Function


Public Sub ShowSized()
With Thumbfrm2
.Image1.Picture = LoadPicture("")
.Image1.Width = LL
.Image1.Height = HH
.Image1.Left = (Thumbfrm2.ScaleWidth - LL) / 2
.Image1.Top = (Thumbfrm2.ScaleHeight - HH) / 2
DoEvents
.Image1.Picture = ThumbFrm.Thumb(Idx).Picture
End With
End Sub

