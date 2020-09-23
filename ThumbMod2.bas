Attribute VB_Name = "ThumbMod2"
Public Enum FrameStyle
FrameRaiseRaise
FrameRaiseInset
FrameInsetRaise
FrameInsetInset
End Enum

Public Enum FrFontStyle
FrNormal
FrBold
FrUnderline
FrItalic
FrBoldUnderline
FrBoldItalic
FrUnderlineItalic
FrBoldUnderlineItalic
End Enum

Public Function Frame3D(Frm As Form, Ct1 As Object, Ct2 As Object, Optional FrameTxt$, Optional FontStyle As FrFontStyle, Optional FrameTxtCol&, Optional Style As FrameStyle)
' ******** Code by Stephan Swertvaegher *********
'frm: the form to put the Frame3D
'Ct1: the first control (up-leftmost)
'Ct2: the last control(low-rightmost)
'FrameTxt: The caption of the Frame3D (optional)
'FontStyle: The style of the textcaption (bold,underline,...) (optional)
'FrameTxtCol: The color of the textcaption (optional)
'Style: The style of the Frame, from 0 to 3 (optional)


Dim R1%, G1%, B1%, R2%, G2%, B2%, R3%, G3%, B3%, R4%, G4%, B4%
Dim FrText$
Dim TempScaleMode%, TempForeColor
Dim FrX1%, FrX2%, FrY1%, FrY2%
Dim FrFontTransparent As Boolean
Dim FrFontBold As Boolean
Dim FrFontUnderline As Boolean
Dim FrFontItalic As Boolean

On Error GoTo Frame3D_Error

TempScaleMode = Frm.ScaleMode
TempForeColor = Frm.ForeColor
FrFontTransparent = Frm.FontTransparent
FrFontBold = Frm.FontBold
FrFontUnderline = Frm.FontUnderline
FrFontItalic = Frm.FontItalic
Frm.ScaleMode = 3 'pixel
Frm.AutoRedraw = True
Frm.FontTransparent = False
FrText = FrameTxt

If IsMissing(Style) Then Style = 0
If Style > 3 Then Style = 0
If IsMissing(FrameTxt) Then FrText = ""
FrText = Trim(FrText)
If IsMissing(FontStyle) Then FontStyle = FrNormal
If IsMissing(FrameTxtCol) Then FrameTxtCol = 0

'Framecolors depending on style
If Style = 0 Then 'FrameRaiseRaise
R1 = 240: R2 = 128: R3 = 240: R4 = 128
End If
If Style = 1 Then 'FrameRaiseInset
R1 = 240: R2 = 128: R4 = 240: R3 = 128
End If
If Style = 2 Then 'FrameInsetRaise
R2 = 240: R1 = 128: R3 = 240: R4 = 128
End If
If Style = 3 Then 'FrameInsetInset
R2 = 240: R1 = 128: R4 = 240: R3 = 128
End If
G1 = R1: B1 = R1
G2 = R2: B2 = R2
G3 = R3: B3 = R3
G4 = R4: B4 = R4
'look for the best corners of the frame
If Ct1.Left < Ct2.Left Then
FrX1 = Ct1.Left
Else
FrX1 = Ct2.Left
End If
If Ct1.Left + Ct1.Width > Ct2.Left + Ct2.Width Then
FrX2 = Ct1.Left + Ct1.Width
Else
FrX2 = Ct2.Left + Ct2.Width
End If
If Ct1.Top < Ct2.Top Then
FrY1 = Ct1.Top
Else
FrY1 = Ct2.Top
End If
If Ct1.Top + Ct1.Height > Ct2.Top + Ct2.Height Then
FrY2 = Ct1.Top + Ct2.Height
Else
FrY2 = Ct2.Top + Ct2.Height
End If
'draw outer border
    Frm.Line (FrX1 - 14, FrY1 - 14)-(FrX1 - 14, FrY2 + 14), RGB(R1, G1, B1)
    Frm.Line (FrX1 - 14, FrY1 - 14)-(FrX2 + 14, FrY1 - 14), RGB(R1, G1, B1)
    Frm.Line (FrX2 + 14, FrY1 - 14)-(FrX2 + 14, FrY2 + 14), RGB(R2, G2, B2)
    Frm.Line (FrX1 - 14, FrY2 + 14)-(FrX2 + 15, FrY2 + 14), RGB(R2, G2, B2)
'draw inner border
    Frm.Line (FrX1 - 10, FrY1 - 10)-(FrX1 - 10, FrY2 + 10), RGB(R3, G3, B3)
    Frm.Line (FrX1 - 10, FrY1 - 10)-(FrX2 + 10, FrY1 - 10), RGB(R3, G3, B3)
    Frm.Line (FrX2 + 10, FrY1 - 10)-(FrX2 + 10, FrY2 + 10), RGB(R4, G4, B4)
    Frm.Line (FrX1 - 10, FrY2 + 10)-(FrX2 + 11, FrY2 + 10), RGB(R4, G4, B4)
'place Framecaption in FontStyle
If FrText <> "" Then
FrText = " " & FrText & " "
Frm.ForeColor = FrameTxtCol
Frm.CurrentX = FrX1 - 5
Frm.CurrentY = FrY1 - 12 - Frm.FontSize
Select Case FontStyle
Case FrNormal
    Frm.FontBold = False
    Frm.FontUnderline = False
    Frm.FontItalic = False
Case FrBold
    Frm.FontBold = True
    Frm.FontUnderline = False
    Frm.FontItalic = False
Case FrUnderline
    Frm.FontBold = False
    Frm.FontUnderline = True
    Frm.FontItalic = False
Case FrItalic
    Frm.FontBold = False
    Frm.FontUnderline = False
    Frm.FontItalic = True
Case FrBoldUnderline
    Frm.FontBold = True
    Frm.FontUnderline = True
    Frm.FontItalic = False
Case FrBoldItalic
    Frm.FontBold = True
    Frm.FontUnderline = False
    Frm.FontItalic = True
Case FrUnderlineItalic
    Frm.FontBold = False
    Frm.FontUnderline = True
    Frm.FontItalic = True
Case FrBoldUnderlineItalic
    Frm.FontBold = True
    Frm.FontUnderline = True
    Frm.FontItalic = True
End Select
Frm.Print FrText
End If
'restore original values
Frm.ScaleMode = TempScaleMode
Frm.ForeColor = TempForeColor
Frm.FontTransparent = FrFontTransparent
Frm.FontBold = FrFontBold
Frm.FontUnderline = FrFontUnderline
Frm.FontItalic = FrFontItalic
Exit Function
'error
Frame3D_Error:
MsgBox "Error " & Err.Number & vbCr & Err.Description & vbCr & "Error while processing Frame3D - function", vbOKOnly, "Frame3D"
End Function


