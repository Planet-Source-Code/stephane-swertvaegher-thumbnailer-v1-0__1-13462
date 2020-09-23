VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form ThumbFrm 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   8385
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   9585
   Icon            =   "ThumbFrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   559
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   639
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   1035
      Top             =   5805
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   14
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ThumbFrm.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ThumbFrm.frx":0C56
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   45
      Picture         =   "ThumbFrm.frx":0FE2
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   634
      TabIndex        =   10
      Top             =   45
      Width           =   9510
      Begin VB.Image Image2 
         Height          =   210
         Left            =   9135
         Top             =   45
         Width           =   240
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   315
      Top             =   5715
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ThumbFrm.frx":245D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ThumbFrm.frx":28B1
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ThumbFrm.frx":2D05
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ThumbFrm.frx":3159
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ThumbFrm.frx":32B5
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ThumbFrm.frx":3411
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   330
      Left            =   540
      TabIndex        =   9
      Top             =   450
      Width           =   2565
      _ExtentX        =   4524
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "keyPrev"
            Object.ToolTipText     =   "Previous 30 pictures"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "keyNext"
            Object.ToolTipText     =   "Next 30 pictures"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "keyView"
            Object.ToolTipText     =   "View selected thumb in normal format"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "keySave"
            Object.ToolTipText     =   "Save selected picture"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "keyAbout"
            Object.ToolTipText     =   "About..."
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "keyHelp"
            Object.ToolTipText     =   "Show me some help"
            ImageIndex      =   5
         EndProperty
      EndProperty
   End
   Begin VB.DirListBox Dir1 
      BackColor       =   &H00FFFF80&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   2115
      Left            =   405
      TabIndex        =   4
      Top             =   1575
      Width           =   2715
   End
   Begin VB.DriveListBox Drive1 
      BackColor       =   &H00FFFF80&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   315
      Left            =   405
      TabIndex        =   3
      Top             =   1215
      Width           =   2715
   End
   Begin VB.FileListBox File1 
      BackColor       =   &H00FFFF80&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   2040
      Left            =   405
      TabIndex        =   2
      Top             =   3690
      Width           =   2715
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   7950
      Left            =   3735
      ScaleHeight     =   530
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   388
      TabIndex        =   0
      Top             =   225
      Width           =   5820
      Begin VB.Label TLabel 
         Alignment       =   2  'Center
         Caption         =   "Label1"
         Height          =   195
         Index           =   0
         Left            =   45
         TabIndex        =   1
         Top             =   1125
         Width           =   1050
      End
      Begin VB.Image Thumb 
         Height          =   1050
         Index           =   0
         Left            =   45
         Stretch         =   -1  'True
         Top             =   45
         Width           =   1050
      End
   End
   Begin VB.PictureBox Picture3 
      Height          =   3210
      Left            =   4140
      ScaleHeight     =   3150
      ScaleWidth      =   2250
      TabIndex        =   11
      Top             =   1845
      Visible         =   0   'False
      Width           =   2310
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   0
      Top             =   540
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "ThumbNailer V1.0"
      Flags           =   2
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00606000&
      Height          =   195
      Left            =   405
      TabIndex        =   8
      Top             =   6075
      Width           =   2670
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Picture info"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   270
      TabIndex        =   7
      Top             =   6705
      Width           =   2985
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Label2"
      ForeColor       =   &H00C00000&
      Height          =   1185
      Left            =   270
      TabIndex        =   6
      Top             =   7155
      Width           =   2985
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00606000&
      Height          =   195
      Left            =   405
      TabIndex        =   5
      Top             =   5805
      Width           =   2670
   End
   Begin VB.Image Image1 
      Height          =   3000
      Left            =   4050
      Top             =   2520
      Visible         =   0   'False
      Width           =   3000
   End
End
Attribute VB_Name = "ThumbFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NewStart%, T%, Error%, ff%

Private Sub ShowNextThumbs()
Error = 0
On Error GoTo err_handler
Screen.MousePointer = 11
ReDim Info(File1.ListCount - 1, 4)
ClearThumbs
T = 0
NewStart = Start
For xx = Start To Start + 29
Image1.Picture = LoadPicture(File1.Path & "\" & File1.List(xx))
Info(T, 0) = File1.List(xx)
Info(T, 1) = FileLen(File1.Path & "\" & File1.List(xx))
Info(T, 2) = Str(Image1.Width)
Info(T, 3) = Str(Image1.Height)
Info(T, 4) = FileDateTime(File1.Path & "\" & File1.List(xx))
Dimensions Image1.Width, Image1.Height
Thumb(T).Move Pos(T, 0) + ((70 - NewL) / 2), Pos(T, 1) + ((70 - NewH) / 2), NewL, NewH
Thumb(T).Picture = Image1.Picture
TLabel(T).Caption = File1.List(xx)
T = T + 1
Next xx
Start = Start + 30
Label4.Caption = "Showing items: " & NewStart + 1 & " to " & xx
Toolbar1.Buttons(2).Enabled = True
Screen.MousePointer = 1
If xx = File1.ListCount Then Error = 1
Exit Sub
err_handler:
If Err <> 76 Then
MsgBox "File: " & File1.List(xx) & vbCr & Err.Description, , "ThumbNailer V1.0"
Resume Next
Else
Error = 1
Exit Sub
End If

End Sub

Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.Picture = ImageList2.ListImages(2).Picture
End Sub

Private Sub Image2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.Picture = ImageList2.ListImages(1).Picture
End
End Sub

Private Sub Thumb_DblClick(Index As Integer)
Idx = Index
If TLabel(Idx).Caption = "" Then Exit Sub
ApiMenu05_Click
End Sub

Private Sub Thumb_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Idx = Index
Picture3.Picture = LoadPicture(Dir1.Path & "\" & TLabel(Idx).Caption)
For xx = 0 To File1.ListCount - 1
If TLabel(Index).Caption = File1.List(xx) Then
File1.Selected(xx) = True
Exit For
End If
Next xx
If Button <> 2 Then Exit Sub
If TLabel(Idx).Caption = "" Then Exit Sub
CreateApiMenu2
End Sub

Private Sub TLabel_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Idx = Index
Picture3.Picture = LoadPicture(Dir1.Path & "\" & TLabel(Idx).Caption)
File1.Selected(Index) = True
If Button <> 2 Then Exit Sub
If TLabel(Idx).Caption = "" Then Exit Sub
CreateApiMenu2
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
Case "keyNext"
Toolbar1.Buttons(2).Enabled = False
If Start = File1.ListCount Then Exit Sub
Toolbar1.Buttons(1).Enabled = False
Label4.Caption = ""
DoEvents
If File1.ListCount = 0 Then Exit Sub
ShowNextThumbs
If Start > 30 Then Toolbar1.Buttons(1).Enabled = True

    If Start >= File1.ListCount Then
    Toolbar1.Buttons(2).Enabled = True
    End If
If Error = 1 Then GoTo EndThumbs
'----------
Case "keyPrev"
Start = Start - 60
Toolbar1.Buttons(1).Enabled = False
Toolbar1.Buttons(2).Enabled = False
Label4.Caption = ""
DoEvents
ShowNextThumbs
If Start > 30 Then Toolbar1.Buttons(1).Enabled = True
'-----------
Case "keyView"
For xx = 0 To File1.ListCount - 1
If File1.Selected(xx) = True Then GoTo keyview2
Next xx
MsgBox "No picture selected !", vbExclamation, "ThumbNailer V1.0"
Exit Sub
keyview2:
ApiMenu05_Click
'---------
Case "keyAbout"
ThumbFrm4.Show 1
'---------
Case "keyHelp"
ThumbFrm5.Show 1
'---------
Case "keySave"
For xx = 0 To File1.ListCount - 1
If File1.Selected(xx) = True Then GoTo keysave2
Next xx
MsgBox "No picture selected !", vbExclamation, "ThumbNailer V1.0"
Exit Sub
keysave2:
SavePic

End Select
Exit Sub
EndThumbs:
Label4.Caption = "Showing items: " & NewStart + 1 & " to " & xx
If Start >= 30 Then Toolbar1.Buttons(1).Enabled = True
Toolbar1.Buttons(2).Enabled = False
Start = Start + 30
Screen.MousePointer = 1
End Sub

Private Sub SavePic()
On Error GoTo SavePic_error
CD1.FileName = File1.List(xx)
CD1.ShowSave
SavePicture Picture3.Picture, CD1.FileName
SavePic_error:
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
CD1.InitDir = Dir1.Path
Label4.Caption = ""
On Error Resume Next
ClearThumbs
Toolbar1.Buttons(1).Enabled = False
If File1.ListCount <> 0 Then
Toolbar1.Buttons(2).Enabled = True
Else
Toolbar1.Buttons(2).Enabled = False
End If
Start = 0
Label1.Caption = File1.ListCount & " Items"
End Sub

Private Sub Drive1_Change()
On Error GoTo Dr_error
Dir1.Path = Left(Drive1.Drive, 2) & "\"
Exit Sub
Dr_error:
MsgBox "Device not ready", vbCritical, "ThumbNailer"
End Sub

Private Sub Form_Load()
Dim Nx%, Ny%, T%
Image2.Picture = ImageList2.ListImages(1).Picture
Start = 0
ThumbFrm.Move 1150, 0, 9570, 8525
Thumbfrm2.Move 0, 0, 12000, 9000
ThumbFrm.Line (0, 0)-(ThumbFrm.ScaleWidth - 1, ThumbFrm.ScaleHeight - 1), 0, B
ThumbFrm3.Move 0, 0
ThumbFrm3.Line (0, 0)-(ThumbFrm3.ScaleWidth - 1, ThumbFrm3.ScaleHeight - 1), &H800000, B
ThumbFrm3.Line (3, 3)-(ThumbFrm3.ScaleWidth - 4, ThumbFrm3.ScaleHeight - 4), &H800000, B
ThumbFrm5.Line (3, 3)-(ThumbFrm5.ScaleWidth - 4, ThumbFrm5.ScaleHeight - 4), &H800000, B
ThumbFrm5.Move (Screen.Width / 2) - (ThumbFrm5.Width / 2), (Screen.Height / 2) - (ThumbFrm5.Height / 2)
File1.Pattern = "*.bmp;*.jpg;*.gif;*.wmf;*.ico"
Dir1.Path = Left(Drive1.Drive, 2) & "\"
Label1.Caption = File1.ListCount & " Items"
GetPositions
Picture1.Move 244, 30, (77 * 5) - 3, (88 * 6) - 2
Picture2.Move 2, 2
Picture2.Line (0, 0)-(Picture2.ScaleWidth - 1, Picture2.ScaleHeight - 1), 0, B
For xx = 1 To 29
Load Thumb(xx)
Thumb(xx).Visible = True
Load TLabel(xx)
TLabel(xx).Visible = True
Next xx
SetThumbs
ClearThumbs
T3D ThumbFrm, Label2, 5, T3dRaiseInset
T3D ThumbFrm, Toolbar1, 5, T3dRaiseRaise
T3D ThumbFrm, Label3, 5, T3dRaiseRaise
T3D ThumbFrm, Picture1, 5, T3dRaiseInset
T3D ThumbFrm3, ThumbFrm3.Label1, 5, T3dInsetRaise
T3D ThumbFrm5, ThumbFrm5.Label3, 5, T3dInsetRaise
ThumbFrm.FontSize = 10
Frame3D ThumbFrm, Drive1, Label4, " Directory ", FrBold, &HC00000, FrameRaiseInset
Label2.Caption = ""
Label4.Caption = ""
Toolbar1.Buttons(1).Enabled = False
Toolbar1.Buttons(2).Enabled = False
With ThumbFrm4
.Label1.Caption = "Coded in December 2000" & vbCr & vbCr
.Label1.Caption = .Label1.Caption & "Contact me at:" & vbCr
.Label1.Caption = .Label1.Caption & "gumming@compaqnet.be"
.Label2.Caption = .Label1.Caption
End With
ff = FreeFile
    On Error GoTo Load2
    Open App.Path & "\Help.txt" For Input As #ff
    ThumbFrm5.Text1.Text = Input(LOF(ff), 1)
Load2:
    Close #ff
ThumbFrm.Show
ThumbFrm4.Label3.ForeColor = &H108080
ThumbFrm4.Show 1
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.Caption = ""
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
End
End Sub


Private Sub Thumb_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If TLabel(Index).Caption = "" Then
Label2.Caption = ""
Exit Sub
End If
Label2.Caption = "Item: " & Index + 1 & vbCr
Label2.Caption = Label2.Caption & "Filename: " & Info(Index, 0) & vbCr
Label2.Caption = Label2.Caption & "Filelength: " & Info(Index, 1) & " bytes" & vbCr
Label2.Caption = Label2.Caption & "Picture Width: " & Info(Index, 2) & vbCr
Label2.Caption = Label2.Caption & "Picture Height: " & Info(Index, 3) & vbCr
Label2.Caption = Label2.Caption & "Last Modified:: " & Info(Index, 4) & vbCr
End Sub

Private Sub TLabel_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If TLabel(Index).Caption = "" Then
Label2.Caption = ""
Exit Sub
End If
Label2.Caption = "Item: " & Index + 1 & vbCr
Label2.Caption = Label2.Caption & "Filename: " & Info(Index, 0) & vbCr
Label2.Caption = Label2.Caption & "Filelength: " & Info(Index, 1) & " bytes " & vbCr
Label2.Caption = Label2.Caption & "Picture Width: " & Info(Index, 2) & vbCr
Label2.Caption = Label2.Caption & "Picture Height: " & Info(Index, 3) & vbCr
Label2.Caption = Label2.Caption & "Last Modified:: " & Info(Index, 4) & vbCr
End Sub

