Attribute VB_Name = "PopupMod1"
'Declaration section
Public Declare Function CreatePopupMenu Lib "user32.dll" () As Long
Public Declare Function DestroyMenu Lib "user32.dll" (ByVal hMenu As Long) As Long
Public Type MENUITEMINFO
        cbSize As Long
        fMask As Long
        fType As Long
        fState As Long
        wID As Long
        hSubMenu As Long
        hbmpChecked As Long
        hbmpUnchecked As Long
        dwItemData As Long
        dwTypeData As String
        cch As Long
End Type
Public Const MIIM_STATE = &H1
Public Const MIIM_ID = &H2
Public Const MIIM_SUBMENU = &H4
Public Const MIIM_CHECKMARKS = &H8
Public Const MIIM_DATA = &H20
Public Const MIIM_TYPE = &H10
Public Const MFT_BITMAP = &H4
Public Const MFT_MENUBARBREAK = &H20
Public Const MFT_MENUBREAK = &H40
Public Const MFT_OWNERDRAW = &H100
Public Const MFT_RADIOCHECK = &H200
Public Const MFT_RIGHTJUSTIFY = &H4000
Public Const MFT_RIGHTORDER = &H2000
Public Const MFT_SEPARATOR = &H800
Public Const MFT_STRING = &H0
Public Const MFS_CHECKED = &H8
Public Const MFS_DEFAULT = &H1000
Public Const MFS_DISABLED = &H2
Public Const MFS_ENABLED = &H0
Public Const MFS_GRAYED = &H1
Public Const MFS_HILITE = &H80
Public Const MFS_UNCHECKED = &H0
Public Const MFS_UNHILITE = &H0
Public Declare Function InsertMenuItem Lib "user32.dll" Alias "InsertMenuItemA" _
(ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Long, lpmii As MENUITEMINFO) As Long
Public Declare Function TrackPopupMenu Lib "user32.dll" _
(ByVal hMenu As Long, ByVal uFlags As Long, ByVal X As Long, ByVal Y As Long, _
ByVal nReserved As Long, ByVal hWnd As Long, ByVal prcRect As Long) As Long
Public Const TPM_RIGHTALIGN = &H8&
Public Const TPM_CENTERALIGN = &H4&
Public Const TPM_LEFTALIGN = &H0
Public Const TPM_TOPALIGN = &H0
Public Const TPM_NONOTIFY = &H80
Public Const TPM_RETURNCMD = &H100
Public Const TPM_LEFTBUTTON = &H0
Public Const TPM_RIGHTBUTTON = &H2&
Public Type POINT_TYPE
X As Long
Y As Long
End Type
Public Declare Function GetCursorPos Lib "user32.dll" (lpPoint As POINT_TYPE) As Long
Public Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long

Public Sub CreateApiMenu()
Dim hPopupMenu1 As Long
Dim Menu1 As MENUITEMINFO
Dim curpos As POINT_TYPE
Dim menusel As Long
Dim retval As Long

hPopupMenu1 = CreatePopupMenu()

With Menu1
.cbSize = Len(Menu1)
.fMask = MIIM_STATE Or MIIM_ID Or MIIM_TYPE Or MIIM_SUBMENU
End With

With Menu1
.fType = MFT_STRING
.fState = MFS_ENABLED
.wID = 1011 ' Assign this item an item identifier.
.dwTypeData = "View normal size"
.cch = Len("View normal size")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu1, 0, 1, Menu1)

With Menu1
.fType = MFT_SEPARATOR
.fState = MFS_ENABLED
.wID = 1010 ' Assign this item an item identifier.
.dwTypeData = "/separator/"
.cch = Len("/separator/")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu1, 1, 1, Menu1)

With Menu1
.fType = MFT_STRING
.fState = MFS_ENABLED
.wID = 1009
.dwTypeData = "Reduce to 75%"
.cch = Len("Reduce to 75%")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu1, 2, 1, Menu1)

With Menu1
.fType = MFT_STRING
.fState = MFS_ENABLED
.wID = 1008
.dwTypeData = "Reduce to 50%"
.cch = Len("Reduce to 50%")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu1, 3, 1, Menu1)

With Menu1
.fType = MFT_STRING
.fState = MFS_ENABLED
.wID = 1007
.dwTypeData = "Reduce to 25%"
.cch = Len("Reduce to 25%")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu1, 4, 1, Menu1)

With Menu1
.fType = MFT_SEPARATOR
.fState = MFS_ENABLED
.wID = 1006
.dwTypeData = "/separator/"
.cch = Len("/separator/")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu1, 5, 1, Menu1)

With Menu1
.fType = MFT_STRING
.fState = MFS_ENABLED
.wID = 1005
.dwTypeData = "Enlarge to 150%"
.cch = Len("Enlarge to 150%")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu1, 6, 1, Menu1)

With Menu1
.fType = MFT_STRING
.fState = MFS_ENABLED
.wID = 1004 ' Assign this item an item identifier.
.dwTypeData = "Enlarge to 200%"
.cch = Len("Enlarge to 200%")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu1, 7, 1, Menu1)

With Menu1
.fType = MFT_STRING
.fState = MFS_ENABLED
.wID = 1003 ' Assign this item an item identifier.
.dwTypeData = "Enlarge to 250%"
.cch = Len("Enlarge to 250%")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu1, 8, 1, Menu1)

With Menu1
.fType = MFT_STRING
.fState = MFS_ENABLED
.wID = 1002 ' Assign this item an item identifier.
.dwTypeData = "Enlarge to 300%"
.cch = Len("Enlarge to 300%")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu1, 9, 1, Menu1)

With Menu1
.fType = MFT_SEPARATOR
.fState = MFS_ENABLED
.wID = 1001 ' Assign this item an item identifier.
.dwTypeData = "/separator/"
.cch = Len("/separator/")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu1, 10, 1, Menu1)

With Menu1
.fType = MFT_STRING
.fState = MFS_ENABLED
.wID = 1000 ' Assign this item an item identifier.
.dwTypeData = "Main"
.cch = Len("Main")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu1, 11, 1, Menu1)

With Menu1
.fType = MFT_SEPARATOR
.fState = MFS_ENABLED
.wID = 1012 ' Assign this item an item identifier.
.dwTypeData = "/separator/"
.cch = Len("/separator/")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu1, 12, 1, Menu1)

With Menu1
.fType = MFT_STRING
.fState = MFS_ENABLED
.wID = 1013 ' Assign this item an item identifier.
.dwTypeData = "Show Info"
.cch = Len("Show Info")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu1, 13, 1, Menu1)

retval = GetCursorPos(curpos)
menusel = TrackPopupMenu(hPopupMenu1, TPM_TOPALIGN Or TPM_NONOTIFY Or TPM_RETURNCMD Or TPM_RIGHTALIGN Or TPM_RIGHTBUTTON, curpos.X, curpos.Y, 0, Thumbfrm2.hWnd, 0)
retval = DestroyMenu(hPopupMenu3)
Select Case menusel
Case 1013
Call ApiMenu10_Click
Case 1011
Call ApiMenu1_Click
Case 1009
Call ApiMenu2_Click
Case 1008
Call ApiMenu3_Click
Case 1007
Call ApiMenu4_Click
Case 1005
Call ApiMenu5_Click
Case 1004
Call ApiMenu6_Click
Case 1003
Call ApiMenu7_Click
Case 1002
Call ApiMenu8_Click
Case 1000
Call ApiMenu9_Click
Case Else
End Select
End Sub
'execute items in popupmenu.

Public Sub ApiMenu1_Click()
GetNewDim 100
ShowSized
End Sub

Public Sub ApiMenu2_Click()
GetNewDim 75
ShowSized
End Sub

Public Sub ApiMenu3_Click()
GetNewDim 50
ShowSized
End Sub

Public Sub ApiMenu4_Click()
GetNewDim 25
ShowSized
End Sub

Public Sub ApiMenu5_Click()
GetNewDim 150
ShowSized
End Sub

Public Sub ApiMenu05_Click()
GetNewDim 100
ShowSized
Thumbfrm2.Show 1
End Sub

Public Sub ApiMenu6_Click()
GetNewDim 200
ShowSized
End Sub

Public Sub ApiMenu06_Click()
GetNewDim 75
ShowSized
Thumbfrm2.Show 1
End Sub

Public Sub ApiMenu7_Click()
GetNewDim 250
ShowSized
End Sub

Public Sub ApiMenu07_Click()
GetNewDim 50
ShowSized
Thumbfrm2.Show 1
End Sub

Public Sub ApiMenu8_Click()
GetNewDim 300
ShowSized
End Sub

Public Sub ApiMenu08_Click()
GetNewDim 25
ShowSized
Thumbfrm2.Show 1
End Sub

Public Sub ApiMenu9_Click()
Thumbfrm2.Hide
End Sub

Public Sub ApiMenu10_Click()
With ThumbFrm3
.Label2.Caption = "Item: " & Idx + 1 & vbCr & vbCr
.Label2.Caption = .Label2.Caption & "Filename: " & Info(Idx, 0) & vbCr
.Label2.Caption = .Label2.Caption & "Filelength: " & Info(Idx, 1) & " bytes" & vbCr
.Label2.Caption = .Label2.Caption & "Picture normal width: " & Info(Idx, 2) & vbCr
.Label2.Caption = .Label2.Caption & "Picture normal height: " & Info(Idx, 3) & vbCr & vbCr
.Label2.Caption = .Label2.Caption & "Current " & Thumbfrm2.Image1.ToolTipText & vbCr
.Label2.Caption = .Label2.Caption & "Picture current width: " & Thumbfrm2.Image1.Width & vbCr
.Label2.Caption = .Label2.Caption & "Picture current height: " & Thumbfrm2.Image1.Height & vbCr & vbCr
.Label2.Caption = .Label2.Caption & "Last Modified: " & Info(Idx, 4) & vbCr
ThumbFrm3.Show 1
End With
End Sub

Public Sub CreateApiMenu2()
Dim hPopupMenu1 As Long
Dim Menu4 As MENUITEMINFO
Dim curpos As POINT_TYPE
Dim menusel As Long
Dim retval As Long

hPopupMenu1 = CreatePopupMenu()

With Menu4
.cbSize = Len(Menu4)
.fMask = MIIM_STATE Or MIIM_ID Or MIIM_TYPE Or MIIM_SUBMENU ' Which elements of the structure to use.
End With

With Menu4
.fType = MFT_STRING
.fState = MFS_ENABLED
.wID = 1004
.dwTypeData = "View full picture"
.cch = Len("View full picture")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu1, 0, 1, Menu4)

With Menu4
.fType = MFT_SEPARATOR
.fState = MFS_ENABLED
.wID = 1003
.dwTypeData = "/separator/"
.cch = Len("/separator/")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu1, 1, 1, Menu4)

With Menu4
.fType = MFT_STRING
.fState = MFS_ENABLED
.wID = 1002
.dwTypeData = "View 75%"
.cch = Len("View 75%")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu1, 2, 1, Menu4)

With Menu4
.fType = MFT_STRING
.fState = MFS_ENABLED
.wID = 1001
.dwTypeData = "View 50%"
.cch = Len("View 50%")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu1, 3, 1, Menu4)

With Menu4
.fType = MFT_STRING
.fState = MFS_ENABLED
.wID = 1000
.dwTypeData = "View 25%"
.cch = Len("View 25%")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu1, 4, 1, Menu4)

retval = GetCursorPos(curpos)
menusel = TrackPopupMenu(hPopupMenu1, TPM_TOPALIGN Or TPM_NONOTIFY Or TPM_RETURNCMD Or TPM_RIGHTALIGN Or TPM_RIGHTBUTTON, curpos.X, curpos.Y, 0, ThumbFrm.hWnd, 0)
retval = DestroyMenu(hPopupMenu1)
Select Case menusel
Case 1004
Call ApiMenu05_Click
Case 1002
Call ApiMenu06_Click
Case 1001
Call ApiMenu07_Click
Case 1000
Call ApiMenu08_Click
Case Else
End Select
End Sub



