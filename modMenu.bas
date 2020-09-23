Attribute VB_Name = "modMenu"
Public Declare Function GetMenu Lib "user32" _
(ByVal hWnd As Long) As Long

Public Declare Function GetSubMenu Lib "user32" _
(ByVal hMenu As Long, ByVal nPos As Long) As Long

Public Declare Function GetMenuItemID Lib "user32" _
(ByVal hMenu As Long, ByVal nPos As Long) As Long

Public Declare Function SetMenuItemBitmaps Lib "user32" _
(ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As _
Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked _
As Long) As Long



