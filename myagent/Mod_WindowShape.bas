Attribute VB_Name = "Mod_WindowShape"
'General Api Declarations
Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long) As Long

Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2

'This the Main Code to make an Image Shaped Form
'What it does is scan the Image passed to it and then
'remove all lines that correspond to the Transparent
'Color, creating a new virtual image, but without a
'particular color

Public Function GetBitmapRegion(cPicture As StdPicture, cTransparent As Long)
'Variable Declaration
    Dim hRgn As Long, tRgn As Long
    Dim x As Integer, Y As Integer, X0 As Integer
    Dim hdc As Long, BM As BITMAP
'Create a new memory DC, where we will scan the picture
    hdc = CreateCompatibleDC(0)
    If hdc Then
'Let the new DC select the Picture
        SelectObject hdc, cPicture
'Get the Picture dimensions and create a new rectangular
'region
        GetObject cPicture, Len(BM), BM
        hRgn = CreateRectRgn(0, 0, BM.bmWidth, BM.bmHeight)
'Start scanning the picture from top to bottom
        For Y = 0 To BM.bmHeight
            For x = 0 To BM.bmWidth
'Scan a line of non transparent pixels
                While x <= BM.bmWidth And GetPixel(hdc, x, Y) <> cTransparent
                    x = x + 1
                Wend
'Mark the start of a line of transparent pixels
                X0 = x
'Scan a line of transparent pixels
                While x <= BM.bmWidth And GetPixel(hdc, x, Y) = cTransparent
                    x = x + 1
                Wend
'Create a new Region that corresponds to the row of
'Transparent pixels and then remove it from the main
'Region
                If X0 < x Then
                    tRgn = CreateRectRgn(X0, Y, x, Y + 1)
                    CombineRgn hRgn, hRgn, tRgn, 4
'Free the memory used by the new temporary Region
                    DeleteObject tRgn
                End If
            Next x
        Next Y
'Return the memory address to the shaped region
        GetBitmapRegion = hRgn
'Free memory by deleting the Picture
        DeleteObject SelectObject(hdc, cPicture)
    End If
'Free memory by deleting the created DC
    DeleteDC hdc
End Function


