Attribute VB_Name = "modLoadGraphics"

Option Explicit

'#########################################################################
'# Name:          ModLoadGraphics                                        #
'# Description:   Loads and unloads graphics into main memory            #
'#                from resource files, returning necessary data:         #
'#                Handle, DC, width and Height                           #
'#########################################################################


'API declerationd for ModLoadGraphics
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long

Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long

Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

Private Type BITMAP
    bmType         As Long     'Bitmap type
    bmWidth        As Long     'Width/Pixels
    bmHeight       As Long     'Height/Pixels
    bmWidthBytes   As Long     'Width/Bytes
    bmPlanes       As Integer  'Planes
    bmBitsPixel    As Integer  'Bits per Pixel
    bmBits         As Long     'Bits
End Type

Public Type GRAPHIC
    hGraphic       As Long     'Handle
    hdc            As Long     'Device context
    hWidth         As Long     'Width/Pixels
    hHeight        As Long     'Height/Pixels
End Type

Private Const LR_LOADFROMFILE  As Integer = &H10
Private Const IMAGE_BITMAP     As Integer = 0


Public Sub ConvertToBitmap(FileName As String, Extension As String)

  Dim Pic As IPictureDisp

    'Create new picture object
    Set Pic = New StdPicture
    'Load picture
    Set Pic = LoadPicture(FileName & "." & Extension)
    'Delete old one
    Kill FileName & "." & Extension
    'Save bitmap to new file
    SavePicture Pic, FileName & ".bmp"

End Sub


Public Sub ExtractResource(FileName As String, ResourceName As Variant)

  Dim Buffer() As Byte
    
    'Load data into buffer
    Buffer = LoadResData(ResourceName, "CUSTOM")
    
    'Put buffer into file
    Open FileName For Binary As #1
        Put #1, , Buffer
    Close #1
    
    'Erase buffer
    Erase Buffer
    
End Sub


Public Sub DeleteGraphic(ByRef udtGraphic As GRAPHIC)

    DeleteDC udtGraphic.hdc                 'Delete the Device Context
    DeleteObject udtGraphic.hGraphic        'Delete the object

End Sub


Public Function LoadBitmap(ByVal Path As String) As GRAPHIC

  Dim hDesktop As Long
  Dim bmBitmap As BITMAP

    'Load the bitmap, this will return a handle if successful
    LoadBitmap.hGraphic = LoadImage(App.hInstance, Path, IMAGE_BITMAP, 0, 0, LR_LOADFROMFILE)

    If LoadBitmap.hGraphic = 0 Then     'If no handle, then the loadimage function must have failed _
       Most probable cause is because it dousn't exist or the given path is wrong
        Err.Raise _
                  Number:=vbObjectError + 513, _
                  Source:="ModLoadGraphics.LoadBitmap(''" & Path & "'')", _
                  Description:=("The graphic cannot be loaded. Please ensure that ''" & Path & "'' exists, as this is the most probable cause.")
      Else                                                           'We got a handle, the bitmap's loaded. So...
        hDesktop = GetDesktopWindow                                 'Get desktop handle
        LoadBitmap.hdc = CreateCompatibleDC(GetDC(hDesktop))        'Create a new DC compatible with the desktop
        SelectObject LoadBitmap.hdc, LoadBitmap.hGraphic            'Select the graphic, and the new DC into an object

        GetObject LoadBitmap.hGraphic, Len(bmBitmap), bmBitmap      'Get the width and height of the graphic
        LoadBitmap.hWidth = bmBitmap.bmWidth
        LoadBitmap.hHeight = bmBitmap.bmHeight
    End If

End Function
