VERSION 5.00
Begin VB.Form fMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Sample Project"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6825
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   6825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2895
      Left            =   120
      ScaleHeight     =   191
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   439
      TabIndex        =   0
      Top             =   120
      Width           =   6615
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Dim bmBitmap As modLoadGraphics.GRAPHIC

Private Sub Form_Load()
    
    'Using my module, a graphic can be loaded into main memory _
    from a resource file and drawn in just 6 easy steps!
    
    'Step 1, extract the resource to the hard-drive
    modLoadGraphics.ExtractResource "C:\Temp.jpg", "jpg_Graphic"
    'Step 2, Convert the resource to a bitmap
    modLoadGraphics.ConvertToBitmap "C:\Temp", "jpg"
    'Step 3, load the graphic into main memory
    bmBitmap = modLoadGraphics.LoadBitmap("C:\Temp.Bmp")
    'Step 4, draw the graphic
    BitBlt Pic.hdc, 0, 0, bmBitmap.hWidth, bmBitmap.hHeight, bmBitmap.hdc, 0, 0, vbSrcCopy
    Pic.Refresh
    'Step 5, Finished? Well delete the graphic
    modLoadGraphics.DeleteGraphic bmBitmap
    'Step 6, Don't want all your graphics to remain on the hard-drive? _
    Then delete them when you are finished or once they have been loaded!
    Kill "C:\Temp.Bmp"
    
End Sub
