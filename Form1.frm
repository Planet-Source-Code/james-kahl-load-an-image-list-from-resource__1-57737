VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5490
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   7485
   LinkTopic       =   "Form1"
   ScaleHeight     =   5490
   ScaleWidth      =   7485
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picImage 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3795
      Left            =   1575
      ScaleHeight     =   3795
      ScaleWidth      =   3480
      TabIndex        =   1
      Top             =   945
      Width           =   3480
   End
   Begin MSComctlLib.Toolbar tbrMain 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7485
      _ExtentX        =   13203
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Map Drive"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Disconnect Drive"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlImages 
      Left            =   1995
      Top             =   1050
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.Menu mnuTop 
      Caption         =   "Button Size"
      NegotiatePosition=   2  'Middle
      WindowList      =   -1  'True
      Begin VB.Menu mnuSize 
         Caption         =   "Small"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu mnuSize 
         Caption         =   "Medium"
         Index           =   1
      End
      Begin VB.Menu mnuSize 
         Caption         =   "Large"
         Index           =   2
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    loadImages
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'release the images in imagelist
    Set tbrMain.ImageList = Nothing
    imlImages.ListImages.Clear
End Sub

Private Sub loadImages()
    AddResImageToImageList imlImages, bmpNew, vbResBitmap, "New"
    AddResImageToImageList imlImages, bmpOpen, vbResBitmap, "Open"
    AddResImageToImageList imlImages, bmpSave, vbResBitmap, "Save"
    AddResImageToImageList imlImages, icoMapDrive, vbResIcon, "Map Network Drive"
    AddResImageToImageList imlImages, icoDiscDrive, vbResIcon, "Disconnect Network Drive"
    AddResImageToImageList imlImages, curArrow, vbResCursor
    AddResImageToImageList imlImages, bmpVB6, vbResBitmap
    
    Set tbrMain.ImageList = imlImages
    tbrMain.Buttons(1).Image = 1
    tbrMain.Buttons(2).Image = 2
    tbrMain.Buttons(3).Image = 3
    tbrMain.Buttons(5).Image = 4
    tbrMain.Buttons(6).Image = 5
    Set picImage.Picture = imlImages.ListImages(7).Picture
    Me.MousePointer = vbCustom
    Me.MouseIcon = imlImages.ListImages(6).Picture
End Sub

Private Sub mnuSize_Click(Index As Integer)
    Dim lSize As Long
    Dim lIdx As Long
    
    'release the image list so we can change the size
    Set tbrMain.ImageList = Nothing
    imlImages.ListImages.Clear

    'set the size of the images
    Select Case Index
        Case 0
            lSize = 16
        Case 1
            lSize = 32
        Case 2
            lSize = 48
    End Select
    
    'make sure the appropriate item is checked
    For lIdx = 0 To 2
        If lIdx = Index Then
            mnuSize(lIdx).Checked = True
        Else
            mnuSize(lIdx).Checked = False
        End If
    Next lIdx
    
    imlImages.ImageHeight = lSize
    imlImages.ImageWidth = lSize
    
    'reload the images
    loadImages
End Sub
