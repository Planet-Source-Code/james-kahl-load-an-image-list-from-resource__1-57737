Attribute VB_Name = "mMain"
Option Explicit

Public Enum ResImages
    bmpNew = 101
    bmpOpen = 102
    bmpSave = 103
    bmpVB6 = 104
    icoMapDrive = 101
    icoDiscDrive = 102
    curArrow = 101
    #If False Then
        Private bmpNew
        Private bmpOpen
        Private bmpSave
        Private icoMapDrive
        Private icoDiscDrive
        Private curArrow
    #End If
End Enum

Public Sub AddResImageToImageList(ByVal Iml As ImageList, ByVal ResImage As ResImages, ByVal ResType As LoadResConstants, Optional Key As Variant)
    Dim lIndex As Long
    Dim li As ListImage
    
    'get the index for the new image
    lIndex = Iml.ListImages.Count + 1
    'add the image to the image list
    Set li = Iml.ListImages.Add(lIndex, Key, LoadResPicture(ResImage, ResType))
    
    Set li = Nothing
End Sub

