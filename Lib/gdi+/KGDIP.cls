VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "KGDIP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Sub ReSizePicture(ByVal src As String, ByVal dest As String, Optional ByVal width = 0, Optional ByVal height = 0)
    
    Dim bitmap, w As Long, h As Long, graphic
    
    bitmap = GetImage(src, w, h)
    
    If width = 0 And height = 0 Then
        width = w
        height = h
    End If
    
    If width > 0 And height = 0 Then
        If w <= width Then
            gdip.RemoveImage src
            Exit Sub
        End If
        height = CLng(width / w * h)
        GoTo Resize
    End If
    
    If width = 0 And height > 0 Then
        If h <= height Then
            gdip.RemoveImage src
            Exit Sub
        End If
        width = CLng(height / h * w)
        GoTo Resize
    End If
    
    
    
Resize:
    Dim cBitmap As Long
    graphic = GetGraphic(, width, height, cBitmap)
    GdipDrawImageRectRectI graphic, bitmap, 0, 0, width, height, 0, 0, w, h, UnitPixel
    gdip.RemoveImage src
    
    writable dest
    On Error Resume Next
    Kill dest
    SaveBitMap cBitmap, dest
    gdip.RemoveImage cBitmap
    gdip.RemoveGraphic graphic
    
End Sub

Sub ReSizePictureByMulti(ByVal src As String, ByVal dest As String, multi As Double)
    
    Dim bitmap, w As Long, h As Long, graphic, width As Long, height As Long
    
    bitmap = GetImage(src, w, h)
    
    width = multi * w
    height = multi * h
Resize:
    Dim cBitmap As Long
    graphic = GetGraphic(, width, height, cBitmap)
    
    GdipDrawImageRectRectI graphic, bitmap, 0, 0, width, height, 0, 0, w, h, UnitPixel
    gdip.RemoveImage src
    
    
    writable dest
    On Error Resume Next
    Kill dest
    SaveBitMap cBitmap, dest
    gdip.RemoveImage cBitmap
    gdip.RemoveGraphic graphic
    
End Sub

Sub TrimPicture(ByVal src As String, ByVal dest As String, Optional ByVal top As Long = 0, Optional ByVal Bottom As Long = 0, Optional ByVal left As Long = 0, Optional ByVal Right As Long = 0)
    If Abs(top) + Abs(Bottom) + Abs(left) + Abs(Right) = 0 Then Exit Sub
    
    
    Dim bitmap, width As Long, height As Long, graphic
    
    bitmap = GetImage(src, width, height)
    
    If top + Bottom >= height Or left + Right >= width Then
        gdip.RemoveImage src
        Exit Sub
    End If
    Dim w, h
    w = width - left - Right
    h = height - top - Bottom
    Dim cBitmap As Long
    graphic = GetGraphic(, w, h, cBitmap)
    GdipDrawImageRectRectI graphic, bitmap, 0, 0, w, h, left, top, w, h, UnitPixel
    gdip.RemoveImage src
    
    writable dest
    On Error Resume Next
    Kill dest
    SaveBitMap cBitmap, dest
    gdip.RemoveImage cBitmap
    gdip.RemoveGraphic graphic
End Sub


Function CopyImage(ByVal bitmap As Long) As Long
    Dim g As Long, w As Long, h As Long
    Dim cBitmap As Long
    GdipGetImageWidth bitmap, w
    GdipGetImageHeight bitmap, h
    g = GetGraphic(, w, h, cBitmap)
    GdipDrawImageRectI g, bitmap, 0, 0, w, h
    RemoveGraphic g
    CopyImage = cBitmap
    
End Function


Public Sub Init()

End Sub
Private Sub Class_Initialize()
    
    
    If Not dicsystem.Exists("gdip") Then
        Dim uInput As GdiplusStartupInput
        uInput.GdiplusVersion = 1
        Dim mToken As Long
        GdiplusStartup mToken, uInput
        dicsystem("gdip") = mToken
    End If
End Sub


Private Sub Class_Terminate()
        Dim k
    If dicsystem.Exists("graphics") Then

        For Each k In dicsystem("graphics")
            GdipDeleteGraphics dicsystem("graphics")(k)
            dicsystem("graphics").Remove k
        Next
        dicsystem.Remove ("graphics")
    End If

    If dicsystem.Exists("images") Then
        For Each k In dicsystem("images")
            GdipDisposeImage dicsystem("images")(k)
            dicsystem("images").Remove k
        Next
        dicsystem.Remove ("images")
    End If
    
    
    If dicsystem.Exists("pens") Then
        For Each k In dicsystem("pens")
            GdipDeletePen dicsystem("pens")(k)
            dicsystem("pens").Remove k
        Next
        dicsystem.Remove ("pens")
    End If
    
    If dicsystem.Exists("brushes") Then
        For Each k In dicsystem("brushes")
            GdipDeleteBrush dicsystem("brushes")(k)
            dicsystem("brushes").Remove k
        Next
        dicsystem.Remove ("brushes")
    End If
    
    If dicsystem.Exists("stringformats") Then
        For Each k In dicsystem("stringformats")
            GdipDeleteStringFormat dicsystem("stringformats")(k)
            dicsystem("stringformats").Remove k
        Next
        dicsystem.Remove ("stringformats")
    End If
    
    
    If dicsystem.Exists("fontfamilies") Then
        For Each k In dicsystem("fontfamilies")
            GdipDeleteFontFamily dicsystem("fontfamilies")(k)
            dicsystem("fontfamilies").Remove k
        Next
        dicsystem.Remove ("fontfamilies")
    End If
    
    
    
    
    If dicsystem.Exists("fonts") Then
        For Each k In dicsystem("fonts")
            GdipDeleteFont dicsystem("fonts")(k)
            dicsystem("fonts").Remove k
        Next
        dicsystem.Remove ("fonts")
    End If


    If dicsystem.Exists("gdip") Then
        GdiplusShutdown dicsystem("gdip")
        dicsystem.Remove "gdip"
        
        
        
    End If
End Sub



Public Sub RemoveImage(ByVal key)
    If Not dicsystem.Exists("images") Then Exit Sub
    If Not dicsystem("images").Exists(key) Then Exit Sub
    GdipDisposeImage dicsystem("images")(key)
    dicsystem("images").Remove key
End Sub

Public Sub RemoveGraphic(ByVal key)
    If Not dicsystem.Exists("graphics") Then Exit Sub
    If Not dicsystem("graphics").Exists(key) Then Exit Sub
    GdipDeleteGraphics dicsystem("graphics")(key)
    dicsystem("graphics").Remove key
End Sub

