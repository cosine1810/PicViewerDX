VERSION 5.00
Begin VB.Form frmD3D 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   8970
   ClientLeft      =   7380
   ClientTop       =   2235
   ClientWidth     =   12255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   598
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   817
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   3120
      Top             =   3120
   End
End
Attribute VB_Name = "frmD3D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public myDx As KDirectX8
Public D3DDevice As Direct3DDevice8
'Private TriList() As TLVERTEX '//Note that, like in TriStrip, we're generating 2 triangles - yet using 2 more vertices...
Private TriFan() As TLVERTEX


Const FVF = D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR


Private gdi As New KGDIP


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyUp
    
    Case vbKeyDown
    
    End Select
End Sub

Private Sub Form_Load()
    gdi.Init
    Set myDx = New KDirectX8
    Set D3DDevice = myDx.Initialise(hwnd)

    Dim matWorld As D3DMATRIX
    D3DXMatrixRotationY matWorld, 0
    D3DDevice.SetTransform D3DTS_WORLD, matWorld
    
'    SetPerspective
'    SetCamera
    SetViewMatrix
    Dim sq As K3D
    Set sq = myDx.Add3D
    
    Dim path As String, w As Long, h As Long, bitmap As Long
    path = "C:\shinekara\Resources\Desktop\2014_013.jpg"
    bitmap = GetImage(path, w, h)
    sq.Square path, w, h
    
End Sub

Private Sub SetViewMatrix()
    Dim ViewMatrix As D3DMATRIX
    Dim EyePosition As D3DVECTOR
    Dim LookAtPosition As D3DVECTOR
    Dim UpVector As D3DVECTOR
    

    
    ' 修改为轴侧角的镜头位置和观察方向
    EyePosition.X = 5 ' X坐标
    EyePosition.Y = 5 ' Y坐标
    EyePosition.z = 5 ' Z坐标
    LookAtPosition.X = 0
    LookAtPosition.Y = 0
    LookAtPosition.z = 0
    UpVector.X = 0
    UpVector.Y = 1
    UpVector.z = 0
    
    D3DXMatrixLookAtLH ViewMatrix, EyePosition, LookAtPosition, UpVector
    D3DDevice.SetTransform D3DTS_VIEW, ViewMatrix
End Sub

'Private Sub SetCamera()
'    ' 定义镜头的位置
'    Dim eye As D3DVECTOR
'    eye.X = 0
'    eye.Y = 0
'    eye.z = -5 ' 假设镜头在 z 轴负方向上距离原点 5 个单位
'
'    ' 定义镜头的目标点
'    Dim at As D3DVECTOR
'    at.X = 0
'    at.Y = 0
'    at.z = 0 ' 镜头指向原点
'
'    ' 定义向上的方向向量
'    Dim up As D3DVECTOR
'    up.X = 0
'    up.Y = 1
'    up.z = 0 ' y 轴正方向为向上的方向
'
'    ' 创建观察矩阵
'    Dim viewMatrix As D3DMATRIX
'    D3DXMatrixLookAtLH viewMatrix, eye, at, up
'
'    ' 设置观察矩阵
'    D3DDevice.SetTransform D3DTS_VIEW, viewMatrix
'End Sub

'Private Sub SetPerspective()
'    Dim projectionMatrix As D3DMATRIX
'    Dim fovy As Single
'    fovy = 1.0472 ' 约 60 度视角，转换为弧度
'    Dim aspect As Single
'    aspect = Me.ScaleWidth / Me.ScaleHeight
'    Dim zn As Single
'    zn = 1#
'    Dim zf As Single
'    zf = 100#
'    D3DXMatrixPerspectiveFovLH projectionMatrix, fovy, aspect, zn, zf
'    D3DDevice.SetTransform D3DTS_PROJECTION, projectionMatrix
'End Sub
'
'Private Sub SetCamera()
'    Const c_VerticalAngle As Single = -PI / 4
'    Const c_HorizontalAngle As Single = 0
'    Dim CameraPos As D3DVECTOR
'    Dim MyD3DMATRIX As D3DMATRIX
'
'    D3DXMatrixRotationX MyD3DMATRIX, c_VerticalAngle  'MyD3DMATRIX 乘以了一个绕 X 轴旋转的变换矩阵
'    D3DDevice.SetTransform D3DTS_VIEW, MyD3DMATRIX  '初始化视图矩阵
'
'    D3DXMatrixRotationY MyD3DMATRIX, c_HorizontalAngle 'MyD3DMATRIX 乘以了一个绕 Y 轴旋转的变换矩阵
'    D3DDevice.MultiplyTransform D3DTS_VIEW, MyD3DMATRIX '视图矩阵乘上变换矩阵MyD3DMATRIX
'
'    With CameraPos
'     .X = 0
'     .Y = 0
'     .z = -10
'    End With
'
'    D3DXMatrixTranslation MyD3DMATRIX, -CameraPos.X, -CameraPos.Y, -CameraPos.z '使矩阵 MyD3DMATRIX 平移 -CameraPos.x, -CameraPos.y, -CameraPos.z
'    D3DDevice.MultiplyTransform D3DTS_VIEW, MyD3DMATRIX '视图矩阵乘上变换矩阵MyD3DMATRIX
'End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        MousePos = GetCursorByObjectPos(Me)
'        LastTouchTime = GetTickCount
    End If
End Sub

Private Sub Timer1_Timer()
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, &H0, 1#, 0
    D3DDevice.BeginScene
    D3DDevice.SetVertexShader FVF
    
    myDx.Render
    
    D3DDevice.EndScene
    D3DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0
End Sub

Public Sub Render()

        
    '画背景
    D3DDevice.SetTexture 0, myDx.GetTexture("Temp\back.jpg")
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, TriStrip(0), Len(TriStrip(0))

    'Alpha透明混合
    
    
    D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
    
    D3DDevice.SetTexture 0, myDx.GetTexture("mask.png")
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, TriStrip(0), Len(TriStrip(0))
    
    D3DDevice.SetTexture 0, myDx.GetTexture(PicName)
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, TriStrip(0), Len(TriStrip(0))
    
    
    With TriFan(0)
    
        offset.X = offset.X * 0.97
        offset.Y = offset.Y * 0.97
        
        .X = offset.X * WaveWidth / 2 * Math.Cos(((GetTickCount - LastTouchTime) Mod 1700) / 1700 * 2 * 3.1415926) + CenterX
        .Y = offset.Y * WaveHeight / 2 * Math.Cos(((GetTickCount - LastTouchTime) Mod 1000) / 1000 * 2 * 3.1415926) + CenterY
    
    End With

    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLEFAN, UBound(TriFan) + 1 - 2, TriFan(0), Len(TriFan(0))
    D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 0
        
End Sub
