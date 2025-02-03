VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "KDirectX8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private dx As DirectX8
Private D3D As Direct3D8
Private D3DDevice As Direct3DDevice8


Private bRunning As Boolean

Private D3DX As D3DX8

Private Enum CustomSetting
    zbuffer = 1
    fullscreen = 2
End Enum
Private mySettings As CustomSetting

Private root As K3D
Private m3DMgr As New Dictionary

Function Add3D(Optional ByVal pt As K3D) As K3D
    If pt Is Nothing Then Set pt = root
    Dim n3D As New K3D
    Set n3D.parent = pt
    Set n3D.myDx = Me
    pt.Children.Add n3D
    m3DMgr.Add n3D.id, n3D
    Set Add3D = n3D
End Function

Sub Delete3D(ByVal m3D As K3D)
    m3DMgr.Remove m3D.id
    m3D.Terminate
End Sub

Sub Render(Optional ByVal v As K3D)
    If v Is Nothing Then Set v = root
    Dim txt As Direct3DTexture8

    Select Case v.category
    Case "Square"
        Set txt = GetTexture(v.texturePath)
        D3DDevice.SetTexture 0, txt
        D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, ByVal v.TriStripPtr, v.TriStripLen
    End Select
    
    Dim k
    For Each k In v.Children
        Render k
    Next
End Sub


Public Function Initialise(hwnd As Long) As Direct3DDevice8
    
    
    Dim DispMode As D3DDISPLAYMODE
    Dim D3DWindow As D3DPRESENT_PARAMETERS
    Set dx = New DirectX8
    Set D3D = dx.Direct3DCreate()
    
    Set D3DX = New D3DX8 '后面有东西用到这个
    
    D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispMode
    
    If (mySettings And fullscreen) <> fullscreen Then
        
        D3DWindow.Windowed = 1
        D3DWindow.SwapEffect = D3DSWAPEFFECT_COPY_VSYNC
        D3DWindow.BackBufferFormat = DispMode.format
    
    Else

        
        'DispMode.Format = D3DFMT_X8R8G8B8
        DispMode.format = D3DFMT_R5G6B5
'        DispMode.Width = 640
'        DispMode.Height = 480
        
        D3DWindow.SwapEffect = D3DSWAPEFFECT_FLIP
        D3DWindow.BackBufferCount = 1
        D3DWindow.BackBufferFormat = DispMode.format
        D3DWindow.BackBufferWidth = DispMode.width
        D3DWindow.BackBufferHeight = DispMode.height
        
        D3DWindow.hDeviceWindow = hwnd
    End If

    
    
    
    
    
    If (mySettings And zbuffer) = zbuffer Then
        D3DWindow.EnableAutoDepthStencil = 1
        D3DWindow.AutoDepthStencilFormat = D3DFMT_D16 '//16 bit Z-Buffer
    End If
    
    Dim device As Direct3DDevice8
    
    Set device = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, hwnd, _
                                                          D3DCREATE_SOFTWARE_VERTEXPROCESSING, _
                                                          D3DWindow)
    device.SetVertexShader FVF
    device.SetRenderState D3DRS_LIGHTING, False
    
    If mySettings And zbuffer = zbuffer Then
        device.SetRenderState D3DRS_ZENABLE, 1
    End If
    
    Set D3DDevice = device
    Set Initialise = device
    
    Set root = New K3D
    Set root.myDx = Me
    m3DMgr.Add ObjPtr(root), root
End Function




Public Function GetTexture(ByVal path) As Direct3DTexture8
    Dim Name As String
    Name = MarkName(path)
    
    If Not dicsystem.Exists("d3dtextures") Then Set dicsystem("d3dtextures") = New Dictionary
    If Not dicsystem("d3dtextures").Exists(Name) Then
        Set dicsystem("d3dtextures")(Name) = D3DX.CreateTextureFromFile(D3DDevice, path)
    End If
    Set GetTexture = dicsystem("d3dtextures")(Name)
End Function



Private Sub Class_Terminate()
    Delete3D root

    If dicsystem.Exists("d3dtextures") Then
        dicsystem("d3dtextures").RemoveAll
    End If

    Set D3DDevice = Nothing
    Set D3D = Nothing
    Set dx = Nothing

End Sub


