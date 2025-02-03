VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "K3D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public id As Long
Public category As String
'Public data As New Collection
Public parent As K3D
Public pos As Collection
Public Children As New Collection
Private TriStrip(3) As TLVERTEX
Public TriStripPtr As Long
Public TriStripLen As Long
Public texturePath As String

Public myDx As KDirectX8

'这里没做方向的处理


'Function TransformVertex(ByVal vertex As D3DTLVERTEX, ByVal tx As Single, ByVal ty As Single, ByVal tz As Single, ByVal rotationAngle As Single) As D3DTLVERTEX
'    ' 创建旋转矩阵（绕 Y 轴旋转）
'    Dim rotationMatrix As D3DMATRIX
'    D3DXMatrixRotationY rotationMatrix, rotationAngle
'
'    ' 创建平移矩阵
'    Dim translationMatrix As D3DMATRIX
'    D3DXMatrixTranslation translationMatrix, tx, ty, tz
'
'    ' 矩阵相乘得到最终变换矩阵
'    Dim finalMatrix As D3DMATRIX
'    D3DXMatrixMultiply finalMatrix, rotationMatrix, translationMatrix
'
'    ' 提取顶点的坐标向量
'    Dim vertexVector As D3DVECTOR
'    vertexVector.x = vertex.x
'    vertexVector.y = vertex.y
'    vertexVector.z = vertex.z
'
'    ' 应用变换矩阵到顶点向量
'    Dim transformedVector As D3DVECTOR
'    D3DXVec3TransformCoord transformedVector, vertexVector, finalMatrix
'
'    ' 更新顶点坐标
'    Dim resultVertex As D3DTLVERTEX
'    resultVertex.x = transformedVector.x
'    resultVertex.y = transformedVector.y
'    resultVertex.z = transformedVector.z
'    resultVertex.rhw = vertex.rhw
'    resultVertex.color = vertex.color
'
'    TransformVertex = resultVertex
'End Function

Sub Square(Optional ByVal path As String = "", Optional width As Long = 512, Optional height As Long = 512)

    category = "Square"

    TriStrip(0) = NewTLVertex(0, 0, 0, 1, RGB(255, 255, 255), 0, 0, 0)
    TriStrip(1) = NewTLVertex(width - 1, 0, 0, 1, RGB(255, 255, 255), 0, 1, 0)
    TriStrip(2) = NewTLVertex(0, height - 1, 0, 1, RGB(255, 255, 255), 0, 0, 1)
    TriStrip(3) = NewTLVertex(width - 1, height - 1, 0, 1, RGB(255, 255, 255), 0, 1, 1)
    texturePath = path
    myDx.GetTexture path
    TriStripPtr = VarPtr(TriStrip(0))
    TriStripLen = Len(TriStrip(0))
End Sub


Private Sub Class_Initialize()
    Set pos = iC(0, 0, 0)
    id = ObjPtr(Me)
End Sub

Sub Terminate()
    Dim v As K3D
    If Not Children Is Nothing Then
        While Children.count > 0
            myDx.Delete3D Children(1)
            Children.Remove 1
        Wend
    End If
    Set myDx = Nothing
    Set parent = Nothing
End Sub
