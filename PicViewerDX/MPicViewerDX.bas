Attribute VB_Name = "MPicViewerDX"
Type Vector2
    x As Single
    y As Single
End Type

Public Type TLVERTEX
    x As Single
    y As Single
    z As Single
    rhw As Single
    color As Long
    specular As Long
    tu As Single
    tv As Single
End Type

Type Vector3
    x As Single
    y As Single
    z As Single
End Type

Function MarkName(ByVal path)
    MarkName = suffix(path, App.path & "\")
End Function


Function NewTLVertex(x As Single, y As Single, z As Single, rhw As Single, color As Long, _
                                               specular As Long, tu As Single, tv As Single) As TLVERTEX
    NewTLVertex.x = x
    NewTLVertex.y = y
    NewTLVertex.z = z
    NewTLVertex.rhw = rhw
    NewTLVertex.color = color
    NewTLVertex.specular = specular
    NewTLVertex.tu = tu
    NewTLVertex.tv = tv
End Function

