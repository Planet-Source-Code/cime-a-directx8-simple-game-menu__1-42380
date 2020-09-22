Attribute VB_Name = "modKonstante2D"
'Flexible-Vertex-Format opis za 2D vertex (Transformed in Lit)
Public Const FVF = D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR

'Ta struktura predstavlja transformed in lit vertex
Public Type TLVERTEX
    X As Single
    Y As Single
    Z As Single
    rhw As Single
    color As Long
    specular As Long
    tu As Single
    tv As Single
End Type

'//This is just a simple wrapper function that makes filling the structures much much easier...
Public Function CreateTLVertex(X As Single, Y As Single, Z As Single, rhw As Single, color As Long, _
                                               specular As Long, tu As Single, tv As Single) As TLVERTEX
    '//NB: whilst you can pass floating point values for the coordinates (single)
    '       there is little point - Direct3D will just approximate the coordinate by rounding
    '       which may well produce unwanted results....
CreateTLVertex.X = X
CreateTLVertex.Y = Y
CreateTLVertex.Z = Z
CreateTLVertex.rhw = rhw
CreateTLVertex.color = color
CreateTLVertex.specular = specular
CreateTLVertex.tu = tu
CreateTLVertex.tv = tv
End Function

