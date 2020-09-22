Attribute VB_Name = "modOsnovniMenu"
Public TriStrip(0 To 3) As TLVERTEX

Public Sub MOsnovni() '//draw zaèetni menu
    '//za kvadrat potrebujemo 4 vertecise
    
    
    g_D3Device.SetVertexShader FVF
    g_D3Device.SetRenderState D3DRS_LIGHTING, False
    
    '////////////////////////////////////////////////
    '///////        Ozadje menuja             ///////
    TriStrip(0) = CreateTLVertex(0, 0, 0, 1, RGB(105, 0, 105), 0, 0, 0)
    TriStrip(1) = CreateTLVertex(1024, 0, 0, 1, RGB(255, 0, 0), 0, 0, 0)
    TriStrip(2) = CreateTLVertex(0, 768, 0, 1, RGB(0, 255, 0), 0, 0, 0)
    TriStrip(3) = CreateTLVertex(1024, 768, 0, 1, RGB(0, 0, 255), 0, 0, 0)
    '//draw
    g_D3Device.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, TriStrip(0), Len(TriStrip(0))
    '////////////////////////////////////////////////

    '////////////////////////////////////////////////
    '///////        Gumb nova igra            ///////
    TriStrip(0) = CreateTLVertex(20, 30, 0, 1, RGB(0, 0, 255), 0, 0, 0)
    TriStrip(1) = CreateTLVertex(200, 30, 0, 1, RGB(0, 0, 255), 0, 0, 0)
    TriStrip(2) = CreateTLVertex(20, 80, 0, 1, RGB(0, 0, 255), 0, 0, 0)
    TriStrip(3) = CreateTLVertex(200, 80, 0, 1, RGB(0, 0, 255), 0, 0, 0)
    Buttons(0).X = 20
    Buttons(0).Y = 30
    '//draw
    g_D3Device.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, TriStrip(0), Len(TriStrip(0))
    '////////////////////////////////////////////////

For i = 1 To 4
    '////////////////////////////////////////////////
    '///////         Podlage za Gumbe         ///////
    TriStrip(0) = CreateTLVertex(20, 30 + i * 80, 0, 1, RGB(0, 0, 255), 0, 0, 0)
    TriStrip(1) = CreateTLVertex(200, 30 + i * 80, 0, 1, RGB(0, 0, 255), 0, 0, 0)
    TriStrip(2) = CreateTLVertex(20, 80 + i * 80, 0, 1, RGB(0, 0, 255), 0, 0, 0)
    TriStrip(3) = CreateTLVertex(200, 80 + i * 80, 0, 1, RGB(0, 0, 255), 0, 0, 0)
    Buttons(i).X = 20
    Buttons(i).Y = 30 + (i * 80)
    '//draw
    g_D3Device.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, TriStrip(0), Len(TriStrip(0))
    '////////////////////////////////////////////////
Next i

    Call Cursor(True)

    g_D3Device.SetRenderState D3DRS_LIGHTING, True

End Sub
