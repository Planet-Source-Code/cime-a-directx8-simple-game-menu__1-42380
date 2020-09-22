Attribute VB_Name = "modKazalec"
Public Cursor1(0 To 2) As TLVERTEX
Public Cursor2(0 To 2) As TLVERTEX
Public Cursor3(0 To 2) As TLVERTEX

Public ColorCursor As Long

Public Type Buttonses
    X As Single
    Y As Single
End Type

Public Buttons(0 To 5) As Buttonses

Public Function MenuCollision(X As Single, Y As Single) As Boolean
    For i = LBound(Buttons) To UBound(Buttons)
        If X > Buttons(i).X - 10 And Y > Buttons(i).Y - 20 And X <= (Buttons(i).X + 170) And Y <= (Buttons(i).Y + 30) Then
            MenuCollision = True
            GoTo ok:
        Else
            MenuCollision = False
        End If
    Next i
ok:
End Function

Public Sub Cursor(visible As Boolean)

    If MenuCollision(Miska.X, Miska.Y) = True Then
        ColorCursor = RGB(60, 220, 230)
    Else
        ColorCursor = RGB(250, 250, 0)
    End If
    
    If Miska.X <= 30 Then Miska.X = 30
    If Miska.Y <= 10 Then Miska.Y = 10
    If Miska.X >= DispMode.Width - 30 Then Miska.X = DispMode.Width - 30
    If Miska.Y >= DispMode.Height - 50 Then Miska.Y = DispMode.Height - 50
        
If visible = True Then
    '////////////////////////////////////////////////
    '///////          Mouse Cursor            ///////
    Cursor1(0) = CreateTLVertex(Miska.X - 5, Miska.Y + 55, 0, 1, ColorCursor, 0, 0, 0)
    Cursor1(1) = CreateTLVertex(Miska.X + 8, Miska.Y + 25, 0, 1, RGB(0, 0, 255), 0, 0, 0)
    Cursor1(2) = CreateTLVertex(Miska.X + 20, Miska.Y + 55, 0, 1, ColorCursor, 0, 0, 0)
    
    Cursor2(0) = CreateTLVertex(Miska.X - 25, Miska.Y + 10, 0, 1, ColorCursor, 0, 0, 0)
    Cursor2(1) = CreateTLVertex(Miska.X - 10, Miska.Y - 10, 0, 1, ColorCursor, 0, 0, 0)
    Cursor2(2) = CreateTLVertex(Miska.X + 5, Miska.Y + 20, 0, 1, RGB(0, 0, 255), 0, 0, 0)
    
    Cursor3(0) = CreateTLVertex(Miska.X + 40, Miska.Y + 10, 0, 1, ColorCursor, 0, 0, 0)
    Cursor3(1) = CreateTLVertex(Miska.X + 25, Miska.Y - 10, 0, 1, ColorCursor, 0, 0, 0)
    Cursor3(2) = CreateTLVertex(Miska.X + 10, Miska.Y + 20, 0, 1, RGB(0, 0, 255), 0, 0, 0)
    
    '//draw
    g_D3Device.DrawPrimitiveUP D3DPT_TRIANGLELIST, 2, Cursor1(0), Len(Cursor1(0))
    g_D3Device.DrawPrimitiveUP D3DPT_TRIANGLELIST, 2, Cursor2(0), Len(Cursor2(0))
    g_D3Device.DrawPrimitiveUP D3DPT_TRIANGLELIST, 2, Cursor3(0), Len(Cursor3(0))
    '////////////////////////////////////////////////
End If

End Sub


