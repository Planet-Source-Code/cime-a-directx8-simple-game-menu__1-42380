Attribute VB_Name = "modInput"
Public Tipke(255) As Boolean

Public Type MouseType
    X As Single
    Y As Single
    Z As Single
    Buttons(3) As Boolean
End Type

Public Miska As MouseType

Public Sub TipkovnicaMiska()
    '//poklièe funkcijo, ki prebere informacije
    Call DobiInput
    
    If Not Tipke(DIK_ESCAPE) = False Then
        If Not Izbira = MenuOsnovni Then
            Izbira = MenuOsnovni
        Else
            bRunning = False
        End If
    End If
End Sub

Public Sub DobiInput()

    '//nastavi vse tipke na false
    For i = 0 To 255
        Tipke(i) = False
    Next i
    
    '//dobi sveže informacije o tipkah
    Keyboard.GetDeviceStateKeyboard KeyboardState

    '//Dobi informacije o pritisnjenih Buttonsh
    For i = 0 To 255
        If KeyboardState.Key(i) <> 0 Then
            Tipke(i) = True
        End If
    Next i
    
    '//dobi sveže informacije o miški
    Mouse.GetDeviceStateMouse MouseState
        
    '//nastavimo vse gumbe na false
    For i = 0 To 3
        Miska.Buttons(i) = False
    Next i
        
    '//dobi informacije o Buttonsh na miski
    For i = 0 To 3
        If MouseState.Buttons(i) <> 0 Then
            Miska.Buttons(i) = True
        End If
    Next i
    
    '//dobi informacije o premiku miške
    Miska.X = Miska.X + MouseState.lX
    Miska.Y = Miska.Y + MouseState.lY
    Miska.Z = Miska.Z + MouseState.lZ
    
    DoEvents '//poèakamo, da se vse sprocesira
End Sub
