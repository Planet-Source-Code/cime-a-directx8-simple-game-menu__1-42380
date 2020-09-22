Attribute VB_Name = "modRender"
Public Const MenuOsnovni = 0
Public Const MenuNovaIgra = 1
Public Const MenuNastavitve = 2
Public Const MenuAvtor = 3
Public Const MenuIzhod = 4
Public Const Igra = 5

Public Izbira As Integer

Public Sub Render()
    '//poèisti celoten zaslon
    g_D3Device.Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, RGB(200, 0, 0), 1#, 0
    
    g_D3Device.BeginScene ' tukaj se zaène scena

        Select Case Izbira
            Case Menu
                Call MOsnovni
            Case MenuNovaIgra
                Call MNovaIgra
            Case MenuNastavitve
                Call MNastavitve
            Case MenuAvtor
                Call MAvtor
            Case MenuIzhod
                Call MIzhod
            Case Igra
                Call RenderIgra
        End Select
    
    g_D3Device.EndScene 'tukaj se scena konèa
    
    '//vse skupaj prikažemo na zaslon
    g_D3Device.Present ByVal 0, ByVal 0, 0, ByVal 0
End Sub

