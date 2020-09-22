Attribute VB_Name = "modInit"
Public Function InitD3D(frm As Form, bWindowed As Boolean) As Boolean
    '//ustvarimo in preverimo D3D object
    Set g_D3D = g_Dx.Direct3DCreate()
    If g_D3D Is Nothing Then GoTo e:
    '//
    
    '//dobimo trenutni naèin prikaza
    g_D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispMode
    '//
    
    '//parametri za ustvaritev nove naprave
    d3dpp.BackBufferHeight = DispMode.Height 'npr. 768
    d3dpp.BackBufferWidth = DispMode.Width 'npr. 1024
    d3dpp.BackBufferFormat = DispMode.Format
    d3dpp.EnableAutoDepthStencil = 1
    d3dpp.AutoDepthStencilFormat = D3DFMT_D16
    d3dpp.SwapEffect = D3DSWAPEFFECT_COPY_VSYNC
    d3dpp.Windowed = bWindowed 'ali na program deluje v fullscrenu
    '//
    
    '//ustvarimo in preverimo napravo
    Set g_D3Device = g_D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, frm.hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, d3dpp)
    If g_D3Device Is Nothing Then GoTo e:
    '//
    
    '//nastavimo parametre renderiranja
    g_D3Device.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
    g_D3Device.SetRenderState D3DRS_ZENABLE, 1
    g_D3Device.SetRenderState D3DRS_FILLMODE, D3DFILL_SOLID
  
    g_D3Device.SetTextureStageState 0, D3DTSS_ADDRESSU, D3DTADDRESS_WRAP
    g_D3Device.SetTextureStageState 0, D3DTSS_ADDRESSV, D3DTADDRESS_WRAP
  
    g_D3Device.SetRenderState D3DRS_LIGHTING, 1 'Turn on lighting
    g_D3Device.SetVertexShader D3DFVF_CUSTOMVERTEX 'Set vertex shader to our custom type
    '//
    
    '//vse je v redu
    InitD3D = True
    Exit Function

e: '//napaka
    InitD3D = False
End Function


Public Function InitDI(frm As Form) As Boolean
    '//ustvarimo object Input
    Set g_DI = g_Dx.DirectInputCreate()
    '//
    
    
    '//ustvarimo napravo za miško
    Set Mouse = g_DI.CreateDevice("GUID_SysMouse")
    If Mouse Is Nothing Then GoTo e:
    '//ustvarimo napravo za tipkovnico
    Set Keyboard = g_DI.CreateDevice("GUID_SysKeyboard")
    If Keyboard Is Nothing Then GoTo e:
    
    
    '//nastavimo standardni format tipkovnice in miške
    Mouse.SetCommonDataFormat DIFORMAT_MOUSE
    Keyboard.SetCommonDataFormat DIFORMAT_KEYBOARD
    '//
    
    '//nastavimo kontrolo, ki jo želimo nad napravama
    Mouse.SetCooperativeLevel frm.hWnd, DISCL_NONEXCLUSIVE Or DISCL_BACKGROUND
    Keyboard.SetCooperativeLevel frm.hWnd, DISCL_NONEXCLUSIVE Or DISCL_BACKGROUND
    '//
    
    '//ko je vse nastavljeno, vkljuèimo napravi
    Mouse.Acquire
    Keyboard.Acquire
    '//
    
    '//vse je v redu
    InitDI = True
    Exit Function

e: '//napaka
    InitDI = False
End Function

Public Function InitDS(frm As Form) As Boolean
    '//ustvarimo object DirectSound
    Set s_DS = g_Dx.DirectSoundCreate("")
    If s_DS Is Nothing Then GoTo e:
    '//
    
    
    '//nastavimo DirectSound object
    s_DS.SetCooperativeLevel frm.hWnd, DSSCL_NORMAL
    '//
    
    '//ustvarimo še zvoke
    'Set s_Secondary = s_DS.CreateSoundBufferFromFile(App.Path & "\zvoki\datoteka2.wav", s_DescSecondary)
    
    '//vse je v redu
    InitDS = True
    Exit Function

e: '//napaka
    InitDS = False
End Function
