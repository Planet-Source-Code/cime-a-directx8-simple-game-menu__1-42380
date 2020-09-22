Attribute VB_Name = "modMain"
'######################################################
'#              A SIMPLE GAME MENU EXAMPLE            #
'#              By JasonX (on PSC) / Cime             #
'#          email : cimethebest@volja.net             #
'#  you can mail me if you have answer for me         #
'#                                                    #
'# And sorry for my bad English, I'm from Slovenia.   #
'#                                                    #
'######################################################
'# I have seen many of good 3D engine on PSC, but     #
'# nono of them had a menu, so I decided to make an   #
'# exaple on how to make a simple game menu in        #
'# DirectX 8.                                         #
'######################################################
'#                                                    #
'# You can freely use this code in your games, but    #
'# please mail me, if you will use it.                #
'#          email : cimethebest@volja.net             #
'######################################################
'#                                                    #
'#              Vote for me on PSC :                  #
'#  http://www.planet-source-code.com/vb              #
'#                                                    #
'######################################################

'//main object of DirectX
Public g_Dx As New DirectX8

'------------------------
'|  objectI ZA PRIKAZ   |
'------------------------

'//D3D object
Public g_D3D As Direct3D8
'//D3DX object, ki vsebuje pomožne funkcije
Public g_D3Dx As D3DX8

'//naprava za prikazovanje
Public g_D3Device As Direct3DDevice8

'//naèin prikaza
Public DispMode As D3DDISPLAYMODE
'//parametri za prikaz
Public d3dpp As D3DPRESENT_PARAMETERS

'--------------------------------
'|  objectI ZA VNOSNE NAPRAVE   |
'--------------------------------

'//glavni object za ustvarjanje input naprav
Public g_DI As DirectInput8

'//object for mouse
Public Mouse As DirectInputDevice8
'//object za tipkovnico
Public Keyboard As DirectInputDevice8

'//mouse state
Public MouseState As DIMOUSESTATE
'//keyboard state
Public KeyboardState As DIKEYBOARDSTATE

'----------------------
'|  objectI ZA ZVOK   |
'----------------------

'//glavni object za zvok-sound
Public s_DS As DirectSound8

'//ustvarimo primarni in sekundarni buffer
Public s_Secondary As DirectSoundSecondaryBuffer8

'//ustvarimo opis bufferjev za zvok
Public s_DescSecondary As DSBUFFERDESC

'----------------------------
'|  KONSTANTE IN STRUKTURE  |
'----------------------------
Public Const D3DFVF_CUSTOMVERTEX = (D3DFVF_XYZ Or D3DFVF_DIFFUSE Or D3DFVF_TEX1)

Public Const PI = 3.141592576 'PI konstanta

Public Type CustomVertex
    Position As D3DVECTOR 'Pozicija
    color As Long 'Color
    tu As Single 'Koordinate texture
    tv As Single
End Type

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type POINTAPI
    X As Single
    Y As Single
End Type

Public bRunning As Boolean '//pove nam ali naj igra še deluje

Sub Main()
    '//nastavimo na celozaslonski naèin
    Call Startup(frmMain, False)
    
    '//ponavljanje zanke
    bRunning = True
    
    '//izberemo osnovni menu za zaèetek
    Izbira = MenuOsnovni
    
    Do '//glavna zanka v kateri se vse izvede
    
        If bRunning = False Then GoTo izhod:
    
        '//predela podatke iz tipkovnice in miške
        Call TipkovnicaMiska

        
        '//nastavi nove pozicije/velikosti
        Call NastaviPredmete
        
        '//renderira grafiko
        Call Render
        
        DoEvents '//poèakamo, da se vse konèa
    Loop While bRunning
    
izhod:
    Set g_DI = Nothing
    Set g_Dx = Nothing
    Set s_DS = Nothing
    Set g_D3Device = Nothing
    Unload frmMain
End Sub

Public Function Startup(frm As Form, bWindowed As Boolean) As Boolean
Dim AllOK(0 To 5) As Boolean
    frm.Show
    '//preverimo inicializacijo vsega
    AllOK(0) = InitD3D(frm, bWindowed) '//inicializiramo D3D
    AllOK(1) = InitDI(frm) '//inicializiramo vnosne naprave
    AllOK(2) = InitDS(frm) '//inicializiramo DirectSound
    
    For i = LBound(AllOK) To UBound(AllOK)
        If Not AllOK(i) = True Then GoTo e:
    Next i
    
    '//vse je v redu
    Startup = True
    Exit Function
    
e: '//napaka
    Startup = False
End Function
