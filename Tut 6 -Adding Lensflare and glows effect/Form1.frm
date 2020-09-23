VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "NemoX engine Tutrial 6"
   ClientHeight    =   5625
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7020
   LinkTopic       =   "Form1"
   ScaleHeight     =   5625
   ScaleWidth      =   7020
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'========================================
'Tutorial# 6 :Adding LensFlare'
'========================================
'
'
'Welcome to the NemoX engine Tutorial Series
'
'This is the Tutorial# 6 :Adding Lensflare
'for questions and remark go to
'the Engine Website:http://perso.wanadoo.fr/malakoff/index.htm
'
'Or mail me at Johna_pop@yahoo.fr
'
'If you are interresting of making demos,sample,tuts for NemoX engine
'or you want to help me mail Me Johna_pop@yahoo.fr
'
'======================================================
'MAKE SURE YOU HAVE DOWNLOAD THE ENGINE FILES AND HAVE
'SEEN THE TUT #1 at http://perso.wanadoo.fr/malakoff/index.htm
'project section
'=======================================================
'
'
'KEY Arrow to Move left,right,Forward,BackWard
'KEY Numpad 8/2  turn UP/DOWN
'KEY Numpad +/- to move camera UP/DOWN
'


'The first thing to do is to Declarate the Object that
'we will use for this demo


'The Main One is The NemoX renderer
Dim Nemo As NemoX

'we're gonna use a Nemo class for rendering a mesh or polygons
'we use cNemo_Mesh class

Dim Mesh As cNemo_Mesh


'we will need a quick acces to very important and useful functions
Dim Tool As cNemo_Tools




'the sky object
Dim Sky As cNemo_SkyBox



'the lensflare object
'====NEW OBJECT====

Dim LENS As cNemo_LensFlare




Private Sub Form_Load()
 
  
  Me.Show
  Me.Refresh

  'call the initializer sub
  Call InitEngine
  
  
    'make geometry
  Call BuilGeometry
  
  'call the main game Loop
   Call gameLoop
   
 
 
End Sub




'we build our mesh here
Sub BuilGeometry()
        'allocate memory for our meshbuilder
        Set Mesh = New cNemo_Mesh
        
        
        'adding a plane surface for a simple floor$
        
        'first off very important we pass a texture to the meshbuilder
        Mesh.Add_Texture (App.Path + "\Relief_8.jpg")
        Mesh.Add_WallFloor Tool.Vector(-5000, -1, -5000), Tool.Vector(5000, -1, 5000), 10, 10, 0
        
        Mesh.Add_Light 254, 251, 0, 800, Tool.Vector(500, 500, 0), 50
        
        
        'then we build our mesh
        Mesh.BuilMesh
        
        
        
        
        
        
        Set Sky = New cNemo_SkyBox
        Sky.Init_SkyBox 10, 10, 10
        
        Sky.Add_SkyBOX NEMO_SKY_LEFT, App.Path + "\\siege\siege_lf.jpg"
        Sky.Add_SkyBOX NEMO_SKY_RIGHT, App.Path + "\siege\siege_rt.jpg"
        Sky.Add_SkyBOX NEMO_SKY_BACK, App.Path + "\siege\siege_bk.jpg"
        Sky.Add_SkyBOX NEMO_SKY_FRONT, App.Path + "\siege\siege_ft.jpg"
        Sky.Add_SkyBOX NEMO_SKY_DOWN, App.Path + "\siege\siege_dn.jpg"
        Sky.Add_SkyBOX NEMO_SKY_TOP, App.Path + "\siege\siege_up.jpg"
        
        Sky.Set_Scale 9500, 9500, 9500
        
        
        
        
        '============NEW CODE========
        ' setting up the Lensflare
        '===========================
        
        'allocate memory for our Lensflare class
        Set LENS = New cNemo_LensFlare
        
        
        'set our SunGlow textureFile,Position, and size
        LENS.Add_SunTEX App.Path + "\flare\flare_sun.bmp", Tool.Vector(800, 900, 0), 500
        
        'configurate our cameraLensburningout color default=white &HFFFFFFFF
        LENS.Set_SunburningEffectColor 255, 150, 255
        
        'specify that we wil use a static position for our lenslare
        'this allow to simultate Ray of lens from any Static lightSource
        'In this demo we will use a static source
        LENS.Set_LensPositionStatus STATIC_SOURCE
        
        'activate randomizer
        Randomize Timer
        
        'we add 6 Lens Spark.....TextureFile,index,size,position,Color,SpecularColor
        'TextureFile =any TGA,BMP,JPG file
        'index       =the Index for the lensSpark
        'size        =the size of our spark
        'position    =the position the sun pos is 100, middle distance=50
        'Color       =main Color for the spark
        'highlight   =color for the outter of the circle
        
        LENS.Add_Lens App.Path + "\flare\flare1.jpg", 0, 15, 90, &HFFFFFFFF * Rnd, &HFFFFFFFF * Rnd
        LENS.Add_Lens App.Path + "\flare\flare2.jpg", 1, 10, 50, &HFFFFFFFF * Rnd, &HFFFFFFFF * Rnd
        LENS.Add_Lens App.Path + "\flare\flare3.jpg", 2, 7, 30, &HFFFFFFFF * Rnd, &HFFFFFFFF * Rnd
        LENS.Add_Lens App.Path + "\flare\flare4.jpg", 3, 5, 15, &HFFFFFFFF * Rnd, &HFFFFFFFF * Rnd
        LENS.Add_Lens App.Path + "\flare\flare5.jpg", 4, 3, 0, &HFFFFFFFF * Rnd, &HFFFFFFFF * Rnd
        LENS.Add_Lens App.Path + "\flare\flare6.jpg", 5, 1, -4, &HFFFFFFFF * Rnd, &HFFFFFFFF * Rnd
        
      
      'this next line allow to specify the blending mode
      'LENS.Set_LenflareBlend FLARES_BLENDMODE, D3DBLEND_ONE, D3DBLEND_DESTALPHA

End Sub




'we will used that sub for the engine initialization

Sub InitEngine()

 'first thing allocate memory for the main Object
  
  Set Nemo = New NemoX
  
  Set Tool = New cNemo_Tools


'we use this method
'now we allow the user to choose options
'32/16 bit backbuffer
  
  If Not (Nemo.INIT_ShowDeviceDLG(Form1.hWnd)) Then
   End 'terminate here if error
  End If
  
  
  'set the back clearcolor
  Nemo.BackBuffer_ClearCOLOR = &HFFC0C0C0
  
  
  'set some parameters
  Nemo.Set_ViewFrustum 1, 17000, 3.14 / 4
  Nemo.Set_light True
  
  'set our camera
  Nemo.Camera_set_EYE Tool.Vector(0, 15, 0)
  
  
End Sub



'this sub is the main loop for a game or 3d apllication
Sub gameLoop()

            'loop untill player press 'ESCAPE'
            
    Do
            
               '=====Keyboard handler can be added here
               Call GetKey
               DoEvents
               
               
               
               'start the 3d renderer
               Nemo.Begin3D
                     '===============ADD game rendering mrthod here
               
               'draw our plane_space
               
               
               'render the sky
               Sky.RenderSky
               
               'draw our floor
               Mesh.Render
               
               
               'Render lastly our lensflare
               LENS.Render
               
               
               
               
               
               'show the FPS at pixel(5,10) color White
               Nemo.Draw_Text "FPS:" + Str(Nemo.Framesperseconde), 5, 10, &HFFFF0000
               'Nemo.Draw_Text "Position X=" + Str(Nemo.Camera_GetPosition.x), 5, 25, &HFFFFFFFF
               'Nemo.Draw_Text "Position Y=" + Str(Nemo.Camera_GetPosition.y), 5, 40, &HFFFFFFFF
               'Nemo.Draw_Text "Position Z=" + Str(Nemo.Camera_GetPosition.z), 5, 55, &HFFFFFFFF
               
               
               Nemo.End3D
               'end the 3d renderer
            
            'check the player keyPressed
    Loop Until Nemo.Get_KeyPress(NEMO_KEY_ESCAPE)
            
            
    Call EndGame

End Sub



'----------------------------------------
'Name: GetKey
'----------------------------------------
Sub GetKey()



   'just Rotate left and Right
If Nemo.Get_KeyPress(NEMO_KEY_LEFT) Then _
    Nemo.Camera_Turn_Left 1.5 / 50
If Nemo.Get_KeyPress(NEMO_KEY_RIGHT) Then _
    Nemo.Camera_Turn_Right 1.5 / 50
    
    
    'just move Forward and Backward
    If Nemo.Get_KeyPress(NEMO_KEY_UP) Then _
    Nemo.Camera_Move_Foward 1
    If Nemo.Get_KeyPress(NEMO_KEY_RCONTROL) Then _
    Nemo.Camera_Move_Foward 10
   If Nemo.Get_KeyPress(NEMO_KEY_DOWN) Then _
    Nemo.Camera_Move_Backward 1


   'just move uP and Down
  If Nemo.Get_KeyPress(NEMO_KEY_ADD) Then _
    Nemo.Camera_Strafe_UP 1
  If Nemo.Get_KeyPress(NEMO_KEY_SUBTRACT) Then _
    Nemo.Camera_Strafe_DOWN 1


      'to rotate UP/DOWN
     If Nemo.Get_KeyPress(NEMO_KEY_NUMPAD8) Then
        Nemo.Camera_Turn_UP 1 / 50
    End If
    If Nemo.Get_KeyPress(NEMO_KEY_NUMPAD2) Then _
        Nemo.Camera_Turn_DOWN 1 / 50
        
        'to move 0,8,-8
    If Nemo.Get_KeyPress(NEMO_KEY_SPACE) Then _
     Nemo.Camera_SetPosition Vector(0#, 8#, -8#), _
                                    Vector(0#, 8#, 500#)
    
    'to take a snapshot
    If Nemo.Get_KeyPress(NEMO_KEY_S) Then _
     Nemo.Take_SnapShot App.Path + "\Shot.bmp"




End Sub



Sub EndGame()
 'end of the demo
            Nemo.Free  'free resources used by the engine
            End
End Sub



