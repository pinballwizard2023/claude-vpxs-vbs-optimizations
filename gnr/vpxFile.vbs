  ' Based on the PM5 Table :
 ' Tabla Guns And Roses creada por Data East en 1994
 ' Graficos para el VP por Lord Hiryu
 ' Script por JP Salas


' Possible issue : ball not be launched in the right launch lane. Wait a short moment and it will be.


' VPX version :
' translated for VPX by Team PP 2017 (NeoFR45, Aetios, Chucky87, Arngrim and JPJ)
' Thanks to JP Salas to permit this translation
' Thanks Knorr for your Sound pack !!!! Very Usefull, and very clean !!!
' Thanks to Brian Ryan for helping in choosing environnement file, and for the Shadow Method ;) ;) ;)
' Very big Thanks Chucky for your lynx's eyes !!! :)
' Graphism : Néo and JPJ
' 3d : Aetios and JPJ At first, and finaly a BIG HELP to improve all from Ninuzzu and Tom Tower !!!!
' Dof : Arngrim
' Script : JP Salas and JPJ for adaptation in VPX and scripting animations and lights and more ;)
'		   And Ninuzzu who helped me resolving the plunger launch method !!!!
' Big Thanks for beta testing : Jens Leiensetter (see his youtube channel : Bambi Platfuss)
'								Peskopat, Mariopourlavie
'								My Super Chucky, always there
'								Shadow, Arngrim, Bertie
' Thanks to Bertie for his GNR death picture in the spinner of the Yellow ramp !!!
' Thanks to the Pinball community for all part of script found everywhere, and for all your tips and tricks ;) ;) ;) and your kindness for answer when support is needed (Ninuzzu, Brian, JPS)
' Another Thanks for Jens for all of his photos of the real pinball !!! Helpful for finalising, and searching some finition's detail (not all but a lot i've seen).
' Thanks for all of you i forget :)
'
' JPJ - Team PP

Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

'Constantes
'///////////////

Const VolumeDial = 0.8				'fleep
Const BallRollVolume = 0.5 			'fleep'Level of ball rolling volume. Value between 0 and 1
Const RampRollVolume = 0.5 			'fleep'Level of ramp rolling volume. Value between 0 and 1
Dim tablewidth: tablewidth = Table1.width 'fleep
Dim tableheight: tableheight = Table1.height 'fleep
Dim LUTset, LutToggleSound 'LUT
Dim Ballsize,BallMass
BallSize = 50
BallMass = 1 '(Ballsize^3)/125000

'----- Phsyics Mods -----
Const RubberizerEnabled = 0			'0 = normal flip rubber, 1 = more lively rubber for flips
Const FlipperCoilRampupMode = 0   	'0 = fast, 1 = medium, 2 = slow (tap passes should work)

' ****** LUT overall contrast & brightness setting **************************************************************************************
Dim luts, lutpos
luts = array("LUTVogliadicane80", "LUTVogliadicane70", "1to1", "Fleep Natural Dark 1", "Fleep Natural Dark 2", "Fleep Warm Bright", "Fleep Warm Dark", "3rdaxis Referenced THX Standard", "CalleV Punchy Brightness and Contrast", "Skitso Natural and Balanced", "Skitso Natural High Contrast", "LUTbassgeige1", "LUTbassgeige2", "LUTbassgeigemeddark", "LUTbassgeigemeddarkwhite", "LUTbassgeigeultrdark", "LUTbassgeigeultrdarkwhite", "LUTblacklight", "LUTfleep", "LUTmandolin", "LUTmlager8", "LUTmlager8night", "LUTrobertmstotan0_darker", "luttotan4mhcontrastfilmic2", "LUTtotan1", "LUTtotan2", "LUTtotan4", "LUTtotan5", "LUTtotan6")
'lutpos = 0						'  set the nr of the LUT you want to use (0 = first in the list above, 1 = second, etc); 0 is the default
'Table1.ColorGradeImage = luts(lutpos)
Const EnableMagnasave = 1		' 1 - on; 0 - off; if on then the magnasave button let's you rotate all LUT's
LoadLUT 'LUT
SetLUT 'LUT
'****************************************************************************************************************************************

Const lob = 1	'locked balls on start; might need some fiddling depending on how your locked balls are done
Const DynamicBallShadowsOn = 1		'0 = no dynamic ball shadow ("triangles" near slings and such), 1 = enable dynamic ball shadow
Const AmbientBallShadowOn = 1		'0 = Static shadow under ball ("flasher" image, like JP's)
'									'1 = Moving ball shadow ("primitive" object, like ninuzzu's) - This is the only one that shows up on the pf when in ramps and fades when close to lights!
'									'2 = flasher image shadow, but it moves like ninuzzu's

' Thalamus 2019 November : Improved directional sounds
' !! NOTE : Table not verified yet !!

' Options
' Volume devided by - lower gets higher sound

Const VolDiv = 2000    ' Lower number, louder ballrolling/collition sound
Const VolCol = 10      ' Ball collition divider ( voldiv/volcol )

'Dim DesktopMode:DesktopMode = Table1.ShowDT

'*********************************************************************************************************
'*** to show dmd in desktop Mod - Taken from ACDC Ninuzzu (THX) And Thanks To Rob Ross for Helping *******
Dim UseVPMDMD, DesktopMode
DesktopMode = Table1.ShowDT
If NOT DesktopMode Then UseVPMDMD = False		'hides the internal VPMDMD when using the color ROM or when table is in Full Screen and color ROM is not in use
If DesktopMode Then UseVPMDMD = True							'shows the internal VPMDMD when in desktop mode
'*********************************************************************************************************

LoadVPM "01120100", "de.vbs", 3.02

' VR Room Auto-Detect - Suggested LUT 00 or 01 for darker room - LUT 02 for lighter room.
Dim VR_Obj, VR_Room

If RenderingMode = 2 Then
	VR_Room = 1
	Primitive005.Visible = 0
	Ramp15.Visible = 0
	Ramp16.Visible = 0
	For Each VR_Obj in VRCabinet : VR_Obj.Visible = 1 : Next
	For Each VR_Obj in VRRoom : VR_Obj.Visible = 1 : Next
Else
	VR_Room = 0
	For Each VR_Obj in VRCabinet : VR_Obj.Visible = 0 : Next
	For Each VR_Obj in VRRoom : VR_Obj.Visible = 0 : Next
End If


Dim bsTrough, bsBallRelease, bsCenterScoop, bsEject, dtLBank, dtRBank, bsVuk, mLeftMagnet, mRightMagnet, mCenterMagnet, cbCaptive
Dim x, bump1, bump2, bump3, BallInVuk, VukStep, Hole, seq, seqb, swrot, grvar, gr, balldropA, balldropB, testJP
Dim SndRedRamp, SndYellowRamp, SndRedRampV, SndYellowRampV, SndLaunchRamp, SndLaunchRampV, MH
'Dim GlobalSoundLevel 'Diner's Method for amplify mechanical sounds, thanks to them :
Dim Myst, Mystv, Mystnudge, numcap
Dim OptionOpacity
Dim Cap(17)
Dim light
Dim FlashClearEtape
Dim FlashClearEtapeR
Dim FlashGreenEtape
Dim FlashRedEtape
Dim sides
'Dim FastFlips
GlobalSoundLevel = 2


'**********************************************
'** TRY Option Myst Mod : 0 = off or 1 = On  **
'**********************************************
Myst = 0                      			         '**
'**********************************************************************
'**     Option sides or not sides in FS : 0 = without or 1 = with    **
'**********************************************************************
sides = 1															'**
'**********************************************************************

 Const cGameName = "gnr_300"

 Const UseSolenoids = 2
 Const UseLamps = 1
 Const UseGI = 1
 Const UseSync = 0
 Const HandleMech = 0


 ' Standard Sounds
Const SSolenoidOn="solon",SSolenoidOff="soloff",SFlipperOn="FlipperUp",SFlipperOff="FlipperDown",SCoin="coin_in_1" 'fleep

 ' TT addition to fix ball hangs in kickers
 Dim BallInSw38


 'Dim BallInCapKicker
 '************
 ' Table init.
 '************

 Sub Table1_Init
vpmInit Me
'dm


     With Controller
         .GameName = cGameName
		If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description:Exit Sub
         .SplashInfoLine = "Guns 'N' Roses, Data East 1994" & vbNewLine & "VPM table by Lord Hiryu & JPSalas v.1.0"
		.Games(cGameName).Settings.Value("sound") = 1
         .HandleMechanics = 0
         .HandleKeyboard = 0
         .ShowDMDOnly = 1
         .ShowFrame = 0
         .ShowTitle = 0
         .Hidden = DesktopMode
		  If Err Then MsgBox Err.Description
		On Error Goto 0
     End With

	Controller.SolMask(0) = 0
     vpmTimer.AddTimer 2000, "Controller.SolMask(0)=&Hffffffff'" 'ignore all solenoids - then add the Timer to renable all the solenoids after 2 seconds
     Controller.Run

     ' Nudging
     vpmNudge.TiltSwitch = 1
     vpmNudge.Sensitivity = 5
     vpmNudge.TiltObj = Array(Bumper1, Bumper2, Bumper3, LeftSlingshot, RightSlingshot, LeftSlingShotH)


     ' Trough & Ball Release
 Set bsTrough = New cvpmBallStack
	With bsTrough
		.InitSw 0, 14, 13, 12, 11, 10, 9, 0
		'.InitSw 0, 14, 13, 12, 11, 10, 9, 0
 		.Balls = 6
	End With


     Set bsBallRelease = New cvpmBallStack
     With bsBallRelease
         .InitSaucer BallRelease, 150, 110, 20
         .InitEntrySnd "solenoid", "solenoid"
         .InitExitSnd SoundFX("ballrelease1",DOFContactors), SoundFX("solenoid",DOFContactors)
     End With

     ' Center Scoop
     Set bsCenterScoop = New cvpmBallStack
     With bsCenterScoop
         .InitSw 0, 38, 0, 0, 0, 0, 0, 0
         .InitKick sw38a, 192, 20
         .KickZ = 0.4
         .KickForceVar = 2
		.KickBalls = 3
         .InitExitSnd SoundFX("saucer_kick",DOFContactors), SoundFX("solenoid",DOFContactors)
     End With

     Set bsEject = New cvpmBallStack
     With bsEject
         .InitSaucer sw37, 37, 280, 10
		.KickZ = 0.4
         .KickForceVar = 2
		.InitExitSnd SoundFX("saucer_kick",DOFContactors), SoundFX("solenoid",DOFContactors)
     End With

     'vuk

     Set bsVuk = New cvpmBallStack
     With bsVuk
         .InitSw 0, 39, 0, 0, 0, 0, 0, 0
         .InitKick sw39a, 0, 90
		.KickZ = 1.56
		.KickForceVar = 2
         .InitExitSnd SoundFX("saucer_kick",DOFContactors), SoundFX("solenoid",DOFContactors)
     End With

 '   Set bsVuk = New cvpmBallStack
 '   With bsVuk
'
'        .InitSaucer Sw39a, 39, 0, 60
'        .KickZ = 1.56
'        '.KickZ = 1.57
'        .KickForceVar = 2
'        .InitExitSnd SoundFX("ballrel",DOFContactors), SoundFX("solenoid",DOFContactors)
'    End With


     ' Magnets
     Set mLeftMagnet = New cvpmMagnet
     With mLeftMagnet
         .InitMagnet LeftMagnet, 16
         .Solenoid = 51
         .GrabCenter = 0
         .CreateEvents "mLeftMagnet"
     End With

     Set mRightMagnet = New cvpmMagnet
     With mRightMagnet
         .InitMagnet RightMagnet, 16
         .Solenoid = 53
         .GrabCenter = 0
         .CreateEvents "mRightMagnet"
     End With

     Set mCenterMagnet = New cvpmMagnet
     With mCenterMagnet
         .InitMagnet CenterMagnet, 16
         .Solenoid = 52
         .GrabCenter = 0
         .CreateEvents "mCenterMagnet"
     End With

''**** Fastflips
'    Set FastFlips = new cFastFlips
'    with FastFlips
'       .CallBackL = "SolLflipper"  'Point these to flipper subs
'       .CallBackR = "SolRflipper"  '...
''       .CallBackUL = "SolULflipper"'...(upper flippers, if needed)
''       .CallBackUR = "SolURflipper"'...
'       .TiltObjects = True 'Optional, if True calls vpmnudge.solgameon automatically. IF YOU GET A LINE 1 ERROR, DISABLE THIS! (or setup vpmNudge.TiltObj!)
' '      .DebugOn = True        'Debug, always-on flippers. Call FastFlips.DebugOn True or False in debugger to enable/disable.
'    end with

     'Captive Ball
	'CapKicker.CreateSizedBallWithMass Ballsize, BallMass:CapKicker.Kick 0,3:CapKicker.enabled = 0

	' replaced captive ball with directly creating ball to kicker
	 vpmCreateBall CapKicker:CapKicker.Kick 0,3:CapKicker.enabled = 0



     ' Init GI
     GiOff


     ' Init Plungers & Div
'	Plunger.Pullback
	KickBack.Pullback
     SolTrapDoor 0

     'Mas flashers (27-04-09)
     f8.state = 0
     f8a.state = 0
     f7.state = 0
     f7a.state = 0

     ' Turn of Flashers
     SolFlash25 0:SolFlash26 0:SolFlash27 0:SolFlash28 0:SolFlash29 0:SolFlash30 0

     ' Setup Lamps
     vpmMapLights AllLamps

     ' Main Timer init
     PinMAMETimer.Interval = PinMAMEInterval
     PinMAMETimer.Enabled = 1

     ' GI's
	flash5.state = 0
	flash6.state = 0:
	flash5i.state = 1:flash6i.state = 1
	Flash5i.intensity = 0.3:flash6i.intensity = 0.3
	flash2veille.state = 0 :flash2veille1.state = 0
	flash2veille.intensity = 0.75 :flash2veille1.intensity = 0.75
	BallInSw38 = 0
	 'BallInCapKicker = 1
	slingflashL.state = 0
	slingflashR.state = 0
		GIBWL.color = RGB(212,212,212)
		GIBWL.intensity = 2
		GIBWR.color = RGB(212,212,212)
		GIBWR.intensity = 2
	GIBWL.state = 0
	GIBWR.state = 0


'********* rm = reflection Mod *********
	rm.enabled = 1 '  ****Activation****
'***************************************

'********    FLippers DE Anim     ******
	logo.enabled = 1' ****Activation****
'***************************************

FlashClearAnim.enabled = 1
FlashClearAnim2.enabled = 1
FlashGreenAnim.enabled = 1


	hole=0
	seq=1
	seqb=1
	grvar=1
	gr=2
	balldropA=0
	balldropB=0
	SndRedRamp=0
	SndRedRampV=0
	SndYellowRamp=0
	SndYellowRampV=0
	SndLaunchRamp=0
	SndLaunchRampV=0
	swrot=-4
	MH = 0

	MystHole1.visible = 0
	MystHole2.visible = 0
	MystHole3.visible = 0

	HoleRampLight.intensity = 0
	HoleRampLight1.intensity = 0
	HoleRampLight2.intensity = 0
	HoleRampLight3.intensity = 0
	HoleRampLight4.intensity = 0
	GIcenter4.state = 0
	Mystv=0
	Mystnudge = 1

	testjp = 0
	cache.isdropped = 1
	cache1.isdropped = 1

	mm01.visible = 0
	mm02.visible = 0
	mm03.visible = 0
	mm04.visible = 0
	mml01.state = 0
	mml01.intensity = 0

	FlashClearEtape = 0
	FlashClearEtapeR = 0
	FlashRedEtape = 0
	FlashGreenEtape = 0

End Sub

'****************** End Table Ini ************

Sub table1_Paused:Controller.Pause = 1:End Sub
Sub table1_unPaused:Controller.Pause = 0:End Sub

Dim BISPL
BISPL = 0

' Snake Pit Plunger Trigger
Sub SnakePitPlunger_Hit
	BISPL = 1
End Sub

Sub SnakePitPlunger_Unhit
	BISPL = 0
End Sub

Sub HoleLight_Timer()
if hole=1 then LightSequence:GreenRamp
if hole=0 then Hole01.state=0:Hole02.state=0:Hole03.state=0:Hole04.state=0:Hole1r.state=0:Hole2r.state=0:Hole3r.state=0:Hole4r.state=0:seq=1
End Sub

Sub LightSequence
Select Case seq
Case 1:
Hole01.state= 1:Hole02.State=0:Hole03.state= 0:Hole04.State=0
Hole1r.state= 1:Hole2r.State=0:Hole3r.state= 0:Hole4r.State=0
seq = 2
Case 2:
Hole01.state= 0:Hole02.State=1:Hole03.state= 0:Hole04.State=0
Hole1r.state= 0:Hole2r.State=1:Hole3r.state= 0:Hole4r.State=0
seq = 3
Case 3:
Hole01.state= 0:Hole02.State=0:Hole03.state= 1:Hole04.State=0
Hole1r.state= 0:Hole2r.State=0:Hole3r.state= 1:Hole4r.State=0
seq = 4
Case 4:
Hole01.state= 0:Hole02.State=0:Hole03.state= 0:Hole04.State=1
Hole1r.state= 0:Hole2r.State=0:Hole3r.state= 0:Hole4r.State=1
seq = 5
Case 5:
Hole01.state= 0:Hole02.State=0:Hole03.state= 1:Hole04.State=0
Hole1r.state= 0:Hole2r.State=0:Hole3r.state= 1:Hole4r.State=0
seq = 6
Case 6:
Hole01.state= 0:Hole02.State=1:Hole03.state= 0:Hole04.State=0
Hole1r.state= 0:Hole2r.State=1:Hole3r.state= 0:Hole4r.State=0
seq = 7
Case 7:
Hole01.state= 1:Hole02.State=0:Hole03.state= 0:Hole04.State=0
Hole1r.state= 1:Hole2r.State=0:Hole3r.state= 0:Hole4r.State=0
seq = 8
Case 8:
Hole01.state= 0:Hole02.State=1:Hole03.state= 0:Hole04.State=0
Hole1r.state= 0:Hole2r.State=1:Hole3r.state= 0:Hole4r.State=0
seq = 3
End Select
End Sub

Sub GreenRamp
gr=gr+grvar
HoleRampLight.intensity = gr
HoleRampLight1.intensity = gr
HoleRampLight2.intensity = gr
HoleRampLight3.intensity = gr
HoleRampLight4.intensity = gr/4
if gr=>20 then grvar=-1
if gr=<2 then grvar=1
End Sub


' **************************************************************************
' ****************** Drop Targets ******************************************
' Left Drop Targets


	dim sw33Dir, sw34Dir, sw59Dir
	dim sw33Pos, sw34Pos, sw59Pos

	sw33Dir = 1:sw34Dir = 1:sw59Dir = 1
	sw33Pos = 0:sw34Pos = 0:sw59Pos = 0

  'Targets Init
	sw33a.TimerEnabled = 1:sw34a.TimerEnabled = 1:sw59a.TimerEnabled = 1


   	Set DTLBank = New cvpmDropTarget
   	  With DTLBank
   		.InitDrop Array(Array(sw33,sw33a),Array(sw34,sw34a),Array(sw59,sw59a)), Array(33,34,59)
		.InitSnd SoundFX("drop_target_down_1",DOFDropTargets),SoundFX("drop_target_reset_1",DOFContactors)
       End With


  Sub sw33_Hit:DTLBank.Hit 1:sw33Dir = 0:sw33a.TimerEnabled = 1:End Sub
  Sub sw34_Hit:DTLBank.Hit 2:sw34Dir = 0:sw34a.TimerEnabled = 1:End Sub
  Sub sw59_Hit:DTLBank.Hit 3:sw59Dir = 0:sw59a.TimerEnabled = 1:End Sub



 Sub sw33a_Timer()

  Select Case sw33Pos
        Case 0: sw33P.z=24
				If sw33Dir = 1 then
					sw33a.TimerEnabled = 0
				else
					sw33Dir = 0
					sw33a.TimerEnabled = 1
				end if
        Case 1: sw33P.z=26
        Case 2: sw33P.z=29
        Case 3: sw33P.z=26
        Case 4: sw33P.z=22
        Case 5: sw33P.z=18
        Case 6: sw33P.z=14
        Case 7: sw33P.z=10
        Case 8: sw33P.z=6
        Case 9: sw33P.z=2
        Case 10: sw33P.z=-4
        Case 11: sw33P.z=--10
        Case 12: sw33P.z=-16
        Case 13: sw33P.z=-19:sw33P.ReflectionEnabled = true
        Case 14: sw33P.z=-22:sw33P.ReflectionEnabled = false
				 If sw33Dir = 1 then
				 else
					sw33a.TimerEnabled = 0
			     end if
End Select
	If sw33Dir = 1 then
		If sw33pos>0 then sw33pos=sw33pos-1
	else
		If sw33pos<14 then sw33pos=sw33pos+1
	end if
  End Sub





 Sub sw34a_Timer()
  Select Case sw34Pos
        Case 0: sw34P.z=24
				 If sw34Dir = 1 then
					sw34a.TimerEnabled = 0
				 else
					sw34Dir = 0
					sw34a.TimerEnabled = 1
			     end if
        Case 1: sw34P.z=26
        Case 2: sw34P.z=29
        Case 3: sw34P.z=26
        Case 4: sw34P.z=22
        Case 5: sw34P.z=18
        Case 6: sw34P.z=14
        Case 7: sw34P.z=10
        Case 8: sw34P.z=6
        Case 9: sw34P.z=2
        Case 10: sw34P.z=-4
        Case 11: sw34P.z=-10
        Case 12: sw34P.z=-16
        Case 13: sw34P.z=-19:sw34P.ReflectionEnabled = true
        Case 14: sw34P.z=-22:sw34P.ReflectionEnabled = false
				 If sw34Dir = 1 then
				 else
					sw34a.TimerEnabled = 0
			     end if
End Select
	If sw34Dir = 1 then
		If sw34pos>0 then sw34pos=sw34pos-1
	else
		If sw34pos<14 then sw34pos=sw34pos+1
	end if
  End Sub


 Sub sw59a_Timer()
  Select Case sw59Pos
        Case 0: sw59P.z=24
				 If sw59Dir = 1 then
					sw59a.TimerEnabled = 0
				 else
					sw59Dir = 0
					sw59a.TimerEnabled = 1
			     end if
        Case 1: sw59P.z=26
        Case 2: sw59P.z=29
        Case 3: sw59P.z=26
        Case 4: sw59P.z=22
        Case 5: sw59P.z=18
        Case 6: sw59P.z=14
        Case 7: sw59P.z=10
        Case 8: sw59P.z=6
        Case 9: sw59P.z=2
        Case 10: sw59P.z=-4
        Case 11: sw59P.z=-10
        Case 12: sw59P.z=-16
        Case 13: sw59P.z=-19:sw59P.ReflectionEnabled = true
        Case 14: sw59P.z=-22:sw59P.ReflectionEnabled = false
				 If sw59Dir = 1 then
				 else
					sw59a.TimerEnabled = 0
			     end if
End Select
	If sw59Dir = 1 then
		If sw59pos>0 then sw59pos=sw59pos-1
	else
		If sw59pos<14 then sw59pos=sw59pos+1
	end if
  End Sub



'DT Subs
   Sub ResetDropsL(Enabled)
		If Enabled Then
			sw33Dir = 1:sw34Dir = 1:sw59Dir = 1
			sw33a.TimerEnabled = 1:sw34a.TimerEnabled = 1:sw59a.TimerEnabled = 1
			DTLBank.DropSol_On
		End if
   End Sub

' **************************************************************************
' Right Drop Targets


	dim sw36Dir, sw35Dir, sw57Dir
	dim sw36Pos, sw35Pos, sw57Pos

	sw36Dir = 1:sw35Dir = 1:sw57Dir = 1
	sw36Pos = 0:sw35Pos = 0:sw57Pos = 0

  'Targets Init
	sw36a.TimerEnabled = 1:sw35a.TimerEnabled = 1:sw57a.TimerEnabled = 1


   	Set DTRBank = New cvpmDropTarget
   	  With DTRBank
   		.InitDrop Array(Array(sw36,sw36a),Array(sw35,sw35a),Array(sw57,sw57a)), Array(36,35,57)
		.InitSnd SoundFX("drop_target_down_1",DOFDropTargets),SoundFX("drop_target_reset_1",DOFContactors)
       End With


  Sub sw36_Hit:DTRBank.Hit 1:sw36Dir = 0:sw36a.TimerEnabled = 1:End Sub
  Sub sw35_Hit:DTRBank.Hit 2:sw35Dir = 0:sw35a.TimerEnabled = 1:End Sub
  Sub sw57_Hit:DTRBank.Hit 3:sw57Dir = 0:sw57a.TimerEnabled = 1:End Sub



 Sub sw36a_Timer()

  Select Case sw36Pos
        Case 0: sw36P.z=24
				If sw36Dir = 1 then
					sw36a.TimerEnabled = 0
				else
					sw36Dir = 0
					sw36a.TimerEnabled = 1
				end if
        Case 1: sw36P.z=26
        Case 2: sw36P.z=29
        Case 3: sw36P.z=26
        Case 4: sw36P.z=22
        Case 5: sw36P.z=18
        Case 6: sw36P.z=14
        Case 7: sw36P.z=10
        Case 8: sw36P.z=6
        Case 9: sw36P.z=2
        Case 10: sw36P.z=-4
        Case 11: sw36P.z=-10
        Case 12: sw36P.z=-16
        Case 13: sw36P.z=-19:sw36P.ReflectionEnabled = true
        Case 14: sw36P.z=-22:sw36P.ReflectionEnabled = false
				 If sw36Dir = 1 then
				 else
					sw36a.TimerEnabled = 0
			     end if
End Select
	If sw36Dir = 1 then
		If sw36pos>0 then sw36pos=sw36pos-1
	else
		If sw36pos<14 then sw36pos=sw36pos+1
	end if
  End Sub





 Sub sw35a_Timer()
  Select Case sw35Pos
        Case 0: sw35P.z=24
				 If sw35Dir = 1 then
					sw35a.TimerEnabled = 0
				 else
					sw35Dir = 0
					sw35a.TimerEnabled = 1
			     end if
        Case 1: sw35P.z=26
        Case 2: sw35P.z=29
        Case 3: sw35P.z=26
        Case 4: sw35P.z=22
        Case 5: sw35P.z=18
        Case 6: sw35P.z=14
        Case 7: sw35P.z=10
        Case 8: sw35P.z=6
        Case 9: sw35P.z=2
        Case 10: sw35P.z=-4
        Case 11: sw35P.z=-10
        Case 12: sw35P.z=-16
        Case 13: sw35P.z=-19:sw35P.ReflectionEnabled = true
        Case 14: sw35P.z=-22:sw35P.ReflectionEnabled = false
				 If sw35Dir = 1 then
				 else
					sw35a.TimerEnabled = 0
			     end if
End Select
	If sw35Dir = 1 then
		If sw35pos>0 then sw35pos=sw35pos-1
	else
		If sw35pos<14 then sw35pos=sw35pos+1
	end if
  End Sub


 Sub sw57a_Timer()
  Select Case sw57Pos
        Case 0: sw57P.z=24
				 If sw57Dir = 1 then
					sw57a.TimerEnabled = 0
				 else
					sw57Dir = 0
					sw57a.TimerEnabled = 1
			     end if
        Case 1: sw57P.z=26
        Case 2: sw57P.z=29
        Case 3: sw57P.z=26
        Case 4: sw57P.z=22
        Case 5: sw57P.z=18
        Case 6: sw57P.z=14
        Case 7: sw57P.z=10
        Case 8: sw57P.z=6
        Case 9: sw57P.z=2
        Case 10: sw57P.z=-4
        Case 11: sw57P.z=-10
        Case 12: sw57P.z=-16
        Case 13: sw57P.z=-19:sw57P.ReflectionEnabled = true
        Case 14: sw57P.z=-22:sw57P.ReflectionEnabled = false
				 If sw57Dir = 1 then
				 else
					sw57a.TimerEnabled = 0
			     end if
End Select
	If sw57Dir = 1 then
		If sw57pos>0 then sw57pos=sw57pos-1
	else
		If sw57pos<14 then sw57pos=sw57pos+1
	end if
  End Sub



'DT Subs
   Sub ResetDropsR(Enabled)
		If Enabled Then
			sw36Dir = 1:sw35Dir = 1:sw57Dir = 1
			sw36a.TimerEnabled = 1:sw35a.TimerEnabled = 1:sw57a.TimerEnabled = 1
			DTRBank.DropSol_On
		End if
   End Sub

' **************************************************************************


 '**********
 ' Keys
 '**********
'*****************
' TT nudge
' ignore this -> used only for my cab
'*****************

' StopShake

Dim NudgeDirection
Dim MyNudge

Sub MyNudgeTimer_Timer()
    me.enabled = false
    if NudgeDirection = "L" Then LeftNudge 90, 2, 20:PlaySound SoundFX("fx_nudge_left",0):end if
    if NudgeDirection = "R" Then RightNudge 270, 2, 20:PlaySound SoundFX("fx_nudge_right",0):end if
End Sub

Sub table1_KeyDown(ByVal Keycode)

	If vpmKeyDown(keycode) Then Exit Sub
	If keycode = keyFront Then vpmTimer.pulsesw 8 'Buy-in Button - 2 key

	If keycode = keyInsertCoin1 or keycode = keyInsertCoin2 or keycode = keyInsertCoin3 or keycode = keyInsertCoin4 Then
		Select Case Int(rnd*3)
			Case 0: PlaySound ("Coin_In_1"), 0, CoinSoundLevel, 0, 0.25
			Case 1: PlaySound ("Coin_In_2"), 0, CoinSoundLevel, 0, 0.25
			Case 2: PlaySound ("Coin_In_3"), 0, CoinSoundLevel, 0, 0.25
		End Select
	End If

	If keycode = PlungerKey Then
		Plunger2.Pullback
		vpmTimer.PulseSw 62
		' Only show VR plunger animation when ball is in snake plunger lane.
		If BISPL = 1 Then
			TimerPlunger.Enabled = True
			TimerPlunger2.Enabled = False
		End If
	End If

	If keycode = LeftTiltKey Then LeftNudge 80, 1.2, 20
	If keycode = RightTiltKey Then RightNudge 280, 1.2, 20
	If keycode = CenterTiltKey Then CenterNudge 0, 1.6, 25

	If keycode = RightMagnaSave and EnableMagnasave = 1 then
		lutpos = lutpos + 1 : If lutpos > ubound(luts) Then lutpos = 0 : end if
		call myChangeLut
		playsound "Lut_Toggle"
	End If

	If keycode = LeftMagnaSave and EnableMagnasave = 1 then
		lutpos = lutpos - 1 : If lutpos < 0 Then lutpos = ubound(luts) : end if 
		call myChangeLut
		playsound "LUT_Toggle"
    End If

	If keycode = LeftFlipperKey Then FlipperActivate LeftFlipper, LFPress:PinCab_LeftFlipperButton.X = PinCab_LeftFlipperButton.X + 10
	If keycode = RightFlipperKey Then FlipperActivate RightFlipper, RFPress:PinCab_RightFlipperButton.X = PinCab_RightFlipperButton.X - 10

End Sub

Sub Table1_KeyUp(ByVal Keycode)
	If vpmKeyUp(keycode) Then Exit Sub

	If keycode = LeftFlipperKey Then FlipperDeActivate LeftFlipper, LFPress:PinCab_LeftFlipperButton.X = PinCab_LeftFlipperButton.X - 10
	If keycode = RightFlipperKey Then FlipperDeActivate RightFlipper, RFPress:PinCab_RightFlipperButton.X = PinCab_RightFlipperButton.X + 10

	If keycode = PlungerKey Then
		Plunger2.Fire
		PlaySoundAtVol SoundFX("plunger",DOFcontactors), Plunger2, 1
		' Only show VR plunger animation when ball is in snake plunger lane.
		If BISPL = 1 Then
			TimerPlunger.Enabled = False
			TimerPlunger2.Enabled = True
			PinCab_Primary_Plunger.Y = -250
		End If
	End If

End Sub



 '*********
 ' Switches
 '*********

 ' Slings

'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep, Hstep

Sub RightSlingShot_Slingshot
    vpmTimer.PulseSw 28
    RandomSoundSlingshotRight sling1
	slingflashR.state = 0
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
	slingflashR.state = 1
		GIBWR.color = RGB(255,0,0):GIBWR.intensity = 15:GIBWR.state = 1

if RSLing1.Visible = 0 then slingflashR.state = 0:GIBWR.color = RGB(212,212,212):GIBWR.intensity = 2:GIBWR.state = 1:end If
End Sub

Sub LeftSlingShot_Slingshot
    vpmTimer.PulseSw 29
    RandomSoundSlingshotLeft sling2
	slingflashL.state = 0
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
	slingflashL.state = 1
		GIBWL.state = 1:GIBWL.color = RGB(0,128,0):GIBWL.intensity = 15


if LSLing1.Visible = 0 then slingflashL.state = 0:GIBWL.state = 1:GIBWL.color = RGB(212,212,212):GIBWL.intensity = 2:end If
End Sub

Sub LeftSlingShotH_Slingshot
    vpmTimer.PulseSw 30
    RandomSoundSlingshotLeft slingh
    LSlingH.Visible = 0
    LSlingH1.Visible = 1
    slingH.TransZ = -20
    HStep = 0
    LeftSlingShotH.TimerEnabled = 1
End Sub

Sub LeftSlingShotH_Timer
    Select Case HStep
        Case 3:LSLingH1.Visible = 0:LSLingH2.Visible = 1:slingH.TransZ = -10
        Case 4:FlashGreenEtape = 1:LSLingH2.Visible = 0:LSLingH.Visible = 1:slingH.TransZ = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    HStep = HStep + 1
		flash6.state = 1
		GIBWR.color = RGB(255,0,0):GIBWR.intensity = 15:GIBWR.state = 1
		GIBWL.color = RGB(0,128,0):GIBWL.intensity = 15:GIBWL.state = 1
		flashvertref.amount = 70
		flashvertref.intensityScale = 1
		if LSLingH1.Visible = 0 then GIBWL.color = RGB(212,212,212):GIBWL.intensity = 2:GIBWL.state = 1:GIBWR.color = RGB(212,212,212):GIBWR.intensity = 2:GIBWR.state = 1:flash6.state = 0:flashvertref.amount = 0:flashvertref.intensityScale = 0:end If


End Sub

'**********************************************************
'************************ Bumpers *************************
'**********************************************************

Sub Bumper1_Hit
    vpmTimer.PulseSw 25
	RandomSoundBumperTop Bumper1
	Me.TimerEnabled = 1
	Flasher101.state = 1
		GIBWL.color = RGB(0,128,0):GIBWL.intensity = 15:GIBWL.state = 1
	BackWall.image = "backwallon4Bump"
End Sub

Sub Bumper1_Timer
	Me.Timerenabled = 0
	Flasher101.state = 0
		GIBWL.color = RGB(212,212,212):GIBWL.intensity = 2:GIBWL.state = 1
	BackWall.image = "backwallon"
End Sub

Sub Bumper2_Hit
    vpmTimer.PulseSw 26
	RandomSoundBumperMiddle Bumper2
	Me.TimerEnabled = 1
	Flasher102.state = 1
		GIBWL.color = RGB(0,128,0):GIBWL.intensity = 15:GIBWL.state = 1
		GIBWR.color = RGB(255,0,0):GIBWR.intensity = 15:GIBWR.state = 1
	BackWall.image = "backwallon4Bump"
End Sub

Sub Bumper2_Timer
	Me.Timerenabled = 0
	Flasher102.state = 0
		GIBWL.color = RGB(212,212,212):GIBWL.intensity = 2:GIBWL.state = 1
		GIBWR.color = RGB(212,212,212):GIBWR.intensity = 2:GIBWR.state = 1
	BackWall.image = "backwallon"
End Sub

Sub Bumper3_Hit
    vpmTimer.PulseSw 27
	RandomSoundBumperBottom Bumper3
	Me.TimerEnabled = 1
	Flasher103.state = 1
		GIBWR.color = RGB(255,0,0):GIBWR.intensity = 15:GIBWR.state = 1
	BackWall.image = "backwallon4Bump"
End Sub

Sub Bumper3_Timer
	Me.Timerenabled = 0
	Flasher103.state = 0
	BackWall.image = "backwallon"
	GIBWR.color = RGB(212,212,212):GIBWR.intensity = 2:GIBWR.state = 1
End Sub

sub BumperSounds()
	Select Case Int(Rnd*2)
		Case 0 : PlaySoundAtVol SoundFX("Bumpers_Top_1",DOFContactors), ActiveBall, 1
		Case 1 : PlaySoundAtVol SoundFX("Bumpers_Top_2",DOFContactors), ActiveBall, 1
	End Select
End Sub


' Rollovers
 Sub sw16_Hit():Controller.Switch(16) = 1:PlaySoundAtVol "", ActiveBall, 1:End Sub
 Sub sw16_UnHit:Controller.Switch(16) = 0:End Sub

 Sub sw53_Hit():Controller.Switch(53) = 1:PlaySoundAtVol "", ActiveBall, 1:End Sub
 Sub sw53_UnHit:Controller.Switch(53) = 0:End Sub

 Sub sw54_Hit():Controller.Switch(54) = 1:PlaySoundAtVol "", ActiveBall, 1:End Sub
 Sub sw54_UnHit:Controller.Switch(54) = 0:End Sub

 Sub sw55_Hit():Controller.Switch(55) = 1:PlaySoundAtVol "", ActiveBall, 1:End Sub
 Sub sw55_UnHit:Controller.Switch(55) = 0:End Sub

 Sub sw56_Hit():Controller.Switch(56) = 1:PlaySoundAtVol "", ActiveBall, 1:End Sub
 Sub sw56_UnHit:Controller.Switch(56) = 0:End Sub

 Sub sw58_Hit():Controller.Switch(58) = 1:PlaySoundAtVol "", ActiveBall, 1:End Sub
 Sub sw58_UnHit:Controller.Switch(58) = 0:End Sub

 Sub sw60_Hit():Controller.Switch(60) = 1:PlaySoundAtVol "", ActiveBall, 1:End Sub
 Sub sw60_UnHit:Controller.Switch(60) = 0:End Sub

 Sub sw24_Hit():Controller.Switch(24) = 1:PlaySoundAtVol "", ActiveBall, 1:hole=1:HoleLight.enabled = 1:End Sub
 Sub sw24_UnHit:Controller.Switch(24) = 0:End Sub

 Sub sw21_Hit():Controller.Switch(21) = 1:PlaySoundAtVol "", ActiveBall, 1:End Sub
 Sub sw21_UnHit:Controller.Switch(21) = 0:End Sub

 Sub sw22_Hit():Controller.Switch(22) = 1:PlaySoundAtVol "", ActiveBall, 1:End Sub
 Sub sw22_UnHit:Controller.Switch(22) = 0:End Sub

 Sub sw23_Hit():Controller.Switch(23) = 1:PlaySoundAtVol "", ActiveBall, 1:End Sub
 Sub sw23_UnHit:Controller.Switch(23) = 0:End Sub

 Sub sw48_Hit():Controller.Switch(48) = 1:PlaySoundAtVol "", ActiveBall, 1:End Sub
 Sub sw48_UnHit:Controller.Switch(48) = 0:End Sub

 ' Ramp Switches
 Sub sw49_Hit():Controller.Switch(49) = 1:PlaySoundAtVol "gate", ActiveBall, 1:sndredrampV=1:End Sub
 Sub sw49_UnHit:Controller.Switch(49) = 0:End Sub


' Switch metalic ramp red JPJ

Sub sw50_Hit():Controller.Switch(50) = 1:PlaySoundAtVol "sensor", ActiveBall, 1:seqb=1:RedMetalAnim.enabled = 1: 'anim
Mystv = Mystv + 1
if Mystv=11 then resetMyst
if Myst=1 then
	if Mystv=1 or Mystv=1 or Mystv=3 or Mystv=5 or Mystv=7 or Mystv=9 then playsound "biere":capchoice
end if
if Myst=1 Then
	if Mystv=2 or Mystv=4 or Mystv=6 or Mystv=8 or Mystv=10 then playsound "pchit":capchoice
end if
if Mystv=11 then resetMyst

End Sub

Sub sw50_UnHit:Controller.Switch(50) = 0:End Sub

Sub RedMetalAnim_timer
Select Case seqb
Case 1:seqb = 2:
swRedMetal.rotz = -87
Case 2:seqb = 3:
swRedMetal.rotz = -95
Case 3:seqb = 4:
swRedMetal.rotz = -101
Case 4:seqb = 5:
swRedMetal.rotz = -97
Case 5:seqb = 6:
swRedMetal.rotz = -91
Case 6:
swRedMetal.rotz = -85:RedMetalAnim.enabled = 0
End Select
'RedMetalAnim.enabled = 0
End Sub

'Sub Anim
'Select Case seqb
'Case 1:seqb = 2:
'swRedMetal.rotz = -87
'Case 2:seqb = 3:
'swRedMetal.rotz = -95
'Case 3:seqb = 4:
'swRedMetal.rotz = -101
'Case 4:seqb = 5:
'swRedMetal.rotz = -97
'Case 5:seqb = 6:
'swRedMetal.rotz = -91
'Case 6:
'swRedMetal.rotz = -85
'End Select
'End Sub



 Sub sw51_Hit():Controller.Switch(51) = 1:PlaySoundAtVol "gate", ActiveBall, 1:SndYellowRampV=1:Controller.Switch(52) = 0::End Sub
 Sub sw51_UnHit:Controller.Switch(51) = 0:End Sub

 Sub sw52_Hit():Controller.Switch(52) = 1 'PlaySound "gate" enlevé car plus mis au même endroit
	End Sub
' Sub sw52_UnHit:debug.print "sorti du switch 52":Controller.Switch(52) = 0:
'	End Sub
Sub gate5_hit:Controller.Switch(52) = 1:	End Sub


 Sub sw40_Hit():Controller.Switch(40) = 1:PlaySoundAtVol "sensor", ActiveBall, 1:balldropB=1:End Sub
 Sub sw40_UnHit:Controller.Switch(40) = 0:End Sub

 ' Targets

Sub SW17_hit():sw17p.transx = -10:Me.TimerEnabled = 1:vpmTimer.PulseSw 17:End Sub
Sub SW17_Timer():sw17p.transx = 0:Me.TimerEnabled = 0:End Sub

Sub SW18_hit():sw18p.transx = -10:Me.TimerEnabled = 1:vpmTimer.PulseSw 18:End Sub
Sub SW18_Timer():sw18p.transx = 0:Me.TimerEnabled = 0:End Sub

Sub SW19_hit():sw19p.transx = -10:Me.TimerEnabled = 1:vpmTimer.PulseSw 19:End Sub
Sub SW19_Timer():sw19p.transx = 0:Me.TimerEnabled = 0:End Sub

Sub SW20_hit():sw20p.transx = -10:Me.TimerEnabled = 1:vpmTimer.PulseSw 20:End Sub
Sub SW20_Timer():sw20p.transx = 0:Me.TimerEnabled = 0:End Sub







 ' Drain & holes
 Sub Drain_Hit:RandomSoundDrain Drain:ClearBallID:bsTrough.AddBall Me:End Sub
 Sub sw37_Hit:playSoundAtVol "Saucer_Enter_1", sw37, 1:
	'ClearBallID:
	'sw37.DestroyBall:
	bsEject.AddBall 0
End Sub

Sub sw39_Hit:playSoundAtVol "Saucer_Enter_2", sw39, 1:
	cache1.isdropped = 0
	 ClearBallID
bsVuk.AddBall Me
sw39a.enabled=1
sw39.enabled=0
End Sub
sub sw39a_unhit
sw39a.enabled=0
sw39.enabled = 1
cache1.isdropped = 1
end Sub

  ' Center Scoop
 Dim aBall, aZpos

 Sub sw38_Hit
		cache.isdropped = 0
		MH=1
		 MystAnimTimer.Enabled=1
	If BallInSw38 = 0 Then
		 me.enabled = 0
		 BallInSw38 = 1
		 playSoundAtVol("Saucer_Enter_1"), sw38, 1
		 Set aBall = ActiveBall
		 aZpos = 35
		 ClearBallID
		 sw38.Destroyball
		 bsCenterScoop.AddBall Me
		Me.TimerInterval = 10000 '2000
		Me.TimerEnabled = 1
		 sw38a.Enabled = 1
	End If
 End Sub


 Sub sw38_Timer
      Me.TimerEnabled = 0
	  BallInSw38 = 0
	  me.enabled = 1
 End Sub


Sub sw38a_Unhit

sw38a.Enabled = 0
MH=0
MystAnimTimer.Enabled=0
MH=0:MystHole1.visible = 0:MystHole2.visible = 0:MystHole3.visible = 0:BallInSw38 = 0

sw38.Enabled = 1
		cache.isdropped = 1
End Sub



'************************************
'**** Mystery Ball Animation JPJ ****
'************************************

Sub MystAnimTimer_Timer
MystAnimation
End Sub


Sub MystAnimation
Select Case MH
Case 1:MH = 2:
MystHole1.visible = 1:MystHole2.visible = 0:MystHole3.visible = 0
Case 2:MH = 3:
MystHole1.visible = 0:MystHole2.visible = 1:MystHole3.visible = 0
Case 3:MH = 4:
MystHole1.visible = 0:MystHole2.visible = 0:MystHole3.visible = 1
Case 4:MH = 5:
MystHole1.visible = 0:MystHole2.visible = 1:MystHole3.visible = 0
Case 5:MH = 6:
MystHole1.visible = 1:MystHole2.visible = 0:MystHole3.visible = 0
Case 6:MH=1:
MystHole1.visible = 0:MystHole2.visible = 1:MystHole3.visible = 0
End Select
End Sub



 ' Ramp Helpers
  Sub RHelp1_Hit:ClearBallID:RHelp1.DestroyBall:vpmCreateBall RHelp1a:RHelp1a.kick 0, 0:PlaySoundAtVol "Ball_Bounce_Playfield_Soft_1", ActiveBall, 1:End Sub
  Sub RHelp2_Hit:ClearBallID:RHelp2.DestroyBall:vpmCreateBall RHelp2a:RHelp2a.kick 0, 0:PlaySoundAtVol "Ball_Bounce_Playfield_Soft_2", ActiveBall, 1:End Sub


Sub Trigger1_Hit:sndLaunchrampV=1:End Sub
Sub Trigger2_Hit:playsoundAtVol "rail", ActiveBall, 1:End Sub
Sub TriggerTest_Hit:testJP=1:End Sub
Sub TriggerTest2_Hit:testJP=0:End Sub
Sub Trigger9_Hit
if balldropA=1 then balldropBig():balldropA=0
End Sub
Sub Trigger10_Hit
if balldropB = 1 then balldropSides()
hole=0:HoleLight.enabled = 0:Hole01.state=0:Hole02.state=0:Hole03.state=0:Hole04.state=0:Hole1r.state=0:Hole2r.state=0:Hole3r.state=0:Hole4r.state=0:HoleRampLight.intensity = 0:HoleRampLight1.intensity = 0:HoleRampLight2.intensity = 0:HoleRampLight3.intensity = 0:balldropB=0
End Sub
Sub Trigger4_Hit:playsoundAtVol "WirerampLeft", ActiveBall, 1
Mystv = Mystv + 1
if Myst=1 then
	if Mystv=1 or Mystv=1 or Mystv=3 or Mystv=5 or Mystv=7 or Mystv=9 then playsound "biere":capchoice
end if
if Myst=1 Then
	if Mystv=2 or Mystv=4 or Mystv=6 or Mystv=8 or Mystv=10 then playsound "pchit":capchoice
end if
if Mystv=11 then resetMyst
End Sub
Sub Trigger5_Hit:playsoundAtVol "Wireramp2jpmod", ActiveBall, 1:End Sub
Sub Trigger6_Hit:balldropSides():stopsound "WirerampLeft":playsoundAtVol "WirerampStop", ActiveBall, 1:End Sub
Sub Trigger7_Hit:balldropSides():stopsound "Wireramp2jpmod":playsoundAtVol "WirerampStop", ActiveBall, 1:End Sub
Sub Trigger8_Hit:stopsound "LaunchWire":balldropA=1:sndLaunchramp=0:sndLaunchrampV=0:End Sub

Sub Trigger11RedRamp_Hit:
	if sndredramp=0 and sndredrampV=0 and myst = 0 then playsound "plasticramp", 2, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall):SndRedRamp=1
	if sndredramp=1 and sndredrampV=1 and myst = 0 then stopsound "plasticramp":SndRedRamp=0:SndRedRampV=0
	if sndredramp=0 and sndredrampV=0 and myst = 1 then playsound "mystramp", 2, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall):SndRedRamp=1
	if sndredramp=1 and sndredrampV=1 and myst = 1 then stopsound "mystramp":SndRedRamp=0:SndRedRampV=0

End Sub

Sub Trigger12RedRampOut_Hit:stopsound "plasticramp":stopsound "mystramp":SndRedRamp=0:SndRedRampV=0:End Sub

Sub Trigger13YellowRamp_Hit:
	If ActiveBall.VelX<-11 and ActiveBall.VelY<-29 then ActiveBall.VelX=-11:ActiveBall.Vely=-29:End If
	'If ActiveBall.VelY<-29 then ActiveBall.Vely=-29:End If
	if sndyellowramp=0 and sndyellowrampV=0 and myst = 0 then playsound "plasticramp2", 2, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall):SndYellowRamp=1
	if sndyellowramp=1 and sndyellowrampV=1 and myst = 0 then stopsound "plasticramp2":SndyellowRamp=0:SndYellowRampV=0
	if sndyellowramp=0 and sndyellowrampV=0 and myst = 1 then playsound "mystramp", 2, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall):SndYellowRamp=1
	if sndyellowramp=1 and sndyellowrampV=1 and myst = 1 then stopsound "mystramp":SndyellowRamp=0:SndYellowRampV=0
End Sub

Sub Trigger14YellowRampOut_Hit:stopsound "plasticramp2":stopsound "mystramp":SndYellowRamp=0:SndYellowRampV=0:Controller.Switch(52) = 0:End Sub

Sub Trigger15LaunchRamp_Hit:
	if sndLaunchramp=0 and sndLaunchrampV=0 then playsound "LaunchWire", 3, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall):SndLaunchRamp=1
	if sndLaunchramp=1 and sndLaunchrampV=1 then stopsound "LaunchWire":SndLaunchRamp=0:SndLaunchRampV=0
End Sub

sub balldropSides()
	Select Case Int(Rnd*2)
		Case 0 : PlaySoundAtBallVol "Ball_Bounce_Playfield_Soft_1", 1
		Case 1 : PlaySoundAtBallVol "Ball_Bounce_Playfield_Soft_2", 1
	End Select
End Sub

 sub balldropBig()
	Select Case Int(Rnd*3)
		Case 0 : PlaySoundAtBallVol "Ball_Bounce_Playfield_Hard_1", 1
		Case 1 : PlaySoundAtBallVol "Ball_Bounce_Playfield_Hard_2", 1
		Case 2 : PlaySoundAtBallVol "Ball_Bounce_Playfield_Hard_3", 1
	End Select
End Sub

 'FLIPPERS *************************************************************************************************************************************************

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Const ReflipAngle = 20
Sub SolLFlipper(Enabled)
        If Enabled Then
		LF.Fire
		LeftFlipper1.RotateToEnd
		'LF3.Fire 
        
                If leftflipper.currentangle < leftflipper.endangle + ReflipAngle Then 
                        RandomSoundReflipUpLeft LeftFlipper
                Else 
                        SoundFlipperUpAttackLeft LeftFlipper
                        RandomSoundFlipperUpLeft LeftFlipper
                End If                
        Else
                LeftFlipper.RotateToStart
				LeftFlipper1.RotateToStart  'voir ATTENTION
				'LeftFlipper3.RotateToStart
                If LeftFlipper.currentangle < LeftFlipper.startAngle - 5 Then
                        RandomSoundFlipperDownLeft LeftFlipper
                End If
                FlipperLeftHitParm = FlipperUpSoundLevel
    End If
End Sub

Sub SolRFlipper(Enabled)
        If Enabled Then
                RF.Fire
				'RF2.Fire
				'RF3.Fire

                If rightflipper.currentangle > rightflipper.endangle - ReflipAngle Then
                        RandomSoundReflipUpRight RightFlipper
                Else 
                        SoundFlipperUpAttackRight RightFlipper
                        RandomSoundFlipperUpRight RightFlipper
                End If
        Else
                RightFlipper.RotateToStart
				'RightFlipper2.RotateToStart
				'RightFlipper3.RotateToStart
                If RightFlipper.currentangle > RightFlipper.startAngle + 5 Then
                        RandomSoundFlipperDownRight RightFlipper
                End If        
                FlipperRightHitParm = FlipperUpSoundLevel
        End If
End Sub

Sub LeftFlipper_Collide(parm)
		CheckLiveCatch Activeball, LeftFlipper, LFCount, parm
        LeftFlipperCollide parm
	if RubberizerEnabled = 1 then Rubberizer(parm)
	if RubberizerEnabled = 2 then Rubberizer2(parm)

End Sub

Sub RightFlipper_Collide(parm)
		CheckLiveCatch Activeball, RightFlipper, RFCount, parm
        RightFlipperCollide parm
	if RubberizerEnabled = 1 then Rubberizer(parm)
	if RubberizerEnabled = 2 then Rubberizer2(parm)

End Sub

'***********************************

sub cache_hit
	PlaySoundAtBallVol("fx_collide"),1
end sub
sub cache1_hit
	PlaySoundAtBallVol("fx_collide"), 1
end sub

 '********************
 'Solenoids & Flashers
 '********************

 SolCallback(1) = "SolRelease"
 SolCallback(2) = "bsBallRelease.SolOut"
 SolCallback(3) = "SolAutoLaunch"
 SolCallback(4) = "bsEject.SolOut"
 SolCallback(5) = "bsVuk.SolOut"
 SolCallback(6) = "bsCenterScoop.SolOut"
 SolCallback(7) = "SolTrapDoor"
 SolCallback(8) = "vpmSolSound SoundFX(""knocker"",DOFKnocker),"
 SolCallback(9) = "ResetDropsR"  'old "dtRBank.SolDropUp"
 SolCallback(11) = "SolGi"
 SolCallback(12) = "ResetDropsL" ' old "dtLBank.SolDropUp"
 SolCallback(14) = "SolKickBack" ' "vpmSolAutoPlunger KickBack,10,"
' SolCallback(23) = "FastFlips.TiltSol" ' Fastflips
 SolCallback(25) = "SolFlash25" ' Left Bumper
 SolCallback(26) = "SolFlash26" ' Mid Bumper
 SolCallback(27) = "SolFlash27" ' Right Bumper
 SolCallback(28) = "SolFlash28" ' Left Slingshot
 SolCallback(29) = "SolFlash29" ' Right Slingshot
 SolCallback(30) = "SolFlash30" ' Top Slingshot In Duff zone
 SolCallback(31) = "SolFlash31" ' Not Use ???
 SolCallback(32) = "SolFlash32" ' Not Use ???
' SolCallback(sLRFlipper) = "SolRFlipper"
' SolCallback(sLLFlipper) = "SolLFlipper"
'
'Sub SolLFlipper(Enabled)
'		 If Enabled Then
'			 PlaySoundAtVol SoundFX("fx_FlipperUp",DOFFlippers), LeftFlipper, 1:LeftFlipper.RotateToEnd:LeftFlipper1.RotateToEnd
'       PlaySoundAtVol "fx_FlipperUp", LeftFlipper1, 1
'	else
'			 PlaySoundAtVol SoundFX("fx_FlipperDown",DOFFlippers), LeftFlipper, 1:LeftFlipper.RotateToStart:LeftFlipper1.RotateToStart
'       PlaySoundAtVol "fx_FlipperDown", LeftFlipper1, 1
'		 End If
'	  End Sub
'
'	Sub SolRFlipper(Enabled)
'		 If Enabled Then
'			 PlaySoundAtVol SoundFX("fx_FlipperUp",DOFFlippers), RightFlipper, 1:RightFlipper.RotateToEnd
'		 Else
'			 PlaySoundAtVol SoundFX("fx_FlipperDown",DOFFlippers), RightFlipper, 1:RightFlipper.RotateToStart
'		 End If
'	End Sub


Sub SolAutoLaunch (enabled)
	If enabled Then Plunger.fire:'PlaysoundAtVol SoundFX ("plunger",DOFContactors), Plunger, 1
End Sub

Sub SolKickBack(enabled)
    If enabled Then
       Kickback.Fire
       PlaysoundAtVol SoundFX ("plunger",DOFContactors), Kickback, 1
    Else
       Kickback.PullBack
    End If
End Sub


Sub SolRelease(Enabled)
   If Enabled And TestJP=0 and bsBallrelease <> 1 And bsTrough.Balls > 0 Then 'pb with multiball when ballrelease was allready waiting with a ball
		bsTrough.ExitSol_On
        vpmCreateBall BallRelease
	bsBallRelease.AddBall 0
    End If
End Sub

 Sub SolTrapDoor(Enabled)
     If Enabled Then
         TrapDoor.Enabled = 1
		Hole=1
         SnakeTrapDoor.IsDropped = 1
         SP.Material = "SPOn"
         TrapDoorOpen.IsDropped = 0
     Else
         TrapDoor.Enabled = 0
		Hole=0
         SnakeTrapDoor.IsDropped = 0
         SP.Material = "SPOff"
         TrapDoorOpen.IsDropped = 1
     End If
 End Sub

 Sub TrapDoor_Hit
	ClearBallID
     TrapDoor.DestroyBall
     PlaySoundAtVol "Ballhit", TrapDoorA, 1
	stopsound "plasticramp2":SndYellowRamp=0:SndYellowRampV=0
     vpmCreateBall TrapDoorA
     TrapDoorA.kick 0, 2
 End Sub

 Sub SolFlash25(Enabled)
     If Enabled Then
		flash5.state = 1
        flash5i.state = 1
		flashrougeref.amount = 70
		flashrougeref.intensityScale = 1
		flash6.state = 1
        flash6i.state = 1
		FlashGreenEtape = 1
		flashvertref.amount = 70
		flashvertref.intensityScale = 1
		FlashRedEtape = 1
'		GIBWL.color = RGB(0,128,0):GIBWL.intensity = 10:
     Else
		flash5.state = 0
        flash5i.state = 1
		flashrougeref.amount = 0
		flashrougeref.intensityScale = 0
		flash6.state = 0
        flash6i.state = 0
		flashvertref.amount = 0
		flashvertref.intensityScale = 0
'		GIBWL.color = RGB(212,212,212):GIBWL.intensity = 2
     end If
 End Sub

 Sub SolFlash26(Enabled)
     If Enabled Then
          f26s.state = 1
		flash2.state = 1
		flash4.state = 1
		flash4a.Amount = 50
		flash4a.IntensityScale = 1
		flash4b.Amount = 70
		flash4b.IntensityScale = 1
		FlashClearEtapeR = 1
'		GIBWL.color = RGB(0,128,0):GIBWL.intensity = 10:
'		GIBWR.color = RGB(255,0,0):GIBWR.intensity = 10
     Else
          f26s.state = 0
		flash2.state = 0
		flash4.state = 0
		flash4a.Amount = 0
		flash4a.IntensityScale = 0
		flash4b.Amount = 0
		flash4b.IntensityScale = 0
'		GIBWL.state = 1:GIBWL.color = RGB(212,212,212):GIBWL.intensity = 2
'		GIBWR.state = 1:GIBWR.color = RGB(212,212,212):GIBWR.intensity = 2
     End If
 End Sub

 Sub SolFlash27(Enabled)
     If Enabled Then
		flash1.state = 1
		flash1a.Amount = 50
		flash1a.IntensityScale = 1
		flash1b.Amount = 70
		flash1b.IntensityScale = 1
		flash1c.Amount = 30
		flash1c.IntensityScale = 1
		GIBWR.color = RGB(255,0,0):GIBWR.intensity = 15:GIBWR.state = 1
		FlashClearEtape = 1
     Else
		flash1.state = 0
		flash1a.Amount = 0
		flash1a.IntensityScale = 0
		flash1b.Amount = 0
		flash1b.IntensityScale = 0
		flash1c.Amount = 0
		flash1c.IntensityScale = 0
		GIBWR.color = RGB(212,212,212):GIBWR.intensity = 2:GIBWR.state = 1
     End If

 End Sub




 Sub SolFlash28(Enabled)
     If Enabled Then
         f28a.State = 1
         f28b.State = 1
		f26S.state = 1
		flash1.state = 1
		flash1a.Amount = 50
		flash1a.IntensityScale = 1
		flash1b.Amount = 70
		flash1b.IntensityScale = 1
		flash1c.Amount = 30
		flash1c.IntensityScale = 1
		flash3.state = 1
		FlashClearEtape = 1

     Else
         f28a.State = 0
         f28b.State = 0
		f26S.state = 0
		flash1.state = 0
		flash1a.Amount = 0
		flash1a.IntensityScale = 0
		flash1b.Amount = 0
		flash1b.IntensityScale = 0
		flash1c.Amount = 0
		flash1c.IntensityScale = 0
		flash3.state = 0
     End If
 End Sub



 Sub SolFlash29(Enabled)
     If Enabled Then
         JackpotRojo.State = 1
		'flash2.state = 1
		flash4.state = 1
		flash4a.Amount = 50
		flash4a.IntensityScale = 1
		flash4b.Amount = 70
		flash4b.IntensityScale = 1
		FlashClearEtapeR = 1
     Else
         JackpotRojo.State = 0
		'flash2.state = 0
		flash4.state = 0
		flash4a.Amount = 0
		flash4a.IntensityScale = 0
		flash4b.Amount = 0
		flash4b.IntensityScale = 0
     End If
 End Sub

 Sub SolFlash30(Enabled)
     If Enabled Then
         JackpotAm.State = 1
     Else
         JackpotAm.State = 0
     End If
 End Sub

  Sub SolFlash31(Enabled)
     If Enabled Then
     f7.state = 0
     f7a.state = 0
     Else
     f7.state = 1
     f7a.state = 1
     End If
 End Sub

  Sub SolFlash32(Enabled)
     If Enabled Then
		F8.state = 1
		F8a.state = 1
     Else
		F8.state = 0
		F8a.state = 0
     End If
 End Sub

 '*****
 ' GI
 '*****

 Sub SolGi(Enabled)
     If Enabled Then
         GiOff
     Else
         GiOn
     End If
 End Sub

 Sub GiOn
     For Each x in GILights:
		x.State = 1
		if l62.state = 1 Then
			RefLaunchBall.Amount = 10
			RefLaunchBall.IntensityScale = 1
		else
			RefLaunchBall.Amount = 0
			RefLaunchBall.IntensityScale = 0
		end If
	Next
     For Each x in LightAmbiant:
		x.State = 1
	Next
	BackWall.Image = "BackWallOn"
	Ombreall.Image= "gunsombreon"
'	Flash5i.intensity = 0.75:Flash6i.intensity = 1
	flash2veille.state = 1 :flash2veille1.state = 1
	flash2veille.intensity = 0.75 :flash2veille1.intensity = 0.75
	GIBWL.color = RGB(212,212,212):GIBWL.intensity = 2:GIBWL.state = 1:
	GIBWR.color = RGB(212,212,212):GIBWR.intensity = 2:GIBWR.state = 1
	Pincab_Backglass.blenddisablelighting = 1
 End Sub

 Sub GiOff
     For Each x in GILights
		x.State = 0
		if l62.state = 1 Then
			RefLaunchBall.Amount = 10
			RefLaunchBall.IntensityScale = 1
		else
			l62z.state = 0
			RefLaunchBall.Amount = 0
			RefLaunchBall.IntensityScale = 0
		end If
	Next
     For Each x in LightAmbiant:
		x.State = 0
	Next
	if FlashClearEtape = 0 then
		flash2veille1.state = 0
	end If
	if FlashClearEtapeR = 0 Then
		flash2veille.state = 0
	end If
	if FlashGreenEtape = 0 then
		flash6i.intensity = 0.3
	end If
	if FlashRedEtape = 0 Then
		Flash5i.intensity = 0.3
	end if
	BackWall.Image = "BackWallOn"
	Ombreall.Image = "gunsombreoff"
	flash2veille.intensity = 0 :flash2veille1.intensity = 0
	GIBWL.color = RGB(212,212,212):GIBWL.intensity = 1:GIBWL.state = 1
	GIBWR.color = RGB(212,212,212):GIBWR.intensity = 1:GIBWR.state = 1
	Pincab_Backglass.blenddisablelighting = 0.20
End Sub

 '*************************************
 '          Nudge System
 ' based on Noah's nudgetest table
 '*************************************

 Dim LeftNudgeEffect, RightNudgeEffect, NudgeEffect

 Sub LeftNudge(angle, strength, delay)
     vpmNudge.DoNudge angle, (strength * (delay-LeftNudgeEffect) / delay) + RightNudgeEffect / delay
     LeftNudgeEffect = delay
     RightNudgeEffect = 0
     RightNudgeTimer.Enabled = 0
     LeftNudgeTimer.Interval = delay
     LeftNudgeTimer.Enabled = 1
 End Sub

Cap(1) = "cap01"
Cap(2) = "cap02"
Cap(3) = "cap03"
Cap(4) = "cap04"
Cap(5) = "cap05"
Cap(6) = "cap06"
Cap(7) = "cap07"
Cap(8) = "cap08"
Cap(9) = "cap09"
Cap(10) = "cap10"
Cap(11) = "cap11"
Cap(12) = "cap12"
Cap(13) = "cap13"
Cap(14) = "cap14"
Cap(15) = "cap15"
Cap(16) = "cap16"
Cap(17) = "cap17"


Sub capchoice
	Select Case Int(Rnd*17)+1
		Case 1 : numcap = 1:Table1.BallFrontDecal = Cap(numcap)
		Case 2 : numcap = 2:Table1.BallFrontDecal = Cap(numcap)
		Case 3 : numcap = 3:Table1.BallFrontDecal = Cap(numcap)
		Case 4 : numcap = 4:Table1.BallFrontDecal = Cap(numcap)
		Case 5 : numcap = 5:Table1.BallFrontDecal = Cap(numcap)
		Case 6 : numcap = 6:Table1.BallFrontDecal = Cap(numcap)
		Case 7 : numcap = 7:Table1.BallFrontDecal = Cap(numcap)
		Case 8 : numcap = 8:Table1.BallFrontDecal = Cap(numcap)
		Case 9 : numcap = 9:Table1.BallFrontDecal = Cap(numcap)
		Case 10 : numcap = 10:Table1.BallFrontDecal = Cap(numcap)
		Case 11 : numcap = 11:Table1.BallFrontDecal = Cap(numcap)
		Case 12 : numcap = 12:Table1.BallFrontDecal = Cap(numcap)
		Case 13 : numcap = 13:Table1.BallFrontDecal = Cap(numcap)
		Case 14 : numcap = 14:Table1.BallFrontDecal = Cap(numcap)
		Case 15 : numcap = 15:Table1.BallFrontDecal = Cap(numcap)
		Case 16 : numcap = 16:Table1.BallFrontDecal = Cap(numcap)
		Case 17 : numcap = 17:Table1.BallFrontDecal = Cap(numcap)
	End Select
End Sub



 Sub RightNudge(angle, strength, delay)
     vpmNudge.DoNudge angle, (strength * (delay-RightNudgeEffect) / delay) + LeftNudgeEffect / delay
     RightNudgeEffect = delay
     LeftNudgeEffect = 0
     LeftNudgeTimer.Enabled = 0
     RightNudgeTimer.Interval = delay
     RightNudgeTimer.Enabled = 1
 End Sub

 Sub CenterNudge(angle, strength, delay)
     vpmNudge.DoNudge angle, strength * (delay-NudgeEffect) / delay
     NudgeEffect = delay
     CenterNudgeTimer.Interval = delay
     CenterNudgeTimer.Enabled = 1
 End Sub

 Sub LeftNudgeTimer_Timer()
     LeftNudgeEffect = LeftNudgeEffect-1
     If LeftNudgeEffect = 0 then LeftNudgeTimer.Enabled = 0

 End Sub

 Sub RightNudgeTimer_Timer()
     RightNudgeEffect = RightNudgeEffect-1
     If RightNudgeEffect = 0 then RightNudgeTimer.Enabled = 0
 End Sub

 Sub CenterNudgeTimer_Timer()
     NudgeEffect = NudgeEffect-1
     If NudgeEffect = 0 then CenterNudgeTimer.Enabled = 0
 End Sub

' *******************************************************************************************************
' Positional Sound Playback Functions by DJRobX and Rothbauerw
' PlaySound sound, 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
' *******************************************************************************************************

' Play a sound, depending on the X,Y position of the table element (especially cool for surround speaker setups, otherwise stereo panning only)
' parameters (defaults): loopcount (1), volume (1), randompitch (0), pitch (0), useexisting (0), restart (1))
' Note that this will not work (currently) for walls/slingshots as these do not feature a simple, single X,Y position

Sub PlayXYSound(soundname, tableobj, loopcount, volume, randompitch, pitch, useexisting, restart)
  PlaySound soundname, loopcount, volume, AudioPan(tableobj), randompitch, pitch, useexisting, restart, AudioFade(tableobj)
End Sub

' Set position as table object (Use object or light but NOT wall) and Vol to 1

'Sub PlaySoundAt(soundname, tableobj)
'  PlaySound soundname, 1, 1, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
'End Sub

'Set all as per ball position & speed.

'Sub PlaySoundAtBall(soundname)
'  PlaySoundAt soundname, ActiveBall
'End Sub

'Set position as table object and Vol manually.

'Sub PlaySoundAtVol(sound, tableobj, Volume)
'  PlaySound sound, 1, Volume, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
'End Sub

'Set all as per ball position & speed, but Vol Multiplier may be used eg; PlaySoundAtBallVol "sound",3

'Sub PlaySoundAtBallVol(sound, VolMult)
'  PlaySound sound, 0, Vol(ActiveBall) * VolMult, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
'End Sub

'Set position as bumperX and Vol manually.

Sub PlaySoundAtBumperVol(sound, tableobj, Vol)
  PlaySound sound, 1, Vol, AudioPan(tableobj), 0,0,1, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBOTBallZ(sound, BOT)
    PlaySound sound, 0, ABS(BOT.velz)/17, Pan(BOT), 0, Pitch(BOT), 1, 0, AudioFade(BOT)
End Sub

' play a looping sound at a location with volume
Sub PlayLoopSoundAtVol(sound, tableobj, Vol)
	PlaySound sound, -1, Vol, AudioPan(tableobj), 0, 0, 1, 0, AudioFade(tableobj)
End Sub

'*********************************************************************
'                     Supporting Ball & Sound Functions
'*********************************************************************

'Function RndNum(min, max)
'    RndNum = Int(Rnd() * (max-min + 1) ) + min ' Sets a random number between min and max
'End Function

'Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "table1" is the name of the table
'  Dim tmp
'  On Error Resume Next
'  tmp = tableobj.y * 2 / table1.height-1
'  If tmp > 0 Then
'    AudioFade = Csng(tmp ^10)
'  Else
'    AudioFade = Csng(-((- tmp) ^10) )
'  End If
'End Function
'
'Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "table1" is the name of the table
'  Dim tmp
'  On Error Resume Next
'  tmp = tableobj.x * 2 / table1.width-1
'  If tmp > 0 Then
'    AudioPan = Csng(tmp ^10)
'  Else
'    AudioPan = Csng(-((- tmp) ^10) )
'  End If
'End Function
'
' OPT FIX 6b: Same ^10 elimination as AudioPan. Pre-computed InvTWHalf.
Function Pan(ball)
  Dim tmp
  On Error Resume Next
  tmp = ball.x * InvTWHalf - 1
  If tmp > 0 Then
    Dim t2,t4,t8 : t2=tmp*tmp : t4=t2*t2 : t8=t4*t4
    Pan = Csng(t8 * t2)
  Else
    Dim nt,n2,n4,n8 : nt=-tmp : n2=nt*nt : n4=n2*n2 : n8=n4*n4
    Pan = Csng(-(n8 * n2))
  End If
End Function
'
'Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
'  Vol = Csng(BallVel(ball) ^2 / VolDiv)
'End Function
'
'Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
'  Pitch = BallVel(ball) * 20
'End Function
'
'Function BallVel(ball) 'Calculates the ball speed
'  BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
'End Function
'
'Function BallVelZ(ball) 'Calculates the ball speed in the -Z
'    BallVelZ = INT((ball.VelZ) * -1 )
'End Function
'
'Function VolZ(ball) ' Calculates the Volume of the sound based on the ball speed in the Z
'    VolZ = Csng(BallVelZ(ball) ^2 / 200)*1.2
'End Function

'*** Determines if a Points (px,py) is inside a 4 point polygon A-D in Clockwise/CCW order

Function InRect(px,py,ax,ay,bx,by,cx,cy,dx,dy)
  Dim AB, BC, CD, DA
  AB = (bx*py) - (by*px) - (ax*py) + (ay*px) + (ax*by) - (ay*bx)
  BC = (cx*py) - (cy*px) - (bx*py) + (by*px) + (bx*cy) - (by*cx)
  CD = (dx*py) - (dy*px) - (cx*py) + (cy*px) + (cx*dy) - (cy*dx)
  DA = (ax*py) - (ay*px) - (dx*py) + (dy*px) + (dx*ay) - (dy*ax)

  If (AB <= 0 AND BC <=0 AND CD <= 0 AND DA <= 0) Or (AB >= 0 AND BC >=0 AND CD >= 0 AND DA >= 0) Then
    InRect = True
  Else
    InRect = False
  End If
End Function


'*****************************************
'      JP's VP10 Rolling Sounds
'*****************************************

Const tnob = 10 ' total number of balls

ReDim rolling(tnob)
InitRolling

Dim DropCount
ReDim DropCount(tnob)

' OPT FIX 5: Pre-built string arrays — eliminates string concatenation in per-tick loop.
' "BallRoll_" & b = new string allocation every call. Array lookup = zero allocation.
Dim BallRollStr(10), RampLoopStr(10)
Dim brsI : For brsI = 0 To tnob
	BallRollStr(brsI) = "BallRoll_" & brsI
	RampLoopStr(brsI) = "ramploop" & brsI
Next

' OPT FIX 11: Pre-computed BallSize constants — eliminates division in per-ball loop.
Dim BS_d10 : BS_d10 = BallSize / 10
Dim BS_d5 : BS_d5 = BallSize / 5
Dim BS_d4 : BS_d4 = BallSize / 4
Dim BS_d2 : BS_d2 = BallSize / 2
Dim TW_d2 : TW_d2 = tablewidth / 2
' NOTE: BS_dAM, DynBSFactor2, DynBSFactor3 declared after shadow Consts (line ~4713)
Dim BS_dAM, DynBSFactor2, DynBSFactor3

Sub InitRolling
        Dim i
        For i = 0 to tnob
                rolling(i) = False
        Next
End Sub

' OPT FIX 5: RollingTimer rewrite.
' - Cache UBound(BOT) once at entry (was re-evaluated 3x per call).
' - Pre-built BallRollStr()/RampLoopStr() arrays eliminate string concatenation
'   ("BallRoll_" & b was 7 allocs/tick with 5 balls + cleanup; now zero).
' - Inline BallVel: cache ball.VelX/VelY into locals, compute Sqr once.
'   Eliminates redundant BallVel->VolPlayfieldRoll->PitchPlayfieldRoll chain
'   (was 6 COM reads + 3 Sqr per ball; now 4 COM reads + 1 Sqr).
' - Cache ball.z and ball.VelZ into bz/bvz for drop sound checks.
' - Use pre-computed BS_d4/BS_d2/BS_d5 for static shadow math.
Sub RollingTimer_Timer()
        Dim BOT, b
        BOT = GetBalls
        Dim ubBot : ubBot = UBound(BOT)

        ' stop the sound of deleted balls
        For b = ubBot + 1 to tnob
            If AmbientBallShadowOn = 0 Then BallShadowA(b).visible = 0
            rolling(b) = False
            StopSound(BallRollStr(b))
            StopSound(RampLoopStr(b))
            StopSound("plasticramp")
            StopSound("plasticramp2")
        Next

        ' exit the sub if no balls on the table
        If ubBot = -1 Then Exit Sub

        ' play the rolling sound for each ball
        Dim bvx, bvy, bvz, bz, bv, rollVol, rollPitch

        For b = 0 to ubBot
            bvx = BOT(b).VelX : bvy = BOT(b).VelY
            bv = INT(SQR(bvx * bvx + bvy * bvy))
            If bv > 1 Then
                rolling(b) = True
                bz = BOT(b).z
                rollVol = RollingSoundFactor * 0.0005 * Csng(bv * bv * bv)
                rollPitch = bv * bv * 15
                If bz < 30 Then
                    PlaySound BallRollStr(b), -1, rollVol * 1.1 * VolumeDial, AudioPan(BOT(b)), 0, rollPitch, 1, 0, AudioFade(BOT(b))
                    StopSound("Wireramp2jpmod")
                    StopSound("plasticramp")
                    StopSound("plasticramp2")
                    StopSound(RampLoopStr(b))
                Else
                    PlaySound RampLoopStr(b), -1, rollVol * 2.1 * VolumeDial, AudioPan(BOT(b)), 0, rollPitch, 1, 0, AudioFade(BOT(b))
                    StopSound(BallRollStr(b))
                End If
            Else
                If rolling(b) = True Then
                    StopSound(BallRollStr(b))
                    StopSound("Wireramp2jpmod")
                    rolling(b) = False
                End If
            End If

            '***Ball Drop Sounds***
            bvz = BOT(b).VelZ
            bz = BOT(b).z
            If bvz < -1 and bz < 55 and bz > 27 Then
                If DropCount(b) >= 5 Then
                    DropCount(b) = 0
                    If bvz > -7 Then
                        RandomSoundBallBouncePlayfieldSoft BOT(b)
                    Else
                        RandomSoundBallBouncePlayfieldHard BOT(b)
                    End If
                End If
            End If
            If DropCount(b) < 5 Then
                DropCount(b) = DropCount(b) + 1
            End If

            ' "Static" Ball Shadows
            If AmbientBallShadowOn = 0 Then
                If bz > 30 Then
                    BallShadowA(b).height = bz - BS_d4
                Else
                    BallShadowA(b).height = bz - BS_d2 + 5
                End If
                BallShadowA(b).Y = BOT(b).Y + BS_d5 + fovY
                BallShadowA(b).X = BOT(b).X
                BallShadowA(b).visible = 1
            End If

        Next
End Sub



'    For b = 0 to UBound(BOT)
'      If BallVel(BOT(b) ) > 1 Then
'        rolling(b) = True
'        if BOT(b).z < 30 Then ' Ball on playfield
'          PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) ), AudioPan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b) )
'        Else ' Ball on raised ramp
'          PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) )*.5, AudioPan(BOT(b) ), 0, Pitch(BOT(b) )+50000, 1, 0, AudioFade(BOT(b) )
'        End If
'      Else
'        If rolling(b) = True Then
'          StopSound("fx_ballrolling" & b)
'          rolling(b) = False
'        End If
'      End If
'    Next
'End Sub


'******************************
' destruk's new vpmCreateBall
' use it: vpmCreateBall kicker
' Use it in vpm tables instead
' of CreateBallID kickername
'******************************

Set vpmCreateBall = GetRef("mypersonalcreateballroutine")
Function mypersonalcreateballroutine(aKicker)
    For cnt = 1 to ubound(ballStatus)        ' Loop through all possible ball IDs
        If ballStatus(cnt) = 0 Then            ' If ball ID is available...
        If Not IsEmpty(vpmBallImage) Then
            Set currentball(cnt) = aKicker.CreateBall.Image            ' Set ball object with the first available ID
        Else
            Set currentball(cnt) = aKicker.CreateBall
        End If
        Set mypersonalcreateballroutine = aKicker
        currentball(cnt).uservalue = cnt            ' Assign the ball's uservalue to it's new ID
        ballStatus(cnt) = 1                ' Mark this ball status active
        ballStatus(0) = ballStatus(0)+1         ' Increment ballStatus(0), the number of active balls
    If coff = False Then                ' If collision off, overrides auto-turn on collision detection
                            ' If more than one ball active, start collision detection process
    If ballStatus(0) > 1 and XYdata.enabled = False Then XYdata.enabled = True
    End If
    Exit For                    ' New ball ID assigned, exit loop
        End If
        Next
End Function

'****************************************
' B2B Collision by Steely & Pinball Ken
' jpsalas: added destruk's changes
'  & ball height check
'****************************************

Dim tnopb, nosf
'
tnopb = 10
nosf = 10

Dim currentball(10), ballStatus(10)
Dim iball, cnt, coff, errMessage

XYdata.interval = 1
coff = False

For cnt = 0 to ubound(ballStatus)
    ballStatus(cnt) = 0
Next

' Create ball in kicker and assign a Ball ID used mostly in non-vpm tables
Sub CreateBallID(Kickername)
    For cnt = 1 to ubound(ballStatus)
        If ballStatus(cnt) = 0 Then
			If Not IsEmpty(vpmBallImage) Then		' Set ball object with the first available ID
				Set currentball(cnt) = Kickername.Createsizedball(15*brc)
			Else
				Set currentball(cnt) = Kickername.Createsizedball(15*brc)
			end If
			'Set currentball(cnt) = Kickername.createball
            currentball(cnt).uservalue = cnt
            ballStatus(cnt) = 1
            ballStatus(0) = ballStatus(0) + 1
            If coff = False Then
                If ballStatus(0) > 1 and XYdata.enabled = False Then XYdata.enabled = True
            End If
		End If
     Exit For

    Next
End Sub

Sub ClearBallID
    On Error Resume Next
    iball = ActiveBall.uservalue
    currentball(iball).UserValue = 0
    'If Err Then Msgbox Err.description & vbCrLf & iball
    ballStatus(iBall) = 0
    ballStatus(0) = ballStatus(0) -1
    On Error Goto 0
End Sub

' Ball data collection and B2B Collision detection. jpsalas: added height check
' ReDim baX(tnopb, 4), baY(tnopb, 4), baZ(tnopb, 4), bVx(tnopb, 4), bVy(tnopb, 4), TotalVel(tnopb, 4)
' Dim cForce, bDistance, xyTime, cFactor, id, id2, id3, B1, B2

' Sub XYdata_Timer()
'     xyTime = Timer + (XYdata.interval * .001)
'     If id2 >= 4 Then id2 = 0
'     id2 = id2 + 1
'     For id = 1 to ubound(ballStatus)
'         If ballStatus(id) = 1 Then
'             baX(id, id2) = round(currentball(id).x, 2)
'             baY(id, id2) = round(currentball(id).y, 2)
'             baZ(id, id2) = round(currentball(id).z, 2)
'             bVx(id, id2) = round(currentball(id).velx, 2)
'             bVy(id, id2) = round(currentball(id).vely, 2)
'             TotalVel(id, id2) = (bVx(id, id2) ^2 + bVy(id, id2) ^2)
'             If TotalVel(id, id2) > TotalVel(0, 0) Then TotalVel(0, 0) = int(TotalVel(id, id2) )
'         End If
'     Next
'
'     id3 = id2:B2 = 2:B1 = 1
'     Do
'         If ballStatus(B1) = 1 and ballStatus(B2) = 1 Then
'             bDistance = int((TotalVel(B1, id3) + TotalVel(B2, id3) ) ^1.04)
' 			If ABS(baZ(B1, id3) - baZ(B2, id3)) < 50 Then
' 				If((baX(B1, id3) - baX(B2, id3) ) ^2 + (baY(B1, id3) - baY(B2, id3) ) ^2) < 2800 + bDistance Then collide B1, B2:Exit Sub
' 			End If
'         End If
'         B1 = B1 + 1
'         If B1 >= ballStatus(0) Then Exit Do
'         If B1 >= B2 then B1 = 1:B2 = B2 + 1
'     Loop
'
'     If ballStatus(0) <= 1 Then XYdata.enabled = False
'
'     If XYdata.interval >= 40 Then coff = True:XYdata.enabled = False
'     If Timer > xyTime * 3 Then coff = True:XYdata.enabled = False
'     If Timer > xyTime Then XYdata.interval = XYdata.interval + 1
' End Sub

' 'Calculate the collision force and play sound
' Dim cTime, cb1, cb2, avgBallx, cAngle, bAngle1, bAngle2
'
' Sub Collide(cb1, cb2)
'     If TotalVel(0, 0) / 1.8 > cFactor Then cFactor = int(TotalVel(0, 0) / 1.8)
'     avgBallx = (bvX(cb2, 1) + bvX(cb2, 2) + bvX(cb2, 3) + bvX(cb2, 4) ) / 4
'     If avgBallx < bvX(cb2, id2) + .1 and avgBallx > bvX(cb2, id2) -.1 Then
'         If ABS(TotalVel(cb1, id2) - TotalVel(cb2, id2) ) < .000005 Then Exit Sub
'     End If
'     If Timer < cTime Then Exit Sub
'     cTime = Timer + .1
'     GetAngle baX(cb1, id3) - baX(cb2, id3), baY(cb1, id3) - baY(cb2, id3), cAngle
'     id3 = id3 - 1:If id3 = 0 Then id3 = 4
'     GetAngle bVx(cb1, id3), bVy(cb1, id3), bAngle1
'     GetAngle bVx(cb2, id3), bVy(cb2, id3), bAngle2
'     cForce = Cint((abs(TotalVel(cb1, id3) * Cos(cAngle-bAngle1) ) + abs(TotalVel(cb2, id3) * Cos(cAngle-bAngle2) ) ) )
'     If cForce < 4 Then Exit Sub
'     cForce = Cint((cForce) / (cFactor / nosf) )
'     If cForce > nosf-1 Then cForce = nosf-1
'     PlaySound("collide" & cForce)
' End Sub

' Get angle
' Dim Xin, Yin, rAngle, Radit, wAngle, Pi
' Pi = Round(4 * Atn(1), 6) '3.1415926535897932384626433832795
'
' Sub GetAngle(Xin, Yin, wAngle)
'     If Sgn(Xin) = 0 Then
'         If Sgn(Yin) = 1 Then rAngle = 3 * Pi / 2 Else rAngle = Pi / 2
'         If Sgn(Yin) = 0 Then rAngle = 0
'         Else
'             rAngle = atn(- Yin / Xin)
'     End If
'     If sgn(Xin) = -1 Then Radit = Pi Else Radit = 0
'     If sgn(Xin) = 1 and sgn(Yin) = 1 Then Radit = 2 * Pi
'     wAngle = round((Radit + rAngle), 4)
' End Sub


'***************************************************************
'*******                 JPJ - more lights               *******
'***************************************************************
Sub GiHelper_Timer()
	if l54.state = 1 then
		l54_reflet.state = 1
		GIcenter4.state=1
	else
		l54_reflet.state = 0
		GIcenter4.state = 0
	end if

	if l62.state = 1 then
		GIShooter.state = 1
		l62z.state = 1
		RefLaunchBall.Amount = 10
		RefLaunchBall.IntensityScale = 1
	else
		GIShooter.state = 0
		l62z.state = 0
		RefLaunchBall.Amount = 0
		RefLaunchBall.IntensityScale = 0
	end If

	if l55.state = 1 then
		l55a.state = 1
		GIShooter1.state = 1
	else
		l55a.state = 0
		GIShooter1.state = 0
	end if

	if l35.state = 1 then
		l35_reflet.state = 1
	else
		l35_reflet.state = 0
	end if

	if l40.state = 1 then
		l40_reflet.state = 1
	else
		l40_reflet.state = 0
	end if

	if l55a.state = 1 then
		l55a_reflet.state = 1
	else
		l55a_reflet.state = 0
	end if

'	if GIright11.state = 1 Then
'		sider.Image = "SideRWoodOn"
'		sidel.Image = "SideLWoodOn"
'	Else
'		sider.Image = "SideRWoodOff"
'		sidel.Image = "SideLWoodOff"
'	End If

End Sub




'**************************************************************
'************************    Sounds    ************************
'**************************************************************

' Extra Sounds
Sub Sound1_Hit:PlaySoundAtBall "metalrolling", 1:End Sub
Sub Sound2_Hit:PlaySoundAtBall "metalrolling", 1:End Sub
Sub Sound3_Hit:PlaySoundAtBall "metalrolling", 1:End Sub
Sub BallRol1_Hit:PlaySoundAtBall "ballrolling", 1:End Sub


'Sub BalldropSound: PlaySoundAtBallVol "fx_ball_drop", 1 : End Sub
' Sub OnBallBallCollision(ball1, ball2, velocity) : PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, Pan(ball1), 0, Pitch(ball1), 0, 0 : End Sub

'Sub Trigger1_Hit() : Stopsound "plasticrolling" : PlaySoundAtBallVol "metalrolling", GlobalSoundLevel * 0.3 : End Sub

'Sub Pins_Hit (idx)
'	PlaySound "plastic", 0, Vol(ActiveBall)*GlobalSoundLevel, Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
'End Sub
'
'Sub Targets_Hit (idx)
'	PlaySound "target", 0, Vol(ActiveBall)*GlobalSoundLevel, Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
'End Sub
'
'Sub MetalsThin_Hit (idx)
'	PlaySound "metalhit_thin", 0, Vol(ActiveBall)*GlobalSoundLevel, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
'End Sub
'
'Sub MetalsMedium_Hit (idx)
'	PlaySound "metalhit_medium", 0, Vol(ActiveBall)*GlobalSoundLevel, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
'End Sub
'
'Sub Metals2_Hit (idx)
'	PlaySound "metalhit2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
'End Sub
'
'Sub Gates_Hit (idx)
'	PlaySound "gate4", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
'End Sub

Sub Spinner1_Hit ()
	'PlaySound "spinner", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
SoundSpinner Spinner1
End Sub

Sub Spinner2_Hit ()
	'PlaySound "spinner", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
SoundSpinner Spinner2
End Sub

'Sub Rubbers_Hit(idx)
' 	dim finalspeed
'  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
' 	If finalspeed > 20 then
'		PlaySound "rubber_hit_1", 0, Vol(ActiveBall)*10*GlobalSoundLevel , pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
'	End if
'	If finalspeed >= 1 AND finalspeed <= 20 then
' 		RandomSoundRubber()
' 	End If
'End Sub

'Sub dPosts_Hit(idx)
'Rubbers_hit (idx)
'End Sub

'Sub RandomSoundRubber()
'	Select Case Int(Rnd*3)+1
'		Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall)*10*GlobalSoundLevel , pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
'		Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall)*10*GlobalSoundLevel , pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
'		Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall)*10*GlobalSoundLevel , pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
'	End Select
'End Sub


'Sub RandomSoundFlipper()
'	Select Case Int(Rnd*3)+1
'		Case 1 :PlaySoundAtBallVol "flip_hit_1", GlobalSoundLevel * ballvel(ActiveBall) / 50
'		Case 2 :PlaySoundAtBallVol "flip_hit_2", GlobalSoundLevel * ballvel(ActiveBall) / 50
'		Case 3 :PlaySoundAtBallVol "flip_hit_3", GlobalSoundLevel * ballvel(ActiveBall) / 50
'	End Select
'End Sub


'***************************************
'**     Data-east Flippers Ninuzzu    **
'***************************************

' OPT FIX 13: Cache currentangle into locals. Guard writes — when flippers
' are at rest (majority of gameplay), 6 COM writes/tick are eliminated.
Dim lastLFAngle : lastLFAngle = -9999
Dim lastRFAngle : lastRFAngle = -9999
Dim lastLF1Angle : lastLF1Angle = -9999

Sub Logo_Timer()
    Dim aLF : aLF = LeftFlipper.CurrentAngle
    If aLF <> lastLFAngle Then
        lastLFAngle = aLF
        LFLogo.RotY = aLF
        pLeftFlipperShadow.ObjRotZ = aLF + 1
    End If
    Dim aRF : aRF = RightFlipper.CurrentAngle
    If aRF <> lastRFAngle Then
        lastRFAngle = aRF
        RFLogo.RotY = aRF
        pRightFlipperShadow.ObjRotZ = aRF + 1
    End If
    Dim aLF1 : aLF1 = LeftFlipper1.CurrentAngle
    If aLF1 <> lastLF1Angle Then
        lastLF1Angle = aLF1
        LFLogo001.RotY = aLF1
        pLeftFlipperShadow001.ObjRotZ = aLF1 + 1
    End If
End Sub


'****************************
'**     Anim Flashers      **
'****************************

sub FlashClearAnim_timer()
select Case FlashClearEtape
	Case 0:Primitive13.image = "OctaDome_Clear_00":FlashClearEtape=0
	flash2veille1.state = 1
	flash2veille1.intensity = 0.75
	Case 1:Primitive13.image = "OctaDome_Clear_10":FlashClearEtape=2
	flash2veille1.state = 1
	flash2veille1.intensity = 14
	Case 2:Primitive13.image = "OctaDome_Clear_09":FlashClearEtape=3
	flash2veille1.state = 1
	flash2veille1.intensity = 12.6
	Case 3:Primitive13.image = "OctaDome_Clear_08":FlashClearEtape=4
	flash2veille1.state = 1
	flash2veille1.intensity = 11.2
	Case 4:Primitive13.image = "OctaDome_Clear_07":FlashClearEtape=5
	flash2veille1.state = 1
	flash2veille1.intensity = 9.8
	Case 5:Primitive13.image = "OctaDome_Clear_06":FlashClearEtape=6
	flash2veille1.state = 1
	flash2veille1.intensity = 8.4
	Case 6:Primitive13.image = "OctaDome_Clear_05":FlashClearEtape=7
	flash2veille1.state = 1
	flash2veille1.intensity = 7
	Case 7:Primitive13.image = "OctaDome_Clear_04":FlashClearEtape=8
	flash2veille1.state = 1
	flash2veille1.intensity = 5.6
	Case 8:Primitive13.image = "OctaDome_Clear_03":FlashClearEtape=9
	flash2veille1.state = 1
	flash2veille1.intensity = 4.2
	Case 9:Primitive13.image = "OctaDome_Clear_02":FlashClearEtape=10
	flash2veille1.state = 1
	flash2veille1.intensity = 2.8
	Case 10:Primitive13.image = "OctaDome_Clear_01":FlashClearEtape=11
	flash2veille1.state = 1
	flash2veille1.intensity = 1.4
	Case 11:Primitive13.image = "OctaDome_Clear_00":FlashClearEtape=0
	flash2veille1.state = 0.75
	flash2veille1.intensity = 0.75
End select
End Sub

sub FlashClearAnim2_timer()
select Case FlashClearEtapeR
	Case 0:Primitive16.image = "OctaDome_Clear_00":FlashClearEtapeR=0
	flash2veille.state = 1
	flash2veille.intensity = 0.75
	Case 1:Primitive16.image = "OctaDome_Clear_10":FlashClearEtapeR=2
	flash2veille.state = 1
	flash2veille.intensity = 20
	Case 2:Primitive16.image = "OctaDome_Clear_09":FlashClearEtapeR=3
	flash2veille.state = 1
	flash2veille.intensity = 18
	Case 3:Primitive16.image = "OctaDome_Clear_08":FlashClearEtapeR=4
	flash2veille.state = 1
	flash2veille.intensity = 16
	Case 4:Primitive16.image = "OctaDome_Clear_07":FlashClearEtapeR=5
	flash2veille.state = 1
	flash2veille.intensity = 14
	Case 5:Primitive16.image = "OctaDome_Clear_06":FlashClearEtapeR=6
	flash2veille.state = 1
	flash2veille.intensity = 12
	Case 6:Primitive16.image = "OctaDome_Clear_05":FlashClearEtapeR=7
	flash2veille.state = 1
	flash2veille.intensity = 10
	Case 7:Primitive16.image = "OctaDome_Clear_04":FlashClearEtapeR=8
	flash2veille.state = 1
	flash2veille.intensity = 8
	Case 8:Primitive16.image = "OctaDome_Clear_03":FlashClearEtapeR=9
	flash2veille.state = 1
	flash2veille.intensity = 6
	Case 9:Primitive16.image = "OctaDome_Clear_02":FlashClearEtapeR=10
	flash2veille.state = 1
	flash2veille.intensity = 4
	Case 10:Primitive16.image = "OctaDome_Clear_01":FlashClearEtapeR=11
	flash2veille.state = 1
	flash2veille.intensity = 2
	Case 11:Primitive16.image = "OctaDome_Clear_00":FlashClearEtapeR=0
	flash2veille.state = 0.75
	flash2veille.intensity = 0.75
End select
End Sub

'sub FlashGreenAnim_timer()
select Case FlashGreenEtape
	Case 0:Primitive14.image = "OctaDome_Green_00":FlashGreenEtape=0
	flash6i.state = 1
	flash6i.intensity = 0.3
	Case 1:Primitive14.image = "OctaDome_Green_10":FlashGreenEtape=2
	flash6i.state = 1
	flash6i.intensity = 12
	Case 2:Primitive14.image = "OctaDome_Green_09":FlashGreenEtape=3
	flash6i.state = 1
	flash6i.intensity = 10.8
	Case 3:Primitive14.image = "OctaDome_Green_08":FlashGreenEtape=4
	flash6i.state = 1
	flash6i.intensity = 9.6
	Case 4:Primitive14.image = "OctaDome_Green_07":FlashGreenEtape=5
	flash6i.state = 1
	flash6i.intensity = 8.4
	Case 5:Primitive14.image = "OctaDome_Green_06":FlashGreenEtape=6
	flash6i.state = 1
	flash6i.intensity = 7.2
	Case 6:Primitive14.image = "OctaDome_Green_05":FlashGreenEtape=7
	flash6i.state = 1
	flash6i.intensity = 6
	Case 7:Primitive14.image = "OctaDome_Green_04":FlashGreenEtape=8
	flash6i.state = 1
	flash6i.intensity = 4.8
	Case 8:Primitive14.image = "OctaDome_Green_03":FlashGreenEtape=9
	flash6i.state = 1
	flash6i.intensity = 3.6
	Case 9:Primitive14.image = "OctaDome_Green_02":FlashGreenEtape=10
	flash6i.state = 1
	flash6i.intensity = 2.4
	Case 10:Primitive14.image = "OctaDome_Green_01":FlashGreenEtape=11
	flash6i.state = 1
	flash6i.intensity = 1.2
	Case 11:Primitive14.image = "OctaDome_Green_00":FlashGreenEtape=0
	flash6i.state = 1
	flash6i.intensity = 0.75
End select
'End Sub

'sub FlashRedAnim_timer()
select Case FlashRedEtape
	Case 0:Primitive15.image = "OctaDome_Red_00":FlashRedEtape=0
	flash5i.state = 1
	flash5i.intensity = 0.75
	Case 1:Primitive15.image = "OctaDome_Red_10":FlashRedEtape=2
	flash5i.state = 1
	flash5i.intensity = 12
	Case 2:Primitive15.image = "OctaDome_Red_09":FlashRedEtape=3
	flash5i.state = 1
	flash5i.intensity = 10.8
	Case 3:Primitive15.image = "OctaDome_Red_08":FlashRedEtape=4
	flash5i.state = 1
	flash5i.intensity = 9.6
	Case 4:Primitive15.image = "OctaDome_Red_07":FlashRedEtape=5
	flash5i.state = 1
	flash5i.intensity = 8.4
	Case 5:Primitive15.image = "OctaDome_Red_06":FlashRedEtape=6
	flash5i.state = 1
	flash5i.intensity = 7.2
	Case 6:Primitive15.image = "OctaDome_Red_05":FlashRedEtape=7
	flash5i.state = 1
	flash5i.intensity = 6
	Case 7:Primitive15.image = "OctaDome_Red_04":FlashRedEtape=8
	flash5i.state = 1
	flash5i.intensity = 4.8
	Case 8:Primitive15.image = "OctaDome_Red_03":FlashRedEtape=9
	flash5i.state = 1
	flash5i.intensity = 3.6
	Case 9:Primitive15.image = "OctaDome_Red_02":FlashRedEtape=10
	flash5i.state = 1
	flash5i.intensity = 2.4
	Case 10:Primitive15.image = "OctaDome_Red_01":FlashRedEtape=11
	flash5i.state = 1
	flash5i.intensity = 1.2
	Case 11:Primitive15.image = "OctaDome_Red_00":FlashRedEtape=0
	flash5i.state = 1
	flash5i.intensity = 0.75
End select
'End Sub



'****************************
'** Reflection Mod Routine **
'****************************
Sub MystAnimationB
if Myst=1 and Mystv=1 then Myst1.visible = 1:mm01.visible = 0:mm02.visible = 0:mm03.visible = 0:mm04.visible = 0:mml01.state = 0:end If
if Myst=1 and Mystv=>1 then Myst1.rotz = Myst1.rotz +3
if Myst=1 and Mystv=2 then Myst2.visible = 1:mm01.visible = 0:mm02.visible = 0:mm03.visible = 0:mm04.visible = 0:mml01.state = 0:end If
if Myst=1 and Mystv=>2 then Myst2.rotz = Myst2.rotz +5
if Myst=1 and Mystv=3 then Myst3.visible = 1:mm01.visible = 1:mm02.visible = 0:mm03.visible = 0:mm04.visible = 0:mml01.state = 1:mml01.intensity = 5:end If
if Myst=1 and Mystv=>3 then Myst3.rotz = Myst3.rotz +2
if Myst=1 and Mystv=4 then Myst4.visible = 1:mm01.visible = 1:mm02.visible = 0:mm03.visible = 0:mm04.visible = 0:mml01.state = 1:mml01.intensity = 7:end If
if Myst=1 and Mystv=>4 then Myst4.rotz = Myst4.rotz +3
if Myst=1 and Mystv=5 then Myst5.visible = 1:mm01.visible = 0:mm02.visible = 1:mm03.visible = 0:mm04.visible = 0:mml01.state = 1:mml01.intensity = 9:end If
if Myst=1 and Mystv=>5 then Myst5.rotz = Myst5.rotz +4
if Myst=1 and Mystv=6 then Myst6.visible = 1:mm01.visible = 0:mm02.visible = 1:mm03.visible = 0:mm04.visible = 0:mml01.state = 1:mml01.intensity = 11:end If
if Myst=1 and Mystv=>6 then Myst6.rotz = Myst6.rotz +2
if Myst=1 and Mystv=7 then Myst7.visible = 1:mm01.visible = 0:mm02.visible = 0:mm03.visible = 1:mm04.visible = 0:mml01.state = 1:mml01.intensity = 13:end If
if Myst=1 and Mystv=>7 then Myst7.rotz = Myst7.rotz +3
if Myst=1 and Mystv=8 then Myst8.visible = 1:mm01.visible = 0:mm02.visible = 0:mm03.visible = 1:mm04.visible = 0:mml01.state = 1:mml01.intensity = 15:end If
if Myst=1 and Mystv=>8 then Myst8.rotz = Myst8.rotz +4
if Myst=1 and Mystv=9 then Myst9.visible = 1:mm01.visible = 0:mm02.visible = 0:mm03.visible = 0:mm04.visible = 1:mml01.state = 1:mml01.intensity = 17:end If
if Myst=1 and Mystv=>9 then Myst9.rotz = Myst9.rotz +2
if Myst=1 and Mystv=10 then Myst10.visible = 1:mm01.visible = 0:mm02.visible = 0:mm03.visible = 0:mm04.visible = 1:mml01.state = 1:mml01.intensity = 20:end If
if Myst=1 and Mystv=10 then Myst10.rotz = Myst10.rotz +6
end Sub

Sub RM_Timer()

if myst = 1 then MystAnimationB
pgate5.Rotx = gate5.CurrentAngle + 180

If l1.state = 1 then l1z.state = 1 Else l1z.state = 0: end If
If l2.state = 1 then l2z.state = 1 Else l2z.state = 0: end If
If l3.state = 1 then l3z.state = 1 Else l3z.state = 0: end If
If l4.state = 1 then l4z.state = 1 Else l4z.state = 0: end If
If l5.state = 1 then l5z.state = 1 Else l5z.state = 0: end If
If l6.state = 1 then l6z.state = 1 Else l6z.state = 0: end If
If l7.state = 1 then l7z.state = 1 Else l7z.state = 0: end If
If l8.state = 1 then l8z.state = 1 Else l8z.state = 0: end If
'If l9.state = 1 then l9z.state = 1 Else l9z.state = 0: end If
If l10.state = 1 then l10z.state = 1 Else l10z.state = 0: end If
If l11.state = 1 then l11z.state = 1 Else l11z.state = 0: end If
'If l12.state = 1 then l12z.state = 1 Else l12z.state = 0: end If
'If l13.state = 1 then l13z.state = 1 Else l13z.state = 0: end If
'If l14.state = 1 then l14z.state = 1 Else l14z.state = 0: end If
'If l15.state = 1 then l15z.state = 1 Else l15z.state = 0: end If
'If l16.state = 1 then l16z.state = 1 Else l16z.state = 0: end If
If l17.state = 1 then l17z.state = 1 Else l17z.state = 0: end If
If l17.state = 1 then l17r.state = 1 Else l17r.state = 0: end If
If l18.state = 1 then l18z.state = 1 Else l18z.state = 0: end If
If l18.state = 1 then l18r.state = 1 Else l18r.state = 0: end If
If l19.state = 1 then l19z.state = 1 Else l19z.state = 0: end If
'If l20.state = 1 then l20z.state = 1 Else l20z.state = 0: end If
'If l21.state = 1 then l21z.state = 1 Else l21z.state = 0: end If
'If l22.state = 1 then l22z.state = 1 Else l22z.state = 0: end If
'If l23.state = 1 then l23z.state = 1 Else l23z.state = 0: end If
'If l24.state = 1 then l24z.state = 1 Else l24z.state = 0: end If
If l25.state = 1 then l25z.state = 1 Else l25z.state = 0: end If
If l26.state = 1 then l26z.state = 1 Else l26z.state = 0: end If
If l27.state = 1 then l27z.state = 1 Else l27z.state = 0: end If
If l28.state = 1 then l28z.state = 1 Else l28z.state = 0: end If
If f28b.state = 1 then f28bz.state = 1 Else f28bz.state = 0: end If
If f28a.state = 1 then f28az.state = 1 Else f28az.state = 0: end If
If l29.state = 1 then l29z.state = 1 Else l29z.state = 0: end If
If l29.state = 1 then l29a.state = 1 Else l29a.state = 0: end If
If l30.state = 1 then l30z.state = 1 Else l30z.state = 0: end If
If l30.state = 1 then l30a.state = 1 Else l30a.state = 0: end If
If l31.state = 1 then l31z.state = 1 Else l31z.state = 0: end If
If l31.state = 1 then l31a.state = 1 Else l31a.state = 0: end If
If l32.state = 1 then l32z.state = 1 Else l32z.state = 0: end If
If l32.state = 1 then l32a.state = 1 Else l32a.state = 0: end If
If l33.state = 1 then l33z.state = 1 Else l33z.state = 0: end If
If l34.state = 1 then l34z.state = 1 Else l34z.state = 0: end If
'If l35.state = 1 then l35z.state = 1 Else l35z.state = 0: end If
If l36.state = 1 then
		l36r.state = 1
		flash36r.amount = 120
		flash36r.intensityScale = 2
	Else
		l36r.state = 0
		flash36r.amount = 0
		flash36r.intensityScale = 0
end If
If l37.state = 1 then l37z.state = 1 Else l37z.state = 0: end If
If l38.state = 1 then l38z.state = 1 Else l38z.state = 0: end If
If l39.state = 1 then l39z.state = 1 Else l39z.state = 0: end If
'If l40.state = 1 then l40z.state = 1 Else l40z.state = 0: end If
If l41.state = 1 then l41z.state = 1 Else l41z.state = 0: end If
If l42.state = 1 then l42z.state = 1 Else l42z.state = 0: end If
If l43.state = 1 then l43z.state = 1 Else l43z.state = 0: end If
If l44.state = 1 then l44z.state = 1 Else l44z.state = 0: end If
If l45.state = 1 then l45z.state = 1 Else l45z.state = 0: end If
If l46.state = 1 then l46z.state = 1 Else l46z.state = 0: end If
If l47.state = 1 then l47z.state = 1 Else l47z.state = 0: end If
If l48.state = 1 then l48z.state = 1 Else l48z.state = 0: end If
If l49.state = 1 then l49z.state = 1 Else l49z.state = 0: end If
If l50.state = 1 then l50z.state = 1 Else l50z.state = 0: end If
If l51.state = 1 then l51z.state = 1 Else l51z.state = 0: end If
If l52.state = 1 then l52z.state = 1 Else l52z.state = 0: end If
If l53.state = 1 then l53z.state = 1 Else l53z.state = 0: end If
'If l54.state = 1 then l54z.state = 1 Else l54z.state = 0: end If
If l55.state = 1 then GIShooter1.state = 1 Else GIShooter1.state = 0: end If
If l56.state = 1 then l56z.state = 1 Else l56z.state = 0: end If
If l57.state = 1 then l57z.state = 1 Else l57z.state = 0: end If
If l57.state = 1 then l57r.state = 1 Else l57r.state = 0: end If
If l58.state = 1 then l58z.state = 1 Else l58z.state = 0: end If
If l58.state = 1 then l58r.state = 1 Else l58r.state = 0: end If
If l59.state = 1 then l59z.state = 1 Else l59z.state = 0: end If
If l59.state = 1 then l59r.state = 1 Else l59r.state = 0: end If
If l60.state = 1 then l60z.state = 1 Else l60z.state = 0: end If
If l60.state = 1 then l60r.state = 1 Else l60r.state = 0: end If
If l61.state = 1 then l61z.state = 1 Else l61z.state = 0: end If
If f26S.state = 1 then f26Sz.state = 1 Else f26Sz.state = 0: end If
If JackpotRojo.state = 1 then 	JackpotRojoz.state = 1 Else JackpotRojoz.state = 0: end If
If JackpotAm.state = 1 then JackpotAmz.state = 1 Else JackpotAmz.state = 0: end If
End Sub

Sub resetMyst

Myst1.visible = 0
Myst2.visible = 0
Myst3.visible = 0
Myst4.visible = 0
Myst5.visible = 0
Myst6.visible = 0
Myst7.visible = 0
Myst8.visible = 0
Myst9.visible = 0
Myst10.visible = 0
Mystv=0
if myst =1 then mm01.visible = 0:mm02.visible = 0:mm03.visible = 0:mm04.visible = 0:mml01.state = 0:mml01.intensity = 0:WinSounds:end If
end Sub

sub WinSounds()
	Select Case Int(Rnd*3)
		Case 0 : PlaySound SoundFX("bierwin",DOFContactors)
		Case 1 : PlaySound SoundFX("bierwin2",DOFContactors)
		Case 2 : PlaySound SoundFX("bierwin3",DOFContactors)
	End Select
End Sub

'****************************************
'Flasher Routines
'****************************************
Dim FlashState(200), FlashLevel(200)
Dim LampState(200), FadingLevel(200)
Dim FlashSpeedUp, FlashSpeedDown

FlashInit()
FlasherTimer.Interval = 10 'flash fading speed
FlasherTimer.Enabled = 1
'
'' Lamp & Flasher Timers
'
Sub FlasherTimer_Timer()
	Flash  101, Flasher101
	Flash  102, Flasher102
	NfadeL 103, Flasher103

End Sub





Sub FlashInit
    Dim i
    For i = 0 to 200
        FlashState(i) = 0
        FlashLevel(i) = 0
    Next

    FlashSpeedUp = 50   ' fast speed when turning on the flasher
    FlashSpeedDown = 10 ' slow speed when turning off the flasher, gives a smooth fading
    AllFlashOff()
End Sub

Sub AllFlashOff
    Dim i
    For i = 0 to 200
        FlashState(i) = 0
    Next
End Sub

Sub SetFlash(nr, stat)
    FlashState(nr) = ABS(stat)
End Sub

Sub Flash(nr, object)
    Select Case FlashState(nr)
        Case 0 'off
            FlashLevel(nr) = FlashLevel(nr) - FlashSpeedDown
            If FlashLevel(nr) < 0 Then
                FlashLevel(nr) = 0
                FlashState(nr) = -1 'completely off
            End if
 '           Object.alpha = FlashLevel(nr)
        Case 1 ' on
            FlashLevel(nr) = FlashLevel(nr) + FlashSpeedUp
            If FlashLevel(nr) > 255 Then
                FlashLevel(nr) = 255
                FlashState(nr) = -2 'completely on
            End if
 '           Object.alpha = FlashLevel(nr)
    End Select
End Sub

Sub NFadeL(nr, a)
    Select Case FadingLevel(nr)
        Case 4:a.state = 0:FadingLevel(nr) = 0
        Case 5:a.State = 1:FadingLevel(nr) = 1
    End Select
End Sub

''cFastFlips by nFozzy
''Bypasses pinmame callback for faster and more responsive flippers
''Version 1.1 beta2 (More proper behaviour, extra safety against script errors)
''*************************************************
'Function NullFunction(aEnabled):End Function    '1 argument null function placeholder
'Class cFastFlips
'    Public TiltObjects, DebugOn, hi
'    Private SubL, SubUL, SubR, SubUR, FlippersEnabled, Delay, LagCompensation, Name, FlipState(3)
'
'    Private Sub Class_Initialize()
'        Delay = 0 : FlippersEnabled = False : DebugOn = False : LagCompensation = False
'        Set SubL = GetRef("NullFunction"): Set SubR = GetRef("NullFunction") : Set SubUL = GetRef("NullFunction"): Set SubUR = GetRef("NullFunction")
'    End Sub
'
'    'set callbacks
'    Public Property Let CallBackL(aInput)  : Set SubL  = GetRef(aInput) : Decouple sLLFlipper, aInput: End Property
'    Public Property Let CallBackUL(aInput) : Set SubUL = GetRef(aInput) : End Property
'    Public Property Let CallBackR(aInput)  : Set SubR  = GetRef(aInput) : Decouple sLRFlipper, aInput:  End Property
'    Public Property Let CallBackUR(aInput) : Set SubUR = GetRef(aInput) : End Property
'    Public Sub InitDelay(aName, aDelay) : Name = aName : delay = aDelay : End Sub   'Create Delay
'    'Automatically decouple flipper solcallback script lines (only if both are pointing to the same sub) thanks gtxjoe
'    Private Sub Decouple(aSolType, aInput)  : If StrComp(SolCallback(aSolType),aInput,1) = 0 then SolCallback(aSolType) = Empty End If : End Sub
'
'    'call callbacks
'    Public Sub FlipL(aEnabled)
'        FlipState(0) = aEnabled 'track flipper button states: the game-on sol flips immediately if the button is held down (1.1)
'        If not FlippersEnabled and not DebugOn then Exit Sub
'        subL aEnabled
'    End Sub
'
'    Public Sub FlipR(aEnabled)
'        FlipState(1) = aEnabled
'        If not FlippersEnabled and not DebugOn then Exit Sub
'        subR aEnabled
'    End Sub
'
'    Public Sub FlipUL(aEnabled)
'        FlipState(2) = aEnabled
'        If not FlippersEnabled and not DebugOn then Exit Sub
'        subUL aEnabled
'    End Sub
'
'    Public Sub FlipUR(aEnabled)
'        FlipState(3) = aEnabled
'        If not FlippersEnabled and not DebugOn then Exit Sub
'        subUR aEnabled
'    End Sub
'
'    Public Sub TiltSol(aEnabled)    'Handle solenoid / Delay (if delayinit)
'        If delay > 0 and not aEnabled then  'handle delay
'            vpmtimer.addtimer Delay, Name & ".FireDelay" & "'"
'            LagCompensation = True
'        else
'            If Delay > 0 then LagCompensation = False
'            EnableFlippers(aEnabled)
'        end If
'    End Sub
'
'    Sub FireDelay() : If LagCompensation then EnableFlippers False End If : End Sub
'
'    Private Sub EnableFlippers(aEnabled)
'        If aEnabled then SubL FlipState(0) : SubR FlipState(1) : subUL FlipState(2) : subUR FlipState(3)
'        FlippersEnabled = aEnabled
'        If TiltObjects then vpmnudge.solgameon aEnabled
'        If Not aEnabled then
'            subL False
'            subR False
'            If not IsEmpty(subUL) then subUL False
'            If not IsEmpty(subUR) then subUR False
'        End If
'    End Sub
'
'
'    End Class
'
Sub FlashL(nr, a)
    Select Case FadingLevel(nr)
        Case 4:a.state = 0:FadingLevel(nr) = 0 ':b.state = 0:FadingLevel(nr) = 0 pour une deuxième
        Case 5:a.state = 1:FadingLevel(nr) = 1 ':b.state = 1:FadingLevel(nr) = 1
    End Select
End Sub

'******************************************************
'****  GENEREAL ADVICE ON PHYSICS
'******************************************************
'
' It's advised that flipper corrections, dampeners, and general physics settings should all be updated per these 
' examples as all of these improvements work together to provide a realistic physics simulation.
'
' Tutorial videos provided by Bord
' Flippers: 	https://www.youtube.com/watch?v=FWvM9_CdVHw
' Dampeners: 	https://www.youtube.com/watch?v=tqsxx48C6Pg
' Physics: 		https://www.youtube.com/watch?v=UcRMG-2svvE
'
'
' Note: BallMass must be set to 1. BallSize should be set to 50 (in other words the ball radius is 25) 
'
' Recommended Table Physics Settings
' | Gravity Constant             | 0.97      |
' | Playfield Friction           | 0.15-0.25 |
' | Playfield Elasticity         | 0.25      |
' | Playfield Elasticity Falloff | 0         |
' | Playfield Scatter            | 0         |
' | Default Element Scatter      | 2         |
'
' Bumpers
' | Force         | 9.5-10.5 |
' | Hit Threshold | 1.6-2    |
' | Scatter Angle | 2        |
' 
' Slingshots
' | Hit Threshold      | 2    |
' | Slingshot Force    | 4-5  |
' | Slingshot Theshold | 2-3  |
' | Elasticity         | 0.85 |
' | Friction           | 0.8  |
' | Scatter Angle      | 1    |



'******************************************************
'****  FLIPPER CORRECTIONS by nFozzy
'******************************************************
'
' There are several steps for taking advantage of nFozzy's flipper solution.  At a high level weÂ?ll need the following:
'	1. flippers with specific physics settings
'	2. custom triggers for each flipper (TriggerLF, TriggerRF)
'	3. an object or point to tell the script where the tip of the flipper is at rest (EndPointLp, EndPointRp)
'	4. and, special scripting
'
' A common mistake is incorrect flipper length.  A 3-inch flipper with rubbers will be about 3.125 inches long.  
' This translates to about 147 vp units.  Therefore, the flipper start radius + the flipper length + the flipper end 
' radius should  equal approximately 147 vp units. Another common mistake is is that sometimes the right flipper
' angle was set with a large postive value (like 238 or something). It should be using negative value (like -122).
'
' The following settings are a solid starting point for various eras of pinballs.
' |                    | EM's           | late 70's to mid 80's | mid 80's to early 90's | mid 90's and later |
' | ------------------ | -------------- | --------------------- | ---------------------- | ------------------ |
' | Mass               | 1              | 1                     | 1                      | 1                  |
' | Strength           | 500-1000 (750) | 1400-1600 (1500)      | 2000-2600              | 3200-3300 (3250)   |
' | Elasticity         | 0.88           | 0.88                  | 0.88                   | 0.88               |
' | Elasticity Falloff | 0.15           | 0.15                  | 0.15                   | 0.15               |
' | Fricition          | 0.8-0.9        | 0.9                   | 0.9                    | 0.9                |
' | Return Strength    | 0.11           | 0.09                  | 0.07                   | 0.055              |
' | Coil Ramp Up       | 2.5            | 2.5                   | 2.5                    | 2.5                |
' | Scatter Angle      | 0              | 0                     | 0                      | 0                  |
' | EOS Torque         | 0.3            | 0.3                   | 0.275                  | 0.275              |
' | EOS Torque Angle   | 4              | 4                     | 6                      | 6                  |
'


'******************************************************
' Flippers Polarity (Select appropriate sub based on era) 
'******************************************************

dim LF : Set LF = New FlipperPolarity
dim RF : Set RF = New FlipperPolarity
dim RF1: Set RF1 = New FlipperPolarity
dim LF1: Set LF1 = New FlipperPolarity
dim LF2 : Set LF2 = New FlipperPolarity
dim RF2 : Set RF2 = New FlipperPolarity
dim LF3 : Set LF3 = New FlipperPolarity
dim RF3 : Set RF3 = New FlipperPolarity


InitPolarity

'
''*******************************************
'' Late 70's to early 80's
'
'Sub InitPolarity()
'        dim x, a : a = Array(LF, RF, LF2, RF2, LF3, RF3)
'        for each x in a
'                x.AddPoint "Ycoef", 0, RightFlipper.Y-65, 1        'disabled
'                x.AddPoint "Ycoef", 1, RightFlipper.Y-11, 1
'                x.enabled = True
'                x.TimeDelay = 80  '*****Important, this variable is an offset for the speed that the ball travels down the table to determine if the flippers have been fired 
'							'This is needed because the corrections to ball trajectory should only applied if the flippers have been fired and the ball is in the trigger zones.
'							'FlipAT is set to GameTime when the ball enters the flipper trigger zones and if GameTime is less than FlipAT + this time delay then changes to velocity
'							'and trajectory are applied.  If the flipper is fired before the ball enters the trigger zone then with this delay added to FlipAT the changes
'							'to tragectory and velocity will not be applied.  Also if the flipper is in the final 20 degrees changes to ball values will also not be applied.
'							'"Faster" tables will need a smaller value while "slower" tables will need a larger value to give the ball more time to get to the flipper. 		
'							'If this value is not set high enough the Flipper Velocity and Polarity corrections will NEVER be applied.
'
'        Next
'
'        AddPt "Polarity", 0, 0, 0
'        AddPt "Polarity", 1, 0.05, -2.7        
'        AddPt "Polarity", 2, 0.33, -2.7
'        AddPt "Polarity", 3, 0.37, -2.7        
'        AddPt "Polarity", 4, 0.41, -2.7
'        AddPt "Polarity", 5, 0.45, -2.7
'        AddPt "Polarity", 6, 0.576,-2.7
'        AddPt "Polarity", 7, 0.66, -1.8
'        AddPt "Polarity", 8, 0.743, -0.5
'        AddPt "Polarity", 9, 0.81, -0.5
'        AddPt "Polarity", 10, 0.88, 0
'
'        addpt "Velocity", 0, 0,         1
'        addpt "Velocity", 1, 0.16, 1.06
'        addpt "Velocity", 2, 0.41,         1.05
'        addpt "Velocity", 3, 0.53,         1'0.982
'        addpt "Velocity", 4, 0.702, 0.968
'        addpt "Velocity", 5, 0.95,  0.968
'        addpt "Velocity", 6, 1.03,         0.945
'
'        LF.Object = LeftFlipper 
'        LF.EndPoint = EndPointLp  'you can use just a coordinate, or an object with a .x property. Using a couple of simple primitive objects
'        RF.Object = RightFlipper
'        RF.EndPoint = EndPointRp
'		'LF2.Object = LeftFlipper2        
'        'LF2.EndPoint = EndPointLp2
'        'RF2.Object = RightFlipper2
'        'RF2.EndPoint = EndPointRp2
'		'LF3.Object = LeftFlipper3        
'        'LF3.EndPoint = EndPointLp3
'        'RF3.Object = RightFlipper3
'        'RF3.EndPoint = EndPointRp3
'
'End Sub



''*******************************************
'' Mid 80's
'
'Sub InitPolarity()
'        dim x, a : a = Array(LF, RF, LF2, RF2, LF3, RF3)
'        for each x in a
'                x.AddPoint "Ycoef", 0, RightFlipper.Y-65, 1        'disabled
'                x.AddPoint "Ycoef", 1, RightFlipper.Y-11, 1
'                x.enabled = True
'                x.TimeDelay = 80
'        Next
'
'        AddPt "Polarity", 0, 0, 0
'        AddPt "Polarity", 1, 0.05, -3.7        
'        AddPt "Polarity", 2, 0.33, -3.7
'        AddPt "Polarity", 3, 0.37, -3.7
'        AddPt "Polarity", 4, 0.41, -3.7
'        AddPt "Polarity", 5, 0.45, -3.7 
'        AddPt "Polarity", 6, 0.576,-3.7
'        AddPt "Polarity", 7, 0.66, -2.3
'        AddPt "Polarity", 8, 0.743, -1.5
'        AddPt "Polarity", 9, 0.81, -1
'        AddPt "Polarity", 10, 0.88, 0
'
'        addpt "Velocity", 0, 0,         1
'        addpt "Velocity", 1, 0.16, 1.06
'        addpt "Velocity", 2, 0.41,         1.05
'        addpt "Velocity", 3, 0.53,         1'0.982
'        addpt "Velocity", 4, 0.702, 0.968
'        addpt "Velocity", 5, 0.95,  0.968
'        addpt "Velocity", 6, 1.03,         0.945
'
'        LF.Object = LeftFlipper        
'        LF.EndPoint = EndPointLp
'        RF.Object = RightFlipper
'        RF.EndPoint = EndPointRp
'		'LF2.Object = LeftFlipper2        
'       'LF2.EndPoint = EndPointLp2
'       'RF2.Object = RightFlipper2
'       'RF2.EndPoint = EndPointRp2
'		'LF3.Object = LeftFlipper3        
'       'LF3.EndPoint = EndPointLp3
'       'RF3.Object = RightFlipper3
'       'RF3.EndPoint = EndPointRp3
'End Sub
'
'


'*******************************************
'  Late 80's early 90's
'
'Sub InitPolarity()
'        dim x, a : a = Array(LF, RF, LF2, RF2, LF3, RF3)
'	for each x in a
'		x.AddPoint "Ycoef", 0, RightFlipper.Y-65, 1        'disabled
'		x.AddPoint "Ycoef", 1, RightFlipper.Y-11, 1
'		x.enabled = True
'		x.TimeDelay = 60
'	Next
'
'	AddPt "Polarity", 0, 0, 0
'	AddPt "Polarity", 1, 0.05, -5
'	AddPt "Polarity", 2, 0.4, -5
'	AddPt "Polarity", 3, 0.6, -4.5
'	AddPt "Polarity", 4, 0.65, -4.0
'	AddPt "Polarity", 5, 0.7, -3.5
'	AddPt "Polarity", 6, 0.75, -3.0
'	AddPt "Polarity", 7, 0.8, -2.5
'	AddPt "Polarity", 8, 0.85, -2.0
'	AddPt "Polarity", 9, 0.9,-1.5
'	AddPt "Polarity", 10, 0.95, -1.0
'	AddPt "Polarity", 11, 1, -0.5
'	AddPt "Polarity", 12, 1.1, 0
'	AddPt "Polarity", 13, 1.3, 0
'
'	addpt "Velocity", 0, 0,         1
'	addpt "Velocity", 1, 0.16, 1.06
'	addpt "Velocity", 2, 0.41,         1.05
'	addpt "Velocity", 3, 0.53,         1'0.982
'	addpt "Velocity", 4, 0.702, 0.968
'	addpt "Velocity", 5, 0.95,  0.968
'	addpt "Velocity", 6, 1.03,         0.945
'
'        LF.Object = LeftFlipper        
'        LF.EndPoint = EndPointLp
'        RF.Object = RightFlipper
'        RF.EndPoint = EndPointRp
'		'LF2.Object = LeftFlipper2        
'       'LF2.EndPoint = EndPointLp2
'       'RF2.Object = RightFlipper2
'       'RF2.EndPoint = EndPointRp2
'		'LF3.Object = LeftFlipper3        
'       'LF3.EndPoint = EndPointLp3
'       'RF3.Object = RightFlipper3
'       'RF3.EndPoint = EndPointRp3
'End Sub
'
'
'
'
'*******************************************
'' Early 90's and after
'
Sub InitPolarity()
        dim x, a : a = Array(LF, RF, LF2, RF2, LF3, RF3)
        for each x in a
                x.AddPoint "Ycoef", 0, RightFlipper.Y-65, 1        'disabled
                x.AddPoint "Ycoef", 1, RightFlipper.Y-11, 1
                x.enabled = True
                x.TimeDelay = 60
        Next

        AddPt "Polarity", 0, 0, 0
        AddPt "Polarity", 1, 0.05, -5.5
        AddPt "Polarity", 2, 0.4, -5.5
        AddPt "Polarity", 3, 0.6, -5.0
        AddPt "Polarity", 4, 0.65, -4.5
        AddPt "Polarity", 5, 0.7, -4.0
        AddPt "Polarity", 6, 0.75, -3.5
        AddPt "Polarity", 7, 0.8, -3.0
        AddPt "Polarity", 8, 0.85, -2.5
        AddPt "Polarity", 9, 0.9,-2.0
        AddPt "Polarity", 10, 0.95, -1.5
        AddPt "Polarity", 11, 1, -1.0
        AddPt "Polarity", 12, 1.05, -0.5
        AddPt "Polarity", 13, 1.1, 0
        AddPt "Polarity", 14, 1.3, 0

        addpt "Velocity", 0, 0,         1
        addpt "Velocity", 1, 0.16, 1.06
        addpt "Velocity", 2, 0.41,         1.05
        addpt "Velocity", 3, 0.53,         1'0.982
        addpt "Velocity", 4, 0.702, 0.968
        addpt "Velocity", 5, 0.95,  0.968
        addpt "Velocity", 6, 1.03,         0.945

        LF.Object = LeftFlipper        
        LF.EndPoint = EndPointLp
        RF.Object = RightFlipper
        RF.EndPoint = EndPointRp
'		LF2.Object = LeftFlipper1       
'       LF2.EndPoint = EndPointLp001
       'RF2.Object = RightFlipper2
       'RF2.EndPoint = EndPointRp2
		'LF3.Object = LeftFlipper3        
       'LF3.EndPoint = EndPointLp3
       'RF3.Object = RightFlipper3
       'RF3.EndPoint = EndPointRp3
End Sub


' Flipper trigger hit subs
Sub TriggerLF_Hit() : LF.Addball activeball : End Sub
Sub TriggerLF_UnHit() : LF.PolarityCorrect activeball : End Sub
Sub TriggerRF_Hit() : RF.Addball activeball : End Sub
Sub TriggerRF_UnHit() : RF.PolarityCorrect activeball : End Sub
'Sub TriggerLF2_Hit() : LF.Addball activeball : End Sub
'Sub TriggerLF2_UnHit() : LF.PolarityCorrect activeball : End Sub
'Sub TriggerRF2_Hit() : RF.Addball activeball : End Sub
'Sub TriggerRF2_UnHit() : RF.PolarityCorrect activeball : End Sub
'Sub TriggerLF3_Hit() : LF.Addball activeball : End Sub
'Sub TriggerLF3_UnHit() : LF.PolarityCorrect activeball : End Sub
'Sub TriggerRF3_Hit() : RF.Addball activeball : End Sub
'Sub TriggerRF3_UnHit() : RF.PolarityCorrect activeball : End Sub

'Methods:
'.TimeDelay - Delay before trigger shuts off automatically. Default = 80 (ms)
'.AddPoint - "Polarity", "Velocity", "Ycoef" coordinate points. Use one of these 3 strings, keep coordinates sequential. x = %position on the flipper, y = output
'.Object - set to flipper reference. Optional.
'.StartPoint - set start point coord. Unnecessary, if .object is used.

'Called with flipper - 
'ProcessBalls - catches ball data. 
' - OR - 
'.Fire - fires flipper.rotatetoend automatically + processballs. Requires .Object to be set to the flipper.


'******************************************************
'  FLIPPER CORRECTION FUNCTIONS
'******************************************************

Sub AddPt(aStr, idx, aX, aY)        'debugger wrapper for adjusting flipper script in-game
	dim a : a = Array(LF, RF, LF1, RF1, LF2, RF2)
	dim x : for each x in a
		x.addpoint aStr, idx, aX, aY
	Next
End Sub

Class FlipperPolarity
	Public DebugOn, Enabled
	Private FlipAt        'Timer variable (IE 'flip at 723,530ms...)
	Public TimeDelay        'delay before trigger turns off and polarity is disabled TODO set time!
	private Flipper, FlipperStart,FlipperEnd, FlipperEndY, LR, PartialFlipCoef
	Private Balls(20), balldata(20)

	dim PolarityIn, PolarityOut
	dim VelocityIn, VelocityOut
	dim YcoefIn, YcoefOut
	Public Sub Class_Initialize 
		redim PolarityIn(0) : redim PolarityOut(0) : redim VelocityIn(0) : redim VelocityOut(0) : redim YcoefIn(0) : redim YcoefOut(0)
		Enabled = True : TimeDelay = 50 : LR = 1:  dim x : for x = 0 to uBound(balls) : balls(x) = Empty : set Balldata(x) = new SpoofBall : next 
	End Sub

	Public Property let Object(aInput) : Set Flipper = aInput : StartPoint = Flipper.x : End Property
	Public Property Let StartPoint(aInput) : if IsObject(aInput) then FlipperStart = aInput.x else FlipperStart = aInput : end if : End Property
	Public Property Get StartPoint : StartPoint = FlipperStart : End Property
	Public Property Let EndPoint(aInput) : FlipperEnd = aInput.x: FlipperEndY = aInput.y: End Property
	Public Property Get EndPoint : EndPoint = FlipperEnd : End Property        
	Public Property Get EndPointY: EndPointY = FlipperEndY : End Property

	Public Sub AddPoint(aChooseArray, aIDX, aX, aY) 'Index #, X position, (in) y Position (out) 
		Select Case aChooseArray
			case "Polarity" : ShuffleArrays PolarityIn, PolarityOut, 1 : PolarityIn(aIDX) = aX : PolarityOut(aIDX) = aY : ShuffleArrays PolarityIn, PolarityOut, 0
			Case "Velocity" : ShuffleArrays VelocityIn, VelocityOut, 1 :VelocityIn(aIDX) = aX : VelocityOut(aIDX) = aY : ShuffleArrays VelocityIn, VelocityOut, 0
			Case "Ycoef" : ShuffleArrays YcoefIn, YcoefOut, 1 :YcoefIn(aIDX) = aX : YcoefOut(aIDX) = aY : ShuffleArrays YcoefIn, YcoefOut, 0
		End Select
		if gametime > 100 then Report aChooseArray
	End Sub 

	Public Sub Report(aChooseArray)         'debug, reports all coords in tbPL.text
		if not DebugOn then exit sub
		dim a1, a2 : Select Case aChooseArray
			case "Polarity" : a1 = PolarityIn : a2 = PolarityOut
			Case "Velocity" : a1 = VelocityIn : a2 = VelocityOut
			Case "Ycoef" : a1 = YcoefIn : a2 = YcoefOut 
				case else :tbpl.text = "wrong string" : exit sub
		End Select
		dim str, x : for x = 0 to uBound(a1) : str = str & aChooseArray & " x: " & round(a1(x),4) & ", " & round(a2(x),4) & vbnewline : next
		tbpl.text = str
	End Sub

'********Triggered by a ball hitting the flipper trigger area
	Public Sub AddBall(aBall) : dim x : for x = 0 to uBound(balls) : if IsEmpty(balls(x)) then set balls(x) = aBall : exit sub :end if : Next  : End Sub

	Private Sub RemoveBall(aBall)
		dim x : for x = 0 to uBound(balls)
			if TypeName(balls(x) ) = "IBall" then 
				if aBall.ID = Balls(x).ID Then
					balls(x) = Empty
					Balldata(x).Reset
				End If
			End If
		Next
	End Sub

'*********Used to rotate flipper since this is removed from the key down for the flippers
	Public Sub Fire() 
		Flipper.RotateToEnd
		processballs
	End Sub

	Public Property Get Pos 'returns % position a ball. For debug stuff.
		dim x : for x = 0 to uBound(balls)
			if not IsEmpty(balls(x) ) then
				pos = pSlope(Balls(x).x, FlipperStart, 0, FlipperEnd, 1)
			End If
		Next                
	End Property

	Public Sub ProcessBalls() 'save data of balls in flipper range
		FlipAt = GameTime
		dim x : for x = 0 to uBound(balls)
			if not IsEmpty(balls(x) ) then
				balldata(x).Data = balls(x)
			End If
		Next
		PartialFlipCoef = ((Flipper.StartAngle - Flipper.CurrentAngle) / (Flipper.StartAngle - Flipper.EndAngle))   '% of flipper swing
		PartialFlipCoef = abs(PartialFlipCoef-1)
	End Sub

'***********gameTime is a global variable of how long the game has progressed in ms
'***********This function lets the table know if the flipper has been fired
	Private Function FlipperOn() : if gameTime < FlipAt+TimeDelay then FlipperOn = True : End If : End Function        'Timer shutoff for polaritycorrect

'***********This is turned on when a ball leaves the flipper trigger area
	Public Sub PolarityCorrect(aBall)
		if FlipperOn() then 
			dim tmp, BallPos, x, IDX, Ycoef : Ycoef = 1

			'y safety Exit
			if aBall.VelY > -8 then 'ball going down
				RemoveBall aBall
				exit Sub
			end if

			'Find balldata. BallPos = % on Flipper
			for x = 0 to uBound(Balls)
				if aBall.id = BallData(x).id AND not isempty(BallData(x).id) then 
					idx = x
					BallPos = PSlope(BallData(x).x, FlipperStart, 0, FlipperEnd, 1)
					if ballpos > 0.65 then  Ycoef = LinearEnvelope(BallData(x).Y, YcoefIn, YcoefOut)                                'find safety coefficient 'ycoef' data
				end if
			Next

			If BallPos = 0 Then 'no ball data meaning the ball is entering and exiting pretty close to the same position, use current values.
				BallPos = PSlope(aBall.x, FlipperStart, 0, FlipperEnd, 1)
				if ballpos > 0.65 then  Ycoef = LinearEnvelope(aBall.Y, YcoefIn, YcoefOut)                                                'find safety coefficient 'ycoef' data
			End If

			'Velocity correction
			if not IsEmpty(VelocityIn(0) ) then
				Dim VelCoef
				VelCoef = LinearEnvelope(BallPos, VelocityIn, VelocityOut)

				if partialflipcoef < 1 then VelCoef = PSlope(partialflipcoef, 0, 1, 1, VelCoef)

				if Enabled then aBall.Velx = aBall.Velx*VelCoef
				if Enabled then aBall.Vely = aBall.Vely*VelCoef
			End If

			'Polarity Correction (optional now)
			if not IsEmpty(PolarityIn(0) ) then
				If StartPoint > EndPoint then LR = -1        'Reverse polarity if left flipper
				dim AddX : AddX = LinearEnvelope(BallPos, PolarityIn, PolarityOut) * LR

				if Enabled then aBall.VelX = aBall.VelX + 1 * (AddX*ycoef*PartialFlipcoef)
			End If
		End If
		RemoveBall aBall
	End Sub
End Class

'******************************************************
'  FLIPPER POLARITY AND RUBBER DAMPENER SUPPORTING FUNCTIONS 
'******************************************************

' Used for flipper correction and rubber dampeners
Sub ShuffleArray(ByRef aArray, byVal offset) 'shuffle 1d array
	dim x, aCount : aCount = 0
	redim a(uBound(aArray) )
	for x = 0 to uBound(aArray)        'Shuffle objects in a temp array
		if not IsEmpty(aArray(x) ) Then
			if IsObject(aArray(x)) then 
				Set a(aCount) = aArray(x)   'Set creates an object in VB
			Else
				a(aCount) = aArray(x)
			End If
			aCount = aCount + 1
		End If
	Next
	if offset < 0 then offset = 0
	redim aArray(aCount-1+offset)        'Resize original array
	for x = 0 to aCount-1                'set objects back into original array
		if IsObject(a(x)) then 
			Set aArray(x) = a(x)
		Else
			aArray(x) = a(x)
		End If
	Next
End Sub

' Used for flipper correction and rubber dampeners
'**********Takes in more than one array and passes them to ShuffleArray
Sub ShuffleArrays(aArray1, aArray2, offset)
	ShuffleArray aArray1, offset
	ShuffleArray aArray2, offset
End Sub

' Used for flipper correction, rubber dampeners, and drop targets
'**********Calculate ball speed as hypotenuse of velX/velY triangle
' OPT FIX 9: Cache COM reads, replace ^2 with x*x
Function BallSpeed(ball) 'Calculates the ball speed
	Dim vx, vy, vz : vx = ball.VelX : vy = ball.VelY : vz = ball.VelZ
	BallSpeed = SQR(vx*vx + vy*vy + vz*vz)
End Function

' Used for flipper correction and rubber dampeners
'**********Calculates the value of Y for an input x using the slope intercept equation
Function PSlope(Input, X1, Y1, X2, Y2)        'Set up line via two points, no clamping. Input X, output Y
	dim x, y, b, m : x = input : m = (Y2 - Y1) / (X2 - X1) : b = Y2 - m*X2
	Y = M*x+b
	PSlope = Y
End Function

' Used for flipper correction
Class spoofball 
	Public X, Y, Z, VelX, VelY, VelZ, ID, Mass, Radius 
	Public Property Let Data(aBall)
		With aBall
			x = .x : y = .y : z = .z : velx = .velx : vely = .vely : velz = .velz
			id = .ID : mass = .mass : radius = .radius
		end with
	End Property
	Public Sub Reset()
		x = Empty : y = Empty : z = Empty  : velx = Empty : vely = Empty : velz = Empty 
		id = Empty : mass = Empty : radius = Empty
	End Sub
End Class

' Used for flipper correction and rubber dampeners
'********Interpolates the value for areas between the low and upper bounds sent to it
Function LinearEnvelope(xInput, xKeyFrame, yLvl)
	dim y 'Y output
	dim L 'Line
	dim ii : for ii = 1 to uBound(xKeyFrame)        'find active line
		if xInput <= xKeyFrame(ii) then L = ii : exit for : end if
	Next
	if xInput > xKeyFrame(uBound(xKeyFrame) ) then L = uBound(xKeyFrame)        'catch line overrun
	Y = pSlope(xInput, xKeyFrame(L-1), yLvl(L-1), xKeyFrame(L), yLvl(L) )

	if xInput <= xKeyFrame(lBound(xKeyFrame) ) then Y = yLvl(lBound(xKeyFrame) )         'Clamp lower
	if xInput >= xKeyFrame(uBound(xKeyFrame) ) then Y = yLvl(uBound(xKeyFrame) )        'Clamp upper

	LinearEnvelope = Y
End Function


'******************************************************
'  FLIPPER TRICKS 
'******************************************************

RightFlipper.timerinterval=10   ' OPT FIX 1: was 1 (1000Hz → 100Hz). 100Hz is visually identical for flipper correction.
Rightflipper.timerenabled=True

sub RightFlipper_timer()
	FlipperTricks LeftFlipper, LFPress, LFCount, LFEndAngle, LFState
	FlipperTricks RightFlipper, RFPress, RFCount, RFEndAngle, RFState
	' OPT FIX 1: Press guard — skip FlipperNudge entirely when neither flipper is pressed.
	' Shared GetBalls — one call instead of two (one per FlipperNudge).
	If LFPress = 0 And RFPress = 0 Then Exit Sub
	Dim fBOT : fBOT = GetBalls
	FlipperNudge fBOT, RightFlipper, RFEndAngle, RFEOSNudge, LeftFlipper, LFEndAngle
	FlipperNudge fBOT, LeftFlipper, LFEndAngle, LFEOSNudge,  RightFlipper, RFEndAngle
end sub

Dim LFEOSNudge, RFEOSNudge

' OPT FIX 1+3: Accept shared BOT array as first param. Cache flipper currentangle into locals.
Sub FlipperNudge(BOT, Flipper1, Endangle1, EOSNudge1, Flipper2, EndAngle2)
	Dim b, bx, by
	Dim ca1 : ca1 = Flipper1.currentangle    ' OPT FIX 3: cache COM read

	If ca1 = Endangle1 and EOSNudge1 <> 1 Then
		EOSNudge1 = 1
		If Flipper2.currentangle = EndAngle2 Then 
			For b = 0 to Ubound(BOT)
				bx = BOT(b).x : by = BOT(b).y   ' OPT FIX 3: cache ball position
				If FlipperTrigger(bx, by, Flipper1) Then
					exit Sub
				end If
			Next
			For b = 0 to Ubound(BOT)
				bx = BOT(b).x : by = BOT(b).y
				If FlipperTrigger(bx, by, Flipper2) Then
					BOT(b).velx = BOT(b).velx / 1.3
					BOT(b).vely = BOT(b).vely - 0.5
				end If
			Next
		End If
	Else 
		If Abs(ca1) > Abs(EndAngle1) + 30 then EOSNudge1 = 0
	End If
End Sub

'*****************
' Maths
'*****************
'Dim PI: PI = 4*Atn(1)

' OPT FIX 7: Use pre-computed PIover180 instead of Pi/180
Function dSin(degrees)
	dsin = sin(degrees * PIover180)
End Function

Function dCos(degrees)
	dcos = cos(degrees * PIover180)
End Function

'Function Atn2(dy, dx)
'	If dx > 0 Then
'		Atn2 = Atn(dy / dx)
'	ElseIf dx < 0 Then
'		If dy = 0 Then 
'			Atn2 = pi
'		Else
'			Atn2 = Sgn(dy) * (pi - Atn(Abs(dy / dx)))
'		end if
'	ElseIf dx = 0 Then
'		if dy = 0 Then
'			Atn2 = 0
'		else
'			Atn2 = Sgn(dy) * pi / 2
'		end if
'	End If
'End Function

'*************************************************
'  Check ball distance from Flipper for Rem
'*************************************************

' OPT FIX 8: Replace ^2 with x*x (^2 = Exp(2*Log(x)) in VBS)
Function Distance(ax,ay,bx,by)
	Dim dx, dy : dx = ax - bx : dy = ay - by
	Distance = SQR(dx * dx + dy * dy)
End Function

Function DistancePL(px,py,ax,ay,bx,by) ' Distance between a point and a line where point is px,py
	DistancePL = ABS((by - ay)*px - (bx - ax) * py + bx*ay - by*ax)/Distance(ax,ay,bx,by)
End Function

' OPT FIX 7: Use pre-computed PIover180
Function Radians(Degrees)
	Radians = Degrees * PIover180
End Function

' OPT FIX 7: Use pre-computed d180overPI
Function AnglePP(ax,ay,bx,by)
	AnglePP = Atn2((by - ay),(bx - ax))*d180overPI
End Function

Function DistanceFromFlipper(ballx, bally, Flipper)
	DistanceFromFlipper = DistancePL(ballx, bally, Flipper.x, Flipper.y, Cos(Radians(Flipper.currentangle+90))+Flipper.x, Sin(Radians(Flipper.currentangle+90))+Flipper.y)
End Function

' OPT FIX 4: Cache all flipper COM properties into locals. Inline DistanceFromFlipper.
' Eliminates ~6-10 redundant COM reads per call. At 100Hz × 2 flippers × N balls: major savings.
Function FlipperTrigger(ballx, bally, Flipper)
	Dim fx, fy, fca, flen, frad, dfl, DiffAngle
	fx = Flipper.x : fy = Flipper.y
	fca = Flipper.currentangle : flen = Flipper.Length
	frad = (fca + 90) * PIover180
	DiffAngle = ABS(fca - Atn2((bally - fy),(ballx - fx))*d180overPI - 90)
	If DiffAngle > 180 Then DiffAngle = DiffAngle - 360
	dfl = DistancePL(ballx, bally, fx, fy, Cos(frad)+fx, Sin(frad)+fy)
	If dfl < 48 and DiffAngle <= 90 and Distance(ballx, bally, fx, fy) < flen Then
		FlipperTrigger = True
	Else
		FlipperTrigger = False
	End If        
End Function


'*************************************************
'  End - Check ball distance from Flipper for Rem
'*************************************************

dim LFPress, RFPress, LFCount, RFCount
dim LFState, RFState
dim EOST, EOSA,Frampup, FElasticity,FReturn
dim RFEndAngle, LFEndAngle, LF1EndAngle, RF1EndAngle

'Const FlipperCoilRampupMode = 1   	'0 = fast, 1 = medium, 2 = slow (tap passes should work)

LFState = 1
RFState = 1
EOST = leftflipper.eostorque			'End of Swing Torque
EOSA = leftflipper.eostorqueangle		'End of Swing Torque Angle
Frampup = LeftFlipper.rampup			'Flipper Stregth Ramp Up
FElasticity = LeftFlipper.elasticity	'Flipper Elasticity
FReturn = LeftFlipper.return			'Flipper Return Strength
'Const EOSTnew = 1 'EM's to late 80's
Const EOSTnew = 0.8 '90's and later
Const EOSAnew = 1
Const EOSRampup = 0
Dim SOSRampup
Select Case FlipperCoilRampupMode 		'determines strength of coil field at start of swing
	Case 0:
		SOSRampup = 2.5
	Case 1:
		SOSRampup = 6
	Case 2:
		SOSRampup = 8.5
End Select

Const LiveCatch = 16					'variable to check elapsed time
Const LiveElasticity = 0.45
Const SOSEM = 0.815						
'Const EOSReturn = 0.055  'EM's
'Const EOSReturn = 0.045  'late 70's to mid 80's
'Const EOSReturn = 0.035  'mid 80's to early 90's
Const EOSReturn = 0.025  'mid 90's and later

LFEndAngle = Leftflipper.endangle
RFEndAngle = RightFlipper.endangle

Sub FlipperActivate(Flipper, FlipperPress)
	FlipperPress = 1
	Flipper.Elasticity = FElasticity

	Flipper.eostorque = EOST         
	Flipper.eostorqueangle = EOSA         
End Sub

Sub FlipperDeactivate(Flipper, FlipperPress)
	FlipperPress = 0
	Flipper.eostorqueangle = EOSA
	Flipper.eostorque = EOST*EOSReturn/FReturn

	If Abs(Flipper.currentangle) <= Abs(Flipper.endangle) + 0.1 Then
		Dim b, BOT
		BOT = GetBalls
		' OPT FIX 15: Cache flipper position for Distance check
		Dim fx : fx = Flipper.x : Dim fy : fy = Flipper.y
		For b = 0 to UBound(BOT)
			If Distance(BOT(b).x, BOT(b).y, fx, fy) < 55 Then
				If BOT(b).vely >= -0.4 Then BOT(b).vely = -0.4
			End If
		Next
	End If
End Sub

Sub FlipperTricks (Flipper, FlipperPress, FCount, FEndAngle, FState) 
	' OPT FIX 2: Cache COM reads of startangle/currentangle into locals.
	' Eliminates ~10 COM reads per call (×2 flippers ×100Hz = ~2000 reads/sec).

	Dim sa : sa = Flipper.startangle
	Dim Dir : Dir = sa / Abs(sa)                     '-1 for Right Flipper
	Dim ca : ca = Abs(Flipper.currentangle)
	Dim absSa : absSa = Abs(sa)

	If ca > absSa - 0.05 Then  'If the flipper has started its swing, make it swing fast to nearly the end...
		If FState <> 1 Then
			Flipper.rampup = SOSRampup
			Flipper.endangle = FEndAngle - 3*Dir
			Flipper.Elasticity = FElasticity * SOSEM
			FCount = 0 
			FState = 1
		End If
	ElseIf ca <= Abs(Flipper.endangle) and FlipperPress = 1 then   'If the flipper is fully swung and the flipper button is pressed then
		if FCount = 0 Then FCount = GameTime

		If FState <> 2 Then
			Flipper.eostorqueangle = EOSAnew
			Flipper.eostorque = EOSTnew
			Flipper.rampup = EOSRampup                        
			Flipper.endangle = FEndAngle
			FState = 2
		End If
	Elseif ca > Abs(Flipper.endangle) + 0.01 and FlipperPress = 1 Then
		If FState <> 3 Then
			Flipper.eostorque = EOST
			Flipper.eostorqueangle = EOSA
			Flipper.rampup = Frampup
			Flipper.Elasticity = FElasticity
			FState = 3
		End If
	End If
End Sub

Const LiveDistanceMin = 30  'minimum distance in vp units from flipper base live catch dampening will occur
Const LiveDistanceMax = 114  'maximum distance in vp units from flipper base live catch dampening will occur (tip protection)

' OPT FIX 14: Cache Flipper.x/startangle. Eliminate redundant ABS(Flipper.x - ball.x) read.
Sub CheckLiveCatch(ball, Flipper, FCount, parm)
	Dim sa : sa = Flipper.startangle
	Dim Dir : Dir = sa / Abs(sa)
	Dim LiveCatchBounce
	Dim CatchTime : CatchTime = GameTime - FCount
	Dim flipDist : flipDist = ABS(Flipper.x - ball.x)

	if CatchTime <= LiveCatch and parm > 6 and flipDist > LiveDistanceMin and flipDist < LiveDistanceMax Then
		if CatchTime <= LiveCatch*0.5 Then
			LiveCatchBounce = 0
		else
			LiveCatchBounce = Abs((LiveCatch/2) - CatchTime)
		end If

		If LiveCatchBounce = 0 and ball.velx * Dir > 0 Then ball.velx = 0
		ball.vely = LiveCatchBounce * (32 / LiveCatch)
		ball.angmomx= 0
		ball.angmomy= 0
		ball.angmomz= 0
    Else
        If Abs(Flipper.currentangle) <= Abs(Flipper.endangle) + 1 Then FlippersD.Dampenf Activeball, parm
	End If
End Sub

'**************************************************
'        Flipper Collision Subs
'NOTE: COpy and overwrite collision sound from original collision subs over
'RandomSoundFlipper()' below
'**************************************************'

'Sub LeftFlipper_Collide(parm)
'    CheckLiveCatch Activeball, LeftFlipper, LFCount, parm
'    'RandomSoundFlipper() 'Remove this line if Fleep is integrated
'    LeftFlipperCollide parm   'This is the Fleep code
'End Sub
'
'Sub RightFlipper_Collide(parm)
'    CheckLiveCatch Activeball, RightFlipper, RFCount, parm
'    'RandomSoundFlipper() 'Remove this line if Fleep is integrated
'    RightFlipperCollide parm  'This is the Fleep code
'End Sub

' iaakki Rubberizer
sub Rubberizer(parm)
	if parm < 10 And parm > 2 And Abs(activeball.angmomz) < 10 then
		'debug.print "parm: " & parm & " momz: " & activeball.angmomz &" vely: "& activeball.vely
		activeball.angmomz = activeball.angmomz * 1.2
		activeball.vely = activeball.vely * 1.2
		'debug.print ">> newmomz: " & activeball.angmomz&" newvely: "& activeball.vely
	Elseif parm <= 2 and parm > 0.2 Then
		'debug.print "* parm: " & parm & " momz: " & activeball.angmomz &" vely: "& activeball.vely
		activeball.angmomz = activeball.angmomz * -1.1
		activeball.vely = activeball.vely * 1.4
		'debug.print "**** >> newmomz: " & activeball.angmomz&" newvely: "& activeball.vely
	end if
end sub

' apophis rubberizer
sub Rubberizer2(parm)
	if parm < 10 And parm > 2 And Abs(activeball.angmomz) < 10 then
		'debug.print "parm: " & parm & " momz: " & activeball.angmomz &" vely: "& activeball.vely
		activeball.angmomz = -activeball.angmomz * 2
		activeball.vely = activeball.vely * 1.2
		'debug.print ">> newmomz: " & activeball.angmomz&" newvely: "& activeball.vely
	Elseif parm <= 2 and parm > 0.2 Then
		'debug.print "* parm: " & parm & " momz: " & activeball.angmomz &" vely: "& activeball.vely
		activeball.angmomz = -activeball.angmomz * 0.5
		activeball.vely = activeball.vely * (1.2 + rnd(1)/3 )
		'debug.print "**** >> newmomz: " & activeball.angmomz&" newvely: "& activeball.vely
	end if
end sub

'******************************************************
'****  END FLIPPER CORRECTIONS
'******************************************************

'******************************************************
'****  PHYSICS DAMPENERS
'******************************************************
'
' These are data mined bounce curves, 
' dialed in with the in-game elasticity as much as possible to prevent angle / spin issues.
' Requires tracking ballspeed to calculate COR



Sub dPosts_Hit(idx) 
	RubbersD.dampen Activeball
Rubbers_hit (idx)
End Sub

Sub dSleeves_Hit(idx) 
	SleevesD.Dampen Activeball
End Sub

'*********This sets up the rubbers:
dim RubbersD : Set RubbersD = new Dampener     'Makes a Dampener Class Object 
RubbersD.name = "Rubbers"
RubbersD.debugOn = False        'shows info in textbox "TBPout"
RubbersD.Print = False        'debug, reports in debugger (in vel, out cor)
'cor bounce curve (linear)
'for best results, try to match in-game velocity as closely as possible to the desired curve
'RubbersD.addpoint 0, 0, 0.935        'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 0, 0, 1.1        'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 1, 3.77, 0.97
RubbersD.addpoint 2, 5.76, 0.967        'dont take this as gospel. if you can data mine rubber elasticitiy, please help!
RubbersD.addpoint 3, 15.84, 0.874
RubbersD.addpoint 4, 56, 0.64        'there's clamping so interpolate up to 56 at least

dim SleevesD : Set SleevesD = new Dampener        'this is just rubber but cut down to 85%...
SleevesD.name = "Sleeves"
SleevesD.debugOn = False        'shows info in textbox "TBPout"
SleevesD.Print = False        'debug, reports in debugger (in vel, out cor)
SleevesD.CopyCoef RubbersD, 0.85

'######################### Add new FlippersD Profile
'#########################    Adjust these values to increase or lessen the elasticity

dim FlippersD : Set FlippersD = new Dampener
FlippersD.name = "Flippers"
FlippersD.debugOn = False
FlippersD.Print = False	
FlippersD.addpoint 0, 0, 1.1	
FlippersD.addpoint 1, 3.77, 0.99
FlippersD.addpoint 2, 6, 0.99

Class Dampener
	Public Print, debugOn 'tbpOut.text
	public name, Threshold         'Minimum threshold. Useful for Flippers, which don't have a hit threshold.
	Public ModIn, ModOut
	Private Sub Class_Initialize : redim ModIn(0) : redim Modout(0): End Sub 

	Public Sub AddPoint(aIdx, aX, aY) 
		ShuffleArrays ModIn, ModOut, 1 : ModIn(aIDX) = aX : ModOut(aIDX) = aY : ShuffleArrays ModIn, ModOut, 0
		if gametime > 100 then Report
	End Sub

	public sub Dampen(aBall)
		if threshold then if BallSpeed(aBall) < threshold then exit sub end if end if
		dim RealCOR, DesiredCOR, str, coef
'       		 Uses the LinearEnvelope function to calculate the correction based upon where it's value sits in relation
'       		 to the addpoint parameters set above.  Basically interpolates values between set points in a linear fashion
		DesiredCor = LinearEnvelope(cor.ballvel(aBall.id), ModIn, ModOut )
'       		 Uses the function BallSpeed's value at the point of impact/the active ball's velocity which is constantly being updated	
'				 RealCor is always less than 1 
		RealCOR = BallSpeed(aBall) / (cor.ballvel(aBall.id)+0.0001)
'                Divides the desired CoR by the real COR to make a multiplier to correct velocity in x and y
		coef = desiredcor / realcor 
		if debugOn then str = name & " in vel:" & round(cor.ballvel(aBall.id),2 ) & vbnewline & "desired cor: " & round(desiredcor,4) & vbnewline & _
		"actual cor: " & round(realCOR,4) & vbnewline & "ballspeed coef: " & round(coef, 3) & vbnewline 
		if Print then debug.print Round(cor.ballvel(aBall.id),2) & ", " & round(desiredcor,3)
'             	  Applies the coef to x and y velocities
		aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef
		if debugOn then TBPout.text = str
	End Sub

	public sub Dampenf(aBall, parm) 'Rubberizer is handled here
		dim RealCOR, DesiredCOR, str, coef
		DesiredCor = LinearEnvelope(cor.ballvel(aBall.id), ModIn, ModOut )
		RealCOR = BallSpeed(aBall) / (cor.ballvel(aBall.id)+0.0001)
		coef = desiredcor / realcor 
		If abs(aball.velx) < 2 and aball.vely < 0 and aball.vely > -3.75 then 
			aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef
		End If
	End Sub

	Public Sub CopyCoef(aObj, aCoef) 'alternative addpoints, copy with coef
		dim x : for x = 0 to uBound(aObj.ModIn)
			addpoint x, aObj.ModIn(x), aObj.ModOut(x)*aCoef
		Next
	End Sub


	Public Sub Report()         'debug, reports all coords in tbPL.text
		if not debugOn then exit sub
		dim a1, a2 : a1 = ModIn : a2 = ModOut
		dim str, x : for x = 0 to uBound(a1) : str = str & x & ": " & round(a1(x),4) & ", " & round(a2(x),4) & vbnewline : next
		TBPout.text = str
	End Sub

End Class

'******************************************************
'  TRACK ALL BALL VELOCITIES
'  FOR RUBBER DAMPENER AND DROP TARGETS
'******************************************************
'*********CoR is Coefficient of Restitution defined as "how much of the kinetic energy remains for the objects 
'to rebound from one another vs. how much is lost as heat, or work done deforming the objects 
dim cor : set cor = New CoRTracker

' OPT FIX 10b: Pre-allocate arrays to tnob in Class_Initialize.
' Keep conditional ReDim for ball IDs > tnob (VPX .id auto-increments and
' can exceed tnob during multiball with ball create/destroy cycles).
' Eliminated: per-tick highestID scan loop (iterated all balls twice).
Class CoRTracker
	public ballvel, ballvelx, ballvely

	Private Sub Class_Initialize : redim ballvel(tnob) : redim ballvelx(tnob): redim ballvely(tnob) : End Sub 

	Public Sub Update()	'tracks in-ball-velocity
		dim b, AllBalls, bid : AllBalls = getballs
		If UBound(AllBalls) = -1 Then Exit Sub

		for each b in AllBalls
			bid = b.id
			If bid > UBound(ballvel) Then
				redim Preserve ballvel(bid)
				redim Preserve ballvelx(bid)
				redim Preserve ballvely(bid)
			End If
			ballvel(bid) = BallSpeed(b)
			ballvelx(bid) = b.velx
			ballvely(bid) = b.vely
		Next
	End Sub
End Class

' Note, cor.update must be called in a 10 ms timer. The example table uses the GameTimer for this purpose, but sometimes a dedicated timer call RDampen is used.
'
Sub RDampen_Timer
	Cor.Update
End Sub

'******************************************************
'  END NFOZZY PHYSICS
'******************************************************
'//////////////////////////////////////////////////////////////////////
'// Mechanic Sounds
'//////////////////////////////////////////////////////////////////////

' This part in the script is an entire block that is dedicated to the physics sound system.
' Various scripts and sounds that may be pretty generic and could suit other WPC systems, but the most are tailored specifically for the TOM table

' Many of the sounds in this package can be added by creating collections and adding the appropriate objects to those collections.  
' Create the following new collections:
' 	Metals (all metal objects, metal walls, metal posts, metal wire guides)
' 	Apron (the apron walls and plunger wall)
' 	Walls (all wood or plastic walls)
' 	Rollovers (wire rollover triggers, star triggers, or button triggers)
' 	Targets (standup or drop targets, these are hit sounds only ... you will want to add separate dropping sounds for drop targets)
' 	Gates (plate gates)
' 	GatesWire (wire gates)
' 	Rubbers (all rubbers including posts, sleeves, pegs, and bands)
' When creating the collections, make sure "Fire events for this collection" is checked.  
' You'll also need to make sure "Has Hit Event" is checked for each object placed in these collections (not necessary for gates and triggers).  
' Once the collections and objects are added, the save, close, and restart VPX.
'
' Tutorial vides by Apophis
' Part 1: 	https://youtu.be/PbE2kNiam3g
' Part 2: 	https://youtu.be/B5cm1Y8wQsk
' Part 3: 	https://youtu.be/eLhWyuYOyGg


'///////////////////////////////  SOUNDS PARAMETERS  //////////////////////////////
Dim GlobalSoundLevel, CoinSoundLevel, PlungerReleaseSoundLevel, PlungerPullSoundLevel, NudgeLeftSoundLevel
Dim NudgeRightSoundLevel, NudgeCenterSoundLevel, StartButtonSoundLevel, RollingSoundFactor

CoinSoundLevel = 1														'volume level; range [0, 1]
NudgeLeftSoundLevel = 1													'volume level; range [0, 1]
NudgeRightSoundLevel = 1												'volume level; range [0, 1]
NudgeCenterSoundLevel = 1												'volume level; range [0, 1]
StartButtonSoundLevel = 0.1												'volume level; range [0, 1]
PlungerReleaseSoundLevel = 0.8 '1 wjr											'volume level; range [0, 1]
PlungerPullSoundLevel = 1												'volume level; range [0, 1]
RollingSoundFactor = 1.1/5		

'///////////////////////-----Solenoids, Kickers and Flash Relays-----///////////////////////
Dim FlipperUpAttackMinimumSoundLevel, FlipperUpAttackMaximumSoundLevel, FlipperUpAttackLeftSoundLevel, FlipperUpAttackRightSoundLevel
Dim FlipperUpSoundLevel, FlipperDownSoundLevel, FlipperLeftHitParm, FlipperRightHitParm
Dim SlingshotSoundLevel, BumperSoundFactor, KnockerSoundLevel

FlipperUpAttackMinimumSoundLevel = 0.010           						'volume level; range [0, 1]
FlipperUpAttackMaximumSoundLevel = 0.635								'volume level; range [0, 1]
FlipperUpSoundLevel = 1.0                        						'volume level; range [0, 1]
FlipperDownSoundLevel = 0.45                      						'volume level; range [0, 1]
FlipperLeftHitParm = FlipperUpSoundLevel								'sound helper; not configurable
FlipperRightHitParm = FlipperUpSoundLevel								'sound helper; not configurable
SlingshotSoundLevel = 0.95												'volume level; range [0, 1]
BumperSoundFactor = 4.25												'volume multiplier; must not be zero
KnockerSoundLevel = 1 													'volume level; range [0, 1]

'///////////////////////-----Ball Drops, Bumps and Collisions-----///////////////////////
Dim RubberStrongSoundFactor, RubberWeakSoundFactor, RubberFlipperSoundFactor,BallWithBallCollisionSoundFactor
Dim BallBouncePlayfieldSoftFactor, BallBouncePlayfieldHardFactor, PlasticRampDropToPlayfieldSoundLevel, WireRampDropToPlayfieldSoundLevel, DelayedBallDropOnPlayfieldSoundLevel
Dim WallImpactSoundFactor, MetalImpactSoundFactor, SubwaySoundLevel, SubwayEntrySoundLevel, ScoopEntrySoundLevel
Dim SaucerLockSoundLevel, SaucerKickSoundLevel

BallWithBallCollisionSoundFactor = 3.2									'volume multiplier; must not be zero
RubberStrongSoundFactor = 0.055/5											'volume multiplier; must not be zero
RubberWeakSoundFactor = 0.075/5											'volume multiplier; must not be zero
RubberFlipperSoundFactor = 0.075/5										'volume multiplier; must not be zero
BallBouncePlayfieldSoftFactor = 0.025									'volume multiplier; must not be zero
BallBouncePlayfieldHardFactor = 0.025									'volume multiplier; must not be zero
DelayedBallDropOnPlayfieldSoundLevel = 0.8									'volume level; range [0, 1]
WallImpactSoundFactor = 0.075											'volume multiplier; must not be zero
MetalImpactSoundFactor = 0.075/3
SaucerLockSoundLevel = 0.8
SaucerKickSoundLevel = 0.8

'///////////////////////-----Gates, Spinners, Rollovers and Targets-----///////////////////////

Dim GateSoundLevel, TargetSoundFactor, SpinnerSoundLevel, RolloverSoundLevel, DTSoundLevel

GateSoundLevel = 0.5/5													'volume level; range [0, 1]
TargetSoundFactor = 0.0025 * 10											'volume multiplier; must not be zero
DTSoundLevel = 0.25														'volume multiplier; must not be zero
RolloverSoundLevel = 0.25                              					'volume level; range [0, 1]
SpinnerSoundLevel = 0.5                              					'volume level; range [0, 1]

'///////////////////////-----Ball Release, Guides and Drain-----///////////////////////
Dim DrainSoundLevel, BallReleaseSoundLevel, BottomArchBallGuideSoundFactor, FlipperBallGuideSoundFactor 

DrainSoundLevel = 0.8														'volume level; range [0, 1]
BallReleaseSoundLevel = 1												'volume level; range [0, 1]
BottomArchBallGuideSoundFactor = 0.2									'volume multiplier; must not be zero
FlipperBallGuideSoundFactor = 0.015										'volume multiplier; must not be zero

'///////////////////////-----Loops and Lanes-----///////////////////////
Dim ArchSoundFactor
ArchSoundFactor = 0.025/5													'volume multiplier; must not be zero


'/////////////////////////////  SOUND PLAYBACK FUNCTIONS  ////////////////////////////
'/////////////////////////////  POSITIONAL SOUND PLAYBACK METHODS  ////////////////////////////
' Positional sound playback methods will play a sound, depending on the X,Y position of the table element or depending on ActiveBall object position
' These are similar subroutines that are less complicated to use (e.g. simply use standard parameters for the PlaySound call)
' For surround setup - positional sound playback functions will fade between front and rear surround channels and pan between left and right channels
' For stereo setup - positional sound playback functions will only pan between left and right channels
' For mono setup - positional sound playback functions will not pan between left and right channels and will not fade between front and rear channels

' PlaySound full syntax - PlaySound(string, int loopcount, float volume, float pan, float randompitch, int pitch, bool useexisting, bool restart, float front_rear_fade)
' Note - These functions will not work (currently) for walls/slingshots as these do not feature a simple, single X,Y position
Sub PlaySoundAtLevelStatic(playsoundparams, aVol, tableobj)
	PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(tableobj), 0, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelExistingStatic(playsoundparams, aVol, tableobj)
	PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(tableobj), 0, 0, 1, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelStaticLoop(playsoundparams, aVol, tableobj)
	PlaySound playsoundparams, -1, aVol * VolumeDial, AudioPan(tableobj), 0, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelStaticRandomPitch(playsoundparams, aVol, randomPitch, tableobj)
	PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(tableobj), randomPitch, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelActiveBall(playsoundparams, aVol)
	PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(ActiveBall), 0, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtLevelExistingActiveBall(playsoundparams, aVol)
	PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(ActiveBall), 0, 0, 1, 0, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtLeveTimerActiveBall(playsoundparams, aVol, ballvariable)
	PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(ballvariable), 0, 0, 0, 0, AudioFade(ballvariable)
End Sub

Sub PlaySoundAtLevelTimerExistingActiveBall(playsoundparams, aVol, ballvariable)
	PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(ballvariable), 0, 0, 1, 0, AudioFade(ballvariable)
End Sub

Sub PlaySoundAtLevelRoll(playsoundparams, aVol, pitch)
	PlaySound playsoundparams, -1, aVol * VolumeDial, AudioPan(tableobj), randomPitch, 0, 0, 0, AudioFade(tableobj)
End Sub

' Previous Positional Sound Subs

Sub PlaySoundAt(soundname, tableobj)
	PlaySound soundname, 1, 1 * VolumeDial, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtVol(soundname, tableobj, aVol)
	PlaySound soundname, 1, aVol * VolumeDial, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBall(soundname)
	PlaySoundAt soundname, ActiveBall
End Sub

Sub PlaySoundAtBallVol (Soundname, aVol)
	Playsound soundname, 1,aVol * VolumeDial, AudioPan(ActiveBall), 0,0,0, 1, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtBallVolM (Soundname, aVol)
	Playsound soundname, 1,aVol * VolumeDial, AudioPan(ActiveBall), 0,0,0, 0, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtVolLoops(sound, tableobj, Vol, Loops)
	PlaySound sound, Loops, Vol * VolumeDial, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub


'******************************************************
'  Fleep  Supporting Ball & Sound Functions
'******************************************************

' OPT FIX 11: Pre-computed table dimension inverses for AudioFade/AudioPan
Dim InvTWHalf : InvTWHalf = 2 / tablewidth
Dim InvTHHalf : InvTHHalf = 2 / tableheight

' OPT FIX 6: Replace ^10 with chained multiplication (^10 = Exp(10*Log(x)) in VBS).
' Pre-compute InvTHHalf to eliminate division per call.
Function AudioFade(tableobj) ' Fades between front and back of the table
  Dim tmp
    tmp = tableobj.y * InvTHHalf - 1

	if tmp > 7000 Then
		tmp = 7000
	elseif tmp < -7000 Then
		tmp = -7000
	end if

    If tmp > 0 Then
		Dim t2,t4,t8 : t2=tmp*tmp : t4=t2*t2 : t8=t4*t4
		AudioFade = Csng(t8 * t2)
    Else
		Dim nt,n2,n4,n8 : nt=-tmp : n2=nt*nt : n4=n2*n2 : n8=n4*n4
        AudioFade = Csng(-(n8 * n2))
    End If
End Function

' OPT FIX 6: Same ^10 elimination + InvTWHalf pre-computed inverse.
Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table.
    Dim tmp
    tmp = tableobj.x * InvTWHalf - 1

	if tmp > 7000 Then
		tmp = 7000
	elseif tmp < -7000 Then
		tmp = -7000
	end if

    If tmp > 0 Then
		Dim t2,t4,t8 : t2=tmp*tmp : t4=t2*t2 : t8=t4*t4
        AudioPan = Csng(t8 * t2)
    Else
		Dim nt,n2,n4,n8 : nt=-tmp : n2=nt*nt : n4=n2*n2 : n8=n4*n4
        AudioPan = Csng(-(n8 * n2))
    End If
End Function

' OPT FIX 9: Cache COM reads into locals, replace ^2 with x*x
Function Vol(ball) ' Calculates the volume of the sound based on the ball speed
	Dim bv : bv = BallVel(ball)
	Vol = Csng(bv * bv)
End Function

Function Volz(ball) ' Calculates the volume of the sound based on the ball speed
	Dim vz : vz = ball.velz
	Volz = Csng(vz * vz)
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
	Pitch = BallVel(ball) * 20
End Function

' OPT FIX 9: Cache VelX/VelY, replace ^2 with x*x
Function BallVel(ball) 'Calculates the ball speed
	Dim vx, vy : vx = ball.VelX : vy = ball.VelY
	BallVel = INT(SQR(vx * vx + vy * vy))
End Function

' OPT FIX 9b: ^3 → bv*bv*bv, ^2 → bv*bv. Note: hot-path RollingTimer inlines these.
' These remain for non-hot-path callers.
Function VolPlayfieldRoll(ball)
	Dim bv : bv = BallVel(ball)
	VolPlayfieldRoll = RollingSoundFactor * 0.0005 * Csng(bv * bv * bv)
End Function

Function PitchPlayfieldRoll(ball)
	Dim bv : bv = BallVel(ball)
	PitchPlayfieldRoll = bv * bv * 15
End Function

Function RndInt(min, max)
	RndInt = Int(Rnd() * (max-min + 1) + min)' Sets a random number integer between min and max
End Function

Function RndNum(min, max)
	RndNum = Rnd() * (max-min) + min' Sets a random number between min and max
End Function

'/////////////////////////////  GENERAL SOUND SUBROUTINES  ////////////////////////////
Sub SoundStartButton()
	PlaySound ("Start_Button"), 0, StartButtonSoundLevel, 0, 0.25
End Sub

Sub SoundNudgeLeft()
	PlaySound ("Nudge_" & Int(Rnd*2)+1), 0, NudgeLeftSoundLevel * VolumeDial, -0.1, 0.25
End Sub

Sub SoundNudgeRight()
	PlaySound ("Nudge_" & Int(Rnd*2)+1), 0, NudgeRightSoundLevel * VolumeDial, 0.1, 0.25
End Sub

Sub SoundNudgeCenter()
	PlaySound ("Nudge_" & Int(Rnd*2)+1), 0, NudgeCenterSoundLevel * VolumeDial, 0, 0.25
End Sub


Sub SoundPlungerPull()
	PlaySoundAtLevelStatic ("Plunger_Pull_1"), PlungerPullSoundLevel, Plunger
End Sub

Sub SoundPlungerReleaseBall()
	PlaySoundAtLevelStatic ("Plunger_Release_Ball"), PlungerReleaseSoundLevel, Plunger	
End Sub

Sub SoundPlungerReleaseNoBall()
	PlaySoundAtLevelStatic ("Plunger_Release_No_Ball"), PlungerReleaseSoundLevel, Plunger
End Sub


'/////////////////////////////  KNOCKER SOLENOID  ////////////////////////////
Sub KnockerSolenoid()
	PlaySoundAtLevelStatic SoundFX("Knocker_1",DOFKnocker), KnockerSoundLevel, KnockerPosition
End Sub

'/////////////////////////////  DRAIN SOUNDS  ////////////////////////////
Sub RandomSoundDrain(drainswitch)
	PlaySoundAtLevelStatic ("Drain_" & Int(Rnd*11)+1), DrainSoundLevel, drainswitch
End Sub

'/////////////////////////////  TROUGH BALL RELEASE SOLENOID SOUNDS  ////////////////////////////

Sub RandomSoundBallRelease(drainswitch)
	PlaySoundAtLevelStatic SoundFX("BallRelease" & Int(Rnd*7)+1,DOFContactors), BallReleaseSoundLevel, drainswitch
End Sub

'/////////////////////////////  SLINGSHOT SOLENOID SOUNDS  ////////////////////////////
Sub RandomSoundSlingshotLeft(sling)
	PlaySoundAtLevelStatic SoundFX("Sling_L" & Int(Rnd*10)+1,DOFContactors), SlingshotSoundLevel, Sling
End Sub

Sub RandomSoundSlingshotRight(sling)
	PlaySoundAtLevelStatic SoundFX("Sling_R" & Int(Rnd*8)+1,DOFContactors), SlingshotSoundLevel, Sling
End Sub

'/////////////////////////////  BUMPER SOLENOID SOUNDS  ////////////////////////////
Sub RandomSoundBumperTop(Bump)
	PlaySoundAtLevelStatic SoundFX("Bumpers_Top_" & Int(Rnd*5)+1,DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
End Sub

Sub RandomSoundBumperMiddle(Bump)
	PlaySoundAtLevelStatic SoundFX("Bumpers_Middle_" & Int(Rnd*5)+1,DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
End Sub

Sub RandomSoundBumperBottom(Bump)
	PlaySoundAtLevelStatic SoundFX("Bumpers_Bottom_" & Int(Rnd*5)+1,DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
End Sub

'/////////////////////////////  SPINNER SOUNDS  ////////////////////////////
Sub SoundSpinner(spinnerswitch)
	PlaySoundAtLevelStatic ("Spinner"), SpinnerSoundLevel, spinnerswitch
End Sub


'/////////////////////////////  FLIPPER BATS SOUND SUBROUTINES  ////////////////////////////
'/////////////////////////////  FLIPPER BATS SOLENOID ATTACK SOUND  ////////////////////////////
Sub SoundFlipperUpAttackLeft(flipper)
	FlipperUpAttackLeftSoundLevel = RndNum(FlipperUpAttackMinimumSoundLevel, FlipperUpAttackMaximumSoundLevel)
	PlaySoundAtLevelStatic ("Flipper_Attack-L01"), FlipperUpAttackLeftSoundLevel, flipper
End Sub

Sub SoundFlipperUpAttackRight(flipper)
	FlipperUpAttackRightSoundLevel = RndNum(FlipperUpAttackMinimumSoundLevel, FlipperUpAttackMaximumSoundLevel)
	PlaySoundAtLevelStatic ("Flipper_Attack-R01"), FlipperUpAttackLeftSoundLevel, flipper
End Sub

'/////////////////////////////  FLIPPER BATS SOLENOID CORE SOUND  ////////////////////////////
Sub RandomSoundFlipperUpLeft(flipper)
	PlaySoundAtLevelStatic SoundFX("Flipper_L0" & Int(Rnd*9)+1,DOFFlippers), FlipperLeftHitParm, Flipper
End Sub

Sub RandomSoundFlipperUpRight(flipper)
	PlaySoundAtLevelStatic SoundFX("Flipper_R0" & Int(Rnd*9)+1,DOFFlippers), FlipperRightHitParm, Flipper
End Sub

Sub RandomSoundReflipUpLeft(flipper)
	PlaySoundAtLevelStatic SoundFX("Flipper_ReFlip_L0" & Int(Rnd*3)+1,DOFFlippers), (RndNum(0.8, 1))*FlipperUpSoundLevel, Flipper
End Sub

Sub RandomSoundReflipUpRight(flipper)
	PlaySoundAtLevelStatic SoundFX("Flipper_ReFlip_R0" & Int(Rnd*3)+1,DOFFlippers), (RndNum(0.8, 1))*FlipperUpSoundLevel, Flipper
End Sub

Sub RandomSoundFlipperDownLeft(flipper)
	PlaySoundAtLevelStatic SoundFX("Flipper_Left_Down_" & Int(Rnd*7)+1,DOFFlippers), FlipperDownSoundLevel, Flipper
End Sub

Sub RandomSoundFlipperDownRight(flipper)
	PlaySoundAtLevelStatic SoundFX("Flipper_Right_Down_" & Int(Rnd*8)+1,DOFFlippers), FlipperDownSoundLevel, Flipper
End Sub

'/////////////////////////////  FLIPPER BATS BALL COLLIDE SOUND  ////////////////////////////

Sub LeftFlipperCollide(parm)
	FlipperLeftHitParm = parm/10
	If FlipperLeftHitParm > 1 Then
		FlipperLeftHitParm = 1
	End If
	FlipperLeftHitParm = FlipperUpSoundLevel * FlipperLeftHitParm
	RandomSoundRubberFlipper(parm)
	
End Sub

Sub RightFlipperCollide(parm)
	FlipperRightHitParm = parm/10
	If FlipperRightHitParm > 1 Then
		FlipperRightHitParm = 1
	End If
	FlipperRightHitParm = FlipperUpSoundLevel * FlipperRightHitParm
	RandomSoundRubberFlipper(parm)
	
End Sub

Sub RandomSoundRubberFlipper(parm)
	PlaySoundAtLevelActiveBall ("Flipper_Rubber_" & Int(Rnd*7)+1), parm  * RubberFlipperSoundFactor
End Sub

'/////////////////////////////  ROLLOVER SOUNDS  ////////////////////////////
Sub RandomSoundRollover()
	PlaySoundAtLevelActiveBall ("Rollover_" & Int(Rnd*4)+1), RolloverSoundLevel
End Sub

Sub Rollovers_Hit(idx)
	RandomSoundRollover
End Sub

'/////////////////////////////  VARIOUS PLAYFIELD SOUND SUBROUTINES  ////////////////////////////
'/////////////////////////////  RUBBERS AND POSTS  ////////////////////////////
'/////////////////////////////  RUBBERS - EVENTS  ////////////////////////////
Sub Rubbers_Hit(idx)
	dim finalspeed
	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
	If finalspeed > 5 then		
		RandomSoundRubberStrong 1
	End if
	If finalspeed <= 5 then
		RandomSoundRubberWeak()
	End If	
End Sub

'/////////////////////////////  RUBBERS AND POSTS - STRONG IMPACTS  ////////////////////////////
Sub RandomSoundRubberStrong(voladj)
	Select Case Int(Rnd*10)+1
		Case 1 : PlaySoundAtLevelActiveBall ("Rubber_Strong_1"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
		Case 2 : PlaySoundAtLevelActiveBall ("Rubber_Strong_2"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
		Case 3 : PlaySoundAtLevelActiveBall ("Rubber_Strong_3"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
		Case 4 : PlaySoundAtLevelActiveBall ("Rubber_Strong_4"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
		Case 5 : PlaySoundAtLevelActiveBall ("Rubber_Strong_5"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
		Case 6 : PlaySoundAtLevelActiveBall ("Rubber_Strong_6"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
		Case 7 : PlaySoundAtLevelActiveBall ("Rubber_Strong_7"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
		Case 8 : PlaySoundAtLevelActiveBall ("Rubber_Strong_8"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
		Case 9 : PlaySoundAtLevelActiveBall ("Rubber_Strong_9"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
		Case 10 : PlaySoundAtLevelActiveBall ("Rubber_1_Hard"), Vol(ActiveBall) * RubberStrongSoundFactor * 0.6*voladj
	End Select
End Sub

'/////////////////////////////  RUBBERS AND POSTS - WEAK IMPACTS  ////////////////////////////
Sub RandomSoundRubberWeak()
	PlaySoundAtLevelActiveBall ("Rubber_" & Int(Rnd*9)+1), Vol(ActiveBall) * RubberWeakSoundFactor
End Sub

'/////////////////////////////  WALL IMPACTS  ////////////////////////////
Sub Walls_Hit(idx)
	RandomSoundWall()      
End Sub

Sub RandomSoundWall()
	dim finalspeed
	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
	If finalspeed > 16 then 
		Select Case Int(Rnd*5)+1
			Case 1 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_1"), Vol(ActiveBall) * WallImpactSoundFactor
			Case 2 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_2"), Vol(ActiveBall) * WallImpactSoundFactor
			Case 3 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_5"), Vol(ActiveBall) * WallImpactSoundFactor
			Case 4 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_7"), Vol(ActiveBall) * WallImpactSoundFactor
			Case 5 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_9"), Vol(ActiveBall) * WallImpactSoundFactor
		End Select
	End if
	If finalspeed >= 6 AND finalspeed <= 16 then
		Select Case Int(Rnd*4)+1
			Case 1 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_3"), Vol(ActiveBall) * WallImpactSoundFactor
			Case 2 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_4"), Vol(ActiveBall) * WallImpactSoundFactor
			Case 3 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_6"), Vol(ActiveBall) * WallImpactSoundFactor
			Case 4 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_8"), Vol(ActiveBall) * WallImpactSoundFactor
		End Select
	End If
	If finalspeed < 6 Then
		Select Case Int(Rnd*3)+1
			Case 1 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_4"), Vol(ActiveBall) * WallImpactSoundFactor
			Case 2 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_6"), Vol(ActiveBall) * WallImpactSoundFactor
			Case 3 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_8"), Vol(ActiveBall) * WallImpactSoundFactor
		End Select
	End if
End Sub

'/////////////////////////////  METAL TOUCH SOUNDS  ////////////////////////////
Sub RandomSoundMetal()
	PlaySoundAtLevelActiveBall ("Metal_Touch_" & Int(Rnd*13)+1), Vol(ActiveBall) * MetalImpactSoundFactor
End Sub

'/////////////////////////////  METAL - EVENTS  ////////////////////////////

Sub Metals_Hit (idx)
	RandomSoundMetal
End Sub

Sub ShooterDiverter_collide(idx)
	RandomSoundMetal
End Sub

'/////////////////////////////  BOTTOM ARCH BALL GUIDE  ////////////////////////////
'/////////////////////////////  BOTTOM ARCH BALL GUIDE - SOFT BOUNCES  ////////////////////////////
Sub RandomSoundBottomArchBallGuide()
	dim finalspeed
	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
	If finalspeed > 16 then 
		PlaySoundAtLevelActiveBall ("Apron_Bounce_"& Int(Rnd*2)+1), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
	End if
	If finalspeed >= 6 AND finalspeed <= 16 then
		Select Case Int(Rnd*2)+1
			Case 1 : PlaySoundAtLevelActiveBall ("Apron_Bounce_1"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
			Case 2 : PlaySoundAtLevelActiveBall ("Apron_Bounce_Soft_1"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
		End Select
	End If
	If finalspeed < 6 Then
		Select Case Int(Rnd*2)+1
			Case 1 : PlaySoundAtLevelActiveBall ("Apron_Bounce_Soft_1"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
			Case 2 : PlaySoundAtLevelActiveBall ("Apron_Medium_3"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
		End Select
	End if
End Sub

'/////////////////////////////  BOTTOM ARCH BALL GUIDE - HARD HITS  ////////////////////////////
Sub RandomSoundBottomArchBallGuideHardHit()
	PlaySoundAtLevelActiveBall ("Apron_Hard_Hit_" & Int(Rnd*3)+1), BottomArchBallGuideSoundFactor * 0.25
End Sub

Sub Apron_Hit (idx)
	If Abs(cor.ballvelx(activeball.id) < 4) and cor.ballvely(activeball.id) > 7 then
		RandomSoundBottomArchBallGuideHardHit()
	Else
		RandomSoundBottomArchBallGuide
	End If
End Sub

'/////////////////////////////  FLIPPER BALL GUIDE  ////////////////////////////
Sub RandomSoundFlipperBallGuide()
	dim finalspeed
	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
	If finalspeed > 16 then 
		Select Case Int(Rnd*2)+1
			Case 1 : PlaySoundAtLevelActiveBall ("Apron_Hard_1"),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
			Case 2 : PlaySoundAtLevelActiveBall ("Apron_Hard_2"),  Vol(ActiveBall) * 0.8 * FlipperBallGuideSoundFactor
		End Select
	End if
	If finalspeed >= 6 AND finalspeed <= 16 then
		PlaySoundAtLevelActiveBall ("Apron_Medium_" & Int(Rnd*3)+1),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
	End If
	If finalspeed < 6 Then
		PlaySoundAtLevelActiveBall ("Apron_Soft_" & Int(Rnd*7)+1),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
	End If
End Sub

'/////////////////////////////  TARGET HIT SOUNDS  ////////////////////////////
Sub RandomSoundTargetHitStrong()
	PlaySoundAtLevelActiveBall SoundFX("Target_Hit_" & Int(Rnd*4)+5,DOFTargets), Vol(ActiveBall) * 0.45 * TargetSoundFactor
End Sub

Sub RandomSoundTargetHitWeak()		
	PlaySoundAtLevelActiveBall SoundFX("Target_Hit_" & Int(Rnd*4)+1,DOFTargets), Vol(ActiveBall) * TargetSoundFactor
End Sub

Sub PlayTargetSound()
	dim finalspeed
	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
	If finalspeed > 10 then
		RandomSoundTargetHitStrong()
		RandomSoundBallBouncePlayfieldSoft Activeball
	Else 
		RandomSoundTargetHitWeak()
	End If	
End Sub

Sub Targets_Hit (idx)
	PlayTargetSound	
End Sub

'/////////////////////////////  BALL BOUNCE SOUNDS  ////////////////////////////
Sub RandomSoundBallBouncePlayfieldSoft(aBall)
	Select Case Int(Rnd*9)+1
		Case 1 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_1"), volz(aBall) * BallBouncePlayfieldSoftFactor, aBall
		Case 2 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_2"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.5, aBall
		Case 3 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_3"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.8, aBall
		Case 4 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_4"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.5, aBall
		Case 5 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_5"), volz(aBall) * BallBouncePlayfieldSoftFactor, aBall
		Case 6 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_1"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.2, aBall
		Case 7 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_2"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.2, aBall
		Case 8 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_5"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.2, aBall
		Case 9 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_7"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.3, aBall
	End Select
End Sub

Sub RandomSoundBallBouncePlayfieldHard(aBall)
	PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_" & Int(Rnd*7)+1), volz(aBall) * BallBouncePlayfieldHardFactor, aBall
End Sub

'/////////////////////////////  DELAYED DROP - TO PLAYFIELD - SOUND  ////////////////////////////
Sub RandomSoundDelayedBallDropOnPlayfield(aBall)
	Select Case Int(Rnd*5)+1
		Case 1 : PlaySoundAtLevelStatic ("Ball_Drop_Playfield_1_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
		Case 2 : PlaySoundAtLevelStatic ("Ball_Drop_Playfield_2_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
		Case 3 : PlaySoundAtLevelStatic ("Ball_Drop_Playfield_3_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
		Case 4 : PlaySoundAtLevelStatic ("Ball_Drop_Playfield_4_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
		Case 5 : PlaySoundAtLevelStatic ("Ball_Drop_Playfield_5_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
	End Select
End Sub

'/////////////////////////////  BALL GATES AND BRACKET GATES SOUNDS  ////////////////////////////

Sub SoundPlayfieldGate()			
	PlaySoundAtLevelStatic ("Gate_FastTrigger_" & Int(Rnd*2)+1), GateSoundLevel, Activeball
End Sub

Sub SoundHeavyGate()
	PlaySoundAtLevelStatic ("Gate_2"), GateSoundLevel, Activeball
End Sub

Sub Gates_hit(idx)
	SoundHeavyGate
End Sub

Sub GatesWire_hit(idx)	
	SoundPlayfieldGate	
End Sub	

'/////////////////////////////  LEFT LANE ENTRANCE - SOUNDS  ////////////////////////////

Sub RandomSoundLeftArch()
	PlaySoundAtLevelActiveBall ("Arch_L" & Int(Rnd*4)+1), Vol(ActiveBall) * ArchSoundFactor
End Sub

Sub RandomSoundRightArch()
	PlaySoundAtLevelActiveBall ("Arch_R" & Int(Rnd*4)+1), Vol(ActiveBall) * ArchSoundFactor
End Sub


Sub Arch1_hit()
	If Activeball.velx > 1 Then SoundPlayfieldGate
	StopSound "Arch_L1"
	StopSound "Arch_L2"
	StopSound "Arch_L3"
	StopSound "Arch_L4"
End Sub

Sub Arch1_unhit()
	If activeball.velx < -8 Then
		RandomSoundRightArch
	End If
End Sub

Sub Arch2_hit()
	If Activeball.velx < 1 Then SoundPlayfieldGate
	StopSound "Arch_R1"
	StopSound "Arch_R2"
	StopSound "Arch_R3"
	StopSound "Arch_R4"
End Sub

Sub Arch2_unhit()
	If activeball.velx > 10 Then
		RandomSoundLeftArch
	End If
End Sub

'/////////////////////////////  SAUCERS (KICKER HOLES)  ////////////////////////////

Sub SoundSaucerLock()
	PlaySoundAtLevelStatic ("Saucer_Enter_" & Int(Rnd*2)+1), SaucerLockSoundLevel, Activeball
End Sub

Sub SoundSaucerKick(scenario, saucer)
	Select Case scenario
		Case 0: PlaySoundAtLevelStatic SoundFX("Saucer_Empty", DOFContactors), SaucerKickSoundLevel, saucer
		Case 1: PlaySoundAtLevelStatic SoundFX("Saucer_Kick", DOFContactors), SaucerKickSoundLevel, saucer
	End Select
End Sub

'/////////////////////////////  BALL COLLISION SOUND  ////////////////////////////
Sub OnBallBallCollision(ball1, ball2, velocity)
	Dim snd
	Select Case Int(Rnd*7)+1
		Case 1 : snd = "Ball_Collide_1"
		Case 2 : snd = "Ball_Collide_2"
		Case 3 : snd = "Ball_Collide_3"
		Case 4 : snd = "Ball_Collide_4"
		Case 5 : snd = "Ball_Collide_5"
		Case 6 : snd = "Ball_Collide_6"
		Case 7 : snd = "Ball_Collide_7"
	End Select

	' OPT FIX 16: velocity^2 → v*v
	Dim v : v = Csng(velocity)
	PlaySound (snd), 0, v * v / 200 * BallWithBallCollisionSoundFactor * VolumeDial, AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub


'/////////////////////////////////////////////////////////////////
'					End Mechanical Sounds
'/////////////////////////////////////////////////////////////////
'********************************************
'*   LUT Selector
'********************************************																																				"

dim textindex : textindex = 1
dim charobj(55), glyph(201)
InitDisplayText

Sub myChangeLut
	Table1.ColorGradeImage = luts(lutpos)
	DisplayText lutpos, luts(lutpos)
	vpmTimer.AddTimer 2000, "If lutpos = " & lutpos & " then for anr = 10 to 54 : charobj(anr).visible = 0 : next'"
End Sub

Sub InitDisplayText
	Dim anr
	For anr = 10 to 54 : set charobj(anr) = eval("text0" & anr) : charobj(anr).visible = 0 : Next
	For anr = 32 to 96 : glyph(anr) = anr : next
	For anr = 0 to 31 : glyph(anr) = 32 : next
	for anr = 97 to 122 : glyph(anr)  = anr - 32 : next
	for anr = 123 to 200 : glyph(anr) = 32 : next
End Sub

Sub DisplayText(nr, luttext)
	dim tekst, anr
	for anr = 10 to 54 : charobj(anr).imageA = 32 : charobj(anr).visible = 1 : next
	If nr > -1 then
		tekst = "lutpos:" & nr
		For anr = 1 to len(tekst) : charobj(43 + anr).imageA = glyph(asc(mid(tekst, anr, 1))) : Next
	End If
	For anr = 1 to len(luttext)
		charobj(9 + anr).imageA = glyph(asc(mid(luttext, anr, 1)))
		If nr = -1 Then
			charobj(9 + anr).y = 1500 + sin(((textindex * 4 + anr)/20)*3.14) * 100
			charobj(9 + anr).height = 150 + cos(((textindex * 4 + anr)/20)*3.14) * 100
		End If
	Next
End Sub

Sub SetLUT
	'Table1.ColorGradeImage = "LUT" & LUTset
	Table1.ColorGradeImage = luts(lutpos)
end sub 

Sub SaveLUT

	Dim FileObj
	Dim ScoreFile

	Set FileObj=CreateObject("Scripting.FileSystemObject")
	If Not FileObj.FolderExists(UserDirectory) then 
		Exit Sub
	End if

	if lutpos = "" then lutpos = 0 'failsafe

	Set ScoreFile=FileObj.CreateTextFile(UserDirectory & "SLUT_" & cGameName & ".txt",True)
	ScoreFile.WriteLine lutpos 'la réf dans le txt
	Set ScoreFile=Nothing
	Set FileObj=Nothing

End Sub

Sub LoadLUT

	Dim FileObj, ScoreFile, TextStr
	dim rLine

	Set FileObj=CreateObject("Scripting.FileSystemObject")
	If Not FileObj.FolderExists(UserDirectory) then 
		lutpos=0
		Exit Sub
	End if
	If Not FileObj.FileExists(UserDirectory & "SLUT_" & cGameName & ".txt") then
		lutpos=0
		Exit Sub
	End if
	Set ScoreFile=FileObj.GetFile(UserDirectory & "SLUT_" & cGameName & ".txt")
	Set TextStr=ScoreFile.OpenAsTextStream(1,0)
		If (TextStr.AtEndOfStream=True) then
			Exit Sub
		End if
		rLine = TextStr.ReadLine
		If rLine = "" then
			lutpos=0
			Exit Sub
		End if
		lutpos = int (rLine) 
		Set ScoreFile = Nothing
	    Set FileObj = Nothing
End Sub
'********************************************	
'***************************************************************
'****  VPW DYNAMIC BALL SHADOWS by Iakki, Apophis, and Wylte
'***************************************************************

'****** INSTRUCTIONS please read ******

'****** Part A:  Table Elements ******
'
' Import the "bsrtx7" and "ballshadow" images
' Import the shadow materials file (3 sets included) (you can also export the 3 sets from this table to create the same file)
' Copy in the BallShadowA flasher set and the sets of primitives named BallShadow#, RtxBallShadow#, and RtxBall2Shadow#
'	* with at least as many objects each as there can be balls, including locked balls
' Ensure you have a timer with a -1 interval that is always running

' Create a collection called DynamicSources that includes all light sources you want to cast ball shadows
'***These must be organized in order, so that lights that intersect on the table are adjacent in the collection***
'***If there are more than 3 lights that overlap in a playable area, exclude the less important lights***
' This is because the code will only project two shadows if they are coming from lights that are consecutive in the collection, and more than 3 will cause "jumping" between which shadows are drawn
' The easiest way to keep track of this is to start with the group on the right slingshot and move anticlockwise around the table
'	For example, if you use 6 lights: A & B on the left slingshot and C & D on the right, with E near A&B and F next to C&D, your collection would look like EBACDF
'
'G				H											^	E
'															^	B
'	A		 C												^	A
'	 B		D			your collection should look like	^	G		because E&B, B&A, etc. intersect; but B&D or E&F do not
'  E		  F												^	H
'															^	C
'															^	D
'															^	F
'		When selecting them, you'd shift+click in this order^^^^^

'****** End Part A:  Table Elements ******


'****** Part B:  Code and Functions ******

' *** Timer sub
' The "DynamicBSUpdate" sub should be called by a timer with an interval of -1 (framerate)
Sub FrameTimer_Timer()
	If DynamicBallShadowsOn Or AmbientBallShadowOn Then DynamicBSUpdate 'update ball shadows
End Sub

' *** These are usually defined elsewhere (ballrolling), but activate here if necessary
'Const tnob = 10 ' total number of balls
'Const lob = 0	'locked balls on start; might need some fiddling depending on how your locked balls are done
'Dim tablewidth: tablewidth = Table1.width
'Dim tableheight: tableheight = Table1.height

' *** User Options - Uncomment here or move to top
'----- Shadow Options -----
'Const DynamicBallShadowsOn = 1		'0 = no dynamic ball shadow ("triangles" near slings and such), 1 = enable dynamic ball shadow
'Const AmbientBallShadowOn = 1		'0 = Static shadow under ball ("flasher" image, like JP's)
'									'1 = Moving ball shadow ("primitive" object, like ninuzzu's) - This is the only one that shows up on the pf when in ramps and fades when close to lights!
'									'2 = flasher image shadow, but it moves like ninuzzu's

Const fovY					= 0		'Offset y position under ball to account for layback or inclination (more pronounced need further back)
Const DynamicBSFactor 		= 0.95	'0 to 1, higher is darker
Const AmbientBSFactor 		= 0.97	'0 to 1, higher is darker
Const AmbientMovement		= 2		'1 to 4, higher means more movement as the ball moves left and right
Const Wideness				= 20	'Sets how wide the dynamic ball shadows can get (20 +5 thinness should be most realistic for a 50 unit ball)
Const Thinness				= 5		'Sets minimum as ball moves away from source

' OPT FIX 11: Now that shadow Consts are declared, compute derived values.
BS_dAM = BallSize / AmbientMovement
DynBSFactor2 = DynamicBSFactor * DynamicBSFactor
DynBSFactor3 = DynamicBSFactor * DynamicBSFactor * DynamicBSFactor


' *** This segment goes within the RollingUpdate sub, so that if Ambient...=0 and Dynamic...=0 the entire DynamicBSUpdate sub can be skipped for max performance
'	' stop the sound of deleted balls
'	For b = UBound(BOT) + 1 to tnob
'		If AmbientBallShadowOn = 0 Then BallShadowA(b).visible = 0
'		...rolling(b) = False
'		...StopSound("BallRoll_" & b)
'	Next
'
'...rolling and drop sounds...

'		If DropCount(b) < 5 Then
'			DropCount(b) = DropCount(b) + 1
'		End If
'
'		' "Static" Ball Shadows
'		If AmbientBallShadowOn = 0 Then
'			If BOT(b).Z > 30 Then
'				BallShadowA(b).height=BOT(b).z - BallSize/4		'This is technically 1/4 of the ball "above" the ramp, but it keeps it from clipping
'			Else
'				BallShadowA(b).height=BOT(b).z - BallSize/2 + 5
'			End If
'			BallShadowA(b).Y = BOT(b).Y + Ballsize/5 + fovY
'			BallShadowA(b).X = BOT(b).X
'			BallShadowA(b).visible = 1
'		End If

' *** Required Functions, enable these if they are not already present elswhere in your table
Function DistanceFast(x, y)
	dim ratio, ax, ay
	ax = abs(x)					'Get absolute value of each vector
	ay = abs(y)
	ratio = 1 / max(ax, ay)		'Create a ratio
	ratio = ratio * (1.29289 - (ax + ay) * ratio * 0.29289)
	if ratio > 0 then			'Quickly determine if it's worth using
		DistanceFast = 1/ratio
	Else
		DistanceFast = 0
	End if
end Function

Function max(a,b)
	if a > b then 
		max = a
	Else
		max = b
	end if
end Function

Dim PI: PI = 4*Atn(1)
' OPT FIX 7: Pre-computed trig constants — eliminates 2 divisions per trig call.
Dim PIover180 : PIover180 = PI / 180
Dim d180overPI : d180overPI = 180 / PI

Function Atn2(dy, dx)
	If dx > 0 Then
		Atn2 = Atn(dy / dx)
	ElseIf dx < 0 Then
		If dy = 0 Then 
			Atn2 = pi
		Else
			Atn2 = Sgn(dy) * (pi - Atn(Abs(dy / dx)))
		end if
	ElseIf dx = 0 Then
		if dy = 0 Then
			Atn2 = 0
		else
			Atn2 = Sgn(dy) * pi / 2
		end if
	End If
End Function

'****** End Part B:  Code and Functions ******


'****** Part C:  The Magic ******
Dim sourcenames, currentShadowCount, DSSources(30), numberofsources, numberofsources_hold
sourcenames = Array ("","","","","","","","","","","","")
currentShadowCount = Array (0,0,0,0,0,0,0,0,0,0,0,0)

' *** Trim or extend these to match the number of balls/primitives/flashers on the table!
dim objrtx1(12), objrtx2(12)
dim objBallShadow(12)
Dim BallShadowA
BallShadowA = Array (BallShadowA0,BallShadowA1,BallShadowA2,BallShadowA3,BallShadowA4,BallShadowA5,BallShadowA6,BallShadowA7,BallShadowA8,BallShadowA9,BallShadowA10,BallShadowA11)

DynamicBSInit

sub DynamicBSInit()
	Dim iii, source

	for iii = 0 to tnob									'Prepares the shadow objects before play begins
		Set objrtx1(iii) = Eval("RtxBallShadow" & iii)
		objrtx1(iii).material = "RtxBallShadow" & iii
		objrtx1(iii).z = iii/1000 + 0.01
		objrtx1(iii).visible = 0

		Set objrtx2(iii) = Eval("RtxBall2Shadow" & iii)
		objrtx2(iii).material = "RtxBallShadow2_" & iii
		objrtx2(iii).z = (iii)/1000 + 0.02
		objrtx2(iii).visible = 0

		currentShadowCount(iii) = 0

		Set objBallShadow(iii) = Eval("BallShadow" & iii)
		objBallShadow(iii).material = "BallShadow" & iii
		UpdateMaterial objBallShadow(iii).material,1,0,0,0,0,0,AmbientBSFactor,RGB(0,0,0),0,0,False,True,0,0,0,0
		objBallShadow(iii).Z = iii/1000 + 0.04
		objBallShadow(iii).visible = 0

		BallShadowA(iii).Opacity = 100*AmbientBSFactor
		BallShadowA(iii).visible = 0
	Next

	iii = 0

	For Each Source in DynamicSources
		DSSources(iii) = Array(Source.x, Source.y)
		iii = iii + 1
	Next
	numberofsources = iii
	numberofsources_hold = iii
end sub


' OPT FIX 10: DynamicBSUpdate rewrite.
' - Cache BOT(s).X/Y/Z into local vars bx/by/bz at top of each ball iteration.
'   Eliminates 10-15 COM property reads per ball per frame.
' - Cache DSSources(iii)(0/1) into local sx/sy before inner source loop.
'   Eliminates double array dereference per light source.
' - Distance-squared gate: compute dx*dx+dy*dy and compare against falloffSq.
'   DistanceFast() only called for balls actually within range.
' - Pre-computed invFalloff (1/falloff) replaces per-shadow division with multiplication.
' - DynBSFactor2/DynBSFactor3 replace DynamicBSFactor^2 and ^3 per shadow.
' - Cache UBound(BOT) once at sub entry.
' - Use pre-computed BS_d10/BS_d4/BS_d5/BS_d2/TW_d2/BS_dAM throughout ambient shadow math.
Sub DynamicBSUpdate
	Dim falloff:	falloff = 150
	Dim falloffSq : falloffSq = falloff * falloff
	Dim invFalloff : invFalloff = 1 / falloff
	Dim ShadowOpacity, ShadowOpacity2 
	Dim s, Source, LSd, currentMat, AnotherSource, BOT, iii
	Dim bx, by, bz, sx, sy, dx, dy, dsq, asx, asy
	BOT = GetBalls
	Dim ubBot : ubBot = UBound(BOT)

	'Hide shadow of deleted balls
	For s = ubBot + 1 to tnob
		objrtx1(s).visible = 0
		objrtx2(s).visible = 0
		objBallShadow(s).visible = 0
		BallShadowA(s).visible = 0
	Next

	If ubBot < lob Then Exit Sub

'The Magic happens now
	For s = lob to ubBot
		bx = BOT(s).X : by = BOT(s).Y : bz = BOT(s).Z

' *** Normal "ambient light" ball shadow
		If AmbientBallShadowOn = 1 Then
			If bz > 30 Then
				If bx < TW_d2 Then
					objBallShadow(s).X = (bx - BS_d10 + ((bx - TW_d2) / BS_dAM)) + 5
				Else
					objBallShadow(s).X = (bx + BS_d10 + ((bx - TW_d2) / BS_dAM)) - 5
				End If
				objBallShadow(s).Y = by + BS_d10 + fovY
				objBallShadow(s).visible = 1

				BallShadowA(s).X = bx
				BallShadowA(s).Y = by + BS_d5 + fovY
				BallShadowA(s).height = bz - BS_d4
				BallShadowA(s).visible = 1
			Elseif bz <= 30 And bz > 20 Then
				objBallShadow(s).visible = 1
				If bx < TW_d2 Then
					objBallShadow(s).X = (bx - BS_d10 + ((bx - TW_d2) / BS_dAM)) + 5
				Else
					objBallShadow(s).X = (bx + BS_d10 + ((bx - TW_d2) / BS_dAM)) - 5
				End If
				objBallShadow(s).Y = by + fovY
				BallShadowA(s).visible = 0
			Else
				objBallShadow(s).visible = 0
				BallShadowA(s).visible = 0
			end if

		Elseif AmbientBallShadowOn = 2 Then
			If bz > 30 Then
				BallShadowA(s).X = bx
				BallShadowA(s).Y = by + BS_d5 + fovY
				BallShadowA(s).height = bz - BS_d4
				BallShadowA(s).visible = 1
			Elseif bz <= 30 And bz > 20 Then
				BallShadowA(s).visible = 1
				If bx < TW_d2 Then
					BallShadowA(s).X = (bx - BS_d10 + ((bx - TW_d2) / BS_dAM)) + 5
				Else
					BallShadowA(s).X = (bx + BS_d10 + ((bx - TW_d2) / BS_dAM)) - 5
				End If
				BallShadowA(s).Y = by + BS_d10 + fovY
				BallShadowA(s).height = bz - BS_d2 + 5
			Else
				BallShadowA(s).visible = 0
			End If
		End If

' *** Dynamic shadows
		If DynamicBallShadowsOn Then
			If bz < 30 Then
				For iii = 0 to numberofsources - 1 
					sx = DSSources(iii)(0) : sy = DSSources(iii)(1)
					dx = bx - sx : dy = by - sy
					dsq = dx*dx + dy*dy
					If dsq < falloffSq Then
						LSd = DistanceFast(dx, dy)
					Else
						LSd = falloff
					End If
					If LSd < falloff And gilvl > 0 Then
						currentShadowCount(s) = currentShadowCount(s) + 1
						if currentShadowCount(s) = 1 Then
							sourcenames(s) = iii
							currentMat = objrtx1(s).material
							objrtx2(s).visible = 0 : objrtx1(s).visible = 1 : objrtx1(s).X = bx : objrtx1(s).Y = by + fovY
							objrtx1(s).rotz = AnglePP(sx, sy, bx, by) + 90
							ShadowOpacity = (falloff-LSd)*invFalloff
							objrtx1(s).size_y = Wideness*ShadowOpacity+Thinness
							UpdateMaterial currentMat,1,0,0,0,0,0,ShadowOpacity*DynBSFactor2,RGB(0,0,0),0,0,False,True,0,0,0,0
							If AmbientBallShadowOn = 1 Then
								currentMat = objBallShadow(s).material
								UpdateMaterial currentMat,1,0,0,0,0,0,AmbientBSFactor*(1-ShadowOpacity),RGB(0,0,0),0,0,False,True,0,0,0,0
							Else
								BallShadowA(s).Opacity = 100*AmbientBSFactor*(1-ShadowOpacity)
							End If

						Elseif currentShadowCount(s) = 2 Then
							currentMat = objrtx1(s).material
							AnotherSource = sourcenames(s)
							asx = DSSources(AnotherSource)(0) : asy = DSSources(AnotherSource)(1)
							objrtx1(s).visible = 1 : objrtx1(s).X = bx : objrtx1(s).Y = by + fovY
							objrtx1(s).rotz = AnglePP(asx, asy, bx, by) + 90
							ShadowOpacity = (falloff-DistanceFast((bx-asx),(by-asy)))*invFalloff
							objrtx1(s).size_y = Wideness*ShadowOpacity+Thinness
							UpdateMaterial currentMat,1,0,0,0,0,0,ShadowOpacity*DynBSFactor3,RGB(0,0,0),0,0,False,True,0,0,0,0

							currentMat = objrtx2(s).material
							objrtx2(s).visible = 1 : objrtx2(s).X = bx : objrtx2(s).Y = by + fovY
							objrtx2(s).rotz = AnglePP(sx, sy, bx, by) + 90
							ShadowOpacity2 = (falloff-LSd)*invFalloff
							objrtx2(s).size_y = Wideness*ShadowOpacity2+Thinness
							UpdateMaterial currentMat,1,0,0,0,0,0,ShadowOpacity2*DynBSFactor3,RGB(0,0,0),0,0,False,True,0,0,0,0
							If AmbientBallShadowOn = 1 Then
								currentMat = objBallShadow(s).material
								UpdateMaterial currentMat,1,0,0,0,0,0,AmbientBSFactor*(1-max(ShadowOpacity,ShadowOpacity2)),RGB(0,0,0),0,0,False,True,0,0,0,0
							Else
								BallShadowA(s).Opacity = 100*AmbientBSFactor*(1-max(ShadowOpacity,ShadowOpacity2))
							End If
						end if
					Else
						currentShadowCount(s) = 0
						BallShadowA(s).Opacity = 100*AmbientBSFactor
					End If
				Next
			Else
				objrtx2(s).visible = 0 : objrtx1(s).visible = 0
			End If
		End If
	Next
End Sub

'****************************************************************
'****  END VPW DYNAMIC BALL SHADOWS by Iakki, Apophis, and Wylte
'****************************************************************
'****************************************************************
'  GI
'****************************************************************
dim gilvl:gilvl = 1
Sub ToggleGI(Enabled)
	dim xx
	If enabled Then
		for each xx in GI:xx.state = 1:Next
		gilvl = 1
	Else
		for each xx in GI:xx.state = 0:Next	
		GITimer.enabled = True
		gilvl = 0
	End If
	Sound_GI_Relay enabled, bumper1
End Sub

Sub GITimer_Timer()
	me.enabled = False
	ToggleGI 1
End Sub
'****************************************************************
'  END GI
'****************************************************************

Sub Table1_exit()
	SaveLUT
	If B2SOn Then
		Controller.Pause = False
		Controller.Stop
	End If
End Sub

' VR PLUNGER ANIMATION
Sub TimerPlunger_Timer
	If PinCab_Primary_Plunger.Y < -150  then
		PinCab_Primary_Plunger.Y = PinCab_Primary_Plunger.Y + 5
	End If
End Sub

Sub TimerPlunger2_Timer
	PinCab_Primary_Plunger.Y = -250 + (5* Plunger.Position) - 20
End Sub
