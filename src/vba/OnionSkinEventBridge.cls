VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OnionSkinEventBridge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents App As Application

Private Sub App_PresentationNewSlide(ByVal Sld As Slide)
    On Error GoTo SafeExit

    HandleNewSlideForAnimationMode Sld

SafeExit:
End Sub
