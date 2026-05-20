Option Explicit

' =================================================================
'  ONION SKIN TOOLS - PowerPoint Animation Helper
'  Import this file in the VBA Editor (File > Import File)
'  Then assign the Subs to Quick Access Toolbar buttons
' =================================================================

' Shape name kept for compatibility with decks that already use this guide
Private Const GUIDE_NAME As String = "GuiaOnionSkin"
Private Const GUIDE_OPACITY As Single = 0.75   ' 0 = opaque | 1 = invisible
Private Const SETTINGS_APP_NAME As String = "OnionSkinTools"
Private Const SETTINGS_SECTION As String = "Preferences"
Private Const SETTINGS_ANIMATION_MODE As String = "AnimationMode"
Private Const SETTINGS_OUTPUT_FOLDER As String = "OutputFolder"
Private Const SETTINGS_BG_EVERY_NEW_SLIDE As String = "BackgroundEveryNewSlide"
Private Const ANIMATION_MODE_OFF As String = "Off"
Private Const ANIMATION_MODE_CLASSIC As String = "Classic"
Private Const ANIMATION_MODE_DUPLICATE As String = "Duplicate"
Private Const BG_TAG As String = "OnionBg"
Private Const BG_GROUP_TAG As String = "OnionBgGroupId"
Private Const BG_EDITABLE_COPY_TAG As String = "OnionBgEditableCopy"

' Hardcoded so the module compiles on PowerPoint builds where these
' enum members were added later than the rest of PpSaveAsFileType /
' PpMediaTaskStatus (e.g. ppSaveAsAnimatedGIF = 40 only exists in M365).
Private Const PPT_SAVE_AS_ANIMATED_GIF As Long = 40
Private Const PPT_SAVE_AS_OPEN_XML_PRESENTATION As Long = 24
Private Const PPT_SAVE_AS_OPEN_XML_PRESENTATION_MACRO_ENABLED As Long = 25
Private Const PPT_MEDIA_TASK_DONE As Long = 3
Private Const PPT_MEDIA_TASK_FAILED As Long = 4

#If VBA7 Then
    Private Declare PtrSafe Sub OST_Sleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub OST_Sleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)
#End If

Private gRibbon As IRibbonUI
Public gEventBridge As OnionSkinEventBridge


' -----------------------------------------------------------------
'  SHOW FLOATING PANEL  (run once to open the toolbar) -> Possible path for future development
' -----------------------------------------------------------------
' Sub ShowPanel()
'     OnionPanel.Show vbModeless
' End Sub


' -----------------------------------------------------------------
'  1. CREATE GUIDE - copies the previous slide as a ghost image
' -----------------------------------------------------------------
Sub CreateGuide()
    Dim idx As Integer
    On Error Resume Next
    idx = ActiveWindow.View.Slide.SlideIndex
    On Error GoTo 0
    CreateGuideForSlideIndex idx, False
End Sub

Public Sub CreateGuideForSlideIndex(ByVal idx As Integer, Optional ByVal silent As Boolean = False)
    Dim sldCurrent       As Slide
    Dim sldPrevious      As Slide
    Dim shp              As Shape
    Dim hiddenOnPrev     As Collection
    Dim imgTemp          As String
    Dim w As Single, h As Single

    If idx = 0 Then
        If Not silent Then
            MsgBox "Could not determine the current slide." & vbCrLf & _
                   "Make sure you are in Normal view with a slide selected.", _
                   vbExclamation, "Onion Skin Tools"
        End If
        Exit Sub
    End If

    If idx <= 1 Then
        If Not silent Then
            MsgBox "This is the first slide - there is no previous slide to use as a guide.", _
                   vbExclamation, "Onion Skin Tools"
        End If
        Exit Sub
    End If

    Set sldCurrent = ActivePresentation.Slides(idx)
    Set sldPrevious = ActivePresentation.Slides(idx - 1)

    ' --- check if a guide already exists and ask for confirmation ---
    Dim shpCheck As Shape
    For Each shpCheck In sldCurrent.Shapes
        If shpCheck.Name = GUIDE_NAME Then
            If Not silent Then
                Dim resp As Integer
                resp = MsgBox("A guide already exists on this slide." & vbCrLf & "Replace it?", _
                              vbYesNo + vbQuestion, "Onion Skin Tools")
                If resp = vbNo Then Exit Sub
            End If
            On Error Resume Next
            shpCheck.Locked = msoFalse
            On Error GoTo 0
            shpCheck.Delete
            Exit For
        End If
    Next shpCheck

    ' --- export previous slide as a temporary image ---
    w = ActivePresentation.PageSetup.SlideWidth
    h = ActivePresentation.PageSetup.SlideHeight

    ' unique file name to avoid conflicts when multiple decks are open
    imgTemp = Environ("TEMP") & "\ost_" & idx & "_" & Format(Now, "hhmmss") & ".png"

    ' --- Hide internal guide/background layers on the previous slide before export ---
    ' Export captures all visible shapes; without this, old guides would stack and
    ' background shapes would be baked into the ghost even though they are copied
    ' forward as a solid layer.
    Set hiddenOnPrev = HideInternalExportShapes(sldPrevious)

    sldPrevious.Export imgTemp, "PNG"

    ' --- Show hidden internal layers again right after export ---
    RestoreHiddenShapes hiddenOnPrev

    ' --- create ghost rectangle ---
    Set shp = sldCurrent.Shapes.AddShape(msoShapeRectangle, 0, 0, w, h)

    With shp
        .Name = GUIDE_NAME
        .Line.Visible = msoFalse
        .Fill.UserPicture imgTemp
        .Fill.Transparency = GUIDE_OPACITY
        .ZOrder msoSendToBack
        ' .Locked exists in Office 2013+; ignore error on older versions
        On Error Resume Next
        .Locked = msoTrue
        On Error GoTo 0
    End With

    ' --- delete temporary file ---
    On Error Resume Next
    Kill imgTemp
    On Error GoTo 0

End Sub


' -----------------------------------------------------------------
'  2a. CLEAR CURRENT GUIDE - removes the guide from the active slide only
' -----------------------------------------------------------------
Sub ClearCurrentGuide()
    Dim sld     As Slide
    Dim i       As Integer
    Dim count   As Integer
    Dim idx     As Integer

    On Error Resume Next
    idx = ActiveWindow.View.Slide.SlideIndex
    On Error GoTo 0

    If idx = 0 Then
        MsgBox "Could not determine the current slide." & vbCrLf & _
               "Make sure you are in Normal view with a slide selected.", _
               vbExclamation, "Onion Skin Tools"
        Exit Sub
    End If

    Set sld = ActivePresentation.Slides(idx)
    count = 0

    For i = sld.Shapes.count To 1 Step -1
        If sld.Shapes(i).Name = GUIDE_NAME Then
            On Error Resume Next
            sld.Shapes(i).Locked = msoFalse
            On Error GoTo 0
            sld.Shapes(i).Delete
            count = count + 1
        End If
    Next i

    If count = 0 Then
        MsgBox "No guide found on the current slide.", _
               vbInformation, "Onion Skin Tools"
    End If
End Sub


' -----------------------------------------------------------------
'  2b. CLEAR GUIDES - removes all guides from the presentation
' -----------------------------------------------------------------
Sub ClearGuides()
    Dim sld     As Slide
    Dim i       As Integer
    Dim count   As Integer
    count = 0

    For Each sld In ActivePresentation.Slides
        For i = sld.Shapes.count To 1 Step -1
            If sld.Shapes(i).Name = GUIDE_NAME Then
                On Error Resume Next
                sld.Shapes(i).Locked = msoFalse
                On Error GoTo 0
                sld.Shapes(i).Delete
                count = count + 1
            End If
        Next i
    Next sld

    If count = 0 Then
        MsgBox "No guides were found in this presentation.", _
               vbInformation, "Onion Skin Tools"
    Else
        MsgBox count & " guide(s) removed successfully.", _
               vbInformation, "Onion Skin Tools"
    End If
End Sub


' -----------------------------------------------------------------
'  3. BACKGROUND SHAPES - persistent fixed scene elements
' -----------------------------------------------------------------
Private Function IsBackgroundShape(ByVal shp As Shape) As Boolean
    On Error Resume Next
    IsBackgroundShape = (shp.Tags(BG_TAG) = "1")
    On Error GoTo 0
End Function

Private Function GetShapeTag(ByVal shp As Shape, ByVal tagName As String) As String
    On Error Resume Next
    GetShapeTag = shp.Tags(tagName)
    On Error GoTo 0
End Function

Private Sub MarkShapeAsBackground(ByVal shp As Shape, ByVal groupId As String)
    On Error Resume Next
    shp.Tags.Delete BG_EDITABLE_COPY_TAG
    shp.Tags.Delete BG_TAG
    shp.Tags.Delete BG_GROUP_TAG
    shp.Tags.Add BG_TAG, "1"
    If groupId <> "" Then shp.Tags.Add BG_GROUP_TAG, groupId
    shp.Locked = msoTrue
    On Error GoTo 0
End Sub

Private Sub UnmarkShapeAsBackground(ByVal shp As Shape)
    On Error Resume Next
    shp.Tags.Delete BG_TAG
    shp.Tags.Delete BG_GROUP_TAG
    shp.Tags.Delete BG_EDITABLE_COPY_TAG
    shp.Locked = msoFalse
    On Error GoTo 0
End Sub

Private Function NewBackgroundGroupId() As String
    Randomize
    NewBackgroundGroupId = "bg_" & Format(Now, "yyyymmddhhmmss") & "_" & CStr(Int(Rnd() * 1000000))
End Function

Private Function GetSelectedShapes() As ShapeRange
    On Error Resume Next
    If ActiveWindow.Selection.Type = ppSelectionShapes Then
        Set GetSelectedShapes = ActiveWindow.Selection.ShapeRange
    End If
    On Error GoTo 0
End Function

Sub SetSelectedBackgroundShapes()
    Dim selectedShapes As ShapeRange
    Set selectedShapes = GetSelectedShapes()

    If selectedShapes Is Nothing Then
        MsgBox "Select one or more shapes first.", vbInformation, "Onion Skin Tools"
        Exit Sub
    End If

    Dim sld As Slide
    On Error Resume Next
    Set sld = ActiveWindow.View.Slide
    On Error GoTo 0

    If sld Is Nothing Then
        MsgBox "Could not determine the current slide.", vbExclamation, "Onion Skin Tools"
        Exit Sub
    End If

    Dim groupId As String
    groupId = NewBackgroundGroupId()

    Dim i As Long
    Dim pastedRange As ShapeRange
    For i = 1 To selectedShapes.Count
        selectedShapes(i).Copy
        Set pastedRange = sld.Shapes.Paste
        MarkShapeAsBackground pastedRange(1), groupId
        pastedRange(1).ZOrder msoSendToBack

        ' Keep the chosen figure as a normal editable shape; only the copy is background.
        UnmarkShapeAsBackground selectedShapes(i)
    Next i

    SendGuideToBack sld
End Sub

Public Function GetBackgroundEveryNewSlide() As Boolean
    ' When True (default), each new slide (Enter) receives a locked copy of the
    ' background shapes from the previous slide. When False, that automatic
    ' propagation is skipped (guide / duplicate behaviour still runs).
    Dim v As String
    v = GetSetting(SETTINGS_APP_NAME, SETTINGS_SECTION, SETTINGS_BG_EVERY_NEW_SLIDE, "1")
    GetBackgroundEveryNewSlide = (v = "1")
End Function

Public Sub SaveBackgroundEveryNewSlide(ByVal enabled As Boolean)
    SaveSetting SETTINGS_APP_NAME, SETTINGS_SECTION, SETTINGS_BG_EVERY_NEW_SLIDE, IIf(enabled, "1", "0")
    If Not gRibbon Is Nothing Then
        gRibbon.Invalidate
    End If
End Sub

Public Sub ToggleBackgroundEveryNewSlide()
    SaveBackgroundEveryNewSlide Not GetBackgroundEveryNewSlide()
End Sub

Sub RemoveSelectedBackgroundShapes()
    Dim selectedShapes As ShapeRange
    Set selectedShapes = GetSelectedShapes()

    If selectedShapes Is Nothing Then
        MsgBox "Select one or more shapes first.", vbInformation, "Onion Skin Tools"
        Exit Sub
    End If

    Dim i As Long
    For i = 1 To selectedShapes.Count
        UnmarkShapeAsBackground selectedShapes(i)
    Next i
End Sub

Sub ClearBackgroundShapesOnCurrentSlide()
    Dim idx As Integer
    On Error Resume Next
    idx = ActiveWindow.View.Slide.SlideIndex
    On Error GoTo 0

    If idx = 0 Then
        MsgBox "Could not determine the current slide.", vbExclamation, "Onion Skin Tools"
        Exit Sub
    End If

    ClearBackgroundShapesOnSlide ActivePresentation.Slides(idx)
End Sub

Sub ClearBackgroundShapesOnAllSlides()
    Dim sld As Slide
    For Each sld In ActivePresentation.Slides
        ClearBackgroundShapesOnSlide sld
    Next sld
End Sub

Private Sub ClearBackgroundShapesOnSlide(ByVal sld As Slide)
    Dim shp As Shape
    For Each shp In sld.Shapes
        If IsBackgroundShape(shp) Then
            UnmarkShapeAsBackground shp
        ElseIf GetShapeTag(shp, BG_EDITABLE_COPY_TAG) = "1" Then
            On Error Resume Next
            shp.Tags.Delete BG_EDITABLE_COPY_TAG
            On Error GoTo 0
        End If
    Next shp
End Sub

Private Function HideInternalExportShapes(ByVal sld As Slide) As Collection
    Dim hiddenShapes As Collection
    Set hiddenShapes = New Collection

    Dim shp As Shape
    For Each shp In sld.Shapes
        If shp.Name = GUIDE_NAME Or IsBackgroundShape(shp) Or GetShapeTag(shp, BG_EDITABLE_COPY_TAG) = "1" Then
            If shp.Visible <> msoFalse Then
                shp.Visible = msoFalse
                hiddenShapes.Add shp
            End If
        End If
    Next shp

    Set HideInternalExportShapes = hiddenShapes
End Function

Private Sub RestoreHiddenShapes(ByVal hiddenShapes As Collection)
    Dim item As Variant

    If hiddenShapes Is Nothing Then Exit Sub

    On Error Resume Next
    For Each item In hiddenShapes
        item.Visible = msoTrue
    Next item
    On Error GoTo 0
End Sub

Private Sub CopyBackgroundShapesToSlide(ByVal sourceSlide As Slide, ByVal targetSlide As Slide)
    Dim shp As Shape
    Dim pastedRange As ShapeRange

    For Each shp In sourceSlide.Shapes
        If IsBackgroundShape(shp) Then
            shp.Copy
            Set pastedRange = targetSlide.Shapes.Paste
            MarkShapeAsBackground pastedRange(1), GetShapeTag(shp, BG_GROUP_TAG)
            pastedRange(1).ZOrder msoSendToBack
        End If
    Next shp

    SendGuideToBack targetSlide
End Sub

Private Sub SendGuideToBack(ByVal sld As Slide)
    Dim shp As Shape
    On Error Resume Next
    Set shp = sld.Shapes(GUIDE_NAME)
    If Not shp Is Nothing Then shp.ZOrder msoSendToBack
    On Error GoTo 0
End Sub

Private Sub CopySlideShapesForDuplicateMode(ByVal sourceSlide As Slide, ByVal targetSlide As Slide)
    Dim shp As Shape
    Dim pastedRange As ShapeRange

    For Each shp In sourceSlide.Shapes
        If shp.Name <> GUIDE_NAME And GetShapeTag(shp, BG_EDITABLE_COPY_TAG) <> "1" Then
            shp.Copy
            Set pastedRange = targetSlide.Shapes.Paste

            ' Background shapes copied as part of Duplicate mode are the editable
            ' top-layer copy. The locked/tagged authoritative copy is added separately.
            If IsBackgroundShape(shp) Then
                UnmarkShapeAsBackground pastedRange(1)
                pastedRange(1).Tags.Add BG_EDITABLE_COPY_TAG, "1"
            End If
        End If
    Next shp
End Sub

Private Sub ClearShapesOnSlide(ByVal sld As Slide)
    Dim i As Long
    For i = sld.Shapes.Count To 1 Step -1
        On Error Resume Next
        sld.Shapes(i).Locked = msoFalse
        sld.Shapes(i).Delete
        On Error GoTo 0
    Next i
End Sub

Public Sub HandleNewSlideForAnimationMode(ByVal sldNew As Slide)
    Dim mode As String
    mode = GetAnimationMode()

    If mode = ANIMATION_MODE_OFF Then Exit Sub
    If sldNew.SlideIndex <= 1 Then Exit Sub

    Dim sldPrevious As Slide
    Set sldPrevious = ActivePresentation.Slides(sldNew.SlideIndex - 1)

    If mode = ANIMATION_MODE_DUPLICATE Then
        ClearShapesOnSlide sldNew
    End If

    CreateGuideForSlideIndex sldNew.SlideIndex, True
    If GetBackgroundEveryNewSlide() Then
        CopyBackgroundShapesToSlide sldPrevious, sldNew
    End If

    If mode = ANIMATION_MODE_DUPLICATE Then
        CopySlideShapesForDuplicateMode sldPrevious, sldNew
    End If
End Sub


' -----------------------------------------------------------------
'  4a. OUTPUT FOLDER SETTING
' -----------------------------------------------------------------
Public Function GetOutputFolder() As String
    Dim path As String
    path = GetSetting(SETTINGS_APP_NAME, SETTINGS_SECTION, SETTINGS_OUTPUT_FOLDER, "")

    ' Fall back to "next to the active presentation" or Desktop if the
    ' configured folder is empty / no longer exists.
    If path = "" Or Not FolderExists(path) Then
        On Error Resume Next
        If Not ActivePresentation Is Nothing Then
            If ActivePresentation.Path <> "" Then path = ActivePresentation.Path
        End If
        On Error GoTo 0
    End If

    If path = "" Then path = Environ("USERPROFILE") & "\Desktop"
    GetOutputFolder = path
End Function

Public Sub SaveOutputFolder(ByVal path As String)
    SaveSetting SETTINGS_APP_NAME, SETTINGS_SECTION, SETTINGS_OUTPUT_FOLDER, path
End Sub

Private Function FolderExists(ByVal path As String) As Boolean
    On Error Resume Next
    Dim attr As Long
    attr = GetAttr(path)
    FolderExists = (Err.Number = 0) And ((attr And vbDirectory) <> 0)
    Err.Clear
    On Error GoTo 0
End Function

Sub SetOutputFolder()
    Dim fd As FileDialog
    On Error Resume Next
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    On Error GoTo 0

    If fd Is Nothing Then
        MsgBox "Folder picker is not available in this PowerPoint build.", _
               vbExclamation, "Onion Skin Tools"
        Exit Sub
    End If

    fd.Title = "Choose default output folder for Onion Tools exports"
    fd.AllowMultiSelect = False

    Dim current As String
    current = GetSetting(SETTINGS_APP_NAME, SETTINGS_SECTION, SETTINGS_OUTPUT_FOLDER, "")
    If current <> "" And FolderExists(current) Then
        fd.InitialFileName = current & "\"
    End If

    If fd.Show = -1 Then
        Dim chosen As String
        chosen = fd.SelectedItems(1)
        SaveOutputFolder chosen
        MsgBox "Exports will now be saved to:" & vbCrLf & vbCrLf & chosen, _
               vbInformation, "Onion Skin Tools"
    End If
End Sub


' -----------------------------------------------------------------
'  3b. EXPORT GIF / EXPORT VIDEO
'      Both formats and both with/without guides go through the same
'      pipeline: temp .pptx copy -> (optional) hide guides -> render
'      -> reveal in Explorer. No native dialog, no popup, output goes
'      to the configured output folder.
' -----------------------------------------------------------------
Sub ExportGIF(ByVal includeGuides As Boolean)
    ExportViaCopy "gif", includeGuides
End Sub

Sub ExportVideo(ByVal includeGuides As Boolean)
    ExportViaCopy "video", includeGuides
End Sub

Private Function HideGuidesInPresentation(ByVal prs As Presentation) As Collection
    Dim hiddenGuides As Collection
    Set hiddenGuides = New Collection

    Dim sld As Slide
    Dim shp As Shape

    For Each sld In prs.Slides
        For Each shp In sld.Shapes
            If shp.Name = GUIDE_NAME Then
                If shp.Visible <> msoFalse Then
                    shp.Visible = msoFalse
                    hiddenGuides.Add shp
                End If
            End If
        Next shp
    Next sld

    Set HideGuidesInPresentation = hiddenGuides
End Function

Private Sub RestoreHiddenGuides(ByVal hiddenGuides As Collection)
    Dim item As Variant

    If hiddenGuides Is Nothing Then Exit Sub

    On Error Resume Next
    For Each item In hiddenGuides
        item.Visible = msoTrue
    Next item
    On Error GoTo 0
End Sub

Private Sub ExportViaCopy(ByVal outputFormat As String, ByVal includeGuides As Boolean)
    Dim originalPres As Presentation
    Set originalPres = ActivePresentation

    ' --- build output path ---
    Dim outputDir As String, baseName As String, suffix As String, ext As String
    outputDir = GetOutputFolder()
    If outputFormat = "gif" Then ext = ".gif" Else ext = ".mp4"
    suffix = IIf(includeGuides, "_with_guides", "_no_guides")

    If originalPres.Path <> "" Then
        baseName = Left(originalPres.Name, InStrRev(originalPres.Name, ".") - 1)
    Else
        baseName = "OnionExport"
    End If

    Dim outputPath As String
    outputPath = outputDir & "\" & baseName & suffix & ext

    ' --- working copy in TEMP ---
    Dim tempPath As String
    tempPath = Environ("TEMP") & "\OnionExport_" & Format(Now, "yyyymmddhhmmss") & ".pptx"

    Dim copyPres As Presentation
    Set copyPres = Nothing
    Dim hiddenGuides As Collection
    Set hiddenGuides = Nothing

    On Error GoTo CleanFail

    If Not includeGuides Then
        ' Hide guides in the real presentation only long enough to write
        ' the temp copy. This is much faster than opening a copy, deleting
        ' shapes, and saving again; the user's deck is restored immediately.
        Set hiddenGuides = HideGuidesInPresentation(originalPres)
    End If

    ' Macro-free temp copies are smaller and quicker to open/render. If
    ' PowerPoint refuses that conversion for a particular deck, fall back
    ' to a macro-enabled temp copy.
    Dim saveErr As Long
    On Error Resume Next
    originalPres.SaveCopyAs tempPath, PPT_SAVE_AS_OPEN_XML_PRESENTATION
    saveErr = Err.Number
    Err.Clear
    On Error GoTo CleanFail

    If saveErr <> 0 Then
        tempPath = Environ("TEMP") & "\OnionExport_" & Format(Now, "yyyymmddhhmmss") & ".pptm"
        originalPres.SaveCopyAs tempPath, PPT_SAVE_AS_OPEN_XML_PRESENTATION_MACRO_ENABLED
    End If

    RestoreHiddenGuides hiddenGuides
    Set hiddenGuides = Nothing

    ' Open the already-prepared copy without a window. The guide visibility
    ' is now baked into the temp file, so the render pipeline sees it without
    ' needing visible window state or custom ribbon callbacks.
    Set copyPres = Presentations.Open(tempPath, _
        ReadOnly:=msoFalse, _
        Untitled:=msoFalse, _
        WithWindow:=msoFalse)

    If outputFormat = "gif" Then
        ' Animated GIF export via SaveAs is synchronous on M365 builds.
        copyPres.SaveAs outputPath, PPT_SAVE_AS_ANIMATED_GIF
        copyPres.Saved = True
        copyPres.Close
        Set copyPres = Nothing

        On Error Resume Next: Kill tempPath: On Error GoTo 0

        On Error Resume Next: originalPres.Windows(1).Activate: On Error GoTo 0

        Shell "explorer.exe /select,""" & outputPath & """", vbNormalFocus
    Else
        ' Video: CreateVideo runs asynchronously. PowerPoint has no Application.OnTime
        ' (that is Excel-only), so we poll CreateVideoStatus until done or failed.
        copyPres.CreateVideo outputPath, True, 5, 1080, 30, 100

        Dim videoStatus As Long
        Dim probeErr As Long
        Do
            DoEvents
            OST_Sleep 250
            On Error Resume Next
            videoStatus = copyPres.CreateVideoStatus
            probeErr = Err.Number
            On Error GoTo CleanFail
            If probeErr <> 0 Then Err.Raise vbObjectError + 3, , "Could not read video export status."
            If videoStatus = PPT_MEDIA_TASK_DONE Then Exit Do
            If videoStatus = PPT_MEDIA_TASK_FAILED Then
                Err.Raise vbObjectError + 2, , "Video encoding failed."
            End If
        Loop

        copyPres.Saved = True
        copyPres.Close
        Set copyPres = Nothing

        On Error Resume Next: Kill tempPath: On Error GoTo 0

        On Error Resume Next: originalPres.Windows(1).Activate: On Error GoTo 0

        Shell "explorer.exe /select,""" & outputPath & """", vbNormalFocus
    End If

    Exit Sub

CleanFail:
    Dim errDesc As String
    errDesc = Err.Description

    On Error Resume Next
    RestoreHiddenGuides hiddenGuides
    If Not copyPres Is Nothing Then
        copyPres.Saved = True
        copyPres.Close
    End If
    Kill tempPath
    On Error Resume Next: originalPres.Windows(1).Activate: On Error GoTo 0
    On Error GoTo 0

    MsgBox "Export failed: " & errDesc, vbExclamation, "Onion Skin Tools"
End Sub

Sub RibbonCreateGuide(control As IRibbonControl)
    InitOnionEventBridge
    CreateGuide
End Sub

Sub RibbonClearCurrentGuide(control As IRibbonControl)
    InitOnionEventBridge
    ClearCurrentGuide
End Sub

Sub RibbonClearGuides(control As IRibbonControl)
    InitOnionEventBridge
    ClearGuides
End Sub

Sub RibbonExportGIFNoGuides(control As IRibbonControl)
    InitOnionEventBridge
    ExportGIF False
End Sub

Sub RibbonExportGIFWithGuides(control As IRibbonControl)
    InitOnionEventBridge
    ExportGIF True
End Sub

Sub RibbonExportVideoNoGuides(control As IRibbonControl)
    InitOnionEventBridge
    ExportVideo False
End Sub

Sub RibbonExportVideoWithGuides(control As IRibbonControl)
    InitOnionEventBridge
    ExportVideo True
End Sub

Sub RibbonSetOutputFolder(control As IRibbonControl)
    InitOnionEventBridge
    SetOutputFolder
End Sub

Sub RibbonSetAnimationModeOff(control As IRibbonControl, pressed As Boolean)
    SaveAnimationMode ANIMATION_MODE_OFF
End Sub

Sub RibbonSetAnimationModeClassic(control As IRibbonControl, pressed As Boolean)
    SaveAnimationMode ANIMATION_MODE_CLASSIC
End Sub

Sub RibbonSetAnimationModeDuplicate(control As IRibbonControl, pressed As Boolean)
    SaveAnimationMode ANIMATION_MODE_DUPLICATE
End Sub

Sub RibbonGetAnimationModeOff(control As IRibbonControl, ByRef returnedVal)
    returnedVal = (GetAnimationMode() = ANIMATION_MODE_OFF)
End Sub

Sub RibbonGetAnimationModeClassic(control As IRibbonControl, ByRef returnedVal)
    returnedVal = (GetAnimationMode() = ANIMATION_MODE_CLASSIC)
End Sub

Sub RibbonGetAnimationModeDuplicate(control As IRibbonControl, ByRef returnedVal)
    returnedVal = (GetAnimationMode() = ANIMATION_MODE_DUPLICATE)
End Sub

Sub RibbonSetBackground(control As IRibbonControl)
    InitOnionEventBridge
    SetSelectedBackgroundShapes
End Sub

Sub RibbonGetBgEveryNewSlideLabel(control As IRibbonControl, ByRef returnedVal)
    If GetBackgroundEveryNewSlide() Then
        returnedVal = "On every new slide (Enter) — On"
    Else
        returnedVal = "On every new slide (Enter) — Off"
    End If
End Sub

Sub RibbonGetBgEveryNewSlide(control As IRibbonControl, ByRef returnedVal)
    returnedVal = GetBackgroundEveryNewSlide()
End Sub

Sub RibbonToggleBgEveryNewSlide(control As IRibbonControl, pressed As Boolean)
    InitOnionEventBridge
    SaveBackgroundEveryNewSlide pressed
End Sub

Sub RibbonGetBackgroundToolsEnabled(control As IRibbonControl, ByRef returnedVal)
    returnedVal = IsAnimationModeEnabled()
End Sub

Sub RibbonRemoveSelectedBackground(control As IRibbonControl)
    InitOnionEventBridge
    RemoveSelectedBackgroundShapes
End Sub

Sub RibbonClearBackgroundCurrentSlide(control As IRibbonControl)
    InitOnionEventBridge
    ClearBackgroundShapesOnCurrentSlide
End Sub

Sub RibbonClearBackgroundAllSlides(control As IRibbonControl)
    InitOnionEventBridge
    ClearBackgroundShapesOnAllSlides
End Sub

Sub RibbonOnLoad(ribbon As IRibbonUI)
    Set gRibbon = ribbon
    InitOnionEventBridge
End Sub

Sub RibbonToggleAnimationMode(control As IRibbonControl, pressed As Boolean)
    SaveAnimationModeEnabled pressed
End Sub

Sub RibbonGetAnimationMode(control As IRibbonControl, ByRef returnedVal)
    returnedVal = IsAnimationModeEnabled()
End Sub

Public Sub InitOnionEventBridge()
    If gEventBridge Is Nothing Then
        Set gEventBridge = New OnionSkinEventBridge
        Set gEventBridge.App = Application
    End If
End Sub

Public Function IsAnimationModeEnabled() As Boolean
    IsAnimationModeEnabled = (GetAnimationMode() <> ANIMATION_MODE_OFF)
End Function

Public Sub SaveAnimationModeEnabled(ByVal enabled As Boolean)
    If enabled Then
        SaveAnimationMode ANIMATION_MODE_CLASSIC
    Else
        SaveAnimationMode ANIMATION_MODE_OFF
    End If
End Sub

Public Function GetAnimationMode() As String
    Dim mode As String
    mode = GetSetting(SETTINGS_APP_NAME, SETTINGS_SECTION, SETTINGS_ANIMATION_MODE, ANIMATION_MODE_OFF)

    ' Backward compatibility with the old checkbox setting.
    If mode = "1" Then mode = ANIMATION_MODE_CLASSIC
    If mode = "0" Or mode = "" Then mode = ANIMATION_MODE_OFF

    Select Case mode
        Case ANIMATION_MODE_CLASSIC, ANIMATION_MODE_DUPLICATE
            GetAnimationMode = mode
        Case Else
            GetAnimationMode = ANIMATION_MODE_OFF
    End Select
End Function

Public Sub SaveAnimationMode(ByVal mode As String)
    Select Case mode
        Case ANIMATION_MODE_CLASSIC, ANIMATION_MODE_DUPLICATE
            SaveSetting SETTINGS_APP_NAME, SETTINGS_SECTION, SETTINGS_ANIMATION_MODE, mode
        Case Else
            SaveSetting SETTINGS_APP_NAME, SETTINGS_SECTION, SETTINGS_ANIMATION_MODE, ANIMATION_MODE_OFF
    End Select

    If Not gRibbon Is Nothing Then
        gRibbon.Invalidate
    End If
End Sub

Sub Auto_Open()
    InitOnionEventBridge
End Sub
