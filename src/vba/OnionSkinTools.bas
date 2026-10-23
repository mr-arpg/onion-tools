Option Explicit

' =================================================================
'  ONION SKIN TOOLS - PowerPoint Animation Helper
'  Import this file in the VBA Editor (File > Import File)
'  Then assign the Subs to Quick Access Toolbar buttons
' =================================================================

' Shape name kept for compatibility with decks that already use this guide
Private Const GUIDE_NAME As String = "GuiaOnionSkin"
Private Const GUIDE_SLIDE_TAG As String = "OnionGuideBg"
Private Const GUIDE_OPACITY As Single = 0.75   ' 0 = opaque | 1 = invisible
Private Const SETTINGS_APP_NAME As String = "OnionSkinTools"
Private Const SETTINGS_SECTION As String = "Preferences"
Private Const SETTINGS_ANIMATION_MODE As String = "AnimationMode"
Private Const SETTINGS_OUTPUT_FOLDER As String = "OutputFolder"
Private Const SETTINGS_EXPORT_SCOPE As String = "ExportScope"
Private Const EXPORT_SCOPE_ALL As String = "All"
Private Const EXPORT_SCOPE_SELECTED As String = "Selected"
Private Const ANIMATION_MODE_OFF As String = "Off"
Private Const ANIMATION_MODE_CLASSIC As String = "Classic"
Private Const ANIMATION_MODE_DUPLICATE As String = "Duplicate"

' Hardcoded so the module compiles on PowerPoint builds where these
' enum members were added later than the rest of PpSaveAsFileType /
' PpMediaTaskStatus (e.g. ppSaveAsAnimatedGIF = 40 only exists in M365).
Private Const PPT_SAVE_AS_ANIMATED_GIF As Long = 40
Private Const PPT_SAVE_AS_OPEN_XML_PRESENTATION As Long = 24
Private Const PPT_SAVE_AS_OPEN_XML_PRESENTATION_MACRO_ENABLED As Long = 25
Private Const PPT_MEDIA_TASK_DONE As Long = 3
Private Const PPT_MEDIA_TASK_FAILED As Long = 4
' Default slide duration for GIF / video export when the slide has no
' user-defined transition timer. (PowerPoint UI: 00:00,01 = 1/100 s)
Private Const DEFAULT_FRAME_ADVANCE_SECONDS As Single = 0.01

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

Public Sub CreateGuideForSlideIndex(ByVal idx As Integer, Optional ByVal silent As Boolean = False, Optional ByVal prs As Presentation = Nothing)
    Dim sldCurrent       As Slide
    Dim sldPrevious      As Slide
    Dim hiddenOnPrev     As Collection
    Dim hadPrevBgGuide   As Boolean
    Dim imgTemp          As String

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

    If prs Is Nothing Then Set prs = ActivePresentation
    Set sldCurrent = prs.Slides(idx)
    Set sldPrevious = prs.Slides(idx - 1)

    If SlideHasGuide(sldCurrent) Then
        If Not silent Then
            Dim resp As Integer
            resp = MsgBox("A guide already exists on this slide." & vbCrLf & "Replace it?", _
                          vbYesNo + vbQuestion, "Onion Skin Tools")
            If resp = vbNo Then Exit Sub
        End If
        ClearGuideOnSlide sldCurrent
    End If

    ' unique file name to avoid conflicts when multiple decks are open
    imgTemp = Environ("TEMP") & "\ost_" & idx & "_" & format(Now, "hhmmss") & ".png"

    ' --- Hide the previous guide before export ---
    ' Export captures the slide background and visible shapes; without this,
    ' old guides would stack from slide to slide.
    Set hiddenOnPrev = HideInternalExportShapes(sldPrevious)
    hadPrevBgGuide = SuspendSlideGuideBackground(sldPrevious)

    sldPrevious.Export imgTemp, "PNG"

    ' --- Show the previous guide again right after export ---
    RestoreSlideGuideBackground sldPrevious, hadPrevBgGuide
    RestoreHiddenShapes hiddenOnPrev

    ' --- Put the ghost in the slide background so it cannot be selected or deleted ---
    ApplyGuideToSlideBackground sldCurrent, imgTemp

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

    If Not ClearGuideOnSlide(sld) Then
        MsgBox "No guide found on the current slide.", _
               vbInformation, "Onion Skin Tools"
    End If
End Sub


' -----------------------------------------------------------------
'  2b. CLEAR GUIDES - removes all guides from the presentation
' -----------------------------------------------------------------
Sub ClearGuides()
    Dim sld     As Slide
    Dim count   As Integer
    count = 0

    For Each sld In ActivePresentation.Slides
        If ClearGuideOnSlide(sld) Then
            count = count + 1
        End If
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
'  3. NEW SLIDE - guide and duplicate mode
' -----------------------------------------------------------------
Private Function SlideHasBackgroundGuide(ByVal sld As Slide) As Boolean
    On Error Resume Next
    SlideHasBackgroundGuide = (sld.Tags(GUIDE_SLIDE_TAG) = "1")
    On Error GoTo 0
End Function

Private Function SlideHasShapeGuide(ByVal sld As Slide) As Boolean
    Dim shp As Shape
    For Each shp In sld.Shapes
        If shp.Name = GUIDE_NAME Then
            SlideHasShapeGuide = True
            Exit Function
        End If
    Next shp
End Function

Private Function SlideHasGuide(ByVal sld As Slide) As Boolean
    SlideHasGuide = SlideHasBackgroundGuide(sld) Or SlideHasShapeGuide(sld)
End Function

Private Sub ApplyGuideToSlideBackground(ByVal sld As Slide, ByVal imagePath As String)
    On Error Resume Next
    sld.FollowMasterBackground = msoFalse
    With sld.Background.Fill
        .Visible = msoTrue
        .UserPicture imagePath
        .Transparency = GUIDE_OPACITY
    End With
    sld.Tags.Delete GUIDE_SLIDE_TAG
    sld.Tags.Add GUIDE_SLIDE_TAG, "1"
    On Error GoTo 0
End Sub

Private Function ClearGuideOnSlide(ByVal sld As Slide) As Boolean
    Dim i As Long
    Dim removed As Boolean

    If SlideHasBackgroundGuide(sld) Then
        On Error Resume Next
        sld.FollowMasterBackground = msoTrue
        sld.Tags.Delete GUIDE_SLIDE_TAG
        On Error GoTo 0
        removed = True
    End If

    For i = sld.Shapes.count To 1 Step -1
        If sld.Shapes(i).Name = GUIDE_NAME Then
            On Error Resume Next
            sld.Shapes(i).Locked = msoFalse
            sld.Shapes(i).Delete
            On Error GoTo 0
            removed = True
        End If
    Next i

    ClearGuideOnSlide = removed
End Function

Private Function SuspendSlideGuideBackground(ByVal sld As Slide) As Boolean
    If Not SlideHasBackgroundGuide(sld) Then Exit Function
    On Error Resume Next
    sld.FollowMasterBackground = msoTrue
    On Error GoTo 0
    SuspendSlideGuideBackground = True
End Function

Private Sub RestoreSlideGuideBackground(ByVal sld As Slide, ByVal wasSuspended As Boolean)
    If Not wasSuspended Then Exit Sub
    On Error Resume Next
    sld.FollowMasterBackground = msoFalse
    On Error GoTo 0
End Sub

Private Function HideInternalExportShapes(ByVal sld As Slide) As Collection
    Dim hiddenShapes As Collection
    Set hiddenShapes = New Collection

    Dim shp As Shape
    For Each shp In sld.Shapes
        If shp.Name = GUIDE_NAME Then
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

Private Function IsEmptyPlaceholder(ByVal shp As Shape) As Boolean
    On Error Resume Next
    If shp.Type = msoPlaceholder Then
        If shp.HasTextFrame = msoTrue Then
            If shp.TextFrame.HasText = msoFalse Then
                IsEmptyPlaceholder = True
            End If
        Else
            IsEmptyPlaceholder = True
        End If
    End If
    On Error GoTo 0
End Function

Private Sub CopySlideShapesForDuplicateMode(ByVal sourceSlide As Slide, ByVal targetSlide As Slide)
    Dim shp As Shape
    Dim pastedRange As ShapeRange

    For Each shp In sourceSlide.Shapes
        If shp.Name <> GUIDE_NAME And Not IsEmptyPlaceholder(shp) Then
            On Error Resume Next
            Err.Clear
            shp.Copy
            If Err.Number = 0 Then
                Set pastedRange = targetSlide.Shapes.Paste
            End If
            On Error GoTo 0
        End If
    Next shp
End Sub

Private Sub ClearShapesOnSlide(ByVal sld As Slide)
    Dim i As Long
    For i = sld.Shapes.count To 1 Step -1
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

    Dim prs As Presentation
    Set prs = sldNew.Parent

    Dim sldPrevious As Slide
    Set sldPrevious = prs.Slides(sldNew.SlideIndex - 1)

    If mode = ANIMATION_MODE_DUPLICATE Then
        ClearShapesOnSlide sldNew
    End If

    CreateGuideForSlideIndex sldNew.SlideIndex, True, prs

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
            If ActivePresentation.path <> "" Then path = ActivePresentation.path
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
'  3b. EXPORT
'      Guides are always left out of the file. GIF and video render a
'      temp copy when only some slides are included. Images are written
'      straight from the open deck. Output goes to the configured folder.
' -----------------------------------------------------------------
Sub ExportGIF()
    ExportViaCopy "gif"
End Sub

Sub ExportVideo()
    ExportViaCopy "video"
End Sub

Sub ExportImages(ByVal imageFormat As String)
    Dim prs As Presentation
    Dim ids() As Long
    Dim i As Long
    Dim sld As Slide
    Dim hidden As Collection
    Dim outDir As String
    Dim baseName As String
    Dim ext As String
    Dim filePath As String
    Dim firstPath As String
    Dim pxW As Long
    Dim pxH As Long
    Dim saveErr As Long

    If ActivePresentation Is Nothing Then Exit Sub
    Set prs = ActivePresentation
    If Not TryGetExportSlideIDs(prs, ids) Then Exit Sub

    outDir = GetOutputFolder()
    baseName = ExportBaseName(prs)
    ext = LCase$(imageFormat)
    pxW = 1920
    If prs.PageSetup.SlideWidth > 0 Then
        pxH = CLng(pxW * prs.PageSetup.SlideHeight / prs.PageSetup.SlideWidth)
    Else
        pxH = 1080
    End If

    Dim hadBgGuide As Boolean
    For i = LBound(ids) To UBound(ids)
        Set sld = FindSlideByID(prs, ids(i))
        If Not sld Is Nothing Then
            Set hidden = HideInternalExportShapes(sld)
            hadBgGuide = SuspendSlideGuideBackground(sld)
            filePath = outDir & "\" & baseName & "_" & Format$(sld.SlideIndex, "00") & "." & ext
            If firstPath = "" Then firstPath = filePath

            On Error Resume Next
            Err.Clear
            If imageFormat = "PNG" Then
                sld.Export filePath, "PNG", pxW, pxH
            Else
                sld.Export filePath, "SVG"
            End If
            saveErr = Err.Number
            On Error GoTo 0
            RestoreSlideGuideBackground sld, hadBgGuide
            RestoreHiddenShapes hidden

            If saveErr <> 0 Then
                MsgBox "Could not export " & imageFormat & "." & vbCrLf & _
                       "This PowerPoint build may not support that format.", _
                       vbExclamation, "Onion Skin Tools"
                Exit Sub
            End If
        End If
    Next i

    If firstPath <> "" Then
        Shell "explorer.exe /select,""" & firstPath & """", vbNormalFocus
    End If
End Sub

Public Function ExportSelectedOnly() As Boolean
    ExportSelectedOnly = (GetSetting(SETTINGS_APP_NAME, SETTINGS_SECTION, SETTINGS_EXPORT_SCOPE, EXPORT_SCOPE_ALL) = EXPORT_SCOPE_SELECTED)
End Function

Public Sub SaveExportScope(ByVal scope As String)
    If scope <> EXPORT_SCOPE_SELECTED Then scope = EXPORT_SCOPE_ALL
    SaveSetting SETTINGS_APP_NAME, SETTINGS_SECTION, SETTINGS_EXPORT_SCOPE, scope
    If Not gRibbon Is Nothing Then
        gRibbon.InvalidateControl "btnExportAllSlides"
        gRibbon.InvalidateControl "btnExportSelectedSlides"
        gRibbon.InvalidateControl "btnExportAllSlidesDraw"
        gRibbon.InvalidateControl "btnExportSelectedSlidesDraw"
    End If
End Sub

Private Function ExportBaseName(ByVal prs As Presentation) As String
    If prs.path <> "" Then
        ExportBaseName = Left(prs.Name, InStrRev(prs.Name, ".") - 1)
    Else
        ExportBaseName = "OnionExport"
    End If
End Function

Private Function FindSlideByID(ByVal prs As Presentation, ByVal slideId As Long) As Slide
    Dim sld As Slide
    For Each sld In prs.Slides
        If sld.SlideID = slideId Then
            Set FindSlideByID = sld
            Exit Function
        End If
    Next sld
End Function

Private Function TryGetExportSlideIDs(ByVal prs As Presentation, ByRef ids() As Long) As Boolean
    Dim i As Long
    Dim n As Long
    Dim sld As Slide

    If Not ExportSelectedOnly() Then
        ReDim ids(1 To prs.Slides.Count)
        For i = 1 To prs.Slides.Count
            ids(i) = prs.Slides(i).SlideID
        Next i
        TryGetExportSlideIDs = True
        Exit Function
    End If

    On Error Resume Next
    If Not ActiveWindow Is Nothing Then
        If ActiveWindow.Selection.Type = ppSelectionSlides Then
            n = ActiveWindow.Selection.SlideRange.Count
        End If
    End If
    On Error GoTo 0

    If n > 0 Then
        ReDim ids(1 To n)
        For i = 1 To n
            ids(i) = ActiveWindow.Selection.SlideRange(i).SlideID
        Next i
        TryGetExportSlideIDs = True
        Exit Function
    End If

    On Error Resume Next
    Set sld = ActiveWindow.View.Slide
    On Error GoTo 0
    If sld Is Nothing Then
        MsgBox "Select one or more slides in the thumbnail pane, or stay on the slide you want to export.", _
               vbInformation, "Onion Skin Tools"
        Exit Function
    End If

    ReDim ids(1 To 1)
    ids(1) = sld.SlideID
    TryGetExportSlideIDs = True
End Function

Private Sub KeepOnlySlidesByID(ByVal prs As Presentation, ByRef keepIDs() As Long)
    Dim i As Long
    Dim j As Long
    Dim keep As Boolean

    For i = prs.Slides.Count To 1 Step -1
        keep = False
        For j = LBound(keepIDs) To UBound(keepIDs)
            If prs.Slides(i).SlideID = keepIDs(j) Then
                keep = True
                Exit For
            End If
        Next j
        If Not keep Then prs.Slides(i).Delete
    Next i
End Sub

Private Sub RemoveGuidesFromPresentation(ByVal prs As Presentation)
    Dim sld As Slide

    For Each sld In prs.Slides
        ClearGuideOnSlide sld
    Next sld
End Sub

Private Sub ApplyExportSlideTimings(ByVal prs As Presentation)
    Dim sld As Slide

    For Each sld In prs.Slides
        With sld.SlideShowTransition
            ' Apply the default frame duration when the slide has no explicit user-set timer.
            ' If no timed advance is set, use the default frame duration.
            If .AdvanceOnTime <> msoTrue Or .AdvanceTime <= 0 Then
                .AdvanceOnClick = msoFalse
                .AdvanceOnTime = msoTrue
                .AdvanceTime = DEFAULT_FRAME_ADVANCE_SECONDS
            End If
        End With
    Next sld

    On Error Resume Next
    prs.SlideShowSettings.AdvanceMode = ppSlideShowUseSlideTimings
    On Error GoTo 0
End Sub

Private Sub ExportViaCopy(ByVal outputFormat As String)
    Dim originalPres As Presentation
    Dim ids() As Long
    Set originalPres = ActivePresentation
    If originalPres Is Nothing Then Exit Sub
    If Not TryGetExportSlideIDs(originalPres, ids) Then Exit Sub

    Dim outputDir As String, baseName As String, suffix As String, ext As String
    outputDir = GetOutputFolder()
    If outputFormat = "gif" Then ext = ".gif" Else ext = ".mp4"
    suffix = ""
    If ExportSelectedOnly() Then suffix = "_selected"
    baseName = ExportBaseName(originalPres)

    Dim outputPath As String
    outputPath = outputDir & "\" & baseName & suffix & ext

    If outputFormat = "gif" And Not ExportSelectedOnly() Then
        ExportGIFDirect originalPres, outputPath
    ElseIf outputFormat = "gif" Then
        ExportGIFFromCopy originalPres, outputPath, ids
    Else
        ExportVideoViaCopy originalPres, outputPath, ids
    End If
End Sub

Private Function SaveTempCopy(ByVal prs As Presentation) As String
    ' Saves a macro-free .pptx copy to TEMP and returns the path.
    ' Falls back to .pptm if the deck cannot be saved without macros.
    Dim path As String
    path = Environ("TEMP") & "\OnionExport_" & format(Now, "yyyymmddhhmmss") & ".pptx"

    Dim saveErr As Long
    On Error Resume Next
    prs.SaveCopyAs path, PPT_SAVE_AS_OPEN_XML_PRESENTATION
    saveErr = Err.Number
    Err.Clear
    On Error GoTo 0

    If saveErr <> 0 Then
        path = Environ("TEMP") & "\OnionExport_" & format(Now, "yyyymmddhhmmss") & ".pptm"
        prs.SaveCopyAs path, PPT_SAVE_AS_OPEN_XML_PRESENTATION_MACRO_ENABLED
    End If

    SaveTempCopy = path
End Function

Private Sub ExportGIFDirect(ByVal originalPres As Presentation, ByVal outputPath As String)
    ' Fast/direct GIF path: hide guide shapes in the open presentation, export,
    ' then restore only the guides that were visible before export.
    Dim hiddenGuides As Collection
    Dim bgGuideSlides As Collection
    Dim item As Variant
    Set hiddenGuides = New Collection
    Set bgGuideSlides = New Collection

    On Error GoTo GIFCleanFail

    Dim sld As Slide
    Dim shp As Shape
    For Each sld In originalPres.Slides
        If SuspendSlideGuideBackground(sld) Then
            bgGuideSlides.Add sld
        End If
        For Each shp In sld.Shapes
            If shp.Name = GUIDE_NAME And shp.Visible <> msoFalse Then
                shp.Visible = msoFalse
                hiddenGuides.Add shp
            End If
        Next shp
    Next sld
    DoEvents

    ' --- use 00:00,01 for slides without a defined transition timer ---
    For Each sld In originalPres.Slides
        With sld.SlideShowTransition
            If .AdvanceOnTime <> msoTrue Or .AdvanceTime <= 0 Then
                .AdvanceOnClick = msoFalse
                .AdvanceOnTime = msoTrue
                .AdvanceTime = DEFAULT_FRAME_ADVANCE_SECONDS
            End If
        End With
    Next sld

    On Error Resume Next
    originalPres.SlideShowSettings.AdvanceMode = ppSlideShowUseSlideTimings
    On Error GoTo GIFCleanFail

    originalPres.SaveAs outputPath, PPT_SAVE_AS_ANIMATED_GIF

    ' --- restore guide visibility immediately after export ---
    On Error Resume Next
    For Each item In bgGuideSlides
        RestoreSlideGuideBackground item, True
    Next item
    For Each item In hiddenGuides
        item.Visible = msoTrue
    Next item
    On Error Resume Next: originalPres.Windows(1).Activate: On Error GoTo 0
    Shell "explorer.exe /select,""" & outputPath & """", vbNormalFocus
    Exit Sub

GIFCleanFail:
    Dim errDesc As String
    errDesc = Err.Description

    ' Restore any guides that were hidden before the failure
    On Error Resume Next
    For Each item In bgGuideSlides
        RestoreSlideGuideBackground item, True
    Next item
    For Each item In hiddenGuides
        item.Visible = msoTrue
    Next item
    On Error Resume Next: originalPres.Windows(1).Activate: On Error GoTo 0
    On Error GoTo 0

    MsgBox "Export failed: " & errDesc, vbExclamation, "Onion Skin Tools"
End Sub

Private Sub ExportGIFFromCopy(ByVal originalPres As Presentation, _
                              ByVal outputPath As String, _
                              ByRef keepIDs() As Long)
    Dim tempPath As String
    Dim copyPres As Presentation

    On Error GoTo GIFCopyCleanFail

    tempPath = SaveTempCopy(originalPres)
    Set copyPres = Presentations.Open(tempPath, _
        ReadOnly:=msoFalse, _
        Untitled:=msoFalse, _
        WithWindow:=msoFalse)

    KeepOnlySlidesByID copyPres, keepIDs
    RemoveGuidesFromPresentation copyPres
    ApplyExportSlideTimings copyPres
    copyPres.SaveAs outputPath, PPT_SAVE_AS_ANIMATED_GIF

    copyPres.Saved = True
    copyPres.Close
    Set copyPres = Nothing
    On Error Resume Next: Kill tempPath: On Error GoTo 0
    On Error Resume Next: originalPres.Windows(1).Activate: On Error GoTo 0
    Shell "explorer.exe /select,""" & outputPath & """", vbNormalFocus
    Exit Sub

GIFCopyCleanFail:
    Dim errDesc As String
    errDesc = Err.Description
    On Error Resume Next
    If Not copyPres Is Nothing Then
        copyPres.Saved = True
        copyPres.Close
    End If
    Kill tempPath
    On Error Resume Next: originalPres.Windows(1).Activate: On Error GoTo 0
    On Error GoTo 0
    MsgBox "Export failed: " & errDesc, vbExclamation, "Onion Skin Tools"
End Sub

Private Sub ExportVideoViaCopy(ByVal originalPres As Presentation, _
                                ByVal outputPath As String, _
                                ByRef keepIDs() As Long)
    Dim tempPath As String
    Dim copyPres As Presentation

    On Error GoTo VideoCleanFail

    tempPath = SaveTempCopy(originalPres)

    Set copyPres = Presentations.Open(tempPath, _
        ReadOnly:=msoFalse, _
        Untitled:=msoFalse, _
        WithWindow:=msoFalse)

    If ExportSelectedOnly() Then KeepOnlySlidesByID copyPres, keepIDs
    RemoveGuidesFromPresentation copyPres
    ApplyExportSlideTimings copyPres

    ' CreateVideo operates on the live in-memory presentation object.
    copyPres.CreateVideo outputPath, True, DEFAULT_FRAME_ADVANCE_SECONDS, 1080, 30, 100

    Dim videoStatus As Long
    Dim probeErr As Long
    Do
        DoEvents
        OST_Sleep 250
        On Error Resume Next
        videoStatus = copyPres.CreateVideoStatus
        probeErr = Err.Number
        On Error GoTo VideoCleanFail
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
    Exit Sub

VideoCleanFail:
    Dim errDesc As String
    errDesc = Err.Description

    On Error Resume Next
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

Sub RibbonExportGIF(control As IRibbonControl)
    InitOnionEventBridge
    ExportGIF
End Sub

Sub RibbonExportVideo(control As IRibbonControl)
    InitOnionEventBridge
    ExportVideo
End Sub

Sub RibbonExportPNG(control As IRibbonControl)
    InitOnionEventBridge
    ExportImages "PNG"
End Sub

Sub RibbonExportSVG(control As IRibbonControl)
    InitOnionEventBridge
    ExportImages "SVG"
End Sub

Sub RibbonSetExportAllSlides(control As IRibbonControl, pressed As Boolean)
    If Not pressed Then Exit Sub
    SaveExportScope EXPORT_SCOPE_ALL
End Sub

Sub RibbonSetExportSelectedSlides(control As IRibbonControl, pressed As Boolean)
    If Not pressed Then Exit Sub
    SaveExportScope EXPORT_SCOPE_SELECTED
End Sub

Sub RibbonGetExportAllSlides(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Not ExportSelectedOnly()
End Sub

Sub RibbonGetExportSelectedSlides(control As IRibbonControl, ByRef returnedVal)
    returnedVal = ExportSelectedOnly()
End Sub

Sub RibbonSetOutputFolder(control As IRibbonControl)
    InitOnionEventBridge
    SetOutputFolder
End Sub

Sub RibbonSetAnimationModeOff(control As IRibbonControl, pressed As Boolean)
    If Not pressed Then Exit Sub
    SaveAnimationMode ANIMATION_MODE_OFF
End Sub

Sub RibbonSetAnimationModeClassic(control As IRibbonControl, pressed As Boolean)
    If Not pressed Then Exit Sub
    SaveAnimationMode ANIMATION_MODE_CLASSIC
End Sub

Sub RibbonSetAnimationModeDuplicate(control As IRibbonControl, pressed As Boolean)
    If Not pressed Then Exit Sub
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
        gRibbon.InvalidateControl "btnAnimationModeOff"
        gRibbon.InvalidateControl "btnAnimationModeClassic"
        gRibbon.InvalidateControl "btnAnimationModeDuplicate"
        gRibbon.InvalidateControl "btnAnimationModeOffDraw"
        gRibbon.InvalidateControl "btnAnimationModeClassicDraw"
        gRibbon.InvalidateControl "btnAnimationModeDuplicateDraw"
        gRibbon.Invalidate
    End If
End Sub

Sub Auto_Open()
    InitOnionEventBridge
End Sub
