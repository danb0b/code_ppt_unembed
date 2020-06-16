Attribute VB_Name = "Module1"
Private Function slide_index() As Integer

    Dim ap As Presentation
    Set ap = ActivePresentation
    Dim ii As Integer
    'Set sl = ap.Slides(2)
    'ii = sl.SlideIndex
    
    'Dim var1 As Variant
    'Set var1 = SlideShowWindows
    ii = ap.Windows(1).View.slide.SlideIndex
    slide_index = ii
    
    'MsgBox SlideShowWindows(1).View.Slide.SlideIndex
End Function

Sub EmbeddedMoviesToLinkedMovies()


Dim oSl As slide
Dim oSh As Shape
Dim x As Long
Dim sPath As String

' Edit this to the path where vids are stored,
' path must end in backslash

sPath = "null"
'sPath = "C:\Users\danaukes\Desktop\"
Dim ii As Integer

'ii = 0
'ii = slide_index

Dim dummy As Boolean


If sPath = "null" Then
    MsgBox "select folder"
    sPath = folder_path
End If

'For Each oSl In ActivePresentation.Slides
For ii = slide_index To (ActivePresentation.Slides.Count) Step 1
    Set oSl = ActivePresentation.Slides(ii)
    oSl.Select
    
    For Each oSh In oSl.Shapes
        dummy = ConvertToLinked_recursive(oSh, sPath, oSl)
        
    Next oSh   ' Shape
Next    ' Slide

If ii <> 0 Then
DeleteEmptyPlaceholders
Else
MsgBox "No Embedded movies"
End If

If check_embedded_videos = True Then
    MsgBox "File Has Embedded Videos"
End If

End Sub

Sub movedown(old_shape As Shape, new_shape As Shape)
    Dim z As Integer
    z = old_shape.ZOrderPosition
    Do Until new_shape.ZOrderPosition = z
        new_shape.ZOrder (msoSendBackward)
    Loop
End Sub
    
Sub ConvertToLinked(oSh As Shape, sPath As String, slide As slide)

Dim oSl As slide
Dim oNewVid As Shape
Dim x As Long
'Dim lZOrder As Long
Dim ii As Integer
ii = 0


Set oSl = oSh.Parent
'lZOrder = oSh.ZOrderPosition

Dim Name As String

Name = sPath & oSh.Name



On Error GoTo ErrHandler:
    Set oNewVid = oSl.Shapes.AddMediaObject2(Name, msoTrue, msoFalse, oSh.Left, oSh.Top, oSh.Width, oSh.Height)
On Error GoTo 0

'movedown oSh, oNewVid
copyanimationsettings oSh, oNewVid, slide

oSh.Delete

Exit Sub

ErrHandler:
    'MsgBox Err.Number
Select Case ii
Case 0
    Name = sPath & "\" & oSh.Name & ".mp4"
    ii = ii + 1
Case 1
    Name = sPath & "\" & oSh.Name & ".avi"
    ii = ii + 1
Case 2
    Name = sPath & "\" & oSh.Name & ".mov"
    ii = ii + 1
Case 3
    Name = sPath & "\" & oSh.Name & ".wmv"
    ii = ii + 1
Case 4
    result = InputBox("Rename the file", "Rename the file", oSh.Name)
    Name = sPath & "\" & result
    ii = ii + 1
Case 5
    Name = file_path
    ii = ii + 1
Case Else
    'MsgBox "Could not find" & oSh.name
    result = InputBox("Could not find" & oSh.Name & ". Continue?")
    If result = "y" Then
    ii = 5
    Else
    Err.Raise vbObjectError + 513, "Module1::Test()", "My custom error."
    End If
    
End Select
   
Resume

End Sub


Function ConvertToLinked_recursive(a As Variant, ByRef sPath As String, slide As slide) As Boolean
        Dim objshape As Shape
        Set objshape = a
        objshape.Select
        
        Dim dummy As Boolean
                
        If objshape.HasTextFrame = True Then
        ElseIf objshape.HasTable = True Then
        ElseIf (objshape.Type = msoPicture) Then
        ElseIf objshape.Type = msoEmbeddedOLEObject Then
        ElseIf (objshape.Type = msoMedia) Then
            If objshape.MediaType = ppMediaTypeMovie Then
                If objshape.MediaFormat.IsEmbedded Then
                    ConvertToLinked objshape, sPath, slide
                End If
            End If
        ElseIf (objshape.Type = msoGroup) Then
            For Each objGrpItem In objshape.GroupItems
                objGrpItem.Select
                
                dummy = ConvertToLinked_recursive(objGrpItem, sPath, slide)
            Next objGrpItem
        End If
        ConvertToLinked_recursive = result
End Function



 Private Sub DeleteEmptyPlaceholders()
    Dim oSlide As slide
    Dim oShape As Shape
     
    For Each oSlide In ActivePresentation.Slides
        For Each oShape In oSlide.Shapes
            With oShape
                If .HasTextFrame Then
                
                If .Type = msoPlaceholder Then
                    If .TextFrame.TextRange.Length = 0 Then
                        .Delete
                    End If
                End If
                End If
                
            End With
        Next
    Next
End Sub


Private Sub SelectAllShapes()
Dim oSh As Shape, shapeString As String, thisSlideNbr As Long

'   GET THE NUMBER OF THE ACTIVE SLIDE
thisSlideNbr = ActiveWindow.View.slide.SlideIndex

shapeString = InputBox("Type the Shape Name you want to select")
With ActivePresentation.Slides(thisSlideNbr)
    For Each oSh In .Shapes
        If InStr(1, oSh.Name, shapeString) > 0 Then
            oSh.Select
        End If
    Next
End With

End Sub

Private Function file_path()
    Dim lngCount As Long

    ' Open the file dialog
    With Application.FileDialog(msoFileDialogOpen)
        .AllowMultiSelect = False
        .Show

        ' Display paths of each file selected
        For lngCount = 1 To .SelectedItems.Count
            file_path = .SelectedItems(lngCount)
        Next lngCount

    End With
    'name = sPath & name
    
End Function

Private Function folder_path()
    Dim lngCount As Long

    ' Open the file dialog
    With Application.FileDialog(msoFileDialogFolderPicker)
    
        .AllowMultiSelect = False
        .Show

        ' Display paths of each file selected
        For lngCount = 1 To .SelectedItems.Count
            folder_path = .SelectedItems(lngCount)
        Next lngCount

    End With
    'name = sPath & name
    
End Function
Sub SwitchPath()

Dim oSl As slide
Dim oSh As Shape
Dim x As Long
Dim sPath As String

sPath = folder_path
'sPath = "C:\Users\danaukes\Dropbox (ASU)\classes\2017-2018_S_EGR_598 Foldable Robotics\lectures\videos"

Dim ii As Integer
ii = 0
Dim l As Integer


For Each oSl In ActivePresentation.Slides
    oSl.Select
    For x = oSl.Shapes.Count To 1 Step -1
        Set oSh = oSl.Shapes(x)
        If oSh.Type = msoMedia Then
            If oSh.MediaType = ppMediaTypeMovie Then
                If Not oSh.MediaFormat.IsEmbedded Then
                    ii = ii + 1
                    ActiveWindow.View.GotoSlide oSl.SlideIndex
                    oSh.Select
                    List = Split(oSh.LinkFormat.SourceFullName, "\")
                    l = UBound(List) - LBound(List) + 1
                    'Debug.Print List
                    oSh.LinkFormat.SourceFullName = sPath & "\" & List(l - 1)
                End If
            End If
        End If
    Next
Next

If ii <> 0 Then
DeleteEmptyPlaceholders
Else
MsgBox "No Embedded movies"
End If
End Sub


Sub ConvertShapeToPNG()
    Dim oSh As Shape
    Set oSh = ActiveWindow.Selection.ShapeRange(1)
    oSh.Copy
    ActiveWindow.Selection.SlideRange.Shapes.PasteSpecial ppPastePNG
    oSh.Delete
End Sub


Sub PNGIfy()

Dim oSl As slide
Dim oSh As Shape
Dim x As Long
Dim sPath As String
Dim oNewVid As Shape
Dim lZOrder As Long

For Each oSl In ActivePresentation.Slides
    oSl.Select
    
    For x = oSl.Shapes.Count To 1 Step -1
        
        Set oSh = oSl.Shapes(x)
        oSh.Select
        
        
        Debug.Print oSh.Type
        
        If oSh.Type = msoPlaceholder Then
             If oSh.PlaceholderFormat.ContainedType = msoPicture Then
            oSh.Copy
            ActiveWindow.Selection.SlideRange.Shapes.PasteSpecial (ppPastePNG)
            Set oNewVid = oSl.Shapes(oSl.Shapes.Count)
            oNewVid.Left = oSh.Left
            oNewVid.Top = oSh.Top
            
            movedown oSh, oNewVid
            
            
            
            oSh.Delete
            
             End If
        End If
    Next
Next
DeleteEmptyPlaceholders
End Sub


Private Sub copyanimationsettings(shape1 As Shape, shape2 As Shape, slide As slide)

    'Dim test As Boolean
    
    'ActiveWindow.Selection.Unselect
    'shape1.Select 'select the video
    'test = Application.CommandBars.GetPressedMso("MoviePlayFullScreen")
    'ActiveWindow.Selection.Unselect
    'shape2.Select
    'If test <> Application.CommandBars.GetPressedMso("MoviePlayFullScreen") Then
    '    Application.CommandBars.ExecuteMso ("MoviePlayFullScreen")
    'End If
    'ActiveWindow.Selection.Unselect
    
    shape2.Rotation = shape1.Rotation

    'shape2.PictureFormat.CropBottom = shape1.PictureFormat.CropBottom
    'shape2.PictureFormat.CropLeft = shape1.PictureFormat.CropLeft
    'shape2.PictureFormat.CropRight = shape1.PictureFormat.CropRight
    'shape2.PictureFormat.CropTop = shape1.PictureFormat.CropTop
    
    shape2.PictureFormat.Crop.ShapeHeight = shape1.PictureFormat.Crop.ShapeHeight
    shape2.PictureFormat.Crop.ShapeLeft = shape1.PictureFormat.Crop.ShapeLeft
    shape2.PictureFormat.Crop.ShapeTop = shape1.PictureFormat.Crop.ShapeTop
    shape2.PictureFormat.Crop.ShapeWidth = shape1.PictureFormat.Crop.ShapeWidth
    
    shape2.PictureFormat.Crop.PictureHeight = shape1.PictureFormat.Crop.PictureHeight
    shape2.PictureFormat.Crop.PictureOffsetX = shape1.PictureFormat.Crop.PictureOffsetX
    shape2.PictureFormat.Crop.PictureOffsetY = shape1.PictureFormat.Crop.PictureOffsetY
    shape2.PictureFormat.Crop.PictureWidth = shape1.PictureFormat.Crop.PictureWidth
    
    shape2.Top = shape1.Top
    shape2.Left = shape1.Left
    shape2.Width = shape1.Width
    shape2.Height = shape1.Height
    
    
    If shape2.MediaFormat.Length > shape1.MediaFormat.EndPoint Then
        shape2.MediaFormat.EndPoint = shape1.MediaFormat.EndPoint
    Else
        shape2.MediaFormat.EndPoint = shape2.MediaFormat.Length
    End If
    shape2.MediaFormat.StartPoint = shape1.MediaFormat.StartPoint
    
    shape2.MediaFormat.Muted = shape1.MediaFormat.Muted
    shape2.MediaFormat.Volume = shape1.MediaFormat.Volume
    
    shape2.MediaFormat.FadeInDuration = shape1.MediaFormat.FadeInDuration
    shape2.MediaFormat.FadeOutDuration = shape1.MediaFormat.FadeOutDuration
    
    
    'shape2.AnimationSettings.AdvanceMode = shape1.AnimationSettings.AdvanceMode
    'shape2.AnimationSettings.AdvanceTime = shape1.AnimationSettings.AdvanceTime
    
    'shape2.AnimationSettings.AfterEffect = shape1.AnimationSettings.AfterEffect
    'shape2.AnimationSettings.Animate = shape1.AnimationSettings.Animate
    'shape2.AnimationSettings.AnimateBackground = shape1.AnimationSettings.AnimateBackground
    'shape2.AnimationSettings.AnimateTextInReverse = shape1.AnimationSettings.AnimateTextInReverse
    'shape2.AnimationSettings.EntryEffect = shape1.AnimationSettings.EntryEffect
    shape2.AnimationSettings.PlaySettings.LoopUntilStopped = shape1.AnimationSettings.PlaySettings.LoopUntilStopped
    shape2.AnimationSettings.PlaySettings.PauseAnimation = shape1.AnimationSettings.PlaySettings.PauseAnimation
    shape2.AnimationSettings.PlaySettings.PlayOnEntry = shape1.AnimationSettings.PlaySettings.PlayOnEntry
    shape2.AnimationSettings.PlaySettings.RewindMovie = shape1.AnimationSettings.PlaySettings.RewindMovie
    shape2.AnimationSettings.PlaySettings.StopAfterSlides = shape1.AnimationSettings.PlaySettings.StopAfterSlides
    shape2.AnimationSettings.PlaySettings.HideWhileNotPlaying = shape1.AnimationSettings.PlaySettings.HideWhileNotPlaying
    
    shape2.AnimationSettings.AnimationOrder = shape1.AnimationSettings.AnimationOrder
    shape2.AnimationSettings.SoundEffect.Type = shape1.AnimationSettings.SoundEffect.Type
    shape2.AnimationSettings.TextLevelEffect = shape1.AnimationSettings.TextLevelEffect
    shape2.AnimationSettings.TextUnitEffect = shape1.AnimationSettings.TextUnitEffect
    
    slide.TimeLine.MainSequence.Item(shape2.AnimationSettings.AnimationOrder).Timing.Accelerate = slide.TimeLine.MainSequence.Item(shape1.AnimationSettings.AnimationOrder).Timing.Accelerate
    slide.TimeLine.MainSequence.Item(shape2.AnimationSettings.AnimationOrder).Timing.AutoReverse = slide.TimeLine.MainSequence.Item(shape1.AnimationSettings.AnimationOrder).Timing.AutoReverse
    slide.TimeLine.MainSequence.Item(shape2.AnimationSettings.AnimationOrder).Timing.BounceEnd = slide.TimeLine.MainSequence.Item(shape1.AnimationSettings.AnimationOrder).Timing.BounceEnd
    slide.TimeLine.MainSequence.Item(shape2.AnimationSettings.AnimationOrder).Timing.BounceEndIntensity = slide.TimeLine.MainSequence.Item(shape1.AnimationSettings.AnimationOrder).Timing.BounceEndIntensity
    slide.TimeLine.MainSequence.Item(shape2.AnimationSettings.AnimationOrder).Timing.Decelerate = slide.TimeLine.MainSequence.Item(shape1.AnimationSettings.AnimationOrder).Timing.Decelerate
    slide.TimeLine.MainSequence.Item(shape2.AnimationSettings.AnimationOrder).Timing.Duration = slide.TimeLine.MainSequence.Item(shape1.AnimationSettings.AnimationOrder).Timing.Duration
    slide.TimeLine.MainSequence.Item(shape2.AnimationSettings.AnimationOrder).Timing.RepeatCount = slide.TimeLine.MainSequence.Item(shape1.AnimationSettings.AnimationOrder).Timing.RepeatCount
    slide.TimeLine.MainSequence.Item(shape2.AnimationSettings.AnimationOrder).Timing.RepeatDuration = slide.TimeLine.MainSequence.Item(shape1.AnimationSettings.AnimationOrder).Timing.RepeatDuration
    slide.TimeLine.MainSequence.Item(shape2.AnimationSettings.AnimationOrder).Timing.Restart = slide.TimeLine.MainSequence.Item(shape1.AnimationSettings.AnimationOrder).Timing.Restart
    slide.TimeLine.MainSequence.Item(shape2.AnimationSettings.AnimationOrder).Timing.SmoothEnd = slide.TimeLine.MainSequence.Item(shape1.AnimationSettings.AnimationOrder).Timing.SmoothEnd
    slide.TimeLine.MainSequence.Item(shape2.AnimationSettings.AnimationOrder).Timing.SmoothStart = slide.TimeLine.MainSequence.Item(shape1.AnimationSettings.AnimationOrder).Timing.SmoothStart
    slide.TimeLine.MainSequence.Item(shape2.AnimationSettings.AnimationOrder).Timing.Speed = slide.TimeLine.MainSequence.Item(shape1.AnimationSettings.AnimationOrder).Timing.Speed
    
    On Error GoTo AnErrHandler1:
        slide.TimeLine.MainSequence.Item(shape2.AnimationSettings.AnimationOrder).Timing.TriggerBookmark = slide.TimeLine.MainSequence.Item(shape1.AnimationSettings.AnimationOrder).Timing.TriggerBookmark
    On Error GoTo 0

    slide.TimeLine.MainSequence.Item(shape2.AnimationSettings.AnimationOrder).Timing.TriggerDelayTime = slide.TimeLine.MainSequence.Item(shape1.AnimationSettings.AnimationOrder).Timing.TriggerDelayTime
    
    On Error GoTo AnErrHandler1:
        slide.TimeLine.MainSequence.Item(shape2.AnimationSettings.AnimationOrder).Timing.TriggerShape = slide.TimeLine.MainSequence.Item(shape1.AnimationSettings.AnimationOrder).Timing.TriggerShape
    On Error GoTo 0
    
    slide.TimeLine.MainSequence.Item(shape2.AnimationSettings.AnimationOrder).Timing.TriggerType = slide.TimeLine.MainSequence.Item(shape1.AnimationSettings.AnimationOrder).Timing.TriggerType
    

Exit Sub

AnErrHandler1:
'MsgBox "asdf"
Resume Next

End Sub

Function file_from_path(path)
    List = Split(path, "/")
    l = UBound(List) - LBound(List) + 1
    file_from_path = List(l - 1)
End Function

Function strip_ext(path)
    Dim strg As String
    List = Split(path, ".")
    'l = UBound(List) - LBound(List)
    'If l = 1 Then
        'strg = List(1)
    'Else
        For ii = LBound(List) To (UBound(List) - 2)
        strg = strg & List(ii) & "."
        Next ii
        strg = strg & List(UBound(List) - LBound(List) - 1)
    'End If
    'MsgBox strg
    strip_ext = strg
End Function

Function strip_chars(path, char, rchar)
    Dim strg As String
    List = Split(path, char)
    For ii = LBound(List) To (UBound(List) - 1)
    strg = strg & List(ii) & rchar
    Next ii
    strg = strg & List(UBound(List) - LBound(List))
    'MsgBox strg
    strip_chars = strg
End Function


