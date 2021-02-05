Attribute VB_Name = "ExportTools"
' this code extracts text from PPT(X) and saves to latex beamer body
' Provided for free with no guarantees or promises
'WARNING: this will overwrite files in the powerpoint file's folder if there are name collisiona
' Original version by Louis from StackExchange (https://tex.stackexchange.com/users/6321/louis) available here: (https://tex.stackexchange.com/questions/66007/any-way-of-converting-ppt-or-odf-to-beamer-or-org)
' Modified by Jason Kerwin (www.jasonkerwin.com) on 20 February 2018:
    ' Takes out extra text that printed in the title line
    ' Switches titles to \frametitle{} instead of including them on the \begin{frame} line (sometimes helps with compiling)
    ' Changes the image names to remove original filename, which might have spaces
    ' Removes "\subsection{}" which was printing before each slide
'NB you must convert your slides to .ppt format before running this code

Public Sub ConvertToBeamer()
    Dim objPresentation As Presentation
    Set objPresentation = Application.ActivePresentation

    Dim objSlide As slide
    Dim objshape As Shape
    Dim objShape4Note As Shape
    Dim hght As Long, wdth As Long
    Dim objFileSystem
    Dim objGrpItem As Shape

    Dim Name As String, Pth As String, Dest As String, IName As String, ln As String, ttl As String, BaseName As String
    Dim txt As String
    Dim p As Integer, l As Integer, ctr As Integer, i As Integer, j As Integer
    Dim il As Long, cl As Long
    Dim Pgh As TextRange

    If check_embedded_videos Then
        MsgBox "File Has Embedded Videos"
    Else

    PresName = Application.ActivePresentation.Name
    PresName = strip_ext(PresName)
    PresName = strip_chars(PresName, ".", "-")
    PresName = strip_chars(PresName, " ", "-")
    PresName = LCase(PresName)
    
    Name = PresName & ".tex"
    yamlname = PresName & "-video-info.yaml"

    'l = Len(Name)
    
    Pth = Application.ActivePresentation.path
    Dest = Pth & "/" & Name
    yamldest = Pth & "/" & yamlname
    ctr = 0
    'Dim objTextFile
    'Set objFileSystem = CreateObject("Scripting.FileSystemObject")
    'Set objTextFile = objFileSystem.CreateTextFile(Dest, True, 1)
    Dim objTextFile As Object
    Set objTextFile = CreateObject("ADODB.Stream")
    objTextFile.Type = 2 'Specify stream type - we want To save text/string data.
    objTextFile.Charset = "utf-8" 'Specify charset For the source text data.
    objTextFile.Open

    objTextFile.WriteText "\section{" & prep_for_latex(Name) & "}" & vbLf
    
    With Application.ActivePresentation.PageSetup
        wdth = .SlideWidth
        hght = .SlideHeight
    End With

    ungroup_all
    
    ans = VBA.FileSystem.Dir(Pth & "\" & PresName & "-images\", vbDirectory)

    If ans = vbNullString Then
        MkDir Pth & "\" & PresName & "-images\"
    End If

    
       
    Dim yamlfile As Object
    Set yamlfile = CreateObject("ADODB.Stream")
    yamlfile.Type = 2 'Specify stream type - we want To save text/string data.
    yamlfile.Charset = "utf-8" 'Specify charset For the source text data.
    yamlfile.Open
    
    
    Dim ii As Integer
    ii = 0

    For Each objSlide In objPresentation.Slides
        objTextFile.WriteText "" & vbLf
        ttl = "No Title"
        If objSlide.Shapes.HasTitle Then
          ttl = objSlide.Shapes.Title.TextFrame.TextRange.Text
        End If

        ' objTextFile.WriteText "\subsection{" & ttl & "}"  & vbLf
        objTextFile.WriteText "\begin{frame}" & vbLf
        objTextFile.WriteText "\frametitle{" & prep_for_latex(ttl) & "}" & vbLf
       
        For Each objshape In objSlide.Shapes

            If objshape.HasTextFrame = True Then
                If Not objshape.TextFrame.TextRange Is Nothing Then
                    il = 0
                    For Each Pgh In objshape.TextFrame.TextRange.Paragraphs

                        If Not objshape.TextFrame.TextRange.Text = ttl Then
                            cl = Pgh.Paragraphs.IndentLevel
                            txt = Pgh.TrimText
                            If cl > il Then
                               objTextFile.WriteText "\begin{itemize}" & vbLf
                                il = cl
                            ElseIf cl < il Then
                                objTextFile.WriteText "\end{itemize}" & vbLf
                                il = cl
                            End If
                            If il = 0 Then
                                objTextFile.WriteText prep_for_latex(txt) & vbLf
                            Else
                                objTextFile.WriteText "\item " + prep_for_latex(txt) & vbLf
                            End If
                        End If
                    Next Pgh
                    If il > 0 Then
                      For i = 1 To il
                        objTextFile.WriteText "\end{itemize}" & vbLf
                      Next i
                    End If
                End If
            ElseIf objshape.HasTable Then
              ln = "\begin{tabular}{|"
              For j = 1 To objshape.Table.Columns.Count
              ln = ln & "l|"
              Next j
              ln = ln & "} \hline"
              objTextFile.WriteText prep_for_latex(ln) & vbLf
              With objshape.Table
                For i = 1 To .Rows.Count
                    If .Cell(i, 1).Shape.HasTextFrame Then
                        ln = prep_for_latex(.Cell(i, 1).Shape.TextFrame.TextRange.Text)
                    End If

                    For j = 2 To .Columns.Count
                        If .Cell(i, j).Shape.HasTextFrame Then
                            ln = ln & " & " & prep_for_latex(.Cell(i, j).Shape.TextFrame.TextRange.Text)
                        End If
                    Next j
                    ln = ln & "  \\ \hline"
                    objTextFile.WriteText ln & vbLf
                Next i
                objTextFile.WriteText "\end{tabular}" & vbLf & vbLf
              End With
            ElseIf (objshape.Type = msoGroup) Then
                For Each objGrpItem In objshape.GroupItems
                    If objGrpItem.HasTextFrame = True Then
                        If Not objGrpItem.TextFrame.TextRange Is Nothing Then
                           shpx = objGrpItem.Top / hght
                           shpy = objGrpItem.Left / wdth
                           ' this could need adjustment (Footers textblocks)
                           If shpx < 0.1 And shpy > 0.5 Then
                            objTextFile.WriteText ("%BookTitle: " & prep_for_latex(objGrpItem.TextFrame.TextRange.Text)) & vbLf
                            ElseIf shpx < 0.1 And shpy < 0.5 Then
                            objTextFile.WriteText ("%FrameTitle: " & prep_for_latex(objGrpItem.TextFrame.TextRange.Text)) & vbLf
                            Else
                            objTextFile.WriteText ("%PartTitle: " & prep_for_latex(objGrpItem.TextFrame.TextRange.Text)) & vbLf
                           End If
                        End If
                    End If
                 Next objGrpItem
            ElseIf (objshape.Type = msoPicture) Then
                IName = "img" & Format(ctr, "0000") & ".png"
                objTextFile.WriteText "\includegraphics{" & PresName & "-images/" & IName & "}" & vbLf
                Call objshape.Export(Pth & "\" & PresName & "-images\" & IName, ppShapeFormatPNG, , , ppRelativeToSlide)
                ctr = ctr + 1
            ElseIf objshape.Type = msoEmbeddedOLEObject Then
                If objshape.OLEFormat.ProgID = "Equation.3" Then
                    IName = "img" & Format(ctr, "0000") & ".png"
                    objTextFile.WriteText "\includegraphic{" & IName & "}" & vbLf
                    Call objshape.Export(Pth & "\" & IName, ppShapeFormatPNG, , , ppRelativeToSlide)
                    ctr = ctr + 1
                End If
        
            ElseIf (objshape.Type = msoMedia) Then
                If objshape.MediaType = ppMediaTypeMovie Then
                    If objshape.MediaFormat.IsLinked Then
                        linked_path = objshape.LinkFormat.SourceFullName
                        linked_path = Replace(linked_path, "\", "/")
                        
                        fnr = file_from_path(linked_path)
                        fnr = strip_ext(fnr)
                        fnr = strip_chars(fnr, ".", "-")
                        fnr = strip_chars(fnr, " ", "-")
                        fnr = LCase(fnr)
                        fnr = fnr & "-" & Format(ctr, "0000")
                        image_path = "thumbs/" & fnr & ".png"
                        video_path = "videos/" & fnr & ".mp4"
                    
                        objTextFile.WriteText "\begin{figure}[h!]" & vbLf
                        objTextFile.WriteText "\centering" & vbLf
                        objTextFile.WriteText "\movie[externalviewer]" & vbLf
                        objTextFile.WriteText "{\includegraphics[width=\columnwidth]{" & image_path & "}}{" & video_path & "}" & vbLf
                        objTextFile.WriteText "\caption{caption}" & vbLf
                        objTextFile.WriteText "\end{figure}" & vbLf
                        
                        ExtractVideoInfoInner objshape, yamlfile, Format(ctr, "0000")

                        ctr = ctr + 1
                    Else
                        MsgBox "This ppt has embedded videos"
                    End If
                End If
            End If
        
        Next objshape
        
        On Error GoTo errMyErrorHandler:
            Notes = objSlide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text
            objTextFile.WriteText vbLf & vbLf & "% " & Notes & vbLf & vbLf
        On Error GoTo 0

DoCleanUp:
        objTextFile.WriteText "\end{frame}" & vbLf & vbLf
        'to test on the first 3 slides
        'If objSlide.SlideIndex >= 3 Then
        '    Exit For
        'End If
        
    ii = ii + 1
    Next objSlide

    'objTextFile.Close
    'Set objTextFile = Nothing
    'Set objFileSystem = Nothing
    objTextFile.SaveToFile Dest, 2 'Save binary data To disk
    yamlfile.SaveToFile yamldest, 2

    End If
Exit Sub

errMyErrorHandler:
    Resume DoCleanUp
End Sub

Private Function prep_for_latex(input1 As String) As String
    Dim output As String
    output = Replace(input1, "&", "\&")
    prep_for_latex = output
    End Function

Public Sub ConvertToMarkdown()
    Dim objPresentation As Presentation
    Set objPresentation = Application.ActivePresentation

    Dim objSlide As slide
    Dim objshape As Shape
    Dim objShape4Note As Shape
    Dim hght As Long, wdth As Long
    Dim objFileSystem
    Dim objGrpItem As Shape

    Dim Name As String, Pth As String, Dest As String, IName As String, ln As String, ttl As String, BaseName As String
    Dim txt As String
    Dim p As Integer, l As Integer, ctr As Integer, i As Integer, j As Integer
    Dim il As Long, cl As Long
    Dim Pgh As TextRange

    Dim indent_level As String
    
    If check_embedded_videos Then
        MsgBox "File Has Embedded Videos"
    Else

    PresName = Application.ActivePresentation.Name
    PresName = strip_ext(PresName)
    PresName = strip_chars(PresName, ".", "-")
    PresName = strip_chars(PresName, " ", "-")
    PresName = LCase(PresName)

    Name = PresName & ".md"
    yamlname = PresName & "-video-info.yaml"
    
    
    Pth = Application.ActivePresentation.path
    Dest = Pth & "/" & Name
    yamldest = Pth & "/" & yamlname
    ctr = 0
    
    Dim objTextFile As Object
    Set objTextFile = CreateObject("ADODB.Stream")
    objTextFile.Type = 2 'Specify stream type - we want To save text/string data.
    objTextFile.Charset = "utf-8" 'Specify charset For the source text data.
    objTextFile.Open

    objTextFile.WriteText "---" & vbLf
    objTextFile.WriteText "---" & vbLf & vbLf
    objTextFile.WriteText "# " & (Name) & vbLf & vbLf
    
    With Application.ActivePresentation.PageSetup
        wdth = .SlideWidth
        hght = .SlideHeight
    End With

    ungroup_all

    ans = VBA.FileSystem.Dir(Pth & "\" & PresName & "-images\", vbDirectory)

    If ans = vbNullString Then
        MkDir Pth & "\" & PresName & "-images\"
    End If
    
    
    Dim yamlfile As Object
    Set yamlfile = CreateObject("ADODB.Stream")
    yamlfile.Type = 2 'Specify stream type - we want To save text/string data.
    yamlfile.Charset = "utf-8" 'Specify charset For the source text data.
    yamlfile.Open
        
    Dim ii As Integer
    ii = 0

    For Each objSlide In objPresentation.Slides
        objSlide.Select
        
        If objSlide.SlideShowTransition.Hidden Then
            objTextFile.WriteText "<!-- " & vbLf & vbLf
        End If
        
        
        
        ttl = ""
        
        If objSlide.Shapes.HasTitle Then
          ttl = objSlide.Shapes.Title.TextFrame.TextRange.Text
        End If
        
        objTextFile.WriteText "## " & (ttl) & vbLf & vbLf

        For Each objshape In objSlide.Shapes
            objshape.Select
            

            If objshape.HasTextFrame = True Then
                If Not objshape.TextFrame.TextRange Is Nothing Then
                    il = 0
                    For Each Pgh In objshape.TextFrame.TextRange.Paragraphs

                        If Not objshape.TextFrame.TextRange.Text = ttl Then
                            cl = Pgh.Paragraphs.IndentLevel
                            indent_level = ""
                            For i = 1 To (cl - 1)
                                indent_level = indent_level & "  "
                            Next
                            txt = Pgh.TrimText
                            objTextFile.WriteText indent_level & "* " & (txt) & vbLf
                       
                       End If
                    
                    Next
                    objTextFile.WriteText vbLf
                End If
            
            ElseIf objshape.HasTable Then
                Dim jj As Integer
                ii = objshape.Table.Rows.Count
                jj = objshape.Table.Columns.Count
                txt = ""
                
                i = 1
                For j = 1 To jj
                    If objshape.Table.Cell(i, j).Shape.HasTextFrame Then
                        txt = txt & "| " & objshape.Table.Cell(i, j).Shape.TextFrame.TextRange.Text & " "
                    Else
                        txt = txt & "| "
                    End If
                Next
                objTextFile.WriteText txt & "|" & vbLf
                
                txt = ""
                For j = 1 To jj
                    txt = txt & "|-"
                Next
                objTextFile.WriteText txt & "|" & vbLf
                
                For i = 2 To ii
                    txt = ""
                    For j = 1 To jj
                        If objshape.Table.Cell(i, j).Shape.HasTextFrame Then
                            txt = txt & "| " & objshape.Table.Cell(i, j).Shape.TextFrame.TextRange.Text & " "
                        Else
                            txt = txt & "| "
                        End If
                    Next
                    objTextFile.WriteText txt & "|" & vbLf
                Next
                
                objTextFile.WriteText vbLf
            
            ElseIf (objshape.Type = msoGroup) Then
                For Each objGrpItem In objshape.GroupItems
                    If objGrpItem.HasTextFrame = True Then
                        If Not objGrpItem.TextFrame.TextRange Is Nothing Then
                            objTextFile.WriteText objGrpItem.TextFrame.TextRange.Text & vbLf
                           End If
                        End If
                 Next objGrpItem
            
            ElseIf (objshape.Type = msoPicture) Then
                Debug.Print objshape.Name
                
                IName = "img" & Format(ctr, "0000") & ".png"
                Call objshape.Export(Pth & "\" & PresName & "-images\" & IName, ppShapeFormatPNG, , , ppRelativeToSlide)
                objTextFile.WriteText "![](" & PresName & "-images/" & IName & "){width=100%}" & vbLf & vbLf
                ctr = ctr + 1
                
            'ElseIf objshape.Type = msoEmbeddedOLEObject Then
                'If objshape.OLEFormat.ProgID = "Equation.3" Then
                    'IName = "img" & Format(ctr, "0000") & ".png"
                    'Call objshape.Export(Pth & "\" & IName, ppShapeFormatPNG, , , ppRelativeToSlide)
                    'objTextFile.WriteText "![](" & IName & "){width=100%}" & vbLf & vbLf
                    'ctr = ctr + 1
                'End If
        
            ElseIf (objshape.Type = msoMedia) Then
                If objshape.MediaType = ppMediaTypeMovie Then
                    If objshape.MediaFormat.IsLinked Then
                        linked_path = objshape.LinkFormat.SourceFullName
                        linked_path = Replace(linked_path, "\", "/")
                        
                        fnr = file_from_path(linked_path)
                        fnr = strip_ext(fnr)
                        fnr = strip_chars(fnr, ".", "-")
                        fnr = strip_chars(fnr, " ", "-")
                        fnr = LCase(fnr)
                        fnr = fnr & "-" & Format(ctr, "0000")
                        image_path = "thumbs/" & fnr & ".png"
                        video_path = "videos/" & fnr & ".mp4"
                        
                        objTextFile.WriteText "![[Video](" & video_path & ")](" & image_path & ")" & vbLf & vbLf
                        
                        ExtractVideoInfoInner objshape, yamlfile, Format(ctr, "0000")

                        ctr = ctr + 1
                    Else
                        MsgBox "This ppt has embedded videos"
                    End If
                End If
            End If
        
        Next objshape

        On Error GoTo errMyErrorHandler2:
            Notes = objSlide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text
            If (Notes <> "") Then
                objTextFile.WriteText vbLf & vbLf & "::: notes" & vbLf & vbLf & Notes & vbLf & vbLf & ":::" & vbLf & vbLf
            End If
                
        On Error GoTo 0
        
        If objSlide.SlideShowTransition.Hidden Then
            objTextFile.WriteText "-->" & vbLf & vbLf
        End If
        

DoCleanUp:


        'to test on the first 3 slides
        'If objSlide.SlideIndex >= 3 Then
        '    Exit For
        'End If
        
        
    ii = ii + 1
    Next objSlide

    'objTextFile.Close
    'Set objTextFile = Nothing
    'Set objFileSystem = Nothing
    objTextFile.SaveToFile Dest, 2 'Save binary data To disk
    yamlfile.SaveToFile yamldest, 2

    End If
Exit Sub

errMyErrorHandler2:
    Resume DoCleanUp
End Sub

Sub WhatIsTheEquationMadeOf()

Dim ii As Integer
Dim objshape As Shape

Set objshape = ActiveWindow.Selection.ShapeRange(1)

ii = objshape.TextFrame2.TextRange.MathZones.Count
Debug.Print ii
For i = 1 To ii
Debug.Print objshape.TextFrame2.TextRange.MathZones(i).Text
Next
End Sub

Public Sub ExtractVideoInfo()

    Dim objPresentation As Presentation
    Set objPresentation = Application.ActivePresentation

    Dim objSlide As slide
    Dim objshape As Shape
    Dim objShape4Note As Shape
    Dim hght As Long, wdth As Long
    Dim objFileSystem
    Dim objGrpItem As Shape

    Dim Name As String, Pth As String, Dest As String, IName As String, ln As String, ttl As String, BaseName As String
    Dim txt As String
    Dim p As Integer, l As Integer, ctr As Integer, i As Integer, j As Integer
    Dim il As Long, cl As Long
    Dim Pgh As TextRange

    If check_embedded_videos Then
        MsgBox "File Has Embedded Videos"
    Else

    PresName = Application.ActivePresentation.Name
    PresName = strip_ext(PresName)
    PresName = strip_chars(PresName, ".", "-")
    PresName = strip_chars(PresName, " ", "-")
    PresName = LCase(PresName)

    Name = PresName & "-video-info.yaml"
       
    Pth = Application.ActivePresentation.path
    Dest = Pth & "/" & Name
    ctr = 0
    
    Dim objTextFile As Object
    Set objTextFile = CreateObject("ADODB.Stream")
    objTextFile.Type = 2 'Specify stream type - we want To save text/string data.
    objTextFile.Charset = "utf-8" 'Specify charset For the source text data.
    objTextFile.Open

    Dim has_video As Boolean
    

    Dim ii As Integer
    ii = 0
    
    For Each objSlide In objPresentation.Slides
        has_video = False
        For Each objshape In objSlide.Shapes
            If (objshape.Type = msoMedia) Then
                If objshape.MediaType = ppMediaTypeMovie Then
                    If objshape.MediaFormat.IsLinked Then
                        has_video = True
                        linked_path = objshape.LinkFormat.SourceFullName
                        linked_path = Replace(linked_path, "\", "/")
                        
                        fnr = file_from_path(linked_path)
                        fnr = strip_ext(fnr)
                        fnr = strip_chars(fnr, ".", "-")
                        fnr = strip_chars(fnr, " ", "-")
                        fnr = LCase(fnr)
                        fnr = fnr & "-" & ii
                        image_path = "thumbs/" & fnr & ".png"
                        video_path = "videos/" & fnr & ".mp4"
                        
                        objTextFile.WriteText "- source: """ & linked_path & """" & vbLf
                        objTextFile.WriteText "  start_point: " & objshape.MediaFormat.StartPoint & vbLf
                        objTextFile.WriteText "  end_point: " & objshape.MediaFormat.EndPoint & vbLf
                        objTextFile.WriteText "  length: " & objshape.MediaFormat.Length & vbLf
                        objTextFile.WriteText "  fade_in_duration: " & objshape.MediaFormat.FadeInDuration & vbLf
                        objTextFile.WriteText "  fade_out_duration: " & objshape.MediaFormat.FadeOutDuration & vbLf
                        objTextFile.WriteText "  loop: " & objshape.AnimationSettings.PlaySettings.LoopUntilStopped & vbLf
                        objTextFile.WriteText "  pause: " & objshape.AnimationSettings.PlaySettings.PauseAnimation & vbLf
                        objTextFile.WriteText "  play_on_entry: " & objshape.AnimationSettings.PlaySettings.PlayOnEntry & vbLf
                        objTextFile.WriteText "  rewind: " & objshape.AnimationSettings.PlaySettings.RewindMovie & vbLf
                        'objTextFile.WriteText "    stop_after: " & objshape.AnimationSettings.PlaySettings.StopAfterSlides & vbLf
                        objTextFile.WriteText "  hide_while_not_playing: " & objshape.AnimationSettings.PlaySettings.HideWhileNotPlaying & vbLf
                        objTextFile.WriteText "  slide_number: " & ii & vbLf
                        objTextFile.WriteText "  thumb_dest_path: " & image_path & vbLf
                        objTextFile.WriteText "  video_dest_path: " & video_path & vbLf
                        
                    Else
                        MsgBox "This ppt has embedded videos"
                    End If
                End If
            End If
        Next objshape
    
        'objTextFile.WriteText "- null" & vbLf
    
    ii = ii + 1
    
    Next objSlide

    objTextFile.SaveToFile Dest, 2
    End If
End Sub


Public Sub ExtractVideoInfoInner(objshape As Shape, objTextFile As Object, ctr As String)
    linked_path = objshape.LinkFormat.SourceFullName
    linked_path = Replace(linked_path, "\", "/")
    
    fnr = file_from_path(linked_path)
    fnr = strip_ext(fnr)
    fnr = strip_chars(fnr, ".", "-")
    fnr = strip_chars(fnr, " ", "-")
    fnr = LCase(fnr)
    fnr = fnr & "-" & ctr
    image_path = "thumbs/" & fnr & ".png"
    video_path = "videos/" & fnr & ".mp4"
    
    objTextFile.WriteText "- source: """ & linked_path & """" & vbLf
    objTextFile.WriteText "  start_point: " & objshape.MediaFormat.StartPoint & vbLf
    objTextFile.WriteText "  end_point: " & objshape.MediaFormat.EndPoint & vbLf
    objTextFile.WriteText "  length: " & objshape.MediaFormat.Length & vbLf
    objTextFile.WriteText "  fade_in_duration: " & objshape.MediaFormat.FadeInDuration & vbLf
    objTextFile.WriteText "  fade_out_duration: " & objshape.MediaFormat.FadeOutDuration & vbLf
    objTextFile.WriteText "  loop: " & objshape.AnimationSettings.PlaySettings.LoopUntilStopped & vbLf
    objTextFile.WriteText "  pause: " & objshape.AnimationSettings.PlaySettings.PauseAnimation & vbLf
    objTextFile.WriteText "  play_on_entry: " & objshape.AnimationSettings.PlaySettings.PlayOnEntry & vbLf
    objTextFile.WriteText "  rewind: " & objshape.AnimationSettings.PlaySettings.RewindMovie & vbLf
    'objTextFile.WriteText "    stop_after: " & objshape.AnimationSettings.PlaySettings.StopAfterSlides & vbLf
    objTextFile.WriteText "  hide_while_not_playing: " & objshape.AnimationSettings.PlaySettings.HideWhileNotPlaying & vbLf
    objTextFile.WriteText "  slide_number: " & ii & vbLf
    objTextFile.WriteText "  thumb_dest_path: " & image_path & vbLf
    objTextFile.WriteText "  video_dest_path: " & video_path & vbLf
End Sub

Public Function ungroup() As Boolean

    Dim objPresentation As Presentation
    Set objPresentation = Application.ActivePresentation

    Dim objSlide As slide
    Dim objshape As Shape

    Dim nogroups As Boolean
    groups = False
    
    For Each objSlide In objPresentation.Slides
       
        For Each objshape In objSlide.Shapes
            
            If (objshape.Type = msoGroup) Then
                groups = True
                objshape.ungroup
            End If
             'asdf (objshape)
            
        Next objshape
    Next objSlide
    'If groups = False Then
     '   MsgBox "no groups"
    'End If
    ungroup = groups
End Function

Public Sub ungroup_all()
    Do
        test = ungroup()
    Loop While test
End Sub

Public Sub export_me()
PNGIfy
'ConvertToBeamer
ConvertToMarkdown
'ExtractVideoInfo
End Sub

