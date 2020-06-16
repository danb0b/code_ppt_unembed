Attribute VB_Name = "display_linked_path"
Function display_linked_path() As Boolean

    Dim oSl As slide
    Dim oSh As Shape
    Dim x As Long
    Dim sPath As String
    
    ' Edit this to the path where vids are stored,
    ' path must end in backslash
    
    sPath = "null"
    'sPath = "C:\Users\danaukes\Desktop\"
    Dim ii As Integer
    ii = 0
    For Each oSl In ActivePresentation.Slides
        oSl.Select
        
        For x = oSl.Shapes.Count To 1 Step -1
            
            Set oSh = oSl.Shapes(x)
            
            If oSh.Type = msoMedia Then
                If oSh.MediaType = ppMediaTypeMovie Then
                    If oSh.MediaFormat.IsLinked Then
                        linked_path = oSh.LinkFormat.SourceFullName
                        MsgBox linked_path
                    Else
                        MsgBox "This ppt has embedded videos"
                    End If
                End If
            End If
        Next
    Next
End Function



Function check_embedded_videos() As Boolean

    Dim oSl As slide
    Dim oSh As Shape
    Dim x As Long
    Dim sPath As String
    
    sPath = "null"
    
    Dim result As Boolean
    result = False
    
    Dim ii As Integer
    ii = 0
    
    For Each oSl In ActivePresentation.Slides
        oSl.Select
        
        For x = oSl.Shapes.Count To 1 Step -1
            
            Set oSh = oSl.Shapes(x)
            result = obstest(oSh, result)
        
        Next
    Next
    check_embedded_videos = result
End Function

Function obstest(a As Variant, ByRef result As Boolean) As Boolean
        Dim objshape As Shape
        Set objshape = a

        'Dim result As Boolean
        'result = False
        objshape.Select
        
        'Set oSh = oSl.Shapes(x)
        
        If objshape.HasTextFrame = True Then
        ElseIf objshape.HasTable = True Then
        ElseIf (objshape.Type = msoPicture) Then
        ElseIf objshape.Type = msoEmbeddedOLEObject Then
        ElseIf (objshape.Type = msoMedia) Then
            If objshape.MediaType = ppMediaTypeMovie Then
                If objshape.MediaFormat.IsEmbedded Then
                    result = True
                End If
            End If
        ElseIf (objshape.Type = msoGroup) Then
            For Each objGrpItem In objshape.GroupItems
                objGrpItem.Select
                
                'objGrpItem.Application

                dummy = obstest(objGrpItem, result)
            Next objGrpItem
        End If
        obstest = result
End Function

Sub obstest_outer()
    If check_embedded_videos = True Then
        MsgBox "File Has Embedded Videos"
    End If

End Sub
