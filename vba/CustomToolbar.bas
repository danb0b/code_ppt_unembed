Attribute VB_Name = "CustomToolbar"
Private Const TOOLBARNAME = "IDEAlab Tools"
Sub make_button(tb As CommandBar, act As String, cap As String, desc As String, style As Integer, face As Integer)
    'add a button to the new toolbar
    Dim button As CommandBarButton
    Set button = tb.Controls.Add(Type:=msoControlButton)
    With button
         .DescriptionText = desc
         .caption = cap
         .OnAction = act
         .faceId = face
    End With
    button.style = msoButtonIconAndCaptionBelow
End Sub

Sub ShowToolbar()
    Dim oToolbar As CommandBar
    Dim MyToolbar As String

    ' Give the toolbar a name
    'MyToolbar = "Kewl Tools"

    On Error Resume Next
    ' so that it doesn't stop on the next line if the toolbar's already there

    ' Create the toolbar; PowerPoint will error if it already exists
    Set oToolbar = CommandBars.Add(Name:=TOOLBARNAME, Position:=msoBarFloating, Temporary:=True)
    If Err.Number <> 0 Then
          ' The toolbar's already there, so we have nothing to do
          Exit Sub
    End If

    On Error GoTo ErrorHandler

    make_button oToolbar, "EmbeddedMoviesToLinkedMovies", "Link All Movies", "Link All Movies", 526, msoButtonIconAndCaption
    make_button oToolbar, "SwitchPath", "Switch Movies Folder", "Switch Movies Folder", 526, msoButtonIconAndCaption
    make_button oToolbar, "PNGIfy", "PNGIfy All Images", "PNGIfy All Images", 526, msoButtonIconAndCaption
    make_button oToolbar, "export_me", "Export PPT", "Export PPT", 526, msoButtonIconAndCaption
    'make_button oToolbar, "ConvertToBeamer", "Convert To Beamer", "Convert To Beamer", 526, msoButtonIconAndCaption
    'make_button oToolbar, "ConvertToMarkdown", "Convert To Markdown", "Convert To Markdown", 526, msoButtonIconAndCaption
    'make_button oToolbar, "ExtractVideoInfo", "Extract Video Info", "Extract Video Info", 526, msoButtonIconAndCaption
    
    oToolbar.Top = 150
    oToolbar.Left = 150
    oToolbar.Visible = True

NormalExit:
    Exit Sub   ' so it doesn't go on to run the errorhandler code

ErrorHandler:
     'Just in case there is an error
     MsgBox Err.Number & vbLfLf & Err.Description
     Resume NormalExit:
End Sub

Public Sub DeleteCommandBar()
    Application.CommandBars(TOOLBARNAME).Delete
End Sub

Sub presentationBeforeClose(Cancel As Boolean)
    DeleteCommandBar
End Sub

Sub PresentationOpen()
    ShowToolbar
End Sub

Sub Auto_Open()
    ShowToolbar
End Sub
