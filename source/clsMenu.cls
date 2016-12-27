VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' -------------------------------------------------------------------------------
'
'  DVD Flick - A DVD authoring program
'  Copyright (C) 2006-2009  Dennis Meuwissen
'
'  This program is free software; you can redistribute it and/or modify
'  it under the terms of the GNU General Public License as published by
'  the Free Software Foundation; either version 2 of the License, or
'  (at your option) any later version.
'
'  This program is distributed in the hope that it will be useful,
'  but WITHOUT ANY WARRANTY; without even the implied warranty of
'  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'  GNU General Public License for more details.
'
'  You should have received a copy of the GNU General Public License
'  along with this program; if not, write to the Free Software
'  Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301, USA.
'
' -------------------------------------------------------------------------------
'
'   File purpose: Base menu object.
'
Option Explicit
Option Compare Binary
Option Base 0


' Properties
Public backgroundImage As String
Public highlightColor As Long
Public selectColor As Long

' Render images
Public imgBack As clsGDIImage
Public imgHighlight As clsGDIImage
Public imgSelect As clsGDIImage

' Menu objects
Public Objects As New Dictionary
Public Buttons As New Dictionary


' Generate multiplexed MPEG for this menu
' TODO: Generalize this, maybe with special functions to convert a PNG to m2v, with audio
Public Function generateFiles(ByVal preFix As String, ByVal Index As Long, Optional ByVal titleIndex As Long = -1) As Boolean

    Dim baseName As String
    Dim Stream As clsTextFile
    Dim cmdLine As String
    Dim Pic As StdPicture
            

    baseName = Project.destinationDir & "\" & preFix & "menu"
    If titleIndex <> -1 Then baseName = baseName & titleIndex
    baseName = baseName & "_"
    
    Dbg.addLog DM_Menus, "Generating files for menu. Basename " & baseName
    
    
    ' Write spumux file
    Set Stream = New clsTextFile
    Stream.fileCreate baseName & Index & ".xml", encodeASCIIorUTF8, CODEPAGE_LATIN1
    Stream.writeText generateButtonsXML(baseName, Index)
    Set Stream = Nothing
    
    
    ' Output images
    Dbg.addLog DM_Menus, "Writing bitmap images"
    Set Pic = imgBack.getPicture
    If Pic Is Nothing Then
        frmDialog.Display "Could not get menu background picture to save.", OkOnly Or Critical
        Exit Function
    Else
        SavePicture Pic, baseName & "b" & Index & ".bmp"
    End If
    
    Set Pic = imgHighlight.getPicture
    If Pic Is Nothing Then
        frmDialog.Display "Could not get menu highlight picture to save.", OkOnly Or Critical
        Exit Function
    Else
        SavePicture Pic, baseName & "h" & Index & ".bmp"
    End If
    
    Set Pic = imgSelect.getPicture
    If Pic Is Nothing Then
        frmDialog.Display "Could not get menu select picture to save.", OkOnly Or Critical
        Exit Function
    Else
        SavePicture Pic, baseName & "s" & Index & ".bmp"
    End If
    
    
    ' Convert to PNG
    Dbg.addLog DM_Menus, "Converting bitmaps to PNG"
    If Not convertToPNG(baseName & "b" & Index & ".bmp", baseName & "b" & Index & ".png", -1, True) Or _
       Not convertToPNG(baseName & "h" & Index & ".bmp", baseName & "h" & Index & ".png", 0, True) Or _
       Not convertToPNG(baseName & "s" & Index & ".bmp", baseName & "s" & Index & ".png", 0, True) Then
        frmDialog.Display "Could not convert menu images to PNG.", OkOnly Or Critical
        Exit Function
    End If
    
    
    ' Convert background to MPEG-2 video files
    Dbg.addLog DM_Menus, "Converting background to MPEG-2"
    cmdLine = cmdLine & " -loop_input -vframes 4"
    cmdLine = cmdLine & " -i " & vbQuote & baseName & "b" & Index & ".png" & vbQuote
    If Project.targetFormat = VF_MIXED Or Project.targetFormat = VF_PAL Then
        cmdLine = cmdLine & " -s 720x576"
        cmdLine = cmdLine & " -r 25"
    Else
        cmdLine = cmdLine & " -s 720x480"
        cmdLine = cmdLine & " -r 29.97"
    End If
    cmdLine = cmdLine & " -bufsize 1835008 -packetsize 2048 -muxrate 10080000"
    cmdLine = cmdLine & " -aspect 4:3 -f mpeg2video -qmin 1.0 -qmax 4.0 -qscale 1.0 -minrate 8000k -b 8000k -maxrate 8000k"
    cmdLine = cmdLine & " " & vbQuote & baseName & Index & ".m2v" & vbQuote
    executeToFile APP_PATH & "bin\ffmpeg.exe", cmdLine, Project.destinationDir & "\ffmpeg_menu" & Index & "_" & titleIndex & ".txt", SM_Nothing, Project.encodePriority, Project.destinationDir
    
    If Not FS.FileExists(baseName & Index & ".m2v") Then
        frmDialog.Display "Could not convert background images to MPEG2 (no file).", OkOnly Or Critical
        Exit Function
    ElseIf FileLen(baseName & Index & ".m2v") = 0 Then
        frmDialog.Display "Could not convert background images to MPEG2 (null file size).", OkOnly Or Critical
        Exit Function
    End If
    
    
    ' Multiplex audio with menu video
    Dbg.addLog DM_Menus, "Multiplexing menu audio with video"
    
    cmdLine = "-f 8 --vbr"
    cmdLine = cmdLine & " -o " & vbQuote & baseName & Index & ".mpg" & vbQuote
    cmdLine = cmdLine & " " & vbQuote & baseName & Index & ".m2v" & vbQuote
    cmdLine = cmdLine & " " & vbQuote & APP_PATH & "data\silence.ac3" & vbQuote
    
    executeToFile APP_PATH & "bin\mplex.exe", cmdLine, Project.destinationDir & "\mplex_menu" & Index & "_" & titleIndex & ".txt", SM_Nothing, Project.encodePriority, Project.destinationDir

    If Not FS.FileExists(baseName & Index & ".mpg") Then
        frmDialog.Display "Could not multiplex menu. No output file.", OkOnly Or Critical
        Exit Function
    ElseIf FileLen(baseName & Index & ".mpg") = 0 Then
        frmDialog.Display "Could not multiplex menu. Null file size.", OkOnly Or Critical
    End If

    killIfExists Project.destinationDir & "\*.m2v"


    ' Multiplex subpictures
    Dbg.addLog DM_Menus, "Multiplexing menu subpictures"
    cmdLine = vbQuote & APP_PATH & "bin\spumux.exe" & vbQuote
    cmdLine = cmdLine & " -m dvd"
    cmdLine = cmdLine & " " & vbQuote & baseName & Index & ".xml" & vbQuote
    cmdLine = cmdLine & " < " & vbQuote & baseName & Index & ".mpg" & vbQuote
    cmdLine = cmdLine & " 1> " & vbQuote & baseName & Index & "d.mpg" & vbQuote
    cmdLine = cmdLine & " 2> " & vbQuote & Project.destinationDir & "\spumux_menu" & Index & "_" & titleIndex & ".txt" & vbQuote
    
    Set Stream = New clsTextFile
    Stream.fileCreate Project.destinationDir & "\spumux.bat", encodeASCIIorUTF8, CODEPAGE_LATIN1
    Stream.writeText cmdLine
    Set Stream = Nothing
    Execute Project.destinationDir & "\spumux.bat", vbNullString, WS_HIDE, True, True
    
    If Not FS.FileExists(baseName & Index & "d.mpg") Or FileLen(baseName & Index & "d.mpg") = 0 Then
        frmDialog.Display "Could not add subpictures to menu files.", OkOnly Or Critical
        Exit Function
    End If
    FS.DeleteFile baseName & Index & ".mpg"
    
    
    generateFiles = True

End Function


' Generate DVDAuthor data
Public Function generateAuthorData() As String

    Dim A As Long
    Dim XML As String
    
    
    Dbg.addLog DM_Menus, "Generating menu authoring data"
    
    For A = 0 To Buttons.Count - 1
        XML = XML & "<button name=""" & Buttons.Keys(A) & """>" & Buttons.Items(A).Action & "</button>" & vbNewLine
    Next A
    
    generateAuthorData = XML

End Function


' Rescale this menu
' Only do this after a menu has been rendered, objects are NOT scaled, only buttons and images
Public Sub Rescale(ByVal modifX As Single, ByVal modifY As Single)

    Dim A As Long

    
    Dbg.addLog DM_Menus, "Rescaling menus using " & modifX & " " & modifY
    
    For A = 0 To Buttons.Count - 1
        With Buttons.Items(A)
            .Left = .Left * modifX
            .Right = .Right * modifX
            .Top = .Top * modifY
            .Bottom = .Bottom * modifY
        
            ' Ensure they are placed at even coordinates
            .Left = .Left + (.Left Mod 2)
            .Right = .Right + (.Right Mod 2)
            .Top = .Top + (.Top Mod 2)
            .Bottom = .Bottom + (.Bottom Mod 2)
        End With
    Next A
    
    imgBack.Resize imgBack.Width * modifX, imgBack.Height * modifY, STRETCH_DELETESCANS
    imgHighlight.Resize imgHighlight.Width * modifX, imgHighlight.Height * modifY, STRETCH_DELETESCANS
    imgSelect.Resize imgSelect.Width * modifX, imgSelect.Height * modifY, STRETCH_DELETESCANS

End Sub


' Generate SPUMux data
Public Function generateButtonsXML(ByVal preFix As String, ByVal Index As Long) As String

    Dim A As Long
    Dim XML As String
    Dim Button As clsMenuButton
    
    
    Dbg.addLog DM_Menus, "Generating button XML data"
    
    XML = XML & "<subpictures><stream><spu start=""00:00:00.00"" highlight=""" & pathEntities(preFix & "h" & Index & ".png") & """ select=""" & pathEntities(preFix & "s" & Index & ".png") & """ force=""yes"" xoffset=""0"" yoffset=""0"">" & vbNewLine
    For A = 0 To Buttons.Count - 1
        Set Button = Buttons.Items(A)
        
        XML = XML & "<button name=""" & Buttons.Keys(A) & """"
        XML = XML & " x0=" & vbQuote & Button.Left & vbQuote
        XML = XML & " y0=" & vbQuote & Button.Top & vbQuote
        XML = XML & " x1=" & vbQuote & Button.Right & vbQuote
        XML = XML & " y1=" & vbQuote & Button.Bottom & vbQuote
        XML = XML & " />" & vbNewLine
    Next A
    XML = XML & "</spu></stream></subpictures>" & vbNewLine
    
    generateButtonsXML = XML

End Function


' Generate menu buttons and their actions
Public Sub generateButtons()

    Dim A As Long
    Dim buttonAction As String
    Dim newButton As clsMenuButton
    Dim Obj As clsMenuObject
    
    
    Dbg.addLog DM_Menus, "Generating button objects and surfaces"
    
    For A = 0 To Objects.Count - 1
        Set Obj = Objects.Items(A)
        If Obj.thumbID > 0 Then Buttons("Thumb" & Obj.thumbID).addSurface Obj
        If Obj.subID > 0 Then Buttons("Sub" & Obj.subID).addSurface Obj
        If Obj.audioID > 0 Then Buttons("Audio" & Obj.audioID).addSurface Obj

        ' Actions for preset object names
        buttonAction = vbNullString
        If Obj.Name = "buttonMainMenu" Then buttonAction = "jump vmgm menu 1;"
        If Obj.Name = "buttonTitleMenu" Then buttonAction = "jump vmgm menu 2;"
        If Obj.Name = "buttonPlay" Then buttonAction = "if (g0 == 0) { g0 = 1; jump vmgm fpc; } resume;"
        If Obj.Name = "buttonResume" Then buttonAction = "resume;"
        If Obj.Name = "buttonTitleSelect" Then buttonAction = "jump vmgm menu 2;"
        If Obj.Name = "disableSubtitle" Then buttonAction = "subtitle = 32; resume;"
        If Left$(Obj.Name, 15) = "buttonAudioMenu" Then buttonAction = "jump titleset " & Val(Mid$(Obj.Name, 16)) & " menu entry audio"
        If Left$(Obj.Name, 18) = "buttonSubtitleMenu" Then buttonAction = "jump titleset " & Val(Mid$(Obj.Name, 19)) & " menu entry subtitle"
        
        ' Add as new button
        If buttonAction <> vbNullString Then
            Set newButton = New clsMenuButton
            Buttons.Add Obj.Name, newButton
            newButton.Action = buttonAction
            newButton.addSurface Obj
        End If
    Next A

End Sub


' Replace captions inside all object text
Public Sub replaceCaptions(ByRef Dict As Dictionary)

    Dim A As Long, B As Long
    

    For A = 0 To Dict.Count - 1
        For B = 0 To Objects.Count - 1
            Objects.Items(B).Text = Replace(Objects.Items(B).Text, Dict.Keys(A), Dict.Items(A))
        Next B
    Next A
    
End Sub


' Copy this menu from another
Public Sub copyFrom(ByRef Menu As clsMenu)

    Dim A As Long
    Dim newObj As clsMenuObject
    Dim copyObj As clsMenuObject
    
    
    With Menu
        backgroundImage = .backgroundImage
        highlightColor = .highlightColor
        selectColor = .selectColor
    End With
    
    For A = 0 To Menu.Objects.Count - 1
        Set newObj = New clsMenuObject
        Set copyObj = Menu.Objects.Items(A)
        
        Objects.Add Menu.Objects.Keys(A), newObj
        newObj.copyFrom copyObj
    Next A

End Sub


' Render this entire menu
Public Function Render(ByRef Template As clsMenuTemplate) As Boolean

    Dim A As Long
    Dim Obj As clsMenuObject
    Dim Button As clsMenuButton
    
    Dim Img As clsGDIImage
    Dim imgTemp As clsGDIImage
    Dim imgText As clsGDIImage
    Dim fntTemp As clsGDIFont
    Dim Pal() As rgbQuad
    
    Dim Width As Long, Height As Long
    Dim Resize As Single, rY As Long, rX As Long
    Dim TEMP_COLOR As Long
    

    TEMP_COLOR = RGB(255, 255, 255)

    ' Create background image
    Set imgBack = New clsGDIImage
    imgBack.createNew MENU_WIDTH, MENU_HEIGHT, 32
    

    Dbg.addLog DM_Menus, "Calculating menu object sizes"

    ' Load images and calculate object sizes
    For A = 0 To Objects.Count - 1
        Set Obj = Objects.Items(A)
        
        ' Image
        If LenB(Obj.Image) <> 0 Then
            
            Set Obj.imageLD = New clsGDIImage
            
            ' Referenced image
            If Obj.imageRef = False Then
                If Not Obj.imageLD.openFrom(Template.Path & Obj.Image) Then Exit Function
            End If
            
            If Obj.Width = 0 Then Obj.Width = Obj.imageLD.Width
            If Obj.Height = 0 Then Obj.Height = Obj.imageLD.Height
        
        End If
        
        ' Text
        If LenB(Obj.Text) <> 0 Then
            If Obj.Width = 0 Then Obj.Width = imgBack.getTextWidth(Obj.getGDIFont, Obj.Text)
            If Obj.Height = 0 Then Obj.Height = imgBack.getTextHeight(Obj.getGDIFont, Obj.Text)
        End If
            
        ' Select image
        If LenB(Obj.selectImage) <> 0 Then
            Set Obj.selectImageLD = New clsGDIImage
            If Not Obj.selectImageLD.openFrom(Template.Path & Obj.selectImage) Then Exit Function
            
            If Obj.Width = 0 Then Obj.Width = Obj.selectImageLD.Width
            If Obj.Height = 0 Then Obj.Height = Obj.selectImageLD.Height
            
        End If
        
        ' Match to button sizes
        If Buttons.Exists(Obj.Name) Then Buttons.Item(Obj.Name).addSurface Obj
    Next A


    ' Backdrop
    Set imgTemp = New clsGDIImage
    If Not imgTemp.openFrom(Template.Path & backgroundImage) Then Exit Function
    imgBack.renderImage imgTemp, 0, 0, MENU_WIDTH, MENU_HEIGHT, Render_Copy, STRETCH_DELETESCANS
    

    ' Render objects
    Dbg.addLog DM_Menus, "Rendering menu objects"
    For A = 0 To Objects.Count - 1
        Set Obj = Objects.Items(A)
        
        ' Image
        If Not Obj.imageLD Is Nothing Then
            
            ' Calculate final size for letterboxing
            Resize = modUtil.getResizeValue(Obj.imageLD.Width, Obj.imageLD.Height, Obj.Width, Obj.Height)
            Width = Obj.imageLD.Width * Resize
            Height = Obj.imageLD.Height * Resize
            rX = (Obj.Width / 2) - (Width / 2)
            rY = (Obj.Height / 2) - (Height / 2)
            
            ' Any outline
            If Obj.Outline Then
                imgBack.renderBox Obj.X - Obj.Outline, Obj.Y - Obj.Outline, Obj.Width + (Obj.Outline * 2), Obj.Height + (Obj.Outline * 2), Obj.outlineColor
            End If
                        
            imgBack.renderImage Obj.imageLD, rX + Obj.X, Obj.Y + rY, Width, Height, Render_Trans, STRETCH_DELETESCANS, Template.transColor
        
        ' Text
        ElseIf LenB(Obj.Text) <> 0 Then
            Set fntTemp = Obj.getGDIFont
            If Obj.selectMethod = selMethod_None Then
                fntTemp.Quality = ANTIALIASED_QUALITY
            Else
                fntTemp.Quality = NONANTIALIASED_QUALITY
            End If
            imgBack.renderText fntTemp, Obj.Text, Obj.X, Obj.Y, Obj.Width, Obj.Height, Obj.textColor, Obj.getAlignment, Obj.Outline, Obj.outlineColor
        
        End If
    Next A

    
    ' Paletted subpicture bitmap
    ReDim Pal(255)
    colorToQuad Pal(1), TEMP_COLOR
    Set Img = New clsGDIImage
    Img.createNew MENU_WIDTH, MENU_HEIGHT, 8, VarPtr(Pal(0))
    
    ' Render subpicture
    Dbg.addLog DM_Menus, "Rendering menu subpicture bitmap"
    For A = 0 To Objects.Count - 1
        Set Obj = Objects.Items(A)
        
        Select Case Obj.selectMethod
        
            ' Outline
            Case selMethod_Outline

                ' Image
                If Obj.Outline And (Not Obj.imageLD Is Nothing) Then
                    Img.renderBox Obj.X - Obj.Outline, Obj.Y - Obj.Outline, Obj.Width + (Obj.Outline * 2), Obj.Height + (Obj.Outline * 2), TEMP_COLOR
                    Img.renderBox Obj.X, Obj.Y, Obj.Width, Obj.Height, 0
                
                ' Text
                ElseIf Obj.Outline And LenB(Obj.Text) <> 0 Then
                    Set fntTemp = Obj.getGDIFont
                    fntTemp.Quality = NONANTIALIASED_QUALITY
                    Img.renderText fntTemp, Obj.Text, Obj.X, Obj.Y, Obj.Width, Obj.Height, 0, Obj.getAlignment, Obj.Outline, TEMP_COLOR
                End If
            
            ' Text
            Case selMethod_Text
                If LenB(Obj.Text) <> 0 Then
                    Set fntTemp = Obj.getGDIFont
                    fntTemp.Quality = NONANTIALIASED_QUALITY
                    Img.renderText fntTemp, Obj.Text, Obj.X + Obj.Outline, Obj.Y + Obj.Outline, Obj.Width, Obj.Height, TEMP_COLOR, Obj.getAlignment
                End If
            
            ' Image
            Case selMethod_Image
                Img.renderImage Obj.selectImageLD, Obj.X, Obj.Y, Obj.Width, Obj.Height, Render_Trans, STRETCH_DELETESCANS, Template.transColor
            
        End Select
    
    Next A
    
    
    Dbg.addLog DM_Menus, "Creating subpicture images"
    
    ' Set highlight color
    colorToQuad Pal(1), highlightColor
    Img.setPalette VarPtr(Pal(0))
    Set imgHighlight = New clsGDIImage
    imgHighlight.copyFrom Img
    
    ' Set select color
    colorToQuad Pal(1), selectColor
    Img.setPalette VarPtr(Pal(0))
    Set imgSelect = New clsGDIImage
    imgSelect.copyFrom Img
    
    
    ' Cleanup
    Set Img = Nothing
    Set imgTemp = Nothing

End Function


' Read this menu from a dictionary
Public Sub readFrom(ByRef Dict As Dictionary)

    Dim A As Long, B As Long
    Dim Index As Long
    
    Dim objDict As Dictionary
    Dim newObj As clsMenuObject
    Dim inhDict As Dictionary
    
    
    For A = 0 To Dict.Count - 1
        
        Select Case Dict.Keys(A)
        
            ' Properties
            Case "backgroundImage"
                backgroundImage = Dict.Items(A)
            Case "highlightColor"
                highlightColor = stringToColor(Dict.Items(A))
            Case "selectColor"
                selectColor = stringToColor(Dict.Items(A))
            
            ' Menu object
            Case Else
                Set objDict = Dict.Items(A)
                Set newObj = New clsMenuObject
                
                ' Inherit properties from other object
                If objDict.Exists("Inherit") Then
                    Set inhDict = Dict.Item(objDict("Inherit"))
                    For B = 0 To inhDict.Count - 1
                        If inhDict.Keys(B) <> "Parent" And Not objDict.Exists(inhDict.Keys(B)) Then objDict.Add inhDict.Keys(B), inhDict.Items(B)
                    Next B
                End If
                
                newObj.readFrom objDict, Dict.Keys(A)
                Objects.Add newObj.Name, newObj
        End Select
        
    Next A
    
    ' Delete parent type objects, they are inherited now and need not be rendered
    Do
        If Objects.Items(Index).Parent Then
            Objects.Remove Objects.Keys(Index)
        Else
            Index = Index + 1
        End If
    Loop Until Index >= Objects.Count

End Sub
