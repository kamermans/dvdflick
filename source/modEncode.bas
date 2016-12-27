Attribute VB_Name = "modEncode"
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
'   File purpose: The main (trans)coding module.
'
Option Explicit
Option Compare Binary
Option Base 0


' Log output file
Private Log As clsLog

' Sizes used by the project
Private Sizes As Dictionary

' Abort error trap
Public cancelError As Boolean

' Number of frames in currently being pulldown'ed video
Private titleFrameNum As Long

' FPS counter and time for video encoding
Private previousFrameTime As Single
Private previousFrames As Long
Private currentFrames As Long

' PSNR value
Private PSNR As Single


' Please accept my sincerest apologies for using GoTos.
Public Sub Start()

    Dim errorMsg As String
    
    
    cancelError = False
    
    ' Setup log
    Set Log = New clsLog
    Log.Start Project.destinationDir & "\dvdflick.log"
    Log.Add App.Title & " " & versionString
    
    On Error GoTo encodeError
    
    ' Setup progress form
    frmMain.Hide
    frmProgress.Setup
    frmProgress.resetStatus
    frmProgress.Show
    
    
    ' Calculate sizes of streams and BitRates
    Log.Add "Calculating stream sizes..."
    
    Set Sizes = Project.calculateSizes
    Log.Add "Disc space used : " & Sizes("sizeUsed") & " KB", 1
    Log.Add "Average bitrate : " & Sizes("avgBitRate") & " Kbit/s", 1
    Log.Add "Total duration  : " & visualTime(Project.Duration), 1
    
    ' Output some general project info
    Log.Add Project.Titles.Count & " title(s)."
    Select Case Project.encodeProfile
        Case VE_Best: Log.Add "Using Best encoding profile."
        Case VE_Fast: Log.Add "Using Fast encoding profile."
        Case VE_Fastest: Log.Add "Using Fastest encoding profile."
        Case VE_Normal: Log.Add "Using Normal encoding profile."
    End Select
    
    Select Case Project.targetFormat
        Case VF_PAL: Log.Add "PAL target format."
        Case VF_NTSC: Log.Add "NTSC target format."
        Case VF_NTSCFILM: Log.Add "NTSC-Film target format."
        Case VF_MIXED: Log.Add "Mixed target format."
    End Select
    
    Log.Add "CPU: " & cpuInfo.getProcessorBrandString
    Log.Add "Threadcount: " & Project.threadCount
    
    frmProgress.Advance
    
    
    ' Prevent standby\sleep mode being entered by the power management (but not by user)
    SetThreadExecutionState ES_CONTINUOUS Or ES_SYSTEM_REQUIRED Or ES_AWAYMODE_REQUIRED
    
    
    ' Encode video
    Log.Add "Encoding video..."
    encodeVideo
    If cancelError = True Then GoTo encodeError
    
    
    ' Concatenate video
    Log.Add "Combining video sources..."
    concatVideo
    If cancelError = True Then GoTo encodeError
    
    frmProgress.Advance
    
    
    ' Encode audio
    Log.Add "Encoding audio..."
    encodeAudio
    If cancelError = True Then GoTo encodeError
 
    
    ' Concatenate audio
    Log.Add "Combining audio sources..."
    concatAudio
    If cancelError = True Then GoTo encodeError
    
    frmProgress.Advance


    ' Mux audio and video streams together
    Log.Add "Combining sources..."
    muxStreams
    If cancelError = True Then GoTo encodeError
    
    frmProgress.Advance
    
    
    ' Mux in subtitles
    Log.Add "Adding subtitles..."
    muxSubs
    If cancelError = True Then GoTo encodeError
    
    frmProgress.Advance

    
    ' Author DVD
    Log.Add "Authoring DVD..."
    authorDVD
    If cancelError = True Then GoTo encodeError
    
    frmProgress.Advance
    
    
    ' Finalize
    Log.Add "Finalizing..."
    Finalize
    If cancelError = True Then GoTo encodeError
    
    frmProgress.Advance
    
    
    ' Done!
    frmProgress.setStatus "Finished."
    Log.Add "Finished."
    
    
    ' Allow standby\sleep again
    SetThreadExecutionState ES_CONTINUOUS
    
    
    ' When done...
    Select Case frmProgress.cmbWhenDone.ListIndex

        Case 1
            Log.Add "Powering off..."
            
            Project.Modified = False
            Set Log = Nothing
            Set Sizes = Nothing
            
            adjustShutdownToken
            ExitWindowsEx EWX_POWEROFF, &H80000000
            
        Case 2
            Log.Add "Rebooting..."
            
            Project.Modified = False
            Set Log = Nothing
            Set Sizes = Nothing
        
            adjustShutdownToken
            ExitWindowsEx EWX_REBOOT, &H80000000
            
        Case 3
            Log.Add "Entering standby mode..."
            SetSystemPowerState 0, 0
            
    End Select
    
    
EndUp:
    frmProgress.Finish
    Set Log = Nothing
    Set Sizes = Nothing
    
    On Error GoTo 0
    
    Exit Sub
    

' Error occured, log, display and end
encodeError:

    If Err.Number <> 0 Then
        errorMsg = Err.Number & " from " & Err.Source & ": " & Err.Description
        Log.Add errorMsg
    ElseIf cancelError Then
        errorMsg = "Aborted by user."
    Else
        errorMsg = "Unknown error, please view the log for more information."
    End If
    
    On Error GoTo 0
    
    frmProgress.tElapsed.Enabled = False
    Set Log = Nothing
    
    createErrorLog Err.Number, Err.Source, Err.Description
    
    SetThreadExecutionState ES_CONTINUOUS
    
    If Not unattendMode Then
        frmEncodeError.Setup errorMsg
        frmEncodeError.Show 1
    End If
    
    GoTo EndUp
    
End Sub


Private Sub Finalize()

    Dim A As Long
    Dim Found As Boolean
    Dim cmdLine As String
    Dim Burner As clsBurnerDevice
    Dim driveL As String
    
    
    frmProgress.setStatus "Finalizing"
    
    ' Create ISO image
    If Project.createISO Then
    
        Log.Add "Creating ISO image", 1
        frmProgress.setSubStatus "Creating ISO image..."
    
        ' Cmdline parameters
        cmdLine = "/mode build /buildmode imagefile /noimagedetails /rootfolder yes /filesystem ""iso9660 + udf"" /udfrevision ""1.02"" /close /nosavesettings /recursesubdirs yes /start"
        cmdLine = cmdLine & " /src " & vbQuote & Project.destinationDir & "\dvd" & vbQuote
        cmdLine = cmdLine & " /volumelabel " & vbQuote & Project.discLabel & vbQuote
        cmdLine = cmdLine & " /portable"
        cmdLine = cmdLine & " /dest " & vbQuote & Project.destinationDir & "\dvd.iso" & vbQuote
        cmdLine = cmdLine & " /log " & vbQuote & Project.destinationDir & "\imgburn_isobuild.txt" & vbQuote
        
        ' Execute
        Log.Add cmdLine, 1
        Execute vbQuote & APP_PATH & "imgburn\imgburn.exe" & vbQuote, cmdLine, WS_Normal, True
        
        ' ISO created?
        If Not FS.FileExists(Project.destinationDir & "\dvd.iso") Then
            Err.Raise -1, "Finalize", "ISO image was not created."
        End If
   
    End If
    
    
    ' Burn to disc
    If Project.enableBurning Then
    
        Log.Add "Burning to disc", 1
        frmProgress.setSubStatus "Burning to disc..."
        
        ' Get recorder drive letter
        Found = False
        For A = 0 To Burners.deviceCount - 1
            Set Burner = Burners.getDevice(A)
            If Project.burnerName = Burner.deviceName & " (" & Burner.deviceDriveChar & ")" Then
                driveL = Burner.deviceDriveChar
                Found = True
            End If
        Next A
        
        If Found = False Then Err.Raise -1, "Finalize", "Could not find burner drive."
        
        
        ' Burn already made ISO
        If Project.createISO Then
            cmdLine = "/mode isowrite"
            cmdLine = cmdLine & " /src " & vbQuote & Project.destinationDir & "\dvd.iso" & vbQuote
            cmdLine = cmdLine & " /log " & vbQuote & Project.destinationDir & "\imgburn_isowrite.txt" & vbQuote
            
        ' Compile files straight to device
        Else
            cmdLine = "/mode build /buildmode device /noimagedetails /rootfolder yes /filesystem ""iso9660 + udf"" /udfrevision ""1.02"" /recursesubdirs yes"
            cmdLine = cmdLine & " /src " & vbQuote & Project.destinationDir & "\dvd" & vbQuote
            cmdLine = cmdLine & " /volumelabel " & vbQuote & Project.discLabel & vbQuote
            cmdLine = cmdLine & " /log " & vbQuote & Project.destinationDir & "\imgburn_write.txt" & vbQuote
        End If
        If Project.eraseRW Then cmdLine = cmdLine & " /erase"
        If Project.deleteISO Then cmdLine = cmdLine & " /deleteimage yes"
        If Project.verifyDisc Then cmdLine = cmdLine & " /verify yes"
        If Project.ejectTray Then cmdLine = cmdLine & " /eject yes"
        cmdLine = cmdLine & " /dest " & driveL & ":"
        cmdLine = cmdLine & " /close /start /nosavesettings"
        cmdLine = cmdLine & " /waitformedia /ignorelockvolume"
        cmdLine = cmdLine & " /portable"
        cmdLine = cmdLine & " /speed " & Project.burnSpeed
    
    
        ' Execute and hope for the best
        Log.Add cmdLine, 1
        Execute vbQuote & APP_PATH & "imgburn\imgburn.exe" & vbQuote, cmdLine, EP_Normal, True
    
    End If
    
    frmProgress.resetStatus

End Sub


Private Sub encodeAudio()

    Dim A As Long
    Dim B As Long
    Dim C As Long
    Dim myTitle As clsTitle
    Dim myTrack As clsAudioTrack
    Dim myAudio As clsAudio
    Dim Info As Dictionary
    Dim Info0 As Dictionary
    Dim channelCount As Long
    
    Dim sourceFile As String
    Dim destFile As String
    Dim logFile As String
    Dim cmdLine As String
    
    Dim AVI As clsAVIDelay
    Dim inFile As String
    Dim outFile As String
    Dim manualExtract As Boolean
    
    
    For A = 0 To Project.Titles.Count - 1
        Set myTitle = Project.Titles.Item(A)
        
        For B = 0 To myTitle.audioTracks.Count - 1
            Set myTrack = myTitle.audioTracks.Item(B)
            Set Info0 = myTrack.Sources.Item(0).streamInfo
            
            For C = 0 To myTrack.Sources.Count - 1
                Set myAudio = myTrack.Sources.Item(C)
                Set Info = myAudio.streamInfo

                logFile = Project.destinationDir & "\ffmpeg_audio_title" & A & "_track" & B & "_source" & C & ".txt"
                
                ' Init progress bar
                frmProgress.prgBar.Value = 0
                frmProgress.prgBar.Max = Info("Duration") * 10
                frmProgress.showBar True
                frmProgress.setStatus "Title " & A + 1 & " of " & Project.Titles.Count & ", track " & B + 1 & " of " & myTitle.audioTracks.Count & ", source " & C + 1 & " of " & myTrack.Sources.Count
                frmProgress.setSubStatus "Encoding..."
                
                Log.Add "Title " & A & ", track " & B & ", source " & C & " of " & myTrack.Sources.Count - 1, 1
                Log.Add "Source     : " & myAudio.Source.fileName, 2
                Log.Add "Properties : " & Info("Channels") & " channels, " & Info("sampleRate") & " Hz, " & Info("Compression") & ", " & Info("Duration") & " seconds", 2
                
                
                sourceFile = vbQuote & myAudio.Source.fileName & vbQuote
                destFile = audioFileName(A, B, C)
                
                cmdLine = ""
                cmdLine = cmdLine & " -i " & sourceFile
                
                
                ' Title's video source for syncing
                cmdLine = cmdLine & " -i " & vbQuote & videoFileName(A, 0) & vbQuote

                ' Stream mapping
                cmdLine = cmdLine & " -map 0:" & myAudio.streamInfo("Index") & ":1:0"

                ' Recompression needed?
                If Info("Compression") = "ac3" And Info("sampleRate") = 48000 And Info("Channels") = myTrack.targetChannels And Project.volumeMod = 100 Then
                    cmdLine = cmdLine & " -acodec copy"
                Else
                    cmdLine = cmdLine & " -acodec ac3 -ab " & myTrack.targetBitrate & "k -ac " & myTrack.targetChannels & " -ar 48000"
                    Log.Add "Recompressing to " & myTrack.targetBitrate & " Kbit\s with " & myTrack.targetChannels & " channel(s)", 2
                End If
                
                ' Volume modification
                If Project.volumeMod <> 100 Then cmdLine = cmdLine & " -vol " & CInt(256 * (Project.volumeMod / 100))
                                
                ' Number of concurrent threads
                cmdLine = cmdLine & " -threads " & Project.threadCount
                
                ' Audio sync method
                ' Disabled, because it does not seem to produce any sound at all
                cmdLine = cmdLine & " -async 0"
                
                ' Finish with shortest input
                cmdLine = cmdLine & " -shortest"
                
                ' Destination file
                cmdLine = cmdLine & " " & vbQuote & destFile & vbQuote
                
                
                ' Execute
                Log.Add cmdLine, 2
                executeToFile APP_PATH & "bin\ffmpeg.exe", cmdLine, logFile, SM_EncodeAudio, Project.encodePriority, ""
                
                ' Success?
                If FS.FileExists(destFile) = False Then
                    Err.Raise -1, "EncodeAudio", "Audio was not encoded."
                ElseIf FS.GetFile(destFile).Size = 0 Then
                    Err.Raise -1, "EncodeAudio", "Audio encoding failed."
                End If
                
                ' Remove input file if it was manually extracted first
                If manualExtract Then killIfExists outFile

                ' Fix delay
                If Info("Delay") <> 0 And myTrack.ignoreDelay = 0 Then

                    frmProgress.setSubStatus "Fixing delay..."
                    frmProgress.showBar False
                    frmProgress.lblPrct.Caption = ""

                    Log.Add "Fixing delay of " & Info("Delay") & " ms", 2

                    cmdLine = "-start " & Info("Delay") & " " & vbQuote & destFile & vbQuote
                    Log.Add cmdLine, 2
                    Execute vbQuote & APP_PATH & "delaycut\delaycut.exe" & vbQuote, cmdLine, WS_HIDE, True

                    outFile = FS.GetParentFolderName(destFile) & "\" & FS.GetBaseName(destFile) & "_fixed.ac3"

                    If Not FS.FileExists(outFile) Or FS.GetFile(outFile).Size = 0 Then
                        Err.Raise -1, "EncodeAudio", "Audio source was not corrected."
                    End If

                    killIfExists destFile
                    FS.MoveFile outFile, destFile

                End If

                ' Cancelled?
                If cancelError Then Exit Sub
            
            Next C
        
        Next B
        
    Next A
    
    
    frmProgress.resetStatus

End Sub


Private Sub encodeVideo()

    Dim A As Long, B As Long
    Dim myTitle As clsTitle
    Dim myVideo As clsVideo
    Dim Info As Dictionary, Encode As Dictionary
    
    Dim cmdLine As String
    Dim outFile As String, inFile As String, logFile As String
    Dim Status As String
    Dim targetBitrate As Long, maxBitRate As Long, minBitRate As Long
    
    Dim vWidth As Long, vHeight As Long, Resizer As Double, widthScale As Double
    Dim realSrcWidth As Long, realSrcHeight As Long, realDestWidth As Long, realDestHeight As Long
    Dim padTop As Long, padBottom As Long, padLeft As Long, padRight As Long
    Dim cropTop As Long, cropBottom As Long, cropLeft As Long, cropRight As Long

    
    For A = 0 To Project.Titles.Count - 1
        Set myTitle = Project.Titles.Item(A)
                
        For B = 0 To myTitle.Videos.Count - 1
            Set myVideo = myTitle.Videos.Item(B)
            Set Info = myVideo.streamInfo
            Set Encode = myTitle.encodeInfo
            targetBitrate = myTitle.videoBitRate(Sizes("avgBitRate"))

            Log.Add "Title " & A & ", source " & B, 1
            Log.Add "Source         : " & myVideo.Source.fileName, 2
            Log.Add "Properties     : " & Info("Width") & "x" & Info("Height") & ", " & visualAspectRatio(myVideo.streamInfo("sourceAR")) & ", " & Info("FPS") & " FPS, " & Info("Compression") & ", " & Info("Duration") & " seconds", 2
            Log.Add "Target BitRate : " & targetBitrate & " Kbit\s", 2

            logFile = Project.destinationDir & "\ffmpeg_video_title" & A & "_source" & B & ".txt"
        
            PSNR = 0
            padTop = 0
            padBottom = 0
            padLeft = 0
            padRight = 0
            cropTop = 0
            cropLeft = 0
            cropBottom = 0
            cropRight = 0
            cmdLine = vbNullString
        
            ' Status
            Status = "Title " & A + 1 & " of " & Project.Titles.Count & ", video " & B + 1 & " of " & myTitle.Videos.Count
            frmProgress.setStatus Status
            frmProgress.prgBar.Value = 0
            frmProgress.showBar True
        
        
            ' If Matroska, extract video stream first
            If myVideo.Source.isMatroska Then
                
                ' Status
                frmProgress.setSubStatus "Extracting video stream"
                frmProgress.prgBar.Max = 100
            
                If Not extractMKVStream(myVideo.Source.fileName, myVideo.streamIndex, Project.destinationDir & "\mkvtemp") Then
                    Err.Raise -1, "encodeVideo", "Could not extract stream from Matroska file."
                End If
                
                inFile = Project.destinationDir & "\mkvtemp"
                
                ' Force input framerate
                cmdLine = cmdLine & " -r " & Trim$(Str$(Info("FPS")))
                
            Else
                inFile = myVideo.Source.fileName
            
            End If
            
            If cancelError Then Exit Sub
        
        
            ' Status
            frmProgress.prgBar.Max = Info("Duration") * 10
            frmProgress.setSubStatus "Encoding..."
            
            ' Input file
            outFile = videoFileName(A, B)
            cmdLine = cmdLine & " -i " & vbQuote & inFile & vbQuote
            
            ' Copy video if requested
            If Project.MPEG2Copy = 1 And Info("Compression") = "mpeg2video" And Info("Width") = Encode("Width") And Info("Height") = Encode("Height") Then
                Log.Add "Copying as requested..."
                cmdLine = cmdLine & " -vcodec copy"
            Else
                Log.Add "Re-encoding", 2
                cmdLine = cmdLine & " -vcodec mpeg2video"
            End If
            
            
            ' Calculate resizing, cropping and padding
            realDestHeight = Encode("Height")
            realDestWidth = realDestHeight * getAspect(myTitle.targetAspect)
            widthScale = Encode("Width") / realDestWidth
            
            realSrcWidth = Info("Width")
            realSrcHeight = Info("Height")
            If myVideo.PAR <> 1 Then realSrcWidth = realSrcHeight * myVideo.PAR
            
            Resizer = getResizeValue(realSrcWidth, realSrcHeight, realDestWidth, realDestHeight)
            vWidth = realSrcWidth * Resizer
            vHeight = realSrcHeight * Resizer
            
            vWidth = vWidth * widthScale

            ' Add overscan borders
            ' First calculation should indeed use vHeight * Aspect
            If Project.overscanBorders Then
                vWidth = vWidth - (((Project.overscanSize / 100) * vHeight) * getAspect(myTitle.targetAspect))
                vHeight = vHeight - ((Project.overscanSize / 100) * vHeight)
            End If

            ' Ensure size is a multiple of 8 (so that padding and cropping will be multiple of 4)
            vWidth = vWidth - (vWidth Mod 8)
            vHeight = vHeight - (vHeight Mod 8)
            
            ' Set padding and\or cropping
            If vWidth < Encode("Width") Then
                padLeft = (Encode("Width") - vWidth) / 2
                padRight = padLeft
            ElseIf vWidth > Encode("Width") Then
                cropLeft = (vWidth - Encode("Width")) / 2
                cropRight = cropLeft
            End If
            If vHeight < Encode("Height") Then
                padTop = (Encode("Height") - vHeight) / 2
                padBottom = padTop
            ElseIf vHeight > Encode("Height") Then
                cropTop = (vHeight - Encode("Height")) / 2
                cropBottom = cropTop
            End If
            
            ' Add the settings to the commandline
            cmdLine = cmdLine & " -s " & vWidth & "x" & vHeight
            If padTop > 0 Then cmdLine = cmdLine & " -padtop " & padTop
            If padBottom > 0 Then cmdLine = cmdLine & " -padbottom " & padBottom
            If padLeft > 0 Then cmdLine = cmdLine & " -padleft " & padLeft
            If padRight > 0 Then cmdLine = cmdLine & " -padright " & padRight
            If cropTop > 0 Then cmdLine = cmdLine & " -croptop " & cropTop
            If cropBottom > 0 Then cmdLine = cmdLine & " -cropbottom " & cropBottom
            If cropLeft > 0 Then cmdLine = cmdLine & " -cropleft " & cropLeft
            If cropRight > 0 Then cmdLine = cmdLine & " -cropright " & cropRight
            
            
            ' Target format options
            cmdLine = cmdLine & " -r " & Trim$(Str$(Encode("FPS")))
            cmdLine = cmdLine & " -g " & Encode("GOP")
            
            ' DVD options
            cmdLine = cmdLine & " -bufsize 1835008 -packetsize 2048 -muxrate 10080000"
            
            ' Aspect ratio
            If myTitle.targetAspect = VA_169 Then cmdLine = cmdLine & " -aspect 16:9" Else cmdLine = cmdLine & " -aspect 4:3"
            
            
            ' Maximum bitrate
            ' -50 as a dirty workaround for encodes that end up being slightly too big
            targetBitrate = targetBitrate - 50
            cmdLine = cmdLine & " -minrate " & targetBitrate & "k -maxrate " & targetBitrate & "k -b " & targetBitrate & "k"
            
            ' Encoding profiles
            ' Tweaks FFmpeg for quality\speed
            cmdLine = cmdLine & " -preme 1 -precmp 2 -subcmp 8"
            If Project.encodeProfile = VE_Fastest Then
                cmdLine = cmdLine & " -mbd 1 -sws_flags fast_bilinear"
            ElseIf Project.encodeProfile = VE_Fast Then
                cmdLine = cmdLine & " -cmp 1 -mbcmp 8 -sws_flags bilinear+accurate_rnd"
            ElseIf Project.encodeProfile = VE_Normal Then
                cmdLine = cmdLine & " -mbcmp 8 -cmp 1 -sws_flags lanczos+accurate_rnd -mbd 2"
            ElseIf Project.encodeProfile = VE_Best Then
                cmdLine = cmdLine & " -mbcmp 8 -cmp 1 -sws_flags spline+accurate_rnd -trellis 1 -mbd 2"
            End If
            
            ' DC precision
            cmdLine = cmdLine & " -sc_threshold -3000 -dc " & Project.dcPrecision
            
            ' Interlacing
            If myVideo.Interlaced And Project.Deinterlace = 0 Then
                cmdLine = cmdLine & " -top -1 -flags +ilme+ildct"
            ElseIf myVideo.Interlaced And Project.Deinterlace = 1 Then
                cmdLine = cmdLine & " -deinterlace"
            End If
            
            ' Calculate PSNR
            If Project.PSNR = 1 Then cmdLine = cmdLine & " -psnr"
            
            ' Do not process audio
            cmdLine = cmdLine & " -an"
            
            ' Number of concurrent threads
            cmdLine = cmdLine & " -threads " & Project.threadCount
            
            ' Generate presentation timestamps
            cmdLine = cmdLine & " -fflags +genpts"
            
            ' Raw mpeg2video output file format
            cmdLine = cmdLine & " -f mpeg2video"
            
            ' Finish with shortest input
            cmdLine = cmdLine & " -shortest"
            
            ' Copy timestamps
            If myTitle.copyTS = 1 Then cmdLine = cmdLine & " -copyts"
            
            ' Stream mapping
            If Not myVideo.Source.isMatroska Then
                cmdLine = cmdLine & " -map 0:" & Info("Index")
            End If
                            
            ' Output file
            cmdLine = cmdLine & " " & vbQuote & outFile & vbQuote
            
            
            ' Execute
            Log.Add cmdLine, 2
            executeToFile APP_PATH & "bin\ffmpeg.exe", cmdLine, logFile, SM_EncodeVideo, Project.encodePriority, ""
            
            ' Successful?
            If FS.FileExists(outFile) = False Then
                Err.Raise -1, "encodeVideo", "File was not encoded."
            ElseIf FS.GetFile(outFile).Size = 0 Then
                Err.Raise -1, "encodeVideo", "Encoding error."
            End If
            
            ' Clean up matroska extract
            If myVideo.Source.isMatroska Then
                FS.DeleteFile Project.destinationDir & "\mkvtemp"
            End If
            
            ' Cancelled?
            If cancelError Then Exit Sub
            
            
            ' Output average PSNR
            If Project.PSNR = 1 Then
                Log.Add "Average PSNR: " & PSNR & " dB", 2
            End If
               
            ' Apply 2:3 pulldown?
            If Encode("Pulldown") = 1 Then
                Log.Add "Applying 2:3 pulldown", 2
                frmProgress.setSubStatus "Applying pulldown..."
                frmProgress.lblPrct.Caption = ""
                
                frmProgress.showBar True
                frmProgress.prgBar.Value = 0
                frmProgress.prgBar.Max = 100
                titleFrameNum = Info("FPS") * Info("Duration")
            
                
                inFile = outFile
                outFile = Project.destinationDir & "\pulldown.m2v"
                cmdLine = vbQuote & inFile & vbQuote & " " & vbQuote & outFile & vbQuote
                
                Log.Add cmdLine, 2
                executeToFile APP_PATH & "bin\pulldown.exe", cmdLine, Project.destinationDir & "\pulldown_title" & A & ".txt", SM_Pulldown, Project.encodePriority, ""
                
                ' Did pulldown succeed?
                If FS.FileExists(outFile) = False Or FS.GetFile(outFile).Size = 0 Then
                    Err.Raise -1, "encodeVideo", "Pulldown was not performed."
                End If
                
                ' Clean up
                FS.DeleteFile inFile
                FS.GetFile(outFile).Move inFile
                
            End If
            
        Next B
    
    Next A
    
    
    Set myTitle = Nothing
    Set myVideo = Nothing
    
    frmProgress.resetStatus

End Sub


Private Sub muxSubs()

    Dim A As Long, B As Long, C As Long
    Dim Stream As clsTextFile
    Dim myTitle As clsTitle
    Dim mySub As clsSubtitle
    Dim Encode As Dictionary
    
    Dim myFile As clsSubFile
    Dim myblock As clsSubBlock
    Dim baseFileName As String
    
    Dim cmdLine As String
    Dim xmlFile As String
    Dim inputFile As String
    Dim outputFile As String
    Dim logFile As String
    
    
    For A = 0 To Project.Titles.Count - 1
        Set myTitle = Project.Titles.Item(A)

        If myTitle.Subtitles.Count > 0 Then
            
            For B = 0 To myTitle.Subtitles.Count - 1
            
                Log.Add "Title " & A & ", subtitle " & B, 1
                frmProgress.setStatus "Title " & A + 1 & " of " & Project.Titles.Count & ", subtitle " & B + 1 & " of " & myTitle.Subtitles.Count
        
                ' Open new spumux XML file
                xmlFile = Project.destinationDir & "\spumux_title" & A & "_sub" & B & ".xml"
                
                Set Stream = New clsTextFile
                Stream.fileCreate xmlFile, encodeASCIIorUTF8, CODEPAGE_LATIN1
                Stream.writeText "<?xml version=""1.0"" ?>", True
                Stream.writeText "<subpictures>", True
        
                Set mySub = myTitle.Subtitles.Item(B)
                
                Log.Add "Source: " & mySub.fileName, 2
                Log.Add "Generating images...", 2
                frmProgress.setSubStatus "Generating images..."
                frmProgress.showBar True

                ' Load subtitle file
                Set myFile = New clsSubFile
                myFile.openFrom mySub.fileName, mySub.codePage
                myFile.FPS = mySub.FPS

                ' Convert framerate and frame start and ends if needed
                If mySub.frameBased = 1 Then
                    Set Encode = myTitle.encodeInfo
                    
                    Log.Add "Converting framerate from " & mySub.FPS & " to " & Encode("FPS") & " FPS", 2
                    If mySub.FPS <> Encode("FPS") Then myFile.convertToFPS Encode("FPS")
                    myFile.convertFramesToTime
                End If

                ' Fix up where possible
                Log.Add "Fixed " & myFile.fixOverlaps & " overlaps", 2
                Log.Add "Fixed " & myFile.fixShorts & " short blocks", 2
                
                
                ' Generate subtitle images
                baseFileName = Project.destinationDir & "\" & "spu" & A & "_" & B & "_"
                modRenderSub.renderSubFile myFile, myTitle, mySub, baseFileName, frmProgress.prgBar
                
                ' Convert each to PNG
                Log.Add "Converting images...", 2
                frmProgress.setSubStatus "Converting images..."
                frmProgress.prgBar.Max = myFile.blockCount
                For C = 0 To myFile.blockCount - 1
                    If cancelError Then Exit For
                    
                    convertToPNG baseFileName & C & ".bmp", baseFileName & C & ".png", mySub.transBack - 1, False
                    frmProgress.prgBar.Value = C
                Next C
                frmProgress.showBar False
                
                If cancelError Then Exit For
                

                ' Add stream to a spumux XML file
                Log.Add "Generating XML...", 2
                Stream.writeText "<stream>", True
                
                For C = 0 To myFile.blockCount - 1
                    Set myblock = myFile.getBlock(C)
                    
                    Stream.writeText "<spu "
                    Stream.writeText "start=" & vbQuote & displayTime(myblock.startTime) & vbQuote & " "
                    Stream.writeText "end=" & vbQuote & displayTime(myblock.endTime) & vbQuote & " "
                    Stream.writeText "image=" & vbQuote & "spu" & A & "_" & B & "_" & C & ".png" & vbQuote & " "
                    Stream.writeText "xoffset=""" & myblock.oX & """ "
                    Stream.writeText "yoffset=""" & myblock.oY & """ "
                    Stream.writeText "/>", True
                
                Next C
                
                Stream.writeText "</stream>", True
                Stream.writeText "</subpictures>", True
                
                Set Stream = Nothing
        
                
                frmProgress.setSubStatus "Inserting..."
                
                ' Generate filenames and mux
                inputFile = "title" & A & ".mpg"
                outputFile = "title" & A & ".temp"
                logFile = "spumux_title" & A & "_sub" & B & ".txt"
                
                ' Write the command line to a batch file, else spumux chokes on the
                ' redirection commands
                Set Stream = New clsTextFile
                Stream.fileCreate Project.destinationDir & "\spumux.bat", encodeASCIIorUTF8, CODEPAGE_LATIN1
                    
                Stream.writeText "cd /D " & vbQuote & Project.destinationDir & vbQuote, True
                cmdLine = vbQuote & APP_PATH & "bin\spumux.exe" & vbQuote
                cmdLine = cmdLine & " -s " & B
                cmdLine = cmdLine & " -m dvd " & vbQuote & xmlFile & vbQuote
                cmdLine = cmdLine & " < " & vbQuote & inputFile & vbQuote
                cmdLine = cmdLine & " 1> " & vbQuote & outputFile & vbQuote
                cmdLine = cmdLine & " 2> " & vbQuote & logFile & vbQuote
                Log.Add cmdLine, 2
                Stream.writeText cmdLine, True
                
                Set Stream = Nothing
        
        
                ' Execute and clean up batch file
                Execute Project.destinationDir & "\spumux.bat", "", WS_HIDE, True, True
                
                ' Output exists?
                inputFile = Project.destinationDir & "\" & inputFile
                outputFile = Project.destinationDir & "\" & outputFile
                If FS.FileExists(outputFile) = False Or FS.GetFile(outputFile).Size < 16 Then
                    Err.Raise -1, "MuxSubs", "Subtitles were not added."
                End If
                
                ' Delete old, rename new
                If FS.FileExists(outputFile) Then
                    killIfExists inputFile
                    FS.MoveFile outputFile, inputFile
                End If
                
                ' Clean up subtitle images
                If Project.keepFiles = 0 Then FS.DeleteFile baseFileName & "*.png*"
                killIfExists Project.destinationDir & "\spumux.bat"
                
            Next B
            
            If cancelError Then Exit For
            
        End If
        
    Next A
    
    
    Set myTitle = Nothing
    Set mySub = Nothing
    Set myFile = Nothing

    frmProgress.resetStatus

End Sub


Private Sub authorDVD()

    Dim A As Long
    Dim Stream As clsTextFile
    Dim cmdLine As String
    
    Dim myTitle As clsTitle
    Dim myTrack As clsAudioTrack
    Dim cTemplate As clsMenuTemplate
    
    
    frmProgress.setStatus "Authoring DVD..."
    
    If Project.menuTemplateName <> STR_DISABLED_MENU Then
        frmProgress.setStatus "Generating menus..."
        
        ' Generate menu images if necessary
        Log.Add "Opening menu template", 1
        Set cTemplate = New clsMenuTemplate
        cTemplate.openFrom APP_PATH & "templates\" & Project.menuTemplateName & "\template.cfg"
        
        ' Generate and render menus
        Log.Add "Generating menu images", 1
        Project.generateMenus cTemplate.Templates
        Project.renderMenus cTemplate
        Project.rescaleMenus
        
        Log.Add "Generating menu files", 1
        If Not Project.generateMenuFiles Then Err.Raise -1, "AuthorDVD", "Could not generate menu files."
    End If
    
    
    frmProgress.showBar True
    frmProgress.prgBar.Max = 100
    
    ' Generate DVDAuthor XML file
    Log.Add "Generating XML file", 1
    Set Stream = New clsTextFile
    Stream.fileCreate Project.destinationDir & "\dvdauthor.xml", encodeASCIIorUTF8, 65001
    Stream.writeText "<?xml version=""1.0"" encoding=""UTF-8""?>", True
    Stream.writeText "<dvdauthor dest=""" & pathEntities(Project.destinationDir) & "\dvd"">", True
    
    Log.Add "Generating author information", 1
    Stream.writeText Project.generateAuthorData
    
    For A = 0 To Project.Titles.Count - 1
        Set myTitle = Project.Titles.Item(A)
        Stream.writeText myTitle.generateAuthorData(A)
    Next A
    Stream.writeText "</dvdauthor>", True
    
    Set Stream = Nothing
    
    
    ' Run DVDAuthor
    cmdLine = "-x " & vbQuote & Project.destinationDir & "\dvdauthor.xml" & vbQuote
    Log.Add cmdLine, 1
    executeToFile APP_PATH & "bin\dvdauthor.exe", cmdLine, Project.destinationDir & "\dvdauthor.txt", SM_DVDAuthor, Project.encodePriority, ""
    
    
    ' Successful?
    If Not FS.FileExists(Project.destinationDir & "\dvd\video_ts\video_ts.ifo") Then
        Err.Raise -1, "AuthorDVD", "DVD was not authored correctly."
    End If
    
    ' Cleanup
    If Project.keepFiles = 0 Then
        Log.Add "Cleaning up...", 1
        For A = 0 To Project.Titles.Count - 1
            killIfExists Project.destinationDir & "\title" & A & ".mpg"
        Next A
        
        If Project.menuTemplateName <> STR_DISABLED_MENU Then FS.DeleteFile Project.destinationDir & "\*menu*d.mpg"
    Else
        Log.Add "Keeping encoded titles", 1
    End If
    
    
    Set myTitle = Nothing
    frmProgress.resetStatus

End Sub


' Concatenate audio files
Private Sub concatAudio()

    Dim A As Long, B As Long, C As Long
    Dim tempBatch As String
    Dim Stream As clsTextFile
    Dim myTitle As clsTitle
    Dim myTrack As clsAudioTrack
    

    Log.Add "Concatenating audio...", 2
    frmProgress.setStatus "Joining audio"
    
    For A = 0 To Project.Titles.Count - 1
    
        Set myTitle = Project.Titles.Item(A)
        tempBatch = TEMP_PATH & "temp.bat"
        
        For B = 0 To myTitle.audioTracks.Count - 1
        
            Log.Add "Title " & A & ", track " & B, 3
            frmProgress.setSubStatus "Title " & A + 1 & " of " & Project.Titles.Count & ", track " & B & " of " & myTitle.audioTracks.Count & "..."
            
            Set myTrack = myTitle.audioTracks.Item(B)
            
            If myTrack.Sources.Count > 1 Then
                For C = 1 To myTrack.Sources.Count - 1
                    
                    Set Stream = New clsTextFile
                    Stream.fileCreate tempBatch, encodeASCIIorUTF8, CODEPAGE_LATIN1
                    Stream.writeText vbQuote & APP_PATH & "bin\cat.exe" & vbQuote & " " & vbQuote & audioFileName(A, B, C) & vbQuote & " >> " & vbQuote & audioFileName(A, B, 0) & vbQuote
                    Set Stream = Nothing
                    
                    Execute tempBatch, "", WS_HIDE, True, True
                                        
                    If Project.keepFiles = 0 Then killIfExists audioFileName(A, B, C)
                Next C
                
                killIfExists tempBatch
            End If
            
            FS.MoveFile audioFileName(A, B, 0), audioFileName(A, B, -1)
            
        Next B
                
    Next A

End Sub


' Concatenate video files
Private Sub concatVideo()

    Dim A As Long, B As Long, C As Long
    Dim tempBatch As String
    Dim Stream As clsTextFile
    Dim myTitle As clsTitle
    Dim myVideo As clsVideo
    

    Log.Add "Concatenating video...", 2
    frmProgress.setStatus "Joining video"
    
    For A = 0 To Project.Titles.Count - 1
    
        Set myTitle = Project.Titles.Item(A)
        If myTitle.Videos.Count > 1 Then
        
            tempBatch = TEMP_PATH & "temp.bat"
        
            For B = 1 To myTitle.Videos.Count - 1
                Set myVideo = myTitle.Videos.Item(B)
                
                Log.Add "Title " & A & ", video " & B, 3
                frmProgress.setSubStatus "Title " & A + 1 & " of " & Project.Titles.Count & ", video " & B & " of " & myTitle.Videos.Count & "..."
                
                Set Stream = New clsTextFile
                Stream.fileCreate tempBatch, encodeASCIIorUTF8, CODEPAGE_LATIN1
                Stream.writeText vbQuote & APP_PATH & "bin\cat.exe" & vbQuote & " " & vbQuote & videoFileName(A, B) & vbQuote & " >> " & vbQuote & videoFileName(A, 0) & vbQuote
                Set Stream = Nothing
                
                Execute tempBatch, "", WS_HIDE, True, True

                If Project.keepFiles = 0 Then killIfExists videoFileName(A, B)
            Next B
            
            killIfExists tempBatch
            
        End If
        
    Next A
    
    frmProgress.resetStatus

End Sub


' Multiplex streams using MPLEX
Private Sub muxStreams()

    Dim A As Long
    Dim B As Long
    
    Dim totalFiles As String
    Dim videoFile As String
    Dim sourceVideo As String
    Dim cmdLine As String
    Dim logFile As String
    
    Dim myFolder As Folder
    Dim myFile As File
    Dim myTitle As clsTitle
    Dim Encode As Dictionary
    

    Log.Add "Multiplexing audio and video", 2
    frmProgress.showBar False
    
    For A = 0 To Project.Titles.Count - 1
    
        Set myTitle = Project.Titles.Item(A)
        Set Encode = myTitle.encodeInfo
        
        Log.Add "Title " & A, 3
        frmProgress.setStatus "Combining title " & A + 1 & " of " & Project.Titles.Count & "..."
        
        ' Concatenate filenames
        sourceVideo = videoFileName(A, 0)
        videoFile = Project.destinationDir & "\title" & A & ".mpg"
        totalFiles = vbQuote & sourceVideo & vbQuote & " "
        
        For B = 0 To myTitle.audioTracks.Count - 1
            totalFiles = totalFiles & " " & vbQuote & audioFileName(A, B, -1) & vbQuote
        Next B
        
        
        ' Construct command line
        cmdLine = "-f 8 --vbr -o " & vbQuote & videoFile & vbQuote & " " & totalFiles

        ' Execute
        logFile = Project.destinationDir & "\mplex_" & "title" & A & ".txt"
        Log.Add cmdLine, 3
        executeToFile APP_PATH & "bin\mplex.exe", cmdLine, logFile, SM_Nothing, Project.encodePriority, TEMP_PATH
        If cancelError Then Exit For
        
        ' Did the muxed file get created?
        If FS.FileExists(videoFile) = False Then
            Err.Raise -1, "muxStreamsMPLEX", "File was not multiplexed."
        End If
        
        ' Clean up
        If Project.keepFiles = 0 Then
            Log.Add "Cleaning up...", 2
           
            ' Delete source files
            killIfExists sourceVideo
            For B = 0 To myTitle.audioTracks.Count - 1
                killIfExists audioFileName(A, B, -1)
            Next B
            
        Else
            Log.Add "Keeping encoded files", 2
        End If
        
    Next A
    
    
    frmProgress.resetStatus

End Sub


' Return required disc space for current project in Kilobytes
' All internal measurements are in Kilobytes
Public Function requiredSpace() As Long

    Dim A As Long, B As Long, C As Long
    Dim maxSize As Long
    
    Dim myTitle As clsTitle
    Dim myTrack As clsAudioTrack
    Dim myAudio As clsAudio
    Dim myVideo As clsVideo
    Dim mySub As clsSubtitle
    
    Dim Sizes As Dictionary
    Dim videoBitRate As Long
    Dim titleSize As Long
    Dim sourceSize As Long
    
    
    Set Sizes = Project.calculateSizes
    
    ' Video streams
    videoBitRate = Sizes("avgBitRate")
    
    ' Video encoding
    For A = 0 To Project.Titles.Count - 1
        Set myTitle = Project.Titles.Item(A)
        
        For B = 0 To myTitle.Videos.Count - 1
            Set myVideo = myTitle.Videos.Item(B)
            
            ' Matroska extract
            sourceSize = FS.GetFile(myVideo.Source.fileName).Size / 1024
            If myVideo.Source.isMatroska Then
                maxSize = maxSize + sourceSize
            Else
                sourceSize = 0
            End If
            
            maxSize = maxSize + myVideo.encodedSize(videoBitRate)
            If myTitle.encodeInfo("Pulldown") = 1 Then
                maxSize = maxSize + myVideo.encodedSize(videoBitRate)
                If maxSize > requiredSpace Then requiredSpace = maxSize
                maxSize = maxSize - myVideo.encodedSize(videoBitRate)
            End If
            
            If sourceSize <> 0 Then maxSize = maxSize - sourceSize
            
        Next B
    Next A
        
    
    ' Concatenate video
    For A = 0 To Project.Titles.Count - 1
        
        ' Skip 0 because all will be concatenated onto that one
        For B = 1 To myTitle.Videos.Count - 1
            Set myVideo = myTitle.Videos.Item(B)
            
            maxSize = maxSize + myVideo.encodedSize(videoBitRate)
            If maxSize > requiredSpace Then requiredSpace = maxSize
            If Project.keepFiles = 0 Then maxSize = maxSize - myVideo.encodedSize(videoBitRate)
        Next B
        
    Next A


    ' Audio encoding
    For A = 0 To Project.Titles.Count - 1
        Set myTitle = Project.Titles.Item(A)
        
        ' Encode stage
        For B = 0 To myTitle.audioTracks.Count - 1
            Set myTrack = myTitle.audioTracks.Item(B)
        
            maxSize = maxSize + myTrack.encodedSize
            
            ' Audio delay
            For C = 0 To myTrack.Sources.Count - 1
                Set myAudio = myTrack.Sources.Item(C)
                If myAudio.streamInfo("Delay") Then
                    maxSize = maxSize + myAudio.encodedSize(myTrack.targetBitrate)
                    If maxSize > requiredSpace Then requiredSpace = maxSize
                    maxSize = maxSize - myAudio.encodedSize(myTrack.targetBitrate)
                End If
            Next C
            
        Next B
        
    Next A
    
    
    ' Stream muxing stage - Add 10% muxing overhead
    maxSize = 1.1 * maxSize
    If maxSize > requiredSpace Then requiredSpace = maxSize
    
    
    ' Concatenate audio
    For A = 0 To Project.Titles.Count - 1
        
        ' Skip 0 because all will be concatenated onto that one
        For B = 1 To myTitle.audioTracks.Count - 1
            Set myTrack = myTitle.audioTracks.Item(B)
            
            maxSize = maxSize + myTrack.encodedSize
            If maxSize > requiredSpace Then requiredSpace = maxSize
            If Project.keepFiles = 0 Then maxSize = maxSize - myTrack.encodedSize
        Next B
        
    Next A
    
    
    ' Multiplex
    For A = 0 To Project.Titles.Count - 1
        Set myTitle = Project.Titles.Item(A)
        
        titleSize = 0
        For B = 1 To myTitle.audioTracks.Count - 1
            titleSize = titleSize + myTitle.audioTracks.Item(B).encodedSize
        Next B
        For B = 1 To myTitle.Videos.Count - 1
            titleSize = titleSize + myTitle.Videos.Item(B).encodedSize(videoBitRate)
        Next B
        
        maxSize = maxSize + titleSize
        If maxSize > requiredSpace Then requiredSpace = maxSize
        If Project.keepFiles = 0 Then maxSize = maxSize - titleSize
        
    Next A
        
    
    ' Subtitle mux stage
    For A = 0 To Project.Titles.Count - 1
        Set myTitle = Project.Titles.Item(A)
        
        titleSize = 0
        For B = 1 To myTitle.audioTracks.Count - 1
            titleSize = titleSize + myTitle.audioTracks.Item(B).encodedSize
        Next B
        For B = 1 To myTitle.Videos.Count - 1
            titleSize = titleSize + myTitle.Videos.Item(B).encodedSize(videoBitRate)
        Next B
        
        maxSize = maxSize + titleSize + (myTitle.Duration * myTitle.subBitRate)
        If maxSize > requiredSpace Then requiredSpace = maxSize
        maxSize = maxSize - titleSize
    Next A
    
    If maxSize > requiredSpace Then requiredSpace = maxSize
    
    
    ' Menus
    ' Require 512Kb per menu (estimated amount of title select menus)
    maxSize = maxSize + 512
    maxSize = maxSize + 512 * Project.Titles.Count / 4
    maxSize = maxSize + 512
    
    
    ' Author stage: twice MaxSize + 10% mux overhead
    ' This includes ISO image creation\burning
    maxSize = maxSize * 2 + (0.1 * maxSize)
    If maxSize > requiredSpace Then requiredSpace = maxSize

End Function


' Handler for Pulldown pipe input
Public Sub pulldownPipe(ByVal Data As String)

    Dim Offset As Long
    Dim crOffset As Long
    Dim frameNum As Long
    
    On Error GoTo endPullPipe
    
    
    Offset = InStrRev(Data, "Frame = ")
    If Offset > 0 Then
        Offset = Offset + 8
        crOffset = InStr(Offset, Data, vbCr)
        If crOffset > 0 Then
            frameNum = Val(Mid$(Data, Offset, crOffset - Offset))
            frameNum = (frameNum / titleFrameNum) * 100
            If frameNum >= 0 And frameNum <= 100 Then frmProgress.prgBar.Value = frameNum
        End If
    End If
    
    On Error GoTo 0
    Exit Sub
    
endPullPipe:

End Sub


' Handler for DVDAuthor pipe input
Public Sub dvdauthorPipe(ByVal Data As String)

    Dim Offset As Long, mbOffset As Long
    Dim mbSize As Long, titleNum As Long
    Dim Check As Boolean
    
    Static destFileLength As Variant
    
    On Error GoTo endDVDPipe
    
    
    ' Pick title to copy
    Offset = InStr(Data, "Picking VTS ")
    If Offset > 0 And Offset + 12 < Len(Data) Then
        titleNum = Val(Mid$(Data, Offset + 12, 2))
        destFileLength = FS.GetFile(Project.destinationDir & "\title" & titleNum - 1 & ".mpg").Size / 1024 / 1024
        frmProgress.setStatus "Authoring title " & titleNum & " of " & Project.Titles.Count & "..."
    End If
    
    ' Copying
    If InStr(Data, "STAT: VOBU ") <> 0 And destFileLength > 0 Then
        frmProgress.setSubStatus "Creating filesystem"
        Offset = InStrRev(Data, " at ")
        If Offset > 0 Then
            mbOffset = InStr(Offset + 4, Data, "MB")
            If mbOffset > 0 Then
                mbSize = Val(Mid$(Data, Offset + 4, mbOffset - (Offset + 4)))
                mbSize = (mbSize / destFileLength) * 100
                If mbSize >= 0 And mbSize <= 100 Then frmProgress.prgBar.Value = mbSize
            End If
        End If
        
    ' Fixing
    ElseIf InStr(Data, "STAT: fixing ") <> 0 Then
        frmProgress.setSubStatus "Fixing"
        
        Offset = InStrRev(Data, "%")
        If Offset > 0 Then
            mbSize = Val(Mid$(Data, Offset - 3, 3))
            If mbSize >= 0 And mbSize <= 100 Then frmProgress.prgBar.Value = mbSize
        End If
        
    End If
    
    On Error GoTo 0
    Exit Sub
    
endDVDPipe:
    
End Sub


' Handler for mkvextract pipe input
Public Sub mkvExtractPipe(ByVal Data As String)

    Dim Offset As Long
    Dim Percentage As Long
    
    
    If InStr(Data, ": ") And InStr(Data, "%") Then
        Offset = InStr(Data, ": ") + 2
        Percentage = Val(Mid$(Data, Offset))
        If Percentage >= 0 And Percentage < 101 Then frmProgress.prgBar.Value = Percentage
    End If
    
End Sub


' Handler for TCMPlex pipe input
Public Sub tcmplexPipe(ByVal Data As String)

    Dim Offset As Long
    Dim Percentage As Long
    
    On Error GoTo endTCMPipe
        
    
    If InStr(Data, "Scanning audio stream") <> 0 Then
        frmProgress.setSubStatus "Examining audio"
    ElseIf InStr(Data, "Scanning video stream") <> 0 Then
        frmProgress.setSubStatus "Examining video"
    ElseIf InStr(Data, "Multiplexing:") <> 0 Then
        frmProgress.setSubStatus "Combining"
    End If
    
    Offset = InStrRev(Data, "%")
    If Offset > 0 Then
        If Offset > 3 Then Percentage = Val(Mid$(Data, Offset - 3, 3))
        If Percentage >= 0 And Percentage <= 100 Then frmProgress.prgBar.Value = Percentage
    End If
    
    If InStr(Data, "ERROR:") <> 0 Then cancelError = True
    
    On Error GoTo 0
    Exit Sub
    
endTCMPipe:

End Sub


' Handler for FFmpeg encoding pipe input
Public Sub ffmpegPipe(ByVal Data As String)

    Dim Have As Long
    Dim Spatial As Long
    Dim timeOffs As Single
    Dim Proc As Long
    Dim Rover As Long
    Dim nFPS As Long
    
    On Error GoTo endFFPipe
    
    
    ' Progress percentage
    Have = InStr(Data, "time=")
    If Have > 0 Then

        Spatial = InStr(Have, Data, " ")
        If Spatial Then

            timeOffs = Val(Mid$(Data, Have + 5, Spatial - (Have + 5))) * 10
            If timeOffs > frmProgress.prgBar.Max Then timeOffs = frmProgress.prgBar.Max

            Proc = (frmProgress.prgBar.Value / frmProgress.prgBar.Max) * 100

            If timeOffs > 0 Then
                frmProgress.prgBar.Value = timeOffs
                frmProgress.Refresh
            End If

            frmProgress.lblPrct.Caption = Proc & "%"

        End If

    End If
    
    ' Encoding framerate
    Have = InStr(Data, "fps=")
    If Have > 0 Then
        
        If Have < Len(Data) - 4 Then nFPS = CLng(Mid$(Data, Have + 4, 3))
        If nFPS > 0 Then frmProgress.lblPrct.Caption = frmProgress.lblPrct.Caption & vbNewLine & CLng(nFPS) & " FPS"
        
    End If
    
    
    ' PSNR
    Rover = 0
    Do
        Rover = InStr(Rover + 1, Data, "LPSNR=")
        
        If Rover Then
            Spatial = InStr(Rover + 32, Data, " ")
            If Spatial Then PSNR = Val(Mid$(Data, Rover + 32, Spatial - Rover))
        End If
    Loop Until Not Rover
    
    On Error GoTo 0
    Exit Sub
    
endFFPipe:

End Sub


' Create a log file with relevant info the currently loaded project and encoding state
Private Sub createErrorLog(ByVal errorNum As Long, ByVal errorSource As String, ByVal errorDesc As String)
 
    Dim A As Long, B As Long
    Dim fileObj As clsBinaryFile
    Dim errLog As clsLog
    Dim mySource As clsSource
    Dim Info As Dictionary
    Dim myFile As File
    Dim myDrive As Drive
    Dim Stream As clsTextFile
    Dim myTitle As clsTitle
    
    Dim Data As String
    Dim Buffer As String * 1024
    Dim seekPos As Long
    
    
    Set errLog = New clsLog
    errLog.Start Project.destinationDir & "\errorlog.txt", False
    
    errLog.Add "[code]"
    errLog.Add "DVD Flick error log"
    errLog.Add "Version " & versionString
    errLog.Add windowsVersion
    
    ' General project info
    errLog.Add "*** Project info"
    errLog.Add "EncodeProf " & Project.encodeProfile
    errLog.Add "Threads " & Project.threadCount
    errLog.Add "OverScan " & Project.overscanBorders & ", " & Project.overscanSize
    errLog.Add "TargetRate " & Project.targetBitrate
    errLog.Add "TargetFormat " & Project.targetFormat
    errLog.Add "TargetSize " & Project.targetSize
    errLog.Add "CreateISO " & Project.createISO
    errLog.Add "CustRate " & Project.customBitrate
    errLog.Add "EnableBurn " & Project.enableBurning
    errLog.Add "EraseRW " & Project.eraseRW
    errLog.Add "HalfRes " & Project.halfRes
    errLog.Add "LoopPlayb " & Project.loopPlayback
    errLog.Add "AutoPlayMenu " & Project.menuAutoPlay
    errLog.Add "MenuTemplt " & Project.menuTemplateName
    errLog.Add "VolumeMod " & Project.volumeMod
    errLog.Add "WhenPlayed " & Project.whenPlayed
    
    ' Warnings given
    errLog.Add "*** Warnings given"
    errLog.Add Warnings
    
    ' Title settings
    errLog.Add "*** Titles"
    For A = 0 To Project.Titles.Count - 1
        Set myTitle = Project.Titles.Item(A)
        
        errLog.Add A & " - " & myTitle.Name
        
        errLog.Add "ChaptCount " & myTitle.chapterCount, 1
        errLog.Add "ChaptInt " & myTitle.chapterInterval, 1
        errLog.Add "ChaptSource " & myTitle.chapterOnSource, 1
        errLog.Add "TargetAR " & myTitle.targetAspect, 1
        errLog.Add "ThumbIndex " & myTitle.thumbTimeIndex, 1

    Next A
    
    ' Write all sources
    errLog.Add "*** Project sources"
    For A = 0 To Project.nSources - 1
        Set mySource = Project.getSourceIndex(A)
        
        errLog.Add mySource.fileName, 1
        
        ' Write each stream's info
        For B = 0 To mySource.streamCount - 1
            Set Info = mySource.streamInfo(B)
            errLog.Add "Stream " & B & " (" & Info("Type") & ")", 1
            
            errLog.Add "Dur " & Info("Duration"), 2
            errLog.Add "StartT " & Info("startTime"), 2
            errLog.Add "Compr " & Info("Compression"), 2
            errLog.Add "BitR " & Info("bitRate"), 2
                                    
            If Info("Type") = ST_Audio Then
                errLog.Add "Chnnls " & Info("Channels"), 2
                errLog.Add "SRate " & Info("sampleRate"), 2
                errLog.Add "Delay " & Info("Delay"), 2
            ElseIf Info("Type") = ST_Video Then
                errLog.Add Info("Width") & "x" & Info("Height") & " " & Info("FPS") & " FPS", 2
                errLog.Add "PAR " & Info("pixelAR"), 2
                errLog.Add "SAR " & Info("sourceAR"), 2
            End If
        Next B
    Next A
    
    
    ' Append dvdflick.log
    errLog.Add "*** dvdflick.log"
    Set Stream = New clsTextFile
    Stream.fileOpen Project.destinationDir & "\dvdflick.log", False
    Data = Stream.readAll
    errLog.Add Data
    Set Stream = Nothing
    
    
    ' Append last 448 bytes from each txt file
    For Each myFile In FS.GetFolder(Project.destinationDir).Files
        If Right$(LCase$(myFile.Name), 4) = ".txt" And myFile.Name <> "errorlog.txt" Then
        
            errLog.Add "*** " & myFile.Name
            
            Set fileObj = New clsBinaryFile
            fileObj.fileOpen Project.destinationDir & "\" & myFile.Name, False
            If fileObj.fileLength - 448 >= 0 Then
                fileObj.fileSeek fileObj.fileLength - 448
            Else
                fileObj.fileSeek 0
            End If
            Buffer = fileObj.readStringData(448)
            errLog.Add Buffer
            Set fileObj = Nothing
            
        End If
    Next myFile
    
    
    ' List files in project destination directry
    errLog.Add "*** Project destination folder"
    For Each myFile In FS.GetFolder(Project.destinationDir).Files
        errLog.Add myFile.Name & vbTab & myFile.Size, 1
    Next myFile
    
    
    ' List drive types and spaces on computer
    errLog.Add "*** Drives"
    For Each myDrive In FS.Drives
        If myDrive.IsReady = True Then
            errLog.Add myDrive.DriveLetter & " - " & myDrive.VolumeName & " - " & myDrive.DriveType, 1
            errLog.Add "FS " & myDrive.FileSystem, 1
            errLog.Add "AvailSpace " & CLng(myDrive.AvailableSpace / 1024 / 1024), 1
            errLog.Add "FreeSpace " & CLng(myDrive.FreeSpace / 1024 / 1024), 1
            errLog.Add "TotalSize " & CLng(myDrive.TotalSize / 1024 / 1024), 1
        Else
            errLog.Add myDrive.DriveLetter & " - not ready"
        End If
    Next myDrive
    
    
    errLog.Add "[/code]"
    Set errLog = Nothing

End Sub


' Extract a stream from a Matroska file, using mkvextract
Private Function extractMKVStream(ByVal fileName As String, ByVal streamIndex As Long, ByVal outFile As String)

    Dim cmdLine As String
    
    
    ' mkvextract tracks "test.mkv" 1:"output"
    cmdLine = cmdLine & "tracks tracks " & vbQuote & fileName & vbQuote
    cmdLine = cmdLine & " " & streamIndex + 1 & ":" & vbQuote & outFile & vbQuote

    modPipes.executeToFile APP_PATH & "mkvextract\mkvextract.exe", cmdLine, Project.destinationDir & "\mkvextract.txt", SM_MKVExtract, Project.encodePriority, ""
    
    If Not FS.FileExists(outFile) Then Exit Function
    
    extractMKVStream = True

End Function
