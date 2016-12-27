Attribute VB_Name = "modStrings"
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
'   File purpose: General utility string-related functions.
'
Option Explicit
Option Compare Binary
Option Base 0


' Return a path and filename with it's extension stripped off
Public Function stripExt(ByVal Data As String) As String

    If InStr(Data, ".") = 0 Then
        stripExt = Data
        Exit Function
    End If
    
    stripExt = Left$(Data, InStr(Data, ".") - 1)

End Function


' Return a byte array as hexadecimal string data
Public Function getHexString(ByRef Data() As Byte) As String

    Dim A As Long
    
    
    For A = LBound(Data) To UBound(Data)
        getHexString = getHexString & Hex$(Data(A)) & " "
    Next A
    
    getHexString = RTrim$(getHexString)

End Function


' Convert a xx:xx string to an aspect ratio value
Public Function getARFromString(ByVal Data As String) As Single

    Dim Offs As Long
    Dim lVal As Long, rVal As Long
    

    Offs = InStr(Data, ":")
    If Offs <= 0 Then Exit Function
    
    lVal = Val(Left$(Data, Offs - 1))
    rVal = Val(Mid$(Data, Offs + 1))
    If rVal <> 0 Then getARFromString = lVal / rVal

End Function


' Clean up an XML file by reindenting
' TODO: Does not yet work
Public Sub tidyXML(ByVal fileName As String)

    Dim A As Long
    Dim Stream As clsTextFile
    Dim inData As String, outData As String
    Dim Indent As Long
    Dim Char As String, lineData As String
    Dim inString As Boolean, inTag As Boolean
    
    
    Set Stream = New clsTextFile
    Stream.fileOpen fileName, False
    inData = Stream.readAll
    Set Stream = Nothing
    
    For A = 1 To Len(inData)
        Char = Mid$(inData, A, 1)
        
        If Char = vbCr Or Char = vbLf Or Char = vbTab Then
            Char = ""
        ElseIf Char = vbQuote And inTag Then
            If inString = True Then
                inString = False
            Else
                inString = True
            End If
        ElseIf Char = "<" Then
            Indent = Indent + 1
            inTag = True
        ElseIf inTag = True And inString = False And Char = "/" Then
            Indent = Indent - 1
        ElseIf Char = ">" Then
            outData = outData & String(Indent - 1, vbTab) & LTrim$(lineData) & ">" & vbNewLine
            
            lineData = ""
            Char = ""
            inTag = False
        End If
        
        lineData = lineData & Char
        
    Next A
    
    Debug.Print outData

End Sub


' Convert byte array of string contents
Public Sub stringToByteArray(ByVal Data As String, ByRef arrayData() As Byte)

    Dim A As Long


    For A = 1 To Len(Data)
        arrayData(A - 1) = AscB(Mid$(Data, A, 1))
    Next A

End Sub


' Remove trailing NULL characters from a string
Public Function stripNull(ByVal fixedString As String) As String

    Dim A As Long
    
    
    For A = Len(fixedString) To 1 Step -1
        If Mid$(fixedString, A, 1) <> vbNullChar Then
            stripNull = Left$(fixedString, A)
            Exit Function
        End If
    Next A
    
    stripNull = fixedString
    
End Function


' Parse a time in 00:00:00.0 format to seconds
Public Function parseTime(ByVal Data As String) As Single

    Dim Temp() As String
    
    
    parseTime = -1

    If Data = "N/A" Then Exit Function
    If InStr(Data, ".") = 0 Then Exit Function
    If Len(Data) < 10 Then Exit Function

    Temp = Split(Data, ":")
    If UBound(Temp) < 2 Then Exit Function
    If Len(Temp(0)) < 2 Then Exit Function
    If Len(Temp(1)) < 2 Then Exit Function
    If InStr(Temp(2), ".") = 0 Then Exit Function

    parseTime = 0
    parseTime = parseTime + (CLng(Temp(0)) * 60 * 60)
    parseTime = parseTime + (CLng(Temp(1)) * 60)
    parseTime = parseTime + (CLng(Left$(Temp(2), 2)))
    parseTime = parseTime + (CLng(Mid$(Temp(2), 4)) * 0.01)
    
End Function


' Return string with aspect ratio, using 4:3 as default
Public Function visualAspectRatio(ByVal Aspect As Single) As String

    If Aspect = 16 / 9 Then
        visualAspectRatio = "Widescreen (16:9)"
    ElseIf Aspect = 4 / 3 Then
        visualAspectRatio = "Normal (4:3)"
    ElseIf Aspect = 2.21 Then
        visualAspectRatio = "Panavision (2.21:1)"
    Else
        visualAspectRatio = Round(Aspect, 2) & ":1"
    End If
        
End Function


' Return whether a path cotnains non standard ASCII characters
Public Function oddPath(ByVal Path As String) As Boolean

    Dim A As Long
    
    
    For A = 1 To Len(Path)
        If AscW(Mid$(Path, A, 1)) > 127 Then
            oddPath = True
            Exit Function
        End If
    Next A

End Function


' Convert allowed special XML characters in path or filenames to entities
Public Function pathEntities(ByVal inPath As String) As String

    Dim A As Long
    

    inPath = Replace(inPath, vbQuote, "&quot;")
    inPath = Replace(inPath, "&", "&amp;")
    
    pathEntities = inPath
    
End Function


' Return a properly generated filename from a title and audio\video index
Public Function audioFileName(ByVal A As Long, ByVal B As Long, ByVal C As Long) As String

    audioFileName = Project.destinationDir & "\" & A & "." & B
    If C <> -1 Then audioFileName = audioFileName & "." & C
    
    audioFileName = audioFileName & ".ac3"
    
End Function

Public Function videoFileName(ByVal A As Long, ByVal B As Long) As String

    videoFileName = Project.destinationDir & "\" & A & "." & B
    videoFileName = videoFileName & ".m2v"
    
End Function


' Return a nicely formatted version string of this app
Public Function versionString() As String

    versionString = App.Major & "." & Mid$(App.Minor, 1, 1)
    versionString = versionString & "." & Mid$(App.Minor, 2, 1)
    versionString = versionString & "." & Mid$(App.Minor, 3, 1)
    
    If APP_BETA Then
        versionString = versionString & " beta"
    ElseIf APP_RC > 0 Then
        versionString = versionString & " RC" & APP_RC
    End If
    
    versionString = versionString & " build " & App.Revision

End Function


' Return a time in seconds in hours:minutes:seconds format
Public Function visualTime(ByVal Duration As Single) As String

    Dim Hours As Long
    Dim Minutes As Long
    Dim Seconds As Long
    
    
    Do
        If Duration >= 3600 Then
            Hours = Hours + 1
            Duration = Duration - 3600
        ElseIf Duration >= 60 Then
            Minutes = Minutes + 1
            Duration = Duration - 60
        Else
            Seconds = Fix(Duration)
            Exit Do
        End If
    Loop

    If Hours Then
        visualTime = Hours & ":" & padLeft(Minutes, "0", 2) & ":" & padLeft(Seconds, "0", 2) & " hours"
    ElseIf Minutes Then
        visualTime = Minutes & ":" & padLeft(Seconds, "0", 2) & " minutes"
    Else
        visualTime = Seconds & " seconds"
    End If

End Function


' Return a time in seconds in HR:MM:SS.MS format
Public Function displayTime(ByVal Duration As Single) As String

    Dim Hours As Long
    Dim Minutes As Long
    Dim Seconds As Long
    Dim mS As Long
    

    Do
        If Duration >= 3600 Then
            Hours = Hours + 1
            Duration = Duration - 3600
        ElseIf Duration >= 60 Then
            Minutes = Minutes + 1
            Duration = Duration - 60
        ElseIf Duration > 1 Then
            Seconds = Fix(Duration)
            Duration = Duration - Seconds
        Else
            mS = Fix(Duration * 1000)
            Exit Do
        End If
    Loop

    displayTime = padLeft(Hours, "0", 2) & ":"
    displayTime = displayTime & padLeft(Minutes, "0", 2) & ":"
    displayTime = displayTime & padLeft(Seconds, "0", 2) & "."
    displayTime = displayTime & padLeft(mS, "0", 3)

End Function


' Pad the left side of a string with a character
Public Function padLeft(ByVal inString As String, ByVal padChar As String, targetLength As Long) As String

    If Len(inString) < targetLength Then
        padLeft = String(targetLength - Len(inString), padChar) & inString
    
    Else
        padLeft = inString
        
    End If

End Function


' Pad the right side of a string with a character
Public Function padRight(ByVal inString As String, ByVal padChar As String, targetLength As Long) As String

    If Len(inString) < targetLength Then
        padRight = inString & String(targetLength - Len(inString), padChar)
    
    Else
        padRight = inString
        
    End If

End Function


' Return a time in seconds as a minutes:seconds display
Public Function minuteTime(ByVal Duration As Single) As String

    Dim Hours As Long
    Dim Minutes As Long
    Dim Seconds As Long
    
    
    Do
        If Duration >= 60 Then
            Minutes = Minutes + 1
            Duration = Duration - 60
        Else
            Seconds = Fix(Duration)
            Exit Do
        End If
    Loop

    If Seconds < 10 Then
        minuteTime = Minutes & ":0" & Seconds
    Else
        minuteTime = Minutes & ":" & Seconds
    End If

End Function
