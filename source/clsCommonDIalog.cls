VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCommonDialog"
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
'   File purpose: Some commonly used dialogs for file\folder\color selection.
'
Option Explicit
Option Compare Binary
Option Base 0


' File Open/Save structure
Private Type cdlOpenFileName
    lStructSize As Long
    Owner As Long
    hInstance As Long
    Filter As Long
    customFilter As Long
    maxCustFilter As Long
    filterIndex As Long
    File As Long
    maxFile As Long
    fileTitle As Long
    maxFileTitle As Long
    initialDir As Long
    Title As Long
    flags As Long
    fileOffset As Integer
    fileExtension As Integer
    defExt As Long
    custData As Long
    Hook As Long
    templateName As Long
End Type

' Folder structure
Private Type browseInfo
    hwndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type

' Color choose
Private Type chooseColorInfo
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As Long
    flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type


' Maximum length of a filename
Private Const MAX_FILELEN = 2048

' Last directory accessed
Private mLastFolder As String


' API calls
Private Declare Function GetOpenFileName Lib "comdlg32" Alias "GetOpenFileNameW" (ByRef pOpenfilename As cdlOpenFileName) As Long
Private Declare Function GetSaveFileName Lib "comdlg32" Alias "GetSaveFileNameW" (ByRef pSavefilename As cdlOpenFileName) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" Alias "SHBrowseForFolderW" (lpbi As browseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function ChooseColor Lib "comdlg32" Alias "ChooseColorA" (lpcc As chooseColorInfo) As Long
Private Declare Function CommDlgExtendedError Lib "comdlg32" () As Long
Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal hMem As Long)


' lastFolder property
Public Property Get lastFolder() As String

    lastFolder = mLastFolder

End Property

Public Property Let lastFolder(ByVal newValue As String)

    mLastFolder = newValue

End Property


' Color picker dialog
Public Function getColor(ByVal hWnd As Long, ByVal flags As colorChooseConstants, ByVal defaultCol As Long) As Long
    
    Dim chooseData As chooseColorInfo
    Dim retVal As Long
    Static userCols(16) As Long
    
    
    ' Init data struct
    With chooseData
        .lStructSize = Len(chooseData)
        .hInstance = App.hInstance
        .hwndOwner = hWnd
        .flags = flags
        .rgbResult = defaultCol
        .lpCustColors = VarPtr(userCols(0))
    End With
    
    ' Show dialog
    If ChooseColor(chooseData) <> 0 Then
        getColor = chooseData.rgbResult
    Else
        getColor = -1
    End If
    
End Function


' Turn a string returned by the dialog functions into a Dictionary of file items
Private Function getFileList(ByVal fileNames As String) As Dictionary

    Dim A As Long
    Dim nullOffset As Long
    Dim fileDir As String
    Dim Char As String
    Dim cFile As String
    
    Dim Files() As String
    Dim fileCount As Long
    Dim noSort As Boolean
    Dim tempFile As String
    
    
    ' Parse the line into separate files and add them
    nullOffset = InStr(fileNames, vbNullChar)
    If nullOffset > 0 Then
        
        ' Parse the directory
        fileDir = Mid$(fileNames, 1, nullOffset - 1)
        If Right$(fileDir, 1) <> "\" Then fileDir = fileDir & "\"
        
        ' Parse the rest of the line
        For A = nullOffset + 1 To Len(fileNames)
            Char = Mid$(fileNames, A, 1)
            
            ' Next file
            If Char = vbNullChar Then
                ReDim Preserve Files(fileCount)
                fileCount = fileCount + 1
                Files(fileCount - 1) = fileDir & cFile
                
                Char = vbNullString
                cFile = vbNullString
            End If
            
            cFile = cFile & Char
        Next A
        
        ' Add the last file
        ReDim Preserve Files(fileCount)
        fileCount = fileCount + 1
        Files(fileCount - 1) = fileDir & cFile
        
    Else
        ReDim Files(0)
        fileCount = 1
        Files(0) = fileNames
        
    End If

    
    ' Sort it using bubble sort algo
    While noSort = True
        noSort = True

        For A = 0 To fileCount - 2
            If StrComp(Files(A), Files(A + 1), vbTextCompare) = 1 Then
                tempFile = Files(A)
                Files(A) = Files(A + 1)
                Files(A + 1) = tempFile

                noSort = False
            End If
        Next A
    Wend
    
    
    ' Compile into dictionary object
    Set getFileList = New Dictionary
    For A = 0 To fileCount - 1
        If LenB(Files(A)) <> 0 Then getFileList.Add A, Files(A)
    Next A

End Function


' Select a folder
Public Function selectFolder(ByVal hWnd As Long, ByVal Prompt As String) As String
     
    Dim intNull As Integer
    Dim Result As Long
    Dim pathResult As Long
    Dim strPath As String
    Dim folderBrowseInfo As browseInfo
    
    
    ' Set browse information
    modCommonDialog.callBackFolder = mLastFolder
    With folderBrowseInfo
        .hwndOwner = hWnd
        .lpszTitle = StrPtr(Prompt)
        .ulFlags = BIF_USENEWUI Or BIF_RETURNONLYFSDIRS
        .lpfnCallback = getAddressOfFunction(AddressOf modCommonDialog.browseFolderCallback)
    End With
    
    ' Select folder
    Result = SHBrowseForFolder(folderBrowseInfo)
    
    ' Check if not cancelled
    If Result Then
    
        ' Get the selected path
        strPath = String(MAX_FILELEN, 0)
        pathResult = SHGetPathFromIDList(Result, strPath)
        
        ' Free used memory
        CoTaskMemFree Result
        
        ' Remove null terminator(s)
        strPath = stripNull(strPath)
    
    End If
    
    selectFolder = strPath
     
End Function


' Open (or select) a file
Public Function openFile(ByVal hWnd As Long, ByVal dialogTitle As String, ByVal Filter As String, ByVal initDir As String, ByVal initFilename As String, ByVal flags As cdlFileOpenConstants, Optional ByVal filterIndex As Long = 1) As Dictionary
     
    Dim fileStruct As cdlOpenFileName
    Dim nullOffset As Long
    Dim fileList As String
    
    
    ' Use last folder
    If LenB(initDir) = 0 Then initDir = mLastFolder
    
    With fileStruct
    
        ' Set default directory
        initDir = padRight(initDir, vbNullChar, MAX_FILELEN)
        initFilename = padRight(initFilename, vbNullChar, MAX_FILELEN)
        
        ' Set all dialog parameters
        .lStructSize = LenB(fileStruct)
        .hInstance = App.hInstance
        .Owner = hWnd
        .Title = StrPtr(dialogTitle)
        .Filter = StrPtr(Replace(Filter, "|", vbNullChar) & vbNullChar & vbNullChar)
        .flags = flags
        .File = StrPtr(initFilename)
        .filterIndex = filterIndex
        .maxFile = MAX_FILELEN
        .initialDir = StrPtr(initDir)
        
        ' Show the dialog and return the result
        If GetOpenFileName(fileStruct) <> 0 Then
            fileList = stripNull(initFilename)
            
            ' Parse the directory out
            nullOffset = InStr(fileList, vbNullChar)
            If nullOffset > 0 Then
                mLastFolder = Mid$(fileList, 1, nullOffset - 1)
            Else
                mLastFolder = FS.GetParentFolderName(fileList)
            End If
            
        Else
            appLog.Add "CommonDialog error 0x" & Hex$(CommDlgExtendedError)
            
        End If
    
    End With
    
    Set openFile = getFileList(fileList)
     
End Function


' Open (or select) a file
Public Function saveFile(ByVal hWnd As Long, ByVal dialogTitle As String, ByVal Filter As String, ByVal initDir As String, ByVal initFilename As String, ByVal flags As cdlFileOpenConstants, Optional ByVal filterIndex As Long = 1) As String
     
    Dim fileStruct As cdlOpenFileName
    Dim nullOffset As Long
    Dim fileTitle As String
    

    ' Use last folder
    If LenB(initDir) = 0 Then initDir = mLastFolder
    
    With fileStruct
        
        ' Pad\init some strings
        initDir = padRight(initDir, vbNullChar, MAX_FILELEN)
        initFilename = padRight(initFilename, vbNullChar, MAX_FILELEN)
        fileTitle = padRight(fileTitle, vbNullChar, MAX_FILELEN)
                
        ' Set all dialog parameters
        .lStructSize = LenB(fileStruct)
        .hInstance = App.hInstance
        .Owner = hWnd
        .Title = StrPtr(dialogTitle)
        .Filter = StrPtr(Replace(Filter, "|", vbNullChar) & vbNullChar & vbNullChar)
        .flags = flags
        .File = StrPtr(initFilename)
        .fileTitle = StrPtr(fileTitle)
        .maxFileTitle = MAX_FILELEN
        .filterIndex = filterIndex
        .maxFile = MAX_FILELEN
        .initialDir = StrPtr(initDir)
        
        ' Show the dialog and return the result
        If GetSaveFileName(fileStruct) <> 0 Then
            saveFile = stripNull(initFilename)
            mLastFolder = FS.GetParentFolderName(saveFile)
        
        Else
            appLog.Add "CommonDialog error 0x" & Hex$(CommDlgExtendedError)
            
        End If
    
    End With
     
End Function
