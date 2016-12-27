Attribute VB_Name = "modCommonDialog"
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
'   File purpose: Common dialog class support.
'
Option Explicit
Option Compare Binary
Option Base 0


' Constants for file open dialog
Public Enum cdlFileOpenConstants
    cdlOFNAllowMultiselect = 512
    cdlOFNCreatePrompt = 8192
    cdlOFNExplorer = 524288
    cdlOFNExtensionDifferent = 1024
    cdlOFNFileMustExist = 4096
    cdlOFNHelpButton = 16
    cdlOFNHideReadOnly = 4
    cdlOFNLongNames = 2097152
    cdlOFNNoChangeDir = 8
    cdlOFNNoDereferenceLinks = 1048576
    cdlOFNNoLongNames = 262144
    cdlOFNNoReadOnlyReturn = 32768
    cdlOFNNoValidate = 256
    cdlOFNOverwritePrompt = 2
    cdlOFNPathMustExist = 2048
    cdlOFNReadOnly = 1
    cdlOFNShareAware = 16384
End Enum

' Folder browsing constants
Public Enum browseFolderConstants
    BIF_RETURNONLYFSDIRS = &H1
    BIF_DONTGOBELOWDOMAIN = &H2
    BIF_STATUSTEXT = &H4
    BIF_RETURNFSANCESTORS = &H8
    BIF_EDITBOX = &H10
    BIF_VALIDATE = &H20
    BIF_NEWDIALOGSTYLE = &H40
    BIF_USENEWUI = (BIF_NEWDIALOGSTYLE Or BIF_EDITBOX)
    BIF_BROWSEINCLUDEURLS = &H80
    BIF_UAHINT = &H100
    BIF_NONEWFOLDERBUTTON = &H200
    BIF_BROWSEFORCOMPUTER = &H1000
    BIF_BROWSEFORPRINTER = &H2000
    BIF_BROWSEINCLUDEFILES = &H4000
    BIF_SHAREABLE = &H8000&
End Enum

' Color chooser constants
Public Enum colorChooseConstants
    CC_ANYCOLOR = &H100
    CC_ENABLEHOOK = &H10
    CC_ENABLETEMPLATE = &H20
    CC_ENABLETEMPLATEHANDLE = &H40
    CC_FULLOPEN = &H2
    CC_PREVENTFULLOPEN = &H4
    CC_RGBINIT = &H1
    CC_SHOWHELP = &H8
    CC_SOLIDCOLOR = &H80
End Enum


' Message constants
Private Const WM_USER = &H400
Private Const BFFM_INITIALIZED = 1
Private Const BFFM_SETSELECTION = (WM_USER + 102)


' Folder to set when callback occurs
Public callBackFolder As String


' API
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long


' Callback function for the browse for folder dialog, so that the starting folder can be set
Public Function browseFolderCallback(ByVal hWnd As Long, ByVal uMsg As Long, ByVal lp As Long, ByVal pData As Long) As Long
  
    On Error Resume Next
    
    
    If uMsg = BFFM_INITIALIZED Then SendMessage hWnd, BFFM_SETSELECTION, 1, callBackFolder
    browseFolderCallback = 0
    
    On Error GoTo 0
  
End Function
