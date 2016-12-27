Attribute VB_Name = "modDefinitions"
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
'   File purpose: Enumerations and global constants and definitions
'
Option Explicit
Option Compare Binary
Option Base 0


' Debugging mode flags
Public Enum enumDebugModes
    DM_GDI = 1
    DM_Encoding = 2
    DM_Menus = 4
    DM_Subtitles = 8
    DM_SourceParser = 16
    DM_Pipes = 32
End Enum

' Constants used throughout
Public Enum enumVideoEncodeFormats
    VF_PAL = 0
    VF_NTSC
    VF_NTSCFILM
    VF_MIXED
End Enum

' Video encoding profiles
Public Enum enumVideoEncodeProfiles
    VE_Fastest = 0
    VE_Fast
    VE_Normal
    VE_Best
End Enum

' Supported DVD display aspect ratios
Public Enum enumVideoAspects
    VA_43 = 0
    VA_169
End Enum

' Audio channel count
Public Enum enumAudioChannels
    AC_Auto = 0
    AC_Mono
    AC_Stereo
    AC_Surround
End Enum

' Audio bitrate list
Public Enum enumAudioBitRates
    AB_Auto = 0
    AB_64
    AB_128
    AB_256
    AB_384
    AB_448
End Enum

' Target disc sizes
Public Enum enumTargetSizes
    TS_Custom = 0
    TS_650
    TS_700
    TS_800
    TS_DVD
    TS_DVDDL
    TS_DVDRAM
    TS_MINIDVD
    TS_MINIDVDDL
End Enum

' Possible target bitrates
Public Enum enumTargetBitRates
    TB_Auto = 0
    TB_2MB
    TB_3MB
    TB_4MB
    TB_5MB
    TB_6MB
    TB_7MB
    TB_8MB
    TB_Custom
End Enum

' Types of streams in a media file
Public Enum enumStreamTypes
    ST_Unknown = 0
    ST_Audio
    ST_Video
End Enum

' Encoder process priorities
Public Enum enumEncodePriorities
    EP_AboveNormal
    EP_Normal
    EP_BelowNormal
    EP_Idle
End Enum

' Actions to do when a title has finished playing
Public Enum enumWhenPlayedActions
    PA_NextTitle = 0
    PA_SameTitle
    PA_Stop
    PA_Menu
End Enum

' Subtitle alignemnt types
Public Enum enumSubAlignment
    SA_TopLeft = 0
    SA_TopCenter
    SA_TopRight
    SA_CenterLeft
    SA_CenterRight
    SA_BottomLeft
    SA_BottomCenter
    SA_BottomRight
End Enum


' General defaults
Public Default_WindowTop As Long
Public Default_WindowLeft As Long
Public Default_WindowWidth As Long
Public Default_WindowHeight As Long
Public Default_LastOutputDir As String
Public Default_LastBrowseDir As String
Public Const Default_WindowState As Long = vbMaximized

' Project general
Public Const Default_ProjectTitle As String = "Unnamed project"
Public Const Default_TargetSize As Long = TS_DVD
Public Const Default_CustomSize As Long = 4489
Public Const Default_EncodePriority As Long = EP_BelowNormal
Public Const Default_KeepFiles As Long = 0
Public Const Default_MPLEX As Long = 0
Public Default_ThreadCount As Long

' Video
Public Default_TargetFormat As enumVideoEncodeFormats
Public Const Default_EncodeProfile As Long = VE_Normal
Public Const Default_TargetBitRate As Long = TB_Auto
Public Const Default_CustomBitRate As Long = 0
Public Const Default_Pulldown As Long = 0

' Audio
Public Const Default_VolumeMod As Long = 100
Public Const Default_ChannelCount As Long = AC_Auto
Public Const Default_AudioBitRate As Long = AB_Auto

' Burning
Public Const Default_CreateISO As Long = 0
Public Const Default_DeleteISO As Long = 0
Public Const Default_EraseRW As Long = 1
Public Const Default_DiscLabel As String = "DVD Video"
Public Const Default_EnableBurning As Long = 0
Public Const Default_BurnSpeed As String = "4x"
Public Const Default_VerifyDisc As Byte = 0
Public Const Default_EjectTray As Byte = 0
Public Default_BurnerName As String

' Playback
Public Const Default_LoopPlayback As Byte = 1
Public Const Default_WhenPlayed As Long = PA_NextTitle
Public Const Default_EnableFirstSub As Byte = 0

' Video advanced
Public Const Default_PSNR As Long = 0
Public Const Default_HalfRes As Long = 0
Public Const Default_OverscanBorders As Long = 0
Public Const Default_OverscanSize As Single = 3
Public Const Default_Deinterlace As Byte = 0
Public Const Default_MPEG2Copy As Byte = 0
Public Const Default_DCPrecision As Byte = 8

' Subtitle defaults
Public Const Default_SubFont As String = "Verdana"
Public Const Default_SubFontSize As String = 26
Public Const default_SubFontBold As Byte = 1
Public Const default_SubCodePage As Long = 1252

Public Const Default_SubColorText As Long = vbWhite
Public Const Default_SubColorOutline As Long = vbBlack
Public Const Default_SubColorBack As Long = vbBlack
Public Const Default_SubOutline As Long = 2
Public Const Default_SubAA As Byte = 1
Public Const Default_SubTransBack As Byte = 1

Public Const Default_SubMarginTop As Long = 40
Public Const Default_SubMarginBottom As Long = 40
Public Const Default_SubMarginLeft As Long = 40
Public Const Default_SubMarginRight As Long = 40
Public Const Default_SubAlignment As Long = SA_BottomCenter

Public Const Default_SubFramerate As Single = 23.976
Public Const Default_SubLanguage As Long = 28

' Chapters
Public Const Default_ChapterCount As Long = -1
Public Const Default_ChapterInterval As Long = -1
Public Const Default_ChapterOnSource As Byte = 1

' Menu defaults
Public Const Default_menuTemplateName As String = "None"
Public Const Default_menuAutoPlay As Byte = 1
Public Const Default_menuShowSubtitleFirst As Byte = 0
Public Const Default_menuShowAudioFirst As Byte = 0

' Error constants (unused I think)
Public Const Err_CreatePipe As Long = 2001
Public Const Err_PipeState As Long = 2002
Public Const Err_PipeProcess As Long = 2003


' Bitrate constraints
Public Const MIN_VIDEO_BITRATE As Long = 2000
Public Const MAX_VIDEO_BITRATE As Long = 9000
Public Const MAX_STREAM_BITRATE As Long = 9300
Public Const MAX_AUDIO_BITRATE As Long = 448

' GOP sizes
Public Const GOPSIZE_PAL As Long = 12
Public Const GOPSIZE_NTSC As Long = 14

' Default codepage
Public Const CODEPAGE_LATIN1 As Long = 1252

' Disabled menu string
Public Const STR_DISABLED_MENU As String = "None"
