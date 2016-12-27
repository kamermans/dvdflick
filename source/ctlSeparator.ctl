VERSION 5.00
Begin VB.UserControl ctlSeparator 
   BackStyle       =   0  'Transparent
   CanGetFocus     =   0   'False
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3510
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   234
   Begin VB.Line lnLow 
      BorderColor     =   &H80000010&
      X1              =   44
      X2              =   44
      Y1              =   4
      Y2              =   232
   End
   Begin VB.Line lnHigh 
      BorderColor     =   &H80000014&
      X1              =   48
      X2              =   48
      Y1              =   4
      Y2              =   232
   End
End
Attribute VB_Name = "ctlSeparator"
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
'   File purpose: Simple 3D look separator.
'
Option Explicit
Option Compare Binary
Option Base 0


Public Enum OrientateConstants
    orientateHorizontal = 0
    orientateVertical
End Enum


Private mOrientation As OrientateConstants


Public Property Get Orientation() As OrientateConstants

    Orientation = mOrientation

End Property

Public Property Let Orientation(ByVal newValue As OrientateConstants)

    mOrientation = newValue
    updateLines

End Property


Private Sub UserControl_Resize()

    updateLines

End Sub


Private Sub updateLines()

    If mOrientation = orientateHorizontal Then
        With lnHigh
            .X1 = 0
            .X2 = UserControl.Width
            .Y1 = 1
            .Y2 = 1
        End With
        
        With lnLow
            .X1 = 0
            .X2 = UserControl.Width
            .Y1 = 0
            .Y2 = 0
        End With
        
    ElseIf mOrientation = orientateVertical Then
        With lnHigh
            .X1 = 1
            .X2 = 1
            .Y1 = 0
            .Y2 = UserControl.Height
        End With
        
        With lnLow
            .X1 = 0
            .X2 = 0
            .Y1 = 0
            .Y2 = UserControl.Height
        End With
        
    End If

End Sub
