// -------------------------------------------------------------------------------
//
//  DVD Flick - A DVD authoring program
//  Copyright (C) 2006-2008  Dennis Meuwissen
//
//  This program is free software; you can redistribute it and/or modify
//  it under the terms of the GNU General Public License as published by
//  the Free Software Foundation; either version 2 of the License, or
//  (at your option) any later version.
//
//  This program is distributed in the hope that it will be useful,
//  but WITHOUT ANY WARRANTY; without even the implied warranty of
//  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
//  GNU General Public License for more details.
//
//  You should have received a copy of the GNU General Public License
//  along with this program; if not, write to the Free Software
//  Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301, USA.
//
// -------------------------------------------------------------------------------
//
//   File purpose: Windows DLLMain
//

//
// Windows DLL export
// http://www.mingw.org/MinGWiki/index.php/VB-MinGW-DLL
//

#include "windows.h"


__declspec (dllexport) BOOL __stdcall DllMain(HANDLE hModule, DWORD  ul_reason_for_call, LPVOID lpReserved)
{
    switch (ul_reason_for_call)
    {
    case DLL_PROCESS_ATTACH:
            break;
    case DLL_THREAD_ATTACH:
            break;
    case DLL_THREAD_DETACH:
            break;
    case DLL_PROCESS_DETACH:
            break;
    }
    return TRUE;
}