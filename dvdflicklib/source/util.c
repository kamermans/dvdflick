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
//   File purpose: Utility functions and macros
//

#include "dll.h"


#define REVERSE_INT(val)	(((val >> 24) & 0xFF) | ((val >> 8) & 0xFF00) | ((val << 8) & 0xFF0000) | ((val << 24) & 0xFF000000))
#define REVERSE_SHORT(val)	(((val >> 8) & 0xFF) | ((val << 8) & 0xFF00))


#define RED(x)		(((x) << 16) & 0xF8000000)
#define GREEN(x)	(((x) << 13) & 0x00FC0000)
#define BLUE(x)		(((x) << 11) & 0x0000F800)


int EXPORT shortToRGB8(short Value)
{
	return RED(Value) | GREEN(Value) | BLUE(Value);
}


int EXPORT getRed(int Value)
{
	return (unsigned char)Value;
}

int EXPORT getGreen(int Value)
{
	return (unsigned char)(Value >> 8);
}

int EXPORT getBlue(int Value)
{
	return (unsigned char)(Value >> 16);
}


unsigned char EXPORT reverseByte(unsigned char V)
{

	V = (V & 0x0F) << 4 | (V & 0xF0) >> 4;
	V = (V & 0x33) << 2 | (V & 0xCC) >> 2;
	V = (V & 0x55) << 1 | (V & 0xAA) >> 1;
	
	return V;
}


// Reverse long\int bits
int EXPORT reverseLong(int Value)
{
	return REVERSE_INT(Value);
}


// Reverse integer\short bits
short EXPORT reverseInteger(short Value)
{
	return REVERSE_SHORT(Value);
}


int EXPORT getBit(int Value, int Bit)
{
	return (Value | (2^Bit));
}