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
//   File purpose: DVD Flick support functions
//

#include "windows.h"
#include "dll.h"


BITMAP	bmpData;
BYTE*	bmpBits;


// Retrieve image's properties
int EXPORT GDIImageSelect(HANDLE bmpHandle)
{
	
	if (!GetObject(bmpHandle, sizeof(BITMAP), &bmpData))
		return 1;
	else
		bmpBits = (BYTE*)bmpData.bmBits;

	return 0;

}


// Get a pixel from the current bitmap
static int getPixel(int x, int y)
{

	LPRGBQUAD	rgbResultQ;
	LPRGBTRIPLE	rgbResultT;


	if (x < 0 || x >= bmpData.bmWidth || y < 0 || y >= bmpData.bmHeight || !bmpBits)
		return(-1);

	switch(bmpData.bmBitsPixel)
	{
	case 8:
		return *(LPBYTE)(bmpBits + bmpData.bmWidthBytes * y + x);
		break;
	case 24:
		rgbResultT = (LPRGBTRIPLE)(bmpBits + bmpData.bmWidthBytes * y + x * 3);
		return RGB(rgbResultT->rgbtRed, rgbResultT->rgbtGreen, rgbResultT->rgbtBlue);
		break;
	case 32:
		rgbResultQ = (LPRGBQUAD)(bmpBits + bmpData.bmWidthBytes * y + (x << 2));
		return RGB(rgbResultQ->rgbRed, rgbResultQ->rgbGreen, rgbResultQ->rgbBlue);
		break;
	default:
		return(-1);
	}

}


// Set a pixel to the current bitmap
static void setPixel(int x, int y, int Color)
{

	LPRGBQUAD	rgbColorQ;
	LPRGBTRIPLE	rgbColorT;


	if (x < 0 || x >= bmpData.bmWidth || y < 0 || y >= bmpData.bmHeight || !bmpBits)
		return;

	switch(bmpData.bmBitsPixel)
	{
	case 8:
		*(LPBYTE)(bmpBits + bmpData.bmWidthBytes * y + x) = (BYTE)Color;
		return;
		break;
	case 24:
		rgbColorT = (LPRGBTRIPLE)(bmpBits + bmpData.bmWidthBytes * y + x * 3);
		rgbColorT->rgbtRed = GetRValue(Color);
		rgbColorT->rgbtGreen = GetGValue(Color);
		rgbColorT->rgbtBlue = GetBValue(Color);
		break;
	case 32:
		rgbColorQ = (LPRGBQUAD)(bmpBits + bmpData.bmWidthBytes * y + (x << 2));
		rgbColorQ->rgbRed = GetRValue(Color);
		rgbColorQ->rgbGreen = GetGValue(Color);
		rgbColorQ->rgbBlue = GetBValue(Color);
		break;
	default:
		return;
	}

}


// Fill a bitmap with a color
int EXPORT GDIColorFill(int Color)
{

	int x, y;


	if (!bmpBits)
		return 1;

	for(x = 0; x < bmpData.bmWidth; x++)
		for(y = 0; y < bmpData.bmHeight; y++)
			setPixel(x, y, Color);

	return 0;

}


// Replace a color in a bitmap with another
int EXPORT GDIColorReplace(int colorReplace, int Color)
{

	int x, y;


	if (!bmpBits)
		return 1;

	for(x = 0; x < bmpData.bmWidth; x++)
		for(y = 0; y < bmpData.bmHeight; y++)
			if (getPixel(x, y) == colorReplace) setPixel(x, y, Color);

	return 0;

}


// Alphablend a 32-bit image onto the current
int EXPORT GDIAlphaBlit(HANDLE imgHandle, int x, int y)
{

	int cX, cY;
	RGBQUAD*	qSource;
	RGBQUAD*	qDest;

	BITMAP	imgData;
	BYTE*	imgBits;


	if (!GetObject(imgHandle, sizeof(BITMAP), &imgData))
		return 1;
	else
		imgBits = (LPBYTE)imgData.bmBits;

	if (imgData.bmBitsPixel != 32 || bmpData.bmBitsPixel != 32 || !bmpBits)
		return 1;

	for(cX = x; cX < x + imgData.bmWidth; cX++)
	{
		for(cY = y; cY < y + imgData.bmHeight; cY++)
		{
			if (cX >= 0 && cX < bmpData.bmWidth && cY >= 0 && cY < bmpData.bmHeight)
			{
				qSource = (LPRGBQUAD)(imgBits + imgData.bmWidthBytes * (imgData.bmHeight - (cY - y)) + ((cX - x) << 2));
				qDest   = (LPRGBQUAD)(bmpBits + bmpData.bmWidthBytes * (bmpData.bmHeight - cY) +  (cX << 2));

				qDest->rgbRed      = ((qDest->rgbRed      * qSource->rgbReserved) + qSource->rgbRed      * (255 - qSource->rgbReserved) ) / 255;
				qDest->rgbGreen    = ((qDest->rgbGreen    * qSource->rgbReserved) + qSource->rgbGreen    * (255 - qSource->rgbReserved) ) / 255;
				qDest->rgbBlue     = ((qDest->rgbBlue     * qSource->rgbReserved) + qSource->rgbBlue     * (255 - qSource->rgbReserved) ) / 255;
				qDest->rgbReserved = ((qDest->rgbReserved * qSource->rgbReserved) + qSource->rgbReserved * (255 - qSource->rgbReserved) ) / 255;
			}
		}
	}

	return 0;

}


// Render an outline on an image
// If baseColor is encountered, a circle is drawn outwards from it of size outlineSize
int EXPORT GDIRenderOutline(int outlineSize, int baseColor, int backColor, int outlineColor)
{

	int		i;
	int		x, y;

	signed int cX, cY, P;


	if (!bmpBits)
		return 1;

	for (i = 1; i <= outlineSize; i++)
	{
		for(x = i; x < bmpData.bmWidth - i; x++)
		{
			for(y = i; y < bmpData.bmHeight - i; y++)
			{
				if (getPixel(x, y) == baseColor)
				{
					cX = 0;
					cY = i;
					P = (3 - (2 * i));

					while (cX < cY)
					{
						if (getPixel(x + cX, y + cY) == backColor) setPixel(x + cX, y + cY, outlineColor);
						if (getPixel(x + cY, y + cX) == backColor) setPixel(x + cY, y + cX, outlineColor);
						if (getPixel(x + cY, y - cX) == backColor) setPixel(x + cY, y - cX, outlineColor);
						if (getPixel(x + cX, y - cY) == backColor) setPixel(x + cX, y - cY, outlineColor);
						if (getPixel(x - cX, y - cY) == backColor) setPixel(x - cX, y - cY, outlineColor);
						if (getPixel(x - cY, y - cX) == backColor) setPixel(x - cY, y - cX, outlineColor);
						if (getPixel(x - cY, y + cX) == backColor) setPixel(x - cY, y + cX, outlineColor);
						if (getPixel(x - cX, y + cY) == backColor) setPixel(x - cX, y + cY, outlineColor);

						cX++;
						if (P < 0)
							P = P + (cX << 2) + 6;
						else
						{
							cY--;
							P = P + (((cX - cY) << 2) + 1);
						}
					}
				}
			}
		}
	}

	return 0;

}