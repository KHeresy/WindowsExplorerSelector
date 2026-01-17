// pch.h: 此為先行編譯的標頭檔。
// 以下所列檔案只會編譯一次，可改善之後組建的組建效能。
// 這也會影響 IntelliSense 效能，包括程式碼完成以及許多程式碼瀏覽功能。
// 但此處所列的檔案，如果其中任一在組建之間進行了更新，即會重新編譯所有檔案。
// 請勿於此處新增會經常更新的檔案，如此將會對於效能優勢產生負面的影響。

#ifndef PCH_H
#define PCH_H

// 請於此新增您要先行編譯的標頭
#include "framework.h"
#include <shlobj.h>
#include <shlwapi.h>
#include <string>
#include <tchar.h>
#include <vector>

#pragma comment(lib, "shlwapi.lib")

extern HINSTANCE g_hInst;
extern long g_cDllRef;

#define CLSID_ExplorerSelector_String L"{3289050C-926F-4C10-8531-9F31E63C2C90}"
// {3289050C-926F-4C10-8531-9F31E63C2C90}
static const GUID CLSID_ExplorerSelector = 
{ 0x3289050c, 0x926f, 0x4c10, { 0x85, 0x31, 0x9f, 0x31, 0xe6, 0x3c, 0x2c, 0x90 } };


#endif //PCH_H