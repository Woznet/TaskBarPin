/*
Copyright (c) 2020 by Gee Law

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
**/

#pragma comment(lib, "ole32")
#pragma comment(lib, "shell32")

#define STRICT_TYPED_ITEMIDS
#define WIN32_LEAN_AND_MEAN
#define UNICODE
#define _UNICODE

#include<objbase.h>
#include<shlobj.h>
#include<cstdio>

struct CoInitializeGuard
{
    HRESULT const hr;
    CoInitializeGuard() : hr(CoInitialize(NULL)) { }
    ~CoInitializeGuard()
    {
        if (SUCCEEDED(hr))
        {
            CoUninitialize();
        }
    }
    operator HRESULT() const { return hr; }
};

struct PIDLFromPath
{
    PIDLIST_ABSOLUTE const pidl;
    PIDLFromPath(PCWSTR path) : pidl(ILCreateFromPathW(path)) { }
    ~PIDLFromPath()
    {
        if (pidl)
        {
            ILFree(pidl);
        }
    }
    operator PIDLIST_ABSOLUTE() const { return pidl; }
};

const GUID CLSID_TaskbandPin =
{
    0x90aa3a4e, 0x1cba, 0x4233,
    { 0xb8, 0xbb, 0x53, 0x57, 0x73, 0xd4, 0x84, 0x49}
};

const GUID IID_IPinnedList3 =
{
    0x0dd79ae2, 0xd156, 0x45d4,
    { 0x9e, 0xeb, 0x3b, 0x54, 0x97, 0x69, 0xe9, 0x40 }
};

enum PLMC { PLMC_EXPLORER = 4 };

struct IPinnedList3Vtbl;
struct IPinnedList3 { IPinnedList3Vtbl *vtbl; };

typedef ULONG STDMETHODCALLTYPE ReleaseFuncPtr(IPinnedList3 *that);
typedef HRESULT STDMETHODCALLTYPE ModifyFuncPtr(IPinnedList3 *that,
    PCIDLIST_ABSOLUTE unpin, PCIDLIST_ABSOLUTE pin, PLMC caller);

struct IPinnedList3Vtbl
{
    void *QueryInterface;
    void *AddRef;
    ReleaseFuncPtr *Release;
    void *MethodSlot4; void *MethodSlot5; void *MethodSlot6;
    void *MethodSlot7; void *MethodSlot8; void *MethodSlot9;
    void *MethodSlot10; void *MethodSlot11; void *MethodSlot12;
    void *MethodSlot13; void *MethodSlot14; void *MethodSlot15;
    void *MethodSlot16;
    ModifyFuncPtr *Modify;
};

wchar_t const *ERR_STR_USAGE = L"Usage:\n"
"  pin \"C:\\path\\to\\file.lnk\"      Pin a Shortcut.\n"
"  pin \"C:\\path\\to\\file.lnk\" u    Unpin a Shortcut.\n"
"\n"
"Exit codes: -1 = printed usage\n"
"             0 = succeeded\n"
"             1 = CoInitialize failed\n"
"             2 = ILCreateFromPathW failed\n"
"             3 = CoCreateInstance failed\n"
"             4 = IPinnedList3::Modify failed\n";
wchar_t const *ERR_STR_COINIT = L"CoInitialize failed.\n";
wchar_t const *ERR_STR_PIDL = L"ILCreateFromPathW failed.\n";
wchar_t const *ERR_STR_CREATE = L"CoCreateInstance failed.\n";
wchar_t const *ERR_STR_MODIFY = L"IPinnedList3::Modify failed.\n";

int const ERR_ID_USAGE = -1, ERR_ID_SUCCEEDED = 0,
    ERR_ID_COINIT = 1, ERR_ID_PIDL = 2,
    ERR_ID_CREATE = 3, ERR_ID_MODIFY = 4;

int wmain(int argc, PWSTR *argv)
{
    if (argc < 2 || argc > 3)
    {
        fputws(ERR_STR_USAGE, stderr);
        return ERR_ID_USAGE;
    }
    bool pinning = (argc != 3 || argv[2][0] != L'u');
    CoInitializeGuard guard;
    if (!SUCCEEDED(guard))
    {
        fputws(ERR_STR_COINIT, stderr);
        return ERR_ID_COINIT;
    }
    PIDLFromPath pidl(argv[1]);
    if (!(bool)pidl)
    {
        fputws(ERR_STR_PIDL, stderr);
        return ERR_ID_PIDL;
    }
    IPinnedList3 *pinnedList;
    if (!SUCCEEDED(CoCreateInstance(
        CLSID_TaskbandPin, NULL, CLSCTX_ALL,
        IID_IPinnedList3, (LPVOID *)(&pinnedList))))
    {
        fputws(ERR_STR_CREATE, stderr);
        return ERR_ID_CREATE;
    }
    HRESULT hr = pinnedList->vtbl->Modify(pinnedList,
        pinning ? NULL : pidl,
        pinning ? pidl : NULL,
        PLMC_EXPLORER);
    pinnedList->vtbl->Release(pinnedList);
    if (!SUCCEEDED(hr))
    {
        fputws(ERR_STR_MODIFY, stderr);
        return ERR_ID_MODIFY;
    }
    return ERR_ID_SUCCEEDED;
}
