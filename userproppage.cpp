//+-------------------------------------------------------------------------
//
//  Microsoft Windows
//
//  Copyright (C) Microsoft Corporation, 1999 - 2000
//
//  File:       userproppage.cpp
//
//--------------------------------------------------------------------------

// dscreate.cpp : Implementation of DLL Exports.

#include "stdafx.h"

// initguid.h requires this.
//
#include <objbase.h>

// this redefines the DEFINE_GUID() macro to do allocation.
//
#include <initguid.h>

//#include <dsshell.h>

#include <dsadmin.h>

#include "resource.h"
#include "userproppage.h"
#include "ProppageUser.h"

// {40EC22BE-2BCC-4FB4-85F7-BC775C59A8C5}
extern const GUID CLSID_ProppageUser = 
{ 0x40ec22be, 0x2bcc, 0x4fb4, { 0x85, 0xf7, 0xbc, 0x77, 0x5c, 0x59, 0xa8, 0xc5 } };

CMyUIExtModule _Module;

BEGIN_OBJECT_MAP(ObjectMap)
  OBJECT_ENTRY(CLSID_ProppageUser,    CProppageUser)
END_OBJECT_MAP()

class CMyUIExtApp : public CWinApp
{
public:
	virtual BOOL InitInstance();
	virtual int ExitInstance();
};

CMyUIExtApp theApp;

BOOL CMyUIExtApp::InitInstance()
{
	_Module.Init(ObjectMap, m_hInstance);
	return CWinApp::InitInstance();
}

int CMyUIExtApp::ExitInstance()
{
	_Module.Term();
	return CWinApp::ExitInstance();
}

/////////////////////////////////////////////////////////////////////////////
// Used to determine whether the DLL can be unloaded by OLE

STDAPI DllCanUnloadNow(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return (AfxDllCanUnloadNow()==S_OK && _Module.GetLockCount()==0) ? S_OK : S_FALSE;
}

/////////////////////////////////////////////////////////////////////////////
// Returns a class factory to create an object of the requested type

STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID* ppv)
{
	return _Module.GetClassObject(rclsid, riid, ppv);
}

/////////////////////////////////////////////////////////////////////////////
// DllRegisterServer - Adds entries to the system registry

BOOL x = TRUE;
STDAPI DllRegisterServer(void)
{
    return _Module.RegisterServer(FALSE);
}

/////////////////////////////////////////////////////////////////////////////
// DllUnregisterServer - Removes entries from the system registry

STDAPI DllUnregisterServer(void)
{
	_Module.UnregisterServer();
	return S_OK;
}


