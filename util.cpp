//+-------------------------------------------------------------------------
//
//  Microsoft Windows
//
//  Copyright (C) Microsoft Corporation, 1999 - 2000
//
//  File:       util.cpp
//
//--------------------------------------------------------------------------

#include "stdafx.h"
#include "util.h"

////////////////////////////////////////////////////////////////////

HRESULT _ClearAttribute(IADs* pIADs, BSTR bstrAttrName, BOOL bPostCommit) 
{
  ASSERT(pIADs != NULL);
	ASSERT(bstrAttrName != NULL);
  HRESULT hr = S_OK;

  VARIANT varVal;
  ::VariantInit(&varVal);

  // see if the attribute is there
  HRESULT hrFind = pIADs->Get(bstrAttrName, &varVal);
  if (SUCCEEDED(hrFind))
  {
    // found, need to clear the property
    if (bPostCommit)
    {
      // remove from the committed object
      hr = pIADs->PutEx(ADS_PROPERTY_CLEAR, bstrAttrName, varVal);
    }
    else
    {
      // remove from the cache of the temporary object
      IADsPropertyList* pPropList = NULL;
      hr = pIADs->QueryInterface(IID_IADsPropertyList, (void**)&pPropList);
      ASSERT(pPropList != NULL);
      if (SUCCEEDED(hr))
      {
        VARIANT v;
        ::VariantInit(&v);
        v.vt = VT_BSTR;
        v.bstrVal = ::SysAllocString(bstrAttrName);
        hr = pPropList->ResetPropertyItem(v);
        ::VariantClear(&v);
        pPropList->Release();
      }
    }
  }

  ::VariantClear(&varVal);
  return hr;
}


////////////////////////////////////////////////////////////////////

HRESULT SetStringAttribute(IADs* pIADs, LPCWSTR lpszAttr, LPCWSTR lpszVal, BOOL bPostCommit)
{
  ASSERT(pIADs != NULL);
  ASSERT(lpszAttr != NULL);

  TRACE(_T("SetStringAttribute() on bPostCommit = %s\n"), bPostCommit ? L"TRUE" : L"FALSE");
	TRACE(_T("\t'%s' = <%s>\n"), lpszAttr, lpszVal);

  HRESULT hr = S_OK;
  BSTR bstrAttr = ::SysAllocString(lpszAttr);
  VARIANT varVal;
  ::VariantInit(&varVal);
  varVal.vt = VT_BSTR;
  varVal.bstrVal = ::SysAllocString(lpszVal);

  if ( (lpszVal == NULL) || ((lpszVal != NULL) && (lpszVal[0] == NULL)) )
  {
    hr = _ClearAttribute(pIADs, bstrAttr, bPostCommit);
  }
  else
  {
    // set the value
    hr = pIADs->Put(bstrAttr, varVal);
  }  
  ::SysFreeString(bstrAttr);
  ::VariantClear(&varVal);
  return hr;
}

HRESULT GetStringAttribute(IADs* pIADs, LPCWSTR lpszAttr, CString& szVal)
{
  ASSERT(pIADs != NULL);
  VARIANT varVal;
  ::VariantInit(&varVal);

  BSTR bstrAttr = ::SysAllocString(lpszAttr);
  
  szVal.Empty();
  HRESULT hr = pIADs->Get(bstrAttr, &varVal);
  if (SUCCEEDED(hr))
  {
    if (varVal.vt == VT_BSTR)
      szVal = varVal.bstrVal;
    else
      hr = E_INVALIDARG;
  }
  VariantClear(&varVal);
  ::SysFreeString(bstrAttr);
  
  TRACE(_T("GetStringAttribute()lpszAttr = <%s>, (LPCWSTR)szVal = <%s>\n"), lpszAttr, szVal);
  return hr;
}



HRESULT SetByteArrayAsOctetString(IADs* pIADs, LPCWSTR lpszAttr, BYTE* pByteArray, DWORD dwSize)
{
  if (pByteArray == NULL)
  {
    CComBSTR bstrAttr = lpszAttr;
    return _ClearAttribute(pIADs, bstrAttr, FALSE);
  }

  HRESULT hr = S_OK;
  SAFEARRAYBOUND rgsabound[1];
  rgsabound[0].lLbound = 0;
  rgsabound[0].cElements = dwSize;

  SAFEARRAY* psa = SafeArrayCreate(VT_UI1, 1, rgsabound);
  if (psa == NULL)
  {
    return E_OUTOFMEMORY;
  }

  CComVariant var;
  V_VT(&var) = VT_UI1|VT_ARRAY;
  V_ARRAY(&var) = psa;

  for (long idx = 0; (DWORD)idx < dwSize; idx++)
  {
    hr = SafeArrayPutElement(psa, &idx, (void*)&pByteArray[idx]);
    if (FAILED(hr))
    {
      return hr;
    }
  }

  hr = pIADs->Put((LPWSTR)lpszAttr, var);
  //SafeArrayDestroy(psa);

  return hr;
}

BOOL LoadFileAsByteArray(PWSTR pszPath, LPBYTE* ppByteArray, DWORD* pdwSize)
{
  CFile file;
  if (!file.Open(pszPath, CFile::modeRead | CFile::shareDenyNone | CFile::typeBinary))
  {
    AfxMessageBox(_T("Failed to open file."));
    return FALSE;
  }

  *pdwSize = file.GetLength();
  *ppByteArray = new BYTE[*pdwSize];
  ASSERT(*ppByteArray != NULL);

  UINT uiCount = file.Read(*ppByteArray, *pdwSize);
  ASSERT(uiCount == *pdwSize);

  return TRUE;
}



