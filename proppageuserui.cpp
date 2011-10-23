//+-------------------------------------------------------------------------
//
//  Microsoft Windows
//
//  Copyright (C) Microsoft Corporation, 1999 - 2000
//
//  File:       proppageuserui.cpp
//
//--------------------------------------------------------------------------

#include "stdafx.h"
#include "ProppageUserUI.h"
#include "util.h"
#include "adsprop.h"

///////////////////////////////////////////////////////////////






///////////////////////////////////////////////////////////////

BEGIN_MESSAGE_MAP(CProppageUserUI, CPropertyPage)
  ON_WM_DESTROY()
  ON_BN_CLICKED(IDC_BUTTON_NEW, &CProppageUserUI::OnBnClickedButtonNew)
  ON_LBN_DBLCLK(IDC_LIST_PROXY_ADDRESSES, &CProppageUserUI::OnLbnDblclkListProxyAddresses)
  ON_BN_CLICKED(IDC_BUTTON_EDIT, &CProppageUserUI::OnBnClickedButtonEdit)
  ON_BN_CLICKED(IDC_BUTTON_REMOVE, &CProppageUserUI::OnBnClickedButtonRemove)
  ON_LBN_SELCHANGE(IDC_LIST_PROXY_ADDRESSES, &CProppageUserUI::OnLbnSelchangeListProxyAddresses)
END_MESSAGE_MAP()


BOOL CProppageUserUI::OnInitDialog()
{
  //
  // Verify that this page is only coming up for users
  //
  ASSERT(m_szClass == _T("user"));

  BOOL bRes = CPropertyPage::OnInitDialog();

  //
  // Display data in UI
  //
  SetUIData();
  return bRes;
}

HRESULT CProppageUserUI::Init(PWSTR pszPath, PWSTR pszClass)
{
  //
  // Store local versions of the path and class
  //
  m_szPath = pszPath;
  m_szClass = pszClass;

  // 
  // Get the property page initialization parameters from the notify window
  //
  ADSPROPINITPARAMS InitParams = {0};
  InitParams.dwSize = sizeof(ADSPROPINITPARAMS);
  if (!ADsPropGetInitInfo(m_hNotifyObj, &InitParams))
  {
    return E_FAIL;
  }

  if (FAILED(InitParams.hr))
  {
    return InitParams.hr;
  }

  //
  // Store local versions of the IDirectoryObject interface pointer
  // and the writable attributes
  //
  m_spDsObj = InitParams.pDsObj;
  m_pWritableAttrs = InitParams.pWritableAttrs;

  return S_OK;
}

void CProppageUserUI::SetUIData()
{
  //
  // Prime the controls with the data from the object
  //
  SetUIProxyAddresses();
}

void CProppageUserUI::SetUIProxyAddresses()
{
  LPWSTR lpszAttrArray[] = { _T("proxyAddresses") };
  PADS_ATTR_INFO pAttrInfo;
  DWORD dwReturned = 0;

  //
  // Get the proxy Addresses
  //
  HRESULT hr = m_spDsObj->GetObjectAttributes(lpszAttrArray, 1, &pAttrInfo, &dwReturned);
  if (FAILED(hr))
  {
    CString szMsg;
    szMsg.Format(_T("Failed to read proxyAddresses attribute, with hr = 0x%x\n"), hr);
    AfxMessageBox(szMsg);
    return;
  }


  //
  // Retrieve the employee ID from the attribute 
  // information and set it in the edit box
  //
  if (dwReturned == 1)
  {
    ASSERT(pAttrInfo != NULL);

    //
    // write data to the edit control

	for (DWORD i = 0; i < pAttrInfo->dwNumValues; i++) {
		GetProxyAddresses()->AddString(pAttrInfo->pADsValues[i].CaseIgnoreString);
	}
	
    //
    // Free the attribute memory
    //
    FreeADsMem(pAttrInfo);
  }

  //
  // Disable the ID edit box if user does not have write permissions on the attribute
  //
  if (!ADsPropCheckIfWritable(_T("proxyAddresses"), m_pWritableAttrs))
  {
    GetProxyAddresses()->EnableWindow(FALSE);
	GetNewButton()->EnableWindow(FALSE);
	GetEditButton()->EnableWindow(FALSE);
	GetRemoveButton()->EnableWindow(FALSE);
  }

}

BOOL CProppageUserUI::OnApply()
{
  ASSERT(m_spDsObj != NULL);

  //
  // update the two attributes
  //
  HRESULT hr = WriteProxyAddresses();
  if (FAILED(hr))
  {
    CString szMsg;
    szMsg.Format(_T("Failed to write proxyAddresses, with hr = 0x%x\n"), hr);
    AfxMessageBox(szMsg);
    return FALSE;
  }


  // Signal the change notification. Note that the notify-apply 
  // message must be sent even if the page is not dirty so that 
  // the notify ref-counting will properly decrement.
  ::SendMessage(m_hNotifyObj, WM_ADSPROP_NOTIFY_APPLY, TRUE, 0);

  return TRUE;
}

void CProppageUserUI::OnDestroy()
{
  if (::IsWindow(m_hNotifyObj))
  {
    ::SendMessage(m_hNotifyObj, WM_ADSPROP_NOTIFY_EXIT, 0, 0);
  }
  CPropertyPage::OnDestroy();
  delete this;
}



HRESULT CProppageUserUI::WriteProxyAddresses()
{
  
  LPWSTR lpszAttr = _T("proxyAddresses");
  LPWSTR szValue;
  DWORD nCount = GetProxyAddresses()->GetCount();
  DWORD nLength = 0;
  ADS_ATTR_INFO attrInfo[1];
  ADSVALUE* values = new ADSVALUE[nCount];
  attrInfo[0].dwADsType = ADSTYPE_CASE_IGNORE_STRING;
  attrInfo[0].pszAttrName = lpszAttr;
  
  for (DWORD i = 0; i < nCount; i++) {
	szValue = new TCHAR[GetProxyAddresses()->GetTextLen(i)];
	nLength = GetProxyAddresses()->GetText(i, szValue);
	values[i].dwType = ADSTYPE_CASE_IGNORE_STRING;
	values[i].CaseIgnoreString = nLength != LB_ERR ? szValue : NULL;
  }
  
  attrInfo[0].pADsValues = values;
  attrInfo[0].dwNumValues = nCount;
  attrInfo[0].dwControlCode = nCount > 0 ? ADS_ATTR_UPDATE : ADS_ATTR_CLEAR;  
  
  //
  // Write the changes to the object
  //
  DWORD dwModified = 0;

  return m_spDsObj->SetObjectAttributes(attrInfo, 1, &dwModified);
}

LPWSTR szEmailAddress;

BOOL CALLBACK DlgProc(HWND hWndDlg, UINT Msg, WPARAM wParam, LPARAM lParam)
{
	DWORD nLength;
	HWND emailAddress = GetDlgItem(hWndDlg, IDC_EDIT_EMAIL_ADDRESS);
	switch(Msg)
	{
		case WM_INITDIALOG:
			if (lParam) {
				SetWindowText(emailAddress, (LPWSTR)lParam);
			}
			return TRUE;

		case WM_COMMAND:
		
			switch(wParam)
			{
				case IDC_BUTTON_OK:
					
					nLength = GetWindowTextLength(emailAddress) + 1;
					szEmailAddress = new TCHAR[nLength];
					if (!GetWindowText(emailAddress, szEmailAddress,  nLength)) {
						//szEmailAddress = _T("");
						CString szMsg = _T("Email address must not be blank.\n");
						AfxMessageBox(szMsg);
						return FALSE;
					}
					EndDialog(hWndDlg, IDC_BUTTON_OK);
				break;
			
				case IDC_BUTTON_CANCEL:
					EndDialog(hWndDlg, IDC_BUTTON_CANCEL);
				break;
			}
		break;
	
		default:
			return FALSE;
	}
	return TRUE;
}

void CProppageUserUI::EditEmailAddress() {
	DWORD nIndex = GetProxyAddresses()->GetCurSel();
	DWORD nCount = GetProxyAddresses()->GetCount();
	LPWSTR szValue = new TCHAR[GetProxyAddresses()->GetTextLen(nIndex)];
	DWORD nLength = GetProxyAddresses()->GetText(nIndex, szValue);
	if (nIndex != LB_ERR && nCount >= 1 && nLength != LB_ERR && DialogBoxParam(AfxGetInstanceHandle(), MAKEINTRESOURCE(IDD_USER_EMAIL_ADDRESS), m_hWnd, (DLGPROC)DlgProc, (LPARAM)szValue) == IDC_BUTTON_OK) {
      if (GetProxyAddresses()->FindStringExact(-1, szEmailAddress) != LB_ERR) {
		  return;
	  }
	  GetProxyAddresses()->DeleteString(nIndex);
	  GetProxyAddresses()->AddString(szEmailAddress);
	  SetModified(TRUE);
    }
}

void CProppageUserUI::OnBnClickedButtonNew()
{
	if (DialogBoxParam(AfxGetInstanceHandle(), MAKEINTRESOURCE(IDD_USER_EMAIL_ADDRESS), m_hWnd, (DLGPROC)DlgProc, (LPARAM)_T("")) == IDC_BUTTON_OK) {
      if (GetProxyAddresses()->FindStringExact(-1, szEmailAddress) != LB_ERR) {
		  return;
	  }
	  GetProxyAddresses()->AddString(szEmailAddress);
	  SetModified(TRUE);
    }
}


void CProppageUserUI::OnBnClickedButtonEdit()
{
	EditEmailAddress();
}


void CProppageUserUI::OnBnClickedButtonRemove()
{
	DWORD nIndex = GetProxyAddresses()->GetCurSel();
	GetProxyAddresses()->DeleteString(nIndex);
	DWORD nCount = GetProxyAddresses()->GetCount();
	if (nCount > 0) {
		GetProxyAddresses()->SetCurSel(min(nIndex, nCount - 1));
	} else {
		GetEditButton()->EnableWindow(FALSE);
		GetRemoveButton()->EnableWindow(FALSE);
	}
	SetModified(TRUE);
}

void CProppageUserUI::OnLbnDblclkListProxyAddresses()
{
	
	EditEmailAddress();
}


void CProppageUserUI::OnLbnSelchangeListProxyAddresses()
{
	GetEditButton()->EnableWindow(TRUE);
	GetRemoveButton()->EnableWindow(TRUE);
	// TODO: Add your control notification handler code here
}
