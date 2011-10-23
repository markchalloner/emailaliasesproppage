//+-------------------------------------------------------------------------
//
//  Microsoft Windows
//
//  Copyright (C) Microsoft Corporation, 1999 - 2000
//
//  File:       proppageuserui.h
//
//--------------------------------------------------------------------------

#ifndef _PROPPAGEUSERUI_H
#define _PROPPAGEUSERUI_H

#include "resource.h"
#include "util.h"




///////////////////////////////////////////////////////////////

class CProppageUserUI : public CPropertyPage
{
public:
  CProppageUserUI(HWND hNotifyObj) 
    : CPropertyPage(IDD_USER_PAGE_EXT),
      m_hNotifyObj(hNotifyObj)
  {
  }

  ~CProppageUserUI()
  {
  }

  afx_msg void OnNewEmail();
  afx_msg void OnEditEmail();
  afx_msg void OnRemoveEmail();

  virtual BOOL OnInitDialog();
  virtual BOOL OnApply();
  virtual void OnDestroy();


  HRESULT Init(PWSTR pszPath, PWSTR pszClass);

private:
  void SetUIData();
  void SetUIProxyAddresses();
  void EditEmailAddress();

  HRESULT WriteProxyAddresses();


  // Control accessors
  CButton* GetNewButton() { return (CButton*)GetDlgItem(IDC_BUTTON_NEW); }
  CButton* GetEditButton() { return (CButton*)GetDlgItem(IDC_BUTTON_EDIT); }
  CButton* GetRemoveButton() { return (CButton*)GetDlgItem(IDC_BUTTON_REMOVE); }
  CListBox* GetProxyAddresses() { return (CListBox*)GetDlgItem(IDC_LIST_PROXY_ADDRESSES); }

private:
  HWND    m_hNotifyObj;
  CString                   m_szPath;
  CString                   m_szClass;

  CComPtr<IDirectoryObject> m_spDsObj;
  PADS_ATTR_INFO            m_pWritableAttrs;

  DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnBnClickedButtonNew();
	afx_msg void OnLbnDblclkListProxyAddresses();
	afx_msg void OnBnClickedButtonEdit();
	afx_msg void OnBnClickedButtonRemove();
	afx_msg void OnLbnSelchangeListProxyAddresses();
};

#endif // _PROPPAGEUSERUI_H