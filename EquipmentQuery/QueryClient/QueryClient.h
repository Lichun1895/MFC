
// QueryClient.h : PROJECT_NAME Ӧ�ó������ͷ�ļ�
//

#pragma once

#include "stdafx.h"

#ifndef __AFXWIN_H__
	#error "�ڰ������ļ�֮ǰ������stdafx.h�������� PCH �ļ�"
#endif

#include "resource.h"		// ������

using namespace   ADO;

const CHAR INFO_FILE[] = _T("EquipmentQueryInfo.dat");

// CQueryClientApp:
// �йش����ʵ�֣������ QueryClient.cpp
//
class CQueryClientApp : public CWinApp
{
public:
	CQueryClientApp();
	~CQueryClientApp();
// ��д
public:
	virtual BOOL InitInstance();

// ʵ��
	DECLARE_MESSAGE_MAP()
public:
	BOOL InitADOConnect( CString csServer, CString csDataBase, CString csUser, CString csPassword );
	BOOL InitADOConnect( CString csConnectStr );
	BOOL CloseADOConnect();
	//
	BOOL RecordInfo( CString csServer, CString csDataBase, CString csUser, CString csPassword );
	BOOL RecordInfo( CString csConnectStr );
	//
	BOOL ReadInfo( CString& csServer, CString& csDataBase, CString& csUser, CString& csPassword );
	BOOL ReadInfo( CString& csConnectStr );
	//
	BOOL SelectAllItems( CListCtrl& clItemsList );
	BOOL SelectKeyValue( CString csKeyString, CListCtrl& clItemsList );
//
protected:
	_ConnectionPtr	m_pConnection;
	_RecordsetPtr	m_pRecordset;
	//File path
	CHAR			m_cDatFilePath[MAX_PATH*2];
};

extern CQueryClientApp theApp;