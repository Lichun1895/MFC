
// QueryClient.h : PROJECT_NAME 应用程序的主头文件
//

#pragma once

#include "stdafx.h"

#ifndef __AFXWIN_H__
	#error "在包含此文件之前包含“stdafx.h”以生成 PCH 文件"
#endif

#include "resource.h"		// 主符号

using namespace   ADO;

const CHAR INFO_FILE[] = _T("EquipmentQueryInfo.dat");

// CQueryClientApp:
// 有关此类的实现，请参阅 QueryClient.cpp
//
class CQueryClientApp : public CWinApp
{
public:
	CQueryClientApp();
	~CQueryClientApp();
// 重写
public:
	virtual BOOL InitInstance();

// 实现
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