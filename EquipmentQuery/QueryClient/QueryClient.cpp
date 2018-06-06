
// QueryClient.cpp : ����Ӧ�ó��������Ϊ��
//

#include "QueryClient.h"
#include "QueryClientDlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

//using namespace ADODB;
CQueryClientDlg* g_pDlgTesk = NULL;

// CQueryClientApp
BEGIN_MESSAGE_MAP(CQueryClientApp, CWinApp)
	ON_COMMAND(ID_HELP, &CWinApp::OnHelp)
END_MESSAGE_MAP()

// CQueryClientApp ����
CQueryClientApp::CQueryClientApp()
{
	// TODO: �ڴ˴���ӹ�����룬
	// ��������Ҫ�ĳ�ʼ�������� InitInstance ��
}

CQueryClientApp::~CQueryClientApp()
{
	CloseADOConnect();
}

// Ψһ��һ�� CQueryClientApp ����
CQueryClientApp theApp;

////////////////////////////////////////////////////////
// CQueryClientApp ��ʼ��
// 2018-05-25	1.00.01		Lichun
////////////////////////////////////////////////////////
BOOL CQueryClientApp::InitInstance()
{
	CWinApp::InitInstance();

//	AfxOleInit();																			//QUERY_CL_1805250
	// ���� shell ���������Է��Ի������
	// �κ� shell ����ͼ�ؼ��� shell �б���ͼ�ؼ���
	CShellManager *pShellManager = new CShellManager;

	// ��׼��ʼ��
	// ���δʹ����Щ���ܲ�ϣ����С
	// ���տ�ִ���ļ��Ĵ�С����Ӧ�Ƴ�����
	// ����Ҫ���ض���ʼ������
	// �������ڴ洢���õ�ע�����
	// TODO: Ӧ�ʵ��޸ĸ��ַ�����
	// �����޸�Ϊ��˾����֯��
	SetRegistryKey(_T("Ӧ�ó��������ɵı���Ӧ�ó���"));

	//Connect to DB
//	InitADOConnect( _T("") );
	//Get current file path
	if ( 0 != GetCurrentDirectory( sizeof(m_cDatFilePath), m_cDatFilePath ) )
	{	//Get current director
		strcat( m_cDatFilePath, "\\" );
		strcat( m_cDatFilePath, INFO_FILE );
	}
	else
	{
		memset( m_cDatFilePath, 0x00, sizeof(m_cDatFilePath) );
	}
	
	//Dialog
	g_pDlgTesk = new(CQueryClientDlg);
	m_pMainWnd = g_pDlgTesk;
	INT_PTR nResponse = g_pDlgTesk->DoModal();
	if (nResponse == IDOK)
	{
		// TODO: �ڴ˷��ô����ʱ��
		//  ��ȷ�������رնԻ���Ĵ���
	}
	else if (nResponse == IDCANCEL)
	{
		// TODO: �ڴ˷��ô����ʱ��
		//  ��ȡ�������رնԻ���Ĵ���
	}

	// ɾ�����洴���� shell ��������
	if (pShellManager != NULL)
	{
		delete pShellManager;
	}

	// ���ڶԻ����ѹرգ����Խ����� FALSE �Ա��˳�Ӧ�ó���
	//  ����������Ӧ�ó������Ϣ�á�
	return FALSE;
}

////////////////////////////////////////////////////////
// InitConnect �������ݿ� - �û�������
// 2018-05-25	1.00.01		Lichun
////////////////////////////////////////////////////////
BOOL CQueryClientApp::InitADOConnect( CString csServer, CString csDataBase, CString csUser, CString csPassword )
{
	HRESULT	hRet = S_OK;
	CString csName;

	::CoInitialize(NULL);
	//Get connection string
	//E.G: Provider=SQLOLEDB.1;Password=840909;Persist Security Info=True;User ID=sa;Initial Catalog=ArcheryEquipments;Data Source=LICHUN-9E4430A0
	csName.Format( _T("Provider=SQLOLEDB.1;Password=%s;Integrated Security=SSPI;Persist Security Info=True;User ID=%s;Initial Catalog=%s;Data Source=%s"),
				csPassword, csUser, csDataBase, csServer );
	
	//Connect
	try
	{	//
		m_pConnection.CreateInstance( _T("ADODB.Connection") );
		_bstr_t strConnect = csName;
		//Open data base
		hRet = m_pConnection->Open( strConnect, _T(""), _T(""), adModeUnknown );
	}
	catch(_com_error err)
	{
		AfxMessageBox(err.Description());
		return FALSE;
	}

	if ( hRet != S_OK )
	{
		return FALSE;
	}

	return TRUE;
}

////////////////////////////////////////////////////////
// InitConnect �������ݿ� - ������Ϣ
// 2018-05-25	1.00.01		Lichun
////////////////////////////////////////////////////////
BOOL CQueryClientApp::InitADOConnect( CString csConnectStr )
{
	HRESULT	hRet = S_OK;

	::CoInitialize(NULL);
	
	//Connect
	try
	{	//
		m_pConnection.CreateInstance( _T("ADODB.Connection") );
		_bstr_t strConnect = csConnectStr;
		//Open data base
		hRet = m_pConnection->Open( strConnect, _T(""), _T(""), adModeUnknown );
	}
	catch(_com_error err)
	{
		AfxMessageBox(err.Description());
		return FALSE;
	}

	if ( hRet != S_OK )
	{
		return FALSE;
	}

	return TRUE;
}

////////////////////////////////////////////////////////
// InitConnect �Ͽ����ݿ�����
// 2018-05-25	1.00.01		Lichun
////////////////////////////////////////////////////////
BOOL CQueryClientApp::CloseADOConnect()
{
	if ( g_pDlgTesk != NULL )
	{
		delete [] g_pDlgTesk;
	}
	//
	if ( m_pConnection != NULL )
	{
		m_pConnection->Close();
	}
	//Uninit COM
	::CoUninitialize();

	return TRUE;
}

////////////////////////////////////////////////////////
// SelectAllItems ��ʾ��������
// 2018-05-28	1.00.01		Lichun
////////////////////////////////////////////////////////
BOOL CQueryClientApp::SelectAllItems( CListCtrl& clItemsList )
{
	_bstr_t bstrSQL = _T("SELECT * FROM Table_Main ORDER BY ITEM_TYPE");

	if ( g_pDlgTesk == NULL )
	{
		return FALSE;
	}

	clItemsList.SetExtendedStyle( LVS_EX_FLATSB
								|LVS_EX_FULLROWSELECT
								|LVS_EX_HEADERDRAGDROP
								|LVS_EX_ONECLICKACTIVATE
								|LVS_EX_GRIDLINES );
	clItemsList.InsertColumn( 0, _T("���"), LVCFMT_LEFT, 100, 0);
	clItemsList.InsertColumn( 1, _T("���"), LVCFMT_LEFT, 100, 1);
	clItemsList.InsertColumn( 2, _T("����"), LVCFMT_LEFT, 100, 2);
	clItemsList.InsertColumn( 3, _T("Ʒ��"), LVCFMT_LEFT, 100, 3);
	clItemsList.InsertColumn( 4, _T("����"), LVCFMT_LEFT, 200, 4);
	clItemsList.InsertColumn( 5, _T("��ɫ"), LVCFMT_LEFT, 100, 5);
	clItemsList.InsertColumn( 6, _T("�ͺ�"), LVCFMT_LEFT, 100, 6);
	clItemsList.InsertColumn( 7, _T("����"), LVCFMT_LEFT, 100, 7);
	clItemsList.InsertColumn( 8, _T("���"), LVCFMT_LEFT, 100, 8);

	//
	m_pRecordset.CreateInstance( __uuidof(Recordset) );
	m_pRecordset->Open( bstrSQL, m_pConnection.GetInterfacePtr(), adOpenDynamic, adLockOptimistic, adCmdText );
	//
	while( !m_pRecordset->EndOfFile )
	{
		clItemsList.InsertItem( 0, "" );
		clItemsList.SetItemText( 0, 0, (char*)(_bstr_t)m_pRecordset->GetCollect("ITEM_ID"));
		clItemsList.SetItemText( 0, 1, (char*)(_bstr_t)m_pRecordset->GetCollect("ITEM_TYPE"));
		clItemsList.SetItemText( 0, 2, (char*)(_bstr_t)m_pRecordset->GetCollect("ITEM_SUBTYPE"));
		clItemsList.SetItemText( 0, 3, (char*)(_bstr_t)m_pRecordset->GetCollect("ITEM_BRAND"));
		clItemsList.SetItemText( 0, 4, (char*)(_bstr_t)m_pRecordset->GetCollect("ITEM_NAME"));
		clItemsList.SetItemText( 0, 5, (char*)(_bstr_t)m_pRecordset->GetCollect("ITEM_COLOR"));
		clItemsList.SetItemText( 0, 6, (char*)(_bstr_t)m_pRecordset->GetCollect("ITEM_SIZE"));
		clItemsList.SetItemText( 0, 7, (char*)(_bstr_t)m_pRecordset->GetCollect("ITEM_HAND"));
		clItemsList.SetItemText( 0, 8, (char*)(_bstr_t)m_pRecordset->GetCollect("ITEM_COUNT"));
		m_pRecordset->MoveNext();
	}

	if( m_pRecordset!=NULL )
	{
		m_pRecordset->Close();
	}

	return TRUE;
}

BOOL CQueryClientApp::SelectKeyValue( CString csKeyString, CListCtrl& clItemsList )
{
	CString csSqlCmd = _T("SELECT * FROM Table_Main ");
	csSqlCmd.Append( _T("WHERE ITEM_TYPE like N'%") );
	csSqlCmd.Append( csKeyString );
	csSqlCmd.Append( _T("%' ") );
	csSqlCmd.Append( _T("OR ITEM_BRAND like N'%") );
	csSqlCmd.Append( csKeyString );
	csSqlCmd.Append( _T("%' ") );
	csSqlCmd.Append( _T("OR ITEM_NAME like N'%") );
	csSqlCmd.Append( csKeyString );
	csSqlCmd.Append( _T("%'") );

	_bstr_t bstrSQL = csSqlCmd;

	if ( g_pDlgTesk == NULL )
	{
		return FALSE;
	}

	clItemsList.DeleteAllItems();
	clItemsList.SetExtendedStyle( LVS_EX_FLATSB
								|LVS_EX_FULLROWSELECT
								|LVS_EX_HEADERDRAGDROP
								|LVS_EX_ONECLICKACTIVATE
								|LVS_EX_GRIDLINES );
	clItemsList.InsertColumn( 0, _T("���"), LVCFMT_LEFT, 100, 0);
	clItemsList.InsertColumn( 1, _T("���"), LVCFMT_LEFT, 100, 1);
	clItemsList.InsertColumn( 2, _T("����"), LVCFMT_LEFT, 100, 2);
	clItemsList.InsertColumn( 3, _T("Ʒ��"), LVCFMT_LEFT, 100, 3);
	clItemsList.InsertColumn( 4, _T("����"), LVCFMT_LEFT, 200, 4);
	clItemsList.InsertColumn( 5, _T("��ɫ"), LVCFMT_LEFT, 100, 5);
	clItemsList.InsertColumn( 6, _T("�ͺ�"), LVCFMT_LEFT, 100, 6);
	clItemsList.InsertColumn( 7, _T("����"), LVCFMT_LEFT, 100, 7);
	clItemsList.InsertColumn( 8, _T("���"), LVCFMT_LEFT, 100, 8);

	//
	m_pRecordset.CreateInstance( __uuidof(Recordset) );
	m_pRecordset->Open( bstrSQL, m_pConnection.GetInterfacePtr(), adOpenDynamic, adLockOptimistic, adCmdText );
	//
	while( !m_pRecordset->EndOfFile )
	{
		clItemsList.InsertItem( 0, "" );
		clItemsList.SetItemText( 0, 0, (CHAR*)((_bstr_t)m_pRecordset->GetCollect("ITEM_ID")) );
		clItemsList.SetItemText( 0, 1, (CHAR*)((_bstr_t)m_pRecordset->GetCollect("ITEM_TYPE")) );
		clItemsList.SetItemText( 0, 2, (CHAR*)((_bstr_t)m_pRecordset->GetCollect("ITEM_SUBTYPE")) );
		clItemsList.SetItemText( 0, 3, (CHAR*)((_bstr_t)m_pRecordset->GetCollect("ITEM_BRAND")) );
		clItemsList.SetItemText( 0, 4, (CHAR*)((_bstr_t)m_pRecordset->GetCollect("ITEM_NAME")) );
		clItemsList.SetItemText( 0, 5, (CHAR*)((_bstr_t)m_pRecordset->GetCollect("ITEM_COLOR")) );
		clItemsList.SetItemText( 0, 6, (CHAR*)((_bstr_t)m_pRecordset->GetCollect("ITEM_SIZE")) );
		clItemsList.SetItemText( 0, 7, (CHAR*)((_bstr_t)m_pRecordset->GetCollect("ITEM_HAND")) );
		clItemsList.SetItemText( 0, 8, (CHAR*)((_bstr_t)m_pRecordset->GetCollect("ITEM_COUNT")) );
		m_pRecordset->MoveNext();
	}

	if( m_pRecordset!=NULL )
	{
		m_pRecordset->Close();
	}

	return TRUE;
}

////////////////////////////////////////////////////////
// RecordInfo
// 2018-05-29	1.00.01		Lichun
////////////////////////////////////////////////////////
BOOL CQueryClientApp::RecordInfo( CString csServer, CString csDataBase, CString csUser, CString csPassword )
{
	if ( 0 == strlen(m_cDatFilePath) )
	{
		return FALSE;
	}

	//Write string to dat file
	WritePrivateProfileString( "USER_MODE", "Server", csServer, INFO_FILE );
	WritePrivateProfileString( "USER_MODE", "DataBas", csDataBase, INFO_FILE );
	WritePrivateProfileString( "USER_MODE", "User", csUser, INFO_FILE );
	WritePrivateProfileString( "USER_MODE", "Password", csPassword, INFO_FILE );

	return TRUE;
}

////////////////////////////////////////////////////////
// RecordInfo
// 2018-05-29	1.00.01		Lichun
////////////////////////////////////////////////////////
BOOL CQueryClientApp::RecordInfo( CString csConnectStr )
{
	if ( 0 == strlen(m_cDatFilePath) )
	{
		return FALSE;
	}

	//Write string to dat file
	WritePrivateProfileString( "INFO_MODE", "ConnectInfo", csConnectStr, m_cDatFilePath );
	
	return TRUE;
}

////////////////////////////////////////////////////////
// RecordInfo
// 2018-05-29	1.00.01		Lichun
////////////////////////////////////////////////////////
BOOL CQueryClientApp::ReadInfo( CString& csServer, CString& csDataBase, CString& csUser, CString& csPassword )
{
	CHAR cReadInfo[MAX_PATH];

	if ( 0 == strlen(m_cDatFilePath) )
	{
		return FALSE;
	}

	//Get string from dat file
	GetPrivateProfileString( "USER_MODE", "Server", "", cReadInfo, sizeof(cReadInfo), m_cDatFilePath );
	csServer.Empty();
	csServer.Append( cReadInfo );

	GetPrivateProfileString( "USER_MODE", "DataBas", "", cReadInfo, sizeof(cReadInfo), m_cDatFilePath );
	csDataBase.Empty();
	csDataBase.Append( cReadInfo );

	GetPrivateProfileString( "USER_MODE", "User", "", cReadInfo, sizeof(cReadInfo), m_cDatFilePath );
	csUser.Empty();
	csUser.Append( cReadInfo );

	GetPrivateProfileString( "USER_MODE", "Password", "", cReadInfo, sizeof(cReadInfo), m_cDatFilePath );
	csPassword.Empty();
	csPassword.Append( cReadInfo );
	
	return TRUE;
}

////////////////////////////////////////////////////////
// RecordInfo
// 2018-05-29	1.00.01		Lichun
////////////////////////////////////////////////////////
BOOL CQueryClientApp::ReadInfo( CString& csConnectStr )
{
	CHAR cReadInfo[1024];

	if ( 0 == strlen(m_cDatFilePath) )
	{
		return FALSE;
	}

	//Get string from dat file
	GetPrivateProfileString( "INFO_MODE", "ConnectInfo", "", cReadInfo, sizeof(cReadInfo), m_cDatFilePath );
	csConnectStr.Empty();
	csConnectStr.Append( cReadInfo );

	return TRUE;
}
