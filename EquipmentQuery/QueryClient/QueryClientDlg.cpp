
// QueryClientDlg.cpp : ʵ���ļ�
//

#include "stdafx.h"
#include "QueryClient.h"
#include "QueryClientDlg.h"
#include "afxdialogex.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

BOOL	CQueryClientDlg::m_bASC = FALSE;
DWORD	CQueryClientDlg::m_dwSelColID = 0;

// CQueryClientDlg �Ի���
CQueryClientDlg::CQueryClientDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(CQueryClientDlg::IDD, pParent)
	, m_csServer(_T(""))
	, m_csUser(_T(""))
	, m_csPassword(_T(""))
	, m_csDataBase(_T(""))
	, m_csConnectInfo(_T(""))
	, m_csKeyString(_T(""))
	, m_iConnectType(0)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CQueryClientDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Text(pDX, IDC_EDIT_SERVER, m_csServer);
	DDX_Text(pDX, IDC_EDIT_USER, m_csUser);
	DDX_Text(pDX, IDC_EDIT_PW, m_csPassword);
	DDX_Text(pDX, IDC_EDIT_DB, m_csDataBase);
	DDX_Text(pDX, IDC_EDIT_CONNECT_INFO, m_csConnectInfo);
	DDX_Text(pDX, IDC_EDIT_KEY_STR, m_csKeyString);
	DDX_Radio(pDX, IDC_RADIO_USER, m_iConnectType);
	DDX_Control(pDX, IDC_LIST_ITEMS, m_clItemsList);
}

void CQueryClientDlg::OnOk()
{
	OnBnClickedButtonSearch();
	return;
}

void CQueryClientDlg::OnCancel()
{
	return;
}

BOOL CQueryClientDlg::PreTranslateMessage(MSG* pMsg)
{
	if( (pMsg->message==WM_KEYDOWN)&&(pMsg->wParam==VK_ESCAPE) )		// ����esc��  
	{     
		return TRUE;// �����κβ���  
	}  
	if ( (pMsg->message==WM_KEYDOWN)&&(pMsg->wParam==VK_RETURN) )		// ����enter��  
	{  
		OnBnClickedButtonSearch();
		return TRUE;// �����κδ���  
	}

	return CDialog::PreTranslateMessage(pMsg);
}


//Message map
BEGIN_MESSAGE_MAP(CQueryClientDlg, CDialogEx)
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(IDC_BTN_CONNECT, &CQueryClientDlg::OnBnClickedBtnConnect)
	ON_BN_CLICKED(IDC_BUTTON_SEARCH, &CQueryClientDlg::OnBnClickedButtonSearch)
	ON_BN_CLICKED(IDC_RADIO_STR, &CQueryClientDlg::OnBnClickedRadioStr)
	ON_BN_CLICKED(IDC_RADIO_USER, &CQueryClientDlg::OnBnClickedRadioUser)
	ON_BN_CLICKED(IDCANCEL, &CQueryClientDlg::OnBnClickedCancel)
	ON_NOTIFY(BCN_DROPDOWN, IDC_RADIO_USER, &CQueryClientDlg::OnBnDropDownRadioUser)
	ON_NOTIFY(BCN_DROPDOWN, IDC_RADIO_STR, &CQueryClientDlg::OnBnDropDownRadioStr)
	ON_NOTIFY(LVN_COLUMNCLICK, IDC_LIST_ITEMS, &CQueryClientDlg::OnLvnColumnclickListItems)
END_MESSAGE_MAP()


// CQueryClientDlg ��Ϣ�������

BOOL CQueryClientDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// ���ô˶Ի����ͼ�ꡣ��Ӧ�ó��������ڲ��ǶԻ���ʱ����ܽ��Զ�
	//  ִ�д˲���
	SetIcon(m_hIcon, TRUE);			// ���ô�ͼ��
	SetIcon(m_hIcon, FALSE);		// ����Сͼ��

	ShowWindow(SW_SHOWNORMAL);

	// TODO: �ڴ���Ӷ���ĳ�ʼ������
	theApp.ReadInfo( m_csConnectInfo );
	theApp.ReadInfo( m_csServer, m_csDataBase, m_csUser, m_csPassword );

	UpdateData(FALSE);
	return TRUE;  // ���ǽ��������õ��ؼ������򷵻� TRUE
}

// �����Ի��������С����ť������Ҫ����Ĵ���
//  �����Ƹ�ͼ�ꡣ����ʹ���ĵ�/��ͼģ�͵� MFC Ӧ�ó���
//  �⽫�ɿ���Զ���ɡ�

void CQueryClientDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // ���ڻ��Ƶ��豸������

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// ʹͼ���ڹ����������о���
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// ����ͼ��
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

//���û��϶���С������ʱϵͳ���ô˺���ȡ�ù��
//��ʾ��
HCURSOR CQueryClientDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

void CQueryClientDlg::OnBnClickedBtnConnect()
{
	// TODO: �ڴ���ӿؼ�֪ͨ����������
	UpdateData(TRUE);

	if ( m_iConnectType == 0 )		//Connect by user name
	{
		if ( m_csServer.IsEmpty() )
		{
			MessageBox( _T("���������������"), _T("Error"), MB_OK );
			return;
		}
		if ( m_csDataBase.IsEmpty() )
		{
			MessageBox( _T("���������ݿ�����"), _T("Error"), MB_OK );
			return;
		}
		if ( m_csUser.IsEmpty() )
		{
			MessageBox( _T("�������û���"), _T("Error"), MB_OK );
			return;
		}
		if ( m_csPassword.IsEmpty() )
		{
			MessageBox( _T("����������"), _T("Error"), MB_OK );
			return;
		}
		//
		if ( theApp.InitADOConnect( m_csServer, m_csDataBase, m_csUser, m_csPassword ) )
		{
			theApp.RecordInfo( m_csServer, m_csDataBase, m_csUser, m_csPassword );
			theApp.SelectAllItems(m_clItemsList);
			UpdateData(FALSE);
		}
	}
	else if ( m_iConnectType == 1 )
	{
		if ( m_csConnectInfo.IsEmpty() )
		{
			MessageBox( _T("������������Ϣ"), _T("Error"), MB_OK );
			return;
		}
		//
		if ( theApp.InitADOConnect( m_csConnectInfo ) )
		{
			theApp.RecordInfo( m_csConnectInfo );
			theApp.SelectAllItems(m_clItemsList);
			UpdateData(FALSE);
		}
	}
}


void CQueryClientDlg::OnBnClickedButtonSearch()
{
	// TODO: �ڴ���ӿؼ�֪ͨ����������
	UpdateData(TRUE);

	if ( m_csKeyString.IsEmpty() )
	{
		MessageBox( _T("������ؼ���"), _T("Error"), MB_OK );
		return;
	}
	theApp.SelectKeyValue( m_csKeyString, m_clItemsList );
}


void CQueryClientDlg::OnBnDropDownRadioUser(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMBCDROPDOWN pDropDown = reinterpret_cast<LPNMBCDROPDOWN>(pNMHDR);
	// TODO: �ڴ���ӿؼ�֪ͨ����������
}


void CQueryClientDlg::OnBnDropDownRadioStr(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMBCDROPDOWN pDropDown = reinterpret_cast<LPNMBCDROPDOWN>(pNMHDR);
	// TODO: �ڴ���ӿؼ�֪ͨ����������
}


void CQueryClientDlg::OnBnClickedRadioUser()
{
	// TODO: �ڴ���ӿؼ�֪ͨ����������
	UpdateData(TRUE);
	//Enable
	GetDlgItem(IDC_STATIC_USER01)->EnableWindow(TRUE);
	GetDlgItem(IDC_STATIC_USER02)->EnableWindow(TRUE);
	GetDlgItem(IDC_STATIC_USER03)->EnableWindow(TRUE);
	GetDlgItem(IDC_STATIC_USER04)->EnableWindow(TRUE);
	GetDlgItem(IDC_EDIT_SERVER)->EnableWindow(TRUE);
	GetDlgItem(IDC_EDIT_DB)->EnableWindow(TRUE);
	GetDlgItem(IDC_EDIT_USER)->EnableWindow(TRUE);
	GetDlgItem(IDC_EDIT_PW)->EnableWindow(TRUE);

	//Disable
	GetDlgItem(IDC_STATIC_CONNECT_INFO)->EnableWindow(FALSE);
	GetDlgItem(IDC_EDIT_CONNECT_INFO)->EnableWindow(FALSE);

	UpdateData(FALSE);
}


void CQueryClientDlg::OnBnClickedRadioStr()
{
	// TODO: �ڴ���ӿؼ�֪ͨ����������
	UpdateData(TRUE);
	//Enable
	GetDlgItem(IDC_STATIC_CONNECT_INFO)->EnableWindow(TRUE);
	GetDlgItem(IDC_EDIT_CONNECT_INFO)->EnableWindow(TRUE);

	//Disable
	GetDlgItem(IDC_STATIC_USER01)->EnableWindow(FALSE);
	GetDlgItem(IDC_STATIC_USER02)->EnableWindow(FALSE);
	GetDlgItem(IDC_STATIC_USER03)->EnableWindow(FALSE);
	GetDlgItem(IDC_STATIC_USER04)->EnableWindow(FALSE);
	GetDlgItem(IDC_EDIT_SERVER)->EnableWindow(FALSE);
	GetDlgItem(IDC_EDIT_DB)->EnableWindow(FALSE);
	GetDlgItem(IDC_EDIT_USER)->EnableWindow(FALSE);
	GetDlgItem(IDC_EDIT_PW)->EnableWindow(FALSE);

	UpdateData(FALSE);
}


void CQueryClientDlg::OnLvnColumnclickListItems(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMLISTVIEW pNMLV = reinterpret_cast<LPNMLISTVIEW>(pNMHDR);
	// TODO: �ڴ���ӿؼ�֪ͨ����������
	INT iCount = m_clItemsList.GetItemCount();

	//Set list item number
	for ( INT iPos=0;iPos<iCount;iPos++ )
	{
		m_clItemsList.SetItemData( iPos, iPos );
	}

	//Check selected column
	if( m_dwSelColID != pNMLV->iSubItem )
	{
		m_dwSelColID = pNMLV->iSubItem;					//New column number
		m_bASC = TRUE;									//First click
	}
	else
	{
		m_bASC = !m_bASC;								//Double click
	}

	//Sort
	m_clItemsList.SortItems( CompareFunc, (LPARAM)(&m_clItemsList) );

	UpdateData(FALSE);
	*pResult = 0;
}

INT CQueryClientDlg::CompareFunc( LPARAM lParam1, LPARAM lParam2, LPARAM lParamSort )
{
	INT			iPos1 = -1;
	INT			iPos2 = -1;
	INT			iRet = 0;
	CString		csItem1;
	CString		csItem2;
	LVFINDINFO	stFindInfo;
	CListCtrl*	pListCtrl = (CListCtrl*)lParamSort;

	//Get item contant
	csItem1 = pListCtrl->GetItemText( lParam1, m_dwSelColID );
	csItem2 = pListCtrl->GetItemText( lParam2, m_dwSelColID );

	//Compare
	if( m_bASC )//����
	{
		iRet = strcmp( csItem1, csItem2 );
	}
	else
	{
		iRet = strcmp( csItem2, csItem1 );
	}
	
	return iRet;
}

void CQueryClientDlg::OnBnClickedCancel()
{
	// TODO: �ڴ���ӿؼ�֪ͨ����������
	CDialogEx::OnCancel();
}
