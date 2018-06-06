
// QueryClientDlg.h : 头文件
//

#pragma once
#include "afxcmn.h"

// CQueryClientDlg 对话框
class CQueryClientDlg : public CDialogEx
{
// 构造
public:
	CQueryClientDlg(CWnd* pParent = NULL);	// 标准构造函数

// 对话框数据
	enum { IDD = IDD_QUERYCLIENT_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV 支持
	virtual void OnOk();
	virtual void OnCancel();
	virtual BOOL PreTranslateMessage(MSG* pMsg);
// 实现
protected:
	HICON m_hIcon;

	// 生成的消息映射函数
	virtual BOOL OnInitDialog();
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	afx_msg void OnBnClickedCancel();
	afx_msg void OnBnClickedRadioStr();
	afx_msg void OnBnClickedRadioUser();
	afx_msg void OnBnClickedBtnConnect();
	afx_msg void OnBnClickedButtonSearch();
	afx_msg void OnBnDropDownRadioUser( NMHDR *pNMHDR, LRESULT *pResult );
	afx_msg void OnBnDropDownRadioStr( NMHDR *pNMHDR, LRESULT *pResult );
	afx_msg void OnLvnColumnclickListItems( NMHDR *pNMHDR, LRESULT *pResult );
	DECLARE_MESSAGE_MAP()
protected:
	static INT CALLBACK CompareFunc( LPARAM lParam1, LPARAM lParam2, LPARAM lParamSort );
public:

public:
	INT		m_iConnectType;
	CString m_csServer;
	CString m_csUser;
	CString m_csPassword;
	CString m_csDataBase;
	CString m_csConnectInfo;
	CString m_csKeyString;
	CListCtrl m_clItemsList;
public:
	static	BOOL	m_bASC;										//是否升序
	static	DWORD	m_dwSelColID;								//选择的列
};
