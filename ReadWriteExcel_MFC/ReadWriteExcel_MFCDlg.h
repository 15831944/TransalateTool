
// ReadWriteExcel_MFCDlg.h : ͷ�ļ�
//

#pragma once
#include <map>
#include ".\excelReader\Application.h"
#include ".\excelReader\Workbooks.h"
#include ".\excelReader\Workbook.h"
#include ".\excelReader\Worksheets.h"
#include ".\excelReader\Worksheet.h"
#include ".\excelReader\Range.h"
using namespace std;

// CReadWriteExcel_MFCDlg �Ի���
class CReadWriteExcel_MFCDlg : public CDialogEx
{
// ����
public:
	CReadWriteExcel_MFCDlg(CWnd* pParent = NULL);	// ��׼���캯��

// �Ի�������
	enum { IDD = IDD_READWRITEEXCEL_MFC_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV ֧��


// ʵ��
protected:
	HICON m_hIcon;

	// ���ɵ���Ϣӳ�亯��
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()
public:
    afx_msg void OnBnClickedFindsource();
    afx_msg void OnBnClickedSetresultpath();
    afx_msg void OnBnClickedTranslate();
    afx_msg void OnClose();
private:
    //��ȡexcel�ļ��е�����
    void ReadExcelFile();
    //��ȡts�е�����
    void ReadTsFile();
    //���з��빤��
    void DoTranslate();
    int GetColumnCount();
    int GetRowCount();
    CString GetCell(long iRow, long iColumn);
    void PreLoadSheet();

private:
    CString m_SourceFilePathName;//�洢�����Ӧ��ϵ��excel�ļ�
    CString m_ResultFilePathName; //��Ҫ�������ļ��ľ���·��
    map<CString, CString> m_TranslateMap;//�����Ӧ��ϵ��ֵ��
    CApplication m_ExcelApp;
    CWorkbooks m_books;
    CWorkbook m_book;
    CWorksheets m_sheets;
    CWorksheet m_sheet;
    CRange m_range;
    ///�Ƿ��Ѿ�Ԥ������ĳ��sheet������
    BOOL          already_preload_;
    ///Create the SAFEARRAY from the VARIANT ret.
    COleSafeArray ole_safe_array_;
};
