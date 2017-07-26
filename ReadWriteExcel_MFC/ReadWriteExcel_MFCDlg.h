
// ReadWriteExcel_MFCDlg.h : ͷ�ļ�
//

#pragma once
#include <map>
#include <list>
#include <vector>
#include ".\excelReader\Application.h"
#include ".\excelReader\Workbooks.h"
#include ".\excelReader\Workbook.h"
#include ".\excelReader\Worksheets.h"
#include ".\excelReader\Worksheet.h"
#include ".\excelReader\Range.h"
#include "afxwin.h"
using namespace std;
class CMyProgressCtrl;
enum AppType
{
	Type_Translator,   //�滻��
	Type_Extrctor      //�ı���ȡ��
};

//�����С��־
typedef struct 
{
	long lWidth;
	long lHeight;
}FontFlag;

#define WM_UNMATCH_TEXT (WM_USER+100)

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
	//��ȡ��ǰ��Ŀ�µ�����ts�ļ�
	void GetAllTsFile();
	//��ȡ�ض����Ե�ts�ļ��ļ���
	vector<string> getTsFileByLanguage(string strLanguageType);
    //���з��빤��
    void DoTranslate();
    int GetColumnCount();
    int GetRowCount();
    CString GetCell(long iRow, long iColumn);
    void PreLoadSheet();
	void TranslateTsFile();
	std::string TraslateRawData(std::string strRawData,std::string strType);
	BOOL WStringToString(const std::wstring &wstr, std::string &str);
	BOOL StringToWString(const std::string &str, std::wstring &wstr);
	//���������ļ�ת����UTF8�����ʽ
	void ConvertTsFileToUTF8();
	std::string string_To_UTF8(const std::string & str);
	void GetAllFormatFiles(string path, vector<string>& files, string format);
	string trim(string& s);
	wstring trim(wstring& s);
	string getFileName(string strFilePath);
	//��ȡ��ǰ������������е�ts�ļ���������һ���ض����ļ�����
    void find(wchar_t* lpPath, std::vector<std::string> &fileList,wchar_t* strFileType);
	//��û��ƥ����ֶ�д���ļ���
	void saveUnMatchFile();
	string getTsFileType(wstring strFileName);
	string ws2s(const std::wstring& wstr);
	char* UnicodeToUtf8(const wchar_t* unicode);
	//��ʼ����Ա����
	void initData();
	//����ȡ����
	void doExtracAction();
	//��ʼ������
	void initUI();
	//��ȡ�ض�����ĳ���
	FontFlag getStringSize(CString strText);
	//��ʾ�ض���tooltip
	void setToolTip();
	//�Ż����滻ʵ�����
	CString ReplaceEntitySymbols(CString strText);
	
private:
    CString m_SourceFilePathName;//�洢�����Ӧ��ϵ��excel�ļ�
    CString m_ResultFilePathName; //��Ҫ�������ļ��ľ���·��
	map<std::string,map<CString, CString>> m_AllLanguageMap;//ȫ�����ֵ䣻
    map<CString, CString> m_TranslateMap;//�����Ӧ��ϵ��ֵ��
	multimap<CString, CString> m_UnMatchMap;//δƥ�䵽���ַ������ֵ�
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
	vector<string> m_AllExcelFile;//���е�excel�ļ��ļ���
	vector<string> m_AllTsFile; //���е�ts�ļ��ļ���
	vector<string> m_AllEnTsFile; //��ǰ��Ŀ�����е�Ӣ�İ汾��ts�ļ�����
	CString        m_CurrentHandleTsFile;//��ǰ���ڴ����ts�ļ���
	CString        m_UnMatchTextFilePath;//����δƥ�䵽���ַ������ļ�·��
	CString        m_CurrentHandleTsPath;//��ǰ���ڱ������Ts���ļ���·��
	CString        m_FilterValue;       //��������ֵ�û������ض���ts�ļ�
	list<CString> m_AllFilter;         //��ǰ֧�ֵĹ��˲���
	CMyProgressCtrl* progress;
	bool             m_IsReOpenExcelFile; //�Ƿ����´�excel�ļ�
	bool             m_ISReoOpenTsFile; //�Ƿ����´�Ts�ļ�
	bool             m_IsExtrcted;     //��ǰts�ļ��Ƿ��Ѿ��滻��
	AppType             m_CurAppType;      //��ǰҳ������
	CButton*         m_ExtractorButton; //��ȡ��ť
	CToolTipCtrl     m_Tooltip;         //��ʾtooltip
	// excel�ļ������ļ���
	CStatic m_SourcePathText;
	// ts�ļ������ļ���
	CStatic m_ResultPathText;
public:
	afx_msg void OnChangeFilterbox();
	afx_msg void OnSelchangeTypeCombo();
	virtual BOOL PreTranslateMessage(MSG* pMsg);
};
