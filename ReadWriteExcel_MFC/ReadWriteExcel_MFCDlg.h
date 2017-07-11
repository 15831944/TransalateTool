
// ReadWriteExcel_MFCDlg.h : 头文件
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
using namespace std;

// CReadWriteExcel_MFCDlg 对话框
class CReadWriteExcel_MFCDlg : public CDialogEx
{
// 构造
public:
	CReadWriteExcel_MFCDlg(CWnd* pParent = NULL);	// 标准构造函数

// 对话框数据
	enum { IDD = IDD_READWRITEEXCEL_MFC_DIALOG };
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV 支持


// 实现
protected:
	HICON m_hIcon;
	// 生成的消息映射函数
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
    //读取excel文件中的内容
    void ReadExcelFile();
    //读取ts中的内容
    void ReadTsFile();
    //进行翻译工作
    void DoTranslate();
    int GetColumnCount();
    int GetRowCount();
    CString GetCell(long iRow, long iColumn);
    void PreLoadSheet();
	void TranslateTsFile();
	std::string TraslateRawData(std::string strRawData);
	BOOL WStringToString(const std::wstring &wstr, std::string &str);
	//将处理后的文件转换成UTF8编码格式
	void ConvertTsFileToUTF8();
	std::string string_To_UTF8(const std::string & str);
	void GetAllFormatFiles(string path, vector<string>& files, string format);
	string trim(string& s);
	string getFileName(string strFilePath);
	//获取当前解决方案中所有的ts文件并拷贝到一个特定的文件夹下
	bool getAllProjectTsFile();


private:
    CString m_SourceFilePathName;//存储翻译对应关系的excel文件
    CString m_ResultFilePathName; //需要被翻译文件的绝对路径
	map<std::string,map<CString, CString>> m_AllLanguageMap;//全语言字典；
    map<CString, CString> m_TranslateMap;//翻译对应关系键值对
    CApplication m_ExcelApp;
    CWorkbooks m_books;
    CWorkbook m_book;
    CWorksheets m_sheets;
    CWorksheet m_sheet;
    CRange m_range;
    ///是否已经预加载了某个sheet的数据
    BOOL          already_preload_;
    ///Create the SAFEARRAY from the VARIANT ret.
    COleSafeArray ole_safe_array_;
	vector<string> m_AllExcelFile;//所有的excel文件的集合
	vector<string> m_AllTsFile; //所有的ts文件的集合
};
