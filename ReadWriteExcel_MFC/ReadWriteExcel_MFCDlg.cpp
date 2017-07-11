
// ReadWriteExcel_MFCDlg.cpp : ʵ���ļ�
//

#include "stdafx.h"
#include <iostream>
#include <fstream>
#include "ReadWriteExcel_MFC.h"
#include "ReadWriteExcel_MFCDlg.h"
#include "tinyxml2.h"
#include "afxdialogex.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

using namespace  std;
using namespace tinyxml2;
#pragma comment(lib,"tinyxml2.lib")

// ����Ӧ�ó��򡰹��ڡ��˵���� CAboutDlg �Ի���

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

// �Ի�������
	enum { IDD = IDD_ABOUTBOX };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV ֧��

// ʵ��
protected:
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialogEx(CAboutDlg::IDD)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialogEx)
END_MESSAGE_MAP()


// CReadWriteExcel_MFCDlg �Ի���



CReadWriteExcel_MFCDlg::CReadWriteExcel_MFCDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(CReadWriteExcel_MFCDlg::IDD, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CReadWriteExcel_MFCDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CReadWriteExcel_MFCDlg, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
    ON_BN_CLICKED(ID_FINDSOURCE, &CReadWriteExcel_MFCDlg::OnBnClickedFindsource)
    ON_BN_CLICKED(ID_SETRESULTPATH, &CReadWriteExcel_MFCDlg::OnBnClickedSetresultpath)
    ON_BN_CLICKED(ID_TRANSLATE, &CReadWriteExcel_MFCDlg::OnBnClickedTranslate)
    ON_WM_CLOSE()
END_MESSAGE_MAP()

// CReadWriteExcel_MFCDlg ��Ϣ�������

BOOL CReadWriteExcel_MFCDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// ��������...���˵�����ӵ�ϵͳ�˵��С�

	// IDM_ABOUTBOX ������ϵͳ���Χ�ڡ�
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != NULL)
	{
		BOOL bNameValid;
		CString strAboutMenu;
		bNameValid = strAboutMenu.LoadString(IDS_ABOUTBOX);
		ASSERT(bNameValid);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// ���ô˶Ի����ͼ�ꡣ  ��Ӧ�ó��������ڲ��ǶԻ���ʱ����ܽ��Զ�
	//  ִ�д˲���
	SetIcon(m_hIcon, TRUE);			// ���ô�ͼ��
	SetIcon(m_hIcon, FALSE);		// ����Сͼ��

	// TODO:  �ڴ���Ӷ���ĳ�ʼ������

	return TRUE;  // ���ǽ��������õ��ؼ������򷵻� TRUE
}

void CReadWriteExcel_MFCDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialogEx::OnSysCommand(nID, lParam);
	}
}

// �����Ի��������С����ť������Ҫ����Ĵ���
//  �����Ƹ�ͼ�ꡣ  ����ʹ���ĵ�/��ͼģ�͵� MFC Ӧ�ó���
//  �⽫�ɿ���Զ���ɡ�

void CReadWriteExcel_MFCDlg::OnPaint()
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
HCURSOR CReadWriteExcel_MFCDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

void CReadWriteExcel_MFCDlg::OnBnClickedFindsource()
{
    CFileDialog dlg(TRUE, //TRUEΪOPEN�Ի���FALSEΪSAVE AS�Ի���
        NULL,
        NULL,
        OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT,
        (LPCTSTR)_TEXT("Excel Files (*.xls)|*.jpg|All Files (*.*)|*.*||"),
        NULL);
    if (dlg.DoModal() == IDOK)
    {
        m_SourceFilePathName = dlg.GetPathName(); //�ļ�����������FilePathName��
    }
    else
    {
        m_SourceFilePathName =  "";
    }
    //���������ؼ���ֵ
    GetDlgItem(IDC_SETTARGETPATH)->SetWindowText(m_SourceFilePathName);
    //UpdateData(TRUE);
}

void CReadWriteExcel_MFCDlg::OnBnClickedSetresultpath()
{
    // TODO:  �ڴ���ӿؼ�֪ͨ����������
    CFileDialog dlg(TRUE, //TRUEΪOPEN�Ի���FALSEΪSAVE AS�Ի���
        NULL,
        NULL,
        OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT,
        (LPCTSTR)_TEXT("Excel Files (*.ts)|*.ts|All Files (*.*)|*.*||"),
        NULL);
    if (dlg.DoModal() == IDOK)
    {
        m_ResultFilePathName = dlg.GetPathName(); //�ļ�����������FilePathName��
    }
    else
    {
        m_ResultFilePathName = "";
    }
    //���������ؼ���ֵ
    GetDlgItem(IDC_SETRESULTPATH_EDIT)->SetWindowText(m_ResultFilePathName);
}

void CReadWriteExcel_MFCDlg::OnBnClickedTranslate()
{
    // TODO:  �ڴ���ӿؼ�֪ͨ����������
	ReadExcelFile();
	TranslateTsFile();
}

void CReadWriteExcel_MFCDlg::OnClose()
{
    // TODO:  �ڴ������Ϣ�����������/�����Ĭ��ֵ
    //SendMessage(WM_QUIT);
    //DestroyWindow();
    EndDialog(0);
   // CDialogEx::OnCancel();
    CDialogEx::OnClose();
}

void CReadWriteExcel_MFCDlg::ReadExcelFile()
{
    LPDISPATCH lpDisp = NULL;
    if (!m_ExcelApp.CreateDispatch(_T("Excel.Application"), NULL))
    {
        AfxMessageBox(_T("����Excel������ʧ��!"));
        return;
    }
    /*�жϵ�ǰExcel�İ汾*/
    CString strExcelVersion = m_ExcelApp.get_Version();
    int iStart = 0;
    strExcelVersion = strExcelVersion.Tokenize(_T("."), iStart);
    if (_T("11") == strExcelVersion)
    {
        AfxMessageBox(_T("��ǰExcel�İ汾��2003��"));
    }
    else if (_T("12") == strExcelVersion)
    {
        //AfxMessageBox(_T("��ǰExcel�İ汾��2007��"));
    }
    else
    {
        AfxMessageBox(_T("��ǰExcel�İ汾�������汾��"));
    }
    m_ExcelApp.put_Visible(TRUE);
    m_ExcelApp.put_UserControl(FALSE);
    m_books.AttachDispatch(m_ExcelApp.get_Workbooks());
    try
    {
        /*��ָ���Ĺ�����*/
        lpDisp = m_books.Open(m_SourceFilePathName,
            vtMissing, vtMissing, vtMissing, vtMissing, vtMissing,
            vtMissing, vtMissing, vtMissing, vtMissing, vtMissing,
            vtMissing, vtMissing, vtMissing, vtMissing);
        m_book.AttachDispatch(lpDisp);
    }
    catch (...)
    {
        AfxMessageBox(L"open excel failed");
        std::cout << "open excel failed" << endl;
        return;
    }
    m_sheets.AttachDispatch(m_book.get_Worksheets());
    lpDisp = m_book.get_ActiveSheet();
    

    //�õ���ǰ��Ծsheet  
    //����е�Ԫ�������ڱ༭״̬�У��˲������ܷ��أ���һֱ�ȴ�  

    m_sheet.AttachDispatch(lpDisp);
   
   
   // VARIANT varRead = m_range.get_Value2();
    //PreLoadSheet();
    //��ȡexcel������
    int iRowNum = GetRowCount();
    int iColumNum = GetColumnCount();
    for (int i = 1; i <= iRowNum;++i)
    {
        CString strKey, strValue;
        for (int j = 1; j <= iColumNum; ++j)
        {
            cout<< GetCell(i, j);
            if (j == 1)
            {
                strKey = GetCell(i, j);
            }
            if (j==2)
            {
                strValue = GetCell(i, j);
                m_TranslateMap.insert(make_pair(strKey, strValue));
            }
        }
    }
    cout << "" << endl;
}

void CReadWriteExcel_MFCDlg::ReadTsFile()
{

}

void CReadWriteExcel_MFCDlg::DoTranslate()
{
	
}

int CReadWriteExcel_MFCDlg::GetColumnCount()
{
    CRange range;
    CRange usedRange;
    usedRange.AttachDispatch(m_sheet.get_UsedRange(), true);
    range.AttachDispatch(usedRange.get_Columns(), true);
    int count = range.get_Count();
    usedRange.ReleaseDispatch();
    range.ReleaseDispatch();
    return count;
}

//�õ��е�����
int CReadWriteExcel_MFCDlg::GetRowCount()
{
    CRange range;
    CRange usedRange;
    usedRange.AttachDispatch(m_sheet.get_UsedRange(), true);
    range.AttachDispatch(usedRange.get_Rows(), true);
    int count = range.get_Count();
    usedRange.ReleaseDispatch();
    range.ReleaseDispatch();
    return count;
}

CString CReadWriteExcel_MFCDlg::GetCell(long iRow, long iColumn)
{
    COleVariant vResult;
    //�ַ���
    if (already_preload_ == FALSE)
    {
        m_range.AttachDispatch(m_sheet.get_Cells());
        m_range.AttachDispatch(m_range.get_Item(COleVariant((long)iRow), COleVariant((long)iColumn)).pdispVal, true);
        vResult = m_range.get_Value2();
    }
    //�����������Ԥ�ȼ�����
    else
    {
        long read_address[2];
        VARIANT val;
        read_address[0] = iRow;
        read_address[1] = iColumn;
        ole_safe_array_.GetElement(read_address, &val);
        vResult = val;
    }

    CString str;
    if (vResult.vt == VT_BSTR)       //�ַ���
    {
        str = vResult.bstrVal;
    }
    else if (vResult.vt == VT_INT)
    {
        str.Format(_T("%d"), vResult.pintVal);
    }
    else if (vResult.vt == VT_R8)     //8�ֽڵ�����
    {
        str.Format(_T("%0.0f"), vResult.dblVal);
        //str.Format("%.0f",vResult.dblVal);
        //str.Format("%1f",vResult.fltVal);
    }
    else if (vResult.vt == VT_DATE)    //ʱ���ʽ
    {
        SYSTEMTIME st;
        VariantTimeToSystemTime(vResult.date, &st);
        CTime tm(st);
        str = tm.Format(_T("%Y-%m-%d"));

    }
    else if (vResult.vt == VT_EMPTY)   //��Ԫ��յ�
    {
        str = _T("");
    }

    m_range.ReleaseDispatch();

    return str;
}

//Ԥ�ȼ���
void CReadWriteExcel_MFCDlg::PreLoadSheet()
{
    CRange used_range;
    used_range = m_sheet.get_UsedRange();

    VARIANT ret_ary = used_range.get_Value2();
    if (!(ret_ary.vt & VT_ARRAY))
    {
        return;
    }
    ole_safe_array_.Clear();
    ole_safe_array_.Attach(ret_ary);
}

void CReadWriteExcel_MFCDlg::TranslateTsFile()
{
	//readTs file 
	tinyxml2::XMLDocument* doc = new tinyxml2::XMLDocument();
	wstring  wstrPath = m_ResultFilePathName.GetString();
	string strTsFile;
	WStringToString(wstrPath, strTsFile);
	tinyxml2::XMLError error = doc->LoadFile(strTsFile.c_str());
	XMLElement* ele = doc->RootElement();
	ele = ele->FirstChildElement("context");
	while (ele != NULL)
	{
		XMLNode* firstEle = ele->FirstChild();
		while (firstEle != NULL)
		{
			string strText = firstEle->ToElement()->Name();
			if (strText == "message")
			{
				XMLElement* child = firstEle->FirstChildElement();
				string strRawText;
				string strTraslateText;
				while (child != NULL)
				{
					if (string(child->Name()) == "source")
					{
						strRawText = child->GetText();
						strTraslateText = TraslateRawData(strRawText);	
					}
					if (string(child->Name()) == "translation")
					{
						child->SetText(strTraslateText.c_str());
					}
					child = child->NextSiblingElement();
				}	
			}
			firstEle = firstEle->NextSibling();
		}
		ele = ele->NextSiblingElement("context");
	}
	//�޸���ɺ���н��޸ı���
	//doc->SetBOM(false);
	doc->SaveFile(strTsFile.c_str());
	//���ļ�����ΪUTF8�����ʽ
	ConvertTsFileToUTF8();
}

std::string CReadWriteExcel_MFCDlg::TraslateRawData(string strRawData)
{
	CString cstrRawData(strRawData.c_str());
	wstring wstrRect;
	string  strRect;
	wstrRect = m_TranslateMap[cstrRawData].GetString();
	WStringToString(wstrRect, strRect);
	return strRect;
}

BOOL CReadWriteExcel_MFCDlg::WStringToString(const std::wstring &wstr, std::string &str)
{
	int nLen = (int)wstr.length();
	DWORD num = WideCharToMultiByte(CP_ACP, 0, (LPCWSTR)wstr.c_str(), -1, NULL, 0, NULL, 0);
	str.resize(num, ' ');
	int nResult = WideCharToMultiByte(CP_ACP, 0, (LPCWSTR)wstr.c_str(), nLen, (LPSTR)str.c_str(), num, NULL, NULL);
	if (nResult == 0)
	{
		return FALSE;
	}
	return TRUE;
}

void CReadWriteExcel_MFCDlg::ConvertTsFileToUTF8()
{
   //���ļ�
	ifstream fileText(m_ResultFilePathName.GetString());
	string strAllText((std::istreambuf_iterator<char>(fileText)), std::istreambuf_iterator<char>());
	string strUTF8 = string_To_UTF8(strAllText);
	//д�ļ�
	ofstream out(m_ResultFilePathName.GetString());
	if (out.is_open())
	{
		out.write(strUTF8.c_str(), strUTF8.length());
		out.close();
	}
}

std::string CReadWriteExcel_MFCDlg::string_To_UTF8(const std::string & str)
{
	int nwLen = ::MultiByteToWideChar(CP_ACP, 0, str.c_str(), -1, NULL, 0);
	wchar_t * pwBuf = new wchar_t[nwLen + 1];//һ��Ҫ��1����Ȼ�����β��  
	ZeroMemory(pwBuf, nwLen * 2 + 2);
	::MultiByteToWideChar(CP_ACP, 0, str.c_str(), str.length(), pwBuf, nwLen);
	int nLen = ::WideCharToMultiByte(CP_UTF8, 0, pwBuf, -1, NULL, NULL, NULL, NULL);
	char * pBuf = new char[nLen + 1];
	ZeroMemory(pBuf, nLen + 1);
	::WideCharToMultiByte(CP_UTF8, 0, pwBuf, nwLen, pBuf, nLen, NULL, NULL);
	std::string retStr(pBuf);
	delete[]pwBuf;
	delete[]pBuf;
	pwBuf = NULL;
	pBuf = NULL;
	return retStr;
}