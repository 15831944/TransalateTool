
// ReadWriteExcel_MFCDlg.cpp : 实现文件
//

#include "stdafx.h"
#include <iostream>
#include <fstream>
#include <vector>
#include <io.h>  
#include <algorithm>  
#include "ReadWriteExcel_MFC.h"
#include "ReadWriteExcel_MFCDlg.h"
#include "tinyxml2.h"
#include "afxdialogex.h"
#include "InfoDiaglog.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

using namespace  std;
using namespace tinyxml2;
#pragma comment(lib,"tinyxml2.lib")

// 用于应用程序“关于”菜单项的 CAboutDlg 对话框

class CAboutDlg : public CDialogEx
{
public:
    CAboutDlg();

    // 对话框数据
    enum { IDD = IDD_ABOUTBOX };

protected:
    virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

    // 实现
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


// CReadWriteExcel_MFCDlg 对话框



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

// CReadWriteExcel_MFCDlg 消息处理程序

BOOL CReadWriteExcel_MFCDlg::OnInitDialog()
{
    CDialogEx::OnInitDialog();

    // 将“关于...”菜单项添加到系统菜单中。

    // IDM_ABOUTBOX 必须在系统命令范围内。
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

    // 设置此对话框的图标。  当应用程序主窗口不是对话框时，框架将自动
    //  执行此操作
    SetIcon(m_hIcon, TRUE);			// 设置大图标
    SetIcon(m_hIcon, FALSE);		// 设置小图标

    // TODO:  在此添加额外的初始化代码

    return TRUE;  // 除非将焦点设置到控件，否则返回 TRUE
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

// 如果向对话框添加最小化按钮，则需要下面的代码
//  来绘制该图标。  对于使用文档/视图模型的 MFC 应用程序，
//  这将由框架自动完成。

void CReadWriteExcel_MFCDlg::OnPaint()
{
    if (IsIconic())
    {
        CPaintDC dc(this); // 用于绘制的设备上下文

        SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

        // 使图标在工作区矩形中居中
        int cxIcon = GetSystemMetrics(SM_CXICON);
        int cyIcon = GetSystemMetrics(SM_CYICON);
        CRect rect;
        GetClientRect(&rect);
        int x = (rect.Width() - cxIcon + 1) / 2;
        int y = (rect.Height() - cyIcon + 1) / 2;

        // 绘制图标
        dc.DrawIcon(x, y, m_hIcon);
    }
    else
    {
        CDialogEx::OnPaint();
    }
}

//当用户拖动最小化窗口时系统调用此函数取得光标
//显示。
HCURSOR CReadWriteExcel_MFCDlg::OnQueryDragIcon()
{
    return static_cast<HCURSOR>(m_hIcon);
}

void CReadWriteExcel_MFCDlg::OnBnClickedFindsource()
{
    CString workingDirectory;
    // OPTOINAL: Let's initialize the directory to the users home directory, assuming a max of 256 characters for path name:  
    wchar_t temp[256];
    GetEnvironmentVariable(_T("userprofile"), temp, 256);
    workingDirectory = temp;
    CFolderPickerDialog dlg(workingDirectory, 0, NULL, 0);
    if (dlg.DoModal())
    {
        m_SourceFilePathName = dlg.GetPathName();
        //AfxMessageBox(m_SourceFilePathName);
    }
    //设置输入框控件的值
    GetDlgItem(IDC_SETTARGETPATH)->SetWindowText(m_SourceFilePathName);
}

void CReadWriteExcel_MFCDlg::OnBnClickedSetresultpath()
{
    // TODO:  在此添加控件通知处理程序代码
    CString workingDirectory;
    // OPTOINAL: Let's initialize the directory to the users home directory, assuming a max of 256 characters for path name:  
    wchar_t temp[256];
    GetEnvironmentVariable(_T("userprofile"), temp, 256);
    workingDirectory = temp;
    CFolderPickerDialog dlg(workingDirectory, 0, NULL, 0);
    if (dlg.DoModal())
    {
        m_ResultFilePathName = dlg.GetPathName();
       // AfxMessageBox(m_ResultFilePathName);
    }
    //设置输入框控件的值
    GetDlgItem(IDC_SETRESULTPATH_EDIT)->SetWindowText(m_ResultFilePathName);
}

void CReadWriteExcel_MFCDlg::OnBnClickedTranslate()
{
    // TODO:  在此添加控件通知处理程序代码
    ReadExcelFile();
    TranslateTsFile();
	//保存没有翻译的文本
	saveUnMatchFile();
	int iCount = m_UnMatchMap.size();
	wchar_t strInfo[255] = {0};
	wsprintf(strInfo, L"Translate finish, total %d string not found in the excel file.", iCount);
	CString strInfoText = strInfo;

	//提示用户结果
	//CInfoDiaglog* dlg = new CInfoDiaglog(this);
	CInfoDiaglog *dlg = new CInfoDiaglog(this);
	dlg->Create(IDD_INFO_DIAOG);
	CWnd* wnd = FindWindow(NULL, _T("Info"));
	::SendMessage(
		*FindWindow(NULL, _T("Info"))//FindWInd(NULL,_T(***))通过窗口名返回窗口的句柄指针
		, WM_UNMATCH_TEXT
		, (WPARAM)&strInfoText
		, (LPARAM)&m_UnMatchTextFilePath);//信息的地址
	dlg->ShowWindow(SW_NORMAL);
}

void CReadWriteExcel_MFCDlg::OnClose()
{
    // TODO:  在此添加消息处理程序代码和/或调用默认值
    //SendMessage(WM_QUIT);
    //DestroyWindow();
    EndDialog(0);
    // CDialogEx::OnCancel();
    CDialogEx::OnClose();
}

void CReadWriteExcel_MFCDlg::ReadExcelFile()
{
    //遍历所有的excel文件
    string strSourceFilePath;
    WStringToString(m_SourceFilePathName.GetString(), strSourceFilePath);
    GetAllFormatFiles(strSourceFilePath, m_AllExcelFile, ".xlsx");
    //GetAllFormatFiles(strSourceFilePath, m_AllExcelFile, ".xls");
    vector<string>::iterator iter = m_AllExcelFile.begin();
    for (; iter != m_AllExcelFile.end(); ++iter)
    {
        LPDISPATCH lpDisp = NULL;
        if (getFileName(*iter).find("~") != string::npos)
        {
            //对于已经打开的excel产生的临时excel文件做过滤
            continue;
        }
        if (!m_ExcelApp.CreateDispatch(_T("Excel.Application"), NULL))
        {
            AfxMessageBox(_T("启动Excel服务器失败!"));
            return;
        }
        /*判断当前Excel的版本*/
        CString strExcelVersion = m_ExcelApp.get_Version();
        int iStart = 0;
        strExcelVersion = strExcelVersion.Tokenize(_T("."), iStart);
        //if (_T("11") == strExcelVersion)
        //{
        //	AfxMessageBox(_T("当前Excel的版本是2003。"));
        //}
        //else if (_T("12") == strExcelVersion)
        //{
        //	//AfxMessageBox(_T("当前Excel的版本是2007。"));
        //}
        //else
        //{
        //	AfxMessageBox(_T("当前Excel的版本是其他版本。"));
        //}
        m_ExcelApp.put_Visible(TRUE);
        m_ExcelApp.put_UserControl(FALSE);
        m_books.AttachDispatch(m_ExcelApp.get_Workbooks());
        try
        {
            /*打开指定的工作簿*/
            CString strPathTmp((*iter).c_str());
            lpDisp = m_books.Open(strPathTmp,
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
        //得到当前活跃sheet  
        //如果有单元格正处于编辑状态中，此操作不能返回，会一直等待  
        m_sheet.AttachDispatch(lpDisp);
        // VARIANT varRead = m_range.get_Value2();
        //PreLoadSheet();
        //读取excel的内容
        int iRowNum = GetRowCount();
        int iColumNum = GetColumnCount();
        for (int itera = 1; itera <= iRowNum; ++itera)
        {
            CString strKey, strValue;
            for (int j = 1; j <= iColumNum; ++j)
            {
                cout << GetCell(itera, j);
                if (j == 1)
                {
                    strKey = GetCell(itera, j);
                }
                if (j == 2)
                {
                    strValue = GetCell(itera, j);
					strKey = strKey.Trim();

                    m_TranslateMap.insert(make_pair(strKey, strValue));
                }
            }
        }
        //获取当前的语言类型
        string strLanguageType = (*iter);
        int iIndex = strLanguageType.find_last_of("\\");
        strLanguageType = strLanguageType.substr(iIndex + 1);
        iIndex = strLanguageType.find_last_of("_");
        int iLastIndex = strLanguageType.find_last_of(".");
        strLanguageType = strLanguageType.substr(iIndex + 1, iLastIndex - iIndex - 1);
		trim(strLanguageType);
        m_AllLanguageMap.insert(make_pair(strLanguageType, m_TranslateMap));
        //插入完成后清空当前map
        m_TranslateMap.clear();
        //m_ExcelApp
        m_ExcelApp.DetachDispatch();
        m_ExcelApp.Quit();
        cout << "" << endl;
    }
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

//得到行的总数
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
    //字符串
    if (already_preload_ == FALSE)
    {
        m_range.AttachDispatch(m_sheet.get_Cells());
        m_range.AttachDispatch(m_range.get_Item(COleVariant((long)iRow), COleVariant((long)iColumn)).pdispVal, true);
        vResult = m_range.get_Value2();
    }
    //如果数据依据预先加载了
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
    if (vResult.vt == VT_BSTR)//字符串
    {
        str = vResult.bstrVal;
    }
    else if (vResult.vt == VT_INT)
    {
        str.Format(_T("%d"), vResult.pintVal);
    }
    else if (vResult.vt == VT_R8)     //8字节的数字
    {
        str.Format(_T("%0.0f"), vResult.dblVal);
        //str.Format("%.0f",vResult.dblVal);
        //str.Format("%1f",vResult.fltVal);
    }
    else if (vResult.vt == VT_DATE)    //时间格式
    {
        SYSTEMTIME st;
        VariantTimeToSystemTime(vResult.date, &st);
        CTime tm(st);
        str = tm.Format(_T("%Y-%m-%d"));

    }
    else if (vResult.vt == VT_EMPTY)   //单元格空的
    {
        str = _T("");
    }

    m_range.ReleaseDispatch();

    return str;
}

//预先加载
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
    wstring strPathTmp = m_ResultFilePathName.GetString();
    find((wchar_t*)strPathTmp.c_str(), m_AllTsFile, L"\\*.ts");
    vector<string>::iterator iter = m_AllTsFile.begin();
    for (; iter != m_AllTsFile.end(); ++iter)
    {
		//设置当前正在处理ts文件的名字
		m_CurrentHandleTsFile = getFileName(*iter).c_str();
		m_CurrentHandleTsPath = (*iter).c_str();
        tinyxml2::XMLDocument* doc = new tinyxml2::XMLDocument();
        tinyxml2::XMLError error = doc->LoadFile((*iter).c_str());
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
							string strSuffix = getTsFileType(m_CurrentHandleTsFile.GetString());//
							trim(strRawText);
							trim(strSuffix);
							strTraslateText = TraslateRawData(strRawText, strSuffix);
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
        //修改完成后进行将修改保存
        doc->SaveFile((*iter).c_str());
        //将文件设置为UTF8编码格式
        ConvertTsFileToUTF8();
    }
}

std::string CReadWriteExcel_MFCDlg::TraslateRawData(string strRawData, string strType)
{
	//统一转换成小写
	//return "$$$$$$$$$$$$$$$$$$$$$$";
	transform(strType.begin(), strType.end(), strType.begin(), ::tolower);
    CString cstrRawData(strRawData.c_str());
    wstring wstrRect;
    string  strRect;
	trim(strType);
	map<CString,CString>  result =   m_AllLanguageMap[strType];
	CString strInfo = cstrRawData;
	strInfo.Append(L"%%%%%%%%%");
	string watrInfo;
	WStringToString(strInfo.GetString(), watrInfo);
	TRACE("king*************is:%s\n", watrInfo.c_str());
	cstrRawData = cstrRawData.Trim();
	string watrInfoData;
	WStringToString(cstrRawData.GetString(), watrInfoData);
	TRACE("king*************is:%s\n", watrInfoData.c_str());
	CString strText = result[cstrRawData];
	wstrRect = m_AllLanguageMap[strType][cstrRawData].GetString();
	if (wstrRect.empty())
	{
		m_UnMatchMap.insert(make_pair(cstrRawData, m_CurrentHandleTsFile));
	}
    //wstrRect = m_TranslateMap[cstrRawData].GetString();
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
    //读文件
    ifstream fileText(m_CurrentHandleTsPath.GetString());
    string strAllText((std::istreambuf_iterator<char>(fileText)), std::istreambuf_iterator<char>());
    string strUTF8 = string_To_UTF8(strAllText);
    //写文件
	ofstream out(m_CurrentHandleTsPath.GetString());
    if (out.is_open())
    {
        out.write(strUTF8.c_str(), strUTF8.length());
        out.close();
    }
}

std::string CReadWriteExcel_MFCDlg::string_To_UTF8(const std::string & str)
{
    int nwLen = ::MultiByteToWideChar(CP_ACP, 0, str.c_str(), -1, NULL, 0);
    wchar_t * pwBuf = new wchar_t[nwLen + 1];//一定要加1，不然会出现尾巴  
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

//获取特定格式的文件名  
void CReadWriteExcel_MFCDlg::GetAllFormatFiles(string path, vector<string>& files, string format)
{
    //文件句柄    
    long   hFile = 0;
    //文件信息    
    struct _finddata_t fileinfo;
    string p;
    path = trim(path);
    p.assign(path).append("\\*" + format);

    if ((hFile = _findfirst(p.c_str(), &fileinfo)) != -1)
    {
        do
        {
            if ((fileinfo.attrib &  _A_SUBDIR))
            {
                if (strcmp(fileinfo.name, ".") != 0 && strcmp(fileinfo.name, "..") != 0)
                {
                    //files.push_back(p.assign(path).append("\\").append(fileinfo.name) );  
                    GetAllFormatFiles(p.assign(path).append("\\").append(fileinfo.name), files, format);
                }
            }
            else
            {
                files.push_back(p.assign(path).append("\\").append(fileinfo.name));
            }
        } while (_findnext(hFile, &fileinfo) == 0);
        _findclose(hFile);
    }
}

//trim操作
string CReadWriteExcel_MFCDlg::trim(string& s)
{
    const string drop = " ";
    // trim right
    s.erase(s.find_last_not_of(drop) + 1);
    // trim left
    return s.erase(0, s.find_first_not_of(drop));
}

std::wstring CReadWriteExcel_MFCDlg::trim(wstring& s)
{
	const wstring drop = L" ";
	// trim right
	s.erase(s.find_last_not_of(drop) + 1);
	// trim left
	return s.erase(0, s.find_first_not_of(drop));
}

std::string CReadWriteExcel_MFCDlg::getFileName(string strFilePath)
{
    strFilePath = trim(strFilePath);
    string strFileName;
    int iIndex = strFilePath.find_last_of("\\");
    strFileName = strFilePath.substr(iIndex + 1);
    return strFileName;
}

void CReadWriteExcel_MFCDlg::find(wchar_t* lpPath, std::vector<std::string> &fileList, wchar_t* strFileType)
{
	wchar_t szFind[MAX_PATH] = {0};
    WIN32_FIND_DATA FindFileData;
    wcscpy_s(szFind, wcslen(lpPath) + 1, lpPath);
    wcscat_s(szFind, MAX_PATH, L"\\*.*");
    HANDLE hFind = ::FindFirstFile(szFind, &FindFileData);
    if (INVALID_HANDLE_VALUE == hFind)
    {
        return;
    }
    while (true)
    {
        if (FindFileData.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY)
        {
            if (FindFileData.cFileName[0] != '.')
            {
                wchar_t szFile[MAX_PATH];
                wcscpy_s(szFile, wcslen(lpPath)+ 1, lpPath);
                wcscat_s(szFile, MAX_PATH, L"\\");
                wcscat_s(szFile, MAX_PATH, (wchar_t*)(FindFileData.cFileName));
                find(szFile, fileList, strFileType);
            }
        }
        else
        {
            wstring strPath(szFind);
            string strRecPath;
            wstring wstrFileName(FindFileData.cFileName);
            string strFileName,strSuffix;
            WStringToString(wstrFileName, strFileName);
            int iIndex = strFileName.find_last_of(".");
			if (iIndex != string::npos)
			{
				strSuffix = strFileName.substr(iIndex);
				trim(strSuffix);
				TRACE("**********  file suffix begin  *******\n");
				TRACE("the suffix is: %s\n",strSuffix.c_str());
				TRACE("**********  file suffix end  *******\n");
				std::cout << strSuffix.c_str() << endl;
				wchar_t buf[MAX_PATH] = { 0 };
				wmemcpy_s(buf, MAX_PATH, szFind, wcslen(szFind) + 2);
				if (strSuffix.compare(".ts") == 0)
				{					
					//wmemcpy_s(buf, MAX_PATH, wstrFileName.c_str(), wcslen(wstrFileName.c_str()) + 2);
					wstring wstrPath(buf);
					int iIndex = wstrPath.find_last_of(L"\\");
					wstrPath = wstrPath.substr(0, iIndex + 1);
					//swprintf_s(wstrPath, wcslen(wstrPath.c_str()), L"\\%s", wstrFileName);
					wstrPath.append(wstrFileName);
					string strRecPath;
					WStringToString(wstrPath, strRecPath);
					fileList.push_back(strRecPath);
				}
			} 
        }
        if (!FindNextFile(hFind, &FindFileData))
            break;
    }
    FindClose(hFind);
}

void CReadWriteExcel_MFCDlg::saveUnMatchFile()
{
	m_UnMatchTextFilePath = L"d:\\infolog.txt";
	multimap<CString, CString>::iterator iter = m_UnMatchMap.begin();
	ofstream out(m_UnMatchTextFilePath);
	for (; iter != m_UnMatchMap.end();++iter)
	{
		if (out.is_open())
		{
			TRACE("$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$:   %s\n",((iter->first).GetString()));
			TRACE("$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$:   %s\n",((iter->second).GetString()));
			string strKey,strValue;
			WStringToString((iter->first).GetString(), strKey);
			WStringToString((iter->second).GetString(), strValue);
			out << strKey.c_str() <<"***" << strValue.c_str()<<endl;
		}
	}
	out.close();
}

std::string CReadWriteExcel_MFCDlg::getTsFileType(wstring strFileName)
{
	int iUnderlineIndex = strFileName.find_last_of(L"_");
	int iDotIndex = strFileName.find_last_of(L".");
	wstring strFileType;
	strFileType = strFileName.substr(iUnderlineIndex + 1, iDotIndex - iUnderlineIndex - 1);
	string strRec;
	WStringToString(strFileType, strRec);
	return strRec;
}
