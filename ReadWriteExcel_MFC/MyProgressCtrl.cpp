#include "stdafx.h"
#include "MyProgressCtrl.h"


CMyProgressCtrl::CMyProgressCtrl()
{
    initData();
}


CMyProgressCtrl::~CMyProgressCtrl()
{
}

int CMyProgressCtrl::SetPos(int nPos)
{
    if (nPos < m_iMin)
    {
        m_iPos = m_iMin;
    }
    else if (nPos > m_iMax)
    {
        m_iPos = m_iMax;
    }
    else
    {
        m_iPos = nPos;
    }
    return m_iPos;
}

void CMyProgressCtrl::SetRange(int nLower, int nUpper)
{
    m_iMax = nUpper;
    m_iMin = nLower;
    m_iPos = nLower;
    m_nBarWidth = 0;
}

void CMyProgressCtrl::initData()
{
    m_freeColor = RGB(255,255,255);
    m_prgsColor = RGB(0,0,255);
    m_prgsTextColor = RGB(255,0,0);
    m_freeTextColor = RGB(0,255,0);
    m_ProgressGif.Load(L"F:\\择善培训\\学习计划\\我的项目\\Translatetools\\TransalateTool\\ReadWriteExcel_MFC\\res\\progress.gif");
}

BEGIN_MESSAGE_MAP(CMyProgressCtrl, CProgressCtrl)
    ON_WM_PAINT()
END_MESSAGE_MAP()


void CMyProgressCtrl::OnPaint()
{
    CPaintDC dc(this); // device context for painting
    // TODO:  在此处添加消息处理程序代码
    // 不为绘图消息调用 CProgressCtrl::OnPaint()
    CRect LeftRect, RightRect, ClientRect;
    GetClientRect(ClientRect);
    LeftRect = RightRect = ClientRect;
    double dFraction = (double)(m_iPos - m_iMin) / ((double)(m_iMax - m_iMin));
    //绘制进度条的完成部分
    LeftRect.right = LeftRect.left + (int)((LeftRect.right - LeftRect.left)*dFraction);
    //绘制Gif图片
    m_ProgressGif.Draw(dc, LeftRect);


    //dc.FillSolidRect(LeftRect, m_prgsColor);
    //绘制剩余部分
    RightRect.left = LeftRect.right;
    dc.FillSolidRect(RightRect, m_freeColor);
    //绘制文本
    m_strText.Format(_T("%d%%"), (int)(dFraction*100.00));
    dc.SetBkMode(TRANSPARENT);

    CRgn rgn;
    rgn.CreateRectRgn(LeftRect.left, LeftRect.top, LeftRect.right, LeftRect.bottom);
    dc.SelectClipRgn(&rgn);
    dc.SetTextColor(m_prgsTextColor);
    dc.DrawText(m_strText, ClientRect, DT_CENTER | DT_VCENTER | DT_SINGLELINE);

    rgn.DeleteObject();
    rgn.CreateRectRgn(RightRect.left, RightRect.top, RightRect.right, RightRect.bottom);
    dc.SelectClipRgn(&rgn);
    dc.SetTextColor(m_freeTextColor);
    dc.DrawText(m_strText, ClientRect, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
}

