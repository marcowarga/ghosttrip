
// ghosttripDlg.cpp : implementation file
//

#include "stdafx.h"
#include "ghosttrip.h"
#include "ghosttripDlg.h"
#include "afxdialogex.h"
#include "tinyxml.h"
#include "enumser.h"
#include "algorithms.h"

#define ONEMETERRAD double(0.000008998719243599958);

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


time_t Now; // the time of starting this program 
 
char buf[128];
 
// calculate a CRC for the line of input
void do_crc(char *pch)
{
    unsigned char crc;
 
    if (*pch != '$') 
        return;		// does not start with '$' - so can't CRC
 
    pch++;			// skip '$'
    crc = 0;
 
    // scan between '$' and '*' (or until CR LF or EOL)
    while ((*pch != '*') && (*pch != '\0') && (*pch != '\r') && (*pch != '\n'))
    { // checksum calcualtion done over characters between '$' and '*'
        crc ^= *pch;
        pch++;
    }
 
    // add or re-write checksum
 
    sprintf(pch,"*%02X\r\n",(unsigned int)crc);
}
 
CStringA UTF16toUTF8(const CStringW& utf16)
{
   CStringA utf8;
   int len = WideCharToMultiByte(CP_UTF8, 0, utf16, -1, NULL, 0, 0, 0);
   if (len>1)
   { 
      char *ptr = utf8.GetBuffer(len-1);
      if (ptr) WideCharToMultiByte(CP_UTF8, 0, utf16, -1, ptr, len, 0, 0);
      utf8.ReleaseBuffer();
   }
   return utf8;
}

CStringW UTF8toUTF16(const CStringA& utf8)
{
   CStringW utf16;
   int len = MultiByteToWideChar(CP_UTF8, 0, utf8, -1, NULL, 0);
   if (len>1)
   { 
      wchar_t *ptr = utf16.GetBuffer(len-1);
      if (ptr) MultiByteToWideChar(CP_UTF8, 0, utf8, -1, ptr, len);
      utf16.ReleaseBuffer();
   }
   return utf16;
}

std::string toAnsi(const std::wstring& utf16)
{
   CStringA utf8;
   int len = WideCharToMultiByte(CP_UTF8, 0, utf16.c_str(), -1, NULL, 0, 0, 0);
   if (len>1)
   { 
      char *ptr = utf8.GetBuffer(len-1);
      if (ptr) WideCharToMultiByte(CP_UTF8, 0, utf16.c_str(), -1, ptr, len, 0, 0);
      utf8.ReleaseBuffer();
   }
   return static_cast<const char*>(utf8);
}

std::wstring toWide(const std::string& utf8)
{
   CStringW utf16;
   int len = MultiByteToWideChar(CP_UTF8, 0, utf8.c_str() , -1, NULL, 0);
   if (len>1)
   { 
      wchar_t *ptr = utf16.GetBuffer(len-1);
      if (ptr) MultiByteToWideChar(CP_UTF8, 0, utf8.c_str(), -1, ptr, len);
      utf16.ReleaseBuffer();
   }
   return static_cast<const wchar_t*>(utf16);
}

// CAboutDlg dialog used for App About

class CAboutDlg : public CDialogEx
{
public:
    CAboutDlg();

// Dialog Data
    enum { IDD = IDD_ABOUTBOX };

    protected:
    virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

// Implementation
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



CghosttripDlg::CghosttripDlg(CWnd* pParent /*=NULL*/)
    : CDialogEx(CghosttripDlg::IDD, pParent),
    m_nTimer(-1),
    m_dwTickNext(0),
    m_bRun(false),
    m_nCurrent(-1),
    m_dwNmeaInterval(1000),
    m_dwNmeaLastTick(0),
    m_nLines(0),
    m_hSer(INVALID_HANDLE_VALUE),
    m_nStepMeters(15),
    m_pIVoice(NULL),
    m_nTimer2(-1),
    m_nDestination(-1)
{
    m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);

    if(FAILED(CoCreateInstance(CLSID_SpVoice, NULL, CLSCTX_ALL, IID_ISpVoice, (void **)&m_pIVoice)))
    {
        m_pIVoice = NULL;
    }
}

CghosttripDlg::~CghosttripDlg()
{
    if(m_pIVoice)
    {
        m_pIVoice->Release();
        m_pIVoice = NULL;
    }
}

void CghosttripDlg::DoDataExchange(CDataExchange* pDX)
{
    CDialogEx::DoDataExchange(pDX);
    DDX_Control(pDX, IDC_LIST_ROUTE, m_lcRoute);
    DDX_Control(pDX, IDC_BUTTON_LOAD, m_btnLoad);
    DDX_Control(pDX, IDC_BUTTON_START, m_btnStart);
    DDX_Control(pDX, IDC_BUTTON_SETPOS, m_btnSetPos);
    DDX_Control(pDX, IDC_BUTTON_CLIPB, m_btnToClipboard);
    DDX_Control(pDX, IDC_BUTTON_STEPN, m_btnStepN);
    DDX_Control(pDX, IDC_BUTTON_STEPS, m_btnStepS);
    DDX_Control(pDX, IDC_BUTTON_STEPE, m_btnStepE);
    DDX_Control(pDX, IDC_BUTTON_STEPW, m_btnStepW);
    DDX_Control(pDX, IDC_EDIT_NMEA, m_editNMEA);
    DDX_Control(pDX, IDC_EDIT_LAT, m_editLat);
    DDX_Control(pDX, IDC_EDIT_LON, m_editLon);
    DDX_Control(pDX, IDC_EDIT_ALT, m_editAlt);
    DDX_Control(pDX, IDC_COMBO_SER, m_cbSerial);
    DDX_Control(pDX, IDC_BUTTON_SEROPEN, m_btnSerOpen);
}

BEGIN_MESSAGE_MAP(CghosttripDlg, CDialogEx)
    ON_WM_SYSCOMMAND()
    ON_WM_PAINT()
    ON_WM_QUERYDRAGICON()
    ON_BN_CLICKED(IDC_BUTTON_LOAD, &CghosttripDlg::OnBnClickedButtonLoad)
    ON_BN_CLICKED(IDC_BUTTON_START, &CghosttripDlg::OnBnClickedButtonStart)
    ON_NOTIFY(LVN_ITEMCHANGED, IDC_LIST_ROUTE, &CghosttripDlg::OnItemchangedListRoute)
    ON_WM_DESTROY()
    ON_WM_TIMER()
    ON_BN_CLICKED(IDC_BUTTON_CLIPB, &CghosttripDlg::OnBnClickedButtonClipb)
    ON_BN_CLICKED(IDC_BUTTON_SETPOS, &CghosttripDlg::OnBnClickedButtonSetpos)
    ON_BN_CLICKED(IDC_BUTTON_SEROPEN, &CghosttripDlg::OnBnClickedButtonSerOpen)
	ON_NOTIFY(NM_DBLCLK, IDC_LIST_ROUTE, OnDblclkListRoute)
    ON_BN_CLICKED(IDC_BUTTON_STEPN, &CghosttripDlg::OnBnClickedButtonStepn)
    ON_BN_CLICKED(IDC_BUTTON_STEPS, &CghosttripDlg::OnBnClickedButtonSteps)
    ON_BN_CLICKED(IDC_BUTTON_STEPE, &CghosttripDlg::OnBnClickedButtonStepe)
    ON_BN_CLICKED(IDC_BUTTON_STEPW, &CghosttripDlg::OnBnClickedButtonStepw)
    ON_NOTIFY(NM_CUSTOMDRAW, IDC_LIST_ROUTE, OnCustomdrawListRoute)
END_MESSAGE_MAP()

void AutoSizeListCtrl(CListCtrl& list, int nSortCol = -1)
{
    // autosize all columns and add
    // space for the sort arrow in the
    // sorted column (except if its the rightmost column)
    int nLastCol = list.GetHeaderCtrl()->GetItemCount() - 1;
    for(int nCol = 0; nCol <= nLastCol; nCol++)
    {
        list.SetColumnWidth(nCol, LVSCW_AUTOSIZE_USEHEADER);
        if((nSortCol != -1) && (nCol == nSortCol) && (nCol != nLastCol))
        {
            list.SetColumnWidth(nCol, list.GetColumnWidth(nCol) + 32);
        }
    }
}

void CghosttripDlg::UpdateWaypointListControl()
{
    m_lcRoute.SetRedraw(FALSE);

    m_lcRoute.DeleteAllItems();
    WPTVEC::iterator it;
    int nMonitored = 0;
    for(int i = 0; i < (int)m_vecWPT.size(); i++)
    {
        const WPT& wpt = m_vecWPT[i];

        CString strTemp;
        int nItem = m_lcRoute.InsertItem(m_lcRoute.GetItemCount(), strTemp);
        if(nItem >= 0)
        {
            m_lcRoute.SetItemText(nItem, 1, wpt.strName.c_str());
            strTemp.Format(_T("%.8lf"), wpt.dLongitude);
            m_lcRoute.SetItemText(nItem, 2, strTemp);
            strTemp.Format(_T("%.8lf"), wpt.dLatitude);
            m_lcRoute.SetItemText(nItem, 3, strTemp);
            strTemp.Format(_T("%.2lf m"), wpt.dAltitude);
            m_lcRoute.SetItemText(nItem, 4, strTemp);

            double dDistance = 0.0;
            if((i + 1) < (int)m_vecWPT.size())
            {
                const WPT& wptNext = m_vecWPT[i + 1];
                dDistance = wpt.GetDistanceTo(wptNext);
            }
            else if(i != 0)
            {
                const WPT& wptNext = m_vecWPT[0];
                dDistance = wpt.GetDistanceTo(wptNext);
            }
            strTemp.Format(_T("%.2f m"), dDistance);
            m_lcRoute.SetItemText(nItem, 5, strTemp);
            m_lcRoute.SetCheck(nItem, TRUE);
        }
    }

    // fit the column widths to their content/header
    AutoSizeListCtrl(m_lcRoute);

    m_lcRoute.SetRedraw(TRUE);
}

BOOL CghosttripDlg::OnInitDialog()
{
    CDialogEx::OnInitDialog();

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
    SetIcon(m_hIcon, TRUE);			// Set big icon
    SetIcon(m_hIcon, FALSE);		// Set small icon

    m_lcRoute.SetExtendedStyle(LVS_EX_FULLROWSELECT | LVS_EX_CHECKBOXES);

    m_lcRoute.InsertColumn(0, _T("Wait"));
    m_lcRoute.InsertColumn(1, _T("Name"));
    m_lcRoute.InsertColumn(2, _T("Longitude"));
    m_lcRoute.InsertColumn(3, _T("Latitude"));
    m_lcRoute.InsertColumn(4, _T("Altitude"));
    m_lcRoute.InsertColumn(5, _T("Distance"));
    m_lcRoute.InsertColumn(6, _T("To go"));

    m_cbSerial.ResetContent();
    std::vector<std::wstring> vecPorts;
    CEnumerateSerial::UsingRegistry(vecPorts);
    for(int i = 0; i < (int)vecPorts.size(); i++)
    {
        m_cbSerial.AddString(vecPorts[i].c_str());
    }
    if(m_cbSerial.GetCount())
    {
        m_cbSerial.SetCurSel(0);
    }
    UpdateWaypointListControl();

    UpdateControls();

    m_nTimer = SetTimer(1, 500, NULL);
    m_nTimer = SetTimer(2, 20000, NULL);


    return TRUE;  // return TRUE  unless you set the focus to a control
}

void CghosttripDlg::OnSysCommand(UINT nID, LPARAM lParam)
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
void CghosttripDlg::OnPaint()
{
    if (IsIconic())
    {
        CPaintDC dc(this); // device context for painting

        SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

        // Center icon in client rectangle
        int cxIcon = GetSystemMetrics(SM_CXICON);
        int cyIcon = GetSystemMetrics(SM_CYICON);
        CRect rect;
        GetClientRect(&rect);
        int x = (rect.Width() - cxIcon + 1) / 2;
        int y = (rect.Height() - cyIcon + 1) / 2;

        // Draw the icon
        dc.DrawIcon(x, y, m_hIcon);
    }
    else
    {
        CDialogEx::OnPaint();
    }
}

HCURSOR CghosttripDlg::OnQueryDragIcon()
{
    return static_cast<HCURSOR>(m_hIcon);
}

void CghosttripDlg::Log(LOGLEVEL enLevel, const wchar_t* pszFormat, ...)
{
    std::wstring strTitle;
    UINT unType = 0;
    switch(enLevel)
    {
    case LOGLEVEL_ERROR:
    {
        unType = MB_OK | MB_ICONERROR;
        strTitle = L"Error";
        break;
    }
    case LOGLEVEL_WARNING:
    {
        unType = MB_OK | MB_ICONEXCLAMATION;
        strTitle = L"Warning";
        break;
    }
    case LOGLEVEL_INFO:
    default:
    {
        unType = MB_OK | MB_ICONHAND;
        strTitle = L"Info";
        break;
    }
    }

    va_list args;
    va_start(args, pszFormat);
    wchar_t szBuffer[1024] = {0};
    _vsnwprintf(szBuffer, 1024 - 1, pszFormat, args);
    szBuffer[1024 - 1] = 0;
    va_end(args);
    ::MessageBoxW(m_hWnd, szBuffer, strTitle.c_str(), unType);
}

void CghosttripDlg::OnBnClickedButtonLoad()
{
    int nResult = 0;

    CFileDialog dlg(TRUE, NULL, NULL, OFN_FILEMUSTEXIST);
    dlg.m_ofn.lpstrFilter = L"KML Files (*.kml)\0*.kml\0All Files (*.*)\0*.*\0\0";
    dlg.m_ofn.lpstrTitle = L"Open KML File";

    if(dlg.DoModal() == IDOK)
    {
        m_strFile = dlg.GetPathName();
    }
    else
    {
        return;
    }

    //m_strFile = L"C:\\Users\\watz\\Desktop\\Portale.kml";
    TiXmlDocument doc(toAnsi((const wchar_t*)m_strFile));
    if(!doc.LoadFile())
    {
        Log(LOGLEVEL_ERROR, L"Failed to parse file \"%s\":\r\n%d(%d)=%S.\n", (const wchar_t*)m_strFile, doc.ErrorRow(), doc.ErrorCol(), doc.ErrorDesc());
        return;
    }

    m_vecWPT.clear();
    m_lcRoute.DeleteAllItems();

    TiXmlNode* pNodeKml = doc.FirstChild("kml");
    if(pNodeKml)
    {
        TiXmlNode* pNodeDoc = pNodeKml->FirstChild("Document");
        if(pNodeDoc)
        {
            TiXmlNode* pNodeFolder = pNodeDoc->FirstChild("Folder");
            if(!pNodeFolder)
            {
                pNodeFolder = pNodeDoc;
            }
            while(pNodeFolder)
            {
                TiXmlNode* pNodePlacemark = pNodeFolder->FirstChild("Placemark");
                while(pNodePlacemark)
                {
                    WPT wpt;
                    TiXmlNode* pNodeName = pNodePlacemark->FirstChild("name");
                    if(pNodeName)
                    {
                        TiXmlElement* pElemName = pNodeName->ToElement();
                        if(pElemName)
                        {
                            wpt.strName = toWide(pElemName->GetText());
                        }
                    }

                    TiXmlNode* pNodePoint = pNodePlacemark->FirstChild("Point");
                    if(pNodePoint)
                    {
                        TiXmlNode* pNodeCoordinates = pNodePoint->FirstChild("coordinates");
                        if(pNodeCoordinates)
                        {
                            TiXmlElement* pElemCoordinates = pNodeCoordinates->ToElement();
                            if(pElemCoordinates)
                            {
                                std::string strCoords = pElemCoordinates->GetText();
                                if(sscanf(strCoords.c_str(), "%lf,%lf,%lf", &wpt.dLongitude, &wpt.dLatitude, &wpt.dAltitude) == 3)
                                {
                                    m_vecWPT.push_back(wpt);
                                }
                            }
                        }
                    }

                    pNodePlacemark = pNodePlacemark->NextSibling("Placemark");
                }

                pNodeFolder = pNodeFolder->NextSibling("Folder");
            }
        }
    }

    if(m_vecWPT.size() == 0)
    {
        Log(LOGLEVEL_ERROR, L"File does not contain any waypoints");
        return;
    }

    tsp t;
    for(size_t i = 0; i < m_vecWPT.size(); i++)
    {
        t.add(m_vecWPT[i].dLongitude, m_vecWPT[i].dLatitude);
        t.nearest_neighbor();
    }

    WPTVEC vecWPT;
    /*
    for(size_t i = m_vecWPT.size(); i > 0; i--)
    {
        city* pc = t.getSolution(i - 1);
        vecWPT.push_back(m_vecWPT[pc->get_id()]);
    }
    */
    for(size_t i = 0; i < m_vecWPT.size(); i++)
    {
        city* pc = t.getSolution(i);
        vecWPT.push_back(m_vecWPT[pc->get_id()]);
    }
    m_vecWPT.swap(vecWPT);

    UpdateWaypointListControl();

    m_bRun = false;
    m_dwTickNext = GetTickCount() + m_dwNmeaInterval;
    m_nCurrent = 0;
    m_nDestination = GetNextWaypointIdx(m_nCurrent);
    m_wptPos = m_vecWPT[0];
    m_wptPos.strName = L"pos";
    UpdateDestDistance();
    UpdateControls();
    m_dwNmeaLastTick = ::GetTickCount();
    OnSendNmea();
}

void CghosttripDlg::OnBnClickedButtonStart()
{
    CString strTemp;
    if(m_bRun)
    {
        Speak(L"halt");
        m_bRun = false;
        UpdateDestDistance();
    }
    else
    {
        Speak(L"go");
        m_bRun = true;
        UpdateDestDistance(true);
    }
    UpdateControls();
}

void CghosttripDlg::OnBnClickedButtonSetpos()
{
    if(m_lcRoute.GetSelectedCount() == 1)
    {
        int nCurrent = m_lcRoute.GetNextItem(0, LVNI_SELECTED);
        if(nCurrent >= 0)
        {
            m_bRun = false;
            m_lcRoute.SetItemText(m_nCurrent, 6, L"");
            m_nCurrent = nCurrent;
            m_dwTickNext = GetTickCount() + m_dwNmeaInterval;
            m_wptPos = m_vecWPT[nCurrent];
            m_wptPos.strName = L"pos";
            UpdateDestDistance(true);
            m_dwNmeaLastTick = ::GetTickCount();
            OnSendNmea();
        }
    }
}

void CghosttripDlg::OnBnClickedButtonSerOpen()
{
    if(m_hSer != INVALID_HANDLE_VALUE)
    {
        ::CloseHandle(m_hSer);
        m_hSer = INVALID_HANDLE_VALUE;
    }
    else
    {
        CString strName;
        m_cbSerial.GetWindowText(strName);
        strName = L"\\\\.\\" + strName;
        m_hSer = CreateFile(strName, GENERIC_READ | GENERIC_WRITE, 0, NULL, OPEN_EXISTING, 0, NULL);
        if(m_hSer != INVALID_HANDLE_VALUE)
        {
            DCB cfg;
            ZeroMemory(&cfg, sizeof(cfg));
            if(GetCommState(m_hSer, &cfg))
            {
                cfg.BaudRate = 4800;
                cfg.ByteSize = 8;
                cfg.Parity = NOPARITY;
                cfg.StopBits = ONESTOPBIT;
                cfg.fBinary = TRUE;
                cfg.fParity = FALSE;
                cfg.fOutxCtsFlow = FALSE;
                cfg.fOutxDsrFlow = FALSE;
                cfg.fDtrControl = DTR_CONTROL_DISABLE;
                cfg.fRtsControl = RTS_CONTROL_DISABLE;
                cfg.fAbortOnError = FALSE;

                SetCommState(m_hSer, &cfg);
            }
        }
    }
    UpdateControls();
}

void CghosttripDlg::MoveCurrentWaypoint()
{
    if(!m_bRun)
    {
        CString strTemp;
        strTemp.Format(L"%.2f m (waiting)", GetDestDistance());
        m_lcRoute.SetItemText(m_nCurrent, 6, strTemp);
    }
    else
    {
        CString strTemp;
        strTemp.Format(L"%.2f m", GetDestDistance());
        m_lcRoute.SetItemText(m_nCurrent, 6, strTemp);
    }
    m_dwNmeaLastTick = ::GetTickCount();
    OnSendNmea();
    m_dwTickNext = GetTickCount() + m_dwNmeaInterval;
}

void CghosttripDlg::UpdateDestDistance(bool bSpeak /*= false*/)
{
    CString strTemp;
    if(!m_bRun)
    {
        strTemp.Format(L"%.2f m (waiting)",  GetDestDistance());
    }
    else
    {
        strTemp.Format(L"%.2f m",  GetDestDistance());
    }
    m_lcRoute.SetItemText(m_nCurrent, 6, strTemp);

    if(bSpeak)
    {
        if(!m_bRun)
        {
            Speak(L"waiting");
        }
        else
        {
            Speak(L"%d", (int)GetDestDistance());
        }
    }
}

void CghosttripDlg::Speak(const wchar_t* pszFormat, ...)
{
    if(m_pIVoice)
    {
        va_list args;
        va_start(args, pszFormat);
        wchar_t szBuffer[1024] = {0};
        _vsnwprintf(szBuffer, 1024 - 1, pszFormat, args);
        szBuffer[1024 - 1] = 0;
        va_end(args);
        m_pIVoice->Speak(_bstr_t(szBuffer), SPF_ASYNC, NULL);
    }
}

void CghosttripDlg::OnBnClickedButtonStepn()
{
    m_wptPos.dLatitude += ((double)m_nStepMeters) * ONEMETERRAD;
    MoveCurrentWaypoint();
}


void CghosttripDlg::OnBnClickedButtonSteps()
{
    m_wptPos.dLatitude -= ((double)m_nStepMeters) * ONEMETERRAD;
    MoveCurrentWaypoint();
}


void CghosttripDlg::OnBnClickedButtonStepe()
{
    m_wptPos.dLongitude += ((double)m_nStepMeters) * ONEMETERRAD;
    MoveCurrentWaypoint();
}


void CghosttripDlg::OnBnClickedButtonStepw()
{
    m_wptPos.dLongitude -= ((double)m_nStepMeters) * ONEMETERRAD;
    MoveCurrentWaypoint();
}

BOOL CghosttripDlg::PreTranslateMessage(MSG* pMsg)
{
   if(pMsg->message == WM_KEYDOWN   )  
   {
      if(pMsg->wParam == VK_NUMPAD8)
      {
            m_wptPos.dLatitude += ((double)0.5) * ONEMETERRAD;
            MoveCurrentWaypoint();
          return TRUE;
      }
      else if(pMsg->wParam == VK_NUMPAD2)
      {
            m_wptPos.dLatitude -= ((double)0.5) * ONEMETERRAD;
            MoveCurrentWaypoint();
          return TRUE;
      }
      else if(pMsg->wParam == VK_NUMPAD6)
      {
            m_wptPos.dLongitude += ((double)0.5) * ONEMETERRAD;
            MoveCurrentWaypoint();
          return TRUE;
      }
      else if(pMsg->wParam == VK_NUMPAD4)
      {
            m_wptPos.dLongitude -= ((double)0.5) * ONEMETERRAD;
            MoveCurrentWaypoint();
          return TRUE;
      }
      if(pMsg->wParam == VK_NUMPAD8 || pMsg->wParam == VK_UP)
      {
          OnBnClickedButtonStepn();
          return TRUE;
      }
      else if(pMsg->wParam == VK_NUMPAD2 || pMsg->wParam == VK_DOWN)
      {
          OnBnClickedButtonSteps();
          return TRUE;
      }
      else if(pMsg->wParam == VK_NUMPAD6 || pMsg->wParam == VK_RIGHT)
      {
          OnBnClickedButtonStepe();
          return TRUE;
      }
      else if(pMsg->wParam == VK_NUMPAD4 || pMsg->wParam == VK_LEFT)
      {
          OnBnClickedButtonStepw();
          return TRUE;
      }
      else if(pMsg->wParam == VK_NUMPAD5 || pMsg->wParam == VK_SPACE)
      {
          OnBnClickedButtonStart();
          return TRUE;
      }
   }

   return CDialog::PreTranslateMessage(pMsg);  
}

void CghosttripDlg::OnDblclkListRoute(NMHDR* pNMHDR, LRESULT* pResult) 
{
    NM_LISTVIEW* pNML = (NM_LISTVIEW*)pNMHDR;
    if(pNML != NULL)
    {
        int nItem = m_lcRoute.HitTest(pNML->ptAction);
        if(nItem > -1 && nItem != m_nCurrent)
        {
            m_nDestination = nItem;
            m_lcRoute.RedrawWindow();
            UpdateDestDistance();
            m_dwTickNext = GetTickCount() + m_dwNmeaInterval;
            m_dwNmeaLastTick = ::GetTickCount();
            OnSendNmea();
        }
        UpdateControls();
    }
    *pResult = 0;
}

void CghosttripDlg::OnTimer(UINT nIDEvent) 
{
    if((nIDEvent == 1) && (m_nTimer != 0))
    {
        if(m_dwTickNext != 0)
        {
            if(GetTickCount() >= m_dwTickNext)
            {
                if(m_vecWPT.size() > 0)
                {
                    OnSendNmea();
                }
                m_dwTickNext = GetTickCount() + m_dwNmeaInterval;
            }
        }
    }
    if((nIDEvent == 2) && (m_nTimer2 != 0))
    {
        UpdateDestDistance(true);
    }

    CDialog::OnTimer(nIDEvent);
}

void CghosttripDlg::OnDestroy() 
{
    if(m_nTimer > 0)
    {
        KillTimer(m_nTimer);
        m_nTimer = NULL;
    }

    CDialog::OnDestroy();
}

void CghosttripDlg::OnItemchangedListRoute(NMHDR *pNMHDR, LRESULT *pResult)
{
    NM_LISTVIEW* pNMListView = (NM_LISTVIEW*)pNMHDR;

    UpdateControls();

    *pResult = 0;
}

int CghosttripDlg::GetPrevWaypointIdx(int nWaypoint)
{
    if(nWaypoint == 0)
    {
        return m_vecWPT.size() - 1;
    }
    else
    {
        return nWaypoint - 1;
    }
}

int CghosttripDlg::GetNextWaypointIdx(int nWaypoint)
{
    int nNext = 0;
    if((nWaypoint + 1) < (int)m_vecWPT.size())
    {
        nNext = nWaypoint + 1;
    }
    return nNext;
}

WPT& CghosttripDlg::GetNextWaypoint(int nWaypoint)
{
    return m_vecWPT[GetNextWaypointIdx(nWaypoint)];
}

WPT& CghosttripDlg::GetDestWaypoint()
{
    return m_vecWPT[m_nDestination];
}

double CghosttripDlg::GetDestDistance()
{
    return m_wptPos.GetDistanceTo(GetDestWaypoint());
}

void CghosttripDlg::OnSendNmea()
{
    bool bDone = false;
    double dCourse = 0.0;
    double dSpeed = 0.0;
    if(m_bRun)
    {
        double dSecondsPassed = (double)(::GetTickCount() - m_dwNmeaLastTick) / 1000.0;

        const WPT& wptCurrent = m_vecWPT[m_nCurrent];
        const WPT& wptNext = GetDestWaypoint();

        dCourse = wptCurrent.GetCourseTo(wptNext);
        double dDistanceTotal = wptCurrent.GetDistanceTo(wptNext);
        double dDistanceLeft = m_wptPos.GetDistanceTo(wptNext);

        dSpeed = 14.0;
        //dSpeed = 30.0;
        double dDistanceMade = dSecondsPassed * dSpeed;
        if(dDistanceMade < dDistanceLeft)
        {
            double dFactor =  dDistanceMade / dDistanceLeft;

            double dDeltaAlt = (wptNext.dAltitude - m_wptPos.dAltitude) * dFactor;
            double dDeltaLon = (wptNext.dLongitude - m_wptPos.dLongitude) * dFactor;
            double dDeltaLat = (wptNext.dLatitude - m_wptPos.dLatitude) * dFactor;

            m_wptPos.dAltitude += dDeltaAlt;
            m_wptPos.dLongitude += dDeltaLon;
            m_wptPos.dLatitude += dDeltaLat;
        }
        else
        {
            m_wptPos.dAltitude = wptNext.dAltitude;
            m_wptPos.dLongitude = wptNext.dLongitude;
            m_wptPos.dLatitude = wptNext.dLatitude;
            dDistanceMade = dDistanceLeft;
            bDone = true;
        }
        dDistanceLeft = m_wptPos.GetDistanceTo(wptNext);
        double dDistanceDone = dDistanceTotal - dDistanceLeft;

        CString strTemp;
        strTemp.Format(L"%.2f m", dDistanceLeft);
        m_lcRoute.SetItemText(m_nCurrent, 6, strTemp);

        Output_NEMA(m_wptPos.dLatitude, m_wptPos.dLongitude, m_wptPos.dAltitude, dCourse, dSpeed);

        if(bDone)
        {
            strTemp.Empty();
            m_lcRoute.SetItemText(m_nCurrent, 6, strTemp);

            m_nCurrent = m_nDestination;
            m_nDestination = GetNextWaypointIdx(m_nCurrent);
            m_wptPos = m_vecWPT[m_nCurrent];
            m_wptPos.strName = L"pos";

            if(m_pIVoice)
            {
                m_pIVoice->Speak(L"Destination reached", SPF_ASYNC, NULL);
            }
            if(m_lcRoute.GetCheck(m_nCurrent))
            {
                m_bRun = false;
            }
            UpdateDestDistance(true);
        }
    }
    else
    {
        Output_NEMA(m_wptPos.dLatitude, m_wptPos.dLongitude, m_wptPos.dAltitude, 0.0, 0.0);
    }
    UpdateControls();

    m_dwNmeaLastTick = ::GetTickCount();
}

CString FormatLat(double dLat)
{
    wchar_t cDir = L'S';
    if(dLat >= 0)
    {
        cDir = L'N';
    }
 
    dLat = fabs(dLat);
    int nLatDeg = (int)(dLat);
    int nLatMin = (int)((dLat - (double)nLatDeg) * 60.0);
    double dLatSec = ((((dLat - (double)nLatDeg) * 60.0) - (double)nLatMin) * 60);
 
    CString strResult;
    strResult.Format(L"%02d°%02d'%2.2lf'' %c", nLatDeg, nLatMin, dLatSec, cDir);

    return strResult;
}

CString FormatLon(double dLon)
{
    wchar_t cDir = L'W';
    if(dLon >= 0)
    {
        cDir = L'E';
    }
 
    dLon = fabs(dLon);
    int nLonDeg = (int)(dLon);
    int nLonMin = (int)((dLon - (double)nLonDeg) * 60.0);
    double dLonSec = ((((dLon - (double)nLonDeg) * 60.0) - (double)nLonMin) * 60);

    CString strResult;
    strResult.Format(L"%02d°%02d'%2.2lf'' %c", nLonDeg, nLonMin, dLonSec, cDir);

    return strResult;
}

void CghosttripDlg::UpdateControls()
{
    if(!m_bRun)
    {
        m_btnStart.SetWindowText(L"Go");
    }
    else
    {
        m_btnStart.SetWindowText(L"Halt");
        m_btnStart.EnableWindow(TRUE);
    }
    if(m_lcRoute.GetSelectedCount() == 1)
    {
        m_btnSetPos.EnableWindow(TRUE);
    }
    else
    {
        m_btnSetPos.EnableWindow(FALSE);
    }

    if(m_nCurrent >= 0)
    {
        CString strTemp;
        strTemp.Format(L"%.14lf (%s)", m_wptPos.dLatitude, (LPCWSTR)FormatLat(m_wptPos.dLatitude));
        m_editLat.SetWindowText(strTemp);
        strTemp.Format(L"%.14lf (%s)", m_wptPos.dLongitude, (LPCWSTR)FormatLon(m_wptPos.dLongitude));
        m_editLon.SetWindowText(strTemp);
        strTemp.Format(L"%.2lf m", m_wptPos.dAltitude);
        m_editAlt.SetWindowText(strTemp);
    }
    else
    {
        m_editLat.SetWindowText(L"");
        m_editLon.SetWindowText(L"");
        m_editAlt.SetWindowText(L"");
    }

    if(m_hSer != INVALID_HANDLE_VALUE)
    {
        m_btnSerOpen.SetWindowText(L"Close");
        m_btnSerOpen.EnableWindow(TRUE);
        m_cbSerial.EnableWindow(FALSE);
    }
    else
    {
        m_btnSerOpen.SetWindowText(L"Open");
        if(m_cbSerial.GetCount())
        {
            m_btnSerOpen.EnableWindow(TRUE);
            m_cbSerial.EnableWindow(TRUE);
        }
        else
        {
            m_btnSerOpen.EnableWindow(FALSE);
            m_cbSerial.EnableWindow(FALSE);
        }
    }

    if(m_nCurrent < 0)
    {
        m_btnStepN.EnableWindow(FALSE);
        m_btnStepS.EnableWindow(FALSE);
        m_btnStepE.EnableWindow(FALSE);
        m_btnStepW.EnableWindow(FALSE);
    }
    else
    {
        m_btnStepN.EnableWindow(TRUE);
        m_btnStepS.EnableWindow(TRUE);
        m_btnStepE.EnableWindow(TRUE);
        m_btnStepW.EnableWindow(TRUE);
    }
}

void CghosttripDlg::Output_NEMA(double Lat, double Lon, double Alt, double Course, double Speed)
{
    int LatDeg;					// latitude - degree part
    double LatMin;				// latitude - minute part
    char LatDir;				// latitude - direction N/S
    int LonDeg;					// longtitude - degree part
    double LonMin;				// longtitude - minute part
    char LonDir;				// longtitude - direction E/W
    struct tm *ptm;
    time_t Time;
 
    time(&Time);
    ptm = gmtime(&Time);
 
    if (Lat >= 0)
        LatDir = 'N';
    else
        LatDir = 'S';
 
    Lat = fabs(Lat);
    LatDeg = (int)(Lat);
    LatMin = (Lat - (double)LatDeg) * 60.0;
 
    if (Lon >= 0)
        LonDir = 'E';
    else
        LonDir = 'W';
 
    Lon = fabs(Lon);
    LonDeg = (int)(Lon);
    LonMin = (Lon - (double)LonDeg) * 60.0;
 
    // $GPGGA - 1st in epoc - 5 satellites in view, FixQual = 1, 45m Geoidal separation HDOP = 2.4
    sprintf(buf,"$GPGGA,%02d%02d%02d.000,%02d%07.4f,%c,%03d%07.4f,%c,1,05,0.9,%.1f,M,45.0,M,,*",
        ptm->tm_hour,ptm->tm_min,ptm->tm_sec,LatDeg,LatMin,LatDir,LonDeg,LonMin,LonDir,Alt - 45.0);
    do_crc(buf); // add CRC to buf
    AddNmea(buf);
 
 
    switch((int)Time % 3)
    { // include 'none' or $GPGSA or $GPGSV in 3 second cycle
    case 0:
        break;
 
    case 1:
        // 3D fix - 5 satellites (3,7,18,19 & 22) in view. PDOP = 3.3,HDOP = 2.4, VDOP = 2.3
        sprintf(buf,"$GPGSA,A,3,03,07,18,19,22,,,,,,,,3.3,2.4,2.3*");
        do_crc(buf); // add CRC to buf
        AddNmea(buf);
        break;
 
    case 2:
        // two lines og GPGSV messages - 1st line of 2, 8 satellites being tracked in total
        // 03,07 in view 11,12 being tracked
        sprintf(buf,"$GPGSV,2,1,08,03,89,276,30,07,63,181,22,11,,,,12,,,*");
        do_crc(buf); // add CRC to buf
        AddNmea(buf);
 
        // GPGSV 2nd line of 2, 8 satellites being tracked in total
        // 18,19,22 in view 27 being tracked
        sprintf(buf,"$GPGSV,2.2,08,18,73,111,35,19,33,057,27,22,57,173,37,27,,,*");
        do_crc(buf); // add CRC to buf
        AddNmea(buf);
        break;
    }
 
    //$GPRMC
    sprintf(buf,"$GPRMC,%02d%02d%02d.000,A,%02d%07.4f,%c,%03d%07.4f,%c,%.2f,%.2f,%02d%02d%02d,,,A*",
        ptm->tm_hour,ptm->tm_min,ptm->tm_sec,LatDeg,LatMin,LatDir,LonDeg,LonMin,LonDir,Speed * 1.943844,Course,ptm->tm_mday,ptm->tm_mon + 1,ptm->tm_year % 100);
    do_crc(buf); // add CRC to buf
    AddNmea(buf);
 
    // $GPVTG message last in epoc
    sprintf(buf,"$GPVTG,%.2f,T,,,%.2f,N,%.2f,K,A*",Course,Speed * 1.943844,Speed * 3.6);
    do_crc(buf); // add CRC to buf
    //fputs(buf,stdout);
    AddNmea(buf);
}

void CghosttripDlg::AddNmea(const char* pszNmea)
{
    if(m_hSer != INVALID_HANDLE_VALUE)
    {
        DWORD dwWrite = strlen(pszNmea);
        if(dwWrite)
        {
            DWORD dwWritten = 0;
            DWORD dwStart = GetTickCount();

            BOOL bRes = ::WriteFile(m_hSer, pszNmea, dwWrite, &dwWritten, NULL);
            if(!bRes)
            {
                OutputDebugStringW(L"send failed\n");
            }
            CString str;
            str.Format(L"serial send of %d bytes took %d ms\n", dwWrite, GetTickCount() - dwStart);
            OutputDebugStringW(str);
        }
    }


    CString str;
    m_editNMEA.GetWindowText(str);
    str += toWide(pszNmea).c_str();
    if(m_nLines == 500)
    {
        int nPos = str.Find(L"\n");
        if(nPos)
        {
            str = str.Mid(nPos + 1);
        }
    }
    else
    {
        m_nLines++;
    }
    m_editNMEA.SetWindowText(str);
    m_editNMEA.SetSel(-1);
}

void SetClipboardText(CString & szData)
{
    HGLOBAL h;
    LPTSTR arr;

    size_t bytes = (szData.GetLength()+1)*sizeof(TCHAR);
    h=GlobalAlloc(GMEM_MOVEABLE, bytes);
    arr=(LPTSTR)GlobalLock(h);
    ZeroMemory(arr,bytes);
    _tcscpy_s(arr, szData.GetLength()+1, szData);
    szData.ReleaseBuffer();
    GlobalUnlock(h);

    ::OpenClipboard (NULL);
    EmptyClipboard();
    SetClipboardData(CF_UNICODETEXT, h);
    CloseClipboard();
}


void CghosttripDlg::OnBnClickedButtonClipb()
{
    CString str;
    m_editNMEA.GetWindowText(str);
    SetClipboardText(str);
}

void CghosttripDlg::OnCustomdrawListRoute(NMHDR* pNMHDR, LRESULT* pResult)
{
   NMLVCUSTOMDRAW* pLVCD = reinterpret_cast<NMLVCUSTOMDRAW*>(pNMHDR);
   *pResult = CDRF_DODEFAULT;

   switch(pLVCD->nmcd.dwDrawStage)
   {
   case CDDS_PREPAINT:
      *pResult = CDRF_NOTIFYITEMDRAW;
      break;

   case CDDS_ITEMPREPAINT:
      *pResult = CDRF_NOTIFYSUBITEMDRAW;
      break;

   case (CDDS_ITEMPREPAINT | CDDS_SUBITEM):
      {
         if(pLVCD->nmcd.dwItemSpec == m_nDestination)
         {
            pLVCD->clrText = RGB(255, 0, 0);
         }
      }
      break;
   }
}
