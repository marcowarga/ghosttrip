
// ghosttripDlg.h : header file
//

#pragma once
#include "afxcmn.h"
#include "afxwin.h"

// radians to degrees
#define DEGREES(x) ((x) * 57.295779513082320877) 
// degrees to radians
#define RADIANS(x) ((x) / 57.295779513082320877)
#define LOG_BASE 1.4142135623730950488
#define LOG_POWER 5300.0


struct WPT
{
    std::wstring strName;
    double dLongitude;
    double dLatitude;
    double dAltitude;
    WPT()
        : dLongitude(0.0),
          dLatitude(0.0),
          dAltitude(0.0)
    {
    }
    double GetDistanceTo(const WPT& wpt) const
    {
        double dDeltaLon = wpt.dLongitude - dLongitude;
        double dDeltaLat = wpt.dLatitude - dLatitude;
        double dDeltaAlt = wpt.dAltitude - dAltitude;
 
        double dAdjLon = cos(RADIANS((dLatitude + wpt.dLatitude) / 2.0)) * dDeltaLon;
        double dDistance = (sqrt((dDeltaLat * dDeltaLat) + (dAdjLon * dAdjLon)) * 111194.9266); // result in meters

        return dDistance;
    }
    double GetCourseTo(const WPT& wpt) const
    {
        if(wpt.dLongitude == dLongitude && wpt.dLatitude == dLatitude)
        {
            return 0.0;
        }

        double dDeltaLon = wpt.dLongitude - dLongitude;
        double dDeltaLat = wpt.dLatitude - dLatitude;
 
        double dCourse = atan2(dDeltaLat, dDeltaLon);   // result is +PI radians (cw) to -PI radians (ccw) from x (Longtitude) axis
        dCourse = DEGREES(dCourse); // convert radians to degrees
        if(dCourse <= 90.0)
        {
            dCourse = 90.0 - dCourse;
        }
        else
        {
            dCourse = 450.0 - dCourse;		// convert to 0 - 360 clockwise from north
        }

        return dCourse;
    }
};
typedef std::vector<WPT> WPTVEC;

enum LOGLEVEL
{
    LOGLEVEL_INFO = 0,
    LOGLEVEL_ERROR,
    LOGLEVEL_WARNING
};

class CghosttripDlg : public CDialogEx
{
    public:
        enum { IDD = IDD_GHOSTTRIP_DIALOG };

    protected:
        HICON m_hIcon;
        CListCtrl m_lcRoute;
        CButton m_btnLoad;
        CButton m_btnStart;
        CButton m_btnSetPos;
        CButton m_btnToClipboard;
        CButton m_btnStepN;
        CButton m_btnStepS;
        CButton m_btnStepE;
        CButton m_btnStepW;
        CEdit m_editNMEA;
        CEdit m_editLat;
        CEdit m_editLon;
        CEdit m_editAlt;
        WPTVEC m_vecWPT;
        int m_nTimer;
        int m_nTimer2;
        DWORD m_dwTickNext;
        int m_nCurrent;
        bool m_bRun;
        DWORD m_dwNmeaInterval;
        DWORD m_dwNmeaLastTick;
        WPT m_wptPos;
        CString m_strFile;
        int m_nLines;
        HANDLE m_hSer;
        int m_nStepMeters;
        ISpVoice* m_pIVoice;
        int m_nDestination;

    public:
        CghosttripDlg(CWnd* pParent = NULL);
        virtual ~CghosttripDlg();

    protected:
        void UpdateWaypointListControl();
        void Log(LOGLEVEL enLevel, const wchar_t* pszFormat, ...);
        void OnSendNmea();
        void UpdateControls();
        void Output_NEMA(double Lat, double Lon, double Alt, double Course, double Speed);
        void AddNmea(const char* pszNmea);
        int GetPrevWaypointIdx(int nWaypoint);
        int GetNextWaypointIdx(int nWaypoint);
        WPT& GetNextWaypoint(int nWaypoint);
        WPT& GetDestWaypoint();
        double GetDestDistance();
        void MoveCurrentWaypoint();
        void UpdateDestDistance(bool bSpeak = false);
        void Speak(const wchar_t* pszFormat, ...);

        BOOL PreTranslateMessage(MSG* pMsg);


    protected:
        virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV support


    protected:
        // Generated message map functions
        virtual BOOL OnInitDialog();
        afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
        afx_msg void OnPaint();
        afx_msg HCURSOR OnQueryDragIcon();
        afx_msg void OnBnClickedButtonLoad();
        afx_msg void OnDestroy();
        afx_msg void OnTimer(UINT nIDEvent);

    public:
        DECLARE_MESSAGE_MAP()
        afx_msg void OnBnClickedButtonStart();
        afx_msg void OnItemchangedListRoute(NMHDR *pNMHDR, LRESULT *pResult);
        afx_msg void OnBnClickedButtonClipb();
        afx_msg void OnBnClickedButtonSetpos();
        afx_msg void OnDblclkListRoute(NMHDR* pNMHDR, LRESULT* pResult);
        afx_msg void OnBnClickedButtonSerOpen();
        CComboBox m_cbSerial;
        CButton m_btnSerOpen;
        afx_msg void OnBnClickedButtonStepn();
        afx_msg void OnBnClickedButtonSteps();
        afx_msg void OnBnClickedButtonStepe();
        afx_msg void OnBnClickedButtonStepw();
        afx_msg void OnCustomdrawListRoute(NMHDR* pNMHDR, LRESULT* pResult);
};
