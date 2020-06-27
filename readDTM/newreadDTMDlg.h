// newreadDTMDlg.h : header file
//

#if !defined(AFX_NEWREADDTMDLG_H__24C37FA6_4B82_11D3_A3C5_CA4CD97BAB5E__INCLUDED_)
#define AFX_NEWREADDTMDLG_H__24C37FA6_4B82_11D3_A3C5_CA4CD97BAB5E__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

/////////////////////////////////////////////////////////////////////////////
// CNewreadDTMDlg dialog

class CNewreadDTMDlg : public CDialog
{
// Construction
public:
	char* pszFileName;
	bool TimerSet;
    short worldCD[29];
	FILE *stream, *stream2;
	CNewreadDTMDlg(CWnd* pParent = NULL);	// standard constructor

// Dialog Data
	//{{AFX_DATA(CNewreadDTMDlg)
	enum { IDD = IDD_NEWREADDTM_DIALOG };
	CStatic	m_NewLabel;
	CProgressCtrl	m_Progress;
	CString	m_Label;
	//}}AFX_DATA

	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CNewreadDTMDlg)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV support
	//}}AFX_VIRTUAL
	int m_progresspercent;
	double beglog;
	double endlog;
	double beglat;
	double endlat;
	double lat0;
	double lon0;
	double hgt0;
	int ang;
	double aprn;
	short int landflag;
	short int mode;
	float modeval;
	short IgnoreMissingTiles;
	short TemperatureModel; //=0 for not terrestrial refraction modeling;1 for modeling using an mean ground temp = Tground
	short noVoid; //used for removing radar shadows
	short calcProfile; //1 for calculating profiles after extracting relevant DTM data
	double AzimuthStep; //step size in azimuth for horizon profile
	double Tground; //mean ground temperature, inputed in Celsisu, converted to Kelvin
	int DTMflag; //= 0 for 1 km GTOPO, 1 for 30 m NED DEM or SRTM, 2 for 90m SRTM
	double TRpart; //conserved portion of terrestrial refraction calculation
	BOOL directx;

// Implementation
protected:
	HICON m_hIcon;

	// Generated message map functions
	//{{AFX_MSG(CNewreadDTMDlg)
	virtual BOOL OnInitDialog();
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	virtual void OnOK();
	afx_msg void OnTimer(UINT nIDEvent);
	virtual void OnCancel();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_NEWREADDTMDLG_H__24C37FA6_4B82_11D3_A3C5_CA4CD97BAB5E__INCLUDED_)
