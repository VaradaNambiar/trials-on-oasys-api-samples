// GsaComClient.cpp : Defines the class behaviors for the application.
//

#include "stdafx.h"
#include "GsaComClient.h"

#include <comutil.h>
#include <iostream>

#include "GsaComClientDlg.h"
#include <string>
//#import "C:\\repo\\oasys-combined-main\\gsa\\programs64_d\\GSA.tlb" no_namespace
// bring dlls from prog64 to below loc
// opt 2 GsRegister will reg the dev (debug/release) into registry editor 
#import "C:\\Program Files\\Oasys\\GSA 10.2\\Gsa.tlb" no_namespace


// replace the above with path to GSA.tlb in the program files folder

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// CGsaComClientApp

BEGIN_MESSAGE_MAP(CGsaComClientApp, CWinApp)
	ON_COMMAND(ID_HELP, &CWinApp::OnHelp)
END_MESSAGE_MAP()


// CGsaComClientApp construction

CGsaComClientApp::CGsaComClientApp()
{
	// TODO: add construction code here,
	// Place all significant initialization in InitInstance
}


// The one and only CGsaComClientApp object

CGsaComClientApp theApp;

enum Output_Init_Flags
{
	OP_INIT_2D_BOTTOM = 0x01,     // output 2D stresses at bottom layer
	OP_INIT_2D_MIDDLE = 0x02,     // output 2D stresses at middle layer
	OP_INIT_2D_TOP = 0x04,        // output 2D stresses at top layer
	OP_INIT_2D_BENDING = 0x08,    // output 2D stresses at bending layer
	OP_INIT_2D_AVGE = 0x10,      // average 2D element stresses at nodes
	OP_INIT_1D_AUTO_PTS = 0x20,  // calculate 1D results at interesting points
};

// CGsaComClientApp initialization

BOOL CGsaComClientApp::InitInstance()
{
	CWinApp::InitInstance();

	SetRegistryKey(_T("GSA Com Client"));

	CGsaComClientDlg dlg;
	m_pMainWnd = &dlg;
	INT_PTR nResponse = dlg.DoModal();
	if (nResponse == IDOK)
	{
		try
		{
			invokeGsa(dlg.m_filename, dlg.m_analysed_filename, dlg.m_analysed_filename_report);
		}
		catch (_com_error* e)
		{
			std::string error(e->Description());
			TRACE(error.c_str());
		}

	}
	else if (nResponse == IDCANCEL)
	{
		// TODO: Place code here to handle when the dialog is
		//  dismissed with Cancel
	}

	// Since the dialog has been closed, return FALSE so that we exit the
	//  application, rather than start the application's message pump.
	return FALSE;
}

int SaveDefaultOutputView(IComAutoPtr& pObj)
{
	CString file_type = L"HTML";

	// default op
	CString option = L"VARADA via com2"; // VIEW NAME
	pObj->GwaCommand("SET, OP_VIEW.6,\"[EL1DFRCAUTOPTS, STRIPEYOUTPUT, REFSBYNAME, INTRESONTRUSS, FULLPOPFIELDS, ]\"	0	VARADA via com2	128	128	1607	821	1	9	1	4	1	-4	1	1	-4	2	1	-4	3	1	-4	11	1	-4	1	10004000	0	-10	N	1	m	1	kg	1	s	1	°C	1	m	1	Pa	1	m/s²	1	m	1	-Infinity	Infinity	0	1	4	2	4	1e-12	1	3	0	2");
	//CString node_comm = L"SET, NODE, 25, 10, 4.5, 8";
	//pObj->GwaCommand((LPCTSTR)node_comm);

	// custom op
	//CString headers = L"\"[FULLPOPFIELDS,CUSTOMOUTPUT:FORCE_MEMB_MZZ;FORCE_MEMB_MYY;FORCE_MEMB_FX;]\"";
	//CString strdata(",CustomOp from COM2,0,0,0,0,0,0,0,4,1,1,0,0,0,0,3,3,1,-1,25,0,0,0,-10,N,1,m,1,kg,1,s,1,°C,1,m,1,Pa,1,m/s²,1,m,1,0,0,0,1,4,2,4,1e-12,0,0,0,0");
	//CString command = (L"SET,CUSTOM_OP_VIEW," + headers + strdata);

	//CString alter_command = L"[FULLPOPFIELDS]	1	1D member	\"[CUSTOMOUTPUT:ENTITY_CASE_EXPANSION; ]\"	CustomOpCOM	0	0	0	0	0	0	0	4	1	-4	0	0	0	0	3	1	-4	0	0	0	-10	N	1	m	1	kg	1	s	1	°C	1	m	1	Pa	1	m/s²	1	m	1	0	0	0	1	4	2	4	1e-12	0	0	0	0";
	//alter_command = L"ADD,CUSTOM_OP_VIEW," + alter_command;



	//CString headers = L"\"[FULLPOPFIELDS,CUSTOMOUTPUT:FORCE_MEMB_MZZ;FORCE_MEMB_MYY;FORCE_MEMB_FX;]\"";
	//CString strdata(",CustomOp from COM,0,0,0,0,0,0,0,4,1,1,0,0,0,0,3,3,1,-1,25,0,0,0,-10,N,1,m,1,kg,1,s,1,°C,1,m,1,Pa,1,m/s²,1,m,1,0,0,0,1,4,2,4,1e-12,0,0,0,0");
	//CString command = L"SET,CUSTOM_OP_VIEW," + headers + strdata;

	//CString command = L"SET, CUSTOM_OP_VIEW.7	7	[FULLPOPFIELDS]	1	1D member	\"[CUSTOMOUTPUT:MEMB_NAME;MEMB_TYPE;]\"	test	0	0	0	0	0	0	0	4	2	-46	1	0	0	0	0	3	3	1	-1	10	0	0	0	-10	N	1	m	1	kg	1	s	1	°C	1	m	1	Pa	1	m/s²	1	m	1	0	0	0	1	4	2	4	1e-12	0	0	0	0";

	//pObj->GwaCommand((LPCTSTR)command);

	// gsa10.2.2.33 version
	//CString headers = L"\"[FULLPOPFIELDS, CUSTOMOUTPUT:MEMB_TYPE; MEMB_NAME; DISP_MEMB_DY; ]\"	viaCOM	0	0	0	0	0	0	0	4	2	-46	1	0	0	0	0	3	3	1	-1	10	0	0	0	-10	N	1	m	1	kg	1	s	1	°C	1	m	1	Pa	1	m/s²	1	m	1	0	0	0	1	4	2	4	1e-12	0	0	0	0";
	//CString comm = (L"SET,CUSTOM_OP_VIEW," + headers);
	//pObj->GwaCommand((LPCTSTR)comm);


	//UNREFERENCED_PARAMETER(RET);
	//CString option = L"testerCOM"; // VIEW NAME
	//CString option = L"testerCOM";
	//save to gwb file
	//return (pObj->SaveViewToFile((LPCTSTR)option, (LPCTSTR)file_type)); // adds a new view
	return 0;

}
void CGsaComClientApp::invokeGsa(CString filename, CString analysed_filename, CString analysed_filename_report)
{
	if (FAILED(CoInitializeEx(0, COINIT_APARTMENTTHREADED)))
		return;

	IComAutoPtr pObj(__uuidof(ComAuto));
	short ret_code = 0;

	_bstr_t bsFileName = (LPCTSTR)filename;
	ret_code = pObj->Open(bsFileName);
	if (ret_code == 1)
		return;

	_bstr_t bsContent(_T("RESULTS"));
	ret_code = pObj->Delete(bsContent);
	ASSERT(ret_code != 1);	// file not open!

	_variant_t vCase(0L);
	ret_code = pObj->Analyse(vCase);
	ASSERT(ret_code == 0);
	CFile theFile;
	try
	{
		theFile.Open(analysed_filename_report, CFile::modeCreate | CFile::modeWrite);

		//NODES-NOT WORKING
		//CString node_comm = L"ADD, LOAD_BEAM, 1 to 10, 10, GLOBAL, NO, Z, -1000";
		//pObj->GwaCommand((LPCTSTR)node_comm);


		// default 10.2.3.49 works!!!!
		//CString comm = L"SET,OP_VIEW,\"[EL1DFRCAUTOPTS,STRIPEYOUTPUT,REFSBYNAME,INTRESONTRUSS,FULLPOPFIELDS]\",0,viewFromCOM,128,128,1607,821,1,9,1,4,1,-4,1,1,-4,2,1,-4,3,1,-4,11,1,-4,1,10004000,0,-10,N,1,m,1,kg,1,s,1,°C,1,m,1,Pa,1,m/s²,1,m,1,-Infinity,Infinity,0,1,4,2,4,1e-12,1,3,0,2";
		//pObj->GwaCommand((LPCTSTR)comm);

		//CUSTOM OP 10.2.1.32 works!!!
		CString comm(L"SET,CUSTOM_OP_VIEW,[FULLPOPFIELDS],1,Node,\"[CUSTOMOUTPUT:DISP_DY;DISP_DX;]\",nodesCOM,0,0,0,0,0,0,0,4,2,-46,1,1,3,1,-1,10,0,0,0,0,0,0,0,-10,N,1,m,1,kg,1,s,1,°C,1,m,1,Pa,1,m/s²,1,m,1,0,0,0,1,4,2,4,1e-12,0,0,0,0");
		pObj->GwaCommand((LPCTSTR)comm);
		
		theFile.Close();
	}
	catch (...)
	{
		theFile.Abort();
	}
	analysed_filename.Replace(L".gwb", L".gwa");
	_bstr_t bsAnalysedFileName = (LPCTSTR)analysed_filename;
	ret_code = pObj->SaveAs(bsAnalysedFileName);
	ASSERT(ret_code == 0);
	pObj->Close();

}
void CGsaComClientApp::WriteString(CFile& cfile, CString str)
{
	CString cstr = _T("\r\n");
	cfile.Write((LPCTSTR)str, str.GetLength() * sizeof(TCHAR));
	cfile.Write((LPCTSTR)cstr, cstr.GetLength() * sizeof(TCHAR));
}

