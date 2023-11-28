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
int NewMethods(IComAutoPtr& pObj)
{
	_bstr_t op = "nodesCOM";
	return pObj->SaveViewToFile(op, "CSV");

	// Other methods
	//auto max = pObj->HighestView(op);
	//*/

	/*BSTR op = L"COV";
	long place_holder = 0;
	long* nw_vw_id = &place_holder;

	//auto ret = pObj->CopyView(op, 1, nw_vw_id); 
//max = pObj->HighestView(op); // updated highest view

//ret = pObj->RenameView(op, 2, BSTR("renamedCOM"));
//ret=pObj->DeleteView(op, 3);
//BSTR* caselist = new BSTR();
//pObj->GetViewCaseList(op, 1, caselist);
//auto name = pObj->ViewName(op, 2);
//auto iref = pObj->ViewRefFromName(op, BSTR("nodesCOM"));
	/*_bstr_t nam = "nodesCOM";
	pObj->PrintView(nam);*/
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


		auto res = NewMethods(pObj);
		theFile.Close();
	}
	catch (...)
	{
		theFile.Abort();
	}
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

void WorkingGwaCommands()
{

	//NODES-NOT WORKING
//CString node_comm = L"ADD, LOAD_BEAM, 1 to 10, 10, GLOBAL, NO, Z, -1000";
//pObj->GwaCommand((LPCTSTR)node_comm);


// default 10.2.3.49 works!!!!
//CString comm = L"SET,OP_VIEW,\"[EL1DFRCAUTOPTS,STRIPEYOUTPUT,REFSBYNAME,INTRESONTRUSS,FULLPOPFIELDS]\",0,viewFromCOM,128,128,1607,821,1,9,1,4,1,-4,1,1,-4,2,1,-4,3,1,-4,11,1,-4,1,10004000,0,-10,N,1,m,1,kg,1,s,1,°C,1,m,1,Pa,1,m/s²,1,m,1,-Infinity,Infinity,0,1,4,2,4,1e-12,1,3,0,2";
//pObj->GwaCommand((LPCTSTR)comm);

//CUSTOM OP 10.2.3.49 works!!!
//CString comm(L"SET,CUSTOM_OP_VIEW,[FULLPOPFIELDS],1,Node,\"[CUSTOMOUTPUT:DISP_DY;DISP_DX;]\",NodesComToUI,0,0,0,0,0,0,0,4,2,-46,1,1,3,1,-1,10,0,0,0,0,0,0,0,-10,N,1,m,1,kg,1,s,1,°C,1,m,1,Pa,1,m/s²,1,m,1,0,0,0,1,4,2,4,1e-12,0,0,0,0");
//pObj->GwaCommand((LPCTSTR)comm);

}
