// # Package: RDCOMClient
// # Version: 0.93-0.2
// # Title: R-DCOM Client
// # Author: Duncan Temple Lang <duncan@wald.ucdavis.edu>
// #     Maintainer: Duncan Temple Lang <duncan@wald.ucdavis.edu>
// #     Description: Provides dynamic client-side access to (D)COM applications from within R.
// # License: GPL-2
// # Collate: classes.R COMLists.S COMError.R com.R debug.S zzz.R runTime.S
// # URL: http://www.omegahat.net/RDCOMClient, http://www.omegahat.net
// # http://www.omegahat.net/bugs
// Some parts of code by https://github.com/jototland/ jototland@gmail.com

#include "RCOMObject.h"
#include <windows.h>
#include <oleauto.h>
#include <stdio.h>  /* sprintf() */

#include <tchar.h>

extern "C" int RDCOM_WriteErrors;
int RDCOM_WriteErrors = 1;

extern "C"
SEXP
RDCOM_setWriteError(SEXP value)
{
    int tmp = RDCOM_WriteErrors;
    RDCOM_WriteErrors = asLogical(value);
    return(ScalarLogical(tmp));
}

extern "C"
SEXP
RDCOM_getWriteError(SEXP value)
{
    return(ScalarLogical(RDCOM_WriteErrors));
}



FILE *
getErrorFILE()
{
  static FILE *f = NULL;

  if (f)
    return f;

  TCHAR path[MAX_PATH];
  DWORD result;

  result = GetTempPath(MAX_PATH, path);

  if (result > MAX_PATH-10 || result == 0) {
    f = stderr;
  } else {
    lstrcat(path, _T("RDCOM.err"));
    f = fopen(path, "a");
    if (!f) {
      f = stderr;
    }
  }

  return(f);
}

extern "C" {
SEXP R_createCOMErrorCodes();
}

/* Taken from ErrorUtils.cpp in PyWin32 distribution. */
#include "oaidl.h"


	struct HRESULT_ENTRY
	{
		HRESULT hr;
		LPCTSTR lpszName;
	};
	#define MAKE_HRESULT_ENTRY(hr)    { hr, (#hr) }
	static const HRESULT_ENTRY hrNameTable[] =
	{
		MAKE_HRESULT_ENTRY(S_OK),
		MAKE_HRESULT_ENTRY(S_FALSE),

		MAKE_HRESULT_ENTRY(CACHE_S_FORMATETC_NOTSUPPORTED),
		MAKE_HRESULT_ENTRY(CACHE_S_SAMECACHE),
		MAKE_HRESULT_ENTRY(CACHE_S_SOMECACHES_NOTUPDATED),
		MAKE_HRESULT_ENTRY(CONVERT10_S_NO_PRESENTATION),
		MAKE_HRESULT_ENTRY(DATA_S_SAMEFORMATETC),
		MAKE_HRESULT_ENTRY(DRAGDROP_S_CANCEL),
		MAKE_HRESULT_ENTRY(DRAGDROP_S_DROP),
		MAKE_HRESULT_ENTRY(DRAGDROP_S_USEDEFAULTCURSORS),
		MAKE_HRESULT_ENTRY(INPLACE_S_TRUNCATED),
		MAKE_HRESULT_ENTRY(MK_S_HIM),
		MAKE_HRESULT_ENTRY(MK_S_ME),
		MAKE_HRESULT_ENTRY(MK_S_MONIKERALREADYREGISTERED),
		MAKE_HRESULT_ENTRY(MK_S_REDUCED_TO_SELF),
		MAKE_HRESULT_ENTRY(MK_S_US),
		MAKE_HRESULT_ENTRY(OLE_S_MAC_CLIPFORMAT),
		MAKE_HRESULT_ENTRY(OLE_S_STATIC),
		MAKE_HRESULT_ENTRY(OLE_S_USEREG),
		MAKE_HRESULT_ENTRY(OLEOBJ_S_CANNOT_DOVERB_NOW),
		MAKE_HRESULT_ENTRY(OLEOBJ_S_INVALIDHWND),
		MAKE_HRESULT_ENTRY(OLEOBJ_S_INVALIDVERB),
		MAKE_HRESULT_ENTRY(OLEOBJ_S_LAST),
		MAKE_HRESULT_ENTRY(STG_S_CONVERTED),
		MAKE_HRESULT_ENTRY(VIEW_S_ALREADY_FROZEN),

		MAKE_HRESULT_ENTRY(E_UNEXPECTED),
		MAKE_HRESULT_ENTRY(E_NOTIMPL),
		MAKE_HRESULT_ENTRY(E_OUTOFMEMORY),
		MAKE_HRESULT_ENTRY(E_INVALIDARG),
		MAKE_HRESULT_ENTRY(E_NOINTERFACE),
		MAKE_HRESULT_ENTRY(E_POINTER),
		MAKE_HRESULT_ENTRY(E_HANDLE),
		MAKE_HRESULT_ENTRY(E_ABORT),
		MAKE_HRESULT_ENTRY(E_FAIL),
		MAKE_HRESULT_ENTRY(E_ACCESSDENIED),

		MAKE_HRESULT_ENTRY(CACHE_E_NOCACHE_UPDATED),
		MAKE_HRESULT_ENTRY(CLASS_E_CLASSNOTAVAILABLE),
		MAKE_HRESULT_ENTRY(CLASS_E_NOAGGREGATION),
		MAKE_HRESULT_ENTRY(CLIPBRD_E_BAD_DATA),
		MAKE_HRESULT_ENTRY(CLIPBRD_E_CANT_CLOSE),
		MAKE_HRESULT_ENTRY(CLIPBRD_E_CANT_EMPTY),
		MAKE_HRESULT_ENTRY(CLIPBRD_E_CANT_OPEN),
		MAKE_HRESULT_ENTRY(CLIPBRD_E_CANT_SET),
		MAKE_HRESULT_ENTRY(CO_E_ALREADYINITIALIZED),
		MAKE_HRESULT_ENTRY(CO_E_APPDIDNTREG),
		MAKE_HRESULT_ENTRY(CO_E_APPNOTFOUND),
		MAKE_HRESULT_ENTRY(CO_E_APPSINGLEUSE),
		MAKE_HRESULT_ENTRY(CO_E_BAD_PATH),
		MAKE_HRESULT_ENTRY(CO_E_CANTDETERMINECLASS),
		MAKE_HRESULT_ENTRY(CO_E_CLASS_CREATE_FAILED),
		MAKE_HRESULT_ENTRY(CO_E_CLASSSTRING),
		MAKE_HRESULT_ENTRY(CO_E_DLLNOTFOUND),
		MAKE_HRESULT_ENTRY(CO_E_ERRORINAPP),
		MAKE_HRESULT_ENTRY(CO_E_ERRORINDLL),
		MAKE_HRESULT_ENTRY(CO_E_IIDSTRING),
		MAKE_HRESULT_ENTRY(CO_E_NOTINITIALIZED),
		MAKE_HRESULT_ENTRY(CO_E_OBJISREG),
		MAKE_HRESULT_ENTRY(CO_E_OBJNOTCONNECTED),
		MAKE_HRESULT_ENTRY(CO_E_OBJNOTREG),
		MAKE_HRESULT_ENTRY(CO_E_OBJSRV_RPC_FAILURE),
		MAKE_HRESULT_ENTRY(CO_E_SCM_ERROR),
		MAKE_HRESULT_ENTRY(CO_E_SCM_RPC_FAILURE),
		MAKE_HRESULT_ENTRY(CO_E_SERVER_EXEC_FAILURE),
		MAKE_HRESULT_ENTRY(CO_E_SERVER_STOPPING),
		MAKE_HRESULT_ENTRY(CO_E_WRONGOSFORAPP),
		MAKE_HRESULT_ENTRY(CONVERT10_E_OLESTREAM_BITMAP_TO_DIB),
		MAKE_HRESULT_ENTRY(CONVERT10_E_OLESTREAM_FMT),
		MAKE_HRESULT_ENTRY(CONVERT10_E_OLESTREAM_GET),
		MAKE_HRESULT_ENTRY(CONVERT10_E_OLESTREAM_PUT),
		MAKE_HRESULT_ENTRY(CONVERT10_E_STG_DIB_TO_BITMAP),
		MAKE_HRESULT_ENTRY(CONVERT10_E_STG_FMT),
		MAKE_HRESULT_ENTRY(CONVERT10_E_STG_NO_STD_STREAM),
		MAKE_HRESULT_ENTRY(DISP_E_ARRAYISLOCKED),
		MAKE_HRESULT_ENTRY(DISP_E_BADCALLEE),
		MAKE_HRESULT_ENTRY(DISP_E_BADINDEX),
		MAKE_HRESULT_ENTRY(DISP_E_BADPARAMCOUNT),
		MAKE_HRESULT_ENTRY(DISP_E_BADVARTYPE),
		MAKE_HRESULT_ENTRY(DISP_E_EXCEPTION),
		MAKE_HRESULT_ENTRY(DISP_E_MEMBERNOTFOUND),
		MAKE_HRESULT_ENTRY(DISP_E_NONAMEDARGS),
		MAKE_HRESULT_ENTRY(DISP_E_NOTACOLLECTION),
		MAKE_HRESULT_ENTRY(DISP_E_OVERFLOW),
		MAKE_HRESULT_ENTRY(DISP_E_PARAMNOTFOUND),
		MAKE_HRESULT_ENTRY(DISP_E_PARAMNOTOPTIONAL),
		MAKE_HRESULT_ENTRY(DISP_E_TYPEMISMATCH),
		MAKE_HRESULT_ENTRY(DISP_E_UNKNOWNINTERFACE),
		MAKE_HRESULT_ENTRY(DISP_E_UNKNOWNLCID),
		MAKE_HRESULT_ENTRY(DISP_E_UNKNOWNNAME),
		MAKE_HRESULT_ENTRY(DRAGDROP_E_ALREADYREGISTERED),
		MAKE_HRESULT_ENTRY(DRAGDROP_E_INVALIDHWND),
		MAKE_HRESULT_ENTRY(DRAGDROP_E_NOTREGISTERED),
		MAKE_HRESULT_ENTRY(DV_E_CLIPFORMAT),
		MAKE_HRESULT_ENTRY(DV_E_DVASPECT),
		MAKE_HRESULT_ENTRY(DV_E_DVTARGETDEVICE),
		MAKE_HRESULT_ENTRY(DV_E_DVTARGETDEVICE_SIZE),
		MAKE_HRESULT_ENTRY(DV_E_FORMATETC),
		MAKE_HRESULT_ENTRY(DV_E_LINDEX),
		MAKE_HRESULT_ENTRY(DV_E_NOIVIEWOBJECT),
		MAKE_HRESULT_ENTRY(DV_E_STATDATA),
		MAKE_HRESULT_ENTRY(DV_E_STGMEDIUM),
		MAKE_HRESULT_ENTRY(DV_E_TYMED),
		MAKE_HRESULT_ENTRY(INPLACE_E_NOTOOLSPACE),
		MAKE_HRESULT_ENTRY(INPLACE_E_NOTUNDOABLE),
		MAKE_HRESULT_ENTRY(MEM_E_INVALID_LINK),
		MAKE_HRESULT_ENTRY(MEM_E_INVALID_ROOT),
		MAKE_HRESULT_ENTRY(MEM_E_INVALID_SIZE),
		MAKE_HRESULT_ENTRY(MK_E_CANTOPENFILE),
		MAKE_HRESULT_ENTRY(MK_E_CONNECTMANUALLY),
		MAKE_HRESULT_ENTRY(MK_E_ENUMERATION_FAILED),
		MAKE_HRESULT_ENTRY(MK_E_EXCEEDEDDEADLINE),
		MAKE_HRESULT_ENTRY(MK_E_INTERMEDIATEINTERFACENOTSUPPORTED),
		MAKE_HRESULT_ENTRY(MK_E_INVALIDEXTENSION),
		MAKE_HRESULT_ENTRY(MK_E_MUSTBOTHERUSER),
		MAKE_HRESULT_ENTRY(MK_E_NEEDGENERIC),
		MAKE_HRESULT_ENTRY(MK_E_NO_NORMALIZED),
		MAKE_HRESULT_ENTRY(MK_E_NOINVERSE),
		MAKE_HRESULT_ENTRY(MK_E_NOOBJECT),
		MAKE_HRESULT_ENTRY(MK_E_NOPREFIX),
		MAKE_HRESULT_ENTRY(MK_E_NOSTORAGE),
		MAKE_HRESULT_ENTRY(MK_E_NOTBINDABLE),
		MAKE_HRESULT_ENTRY(MK_E_NOTBOUND),
		MAKE_HRESULT_ENTRY(MK_E_SYNTAX),
		MAKE_HRESULT_ENTRY(MK_E_UNAVAILABLE),
		MAKE_HRESULT_ENTRY(OLE_E_ADVF),
		MAKE_HRESULT_ENTRY(OLE_E_ADVISENOTSUPPORTED),
		MAKE_HRESULT_ENTRY(OLE_E_BLANK),
		MAKE_HRESULT_ENTRY(OLE_E_CANT_BINDTOSOURCE),
		MAKE_HRESULT_ENTRY(OLE_E_CANT_GETMONIKER),
		MAKE_HRESULT_ENTRY(OLE_E_CANTCONVERT),
		MAKE_HRESULT_ENTRY(OLE_E_CLASSDIFF),
		MAKE_HRESULT_ENTRY(OLE_E_ENUM_NOMORE),
		MAKE_HRESULT_ENTRY(OLE_E_INVALIDHWND),
		MAKE_HRESULT_ENTRY(OLE_E_INVALIDRECT),
		MAKE_HRESULT_ENTRY(OLE_E_NOCACHE),
		MAKE_HRESULT_ENTRY(OLE_E_NOCONNECTION),
		MAKE_HRESULT_ENTRY(OLE_E_NOSTORAGE),
		MAKE_HRESULT_ENTRY(OLE_E_NOT_INPLACEACTIVE),
		MAKE_HRESULT_ENTRY(OLE_E_NOTRUNNING),
		MAKE_HRESULT_ENTRY(OLE_E_OLEVERB),
		MAKE_HRESULT_ENTRY(OLE_E_PROMPTSAVECANCELLED),
		MAKE_HRESULT_ENTRY(OLE_E_STATIC),
		MAKE_HRESULT_ENTRY(OLE_E_WRONGCOMPOBJ),
		MAKE_HRESULT_ENTRY(OLEOBJ_E_INVALIDVERB),
		MAKE_HRESULT_ENTRY(OLEOBJ_E_NOVERBS),
		MAKE_HRESULT_ENTRY(REGDB_E_CLASSNOTREG),
		MAKE_HRESULT_ENTRY(REGDB_E_IIDNOTREG),
		MAKE_HRESULT_ENTRY(REGDB_E_INVALIDVALUE),
		MAKE_HRESULT_ENTRY(REGDB_E_KEYMISSING),
		MAKE_HRESULT_ENTRY(REGDB_E_READREGDB),
		MAKE_HRESULT_ENTRY(REGDB_E_WRITEREGDB),
		MAKE_HRESULT_ENTRY(RPC_E_ATTEMPTED_MULTITHREAD),
		MAKE_HRESULT_ENTRY(RPC_E_CALL_CANCELED),
		MAKE_HRESULT_ENTRY(RPC_E_CALL_REJECTED),
		MAKE_HRESULT_ENTRY(RPC_E_CANTCALLOUT_AGAIN),
		MAKE_HRESULT_ENTRY(RPC_E_CANTCALLOUT_INASYNCCALL),
		MAKE_HRESULT_ENTRY(RPC_E_CANTCALLOUT_INEXTERNALCALL),
		MAKE_HRESULT_ENTRY(RPC_E_CANTCALLOUT_ININPUTSYNCCALL),
		MAKE_HRESULT_ENTRY(RPC_E_CANTPOST_INSENDCALL),
		MAKE_HRESULT_ENTRY(RPC_E_CANTTRANSMIT_CALL),
		MAKE_HRESULT_ENTRY(RPC_E_CHANGED_MODE),
		MAKE_HRESULT_ENTRY(RPC_E_CLIENT_CANTMARSHAL_DATA),
		MAKE_HRESULT_ENTRY(RPC_E_CLIENT_CANTUNMARSHAL_DATA),
		MAKE_HRESULT_ENTRY(RPC_E_CLIENT_DIED),
		MAKE_HRESULT_ENTRY(RPC_E_CONNECTION_TERMINATED),
		MAKE_HRESULT_ENTRY(RPC_E_DISCONNECTED),
		MAKE_HRESULT_ENTRY(RPC_E_FAULT),
		MAKE_HRESULT_ENTRY(RPC_E_INVALID_CALLDATA),
		MAKE_HRESULT_ENTRY(RPC_E_INVALID_DATA),
		MAKE_HRESULT_ENTRY(RPC_E_INVALID_DATAPACKET),
		MAKE_HRESULT_ENTRY(RPC_E_INVALID_PARAMETER),
		MAKE_HRESULT_ENTRY(RPC_E_INVALIDMETHOD),
		MAKE_HRESULT_ENTRY(RPC_E_NOT_REGISTERED),
		MAKE_HRESULT_ENTRY(RPC_E_OUT_OF_RESOURCES),
		MAKE_HRESULT_ENTRY(RPC_E_RETRY),
		MAKE_HRESULT_ENTRY(RPC_E_SERVER_CANTMARSHAL_DATA),
		MAKE_HRESULT_ENTRY(RPC_E_SERVER_CANTUNMARSHAL_DATA),
		MAKE_HRESULT_ENTRY(RPC_E_SERVER_DIED),
		MAKE_HRESULT_ENTRY(RPC_E_SERVER_DIED_DNE),
		MAKE_HRESULT_ENTRY(RPC_E_SERVERCALL_REJECTED),
		MAKE_HRESULT_ENTRY(RPC_E_SERVERCALL_RETRYLATER),
		MAKE_HRESULT_ENTRY(RPC_E_SERVERFAULT),
		MAKE_HRESULT_ENTRY(RPC_E_SYS_CALL_FAILED),
		MAKE_HRESULT_ENTRY(RPC_E_THREAD_NOT_INIT),
		MAKE_HRESULT_ENTRY(RPC_E_UNEXPECTED),
		MAKE_HRESULT_ENTRY(RPC_E_WRONG_THREAD),
		MAKE_HRESULT_ENTRY(STG_E_ABNORMALAPIEXIT),
		MAKE_HRESULT_ENTRY(STG_E_ACCESSDENIED),
		MAKE_HRESULT_ENTRY(STG_E_CANTSAVE),
		MAKE_HRESULT_ENTRY(STG_E_DISKISWRITEPROTECTED),
		MAKE_HRESULT_ENTRY(STG_E_EXTANTMARSHALLINGS),
		MAKE_HRESULT_ENTRY(STG_E_FILEALREADYEXISTS),
		MAKE_HRESULT_ENTRY(STG_E_FILENOTFOUND),
		MAKE_HRESULT_ENTRY(STG_E_INSUFFICIENTMEMORY),
		MAKE_HRESULT_ENTRY(STG_E_INUSE),
		MAKE_HRESULT_ENTRY(STG_E_INVALIDFLAG),
		MAKE_HRESULT_ENTRY(STG_E_INVALIDFUNCTION),
		MAKE_HRESULT_ENTRY(STG_E_INVALIDHANDLE),
		MAKE_HRESULT_ENTRY(STG_E_INVALIDHEADER),
		MAKE_HRESULT_ENTRY(STG_E_INVALIDNAME),
		MAKE_HRESULT_ENTRY(STG_E_INVALIDPARAMETER),
		MAKE_HRESULT_ENTRY(STG_E_INVALIDPOINTER),
		MAKE_HRESULT_ENTRY(STG_E_LOCKVIOLATION),
		MAKE_HRESULT_ENTRY(STG_E_MEDIUMFULL),
		MAKE_HRESULT_ENTRY(STG_E_NOMOREFILES),
		MAKE_HRESULT_ENTRY(STG_E_NOTCURRENT),
		MAKE_HRESULT_ENTRY(STG_E_NOTFILEBASEDSTORAGE),
		MAKE_HRESULT_ENTRY(STG_E_OLDDLL),
		MAKE_HRESULT_ENTRY(STG_E_OLDFORMAT),
		MAKE_HRESULT_ENTRY(STG_E_PATHNOTFOUND),
		MAKE_HRESULT_ENTRY(STG_E_READFAULT),
		MAKE_HRESULT_ENTRY(STG_E_REVERTED),
		MAKE_HRESULT_ENTRY(STG_E_SEEKERROR),
		MAKE_HRESULT_ENTRY(STG_E_SHAREREQUIRED),
		MAKE_HRESULT_ENTRY(STG_E_SHAREVIOLATION),
		MAKE_HRESULT_ENTRY(STG_E_TOOMANYOPENFILES),
		MAKE_HRESULT_ENTRY(STG_E_UNIMPLEMENTEDFUNCTION),
		MAKE_HRESULT_ENTRY(STG_E_UNKNOWN),
		MAKE_HRESULT_ENTRY(STG_E_WRITEFAULT),
		MAKE_HRESULT_ENTRY(TYPE_E_AMBIGUOUSNAME),
		MAKE_HRESULT_ENTRY(TYPE_E_BADMODULEKIND),
		MAKE_HRESULT_ENTRY(TYPE_E_BUFFERTOOSMALL),
		MAKE_HRESULT_ENTRY(TYPE_E_CANTCREATETMPFILE),
		MAKE_HRESULT_ENTRY(TYPE_E_CANTLOADLIBRARY),
		MAKE_HRESULT_ENTRY(TYPE_E_CIRCULARTYPE),
		MAKE_HRESULT_ENTRY(TYPE_E_DLLFUNCTIONNOTFOUND),
		MAKE_HRESULT_ENTRY(TYPE_E_DUPLICATEID),
		MAKE_HRESULT_ENTRY(TYPE_E_ELEMENTNOTFOUND),
		MAKE_HRESULT_ENTRY(TYPE_E_INCONSISTENTPROPFUNCS),
		MAKE_HRESULT_ENTRY(TYPE_E_INVALIDSTATE),
		MAKE_HRESULT_ENTRY(TYPE_E_INVDATAREAD),
		MAKE_HRESULT_ENTRY(TYPE_E_IOERROR),
		MAKE_HRESULT_ENTRY(TYPE_E_LIBNOTREGISTERED),
		MAKE_HRESULT_ENTRY(TYPE_E_NAMECONFLICT),
		MAKE_HRESULT_ENTRY(TYPE_E_OUTOFBOUNDS),
		MAKE_HRESULT_ENTRY(TYPE_E_QUALIFIEDNAMEDISALLOWED),
		MAKE_HRESULT_ENTRY(TYPE_E_REGISTRYACCESS),
		MAKE_HRESULT_ENTRY(TYPE_E_SIZETOOBIG),
		MAKE_HRESULT_ENTRY(TYPE_E_TYPEMISMATCH),
		MAKE_HRESULT_ENTRY(TYPE_E_UNDEFINEDTYPE),
		MAKE_HRESULT_ENTRY(TYPE_E_UNKNOWNLCID),
		MAKE_HRESULT_ENTRY(TYPE_E_UNSUPFORMAT),
		MAKE_HRESULT_ENTRY(TYPE_E_WRONGTYPEKIND),
		MAKE_HRESULT_ENTRY(VIEW_E_DRAW),

#if NOT_AVAILABLE
		MAKE_HRESULT_ENTRY(CONNECT_E_NOCONNECTION),
		MAKE_HRESULT_ENTRY(CONNECT_E_ADVISELIMIT),
		MAKE_HRESULT_ENTRY(CONNECT_E_CANNOTCONNECT),
		MAKE_HRESULT_ENTRY(CONNECT_E_OVERRIDDEN),
#endif

#ifndef NO_PYCOM_IPROVIDECLASSINFO
		MAKE_HRESULT_ENTRY(CLASS_E_NOTLICENSED),
		MAKE_HRESULT_ENTRY(CLASS_E_NOAGGREGATION),
		MAKE_HRESULT_ENTRY(CLASS_E_CLASSNOTAVAILABLE),
#endif // NO_PYCOM_IPROVIDECLASSINFO

#ifndef MS_WINCE // ??
#if AVAILABLE
		MAKE_HRESULT_ENTRY(CTL_E_ILLEGALFUNCTIONCALL      ),
		MAKE_HRESULT_ENTRY(CTL_E_OVERFLOW                 ),
		MAKE_HRESULT_ENTRY(CTL_E_OUTOFMEMORY              ),
		MAKE_HRESULT_ENTRY(CTL_E_DIVISIONBYZERO           ),
		MAKE_HRESULT_ENTRY(CTL_E_OUTOFSTRINGSPACE         ),
		MAKE_HRESULT_ENTRY(CTL_E_OUTOFSTACKSPACE          ),
		MAKE_HRESULT_ENTRY(CTL_E_BADFILENAMEORNUMBER      ),
		MAKE_HRESULT_ENTRY(CTL_E_FILENOTFOUND             ),
		MAKE_HRESULT_ENTRY(CTL_E_BADFILEMODE              ),
		MAKE_HRESULT_ENTRY(CTL_E_FILEALREADYOPEN          ),
		MAKE_HRESULT_ENTRY(CTL_E_DEVICEIOERROR            ),
		MAKE_HRESULT_ENTRY(CTL_E_FILEALREADYEXISTS        ),
		MAKE_HRESULT_ENTRY(CTL_E_BADRECORDLENGTH          ),
		MAKE_HRESULT_ENTRY(CTL_E_DISKFULL                 ),
		MAKE_HRESULT_ENTRY(CTL_E_BADRECORDNUMBER          ),
		MAKE_HRESULT_ENTRY(CTL_E_BADFILENAME              ),
		MAKE_HRESULT_ENTRY(CTL_E_TOOMANYFILES             ),
		MAKE_HRESULT_ENTRY(CTL_E_DEVICEUNAVAILABLE        ),
		MAKE_HRESULT_ENTRY(CTL_E_PERMISSIONDENIED         ),
		MAKE_HRESULT_ENTRY(CTL_E_DISKNOTREADY             ),
		MAKE_HRESULT_ENTRY(CTL_E_PATHFILEACCESSERROR      ),
		MAKE_HRESULT_ENTRY(CTL_E_PATHNOTFOUND             ),
		MAKE_HRESULT_ENTRY(CTL_E_INVALIDPATTERNSTRING     ),
		MAKE_HRESULT_ENTRY(CTL_E_INVALIDUSEOFNULL         ),
		MAKE_HRESULT_ENTRY(CTL_E_INVALIDFILEFORMAT        ),
		MAKE_HRESULT_ENTRY(CTL_E_INVALIDPROPERTYVALUE     ),
		MAKE_HRESULT_ENTRY(CTL_E_INVALIDPROPERTYARRAYINDEX),
		MAKE_HRESULT_ENTRY(CTL_E_SETNOTSUPPORTEDATRUNTIME ),
		MAKE_HRESULT_ENTRY(CTL_E_SETNOTSUPPORTED          ),
		MAKE_HRESULT_ENTRY(CTL_E_NEEDPROPERTYARRAYINDEX   ),
		MAKE_HRESULT_ENTRY(CTL_E_SETNOTPERMITTED          ),
		MAKE_HRESULT_ENTRY(CTL_E_GETNOTSUPPORTEDATRUNTIME ),
		MAKE_HRESULT_ENTRY(CTL_E_GETNOTSUPPORTED          ),
		MAKE_HRESULT_ENTRY(CTL_E_PROPERTYNOTFOUND         ),
		MAKE_HRESULT_ENTRY(CTL_E_INVALIDCLIPBOARDFORMAT   ),
		MAKE_HRESULT_ENTRY(CTL_E_INVALIDPICTURE           ),
		MAKE_HRESULT_ENTRY(CTL_E_PRINTERERROR             ),
		MAKE_HRESULT_ENTRY(CTL_E_CANTSAVEFILETOTEMP       ),
		MAKE_HRESULT_ENTRY(CTL_E_SEARCHTEXTNOTFOUND       ),
		MAKE_HRESULT_ENTRY(CTL_E_REPLACEMENTSTOOLONG      ),
#endif
#endif // MS_WINCE
	};
	#undef MAKE_HRESULT_ENTRY


#ifndef _countof
#define _countof(array) (sizeof(array)/sizeof(array[0]))
#endif
void GetScodeString(HRESULT hr, LPTSTR buf, int bufSize)
{
	// first ask the OS to give it to us..
	// ### should we get the Unicode version instead?
	int numCopied = ::FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, NULL, hr, 0, buf, bufSize, NULL );
	if (numCopied>0) {
		if (numCopied<bufSize) {
			// trim trailing crap
			if (numCopied>2 && (buf[numCopied-2]=='\n'||buf[numCopied-2]=='\r'))
				buf[numCopied-2] = '\0';
		}
		return;
	}

	// else look for it in the table
	for (unsigned int i = 0; i < _countof(hrNameTable); i++)
	{
		if (hr == hrNameTable[i].hr) {
			strncpy(buf, hrNameTable[i].lpszName, bufSize);
			return;
		}
	}
	// not found - make one up
	sprintf(buf, ("OLE error 0x%08lx"), hr);
}




void
COMError(HRESULT hr)
{
    TCHAR buf[512];
    GetScodeString(hr, buf, sizeof(buf)/sizeof(buf[0]));
    /*
    PROBLEM buf
    ERROR;
    */
    SEXP e;
    PROTECT(e = allocVector(LANGSXP, 3));
    SETCAR(e, Rf_install("COMStop"));
    SETCAR(CDR(e), mkString(buf));
    SETCAR(CDR(CDR(e)), ScalarInteger(hr));
    Rf_eval(e, R_GlobalEnv);
    UNPROTECT(1); /* Won't come back to here. */
}




/* Determines whether we can use the error information from the
   source object and if so, throws that as an error.
   If serr is non-NULL, then the error is not thrown in R
   but a COMSErrorInfo object is returned with the information in it.
*/
HRESULT
checkErrorInfo(IUnknown *obj, HRESULT status, SEXP *serr)
{
  HRESULT hr;
  ISupportErrorInfo *info;

  fprintf(stderr, "<checkErrorInfo> %lX \n", status);

  if(serr) 
    *serr = NULL;

  hr = obj->QueryInterface(IID_ISupportErrorInfo, (void **)&info);
  if(hr != S_OK) {
    fprintf(stderr, "No support for ISupportErrorInfo\n");fflush(stderr);
    return(hr);
  }

  info->AddRef();
  hr = info->InterfaceSupportsErrorInfo(IID_IDispatch);
  info->Release();
  if(hr != S_OK) {
    fprintf(stderr, "No support for InterfaceSupportsErrorInfo\n");fflush(stderr);
    return(hr);
  }


  IErrorInfo *errorInfo;
  hr = GetErrorInfo(0L, &errorInfo);
  if(hr != S_OK) {
    /*    fprintf(stderr, "GetErrorInfo failed\n");fflush(stderr); */
    COMError(status);
    return(hr);
  }


  /* So there is some information for us. Use it. */
  SEXP klass, ans, tmp;
  BSTR ostr;
  char *str;

  errorInfo->AddRef();

  if(serr) {
   PROTECT(klass = MAKE_CLASS("SCOMErrorInfo"));
   PROTECT(ans = NEW(klass));

   PROTECT(tmp = NEW_CHARACTER(1));
   errorInfo->GetSource(&ostr);
   SET_STRING_ELT(tmp, 0, COPY_TO_USER_STRING(FromBstr(ostr)));
   SET_SLOT(ans, Rf_install("source"), tmp);
   UNPROTECT(1);

   PROTECT(tmp = NEW_CHARACTER(1));
   errorInfo->GetDescription(&ostr);
   SET_STRING_ELT(tmp, 0, COPY_TO_USER_STRING(str = FromBstr(ostr)));
   SET_SLOT(ans, Rf_install("description"), tmp);
   UNPROTECT(1);

   PROTECT(tmp = NEW_NUMERIC(1));
   NUMERIC_DATA(tmp)[0] = status;
   SET_SLOT(ans, Rf_install("status"), tmp);

   *serr = ans;
   UNPROTECT(3);

   errorInfo->Release();

   PROBLEM "%s", str
   WARN;
  } else {
   errorInfo->GetDescription(&ostr);
   str = FromBstr(ostr);
   errorInfo->GetSource(&ostr);
   errorInfo->Release();
   PROBLEM "%s (%s)", str, FromBstr(ostr)
   ERROR;
  }

  return(hr);
}



SEXP
R_createCOMErrorCodes()
{
  SEXP ans, names;
  int n;
        n = _countof(hrNameTable);
        PROTECT(ans = allocVector(REALSXP, n));
        PROTECT(names = allocVector(STRSXP, n));
	for (int i = 0; i < n; i++)
	{
	  REAL(ans)[i] = (double) hrNameTable[i].hr;
	  SET_STRING_ELT(names, i, COPY_TO_USER_STRING(hrNameTable[i].lpszName));
	}
	
	SET_NAMES(ans, names);
	UNPROTECT(2);
	return(ans);
}
