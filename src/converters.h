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

HRESULT R_convertRObjectToDCOM(SEXP obj, VARIANT *var);
SEXP R_convertDCOMObjectToR(VARIANT *var);
char *FromBstr(BSTR str);
BSTR AsBstr(const char *str);
SEXP getArray(SAFEARRAY *arr, int dimNo, int numDims, long *indices);

extern "C" {
  void RDCOM_finalizer(SEXP s);
  SEXP R_create2DArray(SEXP obj);
  SEXP R_createVariant(SEXP type);
  SEXP R_setVariant(SEXP svar, SEXP value, SEXP type);
}

