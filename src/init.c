#include <R.h>
#include <Rinternals.h>
#include <stdlib.h> // for NULL
#include <R_ext/Rdynload.h>

/* FIXME: 
Check these declarations against the C/Fortran source code.
*/

/* .Call calls */
extern SEXP R_connect(SEXP, SEXP);
extern SEXP R_connect_hWnd(SEXP, SEXP, SEXP);
extern SEXP R_create(SEXP);
extern SEXP R_create2DArray(SEXP);
extern SEXP R_getCLSIDFromName(SEXP);
extern SEXP R_getProperty(SEXP, SEXP, SEXP, SEXP);
extern SEXP R_initCOM(SEXP);
extern SEXP R_Invoke(SEXP, SEXP, SEXP, SEXP, SEXP, SEXP);
extern SEXP R_isValidCOMObject(SEXP);
extern SEXP R_setProperty(SEXP, SEXP, SEXP, SEXP);

static const R_CallMethodDef CallEntries[] = {
    {"R_connect",          (DL_FUNC) &R_connect,          2},
    {"R_connect_hWnd",     (DL_FUNC) &R_connect_hWnd,     3},
    {"R_create",           (DL_FUNC) &R_create,           1},
    {"R_create2DArray",    (DL_FUNC) &R_create2DArray,    1},
    {"R_getCLSIDFromName", (DL_FUNC) &R_getCLSIDFromName, 1},
    {"R_getProperty",      (DL_FUNC) &R_getProperty,      4},
    {"R_initCOM",          (DL_FUNC) &R_initCOM,          1},
    {"R_Invoke",           (DL_FUNC) &R_Invoke,           6},
    {"R_isValidCOMObject", (DL_FUNC) &R_isValidCOMObject, 1},
    {"R_setProperty",      (DL_FUNC) &R_setProperty,      4},
    {NULL, NULL, 0}
};

void R_init_excel_link(DllInfo *dll)
{
    R_registerRoutines(dll, NULL, CallEntries, NULL, NULL);
    R_useDynamicSymbols(dll, FALSE);
}
