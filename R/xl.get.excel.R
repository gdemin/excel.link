#' Returns reference to Excel application.
#' 
#' Returns reference to Microsoft Excel application. If there is no running 
#' instance exists it will create a new instance.
#' 
#' @return object of class 'COMIDispatch' (as returned by 
#' \code{COMCreate} from RDCOMClient package).
#' 
#' @examples
#' 
#' \dontrun{
#' xls = xl.get.excel() 
#' }
#' @export
xl.get.excel = function()
    # run Excel if it's not running and
    # return reference to Microsoft Excel
{
#     xls = COMCreate("Excel.Application")
    excel_CLSID = "{00020400-0000-0000-C000-000000000046}"
    excel_hwnd = unlist(options("excel_hwnd"))
    succ = FALSE
    if(!is.null(excel_hwnd)){
        xls = getCOMInstance_hWnd(excel_CLSID,excel_hwnd)
        if (!(is.null(xls) || ("COMErrorString" %in% class(xls)))){
            succ = TRUE
            xls = xls[["Application"]]
#             xls[["Visible"]] = TRUE
        }
        
    }
    if(!succ){
        xls = getCOMInstance("Excel.Application",force = FALSE,silent = TRUE)
        if (is.null(xls) || ("COMErrorString" %in% class(xls))) {
            xls = getCOMInstance("Excel.Application",force = TRUE,silent = TRUE)
            xls[["Visible"]] = TRUE
        } else {
            if (!xls[["Visible"]]){
                xls[["Visible"]] = TRUE
                warning("Connection with hidden Microsoft Excel instance. It may cause problems. Try to kill this instance from task manager.")
            } 
        }    
    }
    if (xls[['workbooks']][['count']] == 0) xls[['workbooks']]$add()
    
    return(xls)
}