#' Returns reference to Excel application.
#' 
#' Returns reference to Microsoft Excel application. If there is no running 
#' instance exists it will create a new instance.
#' 
#' @return object of class 'COMIDispatch' (as returned by 
#' \code{\link[RDCOMClient]{COMCreate}} from RDCOMClient package).
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
    if (xls[['workbooks']][['count']] == 0) xls[['workbooks']]$add()
    
    return(xls)
}