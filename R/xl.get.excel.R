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
  xls = COMCreate("Excel.Application")
  if (xls[['workbooks']][['count']] == 0) xls[['workbooks']]$add()
  if (!xls[["Visible"]]) xls[["Visible"]] = TRUE
  return(xls)
}