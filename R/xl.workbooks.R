#' Basic operations with Excel workbooks
#' 
#' @param filename character. Excel workbook filename.
#' @param password character. Password for password-protected workbook.
#' @param xl.workbook.name character. Excel workbook name.
#'   
#' @return \itemize{ 
#' \item{\code{xl.workbook.add}/\code{xl.workbook.open}/\code{xl.workbook.activate}
#' }{ invisibly return name of created/open/activated workbook.} 
#' \item{\code{xl.workbooks}}{ returns character vector of open workbooks.} 
#' \item{\code{xl.workbook.save}}{ invisibly returns path to the saved workbook}
#' \item{\code{xl.workbook.close}}{ invisibly returns NULL.} }
#' @details \itemize{ \item{\code{xl.workbook.add}}{ adds new workbook and
#' invisibly returns name of this newly created workbook. Added workbook become
#' active. If \code{filename} argument is provided then Excel workbook
#' \code{filename} will be used as template.} 
#' \item{\code{xl.workbook.activate}}{ activates workbook with given name. If
#' workbook with this name doesn't exist error will be generated.} 
#' \item{\code{xl.workbook.save}}{ saves active workbook. If only
#' \code{filename} submitted it saves in the working directory. If name of
#' workbook is omitted than new workbook is saved under its default name in the
#' current working directory. It doesn't prompt about overwriting if file
#' already exists.} \item{\code{xl.workbook.close}}{ closes workbook with given
#' name. If name isn't submitted it closed active workbook.  It doesn't prompt
#' about saving so if you don't save changes before closing all changes will be
#' lost.} }
#' 
#' @seealso \code{\link{xl.sheets}}, \code{\link{xl.read.file}}, 
#'   \code{\link{xl.save.file}}
#' @examples
#' \dontrun{
#' ## senseless actions
#' data(iris)
#' data(cars)
#' xl.workbook.add()
#' xlrc[a1] = iris
#' xl.workbook.save("iris.xlsx")
#' xl.workbook.add()
#' xlrc[a1] = cars
#' xl.workbook.save("cars.xlsx")
#' xl.workbook.activate("iris")
#' xl.workbook.close("cars")
#' xl.workbook.open("cars.xlsx")
#' xl.workbooks()
#' for (wb in xl.workbooks()) xl.workbook.close(wb)
#' unlink("iris.xlsx")
#' unlink("cars.xlsx")
#' 
#' # password-protected workbook
#' data(iris) 
#' xl.workbook.add()
#' xlrc[a1] = iris
#' xl.workbook.save("iris.xlsx", password = "my_password")
#' xl.workbook.close()
#' xl.workbook.open("iris.xlsx", password = "my_password")
#' xl.workbook.close()
#' unlink("iris.xlsx")
#' }
#' @export
xl.workbook.add = function(filename = NULL)
    ### add new workbook and invisibily return it's name
    ### if filename is give, its used as template 
{
    ex = xl.get.excel()
    if (!is.null(filename)) {
        if (isTRUE(grepl("^(http|ftp)s?://", filename))){
            path = filename
        } else {
            path = normalizePath(filename,mustWork = TRUE)  
        }
        xl.wb = ex[['Workbooks']]$Add(path) 
    } else xl.wb = ex[['Workbooks']]$Add()
    invisible(xl.wb[["Name"]])
}


#' @export
#' @rdname xl.workbook.add
xl.workbook.open = function(filename,password = NULL)
    ## open workbook
{
    ex = xl.get.excel()
    if (isTRUE(grepl("^(http|ftp)s?://", filename))){
        path = filename
    } else {
        path = normalizePath(filename,mustWork = TRUE)  
    }
    if(is.null(password)){
        xl.wb = ex[["Workbooks"]]$Open(path)
    } else {
        xl.wb = ex[["Workbooks"]]$Open(path, password = password)
    }
    invisible(xl.wb[['Name']])
}


#' @export
#' @rdname xl.workbook.add
xl.workbook.activate = function(xl.workbook.name)
    ### activate sheet with given name in active workbook 
{
    ex = xl.get.excel()
    on.exit(ex[["DisplayAlerts"]] <- TRUE)
    workbooks.xls = tolower(xl.workbooks())
    workbooks = gsub("\\.([^\\.]+)$","",tolower(xl.workbooks()),perl = TRUE)
    wb.num = which((tolower(xl.workbook.name) == workbooks.xls) | (tolower(xl.workbook.name) == workbooks))
    if (length(wb.num) == 0) stop ('workbook with name "',xl.workbook.name,'" doesn\'t exists.')
    xl.wb = ex[['workbooks']][[wb.num]]
    ex[["DisplayAlerts"]] = FALSE
    xl.wb$activate()
    invisible(xl.wb[['Name']])
}



#' @export
#' @rdname xl.workbook.add
xl.workbooks = function()
    ## names of all opened workbooks
{
    ex = xl.get.excel()
    wb.count = ex[['Workbooks']][['Count']]
    sapply(seq_len(wb.count), function(wb) ex[['Workbooks']][[wb]][['Name']])
}

#' @export
#' @rdname xl.workbook.add
xl.workbook.save = function(filename,password = NULL)
    ### save active workbook under the different name. If path is missing it saves in working directory
    ### doesn't alert if it owerwrite other file
{
    ex = xl.get.excel()
    path = normalizePath(filename,mustWork = FALSE)
    on.exit(ex[["DisplayAlerts"]] <- TRUE)
    ex[["DisplayAlerts"]] = FALSE
    if(is.null(password)) {
        ex[["ActiveWorkbook"]]$SaveAs(path)
    } else {
        ex[["ActiveWorkbook"]]$SaveAs(path,password = password)
    }
    invisible(path)
}



#' @export
#' @rdname xl.workbook.add
xl.workbook.close = function(xl.workbook.name = NULL)
    ### close workbook with given name or active workbook if xl.workbook.name is missing
    ## it doesn't promp to save changes, so changes will be lost if workbook isn't saved
{
    ex = xl.get.excel()
    on.exit(ex[["DisplayAlerts"]] <- TRUE)
    if (!is.null(xl.workbook.name)){
        workbooks.xls = tolower(xl.workbooks())
        workbooks = gsub("\\.([^\\.]+)$","",tolower(xl.workbooks()),perl = TRUE)
        wb.num = which((tolower(xl.workbook.name) == workbooks.xls) | (tolower(xl.workbook.name) == workbooks))
        if (length(wb.num) == 0) stop ('workbook with name "',xl.workbook.name,'" doesn\'t exists.')
        xl.wb = ex[['workbooks']][[wb.num]]
    } else xl.wb = ex[["ActiveWorkbook"]]
    ex[["DisplayAlerts"]] = FALSE
    xl.wb$close(SaveChanges = FALSE)
    invisible(NULL)
}









