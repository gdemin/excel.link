#' Basic operations with Excel workbooks
#' 
#' @param filename character. Excel workbook filename.
#' @param password character. Password for password-protected workbook.
#' @param write.res.password character. Second password for editing workbook.
#' @param file.format integer. Excel file format. By default it is
#'   \code{xl.constants$xlOpenXMLWorkbook}. You can use
#'   \code{xl.constants$xlOpenXMLWorkbookMacroEnabled} for workbooks with macros
#'   (*.xlsm) or \code{xl.constants$xlExcel12} for binary workbook (.xlsb).
#' @param xl.workbook.name character. Excel workbook name.
#' @param full.names logical. Should we return full path to the workbook? FALSE,
#'   by default.
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
    ex = xl.get.excel_no_add_workbook()
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
xl.workbook.open = function(filename, password = NULL, write.res.password = NULL)
    ## open workbook
{
    ex = xl.get.excel_no_add_workbook()
    wb.count = ex[['Workbooks']][['Count']]
    if(wb.count>0){
        wb.names = sapply(seq_len(wb.count), function(wb) ex[['Workbooks']][[wb]][['Name']])
        wb.names = tolower(wb.names)
        new_name = tolower(basename(filename))
        if(new_name %in% wb.names){
            return(invisible(xl.workbook.activate(new_name)))
        }
    }
    if (isTRUE(grepl("^(http|ftp)s?://", filename))){
        path = filename
    } else {
        path = normalizePath(filename, mustWork = TRUE)  
    }
    passwords =paste(!is.null(password), !is.null(write.res.password), sep = "_") 
    xl.wb = switch(passwords, 
                   FALSE_FALSE = ex[["Workbooks"]]$Open(path),
                   TRUE_FALSE = ex[["Workbooks"]]$Open(path, 
                                                            password = password
                   ),
                   FALSE_TRUE = ex[["Workbooks"]]$Open(path, 
                                                            writerespassword = write.res.password
                   ),
                   TRUE_TRUE = ex[["Workbooks"]]$Open(path, 
                                                           password = password, 
                                                           writerespassword = write.res.password
                   )
    )
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
xl.workbooks = function(full.names = FALSE)
    ## names of all opened workbooks
{
    ex = xl.get.excel()
    wb.count = ex[['Workbooks']][['Count']]
    if(full.names){
        sapply(seq_len(wb.count), function(wb) ex[['Workbooks']][[wb]][['FullName']])
    } else {
        sapply(seq_len(wb.count), function(wb) ex[['Workbooks']][[wb]][['Name']])
    }
}

#' @export
#' @rdname xl.workbook.add
xl.workbook.save = function(filename, password = NULL, write.res.password = NULL, file.format = xl.constants$xlOpenXMLWorkbook)
    ### save active workbook under the different name. If path is missing it saves in working directory
    ### doesn't alert if it owerwrite other file
{
    ex = xl.get.excel()
    path = normalizePath(filename,mustWork = FALSE)
    on.exit(ex[["DisplayAlerts"]] <- TRUE)
    ex[["DisplayAlerts"]] = FALSE
    passwords =paste(!is.null(password), !is.null(write.res.password), sep = "_") 
    switch(passwords, 
                   FALSE_FALSE = ex[["ActiveWorkbook"]]$SaveAs(path, FileFormat = file.format),
                   TRUE_FALSE = ex[["ActiveWorkbook"]]$SaveAs(path, 
                                                       password = password,
                                                       FileFormat = file.format
                   ),
                   FALSE_TRUE = ex[["ActiveWorkbook"]]$SaveAs(path, 
                                                       writerespassword = write.res.password,
                                                       FileFormat = file.format
                   ),
                   TRUE_TRUE = ex[["ActiveWorkbook"]]$SaveAs(path, 
                                                      password = password, 
                                                      writerespassword = write.res.password,
                                                      FileFormat = file.format
                   )
    )
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









