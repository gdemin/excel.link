#' Functions for saving and reading data to/from Excel file.
#' 
#' @param filename a character string naming a file
#' @param r.obj R object
#' @param header a logical value indicating whether the file contains the names
#'   of the variables as its first line. If TRUE and top-left corner is empty
#'   cell, first column is considered as row names. Ignored if row.names or
#'   col.names is not NULL.
#' @param row.names a logical value indicating whether the row names of r.obj
#'   are to be read/saved along with r.obj
#' @param col.names a logical value indicating whether the column names of r.obj
#'   are to be read/saved along with r.obj
#' @param xl.sheet character. Name of Excel sheet where data is located/will be
#'   saved. By default it is NULL and data will be read/saved from/to active
#'   sheet.
#' @param top.left.cell character. Top-left corner of data in Excel sheet. By
#'   default is 'A1'.
#' @param na character. NA representation in Excel. By default it is empty
#'   string
#' @param excel.visible a logical value indicating will Excel visible during
#'   this operations. FALSE by default.
#'   
#' @details \code{xl.read.file} reads only rectangular data set. It is highly
#' recommended to have all column names and ids in data set. Orphaned
#' rows/columns located apart from the main data will be ignored.
#' \code{xl.save.file} can save all objects for which \code{xl.write} method exists -
#' see examples.
#' 
#' @return \code{xl.read.file} always returns data.frame. \code{xl.save.file}
#' invisibly returns NULL.
#' @seealso
#' 
#' \code{\link{xl.write}}, \code{\link{xl.workbook.save}},
#' \code{\link{xl.workbook.open}}, \code{\link{current.graphics}}
#' 
#' @examples
#' 
#' 
#' \dontrun{
#' data(iris)
#' xl.save.file(iris,"iris.xlsx")
#' xl.iris = xl.read.file("iris.xlsx")
#' all(iris == xl.iris) # Shoud be TRUE
#' unlink("iris.xlsx")
#' 
#' # Save to file list with different data types 
#' dists = dist(iris[,1:4])
#' clusters = hclust(dists,method="ward")
#' iris$clusters = cutree(clusters,3)
#' png("1.png")
#' plot(clusters)
#' dev.off()
#' pl.clus = current.graphics(filename="1.png")
#' cross = table(iris$Species,iris$clusters)
#' png("2.png")
#' plot(cross)
#' dev.off()
#' pl.cross = current.graphics(filename="2.png")
#' output = list("Iris",pl.clus,cross,pl.cross,"Data:","",iris)
#' xl.save.file(output,"output.xls")
#' xl.workbook.open("output.xls")
#' # xl.workbook.close() # close workbook
#' # unlink("output.xls") # delete file
#' 
#' }
#' @export
xl.read.file = function(filename, header = TRUE, row.names = NULL, col.names = NULL, 
                        xl.sheet = NULL,top.left.cell = "A1", na = "",
                        excel.visible = FALSE)
    # read data from excel file
    # filename - name of the file
    # header if TRUE First row treated as colnames and if top.left.cell is empty then first column treated as rownames.
    # if row.names or col.names not is null header argument will be ignored
    # if row.names is TRUE first column will be treated as rownames
    # if col.names is TRUE first row will be treated as colnames
    # xl.sheet - can be character - sheet name or numeric - number number. if omitted data will be read from active sheet 
    # na - string which will be treated as NA value
    # top.left.cell - top-left corner of region which will be read
    # excel.visible if TRUE Excel will be visible during operation
{
    xl_temp = COMCreate("Excel.Application",existing = FALSE)
    on.exit(xl_temp$quit()) 
    xl_temp[["Visible"]] = excel.visible
    xl_temp[["DisplayAlerts"]] = FALSE
    xl_wb = xl_temp[["Workbooks"]]$Open(normalizePath(filename,mustWork = TRUE))
    # on.exit(xl_wb$close())
    # on.exit(xl_temp$quit(),add = TRUE)
    if (!is.null(xl.sheet)){
        if (!is.character(xl.sheet) & !is.numeric(xl.sheet)) stop('Argument "xl.sheet" should be character or numeric.')
        sh.count = xl_wb[['Sheets']][['Count']]
        sheets = sapply(seq_len(sh.count), function(sh) xl_wb[['Sheets']][[sh]][['Name']])
        if (is.numeric(xl.sheet)){
            if (xl.sheet>length(sheets)) stop ("too large sheet number. In workbook only ",length(sheets)," sheet(s)." )
            xl_wb[["Sheets"]][[xl.sheet]]$Activate()
        } else {
            sheet_num = which(tolower(xl.sheet) == tolower(sheets)) 
            if (length(sheet_num) == 0) stop ("sheet ",xl.sheet," doesn't exist." )
            xl_wb[["Sheets"]][[sheet_num]]$Activate()
        }
    }
    if(is.null(row.names) && is.null(col.names)){
        if(header){
            col.names = TRUE
            temp = xl.read.range(xl_temp[["ActiveSheet"]]$range(top.left.cell),na = "")
            row.names = is.na(temp) || all(grepl("^([\\s\\t]+)$",temp,perl = TRUE))
        } else {
            row.names = FALSE
            col.names = FALSE
        }
    } else {
        if (is.null(row.names)) row.names = FALSE
        if (is.null(col.names)) col.names = FALSE
    }
    top_left_corner = xl_temp$range(top.left.cell)
    xl.rng = top_left_corner[["CurrentRegion"]]
    if (tolower(top.left.cell) !=  "a1") {
        bottom_row = xl.rng[["row"]]+xl.rng[["rows"]][["count"]]-1
        right_column = xl.rng[["column"]]+xl.rng[["columns"]][["count"]]-1
        xl.rng = xl_temp$range(top_left_corner,xl_temp$cells(bottom_row,right_column))
    } 
    xl.read.range(xl.rng,drop = FALSE,na = na,row.names = row.names,col.names = col.names)
}


#' @export
#' @rdname xl.read.file
xl.save.file = function(r.obj,filename, row.names = TRUE, col.names = TRUE, 
                        xl.sheet = NULL, top.left.cell = "A1", na = "",
                        excel.visible = FALSE)
{
    xl_temp = COMCreate("Excel.Application",existing = FALSE)
    on.exit(xl_temp$quit()) 
    xl_temp[["Visible"]] = excel.visible
    xl_temp[["DisplayAlerts"]] = FALSE
    xl_wb = xl_temp[["Workbooks"]]$Add()
    if (!is.null(xl.sheet)){
        sh.count = xl_wb[['Sheets']][['Count']]
        sheets = sapply(seq_len(sh.count), function(sh) xl_wb[['Sheets']][[sh]][['Name']])
        if ((tolower(xl.sheet) %in% sheets)) stop ('sheet with name "',xl.sheet,'" already exists.')
        res = xl_temp[['ActiveWorkbook']][['Sheets']]$Add(Before = xl_temp[['ActiveWorkbook']][['Sheets']][[1]])
        res[['Name']] = substr(xl.sheet,1,63)
    }
    top_left_corner = xl_temp$range(top.left.cell)
    xl.write(r.obj,xl.rng = top_left_corner,row.names = row.names,col.names = col.names,na = na)
    path = normalizePath(filename,mustWork = FALSE)
    xl_temp[["ActiveWorkbook"]]$SaveAs(path)
    invisible(NULL)
}