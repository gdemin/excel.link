#' Live connection with data on Microsoft Excel sheet
#' 
#' @description \code{xl.connect.table} returns object that can be operated as 
#'   usual data.frame object and this operations (e. g. subsetting, assignment) 
#'   will be immediately reflected on connected Excel range. See examples.
#'   Connected range is 'current region', e. g. selection which can be obtained
#'   by pressing \code{Ctrl+Shift+*} when selected \code{str.rng} (or top-left
#'   cell of this range is active).
#'   
#' @param str.rng string which represents Excel range
#' @param row.names a logical value indicating whether the Excel range contains 
#'   row names as its first column
#' @param col.names a logical value indicating whether the Excel range contains 
#'   column names as its first row
#' @param na character. NA representation in Excel. By default it is empty string
#' @param x object of class \code{excel.range}
#' @param decreasing logical. Should the sort be increasing or decreasing?
#' @param column numeric or character. Column by which we will sort. There is 
#'   special value - 'rownames'. In this case 'x' will be sorted by row names if
#'   it has it.
#' @param ...	arguments to be passed to or from methods or (for the default 
#'   methods and objects without a class)
#'   
#' @details Subsetting. Indices in subsetting operations are 
#'   numeric/character/logical vectors or empty (missing). Numeric values are 
#'   coerced to integer as by 'as.integer' (and hence truncated towards zero). 
#'   Character vectors will be matched to the 'colnames' of the object (or Excel
#'   column names if \code{has.colnames = FALSE}). For extraction form if column
#'   name doesn't exist error will be generated. For replacement form new column
#'   will be created. If indices are logical vectors they indicate 
#'   elements/slices to select. Such vectors are recycled if necessary to match 
#'   the corresponding extent. Indices can also be negative integers, indicating
#'   elements/slices to leave out of the selection.
#'   
#' @return \itemize{ \item{\code{xl.connect.table}}{ returns object of 
#'   \code{excel.range} class which represent data on Excel sheet. This object
#'   can be treated similar to data.frame. So you can assign values, delete 
#'   columns/rows and so on. For more information see examples.} 
#'   \item{\code{sort}}{ sorts Excel range by single column (multiple columns 
#'   currently not supported) and invisibly return NULL. }}
#' @examples
#' 
#' 
#' \dontrun{
#' ### session example 
#' 
#' library(excel.link)
#' xl.workbook.add()
#' xl.sheet.add("Iris dataset", before = 1)
#' xlrc[a1] = iris
#' xl.iris = xl.connect.table("a1", row.names = TRUE, col.names = TRUE)
#' dists = dist(xl.iris[, 1:4])
#' clusters = hclust(dists, method = "ward.D")
#' xl.iris$clusters = cutree(clusters, 3)
#' plot(clusters)
#' pl.clus = current.graphics()
#' cross = table(xl.iris$Species, xl.iris$clusters)
#' plot(cross)
#' pl.cross = current.graphics()
#' xl.sheet.add("Results", before = 2)
#' xlrc$a1 = list("Crosstabulation", cross,pl.cross, "Dendrogram", pl.clus)
#' 
#' ### completely senseless actions       
#' ### to demonstrate various operations and  
#' ### compare them with operations on usual data.frame
#' 
#' # preliminary operations 
#' data(iris)
#' rownames(iris) = as.character(rownames(iris))
#' iris$Species = as.character(iris$Species)
#' xl.workbook.add()
#' 
#' # drop dataset to Excel and connect it
#' xlrc[a1] = iris
#' xl.iris = xl.connect.table("a1", row.names = TRUE, col.names = TRUE)
#' identical(xl.iris[], iris)
#' 
#' # dim/colnames/rownames
#' identical(dim(xl.iris),dim(iris))
#' identical(colnames(xl.iris),colnames(iris))
#' identical(rownames(xl.iris),rownames(iris))
#' 
#' # sort datasets
#' iris = iris[order(iris$Sepal.Length), ]
#' sort(xl.iris, column = "Sepal.Length")
#' identical(xl.iris[], iris)
#' 
#' # sort datasets by rownames
#' sort(xl.iris, column = "rownames")
#' iris = iris[order(rownames(iris)), ]
#' identical(xl.iris[], iris)
#' 
#' # different kinds of subsetting
#' identical(xl.iris[,1:3], iris[,1:3])
#' identical(xl.iris[,3], iris[,3])
#' identical(xl.iris[26,1:3], iris[26,1:3])
#' identical(xl.iris[-26,1:3], iris[-26,1:3])
#' identical(xl.iris[50,], iris[50,])
#' identical(xl.iris$Species, iris$Species)
#' identical(xl.iris[,'Species', drop = FALSE], iris[,'Species', drop = FALSE])
#' identical(xl.iris[c(TRUE,FALSE), 'Sepal.Length'], 
#'              iris[c(TRUE,FALSE), 'Sepal.Length'])
#' 
#' # column creation and assignment 
#' xl.iris[,'group'] = xl.iris$Sepal.Length > mean(xl.iris$Sepal.Length)
#' iris[,'group'] = iris$Sepal.Length > mean(iris$Sepal.Length)
#' identical(xl.iris[], iris)
#' 
#' # value recycling
#' xl.iris$temp = c('aa','bb')
#' iris$temp = c('aa','bb')
#' identical(xl.iris[], iris)
#' 
#' # delete column
#' xl.iris[,"temp"] = NULL
#' iris[,"temp"] = NULL
#' identical(xl.iris[], iris)
#' 
#' }
#' @export
xl.connect.table = function(str.rng = "A1",row.names = TRUE,col.names = TRUE,na = "")
    ### return object, wich could be treated similar to data.frame (e. g. subsetting), but
    ### use an Excel data. 
{
    ex = xl.get.excel()
    f = local({
        xl.cell = ex[['Activesheet']]$Range(str.rng)
        hasrownames = row.names
        hascolnames = col.names
        function() { 
            res = xl.cell[['CurrentRegion']]
            has.rownames(res) = hasrownames
            has.colnames(res) = hascolnames
            attr(res,"NA") = na
            res
        }    
    })
    class(f) = c("excel.range",class(f))
    f
}


#' @export
#' @rdname xl.connect.table
sort.excel.range = function(x,decreasing = FALSE,column,...)
    # sort excel.range by given column
    # column may be character (column name), integer (column number), or logical.
    # By now it supports sorting only by single column
{
    if (length(column) !=  1 || is.na(column)) stop ("sorting column is not single or is NA. Please, choose one column for sorting")
    cols = colnames(x)
    if (length(column) == 1 && column == "rownames" && has.rownames(x)) {
        column = 1
    } else {
        if (!is.character(column)) column = cols[column]
        column = which(cols == column)
        if (length(column) == 0) stop ("coudn't find such column in the Excel frame.")
        if (length(column)>1) column = column[1]
        column = column+has.rownames(x)
    }
    xl.range = environment(x)$xl.cell[['currentregion']]
    # xl.cell = xl.range$cells(2,1)
    # sheet.sort = xl.range[["Worksheet"]][["Sort"]]
    # sheet.sort[["SortFields"]]$Clear()
    xl.range$sort(
        Key1 = xl.range[['Columns']][[column]],
        Order1 = decreasing+1, #xlAscending
        Header = 2 - has.colnames(x), #xlYes, xlNo
        OrderCustom = 1,
        MatchCase = TRUE,
        Orientation = 1,	#xlTopToBottom
        DataOption1 = 0 #xlSortNormal
    )
    invisible(NULL)
}





#' @export
'[.excel.range' = function(x, i, j, drop = if (missing(i)) TRUE else !missing(j) && (length(j) == 1))
    ## exctract variables from connected excel range. Similar to data.frame
{
    xl.rng = x()
    na = attr(xl.rng,"NA")
    dim.names = xl.dimnames(xl.rng)
    all.colnames = dim.names[[2]] 
    all.rownames = dim.names[[1]] 
    ncolx = length(all.colnames)
    nrowx = length(all.rownames)
    if (!missing(j)){
        if (is.character(j)) {
            if (!all(j %in% all.colnames)) stop("undefined columns selected")
            colnumber = match(j,all.colnames)
        } else {
            colnumber = 1:ncolx
            if (is.numeric(j)) {
                if (max(abs(j))>max(colnumber)) stop("Too large column index: ",max(abs(j))," vs ",max(colnumber)," columns in Excel table.")
                colnumber = colnumber[j]
            } else {
                if (is.logical(j)){
                    if (length(j)>max(colnumber) | max(colnumber)%%length(j) !=  0) stop('Subset has ',length(j),' columns, data has ',max(colnumber))
                    colnumber = colnumber[rep(j,length.out = max(colnumber))]
                } else stop("Undefined type of column indexing")
            }
        }
    } else {
        colnumber = 1:ncolx
    }	
    if (!missing(i)){
        if (is.character(i)) {
            if (!all(i %in% all.rownames)) stop("undefined rows selected")
        } else {
            rownumber = 1:nrowx
            if (is.numeric(i)) {
                if (max(abs(i))>max(rownumber)) stop("Too large row index: ",max(abs(i))," vs ",max(rownumber)," rows in Excel table.")
            } else {
                if (is.logical(i)){
                    if (length(i)>max(rownumber) | max(rownumber)%%length(i) !=  0) stop('Subset has ',length(i),' rows, data has ',max(rownumber))
                } else stop("Undefined type of row indexing")
            }
        }
    }
    colnumber = colnumber+has.rownames(xl.rng)	
    # if (has.colnames(x)) rownumber = rownumber+1
    raw.data = lapply(colnumber,function(each.col) xl.read.range(xl.rng[['columns']][[each.col]],col.names = TRUE, na = na))
    # raw.data = lapply(raw.data,function(each.col) unlist(each.col[[1]][-1]))
    res = do.call(data.frame,list(raw.data,stringsAsFactors = FALSE))
    colnames(res) = all.colnames[colnumber-has.rownames(xl.rng)]
    # print(all.rownames)
    if (!anyDuplicated(all.rownames)) rownames(res) = all.rownames else warning("There are duplicated rownames. They will be ignored.")
    if(!missing(i)) res = res[i,,drop = FALSE]
    if (drop & (ncol(res)<2)) return(res[,1]) else return(res)
}


#' @export
'$.excel.range' = function(x,value){
    x[,value,drop = TRUE]
}


#' @export
'[<-.excel.range' = function(x,i,j,value)
    ### assignment to columns in connected Excel range. If column doesn't exists it will create the new one. 
{
    #### if value = NULL we delete rows and columns
    delete.items = FALSE
    if (is.null(value)){
        if (!missing(i) & !missing(j)) stop("replacement has zero length.")
        value = NA
        delete.items = TRUE
    }
    if (!is.data.frame(value)) {
        value = as.data.frame(value,stringsAsFactors = FALSE)
    }	
    xl.rng = x()
    app = xl.rng[["Application"]]
    on.exit(make.me.slow(app))
    make.me.quick(app)
    na = attr(xl.rng,"NA")
    dim.names = xl.dimnames(xl.rng)
    all.colnames = dim.names[[2]] 
    all.rownames = dim.names[[1]] 
    ncolx = length(all.colnames)
    nrowx = length(all.rownames)
    ### dealing with columns
    value.colnum = ncol(value)
    new.columns = character(0)
    new.value = NULL
    if (missing(j)) all.cols = length(all.colnames) else all.cols = length(j)
    if (value.colnum>all.cols | all.cols%%value.colnum !=  0 ) stop('provided ',value.colnum,' variables to replace ',all.cols, ' variables.')
    if (all.cols>length(all.colnames)) stop('replacment has ',all.cols,' columns, data has ',length(all.colnames))
    if (all.cols !=  value.colnum) {
        value = value[,rep(1:value.colnum,length.out = all.cols),drop = FALSE]
        value.colnum = ncol(value)
    }
    if (!missing(j)){
        if (is.character(j)) {
            new.columns = j[!(j %in% all.colnames)] 
            if (length(new.columns)>0){
                if(!has.colnames(xl.rng)) stop ('new columns allowed only if range has colnames.')
                new.value = value[,!(j %in% all.colnames),drop = FALSE]
                value = value[,(j %in% all.colnames),drop = FALSE]
                value.colnum = ncol(value)
            }	
            j = j[j %in% all.colnames] 
            colnumber = match(j,all.colnames)			
        } else {
            colnumber = 1:ncolx
            if (is.numeric(j)) {
                if (max(abs(j))>max(colnumber)) stop("too large column index: ",max(abs(j))," vs ",max(colnumber)," columns in Excel table.")
                colnumber = colnumber[j]
            } else {
                if (is.logical(j)){
                    colnumber = colnumber[j]
                } else stop("undefined type of column indexing")
            }
        }
        colnumber = colnumber+has.rownames(xl.rng)	
    } 
    ### dealing with rows
    value.rownum = nrow(value)	
    if (missing(i)) all.rows = length(all.rownames) else if (is.logical(i)) all.rows = sum(i,na.rm = TRUE) else all.rows = length(i)
    if (value.rownum>all.rows | all.rows%%value.rownum !=  0) stop('replacment has ',value.rownum,' rows, data has ',all.rows)
    if (all.rows>length(all.rownames)) stop('replacment has ',all.rows,' rows, data has ',length(all.rownames))
    if (all.rows !=  value.rownum) {
        value = value[rep(1:value.rownum,length.out = all.rows),,drop = FALSE]
        if (length(new.columns)>0) new.value = new.value[rep(1:value.rownum,length.out = all.rows),,drop = FALSE]
        value.rownum = ncol(value)
    }
    if (!missing(i)){	
        if (is.character(i)) {
            if (!all(i %in% all.rownames)) stop("undefined rows selected")
            rownumber = match(i,all.rownames)
        } else {
            rownumber = 1:nrowx
            if (is.numeric(i)) {
                if (max(abs(i))>max(rownumber)) stop("too large row index: ",max(abs(i))," vs ",max(rownumber)," rows in Excel table.")
                rownumber = rownumber[i]
            } else {
                if (is.logical(i)){
                    rownumber = rownumber[rep(i,length.out = max(rownumber))]
                } else stop("undefined type of row indexing")
            }
        }
        rownumber = rownumber+has.colnames(xl.rng)
    } 
    if (delete.items){
        if (!missing(j)){
            colnumber = sort(colnumber,decreasing = TRUE)
            lapply(colnumber,function(k) {
                curr.rng = xl.rng[['Application']]$Range(xl.rng$cells(1,k),xl.rng$cells(length(all.rownames)+has.colnames(x),k))
                curr.rng$delete(Shift = -4159)
            })
            return(invisible(x))
        }
        if (!missing(i)){
            rownumber = sort(rownumber,decreasing = TRUE)
            lapply(rownumber,function(k) {
                curr.rng = xl.rng[['Application']]$Range(xl.rng$cells(k,1),xl.rng$cells(k,length(all.colnames)+has.rownames(x)))
                curr.rng$delete(Shift = -4162)
            })
            return(invisible(x))
        }
    }
    #### write data #####
    if (missing(i) & !missing(j)) {
        mapply (function(k,val) {
            curr.rng = xl.rng$cells(has.colnames(xl.rng)+1,k)
            xl.write.default(val,curr.rng,na = na,col.names = FALSE,row.names = FALSE)
        },colnumber,value
        )
        if (length(new.columns)>0 & !delete.items) {
            mapply(function(k,val) {
                kk = k+length(all.colnames)+has.rownames(xl.rng)
                insert.range = xl.rng[['columns']][[kk]]
                insert.range$insert(Shift = -4161)
                curr.rng = xl.rng$cells(has.colnames(xl.rng)+1,kk)
                dummy = xl.rng$cells(1,kk)
                dummy[['Value']] = new.columns[k]
                xl.write.default(val,curr.rng,na = na,col.names = FALSE,row.names = FALSE)
            }, seq_along(new.columns),new.value
            )
        }	
    }
    if (!missing(i) & missing(j)) {
        mapply (function(k,val) {
            curr.rng = xl.rng$cells(k,1+has.rownames(xl.rng))
            xl.writerow(val,curr.rng,na = na)
        },rownumber,as.data.frame(t(value),stringsAsFactors = FALSE)
        )
    }
    if (!missing(i) & !missing(j)) {
        mapply (function(k1,val1) {
            mapply(function(k2,val2){
                xl.write.default(val2,xl.rng$Cells(k2,k1),na = na,col.names = FALSE,row.names = FALSE)
            },
            rownumber,val1)
        },colnumber,value)
        if (length(new.columns)>0) {
            nv = as.data.frame(matrix(NA,nrow = length(all.rownames),ncol = length(new.columns)))
            nv[rownumber-has.colnames(xl.rng),] = new.value
            colnames(nv) = new.columns
            # browser()
            for (k in seq_along(new.columns)) {
                kk = k+length(all.colnames)+has.rownames(xl.rng)
                insert.range = xl.rng[['columns']][[kk]]
                insert.range$insert(Shift = -4161)
            }
            curr.rng = xl.rng$cells(1,1+length(all.colnames)+has.rownames(xl.rng))
            xl.write.data.frame(nv,curr.rng,na = na,col.names = has.colnames(xl.rng),row.names = FALSE)
        }
    }	
    invisible(x)
}


#' @export
'$<-.excel.range' = function(x,j,value){
    x[,j] = value
    invisible(x)
}
