FIX_ENCODING = "excel.link.fix_encoding"
EXCEL_MAX_ROWS = 1048576
EXCEL_MAX_COLS = 16384

#' Methods for writing data to Excel sheet
#' 
#' @param r.obj R object
#' @param xl.rng An object of class \code{COMIDispatch} (as used in RDCOMClient 
#'   package) - reference to Excel range
#' @param na character. NA representation in Excel. By default it is empty string
#' @param row.names a logical value indicating whether the row names/vector 
#'   names of r.obj should to be written along with r.obj
#' @param col.names a logical value indicating whether the column names of r.obj
#'   should to be written along with r.obj
#' @param delete.file a logical value indicating whether delete file with 
#'   graphics after insertion in Excel
#' @param ...	arguments for further processing
#'   
#' @details 
#' \code{xl.rng} should be COM-reference to Excel range, not string. Method 
#' invisibly returns number of columns and rows occupied by \code{r.obj} on
#' Excel sheet. It's useful for multiple objects writing to prevent their
#' overlapping. It is more convenient to use \code{xl} object. \code{xl.write}
#' aimed mostly for programming purposes, not for interactive usage.
#' 
#' @return c(rows,columns) Invisibly returns rows and columns number ocuppied by
#' \code{r.obj} on Excel sheet.
#' 
#' @seealso \code{\link{xl}},
#' \code{\link{xlr}}, \code{\link{xlc}}, \code{\link{xlrc}}, 
#' \code{\link{current.graphics}}
#' 
#' @examples
#' \dontrun{
#' xls = xl.get.excel()
#' xl.workbook.add()
#' rng = xls[["Activesheet"]]$Cells(1,1)
#' nxt = xl.write(iris,rng,row.names = TRUE,col.names = TRUE)
#' rng = rng$Offset(nxt[1] + 1,0)
#' nxt = xl.write(cars,rng,row.names = TRUE,col.names = TRUE)
#' rng = rng$Offset(nxt[1] + 1,0)
#' nxt = xl.write(as.data.frame(Titanic),rng,row.names = TRUE,col.names = TRUE)
#' 
#' data(iris)
#' data(cars)
#' data(Titanic)
#' xl.sheet.add()
#' rng = xls[["Activesheet"]]$Cells(1,1)
#' data.sets = list("Iris dataset",iris,
#'      "Cars dataset",cars,
#'      "Titanic dataset",as.data.frame(Titanic))
#' xl.write(data.sets,rng,row.names = TRUE,col.names = TRUE)
#' 
#' }
#' @export
xl.write = function(r.obj,xl.rng,na = "",...)
    ## insert values in excel range.
    ## should return c(row,column) - next empty point
{
    app = xl.rng[["Application"]]
    on.exit(make.me.slow(app))
    make.me.quick(app)
    UseMethod("xl.write")
}

#' @export
xl.write.xl.property = function(r.obj, xl.rng, ...){
    for(each_item in seq_along(r.obj)){
        each_name = names(r.obj)[each_item]
        properties = unlist(strsplit(each_name, split = ".", fixed = TRUE))
        curr_obj = xl.rng
        num_of_properties = length(properties)
        if(num_of_properties>1){
            for(each in properties[-num_of_properties]){
                curr_obj = curr_obj[[each]]    
            }
        }
        curr_obj[[properties[num_of_properties]]] = r.obj[[each_item]]
    } 
    invisible(c(xl.rng[["rows"]][["count"]], xl.rng[["columns"]][["count"]]))
}


#' @export
#' @rdname xl.write
xl.write.current.graphics = function(r.obj,xl.rng,na = "",delete.file = FALSE,...)
    ## insert picture at the top-left corner of given range
    ## r.obj - picture filename with "current.graphics" class attribute
    ## by default file will be deleted
{
    app = xl.rng[["Application"]]
    curr.sheet = app[["ActiveSheet"]]
    on.exit(curr.sheet$Activate())
    xl.sheet = xl.rng[["Worksheet"]]
    xl.sheet$Activate()
    top = xl.rng[["Top"]]
    left = xl.rng[["Left"]]
    pic = app[["Activesheet"]][['Pictures']]$Insert(unclass(r.obj))
    height = pic[["Height"]]
    width = pic[["Width"]]
    pic$Delete()
    picname = tolower(attr(r.obj,"picname"))
    if (length(picname)>0){
        shape_count = app[["Activesheet"]][['Shapes']][["Count"]]
        for (i in seq_len(shape_count)){
            curr_shape = app[["Activesheet"]][['Shapes']][[i]]
            if(tolower(curr_shape[["Name"]]) == picname){
                curr_shape$Delete()
                break
            }
        }       
    }
    pic = app[["Activesheet"]][['Shapes']]$AddPicture(unclass(r.obj),0,-1,left,top,width,height)
    if (length(picname)>0){
        pic[["Name"]] = picname   
    }    
    fill = pic[['Fill']] 
    fill[['ForeColor']][['RGB']] = 16777215L
    height = pic[["Height"]]+top
    width = pic[["Width"]]+left
    i = 0
    temp = xl.rng$Offset(i,0)
    while(height>temp[['top']]){
        i = i+1
        temp = xl.rng$Offset(i,0)
    }
    j = 0
    temp = xl.rng$Offset(0,j)
    while(width>temp[['left']]){
        j = j+1
        temp = xl.rng$Offset(0,j)
    }


    if (delete.file) file.remove(r.obj)
    invisible(c(i,j))
}



#' @export
#' @rdname xl.write
xl.write.list = function(r.obj,xl.rng,na = "",...)
    ## insert list into excel sheet. Each element pastes on next empty row 
{
    res = c(0,0)
    list.names = names(r.obj)
    for (each.item in seq_along(r.obj)){
        each.name = list.names[each.item]
        has.name = !is.null(each.name) && each.name  !=  "" && length(each.name)>0
        if (has.name) xl.write(each.name,xl.rng$offset(res[1],0),na,...)
        new.res = xl.write(r.obj[[each.item]],xl.rng$offset(res[1],1*has.name),na,...)
        res[1] = res[1]+new.res[1]
        res[2] = max(res[2],new.res[2])
    }
    invisible(res)
}




#' @export
#' @rdname xl.write
xl.write.matrix = function(r.obj,xl.rng,na = "",row.names = TRUE,col.names = TRUE,...)
    ## insert matrix into excel sheet including column and row names
{
    if (!is.null(r.obj)){
        check_size(r.obj = r.obj, xl.rng = xl.rng, row.names = row.names, col.names = col.names)
        if(isTRUE(getOption(FIX_ENCODING))) {
            r.obj = fix_encoding(r.obj)
            na  = fix_encoding(na)
        }
        xl.colnames = colnames(r.obj)
        xl.rownames = rownames(r.obj)
        has.col = (!is.null(xl.colnames) && col.names)*1
        has.row = (!is.null(xl.rownames) && row.names)*1
        dim.names = names(dimnames(r.obj))
        has.col.dimname =  (!is.null(dim.names[2]) && !(dim.names[2] %in% c("",NA)) && col.names)*1
        has.row.dimname =  (!is.null(dim.names[1]) && !(dim.names[1] %in% c("",NA)) && row.names)*1
        delta_row = has.col + has.col.dimname
        delta_col = has.row + has.row.dimname
        # clear output area
        out.rng = xl.rng[['Application']]$range(xl.rng$cells(1,1),xl.rng$cells(nrow(r.obj)+delta_row,ncol(r.obj)+delta_col))
        out.rng$clear()
        if (has.col.dimname){
            xl.raw.write(dim.names[2],xl.rng$offset(0,delta_col),na)
        }
        if (has.col) {

            xl.raw.write(t(xl.colnames),xl.rng$offset(has.col.dimname,delta_col),na)
        }    

        if (has.row.dimname){
            xl.raw.write(dim.names[1],xl.rng$offset(delta_row,0),na)
        }
        if (has.row) {
            xl.raw.write(xl.rownames,xl.rng$offset(delta_row,has.row.dimname),na)
        }	
        # for (i in seq_len(ncol(r.obj)))	xl.raw.write(r.obj[,i],xl.rng$offset(has.col,i+has.row-1),na)
        xl.raw.write.matrix(r.obj,xl.rng$offset(delta_row,delta_col),na)
        invisible(c(nrow(r.obj)+delta_row,ncol(r.obj)+delta_col))
    } else {
        invisible(c(0,0))
    }
    
}


#' @export
#' @rdname xl.write
xl.write.data.frame = function(r.obj,xl.rng,na = "",row.names = TRUE,col.names = TRUE,...)
    ## insert data.frame into excel sheet including column and row names
{
    # stop("Multi-column (e. g. matrix) data.frame elements currently not supported.")
    if (!is.null(r.obj)){
        if(isTRUE(getOption(FIX_ENCODING))) {
            r.obj = fix_encoding(r.obj)
            na  = fix_encoding(na)
        }
        check_size(r.obj = r.obj, xl.rng = xl.rng, row.names = row.names, col.names = col.names)
        xl.colnames = colnames(r.obj)
        column.numbers = sapply(r.obj,NCOL)
        if (any(column.numbers>1)) {
            xl.colnames = rep(xl.colnames,times = column.numbers)
            suffix = as.list(character(length(column.numbers)))
            suffix[column.numbers>1] = lapply(column.numbers[column.numbers>1],function(x) paste(".",seq(x),sep = ""))
            xl.colnames = paste(xl.colnames,unlist(suffix),sep = "")
        }	
        xl.rownames = rownames(r.obj)
        has.col = (!is.null(xl.colnames) & col.names)*1
        has.row = (!is.null(xl.rownames) & row.names)*1
        if (has.col) xl.raw.write(t(xl.colnames),xl.rng$offset(0,has.row),na)
        if (has.row) xl.raw.write(xl.rownames,xl.rng$offset(has.col,0),na)
        types = rle(sapply(r.obj,function(x) paste(class(x),collapse = "&")))
        lens = types$lengths
        beg = head(c(1,1+cumsum(lens)),-1)
        end = cumsum(lens)
        if (has.col || has.row) xl.rng = xl.rng$offset(has.col,has.row)
        for (i in seq_along(beg)){
            x = beg[i]
            y = end[i]
            col.offset = xl.raw.write.matrix(as.matrix(r.obj[,x:y,drop = FALSE]),xl.rng,na)[2]
            xl.rng = xl.rng$offset(0,col.offset)
        }
    }
    invisible(c(nrow(r.obj)+has.col,ncol(r.obj)+has.row))
}


#' @export
#' @rdname xl.write
xl.write.default = function(r.obj,xl.rng,na = "",row.names = TRUE,...){
    if (is.null(r.obj) || length(r.obj) == 0) r.obj = ""
    if(isTRUE(getOption(FIX_ENCODING))) {
        r.obj = fix_encoding(r.obj)
        na  = fix_encoding(na)
    }
    check_size(r.obj = r.obj, xl.rng = xl.rng, row.names = row.names)
    obj.names = names(r.obj)
    if (!is.null(obj.names) & row.names){
        res = xl.raw.write(obj.names,xl.rng,na)+xl.raw.write(r.obj,xl.rng$offset(0,1),na)
    } else {
        if (length(r.obj)<2) r.obj = matrix(r.obj,nrow = xl.rng[['rows']][['count']],ncol = xl.rng[['columns']][['count']])
        if (length(r.obj)<2) r.obj = drop(r.obj)	
        res = xl.raw.write(r.obj,xl.rng,na)
    }
    invisible(res)
}

#' @export
xl.write.factor = function(r.obj,xl.rng,na = "",row.names = TRUE,...){
    r.obj = as.character(r.obj)
    xl.write(r.obj,xl.rng = xl.rng,na = na,row.names = row.names,...)
}

#' @export
xl.write.table = function(r.obj,xl.rng,na = "",row.names = TRUE,col.names = TRUE,...){
    if(length(dim(r.obj))<3) {
        mat_r.obj = matrix(r.obj, ncol = NCOL(r.obj))
        dimnames(mat_r.obj) = dimnames(r.obj)
        invisible(xl.write.matrix(mat_r.obj,xl.rng,na,row.names = row.names,col.names = col.names, ...))
        
    } else {
        stop ("tables with dimensions greater than 2 currently doesn't supported")
        # if(length(dim(r.obj)) == 3) {
        # dim.names = names(dimnames(r.obj))
        # if (!is.null(dim.names[3])) {
        # xl.rng = xl.rng$offset(xl.write(dim.names[1],xl.rng)[1],0)
        # } 
        # curr.names = dimnames(r.obj)[[3]]
        # if (is.null(curr.names)) curr.names = seq_len(dim(r.obj)[3])
        # for (i in seq_len(dim(r.obj)[3])){
        # xl.write(curr.names[i],xl.rng)
        # xl.rng = xl.rng$offset(0,xl.write(r.obj[,,1],xl.rng,row.names = (i == 1))[2])
        # }
        # }	
    }	
}



# xl.write.ftable = function(r.obj,xl.rng,na = "",...){
# invisible(xl.write(format(r.obj,nsmall = 20,quote = FALSE),xl.rng,na))
# }

#' @export
xl.write.etable  = function(r.obj, xl.rng, na = "",  row.names = FALSE, col.names = TRUE, remove_repeated = TRUE, ...){
    class(r.obj) = setdiff(class(r.obj), "etable")
    header = t(split_labels(colnames(r.obj), remove_repeated = remove_repeated))[,-1, drop = FALSE]
    row_labels = split_labels(r.obj[[1]], remove_repeated = remove_repeated)
    # drop completely empty rows
    header = header[rowSums(!is.na(header) &(header!=""))>0, , drop = FALSE]
    top_left_corner = matrix(NA, ncol= NCOL(row_labels), nrow = NROW(header))
    if(!is.null(colnames(r.obj)) && !(colnames(r.obj)[1] %in% c(NA,"row_labels",""))){
        top_left_corner[nrow(top_left_corner), 1] = colnames(r.obj)[1]    
    }
    r.obj = r.obj[, -1, drop = FALSE]
    if(col.names){
        nxt = xl.write(top_left_corner, xl.rng, na = na, row.names = FALSE, col.names = FALSE)
        rng = xl.rng$Offset(0, nxt[2])
        xl.write(header, rng, na = na, row.names = FALSE, col.names = FALSE)
        rng = xl.rng$Offset(nxt[1], 0)
        xl.write(row_labels, rng, na = na, row.names = FALSE, col.names = FALSE)
        rng = xl.rng$Offset(nxt[1] ,nxt[2])
        nxt2 = xl.write(r.obj, rng, na = na, row.names = FALSE, col.names = FALSE)
        invisible(nxt + nxt2)
    } else {
        rng = xl.rng
        nxt = xl.write(row_labels, rng, na = na, row.names = FALSE, col.names = FALSE)
        rng = xl.rng$Offset(0 ,nxt[2])
        nxt2 = xl.write(r.obj, rng, na = na, row.names = FALSE, col.names = FALSE)
        invisible(c(nxt[1], nxt[2] + nxt2[2]))
    }
}

#####################

check_size = function(r.obj, xl.rng, row.names = FALSE, col.names  = FALSE){
    row_offset = xl.rng[["row"]] - 1
    col_offset = xl.rng[["column"]] - 1
    stopif((NROW(r.obj) + col.names + row_offset) > EXCEL_MAX_ROWS, 
           sprintf("'r.obj' (nrows = %s) doesn't fitted in Excel worksheet (max. rows = %s)",
                   NROW(r.obj), EXCEL_MAX_ROWS))
    stopif((NCOL(r.obj) + row.names + col_offset) > EXCEL_MAX_COLS, 
           sprintf("'r.obj' (ncols = %s) doesn't fitted in Excel worksheet (max. columns = %s)",
                   NCOL(r.obj), EXCEL_MAX_COLS))
}

##############################

xl.writerow = function(r.obj,xl.rng,na = "")
    ## special function for writing single row on excel sheet
{
    if (is.null(r.obj)) return(invisible(c(0,0)))
    if (is.factor(r.obj)) r.obj = as.character(r.obj)
    xl.range = xl.rng[['Application']]$range(xl.rng$cells(1,1),xl.rng$cells(1,length(r.obj)))
    nas = is.na(r.obj)
    # if (!is.numeric(r.obj)) r.obj[nas] = na
    r.list = as.list(r.obj)
    r.list[nas] = na
    xl.range[['Value']] = r.list
    invisible(c(1,length(r.obj)))
}


##########################

xl.raw.write = function(r.obj,xl.rng,na = ""){
    UseMethod('xl.raw.write')
}



xl.raw.write.default = function(r.obj,xl.rng,na = "")
    ### writes vectors (one-dimensional objects)
{
    if (is.null(r.obj)) return(invisible(c(0,0)))
    nas = is.na(r.obj)
    if (is.character(r.obj) || all(nas)){
        xl.range = xl.rng[['Application']]$range(xl.rng$cells(1,1),xl.rng$cells(length(r.obj),1))
        r.obj[nas] = na
        if (all(r.obj == ""))	{
            xl.range$ClearContents()
        } else {
            xl.range = xl.rng[['Application']]$range(xl.rng$cells(1,1),xl.rng$cells(length(r.obj),1))
            xl.range[['Value']] = asCOMArray(r.obj)
        }
    } else	{
        if (!any(nas)){
            xl.range = xl.rng[['Application']]$range(xl.rng$cells(1,1),xl.rng$cells(length(r.obj),1))
            xl.range[['Value']] = asCOMArray(r.obj)
        } 
        else return(xl.raw.write.matrix(as.matrix(r.obj),xl.rng))
    }
    invisible(c(length(r.obj),1))
}

xl.raw.write.POSIXct = function(r.obj,xl.rng,na = ""){
    xl.raw.write(format(r.obj, usetz = FALSE),xl.rng,na)
}

xl.raw.write.POSIXlt = function(r.obj,xl.rng,na = ""){
    xl.raw.write(format(r.obj, usetz = FALSE),xl.rng,na)
}

xl.raw.write.matrix = function(r.obj,xl.rng,na = "")
    ### insert matrix into excel sheet without column and row names
{
    # xl.range = xl.sheet$range(xl.sheet$cells(xl.row,xl.col),xl.sheet$cells(xl.row+NROW(r.obj)-1,xl.col))
    if (is.null(r.obj)) return(invisible(c(0,0)))
    excel = xl.rng[['Application']]
    xl.range = excel$range(xl.rng$cells(1,1),xl.rng$cells(nrow(r.obj),ncol(r.obj)))
    nas = is.na(r.obj)
    if (is.numeric(r.obj)){
        if (!any(nas)) {
            xl.range[["Value"]] = asCOMArray(r.obj)
        } else if (all(nas)) {
            if (na == "") {
                xl.range$clearcontents() 
            } else {
                xl.range[['Value']] = matrix(na,nrow = NROW(r.obj),ncol = NCOL(r.obj))
            }
        } else {
            on.exit(excel[["DisplayAlerts"]] <- TRUE)
            excel[["DisplayAlerts"]] = FALSE
            xl.range = xl.range[["Columns"]][[1]]
            # further code for NA's pasting correction
            r.obj[nas] = na
            if (is.vector(r.obj)) r.obj = as.matrix(r.obj)
            # TextToColumns used to avoid problem with "Numbers stored as text"
            # There is no obvious way to convert such numbers to correct format.
            iter = 1:ncol(r.obj)
            block = 1000
            while(length(iter)>0){
                if (length(iter)>block){
                    temp = apply(r.obj[,iter[1:block],drop = FALSE],1,paste,collapse = "\t")
                    iter = iter[-(1:block)]
                } else {
                    temp = apply(r.obj[,iter,drop = FALSE],1,paste,collapse = "\t")
                    iter = numeric(0)
                }
                xl.range[['Value']] = asCOMArray(temp)
                xlDelimited = 1
                xlDoubleQuote = 1
                xl.range$TextToColumns(Destination = xl.range, 
                                       DataType = xlDelimited,TextQualifier = xlDoubleQuote,ConsecutiveDelimiter = FALSE,
                                       Tab = TRUE,Semicolon = FALSE,Comma = FALSE,Space = FALSE,Other = FALSE,FieldInfo = c(1,1),
                                       DecimalSeparator = ".",TrailingMinusNumbers = TRUE)
                if (length(iter)>0) xl.range = xl.range$offset(0,block)	
            }	
        }	
    } else if (is.character(r.obj) || all(nas)) {
        r.obj[nas] = na
        if (all(r.obj == "")) xl.range$clearcontents() else xl.range[["Value"]] = asCOMArray(r.obj)
    } else {	
        xl.range[["Value"]] = asCOMArray(r.obj)
        if (any(nas)){
            lapply(1:ncol(nas),function(column) {
                na.in.column = which(nas[,column])
                if (length(na.in.column)>0){
                    lapply(na.in.column,function(na.in.row){
                        xl.range = xl.rng$cells(na.in.row,column)
                        xl.range[["Value"]] = na
                    })
                }
            })
        }
    }
    # TextToColumns Destination: = Range("A5"), DataType: = xlDelimited, _
    # TextQualifier: = xlDoubleQuote, ConsecutiveDelimiter: = False, Tab: = True, _
    # Semicolon: = False, Comma: = False, Space: = False, Other: = False, FieldInfo _
    # : = Array(1, 1), TrailingMinusNumbers: = True
    invisible(c(nrow(r.obj),ncol(r.obj)))
}


split_labels = function(x, remove_repeated = TRUE, split = "|", fixed = TRUE, perl = FALSE){
    if(length(x)==0){
        return(matrix(NA, ncol=0, nrow = 0))
    }
    if(length(x)==1 && x[1]=="") x[1] = " "
    x_split = strsplit(x, split = split, fixed = fixed, perl = perl)
    max_length = max(lengths(x_split))
    x_split = lapply(x_split, function(each) {
        if(length(each)<max_length){
            each = c(each, rep(NA, max_length - length(each)))
        }
        each
    })
    res = do.call(rbind, x_split)
    res[is.na(res)] = ""
    if (remove_repeated){
        for(i in rev(seq_len(nrow(res))[-1])){
            repeats = res[i-1, ] ==  res[i, ]
            first_no_repeat = which(!repeats)[1]-1
            if(is.na(first_no_repeat)) first_no_repeat = NCOL(res)
            if(!is.na(first_no_repeat) && first_no_repeat>0){
                res[i, 1:first_no_repeat] = ""
            }
            
        }
    }
    res    
}



#########################


fix_encoding = function(x){
    UseMethod("fix_encoding")
}

#' @export
fix_encoding.data.frame = function(x){
    for (each in seq_along(x)){
        x[[each]] = fix_encoding(x[[each]])
    }
    colnames(x) = fix_encoding(colnames(x))
    rownames(x) = fix_encoding(rownames(x))
    x
}


#' @export
fix_encoding.matrix = function(x){
    if(is.character(x) || is.factor(x)){
        
        res = matrix(enc2native(x), nrow = nrow(x), ncol = ncol(x))
        if(!is.null(colnames(x))) colnames(res) = fix_encoding(colnames(x))
        if(!is.null(rownames(x))) rownames(res) = fix_encoding(rownames(x))
        return(res)
    } else {
        if(!is.null(colnames(x))) colnames(x) = fix_encoding(colnames(x))
        if(!is.null(rownames(x))) rownames(x) = fix_encoding(rownames(x))
        return(x)
    }
}

#' @export
fix_encoding.default = function(x){
    if(!is.null(names(x))){
        names(x) = fix_encoding(names(x))
    }
    if(is.character(x)){
        enc = Encoding(x)
        if(any(enc %in% c("UTF-8", "UTF8"))){
            Encoding(x) = "UTF-8"
            x = enc2native(x)
        }
    }
    x
}

#' @export
fix_encoding.list = function(x){
    for (each in seq_along(x)){
            x[[each]] = fix_encoding(x[[each]])
    }
    if(!is.null(names(x))){
        names(x) = fix_encoding(names(x))
    }
    x
}
