#########################################################
#########################################################
## Gregory Demin, 2011 <excel.link.feedback@gmail.com> ##
#########################################################
#########################################################
##### ver 0.5 - 19.11.2011 - Initial release ###########
##### ver 0.5.2 - 23.03.2012- Bugfix release ###########
##### ver 0.5.3 - 17.07.2012- Bugfix release ###########
##### ver 0.5.4 - 13.08.2012- Bugfix release ###########
########## 1. Current graphics now insert picture instead of link to the file
########## 2. aaa[condition,new.var] = data now works properly
##### ver 0.5.5 - 12.04.2013- Compatibilty fix for R3.0 release ###########
##### ver 0.5.6 - 14.04.2013- names_to_matrix fix ###########
##### ver 0.6 - 26.08.2013- xl.read.file, xl.save.file, filename argument in current.graphics ###########
##### ver 0.6.1 - 12.07.2014- Release on GitHub ###########
##### ver 0.7 - 12.07.2014- ???? ###########

.onAttach = function(...) {
  packageStartupMessage("\nTo Daniela Khazova who constantly inspires me...")
}


xl.get.excel = function()
  # run Excel if it's not running and
  # return reference to Microsoft Excel
{
  xls = COMCreate("Excel.Application")
  if (xls[['workbooks']][['count']] == 0) xls[['workbooks']]$add()
  # xls = getCOMInstance("{000208D5-0000-0000-C000-000000000046}")
  if (!xls[["Visible"]]) xls[["Visible"]] = TRUE
  return(xls)
}
has.colnames = function(x){
  UseMethod("has.colnames")
}

has.rownames = function(x){
  UseMethod("has.rownames")
}

has.colnames.default = function(x)
  # get attribute has.colnames
{
  res = attr(x,'has.colnames')
  if (is.null(res)) res = FALSE
  res
}

has.rownames.default = function(x)
  # get attribute has.rownames
{
  res = attr(x,'has.rownames')
  if (is.null(res)) res = FALSE
  res
}

"has.colnames<-" = function(x,value)
  # set attribute has.colnames
{
  attr(x,'has.colnames') = value
  x
}

"has.rownames<-" = function(x,value)
  # set attribute has.rownames
{
  attr(x,'has.rownames') = value
  x
}

has.colnames.excel.range = function(x)
  # get attribute has.colnames
{
  res = attr(x(),'has.colnames')
  if (is.null(res)) res = FALSE
  res
}


has.rownames.excel.range = function(x)
  # get attribute has.rownames
{
  res = attr(x(),'has.rownames')
  if (is.null(res)) res = FALSE
  res
}



xl = function()
  # run Excel if it's not running and
  # return reference to Microsoft Excel
{
  xl.get.excel()
}

# set class for usage '.[', '.[ = ' etc operators
class(xl) = c('xl',class(xl))
xlrc = xlr = xlc = xl
has.rownames(xl) = FALSE 
has.colnames(xl) = FALSE 

has.rownames(xlc) = FALSE 
has.colnames(xlc) = TRUE 

has.rownames(xlr) = TRUE 
has.colnames(xlr) = FALSE 

has.rownames(xlrc) = TRUE 
has.colnames(xlrc) = TRUE 
# class(xlrc) = class(xlr) = class(xlc) = c('xlrc',class(xl))


'[.xl' = function(x,str.rng,drop = !(has.rownames(x) | has.colnames(x)),na = "")
  ### return range from Microsoft Excel. range.name is character string in form of standard
  ### Excel reference, quotes can be omitted, e. g. [A1:B5], [Sheet1!F8], [[Book3]Sheet7!B1] or range name 
  ### Function is intended to use in interactive environement 
{
  # str.rng = as.character(sys.call())[3]
  str.rng = as.character(as.expression(substitute(str.rng)))
  x[[str.rng,drop = drop,na = na]]
}


'[[.xl' = function(x,str.rng,drop = !(has.rownames(x) | has.colnames(x)),na = "")
  ### return range from Microsoft Excel. range.name is character string in form of standard
  ### Excel reference, e. g. ['A1:B5'], ['Sheet1!F8'], ['[Book3]Sheet7!B1'] or range name 
  ### The difference with '[' is that value should be quoted string. It's intended to use in user define functions
  ### or in cases where value is string variable with Excel range 
{
  xl.rng = x()$Range(str.rng)
  xl.read.range(xl.rng,drop = drop,row.names = has.rownames(x),col.names = has.colnames(x),na = na)
}


'$.xl' = function(x,str.rng)
  ### return range from Microsoft Excel. range.name is character string in form of standard
  ### Excel reference, e. g. xl$'A1:B5', xl$'Sheet1!F8', xl$'[Book3]Sheet7!B1', xl$h3 or range name 
  ### The difference with '[' is that value should be quoted string. It's intended to use in user define functions
  ### or in cases where value is string variable with Excel range 
{
  x[[str.rng]]
}



'[[<-.xl' = function(x,str.rng,na = "",value)
{
  xl.rng = x()$Range(str.rng)
  xl.write(value,xl.rng,row.names = has.rownames(x),col.names = has.colnames(x),na = na)
  x
}



'$<-.xl' = function(x,str.rng,value)
{
  x[[str.rng]] = value
  x
}


'[<-.xl' = function(x,str.rng,na = "",value)
{
  str.rng = as.character(as.expression(substitute(str.rng)))
  x[[str.rng,na = na]] = value
  x
}


make.me.quick = function(app){
  app[["Screenupdating"]] = FALSE
  app[["Calculation"]] = -4135 # xlCalculationManual = -4135
  invisible(NULL)
}


make.me.slow = function(app){
  app[["Screenupdating"]] = TRUE
  app[["Calculation"]] = -4105 # xlCalculationAutomatic = -4105
  invisible(NULL)
}


xl.write = function(r.obj,xl.rng,na = "",...)
  ## insert values in excel range.
  ## shoul return c(row,column) - next emty point
{
  app = xl.rng[["Application"]]
  on.exit(make.me.slow(app))
  make.me.quick(app)
  UseMethod("xl.write")
}


current.graphics = function(type = c("png","emf","jpeg","bmp","tiff"),filename = NULL,...){
  if (is.null(filename)){
    # 		if (!('windows' %in% names(dev.cur()))) stop("there is no graphics on windows device.")
    type = match.arg(type)
    res = paste(tempfile(),".",type,sep = "")
    switch(type,
           png = dev.copy(png,res,...),
           emf = dev.copy(win.metafile,res,...),
           jpeg = dev.copy(jpeg,res,...),
           bmp = dev.copy(bmp,res,...),
           tiff = dev.copy(tiff,res,...)
    )
    dev.off()
    # 		savePlot(filename = res,type = type,...)
    attr(res,"temp.file") = TRUE
  } else {
    res = normalizePath(filename,mustWork = TRUE)
  }
  class(res) = "current.graphics"
  res
}


temp.file = function(r.obj)
  # auxiliary function
  # return TRUE if object has attribute "temp.file" with value TRUE
  # in other cases return FALSE
{
  temp.file = attr(r.obj,"temp.file")
  !is.null(temp.file) && temp.file
}

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
  pic = app[["Activesheet"]][['Shapes']]$AddPicture(unclass(r.obj),0,-1,left,top,width,height)
  fill = pic[['Fill']] # [["Shaperange"]]
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



xl.write.ctable = function(r.obj,xl.rng,na = "",row.names = TRUE,col.names = TRUE,...){
  xl.colnames = t(names_to_matrix(colnames(r.obj)))
  xl.rownames = names_to_matrix(rownames(r.obj))
  has.col = ifelse(!is.null(xl.colnames) && col.names,nrow(xl.colnames),0)
  has.row = ifelse(!is.null(xl.rownames) && row.names,ncol(xl.rownames),0)
  if ((row.names | col.names)){
    # clear output area
    out.rng = xl.rng[['Application']]$range(xl.rng$cells(1,1),xl.rng$cells(nrow(r.obj)+has.col,ncol(r.obj)+has.row))
    out.rng$clear()
  }
  if (has.col) {
    xl.raw.write(xl.colnames,xl.rng$offset(0,has.row),na)
  }	
  if (has.row) {
    xl.raw.write(xl.rownames,xl.rng$offset(has.col,0),na)
  }	
  # for (i in seq_len(ncol(r.obj)))	xl.raw.write(r.obj[,i],xl.rng$offset(has.col,i+has.row-1),na)
  xl.raw.write.matrix(r.obj,xl.rng$offset(has.col,has.row),na)
  invisible(c(nrow(r.obj)+has.col,ncol(r.obj)+has.row))
}



names_to_matrix = function(name,splitter = "|")
  # Convert rownames/colnames with items delimited by symbol 'splitted'
  # to matrix with each label in it's own cell and remove sequentally repeated
  # labels.
{
  if (!is.null(name)){
    splitted = strsplit(as.character(name),split = splitter,fixed = TRUE)
    el.len = max(unlist(lapply(splitted,length)))
    column.nums = 1:el.len
    res = vapply(splitted,"[",column.nums,FUN.VALUE = as.character(column.nums))
    if (!is.matrix(res)) res = t(res)
    res[is.na(res)] = ""
    width = NCOL(res)
    if ((NROW(res)>0) & (width>1)){
      boundary = logical(width-1)
      for (i in (1:NROW(res))){
        boundary = boundary | (res[i,1:(width-1)] !=  res[i,2:width])
        res[i,c(FALSE,!boundary)] = ""
      }
    }
    t(res)
  } else NULL
}


xl.write.matrix = function(r.obj,xl.rng,na = "",row.names = TRUE,col.names = TRUE,...)
  ## insert matrix into excel sheet including column and row names
{
  if (!is.null(r.obj)){
    xl.colnames = colnames(r.obj)
    xl.rownames = rownames(r.obj)
    has.col = (!is.null(xl.colnames) && col.names)*1
    has.row = (!is.null(xl.rownames) && row.names)*1
    dim.names = names(dimnames(r.obj))
    has.dim.names = (!is.null(dim.names))*1
    if ((row.names & col.names) | (has.dim.names & (row.names | col.names))){
      # clear output area
      out.rng = xl.rng[['Application']]$range(xl.rng$cells(1,1),xl.rng$cells(nrow(r.obj)+has.col+has.dim.names,ncol(r.obj)+has.row+has.dim.names))
      out.rng$clear()
    }
    if (has.col) {
      if (has.dim.names){
        has.row = has.row+1
        xl.raw.write(dim.names[2],xl.rng$offset(0,has.row),na)
      }
      xl.raw.write(t(xl.colnames),xl.rng$offset(has.dim.names,has.row),na)
    }	
    if (has.row) {
      if (has.dim.names){
        has.col = has.col+1
        xl.raw.write(dim.names[1],xl.rng$offset(has.col,0),na)
      }	
      xl.raw.write(xl.rownames,xl.rng$offset(has.col,has.dim.names),na)
    }	
    # for (i in seq_len(ncol(r.obj)))	xl.raw.write(r.obj[,i],xl.rng$offset(has.col,i+has.row-1),na)
    xl.raw.write.matrix(r.obj,xl.rng$offset(has.col,has.row),na)
  }
  invisible(c(nrow(r.obj)+has.col+has.dim.names,ncol(r.obj)+has.row+has.dim.names))
}



xl.write.data.frame = function(r.obj,xl.rng,na = "",row.names = TRUE,col.names = TRUE,...)
  ## insert data.frame into excel sheet including column and row names
{
  # stop("Multi-column (e. g. matrix) data.frame elements currently not supported.")
  if (!is.null(r.obj)){
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



xl.write.default = function(r.obj,xl.rng,na = "",row.names = TRUE,...){
  if (is.null(r.obj) || length(r.obj) == 0) r.obj = ""
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


xl.write.factor = function(r.obj,xl.rng,na = "",row.names = TRUE,...){
  r.obj = as.character(r.obj)
  xl.write.default(r.obj,xl.rng = xl.rng,na = na,row.names = row.names,...)
}


xl.write.table = function(r.obj,xl.rng,na = "",...){
  if(length(dim(r.obj))<3) {
    # if (!is.null(dimnames(r.obj)) && all(names(dimnames(r.obj)) %in% c("",NA))) names(dimnames(r.obj)) = NULL
    if(length(dim(r.obj))<2) {
      return(invisible(xl.write.matrix(as.matrix(r.obj),xl.rng,na,row.names = TRUE,col.names = TRUE)))
    } else return(invisible(xl.write.matrix(as.matrix(r.obj),xl.rng,na,row.names = TRUE,col.names = TRUE)))
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



xl.writerow = function(r.obj,xl.rng,na = "")
  ## special function for writing single row on excel sheet
{
  if (is.factor(r.obj)) r.obj = as.character(r.obj)
  xl.range = xl.rng[['Application']]$range(xl.rng$cells(1,1),xl.rng$cells(1,length(r.obj)))
  nas = is.na(r.obj)
  # if (!is.numeric(r.obj)) r.obj[nas] = na
  r.list = as.list(r.obj)
  r.list[nas] = na
  xl.range[['Value']] = r.list
  invisible(c(1,length(r.obj)))
}



xl.raw.write = function(r.obj,xl.rng,na = ""){
  UseMethod('xl.raw.write')
}



xl.raw.write.default = function(r.obj,xl.rng,na = "")
  ### writes vectors (one-dimensional objects)
{
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



xl.raw.write.matrix = function(r.obj,xl.rng,na = "")
  ### insert matrix into excel sheet without column and row names
{
  # xl.range = xl.sheet$range(xl.sheet$cells(xl.row,xl.col),xl.sheet$cells(xl.row+NROW(r.obj)-1,xl.col))
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



xl.selection = function(drop = TRUE,na = "",row.names = FALSE,col.names = FALSE)
  # return current selection from Microsoft Excel
{
  ex = xl.get.excel()
  xl.rng = ex[['Selection']]
  xl.read.range(xl.rng,drop = drop,na = na,row.names = row.names,col.names = col.names)
}



xl.current.region = function(str.rng,drop = TRUE,na = "",row.names = FALSE,col.names = FALSE)
  # return current region from Microsoft Excel (region selected when pressing Ctrl+Shift+*)
{
  ex = xl.get.excel()
  xl.rng = ex$range(str.rng)
  xl.read.range(xl.rng[["CurrentRegion"]],drop = drop,na = na,row.names = row.names,col.names = col.names)
}



xl.read.range = function(xl.rng,drop = TRUE,row.names = FALSE,col.names = FALSE,na = "")
  # return matrix/data.frame/vector from excel from given range
{
  if (col.names && (xl.rng[["rows"]][["count"]]<2)) col.names = FALSE
  if (row.names && (xl.rng[["columns"]][["count"]]<2)) row.names = FALSE
  raw.res = xl.rng[['Value2']]
  if (is.null(raw.res)) data.list = NA else data.list = xl.process.list(raw.res,na = na)
  if (col.names)	{
    colNames = lapply(data.list,"[[",1)
    if (row.names) colNames = colNames[-1]
    data.list = lapply(data.list,"[",-1)
  }
  if (row.names) {
    rowNames = unlist(data.list[[1]])
    data.list = data.list[-1]
  }	
  data.list = lapply(data.list,unlist)
  classes = unique(sapply(data.list,class))
  final.matrix = do.call(data.frame,list(data.list,stringsAsFactors = FALSE))
  if (row.names && anyDuplicated(rowNames)) {
    row.names = FALSE
    warning("There are duplicated rownames. They will be ignored.")
  }	
  if (row.names) rownames(final.matrix) = rowNames else rownames(final.matrix) = xl.rownames(xl.rng)[ifelse(col.names,-1,TRUE)]
  if (col.names) colnames(final.matrix) = colNames else colnames(final.matrix) = xl.colnames(xl.rng)[ifelse(row.names,-1,TRUE)]
  if (ncol(final.matrix)<2 & drop) final.matrix = final.matrix[,1]
  final.matrix
}


xl.process.list = function(data.list,na = "")
  ## intended for processing list from Excel
  ## it's replace NULL's, "" and zero-length elements with NA
{
  lapply(data.list, function(each.col) {
    # each.col = gsub("^[\\t\\s]+$","",each.col,perl = TRUE)
    for.na = unlist(lapply(each.col,function(each.cell) isS4(each.cell) || is.null(each.cell) || length(each.cell) == 0 || each.cell == na))
    each.col[for.na ] = NA # | grepl("^[\\t\\s]+$",each.col,perl = TRUE)
    each.col
  })
}



xl.workbook.add = function(filename = NULL)
  ### add new workbook and invisibily return it's name
  ### if filename is give, its used as template 
{
  ex = xl.get.excel()
  if (!is.null(filename)) {
    path = normalizePath(filename,mustWork = TRUE)
    xl.wb = ex[['Workbooks']]$Add(path) 
  } else xl.wb = ex[['Workbooks']]$Add()
  invisible(xl.wb[["Name"]])
}



xl.workbooks = function()
  ## names of all opened workbooks
{
  ex = xl.get.excel()
  wb.count = ex[['Workbooks']][['Count']]
  sapply(seq_len(wb.count), function(wb) ex[['Workbooks']][[wb]][['Name']])
}

xl.workbook.open = function(filename)
  ## open workbook
{
  ex = xl.get.excel()
  xl.wb = ex[["Workbooks"]]$Open(normalizePath(filename,mustWork = TRUE))
  invisible(xl.wb[['Name']])
}



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



xl.workbook.save = function(filename)
  ### save active workbook under the different name. If path is missing it saves in working directory
  ### doesn't alert if it owerwrite other file
{
  ex = xl.get.excel()
  # if (is.null(filename)) filename = ex[["ActiveWorkbook"]][["Name"]]
  path = normalizePath(filename,mustWork = FALSE)
  on.exit(ex[["DisplayAlerts"]] <- TRUE)
  ex[["DisplayAlerts"]] = FALSE
  ex[["ActiveWorkbook"]]$SaveAs(path)
  invisible(path)
}



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



xl.sheets = function()
  ### Return worksheets names in the active workbook 
{
  ex = xl.get.excel()
  sh.count = ex[['ActiveWorkbook']][['Sheets']][['Count']]
  sapply(seq_len(sh.count), function(sh) ex[['ActiveWorkbook']][['Sheets']][[sh]][['Name']])
}



xl.sheet.exists = function(xl.sheet,all.sheets = xl.sheets())
  ## check exsistense of xl.sheet in all.sheets and return xl.sheet position in all.sheets 
{
  UseMethod("xl.sheet.exists")
}



xl.sheet.exists.numeric = function(xl.sheet,all.sheets = xl.sheets())
{
  if (xl.sheet>length(all.sheets)) stop ("too large sheet number. In workbook only ",length(all.sheets)," sheet(s)." )
  xl.sheet
}



xl.sheet.exists.character = function(xl.sheet,all.sheets = xl.sheets())
{
  xl.sheet = which(tolower(xl.sheet) == tolower(all.sheets)) 
  if (length(xl.sheet) == 0) stop ("sheet ",xl.sheet," doesn't exist." )
  xl.sheet
}



xl.sheet.add = function(xl.sheet.name = NULL,before = NULL)
  ### add new sheet to active workbook after the last sheet with given name and invisibily return reference to it 
{
  ex = xl.get.excel()
  sh.count = ex[['ActiveWorkbook']][['Sheets']][['Count']]
  sheets = tolower(xl.sheets())
  if (!is.null(xl.sheet.name) && (tolower(xl.sheet.name) %in% sheets)) stop ('sheet with name "',xl.sheet.name,'" already exists.')
  if (is.null(before)) {
    res = ex[['ActiveWorkbook']][['Sheets']]$Add(After = ex[['ActiveWorkbook']][['Sheets']][[sh.count]])
  } else {
    before = xl.sheet.exists(before,sheets)
    res = ex[['ActiveWorkbook']][['Sheets']]$Add(Before = ex[['ActiveWorkbook']][['Sheets']][[before]])
  }	
  if (!is.null(xl.sheet.name)) res[['Name']] = substr(xl.sheet.name,1,63)
  invisible(res[['Name']])
}



xl.sheet.delete = function(xl.sheet = NULL)
  ### delete sheet with given name(number) in active workbook 
{
  ex = xl.get.excel()
  on.exit(ex[["DisplayAlerts"]] <- TRUE)
  if (is.null(xl.sheet)) {
    xl.sh = ex[['ActiveWorkbook']][["ActiveSheet"]]
  } else {
    xl.sheet = xl.sheet.exists(xl.sheet)
    xl.sh = ex[['ActiveWorkbook']][['Sheets']][[xl.sheet]]
  }	
  ex[["DisplayAlerts"]] = FALSE
  xl.sh$Delete()
  invisible(NULL)
}


xl.sheet.activate = function(xl.sheet)
  ### activate sheet with given name (number) in active workbook 
{
  ex = xl.get.excel()
  #on.exit(ex[["DisplayAlerts"]] = TRUE)
  xl.sheet = xl.sheet.exists(xl.sheet)
  xl.sh = ex[['ActiveWorkbook']][['Sheets']][[xl.sheet]]
  #ex[["DisplayAlerts"]] = FALSE
  xl.sh$activate()
  invisible(xl.sh[['Name']])
}



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



xl.colnames.excel.range = function(xl.rng,...)
  # return colnames of connected excel table
{
  if (has.colnames(xl.rng)){
    all.colnames = unlist(xl.process.list(xl.rng[['rows']][[1]][['Value2']]))
    all.colnames = gsub("^([\\s]+)","",all.colnames,perl = TRUE)
    all.colnames = gsub("([\\s]+)$","",all.colnames,perl = TRUE)
  } else all.colnames = xl.colnames(xl.rng)
  if (has.rownames(xl.rng)) all.colnames = all.colnames[-1]
  return(all.colnames)
}


dimnames.excel.range = function(x){
  xl.dimnames(x())
}



xl.dimnames = function(xl.rng,...)
  ### x - references on excel range
{
  list(xl.rownames.excel.range(xl.rng),xl.colnames.excel.range(xl.rng))
}


xl.colnames = function(xl.rng)
  ## returns vector of Excel colnames, such as A,B,C etc.
{
  first.col = xl.rng[["Column"]]
  columns.count = xl.rng[["Columns"]][['Count']]
  columns = seq(first.col,length.out = columns.count)
  # columns = index3*26*26+index2*26+index1
  index1 = (columns-1) %% 26+1
  index2 = ifelse(columns<27,0,((columns - index1) %/% 26 -1) %% 26 + 1)
  index3 = ifelse(columns<(26*26+1),0,((columns-26*index2-index1) %/% (26 * 26) -1 ) %% 26 +1 )
  letter1 = letters[index1]	
  letter2 = ifelse(columns<27,"",letters[index2])	
  letter3 = ifelse(columns<(26*26+1),"",letters[index3])	
  paste(letter3,letter2,letter1,sep = "")
}


xl.rownames.excel.range = function(xl.rng,...)
  # return rownames of connected excel table
{
  if (has.rownames(xl.rng)){
    all.rownames = unlist(xl.process.list(xl.rng[['columns']][[1]][['Value2']]))
    all.rownames = gsub("^([\\s]+)","",all.rownames,perl = TRUE)
    all.rownames = gsub("([\\s]+)$","",all.rownames,perl = TRUE)
  } else all.rownames = xl.rownames(xl.rng)
  if (has.colnames(xl.rng)) all.rownames = all.rownames[-1]
  return(all.rownames)
}


xl.rownames = function(xl.rng)
  ## returns vector of Excel rownumbers.
{
  first.row = xl.rng[["Row"]]
  rows.count = xl.rng[["Rows"]][['Count']]
  seq(first.row,length.out = rows.count)
}


dim.excel.range = function(x){
  xl.rng = x()
  c(xl.nrow(xl.rng),xl.ncol(xl.rng))
}


xl.nrow = function(xl.rng){
  res = xl.rng[["Rows"]][["Count"]]
  res-has.colnames(xl.rng)
}


xl.ncol = function(xl.rng){
  res = xl.rng[["Columns"]][["Count"]]
  res-has.rownames(xl.rng)
}


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
  raw.data = lapply(colnumber,function(each.col) xl.process.list(xl.rng[['columns']][[each.col]][['Value2']],na = na))
  raw.data = lapply(raw.data,function(each.col) unlist(each.col[[1]][-1]))
  res = do.call(data.frame,list(raw.data,stringsAsFactors = FALSE))
  colnames(res) = all.colnames[colnumber-has.rownames(xl.rng)]
  # print(all.rownames)
  if (!anyDuplicated(all.rownames)) rownames(res) = all.rownames else warning("There are duplicated rownames. They will be ignored.")
  if(!missing(i)) res = res[i,,drop = FALSE]
  if (drop & (ncol(res)<2)) return(res[,1]) else return(res)
}


'$.excel.range' = function(x,value){
  x[,value,drop = TRUE]
}


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



'$<-.excel.range' = function(x,j,value){
  x[,j] = value
  invisible(x)
}



xl.read.file = function(filename, header = TRUE, row.names = NULL, col.names = NULL, xl.sheet = NULL,top.left.cell = "A1", na = "",excel.visible = FALSE)
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



xl.save.file = function(r.obj,filename, row.names = TRUE, col.names = TRUE, xl.sheet = NULL,top.left.cell = "A1", na = "",excel.visible = FALSE)
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

