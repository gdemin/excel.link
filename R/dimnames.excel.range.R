#' @export
dimnames.excel.range = function(x){
    xl.dimnames(x())
}

#' @export
dim.excel.range = function(x){
    xl.rng = x()
    c(xl.nrow(xl.rng),xl.ncol(xl.rng))
}

#' @export
'dim<-.excel.range' = function(x, value){
    stop("'dim' for excel.range is read-only.")
}

#' @export
'dimnames<-.excel.range' = function(x, value){
    stop("'dimnames' for excel.range is read-only.")
}


xl.colnames.excel.range = function(xl.rng,...)
    # return colnames of connected excel table
{
    if (has.colnames(xl.rng)){
        all.colnames = unname(unlist(xl.read.range(xl.rng[['rows']][[1]])))
        all.colnames = gsub("^([\\s]+)","",all.colnames,perl = TRUE)
        all.colnames = gsub("([\\s]+)$","",all.colnames,perl = TRUE)
    } else all.colnames = xl.colnames(xl.rng)
    if (has.rownames(xl.rng)) all.colnames = all.colnames[-1]
    return(all.colnames)
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
        all.rownames = xl.read.range(xl.rng[['columns']][[1]])
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





xl.nrow = function(xl.rng){
    res = xl.rng[["Rows"]][["Count"]]
    res-has.colnames(xl.rng)
}


xl.ncol = function(xl.rng){
    res = xl.rng[["Columns"]][["Count"]]
    res-has.rownames(xl.rng)
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