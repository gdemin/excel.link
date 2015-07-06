# Idea by Stefan Fritsch (https://github.com/gdemin/excel.link/issues/1) 


#' Title
#'
#' @param sym 
#' @param str.range 
#' @param drop 
#' @param na 
#' @param row.names 
#' @param col.names 
#' @param env 
#'
#' @return sdsd

#'
#' @export
xl.bind.range = function(sym, str.range, drop = TRUE, na = "", row.names = FALSE, col.names = FALSE, env = parent.frame())
{
    
    if (exists(sym, env)) remove(list=sym, envir=env)
    
    
    xl = xl.get.excel()
    xl.rng = xl$Range(str.range)
    assignment = function(value=NULL){
        if (missing(value)) {
            xl.read.range(xl.rng, drop=drop, row.names=row.names, col.names=col.names, na = na)
        } else {
            if(is.null(value)){
                cat(xl.rng$address(External = TRUE))        
            } else {
                xl.rng$clear()
                if(is.atomic(value) && length(value)<2 && is.null(attributes(value))){
                    res = xl.write(value,xl.rng$cells(1,1), na = na, row.names = row.names, col.names = col.names)
                } else {
                    res = xl.write(value,xl.rng, na = na, row.names = row.names, col.names = col.names)
                } 
                if (res[1]>0) res[1] = res[1] - 1
                if (res[2]>0) res[2] = res[2] - 1
                xl.rng <<- xl$range(xl.rng$cells(1,1),xl.rng$cells(1,1)$offset(res[1],res[2]))
            }
        }
    }
    # assign active binding:   
    makeActiveBinding(
        sym,
        assignment,
        env
    )
    
}

#' @export
xl.bind.current.region = function(sym, str.range, drop = TRUE, na = "", row.names = FALSE, col.names = FALSE, env = parent.frame())
{
    
    if (exists(sym, env)) remove(list=sym, envir=env)
    xl = xl.get.excel()
    xl.rng = xl$Range(str.range)
    assignment = function(value){
        curr.rng = xl.rng$CurrentRegion()
        if (missing(value)) {
            xl.read.range(curr.rng, drop=drop, row.names=row.names, col.names=col.names, na = na)
        } else 
        {
            if(is.null(value)){
                
                cat(xl.rng$currentregion()$address(External = TRUE))
            } else {
                curr.rng$clear()
                xl.write(value,curr.rng$cells(1,1), na = na, row.names = row.names, col.names = col.names)
                
            }
        }
    }
    
    # assign active binding:   
    makeActiveBinding(
        sym,
        assignment,
        env
    )
    
}

bind.generator = function(row.names, col.names, fun) {
    function(x, value){
        x = deparse(substitute(x))
        value = substitute(value)
        if (!is.character(value)) value = deparse(value)
        env = parent.frame()
        fun(x, value, row.names = row.names, col.names = col.names, env = env)
        
    }
}

#' @export
"%=cr%" = bind.generator(row.names = FALSE, col.names = FALSE, fun = xl.bind.current.region)
#' @export
"%=crr%" = bind.generator(row.names = TRUE, col.names = FALSE, fun = xl.bind.current.region)
#' @export
"%=crc%" = bind.generator(row.names = FALSE, col.names = TRUE, fun = xl.bind.current.region)
#' @export
"%=crrc%" = bind.generator(row.names = TRUE, col.names = TRUE, fun = xl.bind.current.region)

#' @export
"%=xl%" = bind.generator(row.names = FALSE, col.names = FALSE, fun = xl.bind.range)
#' @export
"%=xlr%" = bind.generator(row.names = TRUE, col.names = FALSE, fun = xl.bind.range)
#' @export
"%=xlc%" = bind.generator(row.names = FALSE, col.names = TRUE, fun = xl.bind.range)
#' @export
"%=xlrc%" = bind.generator(row.names = TRUE, col.names = TRUE, fun = xl.bind.range)

#' @export
"%<-cr%" = `%=cr%`
#' @export
"%<-crr%" = `%=crr%`
#' @export
"%<-crc%" = `%=crc%`
#' @export
"%<-crrc%" = `%=crrc%`

#' @export
"%<-xl%" = `%=xl%`
#' @export
"%<-xlr%" = `%=xlr%`
#' @export
"%<-xlc%" = `%=xlc%`
#' @export
"%<-xlrc%" = `%=xlrc%`

#' @export
xl.binding.address = function(sym){
    sym = substitute(sym)
    if (!is.character(sym)) sym = deparse(sym)
    eval(parse(text = paste0("capture.output(",sym,"<-NULL)")),envir = parent.frame())
}

