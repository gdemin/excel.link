# Package: RDCOMClient
# Version: 0.93-0.2
# Title: R-DCOM Client
# Author: Duncan Temple Lang <duncan@wald.ucdavis.edu>
#     Maintainer: Duncan Temple Lang <duncan@wald.ucdavis.edu>
#     Description: Provides dynamic client-side access to (D)COM applications from within R.
# License: GPL-2
# Collate: classes.R COMLists.S COMError.R com.R debug.S zzz.R runTime.S
# URL: http://www.omegahat.org/RDCOMClient, http://www.omegahat.org
# http://www.omegahat.org/bugs



#' @export
#' @rdname RDCOMClient
setClass("COMList", representation("COMIDispatch"))

#' @export
#' @rdname RDCOMClient
COMList =
function(obj, class = "COMList")
{
  new(class, ref = obj@ref)	
}				

#' @export
#' @rdname RDCOMClient
setMethod("length", "COMList",
           function(x) .COM(x, "Count"))

#' @export
#' @rdname RDCOMClient
setMethod("[[", c("COMList", "numeric"),
            function(x, i, j, ...) {
               if(length(i) != 1)
                stop("COMList[[ ]] requires exactly one index")

              .COM(x,"Item", as.integer(i)) 
            }) 

#' @export
#' @rdname RDCOMClient
setMethod("[[<-", c("COMList", "numeric"),
            function(x, i, j, ..., value) {
               if(i < 0)
                stop("COMList[[ ]] requires a positive index")

                # This is probably not a good thing to try.
                # Just here out of curiosity.
               if(i == .COM(x, "Count") + 1) {
                 .COM(x, "Add", value)
               }

               x

            }) 

#' @export
#' @rdname RDCOMClient
setMethod("length", "COMList", 
	    function(x)  .COM(x, "Count"))


#' @export
#' @rdname RDCOMClient    
setGeneric("lapply", function(X, FUN, ...) standardGeneric("lapply"))

#' @export
#' @rdname RDCOMClient
setGeneric("sapply", 
	      function(X, FUN, ..., simplify = TRUE, USE.NAMES = TRUE) 
	         standardGeneric("sapply"))

	
#' @export
#' @rdname RDCOMClient			
setMethod("lapply", "COMList",
            function(X, FUN, ...) {
              lapply(1:length(X),
                       function(id)
                          FUN(X[[id]], ...))
            })

#' @export
#' @rdname RDCOMClient
setMethod("sapply", "COMList",
function (X, FUN, ..., simplify = TRUE, USE.NAMES = TRUE) 
{
    FUN <- match.fun(FUN)
    answer <- lapply(X, FUN, ...)
    if (USE.NAMES && is.character(X) && is.null(names(answer))) 
        names(answer) <- X
    if (simplify && length(answer) && length(common.len <- unique(unlist(lapply(answer, 
        length)))) == 1) {
        if (common.len == 1) 
            unlist(answer, recursive = FALSE)
        else if (common.len > 1) 
            array(unlist(answer, recursive = FALSE), dim = c(common.len, 
                length(X)), dimnames = if (!(is.null(n1 <- names(answer[[1]])) & 
                is.null(n2 <- names(answer)))) 
                list(n1, n2))
        else answer
    }
    else answer
})

#' @export
#' @rdname RDCOMClient
setMethod("lapply", "COMIDispatch",
         function (X, FUN, ...)  {
           lapply(new("COMList", X), FUN, ...)
  	 })	   
	
#' @export
#' @rdname RDCOMClient
setMethod("sapply", "COMIDispatch",
         function (X, FUN, ..., simplify = TRUE, USE.NAMES = TRUE)  {
           sapply(new("COMList", X), FUN, ..., simplify = simplify, USE.NAMES = TRUE)
  	 })	   

#' @export
#' @rdname RDCOMClient
setClass("COMTypedList", contains = "COMList")

# This method gets the name of the class for the returned value of
# an item in the list. This allows the [[ method to be inherited
# directly by COMTypedNamedList from COMTypedList but to behave
# differently.

#' @export
#' @rdname RDCOMClient
setGeneric("getItemClassName", 
             function(x)  standardGeneric("getItemClassName"))

#' @export
#' @rdname RDCOMClient
setMethod("getItemClassName", "COMTypedList", function(x) gsub("s$", "", class(x)))

#' @export
#' @rdname RDCOMClient
setMethod("[[", c("COMTypedList", "numeric"),
            function(x, i, j, ...) {
              val = callNextMethod()
              new(getItemClassName(x), val)
            })

#' @export
#' @rdname RDCOMClient
setClass("COMTypedNamedList", representation(name = "character"), contains = "COMTypedList")

#' @export
#' @rdname RDCOMClient
setClass("COMTypedParameterizedNamedList", representation(nameProperty = "character"), contains = "COMTypedNamedList")
setValidity("COMTypedParameterizedNamedList",
             function(object) {
                 if(length(object@nameProperty) == 0)
                   "nameProperty must be specified"

                 TRUE
             })

#' @export
#' @rdname RDCOMClient
setMethod("names", "COMTypedParameterizedNamedList", 
           function(x) {
              sapply(x, function(el) el[[x@nameProperty]])
           })

#' @export
#' @rdname RDCOMClient
setMethod("[[", c("COMTypedList", "character"),
            function(x, i, j, ...) {
              val = callNextMethod()
              new(getItemClassName(x), val)
            })	


#' @export
#' @rdname RDCOMClient
setMethod("getItemClassName", "COMTypedNamedList", function(x) x@name)

 # This version ends up calling all sorts of methods and 
if(FALSE)  {
  setMethod("names", c("COMTypedNamedList"),
              function(x) {
	      sapply(x, function(el) el[["Name"]])
            })
}

	# Alternative, "faster" way of doing this.
#' @export
#' @rdname RDCOMClient
setMethod("names", c("COMTypedNamedList"),
            function(x) {
	      n = x$Count
	      if(n == 0)
	         return(character())

	      ans = character(n)
              it = x$Item
	      for(i in 1:n)
                ans[i] = it(i)$Name
	
	      ans
            })
