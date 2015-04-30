#' @export
#' @rdname RDCOMClient
createCOMReference <-
    function(ref, className)
    {
        if(!isClass(className)) {
            className = "COMIDispatch"
            warning("Using COMIDispatch instead of ", className)
        }
        
        obj = new(className)
        obj@ref = ref
        
        obj
    }

#' @export
#' @rdname RDCOMClient
setClass("SCOMErrorInfo", representation(status="numeric",
                                         source="character",
                                         description="character"
))

#' @export
#' @rdname RDCOMClient
setClass("IUnknown", representation(ref = "externalptr"))

#' @export
#' @rdname RDCOMClient
setClass("COMIDispatch", representation("IUnknown"))

#' @export
#' @rdname RDCOMClient
setClass("COMDate", representation("numeric"))

#' @export
#' @rdname RDCOMClient
setClass("COMCurrency", representation("numeric"))

#' @export
#' @rdname RDCOMClient
setClass("COMDecimal", representation("numeric"))

#' @export
#' @rdname RDCOMClient
setClass("HResult", representation("numeric"))


#' @export
#' @rdname RDCOMClient
setClass("VARIANT", representation(ref= "externalptr", kind="integer"),
         prototype=list(kind=integer(1)))

#' @export
#' @rdname RDCOMClient
setClass("CurrencyVARIANT", representation("VARIANT"))

#' @export
#' @rdname RDCOMClient
setClass("DateVARIANT", representation("VARIANT"))




# setClass("CurrencyValue", representation("numeric"))


