# Package: RDCOMClient
# Version: 0.93-0.2
# Title: R-DCOM Client
# Author: Duncan Temple Lang <duncan@wald.ucdavis.edu>
#     Maintainer: Duncan Temple Lang <duncan@wald.ucdavis.edu>
#     Description: Provides dynamic client-side access to (D)COM applications from within R.
# License: GPL-2
# Collate: classes.R COMLists.S COMError.R com.R debug.S zzz.R runTime.S
# URL: http://www.omegahat.net/RDCOMClient, http://www.omegahat.net
# http://www.omegahat.net/bugs



#' @export
#' @rdname RDCOMClient
COMStop =
function(msg, status, class = "COMError")
{
  e = simpleError(msg)
  e$status = status
  class(e) = c(class, class(e))
  stop(e)
}
