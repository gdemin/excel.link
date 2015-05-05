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


.onLoad <-
function(lib, pkg) {
 library.dynam("excel.link", pkg, lib)
 .COMInit()
}


