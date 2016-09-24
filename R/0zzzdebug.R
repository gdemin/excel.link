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



#This is called from the C code when a COMIDispatch object
# is registered or unregistered using the finalizers
#
.comRegistry <-
function(id, register)
{
 if(nargs() == 0)
  return(Table)

 if(id %in% names(Table)) {
   if(register)
     Table[id] <<- Table[id] + 1
   else
     Table[id] <<- Table[id] - 1
 } else if(register) {
   Table[id] <<- 1
 } else
   stop("Unregistering a value that has never been in the table.")
  
}
environment(.comRegistry) = new.env()
assign("Table", integer(0), environment(.comRegistry))

.gcAll <-
function()
{
  old= gctorture(TRUE)
  on.exit(gctorture(old))
  gc()
}
