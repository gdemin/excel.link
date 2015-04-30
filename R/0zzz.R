.onLoad <-
function(lib, pkg) {
 library.dynam("excel.link", pkg, lib)
 .COMInit()
}


