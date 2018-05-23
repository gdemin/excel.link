#' Auxiliary function for export graphics to Microsoft Excel
#' 
#' @param type  file type. Ignored if argument 'filename' provided.
#' @param filename character. filename (or full path) of file with graphics.
#' @param picname character. Picture name in Excel.
#' @param ... arguments for internally used \code{\link{dev.copy}} function
#'   
#' @return Path to file with saved graphics with class attribute 
#'   'current.graphics'. If used with argument \code{type} than result has 
#'   attribute \code{temp.file = TRUE}.
#'   
#' @details If argument \code{type} provided this function will save graphics 
#'   from windows plotting device to temporary file and return path to this 
#'   file. Argument \code{filename} is intended to transfer plots to Excel from 
#'   file-based graphics devices (see Examples) or just insert into Excel file 
#'   with graphics. If argument \code{filename} is provided argument \code{type}
#'   will be ignored and returned value is path to file \code{filename} with 
#'   class attribute 'current.graphics'. So it could be used with expressions 
#'   such \code{xl[a1] = current.graphics(filename="plot.png")}. If 
#'   \code{picname} is provided then picture will be inserted in Excel with this
#'   name. If picture \code{picname} already exists in Excel it will be deleted.
#'   This argument is useful when we need to change old picture in Excel instead
#'   of adding new picture. \code{picname} will be automatically prepended by
#'   "_" to avoid conflicts with Excel range names.
#'   
#' @examples
#' 
#' \dontrun{
#' xl.workbook.add()
#' plot(sin)
#' xl[a1] = current.graphics()
#' plot(cos)
#' cos.plot = current.graphics()
#' xl.sheet.add()
#' xl[a1] = list("Cosine plotting",cos.plot,"End of cosine plotting")
#' 
#' # the same thing without graphic windows 
#' png("1.png")
#' plot(sin)
#' dev.off()
#' sin.plot = current.graphics(filename = "1.png")
#' png("2.png")
#' plot(cos)
#' dev.off()
#' cos.plot = current.graphics(filename = "2.png")
#' output = list("Cosine plotting",cos.plot,"Sine plotting",sin.plot)
#' xl.workbook.add()
#' xl[a1] = output
#' }
#' 
#' @export
current.graphics = function(type = c("png","jpeg","bmp","tiff"),filename = NULL,picname = NULL, ...){
    if (is.null(filename)){
        type = match.arg(type)
        res = paste(tempfile(),".",type,sep = "")
        switch(type,
               png = dev.copy(png,res,...),
               jpeg = dev.copy(jpeg,res,...),
               bmp = dev.copy(bmp,res,...),
               tiff = dev.copy(tiff,res,...)
        )
        dev.off()
        attr(res,"temp.file") = TRUE
    } else {
        res = normalizePath(filename,mustWork = TRUE)
    }
    if(!is.null(picname)){
        attr(res,"picname") = paste0("_",picname)
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



