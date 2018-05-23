

#' Converts Excel address to indexes and vice versa.
#' 
#' @param str.range character Excel range address
#' @param top integer top index of top-left cell
#' @param left integer left index of top-left cell
#' @param bottom integer bottom index of bottom-right cell
#' @param right integer right index of bottom-right cell
#'   
#' @return xl.index2address returns character address (e. g. A1:B150), 
#'   xl.address2index returns vector with four components: top, left, bottom,
#'   right.
#'   
#' @examples
#' 
#' xl.address2index("A1:D150")
#' xl.index2address(top=1, left=1)
#' 
#' \dontrun{
#' a1 %=xl% a1
#' a1 = iris
#' addr = xl.binding.address(a1)$address
#' xl.address2index(addr) 
#' 
#' }
#' 
#' @export
xl.index2address = function(top, left, bottom = NULL, right = NULL){
    res = paste0(column_address(left),top)
    if(!is.null(bottom) && !is.null(right)){
        res = paste0(res,":",column_address(right),bottom)
        
    }
    res
}

#' @export
#' @rdname xl.index2address
xl.address2index = function(str.range){
    str.range = tolower(str.range)    
    str.range = gsub("^(.*?)!","",str.range, perl = TRUE)
    str.range = gsub("\\$|\\s","",str.range, perl = TRUE)
    address_vec = unlist(strsplit(str.range, split=":"))
    if (length(address_vec)<2) address_vec[2] = address_vec[1] 
    res = c(top=NA, left = NA, bottom = NA, right = NA)
    res["top"] = as.numeric(gsub("[^\\d]","",address_vec[1],perl = TRUE))
    res["left"] = column_number(gsub("\\d","",address_vec[1],perl = TRUE))
    res["bottom"] = as.numeric(gsub("[^\\d]","",address_vec[2],perl = TRUE))
    res["right"] = column_number(gsub("\\d","",address_vec[2],perl = TRUE))
    res
}




column_address = function(col)
{
    if (col <= 26) { 
        return(LETTERS[col])
    }
    div = floor(col / 26)
    mod = col %% 26
    if (mod == 0) {
        mod = 26
        div = div - 1
    }
    return(paste0(column_address(div),LETTERS[mod]))
}

column_number = function(col_address)
{
    col_address = tolower(rev(unlist(strsplit(col_address, split=""))))
    digits = numeric(length(col_address))
    for (i in seq_along(digits))
    {
        digits[i] = match(col_address[i],letters)
    }
    
    mul = 1
    res = 0
    for (pos in seq_along(digits))
    {
        res = res + digits[pos] * mul
        mul = mul*26
    }
    res
}


