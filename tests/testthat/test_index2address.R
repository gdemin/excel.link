context("index2address")


brute_force_test = function(){
    counter = 1    
    for(i in 0:26) {
        if (i == 0) sec_let = 0:26 else sec_let = 1:26
        if (i>0) l3 = LETTERS[i] else l3 = ""
        for(j in sec_let) {
            if (j>0) l2 = LETTERS[j] else l2 = ""
            for(k in 1:26){
                l1 = LETTERS[k]
                res = paste0(l3,l2,l1)
                if(res!=excel.link::column_address(counter)){
                    cat(counter,i,l3,j,l2,k,l1,"\n")
                    stop("Not good!")
                }
                if(counter!=excel.link::column_number(excel.link::column_address(counter))){
                    cat(counter,i,l3,j,l2,k,l1,"\n")
                    stop("2 Not good!")
                }
                if(counter!=excel.link::column_number(res)){
                    cat(counter,i,l3,j,l2,k,l1,"\n")
                    stop("3 Not good!")
                }
                counter = counter + 1
            } 
        }
    } 
    "Ok"
    
}

expect_identical(brute_force_test(),"Ok")

xl.workbook.add()

xl_iris %=xl% d3
xl_iris = iris

addr = xl.binding.address(xl_iris)$address

expect_identical(xl.address2index(addr),c(top=3,left=4,bottom=152,right=8))
expect_identical(xl.index2address(top=3,left=4,bottom=152,right=8),"D3:H152")

xl.workbook.close()
