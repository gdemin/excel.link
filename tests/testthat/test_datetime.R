

xl.workbook.open("datetime.xlsx")
xl.sheet.activate("datetime")

context("nan")

expect_identical(all(is.nan(xl[d3:d5])),TRUE)
expect_identical(all(is.nan(xl[e3:e4])),TRUE)
expect_equal_to_reference(xl[b3:e7],"rds/datetime1.rds")
expect_equal_to_reference(xlc[b2:e7],"rds/datetime2.rds")
expect_equal_to_reference(xlr[a3:e7],"rds/datetime3.rds")
expect_equal_to_reference(xlrc[a2:e7],"rds/datetime4.rds")


context("datetime read")
expect_equal_to_reference(xl[b3], "rds/datetime5.rds")
expect_equal_to_reference(xl[b3:b4], "rds/datetime6.rds")
expect_equal_to_reference(xl[b3:b5], "rds/datetime7.rds")
expect_equal_to_reference(xl[b3:b6], "rds/datetime8.rds")
expect_equal_to_reference(xl[b3:c3], "rds/datetime9.rds")
expect_equal_to_reference(xl[b3:c4], "rds/datetime10.rds")
expect_equal_to_reference(xl[b3:c5], "rds/datetime11.rds")
expect_equal_to_reference(xl[b3:c6], "rds/datetime12.rds")
expect_equal_to_reference(xl[b3:c7], "rds/datetime13.rds")


xl.sheet.activate("rcdatetime")
expect_equal_to_reference(xlc[b2:e7],"rds/datetime14.rds")
expect_equal_to_reference(xlr[a3:e7],"rds/datetime15.rds")
expect_equal_to_reference(xlrc[a2:e7],"rds/datetime16.rds")


expect_equal_to_reference(xlc[b11:e16],"rds/datetime17.rds")
expect_equal_to_reference(xlr[a12:e16],"rds/datetime18.rds")
expect_equal_to_reference(xlrc[a11:e16],"rds/datetime19.rds")


xl.sheet.activate("datetime")
xl.workbook.close()