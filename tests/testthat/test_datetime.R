

xl.workbook.open("datetime.xlsx")
xl.sheet.activate("datetime")

context("nan")

expect_identical(all(is.nan(xl[d3:d5])),TRUE)
expect_identical(all(is.nan(xl[e3:e4])),TRUE)
expect_equal_to_reference(xl[b3:e7],"rds/datetime1.rds", update = TRUE)
expect_equal_to_reference(xlc[b2:e7],"rds/datetime2.rds", update = TRUE)
expect_equal_to_reference(xlr[a3:e7],"rds/datetime3.rds", update = TRUE)
expect_equal_to_reference(xlrc[a2:e7],"rds/datetime4.rds", update = TRUE)


context("datetime read")
expect_equal_to_reference(xl[b3], "rds/datetime5.rds", update = TRUE)
expect_equal_to_reference(xl[b3:b4], "rds/datetime6.rds", update = TRUE)
expect_equal_to_reference(xl[b3:b5], "rds/datetime7.rds", update = TRUE)
expect_equal_to_reference(xl[b3:b6], "rds/datetime8.rds", update = TRUE)
expect_equal_to_reference(xl[b3:c3], "rds/datetime9.rds", update = TRUE)
expect_equal_to_reference(xl[b3:c4], "rds/datetime10.rds", update = TRUE)
expect_equal_to_reference(xl[b3:c5], "rds/datetime11.rds", update = TRUE)
expect_equal_to_reference(xl[b3:c6], "rds/datetime12.rds", update = TRUE)
expect_equal_to_reference(xl[b3:c7], "rds/datetime13.rds", update = TRUE)


xl.sheet.activate("rcdatetime")
expect_equal_to_reference(xlc[b2:e7],"rds/datetime14.rds", update = TRUE)
expect_equal_to_reference(xlr[a3:e7],"rds/datetime15.rds", update = TRUE)
expect_equal_to_reference(xlrc[a2:e7],"rds/datetime16.rds", update = TRUE)


expect_equal_to_reference(xlc[b11:e16],"rds/datetime17.rds", update = TRUE)
expect_equal_to_reference(xlr[a12:e16],"rds/datetime18.rds", update = TRUE)
expect_equal_to_reference(xlrc[a11:e16],"rds/datetime19.rds", update = TRUE)


xl.sheet.activate("datetime")
xl.workbook.close()

context("datetime write")

xl.workbook.add()

xl[a1, na = "NA"] = strptime(c("2006-01-08 10:07:52", "2006-08-07 19:33:02", NA),
         "%Y-%m-%d %H:%M:%S", tz = "EST5EDT")

expect_identical(xl[a1:a3, na = "NA"], c("2006-01-08 10:07:52", "2006-08-07 19:33:02", NA))

xl[a1, na = "NA"] = as.POSIXct(strptime(c("2006-01-08 10:07:52", "2006-08-07 19:33:02", NA),
                             "%Y-%m-%d %H:%M:%S", tz = "EST5EDT"))

expect_identical(xl[a1:a3, na = "NA"], c("2006-01-08 10:07:52", "2006-08-07 19:33:02", NA))

xl.workbook.close()

