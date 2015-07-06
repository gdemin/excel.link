context("xl.bind.range")

xl.workbook.add()

a1 %=xl% a1
cra1 %=cr% a1
xl[a1] = 42
expect_identical(a1, cra1)
expect_identical(gsub("^(.+?)!","",xl.binding.address(a1),perl=TRUE), "$A$1")
expect_identical(gsub("^(.+?)!","",xl.binding.address(cra1),perl=TRUE), "$A$1")

xl[b1] = 41
xl[b2] = 41
xl[a2] = 41

expect_identical(a1, 42)
expect_equal_to_reference(cra1, "rds/cra1.rds")
expect_identical(gsub("^(.+?)!","",xl.binding.address(a1),perl=TRUE), "$A$1")
expect_identical(gsub("^(.+?)!","",xl.binding.address(cra1),perl=TRUE), "$A$1:$B$2")

a1 = iris
expect_identical(a1, cra1)
xl.sheet.activate(2)
expect_equal_to_reference(cra1, "rds/cra2.rds")
expect_identical(gsub("^(.+?)!","",xl.binding.address(a1),perl=TRUE), "$A$1:$E$150")
expect_identical(gsub("^(.+?)!","",xl.binding.address(cra1),perl=TRUE), "$A$1:$E$150")

xl.workbook.close()

xl.workbook.add()

a1 %=xlrc% a1
cra1 %=crrc% a1
xl[a1] = 42
expect_identical(a1, cra1)
expect_identical(gsub("^(.+?)!","",xl.binding.address(a1),perl=TRUE), "$A$1")
expect_identical(gsub("^(.+?)!","",xl.binding.address(cra1),perl=TRUE), "$A$1")

xl[b1] = 41
xl[b2] = 41
xl[a2] = 41

expect_identical(a1, 42)
expect_equal_to_reference(cra1, "rds/cra3.rds")
expect_identical(gsub("^(.+?)!","",xl.binding.address(a1),perl=TRUE), "$A$1")
expect_identical(gsub("^(.+?)!","",xl.binding.address(cra1),perl=TRUE), "$A$1:$B$2")

a1 = iris
expect_identical(a1, cra1)
expect_equal_to_reference(cra1, "rds/cra4.rds")
expect_identical(gsub("^(.+?)!","",xl.binding.address(a1),perl=TRUE), "$A$1:$F$151")
expect_identical(gsub("^(.+?)!","",xl.binding.address(cra1),perl=TRUE), "$A$1:$F$151")

a1 = 99
expect_identical(a1, cra1)
expect_identical(gsub("^(.+?)!","",xl.binding.address(a1),perl=TRUE), "$A$1")
expect_identical(gsub("^(.+?)!","",xl.binding.address(cra1),perl=TRUE), "$A$1")

a1 = iris
cra1$new_col = "test"
expect_equal_to_reference(cra1, "rds/cra5.rds")
expect_equal_to_reference(a1, "rds/cra6.rds")

cra1 = NA
suppressWarnings(expect_equal_to_reference(a1, "rds/cra7.rds"))
expect_identical(cra1, NA)

xl.workbook.close()



