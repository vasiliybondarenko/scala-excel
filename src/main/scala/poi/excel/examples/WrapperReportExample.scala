package poi.excel.examples

import info.folone.scala.poi._

import scalaz.syntax.monoid._

/**
 * Created with IntelliJ IDEA.
 * Author: shredinger
 * Date: 1/12/15
 * Time: 2:44 PM 
 * Project: scala-excel
 */
object WrapperReportExample extends App{

  val path = "/Users/shredinger/Downloads/wrapper_example.xls"

  val sheetOne = Workbook {
    Set(
      Sheet("name") {
        Set(
          Row(0) {
            Set(NumericCell(1, 13.0/4), NumericCell(1, 13.0/4))
          },
          Row(1) {
          Set(NumericCell(1, 13.0/5), FormulaCell(2, "ABS(A1)"))
        },
        Row(2) {
            Set(StringCell(1, "data"), StringCell(2, "data2"))
        })
    },
    Sheet("name2") {
        Set(Row(2) {
          Set(BooleanCell(1, true), NumericCell(2, 2.4))
        })
    }
    )
  }

  sheetOne.safeToFile(path).fold(ex ⇒ throw ex, identity).unsafePerformIO

  val sheetTwo = Workbook {
    Set(Sheet("name") {
      Set(Row(1) {
        Set(StringCell(1, "newdata"), StringCell(2, "data2"), StringCell(3, "data3"))
      },
        Row(2) {
          Set(StringCell(1, "data"), StringCell(2, "data2"))
        },
        Row(3) {
          Set(StringCell(1, "data"), StringCell(2, "data2"))
        })
    },
      Sheet("name") {
        Set(Row(2) {
          Set(StringCell(1, "data"), StringCell(2, "data2"))
        })
      })
  }

  val res = Workbook(path).fold(
    ex       => false,
    workbook => (workbook |+| sheetTwo) == (sheetOne |+| sheetTwo)
  )

  res.unsafePerformIO


}
