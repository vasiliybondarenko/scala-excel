package poi.excel.examples

import java.io._

import org.apache.poi.ss.usermodel.charts._
import org.apache.poi.ss.util._
import org.apache.poi.xssf.usermodel._

object ReportWriter extends App {
  // Данные для построения отчета
  val months = Array("октябрь 2012", "ноябрь 2012", "декабрь 2012",
    "январь 2013", "февраль 2013", "март 2013")
  val data = Map(
    (1, "RU-CENTER") -> Array(83318, 80521, 83048, 73638, 82014, 93982),
    (2, "REGRU") -> Array(35621, 37013, 36515, 41595, 45042, 49101),
    (3, "R01") -> Array(44155, 44356, 43199, 39629, 42754, 48528),
    (4, "REGTIME") -> Array(19999, 18587, 18630, 18627, 19886, 20496)
  )

  // Создаем новую книгу с одним листом
  val wb = new XSSFWorkbook()
  val sheet = wb.createSheet("Продление доменов")

  // Создаем шапку таблицы
  val headerRow = sheet.createRow(0)
  headerRow.createCell(0).setCellValue("Регистратор / Месяц")
  for(idx <- 1 to months.size) {
    headerRow.createCell(idx).setCellValue(months(idx - 1))
  }

  // Создаем тело таблицы
  for(key @ (rowNumber, rowName) <- data.keys) {
    val row = sheet.createRow(rowNumber)
    row.createCell(0).setCellValue(rowName)
    for(idx <- 1 to data(key).size) {
      row.createCell(idx).setCellValue(data(key)(idx - 1))
    }
  }

  // Создаем стиль ячеек в подвале таблицы
  val footerStyle = wb.createCellStyle
  val footerFont = wb.createFont
  footerFont.setBold(true)
  footerStyle.setFont(footerFont)

  // Создаем подвал таблицы
  val footerRow = sheet.createRow(data.size + 1)
  val firstCell = footerRow.createCell(0)
  firstCell.setCellValue("Всего")
  firstCell.setCellStyle(footerStyle)
  for(idx <- 1 to months.size) {
    val cell = footerRow.createCell(idx)
    val range = new CellRangeAddress(1, data.size, idx, idx)
    cell.setCellFormula(s"SUM(${range.formatAsString})")
    cell.setCellStyle(footerStyle)
  }

  // Наводим красоту :)
  for(columnIndex <- 0 to months.size) {
    sheet.autoSizeColumn(columnIndex, true)
  }

  // Рисуем диаграмму
  val chartLeftOffset = 0
  val chartTopOffset = data.size + 3
  val chartWidth = 4
  val chartHeight = 10
  val drawing = sheet.createDrawingPatriarch
  val anchor = drawing.createAnchor(
    0, 0, 0, 0,
    chartLeftOffset, chartTopOffset,
    chartLeftOffset + chartWidth, chartTopOffset + chartHeight
  )
  val chart = drawing.createChart(anchor)
  val legend = chart.getOrCreateLegend
  legend.setPosition(LegendPosition.TOP_RIGHT)

  val chartData = chart.getChartDataFactory.createScatterChartData
  val axisFactory = chart.getChartAxisFactory
  val bottomAxis = axisFactory.createValueAxis(AxisPosition.BOTTOM)
  val leftAxis = axisFactory.createValueAxis(AxisPosition.LEFT)
  leftAxis.setCrosses(AxisCrosses.AUTO_ZERO)

  val monthsRange = new CellRangeAddress(0, 0, 1, months.size)
  val xs = DataSources.fromNumericCellRange(sheet, monthsRange)
  for(idx <- 1 to data.size) {
    val rowRange = new CellRangeAddress(idx, idx, 1, months.size)
    val ys = DataSources.fromNumericCellRange(sheet, rowRange)
    chartData.addSerie(xs, ys)
  }
  chart.plot(chartData, bottomAxis, leftAxis)

  // Сохраняем книгу
  val resultFile = new FileOutputStream("/Users/shredinger/Downloads/report.xls")
  wb.write(resultFile)
  resultFile.close
}

