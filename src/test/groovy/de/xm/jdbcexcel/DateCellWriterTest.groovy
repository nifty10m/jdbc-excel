package de.xm.jdbcexcel

import org.apache.poi.ss.usermodel.DateUtil
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import spock.lang.Specification

import java.time.LocalDateTime
import java.time.ZoneOffset

class DateCellWriterTest extends Specification {

  def "should format a cell with a cell date format"() {
    given:
      def testWorkbook = new XSSFWorkbook()
      def testSheet = testWorkbook.createSheet()
      def testCell = testSheet.createRow(0).createCell(0)

      def testLocalDateTime = LocalDateTime.of(2020, 1, 1, 12, 0)
      def testDate = Date.from(testLocalDateTime.toInstant(ZoneOffset.UTC))

    when:
      new DateCellWriter().doWriteCell(testWorkbook, testCell, testDate)
    then:
      DateUtil.isCellDateFormatted(testCell)
  }
}
