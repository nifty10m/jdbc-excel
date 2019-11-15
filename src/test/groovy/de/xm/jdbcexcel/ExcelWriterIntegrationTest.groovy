package de.xm.jdbcexcel

import org.springframework.beans.factory.annotation.Autowired
import org.springframework.boot.test.autoconfigure.data.jdbc.DataJdbcTest
import org.springframework.jdbc.core.JdbcTemplate
import spock.lang.Specification

@DataJdbcTest
class ExcelWriterIntegrationTest extends Specification {

  @Autowired
  JdbcTemplate jdbcTemplate

  ExcelWriter excel
  def setup() {
      excel = new ExcelWriter(jdbcTemplate)
  }

  def "Check reading of all data"() {
    given:
      def file = new File("all-rows.xlsx")
      if (file.exists()) {
        file.delete()
      }
    when:
      def bytes = excel.createExcel(ExcelTab.of("Default", "select * from book"))
      file << bytes
    then:
      file.size()
    cleanup:
      file.delete()
  }

  def "Check reading of a single line with prepared statement"() {
    given:
      def file = new File("one-line.xlsx")
      if (file.exists()) {
        file.delete()
      }
      def excel = new ExcelWriter(jdbcTemplate)
    when:
      def bytes = excel.createExcel(ExcelTab.of("Default", "select * from book where id=?",[42]))
      file << bytes
    then:
      file.size()
    cleanup:
      file.delete()
  }

  def "Check reading of a single line with inline statement"() {
    given:
      def file = new File("one-inline.xlsx")
      if (file.exists()) {
        file.delete()
      }
      def excel = new ExcelWriter(jdbcTemplate)
    when:
      def bytes = excel.createExcel(ExcelTab.of("Default", "select * from book where id=42"))
      file << bytes
    then:
      file.size()
    cleanup:
      file.delete()
  }
}
