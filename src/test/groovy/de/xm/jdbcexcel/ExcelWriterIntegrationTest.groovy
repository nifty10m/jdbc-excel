package de.xm.jdbcexcel

import de.xm.jdbcexcel.tabs.ExcelTab
import org.springframework.beans.factory.annotation.Autowired
import org.springframework.boot.test.autoconfigure.data.jdbc.DataJdbcTest
import org.springframework.boot.test.autoconfigure.jdbc.AutoConfigureTestDatabase
import org.springframework.jdbc.core.JdbcTemplate
import spock.lang.Specification

@DataJdbcTest
@AutoConfigureTestDatabase(replace = AutoConfigureTestDatabase.Replace.NONE)
class ExcelWriterIntegrationTest extends Specification {

  @Autowired
  JdbcTemplate jdbcTemplate

  ExcelWriter excel

  def setup() {
    excel = new ExcelWriter(jdbcTemplate)
  }

  def "Check export of DATE and NUMERIC column type,"() {
    given:
      def file = new File("all-rows.xlsx")
      if (file.exists()) {
        file.delete()
      }
    when:
      def bytes = excel.createExcel(ExcelTab.of("Default", """
      SELECT payment_date, amount, sum(amount) OVER (ORDER BY payment_date)
        FROM (
          SELECT CAST(payment_date AS DATE) AS payment_date, SUM(amount) AS amount
            FROM payment
          GROUP BY CAST(payment_date AS DATE)
        ) p
      ORDER BY payment_date;
      """))
      file << bytes
    then:
      file.size()

    cleanup:
      file.delete()
  }

  def "Check export of boolean column type"() {
    given:
      def file = new File("boolean-type.xlsx")
      if (file.exists()) {
        file.delete()
      }
    when:
      def bytes = excel.createExcel(ExcelTab.of("Default", """
      SELECT * FROM staff;
      """))
      file << bytes
    then:
      file.size()

    cleanup:
      file.delete()
  }

  def "Check export of null values"() {
    given:
      def file = new File("null-type.xlsx")
      if (file.exists()) {
        file.delete()
      }
    when:
      def bytes = excel.createExcel(ExcelTab.of("Default", """
      SELECT * FROM address;
      """))
      file << bytes
    then:
      file.size()

    cleanup:
      file.delete()
  }

  def "Check export of big table"() {
    given:
      def file = new File("much-rows-very-data-wow.xlsx")
      if (file.exists()) {
        file.delete()
      }

    when:
      def thenMilis = System.currentTimeMillis()
      def bytes = excel.createExcel(ExcelTab.of("Default", """
      SELECT * FROM inventory i JOIN store s on i.store_id = s.store_id JOIN staff sf on sf.store_id = s.store_id;
      """))
      def nowMilis = System.currentTimeMillis()
      def elapsed =  nowMilis - thenMilis;

      println("Elapsed time: $elapsed ms")

      file << bytes
    then:
      file.size()

    cleanup:
      file.delete()
  }

  def "Check export of VARCHAR column type,"() {
    given:
      def file = new File("all-rows.xlsx")
      if (file.exists()) {
        file.delete()
      }
    when:
      def bytes = excel.createExcel(ExcelTab.of("Default", """
      SELECT first_name, last_name, count(*) films
        FROM actor AS a
        JOIN film_actor AS fa USING (actor_id)
      GROUP BY actor_id, first_name, last_name
      ORDER BY films DESC;
      """))
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
      def bytes = excel.createExcel(ExcelTab.of("Default", """
      SELECT first_name, last_name, email 
        FROM customer 
        JOIN address 
          ON (customer.address_id = address.address_id)
        JOIN city 
          ON (city.city_id = address.city_id)
        JOIN country
          ON (country.country_id = city.country_id)
       WHERE country.country= ?;
       """, ["Canada"]))
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
      def bytes = excel.createExcel(ExcelTab.of("Default", """
      SELECT actor_id, first_name, last_name 
        FROM actor
       WHERE last_name LIKE '%LI%'
      ORDER BY last_name, first_name;
      """))
      file << bytes
    then:
      file.size()

    cleanup:
      file.delete()
  }
}
