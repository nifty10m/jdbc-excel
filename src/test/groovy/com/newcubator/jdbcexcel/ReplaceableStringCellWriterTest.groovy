package com.newcubator.jdbcexcel

import com.newcubator.jdbcexcel.cellwriters.ReplaceableStringCellWriter
import spock.lang.Specification


class ReplaceableStringCellWriterTest extends Specification {
  def "test replaceAll"() {
    given:
      def writer = new ReplaceableStringCellWriter()
    when:
      def result = writer.replaceAll("http://{baseurl}/ab/de/{image}", ["baseurl": "google", "image": "cat.png"])
    then:
      result == "http://google/ab/de/cat.png"
  }

}
