package com.github.alien11689.spreadsheetbuilder

import org.modelcatalogue.spreadsheet.api.Color
import org.modelcatalogue.spreadsheet.api.Row
import org.modelcatalogue.spreadsheet.builder.poi.PoiSpreadsheetBuilder
import org.modelcatalogue.spreadsheet.query.api.SpreadsheetCriteria
import org.modelcatalogue.spreadsheet.query.api.SpreadsheetCriteriaResult
import org.modelcatalogue.spreadsheet.query.poi.PoiSpreadsheetCriteria
import spock.lang.Specification

import static org.modelcatalogue.spreadsheet.api.Keywords.auto
import static org.modelcatalogue.spreadsheet.api.Keywords.bold
import static org.modelcatalogue.spreadsheet.api.Keywords.bottom
import static org.modelcatalogue.spreadsheet.api.Keywords.center
import static org.modelcatalogue.spreadsheet.api.Keywords.thick

/**
 * http://metadataconsulting.github.io/spreadsheet-builder/
 * https://github.com/MetadataConsulting/spreadsheet-builder
 */
class DemoTest extends Specification {

    def 'should build simple xlsx'() {
        given:
            File file = new File('/tmp/01.xlsx')
        when:
            PoiSpreadsheetBuilder.create(file).build {
                sheet('first') {
                    row {
                        cell 'a'
                        cell 'b'
                        cell 'c'
                    }
                    row {
                        cell 5
                        cell 8
                        cell 13
                    }
                }
            }
        then:
            file.exists()
    }

    def 'should leave some rows and cells'() {
        given:
            File file = new File('/tmp/02.xlsx')
        when:
            PoiSpreadsheetBuilder.create(file).build {
                sheet('first') {
                    row {
                        cell 'a'
                        cell(3) {
                            value 'b'
                        }
                        cell(5) {
                            value 'c'
                        }
                    }
                    row 3, {
                        cell(5)
                        cell('C') {
                            value 8
                        }
                        cell('E') {
                            value 13
                        }
                    }
                }
            }
        then:
            file.exists()

    }

    def 'should freeze header'() {
        given:
            File file = new File('/tmp/03.xlsx')
        when:
            PoiSpreadsheetBuilder.create(file).build {
                sheet('first') {

                    freeze 0, 1

                    row {
                        cell('A') {
                            value 'a'
                        }
                        cell('B') {
                            value 'b'
                        }
                        cell('C') {
                            value 'c'
                        }
                    }
                    row {
                        cell('A') {
                            value 5
                        }
                        cell('B') {
                            value 8
                        }
                        cell('C') {
                            value 13
                        }
                    }
                }
            }
        then:
            file.exists()
    }

    def 'should group'() {
        given:
            File file = new File('/tmp/04.xlsx')
        when:
            PoiSpreadsheetBuilder.create(file).build {
                sheet('first') {

                    freeze 0, 1

                    row {
                        cell('A') {
                            value 'a'
                        }
                        group {
                            cell('B') {
                                value 'b'
                            }
                            cell('C') {
                                value 'c'
                            }
                        }
                    }
                    group {
                        row {
                            cell('A') {
                                value 5
                            }
                            cell('B') {
                                value 8
                            }
                            cell('C') {
                                value 13
                            }
                        }
                        row {
                            cell('A') {
                                value 8
                            }
                            cell('B') {
                                value 13
                            }
                            cell('C') {
                                value 21
                            }
                        }
                    }

                }
            }
        then:
            file.exists()
    }

    def 'should use formula'() {
        given:
            File file = new File('/tmp/05.xlsx')
        when:
            PoiSpreadsheetBuilder.create(file).build {
                sheet('first') {

                    freeze 0, 1

                    row {
                        cell('A') {
                            value 'Fib_part1'
                        }
                        group {
                            cell('B') {
                                value 'Fib_part2'
                            }
                            cell('C') {
                                value 'Fib'
                            }
                        }
                    }
                    group {
                        row {
                            cell('A') {
                                value 1
                            }
                            cell('B') {
                                value 1
                            }
                            cell('C') {
                                formula "A2 + B2"
                            }
                        }
                        (2..100).each { i ->
                            row {
                                cell('A') {
                                    formula "B${i}"
                                }
                                cell('B') {
                                    formula "C${i}"
                                }
                                cell('C') {
                                    formula "A${i + 1} + B${i + 1}"
                                }
                            }
                        }
                    }

                }
            }
        then:
            file.exists()
    }

    def 'should use formula with names'() {
        given:
            File file = new File('/tmp/06.xlsx')
        when:
            PoiSpreadsheetBuilder.create(file).build {
                sheet('first') {

                    freeze 0, 1

                    row {
                        cell('A') {
                            value 'Fib_part1'
                        }
                        group {
                            cell('B') {
                                value 'Fib_part2'
                            }
                            cell('C') {
                                value 'Fib'
                            }
                        }
                    }
                    group {
                        row {
                            cell('A') {
                                value 1
                                name 'fib_1_1'
                            }
                            cell('B') {
                                value 1
                                name 'fib_2_1'
                            }
                            cell('C') {
                                formula "A2 + B2"
                                name 'fib_sum_1'
                            }
                        }
                        (2..100).each { i ->
                            row {
                                cell('A') {
                                    formula "#{fib_2_${i - 1}}"
                                    name "fib_1_${i}"
                                }
                                cell('B') {
                                    formula "#{fib_sum_${i - 1}}"
                                    name "fib_2_${i}"
                                }
                                cell('C') {
                                    formula "sum(#{fib_1_${i}} + #{fib_1_${i}})"
                                    name "fib_sum_${i}"
                                }
                            }
                        }
                    }

                }
            }
        then:
            file.exists()
    }

    def 'should add span'() {
        given:
            File file = new File('/tmp/07.xlsx')
        when:
            PoiSpreadsheetBuilder.create(file).build {
                sheet('first') {

                    freeze 0, 1

                    row {
                        cell('A') {
                            value "fib_1 + fib_2 = fib_sum"
                            colspan 3
                        }
                    }
                    group {
                        row {
                            cell('A') {
                                value 1
                                name 'fib_1_1'
                            }
                            cell('B') {
                                value 1
                                name 'fib_2_1'
                            }
                            cell('C') {
                                formula "A2 + B2"
                                name 'fib_sum_1'
                            }
                        }
                        (2..100).each { i ->
                            row {
                                cell('A') {
                                    formula "#{fib_2_${i - 1}}"
                                    name "fib_1_${i}"
                                }
                                cell('B') {
                                    formula "#{fib_sum_${i - 1}}"
                                    name "fib_2_${i}"
                                }
                                cell('C') {
                                    formula "sum(#{fib_1_${i}} + #{fib_1_${i}})"
                                    name "fib_sum_${i}"
                                }
                            }
                        }
                    }

                }
            }
        then:
            file.exists()
    }

    def 'should build with style'() {
        given:
            File file = new File('/tmp/08.xlsx')
        when:
            PoiSpreadsheetBuilder.create(file).build {
                style('headers') {
                    border(bottom) {
                        style thick
                        color Color.black
                    }
                    font {
                        size 15
                        style bold
                    }
                    background Color.gray
                    align center, center
                }

                style('a') {
                    background Color.beige
                }

                style('b') {
                    background Color.aquamarine
                }

                sheet('first') {

                    freeze 0, 1

                    row {
                        style 'headers'
                        cell('A') {
                            height 20
                            value "fib_1 + fib_2 = fib_sum"
                            colspan 3
                        }
                    }
                    group {
                        row {
                            style 'a'
                            cell('A') {
                                width 20
                                value 1
                                name 'fib_1_1'
                            }
                            cell('B') {
                                value 1
                                width 20
                                name 'fib_2_1'
                            }
                            cell('C') {
                                width 20
                                formula "A2 + B2"
                                name 'fib_sum_1'
                            }
                        }
                        (2..100).each { i ->
                            row {
                                style(i % 2 == 0 ? 'b' : 'a')
                                cell('A') {
                                    formula "#{fib_2_${i - 1}}"
                                    name "fib_1_${i}"
                                }
                                cell('B') {
                                    formula "#{fib_sum_${i - 1}}"
                                    name "fib_2_${i}"
                                }
                                cell('C') {
                                    formula "sum(#{fib_1_${i}} + #{fib_1_${i}})"
                                    name "fib_sum_${i}"
                                }
                            }
                        }
                    }

                }
            }
        then:
            file.exists()
    }

    def 'should build with filters'() {
        given:
            File file = new File('/tmp/09.xlsx')
        when:
            PoiSpreadsheetBuilder.create(file).build {
                sheet('first') {

                    freeze 0, 1

                    row {
                        cell('A') {
                            filter auto
                            value 'ABC'
                        }
                        cell('B') {
                            filter auto
                            value 'XYZ'
                        }
                    }
                    Random random = new Random()
                    (1..100).each { i ->
                        row {
                            cell(['a', 'b', 'c'][random.nextInt(3)])
                            cell(['x', 'y', 'z'][random.nextInt(3)])
                        }
                    }
                }

            }
        then:
            file.exists()
    }

    def 'should query cells'() {
        given:
            File file = new File('/tmp/09.xlsx')
            SpreadsheetCriteria query = PoiSpreadsheetCriteria.FACTORY.forFile(file)
        when:
            SpreadsheetCriteriaResult result = query.query {
                sheet {
                    row {
                        cell('B') {
                            value 'y'
                        }
                    }
                }
            }
        then:
            result.cells.size() in 0..100
            result.cells.every { it.value == 'y' }
    }

    def 'should check fib'() {
        given:
            File file = new File('/tmp/10.xlsx')
        when:
            PoiSpreadsheetBuilder.create(file).build {
                sheet('first') {

                    long a = 1
                    long b = 1
                    long c = 2

                    100.times {
                        row {
                            cell a
                            cell b
                            cell c
                        }
                        a = b
                        b = c
                        c = a + b
                    }
                }
            }
        then:
            file.exists()
        when:
            SpreadsheetCriteriaResult result = PoiSpreadsheetCriteria.FACTORY.forFile(file).query {
                sheet {
                    row(1, 20)
                    row {
                        cell(1, 2)
                    }
                }
            }
        then:
            result.cells.size() == 20 * 2
            result.rows.size() == 20
            Collection<Row> fromSecondRow = result.rows.findAll { it.number > 1 }
            fromSecondRow.size() == 19
            fromSecondRow.every {
                it.cells.findAll { it.columnAsString in ['A', 'B'] }.every {
                    it.value == it.aboveRight.value
                }
            }
    }

    def 'should build fib with template'() {
        given:
            File file = new File('/tmp/11.xlsx')
            InputStream template = this.class.getResourceAsStream('/template.xlsx')
        when:
            PoiSpreadsheetBuilder.create(file, template).build {
                sheet('first') {

                    long a = 1
                    long b = 1
                    long c = 2

                    (2..101).each { i ->
                        row(i) {
                            cell a
                            cell b
                            cell c
                        }
                        a = b
                        b = c
                        c = a + b
                    }
                }
            }
        then:
            file.exists()
    }

}
