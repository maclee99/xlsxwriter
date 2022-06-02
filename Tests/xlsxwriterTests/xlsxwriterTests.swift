import XCTest
import Cxlsxwriter
import Logging
@testable import xlsxwriter

final class xlsxwriterTests: XCTestCase {
    let logger = Logger(label: "test")

    func testExample() {
        // Create a new workbook
        let wb = Workbook(name: "demo.xlsx")
        defer { wb.close() }

        wb.properties(title: "title", subject: "subject", author: "Mac Lee")
        
        // Add a format.
        let f = wb
            .addFormat()
            .bold()
            .border(style: .thin)
            .align(horizontal: .center)
            .align(vertical: .center)
            .background(color: .fillGold)

        let bold = wb
            .addFormat()
            .bold()

        let dateFormat =  wb.addFormat()
            .set(num_format: "yyyy/MM/dd")

        
        // Add a format.
        let f2 = wb.addFormat().center()

        let f3 = wb.addFormat()
            .background(color: .fillGreen)   //.font(color: .white)
            .align(horizontal: .left)
        
        // Add a worksheet.
        let ws = wb.addWorksheet()
            .setDefault(row_height: 25)
            .tab(color: .blue) 
            .gridline(screen: true) 
            .row(0, height: 30)
            .column("A:A", width: 10, format: bold) 
            .column("D:D", width: 12, format: bold) 
            .column([1, 2], width: 15) 
            .showComments() 

        ws.merge("Merged Range", firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 2, format: f3)

        ws.write("Number", "A2", format: f)
            .write("Batch 1", "B2", format: f)
            .write(.string("Batch 2"), "C2", format: f)
            .write(.comment("comment"), Cell(stringLiteral: "C2")) 
            .write(.datetime(Date()), "D2", format: dateFormat)
            .write(.boolean(true), "E2")
            // .column(Cols(stringLiteral: "A:A"), width: 10, format: bold)
            // .column(Cols(arrayLiteral: 1, 2), width: 15)
       
        // Create random data
        let data = (1...100).map {
            [Double($0), Double.random(in: 10...100), Double.random(in: 20...50)]
        }
        
        // Write data to add to plot on the chart.
        data.enumerated().forEach {
            ws.write($0.element, row: $0.offset + 2, format: f2)
        }

        ws.freeze(row: 2, col: 1)
           .activate()
        ws.hideColumns(8) 

        
        // Create a new Chart
        let chart = wb
            .addChart(type: .line)
            .set(x_axis: "Test number")
            .set(y_axis: "Sample length (mm)")
            .set(style: 4)
        
        chart // In simplest case we add one or more data series.
            .addSeries()
            .values(sheet: ws, range: "$B$2:$B$101")
            .name(sheet: ws, cell: "B1")
            .trendline(type: .log)
            .trendline_equation()
        
        chart
            .addSeries(values: "=Sheet1!$C$2:$C$101", name: "=Sheet1!$C$1")
            .set(smooth: true)

        wb.addChartsheet(name: "Second")
            .paper(type: .A3)
            .zoom(scale: 150)
            // .activate()
            .set(chart: chart) // Insert the chart into the chartsheet.

        wb.validate(sheet_name: "Sheet Name")


    }

    func testTable() {
        let wb = Workbook(name: "tables.xlsx")
        defer { wb.close() }

        // let ws6 = wb.addWorksheet()
        // let ws7 = wb.addWorksheet()
        // let ws8 = wb.addWorksheet()
        // let ws9 = wb.addWorksheet()
        // let ws10 = wb.addWorksheet()
        // let ws11 = wb.addWorksheet()
        // let ws12 = wb.addWorksheet()
        // let ws13 = wb.addWorksheet()


        // Test 1. Default table with no data
        // set the columns widths for clarity.
        let ws1 = wb.addWorksheet()
        ws1.column("B:G", width: 12)
        // write the worksheet caption to explain the test.
        ws1.write("Default table with no data.", "B1")
        // Add a table to the worksheet
        ws1.table(range: "B3:F7")

        // Test 2. Default table with data
        // set the columns widths for clarity.
        let ws2 = wb.addWorksheet()
        ws2.column("B:G", width: 12)
        // write the worksheet caption to explain the test.
        ws2.write("Default table with data.", "B1")
        // Add a table to the worksheet
        ws2.table(range: "B3:F7")
        _writeSheetData(ws2)

        // Test 3. Table without default autofilter
        // set the columns widths for clarity.
        let ws3 = wb.addWorksheet()
        ws3.column("B:G", width: 12)
        // write the worksheet caption to explain the test.
        ws3.write("Table without default autofilter.", "B1")
        // Add a table to the worksheet
        ws3.table(range: "B3:F7", autoFilter: false)
        _writeSheetData(ws3)

        // Test 4. Table without default header row
        // set the columns widths for clarity.
        let ws4 = wb.addWorksheet()
        ws4.column("B:G", width: 12)
        // write the worksheet caption to explain the test.
        ws4.write("Table without default header row.", "B1")
        // Add a table to the worksheet
        ws4.table(range: "B4:F7", headerRow: false)
        _writeSheetData(ws4)

        //  Test 5. Default table with "First Column" and "Last Column" options
        // set the columns widths for clarity.
        let ws5 = wb.addWorksheet()
        ws5.column("B:G", width: 12)
        // write the worksheet caption to explain the test.
        ws5.write("Default table with \"First Column\" and \"Last Column\" options.", "B1")
        // Add a table to the worksheet
        ws5.table(range: "B3:F7", firstColumn: true, lastColumn: true)
        _writeSheetData(ws5)

        //  Test 6. Table with banded columns but without default banded rows
        // set the columns widths for clarity.
        let ws6 = wb.addWorksheet()
        ws6.column("B:G", width: 12)
        // write the worksheet caption to explain the test.
        ws6.write("Table with banded columns but without default banded rows.", "B1")
        // Add a table to the worksheet
        ws6.table(range: "B3:F7", bandedColumns: true, bandedRows: false)
        _writeSheetData(ws6)

        //  Test 7. Table with user defined column headers
        // set the columns widths for clarity.
        let ws7 = wb.addWorksheet()
        ws7.column("B:G", width: 12)
        // write the worksheet caption to explain the test.
        ws7.write("Table with user defined column headers.", "B1")
        // Add a table to the worksheet
        let headers: [String] = ["Product", "Quater1", "Quater2", "Quater3", "Quater4"]
        ws7.table(range: "B3:F7", header: headers)
        _writeSheetData(ws7)

        //  Test 8. Table with user defined column headers
        // set the columns widths for clarity.
        let ws8 = wb.addWorksheet()
        ws8.column("B:G", width: 12)
        // write the worksheet caption to explain the test.
        ws8.write("Table with user defined column headers.", "B1")
        // Add a table to the worksheet
        let headers8: [String] = ["Product", "Quater1", "Quater2", "Quater3", "Quater4", "Year"]
        let formula8: [String] = ["", "", "", "", "", "=SUM(Table8[@[Quater1]:[Quater4]])"]
        ws8.table(range: "B3:G7", name: "Table8", header: headers8, formula: formula8)
        _writeSheetData(ws8)

        //  Test 9. Table with totals row (but no caption or totals)
        // set the columns widths for clarity.
        let ws9 = wb.addWorksheet()
        ws9.column("B:G", width: 12)
        // write the worksheet caption to explain the test.
        ws9.write("Table with totals row (but no caption or totals).", "B1")
        // Add a table to the worksheet
        let headers9: [String] = ["Product", "Quater1", "Quater2", "Quater3", "Quater4", "Year"]
        let formula9: [String] = ["", "", "", "", "", "=SUM(Table9[@[Quater1]:[Quater4]])"]
        let totalRows9: [TotalFunction] = [.none]
        ws9.table(range: "B3:G8", header: headers9, totalRow: totalRows9, formula: formula9)
        _writeSheetData(ws9)

        //  Test 10. Table with totals row with user caption and functions
        // set the columns widths for clarity.
        let ws10 = wb.addWorksheet()
        ws10.column("B:G", width: 12)
        // write the worksheet caption to explain the test.
        ws10.write("Table with totals row with user caption and functions.", "B1")
        // Add a table to the worksheet
        let headers10: [String] = ["Product", "Quater1", "Quater2", "Quater3", "Quater4", "Year"]
        let formula10: [String] = ["Totals", "", "", "", "", "=SUM(Table10[@[Quater1]:[Quater4]])"]
        let totalRows10: [TotalFunction] = [.none, .sum, .sum, .sum, .sum, .sum]
        ws10.table(range: "B3:G8", header: headers10, totalRow: totalRows10, formula: formula10)
        _writeSheetData(ws10)

        //  Test 11. Table with alternative Excel style
        let ws11 = wb.addWorksheet()
        // set the columns widths for clarity.
        ws11.column("B:G", width: 12)
        // write the worksheet caption to explain the test.
        ws11.write("Table with alternative Excel style.", "B1")
        // Add a table to the worksheet
        let headers11: [String] = ["Product", "Quater1", "Quater2", "Quater3", "Quater4", "Year"]
        let formula11: [String] = ["Totals", "", "", "", "", "=SUM(Table10[@[Quater1]:[Quater4]])"]
        let totalRows11: [TotalFunction] = [.none, .sum, .sum, .sum, .sum, .sum]
        ws11.table(range: "B3:G8", header: headers11, totalRow: totalRows11, formula: formula11,
            styleType: UInt8(LXW_TABLE_STYLE_TYPE_LIGHT.rawValue), styleNumber: 11)
        _writeSheetData(ws11)

        //  Test 12. Table with Excel style removed
        let ws12 = wb.addWorksheet()
        // set the columns widths for clarity.
        ws12.column("B:G", width: 12)
        // write the worksheet caption to explain the test.
        ws12.write("Table with Excel style removed.", "B1")
        // Add a table to the worksheet
        let headers12: [String] = ["Product", "Quater1", "Quater2", "Quater3", "Quater4", "Year"]
        let formula12: [String] = ["Totals", "", "", "", "", "=SUM(Table10[@[Quater1]:[Quater4]])"]
        let totalRows12: [TotalFunction] = [.none, .sum, .sum, .sum, .sum, .sum]
        ws12.table(range: "B3:G8", header: headers12, totalRow: totalRows12, formula: formula12,
            styleType: UInt8(LXW_TABLE_STYLE_TYPE_LIGHT.rawValue), styleNumber: 0)
        _writeSheetData(ws12)

        //  Test 13. Table with column formats
        let ws13 = wb.addWorksheet()
        // set the columns widths for clarity.
        ws13.column("B:G", width: 12)
        // write the worksheet caption to explain the test.
        ws13.write("Table with column formats.", "B1")
        let currencyFormat =  wb.addFormat()
            .set(num_format: "$#,##0")
        // Add a table to the worksheet
        let headers13: [String] = ["Product", "Quater1", "Quater2", "Quater3", "Quater4", "Year"]
        let formats13: [Format?] = [nil, currencyFormat, currencyFormat, currencyFormat, currencyFormat, currencyFormat]
        let formula13: [String] = ["Totals", "", "", "", "", "=SUM(Table10[@[Quater1]:[Quater4]])"]
        let totalRows13: [TotalFunction] = [.none, .sum, .sum, .sum, .sum, .sum]
        ws13.table(range: "B3:G8", header: headers13, format: formats13, totalRow: totalRows13, formula: formula13)
        _writeSheetData(ws13, format: currencyFormat)


    }

    private func _writeSheetData(_ sheet: Worksheet, format: Format? = nil){
        sheet.write("Apples", "B4")
        sheet.write("Pears", "B5")
        sheet.write("Bananas", "B6")
        sheet.write("Oranges", "B7")

        sheet.write(.number(10000), "C4", format: format)
        sheet.write(.number(2000), "C5", format: format)
        sheet.write(.number(6000), "C6", format: format)
        sheet.write(.number(500), "C7", format: format)

        sheet.write(.number(5000), "D4", format: format)
        sheet.write(.number(3000), "D5", format: format)
        sheet.write(.number(6000), "D6", format: format)
        sheet.write(.number(300), "D7", format: format)

        sheet.write(.number(8000), "E4", format: format)
        sheet.write(.number(4000), "E5", format: format)
        sheet.write(.number(6500), "E6", format: format)
        sheet.write(.number(200), "E7", format: format)

        sheet.write(.number(6000), "F4", format: format)
        sheet.write(.number(5000), "F5", format: format)
        sheet.write(.number(6000), "F6", format: format)
        sheet.write(.number(700), "F7", format: format)


    }

    func testConditionalFormatSimple() {
        let wb = Workbook(name: "conditional_format_simple.xlsx")
        defer { wb.close() }

        // Test 1. Default table with no data
        // set the columns widths for clarity.
        let ws1 = wb.addWorksheet()
        // ws1.column("B:G", width: 12)
        // write the worksheet caption to explain the test.
        ws1.write(34.0, "B1")
        ws1.write(32.0, "B2")
        ws1.write(31.0, "B3")
        ws1.write(35.0, "B4")
        ws1.write(36.0, "B5")
        ws1.write(30.0, "B6")
        ws1.write(38.0, "B7")
        ws1.write(38.0, "B8")
        ws1.write(32.0, "B9")

        let redFontFormat =  wb.addFormat()
            .font(color: .red)


        ws1.conditionFormat(range: "B1:B9", type: .cell, criteria: .lessThan, value: 33, 
            format: redFontFormat )
    }

    func testConditionalFormat() {
        let wb = Workbook(name: "conditional_format.xlsx")
        defer { wb.close() }

        // Test 1. Conditional formatting based on simple cell based criteria
        let ws1 = wb.addWorksheet()
        ws1.write("Cells with values >= 50 are in light red. Values < 50 are in light green.", "A1")
        _writeFormatTestData(ws1)

        let lightRedFormat =  wb.addFormat()
            .font(color: 0x9C0006)
            .background(color: 0xFFC7CE)

        let lightGreenFormat =  wb.addFormat()
            .font(color: 0x006100)
            .background(color: 0xC6EFCE)

        ws1.conditionFormat(range: "B3:K12", type: .cell, criteria: .greaterThanOrEqualTo, value: 50, format: lightRedFormat )
        ws1.conditionFormat(range: "B3:K12", type: .cell, criteria: .lessThan, value: 50, format: lightGreenFormat )

        // Test 2. Conditional formatting based on max and min values.
        let ws2 = wb.addWorksheet()
        ws2.write("Values between 30 and 70 are in light red. Values outside that range are in light green.", "A1")
        _writeFormatTestData(ws2)

        ws2.conditionFormat(range: "B3:K12", type: .cell, criteria: .between, min: 30, max: 70, format: lightRedFormat )
        ws2.conditionFormat(range: "B3:K12", type: .cell, criteria: .notBetween, min: 30, max: 70, format: lightGreenFormat )

        // Test 3. Conditional formatting with duplicate and unique values.
        let ws3 = wb.addWorksheet()
        ws3.write("Duplicate values are in light red. Unique values are in light green.", "A1")
        _writeFormatTestData(ws3)

        ws3.conditionFormat(range: "B3:K12", type: .duplicate, format: lightRedFormat )
        ws3.conditionFormat(range: "B3:K12", type: .unique, format: lightGreenFormat )

        // Test 4. Conditional formatting with above and below average values.
        let ws4 = wb.addWorksheet()
        ws4.write("Above average values are in light red. Below average values are in light green.", "A1")
        _writeFormatTestData(ws4)

        ws4.conditionFormat(range: "B3:K12", type: .average, criteria: .averageAbove, min: 30, max: 70, format: lightRedFormat )
        ws4.conditionFormat(range: "B3:K12", type: .average, criteria: .averageBelow, min: 30, max: 70, format: lightGreenFormat )

        // Test 5. Conditional formatting with top and bottom values.
        let ws5 = wb.addWorksheet()
        ws5.write("Top 10 values are in light red. Bottom 10 values are in light green.", "A1")
        _writeFormatTestData(ws5)
        ws5.conditionFormat(range: "B3:K12", type: .top, value: 10, format: lightRedFormat )
        ws5.conditionFormat(range: "B3:K12", type: .bottom, value: 10, format: lightGreenFormat )

        // Test 6. Conditional formatting with multiple ranges.
        let ws6 = wb.addWorksheet()
        ws6.write("Cells with values >= 50 are in light red. Values < 50 are in light green. Non-contiguous ranges.", "A1")
        _writeFormatTestData(ws6)
        ws6.conditionFormat(range: "B3:K12", type: .cell, criteria: .greaterThanOrEqualTo, value: 50, multiRange: "B3:K6 B9:K12", format: lightRedFormat )
        ws6.conditionFormat(range: "B3:K12", type: .cell, criteria: .lessThan, value: 50, multiRange: "B3:K6 B9:K12", format: lightGreenFormat )

        // Test 7. Conditional formatting with 2 color scale.
        let ws7 = wb.addWorksheet()
        ws7.write("Examples of color scales with default and user colors.", "A1")
        // _writeFormatTestData(ws6)
        for i in 1...12 {
            ws7.write([Double(i)], row: i+1, col: 1)
            ws7.write([Double(i)], row: i+1, col: 3)
            ws7.write([Double(i)], row: i+1, col: 6)
            ws7.write([Double(i)], row: i+1, col: 8)
        }
        ws7.write("2 Color Scale", "B2")
        ws7.write("2 Color Scale + user colors", "D2")
        ws7.write("3 Color Scale", "G2")
        ws7.write("3 Color Scale + user colors", "I2")
        ws7.conditionFormat(range: "B3:B14", type: .twoColorScale)
        ws7.conditionFormat(range: "D3:D14", type: .twoColorScale, minColor: 0xFF0000, maxColor: 0x00FF00 )
        ws7.conditionFormat(range: "G3:G14", type: .threeColorScale)
        ws7.conditionFormat(range: "I3:I14", type: .threeColorScale, minColor: 0xC5D9F1, midColor: 0x8DB4E3, maxColor: 0x538ED5)

        // Test 8. Conditional formatting with data bars.
        let ws8 = wb.addWorksheet()
        ws8.write("Examples of data bars.", "A1")
        // _writeFormatTestData(ws6)
        // let data12: [Double] = (1...12).map{Double($0)}
        // ws8.write(data12, row: 1, col: 1)
        // ws8.write(data12, row: 1, col: 3)
        // ws8.write(data12, row: 1, col: 5)
        // ws8.write(data12, row: 1, col: 7)
        // ws8.write(data12, row: 1, col: 9)
        for i in 1...12 {
            ws8.write([Double(i)], row: i+1, col: 1)
            ws8.write([Double(i)], row: i+1, col: 3)
            ws8.write([Double(i)], row: i+1, col: 5)
            ws8.write([Double(i)], row: i+1, col: 7)
            ws8.write([Double(i)], row: i+1, col: 9)
        }
        let data: [Double] = [-1, -2, -3, -2, -1, 0, 1, 2, 3, 2, 1, 0]
        for i in 1...12 {
            ws8.write([data[i-1]], row: i+1, col: 11)
            ws8.write([data[i-1]], row: i+1, col: 13)
        }

        ws8.write("Default data bars", "B2")
        ws8.write("Bars only", "D2")
        ws8.write("With user color", "F2")
        ws8.write("Solid bars", "H2")
        ws8.write("Right to left", "J2")
        ws8.write("Excel 2010 style", "L2")
        ws8.write("Negative same as positive", "N2")

        ws8.conditionFormat(range: "B3:B14", type: .dataBar)
        ws8.conditionFormat(range: "D3:D14", type: .dataBar, barOnly: true )
        ws8.conditionFormat(range: "F3:F14", type: .dataBar, barColor: 0x63C384 )
        ws8.conditionFormat(range: "H3:H14", type: .dataBar, barSolid: true )
        ws8.conditionFormat(range: "J3:J14", type: .dataBar, barDirection: .rightToLeft )
        ws8.conditionFormat(range: "L3:L14", type: .dataBar, bar2010: true )
        ws8.conditionFormat(range: "N3:N14", type: .dataBar, negativeColorSame: true, negativeBorderColorSame: true )

        // Test 9. Conditional formatting with icon sets.
        let ws9 = wb.addWorksheet()
        ws9.write("Examples of conditional formats with icon sets.", "A1")
        // _writeFormatTestData(ws6)
        let data9: [Double] = (1...3).map{Double($0)}
        let data91: [Double] = (1...4).map{Double($0)}
        let data92: [Double] = (1...5).map{Double($0)}
        // ws8.write(data12, row: 1, col: 1)
        // ws8.write(data12, row: 1, col: 3)
        // ws8.write(data12, row: 1, col: 5)
        // ws8.write(data12, row: 1, col: 7)
        // ws8.write(data12, row: 1, col: 9)
        // for i in 1...3 {
            ws9.write(data9, row: 2, col: 1)
            ws9.write(data9, row: 3, col: 1)
            ws9.write(data9, row: 4, col: 1)
            ws9.write(data9, row: 5, col: 1)
            ws9.write(data91, row: 6, col: 1)
            ws9.write(data91, row: 6, col: 1)
            ws9.write(data92, row: 7, col: 1)
            ws9.write(data92, row: 8, col: 1)
        // }

        ws9.conditionFormat(range: "B3:D3", type: .iconSets, iconStyle: .trafficLightsUnrimmed3)
        ws9.conditionFormat(range: "B4:D4", type: .iconSets, iconStyle: .trafficLightsUnrimmed3, reverseIcons: true)
        ws9.conditionFormat(range: "B5:D5", type: .iconSets, iconStyle: .trafficLightsUnrimmed3, iconOnly: true)
        ws9.conditionFormat(range: "B6:D6", type: .iconSets, iconStyle: .arrowsColored3)
        ws9.conditionFormat(range: "B7:E7", type: .iconSets, iconStyle: .arrowsColored4)
        ws9.conditionFormat(range: "B8:F8", type: .iconSets, iconStyle: .arrowsColored5)
        ws9.conditionFormat(range: "B9:F9", type: .iconSets, iconStyle: .ratings5)

    }

    private func _writeFormatTestData(_ ws: Worksheet) {
        let data: [[Double]] = [
            [34, 72,  38, 30, 75, 48, 75, 66, 84, 86],
            [6,  24,  1,  84, 54, 62, 60, 3, 26,  59],
            [28, 79,  97, 13, 85, 93, 93, 22, 5,  14],
            [27, 71,  40, 17, 18, 79, 90, 93, 29, 47],
            [88, 25,  33, 23, 67, 1,  59, 79, 47, 36],
            [24, 100, 20, 88, 29, 33, 38, 54, 54, 88],
            [6,  57,  88, 28, 10, 26, 37, 7,  41, 48],
            [52, 78,  1,  96, 26, 45, 47, 33, 96, 36],
            [60, 54,  81, 66, 81, 90, 80, 93, 12, 55],
            [70, 5,   46, 14, 71, 19, 66, 36, 41, 21],
        ]

        for row in 0...9 {
            ws.write(data[row], row: row + 2, col: 1)
        }

    }

    func testRichStringFormat() {
        let wb = Workbook(name: "rich_strings.xlsx")
        defer { wb.close() }

        let ws1 = wb.addWorksheet()

        /* Make the first column wider for clarity. */
        ws1.column("A:A", width: 30)

        let bold =  wb.addFormat()
            .bold()
        let italic =  wb.addFormat()
            .italic()
        let red =  wb.addFormat()
            .font(color: .red)
        let blue =  wb.addFormat()
            .font(color: .blue)
        let center =  wb.addFormat()
            .center()
        let superScript =  wb.addFormat()
            .fontScript(.superScript)
        
        // Example 1. Some bold and italic in the same string.
        ws1.richString("A1", string: ["This is ", "bold", " and this is ", "italic"],
            formats: [nil, bold, nil, italic])

        // Example 2. Some red and blud coloring in the same string.
        ws1.richString("A3", string: ["This is ", "red", " and this is ", "blue"],
            formats: [nil, red, nil, blue])

        // Example 3. A rich string plus cell formatting.
        ws1.richString("A5", string: ["Some ", "bold text", " centered "],
            formats: [nil, bold, nil], format: center)

        // Example 4. A math example with a superscript.
        ws1.richString("A7", string: ["j = k", "(n-1)"],
            formats: [italic, superScript], format: center)

    }

    func testArrayFormula() {
        let wb = Workbook(name: "array_formula.xlsx")
        defer { wb.close() }

        let ws1 = wb.addWorksheet()
        ws1.write(.number(500), "B1")
        ws1.write(.number(10),  "B2")
        ws1.write(.number(1),  "B5")
        ws1.write(.number(2),  "B6")
        ws1.write(.number(3),  "B7")

        ws1.write(.number(300), "C1")
        ws1.write(.number(15),  "C2")
        ws1.write(.number(20234),  "C5")
        ws1.write(.number(21003),  "C6")
        ws1.write(.number(10000),  "C7")

        // Write an array formula that returns a single value.
        ws1.arrayFormula("A1:A1", formula: "{=SUM(B1:C1*B2:C2)}")

        // Similar to above but using the range macro
        ws1.arrayFormula("A2:A2", formula: "{=SUM(B1:C1*B2:C2)}")

        // Write an array formula that returns a range of values.
        ws1.arrayFormula("A5:A7", formula: "{=TREND(C5:C7,B5:B7)}")

    }

    /// MARK: Dynamic/Array Formula
    func testDynamicArrayFormula() {
        let wb = Workbook(name: "dynamic_arrays.xlsx")
        defer { wb.close() }

        let header1 = wb.addFormat()
            .font(color: 0xFFFFFF)
            .background(color: 0x74AC4C)

        let header2 = wb.addFormat()
            .font(color: 0xFFFFFF)
            .background(color: 0x528FD3)

        // Example of using the FILTER() function
        let ws1 = wb.addWorksheet(name: "Filter")
        ws1.dynamicFormula("F2", formula: "=_xlfn._xlws.FILTER(A1:D17,C1:C17=K2)")
        // write the data the function will work on.
        ws1.write("Product", "K1", format: header2)
        ws1.write("Apple", "K2")
        ws1.write("Region", "F1", format: header2)
        ws1.write("Sales Rep", "G1", format: header2)
        ws1.write("Product", "H1", format: header2)
        ws1.write("Units", "I1", format: header2)
        _writeDynamicFormulaData(ws1, format: header1)
        ws1.column("E:E", pixel: 20)
        ws1.column("J:J", pixel: 20)

        // Example of using the UNIQUE() function.
        let ws2 = wb.addWorksheet(name: "Unique")
        ws2.dynamicFormula("F2", formula: "=_xlfn.UNIQUE(B2:B17)")
        // A more complex example combining SORT and UNIQUE.
        ws2.dynamicFormula("H2", formula: "=_xlfn._xlws.SORT(_xlfn.UNIQUE(B2:B17))")
        // write the data the function will work on.
        ws2.write("Sales Rep", "F1", format: header2)
        ws2.write("Sales Rep", "H1", format: header2)
        _writeDynamicFormulaData(ws2, format: header1)
        ws2.column("E:E", pixel: 20)
        ws2.column("G:G", pixel: 20)

        // Example of using the SORT() function.
        let ws3 = wb.addWorksheet(name: "Sort")
        ws3.dynamicFormula("F2", formula: "=_xlfn._xlws.SORT(B2:B17)")
        // A more complex example combining SORT and FILTER.
        ws3.dynamicFormula("H2", formula: "=_xlfn._xlws.SORT(_xlfn._xlws.FILTER(C2:D17,D2:D17>5000,\"\"),2,1)")
        // write the data the function will work on.
        ws3.write("Sales Rep", "F1", format: header2)
        ws3.write("Product", "H1", format: header2)
        ws3.write("Units", "I1", format: header2)
        _writeDynamicFormulaData(ws3, format: header1)
        ws3.column("E:E", pixel: 20)
        ws3.column("G:G", pixel: 20)

        // Example of using the SORTBY() function.
        let ws4 = wb.addWorksheet(name: "Sortby")
        ws4.dynamicFormula("D2", formula: "=_xlfn.SORTBY(A2:B9,B2:B9)")
        // write the data the function will work on.
        ws4.write("Name", "A1", format: header1)
        ws4.write("Age", "B1", format: header1)

        ws4.write("Tom", "A2")
        ws4.write("Fred", "A3")
        ws4.write("Amy", "A4")
        ws4.write("Sal", "A5")
        ws4.write("Fritz", "A6")
        ws4.write("Srivan", "A7")
        ws4.write("Xi", "A8")
        ws4.write("Hector", "A9")

        ws4.write(.number(52), "B2")
        ws4.write(.number(65), "B3")
        ws4.write(.number(22), "B4")
        ws4.write(.number(73), "B5")
        ws4.write(.number(19), "B6")
        ws4.write(.number(39), "B7")
        ws4.write(.number(19), "B8")
        ws4.write(.number(66), "B9")

        ws4.write("Name", "D1", format: header2)
        ws4.write("Age", "E1", format: header2)
        ws4.column("C:C", pixel: 20)

        // Example of using the XLOOKUP() function.
        let ws5 = wb.addWorksheet(name: "Xlookup")
        ws5.dynamicFormula("F1", formula: "=_xlfn.XLOOKUP(E1,A2:A9,C2:C9)")
        // write the data the function will work on.
        ws5.write("Country", "A1", format: header1)
        ws5.write("Abr", "B1", format: header1)
        ws5.write("Prefix", "C1", format: header1)

        ws5.write("China", "A2")
        ws5.write("India", "A3")
        ws5.write("United States", "A4")
        ws5.write("Indonesia", "A5")
        ws5.write("Brazil", "A6")
        ws5.write("Pakistan", "A7")
        ws5.write("Nigeria", "A8")
        ws5.write("Bangladesh", "A9")

        ws5.write("CN", "B2")
        ws5.write("IN", "B3")
        ws5.write("US", "B4")
        ws5.write("ID", "B5")
        ws5.write("BR", "B6")
        ws5.write("PK", "B7")
        ws5.write("NG", "B8")
        ws5.write("BD", "B9")

        ws5.write(.number(86), "C2")
        ws5.write(.number(91), "C3")
        ws5.write(.number(1), "C4")
        ws5.write(.number(62), "C5")
        ws5.write(.number(55), "C6")
        ws5.write(.number(92), "C7")
        ws5.write(.number(234), "C8")
        ws5.write(.number(880), "C9")

        ws5.write("Brazil", "E1", format: header2)
        ws5.column("A:A", pixel: 100)
        ws5.column("D:D", pixel: 20)

        // Example of using the XMATCH() function
        let ws6 = wb.addWorksheet(name: "Xmatch")
        ws6.dynamicFormula("D2", formula: "=_xlfn.XMATCH(C2,A2:A6)")
        // write the data the function will work on.
        ws6.write("Product", "A1", format: header1)
        ws6.write("Apple", "A2")
        ws6.write("Grape", "A3")
        ws6.write("Pear", "A4")
        ws6.write("Banana", "A5")
        ws6.write("Cherry", "A6")

        ws6.write("Product", "C1", format: header2)
        ws6.write("Position", "D1", format: header2)
        ws6.write("Grape", "C2")
        ws6.column("B:B", pixel: 20)

        // Example of using the RANDARRAY() function
        let ws7 = wb.addWorksheet(name: "RandArray")
        ws7.dynamicFormula("A1", formula: "=_xlfn.RANDARRAY(5,3,1,100, TRUE)")

        // Example of using the SEQUENCE() function
        let ws8 = wb.addWorksheet(name: "Sequence")
        ws8.dynamicFormula("A1", formula: "=_xlfn.SEQUENCE(4,5)")

        // Example of using Spill range function
        let ws9 = wb.addWorksheet(name: "Spill ranges")
        ws9.dynamicFormula("H2", formula: "=_xlfn.ANCHORARRAY(F2)")
        ws9.dynamicFormula("J2", formula: "=COUNTA(_xlfn.ANCHORARRAY(F2))")
        ws9.dynamicFormula("F2", formula: "=_xlfn.UNIQUE(B2:B17)")
        ws9.write("Unique", "F1", format: header2)
        ws9.write("Spill", "H1", format: header2)
        ws9.write("Spill", "J1", format: header2)
        _writeDynamicFormulaData(ws9, format: header1)
        ws9.column("E:E", pixel: 20)
        ws9.column("G:G", pixel: 20)
        ws9.column("I:I", pixel: 20)

        // Example of using the dynamic ranges with older Excel functions.
        let ws10 = wb.addWorksheet(name: "Older functions")
        ws10.dynamicArrayFormula("B1:B3", formula: "=LEN(A1:A3)")
        ws10.write("Foo", "A1")
        ws10.write("Food", "A2")
        ws10.write("Frood", "A3")

        
    }

    /* A simple function and data structure to populate some of the worksheets. */
    struct worksheetData {
        let col1: String
        let col2: String
        let col3: String
        let col4: Int
    };

    private func _writeDynamicFormulaData(_ ws: Worksheet, format: Format? = nil){
        let data: [worksheetData] = [
            worksheetData(col1: "East",  col2: "Tom",    col3: "Apple",  col4: 6380),
            worksheetData(col1: "West",  col2: "Fred",   col3: "Grape",  col4: 5619),
            worksheetData(col1: "North", col2: "Amy",    col3: "Pear",   col4: 4565),
            worksheetData(col1: "South", col2: "Sal",    col3: "Banana", col4: 5323),

            worksheetData(col1: "East",  col2: "Fritz",  col3: "Apple",  col4: 4394),
            worksheetData(col1: "West",  col2: "Sravan", col3: "Grape",  col4: 7195),
            worksheetData(col1: "North", col2: "Xi",     col3: "Pear",   col4: 5231),
            worksheetData(col1: "South", col2: "Hector", col3: "Banana", col4: 2427),

            worksheetData(col1: "East",  col2: "Tom",    col3: "Banana", col4: 4213),
            worksheetData(col1: "West",  col2: "Fred",   col3: "Pear",   col4: 3239),
            worksheetData(col1: "North", col2: "Amy",    col3: "Grape",  col4: 6520),
            worksheetData(col1: "South", col2: "Sal",    col3: "Apple",  col4: 1310),

            worksheetData(col1: "East",  col2: "Fritz",  col3: "Banana", col4: 6274),
            worksheetData(col1: "West",  col2: "Sravan", col3: "Pear",   col4: 4894),
            worksheetData(col1: "North", col2: "Xi",     col3: "Grape",  col4: 7580),
            worksheetData(col1: "South", col2: "Hector", col3: "Apple",  col4: 9814),
        ]

        ws.write("Region", "A1", format: format)
        ws.write("Sales Rep", "B1", format: format)
        ws.write("Product", "C1", format: format)
        ws.write("Units", "D1", format: format)

        for row in 0...15 {
            ws.write(.string(data[row].col1), Cell(UInt32(row + 1), UInt16(0)))
            ws.write(.string(data[row].col2), Cell(UInt32(row + 1), UInt16(1)))
            ws.write(.string(data[row].col3), Cell(UInt32(row + 1), UInt16(2)))
            ws.write(.number(Double(data[row].col4)), Cell(UInt32(row + 1), UInt16(3)))
        }
    }

    /// MARK: Auto Filter
    func testAutoFilter() {
        let wb = Workbook(name: "autofilter.xlsx")
        defer { wb.close() }

        struct RowData {
            let region: String
            let item: String
            let volume: Int
            let month: String
        };

        let data: [RowData] = [
            RowData(region: "East",  item: "Apple",    volume: 9000,  month: "July"),
            RowData(region: "East",  item: "Apple",    volume: 5000,  month: "July"),
            RowData(region: "South", item: "Orange",   volume: 9000,  month: "September"),
            RowData(region: "North", item: "Apple",    volume: 2000,  month: "November"),
            RowData(region: "West",  item: "Apple",    volume: 9000,  month: "November"),
            RowData(region: "South", item: "Pear",     volume: 7000,  month: "October"),
            RowData(region: "North", item: "Pear",     volume: 9000,  month: "August"),
            RowData(region: "West",  item: "Orange",   volume: 1000,  month: "December"),
            RowData(region: "West",  item: "Grape",    volume: 1000,  month: "November"),
            RowData(region: "South", item: "Pear",     volume: 10000, month: "April"),
            RowData(region: "West",  item: "Grape",    volume: 6000,  month: "January"),
            RowData(region: "South", item: "Orange",   volume: 3000,  month: "May"),
            RowData(region: "North", item: "Apple",    volume: 3000,  month: "December"),
            RowData(region: "South", item: "Apple",    volume: 7000,  month: "February"),
            RowData(region: "West",  item: "Grape",    volume: 1000,  month: "December"),
            RowData(region: "East",  item: "Grape",    volume: 8000,  month: "February"),
            RowData(region: "South", item: "Grape",    volume: 10000, month: "June"),
            RowData(region: "West",  item: "Pear",     volume: 7000,  month: "December"),
            RowData(region: "South", item: "Apple",    volume: 2000,  month: "October"),
            RowData(region: "East",  item: "Grape",    volume: 7000,  month: "December"),
            RowData(region: "North", item: "Grape",    volume: 6000,  month: "April"),
            RowData(region: "East",  item: "Pear",     volume: 8000,  month: "February"),
            RowData(region: "North", item: "Apple",    volume: 7000,  month: "August"),
            RowData(region: "North", item: "Orange",   volume: 7000,  month: "July"),
            RowData(region: "North", item: "Apple",    volume: 6000,  month: "June"),
            RowData(region: "South", item: "Grape",    volume: 8000,  month: "September"),
            RowData(region: "West",  item: "Apple",    volume: 3000,  month: "October"),
            RowData(region: "South", item: "Orange",   volume: 10000, month: "November"),
            RowData(region: "West",  item: "Grape",    volume: 4000,  month: "July"),
            RowData(region: "North", item: "Orange",   volume: 5000,  month: "August"),
            RowData(region: "East",  item: "Orange",   volume: 1000,  month: "November"),
            RowData(region: "East",  item: "Orange",   volume: 4000,  month: "October"),
            RowData(region: "North", item: "Grape",    volume: 5000,  month: "August"),
            RowData(region: "East",  item: "Apple",    volume: 1000,  month: "December"),
            RowData(region: "South", item: "Apple",    volume: 10000, month: "March"),
            RowData(region: "East",  item: "Grape",    volume: 7000,  month: "October"),
            RowData(region: "West",  item: "Grape",    volume: 1000,  month: "September"),
            RowData(region: "East",  item: "Grape",    volume: 10000, month: "October"),
            RowData(region: "South", item: "Orange",   volume: 8000,  month: "March"),
            RowData(region: "North", item: "Apple",    volume: 4000,  month: "July"),
            RowData(region: "South", item: "Orange",   volume: 5000,  month: "July"),
            RowData(region: "West",  item: "Apple",    volume: 4000,  month: "June"),
            RowData(region: "East",  item: "Apple",    volume: 5000,  month: "April"),
            RowData(region: "North", item: "Pear",     volume: 3000,  month: "August"),
            RowData(region: "East",  item: "Grape",    volume: 9000,  month: "November"),
            RowData(region: "North", item: "Orange",   volume: 8000,  month: "October"),
            RowData(region: "East",  item: "Apple",    volume: 10000, month: "June"),
            RowData(region: "South", item: "Pear",     volume: 1000,  month: "December"),
            RowData(region: "North", item: "Grape",    volume: 10000, month: "July"),
            RowData(region: "East",  item: "Grape",    volume: 6000,  month: "February"),
        ]

        let header = wb.addFormat()
            .bold()

        // Example 1. Autofilter without conditions
        let ws1 = wb.addWorksheet()
        // Set up the worksheet data.
        _writeAutoFilterHeader(ws1, format: header)
        // Write the row data
        data.enumerated().forEach { row in
            // ws.write($0.element, row: $0.offset + 2, format: f2)
            ws1.write(.string(row.element.region), [row.offset+1, 0])        
            ws1.write(.string(row.element.item), [row.offset+1, 1])        
            ws1.write(.int(row.element.volume), [row.offset+1, 2])        
            ws1.write(.string(row.element.month), [row.offset+1, 3])        
        }
        // Add the autofilter
        ws1.autofilter(range: [0, 0, 50, 3])

        // Example 2: Autofilter with a filter condition in the first column.
        let ws2 = wb.addWorksheet()
        // Set up the worksheet data.
        _writeAutoFilterHeader(ws2, format: header)
        // Write the row data
        // Add the autofilter
        data.enumerated().forEach { row in
            let rowIndex = row.offset+1
            ws2.write(.string(row.element.region), [rowIndex, 0])
            ws2.write(.string(row.element.item), [rowIndex, 1])
            ws2.write(.int(row.element.volume), [rowIndex, 2])
            ws2.write(.string(row.element.month), [rowIndex, 3])

            // It isn't sufficient to just apply the filter condition below. We
            // must also hide the rows that don't match the criteria since Excel
            // doesn't do that automatically.
            if row.element.region == "East" {
                // Row matches the filter, no further action required.
            } else {
                // Hide rows that don't match the filter
                ws2.rowOption(rowIndex, hidden: true)
            }
        }
        ws2.autofilter(range: [0, 0, 50, 3])
        ws2.filter(0, criteria: .equalTo, string: "East")

        // Example 3: Autofilter with a dual filter condition in one of the column.
        let ws3 = wb.addWorksheet()
        // Set up the worksheet data.
        _writeAutoFilterHeader(ws3, format: header)
        // Write the row data
        // Add the autofilter
        data.enumerated().forEach { row in
            let rowIndex = row.offset+1
            ws3.write(.string(row.element.region), [rowIndex, 0])
            ws3.write(.string(row.element.item), [rowIndex, 1])
            ws3.write(.int(row.element.volume), [rowIndex, 2])
            ws3.write(.string(row.element.month), [rowIndex, 3])

            // It isn't sufficient to just apply the filter condition below. We
            // must also hide the rows that don't match the criteria since Excel
            // doesn't do that automatically.
            if ["East", "South"].contains(row.element.region) {
                // Row matches the filter, no further action required.
            } else {
                // Hide rows that don't match the filter
                ws3.rowOption(rowIndex, hidden: true)
            }
        }
        ws3.autofilter(range: [0, 0, 50, 3])
        ws3.filter2(0, criteria: .equalTo, string: "East", criteria2: .equalTo, string2: "South", andOr: .or)

        // Example 4: Autofilter with filter condition in two columns.
        let ws4 = wb.addWorksheet()
        // Set up the worksheet data.
        _writeAutoFilterHeader(ws4, format: header)
        // Write the row data
        // Add the autofilter
        data.enumerated().forEach { row in
            let rowIndex = row.offset+1
            ws4.write(.string(row.element.region), [rowIndex, 0])
            ws4.write(.string(row.element.item), [rowIndex, 1])
            ws4.write(.int(row.element.volume), [rowIndex, 2])
            ws4.write(.string(row.element.month), [rowIndex, 3])

            // It isn't sufficient to just apply the filter condition below. We
            // must also hide the rows that don't match the criteria since Excel
            // doesn't do that automatically.
            if ["East"].contains(row.element.region) &&
                (row.element.volume > 3000 && row.element.volume < 8000) {
                // Row matches the filter, no further action required.
            } else {
                // Hide rows that don't match the filter
                ws4.rowOption(rowIndex, hidden: true)
            }
        }
        ws4.autofilter(range: [0, 0, 50, 3])
        ws4.filter(0, criteria: .equalTo, string: "East")
        ws4.filter2(2, criteria: .greaterThan, value: 3000, criteria2: .lessThan, value2: 8000, andOr: .and)

        // Example 5: Autofilter with a dual filter condition in one of the columns.
        let ws5 = wb.addWorksheet()
        // Set up the worksheet data.
        _writeAutoFilterHeader(ws5, format: header)
        // Write the row data
        // Add the autofilter
        data.enumerated().forEach { row in
            let rowIndex = row.offset+1
            ws5.write(.string(row.element.region), [rowIndex, 0])
            ws5.write(.string(row.element.item), [rowIndex, 1])
            ws5.write(.int(row.element.volume), [rowIndex, 2])
            ws5.write(.string(row.element.month), [rowIndex, 3])

            // It isn't sufficient to just apply the filter condition below. We
            // must also hide the rows that don't match the criteria since Excel
            // doesn't do that automatically.
            if ["East", "North", "South"].contains(row.element.region) {
                // Row matches the filter, no further action required.
            } else {
                // Hide rows that don't match the filter
                ws5.rowOption(rowIndex, hidden: true)
            }
        }
        ws5.autofilter(range: [0, 0, 50, 3])
        ws5.filterList(0, list: ["East", "North", "South"])

        // Example 6: Autofilter with filter for blanks.
        let ws6 = wb.addWorksheet()
        // Set up the worksheet data.
        _writeAutoFilterHeader(ws6, format: header)
        // Write the row data
        // Add the autofilter
        data.enumerated().forEach { row in
            let rowIndex = row.offset+1
            if row.offset == 5 {
                ws6.write(.blank, [rowIndex, 0])
            } else {
                ws6.write(.string(row.element.region), [rowIndex, 0])
            }
            ws6.write(.string(row.element.item), [rowIndex, 1])
            ws6.write(.int(row.element.volume), [rowIndex, 2])
            ws6.write(.string(row.element.month), [rowIndex, 3])

            // It isn't sufficient to just apply the filter condition below. We
            // must also hide the rows that don't match the criteria since Excel
            // doesn't do that automatically.
            // if [""].contains(row.element.region) {
            if row.offset == 5 {
                // Row matches the filter, no further action required.
            } else {
                // Hide rows that don't match the filter
                ws6.rowOption(rowIndex, hidden: true)
            }
        }
        ws6.autofilter(range: [0, 0, 50, 3])
        ws6.filter(0, criteria: .blanks)

        // Example 7: Autofilter with filter for none-blanks.
        let ws7 = wb.addWorksheet()
        // Set up the worksheet data.
        _writeAutoFilterHeader(ws7, format: header)
        // Write the row data
        // Add the autofilter
        data.enumerated().forEach { row in
            let rowIndex = row.offset+1
            if row.offset == 5 {
                ws7.write(.blank, [rowIndex, 0])
            } else {
                ws7.write(.string(row.element.region), [rowIndex, 0])
            }
            ws7.write(.string(row.element.item), [rowIndex, 1])
            ws7.write(.int(row.element.volume), [rowIndex, 2])
            ws7.write(.string(row.element.month), [rowIndex, 3])

            // It isn't sufficient to just apply the filter condition below. We
            // must also hide the rows that don't match the criteria since Excel
            // doesn't do that automatically.
            // if [""].contains(row.element.region) {
            if row.offset != 5 {
                // Row matches the filter, no further action required.
            } else {
                // Hide rows that don't match the filter
                ws7.rowOption(rowIndex, hidden: true)
            }
        }
        ws7.autofilter(range: [0, 0, 50, 3])
        ws7.filter(0, criteria: .nonBlanks)

    }

    private func _writeAutoFilterHeader(_ ws: Worksheet, format: Format? = nil) {
        // Make the columns wider for clarity.
        ws.column("A:D", width: 12)

        // Write the column headers
        ws.row(0, height: 20, format: format)
        ws.write("Region", "A1")
        ws.write("Item", "B1")
        ws.write("Volume", "C1")
        ws.write( "Month", "D1")
    }

    /// MARK: Validation
    func testValidation() {
        let wb = Workbook(name: "data_validate1.xlsx")
        defer { wb.close() }

        let header = wb.addFormat()
            .border(style: .thin)
            .fg(color: 0xC6EFCE)
            .bold()
            .textWrap()
            .align(vertical: .center)
            .indent(level: 1)


        // Example 1. Limiting input to an integer in a fixed range.
        let ws1 = wb.addWorksheet()
        // write some data for the validations
        _writeValidationData(ws1, format: header)

        // Set up layout of the worksheet
        ws1.column("A:A", width: 55)
        ws1.column("B:B", width: 15)
        ws1.column("D:D", width: 15)
        ws1.row(0, height: 36)

        ws1.write("Enter an integer between 1 and 10", "A3")
        ws1.validation(row: 2, col: 1, type: .integer, criteria: .between, minNumber: 1, maxNumber: 10)

        // Example 2. Limiting input to an integer outside a fixed range.
        ws1.write("Enter an integer not between 1 and 10 (using cell references)", "A5")
        ws1.validation(row: 4, col: 1, type: .integerFormula, criteria: .notBetween, minFormula: "=E3", maxFormula: "=F3")

        // Example 3. Limiting input to an integer greater than a fixed value
        ws1.write("Enter an integer greater than 0", "A7")
        ws1.validation(row: 6, col: 1, type: .integer, criteria: .greaterThan, value: 0)

        // Example 4. Limiting input to an integer less than a fixed value
        ws1.write("Enter an integer less than 10", "A9")
        ws1.validation(row: 8, col: 1, type: .integer, criteria: .lessThan, value: 10)

        // Example 5. Limiting input to a decimal in a fixed range.
        ws1.write("Enter an decimal between 0.1 and 0.5", "A11")
        ws1.validation(row: 10, col: 1, type: .decimal, criteria: .between, minNumber: 0.1, maxNumber: 0.5)

        // Example 6. Limiting input to a value in a dropdown list.
        ws1.write("Select a value from a drop down list", "A13")
        ws1.validation(row: 12, col: 1, type: .list, list: ["open", "high", "close"])

        // Example 7. Limiting input to a value in a dropdown list.
        ws1.write("Select a value from a drop down list (using a cell range)", "A15")
        ws1.validation(row: 14, col: 1, type: .listFormula, valueFormula: "=$E$4:$G$4")

        // Example 8. Limiting input to a date in a fixed range.
        ws1.write("Enter a date between 1/1/2022 and 12/31/2022", "A17")
        // let startDateStr = "2022-01-01T00:00:00+0000"
        // let endDateStr = "2022-12-31T00:00:00+0000"
        // let dateFormatter = ISO8601DateFormatter()
        // let startDate = dateFormatter.date(from: startDateStr)
        // let endDate = dateFormatter.date(from: endDateStr)

        let calendar = Calendar(identifier: .iso8601)
        let dayComponent = DateComponents(year: 2022, month: 1, day: 1, hour: 0, minute: 0, second: 0)
        let startDate = calendar.date(from: dayComponent)
        let dayComponent2 = DateComponents(year: 2022, month: 12, day: 31, hour: 0, minute: 0, second: 0)
        let endDate = calendar.date(from: dayComponent2)
        ws1.validation(row: 16, col: 1, type: .date, criteria: .between, minDate: startDate, maxDate: endDate)

        // Example 9. Limiting input to a time in a fixed range.
        ws1.write("Enter a time between 6:00 and 12:00", "A19")
        ws1.validation(row: 18, col: 1, type: .time, criteria: .between, minTime: [6, 0], maxTime: [12, 0, 0])

        // Example 10. Limiting input to a string greater than a fixed length.
        ws1.write("Enter a string longer than 3 characters", "A21")
        ws1.validation(row: 20, col: 1, type: .length, criteria: .greaterThan, value: 3)

        // Example 11. Limiting input based on a formula.
        ws1.write("Enter a value if the following is true \"=AND(F5=50,G5=60)\"", "A23")
        ws1.validation(row: 22, col: 1, type: .customFormula, valueFormula: "=AND(F5=50,G5=60)")

        // Example 12. Display and modify data validation messages.
        ws1.write("Displays a message when you select the cell", "A25")
        ws1.validation(row: 24, col: 1, type: .integer, criteria: .between, minNumber: 1, maxNumber: 100, title: "Enter an integer:", message: "between 1 and 100")

        // Example 13. Display and modify data validation messages.
        ws1.write("Displays a custom message when integer isn't between 1 and 100", "A27")
        ws1.validation(row: 26, col: 1, type: .integer, criteria: .between, minNumber: 1, maxNumber: 100, 
            title: "Enter an integer:", message: "between 1 and 100", 
            errorTitle: "Input value is not valid!", errorMessage: "It should be an integer between 1 and 100")

        // Example 14. Display and modify data validation messages.
        ws1.write("Displays a custom info message when integer isn't between 1 and 100", "A29")
        ws1.validation(row: 28, col: 1, type: .integer, criteria: .between, minNumber: 1, maxNumber: 100, 
            title: "Enter an integer:", message: "between 1 and 100", 
            errorTitle: "Input value is not valid!", errorMessage: "It should be an integer between 1 and 100",
            errorType: .information)
    }

    private func _writeValidationData(_ ws: Worksheet, format: Format? = nil){
        ws.write("Some examples of data validation in libxlsxwriter", "A1", format: format)
        ws.write("Enter values in this column", "B1", format: format)
        ws.write("Sample Data", "D1", format: format)

        ws.write("Integers", "D3")
        ws.write(.int(1), "E3")
        ws.write(.int(10), "F3")

        ws.write("List Data", "D4")
        ws.write("open", "E4")
        ws.write("high", "F4")
        ws.write("close", "G4")

        ws.write("Formula", "D5")
        ws.write("=AND(F5=50,G5=60)", "E5")
        ws.write(.int(50), "F5")
        ws.write(.int(60), "G5")

    }



    /// MARK: Image
    func testImage() {
        let wb = Workbook(name: "images.xlsx")
        defer { wb.close() }

        let ws1 = wb.addWorksheet()

        // Change some of the column widths for clarity.
        ws1.column("A:A", width: 30)

        // Insert an image.
        ws1.write("Inset an image in a cell:", "A2")
        ws1.image(row: 1, col: 1, fileName: "logo.png")

        // Insert an image offset in the cell
        ws1.write("Inset an offset image.", "A12")
        ws1.imageOpt("B12", fileName: "logo.png", xOffset: 15, yOffset: 10)

        // Insett an image with scaling.
        ws1.write("Inset a scaled image.", "A22")
        ws1.imageOpt("B22", fileName: "logo.png", xScale: 0.5, yScale: 0.5)

        // Insert an image with a hyperlink.
        ws1.write("Inset an image with a hyperlink:", "A32")
        ws1.imageOpt("B32", fileName: "logo.png", url: "https://github.com/jmcnamara")

        // var image_buffer: [CUnsignedChar] = [
        var image_buffer: [UInt8] = [
            0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d,
            0x49, 0x48, 0x44, 0x52, 0x00, 0x00, 0x00, 0x20, 0x00, 0x00, 0x00, 0x20,
            0x08, 0x02, 0x00, 0x00, 0x00, 0xfc, 0x18, 0xed, 0xa3, 0x00, 0x00, 0x00,
            0x01, 0x73, 0x52, 0x47, 0x42, 0x00, 0xae, 0xce, 0x1c, 0xe9, 0x00, 0x00,
            0x00, 0x04, 0x67, 0x41, 0x4d, 0x41, 0x00, 0x00, 0xb1, 0x8f, 0x0b, 0xfc,
            0x61, 0x05, 0x00, 0x00, 0x00, 0x20, 0x63, 0x48, 0x52, 0x4d, 0x00, 0x00,
            0x7a, 0x26, 0x00, 0x00, 0x80, 0x84, 0x00, 0x00, 0xfa, 0x00, 0x00, 0x00,
            0x80, 0xe8, 0x00, 0x00, 0x75, 0x30, 0x00, 0x00, 0xea, 0x60, 0x00, 0x00,
            0x3a, 0x98, 0x00, 0x00, 0x17, 0x70, 0x9c, 0xba, 0x51, 0x3c, 0x00, 0x00,
            0x00, 0x46, 0x49, 0x44, 0x41, 0x54, 0x48, 0x4b, 0x63, 0xfc, 0xcf, 0x40,
            0x63, 0x00, 0xb4, 0x80, 0xa6, 0x88, 0xb6, 0xa6, 0x83, 0x82, 0x87, 0xa6,
            0xce, 0x1f, 0xb5, 0x80, 0x98, 0xe0, 0x1d, 0x8d, 0x03, 0x82, 0xa1, 0x34,
            0x1a, 0x44, 0xa3, 0x41, 0x44, 0x30, 0x04, 0x08, 0x2a, 0x18, 0x4d, 0x45,
            0xa3, 0x41, 0x44, 0x30, 0x04, 0x08, 0x2a, 0x18, 0x4d, 0x45, 0xa3, 0x41,
            0x44, 0x30, 0x04, 0x08, 0x2a, 0x18, 0x4d, 0x45, 0x03, 0x1f, 0x44, 0x00,
            0xaa, 0x35, 0xdd, 0x4e, 0xe6, 0xd5, 0xa1, 0x22, 0x00, 0x00, 0x00, 0x00,
            0x49, 0x45, 0x4e, 0x44, 0xae, 0x42, 0x60, 0x82            
        ]
        ws1.imageBuffer("B42", imageBuffer: &image_buffer, count: image_buffer.count)
        ws1.imageBufferOpt("B52", imageBuffer: &image_buffer, count: image_buffer.count, 
            xOffset: 34, yOffset: 4, xScale: 2, yScale: 1)

    }

    /// MARK: Headers and Footers
//  * The control characters used in the header/footer strings are:
//  *
//  *     Control             Category            Description
//  *     =======             ========            ===========
//  *     &L                  Justification       Left
//  *     &C                                      Center
//  *     &R                                      Right
//  *
//  *     &P                  Information         Page number
//  *     &N                                      Total number of pages
//  *     &D                                      Date
//  *     &T                                      Time
//  *     &F                                      File name
//  *     &A                                      Worksheet name
//  *
//  *     &fontsize           Font                Font size
//  *     &"font,style"                           Font name and style
//  *     &U                                      Single underline
//  *     &E                                      Double underline
//  *     &S                                      Strikethrough
//  *     &X                                      Superscript
//  *     &Y                                      Subscript
//  *
//  *     &[Picture]          Images              Image placeholder
//  *     &G                                      Same as &[Picture]
//  *
//  *     &&                  Miscellaneous       Literal ampersand &
//  *
//  * Copyright 2014-2021, John McNamara, jmcnamara@cpan.org    
    func testHeadersAndFooters() {
        let wb = Workbook(name: "headers_footers.xlsx")
        defer { wb.close() }

        // A simple example to start
        let ws1 = wb.addWorksheet(name: "Simple")
        ws1.header("&CHere is some centered text.")
        ws1.footer("&LHere is some left aligned text.")        
        ws1.column("A:A", width: 50)
        ws1.write("Select Print Preview to see the header and footer", "A1")


        // A simple example of image
        let ws2 = wb.addWorksheet(name: "Image")
        ws2.headerOpt("&L&[Picture]", imageLeft: "logo_small.png")        
        ws2.margins(left: -1, right: -1, top: 1.3, bottom: -1)
        ws2.column("A:A", width: 50)
        ws2.write("Select Print Preview to see the header and footer", "A1")

        // Example of some of the header/footer variables.
        let ws3 = wb.addWorksheet(name: "Variables")
        ws3.header("&LPage &P of &N &CFilename: &F &RSheetname: &A")
        ws3.footer("&LCurrent date: &D &RCurrent time: &T")
        ws3.column("A:A", width: 50)
        ws3.write("Select Print Preview to see the header and footer", "A1")
        var breaks: [UInt32] = [20, 0]
        ws3.pageBreaks(&breaks)
        ws3.write("Next page", "A21")

        // Example of how to use more than on font.
        let ws4 = wb.addWorksheet(name: "Mixed fonts")
        ws4.header("&C&\"Courier New,Bold\"Hello &\"Arial,Italic\"World")
        ws4.footer("&C&\"Symbol\"e&\"Arial\" = mc&X2")
        ws4.column("A:A", width: 50)
        ws4.write("Select Print Preview to see the header and footer", "A1")

        // Example of line wrapping.
        let ws5 = wb.addWorksheet(name: "Word wrap")
        ws5.header("&CHeading 1\nHeading 2")
        ws5.column("A:A", width: 50)
        ws5.write("Select Print Preview to see the header and footer", "A1")

        // Example of inserting a literal ampersand &.
        let ws6 = wb.addWorksheet(name: "Ampersand")
        ws6.header("&CCuriouser && Curiouser - Attorneys at Law")
        ws6.column("A:A", width: 50)
        ws6.write("Select Print Preview to see the header and footer", "A1")

    }

    func testDefinedName() {
        let wb = Workbook(name: "defined_name.xlsx")
        defer { wb.close() }

        // We don't use the returned worksheets in the example and use a generic
        // loop instead.
        let _ = wb.addWorksheet()
        let _ = wb.addWorksheet()
        logger.info("--count: \(wb.sheetCount)")
        logger.info("--SheetNames: \(wb.sheetNames)")

        // Define some global/workbook names.
        wb.defineName(name: "Sales", formula: "=!G1:H10")
        wb.defineName(name: "Exchange_rate", formula: "=0.96")
        wb.defineName(name: "Sheet1!Sales", formula: "=Sheet1!$G$1:$H$10")

        // Define a local/worksheet name
        wb.defineName(name: "Sheet2!Sales", formula: "=Sheet2!$G$1:$H$10")

        // Write some text to the worksheets and one of the defined name in a formula.
        wb.sheetNames.forEach{ sn in
            logger.info("write sheet: \(sn)")
            if let ws = wb[worksheet: sn] {
                logger.info("get sheet: \(ws)")
                ws.column("A:A", width: 45)
                ws.write("This worksheet contains some define names.", "A1")
                ws.write("See Formulas -> Name Manager above.", "A2")
                ws.write("Example formula in cell B3 ->.", "A3")
                ws.write(.formula("=Exchange_rate"), "B3")
            }
        }




        // ws1.header("&CHere is some centered text.")
        // ws1.footer("&LHere is some left aligned text.")        
        // ws1.column("A:A", width: 50)
        // ws1.write("Select Print Preview to see the header and footer", "A1")
    }


// C语言指针类型	         swift语言指针对象类型
// char *	                UnsafeMutablePointer<Int8>
// const char *	            UnsafePointer<Int8>
// unsigned char *	        UnsafeMutablePointer<UInt8>
// const unsigned char *	UnsafePointer<UInt8>
// void *	                UnsafeMutableRawPointer
// const void *	            UnsafeRawPointer


    static var allTests = [
        ("testExample", testExample),
        ("testTable", testTable),
        ("testConditionalFormatSimple", testConditionalFormatSimple),
        ("testConditionalFormat", testConditionalFormat),
        ("testRichStringFormat", testRichStringFormat),
        ("testArrayFormula", testArrayFormula),
        ("testDynamicArrayFormula", testDynamicArrayFormula),
        ("testAutoFilter", testAutoFilter),
        ("testValidation", testValidation),
        ("testImage", testImage),
        ("testHeadersAndFooters", testHeadersAndFooters),
        ("testDefinedName", testDefinedName),
    ]
}

