import XCTest
import Cxlsxwriter
import Logging
@testable import xlsxwriter

final class outlineTests: XCTestCase {
    let logger = Logger(label: "outlineTest")

    // func testExample() {
    //     // Create a new workbook
    //     let wb = Workbook(name: "demo.xlsx")
    //     defer { wb.close() }
    // }
    

    func testOutline() {
        let wb = Workbook(name: "outline.xlsx")
        defer { wb.close() }

        // We don't use the returned worksheets in the example and use a generic
        // loop instead.
        let ws1 = wb.addWorksheet(name: "Outlined Rows")
        let ws2 = wb.addWorksheet(name: "Collapsed Rows")
        let ws3 = wb.addWorksheet(name: "Outline Columns")
        let ws4 = wb.addWorksheet(name: "Outline levels")

        logger.info("--count: \(wb.sheetCount)")
        logger.info("--SheetNames: \(wb.sheetNames)")

        let bold = wb.addFormat()
            .bold()

        /*
        * Example 1: Create a worksheet with outlined rows. It also includes
        * SUBTOTAL() functions so that it looks like the type of automatic
        * outlines that are generated when you use the 'Sub Totals' option.
        *
        * For outlines the important parameters are 'hidden' and 'level'. Rows
        * with the same 'level' are grouped together. The group will be collapsed
        * if 'hidden' is non-zero.
        */
        ws1.column("A:A", width: 20)

        ws1.rowOption(1, hidden: false, level: 2, collapsed: false)
        ws1.rowOption(2, hidden: false, level: 2, collapsed: false)
        ws1.rowOption(3, hidden: false, level: 2, collapsed: false)
        ws1.rowOption(4, hidden: false, level: 2, collapsed: false)
        ws1.rowOption(5, hidden: false, level: 1, collapsed: false)
        ws1.rowOption(6, hidden: false, level: 2, collapsed: false)
        ws1.rowOption(7, hidden: false, level: 2, collapsed: false)
        ws1.rowOption(8, hidden: false, level: 2, collapsed: false)
        ws1.rowOption(9, hidden: false, level: 2, collapsed: false)
        ws1.rowOption(10, hidden: false, level: 1, collapsed: false)

        // Add data and formulas to the worksheet
        ws1.write("Region", "A1", format: bold)
        ws1.write("North", "A2")
        ws1.write("North", "A3")
        ws1.write("North", "A4")
        ws1.write("North", "A5")
        ws1.write("North Total", "A6", format: bold)

        ws1.write("Sales", "B1", format: bold)
        ws1.write(.int(1000), "B2")
        ws1.write(.int(1200), "B3")
        ws1.write(.int(900), "B4")
        ws1.write(.int(1200), "B5")
        ws1.write(.formula("=SUBTOTAL(9,B2:B5)"), "B6", format: bold)

        ws1.write("South", "A7")
        ws1.write("South", "A8")
        ws1.write("South", "A9")
        ws1.write("South", "A10")
        ws1.write("South Total", "A11", format: bold)

        ws1.write(.int(400), "B7")
        ws1.write(.int(600), "B8")
        ws1.write(.int(500), "B9")
        ws1.write(.int(600), "B10")
        ws1.write(.formula("=SUBTOTAL(9,B7:B10)"), "B11", format: bold)

        ws1.write("Grand Total", "A12", format: bold)
        ws1.write(.formula("=SUBTOTAL(9,B2:B10)"), "B12", format: bold)

        /*
        * Example 2: Create a worksheet with outlined rows. This is the same as
        * the previous example except that the rows are collapsed.  Note: We need
        * to indicate the row that contains the collapsed symbol '+' with the
        * optional parameter, 'collapsed'.
        */
        ws2.column("A:A", width: 20)

        ws2.rowOption(1, hidden: true, level: 2, collapsed: false)
        ws2.rowOption(2, hidden: true, level: 2, collapsed: false)
        ws2.rowOption(3, hidden: true, level: 2, collapsed: false)
        ws2.rowOption(4, hidden: true, level: 2, collapsed: false)
        ws2.rowOption(5, hidden: true, level: 1, collapsed: false)

        ws2.rowOption(6, hidden: true, level: 2, collapsed: false)
        ws2.rowOption(7, hidden: true, level: 2, collapsed: false)
        ws2.rowOption(8, hidden: true, level: 2, collapsed: false)
        ws2.rowOption(9, hidden: true, level: 2, collapsed: false)
        ws2.rowOption(10, hidden: true, level: 1, collapsed: false)
        ws2.rowOption(11, hidden: false, level: 0, collapsed: true)

        // Add data and formulas to the worksheet
        ws2.write("Region", "A1", format: bold)
        ws2.write("North", "A2")
        ws2.write("North", "A3")
        ws2.write("North", "A4")
        ws2.write("North", "A5")
        ws2.write("North Total", "A6", format: bold)

        ws2.write("Sales", "B1", format: bold)
        ws2.write(.int(1000), "B2")
        ws2.write(.int(1200), "B3")
        ws2.write(.int(900), "B4")
        ws2.write(.int(1200), "B5")
        ws2.write(.formula("=SUBTOTAL(9,B2:B5)"), "B6", format: bold)

        ws2.write("South", "A7")
        ws2.write("South", "A8")
        ws2.write("South", "A9")
        ws2.write("South", "A10")
        ws2.write("South Total", "A11", format: bold)

        ws2.write(.int(400), "B7")
        ws2.write(.int(600), "B8")
        ws2.write(.int(500), "B9")
        ws2.write(.int(600), "B10")
        ws2.write(.formula("=SUBTOTAL(9,B7:B10)"), "B11", format: bold)

        ws2.write("Grand Total", "A12", format: bold)
        ws2.write(.formula("=SUBTOTAL(9,B2:B10)"), "B12", format: bold)

        // Example 3. Create a worksheet with outlined columns
        ws3.write("Month", "A1")
        ws3.write("Jan", "B1")
        ws3.write("Feb", "C1")
        ws3.write("MarJan", "D1")
        ws3.write("Apr", "E1")
        ws3.write("May", "F1")
        ws3.write("Jun", "G1")
        ws3.write("Total", "H1")

        ws3.write("North", "A2")
        ws3.write(.int(50), "B2")
        ws3.write(.int(20), "C2")
        ws3.write(.int(15), "D2")
        ws3.write(.int(25), "E2")
        ws3.write(.int(65), "F2")
        ws3.write(.int(80), "G2")
        ws3.write(.formula("=SUM(B2:G2)"), "H2")

        ws3.write("South", "A3")
        ws3.write(.int(10), "B3")
        ws3.write(.int(20), "C3")
        ws3.write(.int(30), "D3")
        ws3.write(.int(50), "E3")
        ws3.write(.int(50), "F3")
        ws3.write(.int(50), "G3")
        ws3.write(.formula("=SUM(B3:G3)"), "H3")

        ws3.write("East", "A4")
        ws3.write(.int(45), "B4")
        ws3.write(.int(75), "C4")
        ws3.write(.int(50), "D4")
        ws3.write(.int(15), "E4")
        ws3.write(.int(75), "F4")
        ws3.write(.int(100), "G4")
        ws3.write(.formula("=SUM(B4:G4)"), "H4")

        ws3.write("West", "A5")
        ws3.write(.int(15), "B5")
        ws3.write(.int(15), "C5")
        ws3.write(.int(55), "D5")
        ws3.write(.int(35), "E5")
        ws3.write(.int(20), "F5")
        ws3.write(.int(50), "G5")
        ws3.write(.formula("=SUM(B5:G5)"), "H5")

        ws3.write(.formula("=SUM(H2:H5)"), "H6", format: bold)
        ws3.row(0, format: bold)
        ws3.column("A:A", width: 10, format: bold)
        ws3.columnOpt("B:G", width: 5, format: bold, hidden: false, level: 1, collapsed: false)
        ws3.column("H:H", width: 10)

        // Example 4. Show all possible outline levels
        ws4.write("Level 1", "A1")
        ws4.write("Level 2", "A2")
        ws4.write("Level 3", "A3")
        ws4.write("Level 4", "A4")
        ws4.write("Level 5", "A5")
        ws4.write("Level 6", "A6")
        ws4.write("Level 7", "A7")
        ws4.write("Level 6", "A8")
        ws4.write("Level 5", "A9")
        ws4.write("Level 4", "A10")
        ws4.write("Level 3", "A11")
        ws4.write("Level 2", "A12")
        ws4.write("Level 1", "A13")

        ws4.rowOption(0, hidden: false, level: 1, collapsed: false)
        ws4.rowOption(1, hidden: false, level: 2, collapsed: false)
        ws4.rowOption(2, hidden: false, level: 3, collapsed: false)
        ws4.rowOption(3, hidden: false, level: 4, collapsed: false)
        ws4.rowOption(4, hidden: false, level: 5, collapsed: false)
        ws4.rowOption(5, hidden: false, level: 6, collapsed: false)
        ws4.rowOption(6, hidden: false, level: 7, collapsed: false)
        ws4.rowOption(7, hidden: false, level: 6, collapsed: false)
        ws4.rowOption(8, hidden: false, level: 5, collapsed: false)
        ws4.rowOption(9, hidden: false, level: 4, collapsed: false)
        ws4.rowOption(10, hidden: false, level: 3, collapsed: false)
        ws4.rowOption(11, hidden: false, level: 2, collapsed: false)
        ws4.rowOption(12, hidden: false, level: 1, collapsed: false)
    }


    func testOutlineCollapsed() {
        let wb = Workbook(name: "outline_collapsed.xlsx")
        defer { wb.close() }

        // We don't use the returned worksheets in the example and use a generic
        // loop instead.
        let ws1 = wb.addWorksheet(name: "Outlined Rows")
        let ws2 = wb.addWorksheet(name: "Collapsed Rows 1")
        let ws3 = wb.addWorksheet(name: "Collapsed Rows 2")
        let ws4 = wb.addWorksheet(name: "Collapsed Rows 3")
        let ws5 = wb.addWorksheet(name: "Outline Columns")
        let ws6 = wb.addWorksheet(name: "Collapsed Columns")

        let bold = wb.addFormat()
            .bold()

        // Example 1. Create a worksheet with outlined rows. It also includes
        // SUBTOTAL() functions so that it looks like the type of automatic
        // outlines that are generated when you use the 'Sub Totals' option.
        //
        // For outlines that important parameters are 'hidden' and 'level'. Rows
        // with the same 'level' are grouped together. The group will be collapsed
        // if 'hidden' is none-zero.

        // Set the row outline properties set
        ws1.rowOption(1, hidden: false, level: 2, collapsed: false)
        ws1.rowOption(2, hidden: false, level: 2, collapsed: false)
        ws1.rowOption(3, hidden: false, level: 2, collapsed: false)
        ws1.rowOption(4, hidden: false, level: 2, collapsed: false)
        ws1.rowOption(5, hidden: false, level: 1, collapsed: false)

        ws1.rowOption(6, hidden: false, level: 2, collapsed: false)
        ws1.rowOption(7, hidden: false, level: 2, collapsed: false)
        ws1.rowOption(8, hidden: false, level: 2, collapsed: false)
        ws1.rowOption(9, hidden: false, level: 2, collapsed: false)
        ws1.rowOption(10, hidden: false, level: 1, collapsed: false)

        // write the sub-total data that is common the the row examples.
        self._creteRowExampleData(ws1, format: bold)

        // Example 2. Create a worksheet with collapsed outlined rows.
        // This is the same as the exaples 1 except that the all rows are collapsed.
        ws2.rowOption(1, hidden: true, level: 2, collapsed: false)
        ws2.rowOption(2, hidden: true, level: 2, collapsed: false)
        ws2.rowOption(3, hidden: true, level: 2, collapsed: false)
        ws2.rowOption(4, hidden: true, level: 2, collapsed: false)
        ws2.rowOption(5, hidden: true, level: 1, collapsed: false)

        ws2.rowOption(6, hidden: true, level: 2, collapsed: false)
        ws2.rowOption(7, hidden: true, level: 2, collapsed: false)
        ws2.rowOption(8, hidden: true, level: 2, collapsed: false)
        ws2.rowOption(9, hidden: true, level: 2, collapsed: false)
        ws2.rowOption(10, hidden: true, level: 1, collapsed: false)
        ws2.rowOption(11, hidden: false, level: 0, collapsed: true)

        // write the sub-total data that is common the the row examples.
        self._creteRowExampleData(ws2, format: bold)

        // Example 3. Create a worksheet with collapsed outlined rows. Same as the
        // example 1 except that the two sub-totals are collapsed.
        ws3.rowOption(1, hidden: true, level: 2, collapsed: false)
        ws3.rowOption(2, hidden: true, level: 2, collapsed: false)
        ws3.rowOption(3, hidden: true, level: 2, collapsed: false)
        ws3.rowOption(4, hidden: true, level: 2, collapsed: false)
        ws3.rowOption(5, hidden: false, level: 1, collapsed: true)

        ws3.rowOption(6, hidden: true, level: 2, collapsed: false)
        ws3.rowOption(7, hidden: true, level: 2, collapsed: false)
        ws3.rowOption(8, hidden: true, level: 2, collapsed: false)
        ws3.rowOption(9, hidden: true, level: 2, collapsed: false)
        ws3.rowOption(10, hidden: false, level: 1, collapsed: true)

        // write the sub-total data that is common the the row examples.
        self._creteRowExampleData(ws3, format: bold)

        // Example 4. Create a worksheet with outlined rows. Same as the example 1
        // except that th two sub-totals are collapsed
        ws4.rowOption(1, hidden: true, level: 2, collapsed: false)
        ws4.rowOption(2, hidden: true, level: 2, collapsed: false)
        ws4.rowOption(3, hidden: true, level: 2, collapsed: false)
        ws4.rowOption(4, hidden: true, level: 2, collapsed: false)
        ws4.rowOption(5, hidden: true, level: 1, collapsed: true)

        ws4.rowOption(6, hidden: true, level: 2, collapsed: false)
        ws4.rowOption(7, hidden: true, level: 2, collapsed: false)
        ws4.rowOption(8, hidden: true, level: 2, collapsed: false)
        ws4.rowOption(9, hidden: true, level: 2, collapsed: false)
        ws4.rowOption(10, hidden: true, level: 1, collapsed: true)
        ws4.rowOption(11, hidden: false, level: 0, collapsed: true)

        // write the sub-total data that is common the the row examples.
        self._creteRowExampleData(ws4, format: bold)

        // Example 5. Create a worksheet with outlined columns
        self._createColExampleData(ws5, format: bold)
        ws5.row(0, format: bold)
        ws5.column("A:A", width: 10, format: bold)
        ws5.columnOpt("B:G", width: 5, hidden: false, level: 1, collapsed: false)
        ws5.column("H:H", width: 10, format: bold)

        // Example 6. Create a worksheet with outlined columns
        self._createColExampleData(ws6, format: bold)
        ws6.row(0, format: bold)
        ws6.column("A:A", width: 10, format: bold)
        ws6.columnOpt("B:G", width: 5, hidden: true, level: 1, collapsed: false)
        ws6.columnOpt("H:H", width: 10, hidden: false, level: 0, collapsed: true)

    }

    private func _creteRowExampleData(_ ws: Worksheet, format: Format? = nil) {
        // Set the column width for clarity.
        ws.column("A:A", width: 20)

        // Add data and formulas to the worksheet
        ws.write("Region", "A1", format: format)
        ws.write("North", "A2")
        ws.write("North", "A3")
        ws.write("North", "A4")
        ws.write("North", "A5")
        ws.write("North Total", "A6", format: format)

        ws.write("Sales", "B1", format: format)
        ws.write(.int(1000), "B2")
        ws.write(.int(1200), "B3")
        ws.write(.int(900), "B4")
        ws.write(.int(1200), "B5")
        ws.write(.formula("=SUBTOTAL(9,B2:B5)"), "B6", format: format)

        ws.write("South", "A7")
        ws.write("South", "A8")
        ws.write("South", "A9")
        ws.write("South", "A10")
        ws.write("South Total", "A11", format: format)

        ws.write(.int(400), "B7")
        ws.write(.int(600), "B8")
        ws.write(.int(500), "B9")
        ws.write(.int(600), "B10")
        ws.write(.formula("=SUBTOTAL(9,B7:B10)"), "B11", format: format)

        ws.write("Grand Total", "A12", format: format)
        ws.write(.formula("=SUBTOTAL(9,B2:B10)"), "B12", format: format)
    }

    private func _createColExampleData(_ ws: Worksheet, format: Format? = nil) {
        ws.write("Month", "A1")
        ws.write("Jan", "B1")
        ws.write("Feb", "C1")
        ws.write("Mar", "D1")
        ws.write("Apr", "E1")
        ws.write("May", "F1")
        ws.write("Jun", "G1")
        ws.write("Total", "H1")

        ws.write("North", "A2")
        ws.write(.int(50), "B2")
        ws.write(.int(20), "C2")
        ws.write(.int(15), "D2")
        ws.write(.int(25), "E2")
        ws.write(.int(65), "F2")
        ws.write(.int(80), "G2")
        ws.write(.formula("=SUM(B2:G2)"), "H2")

        ws.write("South", "A3")
        ws.write(.int(10), "B3")
        ws.write(.int(20), "C3")
        ws.write(.int(35), "D3")
        ws.write(.int(50), "E3")
        ws.write(.int(50), "F3")
        ws.write(.int(50), "G3")
        ws.write(.formula("=SUM(B3:G3)"), "H3")

        ws.write("East", "A4")
        ws.write(.int(45), "B4")
        ws.write(.int(75), "C4")
        ws.write(.int(50), "D4")
        ws.write(.int(15), "E4")
        ws.write(.int(75), "F4")
        ws.write(.int(100), "G4")
        ws.write(.formula("=SUM(B4:G4)"), "H4")

        ws.write("West", "A5")
        ws.write(.int(15), "B5")
        ws.write(.int(15), "C5")
        ws.write(.int(55), "D5")
        ws.write(.int(35), "E5")
        ws.write(.int(20), "F5")
        ws.write(.int(50), "G5")
        ws.write(.formula("=SUM(B5:G5)"), "H5")

        ws.write(.formula("=SUM(H2:H5)"), "H6", format: format)
    }



    static var allTests = [
        ("testOutline", testOutline),
        ("testOutlineCollapsed", testOutlineCollapsed),
        // ("testExample", testExample),
    ]
}