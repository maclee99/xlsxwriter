import XCTest
import Cxlsxwriter
import Logging
@testable import xlsxwriter

final class macroTests: XCTestCase {
    let logger = Logger(label: "macroTest")
    

    func testMacro() {
        let wb = Workbook(name: "macro.xlsm")
        defer { wb.close() }

        let ws2 = wb.addWorksheet(name: "Button")
        ws2.column("A:A", width: 30)
        wb.addVBA(file: "vbaProject.bin")
        ws2.write("Press the button to say hello.", "A3")
        ws2.button("B3", caption: "Press Me", macro: "say_hello", width: 80, height: 30)

    }

    static var allTests = [
        ("testMacro", testMacro),
    ]
}


