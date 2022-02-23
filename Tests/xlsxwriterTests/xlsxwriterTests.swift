import XCTest
@testable import xlsxwriter

final class xlsxwriterTests: XCTestCase {
    func testExample() {
        // Create a new workbook
        let wb = Workbook(name: "demo.xlsx")
        defer { wb.close() }
        
        // Add a format.
        let f = wb
            .addFormat()
            .bold()
            .border(style: .thin)
            .align(horizontal: .center)
            .align(vertical: .center)
        
        // Add a format.
        let f2 = wb.addFormat().center()

        let f3 = wb.addFormat().background(color: .fillGreen)   //.font(color: .white)
        
        // Add a worksheet.
        let ws = wb
            .addWorksheet()
            .tab(color: .blue)
            .set_default(row_height: 25)
            .write("Number", "A1", format: f)
            .write("Batch 1", "B1", format: f)
            .write("Batch 2", "C1", format: f)
            .column("A:C", width: 30)
            .gridline(screen: false)

        ws.merge(["Merged Range"], firstRow: 1, firstCol: 0, lastRow: 1, lastCol: 10, format: f3)
        
        // Create random data
        let data = (1...100).map {
            [Double($0), Double.random(in: 10...100), Double.random(in: 20...50)]
        }
        
        // Write data to add to plot on the chart.
        data.enumerated().forEach {
            ws.write($0.element, row: $0.offset + 2, format: f2)
        }
        
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
            .activate()
            .set(chart: chart) // Insert the chart into the chartsheet.
    }

    static var allTests = [
        ("testExample", testExample),
    ]
}
