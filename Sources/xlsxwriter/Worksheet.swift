//
//  Worksheet.swift
//  Created by Daniel Müllenborn on 31.12.20.
//

import Cxlsxwriter

/// Struct to represent an Excel worksheet.
public final class Worksheet {
    private var lxw_worksheet: lxw_worksheet
    // private let worksheet: UnsafeMutablePointer<lxw_worksheet>
    private var sheet: UnsafeMutablePointer<lxw_worksheet> {
        get {
            return withUnsafeMutablePointer(to: &self.lxw_worksheet){ $0 }
        }
    }

    var name: String {
        String(cString: lxw_worksheet.name)
    }

    init(_ lxw_worksheet: lxw_worksheet) {
        self.lxw_worksheet = lxw_worksheet 
    }

    /// Insert a chart object into a worksheet.
    public func insert(chart: Chart, _ pos: (row: Int, col: Int)) -> Worksheet {
        let r = UInt32(pos.row)
        let c = UInt16(pos.col)
        // _ = withUnsafeMutablePointer(to: &lxw_worksheet) {
        let error = worksheet_insert_chart(self.sheet, r, c, chart.lxw_chart)
        if error.rawValue != 0 { 
            print("error when insert(chart): \(String(cString: lxw_strerror(error)))") 
        }
        // }
        return self
    }

    /// Insert a chart object into a worksheet, with options.
    public func insert(chart: Chart, _ pos: (row: Int, col: Int), scale: (x: Double, y: Double)) -> Worksheet {
        let r = UInt32(pos.row)
        let c = UInt16(pos.col)
        var o = lxw_chart_options(x_offset: 0, y_offset: 0, x_scale: scale.x, y_scale: scale.y, object_position: 2, description: nil, decorative: 0)
        // let _ = withUnsafeMutablePointer(to: &lxw_worksheet) { 
        let error = worksheet_insert_chart_opt(self.sheet, r, c, chart.lxw_chart, &o)
        if error.rawValue != 0 { 
            print("error when insert(chart opt): \(String(cString: lxw_strerror(error)))") 
        }
        // }
        return self
    }

    /// Write a column of data starting from (row, col).
    @discardableResult public func write(column values: [Value], _ cell: Cell, format: Format? = nil) -> Worksheet {
        var r = cell.row
        let c = cell.col
        for value in values {
            write(value, .init(r, c), format: format)
            r += 1
        }
        return self
    }

    /// Write a row of data starting from (row, col).
    @discardableResult public func write(row values: [Value], _ cell: Cell, format: Format? = nil) -> Worksheet {
        let r = cell.row
        var c = cell.col
        for value in values {
            write(value, .init(r, c), format: format)
            c += 1
        }
        return self
    }

    /// Write a row of Double values starting from (row, col).
    @discardableResult public func write(_ numbers: [Double], row: Int, col: Int = 0, format: Format? = nil) -> Worksheet {
        let f = format?.lxw_format
        let r = UInt32(row)
        var c = UInt16(col)
        // withUnsafeMutablePointer(to: &lxw_worksheet) {
            for number in numbers {
                worksheet_write_number(self.sheet, r, c, number, f)
                c += 1
            // }
        }
        return self
    }

    /// Write a row of String values starting from (row, col).
    @discardableResult public func write(_ strings: [String], row: Int, col: Int = 0, format: Format? = nil) -> Worksheet {
        let f = format?.lxw_format
        let r = UInt32(row)
        var c = UInt16(col)
        // withUnsafeMutablePointer(to: &lxw_worksheet) { sheet in
            for string in strings {
                let error = string.withCString { s in worksheet_write_string(self.sheet, r, c, s, f) }
                if error.rawValue != 0 { 
                    print("error when write(strings): \(String(cString: lxw_strerror(error)))") 
                }
                c += 1
            }
        // }
        return self
    }

    /// Write data to a worksheet cell by calling the appropriate
    /// worksheet_write_*() method based on the type of data being passed.
    @discardableResult public func write(_ value: Value, _ cell: Cell, format: Format? = nil) -> Worksheet {
        let r = cell.row
        let c = cell.col
        let f = format?.lxw_format

        // withUnsafeMutablePointer(to: &lxw_worksheet) { sheet in 
            let error: lxw_error
            switch value {
                case .number(let number): error = worksheet_write_number(self.sheet, r, c, number, f)
                case .string(let string): error = string.withCString { s in worksheet_write_string(self.sheet, r, c, s, f) }
                case .url(let url): error = url.path.withCString { s in worksheet_write_url(self.sheet, r, c, s, f) }
                case .blank: error = worksheet_write_blank(self.sheet, r, c, f)
                case .comment(let comment): error = comment.withCString { s in worksheet_write_comment(self.sheet, r, c, s) }
                case .boolean(let boolean): error = worksheet_write_boolean(self.sheet, r, c, Int32(boolean ? 1 : 0), f)
                case .formula(let formula): error = formula.withCString { s in worksheet_write_formula(self.sheet, r, c, s, f) }
                case .datetime(let datetime):
                    error = lxw_error(rawValue: 0)
                    let num = (datetime.timeIntervalSince1970 / 86400) + 25569
                    worksheet_write_number(self.sheet, r, c, num, f)
                }
            if error.rawValue != 0 { fatalError(String(cString: lxw_strerror(error))) }
        // }
        return self
    }

    /// Set a worksheet tab as selected.
    @discardableResult public func select() -> Worksheet {
        // withUnsafeMutablePointer(to: &lxw_worksheet) { 
        //     worksheet_select($0) 
        // }
        worksheet_select(self.sheet) 
        return self
    }

    /// Hide the current worksheet.
    @discardableResult public func hide() -> Worksheet {
        // withUnsafeMutablePointer(to: &lxw_worksheet) { 
        //     worksheet_hide($0) 
        // }
        worksheet_hide(self.sheet) 
        return self
    }

    /// Make a worksheet the active, i.e., visible worksheet.
    @discardableResult public func activate() -> Worksheet {
        // withUnsafeMutablePointer(to: &lxw_worksheet) {
        //     worksheet_activate($0) 
        // }
        worksheet_activate(self.sheet) 
        return self
    }

    /// Hide zero values in worksheet cells.
    @discardableResult public func hide_zero() -> Worksheet {
        // withUnsafeMutablePointer(to: &lxw_worksheet) {
        //     worksheet_hide_zero($0) 
        // }
        worksheet_hide_zero(self.sheet) 
        return self
    }

    /// Set the paper type for printing.
    @discardableResult public func paper(type: PaperType) -> Worksheet {
        // withUnsafeMutablePointer(to: &lxw_worksheet) { 
        //     worksheet_set_paper($0, type.rawValue) 
        // }
        worksheet_set_paper(self.sheet, type.rawValue) 
        return self
    }

    /// Set the properties for one or more columns of cells.
    /// MARK: NOT WORK
    @discardableResult public func column(_ cols: Cols, width: Double, format: Format? = nil) -> Worksheet {
        let first = cols.col
        let last = cols.col2
        let f = format?.lxw_format

        // withUnsafeMutablePointer(to: &lxw_worksheet) { 
            let error = worksheet_set_column(self.sheet, first, last, width, f) 
            if error.rawValue != 0 { 
                print("error when column: \(String(cString: lxw_strerror(error)))") 
            }
        // }
        return self
    }

    /// Set the properties for one or more columns of cells.
    @discardableResult public func hide_columns(_ col: Int, width: Double = 8.43) -> Worksheet {
        let first = UInt16(col)
        let cols: Cols = "A:XFD"
        let last = cols.col2
        var o = lxw_row_col_options(hidden: 1, level: 0, collapsed: 0)
        // _ = withUnsafeMutablePointer(to: &lxw_worksheet) { 
        //     worksheet_set_column_opt($0, first, last, width, nil, &o) 
        // }
        let error = worksheet_set_column_opt(self.sheet, first, last, width, nil, &o) 
        if error.rawValue != 0 { 
            print("error when column_opt: \(String(cString: lxw_strerror(error)))") 
        }
        return self
    }

    /// Set the color of the worksheet tab.
    /// MARK: NOT WORK
    @discardableResult public func tab(color: Color) -> Worksheet {
        // withUnsafeMutablePointer(to: &lxw_worksheet) { 
        //     worksheet_set_tab_color($0, color.rawValue) 
        // }
        worksheet_set_tab_color(self.sheet, color.rawValue) 
        return self
    }

    /// Set the default row properties.
    @discardableResult public func set_default(row_height: Double, hide_unused_rows: Bool = true) -> Worksheet {
        let hide: UInt8 = hide_unused_rows ? 1 : 0
        // withUnsafeMutablePointer(to: &lxw_worksheet) { 
        //     worksheet_set_default_row($0, row_height, hide) 
        // }
        worksheet_set_default_row(self.sheet, row_height, hide) 
        return self
    }

    /// Set the print area for a worksheet.
    @discardableResult public func print_area(range: Range) -> Worksheet {
        // withUnsafeMutablePointer(to: &lxw_worksheet) { 
        //     let _ = worksheet_print_area($0, range.row, range.col, range.row2, range.col2) 
        // }
        let error = worksheet_print_area(self.sheet, range.row, range.col, range.row2, range.col2) 
        if error.rawValue != 0 { 
            print("error when print_area: \(String(cString: lxw_strerror(error)))") 
        }
        return self
    }

    /// Set the autofilter area in the worksheet.
    @discardableResult public func autofilter(range: Range) -> Worksheet {
        // withUnsafeMutablePointer(to: &lxw_worksheet) { 
        //     let _ = worksheet_autofilter($0, range.row, range.col, range.row2, range.col2) 
        // }
        let error = worksheet_autofilter(self.sheet, range.row, range.col, range.row2, range.col2) 
        if error.rawValue != 0 { 
            print("error when autofilter: \(String(cString: lxw_strerror(error)))") 
        }
        return self
    }

    /// Set the option to display or hide gridlines on the screen and the printed page.
    @discardableResult public func gridline(screen: Bool, print: Bool = false) -> Worksheet {
        // withUnsafeMutablePointer(to: &lxw_worksheet) { 
        //     worksheet_gridlines($0, UInt8((print ? 2 : 0) + (screen ? 1 : 0))) 
        // }
        worksheet_gridlines(self.sheet, UInt8((print ? 2 : 0) + (screen ? 1 : 0))) 
        return self
    }

    private var table_columns = [lxw_table_column]()
    /// Set a table in the worksheet.
    @discardableResult public func table(range: Range, name: String? = nil, header: [String] = [], totalRow: Bool = false) -> Worksheet {
        var options = lxw_table_options()

        if let name = name { options.name = name.makeCString() } //(from: name) }

        options.style_type = UInt8(LXW_TABLE_STYLE_TYPE_MEDIUM.rawValue)
        options.style_type_number = 7
        options.total_row = 0
        let buffer = UnsafeMutableBufferPointer<UnsafeMutablePointer<lxw_table_column>?>.allocate(capacity: header.count + 1)
        table_columns = Array(repeating: lxw_table_column(), count: header.count)
        for i in header.indices {
            table_columns[i].header = header[i].makeCString()   //(from: header[i])
            withUnsafeMutablePointer(to: &table_columns[i]) { 
                buffer.baseAddress?.advanced(by: i).pointee = $0 
            }
        }
        options.columns = buffer.baseAddress
        // _ = withUnsafeMutablePointer(to: &lxw_worksheet) { 
        //     worksheet_add_table($0, range.row, range.col, range.row2, range.col2, &options) 
        // }
        let error = worksheet_add_table(self.sheet, range.row, range.col, range.row2, range.col2, &options) 
        if error.rawValue != 0 { 
            print("error when table: \(String(cString: lxw_strerror(error)))") 
        }
        return self
    }

    ///  Additional functions by Mac
    ///  allows cells to be merged together so that they act as a single area.
    /// MARK: NOT WORK
    @discardableResult public func merge(_ string: String, firstRow: Int, firstCol: Int = 0, 
      lastRow: Int, lastCol: Int, format: Format? = nil) -> Worksheet {
        let f = format?.lxw_format
        let r1 = UInt32(firstRow)
        let c1 = UInt16(firstCol)
        let r2 = UInt32(lastRow)
        let c2 = UInt16(lastCol)
        // print("format: \(f)")

        // withUnsafeMutablePointer(to: &lxw_worksheet) { sheet in
            let error = string.withCString { s in 
                worksheet_merge_range(self.sheet, r1, c1, r2, c2, s, f) 
            }
            if error.rawValue != 0 { 
                print("error when merge: \(String(cString: lxw_strerror(error)))") 
            }
            // print("merge: \(error)")
        // }
        return self
    }

    ///  allows cells to be merged together so that they act as a single area.
    /// MARK: NOT WORK
    @discardableResult public func freeze(row: Int, col: Int) -> Worksheet {
        let r = UInt32(row)
        let c = UInt16(col)
        // withUnsafeMutablePointer(to: &self.lxw_worksheet) { sheet in
            // worksheet_freeze_panes(sheet, r, c) 
        // }
        worksheet_freeze_panes(self.sheet, r, c) 
        return self
    }

}

// internal func makeCString(from str: String) -> UnsafeMutablePointer<CChar> {
//     let count = str.utf8.count + 1
//     let result = UnsafeMutablePointer<CChar>.allocate(capacity: count)
//     str.withCString { result.initialize(from: $0, count: count) }
//     return result
// }

extension String {
    func makeCString() -> UnsafeMutablePointer<CChar> {
        let count = self.utf8.count + 1
        let result = UnsafeMutablePointer<CChar>.allocate(capacity: count)
        self.withCString { result.initialize(from: $0, count: count) }
        return result
    }
}