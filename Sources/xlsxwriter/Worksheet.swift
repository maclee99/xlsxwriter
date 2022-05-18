//
//  Worksheet.swift
//  Created by Daniel Müllenborn on 31.12.20.
//

import Cxlsxwriter
import Logging

/// Struct to represent an Excel worksheet.
public final class Worksheet {
    // private var lxw_worksheet: lxw_worksheet
    private var lxw_worksheet: UnsafeMutablePointer<lxw_worksheet>
    var name: String { String(cString: lxw_worksheet.pointee.name) }
    let logger = Logger(label: "Worksheet")
    // private let worksheet: UnsafeMutablePointer<lxw_worksheet>
    private var sheet: UnsafeMutablePointer<lxw_worksheet> {
        get {
            // return withUnsafeMutablePointer(to: &self.lxw_worksheet){ $0 }
            return self.lxw_worksheet
        }
    }

    // var name: String {
    //     String(cString: lxw_worksheet.name)
    // }
    init(_ lxw_worksheet: UnsafeMutablePointer<lxw_worksheet>) { 
        self.lxw_worksheet = lxw_worksheet 
    }
    // init(_ lxw_worksheet: lxw_worksheet) {
    //     logger.info("init...")
    //     self.lxw_worksheet = lxw_worksheet 
    // }

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

        for number in numbers {
            worksheet_write_number(self.sheet, r, c, number, f)
            c += 1
        }
        return self
    }

    /// Write a row of String values starting from (row, col).
    @discardableResult public func write(_ strings: [String], row: Int, col: Int = 0, format: Format? = nil) -> Worksheet {
        let f = format?.lxw_format
        let r = UInt32(row)
        var c = UInt16(col)

        for string in strings {
            let error = string.withCString { s in worksheet_write_string(self.sheet, r, c, s, f) }
            if error.rawValue != 0 { 
                print("error when write(strings): \(String(cString: lxw_strerror(error)))") 
            }
            c += 1
        }
        return self
    }

    /// Write data to a worksheet cell by calling the appropriate
    /// worksheet_write_*() method based on the type of data being passed.
    @discardableResult public func write(_ value: Value, _ cell: Cell, format: Format? = nil) -> Worksheet {
        logger.info("write: \(value)|\(cell)|\(String(describing: format))")
        let r = cell.row
        let c = cell.col
        let f = format?.lxw_format

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
        if error.rawValue != 0 { 
            logger.error("error-> write: \(String(cString: lxw_strerror(error)))") 
            fatalError(String(cString: lxw_strerror(error))) 
        }

        return self
    }

    /// Set a worksheet tab as selected.
    @discardableResult public func select() -> Worksheet {
        worksheet_select(self.sheet) 
        return self
    }

    /// Hide the current worksheet.
    @discardableResult public func hide() -> Worksheet {
        worksheet_hide(self.sheet) 
        return self
    }

    /// Make a worksheet the active, i.e., visible worksheet.
    @discardableResult public func activate() -> Worksheet {
        worksheet_activate(self.sheet) 
        return self
    }

    /// Hide zero values in worksheet cells.
    @discardableResult public func hideZero() -> Worksheet {
        worksheet_hide_zero(self.sheet) 
        return self
    }

    /// Set the paper type for printing.
    @discardableResult public func paper(type: PaperType) -> Worksheet {
        worksheet_set_paper(self.sheet, type.rawValue) 
        return self
    }

    /// Set the properties for one or more columns of cells.
    @discardableResult public func column(_ cols: Cols, width: Double = LXW_DEF_COL_WIDTH, format: Format? = nil) -> Worksheet {
        logger.info("column: \(cols)|\(width)|\(String(describing: format))")
        let firstCol = cols.col
        let lastCol = cols.col2
        let f = format?.lxw_format

        let error = worksheet_set_column(self.sheet, firstCol, lastCol, width, f) 
        if error.rawValue != 0 { 
            logger.error("error-> column: \(String(cString: lxw_strerror(error)))") 
        }
        return self
    }

    @discardableResult public func column(_ col: lxw_col_t, _ col2: lxw_col_t, width: Double = LXW_DEF_COL_WIDTH, format: Format? = nil) -> Worksheet {
        logger.info("column: \(col)|\(col2)|\(width)|\(String(describing: format))")
        // let firstCol = UInt16(col)
        // let lastCol = UInt16(col2)
        let f = format?.lxw_format

        let error = worksheet_set_column(self.sheet, col, col2, width, f) 
        if error.rawValue != 0 { 
            logger.error("error--> column: \(String(cString: lxw_strerror(error)))") 
        }
        return self
    }

    /// change the default properties of a row. The most common use for this function is to change the height of a row
    /// The height is specified in character units
    @discardableResult public func row(_ row: UInt32, width: Double, format: Format? = nil) -> Worksheet {
        logger.info("row: \(row)|\(width)|\(String(describing: format))")
        // let rowToSet = UInt32(row)
        let f = format?.lxw_format

        let error = worksheet_set_row(self.sheet, row, width, f) 
        if error.rawValue != 0 { 
            logger.error("error--> row: \(String(cString: lxw_strerror(error)))") 
        }
        return self
    }


    /// Set the properties for one or more columns of cells.
    @discardableResult public func hideColumns(_ col: Int, width: Double = 8.43) -> Worksheet {
        let first = UInt16(col)
        let cols: Cols = "A:XFD"
        let last = cols.col2
        // var o = lxw_row_col_options(hidden: 1, level: 0, collapsed: 0)
        var o = lxw_row_col_options()
        o.hidden = 1

        let error = worksheet_set_column_opt(self.sheet, first, last, width, nil, &o)
        if error.rawValue != 0 { 
            logger.error("error--> column_opt: \(String(cString: lxw_strerror(error)))") 
        }
        return self
    }

    /// Set the color of the worksheet tab.
    @discardableResult public func tab(color: Color) -> Worksheet {
        logger.info("tab: \(color)|\(String(format: "0x%08X", color.rawValue))")
        worksheet_set_tab_color(self.sheet, color.rawValue) 
        return self
    }

    /// Set the default row properties.
    @discardableResult public func setDefault(row_height: Double, hide_unused_rows: Bool = true) -> Worksheet {
        // let hide: UInt8  = UInt8(hide_unused_rows ? LXW_TRUE.rawValue : LXW_FALSE.rawValue)
        let hide: UInt8  = hide_unused_rows ? 1 : 0
        logger.info("setDefault: \(row_height)|\(hide)")
        worksheet_set_default_row(self.sheet, row_height, hide) 
        return self
    }

    /// Set the print area for a worksheet.
    @discardableResult public func printArea(range: Range) -> Worksheet {
        let error = worksheet_print_area(self.sheet, range.row, range.col, range.row2, range.col2) 
        if error.rawValue != 0 { 
            logger.error("error--> print_area: \(String(cString: lxw_strerror(error)))") 
        }
        return self
    }

    /// Set the autofilter area in the worksheet.
    @discardableResult public func autofilter(range: Range) -> Worksheet {
        let error = worksheet_autofilter(self.sheet, range.row, range.col, range.row2, range.col2) 
        if error.rawValue != 0 { 
            logger.error("error--> autofilter: \(String(cString: lxw_strerror(error)))") 
        }
        return self
    }

    /// Set the option to display or hide gridlines on the screen and the printed page.
    @discardableResult public func gridline(screen: Bool, print: Bool = false) -> Worksheet {
        worksheet_gridlines(self.sheet, UInt8((print ? 2 : 0) + (screen ? 1 : 0))) 
        return self
    }

    // Set a table in the worksheet.
    @discardableResult public func table(range: Range, name: String? = nil, header: [(String, Format?)] = []) -> Worksheet {
        table(range: range, name: name, header: header.map { $0.0 }, format: header.map { $0.1 }, totalRow: [])
    }

    /// Set a table in the worksheet.
    @discardableResult public func table(range: Range, name: String? = nil, header: [String] = [], format: [Format?] = [], totalRow: [TotalFunction] = []) -> Worksheet {
        var options = lxw_table_options()

        if let name = name { options.name = name.makeCString() } //(from: name) }

        options.style_type = UInt8(LXW_TABLE_STYLE_TYPE_MEDIUM.rawValue)
        options.style_type_number = 7
        // options.total_row = 0
        options.total_row = totalRow.isEmpty ? UInt8(LXW_FALSE.rawValue) : UInt8(LXW_TRUE.rawValue)

        var table_columns = [lxw_table_column]()
        let buffer = UnsafeMutableBufferPointer<UnsafeMutablePointer<lxw_table_column>?>.allocate(capacity: header.count + 1)
        defer { buffer.deallocate() }

        if !header.isEmpty {
            table_columns = Array(repeating: lxw_table_column(), count: header.count)
            for i in header.indices {
                // table_columns[i].header = (from: header[i].makeCString())
                table_columns[i].header = header[i].makeCString()
                if format.endIndex > i {
                    table_columns[i].header_format = format[i]?.lxw_format
                }
                if totalRow.endIndex > i {
                    table_columns[i].total_function = totalRow[i].rawValue
                }
                withUnsafeMutablePointer(to: &table_columns[i]) {
                    buffer.baseAddress?.advanced(by: i).pointee = $0
                }
            }
            options.columns = buffer.baseAddress
        }

        let error = worksheet_add_table(self.sheet, range.row, range.col, range.row2 + (totalRow.isEmpty ? 0 : 1), range.col2, &options) 
        if error.rawValue != 0 { 
            logger.error("error--> table: \(String(cString: lxw_strerror(error)))") 
        }
        if let _ = name { options.name.deallocate() }
        table_columns.forEach { $0.header.deallocate() }
        return self
    }

    ///  Additional functions by Mac
    ///  allows cells to be merged together so that they act as a single area.
    @discardableResult public func merge(_ string: String, firstRow: Int, firstCol: Int, 
      lastRow: Int, lastCol: Int, format: Format? = nil) -> Worksheet {
        let f = format?.lxw_format
        let r1 = UInt32(firstRow)
        let c1 = UInt16(firstCol)
        let r2 = UInt32(lastRow)
        let c2 = UInt16(lastCol)

        logger.info("merge: \(string)|\(firstCol)|\(firstRow)|\(lastCol)|\(lastRow)|\(String(describing: format))")
        let error = worksheet_merge_range(lxw_worksheet, r1, c1, r2, c2, string.makeCString(), f) 
        if error.rawValue != 0 { 
            logger.error("error--> merge: \(String(cString: lxw_strerror(error)))") 
        }

        return self
    }

    /// Merge a range of cells in the worksheet.
    @discardableResult public func merge(range: Range, string: String, format: Format? = nil) -> Worksheet
    {
        let error = worksheet_merge_range(
            lxw_worksheet, range.row, range.col, range.row2, range.col2, string, format?.lxw_format)
        if error.rawValue != 0 { 
            logger.error("error--> merge: \(String(cString: lxw_strerror(error)))") 
        }
        return self
    }

    /// Make a worksheet the active, i.e., visible worksheet.
    @discardableResult public func showComments() -> Worksheet {
        worksheet_show_comments(self.sheet) 
        return self
    }

    ///  allows cells to be merged together so that they act as a single area.
    @discardableResult public func freeze(row: Int, col: Int) -> Worksheet {
        let r = UInt32(row)
        let c = UInt16(col)
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