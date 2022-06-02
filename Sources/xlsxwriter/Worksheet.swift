//
//  Worksheet.swift
//  Created by Daniel Müllenborn on 31.12.20.
//

import Foundation
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
            case .int(let number): error = worksheet_write_number(self.sheet, r, c, Double(number), f)
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

    /// write strings with multiple formats
    @discardableResult public func richString(_ cell: Cell, string: [String] = [], formats: [Format?] = [], format: Format? = nil) -> Worksheet {
        logger.info("richString: \(cell)|\(string)|\(formats)") 
        var string_fragments = [lxw_rich_string_tuple]()
        // let buffer = UnsafeMutableBufferPointer<UnsafeMutablePointer<lxw_rich_string_tuple>?>.allocate(capacity: string.count + 1)
        // defer { buffer.deallocate() }
        var fragments: [UnsafeMutablePointer<lxw_rich_string_tuple>?] = [UnsafeMutablePointer<lxw_rich_string_tuple>?]()

        if !string.isEmpty {
            fragments = Array(repeating: UnsafeMutablePointer<lxw_rich_string_tuple>.allocate(capacity: 1), count: string.count+1)
            fragments[string.count] = nil
            string_fragments = Array(repeating: lxw_rich_string_tuple(), count: string.count)
            for i in string.indices {
                // table_columns[i].header = (from: header[i].makeCString())
                string_fragments[i] = lxw_rich_string_tuple()
                string_fragments[i].string = string[i].makeCString() 
                if formats.endIndex > i && formats[i] != nil {
                    string_fragments[i].format = formats[i]?.lxw_format
                }
                fragments[i] = withUnsafeMutablePointer(to: &string_fragments[i]){$0}
            }
        } 

        logger.info("call function: \(fragments)") 
        let error = worksheet_write_rich_string(self.sheet, cell.row, cell.col, 
            &fragments,
            format?.lxw_format);

        if error.rawValue != 0 { 
            logger.error("error-> richString: \(String(cString: lxw_strerror(error)))") 
            fatalError(String(cString: lxw_strerror(error))) 
        }

        logger.info("free allocates")
        string_fragments.forEach { if let _ = $0.string {$0.string.deallocate() } }

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

    /// Set the properties for one or more columns of cells.
    @discardableResult public func column(_ cols: Cols, pixel: UInt32, format: Format? = nil) -> Worksheet {
        logger.info("column: \(cols)|pixel: \(pixel)|\(String(describing: format))")
        let firstCol = cols.col
        let lastCol = cols.col2
        let f = format?.lxw_format

        let error = worksheet_set_column_pixels(self.sheet, firstCol, lastCol, pixel, f) 
        if error.rawValue != 0 { 
            logger.error("error-> column(pixel): \(String(cString: lxw_strerror(error)))") 
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

    @discardableResult public func column(_ col: lxw_col_t, _ col2: lxw_col_t, pixel: UInt32, format: Format? = nil) -> Worksheet {
        logger.info("column: \(col)|\(col2)|pixel: \(pixel)|\(String(describing: format))")
        // let firstCol = UInt16(col)
        // let lastCol = UInt16(col2)
        let f = format?.lxw_format

        let error = worksheet_set_column_pixels(self.sheet, col, col2, pixel, f) 
        if error.rawValue != 0 { 
            logger.error("error--> column(pixel): \(String(cString: lxw_strerror(error)))") 
        }
        return self
    }

    /// change the default properties of a row. The most common use for this function is to change the height of a row
    /// The height is specified in character units
    @discardableResult public func row(_ row: UInt32, height: Double, format: Format? = nil) -> Worksheet {
        logger.info("row: \(row)|\(height)|\(String(describing: format))")
        // let rowToSet = UInt32(row)
        let f = format?.lxw_format

        let error = worksheet_set_row(self.sheet, row, height, f) 
        if error.rawValue != 0 { 
            logger.error("error--> row: \(String(cString: lxw_strerror(error)))") 
        }
        return self
    }

    @discardableResult public func rowOption(_ row: Int, height: Double = LXW_DEF_ROW_HEIGHT, 
        hidden: Bool? = nil, level: Int? = nil, collapsed: Bool? = nil,
        format: Format? = nil) -> Worksheet {
        var opt = lxw_row_col_options()
        if let hidden = hidden { opt.hidden = hidden ? LxwBoolean.true.rawValue : LxwBoolean.false.rawValue }
        if let level = level { opt.level = UInt8(level) }
        if let collapsed = collapsed { opt.collapsed = collapsed ? LxwBoolean.true.rawValue : LxwBoolean.false.rawValue }

        let error = worksheet_set_row_opt(self.sheet, UInt32(row), height, format?.lxw_format, &opt) 
        if error.rawValue != 0 { 
            logger.error("error--> rowOption: \(String(cString: lxw_strerror(error)))") 
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

    /// Set the color of the worksheet tab.
    @discardableResult public func margins(left: Double = 0.7, right: Double = 0.7, top: Double = 0.75, bottom: Double = 0.75) -> Worksheet {
        logger.info("margins: \(String(describing: left))|\(String(describing: right))")
        worksheet_set_margins(self.sheet, left, right, top, bottom) 
        return self
    }

    // /// Add horizontal page breaks
    @discardableResult public func pageBreaks(_ breaks: UnsafeMutablePointer<UInt32>) -> Worksheet {
        logger.info("pageBreaks: \(String(describing: breaks))")
        // let buffer = UnsafeMutableBufferPointer<UnsafeMutablePointer<UInt32>>.allocate(capacity: breaks.count)
        // defer { buffer.deallocate() }
        let error = worksheet_set_h_pagebreaks(self.sheet, breaks) 
        if error.rawValue != 0 { 
            logger.error("error--> pageBreaks: \(String(cString: lxw_strerror(error)))") 
        }

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

    @discardableResult public func filter(_ col: Int, criteria: FilterCriteria? = nil,
        string: String? = nil, value: Double? = nil) -> Worksheet {
        logger.info("filter: \(col)|\(String(describing: criteria))|\(String(describing: string))|\(String(describing: value))")

        var rule = lxw_filter_rule()
        if let c = criteria { rule.criteria = c.rawValue  }
        if let s = string {
            let str = s.makeCString()
            // defer { str.deallocate() }
            rule.value_string = str
        }
        if let v = value { rule.value = v }

        // var filter_rule2 = lxw_filter_rule()
        // filter_rule2.criteria = UInt8(LXW_FILTER_CRITERIA_EQUAL_TO.rawValue)
        // filter_rule2.value_string = "East".makeCString()
        // logger.info("filter: \(rule)|\(filter_rule2)")
        logger.info("rule: \(rule)")

        let error = worksheet_filter_column(self.sheet, UInt16(col), &rule) 
        if error.rawValue != 0 { 
            logger.error("error--> filter: \(String(cString: lxw_strerror(error)))") 
        }

        if let _ = string {
            rule.value_string.deallocate()
        }


        return self
    }

    @discardableResult public func filter2(_ col: Int, 
        criteria: FilterCriteria? = nil, string: String? = nil, value: Double? = nil, 
        criteria2: FilterCriteria? = nil, string2: String? = nil, value2: Double? = nil, 
        andOr: FilterOperator = .or) -> Worksheet {

        var rule = lxw_filter_rule()
        if let c = criteria { rule.criteria = c.rawValue  }
        if let s = string {
            let str = s.makeCString()
            // defer { str.deallocate() }
            rule.value_string = str
        }
        if let v = value { rule.value = v }

        var rule2 = lxw_filter_rule()
        if let c2 = criteria2 { rule2.criteria = c2.rawValue  }
        if let s2 = string2 {
            let str2 = s2.makeCString()
            // defer { str.deallocate() }
            rule2.value_string = str2
        }
        if let v2 = value2 { rule2.value = v2 }

        logger.info("rule: \(rule)|\(rule2)|\(andOr.rawValue)|\(LXW_FILTER_OR)")


        let error = worksheet_filter_column2(self.sheet, UInt16(col), &rule, &rule2, andOr.rawValue) 
        if error.rawValue != 0 { 
            logger.error("error--> filter2: \(String(cString: lxw_strerror(error)))") 
        }

        if let _ = string {
            rule.value_string.deallocate()
        }
        if let _ = string2 {
            rule2.value_string.deallocate()
        }

        return self
    }

    @discardableResult public func filterList(_ col: Int, list: [String] = [] ) -> Worksheet {
        logger.info("filterList: \(col)|\(list)")

        //char* list[] = {"March", "April", "May", NULL};
        var filters: [UnsafeMutablePointer<CChar>?] = []
        list.forEach{ str in
            filters.append(str.makeCString())
        }
        filters.append(nil)

        //char ** 	list 
        let error = worksheet_filter_list(self.sheet, UInt16(col), &filters) 
        if error.rawValue != 0 { 
            logger.error("error--> filterList: \(String(cString: lxw_strerror(error)))") 
        }

        filters.forEach { if let _ = $0 {$0?.deallocate() } }

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
    @discardableResult public func table(range: Range, name: String? = nil, header: [String] = [], 
        headerFormat: [Format?] = [], format: [Format?] = [], totalRow: [TotalFunction] = [], formula: [String] = [],
        autoFilter: Bool = true, headerRow: Bool = true, firstColumn: Bool = false, lastColumn: Bool = false,
        bandedColumns: Bool = false, bandedRows: Bool = true, styleType: UInt8? = nil, 
        styleNumber: UInt8? = nil) -> Worksheet {
        
        var defaultOptions = lxw_table_options()

        if let name = name { defaultOptions.name = name.makeCString() } //(from: name) }

        defaultOptions.style_type = styleType != nil ? styleType! : UInt8(LXW_TABLE_STYLE_TYPE_MEDIUM.rawValue)
        defaultOptions.style_type_number = styleNumber != nil ? styleNumber! : 9  //7
        defaultOptions.no_autofilter = autoFilter ? LxwBoolean.false.rawValue : LxwBoolean.true.rawValue
        defaultOptions.no_header_row = headerRow ? LxwBoolean.false.rawValue : LxwBoolean.true.rawValue
        defaultOptions.first_column = firstColumn ? LxwBoolean.true.rawValue : LxwBoolean.false.rawValue
        defaultOptions.last_column  = lastColumn ? LxwBoolean.true.rawValue : LxwBoolean.false.rawValue
        defaultOptions.no_banded_rows = bandedRows ? LxwBoolean.false.rawValue : LxwBoolean.true.rawValue
        defaultOptions.banded_columns = bandedColumns ? LxwBoolean.true.rawValue : LxwBoolean.false.rawValue
        // options.total_row = 0
        defaultOptions.total_row = totalRow.isEmpty ? UInt8(LXW_FALSE.rawValue) : UInt8(LXW_TRUE.rawValue)

        var table_columns = [lxw_table_column]()
        let buffer = UnsafeMutableBufferPointer<UnsafeMutablePointer<lxw_table_column>?>.allocate(capacity: header.count + 1)
        defer { buffer.deallocate() }

        if !header.isEmpty {
            table_columns = Array(repeating: lxw_table_column(), count: header.count)
            for i in header.indices {
                // table_columns[i].header = (from: header[i].makeCString())
                table_columns[i].header = header[i].makeCString()
                if headerFormat.endIndex > i {
                    table_columns[i].header_format = headerFormat[i]?.lxw_format
                }
                if format.endIndex > i {
                    // table_columns[i].header_format = format[i]?.lxw_format
                    table_columns[i].format = format[i]?.lxw_format
                }
                if totalRow.endIndex > i {
                    if totalRow[i] == .none {
                        if (formula.endIndex > i) {
                            if !formula[i].isEmpty && !formula[i].hasPrefix("=") {
                                table_columns[i].total_string = formula[i].makeCString()
                            } else {
                                table_columns[i].total_function = totalRow[i].rawValue
                            }
                        }
                    } else {
                        table_columns[i].total_function = totalRow[i].rawValue
                    }

                }
                if (formula.endIndex > i) {
                    if !formula[i].isEmpty && formula[i].hasPrefix("=") {
                        table_columns[i].formula = formula[i].makeCString()
                    }
                }
                withUnsafeMutablePointer(to: &table_columns[i]) {
                    buffer.baseAddress?.advanced(by: i).pointee = $0
                }
            }
            defaultOptions.columns = buffer.baseAddress
        }

        let rowDelta: UInt32 = 0    // (totalRow.isEmpty ? 0 : 1)
        let error = worksheet_add_table(self.sheet, range.row, range.col, range.row2 + rowDelta, 
            range.col2, &defaultOptions) 
        if error.rawValue != 0 { 
            logger.error("error--> table: \(String(cString: lxw_strerror(error)))") 
        }

        // deallocate
        if let _ = name { defaultOptions.name.deallocate() }
        table_columns.forEach { 
            $0.header.deallocate()
            if let _ = $0.total_string { $0.total_string.deallocate() }
            if let _ = $0.formula { $0.formula.deallocate() }
        }

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
        let cellStr = string.makeCString()
        defer { cellStr.deallocate() }
        let error = worksheet_merge_range(lxw_worksheet, r1, c1, r2, c2, cellStr, f) 
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

    @discardableResult public func arrayFormula(_ range: Range, formula: String, format: Format? = nil) -> Worksheet {

        let formulaStr = formula.makeCString()
        defer { formulaStr.deallocate() }
        let error = worksheet_write_array_formula(
            self.sheet, range.row, range.col, range.row2, range.col2, formulaStr, format?.lxw_format)
        if error.rawValue != 0 { 
            logger.error("error--> arrayFormula: \(String(cString: lxw_strerror(error)))") 
        }
        return self
    }

    @discardableResult public func dynamicArrayFormula(_ range: Range, formula: String, format: Format? = nil) -> Worksheet {

        let formulaStr = formula.makeCString()
        defer { formulaStr.deallocate() }
        let error = worksheet_write_dynamic_array_formula(
            self.sheet, range.row, range.col, range.row2, range.col2, formulaStr, format?.lxw_format)
        if error.rawValue != 0 { 
            logger.error("error--> dynamicArrayFormula: \(String(cString: lxw_strerror(error)))") 
        }
        return self
    }

    @discardableResult public func dynamicFormula(_ cell: Cell, formula: String, format: Format? = nil) -> Worksheet {

        let formulaStr = formula.makeCString()
        defer { formulaStr.deallocate() }
        let error = worksheet_write_dynamic_formula(
            self.sheet, cell.row, cell.col, formulaStr, format?.lxw_format)
        if error.rawValue != 0 { 
            logger.error("error--> dynamicFormula: \(String(cString: lxw_strerror(error)))") 
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

    ///  allows cells to be merged together so that they act as a single area.
    @discardableResult public func validation(_ cell: Cell, type: ValidationTypes = .none,
        criteria: ValidationCriteria = .none, minNumber: Double? = nil, maxNumber: Double? = nil,
        minFormula: String? = nil, maxFormula: String? = nil, value: Double? = nil,
        list: [String] = [], valueFormula: String? = nil, minDate: Date? = nil, maxDate: Date? = nil,
        minTime: Time? = nil, maxTime: Time? = nil, title: String? = nil, message: String? = nil,
        errorTitle: String? = nil, errorMessage: String? = nil, errorType: ValidationErrorTypes? = nil
        ) -> Worksheet {
        return self.validation(row: Int(cell.row), col: Int(cell.col), type: type,
        criteria: criteria, minNumber: minNumber, maxNumber: maxNumber,
        minFormula: minFormula, maxFormula: maxFormula, value: value,
        list: list, valueFormula: valueFormula, minDate: minDate, maxDate: maxDate,
        minTime: minTime, maxTime: maxTime, title: title, message: message,
        errorTitle: errorTitle, errorMessage: errorMessage, errorType: errorType)
    }

    @discardableResult public func validation(row: Int, col: Int, type: ValidationTypes = .none,
        criteria: ValidationCriteria = .none, minNumber: Double? = nil, maxNumber: Double? = nil,
        minFormula: String? = nil, maxFormula: String? = nil, value: Double? = nil,
        list: [String] = [], valueFormula: String? = nil, minDate: Date? = nil, maxDate: Date? = nil,
        minTime: Time? = nil, maxTime: Time? = nil, title: String? = nil, message: String? = nil,
        errorTitle: String? = nil, errorMessage: String? = nil, errorType: ValidationErrorTypes? = nil
        ) -> Worksheet {

        logger.info("validation: \(row)|\(col)|\(type)|\(criteria)|\(String(describing: minNumber))|\(String(describing: maxNumber))|\(String(describing: minFormula))|\(String(describing: maxFormula))")
        logger.info("LXW_VALIDATION_CRITERIA_NOT_BETWEEN=\(LXW_VALIDATION_CRITERIA_NOT_BETWEEN.rawValue)")

        let r = UInt32(row)
        let c = UInt16(col)
        //char *list[] = {"open", "high", "close", NULL};
        let buffer = UnsafeMutableBufferPointer<UnsafeMutablePointer<CChar>?>.allocate(capacity: list.count+1)
        defer { buffer.deallocate() }

        if type != .none {
            var option = lxw_data_validation()
            option.validate = type.rawValue
            option.criteria = criteria.rawValue
            if let minN = minNumber { option.minimum_number = minN }
            if let maxN = maxNumber { option.maximum_number = maxN }
            if let val  = value { option.value_number = val }
            if let minF = minFormula { option.minimum_formula = minF.makeCString() }
            if let maxF = maxFormula { option.maximum_formula = maxF.makeCString() }
            if let valueF = valueFormula { option.value_formula = valueF.makeCString() }
            if list.count > 0 {
                for i in list.indices {
                    logger.info("-->\(i)|\(list[i])")
                    buffer.baseAddress?.advanced(by: i).pointee = list[i].makeCString()
                }
                buffer.baseAddress?.advanced(by: list.count).pointee = nil
                logger.info("\(buffer)|\(String(describing: buffer.baseAddress))")
                //char **value_list;
                option.value_list = buffer.baseAddress
            }
            if let minD = minDate {
                let calendar = Calendar(identifier: .iso8601)   //Calendar.current
                let c = calendar.dateComponents([.year, .month, .day, .hour, .minute, .second], from: minD)
                let d = lxw_datetime(year: Int32(c.year!), month: Int32(c.month!), day: Int32(c.day!), hour: Int32(c.hour!), min: Int32(c.minute!), sec: Double(c.second!))
                option.minimum_datetime = d
            }
            if let maxD = maxDate {
                let calendar = Calendar(identifier: .iso8601)   //Calendar.current
                let c = calendar.dateComponents([.year, .month, .day, .hour, .minute, .second], from: maxD)
                let d = lxw_datetime(year: Int32(c.year!), month: Int32(c.month!), day: Int32(c.day!), hour: Int32(c.hour!), min: Int32(c.minute!), sec: Double(c.second!))
                option.maximum_datetime = d
            }
            if let minT = minTime {
                let d = lxw_datetime(year: Int32(0), month: Int32(0), day: Int32(0), hour: minT.hour, min: minT.min, sec: minT.second)
                option.minimum_datetime = d
            }
            if let maxT = maxTime {
                let d = lxw_datetime(year: Int32(0), month: Int32(0), day: Int32(0), hour: maxT.hour, min: maxT.min, sec: maxT.second)
                option.maximum_datetime = d
            }
            if let title = title { option.input_title = title.makeCString() }
            if let message = message { option.input_message = message.makeCString() }
            if let errtitle = errorTitle { option.error_title = errtitle.makeCString() }
            if let errmessage = errorMessage { option.error_message = errmessage.makeCString() }
            if let errType = errorType { option.error_type = errType.rawValue }


            logger.info("option: \(option)")

            let error = worksheet_data_validation_cell(self.sheet, r, c, &option)
            if error.rawValue != 0 { 
                logger.error("error--> validation: \(String(cString: lxw_strerror(error)))") 
            }

            // free allocated resources by makeCString()
            if let _ = option.minimum_formula { option.minimum_formula.deallocate() }
            if let _ = option.maximum_formula { option.maximum_formula.deallocate() }
            if let _ = option.value_formula   { option.value_formula.deallocate() }
            if let _ = option.input_title     { option.input_title.deallocate() }
            if let _ = option.input_message   { option.input_message.deallocate() }
            if let _ = option.error_title     { option.error_title.deallocate() }
            if let _ = option.error_message   { option.error_message.deallocate() }

        }

        return self
    }


    ///  allows cells to be merged together so that they act as a single area.
    @discardableResult public func conditionFormat(range: Range, type: ConditionalFormatTypes = .none,
        criteria: ConditionalCriteria = .none, value: Double? = nil, valueString: String? = nil,
        min: Double? = nil, max: Double? = nil, multiRange: String? = nil, 
        minColor: UInt32? = nil, midColor: UInt32? = nil, maxColor: UInt32? = nil,
        barOnly: Bool? = nil, barColor: UInt32? = nil, barSolid: Bool? = nil,
        barDirection: conditionalFormatBarDrection? = nil, bar2010: Bool? = nil,
        negativeColorSame: Bool? = nil, negativeBorderColorSame: Bool? = nil,
        iconStyle: conditionalIconTypes? = nil, reverseIcons: Bool? = nil, iconOnly: Bool? = nil,
        format: Format? = nil ) -> Worksheet {

        if type != .none {
            var option = lxw_conditional_format()
            option.type = type.rawValue
            option.criteria = criteria.rawValue
            if let v = value { option.value = v }
            if let s = valueString { option.value_string = s.makeCString() }
            if let min = min { option.min_value = min }
            if let max = max { option.max_value = max }
            if let mr = multiRange { option.multi_range = mr.makeCString() }
            if let minColor = minColor { option.min_color = minColor }
            if let midColor = midColor { option.mid_color = midColor }
            if let maxColor = maxColor { option.max_color = maxColor }
            if let barOnly = barOnly { option.bar_only =  barOnly ? LxwBoolean.true.rawValue : LxwBoolean.false.rawValue }
            if let barColor = barColor { option.bar_color = barColor }
            if let barSolid = barSolid { option.bar_solid =  barSolid ? LxwBoolean.true.rawValue : LxwBoolean.false.rawValue }
            if let barDirection = barDirection { option.bar_direction =  barDirection.rawValue }
            if let bar2010 = bar2010 { option.data_bar_2010 =  bar2010 ? LxwBoolean.true.rawValue : LxwBoolean.false.rawValue }
            if let negativeColorSame = negativeColorSame { option.bar_negative_color_same =  negativeColorSame ? LxwBoolean.true.rawValue : LxwBoolean.false.rawValue }
            if let negativeBorderColorSame = negativeBorderColorSame { option.bar_negative_border_color_same  =  negativeBorderColorSame ? LxwBoolean.true.rawValue : LxwBoolean.false.rawValue }
            if let iconStyle = iconStyle { option.icon_style =  iconStyle.rawValue }
            if let reverseIcons = reverseIcons { option.reverse_icons  =  reverseIcons ? LxwBoolean.true.rawValue : LxwBoolean.false.rawValue }
            if let iconOnly = iconOnly { option.icons_only =  iconOnly ? LxwBoolean.true.rawValue : LxwBoolean.false.rawValue }
            option.format = format?.lxw_format

            let error = worksheet_conditional_format_range(self.sheet, range.row, range.col, range.row2, 
                range.col2, &option)
            if error.rawValue != 0 { 
                logger.error("error--> conditionFormat: \(String(cString: lxw_strerror(error)))") 
            }

            if let _ = option.multi_range { option.multi_range.deallocate() }
            if let _ = option.value_string { option.value_string.deallocate() }
        }
        // let r = UInt32(row)
        // let c = UInt16(col)
        // worksheet_freeze_panes(self.sheet, r, c) 
        return self
    }

    /// MARK: Image
    ///  Inset a image into a worksheet. The image can be in PNG, JPEG, GIF or BMP format
    @discardableResult public func image(_ cell: Cell, fileName: String? = nil) -> Worksheet {
        return self.image(row: Int(cell.row), col: Int(cell.col), fileName: fileName)
    }

    @discardableResult public func image(row: Int, col: Int, fileName: String? = nil) -> Worksheet {
        let r = UInt32(row)
        let c = UInt16(col)

        if let filename = fileName {
            let fn = filename.makeCString()
            let error = worksheet_insert_image(self.sheet, r, c, fn)
            if error.rawValue != 0 { 
                logger.error("error--> image: \(String(cString: lxw_strerror(error)))") 
            }

            // free the allocated resources
            fn.deallocate()
        }

        return self
    }

    @discardableResult public func imageOpt(_ cell: Cell, fileName: String, 
        xOffset: Int? = nil, yOffset: Int? = nil, xScale: Double? = nil, yScale: Double? = nil,
        position: Int? = nil, description: String? = nil, decorative: Int? = nil, 
        url: String? = nil, tip: String? = nil
        ) -> Worksheet {
        return self.imageOpt(row: Int(cell.row), col: Int(cell.col), fileName: fileName, 
            xOffset: xOffset, yOffset: yOffset, xScale: xScale, yScale: yScale,
            position: position, description: description, decorative: decorative, 
            url: url, tip: tip
        )
    }

    @discardableResult public func imageOpt(row: Int, col: Int, fileName: String, 
        xOffset: Int? = nil, yOffset: Int? = nil, xScale: Double? = nil, yScale: Double? = nil,
        position: Int? = nil, description: String? = nil, decorative: Int? = nil, 
        url: String? = nil, tip: String? = nil
        ) -> Worksheet {
        let r = UInt32(row)
        let c = UInt16(col)

        if !fileName.isEmpty {
            let fn = fileName.makeCString()
            var opt = lxw_image_options()
            if let xOffset = xOffset { opt.x_offset = Int32(xOffset) }
            if let yOffset = yOffset { opt.y_offset = Int32(yOffset) }
            if let xScale = xScale { opt.x_scale = xScale }
            if let yScale = yScale { opt.y_scale = yScale }
            if let position = position { opt.object_position = UInt8(position) }
            if let description = description { opt.description = description.makeCString() }
            if let decorative = decorative { opt.decorative = UInt8(decorative) }
            if let url = url { opt.url = url.makeCString() }
            if let tip = tip { opt.tip = tip.makeCString() }

            let error = worksheet_insert_image_opt(self.sheet, r, c, fn, &opt)
            if error.rawValue != 0 { 
                logger.error("error--> imageOpt: \(String(cString: lxw_strerror(error)))") 
            }

            // free the allocated resources
            fn.deallocate()
            if let _ = opt.description { opt.description.deallocate() }
            if let _ = opt.url { opt.url.deallocate() }
            if let _ = opt.tip { opt.tip.deallocate() }
        }

        return self
    }

    @discardableResult public func imageBuffer(_ cell: Cell, imageBuffer: UnsafePointer<UInt8>, count: Int) -> Worksheet {
        return self.imageBuffer(row: Int(cell.row), col: Int(cell.col), imageBuffer: imageBuffer, count: count)
    }

    @discardableResult public func imageBuffer(row: Int, col: Int, imageBuffer: UnsafePointer<UInt8>, count: Int) -> Worksheet {
        let r = UInt32(row)
        let c = UInt16(col)
        logger.info("\(imageBuffer)|\(count)")

        let error = worksheet_insert_image_buffer(self.sheet, r, c, imageBuffer, count)
        if error.rawValue != 0 { 
            logger.error("error--> imageBuffer: \(String(cString: lxw_strerror(error)))") 
        }

        return self
    }

    @discardableResult public func imageBufferOpt(_ cell: Cell, imageBuffer: UnsafePointer<UInt8>, count: Int, 
        xOffset: Int? = nil, yOffset: Int? = nil, xScale: Double? = nil, yScale: Double? = nil,
        position: Int? = nil, description: String? = nil, decorative: Int? = nil, 
        url: String? = nil, tip: String? = nil ) -> Worksheet {
        return self.imageBufferOpt(row: Int(cell.row), col: Int(cell.col), imageBuffer: imageBuffer, count: count,
            xOffset: xOffset, yOffset: yOffset, xScale: xScale, yScale: yScale,
            position: position, description: description, decorative: decorative, 
            url: url, tip: tip
            )
    }

    @discardableResult public func imageBufferOpt(row: Int, col: Int, imageBuffer: UnsafePointer<UInt8>, count: Int, 
        xOffset: Int? = nil, yOffset: Int? = nil, xScale: Double? = nil, yScale: Double? = nil,
        position: Int? = nil, description: String? = nil, decorative: Int? = nil, 
        url: String? = nil, tip: String? = nil
        ) -> Worksheet {
        let r = UInt32(row)
        let c = UInt16(col)

        var opt = lxw_image_options()
        if let xOffset = xOffset { opt.x_offset = Int32(xOffset) }
        if let yOffset = yOffset { opt.y_offset = Int32(yOffset) }
        if let xScale = xScale { opt.x_scale = xScale }
        if let yScale = yScale { opt.y_scale = yScale }
        if let position = position { opt.object_position = UInt8(position) }
        if let description = description { opt.description = description.makeCString() }
        if let decorative = decorative { opt.decorative = UInt8(decorative) }
        if let url = url { opt.url = url.makeCString() }
        if let tip = tip { opt.tip = tip.makeCString() }

        let error = worksheet_insert_image_buffer_opt(self.sheet, r, c, imageBuffer, count, &opt)
        if error.rawValue != 0 { 
            logger.error("error--> imageBufferOpt: \(String(cString: lxw_strerror(error)))") 
        }

        // free the allocated resources
        if let _ = opt.description { opt.description.deallocate() }
        if let _ = opt.url { opt.url.deallocate() }
        if let _ = opt.tip { opt.tip.deallocate() }

        return self
    }

    /// MARK: Header & Footer
    /// Set a worksheet tab as selected.
    @discardableResult public func header(_ header: String) -> Worksheet {
        let str = header.makeCString()
        let error = worksheet_set_header(self.sheet, str)
        if error.rawValue != 0 { 
            logger.error("error--> header: \(String(cString: lxw_strerror(error)))") 
        }

        str.deallocate()

        return self
    }
    @discardableResult public func headerOpt(_ header: String, margin: Double? = nil,
        imageLeft: String? = nil, imageCenter: String? = nil, imageRight: String? = nil) -> Worksheet {
        let str = header.makeCString()
        var opt = lxw_header_footer_options()
        if let margin = margin { opt.margin = margin }
        if let imageLeft = imageLeft { opt.image_left = imageLeft.makeCString() }
        if let imageCenter = imageCenter { opt.image_center = imageCenter.makeCString() }
        if let imageRight = imageRight { opt.image_right = imageRight.makeCString() }

        let error = worksheet_set_header_opt(self.sheet, str, &opt)
        if error.rawValue != 0 { 
            logger.error("error--> header: \(String(cString: lxw_strerror(error)))") 
        }

        str.deallocate()
        if let _ = opt.image_left { opt.image_left.deallocate() }
        if let _ = opt.image_center { opt.image_center.deallocate() }
        if let _ = opt.image_right { opt.image_right.deallocate() }

        return self
    }

    @discardableResult public func footer(_ footer: String) -> Worksheet {
        let str = footer.makeCString()
        let error = worksheet_set_footer(self.sheet, str)
        if error.rawValue != 0 { 
            logger.error("error--> footer: \(String(cString: lxw_strerror(error)))") 
        }

        str.deallocate()

        return self
    }

}

// https://github.com/apple/swift/blob/main/docs/HowSwiftImportsCAPIs.md#fundamental-types

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

// func bindArrayToTuple<T, U>(array: Array<T>, tuple: UnsafeMutablePointer<U>) {
//     tuple.withMemoryRebound(to: T.self, capacity: array.count) {
//         $0.assign(from: array, count: array.count)
//     }
// }