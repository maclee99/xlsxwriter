//
//  Workbook.swift
//  Created by Daniel Müllenborn on 31.12.20.
//

import Cxlsxwriter
import Logging

/// Struct to represent an Excel workbook.
public final class Workbook {

    var lxw_workbook: UnsafeMutablePointer<lxw_workbook>
    let logger = Logger(label: "Workbook")

    /// Create a new workbook object.
    public init(name: String) {
        logger.info("init: \(name)")
        self.lxw_workbook = name.withCString { 
            workbook_new($0) 
        }
    }

    /// Close the Workbook object and write the XLSX file.
    public func close() {
        logger.info("close...")
        let error = workbook_close(lxw_workbook)
        if error.rawValue != 0 { fatalError(String(cString: lxw_strerror(error))) }
    }

    /// Add a new worksheet to the Excel workbook.
    public func addWorksheet(name: String? = nil) -> Worksheet {
        logger.info("addWorksheet: \(String(describing: name))")
        let worksheet: UnsafeMutablePointer<lxw_worksheet>
        if let name = name {
            worksheet = name.withCString { workbook_add_worksheet(lxw_workbook, $0) }
        } else {
            worksheet = workbook_add_worksheet(lxw_workbook, nil)
        }
        return Worksheet(worksheet)
    }

    /// Add a new chartsheet to a workbook.
    public func addChartsheet(name: String? = nil) -> Chartsheet {
        let chartsheet: UnsafeMutablePointer<lxw_chartsheet>
        if let name = name {
            chartsheet = name.withCString { workbook_add_chartsheet(lxw_workbook, $0) }
        } else {
            chartsheet = workbook_add_chartsheet(lxw_workbook, nil)
        }
        return Chartsheet(chartsheet)
    }

    /// Add a new format to the Excel workbook.
    public func addFormat() -> Format { 
        Format(workbook_add_format(lxw_workbook)) 
    }

    /// Create a new chart to be added to a worksheet
    public func addChart(type: Chart_type) -> Chart { 
        Chart(workbook_add_chart(lxw_workbook, type.rawValue)) 
    }
  
    /// Get a worksheet object from its name.
    public subscript(worksheet name: String) -> Worksheet? {
        guard let ws = name.withCString({ s in workbook_get_worksheet_by_name(lxw_workbook, s) }) else { return nil }
        return Worksheet(ws)
    }
  
    /// Get a chartsheet object from its name.
    public subscript(chartsheet name: String) -> Chartsheet? {
        guard let cs = name.withCString({ s in workbook_get_chartsheet_by_name(lxw_workbook, s) }) else { return nil }
        return Chartsheet(cs)
    }
    
    /// Validate a worksheet or chartsheet name.
    func validate(sheet_name: String) { 
        logger.info("validate: \(sheet_name)")
        let error = sheet_name.withCString { workbook_validate_sheet_name(lxw_workbook, $0) } 
        if error.rawValue != 0 { 
            logger.error("error-> validate: \(String(cString: lxw_strerror(error)))") 
        }
    }

    /// Additionam func by Mac Lee
    @discardableResult public func properties(title: String? = nil, subject: String? = nil, 
      author: String? = nil, manager: String? = nil, company: String? = nil,
      category: String? = nil, keywords: String? = nil, comments: String? = nil,
      status: String? = nil) -> Workbook {

        var properties = lxw_doc_properties()
        var doSet = false
        // if title != nil && !title!.isEmpty {
        if let t = title?.makeCString() {
            defer { t.deallocate() }
            properties.title = t  // title!.makeCString()
            doSet = true
        }
        // if subject != nil && !subject!.isEmpty {
        if let subject = subject?.makeCString() {
            defer { subject.deallocate() }
            properties.subject = subject  // makeCString(from: subject!)
            doSet = true
        }
        // if author != nil && !author!.isEmpty {
        if let author = author?.makeCString() {
            defer { author.deallocate() }
            properties.author = author  // makeCString(from: author!)
            doSet = true
        }
        // if manager != nil && !manager!.isEmpty {
        if let manager = manager?.makeCString() {
            defer { manager.deallocate() }
            properties.manager = manager  // makeCString(from: manager!)
            doSet = true
        }
        if let company = company?.makeCString() {
            defer { company.deallocate() }
            properties.company = company  //makeCString(from: company!)
            doSet = true
        }
        if let category = category?.makeCString() {
            defer { category.deallocate() }
            properties.category = category  //makeCString(from: category!)
            doSet = true
        }
        if let keywords = keywords?.makeCString() {
            defer { keywords.deallocate() }
            properties.keywords = keywords  // makeCString(from: keywords!)
            doSet = true
        }
        if let comments = comments?.makeCString() {
            defer { comments.deallocate() }
            properties.comments = comments  // makeCString(from: comments!)
            doSet = true
        }
        if let status = status?.makeCString() {
            defer { status.deallocate() }
            properties.status = status  // makeCString(from: status!)
            doSet = true
        }

        if doSet {
            _ = workbook_set_properties(lxw_workbook, &properties)
        }

        return self
    }


}