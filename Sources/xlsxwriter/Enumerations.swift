//
//  Enumerations.swift
//  Created by Daniel Müllenborn on 31.12.20.
//

import Foundation

public enum Value: ExpressibleByFloatLiteral, ExpressibleByStringLiteral, ExpressibleByIntegerLiteral {
    case url(URL)
    case blank
    case comment(String)
    case number(Double)
    case int(Int)
    case string(String)
    case boolean(Bool)
    case formula(String)
    case datetime(Date)
    public init(floatLiteral value: Double) { self = .number(value) }
    public init(integerLiteral value: Int) { self = .int(value) }
    public init(stringLiteral value: String) { self = .string(value) }
}

public enum Axes { case X, Y }

public enum Trendline_type: UInt8 {
    case linear
    case log
    case poly
    case power
    case exp
    case average
}

public enum LxwBoolean: UInt8 {
    case `false` = 0
    case `true` = 1
}

/// Cell border styles for use with format.set_border()
public enum Border: UInt8 {
    case noBorder
    case thin
    case medium
    case dashed
    case dotted
    case thick
    case double
    case hair
    case medium_dashed
    case dash_dot
    case medium_dash_dot
    case dash_dot_dot
    case medium_dash_dot_dot
    case slant_dash_dot
}

/// Alignment values for format.set(alignment:)
public enum HorizontalAlignment: UInt8 {
    case none = 0
    case left
    case center
    case right
    case fill
    case justify
    case center_across
    case distributed
}

/// Alignment values for format.set(alignment:)
public enum VerticalAlignment: UInt8 {
    case top = 8
    case bottom
    case center
    case justify
    case distributed
}

/// The Excel paper format type.
public enum PaperType: UInt8 {
    case PrinterDefault = 0  // Printer default
    case Letter  // 8 1/2 x 11 in
    case LetterSmall  // 8 1/2 x 11 in
    case Tabloid  // 11 x 17 in
    case Ledger  // 17 x 11 in
    case Legal  // 8 1/2 x 14 in
    case Statement  // 5 1/2 x 8 1/2 in
    case Executive  // 7 1/4 x 10 1/2 in
    case A3  // 297 x 420 mm
    case A4  // 210 x 297 mm
    case A4Small  // 210 x 297 mm
    case A5  // 148 x 210 mm
    case B4  // 250 x 354 mm
    case B5  // 182 x 257 mm
    case Folio  // 8 1/2 x 13 in
    case Quarto  // 215 x 275 mm
    case unnamed1  // 10x14 in
    case unnamed2  // 11x17 in
    case Note  // 8 1/2 x 11 in
    case Envelope_9  // 3 7/8 x 8 7/8
    case Envelope_10  // 4 1/8 x 9 1/2
    case Envelope_11  // 4 1/2 x 10 3/8
    case Envelope_12  // 4 3/4 x 11
    case Envelope_14  // 5 x 11 1/2
    case C_size_sheet  // ---
    case D_size_sheet  // ---
    case E_size_sheet  // ---
    case Envelope_DL  // 110 x 220 mm
    case Envelope_C3  // 324 x 458 mm
    case Envelope_C4  // 229 x 324 mm
    case Envelope_C5  // 162 x 229 mm
    case Envelope_C6  // 114 x 162 mm
    case Envelope_C65  // 114 x 229 mm
    case Envelope_B4  // 250 x 353 mm
    case Envelope_B5  // 176 x 250 mm
    case Envelope_B6  // 176 x 125 mm
    case Envelope  // 110 x 230 mm
    case Monarch  // 3.875 x 7.5 in
    case EnvelopeInch  // 3 5/8 x 6 1/2 in
    case Fanfold  // 14 7/8 x 11 in
    case German_Std_Fanfold  // 8 1/2 x 12 in
    case German_Legal_Fanfold  // 8 1/2 x 13 in
}

/// Predefined values for common colors.
public enum Color: UInt32 {
    case black = 0x1000000
    case blue = 0x0000FF
    case brown = 0x800000
    case cyan = 0x00FFFF
    case gray = 0x808080
    case green = 0x008000
    case lime = 0x00FF00
    case magenta = 0xFF00FF
    case navy = 0x000080
    case orange = 0xFF6600
    case purple = 0x800080
    case red = 0xFF0000
    case silver = 0xC0C0C0
    case white = 0xFFFFFF
    case yellow = 0xFFFF00
    // added by mac
    case fillGreen = 0xC8BAFB29
    case fillOrange = 0xC8FFA500
    case fillRed = 0xC8FF0000
    case fillOrangeRed = 0xC8FF4500
    case fillYellowGreen = 0xC89ACD32
    case fillPaleGreen = 0xC898FB98
    case fillWheat = 0xC8F5DEB3
    case fillGray = 0xC8808080
    case fillGold = 0xC8FFD700
    case fillLightPink = 0xC8FDE9D9
}

/// Available chart types.
public enum Chart_type: UInt8 {
    case none
    case area
    case area_stacked
    case area_percentage_stacked
    case bar
    case bar_stacked
    case bar_percentage_stacked
    case column
    case column_stacked
    case column_percentage_stacked
    case doughnut
    case line
    case line_stacked
    case line_percentage_stacked
    case pie
    case scatter
    case scatter_straight
    case scatter_straight_with_markers
    case scatter_smooth
    case scatter_smooth_with_markers
    case radar
    case radar_with_markers
    case radar_filled
}

public enum Legend_position: UInt8 {
    case none = 0
    case right
    case left
    case top
    case bottom
    case top_right
    case overlay_right
    case overlay_left
    case overlay_top_right
}

public enum TotalFunction: UInt8, ExpressibleByIntegerLiteral {
    case none = 0
    /** Use the average function as the table total. */
    case average = 101
    /** Use the count numbers function as the table total. */
    case nums = 102
    /** Use the count function as the table total. */
    case count = 103
    /** Use the max function as the table total. */
    case max = 104
    /** Use the min function as the table total. */
    case min = 105
    /** Use the standard deviation function as the table total. */
    case std_dev = 107
    /** Use the sum function as the table total. */
    case sum = 109

    // public init(floatLiteral value: Double) { self = .number(value) }
    public init(integerLiteral value: Int) {
        if let function = TotalFunction(rawValue: UInt8(value)) {
            self = function
        } else {
            self = .none
        }
    }
}

//////////////////////////////////////////
/// MARK: Conditional Format

/** @brief Type definitions for conditional formats.
*
* Values used to set the "type" field of conditional format.
*/
public enum ConditionalFormatTypes: UInt8 {
    case none = 0

    /** The Cell type is the most common conditional formatting type. It is
    *  used when a format is applied to a cell based on a simple
    *  criterion.  */
    case cell

    /** The Text type is used to specify Excel's "Specific Text" style
    *  conditional format. */
    case text

    /** The Time Period type is used to specify Excel's "Dates Occurring"
    *  style conditional format. */
    case timePeriod

    /** The Average type is used to specify Excel's "Average" style
    *  conditional format. */
    case average

    /** The Duplicate type is used to highlight duplicate cells in a range. */
    case duplicate

    /** The Unique type is used to highlight unique cells in a range. */
    case unique

    /** The Top type is used to specify the top n values by number or
    *  percentage in a range. */
    case top

    /** The Bottom type is used to specify the bottom n values by number or
    *  percentage in a range. */
    case bottom

    /** The Blanks type is used to highlight blank cells in a range. */
    case blanks

    /** The No Blanks type is used to highlight non blank cells in a range. */
    case NoBlanks

    /** The Errors type is used to highlight error cells in a range. */
    case errors

    /** The No Errors type is used to highlight non error cells in a range. */
    case noErrors

    /** The Formula type is used to specify a conditional format based on a
    *  user defined formula. */
    case formula

    /** The 2 Color Scale type is used to specify Excel's "2 Color Scale"
    *  style conditional format. */
    case twoColorScale

    /** The 3 Color Scale type is used to specify Excel's "3 Color Scale"
    *  style conditional format. */
    case threeColorScale

    /** The Data Bar type is used to specify Excel's "Data Bar" style
    *  conditional format. */
    case dataBar

    /** The Icon Set type is used to specify a conditional format with a set
    *  of icons such as traffic lights or arrows. */
    case iconSets

    case last
}

/** @brief The criteria used in a conditional format.
*
* Criteria used to define how a conditional format works.
*/
public enum ConditionalCriteria: UInt8 {
    case none = 0

    /** Format cells equal to a value. */
    case equalTo

    /** Format cells not equal to a value. */
    case notEqualTo

    /** Format cells greater than a value. */
    case geraterThan

    /** Format cells less than a value. */
    case lessThan

    /** Format cells greater than or equal to a value. */
    case greaterThanOrEqualTo

    /** Format cells less than or equal to a value. */
    case leassThanOrEqualTo

    /** Format cells between two values. */
    case between

    /** Format cells that is not between two values. */
    case notBetween

    /** Format cells that contain the specified text. */
    case textContaining

    /** Format cells that don't contain the specified text. */
    case textNotContaining

    /** Format cells that begin with the specified text. */
    case textBeginsWith

    /** Format cells that end with the specified text. */
    case textEndsWith

    /** Format cells with a date of yesterday. */
    case timePeriodYesterday

    /** Format cells with a date of today. */
    case timePeriodToday

    /** Format cells with a date of tomorrow. */
    case timePeriodTomorrow

    /** Format cells with a date in the last 7 days. */
    case timePeriodLast7Days

    /** Format cells with a date in the last week. */
    case timePeriodLastWeek

    /** Format cells with a date in the current week. */
    case timePeriodThisWeek

    /** Format cells with a date in the next week. */
    case timePeriodNextWeek

    /** Format cells with a date in the last month. */
    case timePeriodLastMonth

    /** Format cells with a date in the current month. */
    case timePeriodThisMonth

    /** Format cells with a date in the next month. */
    case timePeriodNextMonth

    /** Format cells above the average for the range. */
    case averageAbove

    /** Format cells below the average for the range. */
    case averageBelow

    /** Format cells above or equal to the average for the range. */
    case averageAboveOrEqual

    /** Format cells below or equal to the average for the range. */
    case averageBelowOrEqual

    /** Format cells 1 standard deviation above the average for the range. */
    case average1StdDevBelow

    /** Format cells 1 standard deviation below the average for the range. */
    case average1StdDevAbove

    /** Format cells 2 standard deviation above the average for the range. */
    case average2StdDevAbove

    /** Format cells 2 standard deviation below the average for the range. */
    case average2StdDevBelow

    /** Format cells 3 standard deviation above the average for the range. */
    case average3StdDevAbove

    /** Format cells 3 standard deviation below the average for the range. */
    case average3StdDevBelow

    /** Format cells in the top of bottom percentage. */
    case topOrBottomPercent
}

public enum conditionalFormatBarDrection: UInt8 {

    /** Data bar direction is set by Excel based on the context of the data
     *  displayed. */
    case context

    /** Data bar direction is from right to left. */
    case rightToLeft

    /** Data bar direction is from left to right. */
    case leftToRight
};

public enum conditionalIconTypes: UInt8 {

    /** Icon style: 3 colored arrows showing up, sideways and down. */
    case arrowsColored3

    /** Icon style: 3 gray arrows showing up, sideways and down. */
    case arrorGray3

    /** Icon style: 3 colored flags in red, yellow and green. */
    case flags3

    /** Icon style: 3 traffic lights - rounded. */
    case trafficLightsUnrimmed3

    /** Icon style: 3 traffic lights with a rim - squarish. */
    case trafficLightsRimmed3

    /** Icon style: 3 colored shapes - a circle, triangle and diamond. */
    case signs3

    /** Icon style: 3 circled symbols with tick mark, exclamation and
     *  cross. */
    case symbolsCircled3

    /** Icon style: 4 symbols with tick mark, exclamation and cross. */
    case symbolsUncircled3

    /** Icon style: 4 colored arrows showing up, diagonal up, diagonal down
     *  and down. */
    case arrowsColored4

    /** Icon style: 4 gray arrows showing up, diagonal up, diagonal down and
     * down. */
    case arrowsGray4

    /** Icon style: 4 circles in 4 colors going from red to black. */
    case redToBlack4

    /** Icon style: 4 histogram ratings. */
    case ratings4

    /** Icon style: 4 traffic lights. */
    case trafficLights4

    /** Icon style: 5 colored arrows showing up, diagonal up, sideways,
     * diagonal down and down. */
    case arrowsColored5

    /** Icon style: 5 gray arrows showing up, diagonal up, sideways, diagonal
     *  down and down. */
    case arrowsGray5

    /** Icon style: 5 histogram ratings. */
    case ratings5

    /** Icon style: 5 quarters, from 0 to 4 quadrants filled. */
    case quarters5
};


public enum formatScripts: UInt8 {

    /** Superscript font */
    case superScript = 1

    /** Subscript font */
    case `subscript`
};


/** @brief The criteria used in autofilter rules.
 *
 * Criteria used to define an autofilter rule condition.
 */
public enum FilterCriteria: UInt8 {
    case none = 0

    /** Filter cells equal to a value. */
    case equalTo

    /** Filter cells not equal to a value. */
    case notEqualTo

    /** Filter cells greater than a value. */
    case greaterThan

    /** Filter cells less than a value. */
    case lessThan

    /** Filter cells greater than or equal to a value. */
    case greaterThanOrEqualTo

    /** Filter cells less than or equal to a value. */
    case lessThanOrEqualTo

    /** Filter cells that are blank. */
    case blanks

    /** Filter cells that are not blank. */
    case nonBlanks
};

/**
 * @brief And/or operator when using 2 filter rules.
 *
 * And/or operator conditions when using 2 filter rules with
 * worksheet_filter_column2(). In general LXW_FILTER_OR is used with
 * LXW_FILTER_CRITERIA_EQUAL_TO and LXW_FILTER_AND is used with the other
 * filter criteria.
 */
public enum FilterOperator: UInt8 {
    /** Logical "and" of 2 filter rules. */
    case and = 0

    /** Logical "or" of 2 filter rules. */
    case or
};

/* Internal filter types. */
public enum FilterType: UInt8 {
    case none = 0

    case single

    case and

    case or

    case stringList
};
