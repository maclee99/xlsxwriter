import Foundation

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
}

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
}


public enum formatScripts: UInt8 {

    /** Superscript font */
    case superScript = 1

    /** Subscript font */
    case `subscript`
}