import Foundation

// Data validation types
public enum ValidationTypes: UInt8 {
    case none = 0

    /** Restrict cell input to whole/integer numbers only. */
    case integer

    /** Restrict cell input to whole/integer numbers only, using a cell
     *  reference. */
    case integerFormula

    /** Restrict cell input to decimal numbers only. */
    case decimal

    /** Restrict cell input to decimal numbers only, using a cell
     * reference. */
    case decimalFormula

    /** Restrict cell input to a list of strings in a dropdown. */
    case list

    /** Restrict cell input to a list of strings in a dropdown, using a
     * cell range. */
    case listFormula

    /** Restrict cell input to date values only, using a lxw_datetime type. */
    case date

    /** Restrict cell input to date values only, using a cell reference. */
    case dateFormula

    /* Restrict cell input to date values only, as a serial number.
     * Undocumented. */
    case dateNumber

    /** Restrict cell input to time values only, using a lxw_datetime type. */
    case time

    /** Restrict cell input to time values only, using a cell reference. */
    case timeFormula

    /* Restrict cell input to time values only, as a serial number.
     * Undocumented. */
    case timeNumber

    /** Restrict cell input to strings of defined length, using a cell
     * reference. */
    case length

    /** Restrict cell input to strings of defined length, using a cell
     * reference. */
    case lengthFormula

    /** Restrict cell to input controlled by a custom formula that returns
     * `TRUE/FALSE`. */
    case customFormula

    /** Allow any type of input. Mainly only useful for pop-up messages. */
    case any
}

/** Data validation criteria uses to control the selection of data. */
public enum ValidationCriteria: UInt8 {
    case none = 0

    /** Select data between two values. */
    case between

    /** Select data that is not between two values. */
    case notBetween

    /** Select data equal to a value. */
    case equalTo

    /** Select data not equal to a value. */
    case notEqualTo

    /** Select data greater than a value. */
    case greaterThan

    /** Select data less than a value. */
    case lessThan

    /** Select data greater than or equal to a value. */
    case greaterThanOrEqualTo

    /** Select data less than or equal to a value. */
    case lessThanOrEqualTo
}

/** Data validation error types for pop-up messages. */
public enum ValidationErrorTypes: UInt8 {
    /** Show a "Stop" data validation pop-up message. This is the default. */
    case stop = 0

    /** Show an "Error" data validation pop-up message. */
    case warning

    /** Show an "Information" data validation pop-up message. */
    case information
}