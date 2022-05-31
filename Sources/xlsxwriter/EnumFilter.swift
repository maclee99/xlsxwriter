import Foundation


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
}

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
}

/* Internal filter types. */
public enum FilterType: UInt8 {
    case none = 0

    case single

    case and

    case or

    case stringList
}
