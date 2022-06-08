import XCTest

#if !canImport(ObjectiveC)
public func allTests() -> [XCTestCaseEntry] {
    return [
        testCase(xlsxwriterTests.allTests),
        testCase(outlineTests.allTests),
        testCase(macroTests.allTests),
    ]
}
#endif
