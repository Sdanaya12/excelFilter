Feature: Test Excel
  A user of the work team
  Enter an Excel file
  The system filters the document

  @smokeTest

  Scenario: interacting with the document
    Given A user of the work team selects the Excel document
    When The system performs the reading of the document
    Then You should see the document properly filtered


