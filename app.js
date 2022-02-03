const readXlsxFile = require('read-excel-file/node')
const cloneDeep = require('clone-deep')
const XLSXColumn = require('xlsx-column')

// Read the Excel file
readXlsxFile('src/basic-accessibility-checks.xlsx')
  // Then loop through each row
  .then(rows => {
    const data = {
      pages: []
    }

    // Generate pages
    // Loop through all cells in row 1
    for (let i = 0; i < rows[0].length; i++) {
      const name = rows[0][i]
      const colNo = i + 1
      // Check for defined pages (default Page 1 etc get ignored)
      if (name !== null && colNo > 2 && !name.match(/Page /)) {
        // Get the column ID such as A or AA
        const column = new XLSXColumn(colNo)
        // Create a page object in the pages array
        data.pages.push({
          colNo,
          name,
          colID: String(column)
        })
      }
    }

    // Generate test categories
    // Create a holding variable for the category name
    let currentTestCategoryInLoop = ''
    // Create an empty array to hold the tests
    const tests = []
    // Loop through all the rows
    for (let i = 0; i < rows.length; i++) {
      // Setup variables for current iteration
      const testName = rows[i][0]
      const rowNo = i + 1
      let testIsCategory = false
      // Row 1 is skipped as it's just headings
      if (rowNo > 1) {
        // If there is a number in the test name, then it's a category name
        if (testName.match(/\d/)) {
          // Set the current category to the data in this current cell
          // and, set the flag to say this is a test category
          currentTestCategoryInLoop = testName
          testIsCategory = true
        }
        // If it's not a test category
        if (!testIsCategory) {
          // Push the test into the tests array
          tests.push({
            name: testName,
            category: currentTestCategoryInLoop,
            wcag: rows[i][1], // Get the WCAG criterion from the cell to the right
            colNo: 1,
            rowNo
          })
        }
      }
    }

    // Add testCategoriesToPages
    // Loop through each page
    data.pages.forEach(page => {
      // Clone the tests array to avoid persisting the wrong data
      const pageTests = cloneDeep(tests)
      // Loop through each test
      pageTests.forEach(test => {
        // Add the column number for the page
        test.colNo = page.colNo
      })
      // Push the tests for the page into the current page
      page.tests = pageTests
    })

    // Get test results
    // Loop through each page
    data.pages.forEach(page => {
      // Loop through each test for the page
      page.tests.forEach(test => {
        // Loop through each row
        for (let i = 0; i < rows.length; i++) {
          const rowNo = i + 1
          // Row 1 is skipped as it's just headings
          if (rowNo > 1) {
            const row = rows[i]
            // If the cells are empty, make them empty strings rather than null
            const status = row[test.colNo - 1] || ''
            const observations = row[test.colNo] || ''
            // The headings rows are ignored
            if (status !== 'Status' && observations !== 'Observations') {
              const column = new XLSXColumn(test.colNo)
              test.colID = String(column) // Get the column ID such as A or AA
              test.status = status
              test.observations = observations
            }
          }
        }
      })
    })
    console.log(data.pages[0])
  })
