# TODO: Fix Date Filtering in PdfReport.cshtml

- [x] Modify loadData function to remove fromDate and toDate from AJAX data (since controller doesn't filter by dates)
- [x] Update populateTable function to accept fromDate and toDate parameters
- [x] Implement client-side filtering logic in populateTable to filter data based on item.Created (Monitoring DateTime) using fromDate and toDate
- [x] Update calls to populateTable in loadData and other places to pass the date parameters
- [x] Add date parsing logic to compare dates correctly (assuming YYYY-MM-DD format)

# TODO: Add Download Excel Button for Grid Data

- [x] Include xlsx library script in the head section
- [x] Add "Download Excel" button next to the Filter button in the UI
- [x] Add global variable currentData to store filtered data
- [x] Update populateTable to set currentData after filtering
- [x] Implement downloadExcel function with header styling (bold, blue background, center alignment) and export data to Excel
