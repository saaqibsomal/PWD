# TODO: Fix Date Filtering in PdfReport.cshtml

- [x] Modify loadData function to remove fromDate and toDate from AJAX data (since controller doesn't filter by dates)
- [x] Update populateTable function to accept fromDate and toDate parameters
- [x] Implement client-side filtering logic in populateTable to filter data based on item.Created (Monitoring DateTime) using fromDate and toDate
- [x] Update calls to populateTable in loadData and other places to pass the date parameters
- [x] Add date parsing logic to compare dates correctly (assuming YYYY-MM-DD format)
