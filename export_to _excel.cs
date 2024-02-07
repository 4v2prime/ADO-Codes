public void Export2ExcelAllExam()
{
    if (grdvExam.Rows.Count > 0)
    {
        SaveFileDialog sfd = new SaveFileDialog();
        sfd.Filter = "Excel Workbook (*.xlsx)|*.xlsx";
        sfd.FileName = "ExternalexamDetails.xlsx";

        if (sfd.ShowDialog() == DialogResult.OK)
        {
            // Create Excel application object and workbook
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook excelWorkbook = null;

            try
            {
                // Create a new workbook
                excelWorkbook = excelApp.Workbooks.Add();

                // Create a new worksheet
                Microsoft.Office.Interop.Excel.Worksheet excelWorksheet = excelWorkbook.Worksheets.Add();
                /// Add title with font size 16 and bold style
                excelWorksheet.Cells[1, 1] = "External Exam Details";
                excelWorksheet.Cells[1, 1].Font.Size = 16;
                excelWorksheet.Cells[1, 1].Font.Bold = true;
                // Merge and center the title row
                excelWorksheet.Range["A1", "F1"].Merge();
                excelWorksheet.Range["A1", "F1"].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                // Skip a line
                excelWorksheet.Cells[2, 1] = ""; // Empty cell to add a blank row
                excelWorksheet.Rows[2].RowHeight = 20; // Increase row height for spacing

                List<string> columnsToExport = new List<string> { };
                columnsToExport = new List<string> { "SerialNumber", "SkillName", "TestName", "RegisterDate", "TotalMarks", "Duration" };

                // Write headers to the third row in Excel (after title and blank row)
                for (int i = 0; i < columnsToExport.Count; i++)
                {
                    excelWorksheet.Cells[3, i + 1] = columnsToExport[i];
                }

                // Iterate through checked checkbox rows and write visible data to Excel
                int rowNumber = 4; // Start writing data from the fourth row
                foreach (DataGridViewRow row in grdvExam.Rows)
                {
                    // Assuming the checkbox column is named "Checkbox"
                    int checkBoxColumnIndex = grdvExam.Columns["Checkbox"].Index;

                    if (checkBoxColumnIndex != -1 && row.Cells[checkBoxColumnIndex].Value != null)
                    {
                        bool isChecked = (bool)row.Cells[checkBoxColumnIndex].Value;

                        if (isChecked)
                        {
                            // Get visible cell values for the current row for the specified columns
                            List<object> rowValues = new List<object>();
                            foreach (string columnName in columnsToExport)
                            {
                                DataGridViewColumn column = grdvExam.Columns[columnName];
                                if (column != null && column.Visible)
                                {
                                    int cellIndex = column.Index;
                                    object cellValue = (row.Cells[cellIndex].Value != null) ? row.Cells[cellIndex].Value.ToString() : "";
                                    rowValues.Add(cellValue);
                                }
                                else
                                {
                                    // Handle the case where the specified column is not found or not visible
                                    rowValues.Add("");
                                }
                            }

                            // Write visible cell values to the current row in Excel
                            for (int i = 0; i < rowValues.Count; i++)
                            {
                                excelWorksheet.Cells[rowNumber, i + 1] = rowValues[i];
                            }

                            rowNumber++;
                        }
                    }
                }

                // AutoFit columns
                excelWorksheet.Columns.AutoFit();

                // Save Excel file
                excelWorkbook.SaveAs(sfd.FileName);

                // Open Excel file
                Process.Start(sfd.FileName);
                MessageBox.Show("External Mock Data Opened in Excel Successfully !!!", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error exporting to Excel: " + ex.Message);
            }
            finally
            {
                // Release COM objects
                if (excelWorkbook != null)
                {
                    Marshal.ReleaseComObject(excelWorkbook);
                }
                if (excelApp != null)
                {
                    excelApp.Quit();
                    Marshal.ReleaseComObject(excelApp);
                }
            }
        }
    }
    else
    {
        MessageBox.Show("No Record To Export !!!", "Info");
    }
}
