private void btnExportPdf_Click(object sender, EventArgs e)
{
    if (grdvExam.Rows.Count > 0)
    {
        SaveFileDialog sfd = new SaveFileDialog();
        sfd.Filter = "PDF (*.pdf)|*.pdf";
        sfd.FileName = "ExternalexamDetails.pdf";

        if (sfd.ShowDialog() == DialogResult.OK)
        {
            try
            {
                using (FileStream stream = new FileStream(sfd.FileName, FileMode.Create))
                {
                    Document pdfDoc = new Document(PageSize.A4.Rotate(), 10f, 20f, 20f, 10f);  // Set page orientation to horizontal
                    PdfWriter.GetInstance(pdfDoc, stream);
                    pdfDoc.Open();

                    PdfPTable pdfTable = new PdfPTable(9);  // Specify the number of columns you want

                    pdfTable.DefaultCell.Padding = 3;
                    pdfTable.WidthPercentage = 100;
                    pdfTable.HorizontalAlignment = Element.ALIGN_LEFT;

                    // Add specific columns to the PDF table
                    List<string> columnsToExport = new List<string> { "SerialNumber", "SkillName", "FullName", "TestName", "StaffName", "LabName", "TestDate", "TestTime", "TotalMarks" };

                    foreach (string columnName in columnsToExport)
                    {
                        if (grdvExam.Columns.Contains(columnName))
                        {
                            DataGridViewColumn column = grdvExam.Columns[columnName];
                            PdfPCell cell = new PdfPCell(new Phrase(column.HeaderText));
                            pdfTable.AddCell(cell);
                        }
                        else
                        {
                            // Handle the case where the specified column is not found
                            pdfTable.AddCell("");
                        }
                    }

                    // Add data from specific columns to the PDF table for rows with checked checkboxes
                    foreach (DataGridViewRow row in grdvExam.Rows)
                    {
                        // Assuming the checkbox column is named "Checkbox"
                        int checkBoxColumnIndex = grdvExam.Columns["Checkbox"].Index;

                        if (checkBoxColumnIndex != -1 && row.Cells[checkBoxColumnIndex].Value != null)
                        {
                            bool isChecked = (bool)row.Cells[checkBoxColumnIndex].Value;

                            if (isChecked)
                            {
                                foreach (string columnName in columnsToExport)
                                {
                                    if (grdvExam.Columns.Contains(columnName))
                                    {
                                        DataGridViewCell cell = row.Cells[columnName];
                                        if (cell.Value != null)
                                        {
                                            pdfTable.AddCell(cell.Value.ToString());
                                        }
                                        else
                                        {
                                            pdfTable.AddCell("");
                                        }
                                    }
                                }
                            }
                        }
                    }

                    pdfDoc.Add(pdfTable);
                    pdfDoc.Close();
                    stream.Close();

                    MessageBox.Show("External Mock Data Exported Successfully !!!", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, "Error");
            }
        }
    }
    else
    {
        MessageBox.Show("No Record To Export !!!", "Info");
    }

}
