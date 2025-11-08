using System;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using ClosedXML.Excel;

namespace Finder
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            clearProgress(sender, e);

            // Update this path
            string filePath = @"C:\Hi\How\Are\You\Sample_Data.xlsx"; // Update this path
            // Update this path

            string searchValue = txtSearch.Text.Trim();

            // Remove leading zeros from the search value
            searchValue = searchValue.TrimStart('0');

            if (File.Exists(filePath))
            {
                using (var workbook = new XLWorkbook(filePath))
                {
                    var worksheet = workbook.Worksheet(1); // Get the first worksheet
                    bool itemFound = false;

                    // Search through rows
                    foreach (var row in worksheet.RowsUsed())
                    {
                        //assign upc and name values
                        var upc = row.Cell(1).GetString().Trim(); 
                        var productName = row.Cell(2).GetString().Trim();

                        // Remove leading zeros from the UPC for comparison
                        upc = upc.TrimStart('0');

                        // Debug output
                        Console.WriteLine($"Checking UPC: '{upc}' against '{searchValue}'");
                        Console.WriteLine($"Checking Product Name: '{productName}' against '{searchValue}'");

                        // Check for UPC match
                        if (upc.Equals(searchValue, StringComparison.OrdinalIgnoreCase))
                        {
                            // Populate labels with values from the third, fourth, and fifth columns
                            lblShelfChar.Text = row.Cell(3).GetString().ToUpper(); // Location
                            lblRowInt.Text = row.Cell(4).GetString(); // Category
                            lblColumnInt.Text = row.Cell(5).GetString(); // Price

                            itemFound = true;
                            break; // Exit loop after finding the first match
                        }

                        // If the product name is not empty, check for a match
                        if (!string.IsNullOrEmpty(productName) &&
                            productName.Equals(searchValue, StringComparison.OrdinalIgnoreCase))
                        {
                            // Populate labels with values from the third, fourth, and fifth columns
                            lblShelfChar.Text = row.Cell(3).GetString(); // Location
                            lblRowInt.Text = row.Cell(4).GetString(); // Category
                            lblColumnInt.Text = row.Cell(5).GetString(); // Price

                            itemFound = true;
                            break; // Exit loop after finding the first match
                        }
                    }

                   
                    string shelfValue = lblShelfChar.Text.Trim().ToUpper();
                    string rowValue = lblRowInt.Text.Trim();
                    string combinedValue = $"{shelfValue}:{rowValue}"; // Combine for switch case

                    switch (combinedValue)
                    {
                        case "A:1":
                            pbA1.Value = 100;
                            break;

                        case "A:2":
                            pbA2.Value = 100;
                            break;

                        case "A:3":
                            pbA3.Value = 100;
                            break;

                        case "B:1":
                            pbB1.Value = 100;
                            break;

                        case "B:2":
                            pbB2.Value = 100;
                            break;

                        case "B:3":
                            pbB3.Value = 100;
                            break;

                        case "B:4":
                            pbB4.Value = 100;
                            break;

                        case "B:5":
                            pbB5.Value = 100;
                            break;

                        case "B:6":
                            pbB6.Value = 100;
                            break;

                        case "B:":
                            pbB7.Value = 100;
                            break;

                        case "B:8":
                            pbB8.Value = 100;
                            break;

                        case "B:9":
                            pbB9.Value = 100;
                            break;

                        case "B:10":
                            pbB10.Value = 100;
                            break;

                        case "C:1":
                            pbC1.Value = 100;
                            break;

                        case "C:2":
                            pbC2.Value = 100;
                            break;

                        case "C:3":
                            pbC3.Value = 100;
                            break;

                        case "C:4":
                            pbC4.Value = 100;
                            break;

                        case "C:5":
                            pbC5.Value = 100;   
                            break;

                        case "C:6":
                            pbC6.Value = 100;
                            break;

                        case "C:7":
                            pbC7.Value = 100;
                            break;

                        case "D:1":
                            pbD1.Value = 100;
                            break;

                        case "D:2":
                            pbD2.Value = 100;
                            break;

                        case "D:3":
                            pbD3.Value = 100;
                            break;

                        case "D:4":
                            pbD4.Value = 100;
                            break;

                        case "D:5":
                            pbD5.Value = 100;
                            break;

                        case "D:6":
                            pbD6.Value = 100;
                            break;

                        case "D:7":
                            pbD7.Value = 100;
                            break;

                        case "E:1":
                            pbE1.Value = 100;
                            break;

                        case "E:2":
                            pbE2.Value = 100;
                            break;

                        case "E:3":
                            pbE3.Value = 100;
                            break;

                        case "E:4":
                            pbE4.Value = 100;
                            break;

                        case "E:5":
                            pbE5.Value = 100;
                            break;

                        case "E:6":
                            pbE6.Value = 100;
                            break;

                        case "E:7":
                            pbE7.Value = 100;
                            break;

                        case "F:1":
                            pbF1.Value = 100;
                            break;

                        case "F:2":
                            pbF2.Value = 100;
                            break;

                        case "F:3":
                            pbF3.Value = 100;
                            break;

                        case "F:4":
                            pbF4.Value = 100;
                            break;

                        case "F:5":
                            pbF5.Value = 100;
                            break;

                        case "F:6":
                            pbF6.Value = 100;
                            break;

                        case "F:7":
                            pbF7.Value = 100;
                            break;

                        case "G:1":
                            pbG1.Value = 100;
                            break;

                        case "G:2":
                            pbG2.Value = 100;
                            break;

                        case "G:3":
                            pbG3.Value = 100;
                            break;

                        case "G:4":
                            pbG4.Value = 100;
                            break;

                        case "G:5":
                            pbG5.Value = 100;
                            break;

                        case "G:6":
                            pbG6.Value = 100;
                            break;

                        case "G:7":
                            pbG7.Value = 100;
                            break;

                        case "H:1":
                            pbH1.Value = 100;
                            break;

                        case "H:2":
                            pbH2.Value = 100;
                            break;

                        case "H:3":
                            pbH3.Value = 100;
                            break;

                        case "H:4":
                            pbH4.Value = 100;
                            break;

                        case "H:5":
                            pbH5.Value = 100;
                            break;

                        case "H:6":
                            pbH6.Value = 100;
                            break;

                        case "H:7":
                            pbH7.Value = 100;
                            break;

                        case "I:1":
                            pbI1.Value = 100;
                            break;

                        case "I:2":
                            pbI2.Value = 100;
                            break;

                        case "I:3":
                            pbI3.Value = 100;
                            break;

                        case "I:4":
                            pbI4.Value = 100;
                            break;

                        case "I:5":
                            pbI5.Value = 100;
                            break;

                        case "I:6":
                            pbI6.Value = 100;
                            break;

                        case "I:7":
                            pbI7.Value = 100;
                            break;

                        case "J:1":
                            pbJ1.Value = 100;
                            break;

                        case "J:2":
                            pbJ2.Value = 100;
                            break;

                        case "J:3":
                            pbJ3.Value = 100;
                            break;

                        case "J:4":
                            pbJ4.Value = 100;
                            break;

                        case "J:5":
                            pbJ5.Value = 100;
                            break;

                        case "J:6":
                            pbJ6.Value = 100;
                            break;

                        case "J:7":
                            pbJ7.Value = 100;
                            break;
                    }


                    if (!itemFound)
                    {
                        MessageBox.Show("Item not found.");
                        // Clear labels if item is not found
                        lblShelfChar.Text = "";
                        lblRowInt.Text = "";
                        lblColumnInt.Text = "";
                    }
                }
            }
            else
            {
                MessageBox.Show("File not found.");
            }
        }

        private void btnRequest_Click(object sender, EventArgs e)
        {
            // Define the file path where you want to save the issues
            string txtFilePath = @"C:\Users\lildi\Downloads\issues.txt";

            // Get the issue text from the TextBox
            string issue = txtIssues.Text.Trim();

            // Check if the issue text is empty
            if (string.IsNullOrEmpty(issue))
            {
                MessageBox.Show("Please enter an issue.");
                return; // Exit if the issue is empty
            }

            try
            {
                // Append the issue to the text file
                using (StreamWriter writer = new StreamWriter(txtFilePath, true)) // Append mode
                {
                    writer.WriteLine($"{DateTime.Now}: {issue}"); // Save with a timestamp
                }

                MessageBox.Show("Your issue has been recorded.");
                txtIssues.Clear(); // Clear the TextBox after saving
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred while saving your issue: {ex.Message}");
            }
        }
        private void clearProgress(object sender, EventArgs e)
        {
            pbA1.Value = 0;
            pbA2.Value = 0;
            pbA3.Value = 0;
            pbB1.Value = 0;
            pbB2.Value = 0;
            pbB3.Value = 0;
            pbB4.Value = 0;
            pbB5.Value = 0;
            pbB6.Value = 0;
            pbB7.Value = 0;
            pbB8.Value = 0;
            pbB9.Value = 0;
            pbB10.Value = 0;
            pbC1.Value = 0;
            pbC2.Value = 0;
            pbC3.Value = 0;
            pbC4.Value = 0;
            pbC5.Value = 0;
            pbC6.Value = 0;
            pbC7.Value = 0;
            pbD1.Value = 0;
            pbD2.Value = 0;
            pbD3.Value = 0;
            pbD4.Value = 0;
            pbD5.Value = 0;
            pbD6.Value = 0;
            pbD7.Value = 0;
            pbE1.Value = 0;
            pbE2.Value = 0;
            pbE3.Value = 0;
            pbE4.Value = 0;
            pbE5.Value = 0;
            pbE6.Value = 0;
            pbE7.Value = 0;
            pbF1.Value = 0;
            pbF2.Value = 0;
            pbF3.Value = 0;
            pbF4.Value = 0;
            pbF5.Value = 0;
            pbF6.Value = 0;
            pbF7.Value = 0;
            pbG1.Value = 0;
            pbG2.Value = 0;
            pbG3.Value = 0;
            pbG4.Value = 0;
            pbG5.Value = 0;
            pbG6.Value = 0;
            pbG7.Value = 0;
            pbH1.Value = 0;
            pbH2.Value = 0;
            pbH3.Value = 0;
            pbH4.Value = 0;
            pbH5.Value = 0;
            pbH6.Value = 0;
            pbH7.Value = 0;
            pbI1.Value = 0;
            pbI2.Value = 0;
            pbI3.Value = 0;
            pbI4.Value = 0;
            pbI5.Value = 0;
            pbI6.Value = 0;
            pbI7.Value = 0;
            pbJ1.Value = 0;
            pbJ2.Value = 0;
            pbJ3.Value = 0;
            pbJ4.Value = 0;
            pbJ5.Value = 0;
            pbJ6.Value = 0;
            pbJ7.Value = 0;
        }

       
    }
}
