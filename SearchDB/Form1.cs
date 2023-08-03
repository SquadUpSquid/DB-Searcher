using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Common;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using ProgressBar = System.Windows.Forms.ProgressBar;

namespace SearchDB
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

        private void SearchButton_Click(object sender, EventArgs e)
        {
              SearchDB();
            //hide button when clicked
            searchButton.Visible = false;
        }
        private void Label1_Click(object sender, EventArgs e)
        {

        }

        private void SearchBox_TextChanged(object sender, EventArgs e)
        {

        }

        private void SaveFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }

        private async void SearchDB()
        {


            // Connection string to the SQL Server/DB
            string connectionString = "Data Source=10.101.93.61,3300;Initial Catalog='LSS PT Snapshot';Integrated Security=True;"; ///possibly make the 'Data Source' a user input???

            ///TEST CONNECTION///Data Source=R520SANDBOXSVR;Initial Catalog=ps_erdb;Integrated Security=True;///;Network Library=DBMSSOCN
            
            // File directory initialzation and parameters
            SaveFileDialog saveFileDialog1 = new SaveFileDialog
            {
                Title = "Save your file",
                DefaultExt = "xlsx",
                Filter = "Excel files (*.xlsx)|*.xlsx|CSV files (*.csv)|*.csv|All files (*.*)|*.*"
            };
            DialogResult result = saveFileDialog1.ShowDialog();


            // Opens a directory for the user to chose where the file is saved whenever they click the search button. 
            // It will open the dir and then it will search for the terms in the DB
            if (result == DialogResult.OK)
            {   //makes the filepath whatever the user made it in the file dir
                string outputFilePath = saveFileDialog1.FileName;

                // user input for searchTerms
                string searchTermsInput = searchBox.Text;
                List<string> searchTerms = new List<string>(searchTermsInput.Split(','));

                // Create an instance of Excel app and a new workbook
                Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
                Workbook workbook = excelApp.Workbooks.Add();


                try
                {
                    using (SqlConnection connection = new SqlConnection(connectionString))
                    {
                        await connection.OpenAsync();
                        // Get the list of table names out of the DB
                        var tables = await GetTableNamesAsync(connection);

                        //Progress bar setup
                        int totalTables = tables.Count; ///max should be totalTables but it's being weird 
                        int currentTableIndex = 0;
                        int progress = 0;
                        ProgressBar progressBar = new ProgressBar
                        {
                            Minimum = 0,
                            Maximum = 250,
                            Dock = DockStyle.Bottom,
                            Value = progress
                        };
                        Controls.Add(progressBar);
                        int tablesProcessed = 0;


                        // Iterate through each table
                        foreach (var tableName in tables)
                        {

                            //skip massive/redundant tables 
                            ///idk any other way to do it
                            if (tableName.IndexOf("tracking", StringComparison.OrdinalIgnoreCase) >= 0 )
                            {
                                continue;
                            };
                            if (tableName.IndexOf("track", StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                continue;
                            };
                            if (tableName.IndexOf("displaychangesstats", StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                continue;
                            };
                            if (tableName.IndexOf("definitionidentitytable", StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                continue;
                            };
                            if (tableName.IndexOf("instancemetadatachangestable", StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                continue;
                            };
                            if (tableName.IndexOf("identityownertable", StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                continue;
                            };
                            if (tableName.IndexOf("instancepromotedpropertiestable", StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                continue;
                            };
                            if (tableName.IndexOf("instancestable", StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                continue;
                            };
                            if (tableName.IndexOf("keystable", StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                continue;
                            };
                            if (tableName.IndexOf("lockownerstable", StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                continue;
                            };
                            if (tableName.IndexOf("runnableinstancestable", StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                continue;
                            };
                            if (tableName.IndexOf("servicedeploymentstable", StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                continue;
                            };
                            if (tableName.IndexOf("sqlworkflowinstancestoreversiontable", StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                continue;
                            };

                            // Query the 'current' table
                            string query = $"SELECT * FROM {tableName}";

                            //progress bar moving along with the program
                            currentTableIndex++;
                            progress = (currentTableIndex * 100 / totalTables);
                            progressBar.Value = progress;

                            // Execute that query and retrieve the data reader
                            using (SqlCommand command = new SqlCommand(query, connection))
                            {
                                SqlDataReader reader = command.ExecuteReader();

                                // Load the data into a DataTable
                                System.Data.DataTable dataTable = new System.Data.DataTable();
                                dataTable.Load(reader);

                                // Filter the table based on searchTerms
                                System.Data.DataTable filteredTable = FilterTable(dataTable, searchTerms);
                                //Worksheet ws1 = workbook.Sheets.Add();

                                if (filteredTable.Rows.Count > 0)
                                {
                                    // Add new worksheet to the workbook
                                    Worksheet ws2 = workbook.Sheets.Add();
                                    ws2.Name = GetValidWorksheetName(tableName);

                                    // Write the header row to the worksheet (i.e. column names)
                                    for (int i = 0; i < filteredTable.Columns.Count; i++)
                                    {
                                        //Wait(1);
                                        ws2.Cells[1, i + 1].Value = filteredTable.Columns[i].ColumnName;
                                    }


                                    if (filteredTable.Rows.Count > 0)
                                    {
                                        // Write the search results (row data) to the worksheet 
                                        for (int row = 0; row < filteredTable.Rows.Count; row++)
                                        {
                                            for (int col = 0; col < filteredTable.Columns.Count; col++)
                                            {
                                                //Wait(1);
                                                if (filteredTable.Rows[row][col].ToString().Contains("0x"))
                                                {
                                                    MessageBox.Show(filteredTable.Rows[row][col].ToString());
                                                    ws2.Cells[row + 2, col + 1].Value = " ";
                                                }
                                                else
                                                {
                                                    ws2.Cells[row + 2, col + 1].Value = filteredTable.Rows[row][col].ToString();
                                                }
                                            }
                                        }
                                    }
                                    // Auto-firt the columns in the worksheet
                                    Range usedRange = ws2.UsedRange;
                                    usedRange.Columns.AutoFit();
                                }

                                reader.Close();
                            }
                            tablesProcessed++;
                                //Save the file at checkpoint intervals
                                if (tablesProcessed % CheckpointInterval == 0)
                                {
                                    SaveWorkbook(workbook, outputFilePath);
                                    ///workbook.SaveAs(outputFilePath);
                                }
                        }
                        Wait(10);
                        // Saves the workbook in the filePath the user chose
                        workbook.SaveAs(outputFilePath);
                        workbook.Close();
                        excelApp.Quit();
                        
                        MessageBox.Show("File saved to your destination.");
                        //hides the progress bar when the program is complete
                        Controls.Remove(progressBar);
                        searchButton.Visible = true;
                    }
                }
                catch (Exception ex)
                {
                    workbook.Close();
                    excelApp.Quit();
                    //error message
                    MessageBox.Show("An error occurred: " + ex.Message + Environment.NewLine + ex.TargetSite
                        + Environment.NewLine + ex.HResult + Environment.NewLine + ex.Data);
                }
                finally
                {
                    //helps excel close after it saves
                    ReleaseCOMObjects(workbook);
                    ReleaseCOMObjects(excelApp);
                }
            }
        }

        //Checkpoint interval variables 
        private const int CheckpointInterval = 1;

        private void SaveWorkbook(Workbook workbook, string outputFilePath)
        {
            try
            {
                workbook.Save();
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occurred while saving the workbook: " + ex.Message, "Error");
            }
        }



        // Filters the tables based on the search terms
        private System.Data.DataTable FilterTable(System.Data.DataTable table, List<string> searchTerms)
        {
            //clones the table so it can search through it 
            System.Data.DataTable filteredTable = table.Clone();
            foreach (DataRow row in table.Rows)
            {

                bool shouldAddRow = false;
                foreach (DataColumn column in table.Columns)
                {
                    //these are known column names that have a massive value and freak out the program
                    if (column.ColumnName.Equals("ChartData" + "XMLData", StringComparison.OrdinalIgnoreCase))
                    {
                        continue;// Skip the ChartData column
                    } 

                    string cellValue = row[column].ToString();

                    //Check if any search term matches the cell value using wildcards
                    if (searchTerms.Any(searchTerm => IsWildcardMatch(cellValue, searchTerm)))
                    {
                        shouldAddRow = true;
                        break;
                    }
                    if (shouldAddRow)
                    {
                        break;
                    }
                }

                try
                {
                    //print the row if it contains search term
                    if (shouldAddRow)
                    {
                        filteredTable.Rows.Add(row.ItemArray);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("An error occurred: " + ex.Message + Environment.NewLine + ex.TargetSite
                        + Environment.NewLine + ex.HResult + Environment.NewLine + ex.Data);
                }
            }
            return filteredTable;
        }



        private bool IsWildcardMatch(string value, string searchTerm) ///new line
        {
            //Convert wildcard search term to a regular expression pattern
            string pattern = WildcardToRegex(searchTerm);
            // check if the value matches the regular expression pattern
            return Regex.IsMatch(value, pattern, RegexOptions.IgnoreCase);
        }

        private string WildcardToRegex(string wildcard) ///new line
        {
            //convert wildcard to regular expression pattern
            return "^" + Regex.Escape(wildcard)
                                .Replace("\\*", ".*") //replace '*' with '.*' to match any number of characters
                                .Replace("\\?", ".") //replace '?' with '.' to match any single character
                        + "$";
        }




        // Looks for the list of tableNames in the database
        private async Task<List<string>> GetTableNamesAsync(SqlConnection connection)
        {
            List<string> tables = new List<string>();
            //command to grab all of the table names in the database
            using (SqlCommand command = new SqlCommand("SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE' " +
                    "AND TABLE_CATALOG = 'LSS PT Snapshot' ORDER BY TABLE_NAME ASC", connection))
            {
                SqlDataReader reader = await command.ExecuteReaderAsync();

                while (await reader.ReadAsync())
                {
                    string tableName = reader.GetString(0);
                    tables.Add(tableName);
                }

                reader.Close();
            }

            return tables;
        }


        private string GetValidWorksheetName(string tableName)
        {
            // Make sure the table name isn't to long for excel
            //if the worksheet names are over a 31 char then it causes an error
            if (tableName.Length > 30)
            {
                return tableName.Substring(0, 30);
            }
            else
            {
                return tableName;
            }
        }


        // Releases COM objs 
        static void ReleaseCOMObjects(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Error releasing COM object: " + ex.Message);
            }
            finally
            {
                GC.Collect();
            }
        }

        //tells the program to wait wherever you call the function. 
        //helps slow certain parts down if running to efficently
        public void Wait(int Time)
        {
            Thread thread = new Thread(delegate ()
            {
                System.Threading.Thread.Sleep(Time);
            });
            thread.Start();
            while (thread.IsAlive)
                System.Windows.Forms.Application.DoEvents();
        }

        private void helpButton_Click_1(object sender, EventArgs e)
        {
            //show help messages
            string helpMessage = "Welcome to the SearchDB Help Menu!!\n\n" +
                                "1. Enter the search term(s) in the search Box.\n" +
                                "2. When using multiple search terms, deliminate them with a comma(no space after).\n" +
                                "3. Click the Search button to start the search.\n" +
                                "4. Choose a location to save the search results.\n" +
                                "5. Wait for the search to complete and the file to be saved.\n" +
                                "6. You can monitor the progress using the progress bar.\n\n\n" +
                                "NOTE: \n-- If you see the program as 'Not Respinding', it is still working do not close program.\n" +
                                "-- If the progress bar isn't moving, it is still working do not close the program.\n" + 
                                "-- If a message box appears asking to replace 'Book1', click yes to replace.\n";

            MessageBox.Show(helpMessage, "Help", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

    }   
    
}

