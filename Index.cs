

var importFile = main.DisplayFile("Select File to Be Processed", null, ".txt");

            string envNum = "";

            bool envSelected = true;

            #region dateSelect

            DateTime dtpStartDate = new DateTime(1000, 1, 1);

            DateTime dtpEndDate = new DateTime(1000, 1, 1);

            var today = DateTime.Today;

            var month = new DateTime(today.Year, today.Month, 1);

            var first = month.AddMonths(-1);

            var last = month.AddDays(-1);

 

            Form frm = new Form();

            frm.StartPosition = FormStartPosition.CenterScreen;

            frm.Height = 160;

            frm.Width = 300;

 

            Label lbl = new Label();

            lbl.Text = "Select Date Range";

            lbl.AutoSize = true;

            lbl.Width = 165;

            lbl.Height = 13;

            lbl.Font = new Font(FontFamily.GenericSansSerif, 8.25F, FontStyle.Bold);

            lbl.Location = new Point(66, 10);

 

            Label lblStartDate = new Label();

            lblStartDate.Text = "Start Date:";

            lblStartDate.AutoSize = true;

            lblStartDate.Location = new Point(13, 26);

 

            Label lblEndDate = new Label();

            lblEndDate.Text = "End Date:";

            lblEndDate.AutoSize = true;

            lblEndDate.Location = new Point(165, 26);

 

            DateTimePicker dtpStart = new DateTimePicker();

            dtpStart.Location = new Point(12, 45);

            dtpStart.Format = DateTimePickerFormat.Short;

            dtpStart.Width = 100;

            dtpStart.Value = first;

 

            DateTimePicker dtpEnd = new DateTimePicker();

            dtpEnd.Location = new Point(165, 45);

            dtpEnd.Format = DateTimePickerFormat.Short;

            dtpEnd.Width = 100;

            dtpEnd.Value = last;

 

            Button btn = new Button(); // set start date and end date

            btn.Text = "Submit";

            btn.Location = new Point(100, 80);

            btn.Click += (object sender, EventArgs e) =>

            {

                dtpStartDate = dtpStart.Value;

                dtpEndDate = dtpEnd.Value;

                frm.Close();

            };

 

            frm.Controls.Add(lbl);

            frm.Controls.Add(lblStartDate);

            frm.Controls.Add(lblEndDate);

            frm.Controls.Add(dtpStart);

            frm.Controls.Add(dtpEnd);

            frm.Controls.Add(btn);

            frm.ShowDialog();

            frm.Close();

 

            var startDate = dtpStartDate.ToString().Substring(0, dtpStartDate.ToString().Length - 12);

            var endDate = dtpEndDate.ToString().Substring(0, dtpEndDate.ToString().Length - 12);

            var finalStartDate = DateTime.Now;

            var finalEndDate = DateTime.Now;

            var checkdate = DateTime.Now;

            DateTime.TryParse(startDate, out finalStartDate);

            DateTime.TryParse(endDate, out finalEndDate);

            

            //Display Date range for user

            MessageBox.Show($"Only Records between {startDate} - {endDate} Will be Shown");

             #endregion

 

                        //Check to make sure user selected a file

            if (importFile != null && importFile != "")

            {

                try

                {

                    var filePath = Path.GetDirectoryName(importFile) + Path.DirectorySeparatorChar;

                    //create excel workbook

                    ExcelPackage xlpack = new ExcelPackage();

                    //create sheet and select sheet

                    xlpack.Workbook.Worksheets.Add("Sheet1");

                    var ws = xlpack.Workbook.Worksheets["Sheet1"];

                    int row = 1;

 

                    //read file into memory

                    var data = File.ReadAllLines(importFile).ToList();

                    //data.RemoveAt(0);

                    //skip first line which is header and order by call field

                    var newSortData = data.Skip(1).OrderBy(p => p.Split(',')[5]);

 

                    //hard code after removing so records can be sorted

                    ws.Cells[row, 1].Value = "FIELD 1";

                    ws.Cells[row, 2].Value = "FIELD 2";

                    ws.Cells[row, 3].Value = "FIELD 3";

                    ws.Cells[row, 4].Value = "FIELD 4";

                    ws.Cells[row, 5].Value = "FIELD 5";

                    ws.Cells[row, 6].Value = "FIELD 6";

                    ws.Cells[row, 7].Value = "FIELD 7";

                    ws.Cells[row, 8].Value = "FIELD 8";

                    ws.Cells[row, 9].Value = "FIELD 9";

                    ++row;

                    //sort records out that don't meet date criteria and create new list based on remaining records

                    var filteredList = newSortData.Where(x => Convert.ToDateTime(x.Split(',')[5]) >= finalStartDate && Convert.ToDateTime(x.Split(',')[5]) <= finalEndDate).ToList();

                    //loop over each record separating fields by comma

                    

                    foreach (var i in filteredList)

                    {

                        var pieces = i.Split(',');

                    

                            TextInfo textInfo = new CultureInfo("en-US", false).TextInfo;

 

                                for (int col = 1; col <= pieces.Length; col++)

                                {

                                    //handling each columns formatting

                                    switch (col)

                                    {

                                        //formatting all names to have the same case.

                                        case 2:

                                            ws.Cells[row, col].Value = textInfo.ToTitleCase(pieces[col - 1].ToLower());

                                            break;

                                    //formatting ints from duration of call time in seconds

                                        case 5:

                                            ws.Cells[row, col].Value = Convert.ToInt32(pieces[col - 1]);

                                            break;

                                    //formatting balances

                                        case 7:

                                            ws.Cells[row, col].Value = Convert.ToDecimal(pieces[col - 1]);

                                            ws.Cells[row, col].Style.Numberformat.Format = "$#,##0.00";

                                            

                                            break;

                                        default:

                                            ws.Cells[row, col].Value = pieces[col - 1];

                                            break;

                                    }

                                }

                                ++row;

                            }

                        

                    ws.Cells.AutoFitColumns();

                    xlpack.SaveAs(new FileInfo(($"{filePath}MONTHLYCALLLOG_{DateTime.Now.ToString("MMddyyyy")}.xlsx")));

                    xlpack.Dispose();

                    main.txtResults.Text += "Processing Complete\r\n";

                    string strFunction = System.Reflection.MethodBase.GetCurrentMethod().Name;

                    main.UpdateLastRun(strFunction, DateTime.Now);

                }

                catch (Exception theException)

                {

                    String errorMessage;

                    errorMessage = "Error: ";

                    errorMessage = String.Concat(errorMessage, theException.Message);

                    errorMessage = String.Concat(errorMessage, " Line: ");

                    errorMessage = String.Concat(errorMessage, theException.Source);

                    MessageBox.Show(errorMessage, "Error");

                }//end of catch

            }

            else

            {

                MessageBox.Show("You Have not Selected a File to Be Processed");

            }

        }