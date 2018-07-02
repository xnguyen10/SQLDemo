using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Microsoft.Office.Interop;
using Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using Microsoft.Office.Interop.Word;

namespace SQLDemo
{
    public partial class BizContacts : Form
    {
        string connString = @"Data Source=M4800\SQLEXPRESS;Initial Catalog=AddressBook;Integrated Security=True;Connect Timeout=30;Encrypt=False;TrustServerCertificate=True;ApplicationIntent=ReadWrite;MultiSubnetFailover=False";

        SqlDataAdapter dataAdapter;
        System.Data.DataTable table;
        SqlCommandBuilder commandBuilder;  //declare new SqlCommandBuilder
        SqlConnection conn; //declares variable to hold a sql connection
        string selectionStatement = "Select * from BizContacts";

        public BizContacts()
        {
            InitializeComponent();
        }

        private void BizContacts_Load(object sender, EventArgs e)
        {
            cboSearch.SelectedIndex = 0;

            dataGridView1.DataSource = bindingSource1;

            //Line below calls a method called GetData
            //The argument is a string that represents an sql query
            //select * from BizContacts means elect all the data from the biz contacts table
            GetData(selectionStatement);
        }

        private void GetData(string selectCommand)
        {
            try
            {
                dataAdapter = new SqlDataAdapter(selectCommand, connString);  //pass in the select command and the connection string
                table = new System.Data.DataTable(); //make a new data table object
                table.Locale = System.Globalization.CultureInfo.InvariantCulture;
                dataAdapter.Fill(table);  //fill the data table
                bindingSource1.DataSource = table; //set the data source on the binding source to the table
                dataGridView1.Columns[0].ReadOnly = true; //helps ID field from being changed.
            }
            catch (SqlException ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            SqlCommand command; //declares a new sql command object

            //field names in the table
            string insert = @"insert into BizContacts(Date_Added, Company, Website, Title, First_Name, Last_Name, Address, 
                                City, State, Postal_Code, Mobile, Notes, Image)
                            values(@Date_Added, @Company, @Website, @Title, @First_Name, @Last_Name, @Address, 
                                @City, @State, @Postal_Code, @Mobile, @Notes, @Image)"; //parameter names

            using (conn = new SqlConnection(connString)) //using allows disposing of low level resources
            {
                try
                {
                    conn.Open(); //open the connection
                    command = new SqlCommand(insert, conn); //create the new sql command object
                    command.Parameters.AddWithValue(@"Date_Added", dateTimePicker1.Value.Date); //read value from form and save to table
                    command.Parameters.AddWithValue(@"Company", txtCompany.Text); //read value from form and save to table
                    command.Parameters.AddWithValue(@"Website", txtWebsite.Text); //read value from form and save to table
                    command.Parameters.AddWithValue(@"Title", txtTitle.Text); //read value from form and save to table
                    command.Parameters.AddWithValue(@"First_Name", txtFirstName.Text); //read value from form and save to table
                    command.Parameters.AddWithValue(@"Last_Name", txtLastName.Text); //read value from form and save to table
                    command.Parameters.AddWithValue(@"Address", txtAddress.Text); //read value from form and save to table
                    command.Parameters.AddWithValue(@"City", txtCity.Text); //read value from form and save to table
                    command.Parameters.AddWithValue(@"State", txtState.Text); //read value from form and save to table
                    command.Parameters.AddWithValue(@"Postal_Code", txtPostalCode.Text); //read value from form and save to table
                    command.Parameters.AddWithValue(@"Mobile", txtMobile.Text); //read value from form and save to table
                    command.Parameters.AddWithValue(@"Notes", txtNotes.Text); //read value from form and save to table

                    if (dlgOpenImage.FileName != "") //check whether file naem is not empty
                    {
                        command.Parameters.AddWithValue(@"Image", File.ReadAllBytes(dlgOpenImage.FileName)); //convert image to bytes for sql server
                    }
                    else
                    {
                        command.Parameters.Add("@image", SqlDbType.VarBinary).Value = DBNull.Value; //save null  to database
                    }
    
                    command.ExecuteNonQuery(); //push stuff into the table
                }

                catch (Exception ex)

                {
                    MessageBox.Show(ex.Message);
                }         
            }

            GetData(selectionStatement);
            dataGridView1.Update(); //redraws the data grid view so the new record is visible on the bottom

            txtTitle.Clear();
            txtFirstName.Clear();
            txtLastName.Clear();
            txtAddress.Clear();
            txtCity.Clear();
            txtState.Clear();
            txtPostalCode.Clear();
            txtWebsite.Clear();
            txtNotes.Clear();
            txtMobile.Clear();
            txtCompany.Clear();
        }

        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            commandBuilder = new SqlCommandBuilder(dataAdapter);
            dataAdapter.UpdateCommand = commandBuilder.GetUpdateCommand(); //get the update command.
            try
            {
                bindingSource1.EndEdit(); //updates the table that is in memory in our program
                dataAdapter.Update(table); //acctually updates the database.
                MessageBox.Show("Update successful.");
            }

            catch (Exception ex)

            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            DataGridViewRow row = dataGridView1.CurrentCell.OwningRow; //grab a reference to the current row

            string value = row.Cells["ID"].Value.ToString(); //grab the value from the id field of the selected record
            string fname = row.Cells["First_Name"].Value.ToString(); //grab the value from the first name field of the selected record
            string lname = row.Cells["Last_Name"].Value.ToString(); //grab the value from the last name field of the selected record

            DialogResult result = MessageBox.Show("Confirm deletion " + fname + " " + lname + ", record " + value, "Message", 
                MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            string deleteState = @"Delete from BizContacts where  id = " + "'" + value + "'";

            if (result==DialogResult.Yes)
            {
                using (conn = new SqlConnection(connString))
                {
                    try
                    {
                        conn.Open(); //try to open connection
                        SqlCommand comm = new SqlCommand(deleteState, conn);
                        comm.ExecuteNonQuery(); //this line causes the deletion to run.
                        GetData(selectionStatement);
                        dataGridView1.Update();
                    }

                    catch (Exception ex)

                    {

                        MessageBox.Show(ex.Message);
                    }
                }
            }
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            switch(cboSearch.SelectedItem.ToString()) //present because we have a combobox
            {
                case "First Name":
                    GetData("select * from bizcontacts where lower(First_Name) like '%" + txtSearch.Text.ToLower() + "%'");
                    break;
                case "Last Name":
                    GetData("select * from bizcontacts where lower(last_Name) like '%" + txtSearch.Text.ToLower() + "%'");
                    break;
                case "Company":
                    GetData("select * from bizcontacts where lower(company) like '%" + txtSearch.Text.ToLower() + "%'");
                    break;
            }
        }

        private void btnGetImage_Click(object sender, EventArgs e)
        {
            if(dlgOpenImage.ShowDialog() == DialogResult.OK)//show box for selecting image from drive               
                pictureBox1.Load(dlgOpenImage.FileName); //loads image from drive using file name property of dialogbox
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            Form frm = new Form(); //make a new form
            try
            {
                frm.BackgroundImage = pictureBox1.Image; //set background image of new, preview form of image
                frm.Size = pictureBox1.Image.Size; //set the size of the form to the size of the image is wholly visible
                frm.Show(); //show form with image.
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\r\n\r\nImage not detected. \r\nPlease add Image.", "Error Detected");
            }          
        }

        private void btnEdit_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Edit", "Please edit directly in the Output Grid.");
        }

        private void btnExportOpen_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel._Application excel = new Microsoft.Office.Interop.Excel.Application(); //make a new excel object
            _Workbook workbook = excel.Workbooks.Add(Type.Missing); //make a workbook
            _Worksheet worksheet = null; //make a worksheet and for now set it to null

            try
            {
                worksheet = workbook.ActiveSheet; //set active sheet
                worksheet.Name = "Business Contacts";
                //because both datagrids and excel sheeets are tabular, use nested loops to write from one to the other

                for (int rowIndex = 0;rowIndex < dataGridView1.Rows.Count - 1; rowIndex++)
                {
                    for (int colIndex = 0; colIndex < dataGridView1.Columns.Count; colIndex++) //needed to go over the columns of each row
                    {
                        if (rowIndex == 0) //because the first row at index 0 is the header row
                        {
                            //in Excel, row and column indexes begin at 1,1, not 0,0

                            //Write out the header texts from the gridview to excel sheet
                            worksheet.Cells[rowIndex + 1, colIndex + 1] = dataGridView1.Columns[colIndex].HeaderText;
                        }
                        else
                        {
                            //fix the row index at 1, then change the column index over its possible values from 0 - 5
                            worksheet.Cells[rowIndex + 1, colIndex + 1] = dataGridView1.Rows[rowIndex].Cells[colIndex].Value.ToString();
                        }
                    }                
                }

                if (saveFileDialog1.ShowDialog() == DialogResult.OK) //user click Ok
                {
                    workbook.SaveAs(saveFileDialog1.FileName); //save file to drive
                    Process.Start("excel.exe", saveFileDialog1.FileName); //load excel with the exported file
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
            finally //this code always runs
            {
                excel.Quit();
                workbook = null; //make workbook object null
                excel = null;
            }
        }

        private void btnSaveToText_Click(object sender, EventArgs e)
        {
            if (saveFileDialog1.ShowDialog() == DialogResult.OK) //Check whether somebody has clicked the OK button
            {
                using (StreamWriter sw = new StreamWriter(saveFileDialog1.FileName))
                {
                    foreach (DataGridViewRow row in dataGridView1.Rows) //grab each row in the datagrid view
                    {
                        foreach (DataGridViewCell cell in row.Cells) //ocne you have a row grabbed, go throught the cells of that row
                        {
                            sw.Write(cell.Value); //this line actually write the value to a text file
                            sw.WriteLine(); //this pushes the cursor to the next line
                        }
                    }
                }
                Process.Start("notepad.exe", saveFileDialog1.FileName); //open file in notepad once file is written to the drive.
            }
        }

        private void btnOpenWord_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Word._Application word = new Microsoft.Office.Interop.Word.Application(); //make a new word object
            Document doc = word.Documents.Add(); //make a new document
            Microsoft.Office.Interop.Word.Range rng = doc.Range(0, 0);
            Table wdTable = doc.Tables.Add(rng, dataGridView1.Rows.Count, dataGridView1.Columns.Count); //make a new table based on our datagrid view
            wdTable.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleDouble; //make a thick outer border
            wdTable.Borders.InsideLineStyle = WdLineStyle.wdLineStyleSingle; //make the cell lines thin

            try
            {
                doc = word.ActiveDocument; //make an active document in word

                //i is the row index from the datagrid view
                for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
                {
                    for (int j = 0; j < dataGridView1.Columns.Count; j++) //this loop is needed to step through the columns of each row
                        //line below runs several times, each time writing the cell value from the grid to word.
                        wdTable.Cell(i + 1, j + 1).Range.InsertAfter(dataGridView1.Rows[i].Cells[j].Value.ToString());
                }
                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    doc.SaveAs(saveFileDialog1.FileName); //save file to drive
                    Process.Start("winword.exe", saveFileDialog1.FileName); //Open doc in word after the table is made.
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                word.Quit();
                word = null;
                doc = null;
            }
        }
    }
}
