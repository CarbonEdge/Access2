using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Data;

namespace AutodeskAccess
{
    class deskTool
    {
        private string DBFileName="";//write
        string connectionString;

        private OleDbConnection connection = new OleDbConnection();

        private List<string> products = new List<string>();//product names
        private SortedList<string, string> abreviations = new SortedList<string, string>();//product names

        private string[] resultFiles = null;//read from
        private string[] HTMLfile = null;
        private string[] licFiles = null;

        internal string selectDB()
        {
            try
            {
                OpenFileDialog choofdlog = new OpenFileDialog();
                choofdlog.Filter = "All Files (*.*)|*.*";
                choofdlog.FilterIndex = 1;
                if (choofdlog.ShowDialog() == DialogResult.OK)
                {
                    DBFileName = choofdlog.FileName;
                    connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;" +
                @"Data Source=" + DBFileName + ";" +
                @"Persist Security Info=False;";
                    Saveproducts();
                }
                return connectionString;
                MessageBox.Show(connectionString);                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                return connectionString;
            }
        }//select db ends

        internal void Saveproducts()
        {
            try
            {
                string PID = "SELECT PID FROM products";//
                string ABR = "SELECT Abreviation, ProductName FROM Abr";//

                connection = new OleDbConnection(connectionString);//db object

                using (OleDbCommand command = new OleDbCommand(PID, connection))//array of products
                {
                    connection.Open();
                    OleDbDataReader reader = command.ExecuteReader();

                    while (reader.Read())
                    {
                        products.Add(reader[0].ToString());
                    }
                    reader.Close();
                    connection.Close();

                }//end using
                using (OleDbCommand command = new OleDbCommand(ABR, connection))//array of abreviations
                {
                    connection.Open();
                    OleDbDataReader reader = command.ExecuteReader();

                    while (reader.Read())
                    {
                        abreviations.Add(reader[0].ToString(), reader[1].ToString());
                    }

                    reader.Close();
                    connection.Close();
                }//end using
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }//ends connection to DB and writes the list of search items.

        internal string[] selectScript()
        {
            try
            {
                FolderBrowserDialog fbd = new FolderBrowserDialog();
                fbd.RootFolder = Environment.SpecialFolder.MyComputer;//This causes the folder to begin at the root folder or your documents
                if (fbd.ShowDialog() == DialogResult.OK)
                {
                    HTMLfile = Directory.GetFiles(fbd.SelectedPath, "*.html", SearchOption.AllDirectories);
                    resultFiles = Directory.GetFiles(fbd.SelectedPath, "*.txt", SearchOption.AllDirectories);//change this to specify file type
                    licFiles = Directory.GetFiles(fbd.SelectedPath, "*.LIC", SearchOption.AllDirectories);
                    MessageBox.Show("Files Found: " + resultFiles.Count<string>().ToString() + "\n" + "HTML Files Found: " + HTMLfile.Count<string>().ToString() + "\n" + "Lic Files Found: " + licFiles.Count<string>().ToString());
                }
                return resultFiles;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
                return null;
            }
        }//selection of script ends

        internal void writeContents()// read all the text files to the Data base
        {
            try
            {
                MessageBox.Show(connectionString);
                OleDbConnection connection = new OleDbConnection(connectionString);
                OleDbCommand command = new OleDbCommand();
                command.Connection = connection;
                connection.Open();
                foreach (string x in resultFiles)//file path
                {
                    string CPname = Path.GetFileNameWithoutExtension(x);//computer name
                    int c = 0;//line count

                    using (StreamReader sr = new StreamReader(x))
                    {

                        string s;//text line
                        while ((s = sr.ReadLine()) != null)
                        {
                            //this breaks the program
                            command.CommandText = "insert into Script (FileName, LineNumber, Contents) Values ('" + CPname + "','" + c.ToString() + "','" + s + "')";
                            command.ExecuteNonQuery();

                            s = sr.ReadLine();
                            c++;
                        }
                    }

                }//ends loop over each file
                connection.Close();
            }//ends try
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
            }
        }//ends method
    }
}
