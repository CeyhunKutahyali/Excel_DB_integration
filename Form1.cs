using System.Collections;
using System.Data.SqlClient;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace Excel_DB_integration
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnGetData_Click(object sender, EventArgs e)
        {
            Excel.Application excelApp = new Excel.Application();
            excelApp.Visible = true;
            Excel.Workbook workbook = excelApp.Workbooks.Add(System.Reflection.Missing.Value);
            Excel.Worksheet Sheet1 = workbook.Sheets[1];

            string[] titles = { "Personel_Id", "Personel_Numara", "Personel_Ad", "Personel_Soyad", "Personel_Semt", "Personel_Þehir" };
            Excel.Range range;
            for (int i = 0; i < titles.Length; i++)
            {
                range = Sheet1.Cells[1, (1 + i)];
                range.Value2 = titles[i];
            }


            try
            {
                string query = "SELECT * FROM PERSONAL";
                SqlCommand cmd = new SqlCommand(query, ConnectionString.connection());
                SqlDataReader rdr = cmd.ExecuteReader();

                int line = 2; // ilk satýr baþlýk olduðu için 2.satýrdan ilerleyecek.

                while (rdr.Read())
                {
                    string PersonalId = rdr[0].ToString();
                    string PersonalNumber = rdr[1].ToString();
                    string PersonalName = rdr[2].ToString();
                    string PersonalSurname = rdr[3].ToString();
                    string PersonalDistrict = rdr[4].ToString();
                    string PersonalCity = rdr[5].ToString();
                    richTextBox1.Text = richTextBox1.Text + " " + PersonalId + " " + PersonalNumber + " " + PersonalName + " " + PersonalSurname + " " + PersonalDistrict + " " + PersonalCity + "\n";

                    range = Sheet1.Cells[line, 1];
                    range.Value2 = PersonalId;
                    range = Sheet1.Cells[line, 2];
                    range.Value2 = PersonalNumber;
                    range = Sheet1.Cells[line, 3];
                    range.Value2 = PersonalName;
                    range = Sheet1.Cells[line, 4];
                    range.Value2 = PersonalSurname;
                    range = Sheet1.Cells[line, 5];
                    range.Value2 = PersonalDistrict;
                    range = Sheet1.Cells[line, 6];
                    range.Value2 = PersonalCity;
                    line++;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Okuma Ýþlemi Sýrasýnda Bir Hata Oluþtu." + ex.Message);
            }
            finally
            {
                if (ConnectionString.connection != null)
                {
                    ConnectionString.connection().Close();
                }
            }
        }

        private void btnImport_Click(object sender, EventArgs e)
        {
            Excel.Application impExcelapp;
            Excel.Workbook impExcelWorkbook;
            Excel.Worksheet impExcelWorksheet;
            Excel.Range range;

            int rowCount = 0;
            int columnCount = 0;

            impExcelapp = new Excel.Application();
            impExcelWorkbook = impExcelapp.Workbooks.Open("C:\\Users\\ceyhun.kutahyali\\Documents\\Kitap1.xlsx");
            impExcelWorksheet = (Excel.Worksheet)impExcelWorkbook.Worksheets.get_Item(1); //excel sheet 
            range = impExcelWorksheet.UsedRange;

            richTextBox2.Clear();

            //ilk satýr kolon adlarý ise verileri okumaya 2.satýrdan baþlat.

            for (rowCount = 2; rowCount <= range.Rows.Count; rowCount++)
            {
                ArrayList list = new ArrayList();

                for (columnCount = 1; columnCount <= range.Columns.Count; columnCount++)
                {
                    string readCell = Convert.ToString((range.Cells[rowCount, columnCount] as Excel.Range).Value2);
                    richTextBox2.Text = richTextBox2.Text + readCell + " ";
                    list.Add(readCell);
                }
                richTextBox2.Text = richTextBox2.Text + "\n ";

                try
                {
                    SqlCommand cmd = new SqlCommand("INSERT INTO Personal (PersonalNumber, PersonalName, PersonalSurname, PersonalDistrict, PersonalCity) VALUES (@p1, @p2, @p3, @p4, @p5)", ConnectionString.connection());
                    cmd.Parameters.AddWithValue("@p1", list[1]);
                    cmd.Parameters.AddWithValue("@p2", list[2]);
                    cmd.Parameters.AddWithValue("@p3", list[3]);
                    cmd.Parameters.AddWithValue("@p4", list[4]);
                    cmd.Parameters.AddWithValue("@p5", list[5]);
                    cmd.ExecuteNonQuery();
                }
                catch(Exception ex)
                {
                    MessageBox.Show("Veritabanýna Kayýt Esnasýnda Bir Sorun Oluþtu. \n" + ex.Message);
                }
                finally
                {
                    if(ConnectionString.connection != null)
                    {
                        ConnectionString.connection().Close();
                         
                    }
                }
            }

            impExcelapp.Quit();
            ReleaseObject(impExcelWorkbook);
            ReleaseObject(impExcelWorksheet);
            ReleaseObject(impExcelapp);
        }

        private void ReleaseObject(object obj)
        {
            try
            {
                Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}