using System.Data.SqlClient;
using System.Web;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.ApplicationServices;
using static System.Windows.Forms.LinkLabel;
using Excel = Microsoft.Office.Interop.Excel;

namespace Test
{ 
    public partial class Form1 : Form
    {
        public string ConnectionString = "Data Source=DESKTOP-6K3QQ5H\\SQLEXPRESS;Initial Catalog=test;Integrated Security=True;Connect Timeout=30;Encrypt=False;TrustServerCertificate=False;ApplicationIntent=ReadWrite;MultiSubnetFailover=False";
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            SqlConnection conn = new SqlConnection(ConnectionString);
            conn.Open();

            String name = tbName.Text;
            int age = int.Parse(tbAge.Text);

            String Query = "INSERT INTO name (name, age) VALUES('"+name+"', "+age+")";
            SqlCommand cmd = new SqlCommand(Query, conn);
            cmd.ExecuteNonQuery();

            conn.Close();
            MessageBox.Show("data has been send");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //string path = "C:\\Users\\Lin\\Documents\\name.xlsx";
            string path = tbPath.Text;
            Excel.Application excel = new Excel.Application();
            Excel.Workbook wb = excel.Workbooks.Open(path);
            Excel.Worksheet ws = wb.Worksheets[1];
            Excel.Range range = ws.UsedRange;

            SqlConnection conn = new SqlConnection(ConnectionString);
            conn.Open();

            int rows = range.Rows.Count;
            int cols = range.Columns.Count;

            string col_name = ws.Cells[1, 1].Value;
            string col_age = ws.Cells[1, 2].Value;
            string col_woonPlaats= ws.Cells[1, 3].Value;



            for (int r=1; r<=rows; r++)
            {
                for(int c=1; c<=cols; c++)
                {
                    if (r > 1 && c == cols)
                    {
                        string name_value = ws.Cells[r, c-2].Value;
                        int age_value = (int)ws.Cells[r, c-1].Value;
                        string woonPlaats_value = ws.Cells[r, c].Value;
                        string Query = "INSERT INTO name ("+ col_name + ", "+ col_age + ", "+col_woonPlaats+") VALUES('"+ name_value +"', "+ age_value +", '"+ woonPlaats_value+ "')";
                        SqlCommand cmd = new SqlCommand(Query, conn);
                        cmd.ExecuteNonQuery();
                    }
                }
            }

            conn.Close();
            MessageBox.Show("data has been send");
        }

        private void button3_Click(object sender, EventArgs e)
        {
            SqlConnection conn = new SqlConnection(ConnectionString);
            conn.Open();

            String Query = "TRUNCATE TABLE name";
            SqlCommand cmd = new SqlCommand(Query, conn);
            cmd.ExecuteNonQuery();

            conn.Close();
            MessageBox.Show("data has been cleared");
        }
    }
}