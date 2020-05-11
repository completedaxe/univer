using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using Microsoft.Office.Interop.Excel;
using DataTable = System.Data.DataTable;
using System.Data.SqlClient;
using System.Data;


namespace WpfApp1
{
    /// <summary>
    /// Логика взаимодействия для WinAdmin.xaml
    /// </summary>
    public partial class WinAdmin : System.Windows.Window
    {
        int idUser;
        public WinAdmin()
        {
            InitializeComponent();
        }
        //создаем метод для заполнения (обновления) DataGrid
        public void Update()
        {
            using (SqlConnection connect = new SqlConnection(Properties.Settings.Default.DBConnect))
            {
                connect.Open();
                SqlDataAdapter adapter = new SqlDataAdapter("Select * from [users]", connect);
                var data = new DataTable();
                adapter.Fill(data);
                dgUsers.ItemsSource = data.DefaultView;

                SqlCommand command = new SqlCommand("Select DISTINCT role from [users] ", connect);
                SqlDataReader reader = command.ExecuteReader();
                while (reader.Read())
                {
                    if (!cmbRole.Items.Contains(reader.GetValue(0).ToString()))
                        cmbRole.Items.Add(reader.GetValue(0).ToString());

                    if (!cmbRoleUp.Items.Contains(reader.GetValue(0).ToString()))
                        cmbRoleUp.Items.Add(reader.GetValue(0).ToString());
                }
            }
        }

        // заполняем DataGrid при открытии окна
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            using (SqlConnection connect = new SqlConnection(Properties.Settings.Default.DBConnect))
            {
                connect.Open();
                DataTable data = new DataTable();
                SqlDataAdapter adapter = new SqlDataAdapter("Select * from [users]", connect);
                adapter.Fill(data);
                dgUsers.ItemsSource = data.DefaultView;
            }
            Update();
        }

        private void btnInsert_Click(object sender, RoutedEventArgs e)
        {
            using (SqlConnection connect = new SqlConnection(Properties.Settings.Default.DBConnect))
            {
                connect.Open();
                SqlCommand command = new SqlCommand("Insert into [users] ([users].[login], [users].[password], [users].[name_user], [users].[role], [users].[status]) values ('" + txtLogin.Text + "', '" + txtPassword.Text + "', '" + txtFIO.Text + "', '" + cmbRole.Text + "', 1)", connect);
                command.ExecuteNonQuery();
                MessageBox.Show("Пользователь добавлен");
                Update(); //обновляем DataGrid  
            }
        }

        private void btnUpdate_Click(object sender, RoutedEventArgs e)
        {
            using (SqlConnection connect = new SqlConnection(Properties.Settings.Default.DBConnect))
            {
                connect.Open();
                SqlCommand command = new SqlCommand("update [users] set [users].[login]='" + txtLoginUp.Text + "', [users].[password]='" + txtPasswordUp.Text + "', [users].[name_user]='" + txtFIOUp.Text + "', [users].[role]='" + cmbRoleUp.Text + "' where [users].[id_user]=" + idUser  + "", connect);
                command.ExecuteNonQuery();
                MessageBox.Show("Пользователь изменен");
                Update();
            }
        }

        private void dgUsers_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                if (dgUsers.SelectedItem != null)
                {           
                    txtLoginUp.Text = ((DataRowView)dgUsers.SelectedItem)[1].ToString();
                    txtPasswordUp.Text = ((DataRowView)dgUsers.SelectedItem)[2].ToString();
                    txtFIOUp.Text = ((DataRowView)dgUsers.SelectedItem)[3].ToString();
                    cmbRoleUp.Text = ((DataRowView)dgUsers.SelectedItem)[4].ToString();
                    idUser = Convert.ToInt32(((DataRowView)dgUsers.SelectedItem)[0]);
                }
            }
            catch
            {
                MessageBox.Show("Выберите пользователя");
            }
        }

        private void btnDelete_Click(object sender, RoutedEventArgs e)
        {
            using (SqlConnection connect = new SqlConnection(Properties.Settings.Default.DBConnect))
            {
                connect.Open();
                SqlCommand command = new SqlCommand("update [users] set [users].[status]=0 where [users].[id_user]=" + idUser + "", connect);
                command.ExecuteNonQuery();
                MessageBox.Show("Пользователь изменен");
                Update();
            }    
        }

        private void txtPoisk_TextChanged(object sender, TextChangedEventArgs e)
        {
            using (SqlConnection connect = new SqlConnection(Properties.Settings.Default.DBConnect))
            {
                connect.Open();
                SqlDataAdapter adapter = new SqlDataAdapter("select * from [users] where [users].[login] like '%" + txtPoisk.Text + "%' or [users].[password] like '%" +txtPoisk.Text + "%' or [users].[name_user] like '%" + txtPoisk.Text + "%' or [users].[role] like '%" + txtPoisk.Text + "%'", connect);
                var data = new DataTable();
                adapter.Fill(data);
                dgUsers.ItemsSource = data.DefaultView;
            }
        }

        private void txtFIO_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            e.Handled = "йцукенгшщзхъфывапролджэячсмитьбюЙЦУКЕНГШЩЗХЪФЫВАПРОЛДЖЭЯЧСМИТЬБЮ.".IndexOf(e.Text) < 0;
        }

        private void txtFIOUp_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            e.Handled = "йцукенгшщзхъфывапролджэячсмитьбюЙЦУКЕНГШЩЗХЪФЫВАПРОЛДЖЭЯЧСМИТЬБЮ.".IndexOf(e.Text) < 0;
        }

        private void btnExcel_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            app.Visible = true;
            app.WindowState = XlWindowState.xlMaximized;

            Workbook wb = app.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            Worksheet ws = wb.Worksheets[1];
            DateTime currentDate = DateTime.Now;
            ws.Columns.AutoFit();

            ws.Range["A1:A3"].Value = "da";
            ws.Range["A4"].Value = "dasad";
            ws.Range["A5"].Value = currentDate;
            ws.Range["B6"].Value = "завтра =>";
            ws.Range["C6"].FormulaLocal = "=СУММ(D1:D10)";
            for (int i = 1; i <= 10; i++)
            {   
                ws.Range["D" + i].Value = i * 2;
            }

            for (int j = 0; j < dgUsers.Columns.Count; j++)
            {
                Range range = (Range)ws.Cells[1, j + 10];
                ws.Cells[1, j + 10].font.bold = true;
                ws.Cells[1, j + 10].columnwidth = 15;
                range.Value2 = dgUsers.Columns[j].Header;
            }

            for (int i = 0; i < dgUsers.Columns.Count; i++)
            {
                for (int j = 0; j < dgUsers.Items.Count; j++)
                {
                    TextBlock text = dgUsers.Columns[i].GetCellContent(dgUsers.Items[j]) as TextBlock;
                    Range range = (Range)ws.Cells[j+2, i + 10];
                    range.Value2 = text.Text;
                }
            }
        }
    }
}
