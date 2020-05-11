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
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Data.SqlClient;

namespace WpfApp1
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            txtLogin.Focus();
        }

        private void btnEnter_Click(object sender, RoutedEventArgs e)
        {
            string login = txtLogin.Text;
            string password = txtPass.Password;

            try {
                using (SqlConnection connect = new SqlConnection(Properties.Settings.Default.DBConnect))
                {
                    connect.Open();
                    SqlCommand command = new SqlCommand("Select * from [users] where[users].[login]='" + login + "' and  [users].[password]='" + password + "' and [users].[status]=1", connect);
                    SqlDataReader reader = command.ExecuteReader();
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            string role = reader.GetValue(4).ToString();
                            MessageBox.Show("Добро пожаловать, " + role);
                            switch (role)
                            {
                                case "Администратор":
                                    WinAdmin win = new WinAdmin();
                                    win.Show();
                                    break;
                                case "Секретарь":
                                    WinSekretar win1 = new WinSekretar();
                                    win1.Show();
                                    break;
                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show("Пользователя не существует");
                    }
                }
            }
            catch
            {
                MessageBox.Show("Нет соединения с БД");
            }
        }
    }
}
