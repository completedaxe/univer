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
using System.Data.SqlClient;
using System.Data;
using Microsoft.Office.Interop.Excel;
using DataTable = System.Data.DataTable;


namespace WpfApp1
{
    /// <summary>
    /// Логика взаимодействия для WinSekretar.xaml
    /// </summary>
    public partial class WinSekretar : System.Windows.Window
    {
        int idStudent;
        int idBall;
        int idPredmet;
        int idSpec;

        public WinSekretar()
        {
            InitializeComponent();
        }

        public void SUpdate()
        {
            using (SqlConnection connect = new SqlConnection(Properties.Settings.Default.DBConnect))
            {
                connect.Open();
                SqlDataAdapter adapter = new SqlDataAdapter("Select * from [students]", connect);
                var data = new DataTable();
                adapter.Fill(data);
                dgStudents.ItemsSource = data.DefaultView;

                SqlCommand command = new SqlCommand("Select DISTINCT family from [students]", connect);
                SqlDataReader reader = command.ExecuteReader();
                while (reader.Read())
                {
                    if (!ScmbFamily.Items.Contains(reader.GetValue(0).ToString()))
                        ScmbFamily.Items.Add(reader.GetValue(0).ToString());
                }  
            }

            using (SqlConnection connect = new SqlConnection(Properties.Settings.Default.DBConnect))
            {
                connect.Open();
                SqlDataAdapter adapter = new SqlDataAdapter("Select * from [balls]", connect);
                var data = new DataTable();
                adapter.Fill(data);
                dgBalls.ItemsSource = data.DefaultView;
            }

            using (SqlConnection connect6 = new SqlConnection(Properties.Settings.Default.DBConnect))
            {
                connect6.Open();
                SqlDataAdapter adapter6 = new SqlDataAdapter("Select * from [predmets]", connect6);
                var data6 = new DataTable();
                adapter6.Fill(data6);
                dgPredmets.ItemsSource = data6.DefaultView;
                SqlCommand command6 = new SqlCommand("Select DISTINCT Pcodpredmeta from [predmets]", connect6);
                SqlDataReader reader6 = command6.ExecuteReader();
                while (reader6.Read())
                {
                    if (!OcmbPredmetOne.Items.Contains(reader6.GetValue(0).ToString()))
                        OcmbPredmetOne.Items.Add(reader6.GetValue(0).ToString());
                }
            }

            using (SqlConnection connect6 = new SqlConnection(Properties.Settings.Default.DBConnect))
            {
                connect6.Open();
                SqlDataAdapter adapter6 = new SqlDataAdapter("Select * from [predmets]", connect6);
                var data6 = new DataTable();
                adapter6.Fill(data6);
                dgPredmets.ItemsSource = data6.DefaultView;
                SqlCommand command6 = new SqlCommand("Select DISTINCT Pcodpredmeta from [predmets]", connect6);
                SqlDataReader reader6 = command6.ExecuteReader();
                while (reader6.Read())
                {
                    if (!OcmbPredmetTwo.Items.Contains(reader6.GetValue(0).ToString()))
                        OcmbPredmetTwo.Items.Add(reader6.GetValue(0).ToString());
                }
            }

            using (SqlConnection connect6 = new SqlConnection(Properties.Settings.Default.DBConnect))
            {
                connect6.Open();
                SqlDataAdapter adapter6 = new SqlDataAdapter("Select * from [predmets]", connect6);
                var data6 = new DataTable();
                adapter6.Fill(data6);
                dgPredmets.ItemsSource = data6.DefaultView;
                SqlCommand command6 = new SqlCommand("Select DISTINCT Pcodpredmeta from [predmets]", connect6);
                SqlDataReader reader6 = command6.ExecuteReader();
                while (reader6.Read())
                {
                    if (!OcmbPredmetThree.Items.Contains(reader6.GetValue(0).ToString()))
                        OcmbPredmetThree.Items.Add(reader6.GetValue(0).ToString());
                }
            }

            using (SqlConnection connect6 = new SqlConnection(Properties.Settings.Default.DBConnect))
            {
                connect6.Open();
                SqlDataAdapter adapter6 = new SqlDataAdapter("Select * from [students]", connect6);
                var data6 = new DataTable();
                adapter6.Fill(data6);
                dgStudents.ItemsSource = data6.DefaultView;
                SqlCommand command6 = new SqlCommand("Select DISTINCT codstudent from [students]", connect6);
                SqlDataReader reader6 = command6.ExecuteReader();
                while (reader6.Read())
                {
                    if (!OcmbStudent.Items.Contains(reader6.GetValue(0).ToString()))
                        OcmbStudent.Items.Add(reader6.GetValue(0).ToString());
                }
            }

            using (SqlConnection connect6 = new SqlConnection(Properties.Settings.Default.DBConnect))
            {
                connect6.Open();
                SqlDataAdapter adapter6 = new SqlDataAdapter("Select * from [balls]", connect6);
                var data6 = new DataTable();
                adapter6.Fill(data6);
                dgBalls.ItemsSource = data6.DefaultView;
                SqlCommand command6 = new SqlCommand("Select DISTINCT Oball1 from [balls]", connect6);
                SqlDataReader reader6 = command6.ExecuteReader();
                while (reader6.Read())
                {
                    if (!OtxtBallOne.Items.Contains(reader6.GetValue(0).ToString()))
                        OtxtBallOne.Items.Add(reader6.GetValue(0).ToString());
                }
            }

            using (SqlConnection connect6 = new SqlConnection(Properties.Settings.Default.DBConnect))
            {
                connect6.Open();
                SqlDataAdapter adapter6 = new SqlDataAdapter("Select * from [balls]", connect6);
                var data6 = new DataTable();
                adapter6.Fill(data6);
                dgBalls.ItemsSource = data6.DefaultView;
                SqlCommand command6 = new SqlCommand("Select DISTINCT Oball2 from [balls]", connect6);
                SqlDataReader reader6 = command6.ExecuteReader();
                while (reader6.Read())
                {
                    if (!OtxtBallTwo.Items.Contains(reader6.GetValue(0).ToString()))
                        OtxtBallTwo.Items.Add(reader6.GetValue(0).ToString());
                }
            }

            using (SqlConnection connect6 = new SqlConnection(Properties.Settings.Default.DBConnect))
            {
                connect6.Open();
                SqlDataAdapter adapter6 = new SqlDataAdapter("Select * from [balls]", connect6);
                var data6 = new DataTable();
                adapter6.Fill(data6);
                dgBalls.ItemsSource = data6.DefaultView;
                SqlCommand command6 = new SqlCommand("Select DISTINCT Oball3 from [balls]", connect6);
                SqlDataReader reader6 = command6.ExecuteReader();
                while (reader6.Read())
                {
                    if (!OtxtBallThree.Items.Contains(reader6.GetValue(0).ToString()))
                        OtxtBallThree.Items.Add(reader6.GetValue(0).ToString());
                }
            }

            using (SqlConnection connect1 = new SqlConnection(Properties.Settings.Default.DBConnect))
            {
                connect1.Open();
                SqlDataAdapter adapter1 = new SqlDataAdapter("Select * from [students]", connect1);
                var data1 = new DataTable();
                adapter1.Fill(data1);
                dgStudents.ItemsSource = data1.DefaultView;
                SqlCommand command1 = new SqlCommand("Select DISTINCT gender from [students]", connect1);
                SqlDataReader reader1 = command1.ExecuteReader();
                while (reader1.Read())
                {
                    if (!ScmbGender.Items.Contains(reader1.GetValue(0).ToString()))
                        ScmbGender.Items.Add(reader1.GetValue(0).ToString());
                }
            } 

            using (SqlConnection connect2 = new SqlConnection(Properties.Settings.Default.DBConnect))
            {
                connect2.Open();
                SqlDataAdapter adapter2 = new SqlDataAdapter("Select * from [students]", connect2);
                var data2 = new DataTable();
                adapter2.Fill(data2);
                dgStudents.ItemsSource = data2.DefaultView;
                SqlCommand command2 = new SqlCommand("Select DISTINCT course from [students]", connect2);
                SqlDataReader reader2 = command2.ExecuteReader();
                while (reader2.Read())
                {
                    if (!ScmbCourse.Items.Contains(reader2.GetValue(0).ToString()))
                        ScmbCourse.Items.Add(reader2.GetValue(0).ToString());
                }
            }

            using (SqlConnection connect3 = new SqlConnection(Properties.Settings.Default.DBConnect))
            {
                connect3.Open();
                SqlDataAdapter adapter3 = new SqlDataAdapter("Select * from [students]", connect3);
                var data3 = new DataTable();
                adapter3.Fill(data3);
                dgStudents.ItemsSource = data3.DefaultView;
                SqlCommand command3 = new SqlCommand("Select DISTINCT groupp from [students]", connect3);
                SqlDataReader reader3 = command3.ExecuteReader();
                while (reader3.Read())
                {
                    if (!ScmbGroup.Items.Contains(reader3.GetValue(0).ToString()))
                        ScmbGroup.Items.Add(reader3.GetValue(0).ToString());
                }
            }

            using (SqlConnection connect3 = new SqlConnection(Properties.Settings.Default.DBConnect))
            {
                connect3.Open();
                SqlDataAdapter adapter3 = new SqlDataAdapter("Select * from [students]", connect3);
                var data3 = new DataTable();
                adapter3.Fill(data3);
                dgStudents.ItemsSource = data3.DefaultView;
                SqlCommand command3 = new SqlCommand("Select DISTINCT groupp from [students]", connect3);
                SqlDataReader reader3 = command3.ExecuteReader();
                while (reader3.Read())
                {
                    if (!OtxtSpeci.Items.Contains(reader3.GetValue(0).ToString()))
                        OtxtSpeci.Items.Add(reader3.GetValue(0).ToString());
                }
            }

            using (SqlConnection connect5 = new SqlConnection(Properties.Settings.Default.DBConnect))
            {
                connect5.Open();
                SqlDataAdapter adapter5 = new SqlDataAdapter("Select * from [students]", connect5);
                var data5 = new DataTable();
                adapter5.Fill(data5);
                dgStudents.ItemsSource = data5.DefaultView;
                //SqlCommand command5 = new SqlCommand("Select DISTINCT CASE WHEN status = 1 THEN 'Да' ELSE 'Нет' END AS status from [students]", connect5);
                SqlCommand command5 = new SqlCommand("Select DISTINCT status from [students]", connect5);
                SqlDataReader reader5 = command5.ExecuteReader();
                while (reader5.Read())
                {
                    if (!ScmbForm.Items.Contains(reader5.GetValue(0).ToString()))
                        ScmbForm.Items.Add(reader5.GetValue(0).ToString());
                }
            }
                using (SqlConnection connect = new SqlConnection(Properties.Settings.Default.DBConnect))
                {
                    connect.Open();
                    SqlDataAdapter adapter = new SqlDataAdapter("Select * from [predmets]", connect);
                    var data = new DataTable();
                    adapter.Fill(data); 
                    dgPredmets.ItemsSource = data.DefaultView;
                }

                using (SqlConnection connect = new SqlConnection(Properties.Settings.Default.DBConnect))
                {
                    connect.Open();
                    SqlDataAdapter adapter = new SqlDataAdapter("Select * from [speci]", connect);
                    var data = new DataTable();
                    adapter.Fill(data);
                    dgSpeci.ItemsSource = data.DefaultView;
                }
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            SUpdate();
        }

        //добавление студента
        private void SbtnInsert_Click(object sender, RoutedEventArgs e)
        {
            try { 
            using (SqlConnection connect = new SqlConnection(Properties.Settings.Default.DBConnect))
            {
                connect.Open();
                SqlCommand command = new SqlCommand("Insert into [students] ([students].[fio], [students].[gender], [students].[bd], [students].[family], [students].[street], [students].[phone], [students].[passport], [students].[numberz], [students].[datep], [students].[groupp], [students].[course], [students].[status])  values ('" + StxtFIO.Text + "', '" + ScmbGender.Text + "', '" + StxtBD.Text + "', '" + ScmbFamily.Text + "' , '" + StxtAdres.Text + "' , '" + StxtPhone.Text + "' , '" + StxtPassport.Text + "' , '" + StxtZK.Text + "' , '" + StxtDate.Text + "' , '" + ScmbGroup.Text + "' , '" + ScmbCourse.Text + "' ,  '" + ScmbForm.Text + "')", connect);
                command.ExecuteNonQuery();
                MessageBox.Show("Студент добавлен");
                SUpdate();
            }
            }
            catch
            {
                MessageBox.Show("Введите данные");
            }
        }

        //изменение студента
        private void SbtnUpdate_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                using (SqlConnection connect = new SqlConnection(Properties.Settings.Default.DBConnect))
                {
                    connect.Open();
                    SqlCommand command = new SqlCommand("update [students] set [students].[fio]='" + StxtFIO.Text + "', [students].[gender]='" + ScmbGender.Text + "', [students].[bd]='" + StxtBD.Text + "', [students].[family]='" + ScmbFamily.Text + "', [students].[street]='" + StxtAdres.Text + "', [students].[phone]='" + StxtPhone.Text + "', [students].[passport]='" + StxtPassport.Text + "', [students].[numberz]='" + StxtZK.Text + "', [students].[datep]='" + StxtDate.Text + "', [students].[groupp]='" + ScmbGroup.Text + "', [students].[course]='" + ScmbCourse.Text + "', [students].[status]= '" + ScmbForm.Text + "' where [students].[codstudent]=" + idStudent + "", connect);
                    command.ExecuteNonQuery();
                    MessageBox.Show("Студент изменен");
                    SUpdate();
                }
            }
            catch
            {
                MessageBox.Show("Введите данные");
            }
        }

        private void dgStudents_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                if (dgStudents.SelectedItem != null)
                {
                    StxtFIO.Text = ((DataRowView)dgStudents.SelectedItem)[1].ToString();
                    ScmbGender.Text = ((DataRowView)dgStudents.SelectedItem)[2].ToString();
                    StxtBD.Text = ((DataRowView)dgStudents.SelectedItem)[3].ToString();
                    ScmbFamily.Text = ((DataRowView)dgStudents.SelectedItem)[4].ToString();
                    StxtAdres.Text = ((DataRowView)dgStudents.SelectedItem)[5].ToString();
                    StxtPhone.Text = ((DataRowView)dgStudents.SelectedItem)[6].ToString();
                    StxtPassport.Text = ((DataRowView)dgStudents.SelectedItem)[7].ToString();
                    StxtZK.Text = ((DataRowView)dgStudents.SelectedItem)[8].ToString();
                    StxtDate.Text = ((DataRowView)dgStudents.SelectedItem)[9].ToString();
                    ScmbGroup.Text = ((DataRowView)dgStudents.SelectedItem)[10].ToString();
                    ScmbCourse.Text = ((DataRowView)dgStudents.SelectedItem)[11].ToString();
                    ScmbForm.Text = ((DataRowView)dgStudents.SelectedItem)[12].ToString();
                    idStudent = Convert.ToInt32(((DataRowView)dgStudents.SelectedItem)[0]);
                }
            }
            catch
            {
                MessageBox.Show("Выберите студента");
            }
        }

        private void StxtPoisk_TextChanged(object sender, TextChangedEventArgs e)
        {
            using (SqlConnection connect = new SqlConnection(Properties.Settings.Default.DBConnect))
            {
                connect.Open();
                SqlDataAdapter adapter = new SqlDataAdapter("select * from [students] where [students].[fio] like '%" + StxtPoisk.Text + "%' or [students].[gender] like '%" + StxtPoisk.Text + "%' or [students].[bd] like '%" + StxtPoisk.Text + "%' or [students].[family] like '%" + StxtPoisk.Text + "%' or [students].[street] like '%" + StxtPoisk.Text + "%' or [students].[phone] like '%" + StxtPoisk.Text + "%' or [students].[passport] like '%" + StxtPoisk.Text + "%' or [students].[numberz] like '%" + StxtPoisk.Text + "%' or [students].[datep] like '%" + StxtPoisk.Text + "%' or [students].[groupp] like '%" + StxtPoisk.Text + "%' or [students].[course] like '%" + StxtPoisk.Text + "%' or [students].[status] like '%" + StxtPoisk.Text + "%' ", connect);
                var data = new DataTable();
                adapter.Fill(data);
                dgStudents.ItemsSource = data.DefaultView;
            }
        }

        private void StxtBD_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            e.Handled = "0123456789./".IndexOf(e.Text) < 0;
        }

        private void StxtPhone_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            e.Handled = "0123456789+".IndexOf(e.Text) < 0;
        }

        private void StxtPassport_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            e.Handled = "0123456789-".IndexOf(e.Text) < 0;
        }

        private void StxtZK_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            e.Handled = "0123456789".IndexOf(e.Text) < 0;
        }

        private void StxtDate_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            e.Handled = "0123456789./".IndexOf(e.Text) < 0;
        }

        private void StxtFIO_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            e.Handled = "йцукенгшщзхъфывапролджэячсмитьбюЙЦУКЕНГШЩЗХЪФЫВАПРОЛДЖЭЯЧСМИТЬБЮ.".IndexOf(e.Text) < 0;
        }

        private void StxtAdres_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            e.Handled = "йцукенгшщзхъфывапролджэячсмитьбюЙЦУКЕНГШЩЗХЪФЫВАПРОЛДЖЭЯЧСМИТЬБЮ,.0123456789-".IndexOf(e.Text) < 0;
        }

        private void StxtPoisk_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            e.Handled = "йцукенгшщзхъфывапролджэячсмитьбюЙЦУКЕНГШЩЗХЪФЫВАПРОЛДЖЭЯЧСМИТЬБЮ,.0123456789-".IndexOf(e.Text) < 0;
        }

        private void ObtnInsert_Click(object sender, RoutedEventArgs e)
        {
            try { 
            OtxtSrBall.Text = Math.Round((double.Parse(OtxtBallOne.Text)+double.Parse(OtxtBallTwo.Text)+double.Parse(OtxtBallThree.Text))/3).ToString();
            using (SqlConnection connect = new SqlConnection(Properties.Settings.Default.DBConnect))
            {
                connect.Open();
                SqlCommand command = new SqlCommand("Insert into [balls] ([balls].[Odateekz1], [balls].[Ocodpredmeta1], [balls].[Oball1], [balls].[Odateekz2], [balls].[Ocodpredmeta2], [balls].[Oball2], [balls].[Odateekz3], [balls].[Ocodpredmeta3], [balls].[Oball3], [balls].[Osrball], [balls].[Ocodstudent])  values ('" + OtxtDateOne.Text + "', '" + OcmbPredmetOne.Text + "', '" + OtxtBallOne.Text + "', '" + OtxtDateTwo.Text + "' , '" + OcmbPredmetTwo.Text + "' , '" + OtxtBallTwo.Text + "' , '" + OtxtDateThree.Text + "' , '" + OcmbPredmetThree.Text + "' , '" + OtxtBallThree.Text + "', '" + OtxtSrBall.Text + "', '" + OcmbStudent.Text + "')", connect);
                command.ExecuteNonQuery();
                MessageBox.Show("Оценка добавлена");
                SUpdate();
                OtxtSrBall.Text = null;
            }
            }
            catch
            {
                MessageBox.Show("Введите данные");
            }

        }

        private void OtxtDateOne_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            e.Handled = "0123456789./".IndexOf(e.Text) < 0;
        }

        private void OtxtDateTwo_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            e.Handled = "0123456789./".IndexOf(e.Text) < 0;
        }

        private void OtxtDateThree_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            e.Handled = "0123456789./".IndexOf(e.Text) < 0;
        }

        private void OtxtPoisk_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            e.Handled = "0123456789./".IndexOf(e.Text) < 0;
        }

        private void ObtnUpdate_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                OtxtSrBall.Text = Math.Round((double.Parse(OtxtBallOne.Text) + double.Parse(OtxtBallTwo.Text) + double.Parse(OtxtBallThree.Text)) / 3).ToString();
                using (SqlConnection connect = new SqlConnection(Properties.Settings.Default.DBConnect))
                {
                    connect.Open();
                    SqlCommand command = new SqlCommand("update [balls] set [balls].[Ocodstudent]='" + OcmbStudent.Text + "', [balls].[Odateekz1]='" + OtxtDateOne.Text + "', [balls].[Odateekz2]='" + OtxtDateTwo.Text + "', [balls].[Odateekz3]='" + OtxtDateThree.Text + "', [balls].[Ocodpredmeta1]='" + OcmbPredmetOne.Text + "', [balls].[Ocodpredmeta2]='" + OcmbPredmetTwo.Text + "', [balls].[Ocodpredmeta3]='" + OcmbPredmetThree.Text + "', [balls].[Oball1]='" + OtxtBallOne.Text + "', [balls].[Oball2]='" + OtxtBallTwo.Text + "', [balls].[Oball3]='" + OtxtBallThree.Text + "', [balls].[Osrball]='" + OtxtSrBall.Text + "' where [balls].[Ocodstudent]=" + idBall + "", connect);
                    command.ExecuteNonQuery();
                    MessageBox.Show("Оценка изменена");
                    SUpdate();
                    OtxtSrBall.Text = null;
                }
            }
            catch
            {
                MessageBox.Show("Введите данные");
            }
        }

        private void dgBalls_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                if (dgBalls.SelectedItem != null)
                {
                    OcmbStudent.Text = ((DataRowView)dgBalls.SelectedItem)[0].ToString();
                    OtxtDateOne.Text = ((DataRowView)dgBalls.SelectedItem)[1].ToString();
                    OcmbPredmetOne.Text = ((DataRowView)dgBalls.SelectedItem)[2].ToString();
                    OtxtBallOne.Text = ((DataRowView)dgBalls.SelectedItem)[3].ToString();
                    OtxtDateTwo.Text = ((DataRowView)dgBalls.SelectedItem)[4].ToString();
                    OcmbPredmetTwo.Text = ((DataRowView)dgBalls.SelectedItem)[5].ToString();
                    OtxtBallTwo.Text = ((DataRowView)dgBalls.SelectedItem)[6].ToString();
                    OtxtDateThree.Text = ((DataRowView)dgBalls.SelectedItem)[7].ToString();
                    OcmbPredmetThree.Text = ((DataRowView)dgBalls.SelectedItem)[8].ToString();
                    OtxtBallThree.Text = ((DataRowView)dgBalls.SelectedItem)[9].ToString();
                    OtxtSrBall.Text = ((DataRowView)dgBalls.SelectedItem)[10].ToString();
                    idBall = Convert.ToInt32(((DataRowView)dgBalls.SelectedItem)[0]);  
                }
            }
            catch 
            {
                MessageBox.Show("Выберите оценку");
            }
        }

        private void OtxtPoisk_TextChanged(object sender, TextChangedEventArgs e)
        {
            using (SqlConnection connect = new SqlConnection(Properties.Settings.Default.DBConnect))
            {
                connect.Open();
                SqlDataAdapter adapter = new SqlDataAdapter("select * from [balls] where [balls].[Odateekz1] like '%" + OtxtPoisk.Text + "%' or [balls].[Odateekz2] like '%" + OtxtPoisk.Text + "%' or [balls].[Odateekz3] like '%" + OtxtPoisk.Text + "%' or [balls].[Ocodpredmeta1] like '%" + OtxtPoisk.Text + "%' or [balls].[Ocodpredmeta2] like '%" + OtxtPoisk.Text + "%' or [balls].[Ocodpredmeta3] like '%" + OtxtPoisk.Text + "%' or [balls].[Oball1] like '%" + OtxtPoisk.Text + "%' or [balls].[Oball2] like '%" + OtxtPoisk.Text + "%' or [balls].[Oball3] like '%" + OtxtPoisk.Text + "%' ", connect);
                var data = new DataTable();
                adapter.Fill(data);
                dgBalls.ItemsSource = data.DefaultView;
            }
        }

        private void PtxtNaimPredmeta_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            e.Handled = "qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNMйцукенгшщзхъфывапролджэячсмитьбюЙЦУКЕНГШЩЗХЪФЫВАПРОЛДЖЭЯЧСМИТЬБЮ,.0123456789-".IndexOf(e.Text) < 0;
        }

        private void PtxtOpisPredmeta_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            e.Handled = "qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNMйцукенгшщзхъфывапролджэячсмитьбюЙЦУКЕНГШЩЗХЪФЫВАПРОЛДЖЭЯЧСМИТЬБЮ,.0123456789-".IndexOf(e.Text) < 0;
        }

        private void StxtsNaimSpec_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            e.Handled = "qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNMйцукенгшщзхъфывапролджэячсмитьбюЙЦУКЕНГШЩЗХЪФЫВАПРОЛДЖЭЯЧСМИТЬБЮ,.0123456789-".IndexOf(e.Text) < 0;
        }

        private void StxtOpisSpec_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            e.Handled = "qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNMйцукенгшщзхъфывапролджэячсмитьбюЙЦУКЕНГШЩЗХЪФЫВАПРОЛДЖЭЯЧСМИТЬБЮ,.0123456789-".IndexOf(e.Text) < 0;
        }

        private void PtxtPoisk_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            e.Handled = "qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNMйцукенгшщзхъфывапролджэячсмитьбюЙЦУКЕНГШЩЗХЪФЫВАПРОЛДЖЭЯЧСМИТЬБЮ,.0123456789-".IndexOf(e.Text) < 0;
        }

        private void SStxtPoisk_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            e.Handled = "qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNMйцукенгшщзхъфывапролджэячсмитьбюЙЦУКЕНГШЩЗХЪФЫВАПРОЛДЖЭЯЧСМИТЬБЮ,.0123456789-".IndexOf(e.Text) < 0;
        }

        private void PbtnsInsert_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                using (SqlConnection connect = new SqlConnection(Properties.Settings.Default.DBConnect))
                {
                    connect.Open();
                    SqlCommand command = new SqlCommand("Insert into [predmets] ([predmets].[Pnaimpredmeta], [predmets].[Popisaniepredmeta])  values ('" + PtxtNaimPredmeta.Text + "', '" + PtxtOpisPredmeta.Text + "')", connect);
                    command.ExecuteNonQuery();
                    MessageBox.Show("Предмет добавлен");
                    SUpdate();
                }
            }
            catch
            {
                MessageBox.Show("Введите данные");
            }
        }

        private void PbtnsUpdate_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                using (SqlConnection connect = new SqlConnection(Properties.Settings.Default.DBConnect))
                {
                    connect.Open();
                    SqlCommand command = new SqlCommand("update [predmets] set [predmets].[Pnaimpredmeta]='" + PtxtNaimPredmeta.Text + "', [predmets].[Popisaniepredmeta]='" + PtxtOpisPredmeta.Text + "' where [predmets].[Pcodpredmeta]=" + idPredmet + "", connect);
                    command.ExecuteNonQuery();
                    MessageBox.Show("Предмет изменен");
                    SUpdate();
                }
            }
            catch
            {
                MessageBox.Show("Введите данные");
            }
        }

        private void dgPredmets_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                if (dgPredmets.SelectedItem != null)
                {
                    PtxtNaimPredmeta.Text = ((DataRowView)dgPredmets.SelectedItem)[1].ToString();
                    PtxtOpisPredmeta.Text = ((DataRowView)dgPredmets.SelectedItem)[2].ToString();

                    idPredmet = Convert.ToInt32(((DataRowView)dgPredmets.SelectedItem)[0]);
                }
            }
            catch
            {
                MessageBox.Show("Выберите предмет");
            }
        }

        private void SSbtnInsert_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                using (SqlConnection connect = new SqlConnection(Properties.Settings.Default.DBConnect))
                {
                    connect.Open();
                    SqlCommand command = new SqlCommand("Insert into [speci] ([speci].[SnaimSpec], [speci].[SopisSpec])  values ('" + StxtsNaimSpec.Text + "', '" + StxtOpisSpec.Text + "')", connect);
                    command.ExecuteNonQuery();
                    MessageBox.Show("Специальность добавлена");
                    SUpdate();
                }
            }
            catch
            {
                MessageBox.Show("Введите данные");
            }
        }

        private void SSbtnUpdate_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                using (SqlConnection connect = new SqlConnection(Properties.Settings.Default.DBConnect))
                {
                    connect.Open();
                    SqlCommand command = new SqlCommand("update [speci] set [speci].[SnaimSpec]='" + StxtsNaimSpec.Text + "', [speci].[SopisSpec]='" + StxtOpisSpec.Text + "' where [speci].[ScodSpec]=" + idSpec + "", connect);
                    command.ExecuteNonQuery();
                    MessageBox.Show("Специальность изменена");
                    SUpdate();
                }
            }
            catch
            {
                MessageBox.Show("Введите данные");
            }
        }

        private void dgSpeci_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                if (dgSpeci.SelectedItem != null)
                {
                    StxtsNaimSpec.Text = ((DataRowView)dgSpeci.SelectedItem)[1].ToString();
                    StxtOpisSpec.Text = ((DataRowView)dgSpeci.SelectedItem)[2].ToString();
                    idSpec = Convert.ToInt32(((DataRowView)dgSpeci.SelectedItem)[0]);
                }
            }
            catch
            {
                MessageBox.Show("Выберите специальность");
            }
        }

        private void PtxtPoisk_TextChanged(object sender, TextChangedEventArgs e)
        {
            using (SqlConnection connect = new SqlConnection(Properties.Settings.Default.DBConnect))
            {
                connect.Open();
                SqlDataAdapter adapter = new SqlDataAdapter("select * from [predmets] where [predmets].[Pcodpredmeta] like '%" + PtxtPoisk.Text + "%' or [predmets].[Pnaimpredmeta] like '%" + PtxtPoisk.Text + "%' or [predmets].[Popisaniepredmeta] like '%" + PtxtPoisk.Text + "%'", connect);
                var data = new DataTable();
                adapter.Fill(data);
                dgPredmets.ItemsSource = data.DefaultView;
            }
        }

        private void SStxtPoisk_TextChanged(object sender, TextChangedEventArgs e)
        {
            using (SqlConnection connect = new SqlConnection(Properties.Settings.Default.DBConnect))
            {
                connect.Open();
                SqlDataAdapter adapter = new SqlDataAdapter("select * from [speci] where [speci].[ScodSpec] like '%" + SStxtPoisk.Text + "%' or [speci].[SnaimSpec] like '%" + SStxtPoisk.Text + "%' or [speci].[SopisSpec] like '%" + SStxtPoisk.Text + "%'", connect);
                var data = new DataTable();
                adapter.Fill(data);
                dgSpeci.ItemsSource = data.DefaultView;
            }
        }

        private void ObtnSearchFour_Click(object sender, RoutedEventArgs e)
        {
            using (SqlConnection connect = new SqlConnection(Properties.Settings.Default.DBConnect))
            {
                connect.Open();
                SqlDataAdapter adapter = new SqlDataAdapter("Select fio from [students]", connect);
                var data = new DataTable();
                adapter.Fill(data);
                dgStudentsFour.ItemsSource = data.DefaultView;
            }

            using (SqlConnection connect = new SqlConnection(Properties.Settings.Default.DBConnect))
            {
                connect.Open();
                SqlDataAdapter adapter = new SqlDataAdapter("Select * from [balls], [students] where [balls].[Ocodstudent] = [students].[codstudent] AND [balls].[Osrball] > 4", connect);
                var data = new DataTable();
                adapter.Fill(data);
                dgStudentsFour.ItemsSource = data.DefaultView;
            }
        }

        private void ObtnSearchBD_Click(object sender, RoutedEventArgs e)
        {
            using (SqlConnection connect = new SqlConnection(Properties.Settings.Default.DBConnect))
            {
                connect.Open();
                SqlDataAdapter adapter = new SqlDataAdapter("Select * from [students] where [students].[bd] Between '" + OdpOne.Text + "' and '" + OdpTwo.Text + "'", connect);
                var data = new DataTable();
                adapter.Fill(data);
                dgStudentsBD.ItemsSource = data.DefaultView;
            }
        }

        private void ObtnSearchSpeci_Click(object sender, RoutedEventArgs e)
        {
            using (SqlConnection connect = new SqlConnection(Properties.Settings.Default.DBConnect))
            {
                connect.Open();
                SqlDataAdapter adapter = new SqlDataAdapter("Select * from [students] where [students].[groupp] = '" + OtxtSpeci.Text + "'", connect);
                var data = new DataTable();
                adapter.Fill(data);
                dgStudentsSpec.ItemsSource = data.DefaultView;
            }
        }

        private void ObtnGenFour_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            app.Visible = true;
            app.WindowState = XlWindowState.xlMaximized;

            Workbook wb = app.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            Worksheet ws = wb.Worksheets[1];
            DateTime currentDate = DateTime.Now;
            ws.Columns.AutoFit();

            for (int j = 0; j < dgStudentsFour.Columns.Count; j++)
            {
                Range range = (Range)ws.Cells[1, j + 1];
                ws.Cells[1, j + 1].font.bold = true;
                ws.Cells[1, j + 1].columnwidth = 15;
                range.Value2 = dgStudentsFour.Columns[j].Header;
            }

            for (int i = 0; i < dgStudentsFour.Columns.Count; i++)
            {
                for (int j = 0; j < dgStudentsFour.Items.Count; j++)
                {
                    TextBlock text = dgStudentsFour.Columns[i].GetCellContent(dgStudentsFour.Items[j]) as TextBlock;
                    Range range = (Range)ws.Cells[j + 2, i + 1];
                    range.Value2 = text.Text;
                }
            }
        }

        private void ObtnGenBD_Click (object sender, RoutedEventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            app.Visible = true;
            app.WindowState = XlWindowState.xlMaximized;

            Workbook wb = app.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            Worksheet ws = wb.Worksheets[1];
            DateTime currentDate = DateTime.Now;
            ws.Columns.AutoFit();

            for (int j = 0; j < dgStudentsBD.Columns.Count; j++)
            {
                Range range = (Range)ws.Cells[1, j + 1];
                ws.Cells[1, j + 1].font.bold = true;
                ws.Cells[1, j + 1].columnwidth = 15;
                range.Value2 = dgStudentsBD.Columns[j].Header;
            }

            for (int i = 0; i < dgStudentsBD.Columns.Count; i++)
            {
                for (int j = 0; j < dgStudentsBD.Items.Count; j++)
                {
                    TextBlock text = dgStudentsBD.Columns[i].GetCellContent(dgStudentsBD.Items[j]) as TextBlock;
                    Range range = (Range)ws.Cells[j + 2, i + 1];
                    range.Value2 = text.Text;
                }
            }
        }

        private void ObtnGenSpeci_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            app.Visible = true;
            app.WindowState = XlWindowState.xlMaximized;

            Workbook wb = app.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            Worksheet ws = wb.Worksheets[1];
            DateTime currentDate = DateTime.Now;
            ws.Columns.AutoFit();

            for (int j = 0; j < dgStudentsSpec.Columns.Count; j++)
            {
                Range range = (Range)ws.Cells[1, j + 1];
                ws.Cells[1, j + 1].font.bold = true;
                ws.Cells[1, j + 1].columnwidth = 15;
                range.Value2 = dgStudentsSpec.Columns[j].Header;
            }

            for (int i = 0; i < dgStudentsSpec.Columns.Count; i++)
            {
                for (int j = 0; j < dgStudentsSpec.Items.Count; j++)
                {
                    TextBlock text = dgStudentsSpec.Columns[i].GetCellContent(dgStudentsSpec.Items[j]) as TextBlock;
                    Range range = (Range)ws.Cells[j + 2, i + 1];
                    range.Value2 = text.Text;
                }
            }
        }
    }
}
