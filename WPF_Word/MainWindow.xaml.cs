using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.WebSockets;
using System.Reflection;
using System.Reflection.Metadata;
using System.Security.Cryptography.X509Certificates;
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
using System.Xml.Linq;
using static System.Net.Mime.MediaTypeNames;
using Word = Microsoft.Office.Interop.Word;
using System.Text.Json;
using System.IO;
using Newtonsoft.Json;
using CheckBox = System.Windows.Controls.CheckBox;

namespace WPF_Word
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        public MainWindow()
        {
            InitializeComponent();
            tb_ResumeDate.Text = DateTime.Now.ToString("dd.MM.yyyy");


        }

        private void btn_CreateWord_Click(object sender, RoutedEventArgs e)
        {
            // Заполняем текст по разделам___________________________

            string b1_Text, b2_Text, b3_Text, b4_Text, b5_Text, b6_Text, b7_Text, b8_Text, b9_Text, b10_Text;
            int j;

            // Блок 1
            b1_Text = $"ФИО: {tb_FIO.Text} \r\nВозраст: {tb_Age.Text} \r\n" +
                $"Дата рождения: {tb_BirthDate.Text}\r\n" +
                $"Дата обследования: {tb_ResumeDate.Text}\r\n" +
                $"Краткий анамнез со слов мамы: {tb_Block1_T1.Text} \r\n" +
                $"Роды: {tb_Block1_T2.Text} \r\n" +
                $"Моторное развитие: {tb_Block1_T3.Text}\r\n" +
                $"Речевое развитие: {tb_Block1_T4.Text}\r\n" +
                $"Состав семьи: {tb_Block1_T5.Text}\r\n" +
                $"Социальная среда: {tb_Block1_T6.Text}";

            // Блок 2
            //var Checkboxes_b2 = StackPanel_b2.Children.OfType<System.Windows.Controls.CheckBox>().ToList();
            
            
            b2_Text = "";
            j = 2;

            for (int i = 1; i <= 10; i++)
            {
                object wantedCbNode = mainGrid.FindName($"cb_Block{j}_T{i}");
                CheckBox wantedCb = wantedCbNode as CheckBox;

                object wantedTbNode = mainGrid.FindName($"tb_Block{j}_T{i}");
                TextBox wantedTb = wantedTbNode as TextBox;

                if ((bool)wantedCb.IsChecked)
                {
                    b2_Text += wantedTb.Text + "\n";
                }

                
            }




            /*foreach (System.Windows.Controls.CheckBox box in Checkboxes_b2)
            {
                if (box.IsChecked == true)
                {
                    b2_Text += box.Content + "\n";
                }
            }*/

            // Блок 3
            b3_Text = "";
            j = 3;

            for (int i = 1; i <= 10; i++)
            {
                object wantedCbNode = mainGrid.FindName($"cb_Block{j}_T{i}");
                CheckBox wantedCb = wantedCbNode as CheckBox;

                object wantedTbNode = mainGrid.FindName($"tb_Block{j}_T{i}");
                TextBox wantedTb = wantedTbNode as TextBox;

                if ((bool)wantedCb.IsChecked)
                {
                    b3_Text += wantedTb.Text + "\n";
                }


            }

            // Блок 4
            b4_Text = "";
            j = 4;

            for (int i = 1; i <= 10; i++)
            {
                object wantedCbNode = mainGrid.FindName($"cb_Block{j}_T{i}");
                CheckBox wantedCb = wantedCbNode as CheckBox;

                object wantedTbNode = mainGrid.FindName($"tb_Block{j}_T{i}");
                TextBox wantedTb = wantedTbNode as TextBox;

                if ((bool)wantedCb.IsChecked)
                {
                    b4_Text += wantedTb.Text + "\n";
                }


            }

            // Блок 5
            b5_Text = "";
            j = 5;

            for (int i = 1; i <= 10; i++)
            {
                object wantedCbNode = mainGrid.FindName($"cb_Block{j}_T{i}");
                CheckBox wantedCb = wantedCbNode as CheckBox;

                object wantedTbNode = mainGrid.FindName($"tb_Block{j}_T{i}");
                TextBox wantedTb = wantedTbNode as TextBox;

                if ((bool)wantedCb.IsChecked)
                {
                    b5_Text += wantedTb.Text + "\n";
                }


            }

            // Блок 6
            b6_Text = "";
            j = 6;

            for (int i = 1; i <= 10; i++)
            {
                object wantedCbNode = mainGrid.FindName($"cb_Block{j}_T{i}");
                CheckBox wantedCb = wantedCbNode as CheckBox;

                object wantedTbNode = mainGrid.FindName($"tb_Block{j}_T{i}");
                TextBox wantedTb = wantedTbNode as TextBox;

                if ((bool)wantedCb.IsChecked)
                {
                    b6_Text += wantedTb.Text + "\n";
                }


            }

            // Блок 7
            b7_Text = "";
            j = 7;

            for (int i = 1; i <= 10; i++)
            {
                object wantedCbNode = mainGrid.FindName($"cb_Block{j}_T{i}");
                CheckBox wantedCb = wantedCbNode as CheckBox;

                object wantedTbNode = mainGrid.FindName($"tb_Block{j}_T{i}");
                TextBox wantedTb = wantedTbNode as TextBox;

                if ((bool)wantedCb.IsChecked)
                {
                    b7_Text += wantedTb.Text + "\n";
                }


            }

            // Блок 8
            b8_Text = "";
            j = 8;

            for (int i = 1; i <= 10; i++)
            {
                object wantedCbNode = mainGrid.FindName($"cb_Block{j}_T{i}");
                CheckBox wantedCb = wantedCbNode as CheckBox;

                object wantedTbNode = mainGrid.FindName($"tb_Block{j}_T{i}");
                TextBox wantedTb = wantedTbNode as TextBox;

                if ((bool)wantedCb.IsChecked)
                {
                    b8_Text += wantedTb.Text + "\n";
                }


            }

            // Блок 9
            b9_Text = "";
            j = 9;

            for (int i = 1; i <= 10; i++)
            {
                object wantedCbNode = mainGrid.FindName($"cb_Block{j}_T{i}");
                CheckBox wantedCb = wantedCbNode as CheckBox;

                object wantedTbNode = mainGrid.FindName($"tb_Block{j}_T{i}");
                TextBox wantedTb = wantedTbNode as TextBox;

                if ((bool)wantedCb.IsChecked)
                {
                    b9_Text += wantedTb.Text + "\n";
                }


            }

            // Блок 10
            b10_Text = tb_Block10_T1.Text;

            // Работа с Word__________________________________________ 

            object oMissing = Missing.Value;
            object templatePathObj = Environment.CurrentDirectory + "\\Template.dotx";
            object falseObj = false;

            Word._Application oWord;
            Word._Document oDoc;
            oWord = new Word.Application();

            try
            {
                oDoc = oWord.Documents.Add(ref templatePathObj, ref oMissing, ref oMissing, ref oMissing);
                oWord.Visible = true;

                


                //Заполняем документ по закладкам

                Word.Range wrdRng1 = oDoc.Bookmarks.get_Item("b1").Range;
                wrdRng1.Text = b1_Text;

                Word.Range wordRange2 = oDoc.Bookmarks.get_Item("b2").Range;
                wordRange2.Text = b2_Text;

                Word.Range wordRange3 = oDoc.Bookmarks.get_Item("b3").Range;
                wordRange3.Text = b3_Text;

                Word.Range wordRange4 = oDoc.Bookmarks.get_Item("b4").Range;
                wordRange4.Text = b4_Text;

                Word.Range wordRange5 = oDoc.Bookmarks.get_Item("b5").Range;
                wordRange5.Text = b5_Text;

                Word.Range wordRange6 = oDoc.Bookmarks.get_Item("b6").Range;
                wordRange6.Text = b6_Text;

                Word.Range wordRange7 = oDoc.Bookmarks.get_Item("b7").Range;
                wordRange7.Text = b7_Text;

                Word.Range wordRange8 = oDoc.Bookmarks.get_Item("b8").Range;
                wordRange8.Text = b8_Text;

                Word.Range wordRange9 = oDoc.Bookmarks.get_Item("b9").Range;
                wordRange9.Text = b9_Text;

                Word.Range wordRange10 = oDoc.Bookmarks.get_Item("b10").Range;
                wordRange10.Text = b10_Text;

                Word.Range wordRange11 = oDoc.Bookmarks.get_Item("Date").Range;
                wordRange11.Text = tb_ResumeDate.Text;

                Word.Window wordWindow = oWord.ActiveWindow;
                wordWindow.SetFocus();
                wordWindow.Activate();

            }

            catch (Exception err)
            {

                oWord.Quit(ref oMissing, ref oMissing, ref oMissing);
                oDoc = null;
                oWord = null;
                MessageBox.Show("Ошибка работы с шаблоном MS Word: \r\n\r\n" + err.ToString(), "Ошибка!");
                throw;
            }

        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            LoadJson();
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            SaveJson();
        }

        public void LoadJson()
        {
            List<UiTbItem> tbItemList = new List<UiTbItem>();

            tbItemList.Clear();
            string jsonData = File.ReadAllText("text_data.json");
            tbItemList = JsonConvert.DeserializeObject<List<UiTbItem>>(jsonData);

            foreach (UiTbItem item in tbItemList)
            {
                object wantedNode = mainGrid.FindName(item.ElementName);
                TextBox wantedChild = wantedNode as TextBox;
                wantedChild.Text = item.Text;

            }
        }

        public void SaveJson()
        {
            List<UiTbItem> tbItemList = new List<UiTbItem>();
            tbItemList.Clear();

            for (int j = 2; j <= 9; j++)
            {
                for (int i = 1; i <= 10; i++)
                {
                    object wantedNode = mainGrid.FindName($"tb_Block{j}_T{i}");
                    if (wantedNode is TextBox)
                    {
                        TextBox wantedChild = wantedNode as TextBox;
                        tbItemList.Add(new UiTbItem(wantedChild.Name, wantedChild.Text));
                    }
                }
            }

            File.WriteAllText("text_data.json", JsonConvert.SerializeObject(tbItemList, Formatting.Indented));

        }

    }
}
