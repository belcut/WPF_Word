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
            tb_Block8_T1.Text = "На фоне хорошо развитого зрительного восприятия, мыслительных операций, выявлены следующие трудности:\r\n- снижение нейродинамических характеристик, \r\n- несформированность функции программирования и контроля, \r\n- несформированность мнестических, гностических и пространственных функций,\r\n- а также недостаточность межполушарного взаимодействия и двигательной сферы.\r\n";
        }

        private void btn_CreateWord_Click(object sender, RoutedEventArgs e)
        {
            // Заполняем текст по разделам___________________________
            
            string b1_Text, b2_Text, b3_Text, b4_Text, b5_Text, b6_Text, b71_Text, b72_Text, b8_Text, b9_Text;

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
            var Checkboxes_b2 = StackPanel_b2.Children.OfType<System.Windows.Controls.CheckBox>().ToList();
            b2_Text = "";
            foreach (System.Windows.Controls.CheckBox box in Checkboxes_b2)
            {
                if (box.IsChecked == true) 
                {
                    b2_Text += box.Content + "\n";
                }
            }

            // Блок 3
            var Checkboxes_b3 = StackPanel_b3.Children.OfType<System.Windows.Controls.CheckBox>().ToList();
            b3_Text = "";
            foreach (System.Windows.Controls.CheckBox box in Checkboxes_b3)
            {
                if (box.IsChecked == true)
                {
                    b3_Text += box.Content + "\n";
                }
            }

            // Блок 4
            var Checkboxes_b4 = StackPanel_b4.Children.OfType<System.Windows.Controls.CheckBox>().ToList();
            b4_Text = "";
            foreach (System.Windows.Controls.CheckBox box in Checkboxes_b4)
            {
                if (box.IsChecked == true)
                {
                    b4_Text += box.Content + "\n";
                }
            }

            // Блок 5
            var Checkboxes_b5 = StackPanel_b5.Children.OfType<System.Windows.Controls.CheckBox>().ToList();
            b5_Text = "";
            foreach (System.Windows.Controls.CheckBox box in Checkboxes_b5)
            {
                if (box.IsChecked == true)
                {
                    b5_Text += box.Content + "\n";
                }
            }

            // Блок 6
            var Checkboxes_b6 = StackPanel_b6.Children.OfType<System.Windows.Controls.CheckBox>().ToList();
            b6_Text = "";
            foreach (System.Windows.Controls.CheckBox box in Checkboxes_b6)
            {
                if (box.IsChecked == true)
                {
                    b6_Text += box.Content + "\n";
                }
            }
            // Работа с Word__________________________________________ 

            object oMissing = Missing.Value;
            object templatePathObj = Environment.CurrentDirectory + "\\Template.dotm" ;
            object falseObj = false;
            
            Word._Application oWord;
            Word._Document oDoc;
            oWord = new Word.Application();

            try
            {
                oDoc = oWord.Documents.Add(ref templatePathObj, ref oMissing, ref oMissing, ref oMissing);
            }
            catch (Exception error)
            {
                //oDoc.Close(ref falseObj, ref oMissing, ref oMissing);
                oWord.Quit(ref oMissing, ref oMissing, ref oMissing);
                oDoc = null;
                oWord = null;
                throw error;
            }
            
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

        }

    }
}
