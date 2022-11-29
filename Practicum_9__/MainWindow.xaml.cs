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
using Word = Microsoft.Office.Interop.Word;
using System.IO;

namespace Practicum_9__
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }
        Random rnd = new Random();

        private void CreateTableClick(object sender, RoutedEventArgs e)
        {
            //creating a new Word document
            Word.Application application = new Word.Application();
            object oMissing = System.Reflection.Missing.Value;
            Word.Document document = application.Documents.Add(ref oMissing, ref oMissing, ref oMissing, ref oMissing);

            try
            {
                string[] towns = File.ReadAllLines(@"C:\Users\904\Desktop\towns1.txt", Encoding.UTF8); //file to fill in the Word cell
                application.Visible = true; // visibility of the object for reading and writing
                application.Selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter; 
                application.Selection.TypeText("Towns of Russia");
                Word.Table table = document.Tables.Add(application.Selection.Range, new Random().Next(1, 50), new Random().Next(1, 6)); // creating table
                table.Borders.Enable = 1; // displaying table lines
                // filling the table with values
                for (int i = 0; i < table.Rows.Count; i++)
                {
                    for (int j = 0; j < table.Columns.Count; j++)
                    {
                        table.Cell(i + 1, j + 1).Select();
                        application.Selection.TypeText(towns[rnd.Next(towns.Length)]);
                    }
                }
                // the style applied to the text
                object unit = Word.WdUnits.wdStory;
                object extend = Word.WdMovementType.wdExtend;
                application.Selection.HomeKey(ref unit, ref extend);
                application.Selection.Font.Size = 14;
                application.Selection.Font.Name = "Times New Roman";
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                Task.Delay(30000).Wait();
                application.Quit();
            }
        }

        private void ExitClick(object sender, RoutedEventArgs e)
        {
            const string message = "Are you sure you want to close the application?";
            const string caption = "Form closing";
            var result = MessageBox.Show(message, caption, MessageBoxButton.YesNo, MessageBoxImage.Question);
            if (result == MessageBoxResult.Yes)
                Environment.Exit(0);
            
        }
    }
}
