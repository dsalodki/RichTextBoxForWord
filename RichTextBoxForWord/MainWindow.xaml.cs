using Microsoft.Win32;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using static System.Net.Mime.MediaTypeNames;
using Application = Microsoft.Office.Interop.Word.Application;

namespace RichTextBoxForWord
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();

        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog() { ValidateNames = true, Multiselect = false, Filter = "Word Doucment|*.docx|Word 97 - 2003 Document|*.doc" };
            if (ofd.ShowDialog() == true)
            {
                // Open document 
                string originalfilename = System.IO.Path.GetFullPath(ofd.FileName);

                if (ofd.CheckFileExists && new[] { ".docx", ".doc", ".txt", ".rtf" }.Contains(System.IO.Path.GetExtension(originalfilename).ToLower()))
                {
                    Microsoft.Office.Interop.Word.Application wordObject = new Microsoft.Office.Interop.Word.Application();
                    object File = originalfilename;
                    object nullobject = System.Reflection.Missing.Value;
                    Microsoft.Office.Interop.Word.Application wordobject = new Microsoft.Office.Interop.Word.Application();
                    wordobject.DisplayAlerts = Microsoft.Office.Interop.Word.WdAlertLevel.wdAlertsNone;
                    Microsoft.Office.Interop.Word._Document docs = wordObject.Documents.Open(ref File, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject);
                    docs.ActiveWindow.Selection.WholeStory();
                    docs.ActiveWindow.Selection.Copy();
                    rtfMain.Paste();
                    docs.Close(ref nullobject, ref nullobject, ref nullobject);
                    wordobject.Quit(ref nullobject, ref nullobject, ref nullobject);


                    MessageBox.Show("file loaded");
                }
            }
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            Application application = new Application();
            object File = "D:\\DOCs\\testRu.docx"; //this is the path
            object nullobject = System.Reflection.Missing.Value;
            application.DisplayAlerts = Microsoft.Office.Interop.Word.WdAlertLevel.wdAlertsNone; 
            Microsoft.Office.Interop.Word._Document docs =
            application.Documents.Open(ref File, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject); 
            docs.ActiveWindow.Selection.WholeStory();
            docs.ActiveWindow.Selection.Copy();


            MemoryStream stream = new MemoryStream(ASCIIEncoding.Default.GetBytes(Clipboard.GetText(TextDataFormat.Rtf)));
            rtfMain.Selection.Load(stream, DataFormats.Rtf);

            //TextRange tr = new TextRange(rtfMain.Document.ContentStart, rtfMain.Document.ContentEnd);

            //var bytes = UTF8Encoding.UTF8.GetBytes(Clipboard.GetText(TextDataFormat.Rtf));
            //MemoryStream ms = new MemoryStream(bytes);
            //tr.Save(ms, DataFormats.Rtf);
            //string xamlText = UTF8Encoding.UTF8.GetString(ms.ToArray());

            //Clipboard.SetText(docs.ActiveWindow.Selection.Text, TextDataFormat.Rtf);

            //rtfMain.Paste();

            //rtfMain.AppendText(Clipboard.GetText(TextDataFormat.Rtf));

            //rtfMain.Document.Blocks.Clear();
            //rtfMain.Document.Blocks.Add(new Paragraph(new Run(Clipboard.GetText(TextDataFormat.Rtf))));

            docs.Close(ref nullobject, ref nullobject, ref nullobject);
            application.Quit(ref nullobject, ref nullobject, ref nullobject);
        }
    }
}
