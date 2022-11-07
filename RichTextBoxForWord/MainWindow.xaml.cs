using Microsoft.Win32;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Shapes;
using System.Xml.Linq;
using static System.Net.Mime.MediaTypeNames;
using Application = Microsoft.Office.Interop.Word.Application;
using Document = Microsoft.Office.Interop.Word.Document;

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
                    Application wordObject = new Application();
                    object File = originalfilename;
                    object nullobject = System.Reflection.Missing.Value;
                    Application wordobject = new Application();
                    wordobject.DisplayAlerts = Microsoft.Office.Interop.Word.WdAlertLevel.wdAlertsNone;
                    Document docs = wordObject.Documents.Open(ref File, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject);
                    docs.ActiveWindow.Selection.WholeStory();
                    docs.ActiveWindow.Selection.Copy();

                    MemoryStream stream = new MemoryStream(Encoding.Default.GetBytes(Clipboard.GetText(TextDataFormat.Rtf)));
                    rtfMain.Selection.Load(stream, DataFormats.Rtf);

                    docs.Close(ref nullobject, ref nullobject, ref nullobject);
                    wordobject.Quit(ref nullobject, ref nullobject, ref nullobject);
                }
            }
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            var fileName = "result.rtf";
            using(var fs = new FileStream(fileName, FileMode.Create, FileAccess.Write))
            {
                var rtbText = new TextRange(rtfMain.Document.ContentStart, rtfMain.Document.ContentEnd);
                rtbText.Save(fs, DataFormats.Rtf);
            }
        }
    }
}
