using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Win32;
using Html2Xaml;

namespace DocumentEditorTestApp
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

        private void menuItemOpenFile_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            if(dialog.ShowDialog() == true)
            {
                //docEditor.OpenDocxFile(dialog.FileName);
                FlowDocument fd = new FlowDocument();
                FileStream fileStream = File.OpenRead(dialog.FileName);
                fd.LoadOpenXml(fileStream);
                docEditor.rtbDocument.Document = fd;
            }
        }

        private void menuItemSaveFile_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog dialog = new SaveFileDialog();
            if(dialog.ShowDialog() == true)
            {
                FileStream saveFileStream = dialog.OpenFile() as FileStream;
                //docEditor.ConvertToOpenXml(saveFileStream);
                docEditor.rtbDocument.Document.ConvertToOpenXml(saveFileStream);
                //docEditor.rtbDocument.Document.SaveToPdf(dialog.FileName);
            }
        }

        private void menuItemOpenHtml_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            if (dialog.ShowDialog() == true)
            {
                string htmlContent = File.ReadAllText(dialog.FileName);
                string xamlContent = HTMLConverter.HtmlToXamlConverter.ConvertHtmlToXaml(htmlContent, true);
                FlowDocument flowDoc = XamlReader.Parse(xamlContent) as FlowDocument;
                this.docEditor.rtbDocument.Document = flowDoc;
            }
        }

        private void menuItemSaveHtml_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog dialog = new SaveFileDialog();
            if (dialog.ShowDialog() == true)
            {
                string xamlContent = XamlWriter.Save(this.docEditor.rtbDocument.Document);
                string htmlContent = HTMLConverter.HtmlFromXamlConverter.ConvertXamlToHtml(xamlContent);
                string formattedHtml = FormatHtml(htmlContent);
                File.WriteAllText(dialog.FileName, formattedHtml, Encoding.UTF8);
            }
        }

        public static string FormatHtml(string nativeTags)
        {
            string upperTags = nativeTags;//.Replace("<META content=\"MSHTML 6.00.2900.2180\" name=GENERATOR>",
                //"");

            StringBuilder strBuilder = new StringBuilder(upperTags);

            strBuilder.Replace("<HTML", "<html");
            strBuilder.Replace("<html>", "<html>\n");
            strBuilder.Replace("</HTML>", "\n</html>");
            strBuilder.Replace("<BODY", "<body");
            strBuilder.Replace("</BODY>", "\n</body>");
            strBuilder.Replace("<HEAD", "<head");
            strBuilder.Replace("</HEAD>", "</head>");
            strBuilder.Replace("<TITLE", "<title");
            strBuilder.Replace("</TITLE>", "</title>");
            strBuilder.Replace("<TABLE", "<table");
            strBuilder.Replace("</TABLE>", "</table>");
            strBuilder.Replace("<TBODY", "<tbody");
            strBuilder.Replace("</TBODY>", "</tbody>");
            strBuilder.Replace("<DIV", "<div");
            strBuilder.Replace("</DIV>", "</div>");
            strBuilder.Replace("<P", "<p");
            strBuilder.Replace("</P>", "</p>\n");
            strBuilder.Replace("<A", "<a");
            strBuilder.Replace("</A>", "</a>");
            strBuilder.Replace("<UL", "<ul");
            strBuilder.Replace("</UL>", "</ul>");
            strBuilder.Replace("<OL", "<ol");
            strBuilder.Replace("</OL>", "</ol>");
            strBuilder.Replace("<LI", "<li");
            strBuilder.Replace("</LI>", "</li>");
            strBuilder.Replace("<TR", "<tr");
            strBuilder.Replace("</TR>", "</tr>");
            strBuilder.Replace("<TD", "<td");
            strBuilder.Replace("</TD>", "</td>");
            strBuilder.Replace("<HR", "<hr");
            strBuilder.Replace("</HR>", "</hr>");
            strBuilder.Replace("<IMG", "<img");
            strBuilder.Replace("</IMG>", "</img>");
            strBuilder.Replace("<H", "<h");
            strBuilder.Replace("<BR>", "<br />\n");
            strBuilder.Replace("</H1>", "</h1>");
            strBuilder.Replace("</H2>", "</h2>");
            strBuilder.Replace("</H3>", "</h3>");
            strBuilder.Replace("</H4>", "</h4>");
            strBuilder.Replace("</H5>", "</h5>");
            strBuilder.Replace("</H6>", "</h6>");
            strBuilder.Replace("<FORM", "<form");
            strBuilder.Replace("</FORM>", "\n</form>");
            strBuilder.Replace("<META", "<meta");
            strBuilder.Replace("<INPUT", "<input");
            strBuilder.Replace("</INPUT>", "</input>");
            strBuilder.Replace("<STYLE", "<style");
            strBuilder.Replace("</STYLE>", "</style>");
            strBuilder.Replace("<SPAN", "<span");
            strBuilder.Replace("</SPAN>", "</span>");
            strBuilder.Replace("STYLE=", "style=");
            strBuilder.Replace("FONT-WEIGHT:", "font-weight:");
            strBuilder.Replace("FONT-FAMILY:", "font-family:");
            strBuilder.Replace("FONT-STYLE:", "font-style:");
            strBuilder.Replace("FONT-WEIGHT:", "font-weight:");
            strBuilder.Replace("FONT-FAMILY:", "font-family:");
            strBuilder.Replace("FONT-SIZE:", "font-size:");
            strBuilder.Replace("TEXT-ALIGN:", "text-align:");
            strBuilder.Replace("BORDER-WIDTH:", "border-width:");
            strBuilder.Replace("POSITION:", "position:");
            strBuilder.Replace("TEXT-DECORATION:", "text-decoration:");
            strBuilder.Replace("BORDER-TOP-WIDTH:", "border-top-width:");
            strBuilder.Replace("BORDER-LEFT-WIDTH:", "border-left-width:");
            strBuilder.Replace("BORDER-BOTTOM-WIDTH:", "border-bottom-width:");
            strBuilder.Replace("BORDER-RIGHT-WIDTH:", "border-right-width:");
            strBuilder.Replace("BORDER-TOP-COLOR:", "border-top-color:");
            strBuilder.Replace("BORDER-LEFT-COLOR:", "border-left-color:");
            strBuilder.Replace("BORDER-BOTTOM-COLOR:", "border-bottom-color:");
            strBuilder.Replace("BORDER-RIGHT-COLOR:", "border-right-color:");
            strBuilder.Replace("BORDER-TOP:", "border-top:");
            strBuilder.Replace("BORDER-LEFT:", "border-left:");
            strBuilder.Replace("BORDER-BOTTOM:", "border-bottom:");
            strBuilder.Replace("BORDER-RIGHT:", "border-right:");
            strBuilder.Replace("BACKGROUND-COLOR:", "background-color:");
            strBuilder.Replace("COLOR:", "color:");
            strBuilder.Replace("LIST-STYLE-TYPE:", "list-style-type:");
            strBuilder.Replace("LEFT:", "left:");
            strBuilder.Replace("TOP:", "top:");
            strBuilder.Replace("WIDTH:", "width:");
            strBuilder.Replace("HEIGHT:", "height:");
            strBuilder.Replace("about:blank", "");
            strBuilder.Replace("<body>", "<body>\n");
            if (!strBuilder.ToString().Contains("<style"))
            {
                strBuilder.Replace("<head>", "<head>\n<style type=\"text/css\">\n\n</style>\n");
            }

            if (strBuilder.ToString().Contains("<li"))
            {
                strBuilder.Insert(strBuilder.ToString().IndexOf("<li>"), "</li>");
                strBuilder.Remove(strBuilder.ToString().IndexOf("</li>"), 5);
                strBuilder.Replace("</ol>", "\n</ol>");
                strBuilder.Replace("</ul>", "\n</ul>");
            }

            if (strBuilder.ToString().Contains("<table"))
            {
                strBuilder.Replace("</tr>", "\n</tr>");
                strBuilder.Replace("</tbody>", "\n</tbody>");
                strBuilder.Replace("</table>", "\n</table>");
            }

            return strBuilder.ToString();
        }
    }
}
