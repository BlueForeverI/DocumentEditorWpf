using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.IO.Packaging;
using System.Xml;
using System.Xml.Linq;

namespace DocumentEditorTestApp
{
    /// <summary>
    /// Interaction logic for DocumentEditor.xaml
    /// </summary>
    public partial class DocumentEditor : UserControl
    {
        public DocumentEditor()
        {
            InitializeComponent();
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            this.comboFontFamily.ItemsSource = System.Windows.Media.Fonts.SystemFontFamilies;
            this.comboFontSize.ItemsSource = new double[] 
            { 
                3.0, 4.0, 5.0, 6.0, 6.5, 7.0, 7.5, 8.0, 8.5, 9.0, 9.5, 
                10.0, 10.5, 11.0, 11.5, 12.0, 12.5, 13.0, 13.5, 14.0, 15.0,
                16.0, 17.0, 18.0, 19.0, 20.0, 22.0, 24.0, 26.0, 28.0, 30.0,
                32.0, 34.0, 36.0, 38.0, 40.0, 44.0, 48.0, 52.0, 56.0, 60.0, 64.0, 68.0, 72.0, 76.0,
                80.0, 88.0, 96.0, 104.0, 112.0, 120.0, 128.0, 136.0, 144.0
            };
        }

        private void rtbDocument_SelectionChanged(object sender, RoutedEventArgs e)
        {
            UpdateItemCheckedState(btnBold, TextElement.FontWeightProperty, FontWeights.Bold);
            UpdateItemCheckedState(btnItalic, TextElement.FontStyleProperty, FontStyles.Italic);
            UpdateItemCheckedState(btnUnderline, Inline.TextDecorationsProperty, TextDecorations.Underline);

            UpdateItemCheckedState(btnLeftAlign, Paragraph.TextAlignmentProperty, TextAlignment.Left);
            UpdateItemCheckedState(btnRightAlign, Paragraph.TextAlignmentProperty, TextAlignment.Right);
            UpdateItemCheckedState(btnCenterAlign, Paragraph.TextAlignmentProperty, TextAlignment.Center);
            UpdateItemCheckedState(btnJustifyAlign, Paragraph.TextAlignmentProperty, TextAlignment.Justify);

            UpdateSelectionListType();
            UpdateSelectedFontFamily();
            UpdateSelectedFontSize();
            UpdateForegroundColor();
        }

        private void UpdateItemCheckedState(ToggleButton button, DependencyProperty formattingProperty, object expectedValue)
        {
            object currentValue = rtbDocument.Selection.GetPropertyValue(formattingProperty);
            button.IsChecked = (currentValue == DependencyProperty.UnsetValue) ? false : currentValue != null && currentValue.Equals(expectedValue);
        }

        private void UpdateSelectionListType()
        {
            Paragraph startParagraph = rtbDocument.Selection.Start.Paragraph;
            Paragraph endParagraph = rtbDocument.Selection.End.Paragraph;

            if(startParagraph != null && endParagraph != null 
                && (startParagraph.Parent is ListItem) 
                && (endParagraph.Parent is ListItem) 
                && object.ReferenceEquals(((ListItem)startParagraph.Parent).List, ((ListItem)endParagraph.Parent).List))
            {
                TextMarkerStyle markerStyle = ((ListItem) startParagraph.Parent).List.MarkerStyle;
                if(markerStyle == TextMarkerStyle.Disc)
                {
                    btnBullets.IsChecked = true;
                }
                else if(markerStyle == TextMarkerStyle.Decimal)
                {
                    btnNumbering.IsChecked = true;
                }
            }
            else
            {
                btnBullets.IsChecked = false;
                btnNumbering.IsChecked = false;
            }
        }

        private void comboFontFamily_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            FontFamily fontFamily = (FontFamily) e.AddedItems[0];
            ApplyPropertyToText(TextElement.FontFamilyProperty, fontFamily);
        }

        private void comboFontSize_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ApplyPropertyToText(TextElement.FontSizeProperty, e.AddedItems[0]);
        }

        private void ApplyPropertyToText(DependencyProperty property, object value)
        {
            if(value == null)
            {
                return;
            }

            rtbDocument.Selection.ApplyPropertyValue(property, value);
        }

        private void UpdateSelectedFontFamily()
        {
            object value = rtbDocument.Selection.GetPropertyValue(TextElement.FontFamilyProperty);
            FontFamily currentFontFamily = (FontFamily) ((value == DependencyProperty.UnsetValue) ? null : value);
            if(currentFontFamily != null)
            {
                comboFontFamily.SelectedItem = currentFontFamily;
            }
        }

        private void UpdateSelectedFontSize()
        {
            object value = rtbDocument.Selection.GetPropertyValue(TextElement.FontSizeProperty);
            comboFontSize.SelectedValue = (value == DependencyProperty.UnsetValue) ? null : value;
        }

        private void colorPicker_SelectedColorChanged(object sender, RoutedPropertyChangedEventArgs<Color> e)
        {
            ApplyPropertyToText(TextElement.ForegroundProperty, new SolidColorBrush(colorPicker.SelectedColor));
        }

        private void UpdateForegroundColor()
        {
            object value = rtbDocument.Selection.GetPropertyValue(TextElement.ForegroundProperty);
            SolidColorBrush colorBrush = (SolidColorBrush) ((value == DependencyProperty.UnsetValue) ? null : value);
            if(colorBrush != null)
            {
                colorPicker.SelectedColor = colorBrush.Color;
            }
        }

        public void OpenDocxFile(string fileName)
        {
            XElement wordDoc = null;
            
            try
            {
                Package package = Package.Open(fileName);
                Uri documentUri = new Uri("/word/document.xml", UriKind.Relative);
                PackagePart documentPart = package.GetPart(documentUri);
                wordDoc = XElement.Load(new StreamReader(documentPart.GetStream()));
            }
            catch (Exception)
            {
                this.rtbDocument.Document.Blocks.Add(new Paragraph(new Run("Cannot open file!")));
                return;
            }

            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

            var paragraphs = from p in wordDoc.Descendants(w + "p")
                             select p;
            foreach (var p in paragraphs)
            {
                var style = from s in p.Descendants(w + "rPr")
                            select s;

                var parStyle = from s in p.Descendants(w + "pPr")
                               select s;

                var font = (from f in style.Descendants(w + "rFonts")
                            select f.FirstAttribute).FirstOrDefault();

                var size = (from s in style.Descendants(w + "sz")
                            select s.FirstAttribute).FirstOrDefault();

                var rgbColor = (from c in style.Descendants(w + "color")
                             select c.FirstAttribute).FirstOrDefault();

                var bold = (from b in style.Descendants(w + "b")
                           select b).FirstOrDefault();

                var italic = (from i in style.Descendants(w + "i")
                             select i).FirstOrDefault();

                var underline = (from u in style.Descendants(w + "u")
                                select u).FirstOrDefault();

                var alignment = (from a in parStyle.Descendants(w + "jc")
                                 select a).FirstOrDefault();

                var textElements = from t in p.Descendants(w + "t")
                                   select t;

                var list = (from l in parStyle.Descendants(w + "numPr")
                           select l).FirstOrDefault();

                

                StringBuilder text = new StringBuilder();
                foreach (var element in textElements)
                {
                    text.Append(element.Value);
                }

                Paragraph par = new Paragraph();
                Run run = new Run(text.ToString());

                List items = new List();
                if (list != null)
                {
                    //List items = new List();
                    items.ListItems.Add(new ListItem(par));
                    var markerStyleElement = (from l in list.Descendants(w + "numId")
                                       select l).FirstOrDefault();

                    TextMarkerStyle markerStyle = new TextMarkerStyle();
                    if(markerStyleElement != null)
                    {
                        int value = int.Parse(markerStyleElement.Attribute(w + "val").Value);
                        switch (value)
                        {
                            case 1:
                                markerStyle = TextMarkerStyle.Disc;
                                break;
                            case 2:
                                markerStyle = TextMarkerStyle.Decimal;
                                break;
                        }
                    }

                    items.MarkerStyle = markerStyle;
                }

                if(font != null)
                {
                    FontFamilyConverter converter = new FontFamilyConverter();
                    run.FontFamily = (FontFamily) converter.ConvertFrom(font.Value);
                }

                if(size != null)
                {
                    run.FontSize = double.Parse(size.Value);
                }

                if(rgbColor != null)
                {
                    Color color = ConvertRgbToColor(rgbColor.Value);
                    run.Foreground = new SolidColorBrush(color);
                }

                if(bold != null)
                {
                    if (bold.Attribute(w + "val") != null)
                    {
                        if (bold.Attribute(w + "val").Value == "off")
                        {
                            run.FontWeight = FontWeights.Normal;
                        }
                        else
                        {
                            run.FontWeight = FontWeights.Bold;
                        }
                    }
                    else
                    {
                        run.FontWeight = FontWeights.Bold;
                    }
                }

                if(italic != null)
                {
                    if (italic.Attribute(w + "val") != null)
                    {
                        if (italic.Attribute(w + "val").Value == "off")
                        {
                            run.FontStyle = FontStyles.Normal;
                        }
                        else
                        {
                            run.FontStyle = FontStyles.Italic;
                        }
                    }
                    else
                    {
                        run.FontStyle = FontStyles.Italic;
                    }
                }

                if(underline != null)
                {
                    run.TextDecorations.Add(TextDecorations.Underline);
                }

                if(alignment != null)
                {
                    TextAlignment textAlignment = new TextAlignment();
                    string value = alignment.Attribute(w + "val").Value;
                    switch (value)
                    {
                        case "left":
                            textAlignment = TextAlignment.Left;
                            break;

                        case "center":
                            textAlignment = TextAlignment.Center;
                            break;

                        case "right":
                            textAlignment = TextAlignment.Right;
                            break;

                        case "justify":
                            textAlignment = TextAlignment.Justify;
                            break;
                    }

                    par.TextAlignment = textAlignment;
                }
                else
                {
                    par.TextAlignment = TextAlignment.Left;
                }

                par.Inlines.Add(run);
                if(list != null)
                {
                    this.rtbDocument.Document.Blocks.Add(items);
                }
                else
                {
                    this.rtbDocument.Document.Blocks.Add(par);
                }
            }
        }

        private Color ConvertRgbToColor(string rgbColor)
        {
            string rgbValue = rgbColor;

            int redValue = Convert.ToInt32(rgbValue.Substring(0, 2), 16);
            int greenValue = Convert.ToInt32(rgbValue.Substring(2, 2), 16);
            int blueValue = Convert.ToInt32(rgbValue.Substring(4, 2), 16);

            return Color.FromRgb((byte)redValue, (byte)greenValue, (byte)blueValue);
        }

        public void ConvertToOpenXml(FlowDocument flowDoc, Stream openXmlStream)
        {
            TextPointer contentstart = flowDoc.ContentStart;
            TextPointer contentend = flowDoc.ContentEnd;
            if (contentstart == null)
            {
                throw new ArgumentNullException("ContentStart");
            }
            if (contentend == null)
            {
                throw new ArgumentNullException("ContentEnd");
            }

            //Create document

            // document package container
            Package zippackage = null;
            zippackage = Package.Open(openXmlStream, FileMode.Create, FileAccess.ReadWrite);

            // main document.xml 
            Uri uri = new Uri("/word/document.xml", UriKind.Relative);
            PackagePart partDocumentXML = zippackage.CreatePart(uri, "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml");


            //ver 1.2 Numbering
            Uri uriNumbering = new Uri("/word/numbering.xml", UriKind.Relative);
            Uri uriNumberingRelationship = new Uri("numbering.xml", UriKind.Relative);
            PackagePart partNumberingXML = zippackage.CreatePart(uriNumbering, "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml");
            partDocumentXML.CreateRelationship(uriNumberingRelationship, TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering", "rId1");



            using (XmlTextWriter openxmlwriter = new XmlTextWriter(partDocumentXML.GetStream(FileMode.Create, FileAccess.Write), System.Text.Encoding.UTF8))
            {
                openxmlwriter.Formatting = Formatting.Indented;
                openxmlwriter.Indentation = 2;
                openxmlwriter.IndentChar = ' ';

                //ver 1.2 
                using (XmlTextWriter numberingwriter = new XmlTextWriter(partNumberingXML.GetStream(FileMode.Create, FileAccess.Write), System.Text.Encoding.UTF8))
                {
                    numberingwriter.Formatting = Formatting.Indented;
                    numberingwriter.Indentation = 2;
                    numberingwriter.IndentChar = ' ';

                    //Actual Writing
                    new OpenXmlWriter().Write(contentstart, contentend, openxmlwriter, numberingwriter);
                }

            }



            zippackage.Flush();

            // relationship 
            zippackage.CreateRelationship(uri, TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument", "rId1");
            zippackage.Flush();
            zippackage.Close();
        }
    }
}
