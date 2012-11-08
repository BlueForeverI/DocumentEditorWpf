//////////////////////////////////////////////////////////////////////
//
// OpenXmlWriter.cs
//
//////////////////////////////////////////////////////////////////////

using System;
using System.Xml;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Controls;
using System.Windows.Media;

namespace DocumentEditorTestApp
{
    /// <summary>OpenXMLWriter</summary>
    class OpenXmlWriter
    {
        public OpenXmlWriter()
        {
            listLevel = 0;
            listContext = 0;
            lastListLevel = 0;

            
            //30 is more than enough
            listMarkerStyle = new TextMarkerStyle[30];
            int i;
            for (i=0;i<30;i++)
                listMarkerStyle[i] = TextMarkerStyle.Disc;

            
        }

        internal void Write(TextPointer contentstart, TextPointer contentend, XmlWriter xmlwriter, XmlWriter numberingwriter)
        {
            WriteOpenXML(contentstart, contentend, xmlwriter, numberingwriter);
        }


        void WriteOpenXML(TextPointer contentstart, TextPointer contentend, XmlWriter xmlwriter, XmlWriter numberingwriter)
        {
            TextElement textElement;         
            _openxmlwriter = xmlwriter;
            _numberingwriter = numberingwriter;
            WriteOpenXmlHeader();
            WriteNumberingHeader();
            WriteNumberingBody();
            
            _textpointer = contentstart;
            while (_textpointer.CompareTo(contentend) < 0)
            {
                TextPointerContext tpc = _textpointer.GetPointerContext(LogicalDirection.Forward);
                if (tpc==TextPointerContext.ElementStart)
                {                   
                        TextPointer position;
                        position = _textpointer;
                        position = position.GetNextContextPosition(LogicalDirection.Forward);
                        textElement = position.Parent as TextElement;

                        if (textElement is Paragraph)
                        {
                            CloseTextRunTag();
                            CreateParagraphTag();
                        }
                        else if (textElement is Inline)
                        {
                            CloseTextRunTag();
                            EnsureWithinParagraphTag();
                            CreateTextRunTag();
                        }
                        else if (textElement is List)
                        {                               
                            //keep track of list levels (add)                                
                            
                            List lx = (List)textElement;                            
                            TextMarkerStyle mst = (TextMarkerStyle) lx.GetValue(List.MarkerStyleProperty);                            
                            listMarkerStyle[listLevel]=mst;                            
                            listLevel++;


                            if (listLevel > 29)
                                listLevel = 29;                            
                           
                        }
                        else if (textElement is ListItem)
                        {
                            listContext = 1;

                        }
                       
                }
                else if  (tpc==TextPointerContext.ElementEnd)
                {
                 
                        textElement = _textpointer.Parent as TextElement;
                        if (textElement is Inline)
                        {
                            CloseTextRunTag();
                        }
                        else if (textElement is Paragraph)
                        {
                            CloseParagraphTag();
                        }
                        else if (textElement is List)
                        {
                            //keep track of list levels (subtract)                            
                            
                            
                            listLevel--;
                            if (listLevel < 0)
                                listLevel = 0;                             
                           
                        }
                        else if (textElement is ListItem)
                        {


                        }
                }
                else if  (tpc==TextPointerContext.Text)
                {
                        CreateTextRunTag();
                        _openxmlwriter.WriteStartElement("w:t");
                        _openxmlwriter.WriteString(_textpointer.GetTextInRun(LogicalDirection.Forward));
                        _openxmlwriter.WriteEndElement();
                   
               }
               else if  (tpc==TextPointerContext.EmbeddedElement)
               {
                        DependencyObject obj = _textpointer.GetAdjacentElement(LogicalDirection.Forward);
                        if (obj is LineBreak)
                        {
                            CreateTextRunTag();
                            _openxmlwriter.WriteStartElement("w:br");
                            _openxmlwriter.WriteEndElement();
                        }             
                        // other embedded objects here
               }        
               
                _textpointer = _textpointer.GetNextContextPosition(LogicalDirection.Forward);
            
            }//end while
             

            CloseTextRunTag();
            WriteNumberingFooter();
            WriteOpenXmlFooter();
        }

        void WriteParagraphProperties()
        {
            DependencyObject parent;
            TextPointer position;

            position = _textpointer;
            position = position.GetNextContextPosition(LogicalDirection.Forward);
            parent = position.Parent;                                   

            _openxmlwriter.WriteStartElement("w:pPr");

            //ver 1.2 Numbering
            if ((listLevel > 0) && (listContext!=0))
            {
                //Case : This is the 1st paragraph in a list item 
                WriteListProperty();

                //Add vertical spacing before this paragraph if its listlevel different from the previous list
                if (lastListLevel != listLevel)
                {
                    //Handle Vertical Spacing                 
                    _openxmlwriter.WriteStartElement("w:spacing");
                    _openxmlwriter.WriteAttributeString("w:before", "180");
                    _openxmlwriter.WriteEndElement();
                }                
                
                //can WordProcessingML handle the case when justification is "center" with bullet/numbering present
                //and with the bullet not centered with the text ?
                TextAlignment ta = (TextAlignment)parent.GetValue(Block.TextAlignmentProperty);
                WriteTextAlignmentProperty(ta);              

                listContext = 0;
                

            }            
            else if ((listLevel > 0) && (listContext == 0))
            {
                //Case : This is not the 1st paragraph in a list item 
                //but will still be indented 
                 int xx = 720 * listLevel;
                 _openxmlwriter.WriteStartElement("w:ind");
                 _openxmlwriter.WriteAttributeString("w:left", xx.ToString());
                 _openxmlwriter.WriteEndElement();

                 //Add vertical spacing before this paragraph if its listlevel different from the previous list
                 if (lastListLevel != listLevel)
                 {
                     //Handle Vertical Spacing                 
                     _openxmlwriter.WriteStartElement("w:spacing");
                     _openxmlwriter.WriteAttributeString("w:before", "180"); 
                     _openxmlwriter.WriteEndElement();
                 }

                 TextAlignment ta = (TextAlignment)parent.GetValue(Block.TextAlignmentProperty);
                 WriteTextAlignmentProperty(ta);              
                

            }
            else
            {

                TextAlignment ta = (TextAlignment)parent.GetValue(Block.TextAlignmentProperty);
                WriteTextAlignmentProperty(ta);

                //Handle Vertical Spacing
                Thickness mx = (Thickness)parent.GetValue(Block.MarginProperty);
                WriteMarginProperty(mx);

                double indent = (double)parent.GetValue(Paragraph.TextIndentProperty);
                WriteTextIndentProperty(indent);

            }
            
            _openxmlwriter.WriteEndElement();


            lastListLevel = listLevel;
           

            
        }


        void WriteListProperty()
        {
            //Write Numbering Property
            int levelindex = listLevel-1;
            if (levelindex>=0)
            {
                _openxmlwriter.WriteStartElement("w:numPr");
                _openxmlwriter.WriteStartElement("w:ilvl");
                _openxmlwriter.WriteAttributeString("w:val", levelindex.ToString());    
                _openxmlwriter.WriteEndElement();
                _openxmlwriter.WriteStartElement("w:numId");

                //System.Windows.MessageBox.Show(levelindex.ToString());
                //System.Windows.MessageBox.Show(listMarkerStyle[levelindex].ToString());

                //Currently, only 2 TextMarkerStytles are supported, Disc and Decimal
                if (listMarkerStyle[levelindex]==TextMarkerStyle.Disc)
                    _openxmlwriter.WriteAttributeString("w:val", "1");
                else if (listMarkerStyle[levelindex] == TextMarkerStyle.Decimal)
                    _openxmlwriter.WriteAttributeString("w:val", "2");
                else
                    _openxmlwriter.WriteAttributeString("w:val", "1");

                _openxmlwriter.WriteEndElement();
                _openxmlwriter.WriteEndElement();
            }           

        }

        void WriteRunProperties()
        {
            DependencyObject parent;
            TextPointer position;

            position = _textpointer;
            position = position.GetNextContextPosition(LogicalDirection.Forward);
            parent = position.Parent;

            _openxmlwriter.WriteStartElement("w:rPr");

            string fontName = ((FontFamily) parent.GetValue(TextElement.FontFamilyProperty)).ToString();
            WriteFontNameProperty(fontName);
            
            double fontSize = (double) parent.GetValue(TextElement.FontSizeProperty);
            WriteFontSizeProperty(fontSize);
            
            FontWeight fontWeight = (FontWeight) parent.GetValue(TextElement.FontWeightProperty);
            WriteBoldProperty(fontWeight);
            
            FontStyle fontStyle = (FontStyle) parent.GetValue(TextElement.FontStyleProperty);
            WriteItalicProperty(fontStyle);
            
            TextDecorationCollection enumdec;
            enumdec =  parent.GetValue(TextBlock.TextDecorationsProperty) as TextDecorationCollection;
            WriteUnderlineProperty(enumdec);
                        
            SolidColorBrush fontBrush ;
            fontBrush = parent.GetValue(TextElement.ForegroundProperty) as SolidColorBrush;
            WriteFontColorProperty(fontBrush);

            _openxmlwriter.WriteEndElement();
        }

        void WriteTextAlignmentProperty(TextAlignment ta)
        {                
                _openxmlwriter.WriteStartElement("w:jc");
                if (ta==TextAlignment.Left)
                    _openxmlwriter.WriteAttributeString("w:val", "left");
                else if (ta == TextAlignment.Center)
                    _openxmlwriter.WriteAttributeString("w:val", "center");
                else if (ta == TextAlignment.Right)
                    _openxmlwriter.WriteAttributeString("w:val", "right");
                else if (ta == TextAlignment.Justify)
                    _openxmlwriter.WriteAttributeString("w:val", "both");                
                _openxmlwriter.WriteEndElement();

        }
        
      
        void WriteMarginProperty(Thickness mx)
        {             
                _openxmlwriter.WriteStartElement("w:spacing");
                String stxbottom = mx.Bottom.ToString();
                String stxtop = mx.Bottom.ToString();
                String stxleft = mx.Left.ToString();
                double chosenSizeBottom = 180;
                double chosenSizeTop = 180;
                if (double.TryParse(stxbottom, out chosenSizeBottom))
                {
                    int xx = (int) Math.Round(chosenSizeBottom * 15);                            
                    if (xx>=0)
                        _openxmlwriter.WriteAttributeString("w:after", xx.ToString());
                    else
                        _openxmlwriter.WriteAttributeString("w:after", "0");

                }
                else
                    _openxmlwriter.WriteAttributeString("w:after", "0");

                if (double.TryParse(stxtop, out chosenSizeTop))
                {
                    int xx = (int) Math.Round(chosenSizeTop * 15);   
                    if (xx>=0)
                        _openxmlwriter.WriteAttributeString("w:before", xx.ToString());
                    else
                        _openxmlwriter.WriteAttributeString("w:before", "0");

                }
                else
                {
                    _openxmlwriter.WriteAttributeString("w:before", "0");

                }
                
                _openxmlwriter.WriteEndElement();


                double chosenSizeLeft = 0;
                if (double.TryParse(stxleft, out chosenSizeLeft))
                {
                    int xx = (int) Math.Round(chosenSizeLeft * 15);
                    if (xx >= 0)
                    {
                        _openxmlwriter.WriteStartElement("w:ind");
                        _openxmlwriter.WriteAttributeString("w:left", xx.ToString());
                        _openxmlwriter.WriteEndElement();

                    }

                }

        }
            
            
         void WriteTextIndentProperty(double indent)
         {

                //wpf device independent pixel is 1/96 of an inch ?
                //so value * 1440/96 = value * 15 for conversion to twips
                int xx = (int)Math.Round(indent*15);
                if (xx > 0)
                {
                    _openxmlwriter.WriteStartElement("w:ind");
                    _openxmlwriter.WriteAttributeString("w:firstLine", xx.ToString());
                    _openxmlwriter.WriteEndElement();                                     
                }                         

         }
        
        void WriteBoldProperty(FontWeight fontWeight)
        {
            _openxmlwriter.WriteStartElement("w:b");
            if (fontWeight > FontWeights.Medium) 
                _openxmlwriter.WriteAttributeString("w:val", "on");
            else
                _openxmlwriter.WriteAttributeString("w:val", "off");
            _openxmlwriter.WriteEndElement();

        }


        void WriteUnderlineProperty(TextDecorationCollection enumdec)
        {
            if (enumdec != null)
            {
                int zs = enumdec.Count;
                if (zs == 1)
                {
                    //Assume all Textdocorations as underline
                    _openxmlwriter.WriteStartElement("w:u");
                    _openxmlwriter.WriteAttributeString("w:val", "single");
                    _openxmlwriter.WriteEndElement();

                }

            }
        }

        void WriteItalicProperty( FontStyle fontStyle)
        {
            _openxmlwriter.WriteStartElement("w:i");
             if (fontStyle != FontStyles.Normal)
                _openxmlwriter.WriteAttributeString("w:val", "on");
             else
                _openxmlwriter.WriteAttributeString("w:val", "off");
            _openxmlwriter.WriteEndElement();

        }


        void WriteFontColorProperty(SolidColorBrush fontbrush)
        {            
                Color color;                
                if (fontbrush != null)
                {
                    color = fontbrush.Color;
                    string colorstr = String.Format( "{0:x2}{1:x2}{2:x2}", color.R, color.G, color.B);
                    _openxmlwriter.WriteStartElement("w:color");
                    _openxmlwriter.WriteAttributeString("w:val", colorstr);
                    _openxmlwriter.WriteEndElement();
                }         
            
        }

        void WriteFontSizeProperty(double fontSize)
        {       
                string sizeStr = ((int)(Math.Round(fontSize * 72.0 / 96.0) * 2)).ToString();                
                _openxmlwriter.WriteStartElement("w:sz");
                _openxmlwriter.WriteAttributeString("w:val", sizeStr);
                _openxmlwriter.WriteEndElement();
                 
        }
         
        void WriteFontNameProperty(string fontName)
        {
            _openxmlwriter.WriteStartElement("w:rFonts");
            _openxmlwriter.WriteAttributeString("w:ascii", fontName);
            _openxmlwriter.WriteAttributeString("w:hAnsi", fontName);
            _openxmlwriter.WriteAttributeString("w:cs", fontName);
            _openxmlwriter.WriteEndElement();

        }


        void WriteNumberingHeader()
        {
            _numberingwriter.WriteStartDocument(true);
            _numberingwriter.WriteStartElement("w:numbering");
            _numberingwriter.WriteAttributeString("xmlns", "ve", null, "http://schemas.openxmlformats.org/markup-compatibility/2006");
            _numberingwriter.WriteAttributeString("xmlns", "o", null, "urn:schemas-microsoft-com:office:office");
            _numberingwriter.WriteAttributeString("xmlns", "r", null, "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            _numberingwriter.WriteAttributeString("xmlns", "m", null, "http://schemas.openxmlformats.org/officeDocument/2006/math");
            _numberingwriter.WriteAttributeString("xmlns", "v", null, "urn:schemas-microsoft-com:vml");
            _numberingwriter.WriteAttributeString("xmlns", "wp", null, "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            _numberingwriter.WriteAttributeString("xmlns", "w10", null, "urn:schemas-microsoft-com:office:word");
            _numberingwriter.WriteAttributeString("xmlns", "w", null, "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            _numberingwriter.WriteAttributeString("xmlns", "wne", null, "http://schemas.microsoft.com/office/word/2006/wordml");


        }

        void WriteNumberingBody()
        {
             //Generate a default numbering body for only 2 styles Bullets and Decimal
             //and leave all the formatting to be handled by the Paragraph/Run tags 
            
              //write <w:abstractNum> 0 - bullets
             _numberingwriter.WriteStartElement("w:abstractNum");
             _numberingwriter.WriteAttributeString("w:abstractNumId", "0");
            
             _numberingwriter.WriteStartElement("w:nsid");
             _numberingwriter.WriteAttributeString("w:val", "90000000");
             _numberingwriter.WriteEndElement();

             _numberingwriter.WriteStartElement("w:multiLevelType");
             _numberingwriter.WriteAttributeString("w:val", "hybridMultilevel");
             _numberingwriter.WriteEndElement();


             //for each level
             for (int i = 0; i < 9; i++)
             {
                 _numberingwriter.WriteStartElement("w:lvl");
                 _numberingwriter.WriteAttributeString("w:ilvl", i.ToString());

                 //start number
                 _numberingwriter.WriteStartElement("w:start");
                 _numberingwriter.WriteAttributeString("w:val", "1");
                 _numberingwriter.WriteEndElement();

                 //bullet or decimal
                 _numberingwriter.WriteStartElement("w:numFmt");
                 _numberingwriter.WriteAttributeString("w:val", "bullet");
                 _numberingwriter.WriteEndElement();               
                 
                 //numbering text
                 _numberingwriter.WriteStartElement("w:lvlText");
                 _numberingwriter.WriteAttributeString("w:val", "·");
                 _numberingwriter.WriteEndElement();
                 
                 //left justification
                 _numberingwriter.WriteStartElement("w:lvlJc");
                 _numberingwriter.WriteAttributeString("w:val", "left");
                 _numberingwriter.WriteEndElement();

                 //tabs and indentation
                 _numberingwriter.WriteStartElement("w:pPr");                 
                 _numberingwriter.WriteStartElement("w:tabs");

                 //half inch (720 twips) indent for each level 
                 int tabposition = (i+1)*720;
                 _numberingwriter.WriteStartElement("w:tab");
                 _numberingwriter.WriteAttributeString("w:val", "num");
                 _numberingwriter.WriteAttributeString("w:pos", tabposition.ToString());
                 _numberingwriter.WriteEndElement();
                 
                 _numberingwriter.WriteEndElement(); //tabs

                 _numberingwriter.WriteStartElement("w:ind");
                 _numberingwriter.WriteAttributeString("w:left", tabposition.ToString());
                 _numberingwriter.WriteAttributeString("w:hanging", "360");
                 _numberingwriter.WriteEndElement();
                 _numberingwriter.WriteEndElement(); //pPr


                 //use Symbol fonts for bulletted list
                 _numberingwriter.WriteStartElement("w:rPr");
                 
                 _numberingwriter.WriteStartElement("w:rFonts");
                 _numberingwriter.WriteAttributeString("w:ascii", "Symbol");
                 _numberingwriter.WriteAttributeString("w:hAnsi", "Symbol");
                 _numberingwriter.WriteAttributeString("w:hint", "default");
                 _numberingwriter.WriteEndElement(); //rFonts 
                 _numberingwriter.WriteEndElement(); //rPr


                 _numberingwriter.WriteEndElement(); //lvl
             } 
             
                         
             _numberingwriter.WriteEndElement();
             //close <w:abstractNum> 0

             
             //write <w:abstractNum> 1 - numbers
             _numberingwriter.WriteStartElement("w:abstractNum");
             _numberingwriter.WriteAttributeString("w:abstractNumId", "1");


             _numberingwriter.WriteStartElement("w:nsid");
             _numberingwriter.WriteAttributeString("w:val", "90000001");
             _numberingwriter.WriteEndElement();

             _numberingwriter.WriteStartElement("w:multiLevelType");
             _numberingwriter.WriteAttributeString("w:val", "hybridMultilevel");
             //_numberingwriter.WriteAttributeString("w:val", "multilevel");
             _numberingwriter.WriteEndElement();

             //for each level
             for (int i = 0; i < 9; i++)
             {
                 _numberingwriter.WriteStartElement("w:lvl");
                 _numberingwriter.WriteAttributeString("w:ilvl", i.ToString());

                 _numberingwriter.WriteStartElement("w:start");
                 _numberingwriter.WriteAttributeString("w:val", "1");
                 _numberingwriter.WriteEndElement();

                 _numberingwriter.WriteStartElement("w:numFmt");
                 _numberingwriter.WriteAttributeString("w:val", "decimal");
                 _numberingwriter.WriteEndElement();

                 String numberingStr;
                 int currentNumbering = i + 1;
                 numberingStr = "%" + currentNumbering.ToString() + ".";
                 _numberingwriter.WriteStartElement("w:lvlText");
                 _numberingwriter.WriteAttributeString("w:val", numberingStr);
                 _numberingwriter.WriteEndElement();

                 //left justification
                 _numberingwriter.WriteStartElement("w:lvlJc");
                 _numberingwriter.WriteAttributeString("w:val", "left");                 
                 _numberingwriter.WriteEndElement();

                 //tabs and indentation
                 _numberingwriter.WriteStartElement("w:pPr");
                 _numberingwriter.WriteStartElement("w:tabs");

                 //half inch (720 twips) indent for each level 
                 int tabposition = (i + 1) * 720;
                 _numberingwriter.WriteStartElement("w:tab");
                 _numberingwriter.WriteAttributeString("w:val", "num");
                 _numberingwriter.WriteAttributeString("w:pos", tabposition.ToString());
                 _numberingwriter.WriteEndElement();

                 _numberingwriter.WriteEndElement(); //tabs

                 _numberingwriter.WriteStartElement("w:ind");
                 _numberingwriter.WriteAttributeString("w:left", tabposition.ToString());
                 _numberingwriter.WriteAttributeString("w:hanging", "360");
                 _numberingwriter.WriteEndElement();
                 _numberingwriter.WriteEndElement(); //pPr


                 //use Times New Roman fonts for numbered lists
                 _numberingwriter.WriteStartElement("w:rPr");

                 _numberingwriter.WriteStartElement("w:rFonts");                 
                 _numberingwriter.WriteAttributeString("w:cs", "Times New Roman");  
                 _numberingwriter.WriteAttributeString("w:hint", "default");
                 _numberingwriter.WriteEndElement(); //rFonts 
                 _numberingwriter.WriteEndElement(); //rPr

                 _numberingwriter.WriteEndElement(); //lvl
             }             

             
             _numberingwriter.WriteEndElement();
             //close <w:abstractNum> 1           
            
            

             //<w:num Id 1>
             _numberingwriter.WriteStartElement("w:num");
             _numberingwriter.WriteAttributeString("w:numId", "1");
             _numberingwriter.WriteStartElement("w:abstractNumId");
             _numberingwriter.WriteAttributeString("w:val", "0");
             _numberingwriter.WriteEndElement();
             _numberingwriter.WriteEndElement();

            

             //<w:num Id 2>
             _numberingwriter.WriteStartElement("w:num");
             _numberingwriter.WriteAttributeString("w:numId", "2");
             _numberingwriter.WriteStartElement("w:abstractNumId");
             _numberingwriter.WriteAttributeString("w:val", "1");
             _numberingwriter.WriteEndElement();
             _numberingwriter.WriteEndElement();
            
                        

        }
        

        void WriteNumberingFooter()
        {
            _numberingwriter.WriteEndElement();
            
        }      

                
        void WriteOpenXmlHeader()
        {
            _openxmlwriter.WriteStartDocument(true);
            _openxmlwriter.WriteStartElement("w:document");
            _openxmlwriter.WriteAttributeString("xmlns", "ve", null, "http://schemas.openxmlformats.org/markup-compatibility/2006");
            _openxmlwriter.WriteAttributeString("xmlns", "o",null,"urn:schemas-microsoft-com:office:office" );
            _openxmlwriter.WriteAttributeString("xmlns", "r",null,"http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            _openxmlwriter.WriteAttributeString("xmlns", "m",null,"http://schemas.openxmlformats.org/officeDocument/2006/math" );
            _openxmlwriter.WriteAttributeString("xmlns", "v",null,"urn:schemas-microsoft-com:vml");
            _openxmlwriter.WriteAttributeString("xmlns", "wp",null,"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            _openxmlwriter.WriteAttributeString("xmlns", "w10",null,"urn:schemas-microsoft-com:office:word");
            _openxmlwriter.WriteAttributeString("xmlns", "w",null,"http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            _openxmlwriter.WriteAttributeString("xmlns", "wne",null,"http://schemas.microsoft.com/office/word/2006/wordml");

            
            _openxmlwriter.WriteAttributeString("xml:space", "preserve");

            _openxmlwriter.WriteStartElement("w:body");
            _openxmlwriter.WriteStartElement("w:sect");
        }

        
        void WriteOpenXmlFooter()
        {
            _openxmlwriter.WriteEndElement();   
            _openxmlwriter.WriteEndElement();   
            _openxmlwriter.WriteEndElement();   
        }

        
        void CreateTextRunTag()
        {
            EnsureWithinParagraphTag();
            if (!_insideRun)
            {
                _openxmlwriter.WriteStartElement("w:r");
                _insideRun = true;
                WriteRunProperties();
            }
        }

        
        void CreateParagraphTag()
        {
           
            if (!_insideParagraph)
            {
                _openxmlwriter.WriteStartElement("w:p");
                _insideParagraph = true;
                WriteParagraphProperties();
            }
        }

        void EnsureWithinParagraphTag()
        {
            if (!_insideParagraph)
            {
                _openxmlwriter.WriteStartElement("w:p");
                _insideParagraph = true;
                WriteParagraphProperties();
            }
        }

        
        void CloseTextRunTag()
        {
            if (_insideRun)
            {
                // Close w:r.
                _openxmlwriter.WriteEndElement();
                _insideRun = false;
            }
        }

        
        void CloseParagraphTag()
        {
            CloseTextRunTag();
            if (_insideParagraph)
            {
                // Close w:p.
                _openxmlwriter.WriteEndElement();
                _insideParagraph = false;
            }
        }

        
        bool _insideParagraph;        
        bool _insideRun;
        TextPointer _textpointer;        
        XmlWriter _openxmlwriter;
        XmlWriter _numberingwriter;
        int listContext; //indicates next paragraph encountered is the 1st paragraph in a list item and so need to add bullets/numberings
        int listLevel; //current list level, 0 - means not in list , 1 - means list level 0 in WordProcessingML, 2 - means list level 1 in WordProcessingML        
        int lastListLevel;
        TextMarkerStyle[] listMarkerStyle; 

        
    }
}
