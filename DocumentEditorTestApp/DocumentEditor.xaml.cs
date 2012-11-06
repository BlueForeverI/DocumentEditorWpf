using System;
using System.Collections.Generic;
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
    }
}
