using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.IO.Packaging;
using System.Printing;
using System.Windows.Markup;
using System.Windows.Xps.Packaging;
using System.Windows.Xps;
using System.Windows.Documents;
using System.Windows.Xps.Serialization;
using NiPDF;

namespace DocumentEditorTestApp
{
    public static class PdfConvertor
    {
        public static void SaveToPdf(this FlowDocument flowDoc, string filename)
        {
            MemoryStream xamlStream = new MemoryStream();
            XamlWriter.Save(flowDoc, xamlStream);
            File.WriteAllBytes("d:\\file.xaml", xamlStream.ToArray());

            IDocumentPaginatorSource text = flowDoc as IDocumentPaginatorSource;
            xamlStream.Close();

            MemoryStream memoryStream = new MemoryStream();
            Package pkg = Package.Open(memoryStream, FileMode.Create, FileAccess.ReadWrite);

            string pack = "pack://temp.xps";
            PackageStore.AddPackage(new Uri(pack), pkg);

            XpsDocument doc = new XpsDocument(pkg, CompressionOption.SuperFast, pack);
            XpsSerializationManager rsm = new XpsSerializationManager(new XpsPackagingPolicy(doc), false);
            DocumentPaginator pgn = text.DocumentPaginator;
            rsm.SaveAsXaml(pgn);

            MemoryStream xpsStream = new MemoryStream();
            var writer = new XpsSerializerFactory().CreateSerializerWriter(xpsStream);
            writer.Write(doc.GetFixedDocumentSequence());

            MemoryStream outStream = new MemoryStream();
            NiXPS.Converter.XpsToPdf(xpsStream, outStream);
            File.WriteAllBytes("file.pdf", outStream.ToArray());
            
        }
    }
}
