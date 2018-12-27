using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using WordDocumentGeneration.ContentHelper;

namespace WordDocumentGeneration.Helpers
{
    public static class MainDocumentPartHelper
    {
        public static void GenerateMainDocumentPart1Content(MainDocumentPart mainDocumentPart1)
        {
            Document document1 = ContentDocument.CreateDocument();

            Body body1 = new Body();

            SectionProperties sectionProperties1 = ContentSectionProperties.CreateProperties();

            body1.Append(ContentTable1.GenerateTable());
            body1.Append(ContentParagraph1.GenerateParagraph());
            body1.Append(ContentTable2.GenerateTable());
            body1.Append(ContentParagraph2.GenerateParagraph());
            body1.Append(ContentTable3.GenerateTable());
            body1.Append(ContentParagraph3.GenerateParagraph());
            body1.Append(ContentParagraph4.GenerateParagraph());
            body1.Append(ContentTable4.GenerateTable());
            body1.Append(ContentParagraph5.GenerateParagraph());
            body1.Append(ContentTable5.GenerateTable());
            body1.Append(ContentParagraph6.GenerateParagraph());
            body1.Append(ContentParagraph7.GenerateParagraph());
            body1.Append(ContentParagraph8.GenerateParagraph());
            body1.Append(ContentTable6.GenerateTable());
            body1.Append(paragraph124);
            body1.Append(paragraph125);
            body1.Append(paragraph126);
            body1.Append(table8);
            body1.Append(paragraph255);
            body1.Append(table9);
            body1.Append(paragraph261);
            body1.Append(table10);
            body1.Append(paragraph267);
            body1.Append(table11);
            body1.Append(paragraph271);
            body1.Append(table12);
            body1.Append(paragraph276);
            body1.Append(paragraph277);
            body1.Append(sectionProperties1);

            document1.Append(body1);

            mainDocumentPart1.Document = document1;
        }
    }
}
