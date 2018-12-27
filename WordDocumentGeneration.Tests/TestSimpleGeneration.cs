using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace WordDocumentGeneration.Tests
{
    [TestClass]
    public class TestSimpleGeneration
    {
        [TestMethod]
        public void Test_FileStoring_TestRealData()
        {
            var documentManager = new WordDocumentManagerV2();

            var simpleFile = GetGenerationData();

            var filePath = "C:\\temp\\WordGeneration";
            var fileName = $"{Guid.NewGuid().ToString()}.docx";

            documentManager.SaveDocument(simpleFile, filePath, fileName);
        }

        private static GenerationData GetGenerationData()
        {
            return new GenerationData
            {
                DocumentProperties = new DocumentProperties
                {
                  Creator  = "Agnese Zanriba",
                  Category = "",
                  Keywords = "",
                  Subject = "",
                  Description = "",
                  Title = ""
                },
                TitleArea = new TitleArea
                {
                    Date = "December 2018",
                    Name = "Rolands Krumbergs",
                    Title = "Confidential candidate CV"
                }
            };
        }
    }
}
