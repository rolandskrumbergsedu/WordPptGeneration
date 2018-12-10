using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace WordDocumentGeneration.Tests
{
    [TestClass]
    public class TestSimpleGeneration
    {
        [TestMethod]
        public void Test_FileStoring()
        {
            var documentManager = new WordDocumentManager();

            var simpleFile = new GenerationData();

            var filePath = "C:\\temp\\WordGeneration";
            var fileName = $"{Guid.NewGuid().ToString()}.docx";

            documentManager.SaveDocument(simpleFile, filePath, fileName);
        }

        [TestMethod]
        public void Test_FileStoring_TestRealData()
        {
            var documentManager = new WordDocumentManager();

            var simpleFile = GetGenerationData();

            var filePath = "C:\\temp\\WordGeneration";
            var fileName = $"{Guid.NewGuid().ToString()}.docx";

            documentManager.SaveDocument(simpleFile, filePath, fileName);
        }

        private static GenerationData GetGenerationData()
        {
            return new GenerationData
            {
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
