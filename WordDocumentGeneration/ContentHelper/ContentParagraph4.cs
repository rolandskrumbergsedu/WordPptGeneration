﻿using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;

namespace WordDocumentGeneration.ContentHelper
{
    public static class ContentParagraph4
    {
        // Creates an Paragraph instance and adds its children.
        public static Paragraph GenerateParagraph()
        {
            Paragraph paragraph1 = new Paragraph() { RsidParagraphAddition = "009B2C1D", RsidRunAdditionDefault = "009B2C1D", ParagraphId = "2C48C2FF", TextId = "77777777" };
            return paragraph1;
        }


    }
}