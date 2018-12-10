using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordDocumentGeneration
{
    public class GenerationData
    {
        public TitleArea TitleArea { get; set; }
    }

    public class TitleArea
    {
        public string Title { get; set; }
        public string Name { get; set; }
        public string Date { get; set; }
    }
}
