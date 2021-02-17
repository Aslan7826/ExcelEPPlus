using System;
using System.Collections.Generic;
using System.IO;

namespace ExcelEPPlus
{
    class Program
    {
        static void Main(string[] args)
        {
            MemoryStream stream = new MemoryStream();
            var model = new List<Model>() { new Model() { Id = 1 } };
            new ExcelService().ExportToExcel(stream,model);
        }
    }
}
