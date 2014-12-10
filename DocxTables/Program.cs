using Novacode;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocxTables
{
    class Program
    {
        static void Main(string[] args)
        {
            using (DocX document = DocX.Create("C:/Temp/tables.docx"))
            {
                foreach (TableDesign td in (TableDesign[])Enum.GetValues(typeof(TableDesign)))
                {
                    Table aTable = document.InsertTable(5, 5); // because office preview is also 5 x 5
                    aTable.Design = td;
                    aTable.Rows[0].Cells[0].Paragraphs.First().Append(td.ToString());
                }
                document.Save();
            }
        }
    }
}
