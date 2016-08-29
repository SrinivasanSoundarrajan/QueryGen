using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QueryGen
{
    public class QueryBuilderInput
    {
        public string SheetName { get; set; }
        public int StartRow { get; set; }
        public int EndRow { get; set; }
        public List<Column> ExcelColumns { get; set; }
        public string TableNameDB { get; set; }
        public string File { get; set; }
    }

    public class Column
    {
        public Column()
        {
            PassNull = false;
            IsString = true;
        }

        public string ColumnNameDB { get; set; }
        public string ColumnNameExcel { get; set; }
        public bool IsString { get; set; }
        public bool PassNull { get; set; }
    }

}
