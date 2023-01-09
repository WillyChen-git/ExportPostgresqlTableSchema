using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ExportPostgresqlTableSchema
{
    public class TableSchemaModel
    {
        public string table_name { get; set; }
        public string table_description { get; set; }
        public string column_name { get; set; }
        public string column_description { get; set; }
        public int? ordinal_position { get; set; }
        public string data_type { get; set; }
        public int? character_maximum_length { get; set; }
        public string default_value { get; set; }
        public string is_nullable { get; set; }
        public string? constraint_type { get; set; }
        public string? foreign_table_name { get; set; }
        public string? foreign_column_name { get; set; }
    }
}