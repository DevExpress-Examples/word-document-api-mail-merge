using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordProcessingMailMerge
{
    using DevExpress.Office.Services;
    using System;
    using System.Data;
    using System.IO;

    public class ImageStreamProvider : IUriStreamProvider
    {
        static readonly string prefix = "dbimg://";
        DataTable table;
        string columnName;

        public ImageStreamProvider(DataTable sourceTable, string imageColumn)
        {
            this.table = sourceTable;
            this.columnName = imageColumn;
        }


        public Stream GetStream(string uri)
        {
            // Parse the retrieved URI string
            uri = uri.Trim();
            if (!uri.StartsWith(prefix))
                return null;

            // Remove the prefix from the retrieved URI string
            string strId = uri.Substring(prefix.Length).Trim();
            int id;

            // Check if the string contains the primary key
            if (!int.TryParse(strId, out id))
                return null;

            // Retrieve the row that corresponds
            // with the key
            DataRow row = table.Rows.Find(id);
            if (row == null)
                return null;

            // Convert the image string from this row
            // to a byte array
            byte[] bytes = Convert.FromBase64String(row[columnName] as string) as byte[];
            if (bytes == null)
                return null;

            // Return the MemoryStream with an image
            MemoryStream memoryStream = new MemoryStream(bytes);
            return memoryStream;
        }
    }
}
