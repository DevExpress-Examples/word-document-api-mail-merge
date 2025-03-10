using DevExpress.Office.Services;
using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;
using System.Data;
using System.Diagnostics;

namespace WordProcessingMailMerge
{
    internal class Program
    {
        static void Main(string[] args)
        {
            DataSet xmlDataSet = new DataSet();
            xmlDataSet.ReadXml("..\\..\\..\\Employees.xml");
            xmlDataSet.Tables[0].PrimaryKey = new DataColumn[] { xmlDataSet.Tables[0].Columns[0] };

            using (var wordProcessor = new RichEditDocumentServer())
            {
                wordProcessor.LoadDocument("..\\..\\..\\template.docx");
                AddMailMergeFields(wordProcessor.Document);
                wordProcessor.CalculateDocumentVariable += WordProcessor_CalculateDocumentVariable;

                // Register the URI provider service
                IUriStreamService uriStreamService = wordProcessor.GetService<IUriStreamService>();
                uriStreamService.RegisterProvider(new ImageStreamProvider(xmlDataSet.Tables[0], "Photo"));

                MailMergeOptions myMergeOptions =
                    wordProcessor.Document.CreateMailMergeOptions();
                myMergeOptions.DataSource = xmlDataSet.Tables[0];
                myMergeOptions.MergeMode = MergeMode.NewSection;

                wordProcessor.MailMerge(myMergeOptions, "result.docx", DocumentFormat.OpenXml);
            }

            Process.Start(new ProcessStartInfo("result.docx") { UseShellExecute = true });
        }
        private static void AddMailMergeFields(Document document)
        {
            DocumentRange[] pictureRanges = document.FindAll("Photo", SearchOptions.WholeWord);

            foreach (var pictureRange in pictureRanges)
            {
                DocumentPosition picturePosition = pictureRange.End;

                // Delete the placeholder
                document.Delete(pictureRange);

                // Insert a field to a picture from a database
                Field pictureField = document.Fields.Create(picturePosition, @"INCLUDEPICTURE ""dbimg://{ placeholder }""");

                // Find a placeholder for a nested MERGEFIELD
                DocumentRange nestedFieldRange =
                    document.FindAll("{ placeholder }", SearchOptions.WholeWord, pictureField.CodeRange).First();

                // Clear the placeholder range
                document.Delete(nestedFieldRange);

                // Create a nested field
                document.Fields.Create(nestedFieldRange.Start, "MERGEFIELD EmployeeID");
            }

            // Find a placeholder for another image
            // in the footer
            SubDocument footer = document.Sections[0].BeginUpdateFooter();
            DocumentRange[] imageRanges =
                footer.FindAll("image", SearchOptions.WholeWord);
            foreach (var imageRange in imageRanges)
            {
                DocumentPosition imagePosition = imageRange.End;

                // Delete the phrase
                footer.Delete(imageRange);

                // Create a field at the placeholder's position
                footer.Fields.Create(imagePosition, @"INCLUDEPICTURE ""DevExpress.png""");
            }
            footer.EndUpdate();

            // Find a placeholder for the FirstName field:
            DocumentRange[] nameRanges =
                document.FindAll("FirstName", SearchOptions.WholeWord);
            foreach (var nameRange in nameRanges)
            {
                DocumentPosition namePosition = nameRange.End;

                // Delete the phrase
                document.Delete(nameRange);

                // Create a field at the placeholder position
                document.Fields.Create(namePosition, @"MERGEFIELD ""FirstName""");
            }
        }
        private static void WordProcessor_CalculateDocumentVariable(object sender, CalculateDocumentVariableEventArgs e)
        {
            if (e.Arguments.Count > 0)
            {
                // Retrieve the MERGEFIELD field value
                DateTimeOffset hireDate = Convert.ToDateTime(e.Arguments[0].Value);
                DateTimeOffset currentDate = DateTime.Now;

                // Calculate the difference between the current date
                // and the hire date
                var dif = currentDate.Subtract(hireDate);

                // Calculate the number of years
                int years = dif.Days / 365;

                // Specify the DOCVARIABLE field value
                e.Value = years.ToString();
                e.Handled = true;
            }
            e.Handled = true;
        }
    }
}
