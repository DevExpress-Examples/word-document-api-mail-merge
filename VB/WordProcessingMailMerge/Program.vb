Option Infer On

Imports DevExpress.Office.Services
Imports DevExpress.XtraRichEdit
Imports DevExpress.XtraRichEdit.API.Native
Imports System.Data
Imports System.Diagnostics

Namespace WordProcessingMailMerge
	Friend Class Program
		Shared Sub Main(ByVal args() As String)
			Dim xmlDataSet As New DataSet()
			xmlDataSet.ReadXml("..\..\..\Employees.xml")
			xmlDataSet.Tables(0).PrimaryKey = New DataColumn() { xmlDataSet.Tables(0).Columns(0) }

			Using wordProcessor = New RichEditDocumentServer()
				wordProcessor.LoadDocument("..\..\..\template.docx")
				AddMailMergeFields(wordProcessor.Document)
				AddHandler wordProcessor.CalculateDocumentVariable, AddressOf WordProcessor_CalculateDocumentVariable

				' Register the URI provider service
				Dim uriStreamService As IUriStreamService = wordProcessor.GetService(Of IUriStreamService)()
				uriStreamService.RegisterProvider(New ImageStreamProvider(xmlDataSet.Tables(0), "Photo"))

				Dim myMergeOptions As MailMergeOptions = wordProcessor.Document.CreateMailMergeOptions()
				myMergeOptions.DataSource = xmlDataSet.Tables(0)
				myMergeOptions.MergeMode = MergeMode.NewSection

				wordProcessor.MailMerge(myMergeOptions, "result.docx", DocumentFormat.OpenXml)
			End Using

			Process.Start(New ProcessStartInfo("result.docx") With {.UseShellExecute = True})
		End Sub
		Private Shared Sub AddMailMergeFields(ByVal document As Document)
			Dim pictureRanges() As DocumentRange = document.FindAll("Photo", SearchOptions.WholeWord)

			For Each pictureRange In pictureRanges
				Dim picturePosition As DocumentPosition = pictureRange.End

				' Delete the placeholder
				document.Delete(pictureRange)

				' Insert a field to a picture from a database
				Dim pictureField As Field = document.Fields.Create(picturePosition, "INCLUDEPICTURE ""dbimg://{ placeholder }""")

				' Find a placeholder for a nested MERGEFIELD
				Dim nestedFieldRange As DocumentRange = document.FindAll("{ placeholder }", SearchOptions.WholeWord, pictureField.CodeRange).First()

				' Clear the placeholder range
				document.Delete(nestedFieldRange)

				' Create a nested field
				document.Fields.Create(nestedFieldRange.Start, "MERGEFIELD EmployeeID")
			Next pictureRange

			' Find a placeholder for another image
			' in the footer
			Dim footer As SubDocument = document.Sections(0).BeginUpdateFooter()
			Dim imageRanges() As DocumentRange = footer.FindAll("image", SearchOptions.WholeWord)
			For Each imageRange In imageRanges
				Dim imagePosition As DocumentPosition = imageRange.End

				' Delete the phrase
				footer.Delete(imageRange)

				' Create a field at the placeholder's position
				footer.Fields.Create(imagePosition, "INCLUDEPICTURE ""DevExpress.png""")
			Next imageRange
			footer.EndUpdate()

			' Find a placeholder for the FirstName field:
			Dim nameRanges() As DocumentRange = document.FindAll("FirstName", SearchOptions.WholeWord)
			For Each nameRange In nameRanges
				Dim namePosition As DocumentPosition = nameRange.End

				' Delete the phrase
				document.Delete(nameRange)

				' Create a field at the placeholder position
				document.Fields.Create(namePosition, "MERGEFIELD ""FirstName""")
			Next nameRange
		End Sub
		Private Shared Sub WordProcessor_CalculateDocumentVariable(ByVal sender As Object, ByVal e As CalculateDocumentVariableEventArgs)
			If e.Arguments.Count > 0 Then
				' Retrieve the MERGEFIELD field value
				Dim hireDate As DateTimeOffset = Convert.ToDateTime(e.Arguments(0).Value)
				Dim currentDate As DateTimeOffset = Date.Now

				' Calculate the difference between the current date
				' and the hire date
				Dim dif = currentDate.Subtract(hireDate)

				' Calculate the number of years
				Dim years As Integer = dif.Days \ 365

				' Specify the DOCVARIABLE field value
				e.Value = years.ToString()
				e.Handled = True
			End If
			e.Handled = True
		End Sub
	End Class
End Namespace
