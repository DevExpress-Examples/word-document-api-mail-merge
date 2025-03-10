Imports System.IO
Imports System.Data
Imports DevExpress.Office.Services
Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks

Namespace WordProcessingMailMerge

	Public Class ImageStreamProvider
		Implements IUriStreamProvider

		Private Shared ReadOnly prefix As String = "dbimg://"
		Private table As DataTable
		Private columnName As String

		Public Sub New(ByVal sourceTable As DataTable, ByVal imageColumn As String)
			Me.table = sourceTable
			Me.columnName = imageColumn
		End Sub


		Public Function GetStream(ByVal uri As String) As Stream Implements IUriStreamProvider.GetStream
			' Parse the retrieved URI string
			uri = uri.Trim()
			If Not uri.StartsWith(prefix) Then
				Return Nothing
			End If

			' Remove the prefix from the retrieved URI string
			Dim strId As String = uri.Substring(prefix.Length).Trim()
			Dim id As Integer = Nothing

			' Check if the string contains the primary key
			If Not Integer.TryParse(strId, id) Then
				Return Nothing
			End If

			' Retrieve the row that corresponds
			' with the key
			Dim row As DataRow = table.Rows.Find(id)
			If row Is Nothing Then
				Return Nothing
			End If

			' Convert the image string from this row
			' to a byte array
			Dim bytes() As Byte = TryCast(Convert.FromBase64String(TryCast(row(columnName), String)), Byte())
			If bytes Is Nothing Then
				Return Nothing
			End If

			' Return the MemoryStream with an image
			Dim memoryStream As New MemoryStream(bytes)
			Return memoryStream
		End Function
	End Class
End Namespace
