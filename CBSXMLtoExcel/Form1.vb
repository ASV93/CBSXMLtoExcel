'============================================================================
'
'    CBSXMLtoExcel
'    Copyright (C) 2015 Visual Software Corporation
'
'    Author: ASV93
'    File: Form1.vb
'
'    This program is free software; you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation; either version 2 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License along
'    with this program; if not, write to the Free Software Foundation, Inc.,
'    51 Franklin Street, Fifth Floor, Boston, MA 02110-1301 USA.
'
'============================================================================

Imports System.Xml
Imports Microsoft.Office.Interop

Public Class Form1

    Dim filetoload As String = "FILENAMEHERE.xml"
    Dim pkgname As String = ""
    Dim pkgtype As String = ""
    Dim ShellEXE As String
    Dim CommandLineArgs As System.Collections.ObjectModel.ReadOnlyCollection(Of String) = My.Application.CommandLineArgs
    Dim WorkShetNumber As Integer = 1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If My.Application.CommandLineArgs.Count = 0 Then
            'No args
            MessageBox.Show("Error: No file to load. Please drag 'n drop a file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End
        Else
            'Args
            For i As Integer = 0 To CommandLineArgs.Count - 1
                ShellEXE = ShellEXE & " " & CommandLineArgs(i)
            Next
            filetoload = CommandLineArgs(0)
            Button2_Click(sender, e)
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Try
            Dim doc As System.Xml.XmlDocument = New System.Xml.XmlDocument()
            doc.Load(filetoload)
            Dim nodes As XmlNodeList = doc.SelectNodes("/*[local-name()='Servicing']/*[local-name()='cbsItem']/*[local-name()='component']")
            For Each node As XmlNode In nodes
                Dim nodename As String = node.Attributes.GetNamedItem("name").InnerText
                Dim nodever As String = node.Attributes.GetNamedItem("version").InnerText
                Dim nodearch As String = node.Attributes.GetNamedItem("processorArchitecture").InnerText
                Dim nodelang As String = node.Attributes.GetNamedItem("language").InnerText
                Dim xmltoxls As New ListViewItem
                xmltoxls.Text = nodename
                xmltoxls.SubItems.Add(nodever)
                xmltoxls.SubItems.Add(nodearch)
                xmltoxls.SubItems.Add(nodelang)
                ListView1.Items.Add(xmltoxls)
            Next
            Dim nodes2 As XmlNodeList = doc.SelectNodes("/*[local-name()='Servicing']/*[local-name()='cbsItem']")
            For Each node As XmlNode In nodes2
                pkgname = node.Attributes.GetNamedItem("name").InnerText
                pkgtype = node.Attributes.GetNamedItem("type").InnerText
            Next
            Label2.Text = pkgname & "-" & pkgtype & ".xlsx"
            Button1_Click(sender, e)
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End
        End Try
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            Dim xls As New Excel.Application
            Dim sheet As Excel.Worksheet
            Dim i As Integer
            xls.Workbooks.Add()
            sheet = xls.ActiveWorkbook.Worksheets(WorkShetNumber)
            Dim col As Integer = 1
            For j As Integer = 0 To DirectCast(ListView1, ListView).Columns.Count - 1
                sheet.Cells(1, col) = DirectCast(ListView1, ListView).Columns(j).Text.ToString
                col = col + 1
            Next
            For i = 0 To DirectCast(ListView1, ListView).Items.Count - 1
                Dim subitemscount As String = ""
                Dim columnscount As String = DirectCast(ListView1, ListView).Columns.Count
                Dim currentccount As Object = 1
                Dim currentsubcount As Integer = 0
                While currentccount <= columnscount
                    sheet.Cells(i + 2, currentccount) = DirectCast(ListView1, ListView).Items.Item(i).SubItems(currentsubcount).Text
                    currentccount = Val(currentccount) + 1
                    currentsubcount = Val(currentsubcount) + 1
                End While
            Next
            If IO.File.Exists(My.Application.Info.DirectoryPath & "\" & Label2.Text) = True Then
                MessageBox.Show("Error, the output file already exists", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End
            Else
                xls.ActiveWorkbook.SaveAs(My.Application.Info.DirectoryPath & "\" & Label2.Text)
            End If
            xls.Workbooks.Close()
            xls.Quit()
            MessageBox.Show("Finished", "CBS XML to XLSX", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End
        End Try
    End Sub
End Class
