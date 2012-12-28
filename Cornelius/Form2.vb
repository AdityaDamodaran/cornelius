Imports iTextSharp.text.pdf
Imports System.IO
Public Class Form2
    Dim headers As String
    Dim Studobjects(200) As studentobject
    Dim dept As String
    Dim Pass As Integer = 0
    Dim Fail As Integer = 0
    Dim studcount As Integer = 0
    Private Sub Browsebtn_Click(sender As Object, e As EventArgs) Handles Browsebtn.Click
        If (OpenFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK) Then

            Dim sourcePdf As String = OpenFileDialog1.FileName
            Dim traf = New iTextSharp.text.pdf.RandomAccessFileOrArray(sourcePdf)
            Dim treader = New iTextSharp.text.pdf.PdfReader(traf, Nothing)
            Dim tpageCount = treader.NumberOfPages
            Dim i As Integer = 1
            Dim data As String = ""
            Dim tempdata As String
            tempdata = ReadPdfFile(OpenFileDialog1.FileName)
            Dim fwrite As New StreamWriter("C:\Users\Darth\Desktop\test.txt")
            fwrite.Write(tempdata)
            fwrite.Close()
            ParseToObjects(tempdata)
        End If

    End Sub
    Public Structure studentobject
        Dim rollno As String
        Dim sname As String
        Dim temp As String
        Dim Subject() As String
        Dim SubjectC() As String
        Dim SubjectA() As String
        Dim SubjectE() As String
        Dim SubjectT() As String
        Dim Aggregate As String
        Dim TPercent As String
        Dim Result As String
        Dim scount As Integer

    End Structure
    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
    Public Sub ParseToObjects(ByVal rawdata As String)
        Dim i As Integer

        i = rawdata.IndexOf("Year")
        headers = rawdata.Substring(0, i + 4 + 6)
        rawdata = rawdata.Substring(i + 10)
        rawdata = rawdata.TrimStart()
        i = headers.IndexOf("BT")
        If i = -1 Then
            i = headers.IndexOf("MT")
        End If
        dept = headers.Substring(i + 2, 2)
        i = rawdata.IndexOf(dept)
        Dim j As Integer = 0
        Dim k As Integer = 0

        While (i <> -1)
            k = rawdata.IndexOf("CIA")
            If k = -1 Then
                GoTo skip
            End If
            k = rawdata.IndexOf("CIA", k + 4)
            k = rawdata.IndexOf(dept, k + 3)
            If i <> 0 Then
                Studobjects(j).temp = rawdata.Substring(i - 9, (k - 9) - (i - 9))
            Else
                Studobjects(j).temp = rawdata.Substring(0, k - 9)

            End If
            rawdata = rawdata.Substring(k - 9)
            j = j + 1
            i = rawdata.IndexOf(dept)
        End While

skip:   Label10.Text = j.ToString
        studcount = j
        CleanData(j)
    End Sub
    Public Sub CleanData(ByVal count As Integer)
        Dim k As Integer = 0

        For i = 0 To count - 1 Step 1
            k = Studobjects(i).temp.IndexOf("Card")
            If k <> -1 Then
                Studobjects(i).temp = Studobjects(i).temp.Substring(k + 4, Studobjects(i).temp.Length - (k + 4))
                Studobjects(i).temp = Studobjects(i).temp.TrimStart()
            End If
            k = 0
            Studobjects(i).Aggregate = Studobjects(i).temp.Substring(0, 3)
            Studobjects(i).temp = Studobjects(i).temp.Substring(7, Studobjects(i).temp.Length - 7)
            Studobjects(i).temp = Studobjects(i).temp.TrimStart()
            Studobjects(i).scount = 0
step1:      Dim j As Integer = Studobjects(i).temp.IndexOf(dept)
            If j <> -1 Then
                Studobjects(i).scount = Studobjects(i).scount + 1
                j = Studobjects(i).temp.IndexOf(dept, 5)
                If j <> -1 Then
                    ReDim Preserve Studobjects(i).Subject(k)
                    Studobjects(i).Subject(k) = New String(Studobjects(i).temp.Substring(0, j))
                    Studobjects(i).temp = Studobjects(i).temp.Substring(j, Studobjects(i).temp.Length - j)
                Else
                    ReDim Preserve Studobjects(i).Subject(k)
                    Studobjects(i).Subject(k) = New String(Studobjects(i).temp.Substring(0, 8))
                    Studobjects(i).temp = Studobjects(i).temp.Substring(8, Studobjects(i).temp.Length - 8)

                End If
                k = k + 1
                GoTo step1
            End If
            j = Studobjects(i).temp.IndexOf("Distinction")

            If j = -1 Then
                j = Studobjects(i).temp.IndexOf("First Class")
                If j = -1 Then
                    j = Studobjects(i).temp.IndexOf("Pass Class")
                    If j = -1 Then
                        j = Studobjects(i).temp.IndexOf("FAILED")
                        Studobjects(i).Result = "FAILED"
                        k = Studobjects(i).temp.IndexOf("MAX")
                        Fail = Fail + 1

                        If k > j Then
                            Studobjects(i).TPercent = Studobjects(i).temp.Substring(0, j)

                        Else
                            Studobjects(i).TPercent = Studobjects(i).temp.Substring(k + 3, j - (k + 3))

                        End If
                        Studobjects(i).temp = Studobjects(i).temp.Substring(Studobjects(i).temp.IndexOf("AILED") + 5)
                    Else
                        Studobjects(i).Result = "Pass Class"
                        k = Studobjects(i).temp.IndexOf("MAX")
                        Studobjects(i).TPercent = Studobjects(i).temp.Substring(k + 3, j - (k + 3))
                        Pass = Pass + 1
                        Studobjects(i).temp = Studobjects(i).temp.Substring(Studobjects(i).temp.IndexOf("Class") + 5)
                    End If
                Else
                    Studobjects(i).Result = "First Class"
                    k = Studobjects(i).temp.IndexOf("MAX")
                    Studobjects(i).TPercent = Studobjects(i).temp.Substring(k + 3, j - (k + 3))
                    Pass = Pass + 1
                    Studobjects(i).temp = Studobjects(i).temp.Substring(Studobjects(i).temp.IndexOf("Class") + 5)
                End If
            Else
                Studobjects(i).Result = "Distinction"
                k = Studobjects(i).temp.IndexOf("MAX")
                Studobjects(i).TPercent = Studobjects(i).temp.Substring(k + 3, j - (k + 3))
                Pass = Pass + 1
                Studobjects(i).temp = Studobjects(i).temp.Substring(Studobjects(i).temp.IndexOf("ction") + 5)

            End If
            Studobjects(i).temp = Studobjects(i).temp.TrimStart()
            Dim l As Integer = 0
            For k = 0 To (Studobjects(i).scount * 2) - 1 Step 1
                l = Studobjects(i).temp.IndexOf(" ", l + 1)

            Next
            Studobjects(i).temp = Studobjects(i).temp.Substring(l)
            Studobjects(i).temp = Studobjects(i).temp.TrimStart

            'Studobjects(i).temp = Studobjects(i).temp.Substring((3 * (2 * Studobjects(i).scount)) + 3)
            Studobjects(i).rollno = Studobjects(i).temp.Substring(0, 7)
            Studobjects(i).temp = Studobjects(i).temp.Substring(7)
            Studobjects(i).temp = Studobjects(i).temp.TrimStart()
            Studobjects(i).sname = Studobjects(i).temp.Substring(0, Studobjects(i).temp.IndexOf("ESE"))
            Studobjects(i).temp = Studobjects(i).temp.Substring(Studobjects(i).temp.IndexOf("CIA") + 9)
            Studobjects(i).temp = Studobjects(i).temp.TrimStart()
            ReDim Preserve Studobjects(i).SubjectC(Studobjects(i).scount)
            k = 0
            For k = 0 To Studobjects(i).scount - 1 Step 1
                j = Studobjects(i).temp.IndexOf(" ")
                Studobjects(i).SubjectC(k) = Studobjects(i).temp.Substring(0, j)
                Studobjects(i).temp = Studobjects(i).temp.Substring(j)
                Studobjects(i).temp = Studobjects(i).temp.TrimStart()
                j = Studobjects(i).temp.IndexOf(" ")
                Studobjects(i).temp = Studobjects(i).temp.Substring(j + 1)



            Next
            Studobjects(i).temp = Studobjects(i).temp.Substring(Studobjects(i).temp.IndexOf("Marks") + 5)
            Studobjects(i).temp = Studobjects(i).temp.TrimStart()

            ReDim Preserve Studobjects(i).SubjectA(Studobjects(i).scount)

            For k = 0 To Studobjects(i).scount - 1 Step 1
                j = Studobjects(i).temp.IndexOf(" ")
                Studobjects(i).SubjectA(k) = Studobjects(i).temp.Substring(0, j)
                Studobjects(i).temp = Studobjects(i).temp.Substring(j)
                Studobjects(i).temp = Studobjects(i).temp.TrimStart()
                j = Studobjects(i).temp.IndexOf(" ")
                Studobjects(i).temp = Studobjects(i).temp.Substring(j + 1)



            Next
            Studobjects(i).temp = Studobjects(i).temp.Substring(Studobjects(i).temp.IndexOf("Marks") + 5)
            Studobjects(i).temp = Studobjects(i).temp.TrimStart()

            ReDim Preserve Studobjects(i).SubjectE(Studobjects(i).scount)

            For k = 0 To Studobjects(i).scount - 1 Step 1
                j = Studobjects(i).temp.IndexOf(" ")
                Studobjects(i).SubjectE(k) = Studobjects(i).temp.Substring(0, j)
                Studobjects(i).temp = Studobjects(i).temp.Substring(j)
                Studobjects(i).temp = Studobjects(i).temp.TrimStart()
                j = Studobjects(i).temp.IndexOf(" ")
                Studobjects(i).temp = Studobjects(i).temp.Substring(j + 1)



            Next
            Studobjects(i).temp = Studobjects(i).temp.Substring(Studobjects(i).temp.IndexOf("Mark") + 4)
            Studobjects(i).temp = Studobjects(i).temp.TrimStart()

            ReDim Preserve Studobjects(i).SubjectT(Studobjects(i).scount)
            For k = 0 To Studobjects(i).scount - 1 Step 1
                If Studobjects(i).SubjectA(k) = "" Then
                    Studobjects(i).SubjectA(k) = "0"
                End If
            Next

            For k = 0 To Studobjects(i).scount - 1 Step 1
                Studobjects(i).SubjectT(k) = Integer.Parse(Studobjects(i).SubjectA(k)) + Integer.Parse(Studobjects(i).SubjectC(k)) + Integer.Parse(Studobjects(i).SubjectE(k))
                'j = Studobjects(i).temp.IndexOf(" ")
                'Studobjects(i).SubjectT(k) = Studobjects(i).temp.Substring(0, j)
                'Studobjects(i).temp = Studobjects(i).temp.Substring(j)
                'Studobjects(i).temp = Studobjects(i).temp.TrimStart()
                'j = Studobjects(i).temp.IndexOf(" ")
                'Studobjects(i).temp = Studobjects(i).temp.Substring(j + 1)



            Next

        Next
        Label11.Text = Pass
        Label12.Text = Fail
        Label13.Text = ((Pass / (Integer.Parse(Label10.Text))) * 100).ToString
        Label14.Text = ((Fail / (Integer.Parse(Label10.Text))) * 100).ToString
        For k = 0 To studcount - 1 Step 1
            ListBox1.Items.Add(Studobjects(k).sname)
        Next

    End Sub
    Public Function ReadPdfFile(ByVal fileName As String)

        Dim text As String = ""

        If File.Exists(fileName) Then

            Dim pdfReader As New PdfReader(fileName)

            For page As Integer = 1 To pdfReader.NumberOfPages Step 1
                ' Dim its As iTextSharp.text.pdf.parser.SimpleTextExtractionStrategy
                Dim currentText As String = iTextSharp.text.pdf.parser.PdfTextExtractor.GetTextFromPage(pdfReader, page)

                text = text + currentText
            Next
            pdfReader.Close()
        End If
        Return text.ToString()
    End Function

End Class