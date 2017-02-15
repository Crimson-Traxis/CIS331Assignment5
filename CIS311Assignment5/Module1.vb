Imports System.IO
Imports System.Text

Module Module1

    Sub Main()
        Dim studentList As List(Of Student) = New List(Of Student)
        Dim data As String = ""
        Console.WriteLine("Please specify location of file to load.")
        Console.Write("Location: ")
        Try
            data = File.ReadAllText(Console.ReadLine())
        Catch ex As Exception
            Console.WriteLine(ex.Message)
        End Try
        If (data <> "") Then
            Dim reader As StringReader = New StringReader(data)
            Dim line As String = reader.ReadLine()
            While line <> Nothing
                Try
                    Dim args As String() = line.Split(" ")
                    studentList.Add(New Student(args(0),
                                                args(1),
                                                Convert.ToSingle(args(2)),
                                                Convert.ToSingle(args(3)),
                                                Convert.ToSingle(args(4)),
                                                Convert.ToSingle(args(5)),
                                                Convert.ToSingle(args(6))))
                Catch ex As Exception
                    Console.WriteLine(ex.Message)
                End Try
                line = reader.ReadLine()
            End While

            Dim output As StringBuilder = New StringBuilder()
            output.Append(Environment.NewLine)
            output.AppendLine("                            Ye Old Country School")
            output.AppendLine("                        *** Semester Grade Report ***")
            output.AppendLine("                        -----------------------------")
            output.Append(Environment.NewLine)
            output.AppendLine("                         Homework Scores          Exam    Numeric  Letter")
            output.AppendLine("      Name          1       2       3       4     Score   Grade    Grade")
            output.AppendLine("--------------    -----   -----   -----   -----   ------  -------  ------")
            Dim semesterReport = From students In studentList
                                 Order By students.LastName Ascending
                                 Select students
            For Each student As Student In semesterReport
                output.AppendLine(String.Format("{0,-5}{1,-13}{2,-5:#.00}{3,8:#.00}{4,8:#.00}{5,8:#.00}{6,9:#.00}{7,9:#.00}     {8,-2}",
                                                student.Initials,
                                                student.LastName,
                                                student.Homework1,
                                                student.Homework2,
                                                student.Homework3,
                                                student.Homework4,
                                                student.ExamScore,
                                                student.NumericGrade,
                                                student.LetterGrade))
            Next
            output.AppendLine("-------------------------------------------------------------------------")
            output.AppendLine("                      Grade Distribution Statistics")
            output.AppendLine("-------------------------------------------------------------------------")
            output.AppendLine("")
            Dim gradeList As List(Of String) = {"A", "B", "C", "D", "F"}.ToList()
            For Each grade As String In gradeList
                output.AppendLine("Those students earning a " + grade + " grade are:")
                Dim gradeReport = From students In studentList
                                  Order By students.LastName Ascending
                                  Select students Where students.LetterGrade(0) = grade
                For Each student As Student In gradeReport
                    output.AppendLine(String.Format("{0,-5}{1,-13}--> {2,-6:#.00}", student.Initials, student.LastName, student.LetterGrade))
                Next
                If (gradeReport.Count = 0) Then
                    output.AppendLine("None")
                End If
                output.AppendLine("")
            Next
            output.AppendLine("-------------------------------------------------------------------------")
            output.AppendLine("                 Homework/Exam Grade Range Statistics")
            output.AppendLine("-------------------------------------------------------------------------")
            output.AppendLine("")
            Dim HW1 = New With {
                          Key .min = studentList.Min(Function(x) x.Homework1),
                          Key .max = studentList.Max(Function(x) x.Homework1),
                          Key .avg = studentList.Average(Function(x) x.Homework1)
                      }
            Dim HW2 = New With {
                          Key .min = studentList.Min(Function(x) x.Homework2),
                          Key .max = studentList.Max(Function(x) x.Homework2),
                          Key .avg = studentList.Average(Function(x) x.Homework2)
                      }
            Dim HW3 = New With {
                          Key .min = studentList.Min(Function(x) x.Homework3),
                          Key .max = studentList.Max(Function(x) x.Homework3),
                          Key .avg = studentList.Average(Function(x) x.Homework3)
                      }
            Dim HW4 = New With {
                          Key .min = studentList.Min(Function(x) x.Homework4),
                          Key .max = studentList.Max(Function(x) x.Homework4),
                          Key .avg = studentList.Average(Function(x) x.Homework4)
                      }
            Dim Exam = New With {
                          Key .min = studentList.Min(Function(x) x.ExamScore),
                          Key .max = studentList.Max(Function(x) x.ExamScore),
                          Key .avg = studentList.Average(Function(x) x.ExamScore)
                      }
            Dim highestGrade = From students In studentList
                               Order By students.NumericGrade Descending
                               Select students
            output.AppendLine("The highest course grade of " + highestGrade.First().NumericGrade.ToString("#.00") + " was eraned by:")
            output.AppendLine(String.Format("{0,-5}{1,-13}--> {2,-6:#.00}", highestGrade.First().Initials, highestGrade.First().LastName, highestGrade.First().NumericGrade))
            output.AppendLine("")
            Dim lowestGrade = From students In studentList
                              Order By students.NumericGrade Ascending
                              Select students
            output.AppendLine("The lowest course grade of " + lowestGrade.First().NumericGrade.ToString("#.00") + " was eraned by:")
            output.AppendLine(String.Format("{0,-5}{1,-13}--> {2,-6:#.00}", lowestGrade.First().Initials, lowestGrade.First().LastName, lowestGrade.First().NumericGrade))
            Console.WriteLine(output.ToString())
        End If
    End Sub

End Module
