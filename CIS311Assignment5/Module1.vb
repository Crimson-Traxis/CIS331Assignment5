'-                File Name : Modual1.vb                    - 
'-                Part of Project: Assignment 5             - 
'------------------------------------------------------------
'-                Written By: Trent Killinger               - 
'-                Written On: 2-8-17                        - 
'------------------------------------------------------------ 
'- File Purpose:                                            - 
'-                                                          - 
'- This file contains the main function that is ran when the-
'- program is started.
'------------------------------------------------------------
'- Variable Dictionary                                      - 
'- (none)                                                   -
'------------------------------------------------------------

Imports System.IO
Imports System.Text

Module Module1

    Private Const maxHomeworkGrade As Integer = 25
    Private Const maxExamGrade As Integer = 100

    '------------------------------------------------------------ 
    '-                Subprogram Name: Main                     - 
    '------------------------------------------------------------
    '-                Written By: Trent Killinger               - 
    '-                Written On: 2-8-17                        - 
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      - 
    '-                                                          - 
    '- This subroutine does the following                       -
    '- *Prints Grade Report                                     -
    '- *Prints Grade Distribution Statistics                    -
    '- *Prints Homework/Exam Range Statistics                   -
    '- *Prints Overall Course Grade Statistics                  -
    '------------------------------------------------------------ 
    '- Parameter Dictionary:                                    - 
    '- (None)                                                   - 
    '------------------------------------------------------------ 
    '- Local Variable Dictionary:                               - 
    '- studentList - list of students read from user specifed   -
    '-               file                                       -
    '- data - string representation for file                    -
    '- output - StringBuilder object that stores output         -
    '- semesterReport - list of students sorted decending       -
    '- gradeList - list of grades to look for                   -
    '- gradeReport - list of students found matching grade      -
    '- HW1 - stores stats of homework 1                         -
    '- HW2 - stores stats of homework 2                         -
    '- HW3 - stores stats of homework 3                         -
    '- HW4 - stores stats of homework 4                         -
    '- Exam - stores stats of exams                             -
    '- highestGrade - highest grade in class                    -
    '- lowestGrade - lowest grade in class                      -
    '- courseAverage - course average                           -
    '------------------------------------------------------------
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
            output.AppendLine("               Low                      Ave                      High")
            Dim HW1 = New With {
                          .min = (From student In studentList Select student.Homework1).Min(),
                          .max = (From student In studentList Select student.Homework1).Max(),
                          .avg = (From student In studentList Select student.Homework1).Average()
                      }
            Dim HW2 = New With {
                          .min = (From student In studentList Select student.Homework2).Min(),
                          .max = (From student In studentList Select student.Homework2).Max(),
                          .avg = (From student In studentList Select student.Homework2).Average()
                      }
            Dim HW3 = New With {
                          .min = (From student In studentList Select student.Homework3).Min(),
                          .max = (From student In studentList Select student.Homework3).Max(),
                          .avg = (From student In studentList Select student.Homework3).Average()
                      }
            Dim HW4 = New With {
                          .min = (From student In studentList Select student.Homework4).Min(),
                          .max = (From student In studentList Select student.Homework4).Max(),
                          .avg = (From student In studentList Select student.Homework4).Average()
                      }
            Dim Exam = New With {
                          .min = (From student In studentList Select student.ExamScore).Min(),
                          .max = (From student In studentList Select student.ExamScore).Max(),
                          .avg = (From student In studentList Select student.ExamScore).Average()
                      }
            output.AppendLine(String.Format("Homework 1 :{0,9:P2}{1,25:P2}{2,25:P2}",
                                            HW1.min / maxHomeworkGrade,
                                            HW1.avg / maxHomeworkGrade,
                                            HW1.max / maxHomeworkGrade))
            output.AppendLine(String.Format("Homework 2 :{0,9:P2}{1,25:P2}{2,25:P2}",
                                            HW2.min / maxHomeworkGrade,
                                            HW2.avg / maxHomeworkGrade,
                                            HW2.max / maxHomeworkGrade))
            output.AppendLine(String.Format("Homework 3 :{0,9:P2}{1,25:P2}{2,25:P2}",
                                            HW3.min / maxHomeworkGrade,
                                            HW3.avg / maxHomeworkGrade,
                                            HW3.max / maxHomeworkGrade))
            output.AppendLine(String.Format("Homework 4 :{0,9:P2}{1,25:P2}{2,25:P2}",
                                            HW4.min / maxHomeworkGrade,
                                            HW4.avg / maxHomeworkGrade,
                                            HW4.max / maxHomeworkGrade))
            output.AppendLine(String.Format("Exam       :{0,9:P2}{1,25:P2}{2,25:P2}",
                                            Exam.min / maxExamGrade,
                                            Exam.avg / maxExamGrade,
                                            Exam.max / maxExamGrade))
            output.AppendLine("")
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
            output.AppendLine("")
            Dim courseAverage = (From students In studentList Select students.NumericGrade).Average()
            output.AppendLine(String.Format("The average course grade was {0}", courseAverage))
            Console.WriteLine(output.ToString())
        End If
        Console.WriteLine("Press Any Key to Exit")
        Console.ReadKey()
    End Sub

End Module
