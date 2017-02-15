'-                File Name : student.vb                    - 
'-                Part of Project: Assignment 5             - 
'------------------------------------------------------------Private
'-                Written By: Trent Killinger               - 
'-                Written On: 2-8-17                        - 
'------------------------------------------------------------ 
'- File Purpose:                                            - 
'-                                                          - 
'- This file contains the deffinition for the student class -
'------------------------------------------------------------
'- Variable Dictionary                                      - 
'- _initials - student initials                             -
'- _lastName - student last name                            -
'- _homework1 - student hw1 score                           -
'- _homework2 - student hw2 score                           -
'- _homework3 - student hw3 score                           -
'- _homework4 - student hw4 score                           -
'- _examScore - student exam score                          -
'- _numericGrade - student grade in numeric form            -
'- _letterGrade - student grade in letter form              -
'------------------------------------------------------------
Public Class Student
    Private _initials As String
    Private _lastName As String
    Private _homework1 As Single
    Private _homework2 As Single
    Private _homework3 As Single
    Private _homework4 As Single
    Private _examScore As Single
    Private _numericGrade As Single
    Private _letterGrade As String

    '------------------------------------------------------------ 
    '-                Subprogram Name: New                      - 
    '------------------------------------------------------------
    '-                Written By: Trent Killinger               - 
    '-                Written On: 2-8-17                        - 
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      - 
    '-                                                          - 
    '- This subroutine sets the data members values             -
    '------------------------------------------------------------ 
    '- Parameter Dictionary:                                    - 
    '- initials - student initials                              -
    '- lastName - student last name                             -
    '- homework1 - student hw1 score                            -
    '- homework2 - student hw2 score                            -
    '- homework3 - student hw3 score                            -
    '- homework4 - student hw4 score                            -
    '- examScore - student exam score                           -
    '------------------------------------------------------------ 
    '- Local Variable Dictionary:                               - 
    '- (None)                                                   -
    '------------------------------------------------------------
    Public Sub New(initials As String,
                   lastName As String,
                   homework1 As Single,
                   homework2 As Single,
                   homework3 As Single,
                   homework4 As Single,
                   examScore As Single)
        _initials = initials
        _lastName = lastName
        _homework1 = homework1
        _homework2 = homework2
        _homework3 = homework3
        _homework4 = homework4
        _examScore = examScore
        CalculateNumericGrade()
        CalculateLetterGrade()
    End Sub

    '------------------------------------------------------------ 
    '-                Subprogram Name: CalculateNumericGrade    - 
    '------------------------------------------------------------
    '-                Written By: Trent Killinger               - 
    '-                Written On: 2-8-17                        - 
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      - 
    '-                                                          - 
    '- This subroutine calculates student's numeric grade       -
    '------------------------------------------------------------ 
    '- Parameter Dictionary:                                    - 
    '- (None)                                                   -
    '------------------------------------------------------------ 
    '- Local Variable Dictionary:                               - 
    '- (None)                                                   -
    '------------------------------------------------------------
    Public Sub CalculateNumericGrade()
        _numericGrade = ((_homework1 + _homework2 + _homework3 + _homework4) * 0.5) + ((_examScore) * 0.5)
    End Sub

    '------------------------------------------------------------ 
    '-                Subprogram Name: CalculateLetterGrade     - 
    '------------------------------------------------------------
    '-                Written By: Trent Killinger               - 
    '-                Written On: 2-8-17                        - 
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      - 
    '-                                                          - 
    '- This subroutine calculates student's letter grade        -
    '------------------------------------------------------------ 
    '- Parameter Dictionary:                                    - 
    '- (None)                                                   -
    '------------------------------------------------------------ 
    '- Local Variable Dictionary:                               - 
    '- (None)                                                   -
    '------------------------------------------------------------
    Public Sub CalculateLetterGrade()
        If (_numericGrade >= 95) Then
            _letterGrade = "A"
        ElseIf (_numericGrade >= 90 And _numericGrade < 95) Then
            _letterGrade = "A-"
        ElseIf (_numericGrade >= 87 And _numericGrade < 90) Then
            _letterGrade = "B+"
        ElseIf (_numericGrade >= 84 And _numericGrade < 87) Then
            _letterGrade = "B"
        ElseIf (_numericGrade >= 80 And _numericGrade < 84) Then
            _letterGrade = "B-"
        ElseIf (_numericGrade >= 77 And _numericGrade < 80) Then
            _letterGrade = "C+"
        ElseIf (_numericGrade >= 74 And _numericGrade < 77) Then
            _letterGrade = "C"
        ElseIf (_numericGrade >= 70 And _numericGrade < 74) Then
            _letterGrade = "C-"
        ElseIf (_numericGrade >= 67 And _numericGrade < 70) Then
            _letterGrade = "D+"
        ElseIf (_numericGrade >= 64 And _numericGrade < 67) Then
            _letterGrade = "D"
        ElseIf (_numericGrade >= 60 And _numericGrade < 64) Then
            _letterGrade = "D-"
        Else
            _letterGrade = "F"
        End If
    End Sub

    '------------------------------------------------------------ 
    '-                Property Name: Initials                   - 
    '------------------------------------------------------------
    '-                Written By: Trent Killinger               - 
    '-                Written On: 2-1-17                        - 
    '------------------------------------------------------------
    '- Property Purpose:                                        - 
    '-                                                          - 
    '- This Property sets/gets the Initials                     -
    '------------------------------------------------------------ 
    Public Property Initials() As String
        Get
            Return _initials
        End Get
        Set(ByVal value As String)
            _initials = value
        End Set
    End Property

    '------------------------------------------------------------ 
    '-                Property Name: LastName                   - 
    '------------------------------------------------------------
    '-                Written By: Trent Killinger               - 
    '-                Written On: 2-1-17                        - 
    '------------------------------------------------------------
    '- Property Purpose:                                        - 
    '-                                                          - 
    '- This Property sets/gets the LastName                     -
    '------------------------------------------------------------ 
    Public Property LastName() As String
        Get
            Return _lastName
        End Get
        Set(ByVal value As String)
            _lastName = value
        End Set
    End Property

    '------------------------------------------------------------ 
    '-                Property Name: Homework1                  - 
    '------------------------------------------------------------
    '-                Written By: Trent Killinger               - 
    '-                Written On: 2-1-17                        - 
    '------------------------------------------------------------
    '- Property Purpose:                                        - 
    '-                                                          - 
    '- This Property sets/gets the Homework1                    -
    '------------------------------------------------------------ 
    Public Property Homework1() As Single
        Get
            Return _homework1
        End Get
        Set(ByVal value As Single)
            _homework1 = value
        End Set
    End Property

    '------------------------------------------------------------ 
    '-                Property Name: Homework2                  - 
    '------------------------------------------------------------
    '-                Written By: Trent Killinger               - 
    '-                Written On: 2-1-17                        - 
    '------------------------------------------------------------
    '- Property Purpose:                                        - 
    '-                                                          - 
    '- This Property sets/gets the Homework2                    -
    '------------------------------------------------------------ 
    Public Property Homework2() As Single
        Get
            Return _homework2
        End Get
        Set(ByVal value As Single)
            _homework2 = value
        End Set
    End Property

    '------------------------------------------------------------ 
    '-                Property Name: Homework3                  - 
    '------------------------------------------------------------
    '-                Written By: Trent Killinger               - 
    '-                Written On: 2-1-17                        - 
    '------------------------------------------------------------
    '- Property Purpose:                                        - 
    '-                                                          - 
    '- This Property sets/gets the Homework3                    -
    '------------------------------------------------------------ 
    Public Property Homework3() As Single
        Get
            Return _homework3
        End Get
        Set(ByVal value As Single)
            _homework3 = value
        End Set
    End Property

    '------------------------------------------------------------ 
    '-                Property Name: Homework4                  - 
    '------------------------------------------------------------
    '-                Written By: Trent Killinger               - 
    '-                Written On: 2-1-17                        - 
    '------------------------------------------------------------
    '- Property Purpose:                                        - 
    '-                                                          - 
    '- This Property sets/gets the Homework4                    -
    '------------------------------------------------------------ 
    Public Property Homework4() As Single
        Get
            Return _homework4
        End Get
        Set(ByVal value As Single)
            _homework4 = value
        End Set
    End Property

    '------------------------------------------------------------ 
    '-                Property Name: ExamScore                  - 
    '------------------------------------------------------------
    '-                Written By: Trent Killinger               - 
    '-                Written On: 2-1-17                        - 
    '------------------------------------------------------------
    '- Property Purpose:                                        - 
    '-                                                          - 
    '- This Property sets/gets the ExamScore                    -
    '------------------------------------------------------------ 
    Public Property ExamScore() As Single
        Get
            Return _examScore
        End Get
        Set(ByVal value As Single)
            _examScore = value
        End Set
    End Property

    '------------------------------------------------------------ 
    '-                Property Name: NumericGrade               - 
    '------------------------------------------------------------
    '-                Written By: Trent Killinger               - 
    '-                Written On: 2-1-17                        - 
    '------------------------------------------------------------
    '- Property Purpose:                                        - 
    '-                                                          - 
    '- This Property gets the NumericGrade                      -
    '------------------------------------------------------------
    Public ReadOnly Property NumericGrade() As Single
        Get
            Return _numericGrade
        End Get
    End Property

    '------------------------------------------------------------ 
    '-                Property Name: LetterGrade                - 
    '------------------------------------------------------------
    '-                Written By: Trent Killinger               - 
    '-                Written On: 2-1-17                        - 
    '------------------------------------------------------------
    '- Property Purpose:                                        - 
    '-                                                          - 
    '- This Property gets the LetterGrade                       -
    '------------------------------------------------------------
    Public ReadOnly Property LetterGrade() As String
        Get
            Return _letterGrade
        End Get
    End Property
End Class
