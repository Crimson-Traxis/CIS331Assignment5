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

    Public Sub New()
        _initials = ""
        _lastName = ""
        _homework1 = 0
        _homework2 = 0
        _homework3 = 0
        _homework4 = 0
    End Sub

    Public Sub CalculateNumericGrade()
        _numericGrade = ((_homework1 + _homework2 + _homework3 + _homework4) * 0.5) + ((_examScore) * 0.5)
    End Sub

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

    Public Property Initials() As String
        Get
            Return _initials
        End Get
        Set(ByVal value As String)
            _initials = value
        End Set
    End Property

    Public Property LastName() As String
        Get
            Return _lastName
        End Get
        Set(ByVal value As String)
            _lastName = value
        End Set
    End Property

    Public Property Homework1() As Single
        Get
            Return _homework1
        End Get
        Set(ByVal value As Single)
            _homework1 = value
        End Set
    End Property

    Public Property Homework2() As Single
        Get
            Return _homework2
        End Get
        Set(ByVal value As Single)
            _homework2 = value
        End Set
    End Property

    Public Property Homework3() As Single
        Get
            Return _homework3
        End Get
        Set(ByVal value As Single)
            _homework3 = value
        End Set
    End Property

    Public Property Homework4() As Single
        Get
            Return _homework4
        End Get
        Set(ByVal value As Single)
            _homework4 = value
        End Set
    End Property

    Public Property ExamScore() As Single
        Get
            Return _examScore
        End Get
        Set(ByVal value As Single)
            _examScore = value
        End Set
    End Property

    Public Property NumericGrade() As Single
        Get
            Return _numericGrade
        End Get
        Set(ByVal value As Single)
            _numericGrade = value
        End Set
    End Property

    Public Property LetterGrade() As String
        Get
            Return _letterGrade
        End Get
        Set(ByVal value As String)
            _letterGrade = value
        End Set
    End Property
End Class
