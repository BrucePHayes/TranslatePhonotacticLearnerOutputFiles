VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Interpret output of UCLA Phonotactic Learng"
   ClientHeight    =   3990
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   5685
   LinkTopic       =   "Form1"
   ScaleHeight     =   3990
   ScaleWidth      =   5685
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Go"
      Height          =   1335
      Left            =   1560
      TabIndex        =   0
      Top             =   1320
      Width           =   2415
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

    'Translate the difficult-to-read constraint files of the UCLA Phonotactic Learner
    
    'Module-level variables.
        Dim mNumberOfNaturalClasses As Long
        Dim mSizeOfNaturalClasses() As Long
        Dim mSegmentalDescriptions() As String
        Dim mFeaturalDescriptions() As String
        Dim mFeatureCount() As Long
        Dim mExplication As String


Private Sub Command1_Click()

    Dim MatrixCount As Long
    Dim MyTier As String
    Dim MyNaturalClass As String
    Dim TranslatedConstraint As String
    
    Dim MyLine As String, Buffer As String
    Dim GrammarFile As Long, NatClFile As Long, OutFile As Long
    Dim NaturalClassIndex As Long, ConstraintIndex As Long
    
    'Digest the natural classes.
        'Open file.
            Let NatClFile = FreeFile
            Open App.Path + "/NatClassesFile.txt" For Input As #NatClFile
        'Digest.
            Do While Not EOF(NatClFile)
                Line Input #NatClFile, MyLine
                Let Buffer = MyLine
                Let mNumberOfNaturalClasses = mNumberOfNaturalClasses + 1
                ReDim Preserve mFeaturalDescriptions(mNumberOfNaturalClasses)
                ReDim Preserve mFeatureCount(mNumberOfNaturalClasses)
                ReDim Preserve mSegmentalDescriptions(mNumberOfNaturalClasses)
                ReDim Preserve mSizeOfNaturalClasses(mNumberOfNaturalClasses)
                'Grab the featural description.
                    Let Buffer = s.CustomResidue(MyLine, ": ")
                    Let mFeaturalDescriptions(mNumberOfNaturalClasses) = s.Chomp(Buffer)
                    Let mFeatureCount(mNumberOfNaturalClasses) = CountCommasPlus(mFeaturalDescriptions(mNumberOfNaturalClasses))
                    'MsgBox "Features, surrounded by brackets:  [" + mFeaturalDescriptions(mNumberOfNaturalClasses) + "]"
                'Grab the segmental descriptions.
                    Let mSegmentalDescriptions(mNumberOfNaturalClasses) = s.Residue(Buffer)
                    Let mSizeOfNaturalClasses(mNumberOfNaturalClasses) = CountCommasPlus(mSegmentalDescriptions(mNumberOfNaturalClasses))
                    'MsgBox "Segments, surrounded by brackets:  [" + mSegmentalDescriptions(mNumberOfNaturalClasses) + "]"
            Loop
        'Close.
            Close #NatClFile
        
        'Debug.
            Let OutFile = FreeFile
            Open App.Path + "/debug.txt" For Output As #OutFile
            For NaturalClassIndex = 1 To mNumberOfNaturalClasses
                Print #OutFile, mFeaturalDescriptions(NaturalClassIndex); Chr(9); mSegmentalDescriptions(NaturalClassIndex); Chr(9); mSizeOfNaturalClasses(NaturalClassIndex)
            Next NaturalClassIndex
        
            Close OutFile
        
    'Digest and translate the constraints.
        'Open files.
            Let GrammarFile = FreeFile
            Open App.Path + "/Grammar.txt" For Input As #GrammarFile
            Let OutFile = FreeFile
            Open App.Path + "/InterpretedGrammar.txt" For Output As #OutFile
        'Go through grammar file.
            Do While Not EOF(GrammarFile)
                'Initialize.
                    Let MatrixCount = 0
                    Let TranslatedConstraint = "*"
                    Let mExplication = ""
                'Grab the next constraint.
                    Line Input #GrammarFile, MyLine
                'Put the featural description in a mutilable variable.
                    Let Buffer = s.Chomp(MyLine)
                'Also grab the tier.
                    Let MyTier = s.Chomp(s.Residue(MyLine))
                'Ignore the initial asterisk and the initial bracket.
                    Let Buffer = Mid(Buffer, 3)
                'Iteratively grab the natural classes and form the translation.
                    Do
                        Let MyNaturalClass = s.CustomChomp(Buffer, "]")
                        'Add C0 for vowel tier constraints.
                            If LCase(MyTier) = "(tier=vowel)" Then
                                Let TranslatedConstraint = TranslatedConstraint + " Co "
                            End If
                        Let TranslatedConstraint = TranslatedConstraint + Translation(MyNaturalClass)
                        Let MatrixCount = MatrixCount + 1
                        'Throw away this featural description to access next one.
                            Let Buffer = s.CustomResidue(Buffer, "]")
                            'Discard the next left bracket.
                                Let Buffer = Trim(Mid(Buffer, 2))
                            'Terminate if appropriate.
                                If Buffer = "" Then Exit Do
                    Loop
                'Some last-minute cleanups.
                    'No need for C0 initially.
                        Let TranslatedConstraint = Replace(TranslatedConstraint, "* Co ", "*")
                    'Excess spaces
                        Let TranslatedConstraint = Replace(TranslatedConstraint, "* ", "*")
                        Let TranslatedConstraint = Replace(TranslatedConstraint, "   ", " ")
                        Let TranslatedConstraint = Replace(TranslatedConstraint, "  ", " ")
                'Print what you learned, along with the original.
                    Print #OutFile, TranslatedConstraint; Chr(9); Trim(Str(MatrixCount)); Chr(9); MyLine;
                    If mExplication = "" Then
                        Print #OutFile,
                    Else
                        Print #OutFile, Chr(9); mExplication
                    End If
            Loop
        'Close files
            Close #GrammarFile
            Close #OutFile
    
    End
        
End Sub

Function Translation(FeaturalDescription As String) As String
    
    Dim CorrectSegmentalDescription As String
    Dim NaturalClassIndex As Long
    
    'Handle a few cases with conventional abbreviations.
        If FeaturalDescription = "+word_boundary" Then
            Let Translation = " # "
            Exit Function
        ElseIf FeaturalDescription = "-word_boundary" Then
            Let Translation = "[X] "
            Exit Function
        ElseIf FeaturalDescription = "+syllabic" Then
            Let Translation = " V "
            Exit Function
        ElseIf FeaturalDescription = "-syllabic" Then
            Let Translation = " C "
            Exit Function
        Else
            'First, find the segmental description.
                'Loop through the list of featural descriptions and find the current one.
                    For NaturalClassIndex = 1 To mNumberOfNaturalClasses
                        If mFeaturalDescriptions(NaturalClassIndex) = FeaturalDescription Then
                            Let CorrectSegmentalDescription = mSegmentalDescriptions(NaturalClassIndex)
                            Exit For
                        End If
                    Next NaturalClassIndex
                'If it's just one segment, use that, even if it can be identified with a single feature.
                    If CountCommasPlus(CorrectSegmentalDescription) = 1 Then
                        Let Translation = Translation + " " + CorrectSegmentalDescription
                        Exit Function
                    Else
                        'If the featural description is one feature, use that, but explain the natural class in the comment field.
                            If CountCommasPlus(FeaturalDescription) = 1 Then
                                Let Translation = Translation + "[" + FeaturalDescription + "] "
                                Let mExplication = mExplication + "[" + FeaturalDescription + "] means {" + mSegmentalDescriptions(NaturalClassIndex) + "}.    "
                                Exit Function
                            Else
                                'Otherwise, the features are are likely to be hard to read, so just spell out the list.
                                    Let Translation = Translation + "{" + mSegmentalDescriptions(NaturalClassIndex) + "} "
                                    Exit Function
                            End If
                    End If          'One segment, or more?
        End If                      'Special case, or the general procedure?
        
    'You should never get this far.
        MsgBox "Error: I cannot find a listing in NatClassesFile.txt for the natural class [" + FeaturalDescription + "]."
        Let Translation = "ERROR"
        
End Function


Function CountCommasPlus(MyList As String) As Long

    Dim Buffer As Long, i As Long
    For i = 1 To Len(MyList)
        If Mid(MyList, i, 1) = "," Then
            Let Buffer = Buffer + 1
        End If
    Next i
    Let CountCommasPlus = Buffer + 1

End Function

Private Sub mnuHelp_Click()

    Dim Buffer As String
    
    Let Buffer = "This is an auxiliary program to help you read the grammar output file of the UCLA Phonotactic Learner.  It converts feature matrices to segment lists, prints constraints on the"
    Let Buffer = Buffer + " Vowel projection into strings with SPE-style Co, and converts word boundaries into the symbol #. It retains single-feature matrices, but identifies the segments they denote "
    Let Buffer = Buffer + "in a comment field.  The output file for this program is placed in the same folder as the others.  It is a spreadsheet in tab-delimited-text format, named InterpretedGrammar.txt.  You can read it with a spreadsheet program like Excel."
    Let Buffer = Buffer + vbCr + vbLf + vbCr + vbLf
    Let Buffer = Buffer + "To run this program, place the program (named TranslatePhonotacticLearnerNaturalClasses.exe) in the same folder where the UCLA Phonotactic Learner placed its output files "
    Let Buffer = Buffer + "NatClassesFile.txt and grammar.txt.  Click on Go, and then this program will read these two files and use them to produce a hopefully more legible translation in InterpretedGrammar.txt."
    Let Buffer = Buffer + vbCr + vbLf + vbCr + vbLf
    Let Buffer = Buffer + "Suggestions for improvement to Bruce Hayes, bhayes@humnet.ucla.edu."
    
    MsgBox Buffer

End Sub
