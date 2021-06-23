VERSION 5.00
Object = "{F5044E03-D67B-4B0A-873F-0A291E923A75}#2.0#0"; "PayCal.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9130
   ClientLeft      =   60
   ClientTop       =   350
   ClientWidth     =   14950
   LinkTopic       =   "Form1"
   ScaleHeight     =   9130
   ScaleWidth      =   14950
   StartUpPosition =   3  'Windows Default
   Begin Project1.PayCal PayCal1 
      Height          =   8040
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   14430
      _ExtentX        =   25453
      _ExtentY        =   14182
   End
   Begin VB.TextBox Output1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   10
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6490
      Left            =   2520
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   1200
      Width           =   10090
   End
   Begin VB.CommandButton OutputGridStrings 
      Caption         =   "Output Strings to Immediate Window"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   610
      Left            =   9000
      TabIndex        =   4
      Top             =   8280
      Width           =   2050
   End
   Begin VB.CommandButton ToggleVisibility 
      Caption         =   "Toggle Visibility"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   610
      Left            =   6240
      TabIndex        =   3
      Top             =   8280
      Width           =   2050
   End
   Begin VB.CommandButton Calculate 
      Caption         =   "Caculate"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   610
      Left            =   3600
      TabIndex        =   2
      Top             =   8280
      Width           =   2050
   End
   Begin VB.CommandButton PrintPayslip 
      Caption         =   "Print Payslip"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   610
      Left            =   11760
      TabIndex        =   1
      Top             =   8280
      Width           =   2050
   End
   Begin VB.CommandButton SetupValues 
      Caption         =   "Set up user values"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   610
      Left            =   960
      TabIndex        =   0
      Top             =   8280
      Width           =   2170
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Michael Finnegan
'Thursday 11 June 2015
'
'PayCal.OCX is a Payroll calculation ACtiveX. It is designed for the Philippines Payroll market.
'This Demo shows you how you can control the PayCal.OCX component
'It demonstrates setting up values in the component's text boxes as if a user had typed them in.
'Then you can calculate the results and view the results in a Multi-line TextBox that you place on a form.
'You can toggle the visibility of the PayCal component in order to either show or hide your own TexBox
'You can call the calculate and print payslip methods directly.
'You can receive in a two dimensional string array the contents of the component's MSFlexGrid.
'This demo show you how you can display these strings in a multi-line TextBox
'The PayCal.OCX component must be registered on your computer.
'The component depends on files in the Database and Reports folders. These two folders should be copied
'to the same folder where you registered the PayCal.OCX component.
'If PayCal.OCX can't find these folders it will not work.

Option Explicit
Private Sub Form_Load()
    Dim x As Integer
    x = 1
    Form1.Left = (Screen.Width - Form1.Width) / 2   ' Center form horizontally.
    Form1.Top = -300 + (Screen.Height - Form1.Height) / 2  ' Center form vertically.
    PayCal1.Visible = True
    Output1.Visible = False

End Sub

'An example of setting up values in the Textboxes of the PayCal component
'All values should be as strings.
Private Sub SetupValues_Click()

    PayCal1.NumberOfDays = "14"
    PayCal1.RatePerDay = "700"
    PayCal1.OvertimePay = "1500"
    PayCal1.HolidayPay = "1000"
    PayCal1.NightDifferential = "200"
    PayCal1.allowances = "600"
    PayCal1.TardinessDeduction = "150"
    PayCal1.Dependents = "three"
End Sub

'Programatically calculating a payslip based on the values previously setup
'or based on the values the user chose.
'The PayCal component doesn't necessarily have to be visible to calculate the results.
Private Sub Calculate_Click()
    Call PayCal1.Calculate
End Sub

Private Sub PrintPayslip_Click()
    Call PayCal1.PrintPayslip
End Sub

Private Sub ToggleVisibility_Click()
    Dim b As Boolean
    b = PayCal1.Visible
    b = Not b
    PayCal1.Visible = b
    Output1.Visible = Not b

End Sub

Private Sub OutputGridStrings_Click()
'Output the contents of the MSFlexGrid inside the PayCal.ocx component
'using Debug.print to the immediate window.
'
'When you call PayCal1.FlexGridMatrix you receive a two dimentional string array
'of 26 rows and 4 columns. By analysing the array you can see the details of the
'Generated payslip.
'
    Dim r As Integer, c As Integer

    'Declare a dynamic String array. You don't need to specify the dimentions
    'At this stage it is just a String array pointer
    Dim mGrid() As String
    Dim sp As String * 40
    Dim strOut As String
    Dim LF As String
    
    PayCal1.Visible = False
    Output1.Visible = True
    
    sp = "                                        "
    mGrid = PayCal1.FlexGridMatrix
    'The dynamic string array you declared now points to a two dimentional string array
    'Its contents come from a copy of the contents of the MSFlexGrid within the PayCal.ocx

    'The following is an example of how to display the string array.
    LF = Chr$(13) & Chr$(10)
    With Output1
        .Text = ""
        strOut = Left$("Details" & sp, 33) & Right$(sp & "Employee", 14) & Right$(sp & "Employer", 14) _
               & Right$(sp & "Total", 14)
        Debug.Print strOut
        .Text = strOut & LF

        For r = 1 To 26
            For c = 1 To 4
                If c = 1 Then
                    strOut = Left$(mGrid(r, c) & sp, 33)
                    Debug.Print strOut;        'Left align 1st column
                    .Text = .Text & strOut

                Else
                    strOut = Right$(sp & mGrid(r, c), 14)
                    Debug.Print strOut;       'Right align remaining columns
                    .Text = .Text & strOut
                End If
            Next c
            Debug.Print
            .Text = .Text & LF
        Next r

        Debug.Print "-------------Finished--------------"
    End With

End Sub

