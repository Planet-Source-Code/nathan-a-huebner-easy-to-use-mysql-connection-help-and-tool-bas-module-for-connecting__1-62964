VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "MySQL Connection Tool"
   ClientHeight    =   6300
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   7005
   LinkTopic       =   "Form1"
   ScaleHeight     =   6300
   ScaleWidth      =   7005
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Click here to vote for this application"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   60
      TabIndex        =   26
      Top             =   30
      Width           =   6915
   End
   Begin VB.Frame Frame1 
      Caption         =   "Prerequisites"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   3480
      TabIndex        =   11
      Top             =   3960
      Width           =   3525
      Begin VB.CheckBox Check1 
         Caption         =   "MySQL Installed?"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   17
         Top             =   570
         Width           =   2715
      End
      Begin VB.CheckBox Check2 
         Caption         =   "MyODBC Installed?"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   16
         Top             =   840
         Width           =   2715
      End
      Begin VB.CheckBox Check3 
         Caption         =   "MySQL Tested from Prompt?"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   15
         Top             =   1110
         Width           =   2745
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Database / Tables Created?"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   14
         Top             =   1380
         Width           =   2745
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Password / Login Confirmed OK?"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   13
         Top             =   1650
         Width           =   2745
      End
      Begin VB.CheckBox Check6 
         Caption         =   "Cup of Coffee Ready?"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   12
         Top             =   1920
         Width           =   2745
      End
      Begin VB.Label Label4 
         Caption         =   "CHECKLIST BEFORE CONNECTING"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   24
         Top             =   270
         Width           =   3135
      End
      Begin VB.Label Label8 
         Caption         =   "Get it!"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   150
         TabIndex        =   23
         Top             =   600
         Width           =   585
      End
      Begin VB.Label Label9 
         Caption         =   "Get it!"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   150
         TabIndex        =   22
         Top             =   870
         Width           =   585
      End
      Begin VB.Label Label10 
         Caption         =   "Help"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   150
         TabIndex        =   21
         Top             =   1140
         Width           =   585
      End
      Begin VB.Label Label11 
         Caption         =   "Help"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   150
         TabIndex        =   20
         Top             =   1410
         Width           =   585
      End
      Begin VB.Label Label12 
         Caption         =   "Help"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   150
         TabIndex        =   19
         Top             =   1680
         Width           =   585
      End
      Begin VB.Label Label13 
         Caption         =   "Uh oh"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   150
         TabIndex        =   18
         Top             =   1950
         Width           =   585
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Run Test"
      Height          =   465
      Left            =   2490
      TabIndex        =   10
      Top             =   5760
      Width           =   915
   End
   Begin VB.TextBox MyDatabase 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1170
      TabIndex        =   9
      Top             =   5370
      Width           =   2205
   End
   Begin VB.TextBox MyPassword 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1170
      PasswordChar    =   "*"
      TabIndex        =   7
      Top             =   5070
      Width           =   2205
   End
   Begin VB.TextBox MyUser 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1170
      TabIndex        =   5
      Text            =   "root"
      Top             =   4770
      Width           =   2205
   End
   Begin VB.TextBox MyServer 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1170
      TabIndex        =   3
      Text            =   "localhost"
      Top             =   4470
      Width           =   2205
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3405
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "Main.frx":0000
      Top             =   510
      Width           =   6825
   End
   Begin VB.Label Label14 
      Caption         =   "Your login information will not be saved to hard disk, or transmitted over the internet."
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   0
      TabIndex        =   25
      Top             =   5730
      Width           =   2445
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "Database:"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   5400
      Width           =   1005
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Password:"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   5100
      Width           =   1005
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "User Name:"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   4800
      Width           =   1005
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Server:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   4500
      Width           =   1005
   End
   Begin VB.Label Label3 
      Caption         =   "If you have completed the check list, proceed to test the connection:"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   4050
      Width           =   3255
   End
   Begin VB.Menu mnu_about 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim DateStamp As String
Dim TimeStamp As String
Dim VerifyDateStamp As String
Dim VerifyTimeStamp As String
Dim RecordsCounted As Long

' Set your MySQL BAS Module with the needed login information.

MySQL_User = MyUser ' user name to MySQL (normally root or admin)
MySQL_Password = MyPassword ' password to MySQL
MySQL_Server = MyServer ' localhost is normally used
MySQL_Database = MyDatabase ' name of your database


' First test:
' Create verifyconnection table
Dim QueryString As String



' This query will create a table called verifyconnection
' If the table already exists, the operation will skip creating

    QueryString = "CREATE TABLE `verifyconnection` (" _
      & "`connectdate` varchar(25) default NULL," _
      & "`stamp` varchar(45) default NULL" _
      & ") ENGINE=MyISAM DEFAULT CHARSET=utf8;"
    
    ' Now, send your Query to MySQL
    ' Query code will automatically handle your connection
    QueryDatabase QueryString

MsgBox "So if you didn't receive an error just now, I will now continue to verify that MySQL is working (Last Step). Click OK to continue.", vbInformation, "Continue (Last Step)"

DateStamp = Date
TimeStamp = Time
' Now that the query has been run, a table should have been created
QueryString = "INSERT into verifyconnection SET connectdate='" & DateStamp & "', stamp='" & TimeStamp & "';"
QueryDatabase QueryString 'Run the query

' Now we will return the Stamp we just sent to the verification
' This will show you how to READ Data returned from MySQL
QueryString = "Select * From verifyconnection WHERE connectdate='" & DateStamp & "' AND stamp='" & TimeStamp & "';"
QueryDatabase QueryString 'Run the query

RecordsCounted = RecordCount() 'Counts the records returned
' ALL RECORDS START AT 1, NOT ZERO. :P
VerifyDateStamp = getCell("connectdate", 1) ' Returns row 1, column connectdate
VerifyTimeStamp = getCell("stamp", 1) ' Returns row 1, column stamp

If VerifyDateStamp = "" Then
MsgBox "It did not work! :( I am so sorry. Check your MySQL connection check list, login information, and server, to make sure everything looks right. Basically, the program should have created a table called verifyconnection, with 2 fields  connectdate and stamp. It placed todays date, in the connectdate field, and the time, in the stamp field. It should have returned todays date. But it didn't, that's why you got this message.", vbExclamation, "Test Failure"
Else
MsgBox "Test passed 100%. You have confirmed that MySQL can be talked to through Visual Basic. Now that you have tested it, go through my code, and take whatever you need to make your programs work. - I STORNGLY RECOMMEND INDEXING YOUR IMPORTANT COLUMNS, IF YOU WANT FAST MYSQL POWER! To index, get MySQL Administrator, or learn the ADD INDEX Syntax for MySQL. You could even send the syntax to MySQL through this program." & vbNewLine & vbNewLine & "Good luck, and good fortune!" & vbNewLine & vbNewLine & "Feel free to give this program to anyone you want, and don't forget to vote and give me some globes :) Send me e-mail if you have questions: admin@sellchain.com.", vbInformation, "Wooohoo! It worked!"
End If

End Sub

Private Sub Label10_Click()
Dim TheMsg As String
TheMsg = "So you got this far? Congrats. Now you have to test MySQL with the MySQL command line tool, that came with your MySQL installation. " & _
    vbNewLine & vbNewLine & "First, make sure MySQL installed and said that your database engine SERVICE has started and is not at a FAILED status. " & vbNewLine & vbNewLine & "To find the COMMAND LINE TOOL, go to START > PROGRAMS > MySQL > MySQL 5.0 > MySQL Command Line Client. It will ask you for a password if you have one, then type into that command line  ''USE databasename'' Where the DATABASENAME is the name of your database (don't include the quotes). You named the database when you setup the MySQL 5.0 software. That's the DATABASE name. So when you type USE <databasename> it will assume you are looking to access that specific database. Now that you are using that specific database, a response should say 'Database changed'. " & vbNewLine & vbNewLine & "Now at that prompt, you can make MySQL queries, insert records, create tables, and whatever you can think of. Now that you know the database is working, you have completed this step. You learning?" _

MsgBox TheMsg, vbInformation, "Help"

End Sub

Private Sub Label11_Click()
Dim TheMsg As String
Dim Q1 As Long
Dim Q2 As Long
Dim Q3 As Long

TheMsg = "Did you install MySQL / Did MySQL confirm the Service Has Started?"
Q1 = MsgBox(TheMsg, vbInformation + vbYesNoCancel, "Help")

    Select Case Q1 ' Returns 6 as Yes, 7 as No, 2 as Cancel
        
        Case 6
        ' Yes
        
        TheMsg = "Have you assigned your MySQL with a Database name? (it asked you during setup)"
        Q2 = MsgBox(TheMsg, vbInformation + vbYesNoCancel, "Next Step")
            
            Select Case Q2
                
                Case 6
                MsgBox "You are ready to go then. The MySQL test tool will automatically create a table with the database name you specified, so it can test the connection. It will create table name 'verifyconnection'.", vbInformation, "Steps complete"
                Exit Sub
                
                Case 7
                MsgBox "Then you need to create a database at the MySQL Command Line Client, or by using a MySQL Administration tool. You should have specified the database name during setup. I recommend going back through your setup, and figuring out how to do this. It takes too much time for this step to be done manually. And I don't want to explain how to do the MySQL code, it could mean a lot of troubleshooting. Your best bet is to reinstall and choose the mysql database name." & vbNewLine & vbNewLine & "Also, you could try using the database MySQL creates called 'mysql'. Use this database if you have given up.", vbInformation, "Uh oh!"
                Exit Sub
                
                Case 2
                Exit Sub
            End Select
            
        Case 7
        ' No
        MsgBox "Silly, you shouldn't skip steps :) Go and install the MySQL database at your MYSQL CHECKLIST!!", vbCritical, "...Silly you"
        
        Case 2
        ' Cancel
        
    End Select
    
End Sub

Private Sub Label12_Click()
Dim TheMsg As String
TheMsg = "A lot of people go crazy because their login information is incorrect, while their MySQL connection would work fine if they put it correctly. Be sure your login information is 100% correct before testing, so you don't assume the tool is broken. This tool has been tested on 3 seperate systems, all running Windows 2003, PHP, MySQL, MyODBC." & vbNewLine & vbNewLine & "You specified the Administrator Password during MySQL 5.0 Setup. " & vbNewLine & vbNewLine & "MySQL Default User Name: root"
MsgBox TheMsg, vbInformation, "Help"

End Sub

Private Sub Label13_Click()
MsgBox "MySQL good, coffee bad! Go get me some coffee!", vbCritical, "WARNING: No Coffee = APPLICATION FAILURE! ABORT! ABORT!"

End Sub

Private Sub Label8_Click()
' Opens a window to download MySQL 5.0
' Note: This has been tested with 5.0 and works fine.

frmBrowser.Show
frmBrowser.brwWebBrowser.Navigate "http://dev.mysql.com/get/Downloads/MySQL-5.0/mysql-5.0.13-rc-win32.zip/from/pick#mirrors"
frmBrowser.Width = Screen.Width - 200
frmBrowser.Left = 0

End Sub

Private Sub Label9_Click()
' Opens a window to download MyODBC
' MyODBC aka MySQL ODBC
' NOTE: This program uses MyODBC 3.51.10, however the link downloads 3.51.12
' If you have problems using MyODBC 3.51.12, then get 3.51.10.
' You can figure it out im sure ;) Just go to mysql.com, and look for the
' MySQL Tools link on the home page, it will take you there.
' Look for the MySQL ODBC Connectors

frmBrowser.Show
frmBrowser.brwWebBrowser.Navigate "http://dev.mysql.com/get/Downloads/MyODBC3/mysql-connector-odbc-3.51.12-win32.zip/from/pick#mirrors"
frmBrowser.Width = Screen.Width - 200
frmBrowser.Left = 0

End Sub

Private Sub mnu_about_Click()
frmAbout.Show
End Sub
