VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "INI Expert"
   ClientHeight    =   6045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9330
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   9330
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtParams 
      Height          =   1395
      Left            =   6000
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   26
      Top             =   3000
      Width           =   3255
   End
   Begin VB.TextBox txtDescription 
      Height          =   1395
      Left            =   6000
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   23
      Top             =   1140
      Width           =   3255
   End
   Begin VB.CommandButton cmdFunction 
      Caption         =   "GetKeyIndex2"
      Height          =   555
      Index           =   20
      Left            =   3900
      TabIndex        =   21
      Top             =   4380
      Width           =   1875
   End
   Begin VB.CommandButton cmdFunction 
      Caption         =   "GetKeyIndex"
      Height          =   555
      Index           =   19
      Left            =   3900
      TabIndex        =   20
      Top             =   3780
      Width           =   1875
   End
   Begin VB.CommandButton cmdFunction 
      Caption         =   "GetSectionIndex"
      Height          =   555
      Index           =   18
      Left            =   3900
      TabIndex        =   19
      Top             =   3180
      Width           =   1875
   End
   Begin VB.CommandButton cmdFunction 
      Caption         =   "SectionExists"
      Height          =   555
      Index           =   17
      Left            =   3900
      TabIndex        =   18
      Top             =   2580
      Width           =   1875
   End
   Begin VB.CommandButton cmdFunction 
      Caption         =   "KeyExists2"
      Height          =   555
      Index           =   16
      Left            =   3900
      TabIndex        =   17
      Top             =   1980
      Width           =   1875
   End
   Begin VB.CommandButton cmdFunction 
      Caption         =   "KeyExists"
      Height          =   555
      Index           =   15
      Left            =   3900
      TabIndex        =   16
      Top             =   1380
      Width           =   1875
   End
   Begin VB.CommandButton cmdFunction 
      Caption         =   "IsSection"
      Height          =   555
      Index           =   14
      Left            =   3900
      TabIndex        =   15
      Top             =   780
      Width           =   1875
   End
   Begin VB.CommandButton cmdFunction 
      Caption         =   "IsKey"
      Height          =   555
      Index           =   13
      Left            =   1980
      TabIndex        =   14
      Top             =   4380
      Width           =   1875
   End
   Begin VB.CommandButton cmdFunction 
      Caption         =   "GetSection"
      Height          =   555
      Index           =   12
      Left            =   1980
      TabIndex        =   13
      Top             =   3780
      Width           =   1875
   End
   Begin VB.CommandButton cmdFunction 
      Caption         =   "GetKey2"
      Height          =   555
      Index           =   11
      Left            =   1980
      TabIndex        =   12
      Top             =   3180
      Width           =   1875
   End
   Begin VB.CommandButton cmdFunction 
      Caption         =   "GetKey"
      Height          =   555
      Index           =   10
      Left            =   1980
      TabIndex        =   11
      Top             =   2580
      Width           =   1875
   End
   Begin VB.CommandButton cmdFunction 
      Caption         =   "RenameKey"
      Height          =   555
      Index           =   9
      Left            =   1980
      TabIndex        =   10
      Top             =   1980
      Width           =   1875
   End
   Begin VB.CommandButton cmdFunction 
      Caption         =   "RenameSection"
      Height          =   555
      Index           =   8
      Left            =   1980
      TabIndex        =   9
      Top             =   1380
      Width           =   1875
   End
   Begin VB.CommandButton cmdFunction 
      Caption         =   "NumKeys"
      Height          =   555
      Index           =   7
      Left            =   1980
      TabIndex        =   8
      Top             =   780
      Width           =   1875
   End
   Begin VB.CommandButton cmdFunction 
      Caption         =   "TotalKeys"
      Height          =   555
      Index           =   6
      Left            =   60
      TabIndex        =   7
      Top             =   4380
      Width           =   1875
   End
   Begin VB.CommandButton cmdFunction 
      Caption         =   "TotalSections"
      Height          =   555
      Index           =   5
      Left            =   60
      TabIndex        =   6
      Top             =   3780
      Width           =   1875
   End
   Begin VB.CommandButton cmdFunction 
      Caption         =   "DeleteKeyValue"
      Height          =   555
      Index           =   4
      Left            =   60
      TabIndex        =   5
      Top             =   3180
      Width           =   1875
   End
   Begin VB.CommandButton cmdFunction 
      Caption         =   "DeleteKey"
      Height          =   555
      Index           =   3
      Left            =   60
      TabIndex        =   4
      Top             =   2580
      Width           =   1875
   End
   Begin VB.CommandButton cmdFunction 
      Caption         =   "DeleteSection"
      Height          =   555
      Index           =   2
      Left            =   60
      TabIndex        =   3
      Top             =   1980
      Width           =   1875
   End
   Begin VB.CommandButton cmdFunction 
      Caption         =   "AddToINI"
      Height          =   555
      Index           =   1
      Left            =   60
      TabIndex        =   2
      Top             =   1380
      Width           =   1875
   End
   Begin VB.CommandButton cmdFunction 
      Caption         =   "GetKeyVal"
      Height          =   555
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   780
      Width           =   1875
   End
   Begin VB.Line Line1 
      X1              =   1320
      X2              =   8040
      Y1              =   5580
      Y2              =   5580
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Functions affect, INIFile.ini located in this program's folder."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   60
      TabIndex        =   29
      Top             =   5700
      Width           =   9210
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   $"Form1.frx":0000
      Height          =   435
      Left            =   120
      TabIndex        =   28
      Top             =   5040
      Width           =   5415
   End
   Begin VB.Label lblReturn 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6000
      TabIndex        =   27
      Top             =   5040
      Width           =   3195
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Parameters:"
      Height          =   195
      Left            =   6000
      TabIndex        =   25
      Top             =   2700
      Width           =   840
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Function Return:"
      Height          =   195
      Left            =   7020
      TabIndex        =   24
      Top             =   4680
      Width           =   1185
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Description:"
      Height          =   195
      Left            =   6000
      TabIndex        =   22
      Top             =   840
      Width           =   840
   End
   Begin VB.Label Label1 
      Caption         =   $"Form1.frx":008B
      Height          =   435
      Left            =   60
      TabIndex        =   1
      Top             =   180
      Width           =   9270
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------
'-------------------- READ ME ---------------------
'--------------------------------------------------
'There are two ways to call a function.
'1. Intending for a return value  (ex.  ReturnValue = GetKeyVal(FileName, "Section 4", "Key 3")
'2. Just to execute the code  (ex.  DeleteSection FileName, "Section 1"
'All of my functions return a value, but you do not need to do the
'extra work to get that value if you do not wish to.  All functions that do
'not need a return value I have marked with 3 asterisk. *** followed by the
'alternate way to call the function.
'-Functions get called with 1 line of code.

Private Sub cmdFunction_Click(Index As Integer)
Dim FileName As String
FileName = App.Path & "\INIFile.ini"

Select Case Index
    Case 0:
        'GetKeyVal() - Has to Return
        lblReturn.Caption = GetKeyVal(FileName, "Section 4", "Key 3")
        
        txtDescription.Text = "This function retrieves a Key's Value from an INI file." & vbCrLf & vbCrLf & "Function returns: String"
        txtParams.Text = "FileName: String" & vbCrLf & "The location of the INI File." & _
        vbCrLf & vbCrLf & "Section: String" & vbCrLf & "The name of the Section where the Key is." & _
        vbCrLf & vbCrLf & "Key: String" & vbCrLf & "The name of the Key whos value you want to retrieve."
    Case 1:
        'AddToINI() *** AddToINI FileName, "New Section", "Key 1", "Key 1 Value"
        lblReturn.Caption = AddToINI(FileName, "New Section", "Key 1", "Key 1 Value")
        
        txtDescription.Text = "This function can do one of many things.  It can create a new INI File, add a new Section, add a new Key, or change a Key Value." & vbCrLf & vbCrLf & "Function returns: 1 or 0" & vbCrLf & "1=Worked  0=Unsuccessful"
        txtParams.Text = "FileName: String" & vbCrLf & "The location of the INI File." & _
        vbCrLf & vbCrLf & "Section: String" & vbCrLf & "The name of the Section." & _
        vbCrLf & vbCrLf & "Key: String" & vbCrLf & "The name of the Key." & _
        vbCrLf & vbCrLf & "KeyValue: String" & vbCrLf & "The value of the preceeding key."
    Case 2:
        'DeleteSection() *** DeleteSection FileName, "Section 1"
        lblReturn.Caption = DeleteSection(FileName, "Section 1")

        txtDescription.Text = "This function deletes a Section and all of it's Keys." & vbCrLf & vbCrLf & "Function returns: 1 or 0" & vbCrLf & "1=Worked  0=Unsuccessful"
        txtParams.Text = "FileName: String" & vbCrLf & "The location of the INI File." & _
        vbCrLf & vbCrLf & "Section: String" & vbCrLf & "The name of the Section to be deleted."
    Case 3:
        'DeleteKey() *** DeleteKey FileName, "Section 3", "Key 5"
        lblReturn.Caption = DeleteKey(FileName, "Section 3", "Key 5")
        
        txtDescription.Text = "This function deletes a Key from a Section." & vbCrLf & vbCrLf & "Function returns: 1 or 0" & vbCrLf & "1=Worked  0=Unsuccessful"
        txtParams.Text = "FileName: String" & vbCrLf & "The location of the INI File." & _
        vbCrLf & vbCrLf & "Section: String" & vbCrLf & "The name of the Section where the Key is." & _
        vbCrLf & vbCrLf & "Key: String" & vbCrLf & "The name of the Key you wish to delete."
    Case 4:
        'DeleteKeyValue() *** DeleteKeyValue FileName, "Section 4", "Key 4"
        lblReturn.Caption = DeleteKeyValue(FileName, "Section 4", "Key 4")
        
        txtDescription.Text = "This function deletes a Value from a Key." & vbCrLf & vbCrLf & "Function returns: 1 or 0" & vbCrLf & "1=Worked  0=Unsuccessful"
        txtParams.Text = "FileName: String" & vbCrLf & "The location of the INI File." & _
        vbCrLf & vbCrLf & "Section: String" & vbCrLf & "The name of the Section where the Key is." & _
        vbCrLf & vbCrLf & "Key: String" & vbCrLf & "The name of the Key whos Value you wish to delete."
    Case 5:
        'TotalSections() - Has to Return
        lblReturn.Caption = TotalSections(FileName)
        
        txtDescription.Text = "This counts the total number of Sections in an INI file." & vbCrLf & vbCrLf & "Function returns: Integer"
        txtParams.Text = "FileName: String" & vbCrLf & "The location of the INI File."
    Case 6:
        'TotalKeys() - Has to Return
        lblReturn.Caption = TotalKeys(FileName)
        
        txtDescription.Text = "This counts the total number of Keys in an INI file." & vbCrLf & vbCrLf & "Function returns: Integer"
        txtParams.Text = "FileName: String" & vbCrLf & "The location of the INI File."
    Case 7:
        'NumKeys() - Has to Return
        lblReturn.Caption = NumKeys(FileName, "Section 5")
        
        txtDescription.Text = "This counts the total number of Keys in a single Section." & vbCrLf & vbCrLf & "Function returns: Integer"
        txtParams.Text = "FileName: String" & vbCrLf & "The location of the INI File." & _
        vbCrLf & vbCrLf & "Section: String" & vbCrLf & "The name of the Section where the Keys to count are."
    Case 8:
        'RenameSection() *** RenameSection FileName, "Section 2", "Section 999"
        lblReturn.Caption = RenameSection(FileName, "Section 2", "Section 999")

        txtDescription.Text = "This function renames a given Section." & vbCrLf & vbCrLf & "Function returns: 1 or 0" & vbCrLf & "1=Worked  0=Unsuccessful"
        txtParams.Text = "FileName: String" & vbCrLf & "The location of the INI File." & _
        vbCrLf & vbCrLf & "SectionName: String" & vbCrLf & "The name of the Section to rename." & _
        vbCrLf & vbCrLf & "NewSectionName: String" & vbCrLf & "The name you wish to change the Section to."
    Case 9:
        'RenameKey() *** RenameKey FileName, "Section 3", "Key 2", "Key 999"
        lblReturn.Caption = RenameKey(FileName, "Section 3", "Key 2", "Key 999")
        
        txtDescription.Text = "This function renames a given Key in a Section." & vbCrLf & vbCrLf & "Function returns: 1 or 0" & vbCrLf & "1=Worked  0=Unsuccessful"
        txtParams.Text = "FileName: String" & vbCrLf & "The location of the INI File." & _
        vbCrLf & vbCrLf & "Section: String" & vbCrLf & "The name of the Section where the Key to rename is." & _
        vbCrLf & vbCrLf & "KeyName: String" & vbCrLf & "The name of the Key you wish to rename." & _
        vbCrLf & vbCrLf & "NewKeyName: String" & vbCrLf & "The name you wish to change the Key to."
    Case 10:
        'GetKey() - Has to Return
        lblReturn.Caption = GetKey(FileName, "Section 4", 1)
        
        txtDescription.Text = "This function gets the name of a Key using the Key's IndexNumber.  IndexNumbers start at 0 and increment up." & vbCrLf & vbCrLf & "Function returns: String"
        txtParams.Text = "FileName: String" & vbCrLf & "The location of the INI File." & _
        vbCrLf & vbCrLf & "Section: String" & vbCrLf & "The name of the Section where the Key is." & _
        vbCrLf & vbCrLf & "KeyIndexNum: Integer" & vbCrLf & "The Key's IndexNumber."
    Case 11:
        'GetKey2() - Has to Return
        lblReturn.Caption = GetKey2(FileName, 3, 1)
        
        txtDescription.Text = "This function gets the name of a Key using the Key's IndexNumber and the Section's IndexNumber.  IndexNumbers start at 0 and increment up." & vbCrLf & vbCrLf & "Function returns: String"
        txtParams.Text = "FileName: String" & vbCrLf & "The location of the INI File." & _
        vbCrLf & vbCrLf & "SectionIndexNum: Integer" & vbCrLf & "The IndexNumber of the Section where the Key is." & _
        vbCrLf & vbCrLf & "KeyIndexNum: Integer" & vbCrLf & "The Key's IndexNumber."
    Case 12:
        'GetSection() - Has to Return
        lblReturn.Caption = GetSection(FileName, 3)
        
        txtDescription.Text = "This gets a Section's name using it's IndexNumber.  IndexNumbers start at 0 and increment up." & vbCrLf & vbCrLf & "Function returns: String"
        txtParams.Text = "FileName: String" & vbCrLf & "The location of the INI File." & _
        vbCrLf & vbCrLf & "SectionIndexNumber: Integer" & vbCrLf & "The Section's IndexNumber."
    Case 13:
        'IsKey() - Has to Return
        lblReturn.Caption = IsKey("Valid=key")
        
        txtDescription.Text = "This function designates wheather or not a line of text is a valid Key." & vbCrLf & vbCrLf & "Function returns: Boolean" & vbCrLf & "True or False"
        txtParams.Text = "TextLine: String" & vbCrLf & "The Line of Text which is to be tested if it is a Valid Key."
    Case 14:
        'IsSection() - Has to Return
        lblReturn.Caption = IsSection("[Valid Section]")
        
        txtDescription.Text = "This function designates whether or not a line of text is a valid Section." & vbCrLf & vbCrLf & "Function returns: Boolean" & vbCrLf & "True or False"
        txtParams.Text = "TextLine: String" & vbCrLf & "The Line of Text which is to be tested if it is a Valid Section."
    Case 15:
        'KeyExists() - Has to Return
        lblReturn.Caption = KeyExists(FileName, "Section 4", "Key 2")
        
        txtDescription.Text = "This function tests whether or not a Key exists in a given Section." & vbCrLf & vbCrLf & "Function returns: Boolean" & vbCrLf & "True or False"
        txtParams.Text = "FileName: String" & vbCrLf & "The location of the INI File." & _
        vbCrLf & vbCrLf & "Section: String" & vbCrLf & "The name of the Section where the Key is." & _
        vbCrLf & vbCrLf & "Key: String" & vbCrLf & "The name of the Key you wish to test for."
    Case 16:
        'KeyExists2() - Has to Return
        lblReturn.Caption = KeyExists2(FileName, 3, "Key 2")
        
        txtDescription.Text = "This function tests whether or not a Key allready exists in a given Section.  Section is identified by it's IndexNumber." & vbCrLf & vbCrLf & "Function returns: Boolean" & vbCrLf & "True or False"
        txtParams.Text = "FileName: String" & vbCrLf & "The location of the INI File." & _
        vbCrLf & vbCrLf & "SectionIndexNum: Integer" & vbCrLf & "The IndexNumber of the Section where the Key is." & _
        vbCrLf & vbCrLf & "Key: String" & vbCrLf & "The name of the Key you wish to test."
    Case 17:
        'SectionExists() - Has to Return
        lblReturn.Caption = SectionExists(FileName, "NonExistentSection")
        
        txtDescription.Text = "This function tests whether or not a Section allready exists in a given INI File." & vbCrLf & vbCrLf & "Function returns: Boolean" & vbCrLf & "True or False"
        txtParams.Text = "FileName: String" & vbCrLf & "The location of the INI File." & _
        vbCrLf & vbCrLf & "Section: String" & vbCrLf & "The name of the Section to test."
    Case 18:
        'GetSectionIndex() - Has to Return
        lblReturn.Caption = GetSectionIndex(FileName, "Section 2")
        
        txtDescription.Text = "Gets the IndexNumber of a Section." & vbCrLf & vbCrLf & "Function returns: Integer"
        txtParams.Text = "FileName: String" & vbCrLf & "The location of the INI File." & _
        vbCrLf & vbCrLf & "Section: String" & vbCrLf & "The name of the Section of which to get it's IndexNumber."
    Case 19:
        'GetKeyIndex() - Has to Return
        lblReturn.Caption = GetKeyIndex(FileName, "Section 3", "Key 4")
        
        txtDescription.Text = "Get's the IndexNumber of a Key in a given Section." & vbCrLf & vbCrLf & "Function returns: Integer"
        txtParams.Text = "FileName: String" & vbCrLf & "The location of the INI File." & _
        vbCrLf & vbCrLf & "Section: String" & vbCrLf & "The name of the Section where the Key is." & _
        vbCrLf & vbCrLf & "Key: String" & vbCrLf & "The name of the Key of which to get it's IndexNumber."
    Case 20:
        'GetKeyIndex2() - Has to Return
        lblReturn.Caption = GetKeyIndex2(FileName, 2, "Key 4")
        
        txtDescription.Text = "Get's the IndexNumber of a Key in a given Section.  Section is identified by it's IndexNumber." & vbCrLf & vbCrLf & "Function returns: Integer"
        txtParams.Text = "FileName: String" & vbCrLf & "The location of the INI File." & _
        vbCrLf & vbCrLf & "SectionIndexNum: Iteger" & vbCrLf & "The IndexNumber of the Section where the Key is." & _
        vbCrLf & vbCrLf & "Key: String" & vbCrLf & "The name of the Key of which to get it's IndexNumber."
End Select
End Sub
