VERSION 5.00
Begin VB.Form frm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "3TierXML"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7605
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   7605
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdModify 
      Caption         =   "Modify Name..."
      Height          =   435
      Left            =   180
      TabIndex        =   7
      Top             =   3900
      Width           =   1995
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   435
      Left            =   180
      TabIndex        =   6
      Top             =   3360
      Width           =   1995
   End
   Begin VB.CommandButton cmdSaveOrder 
      Caption         =   "Create Order"
      Height          =   435
      Left            =   5400
      TabIndex        =   5
      Top             =   2820
      Width           =   1995
   End
   Begin VB.CommandButton cmdSaveCustomer 
      Caption         =   "New..."
      Height          =   435
      Left            =   180
      TabIndex        =   4
      Top             =   2820
      Width           =   1995
   End
   Begin VB.Frame Frame2 
      Caption         =   "Customers"
      Height          =   2535
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   7215
      Begin VB.Frame Frame1 
         Caption         =   "Orders"
         Height          =   2175
         Left            =   3540
         TabIndex        =   2
         Top             =   180
         Width           =   3495
         Begin VB.ListBox lstOrd 
            Height          =   1620
            Left            =   120
            TabIndex        =   3
            Top             =   360
            Width           =   3195
         End
      End
      Begin VB.ListBox lstCust 
         Height          =   2010
         Left            =   120
         TabIndex        =   1
         Top             =   300
         Width           =   3195
      End
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   3620
      X2              =   3620
      Y1              =   2700
      Y2              =   4320
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000003&
      X1              =   3600
      X2              =   3600
      Y1              =   2700
      Y2              =   4320
   End
End
Attribute VB_Name = "frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public lErrNum As Long, sErrDesc As String, sErrSource As String

Public myController As New UIController

Public Sub FillCustomers()
    Dim i As Integer
    
    lstCust.Clear
    If myController.getCustomers(lErrNum, sErrDesc, sErrSource) Then
        For i = 1 To myController.mcolCustomers.Count
            lstCust.AddItem myController.mcolCustomers(i).CustomerName
            lstCust.ItemData(lstCust.NewIndex) = myController.mcolCustomers(i).IDCustomer
        Next i
        
        lstCust.ListIndex = 0
    Else
        MsgBox lErrNum & ": " & sErrDesc, vbCritical, sErrSource
        lErrNum = 0: sErrDesc = "": sErrSource = ""
    End If
End Sub

Private Sub cmdDelete_Click()
    Dim customerInfo As CustomerUDT
    
    customerInfo = myController.mcolCustomers("CU-" & lstCust.ItemData(lstCust.ListIndex))
    
    customerInfo.IsDeleted = True
    
    myController.mcolCustomers.Remove "CU-" & customerInfo.IDCustomer
    myController.mcolCustomers.Add customerInfo, "CU-" & customerInfo.IDCustomer
    
    If myController.deleteCustomers(lErrNum, sErrDesc, sErrSource) Then
        Call FillCustomers
        MsgBox "The Customer was deleted"
    Else
        MsgBox "Error"
    End If
End Sub

Private Sub cmdSaveCustomer_Click()
    Dim strCustomerName As String
    Dim customerInfo As CustomerUDT
    
    strCustomerName = InputBox("Give me the Customer Name")
    
    Randomize
    With customerInfo
        .IDCustomer = Int(Rnd * 100000)
        .CustomerName = strCustomerName
        .IsNew = True
    End With
    
    myController.mcolCustomers.Add customerInfo, "CU-" & customerInfo.IDCustomer
    If myController.saveCustomers(lErrNum, sErrDesc, sErrSource) Then
        Call FillCustomers
        MsgBox "The Customer was created"
    Else
        MsgBox "Error"
    End If
End Sub

Private Sub cmdSaveOrder_Click()
    Dim orderInfo As OrderUDT
    Dim lngPresentCustomer As Long
    
    lngPresentCustomer = lstCust.ItemData(lstCust.ListIndex)
    
    Randomize
    With orderInfo
        .IDCustomer = lngPresentCustomer
        .DateOrder = Now
        .IDOrder = Int(Rnd * 100000)
        .IsNew = True
    End With
    
    myController.mcolOrders.Add orderInfo, "O-" & orderInfo.IDOrder
    
    If myController.saveOrders(lErrNum, sErrDesc, sErrSource) Then
        Call FillCustomers
        MsgBox "The Order was created"
    Else
        MsgBox "Error"
    End If
End Sub

Private Sub Form_Load()
    Call FillCustomers
End Sub

Private Sub lstCust_Click()
    Dim i As Integer
    
    lstOrd.Clear
    If myController.getCustomer(lstCust.ItemData(lstCust.ListIndex), lErrNum, sErrDesc, sErrSource) Then
        'Fill the Orders ListBox
        For i = 1 To myController.mcolOrders.Count
            lstOrd.AddItem myController.mcolOrders(i).IDOrder & ": " & myController.mcolOrders(i).DateOrder
            lstOrd.ItemData(lstOrd.NewIndex) = myController.mcolOrders(i).IDOrder
        Next i
        
        If myController.mcolOrders.Count > 0 Then
            lstOrd.ListIndex = 0
        End If
    End If
End Sub

Private Sub cmdModify_Click()
    Dim strCustomerName As String
    Dim customerInfo As CustomerUDT
    
    customerInfo = myController.mcolCustomers("CU-" & lstCust.ItemData(lstCust.ListIndex))
    
    strCustomerName = InputBox("Give me the new Customer Name", , customerInfo.CustomerName)
    
    customerInfo.CustomerName = strCustomerName
    customerInfo.IsDirty = True
    'Remove old version of the Customer
    myController.mcolCustomers.Remove "CU-" & customerInfo.IDCustomer
    'Add the new version
    myController.mcolCustomers.Add customerInfo, "CU-" & customerInfo.IDCustomer
    
    If myController.modifyCustomers(lErrNum, sErrDesc, sErrSource) Then
        Call FillCustomers
        MsgBox "The Customer was modified"
    Else
        MsgBox "Error"
    End If
End Sub

