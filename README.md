# Event_Management
1.LOGIN FORM
Dim rs As New ADODB.Recordset
Private Sub Command1_Click ()
Dim username As String
Dim password As String
username = &quot;admin&quot;
password = &quot;admin12&quot;
If username = Text1.Text And password = Text2.Text Then
MsgBox (&quot;successfully logged in&quot;)

Text1.SetFocus
Text2.SetFocus
Form1.Hide
Form2.Show
Else
MsgBox (&quot;invalid username or password&quot;)
Text1.SetFocus
Text2.SetFocus
End If
End Sub

Private Sub Command2_Click ()
End
End Sub

Private Sub Form_Load ()
With rs
rs. ActiveConnection = &quot;Provider=Microsoft. Jet. OLEDB.4.0; Data Source=D:\agape\event
booking.mdb; Persist Security Info=False&quot;
rs.Source = &quot;select *from adminid&quot;
rs.CursorType = adLockOptimistic
rs.LockType = adLockOptimistic
End With

Text1.Text = &quot;&quot;
Text2.Text = &quot;&quot;
End Sub
2.HOME PAGE
Private Sub bookhis_Click ()
Form2.Hide
Form9.Show
End Sub

Private Sub booking Click ()
Form2.Hide
Form7.Show
End Sub

Private Sub cancel_Click ()
Form2.Hide
Form10.Show
End Sub

Private Sub logout_Click ()
Form2.Hide
Form1.Show

End Sub
3.CUSTOMER DETAILS
Option Explicit
Dim conn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Private Sub Command1_Click()
Form8.Text7.Text = Text6.Text
Form8.Text8.Text = Text1.Text
Adodc1.Recordset.MoveLast
Adodc1.Recordset.AddNew
&#39;Exit Sub
rs.Update
MsgBox &quot;successfully added&quot;
Adodc1.Recordset.AddNew
rs.AddNew
rs!cusid = Val(Text6.Text)
rs!cusname = Text1.Text
rs!phoneno = Text2.Text
rs!adress = Text3.Text
rs!email = Text4.Text
rs!pincode = Text5.Text
rs.Update

&#39;MsgBox &quot;successfully added&quot;
Form7.Hide
Form8.Show
End Sub

Private Sub Command3_Click()
Form7.Hide
Form2.Show
End Sub

Private Sub Form_Load ()
&#39;Dim rs As ADODB.Recordset
&#39;Dim conn As New ADODB.Connection
conn.Open &quot;Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\agape\event
booking.mdb;Persist Security Info=False&quot;
Text1.Text = &quot;&quot;
Text2.Text = &quot;&quot;
Text3.Text = &quot;&quot;
Text4.Text = &quot;&quot;
Text5.Text = &quot;&quot;
Dim newid As Integer
newid = Val(Adodc1.Recordset.Fields(0)) + 1
Text6.Text = newid

End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
If Not (KeyAscii &gt;= 65 And KeyAscii &lt;= 90 Or KeyAscii = 8) And Not (KeyAscii &gt;= 97 And
KeyAscii &lt;= 122) And Not (KeyAscii = 32 Or KeyAscii = 45) Then
KeyAscii = 0
MsgBox &quot;please enter a valid name&quot;, vbExclamation
End If
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
Dim phonenum As String
If Not (KeyAscii &gt;= 48 And KeyAscii &lt;= 57 Or KeyAscii = 8) Or Len(Text2.Text) = 10 Then
KeyAscii = 0
If Len(phonenum) &lt;&gt; 10 Or Not IsNumeric(phonenum) Then
MsgBox &quot;please enter a valid phone no&quot;, vbExclamation
End If
End If
End Sub

Private Sub Text2_LostFocus()
Dim phonenum As String
phonenum = Trim(Text2.Text)
If Len(phonenum) &lt;&gt; 10 Or Not IsNumeric(phonenum) Then
MsgBox &quot;please enter a valid 10 digit contact number&quot;, vbExclamation

Text2.SetFocus
End If
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If Not (KeyAscii &gt;= 48 And KeyAscii &lt;= 57 Or KeyAscii = 8) Then
KeyAscii = 0
MsgBox &quot;please enter a valid pincode&quot;
End If
End Sub

Private Sub Text5_LostFocus()
Dim pin As String
pin = Trim(Text5.Text)
If Len(pin) &lt;&gt; 6 Or Not IsNumeric(pin) Then
MsgBox &quot;please enter a valid 6 digit numeric pincode&quot;
Text5.SetFocus
End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
Dim str As String
str = Text4.Text

If str = &quot;[A-Z a-z 0-9 ]*@[A-Z a-z 0-9]*. [a-z A-Z]&quot; Then
End If
End Sub
4.BOOKING FORM
Option Explicit
Dim conn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Private mcusid As Integer
Private isfrozen As Boolean

Private Sub Combo1_Click()
Combo4.Clear
Select Case Combo1.ListIndex
Case 1
Combo4.AddItem &quot;seetha hall&quot;
Combo4.AddItem &quot;kp minihall&quot;
Combo4.AddItem &quot;thamarai madapam&quot;
Case 0
Combo4.AddItem &quot;raj mahal&quot;
Combo4.AddItem &quot;sv mahal&quot;
Combo4.AddItem &quot;rose bowl&quot;
Combo4.AddItem &quot;dia partyhall&quot;

Combo4.AddItem &quot;grand palace&quot;
End Select
End Sub
Private Sub Combo2_click()
Combo5.Clear
Select Case Combo2.ListIndex
Case 0
Combo5.AddItem &quot;silver vegplatter&quot;
Combo5.AddItem &quot;golden vegplatter&quot;
Combo5.AddItem &quot;diamond vegplatter&quot;
Case 1
Combo5.AddItem &quot;biryaani&quot;
Combo5.AddItem &quot;golden nonvegplatter&quot;
Combo5.AddItem &quot;diamond nonvegplatter&quot;
End Select
End Sub
Private Sub Command1_Click()
Dim evename As String
Dim evedate As Date
Dim venue As String
Dim people As Integer
Dim address As String
Dim days As Integer

Dim timing As String
Dim food As String
Dim plates As Integer
Dim decor As String
Dim total As Currency
Dim venuecost(8) As Currency
Dim foodcost(6) As Currency
Dim decorcost(5) As Currency
Dim selectedvenueindex As Integer
Dim selectedfoodindex As Integer
Dim selecteddecorindex As Integer
evename = Combo3.Text
people = Val(Text6.Text)
evedate = DTPicker1.Value
venue = Combo4.Text
address = Text2.Text
days = Val(Text1.Text)
timing = DTPicker2.Value
&#39;timing = Text5.Text
&#39;Text5.Text = Format(Time, &quot;hh:mm:ss AMPM&quot;)
food = Combo5.Text
plates = Val(Text3.Text)
decor = Combo6.Text

If Not validatedate(DTPicker1.Value) Then
MsgBox &quot;please select a valid date&quot;
End If
If Trim(Combo3.Text) = &quot;&quot; Then
MsgBox &quot;please enter the event name&quot;
End If
venuecost(0) = 70000
venuecost(1) = 50000
venuecost(2) = 65000
venuecost(3) = 45000
venuecost(4) = 50000
venuecost(5) = 40000
venuecost(6) = 40000
venuecost(7) = 80000
foodcost(0) = 250
foodcost(1) = 250
foodcost(2) = 150
foodcost(3) = 400
foodcost(4) = 300
foodcost(5) = 500
decorcost(0) = 4000
decorcost(1) = 1500
decorcost(2) = 6000

decorcost(3) = 7000
decorcost(4) = 8000
selectedvenueindex = Combo4.ListIndex
selectedfoodindex = Combo5.ListIndex
selecteddecorindex = Combo6.ListIndex
Label14.Caption = venuecost(selectedvenueindex) * days
Label15.Caption = foodcost(selectedfoodindex) * plates
Label16.Caption = decorcost(selecteddecorindex) * days
total = Val(Label14.Caption) + Val(Label15.Caption) + Val(Label16.Caption)
Text4.Text = total
Label2.Enabled = False
Label3.Enabled = False
Label4.Enabled = False
Label5.Enabled = False
Label6.Enabled = False
Label7.Enabled = False
Label8.Enabled = False
Label9.Enabled = False
Label10.Enabled = False
Label11.Enabled = False
Label12.Enabled = False
Label13.Enabled = False
Label18.Enabled = False

Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
DTPicker2.Enabled = False
Text6.Enabled = False
Combo1.Enabled = False
Combo2.Enabled = False
Combo3.Enabled = False
Combo4.Enabled = False
Combo5.Enabled = False
Combo6.Enabled = False
DTPicker1.Enabled = False
Command3.Enabled = False
End Sub

Private Sub Command1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As
Single)
If Button = 2 Then
unfreezecontrols
End If
End Sub

Private Sub Command2_Click()

Form11.Text1.Text = Text4.Text
Form11.Text4.Text = Text7.Text
Form11.Text5.Text = Text8.Text
Adodc1.Recordset.AddNew
rs.AddNew
rs!cusid = Val(Text7.Text)
rs!evename = Combo3.Text
rs!bookingDate = DTPicker1.Value
rs!people = Val(Text6.Text)
rs!halltype = Combo1.Text
rs!venue = Combo4.Text
rs!address = Text2.Text
rs!days = Val(Text1.Text)
rs!timing = DTPicker2.Value
rs!foodtype = Combo2.Text
rs!food = Combo5.Text
rs!plates = Val(Text3.Text)
rs!decor = Combo6.Text
rs!total = Val(Text4.Text)
rs.Update
MsgBox (&quot;your order is added, thankyou!!!&quot;)
Dim total As Double
total = Val(Text4.Text)

Form8.Hide
Form11.Show
End Sub
Private Sub Command3_Click()
Dim rs As ADODB.Recordset
Dim conn As ADODB.Connection
Dim cmd As ADODB.Command
Dim strsql As String
Dim selecteddate As Date
selecteddate = DTPicker1.Value
Set conn = New ADODB.Connection
conn.ConnectionString = &quot;Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\agape\event
booking.mdb;Persist Security Info=False&quot;
conn.Open
If conn.State &lt;&gt; adStateOpen Then
MsgBox &quot;error connecting to the database&quot;, vbExclamation
Exit Sub
End If
strsql = &quot;select * from booking where venue= ? and bookingdate = ?;&quot;
Set cmd = New ADODB.Command
With cmd
.ActiveConnection = conn
.CommandType = adCmdText

.CommandText = strsql
.Parameters.Append .CreateParameter(&quot;venue&quot;, adVarChar, adParamInput, 255, Combo4.Text)
.Parameters.Append .CreateParameter(&quot;bookingdate&quot;, adDate, adParamInput, , selecteddate)
End With
Set rs = cmd.Execute
If Not rs.EOF Then
MsgBox &quot;venue is not available on selected date.&quot;, vbExclamation
Else
MsgBox &quot;venue is available on selected date.&quot;, vbInformation
End If
rs. Close
conn. Close
Set rs = Nothing
Set conn = Nothing
End Sub

Private Sub Command4_Click()
DataReport1.Show
End Sub

Private Sub DTPicker1_CloseUp ()
Dim selecteddate As Date
selecteddate = DTPicker1.Value

If selecteddate &lt; Date Then
MsgBox &quot;please select a future date&quot;, vbExclamation
DTPicker1.Value = Date
End If
End Sub

Private Sub Form_Click()
Label2.Enabled = True
Label3.Enabled = True
Label4.Enabled = True
Label5.Enabled = True
Label6.Enabled = True
Label7.Enabled = True
Label8.Enabled = True
Label9.Enabled = True
Label10.Enabled = True
Label11.Enabled = True
Label12.Enabled = True
Label13.Enabled = True
Label18.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True

DTPicker2.Enabled = True
Text6.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Combo4.Enabled = True
Combo5.Enabled = True
Combo6.Enabled = True
DTPicker1.Enabled = True
Command3.Enabled = True
End Sub

Private Sub Form_Load()
Dim venuename As String
Dim address As String
Dim price As Double
Dim cusid As Integer
Combo3.AddItem &quot;marriage&quot;
Combo3.AddItem &quot;birthday&quot;
Combo3.AddItem &quot;engagement&quot;
Combo3.AddItem &quot;reception&quot;
Combo3.AddItem &quot;babyshower&quot;
Combo4.AddItem &quot;rajmahal&quot;

Combo4.AddItem &quot;seetha hall&quot;
Combo4.AddItem &quot;sv mahal&quot;
Combo4.AddItem &quot;rosebowl&quot;
Combo4.AddItem &quot;kp minihall&quot;
Combo4.AddItem &quot;dia party hall&quot;
Combo4.AddItem &quot;grand palace&quot;
Combo4.AddItem &quot;thamarai mandapam&quot;
Combo5.AddItem &quot;biryaani&quot;
Combo5.AddItem &quot;golden vegplatter&quot;
Combo5.AddItem &quot;silver vegplatter&quot;
Combo5.AddItem &quot;diamond vegplatter&quot;
Combo5.AddItem &quot;golden nonvegplatter&quot;
Combo5.AddItem &quot;diamond nonvegplatter&quot;
Combo6.AddItem &quot; floral decoration&quot;
Combo6.AddItem &quot;lighting decoration&quot;
Combo6.AddItem &quot;balloon decoration&quot;
Combo6.AddItem &quot;tropical decoration&quot;
Combo6.AddItem &quot; fairy decoration&quot;
Set conn = New ADODB.Connection
conn.ConnectionString = &quot;Provider=Microsoft.Jet. OLEDB.4.0;Data Source=D:\agape\event
booking.mdb;Persist Security Info=False&quot;
conn.Open
Text1.Text = &quot;&quot;

Text2.Text = &quot;&quot;
Text3.Text = &quot;&quot;
Text4.Text = &quot;&quot;
Text6.Text = &quot;&quot;
Text7.Text = cusid
Set rs = New ADODB.Recordset
rs.Open &quot;booking&quot;, conn, adOpenKeyset, adLockOptimistic
End Sub
Private Function validatedate(ByVal selecteddate As Date) As Boolean
If Year(selecteddate) = Year (Date) Then
validatedate = True
Else
validatedate = False
End If
End Function
Private Sub combo4_click()
Dim selectedvenueindex As Integer
Dim selectedvenueaddress As String
selectedvenueindex = Combo4.ListIndex
selectedvenueaddress = getvenueaddress(selectedvenueindex)
Text2.Text = selectedvenueaddress
End Sub

Private Function getvenueaddress(ByVal venueindex As Integer) As String
Select Case venueindex
Case 0
getvenueaddress = &quot; no 42, Ramamoorthy Ave main road, Sakthi Nagar, porur,chennai-116&quot;
Case 1
getvenueaddress = &quot;SBI colony 2nd st, ranga colony, sembakkam, rajkilpakkam, chennai-73&quot;
Case 2
getvenueaddress = &quot; no 94, kundrathur mainroad, SH 113, Vigneswaran Nagar, porur,chennai-
116&quot;
Case 3
getvenueaddress = &quot;no 172, arcot road, vadapalani,chennai-26&quot;
Case 4
getvenueaddress = &quot;no 1/22b, solai amman Nagar, Gandhi nagar, redhills,chennai-52&quot;
Case 5
getvenueaddress = &quot;sambhu, LV prasad road, vadapalani, chennai-26&quot;
Case 6
getvenueaddress = &quot;no 32, mount poonamalle road, manapakkam, chennai-125&quot;
Case 7
getvenueaddress = &quot;no 40, mogappair west, Ambattur industrial estate,chennai-37&quot;
End Select
End Function

Private Sub Text1_KeyPress (KeyAscii As Integer)

If Not (KeyAscii &gt;= 48 And KeyAscii &lt;= 57 Or KeyAscii = 8) Then
KeyAscii = 0
End Sub
Private Sub Text3_KeyPress (KeyAscii As Integer)
If Not (KeyAscii &gt;= 48 And KeyAscii &lt;= 57 Or KeyAscii = 8) Then
KeyAscii = 0
End Sub
Private Sub Text6_KeyPress (KeyAscii As Integer)
If Not (KeyAscii &gt;= 48 And KeyAscii &lt;= 57 Or KeyAscii = 8) Then
KeyAscii = 0
End Sub
ADVANCE PAYMENT
Private Sub Command1_Click()
Dim rs As New ADODB.Recordset
Dim conn As New ADODB.Connection
conn.Open &quot;Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\agape\event
booking.mdb;Persist Security Info=False&quot;
Adodc1.Recordset.AddNew
&#39;rs.AddNew
&#39;rs!total = Val(Text1.Text)
&#39;rs!advance = Val(Text2.Text)
&#39;rs!balance = Val(Text3.Text)
MsgBox &quot;successfully advance is paid&quot;

Form11.Hide
Form2.Show
End Sub

Private Sub Command2_Click()
Dim total As Double
Dim advance As Double
Dim balance As Double
total = Val(Text1.Text)
advance = total * 0.3
Text2.Text = advance
balance = total - advance
Text3.Text = Val(balance)
End Sub

Private Sub Form_Load()
Dim total As Double
Dim advance As Double
Dim balance As Double
Text1.Text = Format(total, &quot;0.00&quot;)
Text2.Text = Format(advance, &quot;0.00&quot;)
Text3.Text = Format(balance, &quot;0.00&quot;)
End Sub
