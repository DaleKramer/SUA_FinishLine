VERSION 5.00
Begin VB.Form FinishLine 
   Caption         =   "Finish Line by Dale Kramer v1.02"
   ClientHeight    =   7296
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   4524
   LinkTopic       =   "Form1"
   ScaleHeight     =   7296
   ScaleWidth      =   4524
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Kpoint 
      Caption         =   "Center Known"
      Height          =   300
      Index           =   1
      Left            =   2544
      TabIndex        =   27
      Top             =   48
      Width           =   1548
   End
   Begin VB.OptionButton Kpoint 
      Caption         =   "Endpoint Known"
      Height          =   300
      Index           =   0
      Left            =   528
      TabIndex        =   26
      Top             =   48
      Width           =   1548
   End
   Begin VB.CommandButton Export 
      Caption         =   "Export Airspace File"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   444
      Left            =   768
      TabIndex        =   25
      Top             =   6816
      Width           =   2844
   End
   Begin VB.TextBox NSlon2 
      Alignment       =   2  'Center
      Height          =   300
      Left            =   3264
      Locked          =   -1  'True
      TabIndex        =   24
      Top             =   5328
      Width           =   972
   End
   Begin VB.TextBox NSlat2 
      Alignment       =   2  'Center
      Height          =   300
      Left            =   3264
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   4896
      Width           =   972
   End
   Begin VB.TextBox NSlon1 
      Alignment       =   2  'Center
      Height          =   300
      Left            =   3264
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   3312
      Width           =   972
   End
   Begin VB.TextBox NSlat1 
      Alignment       =   2  'Center
      Height          =   300
      Left            =   3264
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   2880
      Width           =   972
   End
   Begin VB.CommandButton Calculate 
      Caption         =   "Calculate"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   444
      Left            =   1056
      TabIndex        =   16
      Top             =   2304
      Width           =   2316
   End
   Begin VB.TextBox Elon2 
      Alignment       =   2  'Center
      Height          =   300
      Left            =   3264
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   6336
      Width           =   972
   End
   Begin VB.TextBox Elat2 
      Alignment       =   2  'Center
      Height          =   300
      Left            =   3264
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   5904
      Width           =   972
   End
   Begin VB.TextBox CenterLon 
      Alignment       =   2  'Center
      Height          =   300
      Left            =   3264
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   4320
      Width           =   972
   End
   Begin VB.TextBox CenterLat 
      Alignment       =   2  'Center
      Height          =   300
      Left            =   3264
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   3888
      Width           =   972
   End
   Begin VB.TextBox Lwidth 
      Alignment       =   2  'Center
      Height          =   300
      Left            =   3264
      TabIndex        =   6
      Text            =   "3300"
      Top             =   1776
      Width           =   972
   End
   Begin VB.TextBox FTdir 
      Alignment       =   2  'Center
      Height          =   300
      Left            =   3264
      TabIndex        =   4
      Text            =   "180"
      Top             =   1344
      Width           =   972
   End
   Begin VB.TextBox Elon1 
      Alignment       =   2  'Center
      Height          =   300
      Left            =   3264
      TabIndex        =   2
      Text            =   "083:46.510W"
      Top             =   864
      Width           =   972
   End
   Begin VB.TextBox Elat1 
      Alignment       =   2  'Center
      Height          =   300
      Left            =   3264
      TabIndex        =   0
      Text            =   "31:59.400N"
      Top             =   384
      Width           =   972
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "Nearest 'Second' Endpoint 2 Longitude"
      Height          =   300
      Left            =   288
      TabIndex        =   22
      Top             =   5376
      Width           =   2940
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "Nearest 'Second' Endpoint 2 Latitude"
      Height          =   300
      Left            =   528
      TabIndex        =   21
      Top             =   4944
      Width           =   2700
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Caption         =   "Nearest 'Second' Endpoint 1 Longitude"
      Height          =   300
      Left            =   288
      TabIndex        =   18
      Top             =   3360
      Width           =   2940
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "Nearest 'Second' Endpoint 1 Latitude"
      Height          =   300
      Left            =   528
      TabIndex        =   17
      Top             =   2928
      Width           =   2700
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "Endpoint 2 Longitude"
      Height          =   300
      Left            =   528
      TabIndex        =   11
      Top             =   6384
      Width           =   2700
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "Endpoint 2 Latitude"
      Height          =   300
      Left            =   528
      TabIndex        =   10
      Top             =   5952
      Width           =   2700
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Center Longitude (ddd:mm.sssW)"
      Height          =   300
      Left            =   528
      TabIndex        =   9
      Top             =   4368
      Width           =   2700
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Center Latitude (dd:mm.sssN)"
      Height          =   300
      Left            =   528
      TabIndex        =   8
      Top             =   3936
      Width           =   2700
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Line Length (feet)"
      Height          =   300
      Left            =   528
      TabIndex        =   7
      Top             =   1824
      Width           =   2700
   End
   Begin VB.Label Headlabel 
      Alignment       =   1  'Right Justify
      Caption         =   "Heading to other end of line (degrees True)"
      Height          =   300
      Left            =   96
      TabIndex        =   5
      Top             =   1392
      Width           =   3132
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Endpoint 1 of line Longitude (ddd:mm.sssW)"
      Height          =   300
      Left            =   48
      TabIndex        =   3
      Top             =   912
      Width           =   3132
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Endpoint 1 of line Latitude (dd:mm.sssN)"
      Height          =   300
      Left            =   336
      TabIndex        =   1
      Top             =   432
      Width           =   2892
   End
End
Attribute VB_Name = "FinishLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim pi  As Double
Dim RotCenLat As Double
Dim RotCenLon As Double
Dim NewLat As Double
Dim NewLon As Double
Dim OldLat As Double
Dim OldLon As Double
Dim dd, mm As Integer
Dim ss As Single

Private Sub Calculate_Click()
   Dim la1 As Double
   Dim lo1 As Double
   Dim la2 As Double
   Dim lo2 As Double
   Dim cLat As Double
   Dim cLon As Double
   Dim lat1 As Double
   Dim lon1 As Double
   Dim lat2 As Double
   Dim lon2 As Double
   Dim latang As Double
   Dim rotang As Double
   Dim mind As Double
   Dim miny, minz As Integer
   Dim y, z As Integer
   Dim slat1(2) As Double
   Dim slon1(2) As Double
   Dim slat2(2) As Double
   Dim slon2(2) As Double
   If Kpoint(0).Value = True Then
      la1 = datlattorad(Elat1)
      lo1 = datlontorad(Elon1)
      latang = (((Val(Lwidth.Text) / 2) / 3280.8) / 111.195) * pi / 180
      lo2 = lo1 + latang / Cos(la1)
      la2 = la1
      Debug.Print Round(distance(la1 + 0, lo1 + 0, la2 + 0, lo2 + 0) * 3280.8, 1);
      Debug.Print Round(heading(la1 + 0, lo1 + 0, la2 + 0, lo2 + 0), 0)
      
      lat1 = la1
      lon1 = lo1
      RotCenLat = la1
      RotCenLon = lo1
      OldLat = la2
      OldLon = lo2
      rotang = Val(FTdir.Text) - 90
      
      Call RotateLine(rotang + 0)
         
      
      cLat = datlattorad(radlattodec(NewLat))
      cLon = datlattorad(radlattodec(NewLon))
      
      Debug.Print Round(distance(la1 + 0, lo1 + 0, cLat + 0, cLon + 0) * 3280.8, 1);
      Debug.Print Round(heading(la1 + 0, lo1 + 0, cLat + 0, cLon + 0), 0)
      
      CenterLat.Text = radlattodec(cLat + 0)
      CenterLon.Text = radlontodec(cLon + 0)
      
      latang = (((Val(Lwidth.Text)) / 3280.8) / 111.195) * pi / 180
      lo2 = lo1 + latang / Cos(la1)
      la2 = la1
      Debug.Print Round(distance(la1 + 0, lo1 + 0, la2 + 0, lo2 + 0) * 3280.8, 1);
      Debug.Print Round(heading(la1 + 0, lo1 + 0, la2 + 0, lo2 + 0), 0)
      
      RotCenLat = la1
      RotCenLon = lo1
      OldLat = la2
      OldLon = lo2
      
      Call RotateLine(rotang + 0)
         
      
      Debug.Print Round(distance(la1 + 0, lo1 + 0, NewLat + 0, NewLon + 0) * 3280.8, 1);
      Debug.Print Round(heading(la1 + 0, lo1 + 0, NewLat + 0, NewLon + 0), 0)
         
      lat2 = NewLat
      lon2 = NewLon
      Elat2.Text = radlattodec(NewLat + 0)
      Elon2.Text = radlontodec(NewLon + 0)
   Else
      la1 = datlattorad(CenterLat.Text)
      lo1 = datlontorad(CenterLon.Text)
      latang = ((((Val(Lwidth.Text) / 2) / 3280.8) / 111.195) * pi) / 180
      lo2 = lo1 + latang / Cos(la1)
      la2 = la1
      Debug.Print Round(distance(la1 + 0, lo1 + 0, la2 + 0, lo2 + 0) * 3280.8, 1);
      Debug.Print Round(heading(la1 + 0, lo1 + 0, la2 + 0, lo2 + 0), 0)
      
      lat1 = la1
      lon1 = lo1
      RotCenLat = la1
      RotCenLon = lo1
      OldLat = la2
      OldLon = lo2
      rotang = Val(FTdir.Text) - 270
      If rotang = 0 Then rotang = 180
      
      Call RotateLine(rotang + 0)
         
      
      Debug.Print Round(distance(la1 + 0, lo1 + 0, NewLat + 0, NewLon + 0) * 3280.8, 1);
      Debug.Print Round(heading(la1 + 0, lo1 + 0, NewLat + 0, NewLon + 0), 0)
      
      lat1 = NewLat
      lon1 = NewLon
      Elat1.Text = radlattodec(NewLat + 0)
      Elon1.Text = radlontodec(NewLon + 0)
      
      lo2 = lo1 + latang / Cos(la1)
      la2 = la1
      Debug.Print Round(distance(la1 + 0, lo1 + 0, la2 + 0, lo2 + 0) * 3280.8, 1);
      Debug.Print Round(heading(la1 + 0, lo1 + 0, la2 + 0, lo2 + 0), 0)
      
      RotCenLat = la1
      RotCenLon = lo1
      OldLat = la2
      OldLon = lo2
      
      Call RotateLine(rotang + 180)
      If rotang > 360 Then
         rotang = rotang - 360
      End If
         
      
      Debug.Print Round(distance(la1 + 0, lo1 + 0, NewLat + 0, NewLon + 0) * 3280.8, 1);
      Debug.Print Round(heading(la1 + 0, lo1 + 0, NewLat + 0, NewLon + 0), 0)
         
      lat2 = NewLat
      lon2 = NewLon
      Elat2.Text = radlattodec(NewLat + 0)
      Elon2.Text = radlontodec(NewLon + 0)
      
   End If
   
   Call SetDMS(lat1)
   If Int(ss) = ss Then
      slat1(1) = lat1
      slat1(2) = lat1
   Else
      slat1(1) = (dd + (mm / 60) + Int(ss) / 3600) * pi / 180
      slat1(2) = (dd + (mm / 60) + (Int(ss) + 1) / 3600) * pi / 180
   End If
   Call SetDMS(lon1)
   If Int(ss) = ss Then
      slon1(1) = lon1
      slon1(2) = lon1
   Else
      slon1(1) = (dd + (mm / 60) + Int(ss) / 3600) * pi / 180
      slon1(2) = (dd + (mm / 60) + (Int(ss) + 1) / 3600) * pi / 180
   End If
   mind = 10000000
   For y = 1 To 2
      For z = 1 To 2
         If distance(slat1(y) + 0, slon1(z) + 0, cLat + 0, cLon + 0) < mind Then
            mind = distance(slat1(y) + 0, slon1(z) + 0, cLat + 0, cLon + 0)
            miny = y
            minz = z
         End If
      Next
   Next
   
   NSlat1.Text = radlattosec(slat1(miny) + 0)
   NSlon1.Text = radlontosec(slon1(minz) + 0)
   
   
   Call SetDMS(lat2)
   If Int(ss) = ss Then
      slat2(1) = lat2
      slat2(2) = lat2
   Else
      slat2(1) = (dd + (mm / 60) + Int(ss) / 3600) * pi / 180
      slat2(2) = (dd + (mm / 60) + (Int(ss) + 1) / 3600) * pi / 180
   End If
   Call SetDMS(lon2)
   If Int(ss) = ss Then
      slon2(1) = lon2
      slon2(2) = lon2
   Else
      slon2(1) = (dd + (mm / 60) + Int(ss) / 3600) * pi / 180
      slon2(2) = (dd + (mm / 60) + (Int(ss) + 1) / 3600) * pi / 180
   End If
   
   mind = 10000000
   For y = 1 To 2
      For z = 1 To 2
         If distance(slat2(y) + 0, slon2(z) + 0, cLat + 0, cLon + 0) < mind Then
            mind = distance(slat2(y) + 0, slon2(z) + 0, cLat + 0, cLon + 0)
            miny = y
            minz = z
         End If
      Next
   Next
   
   NSlat2.Text = radlattosec(slat2(miny) + 0)
   NSlon2.Text = radlontosec(slon2(minz) + 0)
   
End Sub
Sub SetDMS(inrad)
   Dim indec
   indec = inrad * 180 / pi
   dd = Int(indec)
   mm = Int((indec - dd) * 60)
   ss = (((indec - dd) * 60) - mm) * 60
End Sub


Function datlattorad(inlat As String)
    If inlat = "" Then
        datlattorad = -1000
        Exit Function
    End If
    Dim sg As Integer
    If Right(inlat, 1) = "S" Then sg = -1 Else sg = 1
    Dim d, s, sp1, sp2 As Integer
    Dim m As Double
    d = Val(Left(inlat, 2))
    m = Val(Mid(inlat, 4, 6))
    datlattorad = sg * (d + m / 60) * pi / 180

End Function
Function datlontorad(inlat As String)
    If inlat = "" Then
        datlontorad = -1000
        Exit Function
    End If
    Dim sg As Integer
    If Right(inlat, 1) = "W" Then sg = -1 Else sg = 1
    Dim d, s, sp1, sp2 As Integer
    Dim m As Double
    d = Val(Left(inlat, 3))
    m = Val(Mid(inlat, 5, 6))
    datlontorad = sg * (d + m / 60) * pi / 180

End Function
Function distance(la1 As Single, lo1 As Single, la2 As Single, lo2 As Single)
   Dim cosa As Double
   Dim tana As Double
    cosa = (Cos(la1)) * (Cos(la2)) * (Cos(lo2 - lo1)) + (Sin(la1)) * (Sin(la2))
    If cosa > 0.99999999999999 Then
        distance = 0
    Else
        tana = (1 - cosa ^ 2) ^ 0.5 / cosa
        distance = (Atn(tana) * 111.195 * 180) / pi 'km
    End If
End Function
Function heading(la1, lo1, la2, lo2)
    Dim dla  As Double
    Dim dlo As Double
    Dim ang As Double
    Dim c As Double
    dla = la2 - la1
    dlo = lo2 - lo1
    c = ARCCOS(Sin(la2) * Sin(la1) + Cos(la2) * Cos(la1) * Cos(lo2 - lo1))
    If dlo <> 0 Then
        ang = Abs(90 - ARCCOS((Sin(la2) - (Sin(la1) * Cos(c))) / (Cos(la1) * Sin(c))) * 180 / pi)
        'ang = (Atn(Abs((dla / (Cos(Abs(la1 + la2) / 2) * dlo))))) * 180 / pi
    End If
    If dla <= 0 Then '90 to 270
        If dlo >= 0 Then
            '90 to 180 q2
            If dlo = 0 Then
                heading = 180
            Else
                heading = 90 + ang
                'heading = 180 + ang
            End If
        Else
            '180+ to 270 q3
            heading = 270 - ang
        End If
    Else  '270 to 90
        If dlo >= 0 Then
            '0 to 90 q1
            If dlo = 0 Then
                heading = 0
            Else
                heading = 90 - ang
            End If
        Else
            '270 to -0 q4
            heading = 270 + ang
        End If
    End If
    'heading = heading + 90
End Function
Sub RotateLine(rota As Single)
    Dim da As Double
    Dim dbb As Double
    Dim a As Double
    Dim b As Double
    Dim c As Double
    Dim psi As Double
    Dim alpha As Double
    Dim lambda(1) As Double
    Dim phi_a(1) As Double
    Dim phi_b(1) As Double
    Dim phi_c(1) As Double
    Dim theta_a(1) As Double
    Dim theta_b(1) As Double
    Dim theta_c(1) As Double
    If rota < 0 Then
        da = 0
        dbb = 1
    Else
        da = 1
        dbb = 0
    End If
    
    phi_a(0) = RotCenLat
    theta_a(0) = RotCenLon
    phi_b(0) = OldLat
    theta_b(0) = OldLon
    
    c = ARCCOS(Sin(phi_a(0)) * Sin(phi_b(0)) + Cos(phi_b(0)) * Cos(phi_a(0)) * Cos(theta_b(0) - theta_a(0)))
    b = c
    a = ARCCOS((Cos((rota * pi) / 180) * Sin(b) * Sin(c)) + Cos(b) * Cos(c))
    If c = 0 Then
    Else
        alpha = ARCCOS((Cos(a) - Cos(b) * Cos(c)) / (Sin(b) * Sin(c)))
        psi = ARCCOS((Sin(phi_b(0)) - Sin(phi_a(0)) * Cos(c)) / (Cos(phi_a(0)) * Sin(c)))
        phi_c(0) = ARCSIN(Cos(b) * Sin(phi_a(0)) + Sin(b) * Cos(phi_a(0)) * Cos(psi - alpha))
        phi_c(1) = ARCSIN(Cos(b) * Sin(phi_a(0)) + Sin(b) * Cos(phi_a(0)) * Cos(psi + alpha))
        lambda(0) = ARCSIN(Sin(b) * Sin(psi - alpha) / Cos(phi_c(0)))
        lambda(1) = ARCSIN(Sin(b) * Sin(psi + alpha) / Cos(phi_c(1)))
        theta_c(0) = theta_a(0) + lambda(0) * Sgn(theta_b(0) - theta_a(0))
        theta_c(1) = theta_a(0) + lambda(1) * Sgn(theta_b(0) - theta_a(0))
        If theta_b(0) > theta_a(0) Then
                'record!lata.Text = tolattext(phi_c(da) + 0)
                'record!lona.Text = tolontext(theta_c(da) + 0)
                NewLat = phi_c(da)
                NewLon = theta_c(da)
        Else
                'record!lata.Text = tolattext(phi_c(dbb) + 0)
                'record!lona.Text = tolontext(theta_c(dbb) + 0)
                NewLat = phi_c(dbb)
                NewLon = theta_c(dbb)
        End If
    End If
End Sub

Private Sub Export_Click()
   Dim apppath As String
   If Right(App.Path, 1) = "\" Then
      apppath = App.Path
   Else
      apppath = App.Path & "\"
   End If
   Dim ms As String
   ms = InputBox("File Name (do not include the .txt extension)", , "FinishLine")
   If ms <> "" Then
      Dim minF, maxF As Integer
ag1:
      minF = Val(InputBox("What is the minimun finish height in feet MSL"))
      If minF = 0 Then GoTo ag1
ag2:
      maxF = Val(InputBox("What is the maximum finish height in feet MSL"))
      If maxF = 0 Then GoTo ag2
      Open apppath & ms & ".txt" For Output As 1
      Print #1, "INCLUDE = YES"
      Print #1, "TYPE=CTA/CTR"
      Print #1, "#"
      Print #1, "TITLE=FINISH LINE"
      Print #1, "#     IS CENTERED ON " & CenterLat.Text & " " & CenterLon.Text
      Print #1, "#     Endpoint 1 is " & Elat1.Text & " " & Elon1.Text
      Print #1, "#     Endpoint 2 is " & Elat2.Text & " " & Elon2.Text
      Print #1, "#     IS " & Lwidth.Text / 5280 & " MILES WIDE"
      Print #1, "#     AND HAS AN END TO END BEARING OF " & FTdir.Text & " DEGREES TRUE"
      Print #1, "#     THE APPROXIMATE ENDPOINTS ARE BELOW"
      Print #1, "#     THESE ARE ROUNDED TO NEAREST SECOND THAT FALLS INSIDE THE ACTUAL ENDPOINTS"
      Print #1, "#"
      Print #1, "BASE=" & minF
      Print #1, "TOPS=" & maxF
      Print #1, "POINT=" & NSlat1.Text & " " & NSlon1.Text
      Print #1, "POINT=" & NSlat2.Text & " " & NSlon2.Text
      Print #1, "End"
      Close 1
      MsgBox ("File written as " & apppath & ms & ".txt")
   End If
End Sub

Private Sub Form_Load()
   pi = 3.14159265358979
   Kpoint(0).Value = True
End Sub
Function ARCCOS(cosa As Double)
'in radians
    If Abs(cosa) > 1 Then
        cosa = cosa \ 1
    End If
    Select Case cosa
        Case 0
            ARCCOS = pi / 2
        Case Is < 0
            ARCCOS = pi + Atn((1 - cosa ^ 2) ^ 0.5 / cosa)
        Case Else
            If Int(100000 * cosa) = 100000 Then
                ARCCOS = 0
            Else
                ARCCOS = Atn((1 - cosa ^ 2) ^ 0.5 / cosa)
            End If
'            arccos = Atn(-cosa / Sqr(-cosa * cosa + 1)) + 2 * Atn(1)
    End Select
End Function
Function ARCSIN(sina As Double)
    Select Case sina
        Case 1, -1
            ARCSIN = pi / 2
        Case Is < 0
            'arcsin = pi + Atn(sina / (1 - sina ^ 2) ^ 0.5)
            ARCSIN = Atn(sina / (1 - sina ^ 2) ^ 0.5)
        Case Else
            ARCSIN = Atn(sina / (1 - sina ^ 2) ^ 0.5)
    End Select
End Function
Function radlattodec(inlat As Double)
    Dim ns As String
    Dim d, m As Integer
    Dim s As Single
    If inlat < 0 Then ns = "S" Else ns = "N"
    inlat = Abs(inlat * 180 / pi)
    d = Int(inlat)
    m = Int((inlat - d) * 60)
    s = (((inlat - d) * 60) - m) * 60
    If s >= 59.97 Then
        m = m + 1
        s = s - 60
    End If
    If m = 60 Then
        d = d + 1
        m = 0
    End If
    'If s >= 60 Then Stop
    radlattodec = String(2 - Len(Trim(d)), "0") & Trim(d) & ":" & String(2 - Len(Trim(m)), "0") & Trim(m) & Format(s / 60, ".000") & ns
End Function

Function radlontodec(inlon As Double)
    Dim ns As String
    Dim d, m As Integer
    Dim s As Single
    If inlon < 0 Then ns = "W" Else ns = "E"
    inlon = Abs(inlon * 180 / pi)
    d = Int(inlon)
    m = Int((inlon - d) * 60)
    s = (((inlon - d) * 60) - m) * 60
    If s >= 59.97 Then
        m = m + 1
        s = s - 60
    End If
    If m = 60 Then
        d = d + 1
        m = 0
    End If
    radlontodec = String(3 - Len(Trim(d)), "0") & Trim(d) & ":" & String(2 - Len(Trim(m)), "0") & Trim(m) & Format(s / 60, ".000") & ns
End Function

Function radlattosec(inlat As Double)
    Dim ns As String
    Dim d, m As Integer
    Dim s As Single
    If inlat < 0 Then ns = "S" Else ns = "N"
    inlat = Abs(inlat * 180 / pi)
    d = Int(inlat)
    m = Int((inlat - d) * 60)
    s = (((inlat - d) * 60) - m) * 60
    If s >= 59.97 Then
        m = m + 1
        s = s - 60
    End If
    If m = 60 Then
        d = d + 1
        m = 0
    End If
    'If s >= 60 Then Stop
    radlattosec = ns & String(2 - Len(Trim(d)), "0") & Trim(d) & String(2 - Len(Trim(m)), "0") & Trim(m) & Format(s, "00")
End Function

Function radlontosec(inlon As Double)
    Dim ns As String
    Dim d, m As Integer
    Dim s As Single
    If inlon < 0 Then ns = "W" Else ns = "E"
    inlon = Abs(inlon * 180 / pi)
    d = Int(inlon)
    m = Int((inlon - d) * 60)
    s = (((inlon - d) * 60) - m) * 60
    If s >= 59.97 Then
        m = m + 1
        s = s - 60
    End If
    If m = 60 Then
        d = d + 1
        m = 0
    End If
    radlontosec = ns & String(3 - Len(Trim(d)), "0") & Trim(d) & String(2 - Len(Trim(m)), "0") & Trim(m) & Format(s, "00")
End Function


Private Sub Kpoint_Click(Index As Integer)
   If Kpoint(0).Value = True Then
      CenterLat.Locked = True
      CenterLon.Locked = True
      CenterLat.Text = ""
      CenterLon.Text = ""
      Elat1.Text = "31:59.400N"
      Elon1.Text = "083:46.510W"
      Elat1.Locked = False
      Elon1.Locked = False
      Headlabel.Caption = "Heading to other end of line (degrees True)"
   Else
      CenterLat.Locked = False
      CenterLon.Locked = False
      CenterLat.Text = "31:59.400N"
      CenterLon.Text = "083:46.510W"
      Elat1.Text = ""
      Elon1.Text = ""
      Elat1.Locked = True
      Elon1.Locked = True
      Headlabel.Caption = "Heading to either end of line (degrees True)"
   End If
   NSlat1.Text = ""
   NSlon1.Text = ""
   NSlat2.Text = ""
   NSlon2.Text = ""
   Elat2.Text = ""
   Elon2.Text = ""
End Sub
