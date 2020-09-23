VERSION 5.00
Begin VB.Form fmain 
   BackColor       =   &H00000000&
   ClientHeight    =   2250
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   4500
   Icon            =   "fmain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   150
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   300
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "fmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Snow flakes by  Eric Sanford aka Lord_illogical


Dim mAmoutOfSnowFall    As Integer  'number of snow flakes to add on each loop
Dim mWeAreGo            As Boolean  'used to stop loop so prog can end
Dim mclsFlake           As cFlake       'line class of a snow flake
Dim mclsFlakes          As cFlakeTable  'table collection class of all the falling snow flakes

Dim mWind          As Single   'if positive wind gose right negitive left

Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long



Private Sub Form_Activate()
    Dim c           As Long
    Dim i           As Integer
    Dim Hsw         As Single
    Dim sh          As Single
    Dim BC          As Long
    Dim SnowColor   As Long
    Dim stestWind   As String
    Randomize Timer     'good seed for randome numbers
    
    'use local varibls for speed
    Hsw = (Me.ScaleWidth / 2)
    sh = Me.ScaleHeight
    BC = Me.BackColor
    SnowColor = QBColor(15)
    

    If mWeAreGo Then
        Do
            DoEvents
            
            For i = 1 To mclsFlakes.Count
            
                DoEvents    'need this and one above to let windows work
                
                
                Set mclsFlake = mclsFlakes.Item(i)
                
                
                If Not mclsFlake Is Nothing Then
                    
                    With mclsFlake
                    
                        If Not .OnGround Then
                            
                            If .y + 1 > sh Then 'check if hit bottum
                            
                                .OnGround = True    'flag for later
                                
                            Else
                                
'                                ''wind not working right dont got time to figure it out
'                                If mWind > 0 Then
'                                    .NextX = .x + 1
'                                Else
'                                    .NextX = .x - 1
'                                End If
                                .NextX = .x
                                
                                'check pixl blow see if we hit something
                                c = GetPixel(Me.hdc, .NextX, .y + 1)     'if something covers the form this must return whats over the for, the snow stops there
                                If c = BC Then  'ok to move there
                                    Me.PSet (.x, .y), BC
                                    .y = .y + 1
                                    .x = .NextX
                                    Me.PSet (.x, .y), SnowColor
                                    
                                Else
                                    
                                    If .x > Hsw Then    'if on right side of screen (check right first)
                                    
                                        c = GetPixel(Me.hdc, .x + 1, .y + 1)
                                        If c = BC Then  'check pixel 1 down 1 to right
                                            Me.PSet (.x, .y), BC
                                            .y = .y + 1
                                            .x = .x + 1
                                            Me.PSet (.x, .y), SnowColor
                                        Else
                                        
                                            c = GetPixel(Me.hdc, .x - 1, .y + 1)
                                            If c = BC Then  'check pixel 1 down 1 left
                                                Me.PSet (.x, .y), BC
                                                .y = .y + 1
                                                .x = .x - 1
                                                Me.PSet (.x, .y), SnowColor
                                            Else
                                                .OnGround = True
                                            End If
                                        End If
                                        
                                    Else    'else left sid of screen (check left first)
                                    
                                        c = GetPixel(Me.hdc, .x - 1, .y + 1)
                                        If c = BC Then  'check pixel 1 down 1 left
                                            Me.PSet (.x, .y), BC
                                            .y = .y + 1
                                            .x = .x - 1
                                            Me.PSet (.x, .y), SnowColor
                                        Else
                                            
                                            c = GetPixel(Me.hdc, .x + 1, .y + 1)
                                            If c = BC Then  'check pixel 1 down 1 right
                                                Me.PSet (.x, .y), BC
                                                .y = .y + 1
                                                .x = .x + 1
                                                Me.PSet (.x, .y), SnowColor
                                            Else
                                                .OnGround = True
                                            End If
                                        End If
                                        
                                        
                                    End If
                                    
                                End If
                                
                            End If
                            
                        End If
                        
                        If .OnGround Then   'dont need to keep trak of this flake anymore so free it
                            mclsFlakes.Remove i
                            
                            'i notices some flakes jump sometimes
                            'think cause they get skiped when the one befor it is deleted
                            'i = i - 1
                            'now im not seeing it
                        End If
                        
                    End With
                    
                End If
                
                If Not mWeAreGo Then Exit Do    'tite loop here so check if we need to quit
            Next i
            
            'add new flakes
            For i = 1 To mAmoutOfSnowFall
                'get a x to start at (y will be 0 so it starts at top)
                Do
                    c = (Rnd * 100) * (Rnd * 10)
                Loop Until c < 300
                mclsFlakes.Add c, 0, False
                
            Next i
            
            
            Me.Caption = "Falling snow flake count =" & mclsFlakes.Count
            
            
            'do wind stuff
            If mWind < 0 Then  'dont let direction change
                mWind = Rnd
                mWind = mWind * -1
            Else
                mWind = Rnd
            End If
            
            
        Loop Until Not mWeAreGo     'loop untile we need to quit program
        
    End If
    
End Sub

Private Sub Form_DblClick()
    Dim frmAbout As fAbout
    Set frmAbout = New fAbout
    frmAbout.Show
    frmAbout.Top = Me.Top + Me.Height
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
    
    Case 189, 109   'minus keys
        mAmoutOfSnowFall = mAmoutOfSnowFall - 1
    
    Case 187, 107   'plus keys
        mAmoutOfSnowFall = mAmoutOfSnowFall + 1
    
    Case vbKeyW     'change wind direction
        mWind = mWind * -1
    End Select
End Sub

Private Sub Form_Load()

    Dim frmAbout As fAbout
    'create flake collection
    Set mclsFlakes = New cFlakeTable
    
    'start with 3 flake a loop
    mAmoutOfSnowFall = 3
    mWind = -1
    
    mWeAreGo = True
    
    Me.Show 'show last so stuff can be setup befor the loop starts
    Set frmAbout = New fAbout
    frmAbout.Show
    frmAbout.Top = Me.Top + Me.Height
    frmAbout.Timer1.Enabled = True
    frmAbout.ZOrder 0
    DoEvents
    
End Sub

Private Sub Form_Paint()
    Dim i%
    
    'if you want to draw something do it here
    'so its redrawn if another window gose over this window
    'autoredraw = true would let you not need to redraw all the time but slows down everthing
    
    For i = 80 To 83
        Me.Line (50, i + 30)-(80, i), QBColor(4)
        Me.Line (80, i)-(110, i + 30), QBColor(4)
    Next i
    For i = 52 To 55
        Me.Line (i, 80)-(i, Me.ScaleHeight), QBColor(4)
        Me.Line (i + 52, 108)-(i + 52, Me.ScaleHeight), QBColor(4)
    Next i
    For i = 77 To 83
        Me.Line (i, 130)-(i, Me.ScaleHeight), QBColor(4)
    Next i
    Me.Line (79, 140)-(79, 141), 0
    
    Me.Line (61, 110)-(72, 120), QBColor(4), B
    Me.Line (88, 110)-(99, 120), QBColor(4), B
    'Me.Line (79, 140)-(79, 141), 0
    'Me.Line (79, 140)-(79, 141), 0
    'Me.Line (79, 140)-(79, 141), 0
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'tell loop to quit
    mWeAreGo = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'free classes
    Set mclsFlake = Nothing
    Set mclsFlakes = Nothing
    
    End
End Sub
