VERSION 5.00
Begin VB.UserControl MelCPB 
   ClientHeight    =   1005
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   855
   ScaleHeight     =   67
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   57
   ToolboxBitmap   =   "MelCPB.ctx":0000
   Begin VB.PictureBox Pic_InitialState 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   525
      Left            =   0
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   34
      TabIndex        =   2
      Top             =   600
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.PictureBox Pic_FinalState 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   525
      Left            =   0
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   34
      TabIndex        =   1
      Top             =   1200
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.PictureBox Pic_Render 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   34
      TabIndex        =   0
      Top             =   0
      Width           =   540
   End
End
Attribute VB_Name = "MelCPB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'
'___________________________________________________________________________
' Program name      : MelCPB OCX.
' Description       : A simple graphical circular progress bar OCX.
' Company           : MELANTECH
' Authors           : Weitten Pascal
'___________________________________________________________________________
'
' Date              : (c) 2005.01.15
' Version N°        : V0.1
' Customer          : Internal stuff.
'
' Last Modification : 2005.01.15
'___________________________________________________________________________
' TODO :
'       -
'       -
'___________________________________________________________________________
'

'Variables definition.
Dim LastInternalAngle As Double, LastExternalAngle As Double
Dim InternalAngle As Double, ExternalAngle As Double

'Constants
Const Pi = 3.141592654
Const MaxAngle_Degrees = 360
Const LineAngle = 1
Const TwipsPixelRatio = 0.067       'Could be done in another way.

'Valeurs de propriétés par défaut:
Const m_def_cpbDrawLine = 0
Const m_def_cpbDrawLineColor = 0
Const m_def_cpbDrawLineSize = 0
Const m_def_cpbInnerCircleMax = 100
Const m_def_cpbOuterCircleMax = 100
Const m_def_cpbPicBorder = 0
Const m_def_cpbSeparatedBars = 0
'Variables de propriétés:
Dim m_cpbDrawLine As Boolean
Dim m_cpbDrawLineColor As OLE_COLOR
Dim m_cpbDrawLineSize As Single
Dim m_cpbInnerCircleMax As Double
Dim m_cpbOuterCircleMax As Double
Dim m_cpbPicInitialState As Picture
Dim m_cpbPicFinalState As Picture
Dim m_cpbPicBorder As Boolean
Dim m_cpbSeparatedBars As Boolean


'ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENTÉES SUIVANTES!
'MemberInfo=0,0,0,0
Public Property Get cpbDrawLine() As Boolean
    cpbDrawLine = m_cpbDrawLine
End Property

Public Property Let cpbDrawLine(ByVal New_cpbDrawLine As Boolean)
    m_cpbDrawLine = New_cpbDrawLine
    PropertyChanged "cpbDrawLine"
End Property

'ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENTÉES SUIVANTES!
'MemberInfo=10,0,0,0
'Public Property Get cpbDrawLineColor() As OLE_COLOR
'    cpbDrawLineColor = m_cpbDrawLineColor
'End Property
'
'Public Property Let cpbDrawLineColor(ByVal New_cpbDrawLineColor As OLE_COLOR)
'    m_cpbDrawLineColor = New_cpbDrawLineColor
'    PropertyChanged "cpbDrawLineColor"
'End Property

'ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENTÉES SUIVANTES!
'MemberInfo=12,0,0,0
'Public Property Get cpbDrawLineSize() As Single
'    cpbDrawLineSize = m_cpbDrawLineSize
'End Property
'
'Public Property Let cpbDrawLineSize(ByVal New_cpbDrawLineSize As Single)
'    m_cpbDrawLineSize = New_cpbDrawLineSize
'    PropertyChanged "cpbDrawLineSize"
'End Property

'ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENTÉES SUIVANTES!
'MemberInfo=4,0,0,0
Public Property Get cpbInnerCircleMax() As Double
    cpbInnerCircleMax = m_cpbInnerCircleMax
End Property

Public Property Let cpbInnerCircleMax(ByVal New_cpbInnerCircleMax As Double)
    m_cpbInnerCircleMax = New_cpbInnerCircleMax
    PropertyChanged "cpbInnerCircleMax"
End Property

'ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENTÉES SUIVANTES!
'MemberInfo=4,0,0,0
Public Property Get cpbOuterCircleMax() As Double
    cpbOuterCircleMax = m_cpbOuterCircleMax
End Property

Public Property Let cpbOuterCircleMax(ByVal New_cpbOuterCircleMax As Double)
    m_cpbOuterCircleMax = New_cpbOuterCircleMax
    PropertyChanged "cpbOuterCircleMax"
End Property

'ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENTÉES SUIVANTES!
'MemberInfo=11,0,0,0
Public Property Get cpbPicInitialState() As Picture
    Set cpbPicInitialState = m_cpbPicInitialState
    Pic_InitialState.Picture = m_cpbPicInitialState
    Pic_Render.Picture = Pic_InitialState.Picture
   Width = CInt(Pic_Render.Width / TwipsPixelRatio)
   Height = CInt(Pic_Render.Height / TwipsPixelRatio)
End Property

Public Property Set cpbPicInitialState(ByVal New_cpbPicInitialState As Picture)
    Set m_cpbPicInitialState = New_cpbPicInitialState
    PropertyChanged "cpbPicInitialState"
End Property

'ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENTÉES SUIVANTES!
'MemberInfo=11,0,0,0
Public Property Get cpbPicFinalState() As Picture
    Set cpbPicFinalState = m_cpbPicFinalState
    Pic_FinalState.Picture = m_cpbPicFinalState
End Property

Public Property Set cpbPicFinalState(ByVal New_cpbPicFinalState As Picture)
    Set m_cpbPicFinalState = New_cpbPicFinalState
    PropertyChanged "cpbPicFinalState"
End Property

'ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENTÉES SUIVANTES!
'MemberInfo=0,0,0,0
Public Property Get cpbPicBorder() As Boolean
    cpbPicBorder = m_cpbPicBorder
    If m_cpbPicBorder Then
        Pic_Render.BorderStyle = 1
    Else
        Pic_Render.BorderStyle = 0
    End If
End Property

Public Property Let cpbPicBorder(ByVal New_cpbPicBorder As Boolean)
    m_cpbPicBorder = New_cpbPicBorder
    PropertyChanged "cpbPicBorder"
End Property

'ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENTÉES SUIVANTES!
'MemberInfo=0,0,0,0
'Public Property Get cpbSeparatedBars() As Boolean
'    cpbSeparatedBars = m_cpbSeparatedBars
'End Property
'
'Public Property Let cpbSeparatedBars(ByVal New_cpbSeparatedBars As Boolean)
'    m_cpbSeparatedBars = New_cpbSeparatedBars
'    PropertyChanged "cpbSeparatedBars"
'End Property


Private Sub UserControl_Initialize()
   Width = CInt(Pic_Render.Width / TwipsPixelRatio)
   Height = CInt(Pic_Render.Height / TwipsPixelRatio)
End Sub

'Initialiser les propriétés pour le contrôle utilisateur
Private Sub UserControl_InitProperties()
    m_cpbDrawLine = m_def_cpbDrawLine
    m_cpbDrawLineColor = m_def_cpbDrawLineColor
    m_cpbDrawLineSize = m_def_cpbDrawLineSize
    m_cpbInnerCircleMax = m_def_cpbInnerCircleMax
    m_cpbOuterCircleMax = m_def_cpbOuterCircleMax
    Set m_cpbPicInitialState = LoadPicture("")
    Set m_cpbPicFinalState = LoadPicture("")
    m_cpbPicBorder = m_def_cpbPicBorder
    m_cpbSeparatedBars = m_def_cpbSeparatedBars
End Sub

Private Sub UserControl_Paint()
    Pic_InitialState.Picture = m_cpbPicInitialState
    Pic_Render.Picture = Pic_InitialState.Picture
    Pic_FinalState.Picture = m_cpbPicFinalState
    If m_cpbPicBorder Then
        Pic_Render.BorderStyle = 1
    Else
        Pic_Render.BorderStyle = 0
    End If
    Width = CInt(Pic_Render.Width / TwipsPixelRatio) + 2
    Height = CInt(Pic_Render.Height / TwipsPixelRatio) + 2
End Sub

'Charger les valeurs des propriétés à partir du stockage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_cpbDrawLine = PropBag.ReadProperty("cpbDrawLine", m_def_cpbDrawLine)
    m_cpbDrawLineColor = PropBag.ReadProperty("cpbDrawLineColor", m_def_cpbDrawLineColor)
    m_cpbDrawLineSize = PropBag.ReadProperty("cpbDrawLineSize", m_def_cpbDrawLineSize)
    m_cpbInnerCircleMax = PropBag.ReadProperty("cpbInnerCircleMax", m_def_cpbInnerCircleMax)
    m_cpbOuterCircleMax = PropBag.ReadProperty("cpbOuterCircleMax", m_def_cpbOuterCircleMax)
    Set m_cpbPicInitialState = PropBag.ReadProperty("cpbPicInitialState", Nothing)
    Set m_cpbPicFinalState = PropBag.ReadProperty("cpbPicFinalState", Nothing)
    m_cpbPicBorder = PropBag.ReadProperty("cpbPicBorder", m_def_cpbPicBorder)
    m_cpbSeparatedBars = PropBag.ReadProperty("cpbSeparatedBars", m_def_cpbSeparatedBars)
End Sub

Private Sub UserControl_Resize()
   Width = CInt(Pic_Render.Width / TwipsPixelRatio) + 2
   Height = CInt(Pic_Render.Height / TwipsPixelRatio) + 2
End Sub

'Écrire les valeurs des propriétés dans le stockage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("cpbDrawLine", m_cpbDrawLine, m_def_cpbDrawLine)
    Call PropBag.WriteProperty("cpbDrawLineColor", m_cpbDrawLineColor, m_def_cpbDrawLineColor)
    Call PropBag.WriteProperty("cpbDrawLineSize", m_cpbDrawLineSize, m_def_cpbDrawLineSize)
    Call PropBag.WriteProperty("cpbInnerCircleMax", m_cpbInnerCircleMax, m_def_cpbInnerCircleMax)
    Call PropBag.WriteProperty("cpbOuterCircleMax", m_cpbOuterCircleMax, m_def_cpbOuterCircleMax)
    Call PropBag.WriteProperty("cpbPicInitialState", m_cpbPicInitialState, Nothing)
    Call PropBag.WriteProperty("cpbPicFinalState", m_cpbPicFinalState, Nothing)
    Call PropBag.WriteProperty("cpbPicBorder", m_cpbPicBorder, m_def_cpbPicBorder)
    Call PropBag.WriteProperty("cpbSeparatedBars", m_cpbSeparatedBars, m_def_cpbSeparatedBars)
End Sub

Private Function Draw_CPB(InternalAngle As Double, ExternalAngle As Double)
    Dim X As Integer, Y As Integer
    Dim X2 As Double, Y2 As Double
    Dim Radius As Double
    Dim i As Long, j As Double, k As Double
    Dim CosX As Double, CosY As Double, SinX As Double, SinY As Double
    Dim FillCircle_Colour As Long
    
    
    On Error Resume Next
    
    'Define the center of the circle position: X,Y.
    X = (Pic_InitialState.Width / 2)
    Y = (Pic_InitialState.Height / 2)
    
    'Defines the radius of the circle. Here X=Y=Radius.
    Radius = Pic_InitialState.Width / 2
    
    'Use of internal circle.
    If InternalAngle > -1 Then
        For i = LastInternalAngle To InternalAngle
            
            If cpbDrawLine Then
                'Draw a little line before
                'drawing the new angle position data.
                
                'Convert the angle: radians
                j = ((i + LineAngle) * Pi) / 180
                
                'For some precision reasons we only keep 5 digits.
                CosX = Format(Cos(j), "0.00000")
                SinX = Format(Sin(j), "0.00000")
                For k = 0 To X / 2
                    X2 = X + (k * CosX)
                    Y2 = Y + (k * SinX)
                    Pic_Render.PSet (X2, Y2), RGB(0, 0, 0)
                Next k
            End If
            
            'Draw the angle data.
            j = (i * Pi) / 180
            CosX = Format(Cos(j), "0.00000")
            SinX = Format(Sin(j), "0.00000")
            
            For k = 0 To X / 2
                X2 = X + (k * CosX)
                Y2 = Y + (k * SinX)
                FillCircle_Colour = Pic_FinalState.Point(X2, Y2)
                Pic_Render.PSet (X2, Y2), FillCircle_Colour
            Next k
        Next i
    End If
    
    
    'Use of external circle.
    If ExternalAngle > -1 Then
        For i = LastExternalAngle To ExternalAngle
            
            If cpbDrawLine Then
                'Draw a little line before
                'drawing the new angle position data.
                
                'Convert the angle: radians
                j = ((i + LineAngle) * Pi) / 180
                
                'For some precision reasons we only keep 5 digits.
                CosX = Format(Cos(j), "0.00000")
                SinX = Format(Sin(j), "0.00000")
                For k = X / 2 To X
                    X2 = X + (k * CosX)
                    Y2 = Y + (k * SinX)
                    Pic_Render.PSet (X2, Y2), RGB(0, 0, 0)
                Next k
            End If
        
            'Draw the angle data.
            j = (i * Pi) / 180
    
            CosX = Format(Cos(j), "0.00000")
            SinX = Format(Sin(j), "0.00000")
    
            For k = X / 2 To X
                X2 = X + (k * CosX)
                Y2 = Y + (k * SinX)
                FillCircle_Colour = Pic_FinalState.Point(X2, Y2)
                Pic_Render.PSet (X2, Y2), FillCircle_Colour
            Next k
        Next i
    End If
End Function

Private Function Restore_InternalCircle()
    Dim X As Integer, Y As Integer
    Dim X2 As Double, Y2 As Double
    Dim Radius As Double
    Dim i As Long, j As Double, k As Double
    Dim CosX As Double, CosY As Double, SinX As Double, SinY As Double
    Dim FillCircle_Colour As Long
    
    
    On Error Resume Next
    'Define the center of the circle position: X,Y.
    X = (Pic_InitialState.Width / 2)
    Y = (Pic_InitialState.Height / 2)
    
    'Defines the radius of the circle. Here X=Y=Radius.
    Radius = Pic_InitialState.Width / 2
    
    'Use the internal circle.
    For i = 0 To MaxAngle_Degrees
        'Convert the angle: radians
        j = (i * Pi) / 180
        
        'For some precision reasons we only keep 5 digits.
        CosX = Format(Cos(j), "0.00000")
        SinX = Format(Sin(j), "0.00000")
        
        For k = 0 To X / 2
            X2 = X + (k * CosX)
            Y2 = Y + (k * SinX)
            FillCircle_Colour = Pic_InitialState.Point(X2, Y2)
            Pic_Render.PSet (X2, Y2), FillCircle_Colour
        Next k
    Next i
End Function

Public Function Init_CircularBar()
    On Error Resume Next
    'Affecte l'image source à l'image résultat
    
    LastInternalAngle = 0
    LastExternalAngle = 0
    InternalAngle = 0
    ExternalAngle = 0
    Pic_Render.Picture = Pic_InitialState.Picture
End Function

Public Function DrawBothCircles(AngleStep As Double)
        If AngleStep = m_cpbOuterCircleMax Then
            AngleStep = MaxAngle_Degrees
        Else
            AngleStep = AngleStep * (MaxAngle_Degrees / m_cpbOuterCircleMax)
        End If
        
        If ExternalAngle < (MaxAngle_Degrees) Then ' - CInt(MaxAngle_Degrees / m_cpbOuterCircleMax)) Then
            InternalAngle = AngleStep
            ExternalAngle = AngleStep
            Call Draw_CPB(InternalAngle, ExternalAngle)
            LastInternalAngle = InternalAngle
            LastExternalAngle = ExternalAngle
       Else
            LastInternalAngle = 0
            LastExternalAngle = 0
            InternalAngle = 0
            ExternalAngle = 0
            Pic_Render.Picture = Pic_InitialState.Picture
        End If
  
End Function

Public Function DrawInnerCircle(AngleStep As Double)
    If AngleStep > 0 Then
        If AngleStep = m_cpbInnerCircleMax Then
            AngleStep = MaxAngle_Degrees
        Else
            AngleStep = AngleStep * (MaxAngle_Degrees / m_cpbInnerCircleMax)
        End If
        AngleStep = CInt(AngleStep)
        
        If InternalAngle < (MaxAngle_Degrees) Then ' - CInt(MaxAngle_Degrees / m_cpbInnerCircleMax)) Then
            InternalAngle = AngleStep
            Call Draw_CPB(InternalAngle, -1)
            LastInternalAngle = InternalAngle
        Else
            LastInternalAngle = 0
            InternalAngle = 0
            Call Restore_InternalCircle
        End If
    End If
End Function

Public Function DrawOuterCircle(AngleStep As Double)
    If AngleStep > 0 Then
        If AngleStep = m_cpbOuterCircleMax Then
            AngleStep = MaxAngle_Degrees
        Else
            AngleStep = AngleStep * (MaxAngle_Degrees / m_cpbOuterCircleMax)
        End If
        
        AngleStep = CInt(AngleStep)
        
        If ExternalAngle < (MaxAngle_Degrees) Then ' - CInt(MaxAngle_Degrees / m_cpbOuterCircleMax)) Then
            ExternalAngle = AngleStep
            Call Draw_CPB(-1, ExternalAngle)
            LastExternalAngle = ExternalAngle
        Else
            LastExternalAngle = 0
            ExternalAngle = 0
        End If
    End If
End Function


