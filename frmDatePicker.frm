VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDatePicker 
   Caption         =   "Calendario"
   ClientHeight    =   4035
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3975
   OleObjectBlob   =   "frmDatePicker.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmDatePicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''El frmDatePicker utiliza el tag del userform (Me.Tag) para almacenar en cada momento la fecha seleccionada (y para recibir la fecha inicial, para devolver la fecha, etc.)

Option Explicit

Private Sub UserForm_Activate()
'Si Me.Tag está vacío, se utilizar la fecha de hoy como fecha inicial
If Me.Tag = "" Then
    Me.Tag = Date
End If

Populate_Form (Me.Tag)

End Sub

Sub Populate_Form(dTracking As Date)

Dim i As Integer
Dim intYear As Integer
Dim intMonth As Integer

Dim dFirstDayCal As Date
Dim intMonthIntinialWeekday As Integer

intYear = year(dTracking)
intMonth = month(dTracking)

Me.txtYear.Caption = intYear
'Originalmente utilizaba una función creada, Month_Esp, que es innecesaria, ya que Format(Date, "mmmm") devuelve los meses en el idioma del sistema
'Me.txtMonth.Caption = Month_Esp(intMonth)
Me.txtMonth.Caption = UCase(Format(dTracking, "mmmm"))

'Establece (empezando por lunes = 1) el día de la semana de la fecha inicial del mes (DateSerial(intYear, intMonth, 1))
intMonthIntinialWeekday = Weekday(DateSerial(intYear, intMonth, 1), vbMonday)

'Establece la primera fecha a mostrar del calendario, restando del día inicial del mes (DateSerial(intYear, intMonth, 1)) los días hasta el lunes (-intMonthIntinialWeekday+1)
dFirstDayCal = DateAdd("d", -intMonthIntinialWeekday + 1, DateSerial(intYear, intMonth, 1))

'Recorre todos los controles y va añadiendo el día correspondiente, desde dFirstDayCal

For i = 0 To 41

    ' Pone en cada controlador el numero del día correspondiente
    Me.Controls("txtDay" & i + 1).Caption = day(dFirstDayCal + i)
    ' Guarda en la tag del controlador la fecha completa (día,mes y año) correspondiente, para futuros cálculos
    Me.Controls("txtDay" & i + 1).Tag = dFirstDayCal + i
    
    
    'Pone en gris los días que no pertenecen al mes actual
    If month(dFirstDayCal + i) <> intMonth Then
    
        Me.Controls("txtDay" & i + 1).ForeColor = 8421504
    
    'Pone en azul claro los días del fin de semana (sabado y domingo)
    ElseIf ((i + 1) Mod 7) = 0 Or ((i + 2) Mod 7) = 0 Then
    
        Me.Controls("txtDay" & i + 1).ForeColor = &HC08B0E
    
    Else
    
    'Pone normal (gris oscuro) el resto de días, para que no se queden coloreados al ir cambiando los meses
    
        Me.Controls("txtDay" & i + 1).ForeColor = &H464646
    
    End If
    
      
    'Sobrea de azul la fecha elegida
    If dFirstDayCal + i = dTracking Then
    
        Me.Controls("txtDay" & i + 1).BackColor = &HFFC0C0
        Me.Controls("txtDay" & i + 1).BorderStyle = 1
        
    ' Sombrea de rosa la fecha de hoy (si no es la fecha elegida)
    ElseIf dFirstDayCal + i = Date Then
    
        Me.Controls("txtDay" & i + 1).BackColor = &HC0C0FF
        Me.Controls("txtDay" & i + 1).BorderStyle = 0
    
    'En el resto de casos, sombrea de blanco (necesarios para que cuando cambie el calendario, no se vayan quedando casillas de colores)
    Else
    
        Me.Controls("txtDay" & i + 1).BackColor = &HFFFFFF
        Me.Controls("txtDay" & i + 1).BorderStyle = 0

    End If
    
Next i

'Introduce la fecha de hoy, indicando día de la semana, empezando en mayúscula (StrConv --> vbProperCase), y con el mes en letra (Format(Date, "mmmm"))

Me.lblHoy.Caption = StrConv(Format(Date, "dddd"), vbProperCase) & ", " & day(Date) & " de " & Format(Date, "mmmm") & " de " & year(Date)

End Sub

Private Sub lblHoySelect_Click()
'Establece el día de hoy como fecha seleccionada, y refresca el calendario
Me.Tag = Date

Populate_Form (Me.Tag)

End Sub

Private Sub lblNextYear_Click()
'Añade un año a la fecha seleccionada, y refresca el calendario
Me.Tag = DateAdd("yyyy", 1, Me.Tag)

Populate_Form (Me.Tag)

End Sub

Private Sub lblPrevYear_Click()
'Resta un año a la fecha seleccionada, y refresca el calendario
Me.Tag = DateAdd("yyyy", -1, Me.Tag)

Populate_Form (Me.Tag)

End Sub

Private Sub lblNextMonth_Click()
'Añade un mes a la fecha seleccionada, y refresca el calendario

Me.Tag = DateAdd("m", 1, Me.Tag)

Populate_Form (Me.Tag)

End Sub
Private Sub lblPrevtMonth_Click()
'Resta un mes a la fecha seleccionada, y refresca el calendario

Me.Tag = DateAdd("m", -1, Me.Tag)

Populate_Form (Me.Tag)

End Sub

'Una serie de funciones para, al clickar en cada uno de los días, remitir a la función DayClick con el número correspondiente de controlador
Private Sub txtDay1_Click(): DayClick (1): End Sub
Private Sub txtDay2_Click(): DayClick (2): End Sub
Private Sub txtDay3_Click(): DayClick (3): End Sub
Private Sub txtDay4_Click(): DayClick (4): End Sub
Private Sub txtDay5_Click(): DayClick (5): End Sub
Private Sub txtDay6_Click(): DayClick (6): End Sub
Private Sub txtDay7_Click(): DayClick (7): End Sub
Private Sub txtDay8_Click(): DayClick (8): End Sub
Private Sub txtDay9_Click(): DayClick (9): End Sub
Private Sub txtDay10_Click(): DayClick (10): End Sub
Private Sub txtDay11_Click(): DayClick (11): End Sub
Private Sub txtDay12_Click(): DayClick (12): End Sub
Private Sub txtDay13_Click(): DayClick (13): End Sub
Private Sub txtDay14_Click(): DayClick (14): End Sub
Private Sub txtDay15_Click(): DayClick (15): End Sub
Private Sub txtDay16_Click(): DayClick (16): End Sub
Private Sub txtDay17_Click(): DayClick (17): End Sub
Private Sub txtDay18_Click(): DayClick (18): End Sub
Private Sub txtDay19_Click(): DayClick (19): End Sub
Private Sub txtDay20_Click(): DayClick (20): End Sub
Private Sub txtDay21_Click(): DayClick (21): End Sub
Private Sub txtDay22_Click(): DayClick (22): End Sub
Private Sub txtDay23_Click(): DayClick (23): End Sub
Private Sub txtDay24_Click(): DayClick (24): End Sub
Private Sub txtDay25_Click(): DayClick (25): End Sub
Private Sub txtDay26_Click(): DayClick (26): End Sub
Private Sub txtDay27_Click(): DayClick (27): End Sub
Private Sub txtDay28_Click(): DayClick (28): End Sub
Private Sub txtDay29_Click(): DayClick (29): End Sub
Private Sub txtDay30_Click(): DayClick (30): End Sub
Private Sub txtDay31_Click(): DayClick (31): End Sub
Private Sub txtDay32_Click(): DayClick (32): End Sub
Private Sub txtDay33_Click(): DayClick (33): End Sub
Private Sub txtDay34_Click(): DayClick (34): End Sub
Private Sub txtDay35_Click(): DayClick (35): End Sub
Private Sub txtDay36_Click(): DayClick (36): End Sub
Private Sub txtDay37_Click(): DayClick (37): End Sub
Private Sub txtDay38_Click(): DayClick (38): End Sub
Private Sub txtDay39_Click(): DayClick (39): End Sub
Private Sub txtDay40_Click(): DayClick (40): End Sub
Private Sub txtDay41_Click(): DayClick (41): End Sub
Private Sub txtDay42_Click(): DayClick (42): End Sub

Sub DayClick(i)
'Establece como fecha seleccionada la que esté guardada en el correspondiente controlador de día
Me.Tag = Me.Controls("txtDay" & i).Tag
Populate_Form (Me.Tag)

End Sub

' Innecesario, porque Format(Date, "mmmm") ya devuelve el nombre en el idioma del sistema
Function Month_Esp(month) As String

Select Case month
Case 1
    Month_Esp = "ENERO"
Case 2
    Month_Esp = "FEBRERO"
Case 3
    Month_Esp = "MARZO"
Case 4
    Month_Esp = "ABRIL"
Case 5
    Month_Esp = "MAYO"
Case 6
    Month_Esp = "JUNIO"
Case 7
    Month_Esp = "JULIO"
Case 8
    Month_Esp = "AGOSTO"
Case 9
    Month_Esp = "SEPTIEMBRE"
Case 10
    Month_Esp = "OCTUBRE"
Case 11
    Month_Esp = "NOVIEMBRE"
Case 12
    Month_Esp = "DICIEMBRE"

End Select

End Function

