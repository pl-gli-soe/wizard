VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PodmienForm 
   Caption         =   "Podmien"
   ClientHeight    =   1830
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3915
   OleObjectBlob   =   "PodmienForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PodmienForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' FORREST SOFTWARE
' Copyright (c) 2016 Mateusz Forrest Milewski
'
' Permission is hereby granted, free of charge,
' to any person obtaining a copy of this software and associated documentation files (the "Software"),
' to deal in the Software without restriction, including without limitation the rights to
' use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software,
' and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
' INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
' IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,
' WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.


Public nazwa_aktywnego_pliku As String
Public ph As PodmienHandler


Private Sub BtnClear_Click()
    hide
    Set ph = New PodmienHandler
    ph.ref_first ""
    ph.just_clear
    
    expand_alerts_on_reconf_i_oncost
End Sub

Private Sub BtnCustomCopy_Click()
    ' MsgBox "implementacja nie jest jeszcze gotowa"
    hide
    
    
    Dim czy_lecimy_z_tym As Boolean
    czy_lecimy_z_tym = True
    ' taka o smieszna blokada
    If czy_lecimy_z_tym Then
        
        hide
        nazwa_aktywnego_pliku = ""
    
        Set ph = New PodmienHandler
        
        With ph
        
            
            ' wypelnij list box wybierzplikform
            With WybierzPlikForm
                .ListBox1.Clear
                
                For Each w In Workbooks
                    If Not w.Name = ThisWorkbook.Name Then
                        .ListBox1.AddItem w.Name
                    End If
                Next w
            End With
            
            WybierzPlikForm.show
            ' ponoc od tego momentu pod nazwa_aktywnego_pliku znajduje sie juz konkretna tresc :D
            ' super uber side effect :P
            
            'prepare ref
            ' tutaj dodatkowo drugi argument zeby wiadomo bylo od razu ze lecimy poza std
            .ref_first nazwa_aktywnego_pliku, True
            
            
            
            
            ' ta procedura jest wtedy kiedy kopiujemy pomiedzy wizardami w tym samym standardzie
            ' .go_on_with_coping_data
            
            
            ' !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
            ' ============================================
            ' metoda ktora odmieni oblicze ziemi
            ' powinna dzialac zarowno dla normalnego wizarda jak i customowego
            .copy_from_raw_td
            ' ============================================
        End With
        
        
        expand_alerts_on_reconf_i_oncost
    End If
End Sub

Private Sub BtnWizardStdCopy_Click()
    hide
    nazwa_aktywnego_pliku = ""
    
    Set ph = New PodmienHandler
    
    With ph
    
        
        ' wypelnij list box wybierzplikform
        With WybierzPlikForm
            .ListBox1.Clear
            
            For Each w In Workbooks
                If Not w.Name = ThisWorkbook.Name Then
                    .ListBox1.AddItem w.Name
                End If
            Next w
        End With
        
        WybierzPlikForm.show
        ' ponoc od tego momentu pod nazwa_aktywnego_pliku znajduje sie juz konkretna tresc :D
        ' super uber side effect :P
        
        'prepare ref
        .ref_first nazwa_aktywnego_pliku
        .go_on_with_coping_data
    End With
    
    expand_alerts_on_reconf_i_oncost
End Sub
