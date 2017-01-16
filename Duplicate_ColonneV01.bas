Attribute VB_Name = "Duplicate_Colonne"
Sub Duplicate_Column()
'Macro qui permet de dupliquer une plage de n Colonnes.
'
'Ces n colonnes sont d�finies en dure dans le code. En revanche, les lignes qui composent ces colonnes ne le sont pas.
'Le "haut" des lignes est d�finit dans le code. En revanche, le "bas" est d�finit � l'aide de la balise /*vide*\.
'Cette balise doit donc �tre pr�sente en bas de chaque colonne et en fin de chaque ligne du tableau.
'
'remarque :  cette m�thode est plus lourde que Selection. etc...(par exemple, le code g�n�r� par l'enregistreur de macro).
'        Mais la m�thode Selection d�pend de l'objet Application qui g�n�re des erreurs.
'
'Algorithme et principe de base:
'    1 - On d�finit la Range qui va servir de Template pour les colonnes � Dupliquer
'        a - le haut des colonnes est indiquer dans le code.
'        b - le bas du tableau est trouv� � l'aide de la fonction FIND param�tr�e pour la balise /*vide*\
'        c - la fonction FIND renvoi le Range du bas du tableau. On construit alors le Range du template � partir de cela.
'        d - On copie cette range.
'    2 - On recherche le point d'insertion pour ajouter les colonnes voulues.
'        a - on proc�de � la recherche de la "derni�re" colonne � l'aide de la fonction FIND param�tr�e pour la balise /*vide*\ comme pr�c�demment (� quelques diff�rences)
'        b - on se place � ce point d'insertion et on utilise la fonction INSERT pour ajouter les colonnes.
'    3 - Afin de "lib�rer" la s�lection, on vient s�lectionner le point d'insertion et on utilise la fonction SENDKEYS.
'    4 - Cette m�me astuce, permet �galement de r�cup�rer le Focus sur ce point d'insertion (bien pratique pour l'aspect visuel). Ceci afin de d�placer directement le fichier Excel sur la zone de travail voulue.
'    5 - On rempli la case "Week" par le num�ro de semaine, � l'aide d'un fonction g�n�rique de r�cup�ration du n� de semaine � partir de la date courante.
'
'remarque :  le fichier contient un second morceau de code dans l'object Workbook-->Open, afin d'avoir le focus directement sur la zone de travail voulue.

    Dim myWS As Worksheet
    Dim myRangeReference As Range
    Dim myRangeInsertion As Range
    Dim myRangeWeek As Range
    
    Dim myRangeRecherche, myRangeTrouve, myRangeTemplate As Range
    
    Set myWS = ThisWorkbook.ActiveSheet
    
    'Define the template which will be copy/paste
    Set myRangeTemplate = myWS.Range("G1:K1").EntireColumn
    'Search the last row with "/*vide*\" inside, to determine the bottom of the Range
    Set myRangeRecherche = myWS.Range("G1").EntireColumn    'Where to look for the String
    Set myRangeTrouve = myRangeRecherche.Find("/*vide*\", After:=Range("G1"), lookAt:=xlWhole, searchorder:=xlByColumns, MatchCase:=True)
    
    
    Set myRangeTemplate = myWS.Range(myRangeTrouve, "K1")  'Column n+1 , Column n
    myRangeTemplate.Copy    'Copy the Template

    'Search for the insertion point
    'Insertion point is the first column before the column containing "/*vide*\" inside
    Set myRangeRecherche = myWS.Range("A1").EntireRow    'Where to look for the String
    Set myRangeTrouve = myRangeRecherche.Find("/*vide*\", After:=Range("A1"), lookAt:=xlWhole, searchorder:=xlByRows, MatchCase:=True)
    
    'Insert the data of the Template range
    myRangeTrouve.Insert shift:=xlShiftToRight
    
    'Tips to unselect the template range
    myRangeTrouve.Copy
    myRangeTrouve.Activate
    SendKeys "{Enter}"
    
        
    'Offset to set Date
    '12 bas et 5 � gauche
    Set myRangeWeek = myRangeTrouve.Offset(rowOffset:=12, columnOffset:=-5)
    Debug.Print "Offset : " & myRangeWeek.AddressLocal & " " & myRangeWeek.Value
    myRangeWeek.Value = "W" & RecuperationDateString(1)
    Debug.Print "Value : " & myRangeWeek.Value
    
    'Format the Column Width to 3.25
    'from column G which contains the Template.
    Range("G1", myRangeTrouve).ColumnWidth = 3.25


    


End Sub



Private Function RecuperationDateString(ByVal ChoixSortie As Integer) As String
'   Function wich permits to get the actual system Date as String value
'   From the date, the function extracts the Year as YY
'   From the date, the function extracts the Week number as WW (from 01 to 52)
'   And the function proposes various String format output

    Dim myDate As Date
    Dim myVariante As Integer
    Dim tmp As String

    'Get the system DATE
    myDate = Date
    



Select Case ChoixSortie
    Case 0
        'Format YY
        myVariante = Right(DatePart("yyyy", myDate), 2)
        RecuperationDateString = myVariante
    Case 1
        'Format WW - Week Number
        
        'Get the Week Number from the current Date
        myVariante = DatePart("ww", myDate, vbUseSystemDayOfWeek, vbFirstFullWeek)
        
        'Format the date to have '0' for value under '10' --> 01, 02, 09, 10, ...52,
        If (myVariante < 10) Then
            RecuperationDateString = "0" & myVariante
        Else
            RecuperationDateString = myVariante
        End If
        
    Case 2
        'Format YYWW -
        
        'Get the Week Number from the current Date
        myVariante = DatePart("ww", myDate, vbUseSystemDayOfWeek, vbFirstFullWeek)
        
        'Format the date to have '0' for value under '10' --> 01, 02, 09, 10, ...52,
        If (myVariante < 10) Then
            tmp = "0" & myVariante
        Else
            tmp = myVariante
        End If
        
        'Format and concatenate the YYWW output
        RecuperationDateString = Right(DatePart("yyyy", myDate), 2) & tmp
    Case Else
        RecuperationDateString = myDate
        
End Select

Debug.Print "Format de sortie choisi pour la fonction : "; RecuperationDateString



End Function



