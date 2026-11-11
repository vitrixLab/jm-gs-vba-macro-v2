Attribute VB_Name = "ImportSuppliersFromJSON"
Option Explicit

Sub ImportSuppliersFromJSON()
    Dim ws As Worksheet
    Dim rowNum As Long
    Dim i As Long
    Dim lastRow As Long
    
    Set ws = ThisWorkbook.Sheets("SUPPLIERS DATA")
    

    ' Build import set in memory before clearing anything.
    Dim records As New Collection
    Dim recValues As Variant
    Dim j As Long
    Dim duplicateFound As Boolean
    Dim placeholderAddr As String
    Dim name As String, address As String, name2 As String, address2 As String
    Dim inner As String

            records.Add Array(1, "008-888-659", "VPHARMA HEALTH AND WELNESS INC", "ALABANG, MUNTINLUPA", "Clinic Materials and Supplies")
            records.Add Array(2, "246-099-058", "TRAVELLERS INTL HOTEL INC", "NEWPORT, PASAY", "Transportation and Travel")
            records.Add Array(3, "008-120-744", "RUSTAN", "SAN LORENZO, MAKATI", "Representation")
            records.Add Array(4, "008-625-341", "THE MESSHALL", "MAGALLANES, MAKATI", "Miscellaneous")
            records.Add Array(5, "252-126-611", "EUROFUEL INC", "183 PASAY", "Fuel and Oil")
            records.Add Array(6, "008-909-992", "SOUTHEASTASIA RETAIL INC", "BGC, TAGUIG", "Supplies")
            records.Add Array(7, "000-388-771", "JOLLIBEE 39TH", "BGC, TAGUIG", "Representation")
            records.Add Array(8, "207-961-175", "SAVEMORE", "ALABANG, MUNTINLUPA", "Supplies")
            records.Add Array(9, "001-799-123", "UNIGLOBE TRAVELWARE INC", "ALABANG, MUNTINLUPA", "Transportation and Travel")
            records.Add Array(10, "148-855-506", "MINDWERKS", "ALABANG, MUNTINLUPA", "Supplies")
            records.Add Array(11, "007-851-409", "PRIMERA CLASE GAS STATION", "ALABANG, MUNTINLUPA", "Fuel and Oil")
            records.Add Array(12, "200-035-311", "ACE HARDWARE", "ALMANZA UNO, LAS PIÑAS", "Repairs and Maintenance")
            records.Add Array(13, "000-388-771", "JOLLIBEE BLUEBAY WALK", "76 PASAY", "Representation")
            records.Add Array(14, "246-969-491", "S AND R", "BGC, TAGUIG", "Representation")
            records.Add Array(15, "648-700-952", "GOOD MORNING FOOD", "BGC, TAGUIG", "Representation")
            records.Add Array(16, "008-087-158", "PNC GARBROS FOODS INC", "ALABANG, MUNTINLUPA", "Representation")
            records.Add Array(17, "208-006-656", "SM STORE", "ORANBO, PASIG", "Supplies")
            records.Add Array(18, "214-706-591", "WATSONS", "ALABANG, MUNTINLUPA", "Supplies")
            records.Add Array(19, "608-395-184", "GIANETO PIZZERIA", "ORANBO, PASIG", "Representation")
            records.Add Array(20, "010-624-401", "MIT SUKOSHI FRESH", "BGC, TAGUIG", "Representation")
            records.Add Array(21, "655-591-162", "MINISO", "ALABANG, MUNTINLUPA", "Supplies")
            records.Add Array(22, "000-405-340", "ROBINSONS", "ALABANG, MUNTINLUPA", "Supplies")
            records.Add Array(23, "000-123-826", "GRAND UNION SUPERMARKET INC", "FILINVEST, MUNTINLUPA", "Supplies")
            records.Add Array(24, "008-188-059", "KNOXPORT INC", "BGC, TAGUIG", "Miscellaneous")
            records.Add Array(25, "009-754-702", "GOOD SMALL SHEEP INC", "DAANG BAKAL, MANDALUYONG", "Miscellaneous")
            records.Add Array(26, "163-955-615", "725 TAGTEAM FUEL STATION", "ALABANG, MUNTINLUPA", "Fuel and Oil")
            records.Add Array(27, "007-875-333", "SOUTHERN ONGVILLE CORP", "CANDELARIA, QUEZON", "Supplies")
            records.Add Array(28, "905-035-955", "N MART GENERAL MERCHANDIZE", "BAMBANG, NUEVA ECIJA", "Supplies")
            records.Add Array(29, "241-838-060", "FACIAL CARE CENTRE", "ALABANG, MUNTINLUPA", "Supplies")
            records.Add Array(30, "009-149--884", "LANDMARK SUPERMARKET", "ALABANG, MUNTINLUPA", "Supplies")
            records.Add Array(31, "240-940-697", "INENGS", "BGC, TAGUIG", "Representation")
            records.Add Array(32, "009-271-384", "OTADER DENTAL AND MEDICAL", "MALATE, MANILA", "Clinic Materials and Supplies")
            records.Add Array(33, "000-340-727", "ALPHADENT CORPORATION", "SAN ANTONIO, QUEZON", "Clinic Materials and Supplies")
            records.Add Array(34, "006-614-702", "PROS APAC CORP", "MALATE, MANILA", "Clinic Materials and Supplies")
            records.Add Array(35, "003-935-334", "ST PATRICK MEDICAL SYSTEMS INC", "SHAW BLVD., MANDALUYONG", "Clinic Materials and Supplies")
            records.Add Array(36, "216-037-865", "DENTAL DOMAIN CORP", "CULIAT, QUEZON CITY", "Clinic Materials and Supplies")
            records.Add Array(37, "226-527-915", "METRO MARKET MARKET", "BGC, TAGUIG", "Supplies")
            records.Add Array(38, "010-176-614", "GREEN DENTAL LAB", "BGC, TAGUIG", "Clinic Materials and Supplies")
            records.Add Array(39, "440-275-225", "AAHA GAS RETAILING", "PAMPLONA UNO, LAS PIÑAS", "Fuel and Oil")
            records.Add Array(40, "010-119-810", "RJD MERRY HARVEST CORP", "SILANG, CAVITE", "Supplies")
            records.Add Array(41, "000-522-680", "ALBA INTERNATIONAL INC", "ALABANG, MUNTINLUPA", "Fuel and Oil")
            records.Add Array(42, "007-551-518", "NUTRITION FOR LIFE", "ALABANG, MUNTINLUPA", "Representation")
            records.Add Array(43, "008-664-864", "GRANDWEN GAS CORP", "BACLARAN, PASAY", "Fuel and Oil")
            records.Add Array(44, "009-195-748", "BIG SKY NATION INC", "183 PASAY", "Fuel and Oil")
            records.Add Array(45, "241-437-266", "YELOW CAB PIZZA", "STA. ROSA, LAGUNA", "Representation")
            records.Add Array(46, "636-985956", "WILSON MOA", "MOA, PASAY", "Supplies")
            records.Add Array(47, "206-114-170", "PETRON EXPRESS CENTER", "BALAGTAS, BULACAN", "Fuel and Oil")
            records.Add Array(48, "005-148-299", "ERNA TRADINGCORP", "TALON 4, LAS PIÑAS", "Repairs and Maintenance")
            records.Add Array(49, "008-367-582", "JAVELLANA ADVANCED AESTHETICS CLINIC", "ALABANG, MUNTINLUPA", "Clinic Materials and Supplies")
            records.Add Array(50, "226-784-068-000", "MDI Group Holdings Inc", "BGC, TAGUIG CITY", "Rental")
            records.Add Array(51, "000-405-340", "ROBINSONS", "TALON UNO, LAS PIÑAS", "Supplies")
            records.Add Array(52, "007-189-834", "KENNY ROGERS", "ALABANG, MUNTINLUPA", "Representation")
            records.Add Array(53, "001-096-140", "ALMON MARINA INC", "ALABANG, MUNTINLUPA", "Representation")
            records.Add Array(54, "236-414-781", "YOKI FARM", "PALOCPOC, CAVITE", "Miscellaneous")
            records.Add Array(55, "203-301-457", "CINNABON", "ALABANG, MUNTINLUPA", "Representation")
            records.Add Array(56, "000-388-474", "MERCURY", "TALON 5, LAS PIÑAS", "Clinic Materials and Supplies")
            records.Add Array(57, "000-123-826", "SOUTH SUPERMARKET", "FILINVEST, MUNTINLUPA", "Supplies")
            records.Add Array(58, "241-498-356", "HONEYBON", "ALABANG, MUNTINLUPA", "Representation")
            records.Add Array(59, "000-388-771", "JOLLIBEE", "ALABANG, MUNTINLUPA", "Representation")
            records.Add Array(60, "610-884-579", "T RIPLE MLD CORP", "ALABANG, MUNTINLUPA", "Transportation and Travel")
            records.Add Array(61, "607-237-567", "RUSTIC PIZZA PASTA RESTAURANT INC", "ALABANG, MUNTINLUPA", "Representation")
            records.Add Array(62, "005-148-299", "ERNA TRADING CORP", "TALON 4, LAS PIÑAS", "Clinic Materials and Supplies")
            records.Add Array(63, "010-892-033", "ELEVATE PILATES COMPANY LIMITED INC", "BGC, TAGUIG", "Clinic Materials and Supplies")
            records.Add Array(64, "200-005-881", "ARISTOCRAT BAKESHOP", "ALMANZA UNO, LAS PIÑAS", "Representation")
            records.Add Array(65, "000-060-696", "RUSTANS", "ALABANG, MUNTINLUPA", "Representation")
            records.Add Array(66, "909-079-780", "INGUSAN GASOLINE STATION", "TALON UNO, LAS PIÑAS", "Fuel and Oil")
            records.Add Array(67, "000-299-299", "ABACUS BOOK AND CARD CORP", "POBLACION, MAKATI", "Supplies")
            records.Add Array(68, "003-058-789", "SM PRIME HOLDINGS INC", "ALMANZA UNO, LAS PIÑAS", "Supplies")
            records.Add Array(69, "183-378-131", "MG PETRON CORDON SERVICE STATION", "CORDON, ISABELA", "Fuel and Oil")
            records.Add Array(70, "630-167-414", "NAIL GARDEN", "76 PASAY", "Miscellaneous")
            records.Add Array(71, "230-577-870", "LALONG FUEL STATION", "SAN PABLO, LAGUNA", "Fuel and Oil")
            records.Add Array(72, "008-909-992", "LANDERS", "TAMBO, PARAÑAQUE", "Supplies")
            records.Add Array(73, "207-961-175", "FESTIVAL MALL", "ALABANG, MUNTINLUPA", "Supplies")
            records.Add Array(74, "010-262-979", "LIN ET VIE PELLE MERCHANDISING", "ALABANG, MUNTINLUPA", "Miscellaneous")
            records.Add Array(75, "005-695-758", "UNICITY", "BGC, TAGUIG", "Miscellaneous")
            records.Add Array(76, "010-176-614", "GREEN DELTA LAB PHILIPPINES INC", "BGC, TAGUIG", "Clinic Materials and Supplies")
            records.Add Array(77, "260-019-682", "BURDA MEDICAL APPAREL", "SANTA CRUZ, MANILA", "Clinic Materials and Supplies")
            records.Add Array(78, "278-353-148", "D AND L MEDICAL APPAREL", "SANTA CRUZ, MANILA", "Clinic Materials and Supplies")
            records.Add Array(79, "778-866-700", "GREEN WAVE COMPUTER SYSTEM", "SANTA CRUZ, MANILA", "Repairs and Maintenance")
            records.Add Array(80, "010-559-959", "FRONTIER DENTAL PRODUCTS CORP", "SAN ANTONIO, MAKATI", "Clinic Materials and Supplies")
            records.Add Array(81, "000-324-286", "GOLDEN PEAK SALES CORP", "SAN NICOLAS, MANILA", "Miscellaneous")
            records.Add Array(82, "225-527-914", "METRO RETAIL STORES GROUP INC", "BGC, TAGUIG", "Supplies")
            records.Add Array(83, "208-544-458", "MKYR TRANSPORT SERVICES", "BATASAN HILLS, QUEZON CITY", "Transportation and Travel")
            records.Add Array(84, "006-614-702", "PROS-APAC CORP", "MALATE, MANILA", "Clinic Materials and Supplies")
            records.Add Array(85, "201-619-152", "VILLACORTA DENTAL LABORATY INC", "SANTA ROSA, LAGUNA", "Clinic Materials and Supplies")
            records.Add Array(87, "120-568-073", "7-ELEVEN", "BGC, TAGUIG", "Miscellaneous")
            records.Add Array(88, "000-388-474", "MERCURY DRUG", "BGC, TAGUIG", "Supplies")
            records.Add Array(89, "233-251-708", "FEDERAL BRENT RETAIL INC", "076 PASAY", "Supplies")
            records.Add Array(90, "214-706-591", "WATSONS", "BGC, TAGUIG", "Miscellaneous")
            records.Add Array(91, "226-784-068", "MDI Group Holdings Inc 2", "BGC, TAGUIG", "Association Dues")
            records.Add Array(92, "TIN NOT FOUND", "Rosalina Ignacio", "—", "Clinician's Fee-Corporators")
            records.Add Array(93, "TIN NOT FOUND", "Daisy Cornejo", "—", "Clinician's Fee-Corporators")
            records.Add Array(94, "TIN NOT FOUND", "Janina Tayag", "—", "Clinician's Fee-Corporators")
            records.Add Array(95, "TIN NOT FOUND", "Jasper Tago", "—", "Clinician's Fee-Corporators")
            records.Add Array(96, "TIN NOT FOUND", "Ana Ramona Locsin", "—", "Clinician's Fee-Visiting")
            records.Add Array(97, "TIN NOT FOUND", "Virgilio Malenab Jr", "—", "Clinician's Fee-Visiting")
            records.Add Array(98, "TIN NOT FOUND", "SBM Consulting Inc", "—", "Professional Fee")
            records.Add Array(99, "TIN NOT FOUND", "Christian Anthony Ermita", "—", "Clinician's Fee-Visiting")
            records.Add Array(100, "TIN NOT FOUND", "Vergilyn Montales", "—", "Clinician's Fee-Visiting")

    ' Detect the placeholder address literal used by this import (non-printing/non-alphanumeric).
    placeholderAddr = ""
    For i = 1 To records.Count
        recValues = records(i)
        If Len(recValues(4)) >= 2 Then
            If Left(recValues(4), 1) = """" And Right(recValues(4), 1) = """" Then
                inner = Mid(recValues(4), 2, Len(recValues(4)) - 2)
                If inner >= ChrW(128) Or inner < Chr(32) Then
                    placeholderAddr = inner
                    Exit For
                End If
            End If
        End If
    Next i

    ' Validate duplicates in memory before any sheet write.
    duplicateFound = False
    For i = 1 To records.Count
        recValues = records(i)
        If Left(recValues(2), 1) = """" Then
            name = Mid(recValues(2), 2, Len(recValues(2)) - 2)
        Else
            name = recValues(2)
        End If
        If Left(recValues(3), 1) = """" Then
            address = Mid(recValues(3), 2, Len(recValues(3)) - 2)
        Else
            address = recValues(3)
        End If
        If name <> "" And address <> "" And address <> placeholderAddr Then
            For j = i + 1 To records.Count
                recValues = records(j)
                If Left(recValues(2), 1) = """" Then
                    name2 = Mid(recValues(2), 2, Len(recValues(2)) - 2)
                Else
                    name2 = recValues(2)
                End If
                If Left(recValues(3), 1) = """" Then
                    address2 = Mid(recValues(3), 2, Len(recValues(3)) - 2)
                Else
                    address2 = recValues(3)
                End If
                If name = name2 And address = address2 Then
                    MsgBox "DUPLICATE FOUND:" & vbCrLf & _
                           "Row " & i & ": " & name & " | " & address & vbCrLf & _
                           "Row " & j & ": " & name2 & " | " & address2 & vbCrLf & vbCrLf & _
                           "Same Name + Address is NOT allowed!", vbCritical
                    duplicateFound = True
                    Exit Sub
                End If
            Next j
        End If
    Next i

    ' Clear old data (keep headers)
    ws.Rows("2:" & ws.Rows.count).Clear
    rowNum = 2
    For i = 1 To records.Count
        recValues = records(i)
        ws.Cells(rowNum, "B").Value = recValues(0)
        ws.Cells(rowNum, "C").Value = recValues(1)
        ws.Cells(rowNum, "D").Value = recValues(2)
        ws.Cells(rowNum, "E").Value = recValues(3)
        ws.Cells(rowNum, "F").Value = recValues(4)
        rowNum = rowNum + 1
    Next i

    If Not duplicateFound Then
        MsgBox "Import complete! " & records.Count & " suppliers imported." & vbCrLf & _
               "All suppliers have unique Name + Address combinations.", vbInformation
    End If

End Sub
