Attribute VB_Name = "ImportSuppliersFromJSON"
Option Explicit

Sub ImportSuppliersFromJSON()
    Dim ws As Worksheet
    Dim rowNum As Long
    Dim i As Long
    Dim lastRow As Long
    
    Set ws = ThisWorkbook.Sheets("SUPPLIERS DATA")
    
    ' Clear old data (keep headers)
    ws.Rows("2:" & ws.Rows.count).Clear
    
    rowNum = 2
    
    ' ============================================================
    ' COLUMNS: B=NO, C=TIN#, D=PARTICULARS, E=ADDRESS, F=CHART OF ACCOUNT
    ' ============================================================
    
    ' 1
    ws.Cells(rowNum, "B").Value = 1
    ws.Cells(rowNum, "C").Value = "008-888-659"
    ws.Cells(rowNum, "D").Value = "VPHARMA HEALTH AND WELNESS INC"
    ws.Cells(rowNum, "E").Value = "ALABANG, MUNTINLUPA"
    ws.Cells(rowNum, "F").Value = "Clinic Materials and Supplies"
    rowNum = rowNum + 1
    
    ' 2
    ws.Cells(rowNum, "B").Value = 2
    ws.Cells(rowNum, "C").Value = "246-099-058"
    ws.Cells(rowNum, "D").Value = "TRAVELLERS INTL HOTEL INC"
    ws.Cells(rowNum, "E").Value = "NEWPORT, PASAY"
    ws.Cells(rowNum, "F").Value = "Transportation and Travel"
    rowNum = rowNum + 1
    
    ' 3
    ws.Cells(rowNum, "B").Value = 3
    ws.Cells(rowNum, "C").Value = "008-120-744"
    ws.Cells(rowNum, "D").Value = "RUSTAN"
    ws.Cells(rowNum, "E").Value = "SAN LORENZO, MAKATI"
    ws.Cells(rowNum, "F").Value = "Representation"
    rowNum = rowNum + 1
    
    ' 4
    ws.Cells(rowNum, "B").Value = 4
    ws.Cells(rowNum, "C").Value = "008-625-341"
    ws.Cells(rowNum, "D").Value = "THE MESSHALL"
    ws.Cells(rowNum, "E").Value = "MAGALLANES, MAKATI"
    ws.Cells(rowNum, "F").Value = "Miscellaneous"
    rowNum = rowNum + 1
    
    ' 5
    ws.Cells(rowNum, "B").Value = 5
    ws.Cells(rowNum, "C").Value = "252-126-611"
    ws.Cells(rowNum, "D").Value = "EUROFUEL INC"
    ws.Cells(rowNum, "E").Value = "183 PASAY"
    ws.Cells(rowNum, "F").Value = "Fuel and Oil"
    rowNum = rowNum + 1
    
    ' 6
    ws.Cells(rowNum, "B").Value = 6
    ws.Cells(rowNum, "C").Value = "008-909-992"
    ws.Cells(rowNum, "D").Value = "SOUTHEASTASIA RETAIL INC"
    ws.Cells(rowNum, "E").Value = "BGC, TAGUIG"
    ws.Cells(rowNum, "F").Value = "Supplies"
    rowNum = rowNum + 1
    
    ' 7
    ws.Cells(rowNum, "B").Value = 7
    ws.Cells(rowNum, "C").Value = "000-388-771"
    ws.Cells(rowNum, "D").Value = "JOLLIBEE 39TH"
    ws.Cells(rowNum, "E").Value = "BGC, TAGUIG"
    ws.Cells(rowNum, "F").Value = "Representation"
    rowNum = rowNum + 1
    
    ' 8
    ws.Cells(rowNum, "B").Value = 8
    ws.Cells(rowNum, "C").Value = "207-961-175"
    ws.Cells(rowNum, "D").Value = "SAVEMORE"
    ws.Cells(rowNum, "E").Value = "ALABANG, MUNTINLUPA"
    ws.Cells(rowNum, "F").Value = "Supplies"
    rowNum = rowNum + 1
    
    ' 9
    ws.Cells(rowNum, "B").Value = 9
    ws.Cells(rowNum, "C").Value = "001-799-123"
    ws.Cells(rowNum, "D").Value = "UNIGLOBE TRAVELWARE INC"
    ws.Cells(rowNum, "E").Value = "ALABANG, MUNTINLUPA"
    ws.Cells(rowNum, "F").Value = "Transportation and Travel"
    rowNum = rowNum + 1
    
    ' 10
    ws.Cells(rowNum, "B").Value = 10
    ws.Cells(rowNum, "C").Value = "148-855-506"
    ws.Cells(rowNum, "D").Value = "MINDWERKS"
    ws.Cells(rowNum, "E").Value = "ALABANG, MUNTINLUPA"
    ws.Cells(rowNum, "F").Value = "Supplies"
    rowNum = rowNum + 1
    
    ' 11
    ws.Cells(rowNum, "B").Value = 11
    ws.Cells(rowNum, "C").Value = "007-851-409"
    ws.Cells(rowNum, "D").Value = "PRIMERA CLASE GAS STATION"
    ws.Cells(rowNum, "E").Value = "ALABANG, MUNTINLUPA"
    ws.Cells(rowNum, "F").Value = "Fuel and Oil"
    rowNum = rowNum + 1
    
    ' 12
    ws.Cells(rowNum, "B").Value = 12
    ws.Cells(rowNum, "C").Value = "200-035-311"
    ws.Cells(rowNum, "D").Value = "ACE HARDWARE"
    ws.Cells(rowNum, "E").Value = "ALMANZA UNO, LAS PIÑAS"
    ws.Cells(rowNum, "F").Value = "Repairs and Maintenance"
    rowNum = rowNum + 1
    
    ' 13
    ws.Cells(rowNum, "B").Value = 13
    ws.Cells(rowNum, "C").Value = "000-388-771"
    ws.Cells(rowNum, "D").Value = "JOLLIBEE BLUEBAY WALK"
    ws.Cells(rowNum, "E").Value = "76 PASAY"
    ws.Cells(rowNum, "F").Value = "Representation"
    rowNum = rowNum + 1
    
    ' 14
    ws.Cells(rowNum, "B").Value = 14
    ws.Cells(rowNum, "C").Value = "246-969-491"
    ws.Cells(rowNum, "D").Value = "S AND R"
    ws.Cells(rowNum, "E").Value = "BGC, TAGUIG"
    ws.Cells(rowNum, "F").Value = "Representation"
    rowNum = rowNum + 1
    
    ' 15
    ws.Cells(rowNum, "B").Value = 15
    ws.Cells(rowNum, "C").Value = "648-700-952"
    ws.Cells(rowNum, "D").Value = "GOOD MORNING FOOD"
    ws.Cells(rowNum, "E").Value = "BGC, TAGUIG"
    ws.Cells(rowNum, "F").Value = "Representation"
    rowNum = rowNum + 1
    
    ' 16
    ws.Cells(rowNum, "B").Value = 16
    ws.Cells(rowNum, "C").Value = "008-087-158"
    ws.Cells(rowNum, "D").Value = "PNC GARBROS FOODS INC"
    ws.Cells(rowNum, "E").Value = "ALABANG, MUNTINLUPA"
    ws.Cells(rowNum, "F").Value = "Representation"
    rowNum = rowNum + 1
    
    ' 17
    ws.Cells(rowNum, "B").Value = 17
    ws.Cells(rowNum, "C").Value = "208-006-656"
    ws.Cells(rowNum, "D").Value = "SM STORE"
    ws.Cells(rowNum, "E").Value = "ORANBO, PASIG"
    ws.Cells(rowNum, "F").Value = "Supplies"
    rowNum = rowNum + 1
    
    ' 18
    ws.Cells(rowNum, "B").Value = 18
    ws.Cells(rowNum, "C").Value = "214-706-591"
    ws.Cells(rowNum, "D").Value = "WATSONS"
    ws.Cells(rowNum, "E").Value = "ALABANG, MUNTINLUPA"
    ws.Cells(rowNum, "F").Value = "Supplies"
    rowNum = rowNum + 1
    
    ' 19
    ws.Cells(rowNum, "B").Value = 19
    ws.Cells(rowNum, "C").Value = "608-395-184"
    ws.Cells(rowNum, "D").Value = "GIANETO PIZZERIA"
    ws.Cells(rowNum, "E").Value = "ORANBO, PASIG"
    ws.Cells(rowNum, "F").Value = "Representation"
    rowNum = rowNum + 1
    
    ' 20
    ws.Cells(rowNum, "B").Value = 20
    ws.Cells(rowNum, "C").Value = "010-624-401"
    ws.Cells(rowNum, "D").Value = "MIT SUKOSHI FRESH"
    ws.Cells(rowNum, "E").Value = "BGC, TAGUIG"
    ws.Cells(rowNum, "F").Value = "Representation"
    rowNum = rowNum + 1
    
    ' 21
    ws.Cells(rowNum, "B").Value = 21
    ws.Cells(rowNum, "C").Value = "655-591-162"
    ws.Cells(rowNum, "D").Value = "MINISO"
    ws.Cells(rowNum, "E").Value = "ALABANG, MUNTINLUPA"
    ws.Cells(rowNum, "F").Value = "Supplies"
    rowNum = rowNum + 1
    
    ' 22
    ws.Cells(rowNum, "B").Value = 22
    ws.Cells(rowNum, "C").Value = "000-405-340"
    ws.Cells(rowNum, "D").Value = "ROBINSONS"
    ws.Cells(rowNum, "E").Value = "ALABANG, MUNTINLUPA"
    ws.Cells(rowNum, "F").Value = "Supplies"
    rowNum = rowNum + 1
    
    ' 23
    ws.Cells(rowNum, "B").Value = 23
    ws.Cells(rowNum, "C").Value = "000-123-826"
    ws.Cells(rowNum, "D").Value = "GRAND UNION SUPERMARKET INC"
    ws.Cells(rowNum, "E").Value = "FILINVEST, MUNTINLUPA"
    ws.Cells(rowNum, "F").Value = "Supplies"
    rowNum = rowNum + 1
    
    ' 24
    ws.Cells(rowNum, "B").Value = 24
    ws.Cells(rowNum, "C").Value = "008-188-059"
    ws.Cells(rowNum, "D").Value = "KNOXPORT INC"
    ws.Cells(rowNum, "E").Value = "BGC, TAGUIG"
    ws.Cells(rowNum, "F").Value = "Miscellaneous"
    rowNum = rowNum + 1
    
    ' 25
    ws.Cells(rowNum, "B").Value = 25
    ws.Cells(rowNum, "C").Value = "009-754-702"
    ws.Cells(rowNum, "D").Value = "GOOD SMALL SHEEP INC"
    ws.Cells(rowNum, "E").Value = "DAANG BAKAL, MANDALUYONG"
    ws.Cells(rowNum, "F").Value = "Miscellaneous"
    rowNum = rowNum + 1
    
    ' 26
    ws.Cells(rowNum, "B").Value = 26
    ws.Cells(rowNum, "C").Value = "163-955-615"
    ws.Cells(rowNum, "D").Value = "725 TAGTEAM FUEL STATION"
    ws.Cells(rowNum, "E").Value = "ALABANG, MUNTINLUPA"
    ws.Cells(rowNum, "F").Value = "Fuel and Oil"
    rowNum = rowNum + 1
    
    ' 27
    ws.Cells(rowNum, "B").Value = 27
    ws.Cells(rowNum, "C").Value = "007-875-333"
    ws.Cells(rowNum, "D").Value = "SOUTHERN ONGVILLE CORP"
    ws.Cells(rowNum, "E").Value = "CANDELARIA, QUEZON"
    ws.Cells(rowNum, "F").Value = "Supplies"
    rowNum = rowNum + 1
    
    ' 28
    ws.Cells(rowNum, "B").Value = 28
    ws.Cells(rowNum, "C").Value = "905-035-955"
    ws.Cells(rowNum, "D").Value = "N MART GENERAL MERCHANDIZE"
    ws.Cells(rowNum, "E").Value = "BAMBANG, NUEVA ECIJA"
    ws.Cells(rowNum, "F").Value = "Supplies"
    rowNum = rowNum + 1
    
    ' 29
    ws.Cells(rowNum, "B").Value = 29
    ws.Cells(rowNum, "C").Value = "241-838-060"
    ws.Cells(rowNum, "D").Value = "FACIAL CARE CENTRE"
    ws.Cells(rowNum, "E").Value = "ALABANG, MUNTINLUPA"
    ws.Cells(rowNum, "F").Value = "Supplies"
    rowNum = rowNum + 1
    
    ' 30
    ws.Cells(rowNum, "B").Value = 30
    ws.Cells(rowNum, "C").Value = "009-149--884"
    ws.Cells(rowNum, "D").Value = "LANDMARK SUPERMARKET"
    ws.Cells(rowNum, "E").Value = "ALABANG, MUNTINLUPA"
    ws.Cells(rowNum, "F").Value = "Supplies"
    rowNum = rowNum + 1
    
    ' 31
    ws.Cells(rowNum, "B").Value = 31
    ws.Cells(rowNum, "C").Value = "240-940-697"
    ws.Cells(rowNum, "D").Value = "INENGS"
    ws.Cells(rowNum, "E").Value = "BGC, TAGUIG"
    ws.Cells(rowNum, "F").Value = "Representation"
    rowNum = rowNum + 1
    
    ' 32
    ws.Cells(rowNum, "B").Value = 32
    ws.Cells(rowNum, "C").Value = "009-271-384"
    ws.Cells(rowNum, "D").Value = "OTADER DENTAL AND MEDICAL"
    ws.Cells(rowNum, "E").Value = "MALATE, MANILA"
    ws.Cells(rowNum, "F").Value = "Clinic Materials and Supplies"
    rowNum = rowNum + 1
    
    ' 33
    ws.Cells(rowNum, "B").Value = 33
    ws.Cells(rowNum, "C").Value = "000-340-727"
    ws.Cells(rowNum, "D").Value = "ALPHADENT CORPORATION"
    ws.Cells(rowNum, "E").Value = "SAN ANTONIO, QUEZON"
    ws.Cells(rowNum, "F").Value = "Clinic Materials and Supplies"
    rowNum = rowNum + 1
    
    ' 34
    ws.Cells(rowNum, "B").Value = 34
    ws.Cells(rowNum, "C").Value = "006-614-702"
    ws.Cells(rowNum, "D").Value = "PROS APAC CORP"
    ws.Cells(rowNum, "E").Value = "MALATE, MANILA"
    ws.Cells(rowNum, "F").Value = "Clinic Materials and Supplies"
    rowNum = rowNum + 1
    
    ' 35
    ws.Cells(rowNum, "B").Value = 35
    ws.Cells(rowNum, "C").Value = "003-935-334"
    ws.Cells(rowNum, "D").Value = "ST PATRICK MEDICAL SYSTEMS INC"
    ws.Cells(rowNum, "E").Value = "SHAW BLVD., MANDALUYONG"
    ws.Cells(rowNum, "F").Value = "Clinic Materials and Supplies"
    rowNum = rowNum + 1
    
    ' 36
    ws.Cells(rowNum, "B").Value = 36
    ws.Cells(rowNum, "C").Value = "216-037-865"
    ws.Cells(rowNum, "D").Value = "DENTAL DOMAIN CORP"
    ws.Cells(rowNum, "E").Value = "CULIAT, QUEZON CITY"
    ws.Cells(rowNum, "F").Value = "Clinic Materials and Supplies"
    rowNum = rowNum + 1
    
    ' 37
    ws.Cells(rowNum, "B").Value = 37
    ws.Cells(rowNum, "C").Value = "226-527-915"
    ws.Cells(rowNum, "D").Value = "METRO MARKET MARKET"
    ws.Cells(rowNum, "E").Value = "BGC, TAGUIG"
    ws.Cells(rowNum, "F").Value = "Supplies"
    rowNum = rowNum + 1
    
    ' 38
    ws.Cells(rowNum, "B").Value = 38
    ws.Cells(rowNum, "C").Value = "010-176-614"
    ws.Cells(rowNum, "D").Value = "GREEN DENTAL LAB"
    ws.Cells(rowNum, "E").Value = "BGC, TAGUIG"
    ws.Cells(rowNum, "F").Value = "Clinic Materials and Supplies"
    rowNum = rowNum + 1
    
    ' 39
    ws.Cells(rowNum, "B").Value = 39
    ws.Cells(rowNum, "C").Value = "440-275-225"
    ws.Cells(rowNum, "D").Value = "AAHA GAS RETAILING"
    ws.Cells(rowNum, "E").Value = "PAMPLONA UNO, LAS PIÑAS"
    ws.Cells(rowNum, "F").Value = "Fuel and Oil"
    rowNum = rowNum + 1
    
    ' 40
    ws.Cells(rowNum, "B").Value = 40
    ws.Cells(rowNum, "C").Value = "010-119-810"
    ws.Cells(rowNum, "D").Value = "RJD MERRY HARVEST CORP"
    ws.Cells(rowNum, "E").Value = "SILANG, CAVITE"
    ws.Cells(rowNum, "F").Value = "Supplies"
    rowNum = rowNum + 1
    
    ' 41
    ws.Cells(rowNum, "B").Value = 41
    ws.Cells(rowNum, "C").Value = "000-522-680"
    ws.Cells(rowNum, "D").Value = "ALBA INTERNATIONAL INC"
    ws.Cells(rowNum, "E").Value = "ALABANG, MUNTINLUPA"
    ws.Cells(rowNum, "F").Value = "Fuel and Oil"
    rowNum = rowNum + 1
    
    ' 42
    ws.Cells(rowNum, "B").Value = 42
    ws.Cells(rowNum, "C").Value = "007-551-518"
    ws.Cells(rowNum, "D").Value = "NUTRITION FOR LIFE"
    ws.Cells(rowNum, "E").Value = "ALABANG, MUNTINLUPA"
    ws.Cells(rowNum, "F").Value = "Representation"
    rowNum = rowNum + 1
    
    ' 43
    ws.Cells(rowNum, "B").Value = 43
    ws.Cells(rowNum, "C").Value = "008-664-864"
    ws.Cells(rowNum, "D").Value = "GRANDWEN GAS CORP"
    ws.Cells(rowNum, "E").Value = "BACLARAN, PASAY"
    ws.Cells(rowNum, "F").Value = "Fuel and Oil"
    rowNum = rowNum + 1
    
    ' 44
    ws.Cells(rowNum, "B").Value = 44
    ws.Cells(rowNum, "C").Value = "009-195-748"
    ws.Cells(rowNum, "D").Value = "BIG SKY NATION INC"
    ws.Cells(rowNum, "E").Value = "183 PASAY"
    ws.Cells(rowNum, "F").Value = "Fuel and Oil"
    rowNum = rowNum + 1
    
    ' 45
    ws.Cells(rowNum, "B").Value = 45
    ws.Cells(rowNum, "C").Value = "241-437-266"
    ws.Cells(rowNum, "D").Value = "YELOW CAB PIZZA"
    ws.Cells(rowNum, "E").Value = "STA. ROSA, LAGUNA"
    ws.Cells(rowNum, "F").Value = "Representation"
    rowNum = rowNum + 1
    
    ' 46
    ws.Cells(rowNum, "B").Value = 46
    ws.Cells(rowNum, "C").Value = "636-985956"
    ws.Cells(rowNum, "D").Value = "WILSON MOA"
    ws.Cells(rowNum, "E").Value = "MOA, PASAY"
    ws.Cells(rowNum, "F").Value = "Supplies"
    rowNum = rowNum + 1
    
    ' 47
    ws.Cells(rowNum, "B").Value = 47
    ws.Cells(rowNum, "C").Value = "206-114-170"
    ws.Cells(rowNum, "D").Value = "PETRON EXPRESS CENTER"
    ws.Cells(rowNum, "E").Value = "BALAGTAS, BULACAN"
    ws.Cells(rowNum, "F").Value = "Fuel and Oil"
    rowNum = rowNum + 1
    
    ' 48
    ws.Cells(rowNum, "B").Value = 48
    ws.Cells(rowNum, "C").Value = "005-148-299"
    ws.Cells(rowNum, "D").Value = "ERNA TRADINGCORP"
    ws.Cells(rowNum, "E").Value = "TALON 4, LAS PIÑAS"
    ws.Cells(rowNum, "F").Value = "Repairs and Maintenance"
    rowNum = rowNum + 1
    
    ' 49
    ws.Cells(rowNum, "B").Value = 49
    ws.Cells(rowNum, "C").Value = "008-367-582"
    ws.Cells(rowNum, "D").Value = "JAVELLANA ADVANCED AESTHETICS CLINIC"
    ws.Cells(rowNum, "E").Value = "ALABANG, MUNTINLUPA"
    ws.Cells(rowNum, "F").Value = "Clinic Materials and Supplies"
    rowNum = rowNum + 1
    
    ' 50
    ws.Cells(rowNum, "B").Value = 50
    ws.Cells(rowNum, "C").Value = "226-784-068-000"
    ws.Cells(rowNum, "D").Value = "MDI Group Holdings Inc"
    ws.Cells(rowNum, "E").Value = "BGC, TAGUIG CITY"
    ws.Cells(rowNum, "F").Value = "Rental"
    rowNum = rowNum + 1
    
    ' 51
    ws.Cells(rowNum, "B").Value = 51
    ws.Cells(rowNum, "C").Value = "000-405-340"
    ws.Cells(rowNum, "D").Value = "ROBINSONS"
    ws.Cells(rowNum, "E").Value = "TALON UNO, LAS PIÑAS"
    ws.Cells(rowNum, "F").Value = "Supplies"
    rowNum = rowNum + 1
    
    ' 52
    ws.Cells(rowNum, "B").Value = 52
    ws.Cells(rowNum, "C").Value = "007-189-834"
    ws.Cells(rowNum, "D").Value = "KENNY ROGERS"
    ws.Cells(rowNum, "E").Value = "ALABANG, MUNTINLUPA"
    ws.Cells(rowNum, "F").Value = "Representation"
    rowNum = rowNum + 1
    
    ' 53
    ws.Cells(rowNum, "B").Value = 53
    ws.Cells(rowNum, "C").Value = "001-096-140"
    ws.Cells(rowNum, "D").Value = "ALMON MARINA INC"
    ws.Cells(rowNum, "E").Value = "ALABANG, MUNTINLUPA"
    ws.Cells(rowNum, "F").Value = "Representation"
    rowNum = rowNum + 1
    
    ' 54
    ws.Cells(rowNum, "B").Value = 54
    ws.Cells(rowNum, "C").Value = "236-414-781"
    ws.Cells(rowNum, "D").Value = "YOKI FARM"
    ws.Cells(rowNum, "E").Value = "PALOCPOC, CAVITE"
    ws.Cells(rowNum, "F").Value = "Miscellaneous"
    rowNum = rowNum + 1
    
    ' 55
    ws.Cells(rowNum, "B").Value = 55
    ws.Cells(rowNum, "C").Value = "203-301-457"
    ws.Cells(rowNum, "D").Value = "CINNABON"
    ws.Cells(rowNum, "E").Value = "ALABANG, MUNTINLUPA"
    ws.Cells(rowNum, "F").Value = "Representation"
    rowNum = rowNum + 1
    
    ' 56
    ws.Cells(rowNum, "B").Value = 56
    ws.Cells(rowNum, "C").Value = "000-388-474"
    ws.Cells(rowNum, "D").Value = "MERCURY"
    ws.Cells(rowNum, "E").Value = "TALON 5, LAS PIÑAS"
    ws.Cells(rowNum, "F").Value = "Clinic Materials and Supplies"
    rowNum = rowNum + 1
    
    ' 57
    ws.Cells(rowNum, "B").Value = 57
    ws.Cells(rowNum, "C").Value = "000-123-826"
    ws.Cells(rowNum, "D").Value = "SOUTH SUPERMARKET"
    ws.Cells(rowNum, "E").Value = "FILINVEST, MUNTINLUPA"
    ws.Cells(rowNum, "F").Value = "Supplies"
    rowNum = rowNum + 1
    
    ' 58
    ws.Cells(rowNum, "B").Value = 58
    ws.Cells(rowNum, "C").Value = "241-498-356"
    ws.Cells(rowNum, "D").Value = "HONEYBON"
    ws.Cells(rowNum, "E").Value = "ALABANG, MUNTINLUPA"
    ws.Cells(rowNum, "F").Value = "Representation"
    rowNum = rowNum + 1
    
    ' 59
    ws.Cells(rowNum, "B").Value = 59
    ws.Cells(rowNum, "C").Value = "000-388-771"
    ws.Cells(rowNum, "D").Value = "JOLLIBEE"
    ws.Cells(rowNum, "E").Value = "ALABANG, MUNTINLUPA"
    ws.Cells(rowNum, "F").Value = "Representation"
    rowNum = rowNum + 1
    
    ' 60
    ws.Cells(rowNum, "B").Value = 60
    ws.Cells(rowNum, "C").Value = "610-884-579"
    ws.Cells(rowNum, "D").Value = "T RIPLE MLD CORP"
    ws.Cells(rowNum, "E").Value = "ALABANG, MUNTINLUPA"
    ws.Cells(rowNum, "F").Value = "Transportation and Travel"
    rowNum = rowNum + 1
    
    ' 61
    ws.Cells(rowNum, "B").Value = 61
    ws.Cells(rowNum, "C").Value = "607-237-567"
    ws.Cells(rowNum, "D").Value = "RUSTIC PIZZA PASTA RESTAURANT INC"
    ws.Cells(rowNum, "E").Value = "ALABANG, MUNTINLUPA"
    ws.Cells(rowNum, "F").Value = "Representation"
    rowNum = rowNum + 1
    
    ' 62
    ws.Cells(rowNum, "B").Value = 62
    ws.Cells(rowNum, "C").Value = "005-148-299"
    ws.Cells(rowNum, "D").Value = "ERNA TRADING CORP"
    ws.Cells(rowNum, "E").Value = "TALON 4, LAS PIÑAS"
    ws.Cells(rowNum, "F").Value = "Clinic Materials and Supplies"
    rowNum = rowNum + 1
    
    ' 63
    ws.Cells(rowNum, "B").Value = 63
    ws.Cells(rowNum, "C").Value = "010-892-033"
    ws.Cells(rowNum, "D").Value = "ELEVATE PILATES COMPANY LIMITED INC"
    ws.Cells(rowNum, "E").Value = "BGC, TAGUIG"
    ws.Cells(rowNum, "F").Value = "Clinic Materials and Supplies"
    rowNum = rowNum + 1
    
    ' 64
    ws.Cells(rowNum, "B").Value = 64
    ws.Cells(rowNum, "C").Value = "200-005-881"
    ws.Cells(rowNum, "D").Value = "ARISTOCRAT BAKESHOP"
    ws.Cells(rowNum, "E").Value = "ALMANZA UNO, LAS PIÑAS"
    ws.Cells(rowNum, "F").Value = "Representation"
    rowNum = rowNum + 1
    
    ' 65
    ws.Cells(rowNum, "B").Value = 65
    ws.Cells(rowNum, "C").Value = "000-060-696"
    ws.Cells(rowNum, "D").Value = "RUSTANS"
    ws.Cells(rowNum, "E").Value = "ALABANG, MUNTINLUPA"
    ws.Cells(rowNum, "F").Value = "Representation"
    rowNum = rowNum + 1
    
    ' 66
    ws.Cells(rowNum, "B").Value = 66
    ws.Cells(rowNum, "C").Value = "909-079-780"
    ws.Cells(rowNum, "D").Value = "INGUSAN GASOLINE STATION"
    ws.Cells(rowNum, "E").Value = "TALON UNO, LAS PIÑAS"
    ws.Cells(rowNum, "F").Value = "Fuel and Oil"
    rowNum = rowNum + 1
    
    ' 67
    ws.Cells(rowNum, "B").Value = 67
    ws.Cells(rowNum, "C").Value = "000-299-299"
    ws.Cells(rowNum, "D").Value = "ABACUS BOOK AND CARD CORP"
    ws.Cells(rowNum, "E").Value = "POBLACION, MAKATI"
    ws.Cells(rowNum, "F").Value = "Supplies"
    rowNum = rowNum + 1
    
    ' 68
    ws.Cells(rowNum, "B").Value = 68
    ws.Cells(rowNum, "C").Value = "003-058-789"
    ws.Cells(rowNum, "D").Value = "SM PRIME HOLDINGS INC"
    ws.Cells(rowNum, "E").Value = "ALMANZA UNO, LAS PIÑAS"
    ws.Cells(rowNum, "F").Value = "Supplies"
    rowNum = rowNum + 1
    
    ' 69
    ws.Cells(rowNum, "B").Value = 69
    ws.Cells(rowNum, "C").Value = "183-378-131"
    ws.Cells(rowNum, "D").Value = "MG PETRON CORDON SERVICE STATION"
    ws.Cells(rowNum, "E").Value = "CORDON, ISABELA"
    ws.Cells(rowNum, "F").Value = "Fuel and Oil"
    rowNum = rowNum + 1
    
    ' 70
    ws.Cells(rowNum, "B").Value = 70
    ws.Cells(rowNum, "C").Value = "630-167-414"
    ws.Cells(rowNum, "D").Value = "NAIL GARDEN"
    ws.Cells(rowNum, "E").Value = "76 PASAY"
    ws.Cells(rowNum, "F").Value = "Miscellaneous"
    rowNum = rowNum + 1
    
    ' 71
    ws.Cells(rowNum, "B").Value = 71
    ws.Cells(rowNum, "C").Value = "230-577-870"
    ws.Cells(rowNum, "D").Value = "LALONG FUEL STATION"
    ws.Cells(rowNum, "E").Value = "SAN PABLO, LAGUNA"
    ws.Cells(rowNum, "F").Value = "Fuel and Oil"
    rowNum = rowNum + 1
    
    ' 72
    ws.Cells(rowNum, "B").Value = 72
    ws.Cells(rowNum, "C").Value = "008-909-992"
    ws.Cells(rowNum, "D").Value = "LANDERS"
    ws.Cells(rowNum, "E").Value = "TAMBO, PARAÑAQUE"
    ws.Cells(rowNum, "F").Value = "Supplies"
    rowNum = rowNum + 1
    
    ' 73
    ws.Cells(rowNum, "B").Value = 73
    ws.Cells(rowNum, "C").Value = "207-961-175"
    ws.Cells(rowNum, "D").Value = "FESTIVAL MALL"
    ws.Cells(rowNum, "E").Value = "ALABANG, MUNTINLUPA"
    ws.Cells(rowNum, "F").Value = "Supplies"
    rowNum = rowNum + 1
    
    ' 74
    ws.Cells(rowNum, "B").Value = 74
    ws.Cells(rowNum, "C").Value = "010-262-979"
    ws.Cells(rowNum, "D").Value = "LIN ET VIE PELLE MERCHANDISING"
    ws.Cells(rowNum, "E").Value = "ALABANG, MUNTINLUPA"
    ws.Cells(rowNum, "F").Value = "Miscellaneous"
    rowNum = rowNum + 1
    
    ' 75
    ws.Cells(rowNum, "B").Value = 75
    ws.Cells(rowNum, "C").Value = "005-695-758"
    ws.Cells(rowNum, "D").Value = "UNICITY"
    ws.Cells(rowNum, "E").Value = "BGC, TAGUIG"
    ws.Cells(rowNum, "F").Value = "Miscellaneous"
    rowNum = rowNum + 1
    
    ' 76
    ws.Cells(rowNum, "B").Value = 76
    ws.Cells(rowNum, "C").Value = "010-176-614"
    ws.Cells(rowNum, "D").Value = "GREEN DELTA LAB PHILIPPINES INC"
    ws.Cells(rowNum, "E").Value = "BGC, TAGUIG"
    ws.Cells(rowNum, "F").Value = "Clinic Materials and Supplies"
    rowNum = rowNum + 1
    
    ' 77
    ws.Cells(rowNum, "B").Value = 77
    ws.Cells(rowNum, "C").Value = "260-019-682"
    ws.Cells(rowNum, "D").Value = "BURDA MEDICAL APPAREL"
    ws.Cells(rowNum, "E").Value = "SANTA CRUZ, MANILA"
    ws.Cells(rowNum, "F").Value = "Clinic Materials and Supplies"
    rowNum = rowNum + 1
    
    ' 78
    ws.Cells(rowNum, "B").Value = 78
    ws.Cells(rowNum, "C").Value = "278-353-148"
    ws.Cells(rowNum, "D").Value = "D AND L MEDICAL APPAREL"
    ws.Cells(rowNum, "E").Value = "SANTA CRUZ, MANILA"
    ws.Cells(rowNum, "F").Value = "Clinic Materials and Supplies"
    rowNum = rowNum + 1
    
    ' 79
    ws.Cells(rowNum, "B").Value = 79
    ws.Cells(rowNum, "C").Value = "778-866-700"
    ws.Cells(rowNum, "D").Value = "GREEN WAVE COMPUTER SYSTEM"
    ws.Cells(rowNum, "E").Value = "SANTA CRUZ, MANILA"
    ws.Cells(rowNum, "F").Value = "Repairs and Maintenance"
    rowNum = rowNum + 1
    
    ' 80
    ws.Cells(rowNum, "B").Value = 80
    ws.Cells(rowNum, "C").Value = "010-559-959"
    ws.Cells(rowNum, "D").Value = "FRONTIER DENTAL PRODUCTS CORP"
    ws.Cells(rowNum, "E").Value = "SAN ANTONIO, MAKATI"
    ws.Cells(rowNum, "F").Value = "Clinic Materials and Supplies"
    rowNum = rowNum + 1
    
    ' 81
    ws.Cells(rowNum, "B").Value = 81
    ws.Cells(rowNum, "C").Value = "000-324-286"
    ws.Cells(rowNum, "D").Value = "GOLDEN PEAK SALES CORP"
    ws.Cells(rowNum, "E").Value = "SAN NICOLAS, MANILA"
    ws.Cells(rowNum, "F").Value = "Miscellaneous"
    rowNum = rowNum + 1
    
    ' 82
    ws.Cells(rowNum, "B").Value = 82
    ws.Cells(rowNum, "C").Value = "225-527-914"
    ws.Cells(rowNum, "D").Value = "METRO RETAIL STORES GROUP INC"
    ws.Cells(rowNum, "E").Value = "BGC, TAGUIG"
    ws.Cells(rowNum, "F").Value = "Supplies"
    rowNum = rowNum + 1
    
    ' 83
    ws.Cells(rowNum, "B").Value = 83
    ws.Cells(rowNum, "C").Value = "208-544-458"
    ws.Cells(rowNum, "D").Value = "MKYR TRANSPORT SERVICES"
    ws.Cells(rowNum, "E").Value = "BATASAN HILLS, QUEZON CITY"
    ws.Cells(rowNum, "F").Value = "Transportation and Travel"
    rowNum = rowNum + 1
    
    ' 84
    ws.Cells(rowNum, "B").Value = 84
    ws.Cells(rowNum, "C").Value = "006-614-702"
    ws.Cells(rowNum, "D").Value = "PROS-APAC CORP"
    ws.Cells(rowNum, "E").Value = "MALATE, MANILA"
    ws.Cells(rowNum, "F").Value = "Clinic Materials and Supplies"
    rowNum = rowNum + 1
    
    ' 85
    ws.Cells(rowNum, "B").Value = 85
    ws.Cells(rowNum, "C").Value = "201-619-152"
    ws.Cells(rowNum, "D").Value = "VILLACORTA DENTAL LABORATY INC"
    ws.Cells(rowNum, "E").Value = "SANTA ROSA, LAGUNA"
    ws.Cells(rowNum, "F").Value = "Clinic Materials and Supplies"
    rowNum = rowNum + 1
    
    ' ROW 86 REMOVED - DUPLICATE
    
    ' 87
    ws.Cells(rowNum, "B").Value = 87
    ws.Cells(rowNum, "C").Value = "120-568-073"
    ws.Cells(rowNum, "D").Value = "7-ELEVEN"
    ws.Cells(rowNum, "E").Value = "BGC, TAGUIG"
    ws.Cells(rowNum, "F").Value = "Miscellaneous"
    rowNum = rowNum + 1
    
    ' 88
    ws.Cells(rowNum, "B").Value = 88
    ws.Cells(rowNum, "C").Value = "000-388-474"
    ws.Cells(rowNum, "D").Value = "MERCURY DRUG"
    ws.Cells(rowNum, "E").Value = "BGC, TAGUIG"
    ws.Cells(rowNum, "F").Value = "Supplies"
    rowNum = rowNum + 1
    
    ' 89
    ws.Cells(rowNum, "B").Value = 89
    ws.Cells(rowNum, "C").Value = "233-251-708"
    ws.Cells(rowNum, "D").Value = "FEDERAL BRENT RETAIL INC"
    ws.Cells(rowNum, "E").Value = "076 PASAY"
    ws.Cells(rowNum, "F").Value = "Supplies"
    rowNum = rowNum + 1
    
    ' 90
    ws.Cells(rowNum, "B").Value = 90
    ws.Cells(rowNum, "C").Value = "214-706-591"
    ws.Cells(rowNum, "D").Value = "WATSONS"
    ws.Cells(rowNum, "E").Value = "BGC, TAGUIG"
    ws.Cells(rowNum, "F").Value = "Miscellaneous"
    rowNum = rowNum + 1
    
    ' 91
    ws.Cells(rowNum, "B").Value = 91
    ws.Cells(rowNum, "C").Value = "226-784-068"
    ws.Cells(rowNum, "D").Value = "MDI Group Holdings Inc 2"
    ws.Cells(rowNum, "E").Value = "BGC, TAGUIG"
    ws.Cells(rowNum, "F").Value = "Association Dues"
    rowNum = rowNum + 1
    
    ' 92
    ws.Cells(rowNum, "B").Value = 92
    ws.Cells(rowNum, "C").Value = "TIN NOT FOUND"
    ws.Cells(rowNum, "D").Value = "Rosalina Ignacio"
    ws.Cells(rowNum, "E").Value = "—"
    ws.Cells(rowNum, "F").Value = "Clinician's Fee-Corporators"
    rowNum = rowNum + 1
    
    ' 93
    ws.Cells(rowNum, "B").Value = 93
    ws.Cells(rowNum, "C").Value = "TIN NOT FOUND"
    ws.Cells(rowNum, "D").Value = "Daisy Cornejo"
    ws.Cells(rowNum, "E").Value = "—"
    ws.Cells(rowNum, "F").Value = "Clinician's Fee-Corporators"
    rowNum = rowNum + 1
    
    ' 94
    ws.Cells(rowNum, "B").Value = 94
    ws.Cells(rowNum, "C").Value = "TIN NOT FOUND"
    ws.Cells(rowNum, "D").Value = "Janina Tayag"
    ws.Cells(rowNum, "E").Value = "—"
    ws.Cells(rowNum, "F").Value = "Clinician's Fee-Corporators"
    rowNum = rowNum + 1
    
    ' 95
    ws.Cells(rowNum, "B").Value = 95
    ws.Cells(rowNum, "C").Value = "TIN NOT FOUND"
    ws.Cells(rowNum, "D").Value = "Jasper Tago"
    ws.Cells(rowNum, "E").Value = "—"
    ws.Cells(rowNum, "F").Value = "Clinician's Fee-Corporators"
    rowNum = rowNum + 1
    
    ' 96
    ws.Cells(rowNum, "B").Value = 96
    ws.Cells(rowNum, "C").Value = "TIN NOT FOUND"
    ws.Cells(rowNum, "D").Value = "Ana Ramona Locsin"
    ws.Cells(rowNum, "E").Value = "—"
    ws.Cells(rowNum, "F").Value = "Clinician's Fee-Visiting"
    rowNum = rowNum + 1
    
    ' 97
    ws.Cells(rowNum, "B").Value = 97
    ws.Cells(rowNum, "C").Value = "TIN NOT FOUND"
    ws.Cells(rowNum, "D").Value = "Virgilio Malenab Jr"
    ws.Cells(rowNum, "E").Value = "—"
    ws.Cells(rowNum, "F").Value = "Clinician's Fee-Visiting"
    rowNum = rowNum + 1
    
    ' 98
    ws.Cells(rowNum, "B").Value = 98
    ws.Cells(rowNum, "C").Value = "TIN NOT FOUND"
    ws.Cells(rowNum, "D").Value = "SBM Consulting Inc"
    ws.Cells(rowNum, "E").Value = "—"
    ws.Cells(rowNum, "F").Value = "Professional Fee"
    rowNum = rowNum + 1
    
    ' 99
    ws.Cells(rowNum, "B").Value = 99
    ws.Cells(rowNum, "C").Value = "TIN NOT FOUND"
    ws.Cells(rowNum, "D").Value = "Christian Anthony Ermita"
    ws.Cells(rowNum, "E").Value = "—"
    ws.Cells(rowNum, "F").Value = "Clinician's Fee-Visiting"
    rowNum = rowNum + 1
    
    ' 100
    ws.Cells(rowNum, "B").Value = 100
    ws.Cells(rowNum, "C").Value = "TIN NOT FOUND"
    ws.Cells(rowNum, "D").Value = "Vergilyn Montales"
    ws.Cells(rowNum, "E").Value = "—"
    ws.Cells(rowNum, "F").Value = "Clinician's Fee-Visiting"
    rowNum = rowNum + 1
    
    ' ---- VALIDATE FOR DUPLICATES ----
    ' Check: Same Name + Same Address = NOT ALLOWED
    lastRow = ws.Cells(ws.Rows.count, "D").End(xlUp).Row
    Dim duplicateFound As Boolean
    duplicateFound = False
    
    For i = 2 To lastRow
        Dim name As String
        Dim address As String
        Dim j As Long
        
        name = ws.Cells(i, "D").Value
        address = ws.Cells(i, "E").Value
        
        If name <> "" And address <> "" And address <> "—" Then
            For j = i + 1 To lastRow
                Dim name2 As String
                Dim address2 As String
                
                name2 = ws.Cells(j, "D").Value
                address2 = ws.Cells(j, "E").Value
                
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
    
    If Not duplicateFound Then
        MsgBox "Import complete! " & (rowNum - 2) & " suppliers imported." & vbCrLf & _
               "All suppliers have unique Name + Address combinations.", vbInformation
    End If
End Sub



