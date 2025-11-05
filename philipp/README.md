**ERP**

*Teilaufgabe 1*

    1. Import und Strukturierung der Excel-Daten (z. B. mit pandas)  

    Funktion : load_data

    - Monate Vergangenheit  -    10.2025 - 09.2026
    - Jeder Monat hat eine Progronese in 1. Jahr und 2. Jahr
    - Daten nicht 100% Konsitent - teilweise abweichungen bei wavor_bstlmengemonat

    2.     Identifikation relevanter Hierarchieebenen (Artikel → Teilegruppe → Kundengruppe → Kunde → Gesamt) 

        1. In der Datei sind für jeden Monat für jeden Artikel die Ist und Soll Zahlen festgehalten.

        2. Es gibt Teilgruppen für Jeden Artikel bzw. None 

        3. Kundengruppe sind die Einzelen Baumärkte

        4. Jeder Baumarkt hat verschiedene Lieferorte
    
    3. 
     - Analyse des Baumarktprogramms
     [! Säuelendiagramm Baumarktprogramm](./output/images/baumarktprogramm_jahresvergleich.png)

     - Analyse auf Teilegruppen-Ebene
     




## ADD Info
- Toom fällt als Kunde raus? - Keine Progrosen (deckung mit dem Plan)
