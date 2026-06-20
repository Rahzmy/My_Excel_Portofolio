Attribute VB_Name = "Module1"
Sub GenerateSummary()

    Dim SelectedMonth As String

    Dim Sales As Double
    Dim Qty As Double
    Dim Rating As Double
    Dim Shipping As Double

    Dim Performance As String
    Dim Insight As String
    Dim Summary As String

    Application.ScreenUpdating = False

    '==================================
    ' Read KPI
    '==================================

    SelectedMonth = Range("D2").Value

    Sales = Range("D5").Value
    Qty = Range("D6").Value
    Rating = Range("D7").Value
    Shipping = Range("D8").Value

    '==================================
    ' Performance Classification
    '==================================

    If Rating >= 4.5 Then

        Performance = "EXCELLENT"

        Insight = "Customer satisfaction berada pada level yang sangat baik. " & _
                  "Penjualan menunjukkan performa yang kuat dan konsisten selama bulan " & _
                  SelectedMonth & "."

    ElseIf Rating >= 4 Then

        Performance = "GOOD"

        Insight = "Kinerja bisnis berada pada level yang baik. " & _
                  "Masih terdapat peluang untuk meningkatkan customer experience."

    Else

        Performance = "NEEDS IMPROVEMENT"

        Insight = "Perlu perhatian lebih terhadap customer experience dan service quality " & _
                  "untuk meningkatkan performa bisnis secara keseluruhan."

    End If

    '==================================
    ' Build Executive Summary
    '==================================

    Summary = "EXECUTIVE SUMMARY - " & UCase(SelectedMonth) & vbCrLf & vbCrLf & _
              "Total Sales mencapai Rp " & Format(Sales, "#,##0") & _
              " dengan Quantity terjual sebanyak " & _
              Format(Qty, "#,##0") & " unit." & vbCrLf & vbCrLf & _
              "Average Rating berada pada level " & _
              Format(Rating, "0.0") & _
              " dengan Total Shipping Fee sebesar Rp " & _
              Format(Shipping, "#,##0") & "." & vbCrLf & vbCrLf & _
              Insight & vbCrLf & vbCrLf & _
              "Overall Performance : " & Performance

    '==================================
    ' Fake Processing Effect
    '==================================

    Application.StatusBar = "Generating Executive Summary..."
    Application.Wait Now + TimeValue("00:00:01")

    '==================================
    ' Write To Text Box
    '==================================

    ActiveSheet.Shapes("txtSummary") _
        .TextFrame.Characters.Text = Summary

    '==================================
    ' KPI Styling
    '==================================

    With Range("D5:D8")

        .Font.Bold = True
        .Font.Size = 11

    End With

    '==================================
    ' Color Logic
    '==================================

    If Rating >= 4.5 Then

        Range("D5:D8").Font.Color = RGB(0, 97, 0)

        ActiveSheet.Shapes("txtSummary") _
            .Fill.ForeColor.RGB = RGB(198, 239, 206)

    ElseIf Rating >= 4 Then

        Range("D5:D8").Font.Color = RGB(156, 101, 0)

        ActiveSheet.Shapes("txtSummary") _
            .Fill.ForeColor.RGB = RGB(255, 235, 156)

    Else

        Range("D5:D8").Font.Color = RGB(156, 0, 6)

        ActiveSheet.Shapes("txtSummary") _
            .Fill.ForeColor.RGB = RGB(255, 199, 206)

    End If

    '==================================
    ' Summary Text Styling
    '==================================

    With ActiveSheet.Shapes("txtSummary").TextFrame.Characters.Font

        .Name = "Calibri"
        .Size = 8
        .Color = RGB(0, 0, 0)

    End With

    Application.StatusBar = False
    Application.ScreenUpdating = True

    '==================================
    ' Performance Message
    '==================================

    If Rating >= 4.5 Then

        MsgBox "EXCELLENT PERFORMANCE!" & vbCrLf & vbCrLf & _
               "Month : " & SelectedMonth & vbCrLf & _
               "Sales : Rp " & Format(Sales, "#,##0") & vbCrLf & _
               "Rating : " & Format(Rating, "0.0") & vbCrLf & vbCrLf & _
               "Customer Satisfaction berada pada level yang sangat baik.", _
               vbInformation, "KarirNex Analytics"

    ElseIf Rating >= 4 Then

        MsgBox "GOOD PERFORMANCE!" & vbCrLf & vbCrLf & _
               "Month : " & SelectedMonth & vbCrLf & _
               "Rating : " & Format(Rating, "0.0") & vbCrLf & vbCrLf & _
               "Masih terdapat peluang untuk meningkatkan customer experience.", _
               vbInformation, "KarirNex Analytics"

    Else

        MsgBox "ATTENTION REQUIRED!" & vbCrLf & vbCrLf & _
               "Month : " & SelectedMonth & vbCrLf & _
               "Rating : " & Format(Rating, "0.0") & vbCrLf & vbCrLf & _
               "Perlu dilakukan improvement plan untuk meningkatkan performa bisnis.", _
               vbExclamation, "KarirNex Analytics"

    End If

    '==================================
    ' Achievement Unlocked
    '==================================

    MsgBox "ACHIEVEMENT UNLOCKED!" & vbCrLf & vbCrLf & _
           "[OK] Build KPI Dashboard" & vbCrLf & _
           "[OK] Import BAS File" & vbCrLf & _
           "[OK] Run VBA Macro" & vbCrLf & _
           "[OK] Generate Executive Summary" & vbCrLf & vbCrLf & _
           "Welcome to VBA Automation!", _
           vbInformation, "KarirNex Bootcamp"

End Sub

