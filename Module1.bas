Attribute VB_Name = "Module1"
Sub Excel_AutoFilter_Multiple_Numbers_Using_VBA_Macro()
'
' Excel AutoFilter Multiple Numbers Using VBA Macro
' Purpose: The purpose of this VBA Macro is to Filter Multiple Numbers from a Column
' Date: 20160525
' Author: Nauman Khan
' Contact: http://naumankhan.blogspot.com
'

ActiveSheet.Range("$A$1:$A$100").AutoFilter Field:=1, _
        Criteria1:=Array("1", "5", "7", "10"), _
                Operator:=xlFilterValues

End Sub
