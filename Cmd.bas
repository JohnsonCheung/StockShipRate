Option Compare Database
Option Explicit
Sub RmvSchmPrpOnEr()
MdRmvPrpOnEr Md("Schm")
MdRmvPrpOnEr Md("SchmT")
MdRmvPrpOnEr Md("SchmF")
End Sub
Sub EnsSchmPrpOnEr()
MdEnsPrpOnEr Md("Schm")
MdEnsPrpOnEr Md("SchmT")
MdEnsPrpOnEr Md("SchmF")
End Sub