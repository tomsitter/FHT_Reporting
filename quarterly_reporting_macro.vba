Public Sub clear()
Worksheets("billings").Range("A2:V9999").ClearContents

Worksheets("IndividualSummary-Detailed").Range("B8:D61").ClearContents
Worksheets("IndividualSummary-Detailed").Range("B2").ClearContents
Worksheets("IndividualSummary-Detailed").Range("J10").ClearContents
Worksheets("IndividualSummary-Detailed").Range("J13").ClearContents

Worksheets("IndividualSummary-Basic").Range("B9:D29").ClearContents
Worksheets("IndividualSummary-Basic").Range("B2").ClearContents

Worksheets("temp").Cells.ClearContents
End Sub

Public Sub CalcIndividualResp()

' New Billing Codes - 2016-12-28

' Physiotherapy
Dim PTNS As Long

' CHF
Dim CHFINI As Long
Dim CHFFU As Long
Dim CHFDU As Long
Dim CHFNS As Long

Dim CHGINI As Long
Dim CHGFU As Long
Dim CHGDU As Long
Dim CHGNS As Long


' Health Links
Dim HKINI As Long
Dim HKFU As Long
Dim HKDUHC As Long
Dim HKDISC As Long

' Mental Health
Dim MHDROP As Long
Dim MHDISC As Long
Dim MHNS As Long

' CKD
Dim CKDINI As Long
Dim CKDFU As Long
Dim CKDDUH As Long
Dim CKDGRP As Long
Dim CKDSCR As Long

' End of new billing codes

' Diabetes
Dim DMGINI As Long
Dim DMGen_B As Long
Dim DMGen_C As Long
Dim DMGS_A As Long
Dim DMGS_B As Long
Dim DMFV_A As Long
Dim DMFV_B As Long
Dim DMFV_C As Long


' Lung Health
Dim LH_A As Long
Dim LH_B As Long
Dim LH_C As Long

' Hypertension
Dim HTNGEN_A As Long
Dim HTNGEN_B As Long
Dim HTNGEN_C As Long

Dim HTNGS_A As Long
Dim HTNGS_B As Long

Dim HTNFV_A As Long
Dim HTNFV_B As Long
Dim HTNFV_C As Long
Dim HTNND As Long


Dim CHOLGEN_A As Long
Dim CHOLGEN_B As Long
Dim CHOLGEN_C As Long

Dim CHOLFV_A As Long
Dim CHOLFV_B As Long
Dim CHOLFV_C As Long

Dim CHOLGS_A As Long
Dim CHOLGS_B As Long

Dim CHF_A As Long
Dim CHF_B As Long
Dim CHF_C As Long

Dim MHI_A As Long
Dim MHI_B As Long
Dim MHI_C As Long

Dim MHDEP_A As Long
Dim MHDEP_B As Long
Dim MHMBSG_A As Long
Dim MHMBSG_B As Long
Dim MHMBCG_A As Long
Dim MHMBCG_B As Long
Dim MHANXG_A As Long
Dim MHANXG_B As Long

Dim SEN_A As Long
Dim SEN_B As Long
Dim SEN_C As Long

Dim SENMGS_A As Long
Dim SENMGS_B As Long
Dim SENEGS_A As Long
Dim SENEGS_B As Long
Dim SENFGS_A As Long
Dim SENFGS_B As Long
Dim SENNGS_A As Long
Dim SENNGS_B As Long
Dim SENCGS_A As Long
Dim SENCGS_B As Long

Dim SENSOC_A As Long
Dim SENSOC_B As Long
Dim SENSOC_C As Long

Dim HLI_A As Long
Dim HLI_B As Long
Dim HLI_C As Long

Dim SC_A As Long
Dim SC_B As Long
Dim SC_C As Long

Dim WM_A As Long
Dim WM_B As Long
Dim WM_C As Long

Dim WMGS_A As Long
Dim WMGS_B As Long

Dim VAC_A As Long
Dim VAC_B As Long
Dim VAC_C As Long

Dim PD_A As Long
Dim PD_B As Long
Dim PD_C As Long

Dim INSOMI_A As Long
Dim INSOMI_B As Long
Dim INSOMI_C As Long

Dim INSOMGS_A As Long
Dim INSOMGS_B As Long

Dim MAT_A As Long
Dim MAT_B As Long
Dim MAT_C As Long

Dim CANCERV_A As Long
Dim CANCERV_B As Long
Dim CANCERV_C As Long

Dim CANCERR_A As Long
Dim CANCERR_B As Long

Dim PAIN_A As Long
Dim PAIN_B As Long
Dim PAIN_C As Long

Dim PAINGS_A As Long
Dim PAINGS_B As Long

Dim medrec_a As Long
Dim medrec_b As Long
Dim medrec_c As Long

Dim pt_a As Long
Dim pt_b As Long
Dim pt_c As Long
Dim pt_disc As Long

Dim mind_a As Long
Dim mind_b As Long
Dim mind_c As Long

Dim MISCEL_A As Long
Dim MISCEL_B As Long
Dim INJECT_A As Long
Dim INJECT_B As Long
Dim BPCHECK_A As Long
Dim BPCHECK_B As Long
Dim SWAB_A As Long
Dim SWAB_B As Long
Dim LAB_A As Long
Dim LAB_B As Long
Dim FORM_A As Long
Dim FORM_B As Long
Dim RNHW_A As Long
Dim RNHW_B As Long

Application.ScreenUpdating = False
Worksheets("IndividualSummary-Detailed").Range("B9:D59").ClearContents
Worksheets("IndividualSummary-Basic").Range("B9:D29").ClearContents
Worksheets("temp").Cells.ClearContents

AllEntries = Application.WorksheetFunction.CountA(Sheets("billings").Range("A:A"))

'unique identifier (lastname, firstname)
Dim lastname As String
Dim firstname As String

With Sheets("billings")
    For i = 2 To AllEntries
        lastname = .Range("A" & i).Value
        firstname = .Range("B" & i).Value
        
        .Range("R" & i).Value = lastname & ", " & firstname
        
    Next i
End With

'copy and paste unique patient names to determine Column B - # of New Patients Included in Column A
Sheets("billings").Range("R:R").Copy
Sheets("temp").Select
Sheets("temp").Range("A1").Select
Sheets("temp").Paste

Sheets("temp").Range("A1:A" & AllEntries).RemoveDuplicates Columns:=1, Header:=xlYes

'DM General - # of Patient Visits
DMGINI = Application.WorksheetFunction.Sum(Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "DMGINI"), Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "DMGFU"), Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "DMGDHC"))
'DM General - # of New Patients Included in Column A
DMGen_B = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "DMGINI")
'DM General - # of Home Visits / DUHC Included in Column A
DMGen_C = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "DMGDHC")

'DM GS - # of Patient Visits
DMGS_A = Application.WorksheetFunction.Sum(Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "DMSM"), Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "DMCM"))
'DM GS - # of New Patients Included in Column A
Call TrueStatusMAT("AA", "DMSM", "DMCM*", "", "")
DMGS_B = Application.WorksheetFunction.CountIf(Sheets("temp").Range("AA:AA"), ">0")

'DM FV - # of Patient Visits
DMFV_A = Application.WorksheetFunction.Sum(Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "DMFINI"), Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "DMFRN1"), Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "DMFFU"), Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "DMFRN2"), Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "DMFDHC"), Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "DMFRNP"))
'DM Initial - # of New Patients Included in Column A
DMFV_B = Application.WorksheetFunction.Sum(Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "DMFINI"), Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "DMFRN1"))
'DM Initial - # of Home Visits / DUHC Included in Column A
DMFV_C = Application.WorksheetFunction.Sum(Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "DMFDHC"), Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "DMFRNP"))

'LH - # of Patient Visits
LH_A = Application.WorksheetFunction.Sum(Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "LHINI"), Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "LHFU"), Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "LHDUHC"))
'LH - # of New Patients Included in Column A
LH_B = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "LHINI")
'LH - # of Home Visits / DUHC Included in Column A
LH_C = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "LHDUHC")

'HTN General - # of Patient Visits
HTNGEN_A = Application.WorksheetFunction.Sum(Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "HTNEIN"), Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "HTNEFU"), Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "HTNEDU"))
'HTN General - # of New Patients Included in Column A
HTNGEN_B = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "HTNEIN")
'HTN General - # of Home Visits / DUHC Included in Column A
HTNGEN_C = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "HTNEDU")

'HTN GS VISIT - # of Patient Visits
HTNGS_A = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "HTNGS")
'HTN VISIT - # of New Patients Included in Column A
Call TrueStatus("AB", "HTNGS")
HTNGS_B = Application.WorksheetFunction.CountIf(Sheets("temp").Range("AB:AB"), ">0")

'HTN FV - # of Patient Visits
HTNFV_A = Application.WorksheetFunction.Sum(Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "HTNNIN"), Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "HTNNFU"), Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "HTNNDU"))
'HTN FV Initial - # of New Patients Included in Column A
HTNFV_B = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "HTNNIN")
'HTN FV Initial - # of Home Visits / DUHC Included in Column A
HTNFV_C = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "HTNNDU")
'HTN Newly Diagnosed - # of Newly Diagnosed Patients
HTNND = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "HTNNDS")

'CHOL General - # of Patient Visits
CHOLGEN_A = Application.WorksheetFunction.Sum(Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "CHOLG1"), Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "CHOLG2"), Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "CHOLDU"))
'CHOL General - # of New Patients Included in Column A
CHOLGEN_B = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "CHOLG1")
'CHOL General - # of Home Visits / DUHC Included in Column A
CHOLGEN_C = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "CHOLDU")

'CHOL FV - # of Patient Visits
CHOLFV_A = Application.WorksheetFunction.Sum(Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "CHOLF1"), Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "CHOLF2"), Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "CHOLFD"))
'CHOL FV Initial - # of New Patients Included in Column A
CHOLFV_B = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "CHOLF1")
'CHOL FV Initial - # of Home Visits / DUHC Included in Column A
CHOLFV_C = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "CHOLFD")

'CHOL GS VISIT - # of Patient Visits
CHOLGS_A = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "CHOLGS")
'CHOL VISIT - # of New Patients Included in Column A
Call TrueStatus("C", "CHOLGS")
CHOLGS_B = Application.WorksheetFunction.CountIf(Sheets("temp").Range("C:C"), ">0")

'CHF - # of Patient Visits
CHF_A = Application.WorksheetFunction.Sum(Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "CHFINI"), Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "CHFFU"), Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "CHFDU"))
'CHF - # of New Patients Included in Column A
CHF_B = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "CHFINI")
'CHF - # of Home Visits / DUHC Included in Column A
CHF_C = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "CHFDU")

'MH indiv - # of Patient Visits
MHI_A = Application.WorksheetFunction.Sum(Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "MHINI"), Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "MHFU"), Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "MHDUHC"))
'MH indiv - # of New Patients Included in Column A
MHI_B = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "MHINI")
'MH indiv - # of Home Visits / DUHC Included in Column A
MHI_C = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "MHDUHC")

'MHGS - # of Patient Visits
MHDEP_A = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "MHDEPG")
MHMBSG_A = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "MHMBSG")
MHMBCG_A = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "MHMBCG")
MHANXG_A = Application.WorksheetFunction.SumIf(Sheets("billings").Range("J:J"), "MHANXG", Sheets("billings").Range("L:L"))

'Mental health GS - # of New Patients Included in Column A
Call TrueStatus("D", "MHDEPG")
MHDEP_B = Application.WorksheetFunction.CountIf(Sheets("temp").Range("D:D"), ">0")
Call TrueStatus("E", "MHMBSG")
MHMBSG_B = Application.WorksheetFunction.CountIf(Sheets("temp").Range("E:E"), ">0")
Call TrueStatus("F", "MHMBCG")
MHMBCG_B = Application.WorksheetFunction.CountIf(Sheets("temp").Range("F:F"), ">0")
Call TrueStatus("G", "MHANXG")
MHANXG_B = Application.WorksheetFunction.CountIf(Sheets("temp").Range("G:G"), ">0")

'SENIORS indiv - # of Patient Visits
SEN_A = Application.WorksheetFunction.Sum(Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "SENINI"), Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "SENFU"), Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "SENDU"), Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "SENDU2"))
'SENIORS indiv - # of New Patients Included in Column A
SEN_B = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "SENINI")
'SENIORS indiv - # of Home Visits / DUHC Included in Column A
SEN_C = Application.WorksheetFunction.Sum(Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "SENDU"), Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "SENDU2"))

'SENIORS GS - # of Patient Visits
SENMGS_A = Application.WorksheetFunction.SumIf(Sheets("billings").Range("J:J"), "SENMGS", Sheets("billings").Range("L:L"))

SENEGS_A = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "SENEGS")
SENFGS_A = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "SENFGS")
SENNGS_A = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "SENNGS")
SENCGS_A = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "SENCGS")

'Senior GS - # of New Patients Included in Column A
Call TrueStatus("H", "SENMGS")
SENMGS_B = Application.WorksheetFunction.CountIf(Sheets("temp").Range("H:H"), ">0")
Call TrueStatus("I", "SENEGS")
SENEGS_B = Application.WorksheetFunction.CountIf(Sheets("temp").Range("I:I"), ">0")
Call TrueStatus("J", "SENFGS")
SENFGS_B = Application.WorksheetFunction.CountIf(Sheets("temp").Range("J:J"), ">0")
Call TrueStatus("K", "SENNGS")
SENNGS_B = Application.WorksheetFunction.CountIf(Sheets("temp").Range("K:K"), ">0")
Call TrueStatus("L", "SENCGS")
SENCGS_B = Application.WorksheetFunction.CountIf(Sheets("temp").Range("L:L"), ">0")

'SENIORS Social Club - # of Patient Visits
SENSOC_A = Application.WorksheetFunction.Sum(Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "SENSOC"), Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "SENSO2"), Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "SENSOD"))
'SENIORS Social Club - # of New Patients Included in Column A
SENSOC_B = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "SENSOC")
'SENIORS Social Club - # of Home Visits / DUHC Included in Column A
SENSOC_C = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "SENSOD")

'HEALTHY LIVING indiv - # of Patient Visits
HLI_A = Application.WorksheetFunction.Sum(Application.WorksheetFunction.SumIf(Sheets("billings").Range("J:J"), "HLINI", Sheets("billings").Range("L:L")), Application.WorksheetFunction.SumIf(Sheets("billings").Range("J:J"), "HLFU", Sheets("billings").Range("L:L")), Application.WorksheetFunction.SumIf(Sheets("billings").Range("J:J"), "HLDUHC", Sheets("billings").Range("L:L")))

'HEALTHY LIVING indiv - # of New Patients Included in Column A
HLI_B = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "HLINI")
'HEALTHY LIVING indiv - # of Home Visits / DUHC Included in Column A
HLI_C = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "HLDUHC")

'SMOKING CESSATION - # of Patient Visits
SC_A = Application.WorksheetFunction.Sum(Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "SCINI"), Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "SCFU"), Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "SCDUHC"))
'SMOKING CESSATION - # of New Patients Included in Column A
SC_B = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "SCINI")
'SMOKING CESSATION - # of Home Visits / DUHC Included in Column A
SC_C = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "SCDUHC")

'WEIGHT MANAGEMENT - # of Patient Visits
WM_A = Application.WorksheetFunction.Sum(Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "WINI"), Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "WFU"), Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "WDUHC"))
'WEIGHT MANAGEMENT - # of New Patients Included in Column A
WM_B = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "WINI")
'WEIGHT MANAGEMENT - # of Home Visits / DUHC Included in Column A
WM_C = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "WDUHC")

'WEIGHT MANAGEMENT GS - # of Patient Visits
WMGS_A = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "WGS")
'weight management GS - # of New Patients Included in Column A
Call TrueStatus("M", "WGS")
WMGS_B = Application.WorksheetFunction.CountIf(Sheets("temp").Range("M:M"), ">0")

'VACCINE - # of Patient Visits
VAC_A = Application.WorksheetFunction.Sum(Application.WorksheetFunction.SumIf(Sheets("billings").Range("J:J"), "RNVAC", Sheets("billings").Range("L:L")), Application.WorksheetFunction.SumIf(Sheets("billings").Range("J:J"), "IMMUNR", Sheets("billings").Range("L:L")))
'VACCINE - # of New Patients Included in Column A
Call TrueStatusCANCER("N", "RNVAC", "IMMUNR", "RNVAC")
VAC_B = Application.WorksheetFunction.CountIf(Sheets("temp").Range("N:N"), ">0")

'Prediabetes - # of Patient Visits
PD_A = Application.WorksheetFunction.Sum(Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "PDINI"), Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "PDFU"), Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "PDDUHC"))
'Prediabetes - # of New Patients Included in Column A
PD_B = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "PDINI")
'Prediabetes - # of Home Visits / DUHC Included in Column A
PD_C = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "PDDUHC")

'Insomnia indiv - # of Patient Visits
INSOMI_A = Application.WorksheetFunction.Sum(Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "INSOIN"), Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "INSOFU"), Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "INSODU"))
'Insomnia indiv - # of New Patients Included in Column A
INSOMI_B = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "INSOIN")
'Insomnia indiv - # of Home Visits / DUHC Included in Column A
INSOMI_C = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "INSODU")

'insomnia GS - # of Patient Visits
INSOMGS_A = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "INSOGS")
'insomnia GS - # of New Patients Included in Column A
Call TrueStatus("O", "INSOGS")
INSOMGS_B = Application.WorksheetFunction.CountIf(Sheets("temp").Range("O:O"), ">0")

'MATERNAL HEALTH - # of Patient Visits
MAT_A = Application.WorksheetFunction.Sum(Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "MCP*"), Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "MCBF"), Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "MCWBC"), Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "MCGEN"), Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "MCDUHC"))
'MATERNAL HEALTH - # of New Patients Included in Column A
Call TrueStatusMAT("Z", "MCBF", "MCP*", "MCWBC", "MCGEN")
MAT_B = Application.WorksheetFunction.CountIf(Sheets("temp").Range("Z:Z"), ">0")
'MATERNAL HEALTH - # of Home Visits / DUHC Included in Column A
MAT_C = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "MCDUHC")

'CANCER VISITS - # of Patient Visits
CANCERV_A = Application.WorksheetFunction.Sum(Application.WorksheetFunction.SumIf(Sheets("billings").Range("J:J"), "CERVV", Sheets("billings").Range("L:L")), Application.WorksheetFunction.SumIf(Sheets("billings").Range("J:J"), "COLV", Sheets("billings").Range("L:L")), Application.WorksheetFunction.SumIf(Sheets("billings").Range("J:J"), "BREASV", Sheets("billings").Range("L:L")), Application.WorksheetFunction.SumIf(Sheets("billings").Range("J:J"), "CANDU", Sheets("billings").Range("L:L")))
'CANCER VISITS - # of New Patients Included in Column A
Call TrueStatusCANCER("P", "COLV", "CERV", "BREASV*")
CANCERV_B = Application.WorksheetFunction.CountIf(Sheets("temp").Range("P:P"), ">0")
'Cancer visits - # of Home Visits / DUHC Included in Column A
CANCERV_C = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "CANDU")

'CANCER REVIEW - # of Patient Visits
CANCERR_A = Application.WorksheetFunction.Sum(Application.WorksheetFunction.SumIf(Sheets("billings").Range("J:J"), "CERVR", Sheets("billings").Range("L:L")), Application.WorksheetFunction.SumIf(Sheets("billings").Range("J:J"), "COLR", Sheets("billings").Range("L:L")), Application.WorksheetFunction.SumIf(Sheets("billings").Range("J:J"), "BREASR", Sheets("billings").Range("L:L")))

'PAIN MANAGEMENT - # of Patient Visits
PAIN_A = Application.WorksheetFunction.Sum(Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "PAININ"), Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "PAINFU"), Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "PAINDU"))
'PAIN MANAGEMENT - # of New Patients Included in Column A
PAIN_B = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "PAININ")
'PAIN MANAGEMENT - # of Home Visits / DUHC Included in Column A
PAIN_C = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "PAINDU")

'PAIN MANAGEMENT - # of Patient Visits
PAINGS_A = Application.WorksheetFunction.SumIf(Sheets("billings").Range("J:J"), "PAINGS", Sheets("billings").Range("L:L"))

'PAIN MANAGEMENT - # of New Patients Included in Column A
Call TrueStatus("Q", "PAINGS")
PAINGS_B = Application.WorksheetFunction.CountIf(Sheets("temp").Range("Q:Q"), ">0")

'medrec - # of Patient Visits
medrec_a = Application.WorksheetFunction.Sum(Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "MEDINI"), Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "MEDFU"), Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "MEDDHC"))
'medrec - # of New Patients Included in Column A
medrec_b = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "MEDINI")
'medrec - # of Home Visits / DUHC Included in Column A
medrec_c = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "MEDDHC")

'PT - # of Patient Visits
pt_a = Application.WorksheetFunction.Sum(Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "PTINI"), Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "PTFU"), Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "PTDUHC"))
'PT - # of New Patients Included in Column A
pt_b = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "PTINI")
'PT - # of Home Visits / DUHC Included in Column A
pt_c = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "PTDUHC")
'PT - # of Discharged Patients
pt_disc = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "PTDISC")

'Mindful Eating GS - # of Patient Visits
mind_a = Application.WorksheetFunction.SumIf(Sheets("billings").Range("J:J"), "MINDGS", Sheets("billings").Range("L:L"))
'Mindful Eating GS - # of New Patients Included in Column A
Call TrueStatus("R", "MINDGS")
mind_b = Application.WorksheetFunction.CountIf(Sheets("temp").Range("R:R"), ">0")

'MISCEL RN - # of Patient Visits
MISCEL_A = Application.WorksheetFunction.SumIf(Sheets("billings").Range("J:J"), "MISCEL", Sheets("billings").Range("L:L"))

'MISCEL RN - # of New Patients Included in Column A
Call TrueStatus("S", "MISCEL")
MISCEL_B = Application.WorksheetFunction.CountIf(Sheets("temp").Range("S:S"), ">0")

'INJECT RN - # of Patient Visits
INJECT_A = Application.WorksheetFunction.SumIf(Sheets("billings").Range("J:J"), "INJECT", Sheets("billings").Range("L:L"))

'INJECT RN - # of New Patients Included in Column A
Call TrueStatus("T", "INJECT")
INJECT_B = Application.WorksheetFunction.CountIf(Sheets("temp").Range("T:T"), ">0")

'BPCHECK RN - # of Patient Visits
BPCHECK_A = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "BPCHEK")

'BPCHECK RN - # of New Patients Included in Column A
Call TrueStatus("U", "BPCHEK")
BPCHECK_B = Application.WorksheetFunction.CountIf(Sheets("temp").Range("U:U"), ">0")

'SWAB RN - # of Patient Visits
SWAB_A = Application.WorksheetFunction.SumIf(Sheets("billings").Range("J:J"), "SWABR", Sheets("billings").Range("L:L"))

'SWAB RN - # of New Patients Included in Column A
Call TrueStatus("V", "SWABR")
SWAB_B = Application.WorksheetFunction.CountIf(Sheets("temp").Range("V:V"), ">0")

'LAB Review RN - # of Patient Visits
LAB_A = Application.WorksheetFunction.SumIf(Sheets("billings").Range("J:J"), "LABR", Sheets("billings").Range("L:L"))

'LAB Review RN - # of New Patients Included in Column A
Call TrueStatus("W", "LABR")
LAB_B = Application.WorksheetFunction.CountIf(Sheets("temp").Range("W:W"), ">0")

'FORM - # of Patient Visits
FORM_A = Application.WorksheetFunction.SumIf(Sheets("billings").Range("J:J"), "FORM", Sheets("billings").Range("L:L"))

'FORM - # of New Patients Included in Column A
Call TrueStatus("X", "FORM")
FORM_B = Application.WorksheetFunction.CountIf(Sheets("temp").Range("X:X"), ">0")

'RN Height/Weight Reveiew - # of Patient Visits
RNHW_A = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "RNHW")

'RN Height/Weight Review - # of New Patients Included in Column A
Call TrueStatus("Y", "RNHW")
RNHW_B = Application.WorksheetFunction.CountIf(Sheets("temp").Range("Y:Y"), ">0")


'ENTERING VALUES INTO APPROPRIATE CELLS in "IndividualSummary-Detailed"
Sheets("IndividualSummary-Detailed").Select

Sheets("IndividualSummary-Detailed").Range("B8").Value = DMGINI
Sheets("IndividualSummary-Detailed").Range("c8").Value = DMGen_B
Sheets("IndividualSummary-Detailed").Range("d8").Value = DMGen_C

Sheets("IndividualSummary-Detailed").Range("B9").Value = DMGS_A
Sheets("IndividualSummary-Detailed").Range("C9").Value = DMGS_B

Sheets("IndividualSummary-Detailed").Range("B10").Value = DMFV_A
Sheets("IndividualSummary-Detailed").Range("c10").Value = DMFV_B
Sheets("IndividualSummary-Detailed").Range("d10").Value = DMFV_C

Sheets("IndividualSummary-Detailed").Range("B11").Value = LH_A
Sheets("IndividualSummary-Detailed").Range("c11").Value = LH_B
Sheets("IndividualSummary-Detailed").Range("d11").Value = LH_C

Sheets("IndividualSummary-Detailed").Range("B13").Value = HTNGEN_A
Sheets("IndividualSummary-Detailed").Range("c13").Value = HTNGEN_B
Sheets("IndividualSummary-Detailed").Range("d13").Value = HTNGEN_C

Sheets("IndividualSummary-Detailed").Range("B14").Value = HTNFV_A
Sheets("IndividualSummary-Detailed").Range("c14").Value = HTNFV_B
Sheets("IndividualSummary-Detailed").Range("d14").Value = HTNFV_C

Sheets("IndividualSummary-Detailed").Range("j10").Value = HTNND

Sheets("IndividualSummary-Detailed").Range("B15").Value = HTNGS_A
Sheets("IndividualSummary-Detailed").Range("c15").Value = HTNGS_B

Sheets("IndividualSummary-Detailed").Range("B16").Value = CHOLGEN_A
Sheets("IndividualSummary-Detailed").Range("c16").Value = CHOLGEN_B
Sheets("IndividualSummary-Detailed").Range("d16").Value = CHOLGEN_C

Sheets("IndividualSummary-Detailed").Range("B17").Value = CHOLFV_A
Sheets("IndividualSummary-Detailed").Range("c17").Value = CHOLFV_B
Sheets("IndividualSummary-Detailed").Range("d17").Value = CHOLFV_C

Sheets("IndividualSummary-Detailed").Range("B18").Value = CHOLGS_A
Sheets("IndividualSummary-Detailed").Range("c18").Value = CHOLGS_B

Sheets("IndividualSummary-Detailed").Range("B19").Value = CHF_A
Sheets("IndividualSummary-Detailed").Range("c19").Value = CHF_B
Sheets("IndividualSummary-Detailed").Range("d19").Value = CHF_C

Sheets("IndividualSummary-Detailed").Range("B21").Value = MHI_A
Sheets("IndividualSummary-Detailed").Range("c21").Value = MHI_B
Sheets("IndividualSummary-Detailed").Range("d21").Value = MHI_C

Sheets("IndividualSummary-Detailed").Range("B22").Value = MHMBSG_A
Sheets("IndividualSummary-Detailed").Range("B23").Value = MHMBCG_A
Sheets("IndividualSummary-Detailed").Range("B24").Value = MHANXG_A
Sheets("IndividualSummary-Detailed").Range("B25").Value = MHDEP_A

Sheets("IndividualSummary-Detailed").Range("C22").Value = MHMBSG_B
Sheets("IndividualSummary-Detailed").Range("C23").Value = MHMBCG_B
Sheets("IndividualSummary-Detailed").Range("C24").Value = MHANXG_B
Sheets("IndividualSummary-Detailed").Range("C25").Value = MHDEP_B

Sheets("IndividualSummary-Detailed").Range("B28").Value = SEN_A
Sheets("IndividualSummary-Detailed").Range("c28").Value = SEN_B
Sheets("IndividualSummary-Detailed").Range("d28").Value = SEN_C

Sheets("IndividualSummary-Detailed").Range("B29").Value = SENMGS_A
Sheets("IndividualSummary-Detailed").Range("B30").Value = SENEGS_A
Sheets("IndividualSummary-Detailed").Range("B31").Value = SENFGS_A
Sheets("IndividualSummary-Detailed").Range("B32").Value = SENNGS_A
Sheets("IndividualSummary-Detailed").Range("B33").Value = SENCGS_A

Sheets("IndividualSummary-Detailed").Range("C29").Value = SENMGS_B
Sheets("IndividualSummary-Detailed").Range("C30").Value = SENEGS_B
Sheets("IndividualSummary-Detailed").Range("C31").Value = SENFGS_B
Sheets("IndividualSummary-Detailed").Range("C32").Value = SENNGS_B
Sheets("IndividualSummary-Detailed").Range("C33").Value = SENCGS_B

Sheets("IndividualSummary-Detailed").Range("B34").Value = SENSOC_A
Sheets("IndividualSummary-Detailed").Range("c34").Value = SENSOC_B
Sheets("IndividualSummary-Detailed").Range("d34").Value = SENSOC_C

Sheets("IndividualSummary-Detailed").Range("B35").Value = medrec_a
Sheets("IndividualSummary-Detailed").Range("c35").Value = medrec_b
Sheets("IndividualSummary-Detailed").Range("d35").Value = medrec_c

Sheets("IndividualSummary-Detailed").Range("B37").Value = HLI_A
Sheets("IndividualSummary-Detailed").Range("c37").Value = HLI_B
Sheets("IndividualSummary-Detailed").Range("d37").Value = HLI_C

Sheets("IndividualSummary-Detailed").Range("B38").Value = SC_A
Sheets("IndividualSummary-Detailed").Range("c38").Value = SC_B
Sheets("IndividualSummary-Detailed").Range("d38").Value = SC_C

Sheets("IndividualSummary-Detailed").Range("B39").Value = WM_A
Sheets("IndividualSummary-Detailed").Range("c39").Value = WM_B
Sheets("IndividualSummary-Detailed").Range("d39").Value = WM_C

Sheets("IndividualSummary-Detailed").Range("B40").Value = WMGS_A
Sheets("IndividualSummary-Detailed").Range("C40").Value = WMGS_B

Sheets("IndividualSummary-Detailed").Range("B41").Value = VAC_A
Sheets("IndividualSummary-Detailed").Range("C41").Value = VAC_B

Sheets("IndividualSummary-Detailed").Range("B42").Value = PD_A
Sheets("IndividualSummary-Detailed").Range("c42").Value = PD_B
Sheets("IndividualSummary-Detailed").Range("d42").Value = PD_C

Sheets("IndividualSummary-Detailed").Range("B43").Value = INSOMI_A
Sheets("IndividualSummary-Detailed").Range("c43").Value = INSOMI_B
Sheets("IndividualSummary-Detailed").Range("d43").Value = INSOMI_C

Sheets("IndividualSummary-Detailed").Range("B44").Value = INSOMGS_A
Sheets("IndividualSummary-Detailed").Range("C44").Value = INSOMGS_B

Sheets("IndividualSummary-Detailed").Range("B45").Value = MAT_A
Sheets("IndividualSummary-Detailed").Range("C45").Value = MAT_B
Sheets("IndividualSummary-Detailed").Range("D45").Value = MAT_C

Sheets("IndividualSummary-Detailed").Range("B46").Value = CANCERV_A
Sheets("IndividualSummary-Detailed").Range("B47").Value = CANCERR_A

Sheets("IndividualSummary-Detailed").Range("C46").Value = CANCERV_B

Sheets("IndividualSummary-Detailed").Range("D46").Value = CANCERV_C

Sheets("IndividualSummary-Detailed").Range("B48").Value = PAIN_A
Sheets("IndividualSummary-Detailed").Range("c48").Value = PAIN_B
Sheets("IndividualSummary-Detailed").Range("d48").Value = PAIN_C

Sheets("IndividualSummary-Detailed").Range("B49").Value = PAINGS_A
Sheets("IndividualSummary-Detailed").Range("c49").Value = PAINGS_B

Sheets("IndividualSummary-Detailed").Range("B50").Value = pt_a
Sheets("IndividualSummary-Detailed").Range("c50").Value = pt_b
Sheets("IndividualSummary-Detailed").Range("d50").Value = pt_c

Sheets("IndividualSummary-Detailed").Range("j13").Value = pt_disc

Sheets("IndividualSummary-Detailed").Range("B51").Value = mind_a
Sheets("IndividualSummary-Detailed").Range("c51").Value = mind_b

Sheets("IndividualSummary-Detailed").Range("B53").Value = MISCEL_A
Sheets("IndividualSummary-Detailed").Range("c53").Value = MISCEL_B
Sheets("IndividualSummary-Detailed").Range("B54").Value = INJECT_A
Sheets("IndividualSummary-Detailed").Range("c54").Value = INJECT_B
Sheets("IndividualSummary-Detailed").Range("B55").Value = BPCHECK_A
Sheets("IndividualSummary-Detailed").Range("c55").Value = BPCHECK_B
Sheets("IndividualSummary-Detailed").Range("B56").Value = SWAB_A
Sheets("IndividualSummary-Detailed").Range("c56").Value = SWAB_B
Sheets("IndividualSummary-Detailed").Range("B57").Value = LAB_A
Sheets("IndividualSummary-Detailed").Range("c57").Value = LAB_B
Sheets("IndividualSummary-Detailed").Range("B58").Value = FORM_A
Sheets("IndividualSummary-Detailed").Range("c58").Value = FORM_B
Sheets("IndividualSummary-Detailed").Range("B59").Value = RNHW_A
Sheets("IndividualSummary-Detailed").Range("c59").Value = RNHW_B

Sheets("IndividualSummary-Detailed").Range("B61").Value = HLI_A + FORM_A + MISCEL_A + INJECT_A + SWAB_A + LAB_A
Sheets("IndividualSummary-Detailed").Range("C61").Value = HLI_B + FORM_B + MISCEL_B + INJECT_B + SWAB_B + LAB_B
Sheets("IndividualSummary-Detailed").Range("D61").Value = HLI_C


'ENTERING VALUES INTO APPROPRIATE CELLS in "IndividualSummary-Basic"
Sheets("IndividualSummary-Basic").Select

Sheets("IndividualSummary-Basic").Range("B9").Value = DMGINI + DMGS_A + DMFV_A
Sheets("IndividualSummary-Basic").Range("C9").Value = DMGen_B + DMGS_B + DMFV_B
Sheets("IndividualSummary-Basic").Range("D9").Value = DMGen_C + DMGS_C + DMFV_C

Sheets("IndividualSummary-Basic").Range("B10").Value = LH_A
Sheets("IndividualSummary-Basic").Range("c10").Value = LH_B
Sheets("IndividualSummary-Basic").Range("d10").Value = LH_C

Sheets("IndividualSummary-Basic").Range("B12").Value = HTNGEN_A + HTNFV_A + HTNGS_A
Sheets("IndividualSummary-Basic").Range("c12").Value = HTNGEN_B + HTNFV_B + HTNGS_B
Sheets("IndividualSummary-Basic").Range("d12").Value = HTNGEN_C + HTNFV_C

Sheets("IndividualSummary-Basic").Range("B13").Value = CHOLGEN_A + CHOLFV_A + CHOLGS_A
Sheets("IndividualSummary-Basic").Range("c13").Value = CHOLGEN_B + CHOLFV_B + CHOLGS_B
Sheets("IndividualSummary-Basic").Range("d13").Value = CHOLGEN_C + CHOLFV_C + CHOLGS_C

Sheets("IndividualSummary-Basic").Range("B14").Value = CHF_A
Sheets("IndividualSummary-Basic").Range("c14").Value = CHF_B
Sheets("IndividualSummary-Basic").Range("d14").Value = CHF_C

Sheets("IndividualSummary-Basic").Range("B15").Value = MHI_A + MHMBSG_A + MHMBCG_A + MHANXG_A + MHDEP_A
Sheets("IndividualSummary-Basic").Range("c15").Value = MHI_B + MHMBSG_B + MHMBCG_B + MHANXG_B + MHDEP_B
Sheets("IndividualSummary-Basic").Range("d15").Value = MHI_C + MHMBSG_C + MHMBCG_C + MHANXG_C + MHDEP_C

Sheets("IndividualSummary-Basic").Range("B18").Value = SEN_A + SENMGS_A + SENEGS_A + SENFGS_A + SENNGS_A + SENCGS_A + SENSOC_A
Sheets("IndividualSummary-Basic").Range("C18").Value = SEN_B + SENMGS_B + SENEGS_B + SENFGS_B + SENNGS_B + SENCGS_B + SENSOC_B
Sheets("IndividualSummary-Basic").Range("D18").Value = SEN_C + SENMGS_C + SENEGS_C + SENFGS_C + SENNGS_C + SENCGS_C + SENSOC_C

Sheets("IndividualSummary-Basic").Range("B20").Value = SC_A
Sheets("IndividualSummary-Basic").Range("c20").Value = SC_B
Sheets("IndividualSummary-Basic").Range("d20").Value = SC_C

Sheets("IndividualSummary-Basic").Range("B21").Value = WM_A + WMGS_A + mind_a
Sheets("IndividualSummary-Basic").Range("c21").Value = WM_B + WMGS_B + mind_b
Sheets("IndividualSummary-Basic").Range("d21").Value = WM_C

Sheets("IndividualSummary-Basic").Range("B22").Value = VAC_A
Sheets("IndividualSummary-Basic").Range("C22").Value = VAC_B

Sheets("IndividualSummary-Basic").Range("B23").Value = PD_A
Sheets("IndividualSummary-Basic").Range("c23").Value = PD_B
Sheets("IndividualSummary-Basic").Range("d23").Value = PD_C

Sheets("IndividualSummary-Basic").Range("B24").Value = INSOMI_A + INSOMGS_A
Sheets("IndividualSummary-Basic").Range("c24").Value = INSOMI_B + INSOMGS_B
Sheets("IndividualSummary-Basic").Range("d24").Value = INSOMI_C

Sheets("IndividualSummary-Basic").Range("B25").Value = MAT_A
Sheets("IndividualSummary-Basic").Range("C25").Value = MAT_B
Sheets("IndividualSummary-Basic").Range("D25").Value = MAT_C

Sheets("IndividualSummary-Basic").Range("B26").Value = CANCERV_A + CANCERR_A
Sheets("IndividualSummary-Basic").Range("C26").Value = CANCERV_B
Sheets("IndividualSummary-Basic").Range("D26").Value = CANCERV_C

Sheets("IndividualSummary-Basic").Range("B27").Value = PAIN_A + PAINGS_A
Sheets("IndividualSummary-Basic").Range("c27").Value = PAIN_B + PAINGS_B
Sheets("IndividualSummary-Basic").Range("d27").Value = PAIN_C

Sheets("IndividualSummary-Basic").Range("B29").Value = HLI_A + FORM_A + MISCEL_A + INJECT_A + SWAB_A + LAB_A
Sheets("IndividualSummary-Basic").Range("C29").Value = HLI_B + FORM_B + MISCEL_B + INJECT_B + SWAB_B + LAB_B
Sheets("IndividualSummary-Basic").Range("D29").Value = HLI_C

'provider name on report
Sheets("IndividualSummary-Basic").Range("B2").Value = Sheets("billings").Range("O2").Value & ", " & Sheets("billings").Range("P2").Value
Sheets("IndividualSummary-Detailed").Range("B2").Value = Sheets("billings").Range("O2").Value & ", " & Sheets("billings").Range("P2").Value

End Sub


Public Sub TrueStatus(column, code)

uniquepat = Sheets("temp").Cells(Rows.Count, 1).End(xlUp).Row

With Sheets("temp")
    For i = 2 To uniquepat + 1
        identifier = Sheets("temp").Range("A" & i).Value
        x = Application.WorksheetFunction.CountIfs(Sheets("billings").Range("R1:R5500"), identifier, Sheets("billings").Range("J1:J5500"), code)
        Sheets("temp").Range(column & i).Value = x
    Next i
End With



End Sub

Public Sub TrueStatusMAT(column, code1, code2, code3, code4)

uniquepat = Sheets("temp").Cells(Rows.Count, 1).End(xlUp).Row

With Sheets("temp")
    For i = 2 To uniquepat + 1
        identifier = Sheets("temp").Range("A" & i).Value
        x = Application.WorksheetFunction.Sum(Application.WorksheetFunction.CountIfs(Sheets("billings").Range("R1:R5500"), identifier, Sheets("billings").Range("J1:J5500"), code1), Application.WorksheetFunction.CountIfs(Sheets("billings").Range("R1:R5500"), identifier, Sheets("billings").Range("J1:J5500"), code2), Application.WorksheetFunction.CountIfs(Sheets("billings").Range("R1:R5500"), identifier, Sheets("billings").Range("J1:J5500"), code3), Application.WorksheetFunction.CountIfs(Sheets("billings").Range("R1:R5500"), identifier, Sheets("billings").Range("J1:J5500"), code4))
        Sheets("temp").Range(column & i).Value = x
    Next i
End With

End Sub

Public Sub TrueStatusCANCER(column, code1, code2, code3)

uniquepat = Sheets("temp").Cells(Rows.Count, 1).End(xlUp).Row

With Sheets("temp")
    For i = 2 To uniquepat + 1
        identifier = Sheets("temp").Range("A" & i).Value
        x = Application.WorksheetFunction.Sum(Application.WorksheetFunction.CountIfs(Sheets("billings").Range("R1:R5500"), identifier, Sheets("billings").Range("J1:J5500"), code1), Application.WorksheetFunction.CountIfs(Sheets("billings").Range("R1:R5500"), identifier, Sheets("billings").Range("J1:J5500"), code2), Application.WorksheetFunction.CountIfs(Sheets("billings").Range("R1:R5500"), identifier, Sheets("billings").Range("J1:J5500"), code3))
        Sheets("temp").Range(column & i).Value = x
    Next i
End With

End Sub

