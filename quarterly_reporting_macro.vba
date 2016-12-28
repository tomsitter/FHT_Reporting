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
Dim DMGINI_TOTAL As Long
Dim DMGINI_NEW As Long
Dim DMGDHC As Long
Dim DMGROUP_TOTAL As Long
Dim DMGROUP_NEW As Long
Dim DMFOCUS_TOTAL As Long
Dim DMFOCUS_NEW As Long
Dim DMFDHC As Long


' Lung Health
Dim LH_TOTAL As Long
Dim LH_NEW As Long
Dim LHDUHC As Long

' Hypertension
Dim HTNE_TOTAL As Long
Dim HTNE_NEW As Long
Dim HTNEDU As Long

Dim HTNGS_TOTAL As Long
Dim HTNGS_NEW As Long

Dim HTNFV_TOTAL As Long
Dim HTNFV_NEW As Long
Dim HTNNDU As Long
Dim HTNNDS As Long


Dim CHOLG_TOTAL As Long
Dim CHOLG_NEW As Long
Dim CHOLDU As Long

Dim CHOLFV_TOTAL As Long
Dim CHOLFV_NEW As Long
Dim CHOLFD As Long

Dim CHOLGS_TOTAL As Long
Dim CHOLGS_NEW As Long

Dim CHF_TOTAL As Long
Dim CHF_NEW As Long
Dim CHFDU As Long

' Mental Health
Dim MH_TOTAL As Long
Dim MH_NEW As Long
Dim MHDUHC As Long

Dim MHDEP_TOTAL As Long
Dim MHDEP_NEW As Long
Dim MHMBSG_TOTAL As Long
Dim MHMBSG_NEW As Long
Dim MHMBCG_TOTAL As Long
Dim MHMBCG_NEW As Long
Dim MHANXG_TOTAL As Long
Dim MHANXG_NEW As Long

' Seniors
Dim SEN_TOTAL As Long
Dim SEN_NEW As Long
Dim SENDU As Long

Dim SENMGS_TOTAL As Long
Dim SENMGS_NEW As Long
Dim SENEGS_TOTAL As Long
Dim SENEGS_NEW As Long
Dim SENFGS_TOTAL As Long
Dim SENFGS_NEW As Long
Dim SENNGS_TOTAL As Long
Dim SENNGS_NEW As Long
Dim SENCGS_TOTAL As Long
Dim SENCGS_NEW As Long

Dim SENSOC_TOTAL As Long
Dim SENSOC_NEW As Long
Dim SENSOD As Long

Dim HLI_TOTAL As Long
Dim HLI_NEW As Long
Dim HLDUHC As Long

Dim SC_TOTAL As Long
Dim SC_NEW As Long
Dim SCDUHC As Long

Dim WM_TOTAL As Long
Dim WM_NEW As Long
Dim WDUHC As Long

Dim WMGS_TOTAL As Long
Dim WMGS_NEW As Long

Dim VAC_TOTAL As Long
Dim VAC_NEW As Long

' Prediabetes
Dim PD_TOTAL As Long
Dim PD_NEW As Long
Dim PDDUHC As Long

Dim INSOMI_TOTAL As Long
Dim INSOMI_NEW As Long
Dim INSODU As Long

Dim INSOMGS_TOTAL As Long
Dim INSOMGS_NEW As Long

Dim MAT_TOTAL As Long
Dim MAT_NEW As Long
Dim MCDUHC As Long

Dim CANCERV_TOTAL As Long
Dim CANCERV_NEW As Long
Dim CANDU As Long

Dim CANCERR_TOTAL As Long
Dim CANCERR_NEW As Long

Dim PAIN_TOTAL As Long
Dim PAIN_NEW As Long
Dim PAINDU As Long

Dim PAINGS_TOTAL As Long
Dim PAINGS_NEW As Long

Dim medrec_TOTAL As Long
Dim medrec_NEW As Long
Dim MEDDHC As Long

Dim pt_TOTAL As Long
Dim pt_NEW As Long
Dim PTDUHC As Long
Dim pt_disc As Long

Dim mind_TOTAL As Long
Dim mind_NEW As Long

Dim MISCEL_TOTAL As Long
Dim MISCEL_NEW As Long
Dim INJECT_TOTAL As Long
Dim INJECT_NEW As Long
Dim BPCHECK_TOTAL As Long
Dim BPCHECK_NEW As Long
Dim SWAB_TOTAL As Long
Dim SWAB_NEW As Long
Dim LAB_TOTAL As Long
Dim LAB_NEW As Long
Dim FORM_TOTAL As Long
Dim FORM_NEW As Long
Dim RNHW_TOTAL As Long
Dim RNHW_NEW As Long

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
DMGINI_TOTAL = Application.WorksheetFunction.Sum(Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "DMGINI"), Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "DMGFU"), Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "DMGDHC"))
'DM General - # of New Patients Included in Column A
DMGINI_NEW = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "DMGINI")
'DM General - # of Home Visits / DUHC Included in Column A
DMGDHC = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "DMGDHC")

'DM GS - # of Patient Visits
DMGROUP_TOTAL = Application.WorksheetFunction.Sum(Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "DMSM"), Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "DMCM"))
'DM GS - # of New Patients Included in Column A
Call TrueStatusMAT("AA", "DMSM", "DMCM*", "", "")
DMGROUP_NEW = Application.WorksheetFunction.CountIf(Sheets("temp").Range("AA:AA"), ">0")

'DM FV - # of Patient Visits
DMFOCUS_TOTAL = Application.WorksheetFunction.Sum(Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "DMFINI"), Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "DMFRN1"), Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "DMFFU"), Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "DMFRN2"), Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "DMFDHC"), Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "DMFRNP"))
'DM Initial - # of New Patients Included in Column A
DMFOCUS_NEW = Application.WorksheetFunction.Sum(Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "DMFINI"), Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "DMFRN1"))
'DM Initial - # of Home Visits / DUHC Included in Column A
DMFDHC = Application.WorksheetFunction.Sum(Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "DMFDHC"), Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "DMFRNP"))

'LH - # of Patient Visits
LH_TOTAL = Application.WorksheetFunction.Sum(Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "LHINI"), Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "LHFU"), Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "LHDUHC"))
'LH - # of New Patients Included in Column A
LH_NEW = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "LHINI")
'LH - # of Home Visits / DUHC Included in Column A
LHDUHC = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "LHDUHC")

'HTN General - # of Patient Visits
HTNE_TOTAL = Application.WorksheetFunction.Sum(Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "HTNEIN"), Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "HTNEFU"), Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "HTNEDU"))
'HTN General - # of New Patients Included in Column A
HTNE_NEW = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "HTNEIN")
'HTN General - # of Home Visits / DUHC Included in Column A
HTNEDU = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "HTNEDU")

'HTN GS VISIT - # of Patient Visits
HTNGS_TOTAL = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "HTNGS")
'HTN VISIT - # of New Patients Included in Column A
Call TrueStatus("AB", "HTNGS")
HTNGS_NEW = Application.WorksheetFunction.CountIf(Sheets("temp").Range("AB:AB"), ">0")

'HTN FV - # of Patient Visits
HTNFV_TOTAL = Application.WorksheetFunction.Sum(Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "HTNNIN"), Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "HTNNFU"), Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "HTNNDU"))
'HTN FV Initial - # of New Patients Included in Column A
HTNFV_NEW = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "HTNNIN")
'HTN FV Initial - # of Home Visits / DUHC Included in Column A
HTNNDU = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "HTNNDU")
'HTN Newly Diagnosed - # of Newly Diagnosed Patients
HTNNDS = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "HTNNDS")

'CHOL General - # of Patient Visits
CHOLG_TOTAL = Application.WorksheetFunction.Sum(Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "CHOLG1"), Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "CHOLG2"), Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "CHOLDU"))
'CHOL General - # of New Patients Included in Column A
CHOLG_NEW = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "CHOLG1")
'CHOL General - # of Home Visits / DUHC Included in Column A
CHOLDU = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "CHOLDU")

'CHOL FV - # of Patient Visits
CHOLFV_TOTAL = Application.WorksheetFunction.Sum(Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "CHOLF1"), Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "CHOLF2"), Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "CHOLFD"))
'CHOL FV Initial - # of New Patients Included in Column A
CHOLFV_NEW = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "CHOLF1")
'CHOL FV Initial - # of Home Visits / DUHC Included in Column A
CHOLFD = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "CHOLFD")

'CHOL GS VISIT - # of Patient Visits
CHOLGS_TOTAL = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "CHOLGS")
'CHOL VISIT - # of New Patients Included in Column A
Call TrueStatus("C", "CHOLGS")
CHOLGS_NEW = Application.WorksheetFunction.CountIf(Sheets("temp").Range("C:C"), ">0")

'CHF - # of Patient Visits
CHF_TOTAL = Application.WorksheetFunction.Sum(Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "CHFINI"), Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "CHFFU"), Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "CHFDU"))
'CHF - # of New Patients Included in Column A
CHF_NEW = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "CHFINI")
'CHF - # of Home Visits / DUHC Included in Column A
CHFDU = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "CHFDU")

'MH indiv - # of Patient Visits
MH_TOTAL = Application.WorksheetFunction.Sum(Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "MHINI"), Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "MHFU"), Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "MHDUHC"))
'MH indiv - # of New Patients Included in Column A
MH_NEW = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "MHINI")
'MH indiv - # of Home Visits / DUHC Included in Column A
MHDUHC = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "MHDUHC")

'MHGS - # of Patient Visits
MHDEP_TOTAL = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "MHDEPG")
MHMBSG_TOTAL = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "MHMBSG")
MHMBCG_TOTAL = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "MHMBCG")
MHANXG_TOTAL = Application.WorksheetFunction.SumIf(Sheets("billings").Range("J:J"), "MHANXG", Sheets("billings").Range("L:L"))

'Mental health GS - # of New Patients Included in Column A
Call TrueStatus("D", "MHDEPG")
MHDEP_NEW = Application.WorksheetFunction.CountIf(Sheets("temp").Range("D:D"), ">0")
Call TrueStatus("E", "MHMBSG")
MHMBSG_NEW = Application.WorksheetFunction.CountIf(Sheets("temp").Range("E:E"), ">0")
Call TrueStatus("F", "MHMBCG")
MHMBCG_NEW = Application.WorksheetFunction.CountIf(Sheets("temp").Range("F:F"), ">0")
Call TrueStatus("G", "MHANXG")
MHANXG_NEW = Application.WorksheetFunction.CountIf(Sheets("temp").Range("G:G"), ">0")

'SENIORS indiv - # of Patient Visits
SEN_TOTAL = Application.WorksheetFunction.Sum(Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "SENINI"), Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "SENFU"), Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "SENDU"), Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "SENDU2"))
'SENIORS indiv - # of New Patients Included in Column A
SEN_NEW = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "SENINI")
'SENIORS indiv - # of Home Visits / DUHC Included in Column A
SENDU = Application.WorksheetFunction.Sum(Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "SENDU"), Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "SENDU2"))

'SENIORS GS - # of Patient Visits
SENMGS_TOTAL = Application.WorksheetFunction.SumIf(Sheets("billings").Range("J:J"), "SENMGS", Sheets("billings").Range("L:L"))

SENEGS_TOTAL = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "SENEGS")
SENFGS_TOTAL = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "SENFGS")
SENNGS_TOTAL = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "SENNGS")
SENCGS_TOTAL = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "SENCGS")

'Senior GS - # of New Patients Included in Column A
Call TrueStatus("H", "SENMGS")
SENMGS_NEW = Application.WorksheetFunction.CountIf(Sheets("temp").Range("H:H"), ">0")
Call TrueStatus("I", "SENEGS")
SENEGS_NEW = Application.WorksheetFunction.CountIf(Sheets("temp").Range("I:I"), ">0")
Call TrueStatus("J", "SENFGS")
SENFGS_NEW = Application.WorksheetFunction.CountIf(Sheets("temp").Range("J:J"), ">0")
Call TrueStatus("K", "SENNGS")
SENNGS_NEW = Application.WorksheetFunction.CountIf(Sheets("temp").Range("K:K"), ">0")
Call TrueStatus("L", "SENCGS")
SENCGS_NEW = Application.WorksheetFunction.CountIf(Sheets("temp").Range("L:L"), ">0")

'SENIORS Social Club - # of Patient Visits
SENSOC_TOTAL = Application.WorksheetFunction.Sum(Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "SENSOC"), Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "SENSO2"), Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "SENSOD"))
'SENIORS Social Club - # of New Patients Included in Column A
SENSOC_NEW = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "SENSOC")
'SENIORS Social Club - # of Home Visits / DUHC Included in Column A
SENSOD = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "SENSOD")

'HEALTHY LIVING indiv - # of Patient Visits
HLI_TOTAL = Application.WorksheetFunction.Sum(Application.WorksheetFunction.SumIf(Sheets("billings").Range("J:J"), "HLINI", Sheets("billings").Range("L:L")), Application.WorksheetFunction.SumIf(Sheets("billings").Range("J:J"), "HLFU", Sheets("billings").Range("L:L")), Application.WorksheetFunction.SumIf(Sheets("billings").Range("J:J"), "HLDUHC", Sheets("billings").Range("L:L")))

'HEALTHY LIVING indiv - # of New Patients Included in Column A
HLI_NEW = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "HLINI")
'HEALTHY LIVING indiv - # of Home Visits / DUHC Included in Column A
HLDUHC = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "HLDUHC")

'SMOKING CESSATION - # of Patient Visits
SC_TOTAL = Application.WorksheetFunction.Sum(Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "SCINI"), Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "SCFU"), Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "SCDUHC"))
'SMOKING CESSATION - # of New Patients Included in Column A
SC_NEW = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "SCINI")
'SMOKING CESSATION - # of Home Visits / DUHC Included in Column A
SCDUHC = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "SCDUHC")

'WEIGHT MANAGEMENT - # of Patient Visits
WM_TOTAL = Application.WorksheetFunction.Sum(Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "WINI"), Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "WFU"), Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "WDUHC"))
'WEIGHT MANAGEMENT - # of New Patients Included in Column A
WM_NEW = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "WINI")
'WEIGHT MANAGEMENT - # of Home Visits / DUHC Included in Column A
WDUHC = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "WDUHC")

'WEIGHT MANAGEMENT GS - # of Patient Visits
WMGS_TOTAL = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "WGS")
'weight management GS - # of New Patients Included in Column A
Call TrueStatus("M", "WGS")
WMGS_NEW = Application.WorksheetFunction.CountIf(Sheets("temp").Range("M:M"), ">0")

'VACCINE - # of Patient Visits
VAC_TOTAL = Application.WorksheetFunction.Sum(Application.WorksheetFunction.SumIf(Sheets("billings").Range("J:J"), "RNVAC", Sheets("billings").Range("L:L")), Application.WorksheetFunction.SumIf(Sheets("billings").Range("J:J"), "IMMUNR", Sheets("billings").Range("L:L")))
'VACCINE - # of New Patients Included in Column A
Call TrueStatusCANCER("N", "RNVAC", "IMMUNR", "RNVAC")
VAC_NEW = Application.WorksheetFunction.CountIf(Sheets("temp").Range("N:N"), ">0")

'Prediabetes - # of Patient Visits
PD_TOTAL = Application.WorksheetFunction.Sum(Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "PDINI"), Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "PDFU"), Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "PDDUHC"))
'Prediabetes - # of New Patients Included in Column A
PD_NEW = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "PDINI")
'Prediabetes - # of Home Visits / DUHC Included in Column A
PDDUHC = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "PDDUHC")

'Insomnia indiv - # of Patient Visits
INSOMI_TOTAL = Application.WorksheetFunction.Sum(Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "INSOIN"), Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "INSOFU"), Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "INSODU"))
'Insomnia indiv - # of New Patients Included in Column A
INSOMI_NEW = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "INSOIN")
'Insomnia indiv - # of Home Visits / DUHC Included in Column A
INSODU = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "INSODU")

'insomnia GS - # of Patient Visits
INSOMGS_TOTAL = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "INSOGS")
'insomnia GS - # of New Patients Included in Column A
Call TrueStatus("O", "INSOGS")
INSOMGS_NEW = Application.WorksheetFunction.CountIf(Sheets("temp").Range("O:O"), ">0")

'MATERNAL HEALTH - # of Patient Visits
MAT_TOTAL = Application.WorksheetFunction.Sum(Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "MCP*"), Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "MCBF"), Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "MCWBC"), Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "MCGEN"), Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "MCDUHC"))
'MATERNAL HEALTH - # of New Patients Included in Column A
Call TrueStatusMAT("Z", "MCBF", "MCP*", "MCWBC", "MCGEN")
MAT_NEW = Application.WorksheetFunction.CountIf(Sheets("temp").Range("Z:Z"), ">0")
'MATERNAL HEALTH - # of Home Visits / DUHC Included in Column A
MCDUHC = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "MCDUHC")

'CANCER VISITS - # of Patient Visits
CANCERV_TOTAL = Application.WorksheetFunction.Sum(Application.WorksheetFunction.SumIf(Sheets("billings").Range("J:J"), "CERVV", Sheets("billings").Range("L:L")), Application.WorksheetFunction.SumIf(Sheets("billings").Range("J:J"), "COLV", Sheets("billings").Range("L:L")), Application.WorksheetFunction.SumIf(Sheets("billings").Range("J:J"), "BREASV", Sheets("billings").Range("L:L")), Application.WorksheetFunction.SumIf(Sheets("billings").Range("J:J"), "CANDU", Sheets("billings").Range("L:L")))
'CANCER VISITS - # of New Patients Included in Column A
Call TrueStatusCANCER("P", "COLV", "CERV", "BREASV*")
CANCERV_NEW = Application.WorksheetFunction.CountIf(Sheets("temp").Range("P:P"), ">0")
'Cancer visits - # of Home Visits / DUHC Included in Column A
CANDU = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "CANDU")

'CANCER REVIEW - # of Patient Visits
CANCERR_TOTAL = Application.WorksheetFunction.Sum(Application.WorksheetFunction.SumIf(Sheets("billings").Range("J:J"), "CERVR", Sheets("billings").Range("L:L")), Application.WorksheetFunction.SumIf(Sheets("billings").Range("J:J"), "COLR", Sheets("billings").Range("L:L")), Application.WorksheetFunction.SumIf(Sheets("billings").Range("J:J"), "BREASR", Sheets("billings").Range("L:L")))

'PAIN MANAGEMENT - # of Patient Visits
PAIN_TOTAL = Application.WorksheetFunction.Sum(Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "PAININ"), Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "PAINFU"), Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "PAINDU"))
'PAIN MANAGEMENT - # of New Patients Included in Column A
PAIN_NEW = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "PAININ")
'PAIN MANAGEMENT - # of Home Visits / DUHC Included in Column A
PAINDU = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "PAINDU")

'PAIN MANAGEMENT - # of Patient Visits
PAINGS_TOTAL = Application.WorksheetFunction.SumIf(Sheets("billings").Range("J:J"), "PAINGS", Sheets("billings").Range("L:L"))

'PAIN MANAGEMENT - # of New Patients Included in Column A
Call TrueStatus("Q", "PAINGS")
PAINGS_NEW = Application.WorksheetFunction.CountIf(Sheets("temp").Range("Q:Q"), ">0")

'medrec - # of Patient Visits
medrec_TOTAL = Application.WorksheetFunction.Sum(Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "MEDINI"), Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "MEDFU"), Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "MEDDHC"))
'medrec - # of New Patients Included in Column A
medrec_NEW = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "MEDINI")
'medrec - # of Home Visits / DUHC Included in Column A
MEDDHC = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "MEDDHC")

'PT - # of Patient Visits
pt_TOTAL = Application.WorksheetFunction.Sum(Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "PTINI"), Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "PTFU"), Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "PTDUHC"))
'PT - # of New Patients Included in Column A
pt_NEW = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "PTINI")
'PT - # of Home Visits / DUHC Included in Column A
PTDUHC = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "PTDUHC")
'PT - # of Discharged Patients
pt_disc = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "PTDISC")

'Mindful Eating GS - # of Patient Visits
mind_TOTAL = Application.WorksheetFunction.SumIf(Sheets("billings").Range("J:J"), "MINDGS", Sheets("billings").Range("L:L"))
'Mindful Eating GS - # of New Patients Included in Column A
Call TrueStatus("R", "MINDGS")
mind_NEW = Application.WorksheetFunction.CountIf(Sheets("temp").Range("R:R"), ">0")

'MISCEL RN - # of Patient Visits
MISCEL_TOTAL = Application.WorksheetFunction.SumIf(Sheets("billings").Range("J:J"), "MISCEL", Sheets("billings").Range("L:L"))

'MISCEL RN - # of New Patients Included in Column A
Call TrueStatus("S", "MISCEL")
MISCEL_NEW = Application.WorksheetFunction.CountIf(Sheets("temp").Range("S:S"), ">0")

'INJECT RN - # of Patient Visits
INJECT_TOTAL = Application.WorksheetFunction.SumIf(Sheets("billings").Range("J:J"), "INJECT", Sheets("billings").Range("L:L"))

'INJECT RN - # of New Patients Included in Column A
Call TrueStatus("T", "INJECT")
INJECT_NEW = Application.WorksheetFunction.CountIf(Sheets("temp").Range("T:T"), ">0")

'BPCHECK RN - # of Patient Visits
BPCHECK_TOTAL = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "BPCHEK")

'BPCHECK RN - # of New Patients Included in Column A
Call TrueStatus("U", "BPCHEK")
BPCHECK_NEW = Application.WorksheetFunction.CountIf(Sheets("temp").Range("U:U"), ">0")

'SWAB RN - # of Patient Visits
SWAB_TOTAL = Application.WorksheetFunction.SumIf(Sheets("billings").Range("J:J"), "SWABR", Sheets("billings").Range("L:L"))

'SWAB RN - # of New Patients Included in Column A
Call TrueStatus("V", "SWABR")
SWAB_NEW = Application.WorksheetFunction.CountIf(Sheets("temp").Range("V:V"), ">0")

'LAB Review RN - # of Patient Visits
LAB_TOTAL = Application.WorksheetFunction.SumIf(Sheets("billings").Range("J:J"), "LABR", Sheets("billings").Range("L:L"))

'LAB Review RN - # of New Patients Included in Column A
Call TrueStatus("W", "LABR")
LAB_NEW = Application.WorksheetFunction.CountIf(Sheets("temp").Range("W:W"), ">0")

'FORM - # of Patient Visits
FORM_TOTAL = Application.WorksheetFunction.SumIf(Sheets("billings").Range("J:J"), "FORM", Sheets("billings").Range("L:L"))

'FORM - # of New Patients Included in Column A
Call TrueStatus("X", "FORM")
FORM_NEW = Application.WorksheetFunction.CountIf(Sheets("temp").Range("X:X"), ">0")

'RN Height/Weight Reveiew - # of Patient Visits
RNHW_TOTAL = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "RNHW")

'RN Height/Weight Review - # of New Patients Included in Column A
Call TrueStatus("Y", "RNHW")
RNHW_NEW = Application.WorksheetFunction.CountIf(Sheets("temp").Range("Y:Y"), ">0")


'ENTERING VALUES INTO APPROPRIATE CELLS in "IndividualSummary-Detailed"
Sheets("IndividualSummary-Detailed").Select

Sheets("IndividualSummary-Detailed").Range("B8").Value = DMGINI_TOTAL
Sheets("IndividualSummary-Detailed").Range("c8").Value = DMGINI_NEW
Sheets("IndividualSummary-Detailed").Range("d8").Value = DMGDHC

Sheets("IndividualSummary-Detailed").Range("B9").Value = DMGROUP_TOTAL
Sheets("IndividualSummary-Detailed").Range("C9").Value = DMGROUP_NEW

Sheets("IndividualSummary-Detailed").Range("B10").Value = DMFOCUS_TOTAL
Sheets("IndividualSummary-Detailed").Range("c10").Value = DMFOCUS_NEW
Sheets("IndividualSummary-Detailed").Range("d10").Value = DMFDHC

Sheets("IndividualSummary-Detailed").Range("B11").Value = LH_TOTAL
Sheets("IndividualSummary-Detailed").Range("c11").Value = LH_NEW
Sheets("IndividualSummary-Detailed").Range("d11").Value = LHDUHC

Sheets("IndividualSummary-Detailed").Range("B13").Value = HTNE_TOTAL
Sheets("IndividualSummary-Detailed").Range("c13").Value = HTNE_NEW
Sheets("IndividualSummary-Detailed").Range("d13").Value = HTNEDU

Sheets("IndividualSummary-Detailed").Range("B14").Value = HTNFV_TOTAL
Sheets("IndividualSummary-Detailed").Range("c14").Value = HTNFV_NEW
Sheets("IndividualSummary-Detailed").Range("d14").Value = HTNNDU

Sheets("IndividualSummary-Detailed").Range("j10").Value = HTNNDS

Sheets("IndividualSummary-Detailed").Range("B15").Value = HTNGS_TOTAL
Sheets("IndividualSummary-Detailed").Range("c15").Value = HTNGS_NEW

Sheets("IndividualSummary-Detailed").Range("B16").Value = CHOLG_TOTAL
Sheets("IndividualSummary-Detailed").Range("c16").Value = CHOLG_NEW
Sheets("IndividualSummary-Detailed").Range("d16").Value = CHOLDU

Sheets("IndividualSummary-Detailed").Range("B17").Value = CHOLFV_TOTAL
Sheets("IndividualSummary-Detailed").Range("c17").Value = CHOLFV_NEW
Sheets("IndividualSummary-Detailed").Range("d17").Value = CHOLFD

Sheets("IndividualSummary-Detailed").Range("B18").Value = CHOLGS_TOTAL
Sheets("IndividualSummary-Detailed").Range("c18").Value = CHOLGS_NEW

Sheets("IndividualSummary-Detailed").Range("B19").Value = CHF_TOTAL
Sheets("IndividualSummary-Detailed").Range("c19").Value = CHF_NEW
Sheets("IndividualSummary-Detailed").Range("d19").Value = CHFDU

Sheets("IndividualSummary-Detailed").Range("B21").Value = MH_TOTAL
Sheets("IndividualSummary-Detailed").Range("c21").Value = MH_NEW
Sheets("IndividualSummary-Detailed").Range("d21").Value = MHDUHC

Sheets("IndividualSummary-Detailed").Range("B22").Value = MHMBSG_TOTAL
Sheets("IndividualSummary-Detailed").Range("B23").Value = MHMBCG_TOTAL
Sheets("IndividualSummary-Detailed").Range("B24").Value = MHANXG_TOTAL
Sheets("IndividualSummary-Detailed").Range("B25").Value = MHDEP_TOTAL

Sheets("IndividualSummary-Detailed").Range("C22").Value = MHMBSG_NEW
Sheets("IndividualSummary-Detailed").Range("C23").Value = MHMBCG_NEW
Sheets("IndividualSummary-Detailed").Range("C24").Value = MHANXG_NEW
Sheets("IndividualSummary-Detailed").Range("C25").Value = MHDEP_NEW

Sheets("IndividualSummary-Detailed").Range("B28").Value = SEN_TOTAL
Sheets("IndividualSummary-Detailed").Range("c28").Value = SEN_NEW
Sheets("IndividualSummary-Detailed").Range("d28").Value = SENDU

Sheets("IndividualSummary-Detailed").Range("B29").Value = SENMGS_TOTAL
Sheets("IndividualSummary-Detailed").Range("B30").Value = SENEGS_TOTAL
Sheets("IndividualSummary-Detailed").Range("B31").Value = SENFGS_TOTAL
Sheets("IndividualSummary-Detailed").Range("B32").Value = SENNGS_TOTAL
Sheets("IndividualSummary-Detailed").Range("B33").Value = SENCGS_TOTAL

Sheets("IndividualSummary-Detailed").Range("C29").Value = SENMGS_NEW
Sheets("IndividualSummary-Detailed").Range("C30").Value = SENEGS_NEW
Sheets("IndividualSummary-Detailed").Range("C31").Value = SENFGS_NEW
Sheets("IndividualSummary-Detailed").Range("C32").Value = SENNGS_NEW
Sheets("IndividualSummary-Detailed").Range("C33").Value = SENCGS_NEW

Sheets("IndividualSummary-Detailed").Range("B34").Value = SENSOC_TOTAL
Sheets("IndividualSummary-Detailed").Range("c34").Value = SENSOC_NEW
Sheets("IndividualSummary-Detailed").Range("d34").Value = SENSOD

Sheets("IndividualSummary-Detailed").Range("B35").Value = medrec_TOTAL
Sheets("IndividualSummary-Detailed").Range("c35").Value = medrec_NEW
Sheets("IndividualSummary-Detailed").Range("d35").Value = MEDDHC

Sheets("IndividualSummary-Detailed").Range("B37").Value = HLI_TOTAL
Sheets("IndividualSummary-Detailed").Range("c37").Value = HLI_NEW
Sheets("IndividualSummary-Detailed").Range("d37").Value = HLDUHC

Sheets("IndividualSummary-Detailed").Range("B38").Value = SC_TOTAL
Sheets("IndividualSummary-Detailed").Range("c38").Value = SC_NEW
Sheets("IndividualSummary-Detailed").Range("d38").Value = SCDUHC

Sheets("IndividualSummary-Detailed").Range("B39").Value = WM_TOTAL
Sheets("IndividualSummary-Detailed").Range("c39").Value = WM_NEW
Sheets("IndividualSummary-Detailed").Range("d39").Value = WDUHC

Sheets("IndividualSummary-Detailed").Range("B40").Value = WMGS_TOTAL
Sheets("IndividualSummary-Detailed").Range("C40").Value = WMGS_NEW

Sheets("IndividualSummary-Detailed").Range("B41").Value = VAC_TOTAL
Sheets("IndividualSummary-Detailed").Range("C41").Value = VAC_NEW

Sheets("IndividualSummary-Detailed").Range("B42").Value = PD_TOTAL
Sheets("IndividualSummary-Detailed").Range("c42").Value = PD_NEW
Sheets("IndividualSummary-Detailed").Range("d42").Value = PDDUHC

Sheets("IndividualSummary-Detailed").Range("B43").Value = INSOMI_TOTAL
Sheets("IndividualSummary-Detailed").Range("c43").Value = INSOMI_NEW
Sheets("IndividualSummary-Detailed").Range("d43").Value = INSODU

Sheets("IndividualSummary-Detailed").Range("B44").Value = INSOMGS_TOTAL
Sheets("IndividualSummary-Detailed").Range("C44").Value = INSOMGS_NEW

Sheets("IndividualSummary-Detailed").Range("B45").Value = MAT_TOTAL
Sheets("IndividualSummary-Detailed").Range("C45").Value = MAT_NEW
Sheets("IndividualSummary-Detailed").Range("D45").Value = MCDUHC

Sheets("IndividualSummary-Detailed").Range("B46").Value = CANCERV_TOTAL
Sheets("IndividualSummary-Detailed").Range("B47").Value = CANCERR_TOTAL

Sheets("IndividualSummary-Detailed").Range("C46").Value = CANCERV_NEW

Sheets("IndividualSummary-Detailed").Range("D46").Value = CANDU

Sheets("IndividualSummary-Detailed").Range("B48").Value = PAIN_TOTAL
Sheets("IndividualSummary-Detailed").Range("c48").Value = PAIN_NEW
Sheets("IndividualSummary-Detailed").Range("d48").Value = PAINDU

Sheets("IndividualSummary-Detailed").Range("B49").Value = PAINGS_TOTAL
Sheets("IndividualSummary-Detailed").Range("c49").Value = PAINGS_NEW

Sheets("IndividualSummary-Detailed").Range("B50").Value = pt_TOTAL
Sheets("IndividualSummary-Detailed").Range("c50").Value = pt_NEW
Sheets("IndividualSummary-Detailed").Range("d50").Value = PTDUHC

Sheets("IndividualSummary-Detailed").Range("j13").Value = pt_disc

Sheets("IndividualSummary-Detailed").Range("B51").Value = mind_TOTAL
Sheets("IndividualSummary-Detailed").Range("c51").Value = mind_NEW

Sheets("IndividualSummary-Detailed").Range("B53").Value = MISCEL_TOTAL
Sheets("IndividualSummary-Detailed").Range("c53").Value = MISCEL_NEW
Sheets("IndividualSummary-Detailed").Range("B54").Value = INJECT_TOTAL
Sheets("IndividualSummary-Detailed").Range("c54").Value = INJECT_NEW
Sheets("IndividualSummary-Detailed").Range("B55").Value = BPCHECK_TOTAL
Sheets("IndividualSummary-Detailed").Range("c55").Value = BPCHECK_NEW
Sheets("IndividualSummary-Detailed").Range("B56").Value = SWAB_TOTAL
Sheets("IndividualSummary-Detailed").Range("c56").Value = SWAB_NEW
Sheets("IndividualSummary-Detailed").Range("B57").Value = LAB_TOTAL
Sheets("IndividualSummary-Detailed").Range("c57").Value = LAB_NEW
Sheets("IndividualSummary-Detailed").Range("B58").Value = FORM_TOTAL
Sheets("IndividualSummary-Detailed").Range("c58").Value = FORM_NEW
Sheets("IndividualSummary-Detailed").Range("B59").Value = RNHW_TOTAL
Sheets("IndividualSummary-Detailed").Range("c59").Value = RNHW_NEW

Sheets("IndividualSummary-Detailed").Range("B61").Value = HLI_TOTAL + FORM_TOTAL + MISCEL_TOTAL + INJECT_TOTAL + SWAB_TOTAL + LAB_TOTAL
Sheets("IndividualSummary-Detailed").Range("C61").Value = HLI_NEW + FORM_NEW + MISCEL_NEW + INJECT_NEW + SWAB_NEW + LAB_NEW
Sheets("IndividualSummary-Detailed").Range("D61").Value = HLDUHC


'ENTERING VALUES INTO APPROPRIATE CELLS in "IndividualSummary-Basic"
Sheets("IndividualSummary-Basic").Select

Sheets("IndividualSummary-Basic").Range("B9").Value = DMGINI_TOTAL + DMGROUP_TOTAL + DMFOCUS_TOTAL
Sheets("IndividualSummary-Basic").Range("C9").Value = DMGINI_NEW + DMGROUP_NEW + DMFOCUS_NEW
Sheets("IndividualSummary-Basic").Range("D9").Value = DMGDHC + DMFDHC

Sheets("IndividualSummary-Basic").Range("B10").Value = LH_TOTAL
Sheets("IndividualSummary-Basic").Range("c10").Value = LH_NEW
Sheets("IndividualSummary-Basic").Range("d10").Value = LHDUHC

Sheets("IndividualSummary-Basic").Range("B12").Value = HTNE_TOTAL + HTNFV_TOTAL + HTNGS_TOTAL
Sheets("IndividualSummary-Basic").Range("c12").Value = HTNE_NEW + HTNFV_NEW + HTNGS_NEW
Sheets("IndividualSummary-Basic").Range("d12").Value = HTNEDU + HTNNDU

Sheets("IndividualSummary-Basic").Range("B13").Value = CHOLG_TOTAL + CHOLFV_TOTAL + CHOLGS_TOTAL
Sheets("IndividualSummary-Basic").Range("c13").Value = CHOLG_NEW + CHOLFV_NEW + CHOLGS_NEW
Sheets("IndividualSummary-Basic").Range("d13").Value = CHOLDU + CHOLFD

Sheets("IndividualSummary-Basic").Range("B14").Value = CHF_TOTAL
Sheets("IndividualSummary-Basic").Range("c14").Value = CHF_NEW
Sheets("IndividualSummary-Basic").Range("d14").Value = CHFDU

Sheets("IndividualSummary-Basic").Range("B15").Value = MH_TOTAL + MHMBSG_TOTAL + MHMBCG_TOTAL + MHANXG_TOTAL + MHDEP_TOTAL
Sheets("IndividualSummary-Basic").Range("c15").Value = MH_NEW + MHMBSG_NEW + MHMBCG_NEW + MHANXG_NEW + MHDEP_NEW
Sheets("IndividualSummary-Basic").Range("d15").Value = MHDUHC

Sheets("IndividualSummary-Basic").Range("B18").Value = SEN_TOTAL + SENMGS_TOTAL + SENEGS_TOTAL + SENFGS_TOTAL + SENNGS_TOTAL + SENCGS_TOTAL + SENSOC_TOTAL
Sheets("IndividualSummary-Basic").Range("C18").Value = SEN_NEW + SENMGS_NEW + SENEGS_NEW + SENFGS_NEW + SENNGS_NEW + SENCGS_NEW + SENSOC_NEW
Sheets("IndividualSummary-Basic").Range("D18").Value = SENDU + SENSOD

Sheets("IndividualSummary-Basic").Range("B20").Value = SC_TOTAL
Sheets("IndividualSummary-Basic").Range("c20").Value = SC_NEW
Sheets("IndividualSummary-Basic").Range("d20").Value = SCDUHC

Sheets("IndividualSummary-Basic").Range("B21").Value = WM_TOTAL + WMGS_TOTAL + mind_TOTAL
Sheets("IndividualSummary-Basic").Range("c21").Value = WM_NEW + WMGS_NEW + mind_NEW
Sheets("IndividualSummary-Basic").Range("d21").Value = WDUHC

Sheets("IndividualSummary-Basic").Range("B22").Value = VAC_TOTAL
Sheets("IndividualSummary-Basic").Range("C22").Value = VAC_NEW

Sheets("IndividualSummary-Basic").Range("B23").Value = PD_TOTAL
Sheets("IndividualSummary-Basic").Range("c23").Value = PD_NEW
Sheets("IndividualSummary-Basic").Range("d23").Value = PDDUHC

Sheets("IndividualSummary-Basic").Range("B24").Value = INSOMI_TOTAL + INSOMGS_TOTAL
Sheets("IndividualSummary-Basic").Range("c24").Value = INSOMI_NEW + INSOMGS_NEW
Sheets("IndividualSummary-Basic").Range("d24").Value = INSODU

Sheets("IndividualSummary-Basic").Range("B25").Value = MAT_TOTAL
Sheets("IndividualSummary-Basic").Range("C25").Value = MAT_NEW
Sheets("IndividualSummary-Basic").Range("D25").Value = MCDUHC

Sheets("IndividualSummary-Basic").Range("B26").Value = CANCERV_TOTAL + CANCERR_TOTAL
Sheets("IndividualSummary-Basic").Range("C26").Value = CANCERV_NEW
Sheets("IndividualSummary-Basic").Range("D26").Value = CANDU

Sheets("IndividualSummary-Basic").Range("B27").Value = PAIN_TOTAL + PAINGS_TOTAL
Sheets("IndividualSummary-Basic").Range("c27").Value = PAIN_NEW + PAINGS_NEW
Sheets("IndividualSummary-Basic").Range("d27").Value = PAINDU

Sheets("IndividualSummary-Basic").Range("B29").Value = HLI_TOTAL + FORM_TOTAL + MISCEL_TOTAL + INJECT_TOTAL + SWAB_TOTAL + LAB_TOTAL
Sheets("IndividualSummary-Basic").Range("C29").Value = HLI_NEW + FORM_NEW + MISCEL_NEW + INJECT_NEW + SWAB_NEW + LAB_NEW
Sheets("IndividualSummary-Basic").Range("D29").Value = HLDUHC

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

