Public Sub clear()
    Worksheets("billings").Range("A2:V9999").ClearContents

    Worksheets("IndividualSummary-Detailed").Range("B9:D66").ClearContents

    Worksheets("IndividualSummary-Detailed").Range("B2").ClearContents
    Worksheets("IndividualSummary-Detailed").Range("J10").ClearContents
    Worksheets("IndividualSummary-Detailed").Range("J13").ClearContents

    Worksheets("IndividualSummary-Basic").Range("B9:D36").ClearContents
    Worksheets("IndividualSummary-Basic").Range("B2").ClearContents

    Worksheets("temp").Cells.ClearContents
End Sub

Public Sub CalcIndividualResp()

    ' ----------------------
    ' Declare all Variables
    ' ----------------------

    ' Home to Stay
    Dim HTS_TOTAL as Long, HTS_NEW  as Long, HTSDUH as Long

    ' Post Hospital Discharge
    Dim HTSPHD as Long
    ' External/Internal Referral
    Dim HTSREF as Long
    ' Med Review
    Dim HTSMR_TOTAL as Long, HTSMR_NEW as Long
    ' Dietitian
    Dim HTSRD_TOTAL as Long, HTSRD_NEW as Long
    ' Occupational Therapist
    Dim HTSOT_TOTAL as Long, HTSOT_NEW as Long
    ' Social Worker
    Dim HTSSW_TOTAL as Long, HTSSW_NEW  as Long
    ' Personal Support Worker
    Dim HTSPSW_TOTAL as Long

    Dim HLYOGA_NEW As Long, HLYOGA_TOTAL As Long ' Healthy Living - Yoga

    ' New Billing Codes - 2016-12-28

    ' CHF
    Dim CHG_TOTAL As Long, CHG_NEW As Long, CHGDU As Long

    ' Health Links
    Dim HK_TOTAL As Long, HK_NEW As Long, HKDUHC As Long

    ' CKD
    Dim CKDGRP_NEW As Long, CKDGRP_TOTAL As Long
    Dim CKD_NEW As Long, CKDGEN_TOTAL As Long, CKDDUH As Long, CKD_TOTAL As Long

    ' End of new billing codes

    ' Diabetes
    Dim DMGINI_TOTAL As Long, DMGINI_NEW As Long, DMGDHC As Long
    Dim DMGROUP_TOTAL As Long, DMGROUP_NEW As Long
    Dim DMFOCUS_TOTAL As Long, DMFOCUS_NEW As Long, DMFDHC As Long

    ' Lung Health
    Dim LH_TOTAL As Long, LH_NEW As Long, LHDUHC As Long

    ' Hypertension
    ' General
    Dim HTNE_TOTAL As Long, HTNE_NEW As Long, HTNEDU As Long, HTNENS As Long

    ' Focus
    Dim HTNNIN As Long, HTNNFU As Long, HTNNDU As Long, HTNNNS As Long

    ' Group
    Dim HTNGS_TOTAL As Long, HTNGS_NEW As Long
    Dim HTNFV_TOTAL As Long, HTNFV_NEW As Long, HTNNDS As Long

    ' Cholesterol
    Dim CHOLG_TOTAL As Long, CHOLG_NEW As Long, CHOLDU As Long

    Dim CHOLFV_TOTAL As Long, CHOLFV_NEW As Long, CHOLFD As Long

    Dim CHOLGS_TOTAL As Long, CHOLGS_NEW As Long

    ' CHF
    Dim CHF_TOTAL As Long, CHF_NEW As Long, CHFDU As Long

    ' Mental Health
    Dim MH_TOTAL As Long, MH_NEW As Long, MHDUHC As Long, MHDISC As Long

    Dim MHDEP_TOTAL As Long, MHDEP_NEW As Long
    ' Dim MHDEPN As Long ' No Show

    ' Mindfulness Based Stress Reduction
    Dim MHMBSG_TOTAL As Long, MHMBSG_NEW As Long
    ' Dim MHMBMN ' No Show

    ' Mindfulness Based Cognitivate Therapy
    Dim MHMBCG_TOTAL As Long, MHMBCG_NEW As Long
    ' Dim MGMBCN ' No Show

    ' Anxiety
    Dim MHANXG_TOTAL As Long, MHANXG_NEW As Long
    ' Dim MHANXN ' No Show'

    ' Seniors
    Dim SEN_TOTAL As Long, SEN_NEW As Long, SENDU As Long 

    Dim SENNS As Long ' No Show
    Dim SENGNS As Long ' Group No Show

    Dim SENMGS_TOTAL As Long, SENMGS_NEW As Long
    Dim SENEGS_TOTAL As Long, SENEGS_NEW As Long
    Dim SENFGS_TOTAL As Long, SENFGS_NEW As Long
    Dim SENNGS_TOTAL As Long, SENNGS_NEW As Long
    Dim SENCGS_TOTAL As Long, SENCGS_NEW As Long

    Dim SENSOC_TOTAL As Long, SENSOC_NEW As Long, SENSOD As Long

    ' Healthy Living
    Dim HLI_TOTAL As Long, HLI_NEW As Long, HLDUHC As Long
    ' Healthy Living - Community Outreach
    Dim HLOUTR As Long

    ' Smoking Cessation
    Dim SC_TOTAL As Long, SC_NEW As Long, SCDUHC As Long

    ' Weight Mgmt
    Dim WM_TOTAL As Long, WM_NEW As Long, WDUHC As Long

    Dim WMGS_TOTAL As Long, WMGS_NEW As Long

    ' Vaccines
    Dim VAC_TOTAL As Long, VAC_NEW As Long

    ' Prediabetes
    Dim PD_TOTAL As Long, PD_NEW As Long, PDDUHC As Long

    ' Insomnia
    Dim INSOMI_TOTAL As Long, INSOMI_NEW As Long, INSODU As Long

    Dim INSOMGS_TOTAL As Long, INSOMGS_NEW As Long

    Dim MAT_TOTAL As Long, MAT_NEW As Long, MCDUHC As Long

    Dim CANCERV_TOTAL As Long, CANCERV_NEW As Long, CANDU As Long

    Dim CANCERR_TOTAL As Long, CANCERR_NEW As Long

    Dim PAIN_TOTAL As Long, PAIN_NEW As Long, PAINDU As Long

    Dim PAINGS_TOTAL As Long, PAINGS_NEW As Long

    Dim MEDREC_TOTAL As Long, MEDREC_NEW As Long, MEDDHC As Long

    Dim PT_TOTAL As Long, PT_NEW As Long, PTDUHC As Long, PTDISC As Long

    Dim MINDGS_TOTAL As Long, MINDGS_NEW As Long

    Dim MISCEL_TOTAL As Long, MISCEL_NEW As Long
    Dim INJECT_TOTAL As Long, INJECT_NEW As Long
    Dim BPCHECK_TOTAL As Long, BPCHECK_NEW As Long
    Dim SWAB_TOTAL As Long, SWAB_NEW As Long
    Dim LAB_TOTAL As Long, LAB_NEW As Long
    Dim FORM_TOTAL As Long, FORM_NEW As Long
    Dim RNHW_TOTAL As Long, RNHW_NEW As Long

    Application.ScreenUpdating = False
    Worksheets("IndividualSummary-Detailed").Range("B9:D66").ClearContents
    Worksheets("IndividualSummary-Basic").Range("B9:D33").ClearContents
    Worksheets("temp").Cells.ClearContents

    AllEntries = Application.WorksheetFunction.CountA(Sheets("billings").Range("A:A"))

    ' create unique identifier (lastname, firstname) and paste it in column R
    Dim lastname As String, firstname As String

    With Sheets("billings")
        For i = 2 To AllEntries
            lastname = .Range("A" & i).Value
            firstname = .Range("B" & i).Value
            
            .Range("R" & i).Value = lastname & ", " & firstname
            
        Next i
    End With

    'copy and paste unique patient names to temp sheet and remove duplicated
    ' this determines Column B - # of New Patients Included in Column A
    Sheets("billings").Range("R:R").Copy
    Sheets("temp").Select
    Sheets("temp").Range("A1").Select
    Sheets("temp").Paste

    Sheets("temp").Range("A1:A" & AllEntries).RemoveDuplicates Columns:=1, Header:=xlYes

    'DM General - # of Patient Visits
    DMGINI_TOTAL = Application.WorksheetFunction.Sum( _
                        Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "DMGINI"), _
                        Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "DMGFU"), _
                        Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "DMGDHC") _
                    )
    'DM General - # of New Patients Included in Column A
    DMGINI_NEW = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "DMGINI")
    'DM General - # of Home Visits / DUHC Included in Column A
    DMGDHC = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "DMGDHC")

    'DM GS - # of Patient Visits
    DMGROUP_TOTAL = Application.WorksheetFunction.Sum( _
                        Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "DMSM"), _
                        Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "DMCM") _
                    )
    'DM GS - # of New Patients Included in Column A
    Call TrueStatusMAT("AA", "DMSM", "DMCM*", "", "")
    DMGROUP_NEW = Application.WorksheetFunction.CountIf(Sheets("temp").Range("AA:AA"), ">0")

    'DM FV - # of Patient Visits
    DMFOCUS_TOTAL = Application.WorksheetFunction.Sum( _
                        Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "DMFINI"), _
                        Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "DMFRN1"), _
                        Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "DMFFU"), _
                        Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "DMFRN2"), _
                        Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "DMFDHC"), _
                        Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "DMFRNP") _
                    )
    'DM Initial - # of New Patients Included in Column A
    DMFOCUS_NEW = Application.WorksheetFunction.Sum( _
                        Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "DMFINI"), _
                        Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "DMFRN1") _
                    )
    'DM Initial - # of Home Visits / DUHC Included in Column A
    DMFDHC = Application.WorksheetFunction.Sum( _
                        Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "DMFDHC"), _
                        Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "DMFRNP") _
                    )

    '--- CKD ---
    CKDGEN_TOTAL = Application.WorksheetFunction.Sum( _
                    Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "CKDINI"), _
                    Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "CKDFU"), _
                    Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "CKDDUH"), _
                    Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "CKDSCR") _
                )
    CKD_NEW = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "CKDINI")
    CKDDUH = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "CKDDUH")

    Call TrueStatus("AC", "CKDGRP")
    CKDGRP_NEW = Application.WorksheetFunction.CountIf(Sheets("temp").Range("AB:AB"), ">0")
    CKDGRP_TOTAL = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "CKDGRP")
    CKD_TOTAL = CKDGEN_TOTAL + CKDGRP_TOTAL

    'LH - # of Patient Visits
    LH_TOTAL = Application.WorksheetFunction.Sum( _
                    Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "LHINI"), _
                    Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "LHFU"), _
                    Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "LHDUHC") _
                )
    'LH - # of New Patients Included in Column A
    LH_NEW = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "LHINI")
    'LH - # of Home Visits / DUHC Included in Column A
    LHDUHC = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "LHDUHC")

    'HTN General - # of Patient Visits
    HTNE_TOTAL = Application.WorksheetFunction.Sum( _
                    Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "HTNEIN"), _
                    Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "HTNEFU"), _
                    Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "HTNEDU") _
                )
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
    HTNFV_TOTAL = Application.WorksheetFunction.Sum( _
                    Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "HTNNIN"), _
                    Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "HTNNFU"), _
                    Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "HTNNDU") _
                )
    'HTN FV Initial - # of New Patients Included in Column A
    HTNFV_NEW = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "HTNNIN")
    'HTN FV Initial - # of Home Visits / DUHC Included in Column A
    HTNNDU = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "HTNNDU")
    'HTN Newly Diagnosed - # of Newly Diagnosed Patients
    HTNNDS = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "HTNNDS")

    'CHOL General - # of Patient Visits
    CHOLG_TOTAL = Application.WorksheetFunction.Sum( _
                    Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "CHOLG1"), _
                    Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "CHOLG2"), _
                    Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "CHOLDU") _
                )
    'CHOL General - # of New Patients Included in Column A
    CHOLG_NEW = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "CHOLG1")
    'CHOL General - # of Home Visits / DUHC Included in Column A
    CHOLDU = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "CHOLDU")

    'CHOL FV - # of Patient Visits
    CHOLFV_TOTAL = Application.WorksheetFunction.Sum( _
                        Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "CHOLF1"), _
                        Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "CHOLF2"), _
                        Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "CHOLFD") _
                    )
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
    CHF_TOTAL = Application.WorksheetFunction.Sum( _
                    Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "CHFINI"), _
                    Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "CHFFU"), _
                    Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "CHFDU") _
                )
    'CHF - # of New Patients Included in Column A
    CHF_NEW = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "CHFINI")
    'CHF - # of Home Visits / DUHC Included in Column A
    CHFDU = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "CHFDU")

    'CHG - CHF Group
    '# Patient Visits
    CHG_TOTAL = Application.WorksheetFunction.Sum( _
                    Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "CHGINI"), _
                    Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "CHGFU"), _
                    Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "CHGDU") _
                )
                
    CHG_NEW = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "CHGINI")

    CHGDU = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "CHGDU")

    'MH indiv - # of Patient Visits
    MH_TOTAL = Application.WorksheetFunction.Sum( _
                    Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "MHINI"), _
                    Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "MHFU"), _
                    Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "MHDUHC") _
                )
    'MH indiv - # of New Patients Included in Column A
    MH_NEW = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "MHINI")
    'MH indiv - # of Home Visits / DUHC Included in Column A
    MHDUHC = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "MHDUHC")
    'MH indiv - # of Discharge visits. May be billed separately or at same time as FU so counted separately
    MHDISC = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "MHDISC")

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
    SEN_TOTAL = Application.WorksheetFunction.Sum( _
                    Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "SENINI"), _
                    Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "SENFU"), _
                    Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "SENDU"), _
                    Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "SENDU2") _
                )
    'SENIORS indiv - # of New Patients Included in Column A
    SEN_NEW = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "SENINI")
    'SENIORS indiv - # of Home Visits / DUHC Included in Column A
    SENDU = Application.WorksheetFunction.Sum( _
                Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "SENDU"), _
                Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "SENDU2") _
            )

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
    SENSOC_TOTAL = Application.WorksheetFunction.Sum( _
                    Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "SENSOC"), _
                    Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "SENSO2"), _
                    Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "SENSOD") _
                    )
    'SENIORS Social Club - # of New Patients Included in Column A
    SENSOC_NEW = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "SENSOC")
    'SENIORS Social Club - # of Home Visits / DUHC Included in Column A
    SENSOD = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "SENSOD")

    'HEALTHY LIVING indiv - # of Patient Visits
    HLI_TOTAL = Application.WorksheetFunction.Sum( _
                    Application.WorksheetFunction.SumIf(Sheets("billings").Range("J:J"), "HLINI", Sheets("billings").Range("L:L")), _
                    Application.WorksheetFunction.SumIf(Sheets("billings").Range("J:J"), "HLFU", Sheets("billings").Range("L:L")), _
                    Application.WorksheetFunction.SumIf(Sheets("billings").Range("J:J"), "HLDUHC", Sheets("billings").Range("L:L")) _
                )
                
    'HEALTHY LIVING - Community Outreach
    HLOUTR = Application.WorksheetFunction.SumIf(Sheets("billings").Range("J:J"), "HLOUTR", Sheets("billings").Range("L:L"))



    'HEALTHY LIVING indiv - # of New Patients Included in Column A
    HLI_NEW = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "HLINI")
    'HEALTHY LIVING indiv - # of Home Visits / DUHC Included in Column A
    HLDUHC = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "HLDUHC")

    'SMOKING CESSATION - # of Patient Visits
    SC_TOTAL = Application.WorksheetFunction.Sum( _
                    Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "SCINI"), _
                    Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "SCFU"), _
                    Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "SCDUHC") _
                )
    'SMOKING CESSATION - # of New Patients Included in Column A
    SC_NEW = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "SCINI")
    'SMOKING CESSATION - # of Home Visits / DUHC Included in Column A
    SCDUHC = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "SCDUHC")

    'WEIGHT MANAGEMENT - # of Patient Visits
    WM_TOTAL = Application.WorksheetFunction.Sum( _
                    Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "WINI"), _
                    Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "WFU"), _
                    Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "WDUHC") _
                )
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
    VAC_TOTAL = Application.WorksheetFunction.Sum( _
                    Application.WorksheetFunction.SumIf(Sheets("billings").Range("J:J"), "RNVAC", Sheets("billings").Range("L:L")), _
                    Application.WorksheetFunction.SumIf(Sheets("billings").Range("J:J"), "IMMUNR", Sheets("billings").Range("L:L")) _
                )
    'VACCINE - # of New Patients Included in Column A
    Call TrueStatusCANCER("N", "RNVAC", "IMMUNR", "RNVAC")
    VAC_NEW = Application.WorksheetFunction.CountIf(Sheets("temp").Range("N:N"), ">0")

    'Prediabetes - # of Patient Visits
    PD_TOTAL = Application.WorksheetFunction.Sum( _
                    Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "PDINI"), _
                    Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "PDFU"), _
                    Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "PDDUHC") _
                )
    'Prediabetes - # of New Patients Included in Column A
    PD_NEW = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "PDINI")
    'Prediabetes - # of Home Visits / DUHC Included in Column A
    PDDUHC = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "PDDUHC")

    'Insomnia indiv - # of Patient Visits
    INSOMI_TOTAL = Application.WorksheetFunction.Sum( _
                        Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "INSOIN"), _
                        Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "INSOFU"), _
                        Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "INSODU") _
                    )
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
    MAT_TOTAL = Application.WorksheetFunction.Sum( _
                    Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "MCP*"), _
                    Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "MCBF"), _
                    Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "MCWBC"), _
                    Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "MCGEN"), _
                    Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "MCDUHC") _
                )
    'MATERNAL HEALTH - # of New Patients Included in Column A
    Call TrueStatusMAT("Z", "MCBF", "MCP*", "MCWBC", "MCGEN")
    MAT_NEW = Application.WorksheetFunction.CountIf(Sheets("temp").Range("Z:Z"), ">0")
    'MATERNAL HEALTH - # of Home Visits / DUHC Included in Column A
    MCDUHC = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "MCDUHC")

    'CANCER VISITS - # of Patient Visits
    CANCERV_TOTAL = Application.WorksheetFunction.Sum( _
                        Application.WorksheetFunction.SumIf(Sheets("billings").Range("J:J"), "CERVV", Sheets("billings").Range("L:L")), _
                        Application.WorksheetFunction.SumIf(Sheets("billings").Range("J:J"), "COLV", Sheets("billings").Range("L:L")), _
                        Application.WorksheetFunction.SumIf(Sheets("billings").Range("J:J"), "BREASV", Sheets("billings").Range("L:L")), _
                        Application.WorksheetFunction.SumIf(Sheets("billings").Range("J:J"), "CANDU", Sheets("billings").Range("L:L")) _
                    )
    'CANCER VISITS - # of New Patients Included in Column A
    Call TrueStatusCANCER("P", "COLV", "CERV", "BREASV*")
    CANCERV_NEW = Application.WorksheetFunction.CountIf(Sheets("temp").Range("P:P"), ">0")
    'Cancer visits - # of Home Visits / DUHC Included in Column A
    CANDU = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "CANDU")

    'CANCER REVIEW - # of Patient Visits
    CANCERR_TOTAL = Application.WorksheetFunction.Sum( _
                        Application.WorksheetFunction.SumIf(Sheets("billings").Range("J:J"), "CERVR", Sheets("billings").Range("L:L")), _
                        Application.WorksheetFunction.SumIf(Sheets("billings").Range("J:J"), "COLR", Sheets("billings").Range("L:L")), _
                        Application.WorksheetFunction.SumIf(Sheets("billings").Range("J:J"), "BREASR", Sheets("billings").Range("L:L")) _
                    )

    'PAIN MANAGEMENT - # of Patient Visits
    PAIN_TOTAL = Application.WorksheetFunction.Sum( _
                    Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "PAININ"), _
                    Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "PAINFU"), _
                    Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "PAINDU") _
                )
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
    MEDREC_TOTAL = Application.WorksheetFunction.Sum( _
                        Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "MEDINI"), _
                        Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "MEDFU"), _
                        Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "MEDDHC") _
                    )
    'medrec - # of New Patients Included in Column A
    MEDREC_NEW = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "MEDINI")
    'medrec - # of Home Visits / DUHC Included in Column A
    MEDDHC = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "MEDDHC")

    'PT - # of Patient Visits
    PT_TOTAL = Application.WorksheetFunction.Sum( _
                    Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "PTINI"), _
                    Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "PTFU"), _
                    Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "PTDUHC") _
                )
    'PT - # of New Patients Included in Column A
    PT_NEW = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "PTINI")
    'PT - # of Home Visits / DUHC Included in Column A
    PTDUHC = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "PTDUHC")
    'PT - # of Discharged Patients
    PTDISC = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "PTDISC")

    'Mindful Eating GS - # of Patient Visits
    MINDGS_TOTAL = Application.WorksheetFunction.SumIf(Sheets("billings").Range("J:J"), "MINDGS", Sheets("billings").Range("L:L"))
    'Mindful Eating GS - # of New Patients Included in Column A
    Call TrueStatus("R", "MINDGS")
    MINDGS_NEW = Application.WorksheetFunction.CountIf(Sheets("temp").Range("R:R"), ">0")

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

    'Health Links
    HK_TOTAL = Application.WorksheetFunction.Sum( _
                    Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "HKINI"), _
                    Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "HKFU"), _
                    Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "HKDUHC") _
                )
                
    HK_NEW = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "HKINI")

    HKDUHC = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "HKDUHC")

    ' Healthy Living - Yoga
    HLYOGA_NEW = Application.WorksheetFunction.CountIf(Sheets("temp").Range("T:T"), ">0")
    HLYOGA_TOTAL = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "HLYOGA")


    ' ---------------------------------------------------------------------
    'ENTERING VALUES INTO APPROPRIATE CELLS in "IndividualSummary-Detailed"
    ' ---------------------------------------------------------------------

    Sheets("IndividualSummary-Detailed").Select

    Sheets("IndividualSummary-Detailed").Range("B8").Value = DMGINI_TOTAL
    Sheets("IndividualSummary-Detailed").Range("C8").Value = DMGINI_NEW
    Sheets("IndividualSummary-Detailed").Range("D8").Value = DMGDHC

    Sheets("IndividualSummary-Detailed").Range("B9").Value = DMGROUP_TOTAL
    Sheets("IndividualSummary-Detailed").Range("C9").Value = DMGROUP_NEW

    Sheets("IndividualSummary-Detailed").Range("B10").Value = DMFOCUS_TOTAL
    Sheets("IndividualSummary-Detailed").Range("C10").Value = DMFOCUS_NEW
    Sheets("IndividualSummary-Detailed").Range("D10").Value = DMFDHC

    Sheets("IndividualSummary-Detailed").Range("B11").Value = CKD_TOTAL
    Sheets("IndividualSummary-Detailed").Range("C11").Value = CKD_NEW
    Sheets("IndividualSummary-Detailed").Range("D11").Value = CKDDUH

    Sheets("IndividualSummary-Detailed").Range("B12").Value = CKDGRP_TOTAL
    Sheets("IndividualSummary-Detailed").Range("C12").Value = CKDGRP_NEW

    Sheets("IndividualSummary-Detailed").Range("B13").Value = LH_TOTAL
    Sheets("IndividualSummary-Detailed").Range("C13").Value = LH_NEW
    Sheets("IndividualSummary-Detailed").Range("D13").Value = LHDUHC

    Sheets("IndividualSummary-Detailed").Range("B15").Value = HTNE_TOTAL
    Sheets("IndividualSummary-Detailed").Range("C15").Value = HTNE_NEW
    Sheets("IndividualSummary-Detailed").Range("D15").Value = HTNEDU

    Sheets("IndividualSummary-Detailed").Range("B16").Value = HTNFV_TOTAL
    Sheets("IndividualSummary-Detailed").Range("C16").Value = HTNFV_NEW
    Sheets("IndividualSummary-Detailed").Range("D16").Value = HTNNDU

    Sheets("IndividualSummary-Detailed").Range("J10").Value = HTNNDS

    Sheets("IndividualSummary-Detailed").Range("B17").Value = HTNGS_TOTAL
    Sheets("IndividualSummary-Detailed").Range("C17").Value = HTNGS_NEW

    Sheets("IndividualSummary-Detailed").Range("B18").Value = CHOLG_TOTAL
    Sheets("IndividualSummary-Detailed").Range("C18").Value = CHOLG_NEW
    Sheets("IndividualSummary-Detailed").Range("D18").Value = CHOLDU

    Sheets("IndividualSummary-Detailed").Range("B19").Value = CHOLFV_TOTAL
    Sheets("IndividualSummary-Detailed").Range("C19").Value = CHOLFV_NEW
    Sheets("IndividualSummary-Detailed").Range("D19").Value = CHOLFD

    Sheets("IndividualSummary-Detailed").Range("B20").Value = CHOLGS_TOTAL
    Sheets("IndividualSummary-Detailed").Range("C20").Value = CHOLGS_NEW

    Sheets("IndividualSummary-Detailed").Range("B21").Value = CHG_TOTAL
    Sheets("IndividualSummary-Detailed").Range("C21").Value = CHG_NEW
    Sheets("IndividualSummary-Detailed").Range("D21").Value = CHGDU

    Sheets("IndividualSummary-Detailed").Range("B22").Value = CHF_TOTAL
    Sheets("IndividualSummary-Detailed").Range("C22").Value = CHF_NEW
    Sheets("IndividualSummary-Detailed").Range("D22").Value = CHFDU

    Sheets("IndividualSummary-Detailed").Range("B24").Value = MH_TOTAL
    Sheets("IndividualSummary-Detailed").Range("C24").Value = MH_NEW
    Sheets("IndividualSummary-Detailed").Range("D24").Value = MHDUHC

    Sheets("IndividualSummary-Detailed").Range("B25").Value = MHMBSG_TOTAL
    Sheets("IndividualSummary-Detailed").Range("C25").Value = MHMBSG_NEW
    Sheets("IndividualSummary-Detailed").Range("B26").Value = MHMBCG_TOTAL
    Sheets("IndividualSummary-Detailed").Range("C26").Value = MHMBCG_NEW
    Sheets("IndividualSummary-Detailed").Range("B27").Value = MHANXG_TOTAL
    Sheets("IndividualSummary-Detailed").Range("C27").Value = MHANXG_NEW
    Sheets("IndividualSummary-Detailed").Range("B28").Value = MHDEP_TOTAL
    Sheets("IndividualSummary-Detailed").Range("C28").Value = MHDEP_NEW

    Sheets("IndividualSummary-Detailed").Range("B29").Value = MHDISC

    Sheets("IndividualSummary-Detailed").Range("B30").Value = HLYOGA_TOTAL
    Sheets("IndividualSummary-Detailed").Range("C30").Value = HLYOGA_NEW

    Sheets("IndividualSummary-Detailed").Range("B33").Value = SEN_TOTAL
    Sheets("IndividualSummary-Detailed").Range("C33").Value = SEN_NEW
    Sheets("IndividualSummary-Detailed").Range("D33").Value = SENDU

    Sheets("IndividualSummary-Detailed").Range("B34").Value = SENMGS_TOTAL
    Sheets("IndividualSummary-Detailed").Range("C34").Value = SENMGS_NEW
    Sheets("IndividualSummary-Detailed").Range("B35").Value = SENEGS_TOTAL
    Sheets("IndividualSummary-Detailed").Range("C35").Value = SENEGS_NEW
    Sheets("IndividualSummary-Detailed").Range("B36").Value = SENFGS_TOTAL
    Sheets("IndividualSummary-Detailed").Range("C36").Value = SENFGS_NEW
    Sheets("IndividualSummary-Detailed").Range("B37").Value = SENNGS_TOTAL
    Sheets("IndividualSummary-Detailed").Range("C37").Value = SENNGS_NEW
    Sheets("IndividualSummary-Detailed").Range("B38").Value = SENCGS_TOTAL
    Sheets("IndividualSummary-Detailed").Range("C38").Value = SENCGS_NEW

    Sheets("IndividualSummary-Detailed").Range("B39").Value = SENSOC_TOTAL
    Sheets("IndividualSummary-Detailed").Range("C39").Value = SENSOC_NEW
    Sheets("IndividualSummary-Detailed").Range("D39").Value = SENSOD

    Sheets("IndividualSummary-Detailed").Range("B40").Value = MEDREC_TOTAL
    Sheets("IndividualSummary-Detailed").Range("C40").Value = MEDREC_NEW
    Sheets("IndividualSummary-Detailed").Range("D40").Value = MEDDHC

    Sheets("IndividualSummary-Detailed").Range("B42").Value = HLI_TOTAL
    Sheets("IndividualSummary-Detailed").Range("C42").Value = HLI_NEW
    Sheets("IndividualSummary-Detailed").Range("D42").Value = HLDUHC

    Sheets("IndividualSummary-Detailed").Range("B43").Value = SC_TOTAL
    Sheets("IndividualSummary-Detailed").Range("C43").Value = SC_NEW
    Sheets("IndividualSummary-Detailed").Range("D43").Value = SCDUHC

    Sheets("IndividualSummary-Detailed").Range("B44").Value = WM_TOTAL
    Sheets("IndividualSummary-Detailed").Range("C44").Value = WM_NEW
    Sheets("IndividualSummary-Detailed").Range("D44").Value = WDUHC

    Sheets("IndividualSummary-Detailed").Range("B45").Value = WMGS_TOTAL
    Sheets("IndividualSummary-Detailed").Range("C45").Value = WMGS_NEW

    Sheets("IndividualSummary-Detailed").Range("B46").Value = VAC_TOTAL
    Sheets("IndividualSummary-Detailed").Range("C46").Value = VAC_NEW

    Sheets("IndividualSummary-Detailed").Range("B47").Value = PD_TOTAL
    Sheets("IndividualSummary-Detailed").Range("C47").Value = PD_NEW
    Sheets("IndividualSummary-Detailed").Range("D47").Value = PDDUHC

    Sheets("IndividualSummary-Detailed").Range("B48").Value = INSOMI_TOTAL
    Sheets("IndividualSummary-Detailed").Range("C48").Value = INSOMI_NEW
    Sheets("IndividualSummary-Detailed").Range("D48").Value = INSODU

    Sheets("IndividualSummary-Detailed").Range("B49").Value = INSOMGS_TOTAL
    Sheets("IndividualSummary-Detailed").Range("C49").Value = INSOMGS_NEW

    Sheets("IndividualSummary-Detailed").Range("B50").Value = MAT_TOTAL
    Sheets("IndividualSummary-Detailed").Range("C50").Value = MAT_NEW
    Sheets("IndividualSummary-Detailed").Range("D50").Value = MCDUHC

    Sheets("IndividualSummary-Detailed").Range("B51").Value = CANCERV_TOTAL

    Sheets("IndividualSummary-Detailed").Range("B52").Value = CANCERR_TOTAL
    Sheets("IndividualSummary-Detailed").Range("C52").Value = CANCERV_NEW
    Sheets("IndividualSummary-Detailed").Range("D52").Value = CANDU

    Sheets("IndividualSummary-Detailed").Range("B53").Value = PAIN_TOTAL
    Sheets("IndividualSummary-Detailed").Range("C53").Value = PAIN_NEW
    Sheets("IndividualSummary-Detailed").Range("D53").Value = PAINDU

    Sheets("IndividualSummary-Detailed").Range("B54").Value = PAINGS_TOTAL
    Sheets("IndividualSummary-Detailed").Range("C54").Value = PAINGS_NEW

    Sheets("IndividualSummary-Detailed").Range("B55").Value = PT_TOTAL
    Sheets("IndividualSummary-Detailed").Range("C55").Value = PT_NEW
    Sheets("IndividualSummary-Detailed").Range("D55").Value = PTDUHC

    Sheets("IndividualSummary-Detailed").Range("J13").Value = PTDISC

    Sheets("IndividualSummary-Detailed").Range("B56").Value = MINDGS_TOTAL
    Sheets("IndividualSummary-Detailed").Range("C56").Value = MINDGS_NEW

    Sheets("IndividualSummary-Detailed").Range("B58").Value = MISCEL_TOTAL
    Sheets("IndividualSummary-Detailed").Range("C58").Value = MISCEL_NEW
    Sheets("IndividualSummary-Detailed").Range("B59").Value = INJECT_TOTAL
    Sheets("IndividualSummary-Detailed").Range("C59").Value = INJECT_NEW
    Sheets("IndividualSummary-Detailed").Range("B60").Value = BPCHECK_TOTAL
    Sheets("IndividualSummary-Detailed").Range("C60").Value = BPCHECK_NEW
    Sheets("IndividualSummary-Detailed").Range("B61").Value = SWAB_TOTAL
    Sheets("IndividualSummary-Detailed").Range("C61").Value = SWAB_NEW
    Sheets("IndividualSummary-Detailed").Range("B62").Value = LAB_TOTAL
    Sheets("IndividualSummary-Detailed").Range("C62").Value = LAB_NEW
    Sheets("IndividualSummary-Detailed").Range("B63").Value = FORM_TOTAL
    Sheets("IndividualSummary-Detailed").Range("C63").Value = FORM_NEW
    Sheets("IndividualSummary-Detailed").Range("B64").Value = RNHW_TOTAL
    Sheets("IndividualSummary-Detailed").Range("C64").Value = RNHW_NEW

    Sheets("IndividualSummary-Detailed").Range("B66").Value = HLI_TOTAL + FORM_TOTAL + MISCEL_TOTAL + INJECT_TOTAL + SWAB_TOTAL + LAB_TOTAL
    Sheets("IndividualSummary-Detailed").Range("C66").Value = HLI_NEW + FORM_NEW + MISCEL_NEW + INJECT_NEW + SWAB_NEW + LAB_NEW
    Sheets("IndividualSummary-Detailed").Range("D66").Value = HLDUHC


    ' ---------------------------------------------------------------------
    'ENTERING VALUES INTO APPROPRIATE CELLS in "IndividualSummary-Basic"
    ' ---------------------------------------------------------------------


    Sheets("IndividualSummary-Basic").Select

    Sheets("IndividualSummary-Basic").Range("B9").Value = DMGINI_TOTAL + DMGROUP_TOTAL + DMFOCUS_TOTAL
    Sheets("IndividualSummary-Basic").Range("C9").Value = DMGINI_NEW + DMGROUP_NEW + DMFOCUS_NEW
    Sheets("IndividualSummary-Basic").Range("D9").Value = DMGDHC + DMFDHC

    Sheets("IndividualSummary-Basic").Range("B10").Value = LH_TOTAL
    Sheets("IndividualSummary-Basic").Range("C10").Value = LH_NEW
    Sheets("IndividualSummary-Basic").Range("D10").Value = LHDUHC

    Sheets("IndividualSummary-Basic").Range("B12").Value = HTNE_TOTAL + HTNFV_TOTAL + HTNGS_TOTAL
    Sheets("IndividualSummary-Basic").Range("C12").Value = HTNE_NEW + HTNFV_NEW + HTNGS_NEW
    Sheets("IndividualSummary-Basic").Range("D12").Value = HTNEDU + HTNNDU

    Sheets("IndividualSummary-Basic").Range("B13").Value = CHOLG_TOTAL + CHOLFV_TOTAL + CHOLGS_TOTAL
    Sheets("IndividualSummary-Basic").Range("C13").Value = CHOLG_NEW + CHOLFV_NEW + CHOLGS_NEW
    Sheets("IndividualSummary-Basic").Range("D13").Value = CHOLDU + CHOLFD

    Sheets("IndividualSummary-Basic").Range("B14").Value = CHF_TOTAL + CHG_TOTAL
    Sheets("IndividualSummary-Basic").Range("C14").Value = CHF_NEW + CHG_NEW
    Sheets("IndividualSummary-Basic").Range("D14").Value = CHFDU + CHGDU

    Sheets("IndividualSummary-Basic").Range("B15").Value = MH_TOTAL + MHMBSG_TOTAL + MHMBCG_TOTAL + MHANXG_TOTAL + MHDEP_TOTAL + HLYOGA_TOTAL
    Sheets("IndividualSummary-Basic").Range("C15").Value = MH_NEW + MHMBSG_NEW + MHMBCG_NEW + MHANXG_NEW + MHDEP_NEW + HLYOGA_NEW
    Sheets("IndividualSummary-Basic").Range("D15").Value = MHDUHC

    Sheets("IndividualSummary-Basic").Range("B18").Value = SEN_TOTAL + SENMGS_TOTAL + SENEGS_TOTAL + SENFGS_TOTAL + SENNGS_TOTAL + SENCGS_TOTAL + SENSOC_TOTAL
    Sheets("IndividualSummary-Basic").Range("C18").Value = SEN_NEW + SENMGS_NEW + SENEGS_NEW + SENFGS_NEW + SENNGS_NEW + SENCGS_NEW + SENSOC_NEW
    Sheets("IndividualSummary-Basic").Range("D18").Value = SENDU + SENSOD

    Sheets("IndividualSummary-Basic").Range("B20").Value = SC_TOTAL
    Sheets("IndividualSummary-Basic").Range("C20").Value = SC_NEW
    Sheets("IndividualSummary-Basic").Range("D20").Value = SCDUHC

    Sheets("IndividualSummary-Basic").Range("B21").Value = WM_TOTAL + WMGS_TOTAL + MINDGS_TOTAL
    Sheets("IndividualSummary-Basic").Range("C21").Value = WM_NEW + WMGS_NEW + MINDGS_NEW
    Sheets("IndividualSummary-Basic").Range("D21").Value = WDUHC

    Sheets("IndividualSummary-Basic").Range("B22").Value = VAC_TOTAL
    Sheets("IndividualSummary-Basic").Range("C22").Value = VAC_NEW

    Sheets("IndividualSummary-Basic").Range("B23").Value = PD_TOTAL
    Sheets("IndividualSummary-Basic").Range("C23").Value = PD_NEW
    Sheets("IndividualSummary-Basic").Range("D23").Value = PDDUHC

    Sheets("IndividualSummary-Basic").Range("B24").Value = INSOMI_TOTAL + INSOMGS_TOTAL
    Sheets("IndividualSummary-Basic").Range("C24").Value = INSOMI_NEW + INSOMGS_NEW
    Sheets("IndividualSummary-Basic").Range("D24").Value = INSODU

    Sheets("IndividualSummary-Basic").Range("B25").Value = MAT_TOTAL
    Sheets("IndividualSummary-Basic").Range("C25").Value = MAT_NEW
    Sheets("IndividualSummary-Basic").Range("D25").Value = MCDUHC

    Sheets("IndividualSummary-Basic").Range("B26").Value = CANCERV_TOTAL + CANCERR_TOTAL
    Sheets("IndividualSummary-Basic").Range("C26").Value = CANCERV_NEW
    Sheets("IndividualSummary-Basic").Range("D26").Value = CANDU

    Sheets("IndividualSummary-Basic").Range("B27").Value = PAIN_TOTAL + PAINGS_TOTAL
    Sheets("IndividualSummary-Basic").Range("C27").Value = PAIN_NEW + PAINGS_NEW
    Sheets("IndividualSummary-Basic").Range("D27").Value = PAINDU

    Sheets("IndividualSummary-Basic").Range("B29").Value = HLI_TOTAL + FORM_TOTAL + MISCEL_TOTAL + INJECT_TOTAL + SWAB_TOTAL + LAB_TOTAL
    Sheets("IndividualSummary-Basic").Range("C29").Value = HLI_NEW + FORM_NEW + MISCEL_NEW + INJECT_NEW + SWAB_NEW + LAB_NEW
    Sheets("IndividualSummary-Basic").Range("D29").Value = HLDUHC

    Sheets("IndividualSummary-Basic").Range("B32").Value = HK_TOTAL
    Sheets("IndividualSummary-Basic").Range("C32").Value = HK_NEW
    Sheets("IndividualSummary-Basic").Range("D32").Value = HKDUHC

    Sheets("IndividualSummary-Basic").Range("B33").Value = CKD_TOTAL
    Sheets("IndividualSummary-Basic").Range("C33").Value = CKD_NEW
    Sheets("IndividualSummary-Basic").Range("D33").Value = CKDDUH

    'provider name on report
    Sheets("IndividualSummary-Basic").Range("B2").Value = Sheets("billings").Range("O2").Value & ", " & Sheets("billings").Range("P2").Value
    Sheets("IndividualSummary-Detailed").Range("B2").Value = Sheets("billings").Range("O2").Value & ", " & Sheets("billings").Range("P2").Value
    

    ' --------------
    ' Home To Stay
    ' --------------

    'Home To Stay - # of Patient Visits by NP/RN/RPN
    HTS_TOTAL = Application.WorksheetFunction.Sum( _
                    Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "HTSINI"), _
                    Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "HTSFU"), _
                    Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "HTSDUH") _
                )
    'Home To Stay - # of New Patients Included in Column A
    HTS_NEW = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "HTSINI")
    'Home To Stay - # of Home Visits / DUHC Included in Column A
    HTSDUH = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "HTSDUH")
    
    ' Post Hospital Discharge
    HTSPHD = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "HTSPHD")
    ' External/Internal Referral
    HTSREF = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "HTSREF")
    ' Med Review
    HTSMR_TOTAL = Application.WorksheetFunction.Sum( _
                    Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "HTSMRI"), _
                    Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "HTSMRF") _
                )
    HTSMR_NEW = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "HTSMRI")
    ' Dietitian
    HTSRD_TOTAL= Application.WorksheetFunction.Sum( _
                    Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "HTSRDI"), _
                    Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "HTSRDF") _
                )
    HTSRD_NEW = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "HTSRDI")
    ' Occupational Therapist
    HTSOT_TOTAL= Application.WorksheetFunction.Sum( _
                    Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "HTSOTI"), _
                    Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "HTSOTF") _
                )
    HTSOT_NEW = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "HTSOTI")
    ' Social Worker
    HTSSW_TOTAL= Application.WorksheetFunction.Sum( _
                    Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "HTSSWI"), _
                    Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "HTSSWF") _
                )
    HTSSW_NEW = Application.WorksheetFunction.CountIf(Sheets("billings").Range("J:J"), "HTSSWI")

    ' Provider Name
    Sheets("Home To Stay").Range("B2").Value = Sheets("billings").Range("O2").Value & ", " & Sheets("billings").Range("P2").Value
    
    ' Nurse Visits
    Sheets("Home To Stay").Range("B9").Value = HTS_TOTAL
    Sheets("Home To Stay").Range("C9").Value = HTS_NEW
    Sheets("Home To Stay").Range("B10").Value = HTSMR_TOTAL
    Sheets("Home To Stay").Range("C10").Value = HTSMR_NEW
    Sheets("Home To Stay").Range("B11").Value = HTSRD_TOTAL
    Sheets("Home To Stay").Range("C11").Value = HTSRD_NEW
    Sheets("Home To Stay").Range("B12").Value = HTSOT_TOTAL
    Sheets("Home To Stay").Range("C12").Value = HTSOT_NEW
    Sheets("Home To Stay").Range("B13").Value = HTSSW_TOTAL
    Sheets("Home To Stay").Range("C13").Value = HTSSW_NEW
    ' 
    Sheets("Home To Stay").Range("B15").Value = HTSPHD
    Sheets("Home To Stay").Range("B16").Value = HTSDUH
    Sheets("Home To Stay").Range("B17").Value = HTSREF


End Sub


' -----------------
' Helper Functions
' -----------------


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
            x = Application.WorksheetFunction.Sum( _
                    Application.WorksheetFunction.CountIfs(Sheets("billings").Range("R1:R5500"), identifier, Sheets("billings").Range("J1:J5500"), code1), _
                    Application.WorksheetFunction.CountIfs(Sheets("billings").Range("R1:R5500"), identifier, Sheets("billings").Range("J1:J5500"), code2), _
                    Application.WorksheetFunction.CountIfs(Sheets("billings").Range("R1:R5500"), identifier, Sheets("billings").Range("J1:J5500"), code3), _
                    Application.WorksheetFunction.CountIfs(Sheets("billings").Range("R1:R5500"), identifier, Sheets("billings").Range("J1:J5500"), code4) _
                )
            Sheets("temp").Range(column & i).Value = x
        Next i
    End With

End Sub


Public Sub TrueStatusCANCER(column, code1, code2, code3)

    uniquepat = Sheets("temp").Cells(Rows.Count, 1).End(xlUp).Row

    With Sheets("temp")
        For i = 2 To uniquepat + 1
            identifier = Sheets("temp").Range("A" & i).Value
            x = Application.WorksheetFunction.Sum( _
                    Application.WorksheetFunction.CountIfs(Sheets("billings").Range("R1:R5500"), identifier, Sheets("billings").Range("J1:J5500"), code1), _
                    Application.WorksheetFunction.CountIfs(Sheets("billings").Range("R1:R5500"), identifier, Sheets("billings").Range("J1:J5500"), code2), _
                    Application.WorksheetFunction.CountIfs(Sheets("billings").Range("R1:R5500"), identifier, Sheets("billings").Range("J1:J5500"), code3) _
                )
            Sheets("temp").Range(column & i).Value = x
        Next i
    End With

End Sub



