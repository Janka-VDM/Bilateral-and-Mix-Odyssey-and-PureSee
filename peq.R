####LIBRARIES####

library("stringr")
library("DBI")
library("odbc")
library("excel.link")
library("dplyr")
library("wktmo")
library("excel.link")
library("lubridate")
library("ggplot2")
library("tidyverse")
library("reshape2")
library("patchwork")
library('readxl')
library('janitor')
library('openxlsx')

## Need to comment/uncomment which PEQ you want:
# 
# 1 month post op is 20-40
postopdatemin<-20
postopdatemax<-40
#
# # 3 month post op is 85-120
# postopdatemin<-85
# postopdatemax<-120

####FILE PATH####

path <- "S:/Clinical Services/Research Assistants/Bilateral-and-Mix-Odyssey-and-PureSee/"

con <- DBI::dbConnect(odbc::odbc(), 
                      Driver = "SQL Server", 
                      Server = "GB-SQL-N2\\I02", 
                      Trusted_Connection = "True")

mindate <- '2024-01-01'
maxdate <- '2026-02-28'

data1 <- read.xlsx(paste(path,"tables/pxs.xlsx",sep=""))
data <- data1 #%>% select(PatientID, SurgeryDateTime, SurgeryLens)
data$SurgeryDateTime <- as.Date(data$SurgeryDateTime, origin='1899-12-30')
sx <- toString(unique(data$SurgeryID))

#### SQL Query ####
query <- paste("SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED
Select Distinct
               /*Con*/
               
			   S.PatientID,
			   floor((datediff(day,0,s.SurgeryDateTime) - datediff(day,0,P.DOB)) / 365.2425) as Age,
			   ese.EyeID,
               iol_press.DominantEyeID,
               lil.DescriptionEnGB As Model,
               s.SurgeryDateTime,
			   lstg.DescriptionEnGB AS Surgery_Group,
               SPOM1.Sph_M1,
               SPOM1.Cyl_M1,
               lc.DescriptionEnGB as Complication,

			   /*PRO*/
        Q1.SurveyDate,
			  Q1.OptionText AS M1_Satisfied,
			  Q2.OptionText AS M1_Recommend,
			  Q3.OptionText AS M1_DoOver,
			  Q10.OptionText AS M1_Consent,
			  Q4.OptionText AS M1_Glare,
			  Q5.OptionText AS M1_Halos,
			  Q6.OptionText AS M1_Ghost,
			  Q7.OptionText AS M1_Starbursts,
			  Q8.OptionText AS M1_DryEyes,
			  Q9.OptionText AS M1_NightDrive,
			  Q12.OptionText AS M1_Hobbies,
			  q13.OptionText AS M1_Sports

               From EMR.dbo.tSurgery S
               
               Left Join Patient.dbo.tPatient P
               On P.PatientID = S.PatientID
               
               Left Join EMR.dbo.vEmployeeView Eply
               On Eply.EmployeeID = S.SurgeonID
               
               LEFT JOIN EMR.dbo.tExam AS e
               ON e.PatientID = s.PatientID
               AND e.ExamID = 
                 (
                   SELECT TOP (1)
                   e2.ExamID
                   FROM EMR.dbo.tExam AS e2
                   LEFT JOIN EMR.dbo.tExam_ExamType AS eet
                   ON eet.ExamID = e2.ExamID
                   WHERE e2.PatientID = e.PatientID
                   AND eet.ExamTypeID in (5,25,40)
                   AND e2.ExamDateTime < s.SurgeryDateTime
                   ORDER BY e2.ExamID DESC
                 )
               
               Left Join Estate.dbo.tLocation Surgery_Lv
               On Surgery_Lv.LocationID = S.SurgeryLocationID
               
               LEFT JOIN Estate.dbo.tLocation_Region as lr
               on lr.LocationID = s.SurgeryLocationID
               
               Left Join Estate.dbo.LocationView Exam_Lv
               On Exam_Lv.LocationID = S.SurgeryLocationID
               
               Left Join EMR.dbo.tExam_ExamType EET
               On EET.ExamID = E.ExamID
               
               Left Join EMR.dbo.tEyeSurgeryEye ESE
               On ESE.SurgeryID = S.SurgeryID
                              
               left join EMR.dbo.tEyeSurgeryEyeComplication as esec
               on esec.SurgeryID = s.SurgeryID and esec.EyeID = ese.EyeID
               
               left join EMR.dbo.tLookupComplication as lc
               on lc.ComplicationID = esec.ComplicationID
               
               Left Join EMR.dbo.tEyeExamSurgery EES
               ON EES.ExamID = e.ExamID
               
               Left Join EMR.dbo.tEyeExam EE
               On EE.ExamID = e.ExamID
               
               Left Join EMR.dbo.tEyeMeasurement EM
               On EM.EyeMeasurementID = EE.EyeMeasurementID
               
               Left Join emr.dbo.tEyeMeasurementEye EME
               ON EME.EyeMeasurementID = EM.EyeMeasurementID and EME.EyeID = ESE.EyeID
               
               LEFT JOIN EMR.dbo.tEyePrescriptionEye AS epeC 
               ON epeC.EyePrescriptionID = ees.ManifestPrescriptionID
               AND epeC.EyeID = ese.EyeID
               AND epeC.EyeID = eme.EyeID

               LEFT JOIN EMR.dbo.tVisualAcuity VA
               ON VA.VisualAcuityID = epeC.VisualAcuityID
               
               LEFT JOIN EMR.dbo.tLookupVA lvaCBCDVA
               ON lvaCBCDVA.VAID = VA.BCDVAID

			   LEFT JOIN EMR.dbo.tLookupVA lvaCDCNVA
               ON lvaCDCNVA.VAID = VA.DCNVAID

			   LEFT JOIN EMR.dbo.tLookupVA lvaCDCIVA
               ON lvaCDCIVA.VAID = VA.DCIVAID

			   LEFT JOIN EMR.dbo.tLookupVA lvaCBCIVA
               ON lvaCBCIVA.VAID = VA.BCIVAID
               
         LEFT JOIN EMR.dbo.tLookupVA lvaCBCNVA
               ON lvaCBCNVA.VAID = VA.BCNVAID

			   LEFT JOIN emr.dbo.tVisualAcuity vaBinocDCNVA 
			   ON vaBinocDCNVA.VisualAcuityID = ees.RefractionVisualAcuityID

			   LEFT JOIN EMR.dbo.tLookupVA lvaBinocCDCNVA
               ON lvaBinocCDCNVA.VAID = vaBinocDCNVA.DCNVAID

			   LEFT JOIN emr.dbo.tVisualAcuity vaBinocDCIVA 
			   ON vaBinocDCIVA.VisualAcuityID = ees.RefractionVisualAcuityID

			   LEFT JOIN EMR.dbo.tLookupVA lvaBinocCDCIVA
               ON lvaBinocCDCIVA.VAID = vaBinocDCIVA.DCIVAID

			   LEFT JOIN emr.dbo.tVisualAcuity vaBinocCBCIVA 
			   ON vaBinocCBCIVA.VisualAcuityID = ees.RefractionVisualAcuityID

			   LEFT JOIN EMR.dbo.tLookupVA lvaBinocCBCIVA
               ON lvaBinocCBCIVA.VAID = vaBinocCBCIVA.BCIVAID

               LEFT Join EMR.dbo.tEyeSurgeryEyeLVC ESL
               ON ESL.SurgeryID = S.SurgeryID and ESL.EyeID = ESE.EyeID
               
               LEFT join EMR.dbo.tLookupAblationType LUAT
               On LUAT.AblationTypeID = ESL.AblationTypeID
               
               Left Join EMR.dbo.tLookupSurgeryType AS lst
               On ese.SurgeryTypeID = lst.SurgeryTypeID
               
               Left Join EMR.dbo.tLookupSurgeryTypeGroup AS lstg
               On lstg.SurgeryTypeGroupID = lst.SurgeryTypeGroupID
               
               Left Join EMR.dbo.tLookupSurgeryTreatmentType AS lstt 
               ON lstt.SurgeryTreatmentTypeID = ESE.SurgeryTreatmentTypeID
               
               Left Join EMR.dbo.tEyeSurgeryEyeIOL ESEI
               On ESEI.SurgeryID = S.SurgeryID and ESEI.EyeID = ESE.EyeID
               
               LEFT JOIN EMR.dbo.tIOLLensCalculation AS ilc
               ON ilc.IOLLensCalculationID = esei.Type1IOLLensCalculationID
               
               LEFT JOIN EMR.dbo.tLookupIOLLens AS lil
               ON lil.IOLLensID = ilc.IOLLensID
               
               Left join emr.dbo.tEyeExamCommonEye eece
               on eece.ExamID = e.ExamID and eece.EyeID = ESE.EyeID
               
               Left Join emr.dbo.tLookupLens ll
               on ll.LensID = eece.LensID
               
               LEFT JOIN emr.dbo.tLookupRecommendedSurgeryType rst
               ON rst.RecommendedSurgeryTypeID = eece.RecommendedProcedureID
               
               OUTER APPLY (SELECT TOP (1)
              e2.PatientID,
              e2.ExamDateTime,
              EES2.DominantEyeID
              
              FROM EMR.dbo.tExam as e2
              
              LEFT JOIN EMR.dbo.tExam_ExamType AS eet2
              ON eet2.ExamID = e2.ExamID
              
              Left Join EMR.dbo.tEyeExamSurgery EES2
              ON EES2.ExamID = e2.ExamID
              
              WHERE eet2.ExamTypeID in (18)
              and e.PatientID = e2.PatientID
              AND e.ExamDateTime < e2.ExamDateTime) AS iol_press
              
               LEFT JOIN 
               (
                 SELECT DISTINCT
                 eeR.ExamID,
                 eeR.RightEyeVisualAcuityID AS VisualAcuityID,
                 eeR.BinocularVisualAcuityID,
                 eeR.EyeMeasurementID,
                 1 AS EyeID
                 FROM EMR.dbo.tEyeExam AS eeR 
                 
                 UNION
                 
                 SELECT DISTINCT
                 eeL.ExamID,
                 eeL.LeftEyeVisualAcuityID AS VisualAcuityID,
                 eeL.BinocularVisualAcuityID,
                 eeL.EyeMeasurementID,
                 2 AS EyeID
                 FROM EMR.dbo.tEyeExam AS eeL
               ) AS ee2
               ON ee2.ExamID = e.ExamID
               AND ee2.EyeID = ese.EyeID
               
               LEFT JOIN EMR.dbo.tVisualAcuity AS vaC
               ON vaC.VisualAcuityID = ee2.VisualAcuityID
               
               LEFT JOIN EMR.dbo.tLookupVA AS lvaCUCDVA
               ON lvaCUCDVA.VAID = vaC.UDVAID
               
               LEFT JOIN EMR.dbo.tLookupVA AS lvaCUCIVA
               ON lvaCUCIVA.VAID = vaC.UIVAID
               
               LEFT JOIN EMR.dbo.tLookupVA AS lvaCUCNVA
               ON lvaCUCNVA.VAID = vaC.UNVAID
               
               LEFT JOIN EMR.dbo.tVisualAcuity AS vaCBinoc
               ON vaCBinoc.VisualAcuityID = ee.BinocularVisualAcuityID
               
               LEFT JOIN EMR.dbo.tLookupVA AS lvaCBinocUCDVA
               ON lvaCBinocUCDVA.VAID = vaCBinoc.UDVAID
               
               LEFT JOIN EMR.dbo.tLookupVA AS lvaCBinocUCIVA
               ON lvaCBinocUCIVA.VAID = vaCBinoc.UIVAID
               
               LEFT JOIN EMR.dbo.tLookupVA AS lvaCBinocUCNVA
               ON lvaCBinocUCNVA.VAID = vaCBinoc.UNVAID

			  LEFT JOIN EMR.dbo.tLookupVA AS lvaCBinocBCIVA
			  ON lvaCBinocBCIVA.VAID = vaCBinoc.BCIVAID

               /*PostOp Day 1*/
                 OUTER APPLY
               (
                 SELECT TOP 1
                 SPO.PostOpDateTime As Day1PostOp, 
                 
                 lvaPUCDVA.SnellenMetre As D1_UCDVA_SM, 
                 lvaPUCDVA.LogMAR As D1_UCDVA_LM, 
                 lvaPUCDVA.Sort As D1_UCDVA_Sort, 
                 lvaPUCIVA.SnellenMetre As D1_UCIVA_SM, 
                 lvaPUCIVA.LogMAR As D1_UCIVA_LM, 
                 lvaPUCIVA.Sort As D1_UCIVA_Sort, 
                 lvaPUCNVA.NNotation As D1_UCNVA_N, 
                 lvaPUCNVA.LogMAR As D1_UCNVA_LM, 
                 lvaPUCNVA.Sort As D1_UCNVA_Sort, 
				 lvaPDCNVA.NNotation As D1_DCNVA_N, 
                 lvaPDCNVA.LogMAR As D1_DCNVA_LM, 
                 lvaPDCNVA.Sort As D1_DCNVA_Sort, 
				 lvaPDCIVA.SnellenMetre As D1_DCIVA_SM, 
                 lvaPDCIVA.LogMAR As D1_DCIVA_LM, 
                 lvaPDCIVA.Sort As D1_DCIVA_Sort, 
                 lvaPBCDVA.SnellenMetre As D1_BCDVA_SM,
                 lvaPBCDVA.LogMAR As D1_BCDVA_LM,
                 lvaPBCDVA.Sort As D1_BCDVA_Sort,
				 lvaPBCIVA.SnellenMetre As D1_BCIVA_SM,
                 lvaPBCIVA.LogMAR As D1_BCIVA_LM,
                 lvaPBCIVA.Sort As D1_BCIVA_Sort,
                 lvaPBinocUCDVA.SnellenMetre As D1_Binoc_UCDVA_SM,
                 lvaPBinocUCDVA.LogMAR As D1_Binoc_UCDVA_LM,
                 lvaPBinocUCDVA.Sort As D1_Binoc_UCDVA_Sort,
                 lvaPBinocUCIVA.SnellenMetre As D1_Binoc_UCIVA_SM,
                 lvaPBinocUCIVA.LogMAR As D1_Binoc_UCIVA_LM,
                 lvaPBinocUCIVA.Sort As D1_Binoc_UCIVA_Sort,
                 lvaPBinocUCNVA.NNotation As D1_Binoc_UCNVA_N,
                 lvaPBinocUCNVA.LogMAR As D1_Binoc_UCNVA_LM,
                 lvaPBinocUCNVA.Sort As D1_Binoc_UCNVA_Sort,
				 lvaPBinocDCNVA.NNotation As D1_Binoc_DCNVA_N,
                 lvaPBinocDCNVA.LogMAR As D1_Binoc_DCNVA_LM,
                 lvaPBinocDCNVA.Sort As D1_Binoc_DCNVA_Sort,
				 lvaPBinocDCIVA.SnellenMetre As D1_Binoc_DCIVA_SM,
                 lvaPBinocDCIVA.LogMAR As D1_Binoc_DCIVA_LM,
                 lvaPBinocDCIVA.Sort As D1_Binoc_DCIVA_Sort,
				 lvaPBinocBCIVA.SnellenMetre as D1_Binoc_BCIVA_SM,
				 lvaPBinocBCIVA.LogMAR as D1_Binoc_BCIVA_LM,
				 lvaPBinocBCIVA.Sort as D1_Binoc_BCIVA_Sort,
                 espoe.EyeID,
                 EPEP.Sphere,
                 EPEP.Cyl
                 
                 FROM emr.dbo.tSurgeryPostOp SPO
                 Left Join EMR.dbo.tEyeSurgeryPostOpEye espoe
                 On espoe.SurgeryPostOpID = SPO.SurgeryPostOpID
                 
                 Left Join EMR.dbo.tEyeSurgeryPostOp ESPO
                 On ESPO.SurgeryPostOpID = SPO.SurgeryPostOpID and ESPO.SurgeryPostOpID = espoe.SurgeryPostOpID
                 
                 Left Join EMR.dbo.tLookupVA lvaPUCDVA
                 On lvaPUCDVA.VAID = espoe.UAVADistanceID
                 
                 Left Join EMR.dbo.tLookupVA lvaPUCIVA
                 On lvaPUCIVA.VAID = espoe.UAVAIntermediateID
                 
                 Left Join EMR.dbo.tLookupVA lvaPUCNVA
                 On lvaPUCNVA.VAID = espoe.UAVANearID

                 
                 Left Join EMR.dbo.tEyePrescriptionEye EPEP
                 On EPEP.EyePrescriptionID = espo.ManifestPrescriptionID and EPEP.EyeID = espoe.EyeID
                 
                 LEFT JOIN EMR.dbo.tVisualAcuity AS vaP 
                 ON vaP.VisualAcuityID = epeP.VisualAcuityID
                 
                 LEFT JOIN EMR.dbo.tLookupVA AS lvaPBCDVA 
                 ON lvaPBCDVA.VAID = vaP.BCDVAID

				 LEFT JOIN EMR.dbo.tLookupVA AS lvaPBCIVA 
                 ON lvaPBCIVA.VAID = vaP.BCIVAID

				 LEFT JOIN EMR.dbo.tLookupVA AS lvaPDCNVA 
                 ON lvaPDCNVA.VAID = vaP.DCNVAID

				 LEFT JOIN EMR.dbo.tLookupVA AS lvaPDCIVA 
                 ON lvaPDCIVA.VAID = vaP.DCIVAID
                 
                 Left Join EMR.dbo.tLookupVA lvaPBinocUCDVA
                 On lvaPBinocUCDVA.VAID = espo.UAVABinocularDistanceID
                 
                 Left Join EMR.dbo.tLookupVA lvaPBinocUCIVA
                 On lvaPBinocUCIVA.VAID = espo.UAVABinocularIntermediateID
                 
                 Left Join EMR.dbo.tLookupVA lvaPBinocUCNVA
                 On lvaPBinocUCNVA.VAID = espo.UAVABinocularNearID
                 
				 Left Join EMR.dbo.tLookupVA lvaPBinocDCNVA
                 On lvaPBinocDCNVA.VAID = espo.DCVABinocularNearID

				 Left Join EMR.dbo.tLookupVA lvaPBinocDCIVA
                 On lvaPBinocDCIVA.VAID = espo.DCVABinocularIntermediateID

				 Left Join EMR.dbo.tLookupVA lvaPBinocBCIVA
                 On lvaPBinocBCIVA.VAID = espo.BCVABinocularIntermediateID
                 
                 WHERE SPO.PatientID = S.PatientID
                 and espoe.EyeID = ESE.EyeID
                 AND DATEDIFF (DAY, S.SurgeryDateTime, SPO.PostOpDateTime) <=6
                 ORDER BY s.SurgeryDateTime asc
               ) AS SPOD1
               
               /*PostOp Month 1*/
                 
                 OUTER APPLY
               (
                 SELECT TOP 1
                 SPO.PostOpDateTime As Month1PostOp, 
                 lvaPUCDVA.SnellenMetre As M1_UCDVA_SM, 
                 lvaPUCDVA.LogMAR As M1_UCDVA_LM, 
                 lvaPUCDVA.NNotation As M1_UCDVA_N, 
                 lvaPUCIVA.SnellenMetre As M1_UCIVA_SM, 
                 lvaPUCIVA.LogMAR As M1_UCIVA_LM, 
                 lvaPUCIVA.NNotation As M1_UCIVA_N, 
                 lvaPUCNVA.SnellenMetre As M1_UCNVA_SM, 
                 lvaPUCNVA.LogMAR As M1_UCNVA_LM, 
                 lvaPUCNVA.NNotation As M1_UCNVA_N, 
				         lvaPDCNVA.SnellenMetre As M1_DCNVA_SM, 
                 lvaPDCNVA.LogMAR As M1_DCNVA_LM, 
                 lvaPDCNVA.NNotation As M1_DCNVA_N, 
				         lvaPDCIVA.SnellenMetre As M1_DCIVA_SM, 
                 lvaPDCIVA.LogMAR As M1_DCIVA_LM, 
                 lvaPDCIVA.NNotation As M1_DCIVA_N, 
                 lvaPBCDVA.SnellenMetre As M1_BCDVA_SM,
                 lvaPBCDVA.LogMAR As M1_BCDVA_LM,
                 lvaPBCDVA.NNotation As M1_BCDVA_N,
				         lvaPBCIVA.SnellenMetre As M1_BCIVA_SM,
                 lvaPBCIVA.LogMAR As M1_BCIVA_LM,
                 lvaPBCIVA.NNotation As M1_BCIVA_N,
                 lvaPBCNVA.SnellenMetre As M1_BCNVA_SM,
                 lvaPBCNVA.LogMAR As M1_BCNVA_LM,
                 lvaPBCNVA.NNotation As M1_BCNVA_N,
                 lvaPBinocUCDVA.SnellenMetre As M1_Binoc_UCDVA_SM,
                 lvaPBinocUCDVA.LogMAR As M1_Binoc_UCDVA_LM,
                 lvaPBinocUCDVA.NNotation As M1_Binoc_UCDVA_N,
                 lvaPBinocUCIVA.SnellenMetre As M1_Binoc_UCIVA_SM,
                 lvaPBinocUCIVA.LogMAR As M1_Binoc_UCIVA_LM,
                 lvaPBinocUCIVA.NNotation As M1_Binoc_UCIVA_N,
                 lvaPBinocUCNVA.SnellenMetre As M1_Binoc_UCNVA_SM,
                 lvaPBinocUCNVA.LogMAR As M1_Binoc_UCNVA_LM,
                 lvaPBinocUCNVA.NNotation As M1_Binoc_UCNVA_N,
				         lvaPBinocDCNVA.SnellenMetre As M1_Binoc_DCNVA_SM,
                 lvaPBinocDCNVA.LogMAR As M1_Binoc_DCNVA_LM,
                 lvaPBinocDCNVA.NNotation As M1_Binoc_DCNVA_N,
				         lvaPBinocDCIVA.SnellenMetre As M1_Binoc_DCIVA_SM,
                 lvaPBinocDCIVA.LogMAR As M1_Binoc_DCIVA_LM,
                 lvaPBinocDCIVA.NNotation As M1_Binoc_DCIVA_N,
				         lvaPBinocBCIVA.SnellenMetre as M1_Binoc_BCIVA_SM,
				         lvaPBinocBCIVA.LogMAR as M1_Binoc_BCIVA_LM,
				         lvaPBinocBCIVA.NNotation as M1_Binoc_BCIVA_N,
				         lvaPBinocBCNVA.SnellenMetre as M1_Binoc_BCNVA_SM,
				         lvaPBinocBCNVA.LogMAR as M1_Binoc_BCNVA_LM,
				         lvaPBinocBCNVA.NNotation as M1_Binoc_BCNVA_N,
                 espoe.EyeID,
                 EPEP.Sphere as Sph_M1,
                 EPEP.Cyl as Cyl_M1
                 
                 FROM emr.dbo.tSurgeryPostOp SPO
                 Left Join EMR.dbo.tEyeSurgeryPostOpEye espoe
                 On espoe.SurgeryPostOpID = SPO.SurgeryPostOpID
                 
                 Left Join EMR.dbo.tEyeSurgeryPostOp ESPO
                 On ESPO.SurgeryPostOpID = SPO.SurgeryPostOpID and ESPO.SurgeryPostOpID = espoe.SurgeryPostOpID
                 
                 Left Join EMR.dbo.tLookupVA lvaPUCDVA
                 On lvaPUCDVA.VAID = espoe.UAVADistanceID
                 
                 Left Join EMR.dbo.tLookupVA lvaPUCIVA
                 On lvaPUCIVA.VAID = espoe.UAVAIntermediateID
                 
                 Left Join EMR.dbo.tLookupVA lvaPUCNVA
                 On lvaPUCNVA.VAID = espoe.UAVANearID

                 
                 Left Join EMR.dbo.tEyePrescriptionEye EPEP
                 On EPEP.EyePrescriptionID = espo.ManifestPrescriptionID and EPEP.EyeID = espoe.EyeID
                 
                 LEFT JOIN EMR.dbo.tVisualAcuity AS vaP 
                 ON vaP.VisualAcuityID = epeP.VisualAcuityID
                 
                 LEFT JOIN EMR.dbo.tLookupVA AS lvaPBCDVA 
                 ON lvaPBCDVA.VAID = vaP.BCDVAID

				 LEFT JOIN EMR.dbo.tLookupVA AS lvaPBCIVA 
                 ON lvaPBCIVA.VAID = vaP.BCIVAID
                 
				 LEFT JOIN EMR.dbo.tLookupVA AS lvaPBCNVA 
                 ON lvaPBCNVA.VAID = vaP.BCNVAID

				 LEFT JOIN EMR.dbo.tLookupVA AS lvaPDCNVA 
                 ON lvaPDCNVA.VAID = vaP.DCNVAID

				 LEFT JOIN EMR.dbo.tLookupVA AS lvaPDCIVA 
                 ON lvaPDCIVA.VAID = vaP.DCIVAID
                 
                 Left Join EMR.dbo.tLookupVA lvaPBinocUCDVA
                 On lvaPBinocUCDVA.VAID = espo.UAVABinocularDistanceID
                 
                 Left Join EMR.dbo.tLookupVA lvaPBinocUCIVA
                 On lvaPBinocUCIVA.VAID = espo.UAVABinocularIntermediateID
                 
                 Left Join EMR.dbo.tLookupVA lvaPBinocUCNVA
                 On lvaPBinocUCNVA.VAID = espo.UAVABinocularNearID
                 
				 Left Join EMR.dbo.tLookupVA lvaPBinocDCNVA
                 On lvaPBinocDCNVA.VAID = espo.DCVABinocularNearID

				 Left Join EMR.dbo.tLookupVA lvaPBinocDCIVA
                 On lvaPBinocDCIVA.VAID = espo.DCVABinocularIntermediateID

				 Left Join EMR.dbo.tLookupVA lvaPBinocBCIVA
                 On lvaPBinocBCIVA.VAID = espo.BCVABinocularIntermediateID
                 
                 Left Join EMR.dbo.tLookupVA lvaPBinocBCNVA
                 On lvaPBinocBCNVA.VAID = espo.BCVABinocularNearID
                 
                 WHERE SPO.PatientID = S.PatientID
                 and espoe.EyeID = ESE.EyeID
                 AND (DATEDIFF (DAY, S.SurgeryDateTime, SPO.PostOpDateTime) >=20 and DATEDIFF (DAY, S.SurgeryDateTime, SPO.PostOpDateTime) <=40)
                 ORDER BY SPO.PostOpDateTime DESC
               ) AS SPOM1

			   			  /* Patient Reported Outcomes */

			   OUTER APPLY ( Select TOP (1)
rh.PatientID, rh.SurveyDate, ol.OptionText
  FROM peq.dbo.tResultHeader rh (NOLOCK)
JOIN peq.dbo.tResultLine rl (NOLOCK) ON rl.ResultHeaderID = rh.ResultHeaderID
JOIN peq.dbo.tQuestion_Option qo (NOLOCK) ON qo.QuestionOptionID = rl.QuestionOptionID
JOIN peq.dbo.tOption_Language ol (NOLOCK) ON ol.OptionID = qo.OptionID  and ol.LanguageID = 2057
WHERE qo.QuestionID in (336,338,1013,1123,1152,1281)
and rh.PatientID = e.PatientID
AND COALESCE(rl.IsDeleted, 0) = 0
and (DATEDIFF (DAY, S.SurgeryDateTime,rh.SurveyDate) >=", postopdatemin," and DATEDIFF (DAY, S.SurgeryDateTime, rh.SurveyDate) <=", postopdatemax,")
and rh.SurveyID in (13,14,17,18,21,28,40,43,44,48,54,56)) AS Q1
			
   OUTER APPLY ( Select TOP (1)
rh.PatientID, rh.SurveyDate, ol.OptionText
  FROM peq.dbo.tResultHeader rh (NOLOCK)
JOIN peq.dbo.tResultLine rl (NOLOCK) ON rl.ResultHeaderID = rh.ResultHeaderID
JOIN peq.dbo.tQuestion_Option qo (NOLOCK) ON qo.QuestionOptionID = rl.QuestionOptionID
JOIN peq.dbo.tOption_Language ol (NOLOCK) ON ol.OptionID = qo.OptionID  and ol.LanguageID = 2057
WHERE qo.QuestionID in (402, 416,1018,1286)
and rh.PatientID = e.PatientID
AND COALESCE(rl.IsDeleted, 0) = 0
and rh.SurveyDate = Q1.SurveyDate
and rh.SurveyID in (1,2,3,4,5,7,8,9,18,21,28,40, 48, 54)) AS Q2  

   OUTER APPLY ( Select TOP (1)
rh.PatientID, rh.SurveyDate, ol.OptionText
  FROM peq.dbo.tResultHeader rh (NOLOCK)
JOIN peq.dbo.tResultLine rl (NOLOCK) ON rl.ResultHeaderID = rh.ResultHeaderID
JOIN peq.dbo.tQuestion_Option qo (NOLOCK) ON qo.QuestionOptionID = rl.QuestionOptionID
JOIN peq.dbo.tOption_Language ol (NOLOCK) ON ol.OptionID = qo.OptionID  and ol.LanguageID = 2057
WHERE qo.QuestionID in (339, 343, 341, 1016, 1100, 1129, 1178, 1284)
and rh.PatientID = e.PatientID
AND COALESCE(rl.IsDeleted, 0) = 0
and rh.SurveyDate = Q1.SurveyDate
and rh.SurveyID in (13,14,17,18,21,28,40,43,44,46,48,54)) AS Q3

   OUTER APPLY ( Select TOP (1)
rh.PatientID, rh.SurveyDate, ol.OptionText
  FROM peq.dbo.tResultHeader rh (NOLOCK)
JOIN peq.dbo.tResultLine rl (NOLOCK) ON rl.ResultHeaderID = rh.ResultHeaderID
JOIN peq.dbo.tQuestion_Option qo (NOLOCK) ON qo.QuestionOptionID = rl.QuestionOptionID
JOIN peq.dbo.tOption_Language ol (NOLOCK) ON ol.OptionID = qo.OptionID  and ol.LanguageID = 2057
WHERE qo.QuestionID in (129,1022,1052,1057,1070,1075,1106,1135,1290)
and rh.PatientID = e.PatientID
AND COALESCE(rl.IsDeleted, 0) = 0
and rh.SurveyDate = Q1.SurveyDate
and rh.SurveyID in (18,21,22,27,28,40,41,41,42,42,43,44,48,54)) AS Q4

   OUTER APPLY ( Select TOP (1)
rh.PatientID, rh.SurveyDate, ol.OptionText
  FROM peq.dbo.tResultHeader rh (NOLOCK)
JOIN peq.dbo.tResultLine rl (NOLOCK) ON rl.ResultHeaderID = rh.ResultHeaderID
JOIN peq.dbo.tQuestion_Option qo (NOLOCK) ON qo.QuestionOptionID = rl.QuestionOptionID
JOIN peq.dbo.tOption_Language ol (NOLOCK) ON ol.OptionID = qo.OptionID  and ol.LanguageID = 2057
WHERE qo.QuestionID in (133,1023,1107,1136,1291)
and rh.PatientID = e.PatientID
AND COALESCE(rl.IsDeleted, 0) = 0
and rh.SurveyDate = Q1.SurveyDate
and rh.SurveyID in (18,21,28,40,43,44,48,54)) AS Q5

   OUTER APPLY ( Select TOP (1)
rh.PatientID, rh.SurveyDate, ol.OptionText
  FROM peq.dbo.tResultHeader rh (NOLOCK)
JOIN peq.dbo.tResultLine rl (NOLOCK) ON rl.ResultHeaderID = rh.ResultHeaderID
JOIN peq.dbo.tQuestion_Option qo (NOLOCK) ON qo.QuestionOptionID = rl.QuestionOptionID
JOIN peq.dbo.tOption_Language ol (NOLOCK) ON ol.OptionID = qo.OptionID  and ol.LanguageID = 2057
WHERE qo.QuestionID in (102,1024,1054,1059,1072,1077,1108,1137,1292)
and rh.PatientID = e.PatientID
AND COALESCE(rl.IsDeleted, 0) = 0
and rh.SurveyDate = Q1.SurveyDate
and rh.SurveyID in (18,21,22,27,28,40,41,41,42,42,43,44,48,54)) AS Q6

   OUTER APPLY ( Select TOP (1)
rh.PatientID, rh.SurveyDate, ol.OptionText
  FROM peq.dbo.tResultHeader rh (NOLOCK)
JOIN peq.dbo.tResultLine rl (NOLOCK) ON rl.ResultHeaderID = rh.ResultHeaderID
JOIN peq.dbo.tQuestion_Option qo (NOLOCK) ON qo.QuestionOptionID = rl.QuestionOptionID
JOIN peq.dbo.tOption_Language ol (NOLOCK) ON ol.OptionID = qo.OptionID  and ol.LanguageID = 2057
WHERE qo.QuestionID in (317,1021,1051,1056,1069,1074,1105,1134,1289)
and rh.PatientID = e.PatientID
AND COALESCE(rl.IsDeleted, 0) = 0
and rh.SurveyDate = Q1.SurveyDate
and rh.SurveyID in (18,21,22,27,28,40,41,41,42,42,43,44,48,54)) AS Q7

   OUTER APPLY ( Select TOP (1)
rh.PatientID, rh.SurveyDate, ol.OptionText
  FROM peq.dbo.tResultHeader rh (NOLOCK)
JOIN peq.dbo.tResultLine rl (NOLOCK) ON rl.ResultHeaderID = rh.ResultHeaderID
JOIN peq.dbo.tQuestion_Option qo (NOLOCK) ON qo.QuestionOptionID = rl.QuestionOptionID
JOIN peq.dbo.tOption_Language ol (NOLOCK) ON ol.OptionID = qo.OptionID  and ol.LanguageID = 2057
WHERE qo.QuestionID in (106,1020,1050,1055,1068,1073,1104,1133,1288)
and rh.PatientID = e.PatientID
AND COALESCE(rl.IsDeleted, 0) = 0
and rh.SurveyDate = Q1.SurveyDate
and rh.SurveyID in (18,21,22,27,28,40,41,41,42,42,43,44,48,54)) AS Q8

   OUTER APPLY ( Select TOP (1)
rh.PatientID, rh.SurveyDate, ol.OptionText
  FROM peq.dbo.tResultHeader rh (NOLOCK)
JOIN peq.dbo.tResultLine rl (NOLOCK) ON rl.ResultHeaderID = rh.ResultHeaderID
JOIN peq.dbo.tQuestion_Option qo (NOLOCK) ON qo.QuestionOptionID = rl.QuestionOptionID
JOIN peq.dbo.tOption_Language ol (NOLOCK) ON ol.OptionID = qo.OptionID  and ol.LanguageID = 2057
WHERE qo.QuestionID in (24,26,928,959,1025,1060,1078,1109,1138,1293)
and rh.PatientID = e.PatientID
AND COALESCE(rl.IsDeleted, 0) = 0
and rh.SurveyDate = Q1.SurveyDate
and rh.SurveyID in (18,21,22,27,28,40,41,41,42,42,43,44,48,54)) AS Q9

   OUTER APPLY ( Select TOP (1)
rh.PatientID, rh.SurveyDate, ol.OptionText
  FROM peq.dbo.tResultHeader rh (NOLOCK)
JOIN peq.dbo.tResultLine rl (NOLOCK) ON rl.ResultHeaderID = rh.ResultHeaderID
JOIN peq.dbo.tQuestion_Option qo (NOLOCK) ON qo.QuestionOptionID = rl.QuestionOptionID
JOIN peq.dbo.tOption_Language ol (NOLOCK) ON ol.OptionID = qo.OptionID  and ol.LanguageID = 2057
WHERE qo.QuestionID in (64,1029,1164,1297)
and rh.PatientID = e.PatientID
AND COALESCE(rl.IsDeleted, 0) = 0
and rh.SurveyDate = Q1.SurveyDate
and rh.SurveyID in (18,21,22,27,28,40,41,41,42,42,43,44,48,54)) AS Q10

   OUTER APPLY ( Select TOP (1)
rh.PatientID, rh.SurveyDate, ol.OptionText
  FROM peq.dbo.tResultHeader rh (NOLOCK)
JOIN peq.dbo.tResultLine rl (NOLOCK) ON rl.ResultHeaderID = rh.ResultHeaderID
JOIN peq.dbo.tQuestion_Option qo (NOLOCK) ON qo.QuestionOptionID = rl.QuestionOptionID
JOIN peq.dbo.tOption_Language ol (NOLOCK) ON ol.OptionID = qo.OptionID  and ol.LanguageID = 2057
WHERE qo.QuestionID in (135, 1017, 1101, 1130, 1179, 1285)
and rh.PatientID = e.PatientID
AND COALESCE(rl.IsDeleted, 0) = 0
and rh.SurveyDate = Q1.SurveyDate
and rh.SurveyID in (13,14,17,18,21,28,40,43,44,46,48,54)) AS Q11

   OUTER APPLY ( Select TOP (1)
rh.PatientID, rh.SurveyDate, ol.OptionText
  FROM peq.dbo.tResultHeader rh (NOLOCK)
JOIN peq.dbo.tResultLine rl (NOLOCK) ON rl.ResultHeaderID = rh.ResultHeaderID
JOIN peq.dbo.tQuestion_Option qo (NOLOCK) ON qo.QuestionOptionID = rl.QuestionOptionID
JOIN peq.dbo.tOption_Language ol (NOLOCK) ON ol.OptionID = qo.OptionID  and ol.LanguageID = 2057
WHERE qo.QuestionID in (109,177,22,1026,1110,1139,1294)
and rh.PatientID = e.PatientID
AND COALESCE(rl.IsDeleted, 0) = 0
and rh.SurveyDate = Q1.SurveyDate
and rh.SurveyID in (13,14,17,18,21,28,40,43,44,46,48,54)) AS Q12


   OUTER APPLY ( Select TOP (1)
rh.PatientID, rh.SurveyDate, ol.OptionText
  FROM peq.dbo.tResultHeader rh (NOLOCK)
JOIN peq.dbo.tResultLine rl (NOLOCK) ON rl.ResultHeaderID = rh.ResultHeaderID
JOIN peq.dbo.tQuestion_Option qo (NOLOCK) ON qo.QuestionOptionID = rl.QuestionOptionID
JOIN peq.dbo.tOption_Language ol (NOLOCK) ON ol.OptionID = qo.OptionID  and ol.LanguageID = 2057
WHERE qo.QuestionID in (111,31,28,1027,1111,1140,1295)
and rh.PatientID = e.PatientID
AND COALESCE(rl.IsDeleted, 0) = 0
and rh.SurveyDate = Q1.SurveyDate
and rh.SurveyID in (13,14,17,18,21,28,40,43,44,46,48,54)) AS Q13
	
where s.SurgeryID in (", sx , ")
/*
S.SurgeryDateTime between   '", mindate , "' AND '", maxdate , "'
and lr.RegionID in (1,2,3,4,5,6)
and ilc.IOLLensID in (101,103,106,105,108,109,110,111,112)
and lstg.SurgeryTypeGroupID in (2)
and ese.SurgeryTreatmentTypeID = 1 --primary
and ese.SurgeryTypeID not in (23,24) --exchanges
and S.IsDeleted = 0
and e.IsDeleted = 0
and ese.IsComplete in (1)
*/
order by S.PatientID desc", sep="")

peq1 <- dbGetQuery(con, query)
peq <- peq1

#### DATA CLEANING ####
data <- merge(data,peq,by=intersect(colnames(data),colnames(peq)),all.x=TRUE)

test <- as.data.frame(table(data$PatientID, data$EyeID)) %>% filter(Freq>1)
if (dim(test)[1]>0){view(test)
  stop("There are duplicates")}

data <- data %>%
  mutate(GroupMonocular1 = case_when(M1_UCDVA_LM == "6/3" ~ "6/5 or better",
                                     M1_UCDVA_SM == "6/3.8" ~ "6/5 or better",
                                     M1_UCDVA_SM == "6/4" ~ "6/5 or better",
                                     M1_UCDVA_SM == "6/4.8" ~ "6/5 or better",
                                     M1_UCDVA_SM == "6/5" ~ "6/5 or better",
                                     M1_UCDVA_SM == "6/6" ~ "6/6 or poorer",
                                     M1_UCDVA_SM == "6/7.5" ~ "6/6 or poorer",
                                     M1_UCDVA_SM == "6/9" ~ "6/6 or poorer",
                                     M1_UCDVA_SM == "6/9.5" ~ "6/6 or poorer",
                                     M1_UCDVA_SM == "6/10" ~ "6/6 or poorer",
                                     M1_UCDVA_SM == "6/12" ~ "6/6 or poorer",
                                     M1_UCDVA_SM == "6/15" ~ "6/6 or poorer",
                                     M1_UCDVA_SM == "6/18" ~ "6/6 or poorer",
                                     M1_UCDVA_SM == "6/19" ~ "6/6 or poorer",
                                     M1_UCDVA_SM == "6/20" ~ "6/6 or poorer",
                                     M1_UCDVA_SM == "6/21" ~ "6/6 or poorer",
                                     M1_UCDVA_SM == "6/24" ~ "6/6 or poorer",
                                     M1_UCDVA_SM == "6/30" ~ "6/6 or poorer",
                                     M1_UCDVA_SM == "6/34.2" ~ "6/6 or poorer",
                                     M1_UCDVA_SM == "6/36" ~ "6/6 or poorer",
                                     M1_UCDVA_SM == "6/38" ~ "6/6 or poorer",
                                     M1_UCDVA_SM == "6/48" ~ "6/6 or poorer",
                                     M1_UCDVA_SM == "6/60" ~ "6/6 or poorer",
                                     M1_UCDVA_SM == "6/75" ~ "6/6 or poorer",
                                     M1_UCDVA_SM == "6/95" ~ "6/6 or poorer",
                                     M1_UCDVA_SM == "6/120" ~ "6/6 or poorer",
                                     M1_UCDVA_SM == "6/240" ~ "6/6 or poorer",
                                     M1_UCDVA_SM == "CF" ~ "6/6 or poorer",
                                     M1_UCDVA_SM == "HM" ~ "6/6 or poorer",
                                     M1_UCDVA_SM == "POL" ~ "6/6 or poorer",
                                     M1_UCDVA_SM == "NPOL" ~ "6/6 or poorer"))

data <- data %>%
  mutate(GroupBinocular1 = case_when(M1_Binoc_UCDVA_SM == "6/3" ~ "6/6 or better",
                                     M1_Binoc_UCDVA_SM == "6/3.8" ~ "6/6 or better",
                                     M1_Binoc_UCDVA_SM == "6/4" ~ "6/6 or better",
                                     M1_Binoc_UCDVA_SM == "6/4.8" ~ "6/6 or better",
                                     M1_Binoc_UCDVA_SM == "6/5" ~ "6/6 or better",
                                     M1_Binoc_UCDVA_SM == "6/6" ~ "6/6 or better",
                                     M1_Binoc_UCDVA_SM == "6/7.5" ~ "6/7.5 or poorer",
                                     M1_Binoc_UCDVA_SM == "6/9" ~ "6/7.5 or poorer",
                                     M1_Binoc_UCDVA_SM == "6/9.5" ~ "6/7.5 or poorer",
                                     M1_Binoc_UCDVA_SM == "6/10" ~ "6/7.5 or poorer",
                                     M1_Binoc_UCDVA_SM == "6/12" ~ "6/7.5 or poorer",
                                     M1_Binoc_UCDVA_SM == "6/15" ~ "6/7.5 or poorer",
                                     M1_Binoc_UCDVA_SM == "6/18" ~ "6/7.5 or poorer",
                                     M1_Binoc_UCDVA_SM == "6/19" ~ "6/7.5 or poorer",
                                     M1_Binoc_UCDVA_SM == "6/20" ~ "6/7.5 or poorer",
                                     M1_Binoc_UCDVA_SM == "6/21" ~ "6/7.5 or poorer",
                                     M1_Binoc_UCDVA_SM == "6/24" ~ "6/7.5 or poorer",
                                     M1_Binoc_UCDVA_SM == "6/30" ~ "6/7.5 or poorer",
                                     M1_Binoc_UCDVA_SM == "6/34.2" ~ "6/7.5 or poorer",
                                     M1_Binoc_UCDVA_SM == "6/36" ~ "6/7.5 or poorer",
                                     M1_Binoc_UCDVA_SM == "6/38" ~ "6/7.5 or poorer",
                                     M1_Binoc_UCDVA_SM == "6/48" ~ "6/7.5 or poorer",
                                     M1_Binoc_UCDVA_SM == "6/60" ~ "6/7.5 or poorer",
                                     M1_Binoc_UCDVA_SM == "6/75" ~ "6/7.5 or poorer",
                                     M1_Binoc_UCDVA_SM == "6/95" ~ "6/7.5 or poorer",
                                     M1_Binoc_UCDVA_SM == "6/120" ~ "6/7.5 or poorer",
                                     M1_Binoc_UCDVA_SM == "6/240" ~ "6/7.5 or poorer",
                                     M1_Binoc_UCDVA_SM == "CF" ~ "6/7.5 or poorer",
                                     M1_Binoc_UCDVA_SM == "HM" ~ "6/7.5 or poorer",
                                     M1_Binoc_UCDVA_SM == "POL" ~ "6/7.5 or poorer",
                                     M1_Binoc_UCDVA_SM == "NPOL" ~ "6/7.5 or poorer"))

#### Motivation vs Satisfaction vs Lens ####

if (postopdatemin==20){
  temp <- read.xlsx('S:\\Clinical Services\\Research Assistants\\Tecnis Odyssey\\Bilateral Odyssey vs Mix PureSee\\tables\\pxsreasons.xlsx')
  temp <- merge(temp,data %>% select(PatientID, EyeID, M1_Satisfied, SurgeryLens),by=c("PatientID","EyeID"))
  
  tab1 <- as.data.frame(table(temp$Reason, temp$M1_Satisfied, temp$SurgeryLens)) %>%
    pivot_wider(names_from="Var2", values_from="Freq") %>%
    select(Var1, Var3, `Very Satisfied`, Satisfied, Neither, Dissatisfied, `Very Dissatisfied`)
  tab2 <- as.data.frame(table(temp$Reason[which(temp$M1_UCDVA_LM<=0)], temp$M1_Satisfied[which(temp$M1_UCDVA_LM<=0)], temp$SurgeryLens[which(temp$M1_UCDVA_LM<=0)])) %>%
    pivot_wider(names_from="Var2", values_from="Freq") %>%
    select(Var1, Var3, `Very Satisfied`, Satisfied, Neither, Dissatisfied, `Very Dissatisfied`)
  
  list <- list("All" = tab1, "UCDVA 6+" = tab2)
  write.xlsx(list,'S:\\Clinical Services\\Research Assistants\\Tecnis Odyssey\\Bilateral Odyssey vs Mix PureSee\\tables\\pxsreasons_tables.xlsx')
}

#### Working ####
completionrate <- dim(data[which(is.na(data$M1_Recommend)==FALSE),])[1]/dim(data)[1]

u <- data %>%
  select(PatientID, M1_Recommend, M1_Satisfied, M1_DoOver, M1_Glare, M1_Halos, M1_DryEyes, M1_NightDrive, GroupMonocular1, GroupBinocular1, SurgeryLens)
u <- unique(u)
u1 <- subset(u,u$GroupBinocular1=='6/6 or better')

#### Recommendation ####
rec <- as.data.frame(table(u$SurgeryLens, u$M1_Recommend)) %>%
  pivot_wider(names_from="Var1", values_from="Freq")

#### Recommendation Mono BO ####
bo <- u[which(u$SurgeryLens=="Bilateral Odyssey"),]
rec_mono_bo <- as.data.frame(table(bo$M1_Recommend, bo$GroupMonocular1))%>%
  pivot_wider(names_from = "Var2", values_from = "Freq")

#### Recommendation Mono BP ####
bp <- u[which(u$SurgeryLens=="Bilateral PureSee"),]
rec_mono_bp <- as.data.frame(table(bp$M1_Recommend, bp$GroupMonocular1))%>%
  pivot_wider(names_from = "Var2", values_from = "Freq")

#### Recommendation Mono MM ####
mm <- u[which(u$SurgeryLens=="Mix and Match"),]
rec_mono_m <- as.data.frame(table(mm$M1_Recommend, mm$GroupMonocular1))%>%
  pivot_wider(names_from = "Var2", values_from = "Freq")

#### Recommendation Bino BO ####
rec_bino_bo <- as.data.frame(table(bo$M1_Recommend, bo$GroupBinocular1)) %>%
  pivot_wider(names_from = "Var2", values_from = "Freq")

#### Recommendation Bino BP ####
rec_bino_bp <- as.data.frame(table(bp$M1_Recommend, bp$GroupBinocular1)) %>%
  pivot_wider(names_from = "Var2", values_from = "Freq")

#### Recommendation Bino MM ####
rec_bino_m <- as.data.frame(table(mm$M1_Recommend, mm$GroupBinocular1)) %>%
  pivot_wider(names_from = "Var2", values_from = "Freq")

#### Satisfaction ####
sat <- as.data.frame(table(u$M1_Satisfied, u$SurgeryLens))%>%
  pivot_wider(names_from="Var2", values_from="Freq")

#### Satisfaction (6/6 or better binocularly) ####
sat1 <- as.data.frame(table(u1$M1_Satisfied, u1$SurgeryLens))%>%
  pivot_wider(names_from="Var2", values_from="Freq")

#### Do Over ####
doover <- as.data.frame(table(u$M1_DoOver, u$SurgeryLens))%>%
  pivot_wider(names_from="Var2", values_from="Freq")

#### Do Over (6/6 or better binocularly) ####
doover1 <- as.data.frame(table(u1$M1_DoOver, u1$SurgeryLens))%>%
  pivot_wider(names_from="Var2", values_from="Freq")

#### Glare ####
glare <- as.data.frame(table(u$M1_Glare, u$SurgeryLens))%>%
  pivot_wider(names_from="Var2", values_from="Freq")

#### Glare (6/6 or better binocularly) ####
glare1 <- as.data.frame(table(u1$M1_Glare, u1$SurgeryLens))%>%
  pivot_wider(names_from="Var2", values_from="Freq")

#### Haloes ####
haloes <- as.data.frame(table(u$M1_Halos, u$SurgeryLens))%>%
  pivot_wider(names_from="Var2", values_from="Freq")

#### Haloes (6/6 or better binocularly) #### 

haloes1 <- as.data.frame(table(u1$M1_Halos, u1$SurgeryLens))%>%
  pivot_wider(names_from="Var2", values_from="Freq")

#### Dry Eye ####
dryeye <- as.data.frame(table(u$M1_DryEyes, u$SurgeryLens))%>%
  pivot_wider(names_from="Var2", values_from="Freq")

#### Night Driving ####
drive <- as.data.frame(table(u$M1_NightDrive, u$SurgeryLens))%>%
  pivot_wider(names_from="Var2", values_from="Freq")

#### Night Driving (6/6 or better binocularly) ####
drive1 <- as.data.frame(table(u1$M1_NightDrive, u1$SurgeryLens))%>%
  pivot_wider(names_from="Var2", values_from="Freq")

#### GRAPHS ####
graph <- list("Completion Rate" = completionrate,
              "Recommendation" = rec,
              "Recommendation - Monoc BO" = rec_mono_bo,
              "Recommendation - Monoc BP" = rec_mono_bp,
              "Recommendation - Monoc MM" = rec_mono_m,
              "Recommendation - Binoc BO" = rec_bino_bo,
              "Recommendation - Binoc BP" = rec_bino_bp,
              "Recommendation - Binoc MM" = rec_bino_m,
              "Satisfaction" = sat,
              "Satisfaction 6" = sat1,
              "Do Over" = doover,
              "Do Over 6" = doover1,
              "Glare" = glare,
              "Glare 6" = glare1,
              "Haloes" = haloes,
              "Haloes 6" = haloes1,
              "Dry Eyes" = dryeye,
              "Night Drive" = drive,
              "Night Drive 6" = drive1)

if (isTRUE(postopdatemin==20)==TRUE){
  write.xlsx(graph, paste(path, "PRO Data 1.xlsx", sep = ""))
  print("Month 1 was updated!")
} else{
  write.xlsx(graph, paste(path, "PRO Data 3.xlsx", sep = ""))
  print("Month 3 was updated!")
}
test <- as.data.frame(table(data$PatientID, data$EyeID))
