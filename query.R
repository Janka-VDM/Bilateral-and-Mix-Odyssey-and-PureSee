'
This is the query thats the same from IMAB C+PROs. It has all the usuals,
but there is also the Motivation required, which is taken either from the IOL
Pre Assessment before the surgery, or its taken from the Refractive Consultation,
which is dealt with by a case. This is only present in the top half of the query,
as the other half is for Clinical Reviews.
'


query <- paste('SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED
Select Distinct
/*Con*/
  
S.PatientID,
1 as QueryType,
P.Gender,
floor((datediff(day,0,s.SurgeryDateTime) - datediff(day,0,P.DOB)) / 365.2425) as Age,
E.ExamDateTime,
ese.EyeID,
case
  when iol_press.DominantEyeID is not null then iol_press.DominantEyeID
  when iol_press.DominantEyeID is null then ees.DominantEyeID
end as DominantEyeID,
case
  when iol_press.Motivation is not null then iol_press.Motivation
  when iol_press.Motivation is null then lm.DescriptionEnGB
end as Motivation,
ese.Target,

/* Con Refractive */
epeM.Sphere,
epeM.Cyl,
epeM.AddNear,

eme.Keratometry1 Keratometry_Power_1,
eme.Keratometry2 Keratometry_Power_2,

lvaCUCDVA.SnellenMetre As UCDVA_SM,
lvaCUCDVA.LogMAR As UCDVA_LM,

lvaCUCIVA.SnellenMetre As UCIVA_SM,
lvaCUCIVA.LogMAR AS UCIVA_LM,

lvaCUCNVA.SnellenMetre As UCNVA_SM,
lvaCUCNVA.LogMAR AS UCNVA_LM,

lvaCBCDVA.LogMAR AS BCDVA_LM,

lvaCBinocUCDVA.SnellenMetre As Binoc_UCDVA_SM,
lvaCBinocUCDVA.LogMAR As Binoc_UCDVA_LM,
lvaCBinocUCDVA.Decimal As Binoc_UCDVA_D,

lvaCBinocUCIVA.SnellenMetre As Binoc_UCIVA_SM,
lvaCBinocUCIVA.LogMAR As Binoc_UCIVA_LM,

lvaCBinocUCNVA.SnellenMetre As Binoc_UCNVA_SM,
lvaCBinocUCNVA.LogMAR As Binoc_UCNVA_LM,

lvaBinocCBCDVA.SnellenMetre as Binoc_BCDVA_SM,
lvaBinocCBCDVA.LogMAR as Binoc_BCDVA_LM,

LL.DescriptionEnGB AS Lens,

/*Surgery*/
  
  S.SurgeryDateTime,
s.SurgeryID,
s.SurgeonID,
s.SurgeryLocationID,
CASE
  WHEN (LVCSurgery.SurgeryDateTime IS NOT NULL) OR (esei.PriorLVC = 1) THEN 1
  ELSE 0
END AS PriorLVC,
lst.DescriptionEnGB As Surgery,
lstt.DescriptionEnGB As Surgery_Type,
lstg.DescriptionEnGB As Surgery_Group,
lil.Multi As Multifocal,
lil.Toric,
lil.DescriptionEnGB As Model,
lc.DescriptionEnGB as Complication,

/*M1 Post Op*/
  
  SPOM1.Month1PostOp,

SPOM1.Sph_M1,
SPOM1.Cyl_M1,

SPOM1.M1_UCDVA_LM,
SPOM1.M1_UCIVA_LM,
SPOM1.M1_UCNVA_LM,
SPOM1.M1_BCDVA_LM,
SPOM1.M1_BCDVA_D,
SPOM1.M1_Binoc_UCDVA_LM,
SPOM1.M1_Binoc_UCDVA_D,
SPOM1.M1_Binoc_UCIVA_LM,
SPOM1.M1_Binoc_UCNVA_LM,
SPOM1.M1_Binoc_BCDVA_D,
SPOM1.M1_Binoc_BCDVA_LM,

Survey.SurveyDate as M1_SurveyDate,

/* Enhancements */
  enh.SurgeryDateTime AS EnhDate,
enh.DescriptionEnGB AS Enh_SurgeryGroup,
enh.Surgery as Enh_Surgery,
enh.Enh_Model,

yag.SurgeryDateTime as YagDate,

/* Exchanges */
exch.SurgeryDateTime AS ExchDate,
exch.DescriptionEnGB AS Exch_SurgeryGroup,
exch.Surgery as Exch_Surgery,
exch.Exch_Model

From EMR.dbo.tSurgery S

-----------------------------------------------------------------------------------------------------------------------------------------------

Left Join Patient.dbo.tPatient P
On P.PatientID = S.PatientID

Left Join EMR.dbo.vEmployeeView Eply
On Eply.EmployeeID = S.SurgeonID

Left Join EMR.dbo.tEyeSurgeryEye ESE
On ESE.SurgeryID = S.SurgeryID

left join EMR.dbo.tEyeSurgeryEyeComplication as esec
on esec.SurgeryID = s.SurgeryID and esec.EyeID = ese.EyeID

left join EMR.dbo.tLookupComplication as lc
on lc.ComplicationID = esec.ComplicationID

LEFT JOIN EMR.dbo.tExam AS e
ON e.PatientID = s.PatientID
AND e.ExamID = 
  (
    SELECT TOP (1)
    e2.ExamID
    FROM EMR.dbo.tExam AS e2
    
    LEFT JOIN EMR.dbo.tExam_ExamType AS eet ON eet.ExamID = e2.ExamID
    left join emr.dbo.tEyeExamSurgery as ees2 on ees2.ExamID = e2.ExamID
    left join emr.dbo.tEyePrescriptionEye as epe2 on epe2.EyeID = ese.EyeID and
                                        epe2.EyePrescriptionID = ees2.ManifestPrescriptionID
    
    WHERE e2.PatientID = e.PatientID
    AND eet.ExamTypeID in (5,6,18,25,40)
    AND e2.ExamDateTime <= s.SurgeryDateTime
    and epe2.Sphere IS NOT NULL
    and epe2.IsDeleted in (0)
    and e2.IsDeleted in (0)
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

Left Join EMR.dbo.tEyeExamSurgery EES
ON EES.ExamID = e.ExamID

left join emr.dbo.tLookupMotivation as lm on lm.MotivationID = ees.MotivationID

Left Join EMR.dbo.tEyeExam EE
On EE.ExamID = e.ExamID

Left Join EMR.dbo.tEyeMeasurement EM
On EM.EyeMeasurementID = EE.EyeMeasurementID

Left Join emr.dbo.tEyeMeasurementEye EME
ON EME.EyeMeasurementID = EM.EyeMeasurementID and EME.EyeID = ESE.EyeID

LEFT JOIN EMR.dbo.tEyePrescriptionEye AS epeM 
ON epeM.EyePrescriptionID = ees.ManifestPrescriptionID
AND epeM.EyeID = ese.EyeID

LEFT JOIN EMR.dbo.tVisualAcuity VA
ON VA.VisualAcuityID = epeM.VisualAcuityID

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

LEFT Join EMR.dbo.tEyeSurgeryEyeLVC ESL
ON ESL.SurgeryID = S.SurgeryID and ESL.EyeID = ESE.EyeID

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

LEFT JOIN emr.dbo.tVisualAcuity vaBinocCBCDVA 
ON vaBinocCBCDVA.VisualAcuityID = ees.RefractionVisualAcuityID

LEFT JOIN EMR.dbo.tLookupVA AS lvaBinocCBCDVA
ON lvaBinocCBCDVA.VAID = vaBinocCBCDVA.BCDVAID


              OUTER APPLY (SELECT TOP (1)
              e3.PatientID,
              e3.ExamDateTime,
              ees3.DominantEyeID,
              lm3.DescriptionEnGB as Motivation
              
              FROM EMR.dbo.tExam as e3
              
              LEFT JOIN EMR.dbo.tExam_ExamType AS eet3
              ON eet3.ExamID = e3.ExamID
              
              Left Join EMR.dbo.tEyeExamSurgery ees3
              ON ees3.ExamID = e3.ExamID
              
              left join emr.dbo.tLookupMotivation as lm3 on lm3.MotivationID = ees3.MotivationID
              
              WHERE eet3.ExamTypeID in (18)
              and e.PatientID = e3.PatientID
              AND e.ExamDateTime <= e3.ExamDateTime) AS iol_press
              
              
/*PostOp Month 1*/
  
  OUTER APPLY
(
  SELECT TOP 1
  SPO.PostOpDateTime As Month1PostOp, 
  lvaPUCDVA.LogMAR As M1_UCDVA_LM, 
  lvaPUCIVA.LogMAR As M1_UCIVA_LM, 
  lvaPUCNVA.LogMAR As M1_UCNVA_LM, 
  lvaPDCNVA.LogMAR As M1_DCNVA_LM, 
  lvaPDCIVA.LogMAR As M1_DCIVA_LM, 
  lvaPBCDVA.LogMAR As M1_BCDVA_LM,
  lvaPBCDVA.Decimal As M1_BCDVA_D,
  lvaPBCIVA.LogMAR As M1_BCIVA_LM,
  lvaPBCNVA.LogMAR As M1_BCNVA_LM,
  lvaPBinocUCDVA.Decimal As M1_Binoc_UCDVA_D,
  lvaPBinocUCDVA.LogMAR As M1_Binoc_UCDVA_LM,
  lvaPBinocUCIVA.LogMAR As M1_Binoc_UCIVA_LM,
  lvaPBinocUCNVA.LogMAR As M1_Binoc_UCNVA_LM,
  lvaPBinocDCNVA.LogMAR As M1_Binoc_DCNVA_LM,
  lvaPBinocDCIVA.LogMAR As M1_Binoc_DCIVA_LM,
  lvaPBinocBCIVA.LogMAR as M1_Binoc_BCIVA_LM,
  lvaPBinocBCNVA.LogMAR as M1_Binoc_BCNVA_LM,
  lvaPBinocBCDVA.Decimal as M1_Binoc_BCDVA_D,
  lvaPBinocBCDVA.LogMar as M1_Binoc_BCDVA_LM,
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
  
  Left Join EMR.dbo.tLookupVA lvaPBinocBCDVA
  On lvaPBinocBCDVA.VAID = espo.BCVABinocularDistanceID
  
  WHERE SPO.PatientID = S.PatientID
  and espoe.EyeID = ESE.EyeID
  AND (DATEDIFF (DAY, S.SurgeryDateTime, SPO.PostOpDateTime) >=20 and DATEDIFF (DAY, S.SurgeryDateTime, SPO.PostOpDateTime) <=40)
  AND SPO.SurgeryID = S.SurgeryID
  ORDER BY SPO.PostOpDateTime DESC
) AS SPOM1

			   OUTER APPLY ( Select TOP (1)
s1.SurgeryDateTime
from emr.dbo.tSurgery as S1

left join emr.dbo.tEyeSurgeryEye as ese1 on ese1.SurgeryID = s1.SurgeryID
left join emr.dbo.tLookupSurgeryType as lst1 on lst1.SurgeryTypeID = ese1.SurgeryTypeID

WHERE s1.PatientID = s.PatientID
and ese1.EyeID = ese.EyeID
and s1.SurgeryDateTime < s.SurgeryDateTime
and lst1.SurgeryTypeGroupID = 1
and s1.IsDeleted = 0
and ese1.IsComplete = 1
) AS LVCSurgery

			   OUTER APPLY ( Select TOP (1)
rh.PatientID, rh.SurveyDate
  FROM peq.dbo.tResultHeader rh (NOLOCK)
JOIN peq.dbo.tResultLine rl (NOLOCK) ON rl.ResultHeaderID = rh.ResultHeaderID
JOIN peq.dbo.tQuestion_Option qo (NOLOCK) ON qo.QuestionOptionID = rl.QuestionOptionID
JOIN peq.dbo.tOption_Language ol (NOLOCK) ON ol.OptionID = qo.OptionID  and ol.LanguageID = 2057
WHERE rh.PatientID = e.PatientID
AND COALESCE(rl.IsDeleted, 0) = 0
and (DATEDIFF (DAY, S.SurgeryDateTime,rh.SurveyDate) >=20 and DATEDIFF (DAY, S.SurgeryDateTime, rh.SurveyDate) <=40)
and rh.SurveyID in (13,14,17,18,21,28,40,43,44,48,54,56)) AS Survey

OUTER APPLY
(
  SELECT TOP 1 su.SurgeryID, lstg_enh.DescriptionEnGB, lstt_enh.SurgeryTreatmentTypeID, lst_enh.DescriptionEnGB AS Surgery,
  lstg_enh.SurgeryTypeGroupID, su.SurgeryDateTime, epe_enh.Sphere, epe_enh.Cyl, lil_enh.DescriptionEnGB as Enh_Model
  FROM emr.dbo.tSurgery su
  
  LEFT JOIN EMR.dbo.tEyeSurgery AS es_enh ON es_enh.SurgeryID = su.SurgeryID
  LEFT JOIN EMR.dbo.tEyeSurgeryEye AS ese_enh ON ese_enh.SurgeryID = su.SurgeryID AND ese_enh.EyeID = ese.EyeID
  LEFT JOIN EMR.dbo.tEyePrescriptionEye AS epe_enh ON epe_enh.EyePrescriptionID = es_enh.TreatedEyePrescriptionID AND epe_enh.EyeID = ese.EyeID
  LEFT JOIN EMR.dbo.tLookupSurgeryTreatmentType AS lstt_enh ON lstt_enh.SurgeryTreatmentTypeID = ese_enh.SurgeryTreatmentTypeID
  LEFT JOIN EMR.dbo.tLookupSurgeryType AS lst_enh ON lst_enh.SurgeryTypeID = ese_enh.SurgeryTypeID
  LEFT JOIN EMR.dbo.tLookupSurgeryTypeGroup AS lstg_enh ON lstg_enh.SurgeryTypeGroupID = lst_enh.SurgeryTypeGroupID
  Left Join EMR.dbo.tEyeSurgeryEyeIOL ESEI_enh On ESEI_enh.SurgeryID = su.SurgeryID and ESEI_enh.EyeID = ese_enh.EyeID
  LEFT JOIN EMR.dbo.tIOLLensCalculation AS ilc_enh ON ilc_enh.IOLLensCalculationID = ESEI_enh.Type1IOLLensCalculationID
  LEFT JOIN EMR.dbo.tLookupIOLLens AS lil_enh ON lil_enh.IOLLensID = ilc_enh.IOLLensID
  
  WHERE su.PatientID = s.PatientID
  AND su.SurgeryDateTime >= s.SurgeryDateTime 
  AND su.IsDeleted = 0
  and lstt_enh.SurgeryTreatmentTypeID = 2
  and lstg_enh.SurgeryTypeGroupID in (1) -- LVC
  ORDER BY su.SurgeryID DESC
) AS Enh

OUTER APPLY
(
  SELECT TOP 1 su.SurgeryID, lstg_enh.DescriptionEnGB, lstt_enh.SurgeryTreatmentTypeID, lst_enh.DescriptionEnGB AS Surgery,
  lstg_enh.SurgeryTypeGroupID, su.SurgeryDateTime, epe_enh.Sphere, epe_enh.Cyl, lil_enh.DescriptionEnGB as Enh_Model
  FROM emr.dbo.tSurgery su
  
  LEFT JOIN EMR.dbo.tEyeSurgery AS es_enh ON es_enh.SurgeryID = su.SurgeryID
  LEFT JOIN EMR.dbo.tEyeSurgeryEye AS ese_enh ON ese_enh.SurgeryID = su.SurgeryID AND ese_enh.EyeID = ese.EyeID
  LEFT JOIN EMR.dbo.tEyePrescriptionEye AS epe_enh ON epe_enh.EyePrescriptionID = es_enh.TreatedEyePrescriptionID AND epe_enh.EyeID = ese.EyeID
  LEFT JOIN EMR.dbo.tLookupSurgeryTreatmentType AS lstt_enh ON lstt_enh.SurgeryTreatmentTypeID = ese_enh.SurgeryTreatmentTypeID
  LEFT JOIN EMR.dbo.tLookupSurgeryType AS lst_enh ON lst_enh.SurgeryTypeID = ese_enh.SurgeryTypeID
  LEFT JOIN EMR.dbo.tLookupSurgeryTypeGroup AS lstg_enh ON lstg_enh.SurgeryTypeGroupID = lst_enh.SurgeryTypeGroupID
  Left Join EMR.dbo.tEyeSurgeryEyeIOL ESEI_enh On ESEI_enh.SurgeryID = su.SurgeryID and ESEI_enh.EyeID = ese_enh.EyeID
  LEFT JOIN EMR.dbo.tIOLLensCalculation AS ilc_enh ON ilc_enh.IOLLensCalculationID = ESEI_enh.Type1IOLLensCalculationID
  LEFT JOIN EMR.dbo.tLookupIOLLens AS lil_enh ON lil_enh.IOLLensID = ilc_enh.IOLLensID
  
  WHERE su.PatientID = s.PatientID
  AND su.SurgeryDateTime >= s.SurgeryDateTime 
  AND su.IsDeleted = 0
  and lstg_enh.SurgeryTypeGroupID in (11) -- YAG
  ORDER BY su.SurgeryID DESC
) AS Yag

OUTER APPLY
(
SELECT TOP 1 su1.SurgeryID, lstg_exch.DescriptionEnGB, lstt_exch.SurgeryTreatmentTypeID, lst_exch.DescriptionEnGB AS Surgery,
lstg_exch.SurgeryTypeGroupID, su1.SurgeryDateTime, epe_exch.Sphere, epe_exch.Cyl, lil_exch.DescriptionEnGB as Exch_Model
FROM emr.dbo.tSurgery su1

LEFT JOIN EMR.dbo.tEyeSurgery AS es_exch ON es_exch.SurgeryID = su1.SurgeryID
LEFT JOIN EMR.dbo.tEyeSurgeryEye AS ese_exch ON ese_exch.SurgeryID = su1.SurgeryID AND ese_exch.EyeID = ese.EyeID
LEFT JOIN EMR.dbo.tEyePrescriptionEye AS epe_exch ON epe_exch.EyePrescriptionID = es_exch.TreatedEyePrescriptionID AND epe_exch.EyeID = ese.EyeID
LEFT JOIN EMR.dbo.tLookupSurgeryTreatmentType AS lstt_exch ON lstt_exch.SurgeryTreatmentTypeID = ese_exch.SurgeryTreatmentTypeID
LEFT JOIN EMR.dbo.tLookupSurgeryType AS lst_exch ON lst_exch.SurgeryTypeID = ese_exch.SurgeryTypeID
LEFT JOIN EMR.dbo.tLookupSurgeryTypeGroup AS lstg_exch ON lstg_exch.SurgeryTypeGroupID = lst_exch.SurgeryTypeGroupID
Left Join EMR.dbo.tEyeSurgeryEyeIOL ESEI_exch On ESEI_exch.SurgeryID = su1.SurgeryID and ESEI_exch.EyeID = ese_exch.EyeID
LEFT JOIN EMR.dbo.tIOLLensCalculation AS ilc_exch ON ilc_exch.IOLLensCalculationID = ESEI_exch.Type1IOLLensCalculationID
LEFT JOIN EMR.dbo.tLookupIOLLens AS lil_exch ON lil_exch.IOLLensID = ilc_exch.IOLLensID

WHERE su1.PatientID = s.PatientID
AND su1.SurgeryDateTime >= s.SurgeryDateTime 
AND su1.IsDeleted = 0
and ese_exch.SurgeryTypeID in (23,24) --exchange mono/multi
ORDER BY su1.SurgeryID asc
) AS Exch

where S.SurgeryDateTime between   ', "'", mindate , "'", ' AND ', "'", maxdate , "'",'
and lr.RegionID in (1,2,3,4,5,6)
and EET.ExamTypeID in (5,6,18,25,40)
and lstg.SurgeryTypeGroupID in (', surgerytype, ')
and ese.SurgeryTreatmentTypeID in (1) --primary
and ilc.IOLLensID in (101,103,106,105,108,109,110,111,112)
and ese.SurgeryTypeID not in (23,24) --exchanges
and S.IsDeleted = 0
and ese.IsComplete in (1)
and e.IsDeleted = 0

-----------------------------------------------------------------------------------------------------------------------------------------------
-----------------------------------------------------------------------------------------------------------------------------------------------

union


Select Distinct
/*Con*/
  
  S.PatientID,
  2 as QueryType,
P.Gender,
floor((datediff(day,0,s.SurgeryDateTime) - datediff(day,0,P.DOB)) / 365.2425) as Age,
E.ExamDateTime,
ese.EyeID,
case
  when iol_press.DominantEyeID is not null then iol_press.DominantEyeID
  when iol_press.DominantEyeID is null then ees.DominantEyeID
end as DominantEyeID,
iol_press.Motivation,
ese.Target,

/* Con Refractive */
  epeM.Sphere,
epeM.Cyl,
epeM.AddNear,

eme.Keratometry1 Keratometry_Power_1,
eme.Keratometry2 Keratometry_Power_2,

lvaCUCDVA.SnellenMetre As UCDVA_SM,
lvaCUCDVA.LogMAR As UCDVA_LM,

lvaCUCIVA.SnellenMetre As UCIVA_SM,
lvaCUCIVA.LogMAR AS UCIVA_LM,

lvaCUCNVA.SnellenMetre As UCNVA_SM,
lvaCUCNVA.LogMAR AS UCNVA_LM,

lvaCBCDVA.LogMAR AS BCDVA_LM,

lvaCBinocUCDVA.SnellenMetre As Binoc_UCDVA_SM,
lvaCBinocUCDVA.LogMAR As Binoc_UCDVA_LM,
lvaCBinocUCDVA.Decimal As Binoc_UCDVA_D,

lvaCBinocUCIVA.SnellenMetre As Binoc_UCIVA_SM,
lvaCBinocUCIVA.LogMAR As Binoc_UCIVA_LM,

lvaCBinocUCNVA.SnellenMetre As Binoc_UCNVA_SM,
lvaCBinocUCNVA.LogMAR As Binoc_UCNVA_LM,

lvaBinocCBCDVA.SnellenMetre as Binoc_BCDVA_SM,
lvaBinocCBCDVA.LogMAR as Binoc_BCDVA_LM,

LL.DescriptionEnGB AS Lens,

/*Surgery*/
  
  S.SurgeryDateTime,
s.SurgeryID,
s.SurgeonID,
s.SurgeryLocationID,
CASE
  WHEN (LVCSurgery.SurgeryDateTime IS NOT NULL) OR (esei.PriorLVC = 1) THEN 1
  ELSE 0
END AS PriorLVC,
lst.DescriptionEnGB As Surgery,
lstt.DescriptionEnGB As Surgery_Type,
lstg.DescriptionEnGB As Surgery_Group,
lil.Multi As Multifocal,
lil.Toric,
lil.DescriptionEnGB As Model,
lc.DescriptionEnGB as Complication,

/*M1 Post Op*/
  
SPOM1.Month1PostOp,

SPOM1.Sph_M1,
SPOM1.Cyl_M1,

SPOM1.M1_UCDVA_LM,
SPOM1.M1_UCIVA_LM,
SPOM1.M1_UCNVA_LM,
SPOM1.M1_BCDVA_LM,
SPOM1.M1_BCDVA_D,
SPOM1.M1_Binoc_UCDVA_LM,
SPOM1.M1_Binoc_UCDVA_D,
SPOM1.M1_Binoc_UCIVA_LM,
SPOM1.M1_Binoc_UCNVA_LM,
SPOM1.M1_Binoc_BCDVA_D,
SPOM1.M1_Binoc_BCDVA_LM,

Survey.SurveyDate as M1_SurveyDate,

/* Enhancements */
  enh.SurgeryDateTime AS EnhDate,
enh.DescriptionEnGB AS Enh_SurgeryGroup,
enh.Surgery as Enh_Surgery,
enh.Enh_Model,

yag.SurgeryDateTime as YagDate,

/* Exchanges */
exch.SurgeryDateTime AS ExchDate,
exch.DescriptionEnGB AS Exch_SurgeryGroup,
exch.Surgery as Exch_Surgery,
exch.Exch_Model

From EMR.dbo.tSurgery S

Left Join Patient.dbo.tPatient P
On P.PatientID = S.PatientID

Left Join EMR.dbo.vEmployeeView Eply
On Eply.EmployeeID = S.SurgeonID

Left Join EMR.dbo.tEyeSurgeryEye ESE
On ESE.SurgeryID = S.SurgeryID

left join EMR.dbo.tEyeSurgeryEyeComplication as esec
on esec.SurgeryID = s.SurgeryID and esec.EyeID = ese.EyeID

left join EMR.dbo.tLookupComplication as lc
on lc.ComplicationID = esec.ComplicationID
               
LEFT JOIN EMR.dbo.tExam AS e
ON e.PatientID = s.PatientID
AND e.ExamID = 
  (
    SELECT TOP (1)
    e2.ExamID
    FROM EMR.dbo.tExam AS e2
    
    LEFT JOIN EMR.dbo.tExam_ExamType AS eet ON eet.ExamID = e2.ExamID
    left join emr.dbo.tEyeExamMedicalReview as eemr2 on eemr2.ExamID = e2.ExamID
    left join emr.dbo.tEyePrescriptionEye as epe2 on epe2.EyeID = ese.EyeID and epe2.EyePrescriptionID = eemr2.ManifestRefractionPrescriptionID
    
    WHERE e2.PatientID = e.PatientID
    AND eet.ExamTypeID in (9)
    AND e2.ExamDateTime <= s.SurgeryDateTime
    and epe2.Sphere IS NOT NULL
    and epe2.IsDeleted in (0)
    and e2.IsDeleted in (0)
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

Left Join EMR.dbo.tEyeExamSurgery EES
ON EES.ExamID = e.ExamID

left join emr.dbo.tEyeExamMedicalReview as eemr on eemr.ExamID = e.ExamID

Left Join EMR.dbo.tEyeExam EE
On EE.ExamID = e.ExamID

Left Join EMR.dbo.tEyeMeasurement EM
On EM.EyeMeasurementID = EE.EyeMeasurementID

Left Join emr.dbo.tEyeMeasurementEye EME
ON EME.EyeMeasurementID = EM.EyeMeasurementID and EME.EyeID = ESE.EyeID

LEFT JOIN EMR.dbo.tEyePrescriptionEye AS epeM 
ON epeM.EyePrescriptionID = eemr.ManifestRefractionPrescriptionID
AND epeM.EyeID = ese.EyeID

LEFT JOIN EMR.dbo.tVisualAcuity VA
ON VA.VisualAcuityID = epeM.VisualAcuityID

LEFT JOIN EMR.dbo.tLookupVA lvaCBCDVA
ON lvaCBCDVA.VAID = VA.BCDVAID


LEFT Join EMR.dbo.tEyeSurgeryEyeLVC ESL
ON ESL.SurgeryID = S.SurgeryID and ESL.EyeID = ESE.EyeID

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

LEFT JOIN emr.dbo.tVisualAcuity vaBinocCBCDVA 
ON vaBinocCBCDVA.VisualAcuityID = ees.RefractionVisualAcuityID

LEFT JOIN EMR.dbo.tLookupVA AS lvaBinocCBCDVA
ON lvaBinocCBCDVA.VAID = vaBinocCBCDVA.BCDVAID

              OUTER APPLY (SELECT TOP (1)
              e3.PatientID,
              e3.ExamDateTime,
              ees3.DominantEyeID,
              lm3.DescriptionEnGB as Motivation
              
              FROM EMR.dbo.tExam as e3
              
              LEFT JOIN EMR.dbo.tExam_ExamType AS eet3
              ON eet3.ExamID = e3.ExamID
              
              Left Join EMR.dbo.tEyeExamSurgery ees3
              ON ees3.ExamID = e3.ExamID
              
              left join emr.dbo.tLookupMotivation as lm3 on lm3.MotivationID = ees3.MotivationID
              
              WHERE eet3.ExamTypeID in (18)
              and e.PatientID = e3.PatientID
              AND e.ExamDateTime <= e3.ExamDateTime) AS iol_press
              
/*PostOp Month 1*/
  
  OUTER APPLY
(
  SELECT TOP 1
  SPO.PostOpDateTime As Month1PostOp, 
  lvaPUCDVA.LogMAR As M1_UCDVA_LM, 
  lvaPUCIVA.LogMAR As M1_UCIVA_LM, 
  lvaPUCNVA.LogMAR As M1_UCNVA_LM,  
  lvaPDCNVA.LogMAR As M1_DCNVA_LM,  
  lvaPDCIVA.LogMAR As M1_DCIVA_LM, 
  lvaPBCDVA.LogMAR As M1_BCDVA_LM,
  lvaPBCDVA.Decimal As M1_BCDVA_D,
  lvaPBCIVA.LogMAR As M1_BCIVA_LM,
  lvaPBCNVA.LogMAR As M1_BCNVA_LM,
  lvaPBinocUCDVA.Decimal As M1_Binoc_UCDVA_D,
  lvaPBinocUCDVA.LogMAR As M1_Binoc_UCDVA_LM,
  lvaPBinocUCIVA.LogMAR As M1_Binoc_UCIVA_LM,
  lvaPBinocUCNVA.LogMAR As M1_Binoc_UCNVA_LM,
  lvaPBinocDCNVA.LogMAR As M1_Binoc_DCNVA_LM,
  lvaPBinocDCIVA.LogMAR As M1_Binoc_DCIVA_LM,
  lvaPBinocBCIVA.LogMAR as M1_Binoc_BCIVA_LM,
  lvaPBinocBCNVA.LogMAR as M1_Binoc_BCNVA_LM,
  lvaPBinocBCDVA.Decimal as M1_Binoc_BCDVA_D,
  lvaPBinocBCDVA.LogMar as M1_Binoc_BCDVA_LM,
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
  
  Left Join EMR.dbo.tLookupVA lvaPBinocBCDVA
  On lvaPBinocBCNVA.VAID = espo.BCVABinocularDistanceID
  
  WHERE SPO.PatientID = S.PatientID
  and espoe.EyeID = ESE.EyeID
  AND (DATEDIFF (DAY, S.SurgeryDateTime, SPO.PostOpDateTime) >=20 and DATEDIFF (DAY, S.SurgeryDateTime, SPO.PostOpDateTime) <=40)
  AND SPO.SurgeryID = S.SurgeryID
  ORDER BY SPO.PostOpDateTime DESC
) AS SPOM1

			   OUTER APPLY ( Select TOP (1)
s1.SurgeryDateTime
from emr.dbo.tSurgery as S1

left join emr.dbo.tEyeSurgeryEye as ese1 on ese1.SurgeryID = s1.SurgeryID
left join emr.dbo.tLookupSurgeryType as lst1 on lst1.SurgeryTypeID = ese1.SurgeryTypeID

WHERE s1.PatientID = s.PatientID
and ese1.EyeID = ese.EyeID
and s1.SurgeryDateTime < s.SurgeryDateTime
and lst1.SurgeryTypeGroupID = 1
) AS LVCSurgery

			   OUTER APPLY ( Select TOP (1)
rh.PatientID, rh.SurveyDate
  FROM peq.dbo.tResultHeader rh (NOLOCK)
JOIN peq.dbo.tResultLine rl (NOLOCK) ON rl.ResultHeaderID = rh.ResultHeaderID
JOIN peq.dbo.tQuestion_Option qo (NOLOCK) ON qo.QuestionOptionID = rl.QuestionOptionID
JOIN peq.dbo.tOption_Language ol (NOLOCK) ON ol.OptionID = qo.OptionID  and ol.LanguageID = 2057
WHERE rh.PatientID = e.PatientID
AND COALESCE(rl.IsDeleted, 0) = 0
and (DATEDIFF (DAY, S.SurgeryDateTime,rh.SurveyDate) >=20 and DATEDIFF (DAY, S.SurgeryDateTime, rh.SurveyDate) <=40)
and rh.SurveyID in (13,14,17,18,21,28,40,43,44,48,54,56)) AS Survey

OUTER APPLY
(
  SELECT TOP 1 su.SurgeryID, lstg_enh.DescriptionEnGB, lstt_enh.SurgeryTreatmentTypeID, lst_enh.DescriptionEnGB AS Surgery,
  lstg_enh.SurgeryTypeGroupID, su.SurgeryDateTime, epe_enh.Sphere, epe_enh.Cyl, lil_enh.DescriptionEnGB as Enh_Model
  FROM emr.dbo.tSurgery su
  
  LEFT JOIN EMR.dbo.tEyeSurgery AS es_enh ON es_enh.SurgeryID = su.SurgeryID
  LEFT JOIN EMR.dbo.tEyeSurgeryEye AS ese_enh ON ese_enh.SurgeryID = su.SurgeryID AND ese_enh.EyeID = ese.EyeID
  LEFT JOIN EMR.dbo.tEyePrescriptionEye AS epe_enh ON epe_enh.EyePrescriptionID = es_enh.TreatedEyePrescriptionID AND epe_enh.EyeID = ese.EyeID
  LEFT JOIN EMR.dbo.tLookupSurgeryTreatmentType AS lstt_enh ON lstt_enh.SurgeryTreatmentTypeID = ese_enh.SurgeryTreatmentTypeID
  LEFT JOIN EMR.dbo.tLookupSurgeryType AS lst_enh ON lst_enh.SurgeryTypeID = ese_enh.SurgeryTypeID
  LEFT JOIN EMR.dbo.tLookupSurgeryTypeGroup AS lstg_enh ON lstg_enh.SurgeryTypeGroupID = lst_enh.SurgeryTypeGroupID
  Left Join EMR.dbo.tEyeSurgeryEyeIOL ESEI_enh On ESEI_enh.SurgeryID = su.SurgeryID and ESEI_enh.EyeID = ese_enh.EyeID
  LEFT JOIN EMR.dbo.tIOLLensCalculation AS ilc_enh ON ilc_enh.IOLLensCalculationID = ESEI_enh.Type1IOLLensCalculationID
  LEFT JOIN EMR.dbo.tLookupIOLLens AS lil_enh ON lil_enh.IOLLensID = ilc_enh.IOLLensID
  
  WHERE su.PatientID = s.PatientID
  AND su.SurgeryDateTime >= s.SurgeryDateTime 
  AND su.IsDeleted = 0
  and lstt_enh.SurgeryTreatmentTypeID = 2
  and lstg_enh.SurgeryTypeGroupID in (1) -- LVC
  ORDER BY su.SurgeryID DESC
) AS Enh

OUTER APPLY
(
  SELECT TOP 1 su.SurgeryID, lstg_enh.DescriptionEnGB, lstt_enh.SurgeryTreatmentTypeID, lst_enh.DescriptionEnGB AS Surgery,
  lstg_enh.SurgeryTypeGroupID, su.SurgeryDateTime, epe_enh.Sphere, epe_enh.Cyl, lil_enh.DescriptionEnGB as Enh_Model
  FROM emr.dbo.tSurgery su
  
  LEFT JOIN EMR.dbo.tEyeSurgery AS es_enh ON es_enh.SurgeryID = su.SurgeryID
  LEFT JOIN EMR.dbo.tEyeSurgeryEye AS ese_enh ON ese_enh.SurgeryID = su.SurgeryID AND ese_enh.EyeID = ese.EyeID
  LEFT JOIN EMR.dbo.tEyePrescriptionEye AS epe_enh ON epe_enh.EyePrescriptionID = es_enh.TreatedEyePrescriptionID AND epe_enh.EyeID = ese.EyeID
  LEFT JOIN EMR.dbo.tLookupSurgeryTreatmentType AS lstt_enh ON lstt_enh.SurgeryTreatmentTypeID = ese_enh.SurgeryTreatmentTypeID
  LEFT JOIN EMR.dbo.tLookupSurgeryType AS lst_enh ON lst_enh.SurgeryTypeID = ese_enh.SurgeryTypeID
  LEFT JOIN EMR.dbo.tLookupSurgeryTypeGroup AS lstg_enh ON lstg_enh.SurgeryTypeGroupID = lst_enh.SurgeryTypeGroupID
  Left Join EMR.dbo.tEyeSurgeryEyeIOL ESEI_enh On ESEI_enh.SurgeryID = su.SurgeryID and ESEI_enh.EyeID = ese_enh.EyeID
  LEFT JOIN EMR.dbo.tIOLLensCalculation AS ilc_enh ON ilc_enh.IOLLensCalculationID = ESEI_enh.Type1IOLLensCalculationID
  LEFT JOIN EMR.dbo.tLookupIOLLens AS lil_enh ON lil_enh.IOLLensID = ilc_enh.IOLLensID
  
  WHERE su.PatientID = s.PatientID
  AND su.SurgeryDateTime >= s.SurgeryDateTime 
  AND su.IsDeleted = 0
  and lstg_enh.SurgeryTypeGroupID in (11) -- YAG
  ORDER BY su.SurgeryID DESC
) AS Yag

OUTER APPLY
(
SELECT TOP 1 su1.SurgeryID, lstg_exch.DescriptionEnGB, lstt_exch.SurgeryTreatmentTypeID, lst_exch.DescriptionEnGB AS Surgery,
lstg_exch.SurgeryTypeGroupID, su1.SurgeryDateTime, epe_exch.Sphere, epe_exch.Cyl, lil_exch.DescriptionEnGB as Exch_Model
FROM emr.dbo.tSurgery su1

LEFT JOIN EMR.dbo.tEyeSurgery AS es_exch ON es_exch.SurgeryID = su1.SurgeryID
LEFT JOIN EMR.dbo.tEyeSurgeryEye AS ese_exch ON ese_exch.SurgeryID = su1.SurgeryID AND ese_exch.EyeID = ese.EyeID
LEFT JOIN EMR.dbo.tEyePrescriptionEye AS epe_exch ON epe_exch.EyePrescriptionID = es_exch.TreatedEyePrescriptionID AND epe_exch.EyeID = ese.EyeID
LEFT JOIN EMR.dbo.tLookupSurgeryTreatmentType AS lstt_exch ON lstt_exch.SurgeryTreatmentTypeID = ese_exch.SurgeryTreatmentTypeID
LEFT JOIN EMR.dbo.tLookupSurgeryType AS lst_exch ON lst_exch.SurgeryTypeID = ese_exch.SurgeryTypeID
LEFT JOIN EMR.dbo.tLookupSurgeryTypeGroup AS lstg_exch ON lstg_exch.SurgeryTypeGroupID = lst_exch.SurgeryTypeGroupID
Left Join EMR.dbo.tEyeSurgeryEyeIOL ESEI_exch On ESEI_exch.SurgeryID = su1.SurgeryID and ESEI_exch.EyeID = ese_exch.EyeID
LEFT JOIN EMR.dbo.tIOLLensCalculation AS ilc_exch ON ilc_exch.IOLLensCalculationID = ESEI_exch.Type1IOLLensCalculationID
LEFT JOIN EMR.dbo.tLookupIOLLens AS lil_exch ON lil_exch.IOLLensID = ilc_exch.IOLLensID

WHERE su1.PatientID = s.PatientID
AND su1.SurgeryDateTime >= s.SurgeryDateTime 
AND su1.IsDeleted = 0
and ese_exch.SurgeryTypeID in (23,24) --exchange mono/multi
ORDER BY su1.SurgeryID asc
) AS Exch

where S.SurgeryDateTime between  ', "'", mindate , "'", ' AND ', "'", maxdate , "'",'
and lr.RegionID in (1,2,3,4,5,6)
and EET.ExamTypeID in (9)
and lstg.SurgeryTypeGroupID in (', surgerytype, ')
and ese.SurgeryTreatmentTypeID in (1) --primary
and ilc.IOLLensID in (101,103,106,105,108,109,110,111,112)
and ese.SurgeryTypeID not in (23,24) --exchanges
and S.IsDeleted = 0
and ese.IsComplete in (1)
and e.IsDeleted = 0

order by S.PatientID desc', sep="")