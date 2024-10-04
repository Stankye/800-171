SELECT Tbl_Requirements.Requirement_Number, Tbl_Objectives.Objective_Number, Tbl_Objectives.Objective_Text, Tbl_Objectives.Objective_Satisfied, Tbl_Objectives.Objective_Other_Than_Satisfied, Tbl_Objectives.Documents_Examined_Details, Tbl_Objectives.SME_Interviewed_Names, Tbl_Objectives.Validation_Text, Tbl_Objectives.Notes, Tbl_Objectives.Standard
FROM Tbl_Requirements INNER JOIN (Tbl_Objectives INNER JOIN LnkTbl_RequirementsToObjectives ON Tbl_Objectives.Objective_Number = LnkTbl_RequirementsToObjectives.Objective_Number) ON Tbl_Requirements.Requirement_Number = LnkTbl_RequirementsToObjectives.Requirement_Number
WHERE (((Tbl_Requirements.Requirement_Number)=[Forms]![Frm_Families_and_Objectives]![Requirement_Number]))
ORDER BY Tbl_Objectives.Objective_Number;
