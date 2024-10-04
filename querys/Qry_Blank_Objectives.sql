SELECT Tbl_Requirements.Req_Sorting, Tbl_Requirements.Requirement_Number, Tbl_Objectives.Objective_Number, Len([Validation_Text]) AS Expr1
FROM Tbl_Objectives INNER JOIN (Tbl_Requirements INNER JOIN LnkTbl_RequirementsToObjectives ON Tbl_Requirements.Requirement_Number = LnkTbl_RequirementsToObjectives.Requirement_Number) ON Tbl_Objectives.Objective_Number = LnkTbl_RequirementsToObjectives.Objective_Number
WHERE (((Len([Validation_Text])) Is Null Or (Len([Validation_Text]))<2))
ORDER BY Tbl_Requirements.Req_Sorting;
