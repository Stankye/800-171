SELECT Tbl_Requirements.Requirement_Number, Tbl_Requirements.Requirement_Score, Tbl_Requirements.Requirement_Special_Considerations_Score, Tbl_Requirements.Requirement_Other_Than_Satisfied, Tbl_Requirements.Requirement_Special_Considerations_Satisfied, IIf([Requirement_Other_Than_Satisfied]=True,[Requirement_Score],([Requirement_Special_Considerations_Score])) AS TotalScore
FROM Tbl_Requirements
WHERE (((Tbl_Requirements.Requirement_Other_Than_Satisfied)=True)) OR (((Tbl_Requirements.Requirement_Special_Considerations_Satisfied)=True));
