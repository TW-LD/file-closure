SELECT 'Client / Matter' = Entities.Shortcode + '/' + CONVERT(varchar, M.Number),
		'Last Bill Date' = M.LastBillPostingDate, 
		'Last Time Posting' = M.LastTimePostingDate, 
		'Last Document (Case Manager) Date' = (SELECT TOP 1 CM.StepCreated FROM View_CaseManagerMP CM WHERE CM.EntityRef = M.EntityRef AND CM.MatterRef = M.Number ORDER BY CM.StepCreated DESC),
	FROM Matters M 
		JOIN Users FeeEarner ON M.FeeEarnerRef = FeeEarner.Code 
		JOIN Users PartnerN ON M.PartnerRef = PartnerN.Code
		JOIN Entities ON M.EntityRef = Entities.Code
