bestellpos = """select bp.[FIXPOSNR]
      ,bp.[BELEGART]
      ,bp.[BELEGNR]
	  ,bp.[WARENEINGANGSNR] as Reference

	   FROM [emea_enventa_live].[dbo].[BESTELLPOS] as bp
	   left join [emea_enventa_live].[dbo].AUFTRAGSKOPF as ak
      on bp.WARENEINGANGSNR = ak.PE4_I_Wareneingangsnr


where (bp.STATUS between 3 and 5) -- Ware bereits verschickt
  
  and (bp.BranchKey = 110) -- kein ICA
  and (bp.[BELEGART] in (2, 191)) --nur PM Hersteller

  and (ak.STATUS < 4 or ak.STATUS is null) --noch nicht angekommen
  and (bp.PE14_CommentEMEA <> 'eRledigt' or bp.PE14_CommentEMEA is null) --noch nicht angekommen
  and bp.WARENEINGANGSNR not in  ( 2403100,2403093,2403092,2403105,2402004,2403123,2403138,2403139) --keine Dachser Lieferungen

  order by bp.WARENEINGANGSNR  """