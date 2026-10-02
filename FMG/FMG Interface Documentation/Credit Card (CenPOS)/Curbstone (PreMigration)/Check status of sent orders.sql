SELECT [CONO],
       [JBAOrderNo],
       [M3OrderNo],
       CO.[CIDESN],
       OS.UCIVNO AS CIINVN,
       [JBACUSN],
       [JBADSEQ],
       [CURN],
       CO.[CIAMFU],
       --'' AS CISTS,
       JCC.CISTS,
       INV.UHIVAM AS CIIAMT
FROM [PRDMODSF3].[dbo].[CurbstoneOrders] CO
    LEFT JOIN
    (
        SELECT UCIVNO,
               UCDLIX,
               UCORNO
        --,SUM(UCSAAM) UCSAAM
        FROM [NGEUMSMI0002.D44D2F32AC81.DATABASE.WINDOWS.NET].[di_staging_prd_us].[dbo].[OSBSTD]
        GROUP BY UCIVNO,
                 UCDLIX,
                 UCORNO
    ) AS OS
        ON CO.M3OrderNo = OS.UCORNO
    LEFT JOIN [NGEUMSMI0002.D44D2F32AC81.DATABASE.WINDOWS.NET].[di_staging_prd_us].[dbo].[OINVOH] AS INV
        ON OS.UCIVNO = INV.UHIVNO
    LEFT JOIN [RAYMOND].[S781CCD0].PRDMODF35.CCINV JCC
        ON RIGHT(OS.UCIVNO,7) = JCC.CIINVN
WHERE OS.UCIVNO IS NOT NULL
      AND OS.UCIVNO <> 0;

