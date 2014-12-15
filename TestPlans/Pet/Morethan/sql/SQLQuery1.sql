SELECT     dbo.tblTranslation.TranslationNo,  dbo.tblTranslation.BGLValue, 
                      dbo.tblTranslation.ProviderValue
FROM         dbo.tblTranslation INNER JOIN
                      dbo.tblAffinity ON dbo.tblTranslation.AffinityId_FK = dbo.tblAffinity.AffinityID_PK INNER JOIN
                      dbo.tblProvider ON dbo.tblAffinity.AffinityID_PK = dbo.tblProvider.AffinityID_FK
WHERE     (dbo.tblProvider.ProviderID_PK = 442)


where affinityid_fk = '442'

provider id
442


tablename = 'Breed' and



where Affinityid_FK = 'MPE1' 



select * from tblAffinity where AffinityName Like '%morethan%'


affinity id

1050-1055