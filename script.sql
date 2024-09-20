
--DETAIL
SELECT F_DOCLIGNE.DO_Piece, AR_Ref,DL_Design,DL_QTE,DL_MontantTTC/DL_qte, DL_MontantTTC,R_Intitule FROM F_DOCLIGNE
inner join F_DOCREGL on F_DOCLIGNE.DO_Piece=F_DOCREGL.DO_Piece and F_DOCLIGNE.DO_Type=F_DOCREGL.DO_Type
inner join P_REGLEMENT on N_Reglement=cbIndice
WHERE F_DOCLIGNE.DO_TYPE IN (6,7,30) and AR_Ref is not null and DL_Qte<>0 and DO_Date='04/24/2024'



--DATE
select R_Intitule, sum(rg_montant) from F_CREGLEMENT
inner join P_REGLEMENT on N_Reglement=cbIndice
where  RG_Date='04/24/2024'
 group by R_Intitule

--MOIS
  select R_Intitule, sum(rg_montant) from F_CREGLEMENT
inner join P_REGLEMENT on N_Reglement=cbIndice
where
 year(RG_Date)=2024 and month(rg_date)=4
group by R_Intitule

--ANNEE
   select R_Intitule, sum(rg_montant) from F_CREGLEMENT
inner join P_REGLEMENT on N_Reglement=cbIndice
where
 year(RG_Date)=2024
 group by R_Intitule
