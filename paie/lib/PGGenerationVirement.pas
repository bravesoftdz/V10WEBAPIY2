{***********UNITE*************************************************
Auteur  ...... : PAIE PGI
Créé le ...... : 31/07/2001
Modifié le ... :   /  /    
Description .. : Chargement des tobs pour la génération des virements
Mots clefs ... : PAIE;VIREMENT
*****************************************************************
PT1    : 31/07/2001 SB V547 Modification champ suffixe MODEREGLE
PT2    : 14/03/2002 SB V571 Fiche de com n°011205M Ajout Champ PVI_RIBSALAIREEMIS
PT3    : 08/02/2005 SB V_60 FQ 11826 Ajout champ PVI_ECHEANCE
PT4    : 14/03/2007 VG V_72 BQ_GENERAL n'est pas forcément unique
PT5    : 03/07/2007 MF V_72 FQ 14094 Génération virement pour un salarié ayant changé d'établissment
}
unit PGGenerationVirement;

interface
uses
{$IFNDEF EAGLCLIENT}
{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}      
      HCtrls,
      HEnt1,
      HMsgBox,
      UTOB,
      PgOutilsTreso;

Function RendTobGenVirement(StWhere : String) : Tob;

implementation

Function RendTobGenVirement(StWhere : String) : Tob;
Var
SQL : TQuery;
Tob_CodBq : Tob;
StPlus : String;
BEGIN
try
   BeginTrans;
   ExecuteSQL ('DELETE FROM VIREMENTS');
   Tob_CodBq:= Tob.Create ('Les virements',nil,-1);

   StPlus:= PGBanqueCP (True);                   //PT4
//Recup les type de virement personnalisé
   Sql:= OpenSql ('SELECT PPU_ETABLISSEMENT AS PVI_ETABLISSEMENT,'+
                  ' PPU_SALARIE AS PVI_SALARIE, PSA_LIBELLE AS PVI_LIBELLE,'+
                  ' PSA_PRENOM AS PVI_PRENOM,'+
                  ' PSA_AUXILIAIRE AS PVI_AUXILIAIRE,'+
                  ' PPU_DATEDEBUT AS PVI_DATEDEBUT,'+
                  ' PPU_DATEFIN AS PVI_DATEFIN, PPU_PAYELE AS PVI_PAYELE,'+
                  ' PPU_CNETAPAYER AS PVI_MONTANT,'+
                  ' PPU_BANQUEEMIS AS PVI_BANQUEEMIS,'+
                  ' BQ_GENERAL AS PVI_RIBSALAIRE, BQ_LIBELLE AS PVI_BANQUE,'+
                  ' PPU_TOPREGLE AS PVI_TOPREGLE,'+
                  ' PSA_AUXILIAIRE AS PVI_AUXILIAIRE,'+
                  ' PPU_RIBSALAIRE AS PVI_RIBSALAIREEMIS,'+
                  ' PPU_ECHEANCE AS PVI_ECHEANCE'+
                  ' FROM BANQUECP'+
                  ' LEFT JOIN SALARIES ON'+
                  ' (PSA_RIBVIRSOC=BQ_GENERAL AND PSA_TYPVIRSOC="PER")'+
                  ' LEFT JOIN PAIEENCOURS ON'+
                  ' PSA_SALARIE=PPU_SALARIE'+
                  ' LEFT JOIN ETABCOMPL ON'+
                  ' (PPU_ETABLISSEMENT=ETB_ETABLISSEMENT) WHERE'+
                  ' PPU_PGMODEREGLE="008" AND'+
                  ' PSA_AUXILIAIRE<>"" AND'+
                  ' PPU_CNETAPAYER>0 '+StPlus+' '+StWhere+       //PT4
                  ' ORDER BY PPU_SALARIE, PPU_DATEDEBUT, PPU_DATEFIN, BQ_GENERAL ',true);

   Tob_CodBq.LoadDetailDB('VIREMENTS','','', SQL, True, False);
   Ferme(SQL);

//Recup les type de virement idem etablissement
   Sql:= OpenSql ('SELECT PPU_ETABLISSEMENT AS PVI_ETABLISSEMENT,'+
                  ' PPU_SALARIE AS PVI_SALARIE, PSA_LIBELLE AS PVI_LIBELLE,'+
                  ' PSA_PRENOM AS PVI_PRENOM,'+
                  ' PSA_AUXILIAIRE AS PVI_AUXILIAIRE,'+
                  ' PPU_DATEDEBUT AS PVI_DATEDEBUT,'+
                  ' PPU_DATEFIN AS PVI_DATEFIN, PPU_PAYELE AS PVI_PAYELE,'+
                  ' PPU_CNETAPAYER AS PVI_MONTANT,'+
                  ' PPU_BANQUEEMIS AS PVI_BANQUEEMIS,'+
                  ' BQ_GENERAL AS PVI_RIBSALAIRE, BQ_LIBELLE AS PVI_BANQUE,'+
                  ' PPU_TOPREGLE AS PVI_TOPREGLE,'+
                  ' PSA_AUXILIAIRE AS PVI_AUXILIAIRE,'+
                  ' PPU_RIBSALAIRE AS PVI_RIBSALAIREEMIS,'+
                  ' PPU_ECHEANCE AS PVI_ECHEANCE'+
                  ' FROM BANQUECP'+
                  ' LEFT JOIN ETABCOMPL ON'+
                  ' ETB_RIBSALAIRE=BQ_GENERAL'+
                  ' LEFT JOIN SALARIES ON'+
                  ' ETB_ETABLISSEMENT=PSA_ETABLISSEMENT'+
                  ' LEFT JOIN PAIEENCOURS ON'+
                  ' PSA_SALARIE=PPU_SALARIE'+
// d PT5
{                 ' PSA_SALARIE=PPU_SALARIE AND'+
                  ' PSA_ETABLISSEMENT=PPU_ETABLISSEMENT WHERE'+}
                  ' WHERE'+
// f PT5
                  ' PPU_PGMODEREGLE="008" AND PSA_TYPVIRSOC="ETB" AND'+
                  ' PSA_AUXILIAIRE<>"" AND PPU_CNETAPAYER>0 '+StPlus+' '+    //PT4
                  StWhere+
                  ' ORDER BY PPU_SALARIE,PPU_DATEDEBUT,PPU_DATEFIN,BQ_GENERAL',true);
   Tob_CodBq.LoadDetailDB ('VIREMENTS','','', SQL, True, False);
   Ferme (SQL);

//Recup les type de virement idem profil : établissement et profil précisés
   Sql:= OpenSql ('SELECT PPU_ETABLISSEMENT AS PVI_ETABLISSEMENT,'+
                  ' PPU_SALARIE AS PVI_SALARIE, PSA_LIBELLE AS PVI_LIBELLE,'+
                  ' PSA_PRENOM AS PVI_PRENOM,'+
                  ' PSA_AUXILIAIRE AS PVI_AUXILIAIRE,'+
                  ' PPU_DATEDEBUT AS PVI_DATEDEBUT,'+
                  ' PPU_DATEFIN AS PVI_DATEFIN, PPU_PAYELE AS PVI_PAYELE,'+
                  ' PPU_CNETAPAYER AS PVI_MONTANT,'+
                  ' PPU_BANQUEEMIS AS PVI_BANQUEEMIS,'+
                  ' BQ_GENERAL AS PVI_RIBSALAIRE,'+
                  ' BQ_LIBELLE AS PVI_BANQUE, PPU_TOPREGLE AS PVI_TOPREGLE,'+
                  ' PSA_AUXILIAIRE AS PVI_AUXILIAIRE,'+
                  ' PPU_RIBSALAIRE AS PVI_RIBSALAIREEMIS,'+
                  ' PPU_ECHEANCE AS PVI_ECHEANCE'+
                  ' FROM BANQUECP'+
                  ' LEFT JOIN DONNEURORDRE ON'+
                  ' PDO_RIBASSOCIE=BQ_GENERAL'+
                  ' LEFT JOIN SALARIES ON'+
                  ' PDO_ETABLISSEMENT=PSA_ETABLISSEMENT AND'+
                  ' PDO_PROFIL=PSA_PROFIL AND PDO_PGMODEREGLE=PSA_PGMODEREGLE'+
                  ' LEFT JOIN PAIEENCOURS ON'+
                  ' PSA_SALARIE=PPU_SALARIE'+
// d PT5
{                 ' PSA_SALARIE=PPU_SALARIE AND'+
                  ' PSA_ETABLISSEMENT=PPU_ETABLISSEMENT WHERE'+}
                  ' WHERE'+
// f PT5
                  ' PPU_PGMODEREGLE="008" AND PSA_TYPVIRSOC="PRO" AND'+
                  ' PSA_AUXILIAIRE<>"" AND PPU_CNETAPAYER>0 '+StPlus+' '+      //PT4
                  StWhere+
                  ' ORDER BY PPU_SALARIE,PPU_DATEDEBUT,PPU_DATEFIN,BQ_GENERAL ',true);
   Tob_CodBq.LoadDetailDB ('VIREMENTS','','', SQL, True, False);
   Ferme(SQL);

   Tob_CodBq.Detail.Sort ('PVI_ETABLISSEMENT;PVI_SALARIE;PVI_DATEDEBUT;PVI_DATEFIN');
   Tob_CodBq.SetAllModifie (TRUE);
   Tob_CodBq.InsertDB (nil,False);
   result:= Tob_CodBq;
   CommitTrans;
Except
   Rollback;
   PGIBox ('Une erreur est survenue lors du chargement des données',
           'Génération des virements');
   result:=nil;
   End;
END;

end.
