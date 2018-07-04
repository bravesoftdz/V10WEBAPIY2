{***********UNITE*************************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 13/10/2004
Modifié le ... : 21/10/2004
Description .. : procédures et fonctions de calcul et d'intégration du maitien
Suite ........ : dans le bulletin  en fct des absence et des règlements
Suite ........ : d'IJSS.
Mots clefs ... : PAIE; IJSS; MAINTIEN
*****************************************************************}
{
PT1   30/11/2004 MF V_60 Correction traitement maintien qd champ catégorie renseigné.
PT2   20/12/2004 PH V_60 Erreur SQL DB2
PT3   29/12/2004 MF V_60 Erreur SQL DB2
PT4   04/01/2005 MF V_60 1- Traitement des absences consécutives:L'ancienneté
                         est calculée avec la date de début de la première
                         absence
                         2- Rectification de la recherche de l'élément
                         CRITMAINTIEN à utiliser
                         3- Traitement des motif d'absence avec paramétrage
                         identique
                         4- Cumul du maintien par Taux
                         5- Modification du traitement de création des lignes
                         de maintien
PT5   12/01/2005 MF V_60 Rectification de la recherche de l'élément
                         CRITMAINTIEN à utiliser ds le cas uniqt règlement IJSS
PT6   26/01/2005 MF V_60 Le % de maintien à appliquer est fonction du cumul des
                           maintien des 12 mois précédent la date de début d'absence.
                           (la date de début d'absence réelle n'est utile que pour
                           déterminer l'ancienneté)
PT7   28/01/2005 MF V_60 On n'intègre pas de règlement d'IJSS ne correspondant à aucun maintien
PT8   08/02/2005 MF V_60 Traitement des commentaires des règlements d'IJSS
PT9   16/02/2005 MF V_60 Correction traitement qd absence sur le mois complet
                         du bulletin et rglt d'ijss intégré sur le même bulletin
                         (pb clé duplliquée)
                         + test limite règle
PT10  21/02/2005 MF V_60 Correction PT9 (limite de règle)
PT11  24/02/2005 Mf V_60 : Contrôle MEMECHECK
PT12  01/03/2005 MF V_60 : Correction calcul carence qd dates consécutives.
PT13  04/03/2005 MF V_602 FQ12075 SQL ORACLE
PT14  08/03/2005 MF V_602 : Correction calcul maintien quand historique +
                            absence consécutive avec carence
PT15  25/04/2005 MF V_602 : FQ 12182 : absences consécutives sur + d'un an.
PT16  18/05/2005 PH V_602 : FQ 118001 gestion de l'arrondi dans le calcul de la paie à l'envers
PT17  24/05/2005 MF V_602 :   Correction PT4-3 "erreur sql (ORPMA_TYPEABS)"
PT18  01/06/2005 MF V_602 : FQ 12342 correction nb jours carence > nb jours
                            calendaires d'absence.
                            Le nb jours de maintien de la
                            ligne de maintien générée = 0 (pas de commantaire sur
                            le bulletin, le maintien à 0 jour est mémorisé dans
                            table MAINTIEN pour conserver l'historique (carence))
PT19  17/06/2005 MF v_602 : FQ 12388 : utilisation du paramètre société SO_PGHISTOMAINTIEN
                            pour connaître la méthode de récupératio de l'historique de
                            maintien à appliquer.
PT20  23/01/2006 SB V_65 FQ 10866 Ajout clause predefini motif d'absence
PT21  20/02/2006 MF V_65   FQ 12115 : En création de la ligne de maintien initialisation
                            du champ PMT_HISTOMAINT à False
                            + Traitement du champ PMT_TYPEMAINTIEN au lieu du
                            champ PMT_CARENCE pour différencier les lignes de
                            maintien des lignes de garantie.
PT22  20/02/2006 MF  V_65   correction memcheck
PT23  23/02/2006 MF  V_65   Traitement d'une Rubrique de garantie utilisée pour
                            assurer le montant du Salaire (net) dans le cas
                            de l'intégration d'IJSS
PT24  28/02/2006 MF  V_65   1-Alimentation rubrique de garantie quand IJSS
                            2-FQ 12827 Exclusion de rub de rem pour calcul maintiens
PT25  16/03/2006 MF  V_65   FQ 12968 : PCM_VALCATEG remplacé par PCM_VALCATEGORIE
PT26  27/03/2006 PH  V_65  Prise en compte des champs motifs absences dédoublés (Heures,Jours)
PT27  11/04/2006 SB  V_65  Modification de la procedure de recherche du paramétrage motifabsence
PT28  25/04/2006 MF  V_65  FQ 13077 : rubrique de garnatie non obligatoire + message non bloquiant
                           + correction du tritement d'intégration des IJ
PT29  02/06/2006 SB  V_65  Optimisation mémoire
PT30  12/07/2006 MF  V_70  FQ 13382 : correction requête qui plante sous DB2
PT31  01/08/2006 MF  V_70  Correction
                           - quand plusieurs lignes de règlement à intégrer et aucun
                             maintien correspondant ==> "Indice de liste hors limites"
                           - Test de la valeur StWhere <> '' (Critère de suppresion du maintien)
                             afin d'éviter un vidage total de la table MAINTIEN.
PT32  13/09/2006 MF  V_70  Correction
                           Cas : Etablissement pratiquant la subrogation,
                           1 règlement d'IJSS concenant une absence déjà intégrée
                           sur un mois précédent --> la ligne d'IJSS nette ne faisait
                           pas partie du bulletin.
PT33  25/09/2006 MF  V_70  Correction du PT28
                           Le calcul de la garantie en cas d'intégration des IJSS sans subrogation
                           était erroné. Il ne tenait pas compte des IJSS nettes versées.
                           Il faut effectuer le calcul de la paie à l'envers en tenant compte
                           des IJSS nettes versées, qu'il y ait subrogation ou non.
                           En cas de non subrogation la rubrique d'IJSS nettes est ensuite
                           supprimée du bulletin.
PT34  15/02/2006  MF  V_702  Quand le paramètre société "Champ critère maintien" n'est pas renseigné
                             pas de contrôle sur ce critère dans la tob des ABS
PT35  05/06/2008 MF V_82  FQ 15459 et FQ 15466
PT37  07/08/2008 MF FQ 013;15514 : correction de la rcherche de la bonne règle en
                    fonction de la valeur du champ PCM_VALCATEGORIE
PT38  09/09/2008 MF FQ 013;15733 - le CODETAPE des absences restait à '...'
PT39  23/09/2008 MF Correction PT38 (Access Vio) + correction PT35 (Pls lignes de commentaire)
PT40  02/10/2008 MF FQ 15781 : pas de calcul du maintien qd ancienneté <= 01/01/1900
}
unit PGIJSSMaintien;

interface
uses
{$IFNDEF EAGLCLIENT}
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ELSE}
{$ENDIF}
  EntPaie,
  HCtrls,
  Hent1,
  HMsgBox,
  P5Util,
  PgCongespayes,
  PGIJSS,
  PGMaintien,
  sysutils,
  UTOB;


  procedure SalEcritIJSS(var tob_IJSS: tob);
  procedure SalEcritMaintien(var Tob_Maintien: tob);
  function RecupereAbsencesIJSS(Datef: tdatetime; Salarie, Action: string): boolean;
  function RecupereRegltIJSS(Datef: tdatetime; Salarie, Action,ChampCateg: string; Tob_maintien,Tob_etab : TOB ): TOB;
  procedure MaintienDuSalarie (const TobSal,TOB_Rub,Tob_etablissement :TOB;var TOB_abs : TOB; Datef, Dated, Mode: string; var Anomalie: integer; const ChampCateg : string; var TOB_IJSS, TOB_Maintien :TOB);
  procedure MajMaintien (var Tob_Maintien : TOB);
  procedure RecupIJSS(const Dated, Datef:TDateTime;const Mode : string; Tob_Rub : TOB;var Tob_IJSS : TOB; const Tob_Etablissement,Tob_Salarie: TOB;var ChampCateg,RubIJNettes: String; Const Subrogation : boolean);
  procedure NvelleLMaintien(const Tob_ReglesMaintien, Tob_etab : TOB;const Wcarence, TotAbsence : double; const MtGarantie : double;const MaintienNet : boolean;const Tabsence,treglt : TOB;var tmaintien : TOB;var ACheval : boolean;var WMtMaintien : double;var NbRestant : double;var DateDebutAbs : TDateTime;const uniqtreglt : integer;dated,datef : TDateTime);
var
  Tob_IJSS,Tob_Maintien                             : tob;
  TbNbj, TbNbjMax, TbCarence, TbWCarence                : array of Integer;
  TbPct                                                 : array of double;
  MaintConsec                                           : boolean;
  MaintienAZero                                         : boolean;

implementation
uses
  UTofPG_PaieEnvers,
  P5Def,
  P5RecupInfos;

var
   RegleNo      : integer;
{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 21/10/2004
Modifié le ... : 21/10/2004
Description .. :  procédure SalEcritIJSS
Suite ........ : mise à jour de la table REGLTIJSS après validation du
Suite ........ : bulletin
Mots clefs ... : PAIE; IJSS
*****************************************************************}
procedure SalEcritIJSS (var tob_IJSS: tob);
begin
  if (assigned(tob_IJSS)) and (tob_IJSS.detail.count > 0) then  
  begin
    try
      Begintrans;
      tob_IJSS.SetAllModifie(TRUE);
      tob_IJSS.InsertOrUpdateDB(false);
      CommitTrans;
    except
      Rollback;
      PGIBox ('Une erreur en mise à jour REGLTIJSS','');
    end;
  end;
  FreeAndNil(Tob_IJSS);         
  exit;
end;
{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 21/10/2004
Modifié le ... :   /  /
Description .. : procédure SalEcritMaintien
Suite ........ : Mise à jour de la table MAINTIEN après intégration dans le
Suite ........ : bulletin.
Mots clefs ... : PAIE; MAINTIEN
*****************************************************************}
procedure SalEcritMaintien(var tob_Maintien: tob);
begin

// PT31  if (assigned(tob_Maintien)) and (tob_Maintien.detail.count > 0) then 
  if (assigned(tob_Maintien)) and (tob_Maintien.detail.count > 0) and (StWhere <> '') then
  begin
    try
//PT30      Executesql('DELETE MAINTIEN WHERE ' + StWhere);
      Executesql('DELETE FROM MAINTIEN WHERE ' + StWhere);
      Begintrans;
      tob_Maintien.SetAllModifie(TRUE);
      tob_Maintien.InsertOrUpdateDB(false);
      CommitTrans;
    except
      Rollback;
      PGIBox('Une erreur en mise à jour MAINTIEN', '');
    end;
  end;
  FreeAndNil(Tob_Maintien); 
  exit;
end;
{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 21/10/2004
Modifié le ... : 21/10/2004
Description .. : fonction RecupereAbsenceIJSS
Suite ........ : vérifie l'existence d'absence de type IJSS
Mots clefs ... : PAIE; IJSS; MAINTIEN
*****************************************************************}
function RecupereAbsencesIJSS(Datef: tdatetime; Salarie, Action: string): boolean;
var
  st: string;
  Q: TQuery;
begin
  result := False;

  if Action = 'M' then
    st := 'SELECT PCN_SALARIE FROM ABSENCESALARIE WHERE' +
      ' PCN_SALARIE = "' + Salarie + '"' +
      ' AND (PCN_TYPEMVT = "ABS" AND PCN_SENSABS="-" ' +
      ' AND  PCN_TYPECONGE <>  "PRI" AND PCN_VALIDRESP="VAL")' +
      ' AND ((PCN_DATEVALIDITE <= "' + usdatetime(Datef) + '"' +
      ' AND PCN_DATEPAIEMENT <= "' + usdatetime(10) + '")' +
      ' OR (PCN_DATEPAIEMENT="' + usdatetime(Datef) + '"' +
      ' AND PCN_CODETAPE="P"))' +
      ' AND PCN_GESTIONIJSS="X"' +
      ' AND PCN_TYPECONGE IN ' +
      '(SELECT PMA_MOTIFABSENCE FROM MOTIFABSENCE' +
      ' WHERE ##PMA_PREDEFINI## (PMA_RUBRIQUE <> ""' +
      ' AND PMA_RUBRIQUE is not null) OR (PMA_RUBRIQUEJ <> "" AND PMA_RUBRIQUEJ is not null)) ' +
      ' ORDER BY PCN_TYPECONGE,PCN_DATEVALIDITE'
  else
    st := 'SELECT PCN_SALARIE FROM ABSENCESALARIE WHERE' +
      ' PCN_SALARIE = "' + Salarie + '"' +
      ' AND (PCN_TYPEMVT = "ABS" AND PCN_SENSABS="-"' +
      ' AND PCN_TYPECONGE <>  "PRI" AND PCN_VALIDRESP="VAL")' +
      ' AND PCN_DATEVALIDITE <= "' + usdatetime(Datef) + '"' +
      ' AND PCN_DATEPAIEMENT <= "' + usdatetime(10) + '"' +
      ' AND PCN_GESTIONIJSS="X"' +
      ' AND PCN_TYPECONGE IN ' +
      '(SELECT PMA_MOTIFABSENCE FROM MOTIFABSENCE' +
      ' WHERE ##PMA_PREDEFINI## (PMA_RUBRIQUE <> ""' +
      ' AND PMA_RUBRIQUE is not null) OR (PMA_RUBRIQUEJ <> "" AND PMA_RUBRIQUEJ is not null)) ' +
      ' ORDER BY PCN_TYPECONGE,PCN_DATEVALIDITE';
  Q := OpenSql(st, TRUE);
  if not Q.eof then
  begin
    result := true;
  end;
  Ferme(Q);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 21/10/2004
Modifié le ... :   /  /
Description .. : fonction RecupereRegltIJSS
Suite ........ : vérifie l'existence de règlements d'IJSS
Mots clefs ... : PAIE; IJSS; MAINTIEN
*****************************************************************}
function RecupereRegltIJSS(Datef: tdatetime; Salarie, Action,ChampCateg: string; Tob_maintien, Tob_etab : TOB ): TOB;

var
  st                                         : string;
  Tob_ReglIJSS                               : Tob;
  Q                                          : TQuery;
  tMaintien, Treglt                          : TOB;
  Wi, i , j                                  : integer;
  Totmaintien                                : integer;
begin
  Tob_ReglIJSS := Tob.create('Les Règlements IJSS', nil, -1);
  if Action = 'M' then
  begin
    st := 'SELECT REGLTIJSS.* '; 

    if (ChampCateg <> '') then
    begin
      st:=st+',SALARIES.'+ChampCateg+' ';
    end;

    st := st+ 'FROM REGLTIJSS ';
    if (ChampCateg <> '') then
    begin
      st := st +'LEFT JOIN SALARIES ON PRI_SALARIE=PSA_SALARIE ';
    end;
    st := st+ 'WHERE' +
          ' PRI_SALARIE = "' + Salarie + '"' +
          ' AND PRI_DATEREGLT <= "' + usdatetime(Datef) + '"' +
          ' AND PRI_DATEINTEGR <= "' + usdatetime(10) + '"';
  end
  else
  begin
    st := 'SELECT REGLTIJSS.* ';
    if (ChampCateg <> '') then
    begin
      st:=st+',SALARIES.'+ChampCateg+' ';
    end;

    st := st+ 'FROM REGLTIJSS ';
    if (ChampCateg <> '') then
    begin
      st := st +'LEFT JOIN SALARIES ON PRI_SALARIE=PSA_SALARIE ';
    end;
    st := st+ 'WHERE' +
          ' PRI_SALARIE = "' + Salarie + '"' +
          ' AND PRI_DATEREGLT <= "' + usdatetime(Datef) + '"' +
          ' AND PRI_DATEINTEGR <= "' + usdatetime(10) + '"';
  end;
  Q := OpenSql(st, TRUE);
  if not Q.eof then
  begin
    Tob_ReglIJSS.LoadDetailDB('REGLTIJSS', '', '', Q, False);
  end;
     if (VH_Paie.Pgmaintien) and (VH_Paie.PgGestIJSS) and
        (Tob_ReglIJSS <> nil) and (Tob_ReglIJSS.Detail.count > 0) then
     begin
       // XP 10.10.2007 Totmaintien := 0 ;
       Wi := 0;
       for i := 0 to Tob_ReglIJSS.detail.count-1 do
       begin
         TReglt := Tob_ReglIJSS.detail[wi];
// PT31         Wi := i ;
// Y-a-t'il eu du maintien sur l'absence dont les IJSS on été réglées ?
// si non on ne doit pas intégrer les retenues IJSS s'il n'y a pas de subrogation
         st := 'SELECT * FROM MAINTIEN WHERE '+
               'PMT_SALARIE = "'+treglt.getValue('PRI_SALARIE')+'" AND '+
               'PMT_DATEDEBUTABS <= "'+
               UsDateTime(treglt.getValue('PRI_DATEFINABS'))+'" AND '+
               'PMT_DATEFINABS >= "'+
               UsDateTime(treglt.getValue('PRI_DATEDEBUTABS'))+'" ';
         Q := opensql(st, True);
         TotMaintien := 0;

         if not Q.eof then
         begin
           Q.first;
           While not Q.eof do
           begin
             if (Q.FindField('PMT_PCTMAINTIEN').AsFloat <> 0) then
               Inc(TotMaintien) ; // XP 10.10.2007 := TotMaintien +1 ;
             Q.Next;
           end;
         end;
         if (Tob_Maintien <> nil) then
         for j:= 0 to Tob_Maintien.detail.count-1 do
         begin
           tMaintien := Tob_Maintien.Detail[j];
           if (tmaintien.GetValue ('PMT_SALARIE') = treglt.getValue('PRI_SALARIE')) and
              (tmaintien.GetValue ('PMT_DATEDEBUTABS') <= treglt.getValue('PRI_DATEFINABS')) and
              (tmaintien.GetValue ('PMT_DATEFINABS') >= treglt.getValue('PRI_DATEDEBUTABS')) then
             if (tmaintien.GetValue ('PMT_MTMAINTIEN') <> 0) then
             Inc(TotMaintien) ; // XP 10.10.2007  := TotMaintien + 1;
         end;
         if (Tob_etab <> nil) and (Tob_etab.GetValue('ETB_SUBROGATION') <> 'X') and
            (TotMaintien = 0) then
         begin
           Treglt.free;
// d PT31
//          Wi := Wi -1;
//          if (Wi < 0) then Wi := 0;
         end
         else
             Wi := Wi+1;
// f PT31
       end;
     end;

  result := Tob_ReglIJSS;

  Ferme(Q);

end;
{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 13/10/2004
Modifié le ... : 13/10/2004
Description .. :
Suite ........ : procédure MaintienDuSalarié
Suite ........ : Calcul du maintien en fonction des absences et des
Suite ........ : règlements d'IJSS à intégrer, du paramétrage des règles de
Suite ........ : maintien (CRITMAINTIEN,REGLEMAINTIEN)
Suite ........ : Utilisation du calcul de la paie à l'envers pour calculer le
Suite ........ : montant de la garantie en fct du mt à garantir (Net, Brut....)
Suite ........ : Création des lignes de maitien et de garantie de salaire
Suite ........ : (table MAINTIEN)
Suite ........ : On doit maintenir le salaire que le salarié aurait percu s'il
Suite ........ : n'était pas malade.
Mots clefs ... : PAIE IJSS MAINTIEN
*****************************************************************}
procedure MaintienDuSalarie (const TobSal,TOB_Rub,Tob_etablissement :TOB;var TOB_abs : TOB; Datef, Dated, Mode : string; var Anomalie : integer; const ChampCateg :string; var TOB_IJSS, TOB_Maintien : TOB);

var
  Anciennete                                            : WORD;
  Q                                                     : TQuery;
  st,wdate                                              : string;
  Predef,NoDossier,CodeMaint,Convention                 : string;
  Standard                                              : string;
  Tob_AbsIJ, tabsence                                     : Tob;
  Tob_ReglesMaintien                                    : Tob;
  tMaintien,tMajMaintien                                : Tob;
  Tob_etab                                              : Tob;
  WDateFin                                              : TDateTime;
  Bornefinanc                                           : integer;
  WCarence, WNbjMaintien                                : double;
  NbRestant,TotAbsence,WMtMaintien                      : double;
  ACheval                                               : boolean;
  Salarie,DateAnciennete                                : string;
  MtAGarantir, PlusApprochant, MtGarantie, Wess         : double;
  LaRubriq,Sauv_PGEnversNet, RubIJSSNettes              : string;
  Tlig,Tob_RegltIJSS, treglt,TR                         : TOB;
  RetEnvers, IJSS, MaintienNet ,subrogation             : boolean;
  typeabs ,CumulAGarantir                               : string;
  Calendrier,StandCalend                                : string;
  Abprof, AbArub,AbAliment                              : string;
  AbgereCom,Abheure                                     : boolean;
  QAb                                                   : TQuery;
  AbsIJOk                                               : integer;
  RechDebutAbs                                          : boolean;
  DateDebutAbsReel                                      : TDateTime;
  DateDebutAbs ,DateFinAbsPrec                          : TDateTime;
  Libelle                                               : string;
  Tob_TypeAbs, TRgl, TTypeAbs                           : TOB;
  i, Poids, PoidsMax, NoCrit                            : Integer;
  Tob_CritMaintien, tCritmaintien                       : TOB;
  WdateDeb,DF                                           : TDateTime;
  LaRubriqGar                                           : string;
  StAlimH,StProfilH,StRubH,StAlimJ,StProfilJ,StRubJ     : string;
  NetH,HetJ                                             : Boolean;
  Tob_AbsIJBis                                          : TOB; // PT35
// d PT37
  VALCATEGORIE,ListeVALCATEGORIE : string;
  CatgegOK  : boolean;
// f PT37
  TypeAbsPreced : string;
begin
   RegleNo := 0;
   DateDebutAbs := IDate1900;
   AbHeure := False;
   AbsIJOk := 0;
   FreeAndNil (Tob_Maintien);

   Salarie := TobSal.GetValue('PSA_SALARIE');
   DateAnciennete := TobSal.GetValue('PSA_DATEANCIENNETE');
// d PT40
   if (StrToDate(DateAnciennete) <= IDate1900) then
   begin
     Anomalie := 4;
     exit;
   end;
// f PT40
   Convention := TobSal.GetValue('PSA_CONVENTION');
   Calendrier := TobSal.GetValue('PSA_CALENDRIER');
   StandCalend := TobSal.GetValue('PSA_STANDCALEND');
   Tob_etab := TOB_Etablissement.findfirst(['ETB_ETABLISSEMENT'],
                                           [TobSal.GetValue('PSA_ETABLISSEMENT')], True);

   WDateFin := IDate1900;
   WCarence := 0;
   TotAbsence := 0;
   WnbjMaintien := 0;
   WMtMaintien := 0;
   NbRestant := 0;

   // Création de la Tob des Absences du salarié pour la période.
   // ===========================================================
   st :='SELECT ABSENCESALARIE.* ';
   if (ChampCateg <> '') then
   begin
     st:=st+',SALARIES.'+ChampCateg+' ';
   end;
   st := st+',MOTIFABSENCE.PMA_TYPEABS '; // PT35
   st := st+'FROM ABSENCESALARIE ';
   if (ChampCateg <> '') then
   begin
     st := st +'LEFT JOIN SALARIES ON PCN_SALARIE=PSA_SALARIE ';
   end;
   st := st +'LEFT JOIN MOTIFABSENCE ON PCN_TYPECONGE=PMA_MOTIFABSENCE ';  // PT35

   st := st+ 'WHERE'+
         ' PCN_SALARIE = "'+   Salarie+ '"'+
         ' AND PCN_DATEVALIDITE <= "' + usdatetime(StrToDate(Datef))+'"' +
         ' AND PCN_DATEPAIEMENT <= "' + usdatetime(10)+'"'+
         ' AND PCN_GESTIONIJSS ="X"';
   Q:=OpenSql(st, TRUE);
   if not Q.eof then
   begin
     Tob_AbsIJ:= Tob.create('les absences du salarié',nil,-1);
     Tob_AbsIJ.LoadDetailDB('ABSENCESALARIE','','',Q,False);
  end;
   Ferme(Q);
// d PT35
    // duplication de la tob des absences pour calculer le total des absences par type
    if (Tob_AbsIJ <> nil) then
    begin
      Tob_AbsIJBis:= Tob.create('absences salarié pout Tot',nil,-1);
      Tob_AbsIJBis.Dupliquer(Tob_AbsIJ,TRUE,TRUE,TRUE) ;
    end;
// f PT35

   // ici calcul nbj calendaires total (il peut y avoir plusieurs absences à traiter)
   if (Assigned(Tob_AbsIJ)) and (Tob_AbsIJ.detail.count > 0) then
   begin
     // tri de la tob des absences par date de début d'absence afin de traiter
     // les absences dans l'ordre chronologique
     Tob_AbsIJ.Detail.Sort('PCN_DATEDEBUTABS;');
// PT35   TotAbsence := Tob_AbsIJ.somme('PCN_NBJCALEND',['PCN_SALARIE'],[salarie],true);;
   end;


   // Pour chaque absence du salarié
   // ==============================
   if (Tob_AbsIJ <> nil) then
   begin
     tabsence:= Tob_AbsIJ.Findfirst([''],[''],false);

     while tabsence <> nil do
     begin
       // récup. type absence
       st := 'SELECT PMA_TYPEABS FROM MOTIFABSENCE WHERE ##PMA_PREDEFINI## PMA_MOTIFABSENCE = "' +
       tabsence.getvalue('PCN_TYPECONGE') + '"';
       Q := opensql(st, True);
       if not Q.eof then
         TypeAbs := Q.findfield('PMA_TYPEABS').AsString;
       ferme(Q);
// d PT39
       // on ne récupère les absences de type IJSS qu'une fois par type d'absence
       if (TypeAbs <> TypeAbsPreced) then
       begin
          AbsIJOk := 0;
          TypeAbsPreced := TypeAbs;
       end;
// f PT39

// d PT35
      //Calcul du total des absences par type
      TotAbsence := Tob_AbsIJBis.somme('PCN_NBJCALEND',['PCN_SALARIE', 'PMA_TYPEABS'],[salarie,TypeAbs],true);;
// f PT35
       Tob_etab := TOB_Etablissement.findfirst(['ETB_ETABLISSEMENT'],
                                               [Tabsence.GetValue('PCN_ETABLISSEMENT')], True);
//     recherche date début absence
       RechDebutAbs := True;
       DateDebutAbsReel := tabsence.getValue('PCN_DATEDEBUTABS');
       DateFinAbsPrec := DateDebutAbsReel-1 ;
       while RechDebutAbs do
       begin
         st := 'SELECT PCN_DATEDEBUTABS FROM ABSENCESALARIE WHERE'+
               ' PCN_SALARIE = "'+   Salarie+ '"'+
               ' AND PCN_GESTIONIJSS ="X"'+
               ' AND PCN_DATEFINABS ="'+UsDateTime(DateFinAbsPrec)+'"'+
               ' AND PCN_TYPECONGE IN (SELECT PMA_MOTIFABSENCE FROM MOTIFABSENCE'+
               ' WHERE ##PMA_PREDEFINI## AND PMA_TYPEABS="'+TypeAbs+'")';
         Q := opensql(st, True);
         if not Q.eof then
         begin
           DateDebutAbsReel := Q.Fields[0].AsDateTime;
           DateFinAbsPrec := DateDebutAbsReel-1;
         end
         else
           RechDebutAbs := False;
         ferme (Q);
       end;

       // calcul de l'ancienneté du salarié traité à la date de début d'absence
       // ----------------------------------------------------------------------
       Anciennete := AncienneteMois (StrToDate(DateAnciennete),PlusDate(DateDebutAbsReel,1,'J'));
       tabsence.AddChampSupValeur('ANCIENNETE', Anciennete, False);

       // récup des paramètres de maintien

       // récup critères de maintien
       // -------------------------------
       st := 'SELECT * FROM CRITMAINTIEN WHERE ##PCM_PREDEFINI## (PCM_TYPEABS = "' +
             TypeAbs+'" OR PCM_TYPEABS="") AND  PCM_BORNEFINANC >= '+
             IntToStr(Anciennete)+' AND PCM_BORNEDEBANC <= '+IntToStr(Anciennete);

       if (Convention <> '') then
           st := st + ' AND (PCM_CONVENTION ="'+Convention+'" OR '+
                 'PCM_CONVENTION ="000")'
         else
           st := st + ' AND PCM_CONVENTION ="000"';
       if (ChampCateg<>'') then
         if (tabsence.GetValue(ChampCateg)= '') then
           st := st + ' AND (PCM_VALCATEGORIE = "" OR PCM_VALCATEGORIE="<<Tous>>")'
         else
           st := st + ' AND (PCM_VALCATEGORIE LIKE "%'+tabsence.GetValue(ChampCateg)+'%"'+
                      ' OR PCM_VALCATEGORIE="" OR PCM_VALCATEGORIE="<<Tous>>")';
       Q:=OpenSql(st, TRUE);

       if not Q.eof then
       begin
         PoidsMax := 0;
         Tob_CritMaintien := Tob.create('les tables de critère ',nil,-1);
         Tob_CritMaintien.LoadDetailDB('CRITMAINTIEN','','',Q,False);
         // plusieurs Critères de maintien peuvent correspondre à la sélection
         // il faut définir l'élément de la table CRITMAINTIEN optimum, on donne
         // donc à chaque élément de la tob un poids
         if (Tob_CritMaintien <> nil) and (Tob_CritMaintien.Detail.count <> 0) then
         begin
           tCritMaintien := Tob_CritMaintien.FindFirst([''],[''],false);
           for i:= 1 to Tob_CritMaintien.Detail.count do
           begin
             Poids := 0;
             // PREDEFINI
             if (tCritMaintien.GetValue('PCM_PREDEFINI') = 'CEG') then
               // CEG
               Poids := Poids + 1
             else
               if (tCritMaintien.GetValue('PCM_PREDEFINI') = 'STD') then
                 // STD
                 Poids := Poids + 2
               else
                 // DOS
                 Poids := Poids + 3;

             // TYPEABS
             if (tCritMaintien.GetValue('PCM_TYPEABS') = TypeAbs) then
               Poids := Poids + 2
             else
               Poids := Poids +1;

             // CONVENTION
             if ((tCritMaintien.GetValue('PCM_CONVENTION') = Convention) or
                 ((tCritMaintien.GetValue('PCM_CONVENTION') = '000') and
                  (Convention = ''))) then
               Poids := Poids + 2
             else
               Poids := Poids +1 ;

             // VALCATEG
// d PT37
{ PCM_VALCATEGORIE est un MultiValComboBox : Il faut vérifier que le
  code tabsence.GetValue(ChampCateg) soit égal à une des valeurs de PCM_VALCATEGORIE}
           ListeVALCATEGORIE := tCritMaintien.GetValue('PCM_VALCATEGORIE');
           if (ListeVALCATEGORIE <> '<<Tous>>') then
           begin
             VALCATEGORIE := READTOKENST(ListeVALCATEGORIE);
             CatgegOK := False;
             while VALCATEGORIE <> '' do
             begin
               if ((ChampCateg<>'') and (VALCATEGORIE = tabsence.GetValue(ChampCateg))) then
               begin
                 Poids := Poids + 2;
                 CatgegOK := True;
                 break;
               end;
               VALCATEGORIE := READTOKENST(ListeVALCATEGORIE);
             end;
             if (not CatgegOK) then
               Poids := Poids + 1;
           end
           else
           // ListeVALCATEGORIE = '<<Tous>>'
             Poids := Poids + 2;

{             if ((ChampCateg<>'') and ((tCritMaintien.GetValue('PCM_VALCATEGORIE') = // PT34
                 tabsence.GetValue(ChampCateg)) or
                 ((tCritMaintien.GetValue('PCM_VALCATEGORIE') = '<<Tous>>') and
                  (tabsence.GetValue(ChampCateg)='')))) then
               Poids := Poids + 2
             else
               Poids := Poids +1; }
// f PT37

             TCritmaintien.AddChampSupValeur('POIDS', Poids, False);
             if (Poids > PoidsMax) then
             // c'est lélément de poids le + élévé qui sera retenu
             begin
               PoidsMax := Poids;
               NoCrit := i;
             end;
           tCritMaintien := Tob_CritMaintien.FindNext([''],[''],false);
           end;
         end;
       end
       else

       // pas de paramétrage CRITMAINTIEN correspondant aux critères
       begin
         Anomalie := 1;
         Ferme(Q);
         exit;
       end;

       tCritMaintien := Tob_CritMaintien.Detail[NoCrit-1];

       tabsence.AddChampSupValeur('PREDEFINI', tCritMaintien.GetValue('PCM_PREDEFINI'), False);
       tabsence.AddChampSupValeur('NODOSSIER', tCritMaintien.GetValue('PCM_NODOSSIER'), False);
       tabsence.AddChampSupValeur('CODEMAINT', tCritMaintien.GetValue('PCM_CODEMAINT'), False);
       tabsence.AddChampSupValeur('CONVENTION', tCritMaintien.GetValue('PCM_CONVENTION'), False);
       tabsence.AddChampSupValeur('BORNEFINANC', tCritMaintien.GetValue('PCM_BORNEFINANC'), False);
       tabsence.AddChampSupValeur('MAINTIENNET', tCritMaintien.GetValue('PCM_MAINTIENNET'), False);
       tabsence.AddChampSupValeur('RUBIJSSNETTE', tCritMaintien.GetValue('PCM_RUBIJSSNETTE'), False);
       tabsence.AddChampSupValeur('RUBIJSSBRUTE', tCritMaintien.GetValue('PCM_RUBIJSSBRUTE'), False);
       tabsence.AddChampSupValeur('RUBMAINTIEN', tCritMaintien.GetValue('PCM_RUBMAINTIEN'), False);
       tabsence.AddChampSupValeur('RUBGARANTIE', tCritMaintien.GetValue('PCM_RUBGARANTIE'), False);
       tabsence.AddChampSupValeur('CUMULPAIE', tCritMaintien.GetValue('PCM_CUMULPAIE'), False);
       Predef := tCritMaintien.GetValue('PCM_PREDEFINI');
       NoDossier := tCritMaintien.GetValue('PCM_NODOSSIER');
       CodeMaint := tCritMaintien.GetValue('PCM_CODEMAINT');
       Bornefinanc := tCritMaintien.GetValue('PCM_BORNEFINANC');
       Convention := tCritMaintien.GetValue('PCM_CONVENTION');
       Libelle := tCritMaintien.GetValue('PCM_LIBELLE')  ;

       if (tCritMaintien.GetValue('PCM_MAINTIENNET') = 'X') then
         MaintienNet := True
       else
         maintienNet := False;

       Ferme (Q);

       if (Tob_CritMaintien <> nil) then
       FreeAndNil(Tob_CritMaintien);
//    création tob des types d'absence à gérer de façon identique
       st := 'SELECT PCM_TYPEABS FROM CRITMAINTIEN WHERE '+
             'PCM_PREDEFINI = "'+Predef+'" AND '+
             'PCM_NODOSSIER = "'+NoDossier+'" AND '+
             'PCM_CONVENTION = "'+Convention+'" AND '+
             'PCM_BORNEFINANC = '+IntToStr(bornefinanc)+' AND '+
             'PCM_LIBELLE ="'+Libelle+'" AND '+
             'PCM_CODEMAINT <> "'+CodeMaint+'"';
       Q:=OpenSql(st, TRUE);
       if not Q.eof then
       begin
         Tob_TypeAbs := Tob.create('les types absences identiques',nil,-1);
         Tob_TypeAbs.LoadDetailDB('REGLESMAINTIEN','','',Q,False);
       end; { fin if not Q.eof du  SELECT * FROM CRITMAINTIEN }
       ferme(Q);

       // Récup. des régles du maintien à appliquer
       // -----------------------------------------
       st := 'SELECT * FROM REGLESMAINTIEN WHERE PAM_PREDEFINI = "'+Predef+'"'+
             ' AND PAM_NODOSSIER = "'+ NoDossier + '"'+
             ' AND PAM_BORNEFINANC = '+ IntToStr(Bornefinanc) +
             ' AND PAM_CODEMAINT = "'+ CodeMaint +'"'+
             ' AND PAM_CONVENTION = "'+ Convention +'"';

       Q:=OpenSql(st, TRUE);
       if not Q.eof then
       begin
         Tob_ReglesMaintien := Tob.create('les régles du maintien',nil,-1);
         Tob_ReglesMaintien.LoadDetailDB('REGLESMAINTIEN','','',Q,False);
       end; { fin if not Q.eof du  SELECT * FROM REGLESMAINTIEN }
       ferme (Q);

//     création des tableaux de cumul du nb jours maintien par taux
       SetLength (TbNbj, Tob_ReglesMaintien.detail.count);
       SetLength (TbNbjMax, Tob_ReglesMaintien.detail.count);
       SetLength (TbPct, Tob_ReglesMaintien.detail.count);
       SetLength (TbCarence, Tob_ReglesMaintien.detail.count);
       SetLength (TbWCArence, Tob_ReglesMaintien.detail.count);

       TRgl := Tob_ReglesMaintien.FindFirst([''],[''],false);
       for i:= 0 to Tob_ReglesMaintien.detail.count-1 do
       begin
         if (Trgl <> nil) then
         begin
           TbPct [i] := TRgl.GetValue('PAM_TXMAINTIEN');
           TbNbjMax [i] := TRgl .GetValue('PAM_NBJMAINTIEN');
           TbCarence [i] := TRgl .GetValue('PAM_CARENCE');
           if (RegleNo = 0) then
           begin
             TbNbj [i] := 0;
             TbWCarence[i] := 0;
           end;
         end;
         TRgl := Tob_ReglesMaintien.FindNext([''],[''],false);
       end;

       wdate :=  DateToStr(tabsence.GetValue ('PCN_DATEDEBUTABS'));

       // raz rub abs IJSS + rub. IJSS et maintien
       //-----------------------------------------

       // absence IJSS
{PT35
       QAb:=OpenSql('SELECT PCN_TYPECONGE FROM ABSENCESALARIE WHERE'+
                    ' PCN_SALARIE = "'+   Salarie+ '"'+
                    ' AND PCN_DATEVALIDITE <= "' + usdatetime(StrToDate(Datef))+'"' +
                    ' AND (PCN_DATEPAIEMENT = "' + usdatetime(StrToDate(Datef))+'"'+
                    ' OR PCN_DATEPAIEMENT <="'+ usdatetime(2)+'")'+
                    ' AND PCN_GESTIONIJSS ="X"',TRUE);}
       QAb:=OpenSql('SELECT PCN_TYPECONGE FROM ABSENCESALARIE WHERE'+
                    ' PCN_SALARIE = "'+   Salarie+ '"'+
                    ' AND PCN_DATEVALIDITE <= "' + usdatetime(StrToDate(Datef))+'"' +
                    ' AND (PCN_DATEPAIEMENT = "' + usdatetime(StrToDate(Datef))+'"'+
                    ' OR PCN_DATEPAIEMENT <="'+ usdatetime(2)+'")'+
                    ' AND PCN_GESTIONIJSS ="X" '+
                    ' AND PCN_TYPECONGE IN '+
                    '(SELECT PMA_MOTIFABSENCE FROM MOTIFABSENCE WHERE PMA_TYPEABS '+
                    '= "'+ typeAbs +'")',TRUE);
       While not QAb.eof do
       begin
         RechercheCarMotifAbsence(Qab.Fields[0].AsString,StProfilH,StRubH,StAlimH,StProfilJ,StRubJ,StAlimJ,AbgereCom,Abheure,NetH,HetJ);
         if NetH then
           LaRubriq := StRubH
         else
           LaRubriq := StRubJ;
         if  (LaRubriq <> '') then
         begin
           Tlig := TOB_Rub.FindFirst(['PHB_NATURERUB', 'PHB_RUBRIQUE'], ['AAA', LaRubriq], FALSE);

           // Remise à 0 la valeur de la ligne pour le calcul de la paie à l'envers
           if TLig <> nil then
           begin
             TLig.putValue('PHB_MTREM', 0);
             TLig.putValue('PHB_BASEREM', 0);
           end;
         end;
         Qab.Next;
       end;
       ferme(Qab);

       if  (LaRubriq <> '') then
       begin
         Tlig := TOB_Rub.FindFirst(['PHB_NATURERUB', 'PHB_RUBRIQUE'], ['AAA', LaRubriq], FALSE);
         // Remise à 0 la valeur de la ligne pour le calcul de la paie à l'envers
         if TLig <> nil then TLig.putValue('PHB_MTREM', 0);
       end;

       // IJSS réglées (IJSS nettes)
       LaRubriq := '';
       LaRubriq := tAbsence.GetValue('RUBIJSSNETTE');
       if  (LaRubriq <> '') then
       begin
         Tlig := TOB_Rub.FindFirst(['PHB_NATURERUB', 'PHB_RUBRIQUE'], ['AAA', LaRubriq], FALSE);
         // Remise à 0 la valeur de la ligne pour le calcul de la paie à l'envers
         if TLig <> nil then TLig.putValue('PHB_MTREM', 0);
       end;

       // IJSS Brutes
       LaRubriq := '';
       LaRubriq := tAbsence.GetValue('RUBIJSSBRUTE');
       if  (LaRubriq <> '') then
       begin
         Tlig := TOB_Rub.FindFirst(['PHB_NATURERUB', 'PHB_RUBRIQUE'], ['AAA', LaRubriq], FALSE);
         // Remise à 0 la valeur de la ligne pour le calcul de la paie à l'envers
         if TLig <> nil then TLig.putValue('PHB_MTREM', 0);
       end;

       // maintien
       LaRubriq := '';
       LaRubriq :=  tAbsence.GetValue('RUBMAINTIEN');
       if  (LaRubriq <> '') then
       begin
         Tlig := TOB_Rub.FindFirst(['PHB_NATURERUB', 'PHB_RUBRIQUE'], ['AAA', LaRubriq], FALSE);
       // Remise à 0 la valeur de la ligne pour le calcul de la paie à l'envers
         if TLig <> nil then TLig.putValue('PHB_MTREM', 0);
       end;

       // garantie
       LaRubriqGar := '';
       LaRubriqGar :=  tAbsence.GetValue('RUBGARANTIE');
       if  (LaRubriqGar <> '') then
       begin
         Tlig := TOB_Rub.FindFirst(['PHB_NATURERUB', 'PHB_RUBRIQUE'], ['AAA', LaRubriqGar], FALSE);
       // Remise à 0 la valeur de la ligne pour le calcul de la paie à l'envers
         if TLig <> nil then TLig.putValue('PHB_MTREM', 0);
       end;

       CalculBulletin(Tob_Rub,True);


       // Calcul du maintien net ou brut à appliquer
       // ------------------------------------------
       // On récupère le montant à garantir (net à payer, brut, brut habituel...en fct du paramétrage)
       CumulAGarantir := tAbsence.GetValue('CUMULPAIE');
       if (CumulAGarantir = 'NET') then
         MtAGarantir := RendCumulSalSess('10')
       else
        if (CumulAGarantir = 'BRU') then
         MtAGarantir := RendCumulSalSess('01')
        else
          if (CumulAGarantir = 'BRH') then
            MtAGarantir := RendCumulSalSess('05')
          else
            MtAGarantir := RendCumulSalSess('07') + RendCumulSalSess('08') + RendCumulSalSess('10');

       // récupération de la Rubrique de garantie du salaire
       LaRubriqGar :=  tAbsence.GetValue('RUBGARANTIE');
       if (MtAGarantir <> 0) and (LaRubriqGar <> '') then
       begin
         Tlig := TOB_Rub.FindFirst(['PHB_NATURERUB', 'PHB_RUBRIQUE'], ['AAA', LaRubriqGar], FALSE);
         if TLig = nil then
         begin
           // on insere la rubrique dans le bulletin
           ChargeProfilSPR(TobSal, TOB_Rub, LaRubriqGar, 'AAA');
           Tlig := TOB_Rub.FindFirst(['PHB_NATURERUB', 'PHB_RUBRIQUE'], ['AAA', LaRubriqGar], FALSE);
         end;
         // Remise à 0 la valeur de la ligne pour le calcul de la paie à l'envers
        if TLig <> nil then TLig.putValue('PHB_MTREM', 0);
       end;

       LaRubriq := '';
       if (MtAGarantir <> 0) then
       // récupération de la Rubrique de garantie du maintien
         LaRubriq :=  tAbsence.GetValue('RUBMAINTIEN');
       if (MtAGarantir <> 0) and (LaRubriq <> '') then
       begin
         Tlig := TOB_Rub.FindFirst(['PHB_NATURERUB', 'PHB_RUBRIQUE'], ['AAA', LaRubriq], FALSE);
         if TLig = nil then
         begin
           // on insere la rubrique dans le bulletin
           ChargeProfilSPR(TobSal, TOB_Rub, LaRubriq, 'AAA');
           Tlig := TOB_Rub.FindFirst(['PHB_NATURERUB', 'PHB_RUBRIQUE'], ['AAA', LaRubriq], FALSE);
         end;
         // Remise à 0 la valeur de la ligne pour le calcul de la paie à l'envers
         if TLig <> nil then TLig.putValue('PHB_MTREM', 0);

       end;

       // récupération des absence de type IJSS
       IJSS := True;
// le RecupereAbsences charge Tob_Abs: Pour ne charger qu'une seule fois les absences de
// type IJSS on utilise l'indicateur AbsIJOk
{PT35
       if (AbsIJOk = 0) then
         if  RecupereAbsences(Tob_Abs,StrToDate(Datef), Salarie,Mode, TRUE, IJSS) then
           AbsIJOk := 1;
       if (AbsIjOk = 1) then}
//       if  RecupereAbsences(Tob_Abs,StrToDate(Datef), Salarie,Mode, TRUE, IJSS) then
// d PT39
       if (AbsIJOk = 0) then
          if  RecupereAbsences(Tob_Abs,StrToDate(Datef), Salarie,Mode, TRUE, IJSS, TypeAbs) then
          AbsIJOk := 1;
       if (AbsIjOk = 1) then
         IntegreAbsenceDansPaye(Tob_Rub, Tob_Abs, TobSal, StrToDate(DateD), StrToDate(DateF), Mode);

// d PT38
{       if (Tob_AbsIJBis <> nil) then
         FreeAndNil (Tob_AbsIJBis);  *}
// f PT38
// f PT39
       CalculBulletin(TOB_Rub,True);
       RetEnvers := TRUE;
       if (MtAGarantir <> 0) and (LaRubriq <> '') then
       begin
         if TLig <> nil then
         begin
           PlusApprochant := 0;

           Sauv_PGEnversNet := VH_Paie.PGEnversNet;
           VH_Paie.PGEnversNet := CumulAGarantir;
           RetEnvers := CalculPEnvers(PlusApprochant, TOB_Rub, TLig, MtAGarantir, LaRubriq, TRUE, FALSE, TRUE);
           //  On relance le calcul avec arrondi si echec
           if (NOT RetEnvers) AND (VH_PAIE.PGArrondiEnvers >0) then
             RetEnvers := CalculPEnvers(PlusApprochant, TOB_Rub, TLig, MtAGarantir, LaRubriq, TRUE, TRUE, TRUE);

           VH_Paie.PGEnversNet := Sauv_PGEnversNet;

           if not RetEnvers then
           begin
             Anomalie := 2;
           end;

           Tlig := TOB_Rub.FindFirst(['PHB_NATURERUB', 'PHB_RUBRIQUE'], ['AAA', LaRubriq], FALSE);
           if TLig <> nil then
           begin
             MtGarantie := TLig.getValue('PHB_MTREM');
           end;
         end;
       end;

       // maj MAINTIEN et raz TOB
       if (Tob_Maintien <> Nil) then
       begin
         MajMaintien(Tob_Maintien);
         Wcarence := 0;
       end;
      if (not Acheval) then // NB: si(ACheval = true) la récup a déjà été faite
      begin
       // Récup du maintien des 12 mois précédents
       // ----------------------------------------
      if (VH_Paie.PGHistoMaintien = '003') then
       // historique depuis début exercice social
         RendDateExerSocial(tabsence.GetValue ('PCN_DATEDEBUTABS'), tabsence.GetValue ('PCN_DATEFINABS'), WdateDeb, DF)
      else
        if (VH_Paie.PGHistoMaintien = '001') then
          // historique glissant sur 12 mois
          WdateDeb := PlusMois(StrToDate(wdate),-12)
        else
          if (VH_Paie.PGHistoMaintien = '002') then
          // historique depuis début année civile
            WdateDeb := StrToDate('01/01/'+Copy(DateToStr(tabsence.GetValue ('PCN_DATEDEBUTABS')),7,4));
      st := 'SELECT * FROM MAINTIEN WHERE PMT_SALARIE = "'+Salarie+'"'+
            ' AND PMT_DATEDEBUTABS >= "'+
             UsDateTime(WdateDeb) +
             '" AND PMT_DATEDEBUTABS <= "'+
             UsDateTime(PlusDate(StrToDate(wdate), -1, 'J'))+'"';

//     récupérer les types absences identiques
       if (Tob_TypeAbs = nil ) or (Tob_TypeAbs.Detail.count = 0) then
       begin
         if (TypeAbs<> '') then
           st := st +' AND PMT_TYPECONGE IN (SELECT PMA_MOTIFABSENCE '+
                     'FROM MOTIFABSENCE WHERE ##PMA_PREDEFINI## '+
                     'PMA_TYPEABS ="'+TypeAbs+'")';
       end
       else
       begin
         st := st +' AND (PMT_TYPECONGE IN (SELECT PMA_MOTIFABSENCE '+
                   ' FROM MOTIFABSENCE WHERE ##PMA_PREDEFINI## (PMA_TYPEABS ="'+TypeAbs+'" OR ';
         TTypeAbs := Tob_TypeAbs.FindFirst([''],[''], false);
         for i:=1 to Tob_TypeAbs.Detail.count do
         begin
           st := st +'PMA_TYPEABS= "'+TTypeAbs.GetValue('PCM_TYPEABS')+'"';
           TTypeAbs := Tob_TypeAbs.FindNext([''],[''], false);
           if (TTypeAbs <> nil) then
             st := st + ' OR ';
         end;
         st := st + ')))';
       end;

       Q:=OpenSql(st, TRUE);
       if not Q.eof then
       begin
         Tob_Maintien := Tob.create('les maintiens du salariés',nil,-1);
         Tob_Maintien.LoadDetailDB('MAINTIEN','','',Q,False);

         Tob_Maintien.Detail.Sort('PMT_DATEDEBUTABS');
         WDateFin := IDate1900;
         WCarence := 0;
         WNbjMaintien := 0;

         MaintConsec := true;
         if (Tob_Maintien.Detail[Tob_Maintien.Detail.count-1].GetValue('PMT_PCTMAINTIEN') = 0) then
         // le dernier maintien de l'historique est à 0
           MaintienAZero := true
         else
           MaintienAZero := false;

         tMaintien:= Tob_Maintien.Findfirst([''],[''],false);
         while tMaintien <> nil do
         begin
           if (tMaintien.GetValue('PMT_TYPEMAINTIEN') = 'MAI') then
           begin
             // Cumul du nbre de jours de maintien des 12 mois précédents
             WNbjMaintien := WNbjMaintien + tMaintien.GetValue('PMT_NBJMAINTIEN');

             // Cumul du nbre de jours de carence des 12 mois précédents
             if (tMaintien.GetValue('PMT_DATEDEBUTABS') <> WDateFin + 1) then
               WCarence := tMaintien.GetValue('PMT_CARENCE')
             else
               WCarence := WCarence +tMaintien.GetValue('PMT_CARENCE');

//           tenue du cumul des nbre de jours de maintien (et carence) par taux
             for i := 0 to Tob_ReglesMaintien.Detail.count-1 do
             begin
               if (TbPct[i] = tMaintien.GetValue('PMT_PCTMAINTIEN')) then
               begin
                 TbNbj[i] := TbNbj[i] + tMaintien.GetValue('PMT_NBJMAINTIEN');
                 break;
               end
             end;
             if (tMaintien.GetValue('PMT_DATEDEBUTABS') <> WDateFin + 1) and
                ( WDateFin <> IDate1900) then
             // les dates de maintien ne sont pas consécutives
               MaintConsec := False;

             WDateFin := tMaintien.GetValue('PMT_DATEFINABS');
           end;
           tMaintien := Tob_Maintien.FindNext([''],[''],false);
         end; { fin while tMaintien <> nil }
       end; { fin if not Q.eof du SELECT FROM MAINTIEN}
       ferme (Q);
      end;

       // raz cumul carence si l'absence traitée n'est pas consécutive à la
       // l'absence correspondant à la derniére ligne de la table MAINTIEN
       if (tAbsence.GetValue('PCN_DATEDEBUTABS') <> WDateFin+1) then
         WCarence := 0;
       if (tAbsence.GetValue('PCN_DATEDEBUTABS') <> WDateFin+1) then
       // la date d'absence n'est pas consécutive au denier maintien
          MaintConsec := False;
       for i := 0 to Tob_ReglesMaintien.Detail.count-1 do
       begin
         if (tAbsence.GetValue('PCN_DATEDEBUTABS') <> WDateFin+1) then
           TbWCarence[i] := 0;
       end;
       // Maj Maintien : création nouvelle ligne
       //========================================
       if (Tob_Maintien = nil) then
         Tob_Maintien := Tob.create('les maintiens du salarié',nil,-1);

       tMaintien := TOB.Create('MAINTIEN',Tob_Maintien, -1);

       NvelleLMaintien(Tob_ReglesMaintien ,Tob_etab, Wcarence,TotAbsence, MtGarantie,MaintienNet,Tabsence,treglt,tmaintien,ACheval,WMtMaintien,NbRestant,DateDebutAbs,0,IDate1900,IDate1900);

       if (ACheval = false) then // NB: Si (ACheval=True) on traite toujours la même absence
       begin
         DateDebutAbs := IDate1900;
         tabsence := Tob_AbsIJ.Findnext([''],[''],false);
         RegleNo := 0;
         NbRestant := 0;
       end;

       if (Tob_ReglesMaintien <> nil) then
         FreeAndNil (Tob_ReglesMaintien);

//     vidage tob_typeabs
       if (Tob_TypeAbs <> nil) then
         FreeAndNil (Tob_TypeABs);
     end; {fin while tabsence <> nil}

// d PT39
     // on libère la tob dupliquée des absences
     if (Tob_AbsIJBis <> nil) then
        FreeAndNil (Tob_AbsIJBis);
// f PT39

     // calcul du maintien (Garantie) en cas de règlement d'IJSS
     //=========================================================
     Tob_RegltIJSS := RecupereRegltIJSS(StrToDate(Datef), Salarie, Mode,ChampCateg, Tob_Maintien, Tob_Etab);
     if (Tob_RegltIJSS <> nil) and (Tob_RegltIJSS.Detail.count > 0) then
     begin
       TReglt := Tob_RegltIJSS.Findfirst([''],[''],false);
       if (TReglt <> nil) then
       begin
         if (LaRubriqGar  = '') then
         // pas de paramétrage de la rubrique de garantie dans CRITMAINTIEN
         begin
           Anomalie := 3;
         end
         else
           Tlig := TOB_Rub.FindFirst(['PHB_NATURERUB', 'PHB_RUBRIQUE'], ['AAA', LaRubriqGar], FALSE);
         if (Tlig <> nil) then
           TLig.putValue('PHB_MTREM',0);
         CalculBulletin(TOB_Rub, True);

         if (CumulAGarantir= 'NET') then
           MtAGarantir := RendCumulSalSess('10')
         else
           if (CumulAGarantir = 'BRU') then
             MtAGarantir := RendCumulSalSess('01')
           else
             if (CumulAGarantir = 'BRH') then
               MtAGarantir := RendCumulSalSess('05')
             else
               MtAGarantir := RendCumulSalSess('07') + RendCumulSalSess('08') + RendCumulSalSess('10');


         subrogation := (Tob_etab.GetValue('ETB_SUBROGATION') = 'X');

         if RecupereIJSS(StrToDate(Datef), Salarie, Mode, TOB_IJSS, Tob_Maintien, Tob_Etab) then
//FQ 11673 IntegreIJSSDansPaye(Tob_Rub, TobSal, StrToDate(DateD), StrToDate(DateF), Mode, ChampCateg, TOB_IJSS,RubIJSSNettes);

//      Le dernier paramètre (subrogation) doit être = True.
//      La rubrique d'ijss nette intégrée pour le calcul de la garantie est ensuite
//      supprimée s'il n'y a pas subrogation
           IntegreIJSSDansPaye(Tob_Rub, TobSal, StrToDate(DateD), StrToDate(DateF), Mode, ChampCateg, TOB_IJSS,RubIJSSNettes, True);
// PT33           IntegreIJSSDansPaye(Tob_Rub, TobSal, StrToDate(DateD), StrToDate(DateF), Mode, ChampCateg, TOB_IJSS,RubIJSSNettes, Subrogation);

         CalculBulletin(TOB_Rub, True);

         PlusApprochant := 0;

         Sauv_PGEnversNet := VH_Paie.PGEnversNet;
         VH_Paie.PGEnversNet := CumulAGarantir;
         RetEnvers := CalculPEnvers(PlusApprochant, TOB_Rub, TLig, MtAGarantir, LaRubriqGar, TRUE, FALSE, TRUE);

         //  On relance le calcul avec arrondi si echec
         if (NOT RetEnvers) AND (VH_PAIE.PGArrondiEnvers >0) then
           RetEnvers := CalculPEnvers(PlusApprochant, TOB_Rub, TLig, MtAGarantir, LaRubriqGar, TRUE, TRUE, TRUE);

         VH_Paie.PGEnversNet := Sauv_PGEnversNet;

         if not RetEnvers then
         begin
           Anomalie := 2;
         end;

         MtGarantie := 0;
         Tlig := TOB_Rub.FindFirst(['PHB_NATURERUB', 'PHB_RUBRIQUE'], ['AAA', LaRubriqGar], FALSE);
         if TLig <> nil then
         begin
           MtGarantie := TLig.getValue('PHB_MTREM');
         end;
         Wess := MtGarantie;
       end; // fin if (TReglt <> nil)

       // Maj Maintien : création nouvelle ligne
       //========================================
       if (Tob_Maintien = nil) then
         Tob_Maintien := Tob.create('les maintiens du salarié',nil,-1);

       tMaintien := TOB.Create('MAINTIEN',Tob_Maintien, -1);
       NvelleLMaintien(Tob_ReglesMaintien,Tob_etab,Wcarence,TotAbsence,MtGarantie,MaintienNet,Tabsence,treglt,tmaintien,ACheval,WMtMaintien,NbRestant,DateDebutAbs,1,StrToDate(DateD), StrToDate(DateF));
       tMaintien.PutValue('PMT_RUBRIQUE',LaRubriqGar);
     end; // fin if (Tob_RegltIJSS <> nil) and (Tob_RegltIJSS.Detail.count > 0)

     // suppression de la ligne de garantie dans le bulletin
     Tlig := TOB_Rub.FindFirst(['PHB_NATURERUB', 'PHB_RUBRIQUE'], ['AAA', LaRubriq], FALSE);
     if (Tlig <> nil) then
       Tlig.free;
   end { fin if (Tob_AbsIJ <> nil)}
   else
   // Pas d'absence pour cette période, au moins un règlement d'IJSS
   begin
     Tob_RegltIJSS := RecupereRegltIJSS(StrToDate(Datef), Salarie, Mode,ChampCateg, Tob_Maintien, Tob_Etab);
     if (Tob_RegltIJSS <> nil) and (Tob_RegltIJSS.Detail.count > 0) then
     begin
       TReglt := Tob_RegltIJSS.Findfirst([''],[''],false);
       if (TReglt <> nil) then
       begin

         TypeAbs := treglt.getValue('PRI_TYPEABS');
        // recherche date début absence
           RechDebutAbs := True;
           DateDebutAbsReel := treglt.getValue('PRI_DATEDEBUTABS');
           DateFinAbsPrec := DateDebutAbsReel-1 ;
           while RechDebutAbs do
           begin
             st := 'SELECT PCN_DATEDEBUTABS FROM ABSENCESALARIE WHERE'+
                   ' PCN_SALARIE = "'+   Salarie+ '"'+
                   ' AND PCN_GESTIONIJSS ="X"'+
                   ' AND PCN_DATEFINABS ="'+UsDateTime(DateFinAbsPrec)+'"'+
                   ' AND PCN_TYPECONGE IN (SELECT PMA_MOTIFABSENCE FROM MOTIFABSENCE'+
                   ' WHERE ##PMA_PREDEFINI## AND PMA_TYPEABS="'+TypeAbs+'")';
             Q := opensql(st, True);
             if not Q.eof then
             begin
               DateDebutAbsReel := Q.Fields[0].AsDateTime;
               DateFinAbsPrec := DateDebutAbsReel-1;
             end
             else
               RechDebutAbs := False;
             ferme (Q);
           end;

           Anciennete := AncienneteMois (StrToDate(DateAnciennete), PlusDate(DateDebutAbsReel,1,'J'));

           // récup critères de maintien
           // --------------------------------
           st := 'SELECT * FROM CRITMAINTIEN WHERE ##PCM_PREDEFINI## (PCM_TYPEABS = "' +
                 TypeAbs+'" OR PCM_TYPEABS="") AND  PCM_BORNEFINANC >= '+
                 IntToStr(Anciennete)+' AND PCM_BORNEDEBANC <= '+IntToStr(Anciennete);

           if (Convention <> '') then
             st := st + ' AND (PCM_CONVENTION ="'+Convention+'" OR '+
                   'PCM_CONVENTION ="000")'
           else
             st := st + ' AND PCM_CONVENTION ="000"';

           if (ChampCateg<>'') then
             if (TReglt.GetValue(ChampCateg)= '') then
               st := st + ' AND (PCM_VALCATEGORIE = "" OR PCM_VALCATEGORIE="<<Tous>>")'
             else
               st := st + ' AND (PCM_VALCATEGORIE LIKE "%'+TReglt.GetValue(ChampCateg)+'%"'+
                          ' OR PCM_VALCATEGORIE="" OR PCM_VALCATEGORIE="<<Tous>>")';

           Q:=OpenSql(st, TRUE);

           if not Q.eof then
           begin
             PoidsMax := 0;
             Tob_CritMaintien := Tob.create('les tables de critère ',nil,-1);
             Tob_CritMaintien.LoadDetailDB('CRITMAINTIEN','','',Q,False);

             // plusieurs Critères de maintien peuvent correspondre à la sélection
             // il faut définir l'élément de la table CRITMAINTIEN optimum, on donne
             // donc à chaque élément de la tob un poids
             if (Tob_CritMaintien <> nil) and (Tob_CritMaintien.Detail.count <> 0) then
             begin
               tCritMaintien := Tob_CritMaintien.FindFirst([''],[''],false);
               for i:= 1 to Tob_CritMaintien.Detail.count do
               begin
                 Poids := 0;
                 // PREDEFINI
                 if (tCritMaintien.GetValue('PCM_PREDEFINI') = 'CEG') then
                   // CEG
                   Poids := Poids + 1
                 else
                   if (tCritMaintien.GetValue('PCM_PREDEFINI') = 'STD') then
                     // STD
                     Poids := Poids + 2
                   else
                     // DOS
                     Poids := Poids + 3;

                 // TYPEABS
                 if (tCritMaintien.GetValue('PCM_TYPEABS') = TypeAbs) then
                   Poids := Poids + 2
                 else
                   Poids := Poids +1;

                 // CONVENTION
                 if ((tCritMaintien.GetValue('PCM_CONVENTION') = Convention) or
                     ((tCritMaintien.GetValue('PCM_CONVENTION') = '000') and
                      (Convention = ''))) then
                   Poids := Poids + 2
                 else
                   Poids := Poids +1 ;

                 // VALCATEG
                 if ((tCritMaintien.GetValue('PCM_VALCATEGORIE') =
                     TReglt.GetValue(ChampCateg)) or
                     ((tCritMaintien.GetValue('PCM_VALCATEGORIE') = '<<Tous>>') and
                      (TReglt.GetValue(ChampCateg)=''))) then
                   Poids := Poids + 2
                 else
                   Poids := Poids +1;

                 TCritmaintien.AddChampSupValeur('POIDS', Poids, False);

                 if (Poids > PoidsMax) then
                 // c'est lélément de poids le + élévé qui sera retenu
                 begin
                   PoidsMax := Poids;
                   NoCrit := i;
                 end;
                 tCritMaintien := Tob_CritMaintien.FindNext([''],[''],false);
               end; {fin for i:= 1 to Tob_CritMaintien.Detail.count}
             end; {fin (Tob_CritMaintien <> nil) and (Tob_CritMaintien.Detail.count <> 0)}
           end
           else
           // pas de paramétrage CRITMAINTIEN correspondant aux critères
           begin
             Anomalie := 1;
             Ferme(Q);
             exit;
           end;

           tCritMaintien := Tob_CritMaintien.Detail[NoCrit-1];
           // Maintien du net Oui ou Non
           if (tCritMaintien.GetValue('PCM_MAINTIENNET') = 'X') then  
             MaintienNet := True
           else
             MaintienNet := False;

           // raz rub abs IJSS + rub. IJSS et maintien
           //-----------------------------------------

           // Absence IJSS
           QAb:=OpenSql('SELECT PCN_TYPECONGE FROM ABSENCESALARIE WHERE'+
                        ' PCN_SALARIE = "'+   Salarie+ '"'+
                        ' AND PCN_DATEVALIDITE <= "' + usdatetime(StrToDate(Datef))+'"' +
                        ' AND PCN_DATEPAIEMENT = "' + usdatetime(StrToDate(Datef))+'"'+
                        ' AND PCN_GESTIONIJSS ="X"',TRUE);
           While not QAb.eof do
           begin
             RechercheCarMotifAbsence(Qab.Fields[0].AsString,StProfilH,StRubH,StAlimH,StProfilJ,StRubJ,StAlimJ,AbgereCom,Abheure,NetH,HetJ);
             if NetH then
               LaRubriq := StRubH
             else
               LaRubriq := StRubJ;

             if  (LaRubriq <> '') then
             begin
               Tlig := TOB_Rub.FindFirst(['PHB_NATURERUB', 'PHB_RUBRIQUE'], ['AAA', LaRubriq], FALSE);

               // Remise à 0 la valeur de la ligne pour le calcul de la paie à l'envers
               if TLig <> nil then
               begin
                 TLig.putValue('PHB_MTREM', 0);
                 TLig.putValue('PHB_BASEREM', 0);
               end;
             end;
             Qab.Next;
           end;
           ferme(Qab);

           // IJSS réglées (IJSS nettes)
           LaRubriq := '';
           LaRubriq := tCritMaintien.GetValue('PCM_RUBIJSSNETTE'); 
           if  (LaRubriq <> '') then
           begin
             Tlig := TOB_Rub.FindFirst(['PHB_NATURERUB', 'PHB_RUBRIQUE'], ['AAA', LaRubriq], FALSE);
             // Remise à 0 la valeur de la ligne pour le calcul de la paie à l'envers
             if TLig <> nil then TLig.putValue('PHB_MTREM', 0);
           end;

           // IJSS brutes
           LaRubriq := '';
           LaRubriq := tCritMaintien.GetValue('PCM_RUBIJSSBRUTE'); 
           if  (LaRubriq <> '') then
           begin
             Tlig := TOB_Rub.FindFirst(['PHB_NATURERUB', 'PHB_RUBRIQUE'], ['AAA', LaRubriq], FALSE);
             // Remise à 0 la valeur de la ligne pour le calcul de la paie à l'envers
             if TLig <> nil then TLig.putValue('PHB_MTREM', 0);
           end;

           // Maintien
           LaRubriq := '';
           LaRubriq := tCritMaintien.GetValue('PCM_RUBGARANTIE');
           if (LaRubriq = '') then
           begin
             anomalie := 3;
           end;
           if  (LaRubriq <> '') then
           begin
             Tlig := TOB_Rub.FindFirst(['PHB_NATURERUB', 'PHB_RUBRIQUE'], ['AAA', LaRubriq], FALSE);
             // Remise à 0 la valeur de la ligne pour le calcul de la paie à l'envers
             if TLig <> nil then TLig.putValue('PHB_MTREM', 0);
           end;

           CalculBulletin(Tob_Rub, True);

           // Calcul du maintien net ou brut à appliquer
           // ------------------------------------------
           // On récupère le montant à garantir (net à payer, brut, brut habituel...en fct du paramétrage)
           if (tCritMaintien.GetValue('PCM_CUMULPAIE') = 'NET') then
             MtAGarantir := RendCumulSalSess('10')
           else
             if (tCritMaintien.GetValue('PCM_CUMULPAIE') = 'BRU') then
               MtAGarantir := RendCumulSalSess('01')
             else
               if (tCritMaintien.GetValue('PCM_CUMULPAIE') = 'BRH') then  
                 MtAGarantir := RendCumulSalSess('05')
               else
                 MtAGarantir := RendCumulSalSess('07') + RendCumulSalSess('08') + RendCumulSalSess('10');

           LaRubriq := '';
           if (MtAGarantir <> 0) then
             // récupération de la Rubrique de garantie du maintien
             LaRubriq :=  tCritMaintien.GetValue('PCM_RUBGARANTIE');

           if (MtAGarantir <> 0) and (LaRubriq <> '') then
           begin
             Tlig := TOB_Rub.FindFirst(['PHB_NATURERUB', 'PHB_RUBRIQUE'], ['AAA', LaRubriq], FALSE);
             if TLig = nil then
             begin
               // on insere la rubrique dans le bulletin
               ChargeProfilSPR(TobSal, TOB_Rub, LaRubriq, 'AAA');
               Tlig := TOB_Rub.FindFirst(['PHB_NATURERUB', 'PHB_RUBRIQUE'], ['AAA', LaRubriq], FALSE);
             end;
             // Remise à 0 la valeur de la ligne pour le calcul de la paie à l'envers
             // modification du libellé (régul garantie sur net)
             if TLig <> nil then
             begin
               TLig.putValue('PHB_MTREM', 0);
               TLig.PutValue('PHB_LIBELLE','Régularisation garantie sur net');
             end;
           end;
           // récupération des absence de type IJSS
           // (celles qui ont été suprimées auparavant dans le bulletin)
           IJSS := True;
{PT35
           RecupereAbsences(Tob_Abs,StrToDate(Datef), Salarie,Mode, TRUE, IJSS);}
           RecupereAbsences(Tob_Abs,StrToDate(Datef), Salarie,Mode, TRUE, IJSS, '');
           if (Tob_Abs <> nil) and (Tob_Abs.Detail.count >= 0) then
             IntegreAbsenceDansPaye(Tob_Rub, Tob_Abs, TobSal, StrToDate(DateD), StrToDate(DateF), Mode);

           subrogation := (Tob_etab.GetValue('ETB_SUBROGATION') = 'X'); // PT32

           // intégration reglt d'IJSS
           if RecupereIJSS(StrToDate(Datef), Salarie, Mode, TOB_IJSS, Tob_Maintien, Tob_Etab) then
//FQ 11673 IntegreIJSSDansPaye(Tob_Rub, TobSal, StrToDate(DateD), StrToDate(DateF), Mode, ChampCateg,TOB_IJSS,RubIJSSNettes);

//      Le dernier paramètre (subrogation) doit être = True.
//      La rubrique d'ijss nette intégrée pour le calcul de la garantie est ensuite
//      supprimée s'il n'y a pas subrogation

             IntegreIJSSDansPaye(Tob_Rub, TobSal, StrToDate(DateD), StrToDate(DateF), Mode, ChampCateg,TOB_IJSS,RubIJSSNettes,True);
//PT33             IntegreIJSSDansPaye(Tob_Rub, TobSal, StrToDate(DateD), StrToDate(DateF), Mode, ChampCateg,TOB_IJSS,RubIJSSNettes,Subrogation);

           CalculBulletin(TOB_Rub, True);

           RetEnvers := TRUE;
           if (MtAGarantir <> 0) and (LaRubriq <> '') then
           begin
             if TLig <> nil then
             begin
               PlusApprochant := 0;

               Sauv_PGEnversNet := VH_Paie.PGEnversNet;
               VH_Paie.PGEnversNet := tCritMaintien.GetValue('PCM_CUMULPAIE'); 
               RetEnvers := CalculPEnvers(PlusApprochant, TOB_Rub, TLig, MtAGarantir, LaRubriq, TRUE, FALSE, TRUE);

               //  On relance le calcul avec arrondi si echec
               if (NOT RetEnvers) AND (VH_PAIE.PGArrondiEnvers >0) then
                 RetEnvers := CalculPEnvers(PlusApprochant, TOB_Rub, TLig, MtAGarantir, LaRubriq, TRUE, TRUE, TRUE);

               VH_Paie.PGEnversNet := Sauv_PGEnversNet;

               if not RetEnvers then
               begin
                 Anomalie := 2;
               end;
               Tlig := TOB_Rub.FindFirst(['PHB_NATURERUB', 'PHB_RUBRIQUE'], ['AAA', LaRubriq], FALSE);
               if TLig <> nil then
               begin
                 MtGarantie := TLig.getValue('PHB_MTREM');
               end;

             end;
           end;
       end; { fin de if (TReglt <> nil)}

         // Maj Maintien : création nouvelle ligne
         //========================================
         if (Tob_Maintien = nil) then
           Tob_Maintien := Tob.create('les maintiens du salarié',nil,-1);

         tMaintien := TOB.Create('MAINTIEN',Tob_Maintien, -1);

         NvelleLMaintien(Tob_ReglesMaintien,Tob_etab,Wcarence,TotAbsence,MtGarantie,MaintienNet,Tabsence,treglt,tmaintien,ACheval,WMtMaintien,NbRestant,DateDebutAbs,2,StrToDate(DateD), StrToDate(DateF));

         tMaintien.PutValue('PMT_RUBRIQUE',LaRubriq);

         // suppression de la ligne de garantie dans le bulletin
         Tlig := TOB_Rub.FindFirst(['PHB_NATURERUB', 'PHB_RUBRIQUE'], ['AAA', LaRubriq], FALSE);
         if (Tlig <> nil) then
           Tlig.free;
     end; // fin (Tob_RegltIJSS <> nil) and (Tob_RegltIJSS.Detail.count > 0)
   end; // Fin Pas d'absence pour cette période, au moins un règlement d'IJSS

   // mise à jour de la TOB Final du maintien qui sera utilsée pour la
   // maj de la table MAINTIEN
   // =================================================================

   if (Tob_Maintien <> Nil) then
     MajMaintien(Tob_Maintien);

   FreeAndNil (Tob_AbsIJ);
   if (Tob_RegltIJSS <> NIL) then
   begin
     FreeAndNil(Tob_RegltIJSS) ; // XP 10.10.2007 .free;
     // XP 10.10.2007 Tob_RegltIJSS := NIL ;
   end;

   if RecupereMaintien(StrToDate(Datef), Salarie, Mode,TOB_Maintien) then
     IntegreMaintienDansPaye(Tob_Rub, TobSal, StrToDate(DateD), StrToDate(DateF), Mode, false,ChampCateg, TOB_Maintien);

   subrogation := (Tob_etab.GetValue('ETB_SUBROGATION') = 'X');

   if not subrogation then
   // L'employeur ne pratique pas la subrogation
   // suppression de la ligne IJSS nette dans le bulletin et des ligne de commentaire
   begin
     EnleveCommIJ(TobSal, Tob_Rub, RubIJSSNettes, 'AAA', StrToDate(DateD), StrToDate(DateF));

     Tlig := TOB_Rub.FindFirst(['PHB_NATURERUB', 'PHB_RUBRIQUE'], ['AAA', RubIJSSNettes], FALSE);
     if (Tlig <> nil) then
       Tlig.free;
   end;
   FreeAndNil(Tob_CritMaintien);
end; { fin MaintienDuSalarie }

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 21/10/2004
Modifié le ... :   /  /
Description .. : procédure MajMaintien
Suite ........ : mise à jour de la table MAINTIEN
Suite ........ : Elle se fait au fur et à mesure des création de lignes de
Suite ........ : maintien
Mots clefs ... : PAIE; IJSS; MAINTIEN
*****************************************************************}
procedure MajMaintien (var Tob_Maintien : TOB);
var
  tMaintien,tMajMaintien,TOB_MAJMAINTIEN                : TOB;
  PremierMaintien                                       : boolean;
begin
 PremierMaintien := True;
 tMaintien:= Tob_Maintien.Findfirst([''],[''],false);
 while (TMaintien <> nil) do
 begin
   if (TOB_MAJMAINTIEN = NIL) then
     TOB_MAJMAINTIEN := Tob.create('Les maintiens de tous les salariés',nil,-1);

   tMajMaintien := TOB.Create('MAINTIEN',TOB_MAJMAINTIEN, -1);
   tMajMaintien.PutValue('PMT_SALARIE',tMaintien.GetValue('PMT_SALARIE'));
   tMajMaintien.PutValue('PMT_DATEDEBUTABS',tMaintien.GetValue('PMT_DATEDEBUTABS'));
   tMajMaintien.PutValue('PMT_DATEFINABS',tMaintien.GetValue('PMT_DATEFINABS'));
   tMajMaintien.PutValue('PMT_DATEDEBUT',tMaintien.GetValue('PMT_DATEDEBUT'));
   tMajMaintien.PutValue('PMT_DATEFIN',tMaintien.GetValue('PMT_DATEFIN'));
   tMajMaintien.PutValue('PMT_TYPECONGE',tMaintien.GetValue('PMT_TYPECONGE'));
   tMajMaintien.PutValue('PMT_LIBELLE',tMaintien.GetValue('PMT_LIBELLE'));
   tMajMaintien.PutValue('PMT_BASEMAINTIEN',tMaintien.GetValue('PMT_BASEMAINTIEN'));
   tMajMaintien.PutValue('PMT_PCTMAINTIEN',tMaintien.GetValue('PMT_PCTMAINTIEN'));
   tMajMaintien.PutValue('PMT_CARENCE',tMaintien.GetValue('PMT_CARENCE'));
   tMajMaintien.PutValue('PMT_NBJMAINTIEN',tMaintien.GetValue('PMT_NBJMAINTIEN'));
   // 2 lignes de maintien pour le même mois - on ne conserve le mt que pour la 1ère ligne
   // sinon mt doublé dans bulletin
   if (tMaintien.GetValue('PMT_DATEDEBUT') = Idate1900) and
      (tMaintien.GetValue('PMT_DATEFIN') = IDate1900) and
      not PremierMaintien then
   begin
     tMajMaintien.PutValue('PMT_MTMAINTIEN',tMaintien.GetValue('PMT_MTMAINTIEN'));
     PremierMaintien := False;
   end
   else
   begin
     if (tMaintien.GetValue('PMT_DATEDEBUT') = Idate1900) and
        (tMaintien.GetValue('PMT_DATEFIN') = IDate1900) then
       PremierMaintien := False;
     tMajMaintien.PutValue('PMT_MTMAINTIEN',tMaintien.GetValue('PMT_MTMAINTIEN'));
   end;

   tMajMaintien.PutValue('PMT_BASEABS',tMaintien.GetValue('PMT_BASEABS'));
   tMajMaintien.PutValue('PMT_NBJABS',tMaintien.GetValue('PMT_NBJABS'));
   tMajMaintien.PutValue('PMT_NBHEURES',tMaintien.GetValue('PMT_NBHEURES'));
   tMajMaintien.PutValue('PMT_RUBRIQUE',tMaintien.GetValue('PMT_RUBRIQUE'));
   tMajMaintien.PutValue('PMT_TYPEMAINTIEN',tMaintien.GetValue('PMT_TYPEMAINTIEN'));
   tMajMaintien.PutValue('PMT_HISTOMAINT',tMaintien.GetValue('PMT_HISTOMAINT'));

   tMaintien:= Tob_Maintien.FindNext([''],[''],false);
 end; {while (TMaintien <> nil) }
 FreeAndNil (Tob_Maintien);

 if (TOB_MAJMAINTIEN <> NIL) then
 begin

 // Mise à jour de la table MAINTIEN
 // ================================
   try
     BeginTrans;
     TOB_MAJMAINTIEN.SetAllModifie (TRUE);
     TOB_MAJMAINTIEN.InsertOrUpdateDB (TRUE);
     CommitTrans;
   except
     Rollback;
   end;
 end;
 FreeAndNil(TOB_MAJMAINTIEN);
end;

// d FQ 11673
{***********UNITE*************************************************
Auteur  ...... : PAIE -MF
Créé le ...... : 21/10/2003
Modifié le ... :   /  /
Description .. : procédure RecupIJSS
Suite ........ : Intégration des IJSS dans le bulletin dans le cas où on ne
Suite ........ : gère que les IJSS et pas le maintien
Mots clefs ... : PAIE; IJSS; MAINTIEN
*****************************************************************}
procedure RecupIJSS(const Dated, Datef:TDateTime;const Mode : string; Tob_Rub : TOB;var Tob_IJSS : TOB; const Tob_Etablissement,Tob_Salarie: TOB;var ChampCateg,RubIJNettes: String; Const Subrogation : boolean);
var
  RubNettes,Salarie                             : string;
begin
  Salarie := Tob_Salarie.GetValue('PSA_SALARIE');
  if RecupereIJSS(Datef, Salarie, Mode, TOB_IJSS, nil,nil) then
    IntegreIJSSDansPaye(Tob_Rub, Tob_Salarie, DateD, DateF, Mode, ChampCateg, TOB_IJSS,RubNettes, Subrogation);
//f FQ 11673
end;

{***********UNITE*************************************************
Auteur  ...... : PAIE- MF
Créé le ...... : 21/10/2004
Modifié le ... :   /  /
Description .. : procédure NvelleLigneMaintien
Suite ........ : Calcul du maintien et alimentation des champs
Mots clefs ... : PAIE; IJSS; MAINTIEN
*****************************************************************}
procedure NvelleLMaintien(const Tob_ReglesMaintien, Tob_etab : TOB;const Wcarence, TotAbsence : double; const MtGarantie : double;const MaintienNet : boolean;const Tabsence,treglt : TOB;var tmaintien : TOB;var ACheval : boolean;var WMtMaintien : double;var NbRestant : double;var DateDebutAbs : TDateTime;const uniqtreglt : integer;dated,datef : TDateTime);
var
  NbJoursCalendairesAbsence, NbBase         : double;
  Reglesuivante, FinRegles                  : boolean;
begin
  Reglesuivante := false;
  FinRegles := False;
//  if not(Uniqtreglt) then
  if (Uniqtreglt = 0) then
  begin
    tMaintien.PutValue('PMT_SALARIE',tabsence.GetValue('PCN_SALARIE'));
    tMaintien.PutValue('PMT_TYPEMAINTIEN','MAI');

    if (DateDebutAbs = IDate1900) then
    begin
      tMaintien.PutValue('PMT_DATEDEBUTABS',tabsence.GetValue('PCN_DATEDEBUTABS'));
      DateDebutAbs :=  tabsence.GetValue('PCN_DATEDEBUTABS');
    end
    else
      // il s'agit d'une absence découpée, la date de début = date fin précédente
      // calculée + 1 jour
      tMaintien.PutValue('PMT_DATEDEBUTABS',DateDebutAbs);


    tMaintien.PutValue('PMT_DATEFINABS',tabsence.GetValue('PCN_DATEFINABS'));
    tMaintien.PutValue('PMT_DATEDEBUT',Idate1900);
    tMaintien.PutValue('PMT_DATEFIN',Idate1900);
    tMaintien.PutValue('PMT_TYPECONGE',tabsence.GetValue('PCN_TYPECONGE'));
    tMaintien.PutValue('PMT_LIBELLE',tabsence.GetValue('PCN_LIBELLE'));
    NbJoursCalendairesAbsence := tabsence.GetValue('PCN_NBJCALEND');
    NbBase := NbJoursCalendairesAbsence;

    if (NbRestant <> 0) then
    // cas d'une absence découpée
      NbJoursCalendairesAbsence := NbRestant;

    if (RegleNo < Tob_ReglesMaintien.Detail.count) then
    //
    begin
      if ((TbNbj[RegleNo] = TbNbjMax[RegleNo]) and (TbNbjMax[RegleNo] <> 0)) or
          ((MaintConsec) and (MaintienAZero)) then
      // le maxi pour la règle n° RegleNo est déjà atteint ou bien le maintien est
      // forcément à 0 (dates consécutives et dernier historique maintien = 0)
      begin
        ACheval := True;  // pour passer à la règle suivante pour la même absence
        Reglesuivante := True;
      end
      else
      begin
        if (NbJoursCalendairesAbsence + TbNbj[RegleNo] -
            (TbCarence[RegleNo]-Wcarence) <= TbNbjMax[RegleNo]) or
            (TbNbjMax[RegleNo]= 0) then
        // le nbre de jours calendaires d'absence correspond à une seule règle
        // on tient comptes du maintien des 12 mois précédents l'absence
        begin
          if (TbNbjMax[RegleNo] <> 0) then
          begin
            tMaintien.PutValue('PMT_NBJMAINTIEN', NbJoursCalendairesAbsence-
                                                 (TbCarence[RegleNo]-Wcarence));
            tMaintien.PutValue('PMT_CARENCE',(TbCarence[RegleNo]-Wcarence));
            if (tMaintien.GetValue('PMT_NBJMAINTIEN') < 0 )  then
            begin
              tMaintien.PutValue('PMT_NBJMAINTIEN',0);
              tMaintien.PutValue('PMT_CARENCE',NbJoursCalendairesAbsence);
            end;
          end
          else
          begin
            tMaintien.PutValue('PMT_NBJMAINTIEN',0);
            tMaintien.PutValue('PMT_CARENCE',0);
            tMaintien.PutValue('PMT_PCTMAINTIEN',0);
          end;
          ACheval := false;
          TbNbj[regleNo] := TbNbj[regleNo] +
                            tMaintien.GetValue('PMT_NBJMAINTIEN');
          NbRestant := 0;
        end
        else
        // le nbre de jours calendaires d'absence est à cheval sur pls règles
        // on tient comptes du maintien des 12 mois précédents l'absence
        begin
          NbRestant := (NbJoursCalendairesAbsence + TbNbj[RegleNo] -
                      (TbCarence[RegleNo]-Wcarence))-TbNbjMax[RegleNo];
          tMaintien.PutValue('PMT_NBJMAINTIEN',NbJoursCalendairesAbsence-
                                               (TbCarence[RegleNo]-Wcarence)-
                                               NbRestant);
          tMaintien.PutValue('PMT_CARENCE',(TbCarence[RegleNo]-Wcarence));
            if (tmaintien.GetValue('PMT_NBJMAINTIEN')< 0 )  then
            begin
              tMaintien.PutValue('PMT_NBJMAINTIEN',0);
              tMaintien.PutValue('PMT_CARENCE', NbJoursCalendairesAbsence);
            end;

          ACheval := true;
          TbNbj[regleNo] := TbNbj[regleNo] +
                           tMaintien.GetValue('PMT_NBJMAINTIEN');
        end;
      end;
    end
    else
    begin
      FinRegles := True;
      tMaintien.PutValue('PMT_NBJMAINTIEN',NbJoursCalendairesAbsence);
      tMaintien.PutValue('PMT_CARENCE',0);
      if (tmaintien.GetValue('PMT_NBJMAINTIEN')< 0 )  then
        tMaintien.PutValue('PMT_NBJMAINTIEN',0);
      tMaintien.PutValue('PMT_PCTMAINTIEN',0);
      ACheval := false;
    end;
    if (Reglesuivante = false) then
    begin
      // la date de fin d'absence est limitée en fct du nbre jours maintien.
      if (tMaintien.GetValue('PMT_CARENCE')+
          tMaintien.GetValue('PMT_NBJMAINTIEN') <> 0) and
         (tMaintien.GetValue('PMT_NBJMAINTIEN') <> 0) then

        tMaintien.PutValue('PMT_DATEFINABS',
                          PlusDate(DateDebutAbs,
                          (tMaintien.GetValue('PMT_CARENCE')+
                          tMaintien.GetValue('PMT_NBJMAINTIEN')- 1)
                          , 'J'));
      DateDebutAbs := tMaintien.GetValue('PMT_DATEFINABS')+1 ;
      if (tMaintien.GetValue('PMT_CARENCE')+
          tMaintien.GetValue('PMT_NBJMAINTIEN') <> 0) and
         (tMaintien.GetValue('PMT_NBJMAINTIEN') <> 0) then
        tMaintien.PutValue('PMT_DATEDEBUTABS',
                         PlusDate(tMaintien.GetValue('PMT_DATEFINABS'),
                         (tMaintien.GetValue('PMT_NBJMAINTIEN')-1)*-1,'J'));

      if not FinRegles then
        tMaintien.PutValue('PMT_PCTMAINTIEN',TbPct[RegleNo]);
      tMaintien.PutValue('PMT_BASEABS',tabsence.GetValue('PCN_BASE'));
      tMaintien.PutValue('PMT_NBJABS',tabsence.GetValue('PCN_JOURS'));
      tMaintien.PutValue('PMT_NBHEURES',tabsence.GetValue('PCN_HEURES'));
      tMaintien.PutValue('PMT_BASEMAINTIEN',((MtGarantie/TotAbsence)*NbBase));

      // Calcul du montant du maintien
      // -----------------------------
      //      Montant maintien =
      //     (Montant de la garantie/nb jour total d'absence du mois) *
      //     (nb jours de maintien) *
      //     (tx du maintien)
      tMaintien.PutValue('PMT_MTMAINTIEN',
                         ((MtGarantie/TotAbsence)
                         *tMaintien.GetValue('PMT_NBJMAINTIEN'))
                         *(tMaintien.GetValue('PMT_PCTMAINTIEN')/100));
      WMtmaintien := WMtMaintien + tMaintien.GetValue('PMT_MTMAINTIEN');

      tMaintien.PutValue('PMT_RUBRIQUE',tabsence.GetValue('RUBMAINTIEN'));
    end;
    if (ACheval = True) then
    begin
      if (Reglesuivante) then
        tmaintien.free;
      RegleNo := RegleNo+1;
    end;

    Reglesuivante := False;

  end {if not(Uniqtreglt)}
  else
  //(uniqtReglt=true) - Traitement de la garantie du salaire
  begin
      tMaintien.PutValue('PMT_SALARIE',treglt.GetValue('PRI_SALARIE'));
      tMaintien.PutValue('PMT_DATEDEBUTABS',Dated);
      tMaintien.PutValue('PMT_DATEFINABS',Datef);

      // pour parer à une clé duppliquée
      // Il n'est plus nécessaire de modifier les dates (ajout de PMT_TYPEMAINTIEN à la clé)

      tMaintien.PutValue('PMT_DATEDEBUT',IDate1900);
      tMaintien.PutValue('PMT_DATEFIN',IDate1900);
      if (MaintienNet) then
        tMaintien.PutValue('PMT_TYPECONGE','NET')
      else
        tMaintien.PutValue('PMT_TYPECONGE','BRU');
      tMaintien.PutValue('PMT_LIBELLE','Garantie du salaire');
      tMaintien.PutValue('PMT_BASEMAINTIEN',MtGarantie);
      tMaintien.PutValue('PMT_PCTMAINTIEN',0);
// Dans le cas garantie du salaire le montant de la garantie est toujours à 100%
      tMaintien.PutValue('PMT_MTMAINTIEN',MtGarantie);
      tMaintien.PutValue('PMT_CARENCE',0);
      tMaintien.PutValue('PMT_TYPEMAINTIEN','GAR');
      tMaintien.PutValue('PMT_NBJMAINTIEN',0);
      tMaintien.PutValue('PMT_BASEABS',0);
      tMaintien.PutValue('PMT_NBJABS',0);
      tMaintien.PutValue('PMT_NBHEURES',0);
      tMaintien.PutValue('PMT_RUBRIQUE','');
  end;

end; {fin  NvelleLMaintien}

end.


