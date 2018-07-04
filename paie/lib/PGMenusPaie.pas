{***********UNITE*************************************************
Auteur  ...... : JL
Créé le ...... : 09/09/2004
Modifié le ... :   /  /
Description .. : Menu formation,primes
Mots clefs ... : PAIE
*****************************************************************}
{
PT1   : 25/07/2005 SB V65  FQ 12399 Ajout Menu des assistantes EConges
PT2   : 15/09/2005 SB V65  Econges : recherche responsable table intérimaire
PT3   : 28/02/2006 SB V70  Optimisation des traitement , ajout rupture resp. abs
                           sur planning du jour
PT4   : 30/03/2006 MF V65  FQ 12585 : affichage étalissements ou << Tous >>
PT5   : 19/06/2006 SB V_65 Refonte appel ligne de menu du module Absence
PT6   : 27/06/2006 SB V_65 Planning du jour sur une semaine
PT7   : 04/07/2006 SB V_65 Nouveaux paramètres d'execution pour Suivi des
                           absences
PT8   : 07/11/2006 SB V_70 Nouvel ligne de menu export etemptation
PT9   : 13/12/2006 SB V_70 Nouveau module : Gestion des présences
PT10  : 20/04/2007 FC V_72 Rajout de la fonction @CHOIXMONTANT pour la synthèse
                           des éléments nationaux
PT11  : 23/05/2007 FL V_72 FQ 14051 Ajout de la fonction LIBEMPLOI pour obtenir
                           les libellés emplois associés au code (dû à un pb
                           dans l'agl)
PT12  : 31/05/2007 JL V_72 FQ 14021 Ajout apramsoc pour gestion intervenants
                           exterieurs
PT13  : 11/06/2007 VG V_72 Adaptation nouvelles méthodes CBPPath
PT14  : 27/09/2007 NA V_80 Ajout au menu paramètres des présences la gestion des regroupements
PT15  : 22/10/2007 NA V_80 Ajout au menu de présence : Intégration en paie
PT16  : 16/11/2007 FC V_80 FQ 14931 Pb filtrage des valeurs d'un élément quand plusieurs prédéfinis paramétrés
PT17  : 18/12/2007 FC V_81 FQ 15065 Pb libellé qualification, niveau, indice et coefficient
PT18  : 28/09/2007 V_7 FL EManager / Report / Ajout cursus + planning
PT19  : 05/10/2007 V_7 FL Emanager / Report / Désactivation des menus d'édition collectif et individuel réalisé
PT20  : 17/10/2007 V_7 FL Report / Fermeture d'une requête
PT21  : 13/02/2008 V_803 FL Changement de certaines tablettes si multidossier formation
PT23  : 03/03/2008 FL V_8  EManager / Ajout du module Salariés (916) et de la consultation des primes
PT23  : 03/04/2008 FL V_8 Inversion des tests dans les changements de tablette dans le cas du partage formation
PT24  : 17/04/2008 GGU V81 FQ 15361 Factorisation du code en vue de la modification
PT25  : 21/04/2008 GGU V81 FQ 15361 Gestion uniformisée des zones libres - tables dynamiques
PT24  : 22/04/2008 FL V_8 Ajout de la tablette PGRESPONSFOR
PT28  : 07/08/2008 NA V_8 FQ 15493 Ne pas faire le remplissage à 0 si le matricule n'est pas numérique
}
unit PGMenusPaie;

interface
uses StdCtrls,
  Controls,
  Classes,
  forms,
  sysutils,
  ComCtrls,
  HCtrls,
  HPanel,
  HEnt1,
  UTofPGMulDetailHisto,
  UTofPGSaisielisteHistorique,
  UtofPGAnalayseZonesLibres,
  ParamSOc,
  UTOB,
{$IFDEF EAGLCLIENT}
  eTablette, MainEAGL, MenuOLX,
{$ELSE}
  Tablette, FE_Main, MenuOLG, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}DB,
{$ENDIF}
  HMsgBox,
  PGRepertoire;

function Paie_Menus(Num: Integer; PRien: THPanel) : boolean ;
function EFormation_menus(Num: Integer; PRien: THPanel) : boolean ;
function EPrimes_menus(Num: Integer; PRien: THPanel) : boolean ;
function EAugmentation_menus(Num: Integer; PRien: THPanel) : boolean ;
function ESalarie_menus(Num: Integer; PRien: THPanel) : boolean ; //PT23
function PgActivAbsences: boolean;
function PgActivFormation: boolean;
function PgActivPrimes: boolean;
function PgActivAugmentation: boolean;
function PgActivSalarie: boolean; //PT23
procedure ChangeMenuFormation(LeSalarie : String);
procedure ChangeMenuAugmentation(LeSalarie : String);
procedure ChangeMenuAbsence(LeSalarie : String);
procedure ChangeMenuSalarie(LeSalarie : String); //PT23
procedure Et_RetireModule(const NumModule: integer);
procedure Et_AjouteModule(const NumModule: integer);
function  PGCalcMul(Func,Params,WhereSQL : hstring ; TT : TDataset ; Total : Boolean) : hstring ;
procedure PGChangeTabletteResponsable;
function  Presence_Menus(Num: Integer; PRien: THPanel) : boolean ;  { PT9 }

var
Images: array[0..4] of integer;

implementation

uses Entpaie, P5Util,PgPlanning,PgPlanningOutils, PGOutilsFormation,
  PGTablesDyna; //PT24

function Paie_Menus(Num: Integer; PRien: THPanel) : boolean ;
var
NiveauRupt : TNiveauRupture;
begin
  // XP 16-09-2004 : Pour virer les droits
  if Num = -1 then
  begin
    result := PgActivAbsences() ;
    Exit ;
  end else result := True ;

  PGChangeTabletteResponsable;
  // Menu 47 FORMATION
  // MENU 43 ECongés ou eSalarié
  case Num of
    { Menu Salariés }
    43150: AglLanceFiche('PAY', 'EABSENCE_MUL', '', '', 'SAL');
    43160: AglLanceFiche('PAY', 'RECAPSALCP', '', LeSalarie, '');

    { Menu Assistantes }
    43215: AglLanceFiche('PAY', 'EABSENCEASS_MUL', '', '', 'ASS');                          //PT1
    43165: AglLanceFiche('PAY', 'SUPERVABSENCES', '', '', 'ASS;' + LeSalarie + ';PLANNING'); //PT1 planning des absences
    43250: AglLanceFiche('PAY', 'MULRECAPSAL', '', '', 'ASS');                              //PT1 Recapitulatif des absences
    43260: Begin  { DEB PT3 }
           NiveauRupt.NiveauRupt := 1;
           NiveauRupt.ChampsRupt[1] := 'PSE_RESPONSABS';
           PGPlanningAbsence(V_Pgi.DateEntree,(V_Pgi.DateEntree + 7),'ASS','','','',LeSalarie,NiveauRupt,True,'ABS'); { PT6 }
           End;   { FIN PT3 }
    43270: AglLanceFiche('PAY', 'SUPERVABSENCES', '', '', 'ASS;;SUIVI');    { PT7 }

    { Menu Responsable }
    //43310: AglLanceFiche('PAY', 'EDITBuL_ETAT', '', '', '');
    43710: AglLanceFiche('PAY', 'EABSENCER_MUL', '', '', 'RESP');
    43720: AglLanceFiche('PAY', 'MULRECAPSAL', '', '', 'RESP');
    43730: AglLanceFiche('PAY', 'SUPERVABSENCES', '', '', 'RESP;' + LeSalarie + ';PLANNING'); // planning des absences
    43735: Begin   { DEB PT3 }
           NiveauRupt.NiveauRupt := 1;
           NiveauRupt.ChampsRupt[1] := 'PSE_RESPONSABS';
           PGPlanningAbsence(V_Pgi.DateEntree,(V_Pgi.DateEntree + 7),'RESP','','','',LeSalarie,NiveauRupt,True,'ABS'); { PT6 }
           End;    { FIN PT3 }
    43770: AglLanceFiche('PAY', 'SUPERVABSENCES', '', '', 'RESP;;SUIVI');   { PT7 }

    { Menu Administrateur }
    43810: AglLanceFiche('PAY', 'ELISTETABLES', '', '', 'I');
    43820: AglLanceFiche('PAY', 'EXPORTABSENCES', '', '', '');
    43825: AglLanceFiche('PAY', 'EXPORTETEMP', '', '', ''); { PT8 }
    43840: AglLanceFiche('PAY', 'EABSENCER_MUL', '', '', 'ADM');
    43850: AglLanceFiche('PAY', 'MULRECAPSAL', '', '', 'ADM');
    43855: AglLanceFiche('PAY', 'SUPERVABSENCES', '', '', 'ADM;;PLANNING'); // planning des absences
    //43860: AglLanceFiche('PAY', 'ADMINEAGL', '', '', '');
    //43865: AglLanceFiche('PAY', 'EABSENCE_MUL', '', '', 'SAL');
    43857: Begin     { DEB PT3 }
           NiveauRupt.NiveauRupt := 1;
           NiveauRupt.ChampsRupt[1] := 'PSE_RESPONSABS';
           PGPlanningAbsence(V_Pgi.DateEntree,(V_Pgi.DateEntree + 7),'ADM','','','','',NiveauRupt,False,'ABS'); { PT6 }
           End;      { FIN PT3 }
    43870: AglLanceFiche('PAY', 'SUPERVABSENCES', '', '', 'ADM;;SUIVI');    { PT7 }

    //EPRIMES
    40110: AglLanceFiche('PAY', 'MUL_SAISPRIM', '', '', 'P;X'); // Saisie des primes
    40114: AglLanceFiche('PAY', 'MUL_SAISPRIM', '', '', 'P'); // Saisie des primes
    40115: AglLanceFiche('PAY', 'MUL_SAISPRIM', '', '', 'C'); // Consultation des primes //PT23
    40120: AglLanceFiche('PAY', 'PGMULHISTOSAL', '', '', 'P'); // ANAL_HISTPRIMconsultation historique des primes

    //EFORMATION
    47401 : AglLanceFiche('PAY', 'MUL_STAGECEGID', 'GRILLE=INSCBUDGET', '', 'CWASINSCBUDGET'); // inscriptions a partir des formations
    47402 : AglLanceFiche('PAY', 'MUL_CURSUS', '', '', '');
    47403 : AglLanceFiche('PAY', 'MUL_INSCFOR', '', '', 'CONSULTATION');
    47405 : AglLanceFiche('PAY', 'MUL_INSCFOR', '', '', 'VALIDATION');
    47408 : AglLanceFiche('PAY', 'EDITSUIVIBUDFORM', '', '', '');
    47409 : AglLanceFiche('PAY', 'EDITCOLFORMATION', '', '', '47521');
  end;
end;

function EFormation_menus(Num: Integer; PRien: THPanel) : boolean ;
begin
  // XP 16-09-2004 : Pour virer les droits
  if Num = -1 then
  begin
    Et_RetireModule(914);
    result := False ;
    Exit ;
  end else result := True ;

  case Num of
    //PT18 - Début
    {47401 : AglLanceFiche('PAY', 'MUL_STAGECEGID', 'GRILLE=INSCBUDGET', '', 'CWASINSCBUDGET'); // inscriptions a partir des formations
    47402 : AglLanceFiche('PAY', 'MUL_CURSUS', '', '', '');
    47403 : AglLanceFiche('PAY', 'EM_MULINSCFOR', '', '', 'CONSULTATION');
    47405 : AglLanceFiche('PAY', 'EM_VALIDINSCFOR', '', '', 'VALIDATION');
    47408 : AglLanceFiche('PAY', 'EDITSUIVIBUDFORM', '', '', '');
    47409 : AglLanceFiche('PAY', 'EDITCOLFORMATION', '', '', '47521');
    47333: AglLanceFiche('PAY', 'EM_MULFORMATIONS', '', '', 'CWASCONSULTATION');
     //dif
    47311 : AglLanceFiche('PAY', 'EM_MULCOMPTDIF', '', '', '');
    47312: AglLanceFiche('PAY', 'EM_MULDEMANDEDIF', '', '', 'DEMANDEDIF');
    47314 : AglLanceFiche('PAY', 'EDITRECAPDIF', '', '', '');
    47315 : AglLanceFiche('PAY', 'EDITCOMPTEURSDIF', '', '', '');
    47180 : AglLanceFiche('PAY', 'MULPARAMFORMABS', '', '', '');
    47317 : AglLanceFiche('PAY', 'EDITDEMANDEDIF', '', '', '');
    47318 : AglLanceFiche('PAY', 'EDITMVTSDIF', '', '', '');}

    // Réalisé
    47301 : AglLanceFiche('PAY', 'EM_MULFORMATIONS', '', '', '');
    47302 : AglLanceFiche('PAY', 'EM_EDITREA', '', '', 'COLLECTIF');
    47303 : AglLanceFiche('PAY', 'EM_EDITREA', '', '', 'INDIVIDUEL');
    47304 : AglLanceFiche('PAY', 'EM_EDITBILAN', '', '', '');
    47305 : AglLanceFiche('PAY', 'EM_EDITCURSUS', '', '', '');
    47306 : AglLanceFiche('PAY', 'EM_EDITFORANIM', '', '', 'COLLECTIF');

    // Catalogues
    47205 : AglLanceFiche('PAY', 'EM_MULCATALOGUE', '', '', 'CONSULTCAT');
    47207 : AglLanceFiche('PAY', 'EM_MULCATCURSUS', '', '', 'CAT');  // Catalogue de cursus

    // Prévisionnel
    47201 : AglLanceFiche('PAY', 'EM_PREVISIONNEL', '', '', 'FOR'); // inscriptions a partir des formations
    47202 : AglLanceFiche('PAY', 'EM_PREVISIONNEL', '', '', 'SAL');// inscriptions a partir des salariés
    47203 : AglLanceFiche('PAY', 'EM_EDITPREV', '', '', 'COLLECTIF');
    47204 : AglLanceFiche('PAY', 'EM_EDITPREV', '', '', 'INDIVIDUEL');
    47206 : AglLanceFiche('PAY', 'EM_EDITPREV', '', '', 'SUIVIVALID');
    47208 : AglLanceFiche('PAY', 'EM_PREVISIONNEL', '', '', 'CUR');  // inscriptions aux cursus

    // DIF
    47311 : AglLanceFiche('PAY', 'EM_MULCOMPTDIF', '', '', '');

    47315 : AglLanceFiche('PAY', 'EMEDITRECAPDIF', '', '', '');

    // Planning
    47321 : AglLanceFiche('PAY', 'EM_PLANNING', '', '', '');
    //PT18 - Fin
  end;
end;


function EPrimes_Menus(Num: Integer; PRien: THPanel) : boolean ;
begin
  // XP 16-09-2004 : Pour virer les droits
  if Num = -1 then
  begin
    Et_RetireModule(913) ;
    result := False;
    Exit ;
  end else result := True ;
  case Num of
    40110: AglLanceFiche('PAY', 'MUL_SAISPRIM', '', '', 'P;X'); // Saisie des primes
    40114: AglLanceFiche('PAY', 'MUL_SAISPRIM', '', '', 'P'); // Saisie des primes
    40115: AglLanceFiche('PAY', 'MUL_SAISPRIM', '', '', 'C'); // Consultation des primes //PT23
    40120: AglLanceFiche('PAY', 'PGMULHISTOSAL', '', '', 'P'); // ANAL_HISTPRIMconsultation historique des primes
  end;
end;

function EAugmentation_Menus(Num: Integer; PRien: THPanel) : boolean ;
begin
  // XP 16-09-2004 : Pour virer les droits
  if Num = -1 then
  begin
    Et_RetireModule(915) ;
    Result := False;
    Exit ;
  end else result := True ;
  case Num of
    303110: AglLanceFiche('PAY', 'MUL_AUGMENTATION', '', '', 'SAISIE');
    303112: AglLanceFiche('PAY', 'MUL_AUGMENTATION', '', '', 'VALIDRESP');
    303111: AglLanceFiche('PAY', 'MUL_AUGMENTATION', '', '', 'PROPOSITION');
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 03/03/2008 / PT23
Modifié le ... :   /  /
Description .. : Gestion des menus du module 916 - Salariés
Mots clefs ... :
*****************************************************************}
function ESalarie_Menus(Num: Integer; PRien: THPanel) : boolean ;
begin
  if Num = -1 then
  begin
    Et_RetireModule(916) ;
    Result := False;
    Exit ;
  end else result := True ;

  case Num of
    42120: AglLanceFiche('PAY', 'EM_MULSALARIE', '', '', 'SALARIES');
    42199: AglLanceFiche('PAY', 'EM_MULSALARIE', '', '', 'GED');
  end;
end;

function PgActivAbsences: boolean;
var
  Q: TQuery;
  LaChaine, LeResp: string;
begin
{PT13
  LaChaine := string(GetParamSoc('SO_PGCHEMINEAGL'));
}
  LaChaine:= VerifieCheminPG (string(GetParamSoc('SO_PGCHEMINEAGL')));
//FIN PT13
  if LaChaine = '' then
  begin
    Q := OpenSql('SELECT SOC_DATA FROM PARAMSOC WHERE SOC_NOM = "SO_PGCHEMINEAGL"', True);
    if not Q.eof then
{PT13
      VH_Paie.PGCheminEagl := Q.FindField('SOC_DATA').AsString;
}
      VH_Paie.PGCheminEagl:= VerifieCheminPG (Q.FindField('SOC_DATA').AsString);
//FIN PT13
    Ferme(Q);
  end;
//  result := FALSE;
  LaChaine := '0000000000'; // Masque du nunéro de salarié
  LeSalarie := V_PGI.UserSalarie;
 // pt28  if Length(LeSalarie) < 10 then LeSalarie := Copy(LaChaine, 1, 10 - Length(LeSalarie)) + LeSalarie;
  if (VH_Paie.PgTypeNumSal='NUM') And  (Length(LeSalarie) < 10) then LeSalarie := Copy(LaChaine, 1, 10 - Length(LeSalarie)) + LeSalarie;  // pt28
  if LeSalarie = '' then
  begin
    {$IFNDEF ESALARIES}
    PgiBox('Vous n''avez pas de numéro salarié en tant qu''utilisateur PGI#13#10 Renseignez vous auprès de l''informatique interne !', 'Saisie des absences');
    {$ENDIF}
    result := FALSE;
    exit;
  end;
  if LeSalarie <> '' then
  begin
    Q := OpenSql('SELECT * FROM DEPORTSAL WHERE PSE_SALARIE = "' + LeSalarie + '"', True);
    if not Q.eof then
    begin
      SalAdm := Q.FindField('PSE_ADMINISTABS').AsString;
      SalVal := Q.FindField('PSE_OKVALIDABS').AsString;
//      result := true;
    end
    else
    begin
      {$IFNDEF ESALARIES}
      PgiBox('Vous êtes inconnu comme salarié#13#10 Renseignez vous au service du personnel', 'Saisie des absences');
      {$ENDIF}
      Ferme(Q);
      result := false;
      exit;
    end;
    LeResp := Q.FindField('PSE_RESPONSABS').AsString;
    Ferme(Q);
    if LeResp = '' then
    begin
      {$IFNDEF ESALARIES}
      PgiBox('Vous n''avez pas de responsable affecté !#13#10 Renseignez vous au service du personnel', 'Saisie des absences');
      {$ENDIF}
      result := false;
      exit;
    end;
    { DEB PT2 }
    Q := OpenSql('SELECT PSA_LIBELLE,PSA_PRENOM FROM SALARIES WHERE PSA_SALARIE = "' + LeSalarie + '"', True);
    if not Q.eof then
      Begin
      LeNomSal := Q.FindField('PSA_LIBELLE').AsString + ' ' + Q.FindField('PSA_PRENOM').AsString;
      GblTypeSal := 'SAL';
      End
    else
      begin
      Ferme(Q);
      Q := OpenSql('SELECT PSI_LIBELLE,PSI_PRENOM FROM INTERIMAIRES WHERE PSI_INTERIMAIRE = "' + LeSalarie + '"', True);
      if not Q.eof then
         Begin
         LeNomSal := Q.FindField('PSI_LIBELLE').AsString + ' ' + Q.FindField('PSI_PRENOM').AsString;
         GblTypeSal := 'INT';
         End
    { FIN PT2 }
      else
        begin
        {$IFNDEF ESALARIES}
        PgiBox('Vous êtes inconnu comme intervenant extèrieur de la paie#13#10 Renseignez vous au service du personnel', 'Saisie des absences');
        {$ENDIF}
        Ferme(Q);
        result := false;
        exit;
        end;
      End;
    Ferme(Q);
  end;
  // Desactivation des fonctionnalités administrateur ou responsable validation selon le salarié
  if SalAdm <> 'X' then FMenuG.RemoveGroup(-43800, TRUE); //  désactiver menu administrateur
  if SalVal <> 'X' then FMenuG.RemoveGroup(43700, TRUE); //  désactiver menu validation absences salarie
  //Salarié secrétaire
  NotSalAss := not (ExisteSql('SELECT PSE_ASSISTABS FROM DEPORTSAL WHERE PSE_ASSISTABS="' + LeSalarie + '"'));
  if NotSalAss then
  begin
    FMenuG.RemoveGroup(43200, TRUE); { PT1 FMenuG.RemoveItem(43165); }
    FMenuG.RemoveItem(43165);
  end;
  if not getparamsocsecur('SO_IFDEFCEGID',False) then { 13/10/2005 norme AGL }
    Begin
    FMenuG.RemoveItem(43215);
    FMenuG.RemoveItem(43250);
    End;
  //Intervenant exterieur
  ChangeMenuAbsence(LeSalarie);

  result := true;
  //BaseEPaie := (GetParamSoc ('SO_PGBASEEPAIE'));

end;

function PgActivPrimes: boolean;
var
  Q: TQuery;
  LaChaine, Lib: string;
  BaseEPaie: Boolean;
begin
  Lib := 'Saisie ' + VH_Paie.PGLibSaisPrim;
{PT13
  LaChaine := string(GetParamSoc('SO_PGCHEMINEAGL'));
}
  LaChaine:= VerifieCheminPG (string(GetParamSoc('SO_PGCHEMINEAGL')));
//FIN PT13
  if LaChaine = '' then
  begin
    Q := OpenSql('SELECT SOC_DATA FROM PARAMSOC WHERE SOC_NOM = "SO_PGCHEMINEAGL"', True);
    if not Q.eof then
{PT13
      VH_Paie.PGCheminEagl := Q.FindField('SOC_DATA').AsString;
}
      VH_Paie.PGCheminEagl:= VerifieCheminPG (Q.FindField('SOC_DATA').AsString);
//FIN PT13
    Ferme(Q);
  end;
//  result := FALSE;
  LaChaine := '0000000000'; // Masque du nunéro de salarié
  LeSalarie := V_PGI.UserSalarie;
 // pt28  if Length(LeSalarie) < 10 then LeSalarie := Copy(LaChaine, 1, 10 - Length(LeSalarie)) + LeSalarie;
  if (VH_Paie.PgTypeNumSal='NUM') And  (Length(LeSalarie) < 10) then LeSalarie := Copy(LaChaine, 1, 10 - Length(LeSalarie)) + LeSalarie;  // pt28
  if LeSalarie = '' then
  begin
    //    PgiBox('Vous n''avez pas de numéro salarié en tant qu''utilisateur PGI#13#10 Renseignez vous auprès de l''informatique interne !', Lib);
    result := FALSE;
    exit;
  end;
  if LeSalarie <> '' then
  begin
    Q := OpenSql('SELECT * FROM DEPORTSAL WHERE PSE_SALARIE = "' + LeSalarie + '"', True);
    if not Q.eof then
    begin
      SalAdm := Q.FindField('PSE_ADMINISTABS').AsString; // Administrateur eAgl des absences et des primes
      SalVal := Q.FindField('PSE_OKVALIDVAR').AsString; // Autorisation de valider les primes = saisie des primes
      ConsultP := FALSE;
      // On va regarder si le salarie valide les absences pour autoriser les accès en consultation uniquement
      if SalVal <> 'X' then
      begin
        SalVal := Q.FindField('PSE_OKVALIDABS').AsString;
        ConsultP := TRUE;
      end;
//      result := true;
    end
    else
    begin
      //      PgiBox('Vous êtes inconnu comme salarié#13#10 Renseignez vous au service du personnel', Lib);
      Ferme(Q);
      result := false;
      exit;
    end;
    Ferme(Q); //PT20
    if (SalAdm <> 'X') and (SalVal <> 'X') then
    begin
      //      PgiBox('Vous n''êtes pas autorisé à vous connecter, !#13#10Renseignez vous au service du personnel', Lib);
      result := false;
      exit;
    end;
    Q := OpenSql('SELECT PSI_LIBELLE,PSI_PRENOM FROM INTERIMAIRES WHERE PSI_INTERIMAIRE = "' + LeSalarie + '"', True);
    if not Q.eof then
      LeNomSal := Q.FindField('PSI_LIBELLE').AsString + ' ' + Q.FindField('PSI_PRENOM').AsString
    else
    begin
      //      PgiBox('Vous êtes inconnu comme salarié de la paie#13#10 Renseignez vous au service du personnel', Lib);
      Ferme(Q);
      result := false;
      exit;
    end;
    Ferme(Q);
  end;
  result := true;
  BaseEPaie := GetParamSoc('SO_PGBASEEPAIE');
  if (not BaseEPaie) then
    if (SalAdm <> 'X') then
    begin
      //      PgiBox('L''accès à la saisie ' + VH_Paie.PGLibSaisPrim + ' est interdit#13#10 Reconnectez vous ultérieurement SVP', Lib);
      Result := FALSE; // Base non accessible en tant que base de saisie déportée
    end;
end;

function PgActivAugmentation: boolean;
var
  Q: TQuery;
  LaChaine, Lib: string;
begin
{$ifdef aucasou}
  // XP 14-02-2005 : AFLINGUER
  if (V_PGI.UserSalarie = '0000002060') OR
  (V_PGI.UserSalarie = '0000000853') OR
  (V_PGI.UserSalarie = '0000002326') OR
  (V_PGI.UserSalarie = '0000002325') OR
  (V_PGI.UserSalarie = '0000000299') OR
  (V_PGI.UserSalarie = '0000001559') then
    result := True
  else
    result := False ;
{$endif}

  Lib := 'Saisie des augmentations';
{PT13
  LaChaine := string(GetParamSoc('SO_PGCHEMINEAGL'));
}
  LaChaine:= VerifieCheminPG (string(GetParamSoc('SO_PGCHEMINEAGL')));
//FIN PT13
  if LaChaine = '' then
  begin
    Q := OpenSql('SELECT SOC_DATA FROM PARAMSOC WHERE SOC_NOM = "SO_PGCHEMINEAGL"', True);
    if not Q.eof then
{PT13
      VH_Paie.PGCheminEagl := Q.FindField('SOC_DATA').AsString;
}
      VH_Paie.PGCheminEagl:= VerifieCheminPG (Q.FindField('SOC_DATA').AsString);
//FIN PT13
    Ferme(Q);
  end;
//  result := FALSE;
  LaChaine := '0000000000'; // Masque du nunéro de salarié
  LeSalarie := V_PGI.UserSalarie;
 // pt28 if Length(LeSalarie) < 10 then LeSalarie := Copy(LaChaine, 1, 10 - Length(LeSalarie)) + LeSalarie;
  if (VH_Paie.PgTypeNumSal='NUM') And  (Length(LeSalarie) < 10) then LeSalarie := Copy(LaChaine, 1, 10 - Length(LeSalarie)) + LeSalarie;  // pt28
  if LeSalarie = '' then
  begin
    result := FALSE;
    exit;
  end;
  if LeSalarie <> '' then
  begin
    If ExisteSQL('SELECT PSE_SALARIE FROM DEPORTSAL WHERE PSE_RESPONSVAR="' + LeSalarie + '"') then Result := True
      else
      begin
        If ExisteSQL('SELECT PSE_SALARIE FROM DEPORTSAL WHERE PSE_RESPONSABS="' + LeSalarie + '"') then Result := True
        else
        begin
          result := false;
          exit;
        end;
      end;
    Q := OpenSql('SELECT PSA_LIBELLE,PSA_PRENOM FROM SALARIES WHERE PSA_SALARIE = "' + LeSalarie + '"', True);
    if not Q.eof then
      LeNomSal := Q.FindField('PSA_LIBELLE').AsString + ' ' + Q.FindField('PSA_PRENOM').AsString
    else
    begin
      Ferme(Q);
      Q := OpenSql('SELECT PSI_LIBELLE,PSI_PRENOM FROM INTERIMAIRES WHERE (PSI_TYPEINTERIM="SAL" OR PSI_TYPEINTERIM="EXT") AND PSI_INTERIMAIRE = "' + LeSalarie + '"', True);
      if not Q.eof then
         LeNomSal := Q.FindField('PSI_LIBELLE').AsString + ' ' + Q.FindField('PSI_PRENOM').AsString
      ELSE
      result := false;
      exit;
    end;
    Ferme(Q);
  end;
  result := true;
end;


function PgActivFormation: boolean;
var
  Q: TQuery;
  LaChaine, Lib: string;
begin
  Lib := 'Saisie formations';
{PT13
  LaChaine := string(GetParamSoc('SO_PGCHEMINEAGL'));
}
  LaChaine:= VerifieCheminPG (string(GetParamSoc('SO_PGCHEMINEAGL')));
//FIN PT13
  if LaChaine = '' then
  begin
    Q := OpenSql('SELECT SOC_DATA FROM PARAMSOC WHERE SOC_NOM = "SO_PGCHEMINEAGL"', True);
    if not Q.eof then
{PT13
      VH_Paie.PGCheminEagl := Q.FindField('SOC_DATA').AsString;
}
      VH_Paie.PGCheminEagl:= VerifieCheminPG (Q.FindField('SOC_DATA').AsString);
//FIN PT13
    Ferme(Q);
  end;
//  result := FALSE;
  LaChaine := '0000000000'; // Masque du nunéro de salarié
  LeSalarie := V_PGI.UserSalarie;
// pt28  if Length(LeSalarie) < 10 then LeSalarie := Copy(LaChaine, 1, 10 - Length(LeSalarie)) + LeSalarie;
  if (VH_Paie.PgTypeNumSal='NUM') And  (Length(LeSalarie) < 10) then LeSalarie := Copy(LaChaine, 1, 10 - Length(LeSalarie)) + LeSalarie;  // pt28
  if LeSalarie = '' then
  begin
  //  PgiBox('Vous n''avez pas de numéro salarié en tant qu''utilisateur PGI#13#10 Renseignez vous auprès de l''informatique interne !', Lib);
    result := FALSE;
    exit;
  end;
  if LeSalarie <> '' then
  begin
    If Not ExisteSQL('SELECT PSE_SALARIE FROM DEPORTSAL WHERE PSE_RESPONSFOR="' + LeSalarie + '"') then //Result := True
//    else
    begin
      If Not ExisteSQL('SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_SECRETAIREFOR="' + LeSalarie + '"') then //Result := True
//      else
      begin
    //  PgiBox('Vous n''êtes pas autorisé à vous connecter, !#13#10Renseignez vous au service du personnel', Lib);
        result := false;
        exit;
      end;
    end;
    Q := OpenSql('SELECT PSI_LIBELLE,PSI_PRENOM FROM INTERIMAIRES WHERE (PSI_TYPEINTERIM="SAL" OR PSI_TYPEINTERIM="EXT") AND  PSI_INTERIMAIRE = "' + LeSalarie + '"', True);
    if not Q.eof then
      LeNomSal := Q.FindField('PSI_LIBELLE').AsString + ' ' + Q.FindField('PSI_PRENOM').AsString
    else
    begin
 //     PgiBox('Vous êtes inconnu comme salarié de la paie#13#10 Renseignez vous au service du personnel', Lib);
      Ferme(Q);
      result := false;
      exit;
    end;
    Ferme(Q);
  end;
  result := true;
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 03/03/2008 / PT23
Modifié le ... :   /  /
Description .. : Activation du module 916 - Salariés
Mots clefs ... :
*****************************************************************}
function PgActivSalarie: boolean;
var
  Q: TQuery;
  LaChaine, Lib: string;
begin
  Lib := 'Consultation des informations salarié';
  LaChaine:= VerifieCheminPG (string(GetParamSoc('SO_PGCHEMINEAGL')));

  if LaChaine = '' then
  begin
    Q := OpenSql('SELECT SOC_DATA FROM PARAMSOC WHERE SOC_NOM = "SO_PGCHEMINEAGL"', True);
    if not Q.eof then
      VH_Paie.PGCheminEagl:= VerifieCheminPG (Q.FindField('SOC_DATA').AsString);
    Ferme(Q);
  end;

  LaChaine := '0000000000'; // Masque du nunéro de salarié
  LeSalarie := V_PGI.UserSalarie;

  // pt28 if Length(LeSalarie) < 10 then LeSalarie := Copy(LaChaine, 1, 10 - Length(LeSalarie)) + LeSalarie;
  if (VH_Paie.PgTypeNumSal='NUM') And  (Length(LeSalarie) < 10) then LeSalarie := Copy(LaChaine, 1, 10 - Length(LeSalarie)) + LeSalarie;  // pt28
  if LeSalarie = '' then
  begin
    result := FALSE;
    exit;
  end;

  if LeSalarie <> '' then
  begin
    If ExisteSQL('SELECT PSE_SALARIE FROM DEPORTSAL WHERE PSE_RESPONSVAR="' + LeSalarie +
                 '" OR PSE_RESPONSFOR="' + LeSalarie + '" OR PSE_RESPONSABS="' + LeSalarie + '"') then
        Result := True
    else
    begin
        result := false;
        exit;
    end;

    Q := OpenSql('SELECT PSA_LIBELLE,PSA_PRENOM FROM SALARIES WHERE PSA_SALARIE = "' + LeSalarie + '"', True);
    if not Q.eof then
      LeNomSal := Q.FindField('PSA_LIBELLE').AsString + ' ' + Q.FindField('PSA_PRENOM').AsString
    else
    begin
      Ferme(Q);
      Q := OpenSql('SELECT PSI_LIBELLE,PSI_PRENOM FROM INTERIMAIRES WHERE (PSI_TYPEINTERIM="SAL" OR PSI_TYPEINTERIM="EXT") AND PSI_INTERIMAIRE = "' + LeSalarie + '"', True);
      if not Q.eof then
         LeNomSal := Q.FindField('PSI_LIBELLE').AsString + ' ' + Q.FindField('PSI_PRENOM').AsString
      ELSE
      result := false;
      exit;
    end;
    Ferme(Q);
  end;
  result := true;
end;

procedure ChangeMenuFormation(LeSalarie : String);
begin
  If Not ExisteSQL('SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_RESPONSFOR="'+LeSalarie+'"'+
  'AND PGS_CODESERVICE IN (SELECT PSO_SERVICESUP FROM SERVICEORDRE)') then FMenuG.RemoveItem(47405);
  //PT19 - Début
  FMenuG.RemoveItem(47302);
  FMenuG.RemoveItem(47303);
  //PT19 - Fin
end;

procedure ChangeMenuAugmentation(LeSalarie : String);
begin
      If Not ExisteSQL('SELECT PGS_CODESERVICE FROM SERVICES '+
      'LEFT JOIN HIERARCHIE ON PGS_HIERARCHIE=PHO_HIERARCHIE '+
      'WHERE PHO_NIVEAUH<=2 AND PGS_RESPONSVAR="'+LeSalarie+'"') then FMenuG.RemoveItem(303112);
      If Not ExisteSQL('SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_RESPONSABS="'+LeSalarie+'" AND PGS_RESPONSVAR<>"'+LeSalarie+'"') then
      FMenuG.RemoveItem(303111);
      If Not ExisteSQL('SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_RESPONSVAR="'+LeSalarie+'"') then
      FMenuG.RemoveItem(303110);

end;

procedure ChangeMenuAbsence(LeSalarie : String);
begin
      If ExisteSQL('SELECT PSI_INTERIMAIRE FROM INTERIMAIRES WHERE '+
      'PSI_INTERIMAIRE="'+LeSalarie+'" AND PSI_TYPEINTERIM="EXT"') then FMenuG.RemoveGroup(43100,True);
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 03/03/2008 / PT23
Modifié le ... :   /  /
Description .. : Désactivation des menus Salariés
Mots clefs ... :
*****************************************************************}
procedure ChangeMenuSalarie(LeSalarie : String);
begin
    If Not ExisteSQL('SELECT PSE_SALARIE FROM DEPORTSAL WHERE PSE_RESPONSVAR="' + LeSalarie +
                 '" OR PSE_RESPONSFOR="' + LeSalarie + '" OR PSE_RESPONSABS="' + LeSalarie + '"') then
        FMenuG.RemoveGroup(-42100,True);
end;

  {***********A.G.L.***********************************************
Auteur  ...... : Xavier PERSOUYRE
Créé le ...... : 24/10/2004
Modifié le ... :   /  /
Description .. : Permet de supprimer un module des menus.
Mots clefs ... :
*****************************************************************}
procedure Et_RetireModule(const NumModule: integer);
var
  i,j: integer;
  Seti1, Seti2: array of integer;
begin
  SetLength(Seti1, 0);
  SetLength(Seti2, 0);

  j := 0 ;
  for i := Low(FMenuG.FModuleNum) to High(FMenuG.FModuleNum) do
  begin
    if FMenuG.FModuleNum[I] <> NumModule then
    begin
      SetLength(Seti1, j + 1);
      Seti1[j] := FMenuG.FModuleNum[I];
      SetLength(Seti2, j + 1);
      Seti2[j] := Images[i];
      inc(j) ;
    end;
  end;
  FMenuG.SetModules(Seti1, Seti2);
  SetLength(Seti1, 0);
  SetLength(Seti2, 0);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Xavier PERSOUYRE
Créé le ...... : 24/10/2004
Modifié le ... :   /  /
Description .. : Permet de supprimer un module des menus.
Mots clefs ... :
*****************************************************************}
procedure Et_AjouteModule(const NumModule: integer);
var
  i,j: integer;
  Seti1, Seti2: array of integer;
begin
  SetLength(Seti1, 0);
  SetLength(Seti2, 0);

  // Contrôle que le module n'est pas déjà présent
  for i := Low(FMenuG.FModuleNum) to High(FMenuG.FModuleNum) do
    if FMenuG.FModuleNum[i] = NumModule then Exit ;

  j := 0 ;
  for i := Low(FMenuG.FModuleNum) to High(FMenuG.FModuleNum) do
  begin
    if FMenuG.FModuleNum[I] = -1 then
    begin
      SetLength(Seti1, j + 1);
      Seti1[j] := NumModule ;
      SetLength(Seti2, j + 1);
      Seti2[j] := -1;
      Break ;
    end
    else
    begin
      SetLength(Seti1, j + 1);
      Seti1[j] := FMenuG.FModuleNum[I];
      SetLength(Seti2, j + 1);
      Seti2[j] := -1;
      inc(j) ;
    end;
  end;
  FMenuG.SetModules(Seti1, Seti2);
  SetLength(Seti1, 0);
  SetLength(Seti2, 0);
end;

function PGCalcMul(Func,Params,WhereSQL : hstring ; TT : TDataset ; Total : Boolean) : hstring ;
var Salarie,Cumul : String;
    Q,Qt : TQuery;
    Heure : TDateTime;
    EtabDecl,CodTabl      : String; // PT4
    Tablette,LaValeur,LeSalarie : String;
    DateVal,DateDyn : TDateTime;
    CCn,Etab,Lib,ConvSal,StWhere : String;
    T : Tob;
    PGTobSal,PGTobTableDyn : Tob;
    Montant,MontantEuro : Double;
    DateApplic : String;
    DateMax:TDateTime;
    St:String;
    Nature:String;
begin
     If Total then exit;
     If Func = 'PGACQUISDIF' then
     begin
          Salarie := TT.FindField('PSA_SALARIE').AsString;
          Cumul := GetParamSOc('SO_PGCUMULDIFACQUIS');
          Q := OpenSQL('SELECT SUM(PHC_MONTANT) ACQUIS '+
          'FROM HISTOCUMSAL WHERE PHC_CUMULPAIE="'+Cumul+'" AND PHC_SALARIE="'+Salarie+'"',True);
          If Not Q.Eof then Result := Q.FindField('ACQUIS').AsString;
          Ferme(Q);
     end
     else If Func = 'PGPRISDIF' then
     begin
          Salarie := TT.FindField('PSA_SALARIE').AsString;
          Q := OpenSQL('SELECT SUM (PFO_NBREHEURE) NBHEURES FROM FORMATIONS '+
          'WHERE PFO_EFFECTUE="X" AND PFO_TYPEPLANPREV="DIF" AND PFO_SALARIE="'+Salarie+'"',True);
          If Not Q.Eof then Result := Q.FindField('NBHEURES').AsString;
          Ferme(Q);
     end
     else If Func = 'PGDEMANDEDIF' then
     begin
          Salarie := TT.FindField('PSA_SALARIE').AsString;
          Q := OpenSQL('SELECT SUM (PFI_DUREESTAGE) NBHEURES FROM INSCFORMATION '+
          'WHERE PFI_ETATINSCFOR="VAL" AND PFI_REALISE="-" AND PFI_TYPEPLANPREV="DIF" AND PFI_SALARIE="'+Salarie+'"',True);
          If Not Q.Eof then Result := Q.FindField('NBHEURES').AsString;
          Ferme(Q);
     end
     else If Func = 'PGHEUREMEDECINE' then
     begin
           Heure := TT.FindField('PVM_HEUREVISITE').AsDateTime;
           Result := FormatDateTime('hh:mm',Heure);
     end
// d PT4
     else if Func = 'ETABDECL' then
     begin
          EtabDecl := TT.FindField('PDA_ETABLISSEMENT').AsString;;
          if EtabDecl = '' then
             Result := '<<Tous>>'
          else
              result := EtabDecl;
     end
     else If Func = 'PGPOPSAL' then
     begin
          try
          Salarie := TT.FindField('PSA_SALARIE').AsString;
//          result := PGPopSalarie (Salarie, 'SAL');
          except;
          result := '';
          end;
     end
// f PT4
     else If Func = 'LIBSAISIEELTDYN' then
     begin
      Result := '';
      Tablette := TT.FindField('PHD_TABLETTE').AsString;
      LaValeur := TT.FindField('PHD_NEWVALEUR').AsString;
      CodTabl := TT.FindField('PHD_CODTABL').AsString;
      DateApplic := TT.FindField('PHD_DATEAPPLIC').AsString;
      Salarie := TT.FindField('PHD_SALARIE').AsString;  //PT16
      If (CodTabl <> '') and (LaValeur <> '') then
      begin
        result := GetLibelleZLTableDyna(StrToDateTime(DateApplic), CodTabl, Salarie, LaValeur); //PT25
//        StWhere := ' SELECT PTD_LIBELLECODE FROM TABLEDIMDET '
//                 + GetPlusPGCOMBOZONELIBRE(StrToDateTime(DateApplic), CodTabl, Salarie)
//                 + ' AND PTD_VALCRIT1="' + LaValeur + '"';
{//PT24
        //DEB PT16
        Q := OpenSQL('SELECT PSA_CONVENTION,PSA_ETABLISSEMENT FROM SALARIES WHERE PSA_SALARIE = "' + Salarie + '"',True,1);
        if not Q.Eof then
        begin
          ConvSal := Q.FindField('PSA_CONVENTION').AsString;
          Etab := Q.FindField('PSA_ETABLISSEMENT').AsString;
        end;
        Ferme(Q);
        //Vérifier s'il n'existe pas des valeurs pour l'établissement du salarié
        if ExisteSQL('SELECT PTE_CODTABL FROM TABLEDIMENT' +
        ' WHERE PTE_PREDEFINI="DOS" AND PTE_NODOSSIER="' + V_PGI.NoDossier + '"' +
        ' AND PTE_CODTABL="'+CodTabl+'"' +
        ' AND PTE_DTVALID<="'+USDATETIME(StrToDateTime(DateApplic))+'"' +
        ' AND PTE_NIVSAIS="ETB" AND PTE_VALNIV="' + Etab + '"') then
        begin
          St := 'SELECT MAX(PTE_DTVALID) AS DTVALID FROM TABLEDIMENT WHERE PTE_CODTABL="'+CodTabl+'"' +
            ' AND PTE_PREDEFINI="DOS" AND PTE_NODOSSIER="' + V_PGI.NoDossier + '"' +
            ' AND PTE_NIVSAIS="ETB" AND PTE_VALNIV="' + Etab + '" '+
            ' AND PTE_DTVALID<="' + USDATETIME(StrToDateTime(DateApplic)) + '"';
          Q := OpenSQL(St,True,1);
          DateMax := iDate1900;
          if not Q.Eof then
            DateMax := Q.FindField('DTVALID').AsDateTime;
          Ferme(Q);
          StWhere := ' SELECT PTD_LIBELLECODE FROM TABLEDIMDET ' +
            ' WHERE PTD_DTVALID = "' + USDATETIME(DateMax) + '" AND PTD_CODTABL="'+CodTabl+'"' +
            ' AND PTD_VALCRIT1="' + LaValeur + '"' +
            ' AND PTD_PREDEFINI="DOS" AND PTD_NODOSSIER="' + V_PGI.NoDossier + '"' +
            ' AND PTD_NIVSAIS="ETB" AND PTD_VALNIV="' + Etab + '" ORDER BY PTD_LIBELLECODE';
        end
        //Vérifier s'il n'existe pas des valeurs pour le dossier en général
        else if ExisteSQL('SELECT PTE_CODTABL FROM TABLEDIMENT' +
        ' WHERE PTE_PREDEFINI="DOS" AND PTE_NODOSSIER="' + V_PGI.NoDossier + '"' +
        ' AND PTE_DTVALID<="'+USDATETIME(StrToDateTime(DateApplic))+'"' +
        ' AND PTE_CODTABL="'+CodTabl+'"' +
        ' AND PTE_NIVSAIS="GEN"') then
        begin
          St := 'SELECT MAX(PTE_DTVALID) AS DTVALID FROM TABLEDIMENT WHERE PTE_CODTABL="'+CodTabl+'"' +
            ' AND PTE_PREDEFINI="DOS" AND PTE_NODOSSIER="' + V_PGI.NoDossier + '"' +
            ' AND PTE_NIVSAIS="GEN"' +
            ' AND PTE_DTVALID<="' + USDATETIME(StrToDateTime(DateApplic)) + '"';
          Q := OpenSQL(St,True,1);
          DateMax := iDate1900;
          if not Q.Eof then
            DateMax := Q.FindField('DTVALID').AsDateTime;
          Ferme(Q);
          StWhere := ' SELECT PTD_LIBELLECODE FROM TABLEDIMDET ' +
            ' WHERE PTD_DTVALID ="' + USDATETIME(DateMax) + '" AND PTD_CODTABL="'+CodTabl+'"' +
            ' AND PTD_VALCRIT1="' + LaValeur + '"' +
            ' AND PTD_PREDEFINI="DOS" AND PTD_NODOSSIER="' + V_PGI.NoDossier + '"' +
            ' AND PTD_NIVSAIS="GEN" ORDER BY PTD_LIBELLECODE';
        end
        //Vérifier s'il n'existe pas des valeurs pour STD + convention
        else if ExisteSQL('SELECT PTE_CODTABL FROM TABLEDIMENT' +
        ' WHERE PTE_PREDEFINI="STD"' +
        ' AND PTE_DTVALID<="'+USDATETIME(StrToDateTime(DateApplic))+'"' +
        ' AND PTE_CODTABL="'+CodTabl+'"' +
        ' AND PTE_NIVSAIS="CON" AND PTE_VALNIV="' + ConvSal + '"') then
        begin
          St := 'SELECT MAX(PTE_DTVALID) AS DTVALID FROM TABLEDIMENT WHERE PTE_CODTABL="'+CodTabl+'"' +
            ' AND PTE_PREDEFINI="STD"' +
            ' AND PTE_NIVSAIS="CON" AND PTE_VALNIV="' + ConvSal + '"'+
            ' AND PTE_DTVALID<="' + USDATETIME(StrToDateTime(DateApplic)) + '"';
          Q := OpenSQL(St,True,1);
          DateMax := iDate1900;
          if not Q.Eof then
            DateMax := Q.FindField('DTVALID').AsDateTime;
          Ferme(Q);
          StWhere := ' SELECT PTD_LIBELLECODE FROM TABLEDIMDET ' +
            ' WHERE PTD_DTVALID ="' + USDATETIME(DateMax) + '" AND PTD_CODTABL="'+CodTabl+'"' +
            ' AND PTD_VALCRIT1="' + LaValeur + '"' +
            ' AND PTD_PREDEFINI="STD"' +
            ' AND PTD_NIVSAIS="CON" AND PTD_VALNIV="' + ConvSal + '" ORDER BY PTD_LIBELLECODE';
        end
        //Vérifier s'il n'existe pas des valeurs pour STD + Convention 000
        else if ExisteSQL('SELECT PTE_CODTABL FROM TABLEDIMENT' +
        ' WHERE PTE_PREDEFINI="STD"' +
        ' AND PTE_DTVALID<="'+USDATETIME(StrToDateTime(DateApplic))+'"' +
        ' AND PTE_CODTABL="'+CodTabl+'"' +
        ' AND PTE_NIVSAIS="CON" AND PTE_VALNIV="' + ConvSal + '"') then
        begin
          St := 'SELECT MAX(PTE_DTVALID) AS DTVALID FROM TABLEDIMENT WHERE PTE_CODTABL="'+CodTabl+'"' +
            ' AND PTE_PREDEFINI="STD"' +
            ' AND PTE_NIVSAIS="CON" AND PTE_VALNIV="000"'+
            ' AND PTE_DTVALID<="' + USDATETIME(StrToDateTime(DateApplic)) + '"';
          Q := OpenSQL(St,True,1);
          DateMax := iDate1900;
          if not Q.Eof then
            DateMax := Q.FindField('DTVALID').AsDateTime;
          Ferme(Q);
          StWhere := ' SELECT PTD_LIBELLECODE FROM TABLEDIMDET ' +
            ' WHERE PTD_DTVALID ="' + USDATETIME(DateMax) + '" AND PTD_CODTABL="'+CodTabl+'"' +
            ' AND PTD_VALCRIT1="' + LaValeur + '"' +
            ' AND PTD_PREDEFINI="STD"' +
            ' AND PTD_NIVSAIS="CON" AND PTD_VALNIV="000" ORDER BY PTD_LIBELLECODE';
        end
        //Vérifier s'il n'existe pas des valeurs pour STD + GEN
        else if ExisteSQL('SELECT PTE_CODTABL FROM TABLEDIMENT' +
        ' WHERE PTE_PREDEFINI="STD"' +
        ' AND PTE_DTVALID<="'+USDATETIME(StrToDateTime(DateApplic))+'"' +
        ' AND PTE_CODTABL="'+CodTabl+'"' +
        ' AND PTE_NIVSAIS="GEN"') then
        begin
          St := 'SELECT MAX(PTE_DTVALID) AS DTVALID FROM TABLEDIMENT WHERE PTE_CODTABL="'+CodTabl+'"' +
            ' AND PTE_PREDEFINI="STD"' +
            ' AND PTE_NIVSAIS="GEN"'+
            ' AND PTE_DTVALID<="' + USDATETIME(StrToDateTime(DateApplic)) + '"';
          Q := OpenSQL(St,True,1);
          DateMax := iDate1900;
          if not Q.Eof then
            DateMax := Q.FindField('DTVALID').AsDateTime;
          Ferme(Q);
          StWhere := ' SELECT PTD_LIBELLECODE FROM TABLEDIMDET ' +
            ' WHERE PTD_DTVALID ="' + USDATETIME(DateMax) + '" AND PTD_CODTABL="'+CodTabl+'"' +
            ' AND PTD_VALCRIT1="' + LaValeur + '"' +
            ' AND PTD_PREDEFINI="STD"' +
            ' AND PTD_NIVSAIS="GEN" ORDER BY PTD_LIBELLECODE';
        end
        //Vérifier s'il n'existe pas des valeurs pour CEG
        else
        begin
          St := 'SELECT MAX(PTE_DTVALID) AS DTVALID FROM TABLEDIMENT WHERE PTE_CODTABL="'+CodTabl+'"'+
            ' AND PTE_PREDEFINI="CEG" AND PTE_NIVSAIS="GEN" '+
            ' AND PTE_DTVALID<="' + USDATETIME(StrToDateTime(DateApplic)) + '"';
          Q := OpenSQL(St,True,1);
          DateMax := iDate1900;
          if not Q.Eof then
            DateMax := Q.FindField('DTVALID').AsDateTime;
          Ferme(Q);
          StWhere := ' SELECT PTD_LIBELLECODE FROM TABLEDIMDET ' +
            ' WHERE PTD_DTVALID ="' + USDATETIME(DateMax) + '" AND PTD_CODTABL="'+CodTabl+'"' +
            ' AND PTD_VALCRIT1="' + LaValeur + '"' +
            ' AND PTD_PREDEFINI="CEG" AND PTD_NIVSAIS="GEN" ORDER BY PTD_LIBELLECODE';
        end;
        //Result := RechDom(Tablette,LaValeur,False);
//        Q := OpenSQL('SELECT PTD_LIBELLECODE FROM TABLEDIMDET WHERE ##PTD_PREDEFINI## AND PTD_CODTABL="' + CodTabl + '" AND PTD_VALCRIT1="' + LaValeur + '"' +
//          ' AND PTD_DTVALID=(SELECT MAX(PTD_DTVALID) FROM TABLEDIMDET WHERE PTD_CODTABL="' + CodTabl + '" AND PTD_DTVALID<="' + USDATETIME(StrToDateTime(DateApplic)) + '")',True);
}
//        Q := OpenSQL(StWhere,True);
//        //FIN PT16
//        If Not Q.Eof then Result := Q.FindField('PTD_LIBELLECODE').AsString;
//        Ferme(Q);
      end;
     end
     else If Func = 'LIBELLEHISTO' then
     begin
      Result := '';
      Tablette := TT.FindField('PHD_TABLETTE').AsString;
      //PT17 FQ 15065 Pb longueur du champ PHD_TABLETTE, il faudrait qu'il fasse 35 caract
      if Tablette = 'PGLIBQUALIFICATIO' then Tablette := 'PGLIBQUALIFICATION';
      LaValeur := TT.FindField('PHD_NEWVALEUR').AsString;
      DateApplic := TT.FindField('PHD_DATEAPPLIC').AsString;
      If (Tablette <> '') and (LaValeur <> '') then
      begin
        //DEB PT17
        if (Tablette = 'PGLIBQUALIFICATION') or (Tablette = 'PGLIBCOEFFICIENT') or (Tablette = 'PGLIBINDICE') or (Tablette = 'PGLIBNIVEAU') then
        begin
          Salarie := TT.FindField('PHD_SALARIE').AsString;
          Q:= OpenSQL('SELECT PSA_CONVENTION FROM SALARIES WHERE PSA_SALARIE="'+Salarie+'"',True);
          if (Tablette = 'PGLIBQUALIFICATION') then
            Nature := 'QUA';
          if (Tablette = 'PGLIBCOEFFICIENT') then
            Nature := 'COE';
          if (Tablette = 'PGLIBINDICE') then
            Nature := 'IND';
          if (Tablette = 'PGLIBNIVEAU') then
            Nature := 'NIV';
          St := 'Select PMI_LIBELLE FROM MINIMUMCONVENT M1' +
            ' WHERE  ##PMI_PREDEFINI## PMI_NATURE="' + Nature + '"' +
            ' AND PMI_CODE="' + LaValeur + '"' +
            ' AND (PMI_CONVENTION="'+Q.FindField('PSA_CONVENTION').AsString+'" OR PMI_CONVENTION="000")' +
            ' AND (PMI_PREDEFINI="DOS"' +
            ' OR (PMI_PREDEFINI="STD" AND PMI_CONVENTION="'+Q.FindField('PSA_CONVENTION').AsString+'" ' +
            '   AND NOT EXISTS(SELECT PMI_CODE FROM MINIMUMCONVENT ' +
            '   WHERE ##PMI_PREDEFINI## PMI_CODE=M1.PMI_CODE AND PMI_PREDEFINI="DOS"))' +
            ' OR (PMI_PREDEFINI="STD" AND PMI_CONVENTION="000" ' +
            '   AND NOT EXISTS(SELECT PMI_CODE FROM MINIMUMCONVENT '+
            '   WHERE ##PMI_PREDEFINI## PMI_CODE=M1.PMI_CODE ' +
            '   AND (PMI_PREDEFINI="DOS" OR (PMI_PREDEFINI="STD" AND PMI_CONVENTION="'+Q.FindField('PSA_CONVENTION').AsString+'")))))' +
            ' ORDER BY PMI_CODE';
          Qt:= OpenSQL(St,True);
          if not Qt.Eof then
            Result := Qt.FindField('PMI_LIBELLE').AsString;
          Ferme(Q);
          Ferme(Qt);
        end
        //FIN PT17
        else
          Result := RechDom(Tablette,LaValeur,False);
      end;
     end
  else if Func = 'CHOIXMONTANT' then     //DEB PT10
  begin
    if (Params = 'SYN') then
    begin
      Montant := TT.FindField('PEY_MONTANT').AsFloat;
      MontantEuro := TT.FindField('PEY_MONTANTEURO').AsFloat;
    end
    else
    begin
      Montant := TT.FindField('PED_MONTANT').AsFloat;
      MontantEuro := TT.FindField('PED_MONTANTEURO').AsFloat;
    end;
    if (Montant <> 0) then
      Result := FloatToStr(Montant)
    else
      Result := FloatToStr(MontantEuro);
  end                                    //FIN PT10
  else If Func = 'LIBELLEZL' then
  begin
    Result := '';
    If PGTobSalMul <> Nil then PGTobSal := PGTobSalMul
    else If PGTobSalListe <> Nil then PGTobSal := PGTobSalListe;
    if PGTobTableDynMul <> Nil then PGTobTableDyn := PGTobTableDynMul
    else if PGTobTableDynListe <> Nil then PGTobTableDyn := PGTobTableDynListe;

    Tablette := TT.FindField('PHD_TABLETTE').AsString;
    If Tablette = 'PGLIBQUALIFICATIO' then tablette := 'PGLIBQUALIFICATION';
    LaValeur := TT.FindField('PHD_NEWVALEUR').AsString;
    DateVal := TT.FindField('PHD_DATEAPPLIC').AsdateTime;
    CodTabl := TT.FindField('PHD_CODTABL').AsString;
    LeSalarie := TT.FindField('PHD_SALARIE').AsString;
    T := PGTobSal.FindFirst(['PSA_SALARIE'],[LeSalarie],False);
    If T <> Nil then
    begin
      Etab := T.GetValue('PSA_ETABLISSEMENT');
      CCn := T.GetValue('PSA_CONVENTION');
    end;
    Lib := '';
    //recherche niveau etab     //@@GGU A étudier/confirmer pour la FQ 15361
    T := PGTobTableDyn.FindFirst(['PTD_CODTABL','PTD_NIVSAIS','PTD_VALNIV','PTD_VALCRIT1'],[CodTabl,'ETB',Etab,LaValeur],False);
    If T <> Nil then
    begin
       DateDyn := T.GetValue('PTD_DTVALID');
       While (T<>Nil) AND (DateDyn > DateVal) do
       begin
          T := PGTobTableDyn.FindNext(['PTD_CODTABL','PTD_NIVSAIS','PTD_VALNIV','PTD_VALCRIT1'],[CodTabl,'ETB',Etab,LaValeur],False);
          If T <> Nil then DateDyn := T.GetValue('PTD_DTVALID');
       end;
       If T <> Nil then Lib := T.GetValue('PTD_LIBELLECODE');
    end;
    If Lib = '' then
    begin
      T := PGTobTableDyn.FindFirst(['PTD_CODTABL','PTD_NIVSAIS','PTD_VALNIV','PTD_VALCRIT1'],[CodTabl,'CON',CCn,LaValeur],False);
      If T <> Nil then
      begin
        DateDyn := T.GetValue('PTD_DTVALID');
        While (T<>Nil) AND (DateDyn > DateVal) do
        begin
          T := PGTobTableDyn.FindNext(['PTD_CODTABL','PTD_NIVSAIS','PTD_VALNIV','PTD_VALCRIT1'],[CodTabl,'CON',CCn,LaValeur],False);
          If T <> Nil then DateDyn := T.GetValue('PTD_DTVALID');
        end;
        If T <> Nil then Lib := T.GetValue('PTD_LIBELLECODE');
      end;
    end;
    If Lib = '' then
    begin
      T := PGTobTableDyn.FindFirst(['PTD_CODTABL','PTD_NIVSAIS','PTD_VALCRIT1'],[CodTabl,'GEN',LaValeur],False);
      If T <> Nil then
      begin
        DateDyn := T.GetValue('PTD_DTVALID');
        While (T<>Nil) AND (DateDyn > DateVal) do
        begin
          T := PGTobTableDyn.FindNext(['PTD_CODTABL','PTD_NIVSAIS','PTD_VALCRIT1'],[CodTabl,'GEN',LaValeur],False);
          If T <> Nil then DateDyn := T.GetValue('PTD_DTVALID');
        end;
        If T <> Nil then Lib := T.GetValue('PTD_LIBELLECODE');
      end;
    end;
    Result := Lib;
{    Q := OpenSQL('SELECT PTD_LIBELLECODE FROM TABLEDIMDET WHERE '+
    'PTD_CODTABL="'+CodTabl+'" AND PTD_DTVALID<="'+UsdateTime(DateVal)+'" '+
    'AND PTD_VALCRIT1="'+laValeur+'" ORDER BY PTD_DTVALID DESC',True);
    If Not Q.Eof then
    begin
        Q.First;
        Result := Q.FindField('PTD_LIBELLECODE').AsString;
    end;
    Ferme(Q);            }
  end
  //PT11 - Début
  else If Func = 'LIBEMPLOI' then
  begin
    Result := '';
    LaValeur := TT.FindField('PMF_LIBELLEEMPLOI').AsString;
    If (LaValeur <> '') then Result := RechDom('PGLIBEMPLOI',LaValeur,False);
  //PT11 - Fin
  end else DefProcCalcMul (Func,Params,WhereSQL, TT, Total) ; // XP 26.11.2007
end;

procedure PGChangeTabletteResponsable;
var
  St: string;
  iChamp: integer;
begin
{$IFNDEF EMANAGER} //PT18
    If (Not GetParamSocSecur('SO_PGINTERVENANTEXT', False)) And (Not PGBundleHierarchie) then exit; //PT12 //PT21
{$ENDIF}
    st := 'PGSALARIE';
    iChamp := TTToNum(St);
    if iChamp > 0 then 
    Begin
        //PT23 Inversion du test
    	If PGBundleHierarchie Then //PT21
    		V_PGI.DECombos[iChamp].champ := 'SALREMPL'
    	Else If GetParamSocSecur('SO_PGINTERVENANTEXT', False) Then
    		V_PGI.DECombos[iChamp].champ := ';SALARIE;SALREMPL';
    End;

    st := 'PGSALARIEINT';
    iChamp := TTToNum(St);
    if iChamp > 0 then 
    Begin
        //PT23 Inversion du test
    	If PGBundleHierarchie Then //PT21
    		V_PGI.DECombos[iChamp].champ := 'SALARIE;RESPONSNDF;RESPONSVAR;RESPONSABS;RESPSERVICE;TUTEURSAL;RESPONSVAL'
    	Else If GetParamSocSecur('SO_PGINTERVENANTEXT', False) Then
    		V_PGI.DECombos[iChamp].champ := ';RESPONSNDF;RESPONSVAR;RESPONSABS;RESPSERVICE;TUTEURSAL';
	End;
    
    st := 'PGRESPONSABS';
    iChamp := TTToNum(St);
    if iChamp > 0 then
    begin
         V_PGI.DECombos[iChamp].ChampLib := 'PSI_LIBELLE';
         V_PGI.DECombos[iChamp].Code := 'PSI_INTERIMAIRE';
         V_PGI.DECombos[iChamp].Prefixe := 'PSI';
         V_PGI.DECombos[iChamp].Where := '(PSI_TYPEINTERIM="EXT" OR PSI_TYPEINTERIM="SAL") AND PSI_INTERIMAIRE IN (SELECT DISTINCT PSE_RESPONSABS FROM DEPORTSAL)';
    end;
    
    //PT24
    st := 'PGRESPONSFOR';
    iChamp := TTToNum(St);
    if iChamp > 0 then
    begin
         V_PGI.DECombos[iChamp].ChampLib := 'PSI_LIBELLE';
         V_PGI.DECombos[iChamp].Code := 'PSI_INTERIMAIRE';
         V_PGI.DECombos[iChamp].Prefixe := 'PSI';
         V_PGI.DECombos[iChamp].Where := '(PSI_TYPEINTERIM="EXT" OR PSI_TYPEINTERIM="SAL") AND PSI_INTERIMAIRE IN (SELECT DISTINCT PSE_RESPONSFOR FROM DEPORTSAL)';
    end;
    
    st := 'PGRESPONSVARIABLE';
    iChamp := TTToNum(St);
    if iChamp > 0 then
    begin
         V_PGI.DECombos[iChamp].ChampLib := 'PSI_LIBELLE';
         V_PGI.DECombos[iChamp].Code := 'PSI_INTERIMAIRE';
         V_PGI.DECombos[iChamp].Prefixe := 'PSI';
         V_PGI.DECombos[iChamp].Where := '(PSI_TYPEINTERIM="EXT" OR PSI_TYPEINTERIM="SAL") AND PSI_INTERIMAIRE IN (SELECT PSE_RESPONSVAR FROM DEPORTSAL)&#@';
    end;
end;

{ DEB PT9 }
function Presence_Menus(Num: Integer; PRien: THPanel) : boolean ;
Begin
  case Num of
    { Menu Cycle }
    347512: AglLanceFiche('PAY', 'JOURNEETYPE_MUL', '', '', '');
    347514: AglLanceFiche('PAY', 'DROITJOURTYPE_FSL', '', '', '');
    347516: AglLanceFiche('PAY', 'DROITJOURNEE_FSL', '', '', '');
    347520: AglLanceFiche('PAY', 'MODELECYCLE_FSL', '', '', '');
    347530: AglLanceFiche('PAY', 'CYCLE_FSL', '', '', '');

    347550: AglLanceFiche('PAY', 'EXCEPTCYCLE_MUL', '', '', '');
    347610: AglLanceFiche('PAY', 'MOTIFPRESENCE_MUL', '', '', '');
    347650: AglLanceFiche('PAY', 'PGREGROUP_MUL', '', '', 'MAB');  //  pt14
    347620: AGLLanceFiche('PAY', 'VARIABLEPRE_MUL', '', '', '');
    347630: AglLanceFiche('PAY', 'COMPTEURPRES_MUL', '', '', '');
    347640: AglLanceFiche('PAY', 'PROFILPRES_MUL', '', '', '');

    347110: AglLanceFiche('PAY', 'AFFPROFILPRES_MUL', '', '', '');
    347120: AglLanceFiche('PAY', 'EXCEPTPRESSAL_MUL', '', '', '');
    347131: AglLanceFiche('PAY', 'PRESENCE_MUL', '', '', ';MENU;;PRE');
    347132: AglLanceFiche('PAY', 'MVTPRESGRP_MUL', '', '', '');
    347140: AglLanceFiche('PAY', 'PRECALCCOMPTEURS', '', '', '');
    347151: AglLanceFiche('PAY', 'INTEGPRESPAIE_MUL', '', '', '');  // PT15
    347152: AglLanceFiche('PAY', 'PRESINTEGPAIE_MUL', '', '', '');  // PT15
    347160: AglLanceFiche('PAY', 'PGPRESCLOTURE', '', '', '');

    347211: AglLanceFiche('PAY', 'COMPTEURSCALC_MUL', '', '', '');
    347212: AglLanceFiche('PAY', 'TOBVIEW_COMPTEURS', '', '', '');
    347213: AglLanceFiche('PAY', 'CUBE_COMPTEURS', '', '', '');
    347214: AglLanceFiche('PAY', 'EDITCOMPTEURSPRES', '', '', '');

    347221: AglLanceFiche('PAY', 'PLANRYTHMETRAV', '', '', '');
    347222: AglLanceFiche('PAY', 'PLANPRESSAL', '', '', '');
   // 347223: AglLanceFiche('PAY', 'SUPERVABSPRES', '', '', 'RESP;;ABSPRES');
    347224: AglLanceFiche('PAY', 'PLANCOMPTEURSPRES', '', '', '');
  End;
End;
{FIN PT9}
end.

