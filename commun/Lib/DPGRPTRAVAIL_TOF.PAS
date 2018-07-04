{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 05/04/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : DP_GRPTRAVAIL ()
Mots clefs ... : TOF;DP_GRPTRAVAIL
*****************************************************************}
Unit DPGRPTRAVAIL_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     Fe_main,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
{$else}
     eMul,
     maineagl,
{$ENDIF}
     uTob,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF;
Type
  TOF_DPGRPTRAVAIL = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  end ;
  function CtrlFiltreDonnees (CdeSql : string):integer;
  procedure DPLanceFiche_GrpTravail (GrpDonnees,NatFiche,CodFiche : String; UserOnly : Boolean = FALSE);

Implementation

////////////////////////////////////////////////////////////////////////////////
{Ctrl des paramétres de la table FILTREDONNEES.
 Retourne un code erreur si il y a, sinon retourne 0.
 Liste code erreur de retour apres vérification des champs :
 1 - FTD_TABLELIEN
 2 - FTD_CLETABLELIEN
 3 - FTD_TABLELIEE
 4 - FTD_CLETABLELIEE}

function CtrlFiltreDonnees(CdeSql: string): integer;
var
  Code_Erreur : integer;
  tob_table : TOB;
  s_TabLien,s_TabLiee,s_CleTabLiee,s_CleTabLien : string;
begin
  Code_Erreur := 0;

  tob_table := tob.Create('FILTREDONNEES',nil,-1);
  tob_table.LoadDetailFromSQL(CdeSql);

  s_TabLien    := tob_table.detail[0].GetValue('FTD_TABLELIEN');
  s_TabLiee    := tob_table.detail[0].GetValue('FTD_TABLELIEE');
  s_CleTabLiee := tob_table.detail[0].GetValue('FTD_CLETABLELIEE');
  s_CleTabLien := tob_table.detail[0].GetValue('FTD_CLETABLELIEN');

  if not ExisteSQL('SELECT 1 FROM DETABLES WHERE DT_NOMTABLE = "'+s_TabLien+'"') then
    Code_Erreur := 1
  else if not ExisteSQL('SELECT 1 FROM DECHAMPS WHERE DH_PREFIXE = "'+TableToPrefixe(s_TabLien)+'" AND DH_NOMCHAMP = "'+s_CleTabLien+'"') then
    Code_Erreur := 2
  else if not ExisteSQL('SELECT 1 FROM DETABLES WHERE DT_NOMTABLE = "'+s_TabLiee+'"') then
    Code_Erreur := 3
  else if not ExisteSQL('SELECT 1 FROM DECHAMPS WHERE DH_PREFIXE = "'+TableToPrefixe(s_TabLiee)+'" AND DH_NOMCHAMP = "'+s_CleTabLiee+'"') then
    Code_Erreur := 4;

  tob_table.Free;

  result := Code_Erreur;
end;
////////////////////////////////////////////////////////////////////////////////
procedure DPLanceFiche_GrpTravail (GrpDonnees,NatFiche,CodFiche : String; UserOnly : Boolean);
var
  s_sql,s_Argument,s_MsgAdm : string;
  CtrlTab : integer;
begin
  // Ctrl de l'existence du groupe de données dans le Table FILTREDONEES.
  s_sql := 'SELECT * FROM FILTREDONNEES WHERE FTD_CODE = "'+GrpDonnees+'"';
  if ExisteSQL(s_sql) then
  begin
    if UserOnly then
      s_Argument := GrpDonnees+'|'+NatFiche+'|'+CodFiche+'|X'
    else
      s_Argument := GrpDonnees+'|'+NatFiche+'|'+CodFiche+'|-';

    // Ctrl des infos.avec retour si erreur.
    CtrlTab := CtrlFiltreDonnees(s_sql);
    s_MsgAdm := 'Paramètrage incorrecte, Merci de vérifier les informations de la table FILTREDONNEES';
    case CtrlTab of
      1: PGIBox(s_MsgAdm,'TABLE FILTREDONNEES : champ FTD_TABLELIEN');
      2: PGIBox(s_MsgAdm,'TABLE FILTREDONNEES : champ FTD_CLETABLELIEN');
      3: PGIBox(s_MsgAdm,'TABLE FILTREDONNEES : champ FTD_TABLELIEE');
      4: PGIBox(s_MsgAdm,'TABLE FILTREDONNEES : champ FTD_CLETABLELIEE');
    else
     // Ctrl si l'utilisateur à les droits d'accès.
      {$IFNDEF FILTREDONNEES} //LM20071008 : on laisse le gestionnaire de droit déterminer
      if not V_PGI.Superviseur then
      begin
        //recherche d'au moins 1 enreg dans la table LIENDONNEES pour l'utilisateur en cours
        s_sql := 'SELECT 1 FROM LIENDONNEES WHERE LND_NOM = "'+GrpDonnees+'" '+
                 'AND LND_USERID = "'+V_PGI.User+'"';
        if not ExisteSQL(s_sql) then
          PGIBox('Vous n''avez pas les droits car vous n''êtes affecté à aucun groupe de travail.'+#10#13+
                 'Veuillez contacter votre administrateur.',TitreHalley)
        else
          AGLLanceFiche ('DP', 'GRPTRAVAIL', '', '', s_Argument);
      end
      else
      {$ENDIF FILTREDONNEES}
        // Droit Admin
        AGLLanceFiche ('DP', 'GRPTRAVAIL', '', '', s_Argument);
    end;
  end
  else
    PGIBox('Lancement de la fiche "GRPTRAVAIL" impossible.'+#10#13+
            'Le groupe de données "'+GrpDonnees+'" est inconnu.',TitreHalley);
end;
////////////////////////////////////////////////////////////////////////////////
procedure TOF_DPGRPTRAVAIL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_DPGRPTRAVAIL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_DPGRPTRAVAIL.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_DPGRPTRAVAIL.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_DPGRPTRAVAIL.OnArgument (S : String ) ;
begin
  Inherited ;
end ;

procedure TOF_DPGRPTRAVAIL.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_DPGRPTRAVAIL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_DPGRPTRAVAIL.OnCancel () ;
begin
  Inherited ;
end ;


Initialization
  registerclasses ( [ TOF_DPGRPTRAVAIL ] ) ;
end.
