{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 17/02/2004
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTCONSOS ()
Mots clefs ... : TOF;BTCONSOS
*****************************************************************}
unit BTConsos_TOF;

interface

uses {$IFDEF VER150} variants,{$ENDIF} StdCtrls,
  Controls,
  Classes,
  {$IFNDEF EAGLCLIENT}
  db,
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  Fe_Main,
  {$ELSE}
  MaineAGL,
  {$ENDIF}
  forms,
  sysutils,
  ComCtrls,
  Menus,
  HCtrls,
  HEnt1,
  HMsgBox,
  UTOF,
  Stat,
  UTOB,
  utobview,
  AGLInit,
  facture,
  uTofAfBaseCodeAffaire;

type
  TOF_BTCONSOS = class(TOF_AFBASECODEAFFAIRE)
    procedure OnNew; override;
    procedure OnDelete; override;
    procedure OnUpdate; override;
    procedure OnLoad; override;
    procedure OnArgument(S: string); override;
    procedure OnClose; override;
    procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_: THEdit); override;

  private
    TV1: TTobViewer;
    TOBCONSO: TOB;
    AffichageAuto, OldAutoSearch : Boolean;
    CodeAffaire, TypeConso: string;
    procedure ChargeLesConsos;
    procedure AjouteChampSup(var TOBTMP: TOB);
  end;

implementation

procedure TOF_BTCONSOS.OnNew;
begin
  inherited;
end;

procedure TOF_BTCONSOS.OnDelete;
begin
  inherited;
end;

procedure TOF_BTCONSOS.OnUpdate;
begin
  inherited;
end;

procedure TOF_BTCONSOS.OnLoad;
begin
  inherited;
  CodeAffaire := GetControltext('BCOAFFAIRE');
  if codeaffaire = '' then
    begin
    HShowMessage('0;?caption?;Veuillez choisir une affaire.;E;O;O;O;','','') ;
    Exit;
    end;

  ChargeLesConsos;
  if TOBCONSO <> nil then
  begin
    TFStat(Ecran).LaTOB := TOBCONSO;
  end;
end;

procedure TOF_BTCONSOS.OnArgument(S: string);
var critere : string;
begin
  inherited;
  AffichageAuto:=False;
  TypeConso:='';

  // Recup des critères
  Critere:=(Trim(ReadTokenSt(S)));
  While (Critere <>'') do
    BEGIN
    if Critere = 'AUTO' then AffichageAuto:=True
    else if Critere<>'' then
      BEGIN
      TypeConso:=Critere;
      END;
    Critere:=(Trim(ReadTokenSt(S)));
    END;

  // Gestion des évènements du TB
  TV1 := TTobViewer(GetControl('TV'));
  TOBCONSO := TOB.create('Ma Tob', nil, -1);

  if AffichageAuto = True then
    begin
    SetControlEnabled('BCOAFFAIRE1', false);
    SetControlEnabled('BCOAFFAIRE2', false);
    SetControlEnabled('BCOAFFAIRE3', false);
    SetControlEnabled('BEFFACEAFF1', false);
    SetControlEnabled('BSELECTAFF1', false);
    end;
  OldAutoSearch:=V_PGI.AutoSearch;
  V_PGI.AutoSearch:=AffichageAuto;
end;

procedure TOF_BTCONSOS.OnClose;
begin
  inherited;
  FreeAndNil(TOBCONSO);
  V_PGI.AutoSearch:=OldAutoSearch;
end;

procedure TOF_BTCONSOS.AjouteChampSup(var TOBTMP: TOB);
begin
  TOBTMP.addchampsup('LIBCATEGORIE', false); TOBTMP.PutValue('LIBCATEGORIE', '');
  TOBTMP.addchampsup('LIBNATURE', false); TOBTMP.PutValue('LIBNATURE', '');
end;

//******************************************************************************
//********************* Zooms / double clic  ***********************************
//******************************************************************************

procedure TOF_BTCONSOS.NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_: THEdit);
begin
  Aff0 := THEdit(GetControl('BCOAFFAIRE0'));
  Aff := THEdit(GetControl('BCOAFFAIRE'));
  Aff1 := THEdit(GetControl('BCOAFFAIRE1'));
  Aff2 := THEdit(GetControl('BCOAFFAIRE2'));
  Aff3 := THEdit(GetControl('BCOAFFAIRE3'));
  Aff4 := THEdit(GetControl('BCOAVENANT'));
end;

procedure TOF_BTCONSOS.ChargeLesConsos;
var
  Req, ReqNature: string;
  QQ: Tquery;
  TOBTMP: TOB;
  i: integer;
begin

  ReqNature:='';

  if (TypeConso = '') or (Copy(TypeConso,1,12) = 'REALISETOTAL') then
    ReqNature:='AND (BCO_NATUREMOUV="EXT" OR BCO_NATUREMOUV="MO" OR BCO_NATUREMOUV="RES") '
  else if Copy(TypeConso,1,9) = 'REALISEMO' then
    ReqNature:='AND BCO_NATUREMOUV="MO" AND BNP_TYPERESSOURCE<>"INT" '
  else if Copy(TypeConso,1,18) = 'REALISEPRESTATIONS' then
    ReqNature:='AND BCO_NATUREMOUV="EXT" OR (BCO_NATUREMOUV="MO" AND BNP_TYPERESSOURCE="INT") '
  else if Copy(TypeConso,1,16) = 'REALISEMATERIELS' then
    ReqNature:='AND BCO_NATUREMOUV="RES" ';

  // Lecture des Consos de type prestations pour le chantier concerné
  Req := 'SELECT BCO_NATUREMOUV, BNP_TYPERESSOURCE AS CATEGORIE, BNP_NATUREPRES AS NATURE, SUM(BCO_MONTANTACH) AS ACHAT,SUM(BCO_MONTANTPR) AS REVIENT, SUM(BCO_MONTANTHT) AS VENTE ' +
         'FROM CONSOMMATIONS ' +
         'LEFT OUTER JOIN ARTICLE ON GA_ARTICLE=BCO_ARTICLE ' +
         'LEFT OUTER JOIN NATUREPREST ON BNP_NATUREPRES=GA_NATUREPRES ' +
         'WHERE BCO_TRANSFORME="-" AND BCO_TRAITEVENTE="-" ' +
         'AND BCO_AFFAIRE="' + CodeAffaire + '" ' +
         'AND BCO_NATUREPIECEG IN ("","LBT","BFC","AFS","AF") ';

  if ReqNature <> '' then Req := Req + ReqNature;

  Req := Req + 'GROUP BY BCO_AFFAIRE, BCO_NATUREMOUV, BNP_TYPERESSOURCE, BNP_NATUREPRES ';
  Req := Req + 'ORDER BY BCO_AFFAIRE, BCO_NATUREMOUV, BNP_TYPERESSOURCE, BNP_NATUREPRES ';
  QQ := OpenSQL(Req, true,-1,'',true);
  TOBCONSO.LoadDetailDB('', '', '', QQ, false);

  ferme(QQ);

  ReqNature:='';
  if (TypeConso = '') or (Copy(TypeConso,1,12) = 'REALISETOTAL') Or (Copy(TypeConso,1,18) = 'REALISEFOURNITURES') then
    ReqNature:='AND BCO_NATUREMOUV="FOU" ';

  // Lecture des Consos de type articles pour le chantier concerné
  Req := 'SELECT BCO_NATUREMOUV, GA_FAMILLENIV1 AS NATURE, SUM(BCO_MONTANTACH) AS ACHAT,SUM(BCO_MONTANTPR) AS REVIENT, SUM(BCO_MONTANTHT) AS VENTE ' +
       'FROM CONSOMMATIONS ' +
       'LEFT OUTER JOIN ARTICLE ON GA_CODEARTICLE=BCO_CODEARTICLE ' +
       'WHERE BCO_AFFAIRE="' + CodeAffaire + '" ' +
       'AND BCO_NATUREPIECEG IN ("","LBT","BFC","AFS","AF") ';

  if ReqNature <> '' then Req := Req + ReqNature;

  Req := Req + 'GROUP BY BCO_AFFAIRE, BCO_NATUREMOUV, GA_FAMILLENIV1 ';
  Req := Req + 'ORDER BY BCO_AFFAIRE, BCO_NATUREMOUV, GA_FAMILLENIV1';
  QQ := OpenSQL(Req, true,-1,'',true);
  TOBCONSO.LoadDetailDB('', '', '', QQ, true);
  ferme(QQ);

  ReqNature:='';
  if (Copy(TypeConso,1,12) = 'REALISETOTAL') then
    ReqNature:='AND BCO_NATUREMOUV="FRS" '
  else if Copy(TypeConso,1,18) = 'REALISEPRESTATIONS' then
    ReqNature:='AND BCO_NATUREMOUV="FRS" AND ARS_TYPERESSOURCE="INT" '
  else if (Copy(TypeConso,1,9) = 'REALISEMO') then
    ReqNature:='AND BCO_NATUREMOUV="FRS" AND ARS_TYPERESSOURCE<>"INT" ';

  // Lecture des Consos de type frais pour le chantier concerné
  Req := 'SELECT BCO_NATUREMOUV, ARS_TYPERESSOURCE AS CATEGORIE, GA_FAMILLENIV1 AS NATURE, SUM(BCO_MONTANTACH) AS ACHAT,SUM(BCO_MONTANTPR) AS REVIENT, SUM(BCO_MONTANTHT) AS VENTE ' +
         'FROM CONSOMMATIONS ' +
         'LEFT OUTER JOIN ARTICLE ON GA_CODEARTICLE=BCO_CODEARTICLE ' +
         'LEFT OUTER JOIN RESSOURCE ON ARS_RESSOURCE=BCO_RESSOURCE ' +
         'WHERE BCO_TRANSFORME="-" AND BCO_TRAITEVENTE="-" ' +
         'AND BCO_AFFAIRE="' + CodeAffaire + '"' +
         'AND BCO_NATUREPIECEG IN ("","LBT","BFC","AFS","AF") ';

  if ReqNature <> '' then Req := Req + ReqNature;

  Req := Req + 'GROUP BY BCO_AFFAIRE, BCO_NATUREMOUV, ARS_TYPERESSOURCE, GA_FAMILLENIV1 ';
  Req := Req + 'ORDER BY BCO_AFFAIRE, BCO_NATUREMOUV, ARS_TYPERESSOURCE, GA_FAMILLENIV1';
  QQ := OpenSQL(Req, true,-1,'',true);
  TOBCONSO.LoadDetailDB('', '', '', QQ, true);
  ferme(QQ);

  TobConso.Detail.Sort('BCO_NATUREMOUV');

  // Mise à jour des libellés type de resssource et natures de prestations
  for i := 0 to TOBCONSO.Detail.Count - 1 do
  begin
    TOBTMP := TOBCONSO.Detail[i];

    if VarIsNull(TOBTMP.getValue('CATEGORIE')) or (VarAsType(TOBTMP.getValue('CATEGORIE'), varString) = #0) then
      TOBTMP.PutValue('CATEGORIE','');
    if VarIsNull(TOBTMP.getValue('NATURE')) or (VarAsType(TOBTMP.getValue('NATURE'), varString) = #0) then
      TOBTMP.PutValue('NATURE','');

    if (TOBTMP <> nil) then
    begin
      AjouteChampSup(TOBTMP);
      if (TOBTMP.GetValue('BCO_NATUREMOUV') = 'FOU') then
      begin
        // Traitement des articles
        TOBTMP.PutValue('LIBCATEGORIE', 'Fournitures');
        TOBTMP.PutValue('LIBNATURE',RechDom('GCFAMILLENIV1',TOBTMP.GetValue('NATURE'),false));
      end
      else if (TOBTMP.GetValue('BCO_NATUREMOUV') = 'FRS') AND (TOBTMP.GetValue('CATEGORIE') <> 'INT') then
      begin
        // Traitement des frais
        TOBTMP.PutValue('LIBCATEGORIE', 'Frais');
        TOBTMP.PutValue('LIBNATURE',RechDom('GCFAMILLENIV1',TOBTMP.GetValue('NATURE'),false));
      end
      else if (TOBTMP.GetValue('BCO_NATUREMOUV') = 'MO') AND (TOBTMP.GetValue('CATEGORIE') <> 'INT') then
      begin
        // Traitement des prestations de type MO interne
        TOBTMP.PutValue('LIBCATEGORIE', 'Main d''oeuvre interne');
        TOBTMP.PutValue('LIBNATURE',RechDom('BTNATPRESTATION',TOBTMP.GetValue('NATURE'),false));
      end
      else if TOBTMP.GetValue('BCO_NATUREMOUV') = 'FAN' then
      begin
        // Traitement des consommations de type frais et recettes annexes
        TOBTMP.PutValue('LIBCATEGORIE', 'Frais et Recettes Annexes');
        TOBTMP.PutValue('LIBNATURE','Frais Annexes');
      end
      else if TOBTMP.GetValue('BCO_NATUREMOUV') = 'RAN' then
      begin
        // Traitement des consommations de type frais et recettes annexes
        TOBTMP.PutValue('LIBCATEGORIE', 'Frais et Recettes Annexes');
        TOBTMP.PutValue('LIBNATURE','Recettes Annexes');
      end
      else
      begin
        // Traitement des autres prestations
        TOBTMP.PutValue('LIBCATEGORIE',RechDom('AFTTYPERESSOURCE',TOBTMP.GetValue('CATEGORIE'),false));
        TOBTMP.PutValue('LIBNATURE',RechDom('BTNATPRESTATION',TOBTMP.GetValue('NATURE'),false));
      end;
    end;
  end;

end;

initialization
  registerclasses([TOF_BTCONSOS]);
end.
