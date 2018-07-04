{***********UNITE*************************************************
Auteur  ...... : B. LOCATELLI
Créé le ...... : 05/07/2006
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTRESCONTRATS ()
Mots clefs ... : TOF;BTRESCONTRATS
*****************************************************************}
unit BTResContrats_TOF;

interface

uses StdCtrls,
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
  Ent1,
  HQry,
  HMsgBox,
  UTOF,
  Stat,
  UTOB,
  utobview,
  AGLInit,
  AGLInitBTP,
  facture,
  Btputil,
  UtilSaisieConso,
  UtilReglementAffaire,
  uTofAfBaseCodeAffaire,
  UtilTableauBord;

type

  TOF_BTRESCONTRATS = class(TOF_AFBASECODEAFFAIRE)
    procedure OnNew; override;
    procedure OnDelete; override;
    procedure OnUpdate; override;
    procedure OnLoad; override;
    procedure OnArgument(S: string); override;
    procedure OnClose; override;
    procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_: THEdit); override;
    procedure bEffaceAff1Click (Sender : TObject); override;
  private
    TV1: TTobViewer;
    TOBBTB: TOB;
    CodeAffaire0,CodeAffaire, CodeAffaire1, CodeAffaire2, CodeAffaire3,CodeAvenant: string;
    CodeClient, Responsable, Apporteur, WhereCritere: string;
    LibreAff1,LibreAff2,LibreAFF3,LibreAff4,LibreAff5,LibreAff6 : String;
    DateDebut,DateFin : TdateTime;
    procedure OnDblClickTV(Sender: TObject);
    procedure ChargeLesChantiers;
    procedure ChargeLesConsos;
    procedure EnregistreCutOffFacture (TOBINT : TOB);
    procedure AddChampsSupresCon(TOBBB: TOB);
    function CalculeCutOff(Mt: double; DateDebMvt, DateFinMvt, DateDeb,
      DateFin: TdateTime): double;
    procedure EnregistreCutOffNonFacture(TOBINT : TOB);
    procedure ChargeLesDevisAfacturer;
  	procedure ChargeLesDevisfacture;
  end;

implementation

uses DateUtils,affaireutil;

procedure TOF_BTRESCONTRATS.OnNew;
begin
  inherited;
end;

procedure TOF_BTRESCONTRATS.OnDelete;
begin
  inherited;
end;

procedure TOF_BTRESCONTRATS.OnUpdate;
begin
  inherited;

 	WhereCritere:=RecupWhereCritere(TFStat(Ecran).Pages);
  //
  TOBBTB.ClearDetail;
	ChargeLesChantiers;
	if TOBBTB <> nil then
 	begin
    //
  	ChargeLesDevisAfacturer;
  	ChargeLesDevisfacture;
    //
   	ChargeLesConsos;
    //
   	TFStat(Ecran).LaTOB := TOBBTB;
 	end;
end;

procedure TOF_BTRESCONTRATS.OnLoad;
begin
  inherited;
  CodeAffaire0 := GetControltext('AFF_AFFAIRE0');
  CodeAffaire := GetControltext('AFF_AFFAIRE');
  CodeAffaire1 := GetControltext('AFF_AFFAIRE1');
  CodeAffaire2 := GetControltext('AFF_AFFAIRE2');
  CodeAffaire3 := GetControltext('AFF_AFFAIRE3');
  CodeAvenant := GetControltext('AFF_AVENANT');
  CodeClient := GetControltext('AFF_TIERS');
  Responsable := GetControltext('AFF_RESPONSABLE');
  Apporteur := GetControltext('BTBAPPORTEUR');
  DateDebut := StrToDate(getControltext('DATEMVTDEB'));
  DateFin := StrToDate(getControltext('DATEMVTDEB_'));
  LibreAff1 := getControltext('AFF_LIBREAFF1');
  LibreAff2 := getControltext('AFF_LIBREAFF2');
  LibreAff3 := getControltext('AFF_LIBREAFF3');
  LibreAff4 := getControltext('AFF_LIBREAFF4');
  LibreAff5 := getControltext('AFF_LIBREAFF5');
  LibreAff6 := getControltext('AFF_LIBREAFF6');
end;

procedure TOF_BTRESCONTRATS.OnArgument(S: string);
Var CC : THValCombobox;
begin
  inherited;
  // Gestion des évènements du TB
  TV1 := TTobViewer(GetControl('TV'));
  TV1.OnDblClick := OnDblClickTV;
  TOBBTB := TOB.create('Ma Tob', nil, -1);
  SetControltext('AFF_AFFAIRE0','I');
  //
  // gestion Etablissement (BTP)
	CC:=THValComboBox(GetControl('BTBETABLISSEMENT')) ;
	if CC<>Nil then
  begin
  	PositionneEtabUser(CC) ;
    if not VH^.EtablisCpta then
    begin
    	if THLabel(GetControl('TBTB_ETABLISSEMENT')) <> nil then THLabel(GetControl('TBTB_ETABLISSEMENT')).Visible := false;
			CC.visible := false;
    end;
	end;

  // gestion Domaine (BTP)
	CC:=THValComboBox(GetControl('AFF_DOMAINE')) ;
	if CC<>Nil then PositionneDomaineUser(CC) ;

end;

procedure TOF_BTRESCONTRATS.OnClose;
begin
  inherited;
  FreeAndNil(TOBBTB);
end;

//******************************************************************************
//********************* Zooms / double clic  ***********************************
//******************************************************************************

procedure TOF_BTRESCONTRATS.OnDblClickTV(Sender: TObject);
begin
end;

procedure TOF_BTRESCONTRATS.NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_: THEdit);
begin
  Aff0 := THEdit(GetControl('AFF_AFFAIRE0'));
  Aff := THEdit(GetControl('AFF_AFFAIRE'));
  Aff1 := THEdit(GetControl('AFF_AFFAIRE1'));
  Aff2 := THEdit(GetControl('AFF_AFFAIRE2'));
  Aff3 := THEdit(GetControl('AFF_AFFAIRE3'));
  Aff4 := THEdit(GetControl('AFF_AVENANT'));
  Tiers := THEdit(GetControl('AFF_TIERS'));
end;

procedure TOF_BTRESCONTRATS.ChargeLesChantiers;
var
  Req,ReqIni: string;
  QQ: Tquery;
  TOBINT : TOB;
  StrInperiode,OnlyDebInPeriode,OnlyFinInPeriode : string;
begin
  // Gestion des contrats dans période ou en partie dans période
  StrInPeriode := '(AFA_DATEDEBUTFAC >= "'+USDATETime(DateDebut)+'" AND AFA_DATEFINFAC <= "'+USDATETime(DateFin)+'")';
  OnlyDebInPeriode := '(AFA_DATEDEBUTFAC >= "'+USDATETime(DateDebut)+'"'+
  										 ' AND AFA_DATEDEBUTFAC <= "'+USDATETime(DateFin)+'"'+
                       ' AND AFA_DATEFINFAC > "'+USDATETime(DateFin)+'")';
  OnlyFinInperiode := '(AFA_DATEDEBUTFAC < "'+USDATETime(DateDebut)+'"'+
  										 ' AND AFA_DATEFINFAC <= "'+USDATETime(DateFin)+'"'+
                       ' AND AFA_DATEFINFAC >= "'+USDATETime(DateDebut)+'")';
  //
  TOBINT := TOB.Create ('LES CONTRATS',nil,-1);
  ReqIni := TFStat(Ecran).stSQL;
(*
  Req := 'SELECT AFF_AFFAIRE AS CODEAFFAIRE, AFF_LIBELLE AS LIBELLEAFFAIRE, '+
         'AFA_DATEECHE AS DATE, AFA_MONTANTECHE AS MONTANTFACTURE, '+
         '0.0 AS Montant_a_facturer, AFA_LIBELLEECHE AS DESIGNATION, '+
         '"" AS APPEL, "" AS CODEARTICLE, 0.0 AS QUANTITE, 0.0 AS COUT, '+
         '0.0 AS FACTURE, 0.0 AS A_FACTURER, 0.0 AS QTEFACTUREE, '+
         '"" AS RESSOURCE, AFA_TIERS + " " +(SELECT T_LIBELLE FROM TIERS WHERE T_TIERS=AFA_TIERS AND T_NATUREAUXI="CLI") AS CLIENT, AFA_DATEDEBUTFAC, AFA_DATEFINFAC '+
         'FROM AFFAIRE, FACTAFF '+
         'WHERE AFA_AFFAIRE=AFF_AFFAIRE '+
         'AND AFA_ECHEFACT="X"';
*)
  req := ReqIni + ' AND AFA_ECHEFACT="X" ';
  req := Req + 'AND ('+StrInperiode+' OR '+OnlYFinINPeriode+' OR '+OnlyDebInPeriode+') ';

  if CodeAffaire <> '' then
  begin
    Req := Req + ' AND AFF_AFFAIRE LIKE "' + CodeAffaire + '%"';
  end else
  begin
    Req := Req + ' AND AFF_AFFAIRE LIKE "I%"';
    if CodeAffaire1 <> '' then
      Req := Req + ' AND AFF_AFFAIRE1 = "' + CodeAffaire1 + '"';
    if CodeAffaire2 <> '' then
      Req := Req + ' AND AFF_AFFAIRE2 = "' + CodeAffaire2 + '"';
    if CodeAffaire3 <> '' then
      Req := Req + ' AND AFF_AFFAIRE3 = "' + CodeAffaire3 + '"';
  end;

  if CodeClient <> '' then
  begin
    Req := Req + ' AND AFF_TIERS = "' + CodeClient + '"';
  end;

  if Responsable <> '' then
  begin
    Req := Req + ' AND AFF_RESPONSABLE = "' + Responsable + '"';
  end;

  if Apporteur <> '' then
  begin
    Req := Req + ' AND AFF_APPORTEUR = "' + Apporteur + '"';
  end;

//
  if LibreAff1 <> '' then
  begin
    Req := Req + ' AND AFF_LIBREAFF1 = "' + LibreAff1 + '"';
  end;
  if LibreAff2 <> '' then
  begin
    Req := Req + ' AND AFF_LIBREAFF2 = "' + LibreAff2 + '"';
  end;
  if LibreAff3 <> '' then
  begin
    Req := Req + ' AND AFF_LIBREAFF3 = "' + LibreAff3 + '"';
  end;
  if LibreAff4 <> '' then
  begin
    Req := Req + ' AND AFF_LIBREAFF4 = "' + LibreAff4 + '"';
  end;
  if LibreAff5 <> '' then
  begin
    Req := Req + ' AND AFF_LIBREAFF5 = "' + LibreAff5 + '"';
  end;
  if LibreAff6 <> '' then
  begin
    Req := Req + ' AND AFF_LIBREAFF6 = "' + LibreAff6 + '"';
  end;
//




  QQ := OpenSQL(Req, true);
  TOBINT.LoadDetailDB('', '', '', QQ, false);
  ferme(QQ);

  EnregistreCutOffFacture (TOBINT);

  TOBINT.ClearDetail;
  (*
  Req := 'SELECT AFF_AFFAIRE AS CODEAFFAIRE, AFF_LIBELLE AS LIBELLEAFFAIRE, '+
         'AFA_DATEECHE AS DATE, 0.0 AS MONTANTFACTURE, '+
         'AFA_MONTANTECHE AS Montant_a_facturer, AFA_LIBELLEECHE AS DESIGNATION, '+
         '"" AS APPEL, "" AS CODEARTICLE, 0.0 AS QUANTITE, 0.0 AS COUT, '+
         '0.0 AS FACTURE, 0.0 AS A_FACTURER, 0.0 AS QTEFACTUREE, '+
         '"" AS RESSOURCE, AFA_TIERS + " " +(SELECT T_LIBELLE FROM TIERS WHERE T_TIERS=AFA_TIERS AND T_NATUREAUXI="CLI") AS CLIENT, AFA_DATEDEBUTFAC, AFA_DATEFINFAC '+
         'FROM AFFAIRE, FACTAFF '+
         'WHERE AFA_AFFAIRE=AFF_AFFAIRE AND AFA_ECHEFACT="-"';
  *)
  req := ReqIni + ' AND AFA_ECHEFACT="-" ';
  req := req + 'AND ('+StrInperiode+' OR '+OnlYFinINPeriode+' OR '+OnlyDebInPeriode+') ';

  if CodeAffaire <> '' then
  begin
    Req := Req + ' AND AFF_AFFAIRE LIKE "' + CodeAffaire + '%"';
  end
  else
  begin
    Req := Req + ' AND AFF_AFFAIRE LIKE "I%"';
    if CodeAffaire1 <> '' then
      Req := Req + ' AND AFF_AFFAIRE1 = "' + CodeAffaire1 + '"';
    if CodeAffaire2 <> '' then
      Req := Req + ' AND AFF_AFFAIRE2 = "' + CodeAffaire2 + '"';
    if CodeAffaire3 <> '' then
      Req := Req + ' AND AFF_AFFAIRE3 = "' + CodeAffaire3 + '"';
  end;

  if CodeClient <> '' then
  begin
    Req := Req + ' AND AFF_TIERS = "' + CodeClient + '"';
  end;

  if Responsable <> '' then
  begin
    Req := Req + ' AND AFF_RESPONSABLE = "' + Responsable + '"';
  end;

  if Apporteur <> '' then
  begin
    Req := Req + ' AND AFF_APPORTEUR = "' + Apporteur + '"';
  end;

//
  if LibreAff1 <> '' then
  begin
    Req := Req + ' AND AFF_LIBREAFF1 = "' + LibreAff1 + '"';
  end;
  if LibreAff2 <> '' then
  begin
    Req := Req + ' AND AFF_LIBREAFF2 = "' + LibreAff2 + '"';
  end;
  if LibreAff3 <> '' then
  begin
    Req := Req + ' AND AFF_LIBREAFF3 = "' + LibreAff3 + '"';
  end;
  if LibreAff4 <> '' then
  begin
    Req := Req + ' AND AFF_LIBREAFF4 = "' + LibreAff4 + '"';
  end;
  if LibreAff5 <> '' then
  begin
    Req := Req + ' AND AFF_LIBREAFF5 = "' + LibreAff5 + '"';
  end;
  if LibreAff6 <> '' then
  begin
    Req := Req + ' AND AFF_LIBREAFF6 = "' + LibreAff6 + '"';
  end;
//

  QQ := OpenSQL(Req, true);
  TOBINT.LoadDetailDB('', '', '', QQ, True);
  ferme(QQ);
  EnregistreCutOffNonFacture (TOBINT);

  TOBINT.free;
end;

// Récupération les lignes de consommations associées au contrat
// soit directement saisies sur le contrat soit saisies sur des appels liés au contrat
procedure TOF_BTRESCONTRATS.ChargeLesConsos;
var
  Req: string;
  QQ: Tquery;
  StrInPeriode,CodeAffaireSaisie : string;
begin
  // ---- Correction LS
  if CodeAffaire <> '' then
  begin
    CodeAffaireSaisie := CodeAffaire;
  end else
  begin
  	CodeAffaireSaisie := ConstitueCodeAffaireSql ('I',CodeAffaire1,CodeAffaire2,CodeAffaire3,CodeAvenant);
  end;

	StrInperiode :=' AND BCO_DATEMOUV >="'+USDATETime(Datedebut)+'" AND BCO_DATEMOUV <="'+USDATETime(DateFin)+'"';

// Lignes non facturables saisies directement sur contrat
  Req := 'SELECT AFF_AFFAIRE AS CODEAFFAIRE, AFF_LIBELLE AS LIBELLEAFFAIRE, '+
         'BCO_DATEMOUV AS DATE, 0.0 AS MONTANTFACTURE, 0.0 AS MONTANT_A_FACTURER, '+
         '"" AS APPEL, BCO_LIBELLE AS DESIGNATION, BCO_CODEARTICLE AS CODEARTICLE, BCO_QUANTITE AS QUANTITE, '+
         'BCO_MONTANTPR AS COUT, 0.0 AS FACTURE, 0.0 AS A_FACTURER, 0.0 AS QTEFACTUREE, ' +
         'BCO_RESSOURCE AS RESSOURCE, '+
         'AFF_TIERS,(SELECT T_LIBELLE FROM TIERS WHERE T_TIERS=AFF_TIERS AND T_NATUREAUXI="CLI") AS CLIENT '+
         'FROM CONSOMMATIONS ' +
         'LEFT JOIN AFFAIRE ON AFF_AFFAIRE=BCO_AFFAIRE ' +
         'WHERE BCO_FACTURABLE="N" ' +
         'AND (BCO_NATUREPIECEG="" OR BCO_NATUREPIECEG="LBT") ';

  req := Req + ' AND ('+StrInperiode+') ';

  if CodeAffaire <> '' then
  begin
    Req := Req + ' AND BCO_AFFAIRE LIKE "' + CodeAffaire + '%"';
  end
  else
  begin
    Req := Req + ' AND BCO_AFFAIRE LIKE "I%"';
    if CodeAffaire1 <> '' then
      Req := Req + ' AND BCO_AFFAIRE1 = "' + CodeAffaire1 + '"';
    if CodeAffaire2 <> '' then
      Req := Req + ' AND BCO_AFFAIRE2 = "' + CodeAffaire2 + '"';
    if CodeAffaire3 <> '' then
      Req := Req + ' AND BCO_AFFAIRE3 = "' + CodeAffaire3 + '"';
  end;

  if CodeClient <> '' then
  begin
    Req := Req + ' AND AFF_TIERS = "' + CodeClient + '"';
  end;

  if Responsable <> '' then
  begin
    Req := Req + ' AND AFF_RESPONSABLE = "' + Responsable + '"';
  end;

  if Apporteur <> '' then
  begin
    Req := Req + ' AND AFF_APPORTEUR = "' + Apporteur + '"';
  end;
//
  if LibreAff1 <> '' then
  begin
    Req := Req + ' AND AFF_LIBREAFF1 = "' + LibreAff1 + '"';
  end;
  if LibreAff2 <> '' then
  begin
    Req := Req + ' AND AFF_LIBREAFF2 = "' + LibreAff2 + '"';
  end;
  if LibreAff3 <> '' then
  begin
    Req := Req + ' AND AFF_LIBREAFF3 = "' + LibreAff3 + '"';
  end;
  if LibreAff4 <> '' then
  begin
    Req := Req + ' AND AFF_LIBREAFF4 = "' + LibreAff4 + '"';
  end;
  if LibreAff5 <> '' then
  begin
    Req := Req + ' AND AFF_LIBREAFF5 = "' + LibreAff5 + '"';
  end;
  if LibreAff6 <> '' then
  begin
    Req := Req + ' AND AFF_LIBREAFF6 = "' + LibreAff6 + '"';
  end;
//
  QQ := OpenSQL(Req, true);
 	TOBBTB.LoadDetailDB('', '', '', QQ, True);
  ferme(QQ);

// Lignes facturées saisies directement sur contrat
  Req := 'SELECT AFF_AFFAIRE AS CODEAFFAIRE, AFF_LIBELLE AS LIBELLEAFFAIRE, '+
         'BCO_DATEMOUV AS DATE, 0.0 AS MONTANTFACTURE, 0.0 AS MONTANT_A_FACTURER, '+
         '"" AS APPEL, BCO_LIBELLE AS DESIGNATION, BCO_CODEARTICLE AS CODEARTICLE, BCO_QUANTITE AS QUANTITE, '+
         'BCO_MONTANTPR AS COUT, BCO_MONTANTHT AS FACTURE, 0.0 AS A_FACTURER, BCO_QTEFACTUREE AS QTEFACTUREE, ' +
         'BCO_RESSOURCE AS RESSOURCE, '+
         'AFF_TIERS ,(SELECT T_LIBELLE FROM TIERS WHERE T_TIERS=AFF_TIERS AND T_NATUREAUXI="CLI") AS CLIENT '+
         'FROM CONSOMMATIONS ' +
         'LEFT JOIN AFFAIRE ON AFF_AFFAIRE=BCO_AFFAIRE ' +
         'WHERE BCO_FACTURABLE="F" ';

  req := Req + ' AND ('+StrInperiode+') ';

  if CodeAffaire <> '' then
  begin
    Req := Req + ' AND BCO_AFFAIRE LIKE "' + CodeAffaire + '%"';
  end
  else
  begin
    Req := Req + ' AND BCO_AFFAIRE LIKE "I%"';
    if CodeAffaire1 <> '' then
      Req := Req + ' AND BCO_AFFAIRE1 = "' + CodeAffaire1 + '"';
    if CodeAffaire2 <> '' then
      Req := Req + ' AND BCO_AFFAIRE2 = "' + CodeAffaire2 + '"';
    if CodeAffaire3 <> '' then
      Req := Req + ' AND BCO_AFFAIRE3 = "' + CodeAffaire3 + '"';
  end;

  if CodeClient <> '' then
  begin
    Req := Req + ' AND AFF_TIERS = "' + CodeClient + '"';
  end;

  if Responsable <> '' then
  begin
    Req := Req + ' AND AFF_RESPONSABLE = "' + Responsable + '"';
  end;

  if Apporteur <> '' then
  begin
    Req := Req + ' AND AFF_APPORTEUR = "' + Apporteur + '"';
  end;

//
  if LibreAff1 <> '' then
  begin
    Req := Req + ' AND AFF_LIBREAFF1 = "' + LibreAff1 + '"';
  end;
  if LibreAff2 <> '' then
  begin
    Req := Req + ' AND AFF_LIBREAFF2 = "' + LibreAff2 + '"';
  end;
  if LibreAff3 <> '' then
  begin
    Req := Req + ' AND AFF_LIBREAFF3 = "' + LibreAff3 + '"';
  end;
  if LibreAff4 <> '' then
  begin
    Req := Req + ' AND AFF_LIBREAFF4 = "' + LibreAff4 + '"';
  end;
  if LibreAff5 <> '' then
  begin
    Req := Req + ' AND AFF_LIBREAFF5 = "' + LibreAff5 + '"';
  end;
  if LibreAff6 <> '' then
  begin
    Req := Req + ' AND AFF_LIBREAFF6 = "' + LibreAff6 + '"';
  end;
//
  QQ := OpenSQL(Req, true);
 	TOBBTB.LoadDetailDB('', '', '', QQ, True);
  ferme(QQ);

// Lignes à facturer saisies directement sur contrat
  Req := 'SELECT AFF_AFFAIRE AS CODEAFFAIRE, AFF_LIBELLE AS LIBELLEAFFAIRE, '+
         'BCO_DATEMOUV AS DATE, 0.0 AS MONTANTFACTURE, 0.0 AS MONTANT_A_FACTURER, '+
         '"" AS APPEL, BCO_LIBELLE AS DESIGNATION, BCO_CODEARTICLE AS CODEARTICLE, BCO_QUANTITE AS QUANTITE, '+
         'BCO_MONTANTPR AS COUT, 0.0 AS FACTURE, BCO_MONTANTHT AS A_FACTURER, BCO_QTEFACTUREE AS QTEFACTUREE, ' +
         'BCO_RESSOURCE AS RESSOURCE, '+
				 'AFF_TIERS, (SELECT T_LIBELLE FROM TIERS WHERE T_TIERS=AFF_TIERS AND T_NATUREAUXI="CLI") AS CLIENT '+
         'FROM CONSOMMATIONS ' +
         'LEFT JOIN AFFAIRE ON AFF_AFFAIRE=BCO_AFFAIRE ' +
         'WHERE BCO_FACTURABLE="A" ' +
         'AND (BCO_NATUREPIECEG="" OR BCO_NATUREPIECEG="LBT") ';

  req := Req + ' AND ('+StrInperiode+') ';

  if CodeAffaire <> '' then
  begin
    Req := Req + ' AND BCO_AFFAIRE LIKE "' + CodeAffaire + '%"';
  end
  else
  begin
    Req := Req + ' AND BCO_AFFAIRE LIKE "I%"';
    if CodeAffaire1 <> '' then
      Req := Req + ' AND BCO_AFFAIRE1 = "' + CodeAffaire1 + '"';
    if CodeAffaire2 <> '' then
      Req := Req + ' AND BCO_AFFAIRE2 = "' + CodeAffaire2 + '"';
    if CodeAffaire3 <> '' then
      Req := Req + ' AND BCO_AFFAIRE3 = "' + CodeAffaire3 + '"';
  end;

  if CodeClient <> '' then
  begin
    Req := Req + ' AND AFF_TIERS = "' + CodeClient + '"';
  end;

  if Responsable <> '' then
  begin
    Req := Req + ' AND AFF_RESPONSABLE = "' + Responsable + '"';
  end;

  if Apporteur <> '' then
  begin
    Req := Req + ' AND AFF_APPORTEUR = "' + Apporteur + '"';
  end;

//
  if LibreAff1 <> '' then
  begin
    Req := Req + ' AND AFF_LIBREAFF1 = "' + LibreAff1 + '"';
  end;
  if LibreAff2 <> '' then
  begin
    Req := Req + ' AND AFF_LIBREAFF2 = "' + LibreAff2 + '"';
  end;
  if LibreAff3 <> '' then
  begin
    Req := Req + ' AND AFF_LIBREAFF3 = "' + LibreAff3 + '"';
  end;
  if LibreAff4 <> '' then
  begin
    Req := Req + ' AND AFF_LIBREAFF4 = "' + LibreAff4 + '"';
  end;
  if LibreAff5 <> '' then
  begin
    Req := Req + ' AND AFF_LIBREAFF5 = "' + LibreAff5 + '"';
  end;
  if LibreAff6 <> '' then
  begin
    Req := Req + ' AND AFF_LIBREAFF6 = "' + LibreAff6 + '"';
  end;
//
  QQ := OpenSQL(Req, true);
 	TOBBTB.LoadDetailDB('', '', '', QQ, True);
  ferme(QQ);


// Lignes non facturables saisies sur appels liés au contrat
  Req := 'SELECT AFF_AFFAIRE AS CODEAFFAIRE, AFF_LIBELLE AS LIBELLEAFFAIRE, '+
         'BCO_DATEMOUV AS DATE, 0.0 AS MONTANTFACTURE, 0.0 AS MONTANT_A_FACTURER, '+
         'BCO_AFFAIRE AS APPEL, BCO_LIBELLE AS DESIGNATION, BCO_CODEARTICLE AS CODEARTICLE, '+
         'BCO_QUANTITE AS QUANTITE, BCO_MONTANTPR AS COUT, '+
         '0.0 AS FACTURE, 0.0 AS A_FACTURER, 0.0 AS QTEFACTUREE, ' +
         'BCO_RESSOURCE AS RESSOURCE, '+
				 'AFF_TIERS, (SELECT T_LIBELLE FROM TIERS WHERE T_TIERS=AFF_TIERS AND T_NATUREAUXI="CLI") AS CLIENT '+
         'FROM CONSOMMATIONS ' +
         'LEFT JOIN AFFAIRE ON AFF_AFFAIRE=BCO_AFFAIRESAISIE ' +
         'WHERE BCO_FACTURABLE="N" ' +
         'AND (BCO_NATUREPIECEG="" OR BCO_NATUREPIECEG="LBT") ';

  req := Req + ' AND ('+StrInperiode+') ';

  if CodeAffaire <> '' then
  begin
    Req := Req + ' AND BCO_AFFAIRESAISIE LIKE "%' + CodeAffaire + '%"';
  end
  else
  begin
    Req := Req + ' AND BCO_AFFAIRESAISIE LIKE "'+CodeAffaireSaisie+'"';
  end;

  if CodeClient <> '' then
  begin
    Req := Req + ' AND AFF_TIERS = "' + CodeClient + '"';
  end;

  if Responsable <> '' then
  begin
    Req := Req + ' AND AFF_RESPONSABLE = "' + Responsable + '"';
  end;

  if Apporteur <> '' then
  begin
    Req := Req + ' AND AFF_APPORTEUR = "' + Apporteur + '"';
  end;

//
  if LibreAff1 <> '' then
  begin
    Req := Req + ' AND AFF_LIBREAFF1 = "' + LibreAff1 + '"';
  end;
  if LibreAff2 <> '' then
  begin
    Req := Req + ' AND AFF_LIBREAFF2 = "' + LibreAff2 + '"';
  end;
  if LibreAff3 <> '' then
  begin
    Req := Req + ' AND AFF_LIBREAFF3 = "' + LibreAff3 + '"';
  end;
  if LibreAff4 <> '' then
  begin
    Req := Req + ' AND AFF_LIBREAFF4 = "' + LibreAff4 + '"';
  end;
  if LibreAff5 <> '' then
  begin
    Req := Req + ' AND AFF_LIBREAFF5 = "' + LibreAff5 + '"';
  end;
  if LibreAff6 <> '' then
  begin
    Req := Req + ' AND AFF_LIBREAFF6 = "' + LibreAff6 + '"';
  end;
//
  QQ := OpenSQL(Req, true);
 	TOBBTB.LoadDetailDB('', '', '', QQ, True);
  ferme(QQ);

// Lignes facturées saisies sur appels liés au contrat
  Req := 'SELECT AFF_AFFAIRE AS CODEAFFAIRE, AFF_LIBELLE AS LIBELLEAFFAIRE, '+
         'BCO_DATEMOUV AS DATE, 0.0 AS MONTANTFACTURE, 0.0 AS MONTANT_A_FACTURER, '+
         'BCO_AFFAIRE AS APPEL, BCO_LIBELLE AS DESIGNATION, BCO_CODEARTICLE AS CODEARTICLE, '+
         'BCO_QUANTITE AS QUANTITE, BCO_MONTANTPR AS COUT, '+
         'BCO_MONTANTHT AS FACTURE, 0.0 AS A_FACTURER, BCO_QTEFACTUREE AS QTEFACTUREE, ' +
         'BCO_RESSOURCE AS RESSOURCE, '+
         'AFF_TIERS ,(SELECT T_LIBELLE FROM TIERS WHERE T_TIERS=AFF_TIERS AND T_NATUREAUXI="CLI") AS CLIENT '+
         'FROM CONSOMMATIONS ' +
         'LEFT JOIN AFFAIRE ON AFF_AFFAIRE=BCO_AFFAIRESAISIE ' +
         'WHERE BCO_FACTURABLE="F" ' ;

  req := Req + ' AND ('+StrInperiode+') ';

  if CodeAffaire <> '' then
  begin
    Req := Req + ' AND BCO_AFFAIRESAISIE LIKE "%' + CodeAffaire + '%"';
  end
  else
  begin
    Req := Req + ' AND BCO_AFFAIRESAISIE LIKE "'+CodeAffaireSaisie+'"';
  end;

  if CodeClient <> '' then
  begin
    Req := Req + ' AND AFF_TIERS = "' + CodeClient + '"';
  end;

  if Responsable <> '' then
  begin
    Req := Req + ' AND AFF_RESPONSABLE = "' + Responsable + '"';
  end;

  if Apporteur <> '' then
  begin
    Req := Req + ' AND AFF_APPORTEUR = "' + Apporteur + '"';
  end;

//
  if LibreAff1 <> '' then
  begin
    Req := Req + ' AND AFF_LIBREAFF1 = "' + LibreAff1 + '"';
  end;
  if LibreAff2 <> '' then
  begin
    Req := Req + ' AND AFF_LIBREAFF2 = "' + LibreAff2 + '"';
  end;
  if LibreAff3 <> '' then
  begin
    Req := Req + ' AND AFF_LIBREAFF3 = "' + LibreAff3 + '"';
  end;
  if LibreAff4 <> '' then
  begin
    Req := Req + ' AND AFF_LIBREAFF4 = "' + LibreAff4 + '"';
  end;
  if LibreAff5 <> '' then
  begin
    Req := Req + ' AND AFF_LIBREAFF5 = "' + LibreAff5 + '"';
  end;
  if LibreAff6 <> '' then
  begin
    Req := Req + ' AND AFF_LIBREAFF6 = "' + LibreAff6 + '"';
  end;
//
  QQ := OpenSQL(Req, true);
 	TOBBTB.LoadDetailDB('', '', '', QQ, True);
  ferme(QQ);

// Lignes à facturer saisies sur appels liés au contrat
  Req := 'SELECT AFF_AFFAIRE AS CODEAFFAIRE, AFF_LIBELLE AS LIBELLEAFFAIRE, '+
         'BCO_DATEMOUV AS DATE, 0.0 AS MONTANTFACTURE, 0.0 AS MONTANT_A_FACTURER, '+
         'BCO_AFFAIRE AS APPEL, BCO_LIBELLE AS DESIGNATION, BCO_CODEARTICLE AS CODEARTICLE, '+
         'BCO_QUANTITE AS QUANTITE, BCO_MONTANTPR AS COUT, '+
         '0.0 AS FACTURE, BCO_MONTANTHT AS A_FACTURER, BCO_QTEFACTUREE AS QTEFACTUREE, ' +
         'BCO_RESSOURCE AS RESSOURCE, '+
         'AFF_TIERS, (SELECT T_LIBELLE FROM TIERS WHERE T_TIERS=AFF_TIERS AND T_NATUREAUXI="CLI") AS CLIENT '+
         'FROM CONSOMMATIONS ' +
         'LEFT JOIN AFFAIRE ON AFF_AFFAIRE=BCO_AFFAIRESAISIE ' +
         'WHERE BCO_FACTURABLE="A" ' +
         'AND (BCO_NATUREPIECEG="" OR BCO_NATUREPIECEG="LBT") ';

  req := Req + ' AND ('+StrInperiode+') ';

  if CodeAffaire <> '' then
  begin
    Req := Req + ' AND BCO_AFFAIRESAISIE LIKE "%' + CodeAffaire + '%"';
  end
  else
  begin
    Req := Req + ' AND BCO_AFFAIRESAISIE LIKE "'+CodeAffaireSaisie+'"';
  end;

  if CodeClient <> '' then
  begin
    Req := Req + ' AND AFF_TIERS = "' + CodeClient + '"';
  end;

  if Responsable <> '' then
  begin
    Req := Req + ' AND AFF_RESPONSABLE = "' + Responsable + '"';
  end;

  if Apporteur <> '' then
  begin
    Req := Req + ' AND AFF_APPORTEUR = "' + Apporteur + '"';
  end;

//
  if LibreAff1 <> '' then
  begin
    Req := Req + ' AND AFF_LIBREAFF1 = "' + LibreAff1 + '"';
  end;
  if LibreAff2 <> '' then
  begin
    Req := Req + ' AND AFF_LIBREAFF2 = "' + LibreAff2 + '"';
  end;
  if LibreAff3 <> '' then
  begin
    Req := Req + ' AND AFF_LIBREAFF3 = "' + LibreAff3 + '"';
  end;
  if LibreAff4 <> '' then
  begin
    Req := Req + ' AND AFF_LIBREAFF4 = "' + LibreAff4 + '"';
  end;
  if LibreAff5 <> '' then
  begin
    Req := Req + ' AND AFF_LIBREAFF5 = "' + LibreAff5 + '"';
  end;
  if LibreAff6 <> '' then
  begin
    Req := Req + ' AND AFF_LIBREAFF6 = "' + LibreAff6 + '"';
  end;
//
  QQ := OpenSQL(Req, true);
 	TOBBTB.LoadDetailDB('', '', '', QQ, True);
  ferme(QQ);
end;

procedure TOF_BTRESCONTRATS.AddChampsSupresCon(TOBBB : TOB);
begin
	TOBBB.AddChampSupValeur ('CODEAFFAIRE','');
	TOBBB.AddChampSupValeur ('LIBELLEAFFAIRE','');
	TOBBB.AddChampSupValeur ('DATE',Idate1900);
	TOBBB.AddChampSupValeur ('MONTANTFACTURE',0.0);
	TOBBB.AddChampSupValeur ('MONTANT_A_FACTURER',0.0);  
	TOBBB.AddChampSupValeur ('DESIGNATION','');
	TOBBB.AddChampSupValeur ('APPEL','');
	TOBBB.AddChampSupValeur ('CODEARTICLE','');
	TOBBB.AddChampSupValeur ('QUANTITE',0.0);
	TOBBB.AddChampSupValeur ('COUT',0.0);
	TOBBB.AddChampSupValeur ('FACTURE',0.0);
	TOBBB.AddChampSupValeur ('A_FACTURER',0.0);
	TOBBB.AddChampSupValeur ('QTEFACTUREE',0.0);
	TOBBB.AddChampSupValeur ('RESSOURCE','');
	TOBBB.AddChampSupValeur ('CLIENT','');
end;

function TOF_BTRESCONTRATS.CalculeCutOff (Mt : double;DateDebMvt,DateFinMvt,DateDeb,DateFin:TdateTime) : double;
var NbJoursTot,NbJoursDansPeriode : double;
begin
	result := Mt;
	NbJoursTot := DaysBetween(DateDebMvt,DateFinMvt);
  if (DateDebMvt>= DateDeb) and (DateFinMvt <= DateFin) then
  begin
  	exit;
  end else if (DateFinMvt >= DateDeb) and (DateFinMvt <= DateFin) and (DateDebMvt < DateDeb) then
  begin
  	// Fin de la période de facturation dans les bornes de prise en charge pas pas le début
  	NbJoursDansPeriode := arrondi(DaySpan  (DateDeb,DateFinMvt),0) ;
  	result := Arrondi(Mt * (NbJoursDansPeriode / NbJoursTot),V_PGI.OkDecV);
  end else if (DateDebMvt >= DateDeb) and (DateDebMvt <= DateFin) and (DateFinMvt > DateFin) then
  begin
  	// Fin de début de la période de facturation dans les bornes de prise en charge pas pas la fin
  	NbJoursDansPeriode := DaysBetween (DateDebmvt,DateFin) ;
  	result := Arrondi(Mt * (NbJoursDansPeriode / NbJoursTot),V_PGI.OkDecV);
  end;
end;

procedure TOF_BTRESCONTRATS.EnregistreCutOffFacture(TOBINT: TOB);
var indice : integer;
		TOBI,TOBBB : TOB;
    MtReel : double;
begin
	for Indice := 0 to TOBINT.detail.count -1 do
  begin
    TOBI := TOBINT.detail[Indice];
    MtReel := CalculeCutOff (TOBI.getValue('MONTANTFACTURE'),
    												 TOBI.getValue('AFA_DATEDEBUTFAC'),
    												 TOBI.getValue('AFA_DATEFINFAC'),
                             DateDebut,DateFin);
    TOBBB := TOB.Create ('',TOBBTB,-1);
    AddChampsSupresCon(TOBBB);
    TOBBB.PutValue('CODEAFFAIRE',TOBI.getValue('CODEAFFAIRE'));
    TOBBB.PutValue('LIBELLEAFFAIRE',TOBI.getValue('LIBELLEAFFAIRE'));
    TOBBB.PutValue('DATE',TOBI.getValue('DATE'));
    TOBBB.PutValue('MONTANTFACTURE',MtReel);
    TOBBB.PutValue('CLIENT',TOBI.getValue('CLIENT'));
    TOBBB.PutValue('DESIGNATION',TOBI.getValue('DESIGNATION'));
  end;
end;

procedure TOF_BTRESCONTRATS.EnregistreCutOffNonFacture (TOBINT : TOB);
var indice : integer;
		TOBI,TOBBB : TOB;
    MtReel : double;
begin
	for Indice := 0 to TOBINT.detail.count -1 do
  begin
    TOBI := TOBINT.detail[Indice];
    MtReel := CalculeCutOff (TOBI.getValue('MONTANT_A_FACTURER'),
    												 TOBI.getValue('AFA_DATEDEBUTFAC'),
    												 TOBI.getValue('AFA_DATEFINFAC'),
                             DateDebut,DateFin);
    TOBBB := TOB.Create ('',TOBBTB,-1);
    AddChampsSupresCon(TOBBB);
    TOBBB.PutValue('CODEAFFAIRE',TOBI.getValue('CODEAFFAIRE'));
    TOBBB.PutValue('LIBELLEAFFAIRE',TOBI.getValue('LIBELLEAFFAIRE'));
    TOBBB.PutValue('DATE',TOBI.getValue('DATE'));
    TOBBB.PutValue('MONTANT_A_FACTURER',MtReel);
    TOBBB.PutValue('DESIGNATION',TOBI.getValue('DESIGNATION'));
    TOBBB.PutValue('CLIENT',TOBI.getValue('CLIENT'));
  end;
end;



procedure TOF_BTRESCONTRATS.ChargeLesDevisAfacturer;
var SQL : string;
		QQ : TQuery;
    TOBINT,TOBI,TOBBB : TOB;
    indice : integer;
begin
	TOBINT := TOB.Create ('LES DEVIS NON FACTURES',nil,-1);

	Sql := 'SELECT GP_AFFAIRE,GP_TOTALHTDEV,GP_MONTANTPR,A1.AFF_AFFAIREINIT,A1.AFF_DATEFIN,A1.AFF_LIBELLE,'+
  			 'GP_TIERS, (SELECT T_LIBELLE FROM TIERS WHERE T_TIERS=GP_TIERS AND T_NATUREAUXI="CLI") AS CLIENT '+
  			 ' FROM PIECE '+
  			 'LEFT JOIN AFFAIRE A1 ON A1.AFF_AFFAIRE=GP_AFFAIRE '+
         'LEFT JOIN AFFAIRE A2 ON A2.AFF_AFFAIRE=A1.AFF_AFFAIREINIT '+
         'WHERE GP_NATUREPIECEG="DAP" ';
	if CodeAffaire <> '' then
  begin
  	Sql := Sql + 'AND A2.AFF_AFFAIRE LIKE "' + CodeAffaire + '%" ';
  end else
  begin
    Sql := Sql + 'AND A2.AFF_AFFAIRE LIKE "I%" ';
    if CodeAffaire1 <> '' then
      Sql := Sql + 'AND A2.AFF_AFFAIRE1 = "' + CodeAffaire1 + '" ';
    if CodeAffaire2 <> '' then
      Sql := Sql + 'AND A2.AFF_AFFAIRE2 = "' + CodeAffaire2 + '" ';
    if CodeAffaire3 <> '' then
      Sql := Sql + 'AND A2.AFF_AFFAIRE3 = "' + CodeAffaire3 + '" ';
  end;
  if CodeClient <> '' then
  begin
    Sql := Sql + 'AND A2.AFF_TIERS = "' + CodeClient + '" ';
  end;
  if Responsable <> '' then
  begin
    Sql := Sql + 'AND A2.AFF_RESPONSABLE = "' + Responsable + '" ';
  end;
  if Apporteur <> '' then
  begin
    Sql := Sql + ' AND A2.AFF_APPORTEUR = "' + Apporteur + '"';
  end;
//
  if LibreAff1 <> '' then
  begin
    Sql := Sql + ' AND A2.AFF_LIBREAFF1 = "' + LibreAff1 + '"';
  end;
  if LibreAff2 <> '' then
  begin
    Sql := Sql + ' AND A2.AFF_LIBREAFF2 = "' + LibreAff2 + '"';
  end;
  if LibreAff3 <> '' then
  begin
    Sql := Sql + ' AND A2.AFF_LIBREAFF3 = "' + LibreAff3 + '"';
  end;
  if LibreAff4 <> '' then
  begin
    Sql := Sql + ' AND A2.AFF_LIBREAFF4 = "' + LibreAff4 + '"';
  end;
  if LibreAff5 <> '' then
  begin
    Sql := Sql + ' AND A2.AFF_LIBREAFF5 = "' + LibreAff5 + '"';
  end;
  if LibreAff6 <> '' then
  begin
    Sql := Sql + ' AND A2.AFF_LIBREAFF6 = "' + LibreAff6 + '"';
  end;
//
  Sql := Sql + 'AND GP_VIVANTE="X" AND A1.AFF_ETATAFFAIRE IN ("REA","AFF") ';
  Sql := Sql + 'AND A1.AFF_DATESIGNE >= "'+USDATETIME(DateDebut)+'" AND A1.AFF_DATESIGNE <= "'+USDATETIME(DateFin)+'"';
  TRY
    QQ := OpenSql (Sql,true,-1,'',True);
    if not QQ.eof then
    begin
    	TOBINT.LoadDetailDB('','','',QQ,false);
      for Indice := 0 to TOBINT.detail.count -1 do
      begin
      	TOBI := TOBINT.detail[Indice];
        TOBBB := TOB.Create ('',TOBBTB,-1);
        AddChampsSupresCon(TOBBB);
        TOBBB.PutValue('CODEAFFAIRE',TOBI.getValue('AFF_AFFAIREINIT'));
        TOBBB.PutValue('LIBELLEAFFAIRE',TOBI.getValue('AFF_LIBELLE'));
        TOBBB.PutValue('DATE',TOBI.getValue('AFF_DATEFIN'));
        TOBBB.PutValue('A_FACTURER',TOBI.GetValue('GP_TOTALHTDEV'));
				TOBBB.PutValue('COUT',TOBI.GetValue('GP_MONTANTPR'));
        TOBBB.PutValue('APPEL',TOBI.GetValue('GP_AFFAIRE'));
    		TOBBB.PutValue('CLIENT',TOBI.getValue('CLIENT'));
      end;
    end;
  FINALLY
  	Ferme (QQ);
    TOBINT.free;
  END;
end;

procedure TOF_BTRESCONTRATS.ChargeLesDevisfacture;
var SQL : string;
		QQ,Q1 : TQuery;
    TOBINT,TOBI,TOBBB : TOB;
    indice : integer;
begin
	TOBINT := TOB.Create ('LES DEVIS FACTURES',nil,-1);
	Sql := 'SELECT DISTINCT GL_NATUREPIECEG,GL_SOUCHE,GL_NUMERO,GL_INDICEG FROM LIGNE '+
				 'LEFT JOIN AFFAIRE A1 ON A1.AFF_AFFAIRE=GL_AFFAIRE '+
         'LEFT JOIN AFFAIRE A2 ON A2.AFF_AFFAIRE=A1.AFF_AFFAIREINIT '+
         'WHERE GL_NATUREPIECEG IN ("FPR","FAC") AND ';

	if CodeAffaire <> '' then
  begin
  	Sql := Sql + 'AND A2.AFF_AFFAIRE LIKE "' + CodeAffaire + '%" ';
  end else
  begin
  	Sql := Sql + 'AND A2.AFF_AFFAIRE LIKE "I%" ';
    if CodeAffaire1 <> '' then
      Sql := Sql + 'AND A2.AFF_AFFAIRE1 = "' + CodeAffaire1 + '" ';
    if CodeAffaire2 <> '' then
      Sql := Sql + 'AND A2.AFF_AFFAIRE2 = "' + CodeAffaire2 + '" ';
    if CodeAffaire3 <> '' then
      Sql := Sql + 'AND A2.AFF_AFFAIRE3 = "' + CodeAffaire3 + '" ';
  end;
  if CodeClient <> '' then
  begin
    Sql := Sql + 'AND A2.AFF_TIERS = "' + CodeClient + '" ';
  end;
  if Responsable <> '' then
  begin
    Sql := Sql + 'AND A2.AFF_RESPONSABLE = "' + Responsable + '" ';
  end;
  if Apporteur <> '' then
  begin
    Sql := Sql + ' AND A2.AFF_APPORTEUR = "' + Apporteur + '" ';
  end;
//
  if LibreAff1 <> '' then
  begin
    Sql := Sql + ' AND A2.AFF_LIBREAFF1 = "' + LibreAff1 + '"';
  end;
  if LibreAff2 <> '' then
  begin
    Sql := Sql + ' AND A2.AFF_LIBREAFF2 = "' + LibreAff2 + '"';
  end;
  if LibreAff3 <> '' then
  begin
    Sql := Sql + ' AND A2.AFF_LIBREAFF3 = "' + LibreAff3 + '"';
  end;
  if LibreAff4 <> '' then
  begin
    Sql := Sql + ' AND A2.AFF_LIBREAFF4 = "' + LibreAff4 + '"';
  end;
  if LibreAff5 <> '' then
  begin
    Sql := Sql + ' AND A2.AFF_LIBREAFF5 = "' + LibreAff5 + '"';
  end;
  if LibreAff6 <> '' then
  begin
    Sql := Sql + ' AND A2.AFF_LIBREAFF6 = "' + LibreAff6 + '"';
  end;
//

//         'A1.AFF_AFFAIREINIT LIKE "'+CodeAffaire+'%" AND '+
	Sql := Sql +' AND A1.AFF_ETATAFFAIRE IN ("FAC","CL1") AND GL_DATEPIECE >= "'+USDATETIME(DateDebut)+'" AND '+
         'GL_DATEPIECE <= "'+USDATETIME(DateFin)+'" AND GL_AFFAIRE LIKE "W%" AND '+
         'GL_PIECEORIGINE <> "" AND GL_TOTALHTDEV <> 0';
  TRY
    QQ := OpenSql (Sql,true,-1,'',True);
    if not QQ.eof then
    begin
    	TOBINT.LoadDetailDB('','','',QQ,false);
      for Indice := 0 to TOBINT.detail.count -1 do
      begin
      	TOBI := TOBINT.detail[Indice];
        Sql := 'SELECT GP_AFFAIRE,GP_AFFAIREDEVIS,GP_TOTALHTDEV,GP_MONTANTPR,A1.AFF_AFFAIREINIT,GP_DATEPIECE,A1.AFF_LIBELLE FROM PIECE '+
               'LEFT JOIN AFFAIRE A1 ON A1.AFF_AFFAIRE=GP_AFFAIRE '+
               'LEFT JOIN AFFAIRE A2 ON A2.AFF_AFFAIRE=A1.AFF_AFFAIREINIT WHERE '+
               'GP_NATUREPIECEG="'+TOBI.GetValue('GL_NATUREPIECEG')+'" AND '+
               'GP_SOUCHE="'+TOBI.GetValue('GL_SOUCHE')+'" AND '+
               'GP_NUMERO='+IntTostr(TOBI.GetValue('GL_NUMERO'))+' AND '+
               'GP_INDICEG='+IntToStr(TOBI.GetValue('GL_INDICEG'));
        Q1 := OpenSql(Sql,true,1,'',true);
        if not Q1.eof then
        begin
          TOBBB := TOB.Create ('',TOBBTB,-1);
          AddChampsSupresCon(TOBBB);
          TOBBB.PutValue('CODEAFFAIRE',Q1.findField('AFF_AFFAIREINIT').AsString);
          TOBBB.PutValue('LIBELLEAFFAIRE',Q1.FindField('AFF_LIBELLE').AsString);
          TOBBB.PutValue('DATE',Q1.FindField('GP_DATEPIECE').AsDateTime);
          TOBBB.PutValue('FACTURE',Q1.FindField('GP_TOTALHTDEV').AsFloat);
          TOBBB.PutValue('APPEL',Q1.FindField('GP_AFFAIRE').AsString);
					TOBBB.PutValue('COUT',Q1.FindField('GP_MONTANTPR').AsFloat);
        end;
        ferme(Q1);
      end;
    end;
  FINALLY
  	Ferme (QQ);
    TOBINT.free;
  END;
end;

procedure TOF_BTRESCONTRATS.bEffaceAff1Click(Sender: TObject);
begin
  inherited;

  SetControltext('AFF_AFFAIRE0','I');
end;

initialization
  registerclasses([TOF_BTRESCONTRATS]);
end.

