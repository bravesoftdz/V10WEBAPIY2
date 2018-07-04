{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 17/02/2004
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTCOMPARATIF ()
Mots clefs ... : TOF;BTCOMPARATIF
*****************************************************************}
unit BTCOMPARATIF_TOF;

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
  HQry,
  UTOF,
  Stat,
  UTOB,
  utobview,
  AGLInit,
  AGLInitBTP,
  facture,
  UtilPhases,
  uTofAfBaseCodeAffaire;

type
  TOF_BTCOMPARATIF = class(TOF_AFBASECODEAFFAIRE)
    procedure OnNew; override;
    procedure OnDelete; override;
    procedure OnUpdate; override;
    procedure OnLoad; override;
    procedure OnArgument(S: string); override;
    procedure OnClose; override;
    procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_: THEdit); override;

  private
    TV1: TTobViewer;
    TOBBCP, TOBConso: TOB;
    AffichageAuto, OldAutoSearch : Boolean;
    CodeAffaire,CodeAffaire1,CodeAffaire2,Codeaffaire3: string;
    procedure OnDblClickTV(Sender: TObject);
    procedure ChargeLePrevu;
    procedure ChargeLesConsos;
    procedure FusionneLesTobs;
    procedure AjouteChampSup(var TOBTMP: TOB; TousLesChamps : Boolean);
  end;

(* COMPOSITION DES TOBS

Désignation du champ      nom dans TOBBCP   nom dans TOBConso

Code affaire              GL_AFFAIRE        BCO_AFFAIRE
Libellé affaire           AFF_LIBELLE       AFF_LIBELLE
Code Phase                BLP_PHASETRA      BCO_PHASETRA
Libellé Phase             BPC_LIBELLE
Code article              GL_CODEARTICLE    BCO_CODEARTICLE
Libellé article           GL_LIBELLE        GA_LIBELLE
Famille article 1         GL_FAMILLENIV1    GA_FAMILLENIV1
Famille article 2         GL_FAMILLENIV2    GA_FAMILLENIV2
Famille article 3         GL_FAMILLENIV3    GA_FAMILLENIV3
table libre article 1     GL_LIBREART1      GA_LIBREART1
table libre article 2     GL_LIBREART2      GA_LIBREART2
table libre article 3     GL_LIBREART3      GA_LIBREART3
Quantité prévue           GL_QTEFACT
Quantité avancée          GL_QTEPREVAVANC
PU achat prévu            GL_DPA
PU revient prévu          GL_DPR
PU vente prévu            GL_PUHT
Montant achat prévu       PREVUPA
Montant revient prévu     PREVUPR
Montant vente prévu       PREVUPV
Montant achat avancé      AVANCEPA
Montant revient avancé    AVANCEPR
Montant vente avancé      AVANCEPV

Quantité consommmée       REALISEQTE        BCO_QUANTITE
PU achat conso            REALISEPUA        BCO_DPA
PU revient conso          REALISEPUR        BCO_DPR
PU vente conso            REALISEPUV        BCO_PUHT
Montant achat conso       REALISEMTA        BCO_MONTANTACH
Montant revient conso     REALISEMTR        BCO_MONTANTPR
Montant vente conso       REALISEMTV        BCO_MONTANTHT

*)

implementation

procedure TOF_BTCOMPARATIF.OnNew;
begin
  inherited;
end;

procedure TOF_BTCOMPARATIF.OnDelete;
begin
  inherited;
end;

procedure TOF_BTCOMPARATIF.OnUpdate;
begin
  inherited;
ChargeLePrevu;
ChargeLesConsos;
FusionneLesTobs;
if TOBBCP <> nil then
  begin
  TFStat(Ecran).LaTOB := TOBBCP
  end;
end;

procedure TOF_BTCOMPARATIF.OnLoad;
var F : TFStat ;
    iCol:integer;
    CC:THLabel ;
begin
inherited ;
CodeAffaire := GetControltext('BCPAFFAIRE');
CodeAffaire1 := GetControltext('BCPAFFAIRE1');
CodeAffaire2 := GetControltext('BCPAFFAIRE2');
CodeAffaire3 := GetControltext('BCPAFFAIRE3');
Ecran.Caption := 'COMPARATIF CHANTIER ' + codeaffaire;

F:=TFStat(Ecran) ;
for iCol:=1 to 3 do
  begin
  CC:=THLabel(TFStat(F).FindComponent('TGA_FAMILLENIV'+InttoStr(iCol)));
  CC.Caption:=RechDom('GCLIBFAMILLE','LF'+InttoStr(iCol),false);
  end;
end;

procedure TOF_BTCOMPARATIF.OnArgument(S: string);
begin
  inherited;
  if S = 'AUTO' then AffichageAuto:=True
  else AffichageAuto:=False;

  // Gestion des évènements du TB
  TV1 := TTobViewer(GetControl('TV'));
  TV1.OnDblClick := OnDblClickTV;
  TOBBCP := TOB.create('Ma Tob', nil, -1);
  TOBConso := TOB.create('Les Consos', nil, -1);

  if AffichageAuto = True then
    begin
    SetControlEnabled('BCPAFFAIRE1', false);
    SetControlEnabled('BCPAFFAIRE2', false);
    SetControlEnabled('BCPAFFAIRE3', false);
    SetControlEnabled('BEFFACEAFF1', false);
    SetControlEnabled('BSELECTAFF1', false);
    end;
  OldAutoSearch:=V_PGI.AutoSearch;
  V_PGI.AutoSearch:=AffichageAuto;

//uniquement en line
//  SetControlProperty('PCOMPLEMENT', 'TabVisible', False);
//  SetControlProperty('PAVANCE', 'TabVisible', False);

end;

procedure TOF_BTCOMPARATIF.OnClose;
begin
  inherited;
  FreeAndNil(TOBConso);
  FreeAndNil(TOBBCP);
  V_PGI.AutoSearch:=OldAutoSearch;
end;

//******************************************************************************
//********************* Zooms / double clic  ***********************************
//******************************************************************************

procedure TOF_BTCOMPARATIF.OnDblClickTV(Sender: TObject);
var
  Req, Arg: string;
  QQ: Tquery;
begin
  if (TV1.ColName[TV1.CurrentCol] = 'GL_QTEFACT') or (Copy(TV1.ColName[TV1.CurrentCol], 1, 5) = 'PREVU') then
  begin
    Req := 'SELECT GP_NATUREPIECEG, GP_DATEPIECE, GP_SOUCHE, GP_NUMERO, GP_INDICEG ' +
      'FROM PIECE ' +
      'WHERE GP_AFFAIRE="' + TV1.AsString[TV1.ColIndex('GL_AFFAIRE'), TV1.CurrentRow] + '" ' +
      'AND GP_NATUREPIECEG="PBT"';
    QQ := OpenSQL(Req, true,-1,'',true);
    if not QQ.EOF then
    begin
      Arg := QQ.Fields[0].AsString + ';' + QQ.Fields[1].AsString + ';' + QQ.Fields[2].AsString + ';' +
        IntToStr(QQ.Fields[3].AsInteger) + ';' + IntToStr(QQ.Fields[4].AsInteger);
      AppelPiece([Arg, 'ACTION=CONSULTATION'], 2);
    end;
    ferme(QQ);
  end
  else if (TV1.ColName[TV1.CurrentCol] = 'GL_QTEPREVAVANC') or (Copy(TV1.ColName[TV1.CurrentCol], 1, 6) = 'AVANCE') then
  begin
    Req := 'SELECT GP_NATUREPIECEG, GP_DATEPIECE, GP_SOUCHE, GP_NUMERO, GP_INDICEG, GP_TIERS, GP_AFFAIRE ' +
      'FROM PIECE ' +
      'WHERE GP_AFFAIRE="' + TV1.AsString[TV1.ColIndex('GL_AFFAIRE'), TV1.CurrentRow] + '" ' +
      'AND GP_NATUREPIECEG="PBT"';
    QQ := OpenSQL(Req, true,-1,'',true);
    if not QQ.EOF then
      begin
      SaisieAvancementChantier(QQ.Fields[0].AsString,QQ.Fields[1].AsString,QQ.Fields[2].AsString,
                               QQ.Fields[5].AsString, QQ.Fields[6].AsString,
                               QQ.Fields[3].AsInteger, QQ.Fields[4].AsInteger, TaConsult);
      end;
    ferme(QQ);
  end
  else if (TV1.ColName[TV1.CurrentCol] = 'GL_AFFAIRE') or (TV1.ColName[TV1.CurrentCol] = 'AFF_LIBELLE') then
  begin
    V_PGI.DispatchTT(5, taConsult, TV1.AsString[TV1.ColIndex('GL_AFFAIRE'), TV1.CurrentRow], '', '');
  end;
end;

procedure TOF_BTCOMPARATIF.NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_: THEdit);
begin
  Aff0 := THEdit(GetControl('BCPAFFAIRE0'));
  Aff := THEdit(GetControl('BCPAFFAIRE'));
  Aff1 := THEdit(GetControl('BCPAFFAIRE1'));
  Aff2 := THEdit(GetControl('BCPAFFAIRE2'));
  Aff3 := THEdit(GetControl('BCPAFFAIRE3'));
  Aff4 := THEdit(GetControl('BCPAVENANT'));
end;

// Ajout des champs dans la TOB

procedure TOF_BTCOMPARATIF.AjouteChampSup(var TOBTMP: TOB; TousLesChamps : Boolean);
begin
  if TousLesChamps then
    begin
    TOBTMP.addchampsup('GL_AFFAIRE', false); TOBTMP.PutValue('GL_AFFAIRE', '');
    TOBTMP.addchampsup('AFF_LIBELLE', false); TOBTMP.PutValue('AFF_LIBELLE', '');
    TOBTMP.addchampsup('BLP_PHASETRA', false); TOBTMP.PutValue('BLP_PHASETRA', '');
    TOBTMP.addchampsup('BPC_LIBELLE', false); TOBTMP.PutValue('BPC_LIBELLE', '');
    TOBTMP.addchampsup('GL_CODEARTICLE', false); TOBTMP.PutValue('GL_CODEARTICLE', '');
    TOBTMP.addchampsup('GL_ARTICLE', false); TOBTMP.PutValue('GL_ARTICLE', '');
    TOBTMP.addchampsup('GL_LIBELLE', false); TOBTMP.PutValue('GL_LIBELLE', '');
    TOBTMP.addchampsup('GL_FAMILLENIV1', false); TOBTMP.PutValue('GL_FAMILLENIV1', '');
    TOBTMP.addchampsup('GL_FAMILLENIV2', false); TOBTMP.PutValue('GL_FAMILLENIV2', '');
    TOBTMP.addchampsup('GL_FAMILLENIV3', false); TOBTMP.PutValue('GL_FAMILLENIV3', '');
    TOBTMP.addchampsup('GL_LIBREART1', false); TOBTMP.PutValue('GL_LIBREART1', '');
    TOBTMP.addchampsup('GL_LIBREART2', false); TOBTMP.PutValue('GL_LIBREART2', '');
    TOBTMP.addchampsup('GL_LIBREART3', false); TOBTMP.PutValue('GL_LIBREART3', '');
    TOBTMP.addchampsup('GL_QTEFACT', false); TOBTMP.PutValue('GL_QTEFACT', 0.0);
    TOBTMP.addchampsup('GL_QTEPREVAVANC', false); TOBTMP.PutValue('GL_QTEPREVAVANC', 0.0);
    TOBTMP.addchampsup('GL_DPA', false); TOBTMP.PutValue('GL_DPA', 0.0);
    TOBTMP.addchampsup('GL_DPR', false); TOBTMP.PutValue('GL_DPR', 0.0);
    TOBTMP.addchampsup('GL_PUHT', false); TOBTMP.PutValue('GL_PUHT', 0.0);
    TOBTMP.addchampsup('PREVUPA', false); TOBTMP.PutValue('PREVUPA', 0.0);
    TOBTMP.addchampsup('PREVUPR', false); TOBTMP.PutValue('PREVUPR', 0.0);
    TOBTMP.addchampsup('PREVUPV', false); TOBTMP.PutValue('PREVUPV', 0.0);
    TOBTMP.addchampsup('AVANCEPA', false); TOBTMP.PutValue('AVANCEPA', 0.0);
    TOBTMP.addchampsup('AVANCEPR', false); TOBTMP.PutValue('AVANCEPR', 0.0);
    TOBTMP.addchampsup('AVANCEPV', false); TOBTMP.PutValue('AVANCEPV', 0.0);
    end;
  TOBTMP.addchampsup('REALISEQTE', false);  TOBTMP.PutValue('REALISEQTE', 0.0);
  TOBTMP.addchampsup('REALISEPUA', false);  TOBTMP.PutValue('REALISEPUA', 0.0);
  TOBTMP.addchampsup('REALISEPUR', false);  TOBTMP.PutValue('REALISEPUR', 0.0);
  TOBTMP.addchampsup('REALISEPUV', false);  TOBTMP.PutValue('REALISEPUV', 0.0);
  TOBTMP.addchampsup('REALISEMTA', false);  TOBTMP.PutValue('REALISEMTA', 0.0);
  TOBTMP.addchampsup('REALISEMTR', false);  TOBTMP.PutValue('REALISEMTR', 0.0);
  TOBTMP.addchampsup('REALISEMTV', false);  TOBTMP.PutValue('REALISEMTV', 0.0);
end;

procedure TOF_BTCOMPARATIF.ChargeLePrevu;
var
  Req: string;
  QQ: Tquery;
begin
  Req := TFStat(Ecran).stSQL;
  if CodeAffaire <> '' then
  begin
    Req := Req + ' AND GL_AFFAIRE = "' + CodeAffaire + '"';
  end else if CodeAffaire1 <> '' then
  begin
    if CodeAffaire1 <> '' then
      Req := Req + ' AND GL_AFFAIRE1 = "' + CodeAffaire1 + '"';
    if CodeAffaire2 <> '' then
      Req := Req + ' AND GL_AFFAIRE2 = "' + CodeAffaire2 + '"';
    if CodeAffaire3 <> '' then
      Req := Req + ' AND GL_AFFAIRE3 = "' + CodeAffaire3 + '"';
  end;

  Req := StringReplace (Req, 'BCP_', 'GL_', [rfReplaceAll]);

  QQ := OpenSQL(Req, true,-1,'',true);
  TOBBCP.LoadDetailDB('', '', '', QQ, false);
  ferme(QQ);
end;

procedure TOF_BTCOMPARATIF.ChargeLesConsos;
var
  Req, WhereSql: string;
  QQ: Tquery;
begin
  Req := 'SELECT BCO_AFFAIRE, BCO_CODEARTICLE,BCO_ARTICLE, BCO_QUANTITE, BCO_DPA, BCO_DPR, BCO_PUHT, BCO_MONTANTACH, BCO_MONTANTPR, BCO_MONTANTHT, ' +
         'AFF_LIBELLE, GA_ARTICLE, GA_CODEARTICLE, GA_LIBELLE, GA_FAMILLENIV1, GA_FAMILLENIV2, GA_FAMILLENIV3, ' +
         'GA_LIBREART1, GA_LIBREART2, GA_LIBREART3 ' ;

// en attente : chargement systématique des phases et présentation par phases gérée dans le TV
//  if TCheckBox(GetControl('BCPPHASE')).checked=True then
  Req := Req + ', BCO_PHASETRA ';

  Req := Req +
    'FROM CONSOMMATIONS ' +
    'LEFT OUTER JOIN AFFAIRE ON AFF_AFFAIRE=BCO_AFFAIRE AND ' +
    'LEFT OUTER JOIN ARTICLE ON GA_ARTICLE=BCO_ARTICLE ';

  WhereSql := RecupWhereCritere (TPageControl(GetCOntrol('PAGES')));
	if WhereSql = '' then
  begin
  	Req := Req + ' WHERE ';
  end else
  begin
    WhereSql := StringReplace (WhereSql, 'GL_', 'GA_', [rfReplaceAll]);
  	Req := Req + WhereSql+' AND ';
  end;

  Req := Req + ' BCO_TRANSFORME="-" AND BCO_TRAITEVENTE="-" AND (BCO_NATUREPIECEG="" OR BCO_NATUREPIECEG="BLC" OR BCO_NATUREPIECEG="LBT"  OR BCO_NATUREPIECEG="BFC" OR BCO_NATUREPIECEG="AF" OR BCO_NATUREPIECEG="AFS")';

  if CodeAffaire <> '' then
  begin
    Req := Req + ' AND BCO_AFFAIRE = "' + CodeAffaire + '"';
  end else
  begin
  	if CodeAffaire1 <> '' then
      Req := Req + ' AND BCO_AFFAIRE1 = "' + CodeAffaire1 + '"';
    if CodeAffaire2 <> '' then
      Req := Req + ' AND BCO_AFFAIRE2 = "' + CodeAffaire2 + '"';
    if CodeAffaire3 <> '' then
      Req := Req + ' AND BCO_AFFAIRE3 = "' + CodeAffaire3 + '"';
  end;

  Req := StringReplace (Req, 'BCP_', 'GA_', [rfReplaceAll]);

  QQ := OpenSQL(Req, true,-1,'',true);
  TOBConso.LoadDetailDB('', '', '', QQ, false);
  ferme(QQ);
end;

// Mise à jour du réalisé dans la TOBBCP
procedure TOF_BTCOMPARATIF.FusionneLesTobs;
var
  TOBTMP, TC: TOB;
  i: integer;
begin

// Ajout des champs sup du réalisé dans la TOB initiale
for i := 0 to TOBBCP.Detail.Count - 1 do
  begin
  TC:=TOBBCP.Detail[i];

  if VarIsNull(TC.getValue('BLP_PHASETRA')) or (VarAsType(TC.getValue('BLP_PHASETRA'), varString) = #0) then
    TC.PutValue('BLP_PHASETRA','');

  if (TC <> nil) then AjouteChampSup(TC, False);
  end;

// lecture de la TOB Conso
for i := 0 to TOBConso.Detail.Count - 1 do
  begin
  TC:=TOBConso.Detail[i];
  if VarIsNull(TC.getValue('BCO_PHASETRA')) or (VarAsType(TC.getValue('BCO_PHASETRA'), varString) = #0) then
    TC.PutValue('BCO_PHASETRA','');
  if (TC <> nil) then
    begin
    // recherche de l'article dans la TOBBCP
    TOBTMP := TOBBCP.FindFirst(['GL_AFFAIRE','BLP_PHASETRA','GL_ARTICLE'],
                               [TC.GetValue('BCO_AFFAIRE'), TC.GetValue('BCO_PHASETRA'), TC.GetValue('BCO_ARTICLE')],
                               True);
    if TOBTMP = nil then
      begin
      // Pas de prévu pour la conso lue : on ajoute un enregistrement
      TOBTMP := TOB.create ('',TOBBCP,-1);
      AjouteChampSup(TOBTMP, True);
      TOBTMP.PutValue('GL_AFFAIRE',TC.GetValue('BCO_AFFAIRE'));
      TOBTMP.PutValue('AFF_LIBELLE',TC.GetValue('AFF_LIBELLE'));
      TOBTMP.PutValue('BLP_PHASETRA',TC.GetValue('BCO_PHASETRA'));
      TOBTMP.PutValue('BPC_LIBELLE',GetLibellePhase(TC.GetValue('BCO_AFFAIRE'),TC.GetValue('BCO_PHASETRA')));
      TOBTMP.PutValue('GL_CODEARTICLE',TC.GetValue('GA_CODEARTICLE'));
      TOBTMP.PutValue('GL_ARTICLE',TC.GetValue('GA_ARTICLE'));
      TOBTMP.PutValue('GL_LIBELLE',TC.GetValue('GA_LIBELLE'));
      TOBTMP.PutValue('GL_FAMILLENIV1',TC.GetValue('GA_FAMILLENIV1'));
      TOBTMP.PutValue('GL_FAMILLENIV2',TC.GetValue('GA_FAMILLENIV2'));
      TOBTMP.PutValue('GL_FAMILLENIV3',TC.GetValue('GA_FAMILLENIV3'));
      TOBTMP.PutValue('GL_LIBREART1',TC.GetValue('GA_LIBREART1'));
      TOBTMP.PutValue('GL_LIBREART2',TC.GetValue('GA_LIBREART2'));
      TOBTMP.PutValue('GL_LIBREART3',TC.GetValue('GA_LIBREART3'));
      TOBTMP.PutValue('GL_QTEFACT',0.0);
      TOBTMP.PutValue('GL_QTEPREVAVANC',0.0);
      TOBTMP.PutValue('GL_DPA',0.0);
      TOBTMP.PutValue('GL_DPR',0.0);
      TOBTMP.PutValue('GL_PUHT',0.0);
      TOBTMP.PutValue('PREVUPA',0.0);
      TOBTMP.PutValue('PREVUPR',0.0);
      TOBTMP.PutValue('PREVUPV',0.0);
      TOBTMP.PutValue('AVANCEPA',0.0);
      TOBTMP.PutValue('AVANCEPR',0.0);
      TOBTMP.PutValue('AVANCEPV',0.0);
      TOBTMP.PutValue('REALISEQTE',TC.GetValue('BCO_QUANTITE'));
      TOBTMP.PutValue('REALISEPUA',TC.GetValue('BCO_DPA'));
      TOBTMP.PutValue('REALISEPUR',TC.GetValue('BCO_DPR'));
      TOBTMP.PutValue('REALISEPUV',TC.GetValue('BCO_PUHT'));
      TOBTMP.PutValue('REALISEMTA',TC.GetValue('BCO_MONTANTACH'));
      TOBTMP.PutValue('REALISEMTR',TC.GetValue('BCO_MONTANTPR'));
      TOBTMP.PutValue('REALISEMTV',TC.GetValue('BCO_MONTANTHT'));
      end
    else
      begin
      // Une prévision existe pour la conso lue : on met à jour
      TOBTMP.PutValue('REALISEQTE',TOBTMP.GetValue('REALISEQTE')+TC.GetValue('BCO_QUANTITE'));
      TOBTMP.PutValue('REALISEMTA',TOBTMP.GetValue('REALISEMTA')+TC.GetValue('BCO_MONTANTACH'));
      TOBTMP.PutValue('REALISEMTR',TOBTMP.GetValue('REALISEMTR')+TC.GetValue('BCO_MONTANTPR'));
      TOBTMP.PutValue('REALISEMTV',TOBTMP.GetValue('REALISEMTV')+TC.GetValue('BCO_MONTANTHT'));

      // recalcul des PU moyens
      if TOBTMP.GetValue('REALISEQTE') <> 0.0 then
        begin
        TOBTMP.PutValue('REALISEPUA',TOBTMP.GetValue('REALISEMTA')/TOBTMP.GetValue('REALISEQTE'));
        TOBTMP.PutValue('REALISEPUR',TOBTMP.GetValue('REALISEMTR')/TOBTMP.GetValue('REALISEQTE'));
        TOBTMP.PutValue('REALISEPUV',TOBTMP.GetValue('REALISEMTV')/TOBTMP.GetValue('REALISEQTE'));
        end
      else
        begin
        TOBTMP.PutValue('REALISEPUA',TOBTMP.GetValue('REALISEMTA'));
        TOBTMP.PutValue('REALISEPUR',TOBTMP.GetValue('REALISEMTR'));
        TOBTMP.PutValue('REALISEPUV',TOBTMP.GetValue('REALISEMTV'));
        end;
      end;

    end;
  end;
end;

initialization
  registerclasses([TOF_BTCOMPARATIF]);
end.
