{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 03/09/2001
Modifié le ... :   /  /
Description .. : Source TOF de la TABLE : CPAFFVENTILANA ()
Mots clefs ... : TOF;CPAFFVENTILANA
*****************************************************************}
unit UTOFAffGrille;

interface

uses StdCtrls,
  Controls,
  Classes,
  Ent1, //VH^
  {$IFDEF EAGLCLIENT}
  MaineAGL, // AGLLanceFiche
  eMul,
  {$ELSE}
  db,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  FE_Main, // AGLLanceFiche
  Mul,
  {$ENDIF}
  HCtrls,
  HEnt1,
  HMsgBox,
  UTOF,
  HTB97,
  UtotVentilType, // ParamVentilType
  Ventil; // ParamVentil

function CPLanceFicheAFFVentilana: string;

type
  TOF_CPAFFVENTILANA = class(TOF)
    Ligne1: TcheckBox;
    Laxe: THValComboBox;
    CCCode: ThEdit;
    BCherche: TToolbarButton97;
    BZoom: TToolbarButton97;
    BSupprimer: TToolbarButton97;
    BNouveau: TToolbarButton97;
    procedure BNouveauOnClick(Sender: TObject);
    procedure BSupprimerOnClick(Sender: TObject);
    procedure SupprimeVentil;
    procedure BZoomOnClick(Sender: TObject);
    procedure OnArgument(S: string); override;
    procedure OnLoad; override;
  end;

var
  LeCodeASupprimer: string;

implementation

function CPLanceFicheAFFVentilana: string;
begin
  Result := AGLLanceFiche('CP', 'CPAFFVENTILANA', '', '', '');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : G.Verdier
Créé le ...... : 05/09/2001
Modifié le ... :   /  /
Description .. : Visualisation et/ou modification des ventilations types
Mots clefs ... :
*****************************************************************}
procedure TOF_CPAFFVENTILANA.BZoomOnClick(Sender: tobject);
var
  F: TFMul;
  LeCodeARenvoyer: string;
begin
  F := TFMul(Ecran);
  LeCodeARenvoyer := '';
  if F <> nil then
    if (not F.Q.Eof) then
    begin
      if F.Q.FindField('CC_CODE') <> nil then LeCodeARenvoyer := F.Q.FindField('CC_CODE').AsString;
    end;
  if LeCodeARenvoyer <> '' then Paramventil('TY', LeCodeARenvoyer, '12345', taModif, True);
  BCherche.Click;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : G.Verdier
Créé le ...... : 06/09/2001
Modifié le ... : 06/09/2001
Description .. : Creation nouvelle ventilation type
Mots clefs ... :
*****************************************************************}
procedure TOF_CPAFFVENTILANA.BNouveauOnClick(Sender: tobject);
begin
  ParamVentiltype;
  BCherche.Click;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : G.Verdier
Créé le ...... : 07/09/2001
Modifié le ... :   /  /
Description .. : Suppression d'1 ventil type.
Mots clefs ... :
*****************************************************************}
procedure TOF_CPAFFVENTILANA.BSupprimerOnClick(Sender: tobject);
var
  F: TFMul;
begin
  F := TFMul(Ecran);
  LeCodeASupprimer := '';
  if F <> nil then
    if (not F.Q.Eof) then
    begin
      if F.Q.FindField('CC_CODE') <> nil then
      begin
        LeCodeASupprimer := F.Q.FindField('CC_CODE').AsString;
      end;
    end;
  if PGIAsk(TraduireMemoire('Confirmez-vous la suppression ?'), TraduireMemoire('Suppression Code Ventilation type')) = mrYes then
  begin
    if LeCodeASupprimer <> '' then
    begin
      if Transactions(SupprimeVentil, 1) = oeUnknown then
        PGIBox(TraduireMemoire('Une erreur est survenue, traitement interrompu'), TraduireMemoire('Suppression code ventilation type'))
      else Bcherche.Click;
    end;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : G.Verdier
Créé le ...... : 07/09/2001
Modifié le ... : 07/09/2001
Description .. : Suppression de la ventil type selon son axe
Mots clefs ... :
*****************************************************************}
procedure TOF_CPAFFVENTILANA.SupprimeVentil;
var
  ReqSql: string;
begin
  ReqSql := 'DELETE FROM VENTIL WHERE V_NATURE LIKE "TY%" AND V_COMPTE="' + LeCodeASupprimer + '"';
  ExecuteSql(ReqSql);
  ReqSql := 'DELETE FROM CHOIXCOD WHERE CC_CODE ="' + LeCodeASupprimer + '" AND CC_TYPE="VTY"';
  ExecuteSql(ReqSql);
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : G.Verdier
Créé le ...... : 05/09/2001
Modifié le ... : 05/09/2001
Description .. : Recherche des enregistrements selon l'axe choisi.
Mots clefs ... :
*****************************************************************}
procedure TOF_CPAFFVENTILANA.OnLoad;
var
  AxeChoisi: Integer;
  TY, XX: string;
begin
  TY := 'TY%';
  if LAXE <> nil then
  begin
    AxeChoisi := THValComboBox(GetControl('LAXE')).ItemIndex;
    case AxeChoisi of
      0: TY := 'TY%';
      1: TY := 'TY1';
      2: TY := 'TY2';
      3: TY := 'TY3';
      4: TY := 'TY4';
      5: TY := 'TY5';
    else TY := 'TY%';
    end; // End du Case
    TY := 'V_NATURE LIKE "' + TY + '"';
  end; // End du If LAXE
  if LIGNE1 <> nil then
  begin
    if GetControlText('LIGNE1') = 'X' then XX := 'AND V_NUMEROVENTIL=1' else XX := '';
  end; // End du If LIGNE
  SetControlText('XX_WHERE', TY + XX);
  inherited;
end;

procedure TOF_CPAFFVENTILANA.OnArgument(S: string);
begin
  inherited;

  //Sg6 17.2.05 Gestion ana croisaxe
  TFMul(Ecran).Q.Manuel := True;

  if VH^.AnaCroisaxe then
    TFMul(Ecran).Q.Liste := 'CPAFFANACROISAXE'
  else
    TFMul(Ecran).Q.Liste := 'CPAFFANA';

  TFMul(Ecran).Q.Manuel := False;

  Laxe := ThValComboBox(GetControl('LAXE'));
  CCCode := ThEdit(GetControl('CC_CODE'));
  Laxe.Itemindex := 0;
  SetFocusControl('CC_CODE');
  Ligne1 := TCheckBox(GetControl('LIGNE1'));
  BCherche := TToolBarButton97(GetControl('BCHERCHE'));
  Bsupprimer := TToolBarButton97(GetControl('BSUPPRIMER'));
  BZoom := TToolBarButton97(GetControl('BZOOM'));
  BNouveau := TToolBarButton97(GetControl('BINSERT'));
  if BZoom <> nil then
  begin
    BZoom.OnClick := BZoomOnClick;
  end;
  if BNouveau <> nil then
  begin
    BNouveau.OnClick := BNouveauOnClick;
  end;
  if BSupprimer <> nil then
  begin
    Bsupprimer.OnClick := BSupprimerOnClick;
  end;
end;


initialization
  registerclasses([TOF_CPAFFVENTILANA]);
end.

