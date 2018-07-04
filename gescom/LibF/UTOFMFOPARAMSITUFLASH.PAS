{***********UNITE*************************************************
Auteur  ...... : O. TARCY
Créé le ...... : 08/08/2001
Modifié le ... : 08/08/2001
Description .. : Source TOF de la FICHE : MFOPARAMSITUFLASH ()
Mots clefs ... : TOF;MFOPARAMSITUFLASH
*****************************************************************}
unit UTOFMFOPARAMSITUFLASH;

interface
uses
  Classes, controls, StdCtrls, Vierge, UTOF;

type
  TOF_MFOPARAMSITUFLASH = class(TOF)
    procedure OnUpdate; override;
    procedure OnArgument(S: string); override;
  private
    sOldCategorie, sNewCategorie: string;
    sOldStats, sNewStats, sOldTypeCateg, sNewTypeCateg: integer;
    sOldReglts, sNewReglts, sOldVendeurs, sNewVendeurs: integer;
    sOldOpFi, sOldPreDiv, sOldCaisBtq: boolean;
    function GetRadioButtonState(NomChamp: string): boolean;
  end;

implementation
uses
  FODefi, FOUtil, Ent1;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : O. TARCY
Créé le ...... : 05/06/2001
Modifié le ... : 17/05/2002
Description .. : Lit une valeur d'une clé de registre
Mots clefs ... : FO
*****************************************************************}

function LitValeurRegistre(NomCle: string; Valeur: variant): variant;
begin
  Result := FOGetFromRegistry(REGSITUFLASH, NomCle, Valeur);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : O. TARCY
Créé le ...... : 05/06/2001
Modifié le ... : 17/05/2002
Description .. : Ecrit une valeur d'une clé de registre
Mots clefs ... : FO
*****************************************************************}

procedure EnregistreValeur(NomCle: string; Valeur: variant);
begin
  FOSaveInRegistry(REGSITUFLASH, NomCle, Valeur);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : O. TARCY
Créé le ...... : 05/06/2001
Modifié le ... : 17/05/2002
Description .. : Lit une valeur d'une clé de registre
Mots clefs ... : FO
*****************************************************************}

function TOF_MFOPARAMSITUFLASH.GetRadioButtonState(NomChamp: string): boolean;
var Ctrl: TControl;
begin
  Result := False;
  Ctrl := GetControl(NomChamp);
  if (Ctrl <> nil) and (Ctrl is TRadioButton) then
    Result := TRadioButton(Ctrl).Checked;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : O. TARCY
Créé le ...... : 05/06/2001
Modifié le ... : 17/05/2002
Description .. : OnUpdate
Mots clefs ... : FO
*****************************************************************}

procedure TOF_MFOPARAMSITUFLASH.OnUpdate;
begin
  inherited;
  // Sauvegarde des paramètres dans la registry s'ils ont été modifiés
  sNewCategorie := GetControlText('COMBOCATEGORIE');
  if sNewCategorie <> sOldCategorie then
    EnregistreValeur(REGCATEGORIEPARDEFAUT, sNewCategorie);

  if GetRadioButtonState('RADIOBUTTON2') then
    sNewStats := STATSINI
  else
    sNewStats := STATSCATEG;
  if sNewStats <> sOldStats then
    EnregistreValeur(REGSTATSPARDEFAUT, sNewStats);

  if GetRadioButtonState('RADIOBUTTON') then
    sNewVendeurs := VENDEURSINI
  else
    sNewVendeurs := VENDEURSACTIFS;
  if sNewVendeurs <> sOldVendeurs then
    EnregistreValeur(REGVENDEURSPARDEFAUT, sNewVendeurs);

  if GetRadioButtonState('RADIOBUTTON4') then
    sNewTypeCateg := TYPECATEGINI
  else
    sNewTypeCateg := TYPECATEGTOUT;
  if sNewTypeCateg <> sOldTypeCateg then
    EnregistreValeur(REGTYPECATEGPARDEFAUT, sNewTypeCateg);

  if GetRadioButtonState('RADIOBUTTON6') then
    sNewReglts := REGLTSINI
  else if GetRadioButtonState('RADIOBUTTON7') then
    sNewReglts := REGLTSTOUS
  else
    sNewReglts := REGLTSAUCUN;
  if sNewReglts <> sOldReglts then
    EnregistreValeur(REGREGLTSPARDEFAUT, sNewReglts);

  sOldPreDiv := (GetCheckBoxState('CHECKBOX1') = cbChecked);
  EnregistreValeur(REGDETAILPRESTATION, sOldPreDiv);
  sOldOpFi := (GetCheckBoxState('CHECKBOX2') = cbChecked);
  EnregistreValeur(REGDETAILOPCAISSE, sOldOpFi);

  if EtabForce = '' then
  begin
    sOldCaisBtq := (GetCheckBoxState('CHECKBOX3') = cbChecked);
    EnregistreValeur(REGCAISSEBTQ, sOldCaisBtq);
  end;

  TFVierge(Ecran).Retour := 'validation'
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : O. TARCY
Créé le ...... : 05/06/2001
Modifié le ... : 17/05/2002
Description .. : OnArgument
Mots clefs ... : FO
*****************************************************************}

procedure TOF_MFOPARAMSITUFLASH.OnArgument(S: string);
begin
  inherited;
  //Initialisation et Mémorisation des paramètres
  // la catégorie
  sOldCategorie := LitValeurRegistre(REGCATEGORIEPARDEFAUT, '');
  if sOldCategorie <> '' then SetControlText('COMBOCATEGORIE', sOldCategorie);

  // les statistiques
  sOldStats := LitValeurRegistre(REGSTATSPARDEFAUT, STATSINI);
  if sOldStats = STATSINI then
    SetControlChecked('RADIOBUTTON2', True)
  else
    SetControlChecked('RADIOBUTTON3', True);

  // les vendeurs
  sOldVendeurs := LitValeurRegistre(REGVENDEURSPARDEFAUT, VENDEURSINI);
  if sOldVendeurs = VENDEURSINI then
    SetControlChecked('RADIOBUTTON', True)
  else
    SetControlChecked('RADIOBUTTON1', True);

  // le type de catégorie
  sOldTypeCateg := LitValeurRegistre(REGTYPECATEGPARDEFAUT, TYPECATEGINI);
  if sOldTypeCateg = TYPECATEGINI then
    SetControlChecked('RADIOBUTTON4', True)
  else
    SetControlChecked('RADIOBUTTON5', True);

  // les règlements
  sOldReglts := LitValeurRegistre(REGREGLTSPARDEFAUT, REGLTSINI);
  if sOldReglts = REGLTSINI then
    SetControlChecked('RADIOBUTTON6', True)
  else if sOldReglts = REGLTSTOUS then
    SetControlChecked('RADIOBUTTON7', True)
  else
    SetControlChecked('RADIOBUTTON8', True);

  // détail des opérations de caisse et prestations
  sOldOpFi := LitValeurRegistre(REGDETAILOPCAISSE, False);
  SetControlChecked('CHECKBOX2', sOldOpFi);
  sOldPreDiv := LitValeurRegistre(REGDETAILPRESTATION, False);
  SetControlChecked('CHECKBOX1', sOldPreDiv);

  // Liste des caisses
  if EtabForce = '' then
  begin
    sOldCaisBtq := LitValeurRegistre(REGCAISSEBTQ, True);
    SetControlChecked('CHECKBOX3', sOldCaisBtq);
  end else
  begin
    SetControlChecked('CHECKBOX3', True);
    SetControlEnabled('CHECKBOX3', False);
  end;
end;

initialization
  registerclasses([TOF_MFOPARAMSITUFLASH]);
end.
