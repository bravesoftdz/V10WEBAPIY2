{***********UNITE*************************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 03/12/2002
Modifié le ... : 03/12/2002
Description .. : Source TOM de la TABLE : AFFCDEENTETE
Suite ........ : (AFFCDEENTETE)
Mots clefs ... : TOM;AFFCDEENTETE;AFFCDE
*****************************************************************}
unit AFFCDEENTETE_TOM;

interface

uses StdCtrls, Controls, Classes, sysutils,
  {$IFNDEF EAGLCLIENT}
  db,
  {$ENDIF}
  HCtrls, UTOM, UTob, HEnt1, LookUp;

type
  TOM_AFFCDEENTETE = class(TOM)
    procedure OnNewRecord; override;
    procedure OnDeleteRecord; override;
    procedure OnUpdateRecord; override;
    procedure OnAfterUpdateRecord; override;
    procedure OnLoadRecord; override;
    procedure OnChangeField(F: TField); override;
    procedure OnArgument(S: string); override;
    procedure OnClose; override;
    procedure OnCancelRecord; override;
  private
    procedure ChangeComboLibre;
    procedure ActivePrioAffCde(Num: integer);
    procedure MetEnErreur(NomChamp: string; NoMsg: integer);
    procedure ActiveChamp(NomChamp: string; Active: Boolean);
    procedure OnChangeGereReliquat;
    function ControleIntegrite: Boolean;
  end;

implementation

uses
  EntGC;

const
  NATURECDECLI = 'CC'; // nature de pièce commande client

const
  // Messages d'erreur
  TexteMessage: array[1..3] of string = (
    {1} 'Le code du dépôt à déstocker doit être renseigné.',
    {2} 'La valeur du pourcentage est incorrecte.',
    {3} 'La nature de pièce est incorrecte.'
  );

  {***********A.G.L.Privé.*****************************************
  Auteur  ...... : N. ACHINO
  Créé le ...... : 03/12/2002
  Modifié le ... : 03/12/2002
  Description .. : Changmenet le titre des tables libres dans les combos des
  Suite ........ : critères de priorité
  Mots clefs ... :
  *****************************************************************}

procedure TOM_AFFCDEENTETE.ChangeComboLibre;
var Ctrl: TControl;
  Ind, JJ: integer;
  NomChamp, Stg, Lib: string;
begin
  inherited;
  for JJ := 1 to 3 do
  begin
    NomChamp := 'GEA_PRIOAFFCDE' + IntToStr(JJ);
    Ctrl := GetControl(NomChamp);
    if (Ctrl <> nil) and (Ctrl is THValComboBox) then
    begin
      THValComboBox(Ctrl).Values.BeginUpdate;
      THValComboBox(Ctrl).Items.BeginUpdate;
      // remplacement du libellé par le titre de la table libre
      for Ind := 0 to THValComboBox(Ctrl).Values.Count - 1 do
      begin
        Lib := '';
        Stg := THValComboBox(Ctrl).Values[Ind];
        Stg := RechDom('GCAFFCDEPRIORITE', Stg, True);
        if Copy(Stg, 1, 13) = 'GP_LIBREPIECE' then
        begin
          Stg := GetInfoParPiece(NATURECDECLI, 'GPP_PIECETABLE' + Copy(Stg, 14, 1));
          if (Stg = '') or (Stg = #0) then
            Lib := '.-'
          else
            Lib := RechDom('GCTYPESTATPIECE', Stg, False);
        end else
        begin
          if Copy(Stg, 1, 15) = 'TABLELIBRETIERS' then Stg := 'YTC_' + Stg;
          if Copy(Stg, 1, 4) = 'YTC_' then Lib := ChampToLibelle(Stg);
        end;
        if Lib <> '' then THValComboBox(Ctrl).Items[Ind] := Lib;
      end;
      // suppression des tables libres non utilisées
      repeat
        Ind := THValComboBox(Ctrl).Items.IndexOf('.-');
        if Ind >= 0 then
        begin
          THValComboBox(Ctrl).Values.Delete(Ind);
          THValComboBox(Ctrl).Items.Delete(Ind);
        end;
      until Ind < 0;
      THValComboBox(Ctrl).Values.EndUpdate;
      THValComboBox(Ctrl).Items.EndUpdate;
    end;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 03/12/2002
Modifié le ... : 03/12/2002
Description .. : Active le choix de critères de tri supplémentaires
Mots clefs ... :
*****************************************************************}

procedure TOM_AFFCDEENTETE.ActivePrioAffCde(Num: integer);
var Valeur, NomChamp: string;
  Ind: integer;
  Actif, DescActif: boolean;
begin
  NomChamp := 'GEA_PRIOAFFCDE' + IntToStr(Num);
  Valeur := GetControlText(NomChamp);
  Actif := ((Valeur <> '') and (Valeur <> '000'));
  DescActif := Actif;
  for Ind := Num to 3 do
  begin
    if Ind > Num then
    begin
      // impact sur les critères suivants
      NomChamp := 'GEA_PRIOAFFCDE' + IntToStr(Ind);
      SetControlEnabled(NomChamp, Actif);
      Valeur := GetControlText(NomChamp);
      if not (Actif) and (Valeur <> '') then
      begin
        Valeur := '';
        SetControlText(NomChamp, Valeur);
      end;
      DescActif := ((Valeur <> '') and (Valeur <> '000'));
    end;
    // impact sur les options de tri
    NomChamp := 'GEA_PRIODESC' + IntToStr(Ind);
    SetControlEnabled(NomChamp, DescActif);
    if not (DescActif) and (GetCheckBoxState(NomChamp) = cbChecked) then
      SetControlChecked(NomChamp, False);
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 03/12/2002
Modifié le ... : 03/12/2002
Description .. : Met un champ en erreur et affiche le message associé
Mots clefs ... :
*****************************************************************}

procedure TOM_AFFCDEENTETE.MetEnErreur(NomChamp: string; NoMsg: integer);
begin
  if NomChamp <> '' then SetFocusControl(NomChamp);
  LastError := NoMsg;
  if (NoMsg >= Low(TexteMessage)) and (NoMsg <= High(TexteMessage)) then
    LastErrorMsg := TexteMessage[NoMsg];
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 03/12/2002
Modifié le ... : 03/12/2002
Description .. : Vérifie l'intégrité des données
Mots clefs ... :
*****************************************************************}

function TOM_AFFCDEENTETE.ControleIntegrite: Boolean;
var
  Pourc: integer;
  Ctrl: TControl;
begin
  Result := False;
  // contrôle du dépôt
  if GetField('GEA_DEPOT') = '' then
  begin
    MetEnErreur('GEA_DEPOT', 1);
    Exit;
  end;
  // contrôle du % maxi de fractionnement
  Pourc := GetField('GEA_POURCENTAGE');
  if (Pourc < 0) or (Pourc > 100) then
  begin
    MetEnErreur('GEA_POURCENTAGE', 2);
    Exit;
  end;
  // contrôle du % minimum en quantité
  Pourc := GetField('GEA_POUMINIQTECDE');
  if (Pourc < 0) or (Pourc > 100) then
  begin
    MetEnErreur('GEA_POUMINIQTECDE', 2);
    Exit;
  end;
  // contrôle du % minimum en montant
  Pourc := GetField('GEA_POUMINIMTCDE');
  if (Pourc < 0) or (Pourc > 100) then
  begin
    MetEnErreur('GEA_POUMINIMTCDE', 2);
    Exit;
  end;
  // contrôle de la nature de pièce
  Ctrl := GetControl('GEA_NATPIECEGRP');
  if (Ctrl = nil) or not (LookupValueExist(Ctrl)) then
  begin
    MetEnErreur('GEA_NATPIECEGRP', 3);
    Exit;
  end;
  Result := True;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 03/12/2002
Modifié le ... : 03/12/2002
Description .. : Active ou désactive un champ et son libellé
Mots clefs ... :
*****************************************************************}

procedure TOM_AFFCDEENTETE.ActiveChamp(NomChamp: string; Active: Boolean);
begin
  SetControlEnabled(NomChamp, Active);
  SetControlEnabled('T' + NomChamp, Active);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 03/12/2002
Modifié le ... : 03/12/2002
Description .. : Modification de l'indicateur de gestion des réliquats
Mots clefs ... :
*****************************************************************}

procedure TOM_AFFCDEENTETE.OnChangeGereReliquat;
var Active: boolean;
begin
  Active := (GetCheckBoxState('GEA_GERERELIQUAT') = cbChecked);
  ActiveChamp('GPK_REMPOURMAX2', Active);
  if not Active then
    SetControlProperty('GEA_POURCENTAGE', 'Value', 0);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 03/12/2002
Modifié le ... : 03/12/2002
Description .. : OnNewRecord
Mots clefs ... :
*****************************************************************}

procedure TOM_AFFCDEENTETE.OnNewRecord;
begin
  inherited;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 03/12/2002
Modifié le ... : 03/12/2002
Description .. : OnDeleteRecord
Mots clefs ... :
*****************************************************************}

procedure TOM_AFFCDEENTETE.OnDeleteRecord;
begin
  inherited;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 03/12/2002
Modifié le ... : 03/12/2002
Description .. : OnUpdateRecord
Mots clefs ... :
*****************************************************************}

procedure TOM_AFFCDEENTETE.OnUpdateRecord;
begin
  inherited;
  // Contrôle de l'intégrité des données
  ControleIntegrite;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 03/12/2002
Modifié le ... : 03/12/2002
Description .. : OnAfterUpdateRecord
Mots clefs ... :
*****************************************************************}

procedure TOM_AFFCDEENTETE.OnAfterUpdateRecord;
begin
  inherited;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 03/12/2002
Modifié le ... : 03/12/2002
Description .. : OnLoadRecord
Mots clefs ... :
*****************************************************************}

procedure TOM_AFFCDEENTETE.OnLoadRecord;
begin
  inherited;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 03/12/2002
Modifié le ... : 03/12/2002
Description .. : OnChangeField
Mots clefs ... :
*****************************************************************}

procedure TOM_AFFCDEENTETE.OnChangeField(F: TField);
begin
  inherited;
  if F.FieldName = 'GEA_PRIOAFFCDE1' then
    ActivePrioAffCde(1)
  else
    if F.FieldName = 'GEA_PRIOAFFCDE2' then
    ActivePrioAffCde(2)
  else
    if F.FieldName = 'GEA_PRIOAFFCDE3' then
    ActivePrioAffCde(3)
  else
    if F.FieldName = 'GEA_GERERELIQUAT' then
    OnChangeGereReliquat;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 03/12/2002
Modifié le ... : 03/12/2002
Description .. : OnArgument
Mots clefs ... :
*****************************************************************}

procedure TOM_AFFCDEENTETE.OnArgument(S: string);
begin
  inherited;
  ChangeComboLibre;
  if V_PGI.SAV then
  begin
    SetControlEnabled('GEA_TERMINE', True);
    SetControlProperty('GEA_TERMINE', 'ReadOnly', False);
    SetControlEnabled('GEA_STATUTAFF', True);
    SetControlProperty('GEA_STATUTAFF', 'Style', csDropDownList);
    SetControlProperty('GEA_STATUTAFF', 'ReadOnly', False);
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 03/12/2002
Modifié le ... : 03/12/2002
Description .. : OnClose
Mots clefs ... :
*****************************************************************}

procedure TOM_AFFCDEENTETE.OnClose;
begin
  inherited;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 03/12/2002
Modifié le ... : 03/12/2002
Description .. : OnCancelRecord
Mots clefs ... :
*****************************************************************}

procedure TOM_AFFCDEENTETE.OnCancelRecord;
var Ind: integer;
begin
  inherited;
  for Ind := 1 to 3 do ActivePrioAffCde(Ind);
  OnChangeGereReliquat;
end;

initialization
  registerclasses([TOM_AFFCDEENTETE]);
end.
