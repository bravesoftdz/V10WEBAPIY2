{***********UNITE*************************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 07/09/2000
Modifié le ... : 23/07/2001
Description .. : Source TOF de la FICHE : ASSISTCLAVECR
Suite ........ : Assistant pour la constitution d'une requête SELECT
Suite ........ : associée à un bouton du pavé tactile.
Mots clefs ... : TOF;UTOFMFOASSISTCLAVECR;FO
*****************************************************************}
unit UtofMFOAssistClavEcr;

interface
uses
  Classes, Controls, StdCtrls, Forms, Vierge, SysUtils, HCtrls, HEnt1, UTOF,
  HTB97, LookUp, HMsgBox;

type
  TOF_MFOASSISTCLAVECR = class(TOF)
    procedure OnArgument(stArgument: string); override;
    procedure OnUpdate; override;
    procedure OnClose; override;
  private
    NomTable: string;
    PrefixeTable: string;
    CleTable: string;
    ChampLibelle: string;
    RequeteSQL: string;
    CanClose: boolean;
    procedure ChargeComboDEChamps;
    procedure AjoutOnChange(sNomChamp: string; FctOnChange: TNotifyEvent);
    procedure BrancheOnChange;
    procedure ConstitueSelect;
    procedure AfficheSelect;
    procedure BTestSelectClick(Sender: TObject);
    procedure ActiveTestRequete;
    procedure ConvertCondition(Chaine: string; NumCond: integer);
    procedure RabLigneSelection(NoLig: Integer);
    procedure RemplaceLigneSelection;
    procedure ImpactTypeChamp(sNomChamp, sNomValeur: string);
    procedure OnChange(Sender: TObject);
    procedure OnChangeSelChamp1(Sender: TObject);
    procedure OnChangeSelChamp2(Sender: TObject);
  end;

implementation

procedure FOAssistClavEcrAction(Parms: array of variant; Nb: integer);
var St: string;
  FF: TForm;
  LaTof: TOF;
begin
  FF := TForm(Integer(Parms[0]));
  if (FF is TFVierge) then LaTof := TFVierge(FF).LaTof else Exit;
  if LaTof is TOF_MFOASSISTCLAVECR then
  begin
    St := UpperCase(string(Parms[1]));
    if St = 'AFFICHESELECT' then TOF_MFOASSISTCLAVECR(LaTof).AfficheSelect;
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  InitialiseCombo : vide le contenu d'un Combo.
///////////////////////////////////////////////////////////////////////////////////////

procedure InitialiseCombo(Combo: THValComboBox);
begin
  Combo.Values.Clear;
  Combo.Items.Clear;
  if Combo.Vide then
  begin
    if Combo.VideString = '' then Combo.Items.Add(Traduirememoire('<<Tous>>'))
    else Combo.Items.Add(Traduirememoire(Combo.VideString));
    Combo.Values.Add('');
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  RecopieCombo : recopie le contenu d'un Combo dans un autre.
///////////////////////////////////////////////////////////////////////////////////////

procedure RecopieCombo(ccOrg, ccDest: TControl);
var Ind: Integer;
  Org, Dest: THValComboBox;
begin
  if (ccOrg = nil) or not (ccOrg is THValComboBox) then Exit;
  if (ccDest = nil) or not (ccDest is THValComboBox) then Exit;
  Org := THValComboBox(ccOrg);
  Dest := THValComboBox(ccDest);
  InitialiseCombo(Dest);
  for Ind := 0 to Org.Items.Count - 1 do if Org.Values[Ind] <> '' then
    begin
      Dest.Values.Add(Org.Values[Ind]);
      Dest.Items.Add(Org.Items[Ind]);
    end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  ChargeComboDEChamps : Charge le combo pour sélectionner les champs de la table.
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOASSISTCLAVECR.ChargeComboDEChamps;
var CC: TControl;
  NumTable, Ind: Integer;
begin
  CC := GetControl('CTEXCHAMP');
  if (CC = nil) or not (CC is THValComboBox) then Exit;
  InitialiseCombo(THValComboBox(CC));
  if PrefixeTable = '' then NumTable := -1 else NumTable := PrefixeToNum(PrefixeTable);
  if NumTable > 0 then
    for Ind := low(V_PGI.DEChamps[NumTable]) to high(V_PGI.DEChamps[NumTable]) do
      if trim(V_PGI.DEChamps[NumTable, Ind].Nom) <> '' then
      begin
        THValComboBox(CC).Values.Add(V_PGI.DEChamps[NumTable, Ind].Nom);
        THValComboBox(CC).Items.Add(V_PGI.DEChamps[NumTable, Ind].Libelle);
      end;
  RecopieCombo(CC, GetControl('CSELCHAMP1'));
  RecopieCombo(CC, GetControl('CSELCHAMP2'));
  RecopieCombo(CC, GetControl('CTRICHAMP'));
end;

///////////////////////////////////////////////////////////////////////////////////////
//  AjoutOnChange : ajoute le traitement de l'événement OnChange sur un contrôle
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOASSISTCLAVECR.AjoutOnChange(sNomChamp: string; FctOnChange: TNotifyEvent);
var CC: TControl;
begin
  CC := GetControl(sNomChamp);
  if CC = nil then Exit;
  if CC is THValComboBox then THValComboBox(CC).OnChange := FctOnChange else
    if CC is THEdit then THEdit(CC).OnExit := FctOnChange else
    if CC is TRadioButton then TRadioButton(CC).OnClick := FctOnChange;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  BrancheOnChange : branche le traitement de l'événement OnChange sur les contrôles de la forme.
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOASSISTCLAVECR.BrancheOnChange;
begin
  AjoutOnChange('CTEXCHAMP', OnChange);
  AjoutOnChange('CSELCHAMP1', OnChangeSelChamp1);
  AjoutOnChange('CSELOPE1', OnChange);
  AjoutOnChange('CSELVALEUR1', OnChange);
  AjoutOnChange('CSELCHAMP2', OnChangeSelChamp2);
  AjoutOnChange('CSELOPE2', OnChange);
  AjoutOnChange('CSELVALEUR2', OnChange);
  AjoutOnChange('CTRICHAMP', OnChange);
  AjoutOnChange('CTRIASC', OnChange);
  AjoutOnChange('CTRIDESC', OnChange);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  TraduitOperateur : traduit l'opérateur saisi en langage SQL
///////////////////////////////////////////////////////////////////////////////////////

function TraduitOperateur(sOpe, sVal: string): string;
begin
  if sOpe = 'C' then // commence par
  begin
    sOpe := ' LIKE ';
    sVal := sVal + '%';
  end else
    if sOpe = 'D' then // ne commence pas par
  begin
    sOpe := ' NOT LIKE ';
    sVal := sVal + '%';
  end else
    if sOpe = 'L' then // contient
  begin
    sOpe := ' LIKE ';
    sVal := '%' + sVal + '%';
  end else
    if sOpe = 'M' then // ne contient pas
  begin
    sOpe := ' NOT LIKE ';
    sVal := '%' + sVal + '%';
  end;
  Result := sOpe + '"' + sVal + '"';
end;

///////////////////////////////////////////////////////////////////////////////////////
//  ConvertOperateur : convertit l'opérateur SQL selon le code de TTCOMPARE
///////////////////////////////////////////////////////////////////////////////////////

function ConvertOperateur(sOpe, sVal: string): string;
var Deb, Fin: boolean;
begin
  Result := sOpe;
  Deb := ((Length(sVal) > 0) and (sVal[1] = '%'));
  Fin := ((Length(sVal) > 0) and (sVal[Length(sVal)] = '%'));
  if sOpe = 'LIKE' then
  begin
    if (Deb) and (Fin) then
      Result := 'L' // contient
    else
      Result := 'C'; // commence par
  end else
    if sOpe = 'NOT LIKE' then
  begin
    if (Deb) and (Fin) then
      Result := 'M' // ne contient pas
    else
      Result := 'D'; // ne commence pas par
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  ConstitueSelect : constitue à partir des données saisies la requête SQL générée
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOASSISTCLAVECR.ConstitueSelect;
var sSql, sChp, sOpe: string;
  CC: TControl;
begin
  // Clause Select
  RequeteSQL := 'SELECT ' + GetControlText('CTEXCHAMP');
  if GetControlText('CTEXCHAMP') <> CleTable then RequeteSQL := RequeteSQL + ',' + CleTable;
  // Clause From
  RequeteSQL := RequeteSQL + ' FROM ' + NomTable;
  // Ajout clause Where
  sSql := ' WHERE ';
  sChp := GetControlText('CSELCHAMP1');
  sOpe := GetControlText('CSELOPE1');
  if (sChp <> '') and (sOpe <> '') then
  begin
    RequeteSQL := RequeteSQL + sSql + sChp;
    RequeteSQL := RequeteSQL + TraduitOperateur(sOpe, GetControlText('CSELVALEUR1'));
    sSql := ' AND ';
  end;
  sChp := GetControlText('CSELCHAMP2');
  sOpe := GetControlText('CSELOPE2');
  if (sChp <> '') and (sOpe <> '') then
  begin
    RequeteSQL := RequeteSQL + sSql + sChp;
    RequeteSQL := RequeteSQL + TraduitOperateur(sOpe, GetControlText('CSELVALEUR2'));
  end;
  if GetControlText('CTRICHAMP') <> '' then
  begin
    RequeteSQL := RequeteSQL + ' ORDER BY ' + GetControlText('CTRICHAMP');
    CC := GetControl('CTRIDESC');
    if (CC <> nil) and (CC is TRadioButton) and (TRadioButton(CC).Checked) then RequeteSQL := RequeteSQL + ' DESC';
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  AfficheSelect : affiche la requête SQL générée
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOASSISTCLAVECR.AfficheSelect;
begin
  ConstitueSelect;
  SetControlText('CSQL', RequeteSQL);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  BTestSelectClick : teste la requête SQL générée
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOASSISTCLAVECR.BTestSelectClick(Sender: TObject);
var CC: TControl;
begin
  SetControlText('CRESULTSQL', '');
  CC := GetControl('CRESULTSQL');
  AfficheSelect;
  LookUpList(CC, 'Résultat de la requête', NomTable, CleTable, ChampLibelle, '', '', True, 0, RequeteSQL, tlDefault);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  ActiveTestRequete : ajoute le test de la requête sur l'événement OnClick du bouton de test
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOASSISTCLAVECR.ActiveTestRequete;
var CC: TControl;
begin
  CC := GetControl('BTESTSELECT');
  if (CC <> nil) and (CC is TToolbarButton97) then TToolbarButton97(CC).OnClick := BTestSelectClick;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  ExtractChamp : extrait un champ d'un select
///////////////////////////////////////////////////////////////////////////////////////

function ExtractChamp(Chaine: string; var Indice: integer; AvecOper: boolean = True): string;
const Liste_delimiters: string = ' ,"'; // liste des separateurs
var Nbc, ii: integer;
  LstDelim: string;
begin
  LstDelim := Liste_delimiters;
  if AvecOper then LstDelim := LstDelim + '=<>';
  Nbc := 0;
  for ii := Indice to Length(Chaine) do
  begin
    if IsDelimiter(LstDelim, Chaine, ii) then Break;
    Inc(Nbc);
  end;
  Result := Copy(Chaine, Indice, Nbc);
  while (ii <= Length(Chaine)) and IsDelimiter(Liste_delimiters, Chaine, ii) do Inc(ii);
  Indice := ii;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  ConvertCondition : convertit une condition d'un select
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOASSISTCLAVECR.ConvertCondition(Chaine: string; NumCond: integer);
var sChamp, sOpe, sVal: string;
  Ind: integer;
begin
  Ind := 1;
  sChamp := ExtractChamp(Chaine, Ind);
  SetControlText('CSELCHAMP' + IntToStr(NumCond), sChamp);
  sOpe := ExtractChamp(Chaine, Ind, False);
  if sOpe = 'NOT' then sOpe := sOpe + ' ' + ExtractChamp(Chaine, Ind, False);
  sVal := ExtractChamp(Chaine, Ind);
  SetControlText('CSELOPE' + IntToStr(NumCond), ConvertOperateur(sOpe, sVal));
  SetControlText('CSELVALEUR' + IntToStr(NumCond), FindEtReplace(sVal, '%', '', True));
  ImpactTypeChamp(sChamp, 'CSELVALEUR' + IntToStr(NumCond));
end;

///////////////////////////////////////////////////////////////////////////////////////
//  OnArgument
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOASSISTCLAVECR.OnArgument(stArgument: string);
var sArg, sSql: string;
  Ind, ii, Nbc: integer;
  OkSelect: boolean;
  CC: TControl;
begin
  inherited;
  // Initialisations
  CanClose := True;
  NomTable := '';
  PrefixeTable := '';
  CleTable := '';
  ChampLibelle := '';
  RequeteSQL := '';
  TFVierge(Ecran).Retour := '';
  sArg := UpperCase(Trim(stArgument));
  // Recherche de la table
  if Copy(sArg, 1, 6) = 'TABLE=' then
  begin
    sArg := Copy(sArg, 7, Length(sArg) - 6);
    NomTable := ReadTokenst(sArg);
  end;
  // Interprétation de la requête
  OkSelect := (Copy(sArg, 1, 6) = 'SELECT');
  if OkSelect then
  begin
    sSql := Copy(sArg, 8, Length(sArg) - 7);
    RequeteSQL := sArg;
    // Affiche dans le memo la requête SQL
    SetControlText('CSQL', sArg);
    // Recherche de la table si besoin
    if NomTable <> '' then
    begin
      Ind := Pos(' FROM ', sSql);
      if Ind < 1 then Exit;
      Inc(Ind, 6);
      NomTable := ExtractChamp(sSql, Ind);
    end;
    // Extraction du nom du champ libellé de la requête
    Ind := 1;
    ChampLibelle := ExtractChamp(sSql, Ind);
  end;
  // Définition de la table
  PrefixeTable := TableToPrefixe(NomTable);
  CleTable := TableToCle1(NomTable);
  if PrefixeTable <> '' then ChargeComboDEChamps;
  // Définition du nom du champ libellé
  if ChampToNum(ChampLibelle) <= 0 then ChampLibelle := CleTable;
  SetControlText('CTEXCHAMP', ChampLibelle);
  if OkSelect then
  begin
    // Extraction des conditions
    Ind := Pos(' WHERE ', sSql);
    if Ind > 0 then
    begin
      Inc(Ind, 7);
      Nbc := 1;
      ii := Pos(' AND ', sSql);
      if ii > Ind then
      begin
        ConvertCondition(Copy(sSql, Ind, (ii - Ind)), Nbc);
        Inc(Nbc);
        Ind := ii + 5;
      end;
      ii := Pos(' ORDER BY ', sSql);
      if ii <= Ind then ii := Length(sSql);
      ConvertCondition(Copy(sSql, Ind, (ii - Ind)), Nbc);
    end;
    // Extraction des ordres de tri
    Ind := Pos(' ORDER BY ', sSql);
    if Ind > 0 then
    begin
      Inc(Ind, 10);
      ChampLibelle := ExtractChamp(sSql, Ind);
      SetControlText('CTRICHAMP', ChampLibelle);
      ChampLibelle := ExtractChamp(sSql, Ind);
      CC := GetControl('CTRIDESC');
      if (CC <> nil) and (CC is TRadioButton) then TRadioButton(CC).Checked := (ChampLibelle = 'DESC');
    end;
  end;
  // Branchement du traitement du OnChange
  BrancheOnChange;
  ActiveTestRequete;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  OnUpDate : traitement du click sur le bouton valider
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOASSISTCLAVECR.OnUpdate;
var OldSql: string;
  Rep: integer;
begin
  inherited;
  OldSql := RequeteSQL;
  AfficheSelect;
  if OldSql <> RequeteSQL then
  begin
    Rep := PGIAskCancel('Confirmez-vous la mise à jour ?');
    if Rep <> mrYes then
    begin
      CanClose := (Rep <> mrCancel);
      LastError := 1;
      Exit;
    end;
  end;
  TFVierge(Ecran).Retour := RequeteSQL;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  OnClose
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOASSISTCLAVECR.OnClose;
begin
  inherited;
  if not CanClose then
  begin
    CanClose := True;
    LastError := 1;
    Exit;
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  OnChange : traitement de l'événement OnChange sur les contrôles de la forme.
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOASSISTCLAVECR.OnChange(Sender: TObject);
begin
  AfficheSelect;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  RabLigneSelection : remise à blanc de l'opérateur et de la valeur de sélection
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOASSISTCLAVECR.RabLigneSelection(NoLig: Integer);
begin
  if NoLig = 1 then
  begin
    SetControlText('CSELOPE1', '');
    SetControlText('CSELVALEUR1', '');
    ImpactTypeChamp('', 'CSELVALEUR1');
  end else
  begin
    SetControlText('CSELOPE2', '');
    SetControlText('CSELVALEUR2', '');
    ImpactTypeChamp('', 'CSELVALEUR2');
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  RemplaceLigneSelection : remplace la 1ère ligne de sélection par la 2ème
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOASSISTCLAVECR.RemplaceLigneSelection;
begin
  if (GetControlText('CSELCHAMP1') = '') and (GetControlText('CSELCHAMP2') <> '') then
  begin
    SetControlText('CSELCHAMP1', GetControlText('CSELCHAMP2'));
    SetControlText('CSELOPE1', GetControlText('CSELOPE2'));
    SetControlText('CSELVALEUR1', GetControlText('CSELVALEUR2'));
    SetControlText('CSELCHAMP2', '');
    RabLigneSelection(2);
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  ImpactTypeChamp : impact du choix du champ de sélection sur le format de saisie de la valeur
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOASSISTCLAVECR.ImpactTypeChamp(sNomChamp, sNomValeur: string);
var sType, sTablette: string;
  CC: TControl;
begin
  CC := GetControl(sNomValeur);
  if (CC = nil) or not (CC is THEdit) then Exit;
  sType := ChampToType(sNomChamp);
  sTablette := '';
  if sType = 'COMBO' then
  begin
    THEdit(CC).CharCase := ecUpperCase;
    THEdit(CC).MaxLength := 3;
    //THEdit(CC).EditMask := '' ;
    sTablette := Get_Join(sNomChamp);
  end else
    if sType = 'BOOLEAN' then
  begin
    THEdit(CC).CharCase := ecUpperCase;
    THEdit(CC).MaxLength := 1;
    //THEdit(CC).EditMask := '' ;
  end else
    if sType = 'DATE' then
  begin
    THEdit(CC).CharCase := ecNormal;
    THEdit(CC).MaxLength := 10;
    //THEdit(CC).EditMask := DateMask ;
  end else
    if sType = 'DOUBLE' then
  begin
    THEdit(CC).CharCase := ecNormal;
    THEdit(CC).MaxLength := 0;
    //THEdit(CC).EditMask := StrfMask(V_PGI.OkDecV, '', False) ;
  end else
    if sType = 'INTEGER' then
  begin
    THEdit(CC).CharCase := ecNormal;
    THEdit(CC).MaxLength := 0;
    //THEdit(CC).EditMask := StrfMask(0, '', False) ;
  end else
  begin
    THEdit(CC).CharCase := ecNormal;
    THEdit(CC).MaxLength := 0;
    //THEdit(CC).EditMask := '' ;
    sTablette := Get_Join(sNomChamp);
  end;
  if sTablette = '' then
  begin
    THEdit(CC).DataType := '';
    THEdit(CC).ElipsisButton := False;
  end else
  begin
    THEdit(CC).DataType := sTablette;
    THEdit(CC).ElipsisButton := True;
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  OnChangeSelChamp1 : traitement de l'événement OnChange sur le champs de sélection n°1.
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOASSISTCLAVECR.OnChangeSelChamp1(Sender: TObject);
var sChp: string;
begin
  sChp := GetControlText('CSELCHAMP1');
  if sChp = '' then
  begin
    if GetControlText('CSELCHAMP2') = '' then RabLigneSelection(1)
    else RemplaceLigneSelection;
  end else
  begin
    ImpactTypeChamp(sChp, 'CSELVALEUR1');
  end;
  AfficheSelect;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  OnChangeSelChamp2 : traitement de l'événement OnChange sur le champs de sélection n°2.
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOASSISTCLAVECR.OnChangeSelChamp2(Sender: TObject);
var sChp: string;
begin
  sChp := GetControlText('CSELCHAMP2');
  if sChp <> '' then
  begin
    if GetControlText('CSELCHAMP1') = '' then RemplaceLigneSelection
    else ImpactTypeChamp(sChp, 'CSELVALEUR2');
  end else RabLigneSelection(2);
  AfficheSelect;
end;

initialization
  RegisterClasses([TOF_MFOASSISTCLAVECR]);
end.
