{Source de la tof TRQRPREVREAL
--------------------------------------------------------------------------------------
  Version    |  Date   | Qui |   Commentaires
--------------------------------------------------------------------------------------
   0.91        06/11/03  JP   création de l'unité
 1.2X.001.001  09/04/04  JP   Problème sur l'initialisation des dates dans le OnLoad
                              car le ValOk n'est pas initialisé
 1.90.xxx.xxx  18/06/04  JP   Modification de l'initialisation des dates FQ 10097
 7.05.001.001  24/10/06  JP   Gestion des filtres multi sociétés
                              Mise en place de l'ancêtre des états
 8.10.001.004  08/08/07  JP   Gestion des confidentialités
--------------------------------------------------------------------------------------}

unit TRQRPREVREAL_TOF ;

interface

uses
  StdCtrls, Controls, Classes,
{$IFNDEF EAGLCLIENT}
  db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} QRS1, FE_Main,
{$ELSE}
  eQRS1, MaineAGL, UTob,
{$ENDIF}
  SysUtils, UTOF, HCtrls, uAncetreEtat, ExtCtrls;

Type
  TOF_TRQRPREVREAL = class (TRANCETREETAT)
    procedure OnUpdate              ; override;
    procedure OnLoad                ; override;
    procedure OnArgument(S : string); override;
  private
    Where     : string;
    cbNature  : THValComboBox;
    cbDate    : THValComboBox;
    cbCodeFx  : THMultiValComboBox;
    gpRupture : TRadioGroup;
    ValOk     : Boolean; {True sur les dates de valeur, False sur les dates d'opération}

    procedure ChargerCodeFlux;
    procedure TypeDateChange(Sender : TObject);
    procedure InitDates;{FQ 10097}
  end ;

procedure TRLanceFiche_PrevRealise(Dom: string; Fiche: string; Range: string; Lequel: string; Arguments: string);

implementation

uses
  {$IFDEF TRCONF}
  ULibConfidentialite,
  {$ENDIF TRCONF}
  Constantes, HEnt1, Commun;

{---------------------------------------------------------------------------------------}
procedure TRLanceFiche_PrevRealise(Dom: string; Fiche: string; Range: string; Lequel: string; Arguments: string);
{---------------------------------------------------------------------------------------}
begin
  AglLanceFiche(Dom, Fiche, Range, Lequel, Arguments);
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_TRQRPREVREAL.OnUpdate ;
{---------------------------------------------------------------------------------------}
begin
  {$IFDEF TRCONF}
  TypeConfidentialite := tyc_Banque + ';';
  {$ENDIF TRCONF}
  inherited;
  Where := TFQRS1(Ecran).WhereSQL;
  if cbNature.Value <> '' then begin
    if Trim(Where ) <> '' then Where := Where + ' AND ';
    if cbNature.Value = na_Realise then Where := Where + 'TE_NATURE = "'+ na_Realise + '"'
                                   else Where := Where + 'TE_NATURE <> "'+ na_Realise + '"'
  end;
  TFQRS1(Ecran).WhereSQL := Where;
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_TRQRPREVREAL.OnLoad ;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  TypeDateChange(cbDate);

  {Gestion des ruptures : elle peuvent étre par compte ou par code flux}
  if gpRupture.ItemIndex = 0 then SetControlText('XX_RUPTURE', 'TE_GENERAL')
                             else SetControlText('XX_RUPTURE', 'TE_CODEFLUX');
  SetControlText('RUPTURE', GetControlText('XX_RUPTURE'));
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_TRQRPREVREAL.OnArgument (S : String ) ;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  Ecran.HelpContext := 150;

  cbNature  := THValComboBox(GetControl('NATURE'));
  cbDate    := THValComboBox(GetControl('OPERATION'));
  cbCodeFx  := THMultiValComboBox(GetControl('TE_CODEFLUX'));
  gpRupture := TRadioGroup(GetControl('RGRUPTURE'));
  ChargerCodeFlux;
  cbDate.OnChange := TypeDateChange;
  cbDate.ItemIndex := 0;
  {24/10/06 : On filtre les comptes en fonction des sociétés du regroupement Tréso}
  if not EtatMD then
    THValComboBox(GetControl('TE_GENERAL')).Plus := FiltreBanqueCp(THValComboBox(GetControl('TE_GENERAL')).DataType, '', '');
end ;

{Chargement de la combo contenant les codes flux à patir des tables FLUXTRESO pour les
 écritures de la trésorerie et RUBRIQUES pour les écritures de la comptabilité
{---------------------------------------------------------------------------------------}
procedure TOF_TRQRPREVREAL.ChargerCodeFlux;
{---------------------------------------------------------------------------------------}
var
  Q : TQuery;
begin
  cbCodeFx.Items.Clear;
  cbCodeFx.Values.Clear;

  Q := OpenSQL('SELECT TFT_FLUX, TFT_LIBELLE FROM FLUXTRESO', True);
  while not Q.EOF do begin
    cbCodeFx.Items.Add(Q.Fields[1].AsString);
    cbCodeFx.Values.Add(Q.Fields[0].AsString);
    Q.Next;
  end;
  Ferme(Q);

  Q := OpenSQL('SELECT RB_RUBRIQUE, RB_LIBELLE FROM RUBRIQUE WHERE RB_NATRUB = "TRE" AND RB_CLASSERUB = "TRE"', True);
  while not Q.EOF do begin
    cbCodeFx.Items.Add(Q.Fields[1].AsString);
    cbCodeFx.Values.Add(Q.Fields[0].AsString);
    Q.Next;
  end;
  Ferme(Q);
end;

{Selon le type de date sur lesquelles on travaille (Opération / Valeur), on rend les critères
 d'impression idoines visibles ou non
{---------------------------------------------------------------------------------------}
procedure TOF_TRQRPREVREAL.TypeDateChange(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  ValOk := cbDate.Value = 'V';
  SetControlVisible('TE_DATEVALEUR' , ValOk);
  SetControlVisible('TE_DATEVALEUR_', ValOk);
  SetControlVisible('TE_DATECOMPTABLE' , not ValOk);
  SetControlVisible('TE_DATECOMPTABLE_', not ValOk);
  if ValOk then SetControlText('TTE_DATEVALEUR', 'Date de valeur')
           else SetControlText('TTE_DATEVALEUR', 'Date d''opération');
  InitDates; {FQ 10097}
end;

{18/06/04 : les dates sont réinitialisées a chaque changement de type pour plus de cohérence (FQ 10097)
{---------------------------------------------------------------------------------------}
procedure TOF_TRQRPREVREAL.InitDates;
{---------------------------------------------------------------------------------------}
begin
  {Pour éviter un message du genre "__/__/____ n'est pas une date correct" :
   - On initialise les champs invisibles
   - On initialise les champs visibles s'il sont vides}
  if ValOk then begin
    SetControlText('TE_DATECOMPTABLE' , StDate1900);
    SetControlText('TE_DATECOMPTABLE_', StDate2099);
    if (Trim(GetControlText('TE_DATEVALEUR_')) = '/  /') or (GetControlText('TE_DATEVALEUR_') = StDate2099) then
      SetControlText('TE_DATEVALEUR_', DateToStr(FinDeMois(V_PGI.DateEntree)));
    if (Trim(GetControlText('TE_DATEVALEUR')) = '/  /') or (GetControlText('TE_DATEVALEUR') = StDate1900)  then
      SetControlText('TE_DATEVALEUR', DateToStr(DebutDeMois(V_PGI.DateEntree)));
    SetControlText('CLE', 'TE_CLEVALEUR');
  end else begin
    SetControlText('TE_DATEVALEUR' , StDate1900);
    SetControlText('TE_DATEVALEUR_', StDate2099);
    if (Trim(GetControlText('TE_DATECOMPTABLE_')) = '/  /') or (GetControlText('TE_DATECOMPTABLE_') = StDate2099) then
      SetControlText('TE_DATECOMPTABLE_', DateToStr(FinDeMois(V_PGI.DateEntree)));
    if (Trim(GetControlText('TE_DATECOMPTABLE')) = '/  /') or (GetControlText('TE_DATECOMPTABLE') = StDate1900) then
      SetControlText('TE_DATECOMPTABLE', DateToStr(DebutDeMois(V_PGI.DateEntree)));
    SetControlText('CLE', 'TE_CLEOPERATION');
  end;
end;

initialization
  RegisterClasses([TOF_TRQRPREVREAL]);

end.

