{Source de la tof TRQRJALCOMMISSION
--------------------------------------------------------------------------------------
  Version    |  Date   | Qui |   Commentaires
--------------------------------------------------------------------------------------
 6.0x.xxx.xxx  29/07/04  JP   Création de l'unité
 7.05.001.001  24/10/06  JP   Gestion des filtres multi sociétés
                              Mise en place de l'ancêtre des états
 8.10.001.004  08/08/07  JP   Gestion des confidentialités
--------------------------------------------------------------------------------------}

unit TRQRJALCOMMISSION_TOF ;

interface

uses
  StdCtrls, Classes,
{$IFNDEF EAGLCLIENT}
  QRS1, FE_Main, 
{$ELSE}
  eQRS1, MaineAGL,
{$ENDIF}
  SysUtils, ExtCtrls, HCtrls, HEnt1, UTOF, uAncetreEtat;

Type
  TOF_TRQRJALCOMMISSION = class (TRANCETREETAT)
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

    procedure TypeDateChange(Sender : TObject);
    procedure InitDates;{FQ 10097}
  end ;

procedure TRLanceFiche_JalCommission(Dom: string; Fiche: string; Range: string; Lequel: string; Arguments: string);

implementation

uses
  {$IFDEF TRCONF}
  ULibConfidentialite,
  {$ENDIF TRCONF}
  Constantes, Commun;

{---------------------------------------------------------------------------------------}
procedure TRLanceFiche_JalCommission(Dom: string; Fiche: string; Range: string; Lequel: string; Arguments: string);
{---------------------------------------------------------------------------------------}
begin
  AglLanceFiche(Dom, Fiche, Range, Lequel, Arguments);
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_TRQRJALCOMMISSION.OnUpdate ;
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
procedure TOF_TRQRJALCOMMISSION.OnLoad ;
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
procedure TOF_TRQRJALCOMMISSION.OnArgument (S : String ) ;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  cbNature  := THValComboBox(GetControl('NATURE'));
  cbDate    := THValComboBox(GetControl('OPERATION'));
  cbCodeFx  := THMultiValComboBox(GetControl('TE_CODEFLUX'));
  gpRupture := TRadioGroup(GetControl('RGRUPTURE'));
  cbDate.OnChange := TypeDateChange;
  cbDate.ItemIndex := 0;
  {24/10/06 : On filtre les comptes en fonction des sociétés du regroupement Tréso}
  THValComboBox(GetControl('TE_GENERAL')).Plus := FiltreBanqueCp(THValComboBox(GetControl('TE_GENERAL')).DataType, '', '');
end;

{Selon le type de date sur lesquelles on travaille (Opération / Valeur), on rend les critères
 d'impression idoines visibles ou non
{---------------------------------------------------------------------------------------}
procedure TOF_TRQRJALCOMMISSION.TypeDateChange(Sender : TObject);
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
procedure TOF_TRQRJALCOMMISSION.InitDates;
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
  RegisterClasses([TOF_TRQRJALCOMMISSION]);

end.

