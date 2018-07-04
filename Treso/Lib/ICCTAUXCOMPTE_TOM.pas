{ Unité : Source TOM de la Table : ICCTAUXCOMPTE
--------------------------------------------------------------------------------------
    Version    |  Date  | Qui |   Commentaires
--------------------------------------------------------------------------------------
 7.05.001.001   17/10/06  JP   Création de l'unité : Tom de la table ICCTAUXCOMPTE
--------------------------------------------------------------------------------------}
unit ICCTAUXCOMPTE_TOM;

interface

uses {$IFDEF VER150} variants,{$ENDIF}
  {$IFDEF EAGLCLIENT}
  MaineAGL, eFiche, UTob,
  {$ELSE}
  db, FE_Main, Fiche, {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  {$ENDIF}
  Classes, UTOM, SysUtils;

type
  TOM_ICCTAUXCOMPTE = class(TOM)
    procedure OnNewRecord              ; override;
    procedure OnArgument   (S : string); override;
    procedure OnClose                  ; override;
    procedure OnUpdateRecord           ; override;
    procedure OnChangeField(F : TField); override;
  private
    function TestCoherenceDates : Boolean;
  end;


function TRLanceFiche_TRTauxIcc(Range, Lequel, Arguments : string) : string;


implementation

uses
  HEnt1, HCtrls, HMsgBox, Constantes, Commun;

{---------------------------------------------------------------------------------------}
function TRLanceFiche_TRTauxIcc(Range, Lequel, Arguments : string) : string;
{---------------------------------------------------------------------------------------}
begin
  Result := AglLanceFiche('TR', 'TRICCTAUXCOMPTE', Range, Lequel, Arguments);
end;

{---------------------------------------------------------------------------------------}
procedure TOM_ICCTAUXCOMPTE.OnArgument(S : string);
{---------------------------------------------------------------------------------------}
begin
  inherited;
  Ecran.HelpContext := 50000155;
  {On limite l'affichage aux comptes courants des dossiers du regroupement Tréso}
  THEdit(GetControl('ICD_GENERAL')).Plus := FiltreBanqueCp(tcp_Tous, tcb_Courant, '');
end;

{---------------------------------------------------------------------------------------}
procedure TOM_ICCTAUXCOMPTE.OnNewRecord ;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  SetField('ICD_GENERAL', '');
  SetField('ICD_DATEDU', DebutAnnee(Date));
  SetField('ICD_DATEAU', FinAnnee(Date));
  SetField('ICD_TAUX', 0);
  SetControlCaption('LBDOM', '');
end;

{---------------------------------------------------------------------------------------}
procedure TOM_ICCTAUXCOMPTE.OnUpdateRecord;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  if VarToStr(GetField('ICD_DATEAU')) = '' then begin
    LastError := 1;
    LastErrorMsg := TraduireMemoire('Veuillez saisir une date de fin de validité.');
    SetFocusControl('ICD_DATEAU');
  end

  else if Valeur(GetField('ICD_DATEAU')) <= Valeur(GetField('ICD_DATEDU')) then begin
    LastError := 1;
    LastErrorMsg := TraduireMemoire('La date de fin doit être supérieure à celle de début.');
    SetFocusControl('ICD_DATEAU');
  end

  else if Valeur(GetField('ICD_TAUX')) = 0 then begin
    LastError := 1;
    LastErrorMsg := TraduireMemoire('Veuillez saisir un taux.');
    SetFocusControl('ICD_TAUX');
  end

  else if not TestCoherenceDates then begin
    LastError := 1;
    LastErrorMsg := TraduireMemoire('Les dates ne sont pas cohérentes avec les taux déjà saisis.');
    SetFocusControl('ICD_DATEDU');
  end;

end;

{---------------------------------------------------------------------------------------}
procedure TOM_ICCTAUXCOMPTE.OnClose;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  TFFiche(Ecran).Retour := GetControlText('ICD_TAUX')
end;

{---------------------------------------------------------------------------------------}
procedure TOM_ICCTAUXCOMPTE.OnChangeField(F: TField);
{---------------------------------------------------------------------------------------}
begin
  inherited;
  if F.FieldName = 'ICD_GENERAL' then begin
    if VarToStr(GetField('ICD_GENERAL')) <> '' then
      SetControlCaption('LBDOM', RechDom('TRBANQUECPALL', VarToStr(GetField('ICD_GENERAL')), False));
  end;
end;

{On s'assure que la période saisie ne chevauche pas une autre période
{---------------------------------------------------------------------------------------}
function TOM_ICCTAUXCOMPTE.TestCoherenceDates : Boolean;
{---------------------------------------------------------------------------------------}
var
  dd : TDateTime;
  df : TDateTime;
begin
  dd := VarToDateTime(GetField('ICD_DATEDU'));
  df := VarToDateTime(GetField('ICD_DATEAU'));
  Result := not ExisteSQl('SELECT ICD_DATEDU, ICD_DATEAU FROM ICCTAUXCOMPTE WHERE ICD_GENERAL = "' +
            VarToStr(GetField('ICD_GENERAL')) + '" AND (("' + UsDatetime(dd) +
            '" BETWEEN ICD_DATEDU AND ICD_DATEAU) OR ("' + UsDatetime(df) +
            '" BETWEEN ICD_DATEDU AND ICD_DATEAU)) AND NOT (ICD_DATEDU = "' + UsDatetime(dd) + {'" ' + 
            'AND ICD_DATEAU = "' + UsDatetime(df) + }'")');
end;

initialization
  RegisterClasses([TOM_ICCTAUXCOMPTE]);

end.

