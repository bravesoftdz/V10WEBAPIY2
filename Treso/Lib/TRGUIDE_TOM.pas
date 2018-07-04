{-------------------------------------------------------------------------------------
    Version    |   Date   | Qui | Commentaires
--------------------------------------------------------------------------------------
 8.01.001.010   29/03/07    JP   Création de l'unité : Tom de la table TRGUIDE
--------------------------------------------------------------------------------------}
unit TRGUIDE_TOM;

interface

uses
  {$IFDEF EAGLCLIENT}
  MaineAGL, UTob ,
  {$ELSE}
   HDB, db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} FE_Main,
  {$ENDIF}
  Controls, Classes, SysUtils, UTOM;

type
  TOM_TRGUIDE = class(TOM)
    procedure OnNewRecord              ; override;
    procedure OnUpdateRecord           ; override;
    procedure OnAfterUpdateRecord      ; override;
    procedure OnArgument   (S : string); override;
    procedure OnDeleteRecord           ; override;
    procedure OnChangeField(F : TField); override;
  private
    FSensCib : string;
    function VerifCoherenceCIBFlux : Boolean;
    function VerifExistence        : Boolean;
  end;

procedure TRLanceFiche_Guide(Range, Lequel, Arguments : string);

implementation

uses
  {$IFDEF VER150} Variants, {$ENDIF}
  HCtrls, Constantes,HMsgBox, HEnt1, Commun;

{---------------------------------------------------------------------------------------}
procedure TRLanceFiche_Guide(Range, Lequel, Arguments : string);
{---------------------------------------------------------------------------------------}
begin
  AglLanceFiche('TR', 'TRGUIDE', Range, Lequel, Arguments);
end;

{---------------------------------------------------------------------------------------}
procedure TOM_TRGUIDE.OnArgument(S : string);
{---------------------------------------------------------------------------------------}
begin
  inherited;
  Ecran.HelpContext := 50000127;
  {$IFDEF EAGLCLIENT}
  THEdit(GetControl('TGU_GENERAL')).Plus := FiltreBanqueCp(THEdit(GetControl('TGU_GENERAL')).DataType, '', '');
  {$ELSE}
  THDBEdit(GetControl('TGU_GENERAL')).Plus := FiltreBanqueCp(THDBEdit(GetControl('TGU_GENERAL')).DataType, '', '');
  {$ENDIF EAGLCLIENT}
end;

{---------------------------------------------------------------------------------------}
procedure TOM_TRGUIDE.OnNewRecord;
{---------------------------------------------------------------------------------------}
begin
  inherited;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_TRGUIDE.OnDeleteRecord;
{---------------------------------------------------------------------------------------}
begin
  inherited;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_TRGUIDE.OnChangeField(F : TField);
{---------------------------------------------------------------------------------------}
begin
  inherited;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_TRGUIDE.OnUpdateRecord;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  if VarToStr(GetField('TGU_CODEGUIDE')) = '' then begin
    LastError := 1;
    LastErrorMsg := TraduireMemoire('Veuillez saisir un code');
    SetFocusControl('TGU_CODEGUIDE');
  end
  else if VarToStr(GetField('TGU_LIBELLE')) = '' then begin
    LastError := 1;
    LastErrorMsg := TraduireMemoire('Veuillez saisir un libellé');
    SetFocusControl('TGU_LIBELLE');
  end
  else if VarToStr(GetField('TGU_CHAMPGUIDE')) = '' then begin
    LastError := 1;
    LastErrorMsg := TraduireMemoire('Veuillez saisir un champ de comparaison');
    SetFocusControl('TGU_CHAMPGUIDE');
  end
  else if VarToStr(GetField('TGU_VALEURCHAMP')) = '' then begin
    LastError := 1;
    LastErrorMsg := TraduireMemoire('Veuillez saisir une valeur de comparaison');
    SetFocusControl('TGU_VALEURCHAMP');
  end
  else if VarToStr(GetField('TGU_CODEFLUX')) = '' then begin
    LastError := 1;
    LastErrorMsg := TraduireMemoire('Veuillez saisir un flux de génération');
    SetFocusControl('TGU_CODEFLUX');
  end
  else if not VerifExistence then begin
    LastError := 1;
    LastErrorMsg := TraduireMemoire('Le cib n''existe pas');
    SetFocusControl('TGU_VALEURCHAMP');
  end
  else if not VerifCoherenceCIBFlux then begin
    LastError := 1;
    LastErrorMsg := TraduireMemoire('Le flux et le cib sont de signe opposé');
    SetFocusControl('TGU_CODEFLUX');
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_TRGUIDE.OnAfterUpdateRecord;
{---------------------------------------------------------------------------------------}
begin
  inherited;
end;

{Vérifie que le CIB est du même sens que le Flux
{---------------------------------------------------------------------------------------}
function TOM_TRGUIDE.VerifCoherenceCIBFlux : Boolean;
{---------------------------------------------------------------------------------------}
var
  Q    : TQuery;
  SQL  : string;
  Flux : string;
begin
  Result := True;
  if VarToStr(GetField('TGU_CHAMPGUIDE')) <> UpperCase(tc_CIB) then Exit;
  {Si le Cib est mixte, pas de problème}
  if FSensCib = 'MIX' then Exit;

  SQL := 'SELECT TTL_SENS FROM TYPEFLUX, FLUXTRESO WHERE TTL_TYPEFLUX = TFT_TYPEFLUX AND ' +
         'TFT_FLUX = "' + Flux + '"';

  Q := OpenSQL(SQL, True);
  if not Q.EOF then
    Result := ((Q.FindField('TTL_SENS').AsString = 'C') and (FSensCib = 'ENC')) or
              ((Q.FindField('TTL_SENS').AsString = 'D') and (FSensCib = 'DEC'));
  Ferme(Q);
end;

{---------------------------------------------------------------------------------------}
function TOM_TRGUIDE.VerifExistence: Boolean;
{---------------------------------------------------------------------------------------}
var
  Q    : TQuery;
  SQL  : string;
  Cpte : string;
begin
  Result := True;
  if VarToStr(GetField('TGU_CHAMPGUIDE')) <> UpperCase(tc_CIB) then Exit;

  Cpte := VarToStr(GetField('TGU_GENERAL'));
  if Cpte = '' then
    SQL := 'SELECT TCI_SENS FROM CIB WHERE TCI_BANQUE = "' + CODECIBREF + '"'
  else
    SQL := 'SELECT TCI_SENS FROM CIB, BANQUECP WHERE BQ_BANQUE = TCI_BANQUE AND ' +
           'BQ_CODE = "' + Cpte + '" AND TCI_CODECIB = "' + VarToStr(GetField('TGU_VALEURCHAMP')) + '"';
  Q := OpenSQL(SQL, True);
  if not Q.EOF then FSensCib := Q.FindField('TCI_SENS').AsString
               else Result := False;
  Ferme(Q);
end;

initialization
  RegisterClasses([TOM_TRGUIDE]);

end.

