unit TomCatTransac ;
{-------------------------------------------------------------------------------------
    Version    |   Date   | Qui  |   Commentaires
--------------------------------------------------------------------------------------
 1.0.0.100.002  04/02/04    JP     Ajout de l'évènement OnUpdateRecord, car le sens des
                                   catégories n'etait pas renseigné.

--------------------------------------------------------------------------------------}
interface

uses
  {$IFDEF EAGLCLIENT}
  MaineAGL, UTob,
  {$ELSE}
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} db,  FE_Main, 
  {$ENDIF}
  Controls, Classes, Forms, HCtrls, UTOM ;

type
  TOM_CATTRANSAC = class (TOM)
    procedure OnArgument(S : string); override;
    procedure OnUpdateRecord        ; override;
  protected
  end ;

procedure TRLanceFiche_CatTransac(Dom: string; Fiche: string; Range: string; Lequel: string; Arguments: string);

implementation

uses
  HMsgBox;

{---------------------------------------------------------------------------------------}
procedure TRLanceFiche_CatTransac(Dom: string; Fiche: string; Range: string; Lequel: string; Arguments: string);
{---------------------------------------------------------------------------------------}
begin
  AglLanceFiche(Dom, Fiche, Range, Lequel, Arguments);
end;

{---------------------------------------------------------------------------------------}
procedure TOM_CATTRANSAC.OnArgument ( S: String ) ;
{---------------------------------------------------------------------------------------}
begin
  inherited ;
  Ecran.HelpContext := 150;
end ;

{---------------------------------------------------------------------------------------}
procedure TOM_CATTRANSAC.OnUpdateRecord;
{---------------------------------------------------------------------------------------}
var
  sens,
  SQL : string ;
  Q   : TQuery ;
begin
  inherited;
  if GetField('TCA_TYPEFLUX') = '' then begin
    HShowMessage('0;' + Ecran.Caption + ';Veuillez renseigner le type de flux.;W;O;O;O', '', '');
    LastError := 1;
    SetFocusControl('TCA_TYPEFLUX');
    Exit;
  end;

  if GetField('TCA_CATTRANSAC') = '' then begin
    HShowMessage('1;' + Ecran.Caption + ';Veuillez renseigner le code de la catégorie.;W;O;O;O', '', '');
    LastError := 1;
    SetFocusControl('TCA_CATTRANSAC');
    Exit;
  end;

  if GetField('TCA_TYPETRANSAC') = '' then begin
    HShowMessage('2;' + Ecran.Caption + ';Veuillez renseigner le type de transaction.;W;O;O;O', '', '');
    LastError := 1;
    SetFocusControl('TCA_TYPETRANSAC');
    Exit;
  end;

  try
    SQL := 'SELECT TTL_SENS FROM TYPEFLUX WHERE TTL_TYPEFLUX="' + GetField('TCA_TYPEFLUX') + '" ' ;
    Q := OpenSQL(SQL,True);
    if not Q.EOF then
      Sens := Q.FindField('TTL_SENS').AsString;
    SetField('TCA_SENS',Sens);
  finally
    Ferme(Q);
  end;
end;

Initialization
  registerclasses ( [ TOM_CATTRANSAC ] ) ;
end.

