{***********UNITE*************************************************
Auteur  ...... : JEAN PASTERIS
Créé le ...... : 30/06/2004
Modifié le ... :   /  /
Description .. : TOM de la fiche TRCONTRAT
Mots clefs ... : TRLANCEFICHE_CONTRAT;
*****************************************************************}
{-------------------------------------------------------------------------------------
    Version    |   Date   | Qui | Commentaires
--------------------------------------------------------------------------------------
 6.xx.xxx.xxx    30/06/04    JP   Création de l'unité
--------------------------------------------------------------------------------------}
unit TRCONTRAT_TOM;

interface

uses
  {$IFDEF EAGLCLIENT}
  MaineAGL, UTob ,
  {$ELSE}
   db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} FE_Main,
  {$ENDIF}
  Controls, Classes, SysUtils, HCtrls, HEnt1, UTOM;

type
  TOM_TRCONTRAT = class(TOM)
    procedure OnNewRecord              ; override;
    procedure OnUpdateRecord           ; override;
    procedure OnAfterUpdateRecord      ; override;
    procedure OnArgument   (S : string); override;
    procedure OnDeleteRecord           ; override;
    procedure OnChangeField(F : TField); override;
  private
    OldContrat  : string;
    Duplication : Boolean;

    procedure DupliquerContrat;
  end ;

procedure TRLanceFiche_Contrat(Dom, Fiche, Range, Lequel, Arguments: string);

implementation

uses
  HMsgBox;

{---------------------------------------------------------------------------------------}
procedure TRLanceFiche_Contrat(Dom, Fiche, Range, Lequel, Arguments: string);
{---------------------------------------------------------------------------------------}
begin
  AglLanceFiche(Dom, Fiche, Range, Lequel, Arguments);
end;

{---------------------------------------------------------------------------------------}
procedure TOM_TRCONTRAT.OnArgument(S : string);
{---------------------------------------------------------------------------------------}
begin
  inherited;
  Ecran.HelpContext := 150;
  {Suppression de l'action}
  ReadTokenSt(S);
  {On regarde si on est en duplication}
  Duplication := ReadTokenSt(S) = 'DUPLI';
  {Récupération du contrat "modèle"}
  OldContrat  := ReadTokenSt(S);
  SetControlVisible('BDELETE', False);
end;

{---------------------------------------------------------------------------------------}
procedure TOM_TRCONTRAT.OnNewRecord;
{---------------------------------------------------------------------------------------}
var
  Q : TQuery;
begin
  inherited;
  if Duplication then begin
    Q := OpenSQL('SELECT * FROM TRCONTRAT WHERE TRC_CODECONTRAT = "' + OldContrat + '"', True);
    try
      if not Q.EOF then begin
        SetField('TRC_AGENCE',  Q.FindField('TRC_AGENCE' ).AsString);
        SetField('TRC_LIBELLE', Q.FindField('TRC_LIBELLE').AsString);
        SetField('TRC_ABREGE',  Q.FindField('TRC_ABREGE' ).AsString);
        SetField('TRC_COMMVT',  Q.FindField('TRC_COMMVT' ).AsFloat );
        SetField('TRC_DEBCONTRAT', Q.FindField('TRC_FINCONTRAT').AsDateTime + 1);
        SetField('TRC_FINCONTRAT', Q.FindField('TRC_FINCONTRAT').AsDateTime + 365);
      end;
    finally
      Ferme(Q);
    end;
  end
  else begin
    SetField('TRC_DEBCONTRAT', DebutAnnee(V_PGI.DateEntree));
    SetField('TRC_FINCONTRAT', FinAnnee(V_PGI.DateEntree));
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_TRCONTRAT.OnDeleteRecord;
{---------------------------------------------------------------------------------------}
begin
  inherited;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_TRCONTRAT.OnChangeField(F : TField);
{---------------------------------------------------------------------------------------}
begin
  inherited;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_TRCONTRAT.OnUpdateRecord;
{---------------------------------------------------------------------------------------}
begin
//  if DS.State = dsEdit then Exit; 19/08/08
  {On s'assure qu'il n'exite pas d'autre contrat pour cette agence et cette période
   19/08/08 : avec le test sur le code c'est quand même plus efficace que le test DS.State = dsEdit !!}
  if ExisteSQL('SELECT TRC_CODECONTRAT FROM TRCONTRAT WHERE TRC_CODECONTRAT <> "' + GetField('TRC_CODECONTRAT') +
               '" AND TRC_AGENCE = "' + GetField('TRC_AGENCE') +
               '" AND ( (TRC_DEBCONTRAT >= "' + UsDateTime(GetField('TRC_DEBCONTRAT')) + '" ' +
                        'AND TRC_DEBCONTRAT <= "' + UsDateTime(GetField('TRC_FINCONTRAT')) +  '") ' +
                      'OR (TRC_FINCONTRAT >= "' + UsDateTime(GetField('TRC_DEBCONTRAT')) +  '" ' +
                        'AND TRC_FINCONTRAT <= "' + UsDateTime(GetField('TRC_FINCONTRAT')) +  '") ' +
                      'OR (TRC_DEBCONTRAT <= "' + UsDateTime(GetField('TRC_DEBCONTRAT')) +  '" ' +
                        'AND TRC_FINCONTRAT >= "' + UsDateTime(GetField('TRC_DEBCONTRAT')) +  '") ' +
                      'OR (TRC_FINCONTRAT >= "' + UsDateTime(GetField('TRC_FINCONTRAT')) +  '" ' +
                        'AND TRC_DEBCONTRAT <= "' + UsDateTime(GetField('TRC_FINCONTRAT')) + '"))') then begin
    LastError := 1;
    LastErrorMsg := TraduireMemoire('Il existe déjà un contrat pour cette période et cette agence.');
  end;
  inherited;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_TRCONTRAT.OnAfterUpdateRecord;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  if Duplication then DupliquerContrat;
  Duplication := False;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_TRCONTRAT.DupliquerContrat;
{---------------------------------------------------------------------------------------}
var
  SQL : string;
begin
  BeginTrans;
  try
    SQL := SQL + 'INSERT INTO FRAIS (TFR_ABREGE, TFR_APARTIRDE, TFR_ARRONDIJOUR, TFR_CODEFLUX, ';
    SQL := SQL + 'TFR_CODEFRAIS, TFR_COMMISSION, TFR_CONTRAT, TFR_COUTFORFAIT, TFR_COUTMAX, ';
    SQL := SQL + 'TFR_COUTMIN, TFR_COUTUNITAIRE, TFR_DEVISE, TFR_FLUXASSOCIE, TFR_LIBELLE, ';
    SQL := SQL + 'TFR_MODEFRAIS, TFR_PLUSJOUR, TFR_TRANCHEFRAIS, TFR_TYPEFRAIS) ';
    SQL := SQL + 'SELECT TFR_ABREGE, TFR_APARTIRDE, TFR_ARRONDIJOUR, TFR_CODEFLUX, TFR_CODEFRAIS, ';
    SQL := SQL + 'TFR_COMMISSION, "' + GetField('TRC_CODECONTRAT') + '", TFR_COUTFORFAIT, TFR_COUTMAX, TFR_COUTMIN, TFR_COUTUNITAIRE, ';
    SQL := SQL + 'TFR_DEVISE, TFR_FLUXASSOCIE, TFR_LIBELLE, TFR_MODEFRAIS, TFR_PLUSJOUR, TFR_TRANCHEFRAIS, ';
    SQL := SQL + 'TFR_TYPEFRAIS FROM FRAIS WHERE TFR_CONTRAT = "' + OldContrat + '"';
    ExecuteSQL(SQL);
    CommitTrans;
  except
    on E : Exception do begin
      RollBack;
      PGIError(E.Message);
    end;
  end;
end;

initialization
  RegisterClasses([TOM_TRCONTRAT]);

end.

