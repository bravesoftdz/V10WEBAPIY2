{ Unité : Source TOM : EQUILIBRAGE_TOM
--------------------------------------------------------------------------------------
    Version    |   Date   | Qui  |   Commentaires
--------------------------------------------------------------------------------------
 0.91            02/10/03    JP     Création de l'unité

--------------------------------------------------------------------------------------}
unit EQUILIBRAGE_TOM ;

interface

uses
  Controls, Classes,
  {$IFNDEF EAGLCLIENT}
  db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} FE_Main,
  {$ELSE}
  MaineAGL, UTob,
  {$ENDIF}
  HCtrls, UTOM, Commun, ExtCtrls;

type
  TOM_EQUILIBRAGE = Class (TOM)
    procedure OnLoadRecord               ; override ;
    procedure OnArgument ( S: String )   ; override ;
  end ;

procedure TRLanceFiche_DetailVirement(Dom, Fiche, Range, Lequel, Arguments: string);

implementation

{---------------------------------------------------------------------------------------}
procedure TRLanceFiche_DetailVirement(Dom, Fiche, Range, Lequel, Arguments: string);
{---------------------------------------------------------------------------------------}
begin
  AglLanceFiche(Dom, Fiche, Range, Lequel, Arguments);
end;

{---------------------------------------------------------------------------------------}
procedure TOM_EQUILIBRAGE.OnLoadRecord ;
{---------------------------------------------------------------------------------------}
var
  Q : TQuery;

    {------------------------------------------------------------------}
    procedure RemplirZone(Prefix : Char);
    {------------------------------------------------------------------}
    begin
      {S'il s'agit d'un compte français, on recupère le RIB ...}
      if Copy(Q.FindField('BQ_PAYS').AsString, 1, 2) = 'FR' then begin
        SetControlText(Prefix + 'ETAB'   , Q.FindField('BQ_ETABBQ').AsString);
        SetControlText(Prefix + 'GUICHET', Q.FindField('BQ_GUICHET').AsString);
        SetControlText(Prefix + 'COMPTE' , Q.FindField('BQ_NUMEROCOMPTE').AsString);
        SetControlText(Prefix + 'CLE'    , Q.FindField('BQ_CLERIB').AsString);
      end
      {... Sinon, on rend les zones invisibles}
      else begin
        SetControlVisible(Prefix + 'ETAB'   , False);
        SetControlVisible(Prefix + 'GUICHET', False);
        SetControlVisible(Prefix + 'COMPTE' , False);
        SetControlVisible(Prefix + 'CLE'    , False);
        SetControlVisible(Prefix + 'LBRIB'  , False);
      end;
      SetControlText(Prefix + 'IBAN'   , Q.FindField('BQ_CODEIBAN').AsString);
      SetControlText(Prefix + 'BIC'    , Q.FindField('BQ_CODEBIC').AsString);
      SetControlText(Prefix + 'LIBELLE', Q.FindField('BQ_LIBELLE').AsString);
      SetControlText(Prefix + 'DEVISE' , Q.FindField('BQ_DEVISE').AsString);
      SetControlText(Prefix + 'DOMIC'  , Q.FindField('BQ_DOMICILIATION').AsString);
    end;

begin
  inherited;
  {On récupère les informations bancaires de la source et de la destination du virement}
  Q := OpenSQL('SELECT * FROM BANQUECP WHERE BQ_CODE IN ("' + GetField('TEQ_SGENERAL') +
               '", "' + GetField('TEQ_DGENERAL') + '")', True);
  try
    {Remplissage de la form}
    while not Q.EOF do begin
      if Q.FindField('BQ_CODE').AsString = GetField('TEQ_SGENERAL') then RemplirZone('S')
                                                                    else RemplirZone('D');
      Q.Next;
    end;
  finally
    Ferme(Q);
  end;
  {Maj des drapeaux}
  AssignDrapeau(TImage(GetControl('IDEV' )), GetControlText('TEQ_DEVISE'));
  AssignDrapeau(TImage(GetControl('IDEV1')), GetControlText('SDEVISE'));
  AssignDrapeau(TImage(GetControl('IDEV2')), GetControlText('DDEVISE'));
end;

{---------------------------------------------------------------------------------------}
procedure TOM_EQUILIBRAGE.OnArgument ( S: String ) ;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  Ecran.HelpContext := 150;
end;

initialization
  RegisterClasses([TOM_EQUILIBRAGE]);

end.
