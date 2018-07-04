{ Unité : Source TOM de la TABLE : TRANSAC
--------------------------------------------------------------------------------------
    Version    |  Date  | Qui |   Commentaires
--------------------------------------------------------------------------------------
    0.91        07/12/01   BT   Création de l'unité
 7.09.001.007   02/03/07   JP   FQ 10410 : ajout de contrôle sur la suppression
--------------------------------------------------------------------------------------}
unit TomTransac ;

interface

uses
  {$IFDEF EAGLCLIENT}
  MaineAGL, UTob, 
  {$ELSE}
  FE_Main, HDB, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} 
  {$ENDIF}
  Controls, Classes, UTOM;

type
  TOM_TRANSAC = class (TOM)
    procedure OnArgument(S: string); override;
    procedure OnUpdateRecord       ; override;
    procedure OnDeleteRecord       ; override; {FQ 10410}
  protected
    procedure CodeCatTransacOnKeyPress(Sender: TObject; var Key: Char);
  end ;

procedure TRLanceFiche_Transac(Dom: string; Fiche: string; Range: string; Lequel: string; Arguments: string);

Implementation

uses
  Commun, Constantes, HCtrls, HEnt1;

{---------------------------------------------------------------------------------------}
procedure TRLanceFiche_Transac(Dom: string; Fiche: string; Range: string; Lequel: string; Arguments: string);
{---------------------------------------------------------------------------------------}
begin
  AglLanceFiche(Dom, Fiche, Range, Lequel, Arguments);
end;

{---------------------------------------------------------------------------------------}
procedure TOM_TRANSAC.OnArgument ( S: String ) ;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  Ecran.HelpContext := 150;
  {$IFDEF EAGLCLIENT}
  THEdit(GetControl('TTR_TRANSAC')).OnKeyPress := CodeCatTransacOnKeyPress;
  {$ELSE}
  THDBEdit(GetControl('TTR_TRANSAC')).OnKeyPress := CodeCatTransacOnKeyPress;
  {$ENDIF}
end ;

{---------------------------------------------------------------------------------------}
procedure TOM_TRANSAC.OnUpdateRecord;
{---------------------------------------------------------------------------------------}
var
  ch : string;
begin
  inherited;
  {On prend quelques précautions avec les codes EQR et EQD}
  ch := GetControlText('TTR_TRANSAC');
  if ch = TRANSACIMPORT then begin
    LastError := 1;
    LastErrorMsg := TraduireMemoire('Ce code est réservé pour les imports de la comptabilité.');
  end
  else if ch = TRANSACEQUI then begin
    LastError := 1;
    LastErrorMsg := TraduireMemoire('Ce code est réservé pour les équilibrages.');
  end
  else if ch = TRANSACSAISIE then begin
    LastError := 1;
    LastErrorMsg := TraduireMemoire('Ce code est réservé pour la saisie d''écriture.');
  end
  else if ch = TRANSACVENTE then begin
    LastError := 1;
    LastErrorMsg := TraduireMemoire('Ce code est réservé pour les ventes d''OPCVM.');
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_TRANSAC.CodeCatTransacOnKeyPress(Sender: TObject; var Key: Char);
{---------------------------------------------------------------------------------------}
begin
  inherited;
  ValidCodeOnKeyPress(Key);
end;

{02/03/07 : FQ 10410 : on s'assure que les transactions ne sont pas utilisées 
{---------------------------------------------------------------------------------------}
procedure TOM_TRANSAC.OnDeleteRecord;
{---------------------------------------------------------------------------------------}
var
  Q : TQuery;
  C : string;
begin
  inherited;
  C := '';
  {Récupération du type de la transaction, ce qui détermine sur quelle table faire le test}
  Q := OpenSQL('SELECT TCA_TYPETRANSAC FROM CATTRANSAC WHERE TCA_CATTRANSAC = "' + GetField('TTR_CATTRANSAC') + '"', True);
  if not Q.Eof then C := Q.Fields[0].AsString;
  Ferme(Q);

  {Tables TROPCVM et TRVENTEOPCVM}
  if C = tca_OPCVM then begin
    if ExisteSQL('SELECT TOP_TRANSACTION FROM TROPCVM LEFT JOIN TRVENTEOPCVM ON TOP_NUMOPCVM = TVE_NUMTRANSAC ' +
                 'WHERE TOP_TRANSACTION = "' + GetField('TTR_TRANSAC') + '" AND ' +
                 '(TVE_VALBO = "" OR TVE_VALBO = "-" OR TVE_VALBO IS NULL)') then begin
      LastError := 1;
      LastErrorMsg := TraduireMemoire('La transaction est utilisée sur un OPCVM non totalement vendu et validé BO');
    end;
  end
  {Table COURTSTERMES}
  else if C = tca_CourtTerme then begin
    if ExisteSQL('SELECT TCT_TRANSAC FROM COURTSTERMES WHERE TCT_TRANSAC = "' + GetField('TTR_TRANSAC') + '" AND ' +
                 '(TCT_VALBO = "" OR TCT_VALBO = "-" OR TCT_VALBO IS NULL)') then begin
      LastError := 1;
      LastErrorMsg := TraduireMemoire('La transaction est utilisée sur un placement / financement court terme non validé BO');
    end;
  end;
end;

initialization
  RegisterClasses ( [ TOM_TRANSAC ] ) ;

end.

