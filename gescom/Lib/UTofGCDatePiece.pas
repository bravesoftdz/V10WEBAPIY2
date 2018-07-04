{***********UNITE*************************************************
Auteur  ...... : Michel RICHAUD
Créé le ...... : 10/08/2001
Modifié le ... :   /  /
Description .. : Source TOF de la TABLE : GCDATEPIECE ()
Suite ........ : Propose de choisir la date des documents qui seront
Suite ........ : établis à partir de la génération automatique
Mots clefs ... : TOF;GCDATEPIECE
*****************************************************************}
unit UTofGCDatePiece;

interface

uses StdCtrls, Controls, Classes,
  {$IFDEF EAGLCLIENT}
  MaineAgl,
  {$ELSE}
  db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} FE_main,
  {$ENDIF}
  forms, sysutils, ComCtrls, Vierge, AglInit,
  HCtrls, HEnt1, HMsgBox, UTOF, SaisUtil, EntGC, UTOB;

function ChoixDatePiece(Caption: string = ''): TdateTime;

type
  TOF_GCDATEPIECE = class(TOF)
    procedure OnNew; override;
    procedure OnDelete; override;
    procedure OnUpdate; override;
    procedure OnLoad; override;
    procedure OnArgument(S: string); override;
    procedure OnClose; override;
  end;

var DatePiece: TDateTime;

const
  // libellés des messages
  TexteMessage: array[1..9] of string = (
    {1}'1;?caption?;La date que vous avez renseignée n''est pas valide;E;O;O;O;'
    {2}, '2;?caption?;La date que vous avez renseignée n''est pas dans un exercice ouvert;E;O;O;O;'
    {3}, '3;?caption?;La date que vous avez renseignée est antérieure à une clôture;E;O;O;O;'
    {4}, '4;?caption?;La date que vous avez renseignée est antérieure à une clôture;E;O;O;O;'
    {5}, '5;?caption?;La date que vous avez renseignée est en dehors des limites autorisées;E;O;O;O;'
    {6}, '6;?caption?;Vous ne pouvez pas saisir avant le ;E;O;O;O;'
    {7}, '7;?caption?;La date est antérieure à celle de dernière clôture de stock;E;O;O;O;'
    {8}, ''
    {9}, ''
    );

implementation

uses paramsoc;

function ChoixDatePiece(Caption: string = ''): TdateTime;
var Retour: string;
  LATOB: TOB;
begin
  Retour := '';
  LATOB := TOB.Create('LES PARAMETRES', nil, -1);
  try
    LaTOB.AddChampSupValeur('CAPTION', Caption, false);
    TheTob := LATOB;
    if not VH_GC.GCIfDefCEGID then
      Retour := AglLanceFiche('GC', 'GCDATEPIECE', '', '', '');
    if Retour = 'VALIDE' then Result := DatePiece else Result := V_PGI.DateEntree;
  finally
    LATOB.free;
    TheTob := nil;
  end;
end;
{==============================================================================================}
{================================= Evenement de la TOF ========================================}
{==============================================================================================}

procedure TOF_GCDATEPIECE.OnNew;
begin
  inherited;
end;

procedure TOF_GCDATEPIECE.OnDelete;
begin
  inherited;
end;

procedure TOF_GCDATEPIECE.OnUpdate;
var Err: integer;
  DD: TDateTime;
begin
  inherited;
  Err := ControleDate(GetControlText('DATEPIECE'));
  if (Err > 0) and not (GetParamSoc('SO_GCDESACTIVECOMPTA')) then // DC Conditionner msg suivant paramsoc
  begin
    HShowMessage(TexteMessage[Err], Ecran.Caption, '');
    TFVierge(Ecran).ModalResult := 0;
    Exit;
  end;
  DD := StrToDate(GetControlText('DATEPIECE'));
  if DD < V_PGI.DateDebutEuro then
  begin
    HShowMessage(TexteMessage[6] + ' ' + DateToStr(V_PGI.DateDebutEuro), Ecran.Caption, '');
    TFVierge(Ecran).ModalResult := 0;
    Exit;
  end;
  if ((DD <= VH_GC.GCDateClotureStock) and (VH_GC.GCDateClotureStock > 100)) then
  begin
    HShowMessage(TexteMessage[7], Ecran.Caption, '');
    TFVierge(Ecran).ModalResult := 0;
    Exit;
  end;
  DatePiece := StrToDate(GetControlText('DATEPIECE'));
  TFVierge(Ecran).Retour := 'VALIDE';
end;

procedure TOF_GCDATEPIECE.OnLoad;
begin
  inherited;
end;

procedure TOF_GCDATEPIECE.OnArgument(S: string);
begin
  inherited;
  S := 'toto';
  if LaTOB.getValue('CAPTION') <> '' then
  begin
    ecran.caption := LATOB.GetValue('CAPTION');
    UpdateCaption(ecran);
  end;
end;

procedure TOF_GCDATEPIECE.OnClose;
begin
  inherited;
end;

initialization
  registerclasses([TOF_GCDATEPIECE]);
end.
