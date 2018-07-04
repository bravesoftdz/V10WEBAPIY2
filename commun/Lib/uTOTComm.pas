{***********UNITE*************************************************
Auteur  ...... : Nicolas Chavanne
Créé le ...... : 02/05/2007
Modifié le ... :   /  /
Description .. : TOT Commune à tous les métiers
Mots clefs ... : TOT;UTOTCOMM

TOT            : TOT AGL
 |
 +- tTOTComm   : TOT Commune à tous les métiers
     |           -
     |           -
     |
     +- twTOT  : TOT Métier W (Gestion de Production)
*****************************************************************}
Unit uTOTComm;

Interface

Uses
  StdCtrls, Controls, Classes,
  Forms, Sysutils, ComCtrls,
  {$IFNDEF EAGLCLIENT}
    db, dbtables, Tablette,
  {$ELSE}
    eTablette,
  {$ENDIF}
  HCtrls, HEnt1, HMsgBox, uTob, uTot
  ;

Type
  tTOTComm = Class(TOT)
    procedure OnNewRecord          ; override;
    procedure OnDeleteRecord       ; override;
    procedure OnUpdateRecord       ; override;
    procedure OnAfterUpdateRecord  ; override;
    procedure OnClose              ; override;
    procedure OnArgument(S: String); override;
  private
    FStArgument: String;
    FLequel: String;
    FTabletteName: String;
    FPrefixe: String;

    procedure GotoLine(const Code: String);
  public
    {$IFNDEF EAGLCLIENT}
      function GetDataSet: TDataSet;
    {$ELSE  !EAGLCLIENT}
      function GetDataSet: Tob;
    {$ENDIF !EAGLCLIENT}
    procedure AvertirTablette;

    property StArgument: String read FStArgument Write FStArgument;
    property Lequel: String read FLequel;
    property TabletteName: String read FTabletteName;
    property Prefixe: String read FPrefixe;
  end;

Implementation

uses
  UtilPgi
  ;

procedure tTOTComm.OnNewRecord;
begin
  Inherited;
end;

procedure tTOTComm.OnDeleteRecord;
begin
  Inherited;

  AvertirTablette()  
end;

procedure tTOTComm.OnUpdateRecord;
begin
  Inherited;
end;

procedure tTOTComm.OnAfterUpdateRecord;
begin
  Inherited;

  AvertirTablette()
end;

procedure tTOTComm.OnClose;
begin
  Inherited;
end;

procedure tTOTComm.OnArgument(S: String);
begin
  Inherited;

  FStArgument := S;
  FLequel := GetArgumentString(FStArgument, 'LEQUEL');
  if Assigned(Ecran) then
  begin
    FTabletteName := TFTablette(Ecran).FQuoi;
    FPrefixe      := TFTablette(Ecran).FPrefixe;
  end;
  { Positionnement de la grille }
  if FLequel <> '' then
    GotoLine(FLequel)
end;

procedure tTOTComm.GotoLine(const Code: String);
var
  DS: TDataSet;
begin
  if FPrefixe <> '' then
  begin
    DS := GetDataSet;
    if Assigned(DS) then
    begin
      {$IFDEF EAGLCLIENT}
        DS.First;
        while (not DS.eof) and (Code <> DS.FindField(FPrefixe + '_CODE').AsString) do
          DS.Next;
        if (DS.CurrentFilleIndex + 1) < TFTablette(Ecran).FListe.RowCount then
          TFTablette(Ecran).FListe.Row := DS.CurrentFilleIndex + 1
        else
          TFTablette(Ecran).FListe.Row := TFTablette(Ecran).FListe.RowCount-1;
      {$ELSE}
        GetDataSet.Locate(FPrefixe + '_CODE', Code, [LoCaseInsensitive]);
      {$ENDIF}
    end
  end
end;

{$IFNDEF EAGLCLIENT}
function tTOTComm.GetDataSet: TDataSet;
{$ELSE  !EAGLCLIENT}
function tTOTComm.GetDataSet: Tob;
{$ENDIF !EAGLCLIENT}
begin
  if Assigned(Ecran) then
  begin
    {$IFNDEF EAGLCLIENT}
      Result := TFTablette(Ecran).FListe.DataSource.DataSet
    {$ELSE  !EAGLCLIENT}
      Result := Tob(TFTablette(Ecran).FListe.Objects[0, TFTablette(Ecran).FListe.FixedRows]);
      if Assigned(Result) then
        Result := Result.Parent
    {$ENDIF !EAGLCLIENT}
  end
  else
    Result := nil
end;

procedure tTOTComm.AvertirTablette;
var
  iTT: Integer;
  TTName: String;
begin
  if TabletteName <> '' then
  begin
    TTName := TabletteName;
    iTT := TTToNum(TTName);
    if V_Pgi.DECombos[iTT].MemLoad = ltMem then
      AvertirTable(TabletteName);
  end
end;

Initialization
  registerclasses([tTOTComm]);
end.
