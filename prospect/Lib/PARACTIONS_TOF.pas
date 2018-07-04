{***********UNITE*************************************************
Auteur  ...... : MNG
Créé le ...... : 18/04/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PARACTIONS ()
Mots clefs ... : TOF;PARACTIONS
*****************************************************************}
Unit PARACTIONS_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,HDB,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
{$else}
     eMul,
     uTob,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF,HQry,HStatus,M3FP ;

Type
  TOF_PARACTIONS = Class (TOF)
    function RameneListeCodesRPA: string;
  end ;

function AGLRameneListeCodesRPA(parms: array of variant; nb: integer): variant;

Implementation

function AGLRameneListeCodesRPA(parms: array of variant; nb: integer): variant;
var
  F: TForm;
  TOTOF: TOF;
begin
  F := TForm(Longint(Parms[0]));
  if (F is TFmul) then
    TOTOF := TFMul(F).LaTOF
  else
    exit;
  if (TOTOF is TOF_PARACTIONS) then
    result := TOF_PARACTIONS(TOTOF).RameneListeCodesRPA
  else
    exit;
end;

function TOF_PARACTIONS.RameneListeCodesRPA: string;
var
  F: TFMul;
  Q: THQuery;
  i: integer;
{$IFDEF EAGLCLIENT}
  L: THGrid;
{$ELSE}
  L: THDBGrid;
{$ENDIF}
  code: string;
begin
  Result := '';
  F := TFMul(Ecran);
  L := F.FListe;
  Q := F.Q;

{$IFDEF EAGLCLIENT}
  if F.bSelectAll.Down then
    if not F.Fetchlestous then
    begin
      F.bSelectAllClick(nil);
      F.bSelectAll.Down := False;
      exit;
     end else
     F.Fliste.AllSelected := true;
{$ENDIF}

  if L.AllSelected then
    begin
    Q.First;
    while not Q.Eof do
    begin
      code:=Q.FindField('RPA_TYPEACTION').asstring ;
      if Result = '' then Result:=code else Result:=Result+';'+code;
      Q.Next;
    end;
    L.AllSelected := False;
    end
  else
    if F.FListe.NbSelected <> 0 then
      begin
      InitMove(L.NbSelected, '');
      for i := 0 to L.NbSelected - 1 do
      begin
        MoveCur(False);
        L.GotoLeBookmark(i);
{$IFDEF EAGLCLIENT}
        Q.TQ.Seek(L.Row - 1);
{$ENDIF}
        code:=TFmul(Ecran).Q.FindField('RPA_TYPEACTION').asstring ;
        if Result = '' then Result:=code else Result:=Result+';'+code;
      end;
      L.ClearSelected;
      end;
  FiniMove;
end;


Initialization
  registerclasses ( [ TOF_PARACTIONS ] ) ;
  RegisterAglFunc('RameneListeCodesRPA', TRUE, 0, AGLRameneListeCodesRPA);
end.

