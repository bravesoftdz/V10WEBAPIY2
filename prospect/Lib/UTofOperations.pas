unit UTofOperations;

interface
uses  Classes,forms,
{$ifdef AFFAIRE}
UTOFAFTRADUCCHAMPLIBRE,
{$endif}
      UTOF,UtilSelection,ParamSoc,hCtrls,HEnt1,HQry,HStatus,M3FP,
{$IFDEF EAGLCLIENT}
      eMul,MaineAGL;
{$ELSE}
      Mul,Fe_Main,hDb;
{$ENDIF}
Type
{$ifdef AFFAIRE}
                //PL le 18/05/07 pour gérer les champs libres si paramétrés
     TOF_Operations = Class (TOF_AFTRADUCCHAMPLIBRE)
 {$else}
     TOF_Operations = Class (TOF)
{$endif}
     public
        procedure OnArgument(Arguments : String ) ; override ;
        function RameneListeCodesROP: string;
     private
        StFiltre : string ;
        procedure AfterShow;
     END ;

Function RTLanceFiche_Operations(Nat,Cod : String ; Range,Lequel,Argument : string) : string;
function AGLRameneListeCodesROP(parms: array of variant; nb: integer): variant;
implementation

uses uTOFComm;


Function RTLanceFiche_Operations(Nat,Cod : String ; Range,Lequel,Argument : string) : string;
begin
AGLLanceFiche(Nat,Cod,Range,Lequel,Argument);
end;

procedure TOF_Operations.OnArgument(Arguments : String ) ;
var F : TForm;
    Critere,ChampMul,ValMul,stArg : String;
    x : integer;
begin
fMulDeTraitement := true;
inherited ;
fTableName := 'OPERATIONS';
stArg:=Arguments;
Critere:=(ReadTokenSt(stArg));
F := TForm (Ecran);
if Critere <> 'GRF' then
  begin
  if GetParamSocSecur('SO_RTGESTINFOS002',False) = True then
      MulCreerPagesCL(F,'NOMFIC=RTOPERATIONS');
  repeat
      Critere:=(ReadTokenSt(stArg));
      if Critere <> '' then
        begin
        x:=pos('=',Critere);
        if x<>0 then
           begin
           ChampMul:=copy(Critere,1,x-1);
           ValMul:=copy(Critere,x+1,length(Critere));
           if ChampMul='FILTRE' then StFiltre := ValMul
           end;
        end;
  until Critere='';
  if (F.Name='RTOPERATIONS_TL') or (F.Name='RFOPERATIONS_TL') then
    begin
    Tfmul(Ecran).FiltreDisabled:=true;
    TFmul(Ecran).OnAfterFormShow := AfterShow;
    end;
  end;
  if (F.Name='RTOPERATION_SELEC') then SetControlText ('XX_WHERE','ROP_OPERATION<>"MODELES D''ACTIONS"');
end;

procedure TOF_Operations.AfterShow;
begin
  Tfmul(Ecran).ForceSelectFiltre(StFiltre , V_PGI.User,false,true);
end ;
function AGLRameneListeCodesROP(parms: array of variant; nb: integer): variant;
var
  F: TForm;
  TOTOF: TOF;
begin
  F := TForm(Longint(Parms[0]));
  if (F is TFmul) then
    TOTOF := TFMul(F).LaTOF
  else
    exit;
  if (TOTOF is TOF_Operations) then
    result := TOF_Operations(TOTOF).RameneListeCodesROP
  else
    exit;
end;

function TOF_Operations.RameneListeCodesROP: string;
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
      code:=Q.FindField('ROP_OPERATION').asstring ;
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
        code:=TFmul(Ecran).Q.FindField('ROP_OPERATION').asstring ;
        if Result = '' then Result:=code else Result:=Result+';'+code;
      end;
      L.ClearSelected;
      end;
  FiniMove;
end;

Initialization
registerclasses([TOF_Operations]);
RegisterAglFunc('RameneListeCodesROP', TRUE, 0, AGLRameneListeCodesROP);
end.
