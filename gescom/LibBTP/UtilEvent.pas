unit UtilEvent;

interface

Uses StdCtrls,Controls,Classes,
{$IFNDEF EAGLCLIENT}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     MenuOLG,
{$ENDIF}
      EntGc,
     forms,sysutils,ComCtrls,HCtrls,HEnt1,HMsgBox,ParamSoc,UTob;

function ActiveEvent(const CodeEvent: string): boolean;

implementation

function ActiveEvent(const CodeEvent: string): boolean;
var Evt : TCollectionEvent;
begin
  try
    Evt := TX_LesEvenements.Find(CodeEvent);
    if assigned(Evt) then
    begin
      if not Evt.SEV_ACTIF then
      begin
        Evt.SEV_ACTIF := True;
        result := Evt.Update();
      end else
        result := true;
    end else
      result := False;
  except
    result := false;
  end;
end;

end.
