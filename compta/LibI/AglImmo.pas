unit AglImmo;

interface
uses SysUtils,AGLInit,M3fp,HmsgBox, StdCtrls, Controls, Classes, db, forms,
{$IFDEF SERIE1}
    {$IFNDEF EAGLCLIENT}
        {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
    {$ENDIF}
{$ELSE}
    {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
{$ENDIF}
ComCtrls, HCtrls, HEnt1, UTOF, DBGrids, HDB, Hqry, Hstatus
{$IFDEF AMORTISSEMENT}
      , IMMO_TOM
{$ENDIF}
     ;

{$IFDEF SERIE1}
procedure initialisationAglImmo ;
{$ENDIF}

implementation
{$IFDEF AMORTISSEMENT}
procedure AGLZoomFicheImmo(parms: array of Variant; nb: integer);
Var CodeImmo: string ;
begin
CodeImmo:=String(parms[0]) ;
AMLanceFiche_FicheImmobilisation(CodeImmo,taConsult,'') ;
End;
{$ENDIF}

function AGLConcatFliste( parms: array of variant; nb: integer ) : variant ;
Var F:TForm; C:THGrid ;i: integer ;Q: THQuery ; AvecInitMove: boolean ; LeChamp,LeToken,LeGuil: string ;
begin
result:='' ;
F:=TForm(LongInt(parms[0]));
C:=THGrid(F.FindComponent('Fliste'));
Q:=THQuery(F.FindComponent('Q'));
LeChamp:=UpperCase(String(Parms[1])) ;
LeToken:=String(Parms[2]) ;
AvecInitMove:=(UpperCase(String(Parms[3]))='TRUE') ;
if  (UpperCase(String(Parms[4]))='TRUE') then LeGuil:='"' else LeGuil:='' ;
if C<>nil then
  begin
  if AvecInitMove then InitMove(100,'');
  if C.AllSelected then
    begin
    Q.First;
    while not Q.EOF do
      begin
      if AvecInitMove then MoveCur(False);
      if Result<>'' then Result:=Result+String(Parms[2]) ;
      Result:=Result+LeGuil+Q.FindField(String(Parms[1])).AsString+LeGuil ;
      Q.Next;
      end;
    FiniMove;
    end
  else
    begin
    if AvecInitMove then InitMove(C.nbSelected,'');
    for i:=0 to C.nbSelected-1 do
      begin
      if AvecInitMove then MoveCur(False);
      C.GotoLeBookmark(i);
      if Result<>'' then Result:=Result+String(Parms[2]) ;
      Result:=Result+LeGuil+Q.FindField(String(Parms[1])).AsString+LeGuil ;
      end;
    end;
  if AvecInitMove then FiniMove;
  end ;
End;

{$IFDEF SERIE1}
procedure initialisationAglImmo ;
begin
{$IFDEF AMORTISSEMENT} //YCP 25/08/05
RegisterAglProc('ZoomFicheImmo', FALSE, 1, AglZoomFicheImmo );
{$ENDIF}
RegisterAglFunc('ConcatFliste', TRUE, 2, AGLConcatFliste );
end;
{$ELSE}
Initialization
{$IFDEF AMORTISSEMENT}
RegisterAglProc('ZoomFicheImmo', FALSE, 1, AglZoomFicheImmo );
{$ENDIF}
RegisterAglFunc('ConcatFliste', TRUE, 2, AGLConcatFliste );
{$ENDIF}

end.
