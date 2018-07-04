unit BTPGetStructure;

interface

uses hctrls,
  sysutils, Windows,
{$IFNDEF EAGLCLIENT}
  uDbxDataSet, DB,
{$ELSE}
  UtilchampsSup,
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} DB,
{$ENDIF}
  uhttp,
  utob;

{$IFDEF EAGLCLIENT}
function BTPGetFieldStructure(const sql: hstring): tob;
{$ELSE}
procedure BTPGetFieldStructure(const sql: hstring; TOBOUT : TOB);
{$ENDIF}

implementation

{$IFDEF EAGLCLIENT}
function BTPGetFieldStructure(const sql: hstring): tob;
var
  It, Ot: tob ;
begin
  result := Nil ;
  It := tob.create ('Données', Nil, -1) ;
  try
    It.AddChampSupValeur ('Sql', Sql) ;
    oT := AppServer.Request('PluginBTPS.GetStructure', 'INFO', iT, '', '');
    if assigned(Ot) and (Ot.detail.count >0 ) then
    begin
      result := TOB.Create('LES CHAMPS',nil,-1);
      result.dupliquer(Ot.detail[0],false,true);
    	Ot.Free ;
    end else raise Exception.Create ('La structure rendue par le serveur n''est pas exploitable.') ;
  finally
    It.Free ;
  end ;
end;
{$ELSE}
procedure BTPGetFieldStructure(const sql: hstring; TOBOUT : TOB);
var
  Q: TQuery;
  i: integer;
  NomChamps : string;
  TOBDO : TOB;
begin
  Q := OpenSql(Sql, True);
  try
  	TOBDO := TOB.Create ('LES CHAMPS',TOBOUT,-1);
    // Boucle sur tous les champs
    for i := 0 to Q.FieldCount - 1 do
    begin
      NomChamps := Q.Fields[I].FieldName;
      TOBDO.AddChampSupValeur (NomChamps,'');
    end;
		//
  finally
    Ferme(Q);
  end;
end;
{$ENDIF}

end.

