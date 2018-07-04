unit uBTPVerrouilleDossier;

interface
uses
	Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
	ComCtrls, StdCtrls, Hctrls, ExtCtrls, UTob,MajTable,Hent1;

const BLOCKTABLE = 'BVERROUBASE';

function ISUtilisable : boolean;
{$IFNDEF EAGLCLIENT}
procedure ForceVerrouille ;
function ISVerrouille : boolean;
procedure Deverrouille;
procedure ISBaseDisponible ( TOBOut: TOB);
{$ENDIF}
                                          

implementation
uses
{$IFNDEF EAGLCLIENT}
DB, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} MajStruc, HQry,DBCtrls,
{$ENDIF}
		uhttp;

{$IFNDEF EAGLCLIENT}
procedure ForceVerrouille ;
begin
  Deverrouille;
  ISVerrouille;
end;

procedure Deverrouille;
begin
//	ExecuteSQL ('UPDATE PARAMSOC SET SOC_DATA="-" WHERE SOC_NOM="SO_BTPVERROU"');
  if TableExiste (BLOCKTABLE) then DbDeleteTable( DBSoc, v_pgi.driver, BLOCKTABLE, false);
end;

function ISVerrouille : boolean;
var PNewStructure,PNewField : TOB;
		i : integer;
begin
//	result := (ExecuteSQL ('UPDATE PARAMSOC SET SOC_DATA="X" WHERE SOC_NOM="SO_BTPVERROU" AND SOC_DATA<>"X"')<>1);
  PNewStructure := TOB.Create(BLOCKTABLE, nil, -1);
  PNewStructure.AddChampSupValeur('DT_NOMTABLE', BLOCKTABLE);
  PNewStructure.AddChampSupValeur('DT_CLE1', 'BVB_ID');
  PNewStructure.AddChampSupValeur('DT_UNIQUE1', '-');
  PNewStructure.AddChampSupValeur('DT_PREFIXE', 'BVB');
  for i := 2 to MaxIndexes do
    PNewStructure.AddChampSupValeur('DT_CLE' + IntToStr(i), '');
  PNewField := TOB.Create('FIELD', pNewStructure, -1);
  PNewField.AddChampSupValeur('DH_NOMCHAMP', 'BVB_ID');
  PNewField.AddChampSupValeur('DH_TYPECHAMP', 'VARCHAR(1)');
  PNewField.AddChampSupValeur('DH_PREFIX', 'BVB');
  PNewField.AddChampSupValeur('DH_NULLABLE', 'X');

	result :=  (not DbCreateTable(DBSOC, PnewStructure, V_PGI.Driver, False));
  freeAndNil(PNewStructure);
end;
{$ENDIF}

function ISUtilisable : boolean;
{$IFDEF EAGLCLIENT}
var OT : TOB;
{$ENDIF}
begin
//	result := ExisteSql ('SELECT 1 FROM PARAMSOC WHERE SOC_NOM="SO_BTPVERROU" AND SOC_DATA="-"');
{$IFDEF EAGLCLIENT}
	result := false;
	oT := AppServer.Request('PluginBTPS.ISBaseDisponible', 'INFO', nil, '', '');
  if assigned(Ot) and (Ot.FieldExists('DISPONIBLE') ) then
  begin
		result := true;
    freeAndNil(Ot);
  end;
{$ELSE}
	result := not TableExisteDB(BLOCKTABLE,dbSoc);
{$ENDIF}
end;

procedure ISBaseDisponible ( TOBOut: TOB);
BEGIN
	if ISUtilisable then
  begin
		TOBOUT.AddChampSupValeur('DISPONIBLE','X');
  end;
END;

end.
