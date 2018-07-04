unit UtilBTPgestChantier;

interface
uses hctrls,
  sysutils,HEnt1,
{$IFNDEF EAGLCLIENT}
  uDbxDataSet, DB,
{$ELSE}
  uWaini,
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} DB,
{$ENDIF}
  uhttp,
  utob;

function EncodeLienDevCHA (TOBPL : TOB ) : string;

implementation

function EncodeLienDevCHA (TOBPL : TOB ) : string;
Var DD  : TDateTime ;
    StD : String ;
begin
  if TOBPL.NomTable = 'LIGNE' then
     begin
       DD:=TOBPL.GetValue('GL_DATEPIECE') ; StD:=FormatDateTime('ddmmyyyy',DD) ;
       Result:=StD+';'+TOBPL.GetValue('GL_NATUREPIECEG')+ ';'+TOBPL.GetValue('GL_SOUCHE')+';'
               +IntToStr(TOBPL.GetValue('GL_NUMERO'))+';'+IntToStr(TOBPL.GetValue('GL_INDICEG'))+';';
     end else if TOBPL.Nomtable = 'PIECE' then
     begin
       DD:=TOBPL.GetValue('GP_DATEPIECE') ; StD:=FormatDateTime('ddmmyyyy',DD) ;
       Result:=StD+';'+TOBPL.GetValue('GP_NATUREPIECEG')+ ';'+TOBPL.GetValue('GP_SOUCHE')+';'
               +IntToStr(TOBPL.GetValue('GP_NUMERO'))+';'+IntToStr(TOBPL.GetValue('GP_INDICEG'))+';';
     end;
end;

end.
