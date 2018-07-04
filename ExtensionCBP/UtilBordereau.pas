unit UtilBordereau;

interface
uses hctrls,
  sysutils,HEnt1,
{$IFNDEF EAGLCLIENT}
  uDbxDataSet, DB,
{$ELSE}
  uWaini,
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} DB,
{$ENDIF}
  uhttp,paramsoc,
  utob,uEntCommun;

function ConstitueRequeteBordereau (fNaturepiece,fAffaire,fTiers,fNatureAuxi: string;fDatePiece : TdateTime; var frequeteSelect : string) : string;

implementation

function ConstitueRequeteBordereau (fNaturepiece,fAffaire,fTiers,fNatureAuxi: string;fDatePiece : TdateTime; var frequeteSelect : string) : string;
var ReqCli : string;
begin
  if fNaturePiece = 'FRC' then exit;
  frequeteSelect := '';
	if (ftiers = '') and (faffaire <> '') then exit;
  if fAffaire <> '' then
  begin
  	if ftiers <> '' then frequeteSelect := '((';
  	frequeteSelect := frequeteSelect + 'BDE_AFFAIRE="'+fAffaire+'"';
  end;
  if ftiers <> '' then
  begin
  	reqCli := 'BDE_CLIENT="'+fTiers+'"' ;

    if fNatureAuxi <> '' then
    begin
      reqCli := reqCli + ' AND BDE_NATUREAUXI="'+fnatureAuxi+'"';
    end;

    if faffaire <> '' then
    begin
      frequeteSelect := frequeteSelect+ ' AND '+reqCli+') OR ('+reqCli+' AND BDE_AFFAIRE=""))';
    end else
    begin
    	frequeteSelect := frequeteSelect+ ReqCli + ' AND BDE_AFFAIRE=""';
    end;
  end;
  if fDatePiece <> iDate1900 then
  begin
  	frequeteSelect := frequeteSelect + ' AND BDE_DATEDEPART <= "'+USDATETIME(fDatePiece)+'"' +
    			 														 ' AND BDE_DATEFIN >= "'+USDATETIME(fDatePiece)+'"';
  end;
  result := 'SELECT * FROM BDETETUDE WHERE '+fRequeteSelect;
end;


end.
