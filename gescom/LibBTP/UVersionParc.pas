unit UVersionParc;

interface
Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
     uTob,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     ParamSoc
          ;

type
	TVersionParc = class (Tobject)
  	private
    	fAction : TactionFiche;
      function RecupParamSoc : integer;
    public
      constructor create;
      destructor destroy; override;
      property ModeGestion : TactionFiche read faction write faction ;
      function Nouveau : TOB;
      function Lire (ID : integer; TOBVersion : TOB) : boolean; overload;
      function Lire (ID : integer) : TOB; overload;
      function Existe (CodeArticle,Version : string) : boolean;
      function Ecrire (TOBVersion : TOB) : boolean;
      function Supprime (ID : integer) : boolean;
  end;

implementation

{ TVersionParc }

constructor TVersionParc.create;
begin
  fAction := TaModif; // par défaut on considère etre en modification
end;


destructor TVersionParc.destroy;
begin
  inherited;
end;

function TVersionParc.Ecrire(TOBVersion: TOB): boolean;
var NumUnique : integer;
begin
	result := false;
  if TOBVersion.NomTable <> 'BVERSIONSPARC' then
  begin
  	PGIError ('Impossible : les données ne sont pas pour cette table');
    exit;
  end;
  if fAction = TaCreat then
  begin
  	// on va d'abord récupérer un identifiant unique
    NumUnique := RecupParamSoc;
    if NumUnique = -1 then
    begin
    	PgiError('Erreur durant la récupération de l''identifiant VERSION');
      exit;
    end;
    TOBVersion.putvalue('BVP_ID',NumUnique);
    if (TOBVersion.getvalue('BVP_CODEARTICLE')='') or (TOBVersion.getvalue('BVP_ARTICLE')='') then
    begin
    	PgiError('Renseignez l''article de parc');
      exit;
    end;
    if (TOBVersion.getvalue('BVP_CODEVERSION')='') then
    begin
    	PgiError('Renseignez le code version');
      exit;
    end;
    if (TOBVersion.getvalue('BVP_TYPEVERSION')='') then
    begin
    	PgiError('Renseignez le type de version');
      exit;
    end;

    TOBVersion.setallModifie(true);
    result := TOBVersion.InsertDB (nil,false);
  end else
  begin
  	result := true;
  	if TOBversion.IsOneModifie(true) then result := TOBVersion.UpdateDB(false); // on protège l'ecriture
  end;
end;

function TVersionParc.Lire(ID: integer; TOBVersion: TOB): boolean;
var SQl : string;
		QQ : TQuery;
begin
	result := true;
	SQL := 'SELECT * FROM BVERSIONSPARC WHERE BVP_ID='+InttoStr(ID);
  QQ := OpenSql(SQl,true,-1,'',true);
  if not QQ.eof then TOBVersion.SelectDB('',QQ) else result := false;
  ferme (QQ);
end;

function TVersionParc.Existe(CodeArticle, Version: string): boolean;
var SQl : string;
		QQ : TQuery;
begin
 	result := false;
	SQL := 'SELECT BVP_ID FROM BVERSIONSPARC WHERE BVP_CODEARTICLE="'+CodeArticle+'" AND BVP_CODEVERSION="'+Version+'"';
  QQ := OpenSql(SQl,true,-1,'',true);
  if not QQ.eof then
  begin
  	result:= True;
  end;
  ferme (QQ);

end;

function TVersionParc.Lire(ID: integer): TOB;
var SQl : string;
		QQ : TQuery;
begin
  result := nil;
	SQL := 'SELECT * FROM BVERSIONSPARC WHERE BVP_ID='+InttoStr(ID);
  QQ := OpenSql(SQl,true,-1,'',true);
  if not QQ.eof then
  begin
  	result:= Nouveau;
  	result.SelectDB('',QQ);
  end;
  ferme (QQ);
end;

function TVersionParc.Nouveau: TOB;
begin
   result := TOB.Create ('BVERSIONSPARC',nil,-1);
   result.putValue('BVP_ACTIVE','X');
end;

function TVersionParc.RecupParamSoc: integer;
var Ok : boolean;
		nbPassage : integer;
begin
	result := -1;
  Ok := false;
  nbPassage := 1;
  repeat
  	if nbpassage > 100 then break;
		if ExecuteSql('UPDATE PARAMSOC SET SOC_DATA="X" WHERE SOC_NOM="SO_BTVERROUCPTREF" AND SOC_DATA="-"')=1 then
    	Ok := true;
    if not Ok then begin sleep(300); inc(nbpassage); end;
  until OK;
  if OK then
  begin
  	result := GetparamSoc ('SO_CPTVERSIONREF');
    inc(result);
    SetParamSoc ('SO_CPTVERSIONREF',result);
		if ExecuteSql('UPDATE PARAMSOC SET SOC_DATA="-" WHERE SOC_NOM="SO_BTVERROUCPTREF"')=1 then
  end;
end;

function TVersionParc.Supprime(ID: integer): boolean;
begin
	result := true;
  if executeSql('DELETE FROM BVERSIONSPARC WHERE BVD_ID='+InttoStr(ID))=0 then
  begin
  	PGIBox('Erreur lors de la suppression');
  	result := false;
  end;
end;

end.
