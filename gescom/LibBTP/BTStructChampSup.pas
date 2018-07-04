unit BTStructChampSup;

interface
uses
  Windows, SysUtils, Classes,
  {$IFDEF EAGLCLIENT}
  UtileAGL, Maineagl,BTPGetStructure,
  {$ELSE}
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} DB,
  {$ENDIF}
  UTOB,HEnt1,EntGC,Hctrls,UtilchampsSup, variants;

// Structure ligne
procedure MemoriseChampsSupLigneETL (NaturePiece : string;WithLigneCompl:boolean=false;WithEntete:boolean=false);
procedure MemoriseChampsSupPARCL (requete : string);
procedure MemoriseChampsSupLigneOUV (NaturePiece : string);
procedure MemoriseChampsSupPIECETRAIT ;

procedure AddlesChampsSupPieceTrait (TOBPT : TOB);

procedure AddlesChampsSupLigneEtl (TOBL : TOB);
procedure InitStructureETL;
procedure AddlesChampsSupParcl (TOBL : TOB);
procedure InitStructurePARCL;
procedure AddlesChampsSupOuv (TOBL : TOB);
procedure InitStructureOUV;

implementation

uses factTOB,FactCommBTP,UtilTOBPiece,UCotraitance;

var TOBStructLig,TOBStructParc,TOBStructLigOuv,TOBSTRUCPTRAIT : TOB;

procedure InitStructureETL;
begin
	if TOBStructLig.detail.count > 0 then TOBSTRUCTLIG.ClearDetail;
end;

procedure InitStructurePARCL;
begin
	if TOBStructParc.detail.count > 0 then TOBStructParc.ClearDetail;
end;

procedure InitStructureOUV;
begin
	if TOBStructLigOuv.detail.count > 0 then TOBStructLigOuv.ClearDetail;
end;


procedure SMemoriseChampsSupLigneETL (QQ : TQuery);
var TOBEtalon : TOB;
		Indice : integer;
    NomChamps : string;
begin
	if TOBStructLig.detail.count > 0 then InitStructureETL;
  //
	TOBEtalon := TOB.Create ('LA LIGNE ETALON',TOBStructLig,-1);
  //
  for Indice := 0 to QQ.FieldCount -1 do
  begin
  	NomChamps := QQ.Fields[Indice].FieldName;
		if ExtractPrefixe (NomChamps) = 'GL' then continue; // pour ne prendre que les champs sup
  	TOBEtalon.AddChampSupValeur (NomChamps,getValInitChampsSup(NomChamps));
  end;
end;

procedure SMemoriseChampsSupParcL (QQ : TQuery);
var TOBEtalon : TOB;
		Indice : integer;
    NomChamps : string;
begin
	if TOBStructParc.detail.count > 0 then InitStructurePARCL;
  //
	TOBEtalon := TOB.Create ('LA LIGNE ETALON',TOBStructParc,-1);
  //
  for Indice := 0 to QQ.FieldCount -1 do
  begin
  	NomChamps := QQ.Fields[Indice].FieldName;
		if ExtractPrefixe (NomChamps) = 'BP1' then continue; // pour ne prendre que les champs sup
  	TOBEtalon.AddChampSupValeur (NomChamps,getValInitChampsSup(NomChamps));
  end;
end;

procedure SMemoriseChampsSupLigneOuv (QQ : TQuery);
var TOBEtalon : TOB;
		Indice : integer;
    NomChamps : string;
begin
	if TOBStructLigOuv.detail.count > 0 then InitStructureOUV;
  //
	TOBEtalon := TOB.Create ('LA LIGNE ETALON',TOBStructLigOuv,-1);
  //
  for Indice := 0 to QQ.FieldCount -1 do
  begin
  	NomChamps := QQ.Fields[Indice].FieldName;
		if ExtractPrefixe (NomChamps) = 'BLO' then continue; // pour ne prendre que les champs sup
  	TOBEtalon.AddChampSupValeur (NomChamps,getValInitChampsSup(NomChamps));
  end;
end;

procedure MemoriseChampsSupLigneOUV (NaturePiece : string);
var WithLigneFac : boolean;
		Sql : String;
    QQ : TQuery;

begin
  WithLigneFac := (pos(NaturePiece,'FBT;DAC;FBP;BAC')>0);
	Sql := MakeSelectLigneOuvBtp (WithLigneFac,'');
  Sql := Sql + ' WHERE 1=2'; // pour ne recup que la structure
  QQ := OpenSQL(Sql, True,-1, '', True);
  SMemoriseChampsSupLigneOuv(QQ);
  ferme (QQ);
end;

procedure MemoriseChampsSupLigneETL (NaturePiece : string;WithLigneCompl:boolean=false;WithEntete:boolean=false);
var WithLigneFac : boolean;
		Sql : String;
    QQ : TQuery;
    Document : Tdocument;

begin
  Document := Tdocument.Create;
  //
  Document.cledoc.NaturePiece := NaturePiece;
  WithLigneFac := (pos(NaturePiece,'FBT;DAC;FBP;BAC')>0);
	Sql := MakeSelectLigneBtp (WithLigneCompl,WithEntete,WithLigneFac,false);
  Sql := Sql + ' WHERE 1=2'; // pour ne recup que la structure
  QQ := OpenSQL(Sql, True,-1, '', True);
  SMemoriseChampsSupLigneETL(QQ);
  ferme (QQ);
  Document.Free;
end;


procedure AddlesChampsSupPieceTrait (TOBPT : TOB);
var Indice : integer;
		TOBEtl : TOB;
    NomChamps : string;
begin
	if TOBSTRUCPTRAIT.detail.count = 0 then exit;
  TOBEtl := TOBSTRUCPTRAIT.detail[0];
  //
  for Indice := 0 to TOBETL.NombreChampSup -1 do
  begin
  	NomChamps := TOBEtl.GetNomChamp(1000+Indice);
    if not TOBPT.FieldExists(NomChamps) then
    begin
  		TOBPT.AddChampSupValeur (NomChamps,TOBEtl.Getvaleur(1000+Indice));
    end;
  end;
end;

procedure InitStructureCptrait;
begin
	if TOBSTRUCPTRAIT.detail.count > 0 then TOBSTRUCPTRAIT.ClearDetail;
end;

procedure SMemoriseChampsPieceTrait(QQ : TQuery);
var TOBEtalon : TOB;
		Indice : integer;
    NomChamps : string;
begin
	if TOBSTRUCPTRAIT.detail.count > 0 then InitStructureCptrait;
  //
	TOBEtalon := TOB.Create ('LA LIGNE ETALON',TOBSTRUCPTRAIT,-1);
  //
  for Indice := 0 to QQ.FieldCount -1 do
  begin
  	NomChamps := QQ.Fields[Indice].FieldName;
		if ExtractPrefixe (NomChamps) = 'BPE' then continue; // pour ne prendre que les champs sup
  	TOBEtalon.AddChampSupValeur (NomChamps,getValInitChampsSup(NomChamps));
  end;
end;

procedure MemoriseChampsSupPIECETRAIT;
var Sql : string;
		QQ : TQuery;
begin
	Sql := GetSqlPieceTrait;
  Sql := Sql + ' WHERE 1=2';
  QQ := OpenSQL(Sql,True,1,'',true);
  SMemoriseChampsPieceTrait(QQ);
  ferme (QQ);
end;

procedure MemoriseChampsSupPARCL (requete : string);
var
    QQ : TQuery;

begin
  QQ := OpenSQL(requete, True,-1, '', True);
  SMemoriseChampsSupParcL(QQ);
  ferme (QQ);
end;

procedure AddlesChampsSupLigneEtl (TOBL : TOB);
var Indice : integer;
		TOBEtl : TOB;
    NomChamps : string;
begin
	if TOBStructLig.detail.count = 0 then exit;
  TOBEtl := TOBStructLig.detail[0];
  //
  for Indice := 0 to TOBETL.NombreChampSup -1 do
  begin
  	NomChamps := TOBEtl.GetNomChamp(1000+Indice);
    if not TOBL.FieldExists(NomChamps) then
    begin
  		TOBL.AddChampSupValeur (NomChamps,TOBEtl.Getvaleur(1000+Indice));
    end else
    begin
      if varisnull(TOBL.GetValue(Nomchamps)) then
  		  TOBL.putvalue (NomChamps,TOBEtl.Getvaleur(1000+Indice));
    end;
  end;

end;

procedure AddlesChampsSupParcl (TOBL : TOB);
var Indice : integer;
		TOBEtl : TOB;
    NomChamps : string;
begin
	if TOBStructParc.detail.count = 0 then exit;
  TOBEtl := TOBStructParc.detail[0];
  //
  for Indice := 0 to TOBETL.NombreChampSup -1 do
  begin
  	NomChamps := TOBEtl.GetNomChamp(1000+Indice);
    if not TOBL.FieldExists(NomChamps) then
    begin
  		TOBL.AddChampSupValeur (NomChamps,TOBEtl.Getvaleur(1000+Indice));
    end;
  end;
end;

procedure AddlesChampsSupOuv (TOBL : TOB);
var Indice : integer;
		TOBEtl : TOB;
    NomChamps : string;
begin
	if TOBStructLigOuv.detail.count = 0 then exit;
  TOBEtl := TOBStructLigOuv.detail[0];
  //
  for Indice := 0 to TOBETL.NombreChampSup -1 do
  begin
  	NomChamps := TOBEtl.GetNomChamp(1000+Indice);
    if not TOBL.FieldExists(NomChamps) then
    begin
  		TOBL.AddChampSupValeur (NomChamps,TOBEtl.Getvaleur(1000+Indice));
    end;
  end;
end;

INITIALIZATION
TOBSTRUCTLIG := TOB.Create (' LA  STRUCTURE LIGNE',nil,-1);
TOBStructParc := TOB.Create (' LA  STRUCTURE PARCL',nil,-1);
TOBStructLigOuv := TOB.Create (' LA  STRUCTURE OUVRAGE',nil,-1);
TOBSTRUCPTRAIT := TOB.Create (' LA  STRUCTURE PIECETRAIT',nil,-1);
FINALIZATION
TOBSTRUCTLIG.free;
TOBStructParc.free;
TOBStructLigOuv.free;
TOBSTRUCPTRAIT.free;

end.
