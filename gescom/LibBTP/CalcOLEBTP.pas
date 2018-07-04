unit CalcOLEBTP;

interface
uses SysUtils, Classes, Affaireutil, UAFO_Ressource,ActiviteUtil, HEnt1,
  {$IFNDEF EAGLCLIENT}
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  {$ENDIF}
    CalcOLEGescom,CalcOLEAffaire,ConfidentAffaire,SyntheseUtil,Dicobtp,HCtrls,UTOB,ParamSoc,uEntCommun;

function BTCalcOLEEtat(sf,sp : string) : variant ;

implementation
uses Entgc,FactComm;

Var CumulPar : Array[1..9] of double;
		CumulSit : Array[1..9] of double;
    CumulPrevu : Array[1..9] of double;
    CumulCumul : Array[1..9] of double;
    CumulMtPrevu : Array[1..9] of double;
		CumulRea : Array[1..9] of double;
    TOBPieces : TOB;

function BTReadParam(var St : String) : string ;
var i : Integer ;
begin
i:=Pos(';',St) ; if i<=0 then i:=Length(St)+1 ; result:=Copy(St,1,i-1) ; Delete(St,1,i) ;
Result := Trim(Result);
end ;


function GereMtCumule (sp : string) : variant;
Var st1,st2,st3,st4 : string;
    niv, i : integer;
    db : double;
begin
  // cumul pour totaux paragraphes :
  // Passer en paramètres :
  // - le type de traitement à effectuer (RAZ=Remise à zéro, REC=Récupération valeur, CUM=Cumul)
  // - le montant à cumuler
  // - le niveau d'imbrication de la ligne
  // - le type de ligne

  st1 := BTReadParam(sp) ;
  st2 := BTReadParam(sp) ;
  st3 := BTReadParam(sp) ;
  st4 := BTReadParam(sp) ;
  niv := StrToInt (st3);
  if st1 = 'REC' then
  begin
  	result:=CumulPar[niv];
  end
  else if st1 = 'CUM' then
  begin
    if (copy(st4,1,2)='DP')  then
    begin
      for i:=niv to 9 do
      begin
        CumulPar[i]:=0;
      end;
  	end;
  	db := valeur(st2);
    for i:=niv downto 0 do
    begin
    	CumulPar[i]:=CumulPar[i]+arrondi(db,V_PGI.OkDecV );
    end;
  end;
end;

function GetPourcentAvanc (sp : string) : variant;
Var st1,st2,st3,st4 : string;
    niv, i : integer;
    db : double;
begin
  result := 0;
  st1 := BTReadParam(sp) ;
  niv := StrToInt (st1);
  if CumulMtPrevu[niv] <> 0 then result:=Arrondi(((CumulCumul[niv] / CumulMtPrevu[Niv]) * 100),V_PGI.OKDecV);
end;

function EnregInfoLig (sp : string) : variant;
Var st1,st2,st3,st4,St5,st6,st7 : string;
    niv, i : integer;
    prevu,Realise,Situation,MtPrevu,MtCumul : double;
begin

  st1 := BTReadParam(sp) ; // Qte prevu
  st2 := BTReadParam(sp) ; // Qte Realisée
  st3 := BTReadParam(sp) ; // Montant Situation
  st4 := BTReadParam(sp) ; // Montant prévu
  st5 := BTReadParam(sp) ; // Montant cumul
  st6 := BTReadParam(sp) ; // Niveau imbrication
  st7 := BTReadParam(sp) ; // Type de ligne
  niv := StrToInt (st6);
  Prevu := valeur(st1);
  Realise := valeur (st2);
  Situation := valeur (St3);
  MtPrevu := valeur (St4);
  MtCumul := valeur(st5);
  if (copy(st7,1,2)='DP')  then
  begin
    for i:=niv to 9 do
    begin
      CumulMtPrevu[i]:=0;
      CumulPrevu[i]:=0;
      CumulRea[i]:=0;
      CumulSit[i]:=0;
      CumulCumul[i]:=0;
    end;
  end;
  if copy (st7,1,2) <> 'TP' then
  begin
    for i:=niv downto 1 do
    begin
      CumulPrevu[i]:=CumulPrevu[i]+Prevu;
      CumulRea[i]:=CumulRea[i]+Realise;
      CumulSit[i]:=CumulSit[i]+Arrondi(Situation,V_PGI.okDecV);
      CumulMtPrevu[i]:=CumulMtPrevu[i]+Arrondi(MtPrevu,V_PGI.OKDecV);
      CumulCumul[i]:=CumulCumul[i]+Arrondi(MtCumul,V_PGI.OKDecV);
    end;
  end;
end;

function GetMtPrevu (sp : string) : variant;
Var st1 : string;
    niv: integer;
begin
  st1 := BTReadParam(sp) ;
  niv := StrToInt (st1);
  result:=CumulMtPrevu[niv];
end;

function GetMtCumul (sp : string) : variant;
Var st1 : string;
    niv: integer;
begin
  st1 := BTReadParam(sp) ;
  niv := StrToInt (st1);
  result:=CumulCumul[niv];
end;

function GetMtSituation (sp : string) : variant;
Var st1 : string;
    niv: integer;
begin
  st1 := BTReadParam(sp) ;
  niv := StrToInt (st1);
  result:=CumulSit[niv];
end;

procedure AddPiece (cledoc : r_cledoc;TOBPieces : TOB);
var TOBPiece : TOB;
begin
  TOBPiece := TOB.Create('UNE PIECE',TOBPieces,-1);
  TOBPiece.addChampSupValeur('NATURE',cledoc.naturepiece);
  TOBPiece.addChampSupValeur('SOUCHE',cledoc.Souche);
  TOBPiece.addChampSupValeur('NUMERO',cledoc.NumeroPiece);
  TOBPiece.addChampSupValeur('MONTANTHT',0);
  TOBPiece.addChampSupValeur('MONTANTHTBRUT',0);
  TOBPiece.addChampSupValeur('MONTANTTTC',0);
end;

function FindPiece (cledoc : r_cledoc;TOBPieces : TOB) : TOB;
begin
  result := TOBPieces.FindFirst (['NATURE','SOUCHE','NUMERO'],[cledoc.naturePiece,cledoc.souche,cledoc.NumeroPiece],true);
end;

function EnregPiece (sp : string) : variant;
var		TOBUnePiece : TOB;
		st1 : string;
    cledoc : r_cledoc;
begin
  st1 := sp ; // pieceprecedente
  if st1 <> '' then
  begin
    DecodeRefPiece (st1,cledoc);
    TOBUnePiece := FindPiece (cledoc,TOBPieces);
    if TOBUnePiece = nil then AddPiece (cledoc,TOBPieces)
  end;
  result := 0;
end;


function GetMontantHtdevprevu (sp : string) : variant;
var QQ2 : TQuery;
    valeur : double;
		TOBUnePiece : TOB;
    indice : integer;
begin
	valeur:=0;
  for indice := 0 to TOBPieces.detail.count -1 do
  begin
    TOBUnePiece := TOBPieces.detail[Indice];
    if TOBUnePiece.getValue('MONTANTHT') = 0 then
    begin
      QQ2 := OpenSql ('SELECT GP_TOTALHTDEV,GP_TOTALTTCDEV,GP_TOTALREMISEDEV FROM PIECE WHERE GP_NATUREPIECEG="'+TOBUnePIece.getValue('NATURE') +'" AND '+
                  'GP_SOUCHE="'+TOBUnePIece.getValue('SOUCHE')+'" AND GP_NUMERO='+IntToStr(TOBUnePIece.getValue('NUMERO')),true,-1,'',true);
      if not QQ2.eof then
      begin
        TOBUnePiece.putValue('MONTANTHT', QQ2.findField('GP_TOTALHTDEV').AsFloat);
        TOBUnePiece.putValue('MONTANTTTC', QQ2.findField('GP_TOTALTTCDEV').AsFloat);
        TOBUnePiece.putValue('MONTANTHTBRUT', QQ2.findField('GP_TOTALHTDEV').AsFloat+QQ2.FindField('GP_TOTALREMISEDEV').AsFloat);
      end;
      ferme (QQ2);
    end;
    valeur := valeur + TOBUnePiece.GetValue('MONTANTHT');
  end;
  result := valeur;
end;

function GetGestionAcompteCpt : variant;
begin
	if (GetParamSoc('SO_BTCOMPTAREGL')) then result := 'X' else result := '-';
end;

function GetMontantTTCdevprevu (sp : string) : variant;
var QQ2 : TQuery;
    valeur : double;
		TOBUnePiece : TOB;
    indice : integer;
begin
	valeur:=0;
  for indice := 0 to TOBPieces.detail.count -1 do
  begin
    TOBUnePiece := TOBPieces.detail[Indice];
    if TOBUnePiece.getValue('MONTANTTTC') = 0 then
    begin
      QQ2 := OpenSql ('SELECT GP_TOTALHTDEV,GP_TOTALTTCDEV,GP_TOTALREMISEDEV FROM PIECE WHERE GP_NATUREPIECEG="'+TOBUnePIece.getValue('NATURE') +'" AND '+
                  'GP_SOUCHE="'+TOBUnePIece.getValue('SOUCHE')+'" AND GP_NUMERO='+IntToStr(TOBUnePIece.getValue('NUMERO')),true,-1,'',true);
      if not QQ2.eof then
      begin
        TOBUnePiece.putValue('MONTANTHT', QQ2.findField('GP_TOTALHTDEV').AsFloat);
        TOBUnePiece.putValue('MONTANTTTC', QQ2.findField('GP_TOTALTTCDEV').AsFloat);
        TOBUnePiece.putValue('MONTANTHTBRUT', QQ2.findField('GP_TOTALHTDEV').AsFloat+QQ2.FindField('GP_TOTALREMISEDEV').AsFloat);
      end;
      ferme (QQ2);
    end;
    valeur := valeur + TOBUnePiece.GetValue('MONTANTTTC');
  end;
  result := valeur;
end;

function GetMontantBrutdevprevu (sp : string) : variant;
var QQ2 : TQuery;
    valeur : double;
		TOBUnePiece : TOB;
    indice : integer;
begin
	valeur:=0;
  for indice := 0 to TOBPieces.detail.count -1 do
  begin
    TOBUnePiece := TOBPieces.detail[Indice];
    if TOBUnePiece.getValue('MONTANTHTBRUT') = 0 then
    begin
      QQ2 := OpenSql ('SELECT GP_TOTALHTDEV,GP_TOTALTTCDEV,GP_TOTALREMISEDEV FROM PIECE WHERE GP_NATUREPIECEG="'+TOBUnePIece.getValue('NATURE') +'" AND '+
                  'GP_SOUCHE="'+TOBUnePIece.getValue('SOUCHE')+'" AND GP_NUMERO='+IntToStr(TOBUnePIece.getValue('NUMERO')),true,-1,'',true);
      if not QQ2.eof then
      begin
        TOBUnePiece.putValue('MONTANTHT', QQ2.findField('GP_TOTALHTDEV').AsFloat);
        TOBUnePiece.putValue('MONTANTTTC', QQ2.findField('GP_TOTALTTCDEV').AsFloat);
        TOBUnePiece.putValue('MONTANTHTBRUT', QQ2.findField('GP_TOTALHTDEV').AsFloat+QQ2.FindField('GP_TOTALREMISEDEV').AsFloat);
      end;
      ferme (QQ2);
    end;
    valeur := valeur + TOBUnePiece.GetValue('MONTANTHTBRUT');
  end;
  result := valeur;
end;

function InitEnregPieces (sp : string ) : variant;
begin
	TOBPieces.cleardetail;
end;

function FreePieces  (sp : string ) : variant;
begin
//	TOBPieces.free;
end;

function BTCalcOLEEtat(sf,sp : string) : variant ;
BEGIN
if copy(sf,1,5) ='GCPDF' then
  Result:=GCGetPortsEdt(sf,sp)
else if sf='GCECHE' then
  Result:=GCGetEchesEdt(sf,sp)
else if copy(sf, 1, 9) = 'GCDIMVALR' then
  Result := GCGetDimValR(sp)
else if copy(sf,1,3) <> 'BT_' then
  Result := AfCalcOLEEtat(sf,sp)
else if sf='BT_CALTOTPAR' then
	result := GereMtCumule (sp)
else if sf='BT_ENREGINFOLIG' then
	result := EnregInfoLig (sp)
else if sf='BT_POURCENTAVANC' then
	result := GetPourcentAvanc (sp)
else if sf='BT_MTPREVU' then
	result := GetMtPrevu (sp)
else if sf='BT_MTSITUATION' then
	result := GetMtSituation (sp)
else if sf='BT_MTCUMULE' then
	result := GetMtCumul (sp)
else if sf='BT_INITENREGPIECE' then
	result := InitEnregPieces(sp)
else if sf='BT_FINENREGPIECE' then
	result := FreePieces(sp)
else if sf='BT_ENREGPIECE' then
	result := EnregPiece(sp)
else if sf='BT_MTDEVHT' then
	result := GetMontantHtdevprevu (sp)
else if sf='BT_MTDEVHTBRUT' then
	result := GetMontantBrutdevprevu (sp)
else if sf='BT_MTDEVTTC' then
	result := GetMontantTTCdevprevu (sp)
else if sf='BT_GESTIONACOMPTECPT' then
	result := GetgestionAcompteCpt;

END;

initialization
TOBPieces := TOB.Create ('LES PIECES',nil,-1);
finalization
TOBPieces.free;
end.
