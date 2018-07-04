unit UspecifPOC;

interface

uses
  Classes,
  SysUtils,
  uTob,
  hEnt1,
  hCtrls,
  EntGC,
  SaisUtil,
  wCommuns,
  uEntCommun,
  UtilConso,
  Db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  affaireutil,FE_Main
  ;

var TOBlesContratsST : TOB;

procedure GetCoefPoc (Affaire : string; var COEFFG,COEFMARGE : double);
procedure CalculeDonneelignePOC (TOBL : TOB; var COEFFG,COEFMARGE : double);
procedure ApliqueCoeflignePOC (TOBL : TOB; DEV : Rdevise);
procedure AppliqueRGPOC (TOBfacture,TOBAffaire,TOBPorcs,TOBTiers,TOBPIECERG : TOB; DEV : RDevise; NumSituation : integer);
procedure EnregistreAvancePOC(TOBFacture,TOBAffaire,TOBPOrcs,TOBBases : TOB;DEV : Rdevise);
function RestitueAvancePOC(TOBFacture,TOBAffaire,TOBPOrcs,TOBBases : TOB;PourcentAvanc : double; DEV : Rdevise; NumSituation : integer) : boolean;
function GetTauxAffairePOC(CodeAffaire : string; DatePiece : TdateTime) : Double;
procedure EnregistreSit0POC(TOBFacture,TOBSituations : TOB; DEV : Rdevise);
function GetInfoMarcheST(Affaire,Fournisseur,CodeMarche,Champs : string) : Variant;
function GetMtPaiementDir (Affaire,Fournisseur,CodeMarche : string) : double;
procedure LibereMemContratST;
procedure AppelsBAST;
function GetNextNumSituation (TOBPiece : TOB) : integer;
procedure LoadLesTOBTS (Cledoc : R_CLEDOC;TOBTSPOC : TOB);
procedure GestionTsPOC(TOBPiece,OneTOB,TOBTSPOC : TOB);
procedure ValideLesTOBTS ( TOBPiece,TOBTSPOC : TOB);
function SetFactureReglePOC (TOBPiece : TOB) : boolean;
procedure AppelMarcheST (CodeChantier,SousTrait,CodeMarche : string);

implementation
uses FactComm,UtilPGI,FactTOB,FactPiece,FactRG,FactUtil,ParamSOc,ENt1,AglInit,M3FP,cbpPath,LicUtil,UtilTOBPiece,UconnectBSV;

procedure GetCoefPoc (Affaire : string; var COEFFG,COEFMARGE : double);
var fCOEFFG,COEFFS,COEFSAV,COEFFD : double;
    TOBAffaire : TOB;
begin
  if Affaire='' then Exit;
  TOBAffaire := FindCetteAffaire (Affaire);
  if TOBAffaire = nil then
  begin
    StockeCetteAffaire (Affaire);
    TOBAffaire := FindCetteAffaire (Affaire);
  end;
  fCOEFFG := TOBAffaire.GetDouble('AFF_COEFFG'); if fCOEFFG = 0 then fCOEFFG := 1;
  COEFFS := TOBAffaire.GetDouble('AFF_COEFFS'); if COEFFS = 0 then CoefFS := 1;
  COEFSAV := TOBAffaire.GetDouble('AFF_COEFSAV'); if COEFSAV = 0 then COEFSAV := 1;
  COEFFD := TOBAffaire.GetDouble('AFF_COEFFD');  if COEFFD = 0 then COEFFD := 1;
  COEFFG := ARRONDI(fCOEFFG * COEFFS * COEFSAV * COEFFD,4);
  COEFMARGE := TOBAffaire.GetDouble('AFF_COEFMARG'); if COEFMARGE = 0 then COEFMARGE := 1;
end;

procedure CalculeDonneelignePOC (TOBL : TOB; var COEFFG,COEFMARGE : double);
var prefixe,Affaire : string;
begin
  prefixe := GetprefixeTable(TOBL);
  if GetInfoParPiece(TOBL.GetString(prefixe+'_NATUREPIECEG'), 'GPP_VENTEACHAT') <> 'VEN' then exit;
  Affaire := TOBL.GetString(prefixe+'_AFFAIRE');
  if Affaire ='' then Exit;
  GetCoefPoc (Affaire,COEFFG,COEFMARGE);
end;

procedure ApliqueCoeflignePOC (TOBL : TOB; DEV : Rdevise);
var coeffg,coefmarg : double;
    prefixe : string;
    fdev : rdevise;
begin
  fdev.Code := '';
  prefixe := GetprefixeTable(TOBL);
  coeffg := 0;
  coefmarg := 0;
  CalculeDonneelignePOC (TOBL,coeffg,coefmarg);
  if (DEV.Code <> '') then
  begin
    fDEv := DEV
  end else
  begin
    if TOBL.FieldExists(prefixe+'_DEVISE') then
    begin
      fdev.Code := TOBL.getString(prefixe+'_DEVISE');
      GetInfosDevise(fDEV);
    end;
  end;
  if coeffg <> 0 then
  begin
    TOBL.putValue(prefixe+'_COEFFG',CoefFg-1);
    TOBL.SetDouble(prefixe+'_DPR',ARRONDI(TOBL.GetDouble(prefixe+'_DPA')*Coeffg,V_PGI.OkDecP));
  end;
  if coefmarg <> 0 then
  begin
    TOBL.putValue(prefixe+'_COEFMARG',coefmarg);
    TOBL.PutValue('POURCENTMARG',Arrondi((coefmarg-1)*100,2));
    TOBL.SetDouble(prefixe+'_PUHT',ARRONDI(TOBL.GetDouble(prefixe+'_DPR')*CoefMarg,V_PGI.OkDecP));
    if fdev.code <> '' then
    begin
      TOBL.SetDouble(prefixe+'_PUHTDEV',pivottodevise(TobL.GetDouble(prefixe+'_PUHT'),DEV.Taux,DEV.quotite,V_PGI.okdecP ));
    end else
    begin
      TOBL.SetDouble(prefixe+'_PUHTDEV',TobL.GetDouble(prefixe+'_PUHT'));
    end;

  end;
end;

function GetMtCautions (Affaire : string;DateMax : TDateTime) : double;
var QQ: TQuery;
begin
  Result := 0;
  QQ:= openSql('SELECT SUM(BAR_CAUTIONMT) FROM AFFAIRERG WHERE BAR_AFFAIRE="'+Affaire+'" AND BAR_DATECAUTION <= "'+USdateTime(DateMax)+'"',True,1,'',true);
  if not QQ.eof then
  begin
    Result := QQ.fields[0].asFloat;
  end;
  ferme (QQ); 
end; 


function GetRGCumulesUsed (AffaireDevis : string; DatePiece : TdateTime; NumSituation : integer) : double;
var TOBBST : TOB;
    SQl : String;
    QQ : TQuery;
begin
  Result := 0;
  TOBBST := TOB.create ('LES SITUATIONS',nil,-1);
  try
    SQL := 'SELECT BST_NATUREPIECE,BST_SOUCHE,BST_NUMEROFAC,BST_NUMEROSIT,SUM(PRG_MTTTCRGDEV) AS MTREGUSED '+
           'FROM BSITUATIONS '+
           'LEFT JOIN PIECERG ON BST_NATUREPIECE=PRG_NATUREPIECEG AND BST_SOUCHE=PRG_SOUCHE AND BST_NUMEROFAC=PRG_NUMERO '+
           'WHERE '+
           'BST_SSAFFAIRE="'+Affairedevis+'" AND '+
           'BST_VIVANTE="X" AND '+
           'BST_NUMEROSIT < '+IntToStr(NumSituation)+' AND '+
           'PRG_CAUTIONMTDEV <> 0 '+
           'GROUP BY BST_NATUREPIECE,BST_SOUCHE,BST_NUMEROFAC,BST_NUMEROSIT';
    QQ := OpenSql (sql,True,-1,'',true);
    if not QQ.eof then
    begin
      QQ.first;
      while not QQ.eof do
      begin
        Result := Result + QQ.fields[4].asFloat;
        QQ.next;
      end;
    end;
    ferme (QQ);
  finally
    TOBBST.free;
  end;
end;

function GetNextNumSituation (TOBPiece : TOB) : integer;
var REq : string;
    Q : Tquery;
begin
  result := -1;
  Req:='SELECT BST_NUMEROSIT '+
       'FROM BSITUATIONS WHERE '+
       'BST_SSAFFAIRE="'+ TOBPiece.GetValue('GP_AFFAIREDEVIS') + '" AND BST_VIVANTE="X" '+
       'ORDER BY BST_SSAFFAIRE,BST_NUMEROSIT DESC';
  Q:=OpenSQL(Req,TRUE,-1,'',true);
  if not Q.EOF then
  begin
    result := Q.Fields[0].AsInteger +1;
  end;
  Ferme(Q) ;
end;


procedure AppliqueRGPOC (TOBfacture,TOBAffaire,TOBPorcs,TOBTiers,TOBPIECERG : TOB; DEV : RDevise; NumSituation : integer);
var TOBR : TOB;
    MtCautions,MtRGUsed,Cautionrestante,PorcTTC,MtPieceTTC : double;
    DatePiece : TDateTime;
begin
  if NumSituation = -1 then NumSituation := GetNextNumSituation(TOBfacture);
  if TOBfacture.GetDateTime('GP_DATEPIECE') < TOBAffaire.GetDouble('AFF_DATERG') then Exit;
  //
  DatePiece := TOBfacture.GetDateTime('GP_DATEPIECE');
  PorcTTC := CalculPort (false,TOBPorcs);

  MtPieceTTC := TOBFacture.getvalue('GP_TOTALTTCDEV') - PorcTTC;
  //
  if TOBAffaire.GetDouble('AFF_TAUXRG')=0 then Exit;
  //
  MtCautions := GetMtCautions (TOBFacture.GetString('GP_AFFAIRE'),DatePiece);
  MtRGUsed := GetRGCumulesUsed (TOBFacture.GetString('GP_AFFAIREDEVIS'),DatePiece,NumSituation);
  Cautionrestante := MtCautions - MtRGUsed; if Cautionrestante < 0 then Cautionrestante := 0;
  //
  TOBR := TOB.Create('PIECERG',TOBPIECERG,-1);
  InitLigneRg (TOBR,TOBfacture);
  TOBR.SetString('PRG_APPLICABLE','X');
  TOBR.SetString('PRG_NATUREPIECEG',TOBfacture.GetString('GP_NATUREPIECEG'));
  TOBR.SetString('PRG_SOUCHE',TOBfacture.GetString('GP_SOUCHE'));
  TOBR.SetInteger('PRG_NUMERO',TOBfacture.GetInteger('GP_NUMERO'));
  TOBR.SetInteger('PRG_INDICEG',TOBfacture.GetInteger('GP_INDICEG'));
  TOBR.SetInteger('PRG_INDICEG',TOBfacture.GetInteger('GP_INDICEG'));
  TOBR.SetString('PRG_TYPERG','TTC');
  TOBR.SetDouble('PRG_TAUXRG',TOBAffaire.GetDouble('AFF_TAUXRG'));
  TOBR.SetDouble('PRG_TAUXDEV',TOBfacture.GetDouble('GP_TAUXDEV'));
  TOBR.SetDouble('PRG_COTATION',TOBfacture.GetDouble('GP_COTATION'));
  TOBR.putvalue('PRG_DATETAUXDEV',TOBfacture.GetValue('GP_DATETAUXDEV'));
  TOBR.putvalue('PRG_DEVISE',TOBfacture.GetValue('GP_DEVISE'));
  TOBR.putvalue('PRG_SAISIECONTRE','-');
  //
  TOBR.SetDouble('PRG_MTTTCRGDEV', ARRONDI(MtPieceTTC*(TOBR.GetDouble('PRG_TAUXRG')/100),DEV.Decimale));
  TOBR.SetDouble('PRG_MTTTCRG', DEVISETOPIVOT(TOBR.GetDouble('PRG_MTTTCRGDEV'),DEV.Taux,DEV.Quotite));
  //
  if Cautionrestante <> 0 then
  begin
    TOBR.SetString('PRG_NUMCAUTION', '111');
    TOBR.SetString('PRG_BANQUECP', 'INT');
    TOBR.SetDouble('PRG_CAUTIONMTDEV', Cautionrestante);
    TOBR.SetDouble('PRG_CAUTIONMT', DEVISETOPIVOT(TOBR.GetDouble('PRG_CAUTIONMTDEV'),DEV.Taux,DEV.Quotite));
  end;
  //

end;

function GetAvanceRest (AffaireDevis ,CodePort : string; NumSituation : integer) : double;
var TOBBST : TOB;
    SQl : String;
    QQ : TQuery;
begin
  Result := 0;
  TOBBST := TOB.create ('LES SITUATIONS',nil,-1);
  try
    SQL := 'SELECT BST_NATUREPIECE,BST_SOUCHE,BST_NUMEROFAC,BST_NUMEROSIT,SUM(GPT_TOTALTTCDEV) AS MTAVANCEUSED '+
           'FROM BSITUATIONS '+
           'LEFT JOIN PIEDPORT ON BST_NATUREPIECE=GPT_NATUREPIECEG AND BST_SOUCHE=GPT_SOUCHE AND BST_NUMEROFAC=GPT_NUMERO '+
           'WHERE '+
           'BST_SSAFFAIRE="'+Affairedevis+'" AND BST_VIVANTE="X" AND BST_NUMEROSIT < '+IntToStr(NumSituation)+' AND GPT_CODEPORT="'+CodePort+'" '+
           'GROUP BY BST_NATUREPIECE,BST_SOUCHE,BST_NUMEROFAC,BST_NUMEROSIT';
    QQ := OpenSql (sql,True,-1,'',true);
    if not QQ.eof then
    begin
      QQ.first;
      while not QQ.eof do
      begin
        Result := Result + QQ.fields[4].asFloat;
        QQ.next;
      end;
    end;
    ferme (QQ);
  finally
    TOBBST.free;
  end;
  if result < 0 then result := result * (-1);
end;

procedure EnregistreAvancePOC(TOBFacture,TOBAffaire,TOBPOrcs,TOBBases : TOB;DEV : Rdevise);
var TOBP,TOBPORT : TOB;
    MtAvance  : double;
    CodePort : String;
    QQ : TQuery;
    DD : TDateTime;
begin
  //
  TOBPORT := TOB.Create('PORT',nil,-1);
  TRY
    MtAvance := TOBAffaire.GetDouble('AFF_MTAVANCE');
    if MtAvance=0 then Exit;
    CodePort := GetparamSocSecur('SO_BTREGAVANCEPOC',''); if CodePort = '' then exit;
    QQ := OpenSQL('SELECT * FROM PORT WHERE GPO_CODEPORT="'+CodePort+'"',True,1,'',true);
    if not QQ.eof then
    begin
      TOBPORT.SelectDB('',QQ);
    end;
    ferme (QQ);
    // UTILISABLE ??
    if TOBPORT.GetString('GPO_CODEPORT')='' then Exit;
    DD := TOBPORT.GetValue('GPO_DATESUP');
    if ((DD < V_PGI.DateEntree) and (DD > iDate1900)) then Exit;
    if (TOBPORT.GetValue('GPO_FRAISREPARTIS')='X')  then Exit;
    if (TOBPORT.GetValue('GPO_FERME') = 'X') then exit;
    //
    TOBP := TOB.Create('PIEDPORT',TOBPOrcs,-1);
    TOBP.SetString('GPT_CODEPORT', CodePort);
    ChargeTobPiedPort (TOBP,TOBPORT,TOBFacture);
    //
    if TOBP.GetBoolean('GPO_RETENUEDIVERSE') then TOBP.SetDouble('GPT_BASETTCDEV', MtAvance)
                                             else TOBP.SetDouble('GPT_BASETTCDEV', MtAvance * -1);
    TOBP.SetDouble('GPT_BASETTC', DeviseToPivot(TOBP.GetDouble('GPT_BASETTCDEV'),DEV.Taux,DEV.Quotite));
    CalculMontantsPiedPort(TobFacture, TOBP,TOBBases);
    //
  FINALLY
    TOBPORT.Free;
  END;

end;


function RestitueAvancePOC(TOBFacture,TOBAffaire,TOBPOrcs,TOBBases : TOB;PourcentAvanc : double; DEV : Rdevise; NumSituation : integer) : boolean;
var TOBP,TOBPORT : TOB;
    MtAvance,DebutRest,FinRest,PourcReel,RelAvance,AvanceRest : double;
    CodePort : String;
    QQ : TQuery;
    DD : TDateTime;
begin
  //
  Result := true;
  if NumSituation = -1 then NumSituation := GetNextNumSituation(TOBfacture);
  TOBPORT := TOB.Create('PORT',nil,-1);
  TRY
    MtAvance := TOBAffaire.GetDouble('AFF_MTAVANCE');
    if MtAvance=0 then Exit;
    CodePort := GetparamSocSecur('SO_BTRACPOC',''); if CodePort = '' then exit;
    QQ := OpenSQL('SELECT * FROM PORT WHERE GPO_CODEPORT="'+CodePort+'"',True,1,'',true);
    if not QQ.eof then
    begin
      TOBPORT.SelectDB('',QQ);
    end;
    ferme (QQ);
    AvanceRest := GetAvanceRest (TOBFacture.GetString('GP_AFFAIREDEVIS'),CodePort,NumSituation);
    // UTILISABLE ??
    if TOBPORT.GetString('GPO_CODEPORT')='' then Exit;
    DD := TOBPORT.GetValue('GPO_DATESUP');
    if ((DD < V_PGI.DateEntree) and (DD > iDate1900)) then Exit;
    if (TOBPORT.GetValue('GPO_FRAISREPARTIS')='X')  then Exit;
    if (TOBPORT.GetValue('GPO_FERME') = 'X') then exit;
    //
    DebutRest := TOBAffaire.GetDouble('AFF_DEBRESTAVANCE');
    FinRest := TOBAffaire.GetDouble('AFF_FINRESTAVANCE');
    //
    if (DebutRest = 0) and (FinREst = 0) then
    begin
      PourcReel := PourcentAvanc;
      RelAvance := (MtAvance * PourcReel);
    end else
    begin
      if (PourcentAvanc < DebutRest) then exit;
      //  Soit : (% Avancement - %Debut restitution) / (%Fin de restitution - %Debut restitution) = % de restitution sur cette facture
      PourcReel  := (PourcentAvanc - DebutRest) / (FinRest - DebutRest);
      if PourcentAvanc < FinRest then RelAvance := (MtAvance * PourcReel) - AvanceRest else RelAvance := MtAvance - AvanceRest;
    end;
    TOBP := TOB.Create('PIEDPORT',TOBPOrcs,-1);
    TOBP.SetString('GPT_CODEPORT', CodePort);
    ChargeTobPiedPort (TOBP,TOBPORT,TOBFacture);
    //
    if TOBP.GetBoolean('GPO_RETENUEDIVERSE') then TOBP.SetDouble('GPT_BASETTCDEV', RelAvance * -1)
                                             else TOBP.SetDouble('GPT_BASETTCDEV', RelAvance);
    TOBP.SetDouble('GPT_BASETTC', DeviseToPivot(TOBP.GetDouble('GPT_BASETTCDEV'),DEV.Taux,DEV.Quotite));
    CalculMontantsPiedPort(TobFacture, TOBP,TOBBases);

    //
  FINALLY
    TOBPORT.Free;
  END;

end;

function GetTauxAffairePOC(CodeAffaire : string; DatePiece : TdateTime) : Double;
var TOBAffaire : TOB;
begin
  Result := 0;
  if CodeAffaire='' then Exit;
  TOBAffaire := FindCetteAffaire (CodeAffaire);
  if TOBAffaire = nil then
  begin
    StockeCetteAffaire (CodeAffaire);
    TOBAffaire := FindCetteAffaire (CodeAffaire);
  end;
  if DatePiece >= TOBaffaire.getDateTime('AFF_DATERG') then
  begin
    result := TOBaffaire.getDouble('AFF_TAUXRG');
  end;
end;

procedure EnregistreSit0POC(TOBFacture,TOBSituations : TOB; DEV : Rdevise);
var TT,TOBT,TOBAffaire : TOB;
    CodeAffaire : string;
    Taux : Double;
begin
  CodeAffaire := TOBFacture.getString('GP_AFFAIRE');
  if CodeAffaire='' then Exit;
  if ExisteSQL('SELECT 1 FROM BSITUATIONS WHERE BST_SSAFFAIRE="'+TOBFacture.getString('GP_AFFAIREDEVIS')+'" AND BST_NUMEROSIT=0 AND BST_NATUREPIECE="FBT" AND BST_INDICESIT=0 AND BST_VIVANTE="X"') then
  begin
    ExecuteSQL('DELETE FROM BSITUATIONS WHERE BST_SSAFFAIRE="'+TOBFacture.getString('GP_AFFAIREDEVIS')+'" AND BST_NUMEROSIT=0 AND BST_NATUREPIECE="FBT" AND BST_INDICESIT=0 AND BST_VIVANTE="X"');
  end;

  TOBAffaire := FindCetteAffaire (CodeAffaire);
  if TOBAffaire = nil then
  begin
    StockeCetteAffaire (CodeAffaire);
    TOBAffaire := FindCetteAffaire (CodeAffaire);
  end;
  TT := TOB.Create ('BSITUATIONS',TOBSituations,-1);   
  TT.SetString('BST_NATUREPIECE','FBT');
  TT.SetInteger('BST_NUMEROFAC',0);
  TT.SetString('BST_SOUCHE',GetSoucheG('FBT', TOBFacture.GetValue('GP_ETABLISSEMENT'), TOBFacture.GetValue('GP_DOMAINE')));
  TT.SetInteger('BST_NUMEROSIT',0);
  TT.SetString('BST_AFFAIRE',TOBFacture.getString('GP_AFFAIRE'));
  TT.SetString('BST_SSAFFAIRE',TOBFacture.getString('GP_AFFAIREDEVIS'));
  TT.SetDouble('BST_MONTANTTTC',TOBAffaire.getDouble('AFF_MTAVANCE'));
  TOBT:=VH^.LaTOBTVA.FindFirst(['TV_TVAOUTPF','TV_REGIME','TV_CODETAUX'],['TX1',TOBFacture.GetValue('GP_REGIMETAXE'),'TN'],False) ;
  if TOBT <> nil then Taux:=1+(TOBT.GetValue('TV_TAUXVTE')/100)
                 else Taux := 0;

  TT.SetDouble('BST_MONTANTHT',Arrondi(TT.getDouble('BST_MONTANTTTC')/ Taux,DEV.decimale));
  TT.SetDouble('BST_MONTANTTVA',TT.getDouble('BST_MONTANTTTC')-TT.getDouble('BST_MONTANTHT'));
  TT.SetDateTime('BST_DATESIT',TOBFacture.geTDateTime('GP_DATEPIECE'));
  TT.SetString('BST_VIVANTE','X');
end;

function GetServerName : string;
var Qs : TQuery;
begin
  try
    Qs := OpenSQL('SELECT CAST (@@SERVERNAME AS VARCHAR) AS MONSERVEUR',true);
    if not Qs.Eof then result := Qs.FindField('MONSERVEUR').AsString
                  else result := '';
  finally
    Ferme(Qs);
  end;
end;

procedure AppelMarcheST (CodeChantier,SousTrait,CodeMarche : string);
var DBName,ServerName : string;
    TheLance,EmailPasswd : string;
begin
  EmailPasswd := DecryptageSt(V_PGI.EMailPassword);
  DBName := V_PGI.DBName;
  ServerName := GetServerName;
  TheLance := IncludeTrailingBackslash(TcbpPath.GetCegid)+'Specif-POC\APP\MarcheST.exe /userLSE='+V_PGI.User+' /EmailPwd='+EmailPasswd +' /Serveur='+ServerName+' /BaseDeDonnees='+DBName+' /CodeChantier="'+CodeChantier+'" /SousTraitant="'+SousTrait+'" /CodeMarche="'+CodeMarche+'" /Action=M';
  FileExecAndWait (TheLance);
end;

function AGLAppelMarcheST( parms: array of variant; nb: integer ) : variant;
var CodeChantier,SousTrait,CodeMarche: string;
begin
  CodeChantier := parms[0];
  SousTrait := parms[1];
  CodeMarche := parms[2];
  AppelMarcheST (CodeChantier,SousTrait,CodeMarche);
end;

function AGLAppelCreationMarcheST( parms: array of variant; nb: integer ) : variant;
var CodeChantier,DBName,ServerName : string;
    TheLance,EmailPasswd : string;
begin
  EmailPasswd := DecryptageSt(V_PGI.EMailPassword);
  CodeChantier := parms[0];
  DBName := V_PGI.DBName;
  ServerName := GetServerName;
  TheLance := IncludeTrailingBackslash(TcbpPath.GetCegid)+'Specif-POC\APP\MarcheST.exe /userLSE='+V_PGI.User +' /EmailPwd='+EmailPasswd+' /Serveur='+ServerName+' /BaseDeDonnees='+DBName+' /CodeChantier="'+CodeChantier+'" /Action=C';
  FileExecAndWait (TheLance);
end;

function GetInfoMarcheST(Affaire,Fournisseur,CodeMarche,Champs : string) : Variant;
var TT : TOB;
    QQ : TQuery;
    Marche : string;
begin
  Result := #0;
  if (affaire = '') or (Fournisseur='') or (CodeMarche='') then Exit;
  Marche := Copy(CodeMarche,1,8);
  if Marche = '' then exit;
  TT := TOBlesContratsST.FindFirst(['BM0_AFFAIRE','BM0_FOURNISSEUR','BM0_MARCHE'],[Affaire,Fournisseur,Marche],true);
  if TT = nil then
  begin
    QQ := OpenSQL('SELECT * FROM BTMARCHEST WHERE BM0_AFFAIRE="'+Affaire+'" AND BM0_FOURNISSEUR="'+Fournisseur+'" AND BM0_MARCHE="'+Marche+'"',True,1,'',true );
    if not QQ.eof then
    begin
      TT := TOB.Create('BTMARCHEST',TOBlesContratsST,-1);
      TT.SelectDB('',QQ);
    end;
    ferme (QQ);
  end;
  if TT = nil then Exit;
  Result := TT.GetValue('BM0_'+Champs);
end;

function GetMtPaiementDir (Affaire,Fournisseur,CodeMarche : string) : double;
var QQ: TQuery;
    SQL : string;
begin
  Result := 0;
  SQL := 'SELECT BM1_MTPAIEDIR FROM BTMARCHESTDET WHERE BM1_AFFAIRE="'+Affaire+'" AND BM1_FOURNISSEUR="'+Fournisseur+'" AND BM1_CODEMARCHE="'+CodeMarche+'"';
  QQ := OpenSQL(SQL,True,1,'',true);
  if not QQ.eof then
  begin
    Result := QQ.fields[0].AsFloat;
  end;
  ferme (QQ);
end;


procedure LibereMemContratST;
begin
  TOBlesContratsST.ClearDetail;
end;

procedure AppelsBAST;
var DBName,ServerName,EmailPasswd : string;
    TheLance : string;
begin
  DBName := V_PGI.DBName;
  EmailPasswd := DecryptageSt(V_PGI.EmailSmtpPassword);
  ServerName := GetServerName;
  TheLance := IncludeTrailingBackslash(TcbpPath.GetCegid)+'Specif-POC\APP\MarcheST.exe /userLSE='+V_PGI.User +' /Serveur='+ServerName+' /EmailPwd='+EmailPasswd+' /BaseDeDonnees='+DBName+' /Action=S';
  FileExecAndWait (TheLance);
end;

function CreTOBTS (TheNumOrdre : string; TOBTSPOC : TOB) : TOB;
begin
  result := TOB.Create('UNE LIGNE',TOBTSPOC,-1);
  result.AddChampSupValeur('NUMORDRE',TheNumOrdre);
  result.AddChampSupValeur('TOTALTS',0);
  result.AddChampSupValeur('MODIFIED','-');
  result.AddChampSupValeur('DESIGNATION','');
end;

procedure DeleteLesTOBTS (cledoc : r_cledoc);
begin
  if ExisteSQL('SELECT 1 FROM BLIGNEETS WHERE '+WherePiece(Cledoc,ttdTSPOC,false)) then
  begin
    ExecuteSQL ('DELETE FROM BLIGNEETS WHERE '+WherePiece(Cledoc,ttdTSPOC,false));
  end;
  if ExisteSQL('SELECT 1 FROM BLIGNETS WHERE '+WherePiece(Cledoc,ttdTSDetPOC,false)) then
  begin
    ExecuteSQL ('DELETE FROM BLIGNETS WHERE '+WherePiece(Cledoc,ttdTSDetPOC,false));
  end;
end;

procedure ValideLesTOBTS ( TOBPiece,TOBTSPOC : TOB);

  procedure ValideUnDetailTS(TOBPiece,TOBENT : TOB);
  var TOBD : TOB;
      II : Integer;
  begin
    for II := 0 to TOBENT.Detail.count - 1 do
    begin
      TOBD := TOBENT.detail[II];
      TOBD.SetString('BLT_NATUREPIECEG',TOBPiece.GetString('GP_NATUREPIECEG'));
      TOBD.SetString('BLT_SOUCHE',TOBPiece.GetString('GP_SOUCHE'));
      TOBD.SetInteger('BLT_NUMERO',TOBPiece.GetInteger('GP_NUMERO'));
      TOBD.SetInteger('BLT_INDICEG',TOBPiece.GetInteger('GP_INDICEG'));
      TOBD.SetString('BLT_REFERENCETS',TOBENT.GetString('BLE_REFERENCETS'));
    end;
  end;

  procedure ValideLesTS(TOBPiece,TOBLIG : TOB);
  var TOBE : TOB;
      II : Integer;
  begin
    for II := 0 to TOBLIG.Detail.count - 1 do
    begin
      TOBE := TOBLIG.detail[II];
      TOBE.SetString('BLE_NATUREPIECEG',TOBPiece.GetString('GP_NATUREPIECEG'));
      TOBE.SetString('BLE_SOUCHE',TOBPiece.GetString('GP_SOUCHE'));
      TOBE.SetInteger('BLE_NUMERO',TOBPiece.GetInteger('GP_NUMERO'));
      TOBE.SetInteger('BLE_INDICEG',TOBPiece.GetInteger('GP_INDICEG'));
      ValideUnDetailTS(TOBPiece,TOBE);
    end;
  end;

var II : Integer;
    TOBLIG: TOB;
    cledoc : R_CLEDOC;
begin
  if (TOBpiece.GetString('GP_NATUREPIECEG') <> 'BCE') or (VH_GC.BTCODESPECIF <> '001') then Exit;
  for II := 0 to TOBTSPOC.Detail.count - 1 do
  begin
    TOBLIG := TOBTSPOC.detail[II];
    ValideLesTS(TOBPiece,TOBLIG);
  end;
  cledoc := TOB2CleDoc(TOBPiece);
  DeleteLesTOBTS (cledoc);
  TOBTSPOC.InsertDbByNivel(false); 
end;

procedure LoadLesTOBTS (Cledoc : R_CLEDOC;TOBTSPOC : TOB);
var QQ : TQuery;
    SQL : String;
    TOBDET,TOBT,TOBE,TOBED,TOBEP,TOBDT : TOB;
    LastRef,TheRef : string;
    LastFather : TOB;
begin
  if (Cledoc.NaturePiece <> 'BCE') or (VH_GC.BTCODESPECIF <> '001') then Exit;
  LastFather := nil;
  LastRef := '';
  TOBDET := TOB.Create ('LES DETAILS',nil,-1);
  TOBE := TOB.Create ('LES Ent',nil,-1);
  //
  SQL := 'SELECT * FROM BLIGNEETS WHERE '+WherePiece(Cledoc,ttdTSPOC,false);
  QQ := OpenSQL(SQL,true,-1,'',true);
  if not QQ.eof then TOBE.LoadDetailDB('BLIGNEETS','','',QQ,false);
  ferme (QQ);
  if TOBE.detail.count > 0 then
  begin
    repeat
      TOBED := TOBE.detail[0];
      TheRef := TOBED.GetString('BLE_NUMORDRE');
      TOBEP := TOBTSPOC.findfirst(['NUMORDRE'],[TheRef],True);
      if TOBEP = nil then
      begin
        TOBEP := CreTOBTS (TheRef,TOBTSPOC);
      end;
      TOBED.ChangeParent(TOBEP,-1);
      TOBEP.SetDouble('TOTALTS',TOBEP.GetDouble('TOTALTS')+TOBED.GetDouble('BLE_MONTANT'));
    until TOBE.Detail.count =0;
    //
    SQL := 'SELECT *,"-" AS NEW FROM BLIGNETS WHERE '+WherePiece(Cledoc,ttdTSDetPOC,false);
    QQ := OpenSQL(SQL,true,-1,'',true);
    if not QQ.eof then TOBDET.LoadDetailDB('BLIGNETS','','',QQ,false);
    ferme (QQ);
    if not TOBDET.detail.Count <> 0 then
    begin
      repeat
        TOBT := TOBDET.detail[0];
        TheRef := TOBT.GetString('BLT_NUMORDRE');
        TOBEP := TOBTSPOC.findfirst(['NUMORDRE'],[TheRef],True);
        if TOBEP = nil then BEGIN TOBT.Free; Continue; end;  // pas de ligne associée ????
        TOBDT := TOBEP.FindFirst(['BLE_REFERENCETS'],[TOBT.GetString('BLT_REFERENCETS')],true);
        if TOBDT = nil then BEGIN TOBT.free; continue; end;// pas de ref ts associé ??
        TOBT.ChangeParent(TOBDT,-1);
      until TOBDET.detail.Count = 0;
    end;
  end;
  TOBDET.Free;
  TOBE.free;
end;

procedure GestionTsPOC(TOBPiece,OneTOB,TOBTSPOC : TOB);
var TheTOBC : TOB;
    TheNumOrdre : String;
    II : integer;
begin
  TheNumOrdre := OneTOB.GetString('GL_NUMORDRE');
  TheTOBC := TOBTSPOC.findfirst(['NUMORDRE'],[TheNumordre],True);
  if TheTOBC = nil then
  begin
    TheTOBC := CreTOBTS(TheNumOrdre,TOBTSPOC);
  end;
  TheTOBC.SetString('DESIGNATION','Ligne '+OneTOB.GetString('GL_NUMLIGNE')+' : '+OneTOB.GetString('GL_CODEARTICLE')+copy(OneTOB.GetString('GL_LIBELLE'),1,35));
  TheTOBC.data := TOBPiece;
  TheTOB := TheTOBC;
  AGLLanceFiche('BTP','BTMULPOCTS','','','ACTION=MODIFICATION');
  TheTOB := nil;
  TheTOBC.SetDouble('TOTALTS',0);
  TheTOBC.Data := nil;
  for II := 0 to TheTOBC.detail.count -1 do
  begin
    TheTOBC.SetDouble('TOTALTS',TheTOBC.GetDouble('TOTALTS')+TheTOBC.detail[II].getDouble('BLE_MONTANT'));
  end;
  OneTOB.SetDouble('SUMTOTALTS',TheTOBC.GetDouble('TOTALTS'));
  if TheTOBC.Detail.count = 0 then TheTOBC.Free;
end;

function SetFactureReglePOC (TOBPiece : TOB) : boolean;
var TheIDZeDoc : string;
begin
  result := True;
  if TOBPiece.getString('GP_NATUREPIECEG')<>'FF' then exit;
  TheIDZeDoc := GetIDBSVDOC (TOBPiece);
  if TheIDZeDoc = '' then Exit;
  result :=  SetFactureRegleBSV (TheIDZeDoc);
end;

Initialization
 TOBlesContratsST := TOB.create ('LES CONTRATS ST',nil,-1);

 RegisterAglFunc('AppelMarcheST',False,4,AGLAppelMarcheST);
 RegisterAglFunc('AppelCreationMarcheST',False,1,AGLAppelCreationMarcheST);

finalization
 TOBlesContratsST.free;

end.
