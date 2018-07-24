unit UCumulCollectifs;

interface

uses hctrls,
  sysutils,HEnt1,
{$IFNDEF EAGLCLIENT}
  uDbxDataSet, DB,
{$ELSE}
  uWaini,
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} DB,
{$ENDIF}
  paramsoc,
  utob,
  EntGC, Ent1,
  uEntCommun,
  UtilsTOB;

procedure CumuleCollectifs(TOBL,TOBTaxesL,TOBVTECOLL,TOBSSTRAIT,TOBARTICLES,TOBTiers,TOBAFFAIRE : TOB);
procedure ReajusteCollectifs ( TOBVTECOLL : TOB; MontantEchEnt,MontantEchEntDev : Double ) ;
procedure PrepareInsertCollectif (TOBPiece,TOBVTECOLL : TOB);
procedure ConstitueVteCollectif (TOBPiece,TOBSSTRAIT,TOBBases,TOBVTECOLL : TOB);
function FindCollectifPlusPlus(TOBARTICLES,TOBTiers,TOBaffaire,TOBL : TOB) : string;
function FindCollectifRDPlus(TOBP,TOBPiece : TOB) : string;
function FindCollectifRGPlus(TOBBRG,TOBPiece : TOB) : string;


implementation

uses FactComm,UCotraitance;

function GetSqlPlus (Nature,CodePort,FamArt,FamTiers,FamAff,Etabl,Regime,FamTaxe : string) : string;
Var wCodeP,wArt,wTiers,wAff,WVenteAchat : String ;
BEGIN

  if CodePort = '' then
  begin
    wCodeP:='AND (BVC_CODEPORT="")';
  end else
  begin
    wCodeP:='AND (BVC_CODEPORT="'+CodePOrt+'" OR BVC_CODEPORT="")';
  end;

  if (GetParamSocSecur('SO_GCVENTCPTAART', False)) then
  begin
    if Famart = '' then
    begin
      wArt:='AND (BVC_COMPTAARTICLE="")';
    end else
    begin
      wArt:='AND (BVC_COMPTAARTICLE="'+FamArt+'" OR BVC_COMPTAARTICLE="")';
    end;
  end else
  begin
    wArt:='' ;
  end;
  if (GetParamSocSecur('SO_GCVENTCPTATIERS', False)) then
  begin
    if Famtiers = '' then
    begin
      wTiers:='AND (BVC_COMPTATIERS="")';
    end else
    begin
      wTiers:='AND (BVC_COMPTATIERS="'+FamTiers+'" OR BVC_COMPTATIERS="")';
    end;
  end else
  begin
    wTiers:='' ;
  end;
  if (GetParamSocSecur('SO_GCVENTCPTAAFF', False)) then
  begin
    if FamAFF = '' then
    begin
      wAff:='AND (BVC_COMPTAAFFAIRE="")';
    end else
    begin
      wAff:='AND (BVC_COMPTAAFFAIRE="'+FamAff+'" OR BVC_COMPTAAFFAIRE="")';
    end;
  end else
  begin
    wAff:='' ;
  end;
  Result:='SELECT * FROM BVENTILCOLL WHERE '+
          '(BVC_NATUREV="'+Nature+'") AND '+
          '(BVC_ETABLISSEMENT="'+Etabl+'" OR BVC_ETABLISSEMENT="") AND '+
          '(BVC_REGIMETAXE="'+Regime+'" OR BVC_REGIMETAXE="") AND '+
          '(BVC_FAMILLETAXE="'+FamTaxe+'" OR BVC_FAMILLETAXE="") '+
          wCodeP+' '+wArt+' '+wTiers+' '+wAff+' '+WVenteAchat+' '+
         'ORDER BY BVC_COMPTAARTICLE DESC, BVC_COMPTATIERS DESC, BVC_COMPTAAFFAIRE DESC, BVC_REGIMETAXE DESC, BVC_FAMILLETAXE DESC, BVC_ETABLISSEMENT DESC' ;
end;

Function FindTOBCodePlus ( TOBDatas : TOB; CodePort,FamArt,FamTiers,FamAff,Etabl,Regime,FamTaxe : string  ) : TOB ;
Var TOBC : TOB ;
    i : integer ;
    fArt,fTiers,fAff : string;
BEGIN
  Result:=Nil ;
  for i:=TOBDatas.Detail.Count-1 downto 0 do
  BEGIN
    TOBC:=TOBDatas.Detail[i] ;
    if ((GetParamSocSecur('SO_GCVENTCPTAART', False)) and (TOBC.GetValue('BVC_COMPTAARTICLE')<>'') and (TOBC.GetValue('BVC_COMPTAARTICLE')<>FamArt)) then Continue ;
    if ((GetParamSocSecur('SO_GCVENTCPTATIERS', False)) and (TOBC.GetValue('BVC_COMPTATIERS')<>'') and (TOBC.GetValue('BVC_COMPTATIERS')<>FamTiers)) then Continue ;
    if ((GetParamSocSecur('SO_GCVENTCPTAAFF', False)) and (TOBC.GetValue('BVC_COMPTAAFFAIRE')<>'') and (TOBC.GetValue('BVC_COMPTAAFFAIRE')<>FamAff)) then Continue ;
    if ((TOBC.GetValue('BVC_CODEPORT')<>'') and (TOBC.GetValue('BVC_CODEPORT')<>CodePort)) then Continue ;
    if ((TOBC.GetValue('BVC_REGIMETAXE')<>'') and (TOBC.GetValue('BVC_REGIMETAXE')<>Regime)) then Continue ;
    if ((TOBC.GetValue('BVC_FAMILLETAXE')<>'') and (TOBC.GetValue('BVC_FAMILLETAXE')<>FamTaxe)) then Continue ;
    if ((TOBC.GetValue('BVC_ETABLISSEMENT')<>'') and (TOBC.GetValue('BVC_ETABLISSEMENT')<>Etabl)) then Continue ;
    Result:=TOBC ; Break ;
  END ;
END ;


function FindCollectifRDPlus(TOBP,TOBPiece : TOB) : string;
var FamArt,FamTiers,FamAff,Etabl,Regime,FamTaxe,CodePort,SQL : string;
    TOBDatas,TOBF : TOB;
begin
  Result := '';
  FamArt := '';
  FamTiers := '';
  FamAff := '';
  Etabl := TOBPiece.GetString('GP_ETABLISSEMENT');
  Regime := TOBPiece.GetString('GP_REGIMETAXE');
  FamTaxe := TOBP.getString('GPT_FAMILLETAXE1');
  CodePort := TOBP.getString('GPT_CODEPORT');
  SQL := GetSqlPlus ('003',CodePort,FamArt,FamTiers,FamAff,Etabl,Regime,FamTaxe);
  TOBDatas := TOB.Create ('LES DATAS',nil,-1);
  TOBDatas.LoadDetailDBFromSQL('BVENTILCOLL',SQL,false);
  TOBF := FindTOBCodePlus (TOBDatas,CodePort,FamArt,FamTiers,FamAff,Etabl,Regime,FamTaxe);
  if TOBF <> nil then
  begin
    if (TOBPiece.GetString('GP_VENTEACHAT')='VEN') then Result := TOBF.GetString('BVC_COLLECTIF')
                                                   else Result := TOBF.GetString('BVC_COLLECTIFAC');
  end;
  TOBDatas.free;
end;

function FindCollectifRGPlus(TOBBRG,TOBPiece : TOB) : string;
var FamArt,FamTiers,FamAff,Etabl,Regime,FamTaxe,SQL : string;
    TOBDatas,TOBF : TOB;
begin
  Result := '';
  FamArt := '';
  FamTiers := '';
  FamAff := '';
  Etabl := TOBPiece.GetString('GP_ETABLISSEMENT');
  Regime := TOBPiece.GetString('GP_REGIMETAXE');
  if TOBBRG.Detail.count = 0 then Exit;
  FamTaxe := TOBBRG.detail[0].getString('PBR_FAMILLETAXE');
  SQL := GetSqlPlus ('002','',FamArt,FamTiers,FamAff,Etabl,Regime,FamTaxe);
  TOBDatas := TOB.Create ('LES DATAS',nil,-1);
  TOBDatas.LoadDetailDBFromSQL('BVENTILCOLL',SQL,false);
  TOBF := FindTOBCodePlus (TOBDatas,'',FamArt,FamTiers,FamAff,Etabl,Regime,FamTaxe);
  if TOBF <> nil then
  begin
    if (TOBPiece.GetString('GP_VENTEACHAT')='VEN') then Result := TOBF.GetString('BVC_COLLECTIF')
                                                   else Result := TOBF.GetString('BVC_COLLECTIFAC');
  end;
  TOBDatas.free;

end;

function FindCollectifPlusPlus(TOBARTICLES,TOBTiers,TOBaffaire,TOBL : TOB) : string;
var TOBA,TOBDatas,TOBF,TOBPIECE : TOB;
    SQL : string;
    FamArt,FamTiers,FamAff,Etabl,Regime,FamTaxe : string;
begin
  result := '';
  if TOBL.GetString('GL_ARTICLE')='' then exit;
  TobA := TOBArticles.FindFirst(['GA_ARTICLE'], [TOBL.GetString('GL_ARTICLE')], False);
  if TOBA = nil then Exit;
  //
  FamArt := TOBA.GetString('GA_COMPTAARTICLE');
  Famtiers := TOBTIers.GetString('T_COMPTATIERS');
  FamAff := TOBAffaire.GetString('AFF_COMPTAAFFAIRE');
  Etabl := TOBL.GetString('GL_ETABLISSEMENT');
  FamTaxe := TOBL.GetString('GL_FAMILLETAXE1');
  Regime := TOBL.GetString('GL_REGIMETAXE');
  //
  SQL := GetSqlPlus ('001','',FamArt,FamTiers,FamAff,Etabl,Regime,FamTaxe);
  TOBPIECE := TOBL.Parent;
  TOBDatas := TOB.Create ('LES DATAS',nil,-1);
  TOBDatas.LoadDetailDBFromSQL('BVENTILCOLL',SQL,false);
  TOBF := FindTOBCodePlus (TOBDatas,'',FamArt,FamTiers,FamAff,Etabl,Regime,FamTaxe);
  if TOBF <> nil then
  begin
    if (TOBPiece.GetString('GP_VENTEACHAT')='VEN') then Result := TOBF.GetString('BVC_COLLECTIF')
                                                   else Result := TOBF.GetString('BVC_COLLECTIFAC');
  end;
  TOBDatas.free;
end;


procedure PrepareInsertCollectif (TOBPiece,TOBVTECOLL : TOB);
var II : Integer;
    TOBL : TOB;
begin
  For II := 0 to TOBVTECOLL.detail.count -1 do
  begin
    TOBL := TOBVTECOLL.detail[II];
    TOBL.SetString('BPB_NATUREPIECEG',TOBPiece.GetString('GP_NATUREPIECEG'));
    TOBL.SetString('BPB_SOUCHE',TOBPiece.GetString('GP_SOUCHE'));
    TOBL.SetInteger('BPB_NUMERO',TOBPiece.GetInteger('GP_NUMERO'));
    TOBL.SetInteger('BPB_INDICEG',TOBPiece.GetInteger('GP_INDICEG'));
    TOBL.SetAllModifie(true);
  end;
end;

procedure ConstitueVteCollectif (TOBPiece,TOBSSTRAIT,TOBBases,TOBVTECOLL : TOB);
var II : Integer;
    TOBB : TOB;
    TOBT,TOBV : TOB;
    FamilleTaxe,RegimeTaxe : string;
    COLLECTIF : string;
begin

  for II := 0 TO TOBBases.detail.count -1 do
  begin
    TOBB := TOBBases.detail[II];
    COLLECTIF := '';
    FamilleTaxe:=TOBB.GetValue('GPB_FAMILLETAXE');
    RegimeTaxe:=TOBPiece.GetValue('GP_REGIMETAXE') ;
    // on ne prends pas en compte les paiements direct
    if (TOBB.GetString('GPB_FOURN')<>'') and (GetPaiementSSTrait (TOBSSTRAIT,TOBB.GetString('GPB_FOURN'))='001') then continue;
    // -----
    TOBT:=VH^.LaTOBTVA.FindFirst(['TV_TVAOUTPF','TV_REGIME','TV_CODETAUX'],['TX1',RegimeTaxe,FamilleTaxe],False) ;
    if TOBT<>Nil then
    BEGIN
      COLLECTIF:= TOBT.GetString('TV_COLLECTIF');
    END;
    TOBV := TOBVTECOLL.FindFirst(['BPB_COLLECTIF'],[COLLECTIF],true);
    if TOBV = nil then
    begin
      TOBV := TOB.Create ('PIEDCOLLECTIF',TOBVTECOLL,-1);
      TOBV.SetString('BPB_NATUREPIECEG',TOBpiece.GetValue('GP_NATUREPIECEG'));
      TOBV.SetString('BPB_SOUCHE',TOBpiece.getString('GP_SOUCHE'));
      TOBV.SetInteger('BPB_NUMERO',TOBpiece.GetInteger('GP_NUMERO'));
      TOBV.SetInteger('BPB_INDICEG',TOBpiece.GetInteger('GP_INDICEG'));
      TOBV.SetString('BPB_COLLECTIF',COLLECTIF);
    end;
    TOBV.SetDouble('BPB_BASETTC',TOBV.GetDouble('BPB_BASETTC')+TOBB.GetDouble('GPB_BASETAXE')+TOBB.GetDouble('GPB_VALEURTAXE'));
    TOBV.SetDouble('BPB_BASETTCDEV',TOBV.GetDouble('BPB_BASETTCDEV')+TOBB.GetDouble('GPB_BASEDEV')+TOBB.GetDouble('GPB_VALEURDEV'));
  end;
end;


procedure CumuleCollectifs(TOBL,TOBTaxesL,TOBVTECOLL,TOBSSTRAIT,TOBARTICLES,TOBTiers,TOBAFFAIRE : TOB);
var TOBT,TOBV : TOB;
    RegimeTaxe,FamilleTaxe,Prefixe,COLLECTIF : string;
    II : Integer;
begin
  COLLECTIF := '';
  prefixe := GetPrefixeTable (TOBL);
  FamilleTaxe:=TOBL.GetValue(prefixe+'_FAMILLETAXE1');
  RegimeTaxe:=TOBL.GetValue(prefixe+'_REGIMETAXE') ;
  // on ne prends pas en compte les paiements direct
  if (TOBL.GetString('GL_FOURNISSEUR')<>'') and (GetPaiementSSTrait (TOBSSTRAIT,TOBL.GetString('GL_FOURNISSEUR'))='001') then exit;
  // -----
  Collectif := FindCollectifPlusPlus(TOBARTICLES,TOBTiers,TOBaffaire,TOBL);
  if Collectif = '' then
  begin
    TOBT:=VH^.LaTOBTVA.FindFirst(['TV_TVAOUTPF','TV_REGIME','TV_CODETAUX'],['TX1',RegimeTaxe,FamilleTaxe],False) ;
    if TOBT<>Nil then
    BEGIN
      COLLECTIF:= TOBT.GetString('TV_COLLECTIF');
    END;
  end;
  TOBV := TOBVTECOLL.FindFirst(['BPB_COLLECTIF'],[COLLECTIF],true);
  if TOBV = nil then
  begin
    TOBV := TOB.Create ('PIEDCOLLECTIF',TOBVTECOLL,-1);
    TOBV.SetString('BPB_NATUREPIECEG',TOBL.GetValue(prefixe+'_NATUREPIECEG'));
    TOBV.SetString('BPB_SOUCHE',TOBL.GetValue(prefixe+'_SOUCHE'));
    TOBV.SetInteger('BPB_NUMERO',TOBL.GetValue(prefixe+'_NUMERO'));
    TOBV.SetInteger('BPB_INDICEG',TOBL.GetValue(prefixe+'_INDICEG'));
    TOBV.SetString('BPB_COLLECTIF',COLLECTIF);
  end;
  for II := 0 to TOBTAXESL.detail.count -1 do
  begin
    TOBV.SetDouble('BPB_BASETTC',TOBV.GetDouble('BPB_BASETTC')+TOBTaxesL.detail[II].GetDouble('BLB_BASETAXE')+TOBTaxesL.detail[II].GetDouble('BLB_VALEURTAXE'));
    TOBV.SetDouble('BPB_BASETTCDEV',TOBV.GetDouble('BPB_BASETTCDEV')+TOBTaxesL.detail[II].GetDouble('BLB_BASEDEV')+TOBTaxesL.detail[II].GetDouble('BLB_VALEURDEV'));
  end;
end;

procedure ReajusteCollectifs ( TOBVTECOLL : TOB; MontantEchEnt,MontantEchEntDev : Double );
var MtGlobal,MtGlobalDEV : Double;
    II,Max : Integer;
    TOBL : TOB;
    Coef,MtMax,Diff : double;
begin
  MtGlobal := 0;
  MtGlobalDEV := 0;
  Max := -1;
  for II := 0 to TOBVTECOLL.Detail.Count -1 do
  begin
    MtGlobal := MtGlobal + TOBVTECOLL.Detail[II].GetDouble('BPB_BASETTC');
    TOBVTECOLL.Detail[II].SetDouble('BPB_REAJUSTE',TOBVTECOLL.Detail[II].GetDouble('BPB_BASETTC'));
    TOBVTECOLL.Detail[II].SetDouble('BPB_REAJUSTEDEV',TOBVTECOLL.Detail[II].GetDouble('BPB_BASETTCDEV'));
    if TOBVTECOLL.Detail[II].GetDouble('BPB_BASETTC') > MtMax then
    begin
      MtMax := TOBVTECOLL.Detail[II].GetDouble('BPB_BASETTC');
      Max := II;
    end;
  end;
  if MontantEchEnt <> MtGlobal then
  begin
    Coef :=  MtGlobal / MontantEchEnt;
    MtGlobal := 0;
    MtGlobalDev := 0;
    for II := 0 to TOBVTECOLL.Detail.Count -1 do
    begin
      TOBVTECOLL.Detail[II].SetDouble('BPB_REAJUSTE',ARRONDI(TOBVTECOLL.Detail[II].GetDouble('BPB_BASETTC')*Coef,V_PGI.OkDecV));
      TOBVTECOLL.Detail[II].SetDouble('BPB_REAJUSTEDEV',ARRONDI(TOBVTECOLL.Detail[II].GetDouble('BPB_BASETTCDEV')*Coef,V_PGI.OkDecV));
      MtGlobal := MtGlobal + TOBVTECOLL.Detail[II].GetDouble('BPB_REAJUSTE');
      MtGlobalDEV := MtGlobalDEV + TOBVTECOLL.Detail[II].GetDouble('BPB_REAJUSTEDEV');
    end;
    Diff := ARRONDI(MontantEchEnt - MtGlobal,V_PGI.OkDecV);
    if (diff <> 0) and (Max <> -1) then
    begin
      TOBVTECOLL.Detail[MAX].SetDouble('BPB_REAJUSTE',TOBVTECOLL.Detail[Max].GetDouble('BPB_REAJUSTE')+ diff);
    end;
    Diff := ARRONDI(MontantEchEntDEV - MtGlobalDev,V_PGI.OkDecV);
    if (diff <> 0) and (Max <> -1) then
    begin
      TOBVTECOLL.Detail[MAX].SetDouble('BPB_REAJUSTEDEV',TOBVTECOLL.Detail[Max].GetDouble('BPB_REAJUSTEDEV')+ diff);
    end;
  end;
end;

end.
