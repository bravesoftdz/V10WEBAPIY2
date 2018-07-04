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

procedure CumuleCollectifs(TOBL,TOBTaxesL,TOBVTECOLL,TOBSSTRAIT : TOB);
procedure ReajusteCollectifs ( TOBVTECOLL : TOB; MontantEchEnt,MontantEchEntDev : Double ) ;
procedure PrepareInsertCollectif (TOBPiece,TOBVTECOLL : TOB);
procedure ConstitueVteCollectif (TOBPiece,TOBSSTRAIT,TOBBases,TOBVTECOLL : TOB);

implementation

uses FactComm,UCotraitance;

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


procedure CumuleCollectifs(TOBL,TOBTaxesL,TOBVTECOLL,TOBSSTRAIT : TOB);
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
  TOBT:=VH^.LaTOBTVA.FindFirst(['TV_TVAOUTPF','TV_REGIME','TV_CODETAUX'],['TX1',RegimeTaxe,FamilleTaxe],False) ;
  if TOBT<>Nil then
  BEGIN
    COLLECTIF:= TOBT.GetString('TV_COLLECTIF');
  END;
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
