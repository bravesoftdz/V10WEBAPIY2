unit factretenues;

interface


uses variants,HEnt1, HCtrls, UTOB, Ent1, LookUp, Controls, ComCtrls, StdCtrls, ExtCtrls, CalcOleGescom,
  factCommBtp,
  forms,
  {$IFDEF EAGLCLIENT}
  HPdfPrev, UtileAGL, Maineagl,
  {$ELSE}
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} FE_Main, DB, EdtEtat, EdtDoc,
  EdtREtat, EdtRDoc,
  {$ENDIF}
  SysUtils, Dialogs, Utiltarif, SaisUtil, UtilPGI, AGLInit, FactUtil, FactTOB, FactArticle,
  FactLotSerie,
  Math, StockUtil, EntGC, Classes, HMsgBox, FactNomen, TiersUtil, FactTiers, FactPiece,
  {$IFDEF CHR}HRReglement, {$ENDIF}
  {$IFDEF GRC}UtilRT,{$ENDIF}
  uRecupSQLModele,
	BTFactImprTOB,
  Echeance, UtilGC, UtilArticle, ParamSoc,uEntCommun,UtilTOBPiece;

procedure GetRetenuesTTC (TOBPorcs : TOB; var RP,RD : double );
procedure ConstitueBasesretenues (TOBPiece,TOBPorcs,TOBBases,TOBBasesRET : TOB; DEV : Rdevise) ;
procedure GetRetenuesCollectif (TOBPorcs : TOB; var RP,RD : double; ZMode : string='');
procedure GetTVARetenuesCol (TOBporcs: TOB; var TXP,TXD : double; Zmode : string ='');
procedure RecalculeRetenues (TOBPiece,TOBporcs,TOBBases: TOB; DEV: RDevise);
procedure ConstitueTVARetenue (TOBPiece,TOBP,TOBBases,TOBBASER : TOB ; DEV : RDevise);
procedure RecalcRetenue (TOBP,TOBBasesRetenues : TOB);
procedure RatioizeTVARetenue (TOBP,TTOBBases,TOBBR : TOB ; DEV : RDevise; Ratio : double);
procedure AddRetenuesHT (TOBpiece,TOBporcs,TOBBases : TOB; DEV : rdevise);
procedure GetTvaRetenuesDivHT (TOBPorcs : TOB; var RP,RD : double; CatTAxe : string);

implementation
uses factcalc;

procedure AddRetenuesHT (TOBpiece,TOBporcs,TOBBases : TOB; DEV : rdevise);
var II,JJ,IMaxBase : integer;
    TOBBASESR,TOBB,TOBBS : TOB;
    TOTBase,Ratio,MaxBase,SumBase,SumTva,MtPorc,MttvaP : double;
begin
  IF TobPorcs = nil then exit;
  TOBBASESR := TOB.Create ('LES BASES',nil,-1);
  //
  TOTBase := 0;
  for II := 0 to TOBBases.detail.count - 1 do
  begin
    TOTBase := TOTBase + TOBBases.detail[II].GetDouble('GPB_BASEDEV');
  end;
  //
  for II := 0 to TOBporcs.detail.count -1 do
  begin
    if TOBPorcs.detail[II].GetBoolean('GPT_RETENUEDIVERSE') then
    begin
      if (Pos(TOBPorcs.detail[II].GetString('GPT_TYPEPORT'),'HT;MI;MT')>0) then
      begin
        TOBBASESR.clearDetail;
        //
        SumBase := 0; SumTva := 0;
        MtPorc := TOBPorcs.detail[II].GetDouble('GPT_TOTALHTDEV');
        MttvaP := TOBPorcs.detail[II].GetDouble('GPT_TOTALTTCDEV') - TOBPorcs.detail[II].GetDouble('GPT_TOTALHTDEV');
        //
        Ratio := TOBPorcs.detail[II].GetDouble('GPT_TOTALHTDEV') / TOTBase;
        //
        for JJ := 0 to TOBBases.detail.count - 1 do
        begin
          TOBBS := TOBBases.detail[JJ];
          TOBB := TOB.Create ('LIGNEBASE',TOBBASESR,-1);
          TOBB.SetString('BLB_CATEGORIETAXE',TOBBS.GetString('GPB_CATEGORIETAXE'));
          TOBB.SetString('BLB_FAMILLETAXE',TOBBS.GetString('GPB_FAMILLETAXE'));
          TOBB.SetDouble('BLB_TAUXTAXE',TOBBS.GetDouble('GPB_TAUXTAXE'));
          TOBB.SetDouble('BLB_TAUXDEV',TOBBS.GetDouble('GPB_TAUXDEV'));
          TOBB.SetDouble('BLB_COTATION',TOBBS.GetDouble('GPB_COTATION'));
          TOBB.SetString('BLB_DEVISE',TOBBS.GetString('GPB_DEVISE'));
          //
          TOBB.SetDouble('BLB_BASEDEV',Arrondi(TOBBS.getDouble('GPB_BASEDEV')*Ratio,DEV.decimale));
          TOBB.SetDouble('BLB_BASETAXE',DeviseToPivotEx(TOBB.GetDouble('BLB_BASEDEV'),TOBB.GetDouble('BLB_TAUXDEV'),DEV.quotite,DEV.decimale));
          TOBB.SetDouble('BLB_VALEURDEV',Arrondi(TOBBS.getDouble('GPB_BASEDEV')*(TOBB.GetDouble('BLB_TAUXTAXE')/100),DEV.decimale));
          TOBB.SetDouble('BLB_VALEURTAXE',DeviseToPivotEx(TOBB.GetDouble('BLB_VALEURDEV'),TOBB.GetDouble('BLB_TAUXDEV'),DEV.quotite,DEV.decimale));
          if TOBB.GetDouble('BLB_BASEDEV') > MaxBase then
          begin
            MaxBase := TOBB.GetDouble('BLB_BASEDEV');
            IMaxBase := JJ;
          end;
          SumBase := SumBase +  TOBB.GetDouble('BLB_BASEDEV');
          SumTva := SumTva + TOBB.GetDouble('BLB_VALEURDEV')
        end;
        if SumBase <> MtPorc then
        begin
          SumTva := SumTva - TOBB.GetDouble('BLB_VALEURDEV');
          TOBB.SetDouble('BLB_BASEDEV',TOBB.getDouble('BLB_BASEDEV') +  (MtPorc - SumBase));
          TOBB.SetDouble('BLB_BASETAXE',DeviseToPivotEx(TOBB.GetDouble('BLB_BASEDEV'),TOBB.GetDouble('BLB_TAUXDEV'),DEV.quotite,DEV.decimale));
          TOBB.SetDouble('BLB_VALEURDEV',Arrondi(TOBBS.getDouble('GPB_BASEDEV')*(TOBB.GetDouble('BLB_TAUXTAXE')/100),DEV.decimale));
          TOBB.SetDouble('BLB_VALEURTAXE',DeviseToPivotEx(TOBB.GetDouble('BLB_VALEURDEV'),TOBB.GetDouble('BLB_TAUXDEV'),DEV.quotite,DEV.decimale));
          SumTva := SumTva + TOBB.GetDouble('BLB_VALEURDEV');
        end;
        if SUmTva <> MttvaP then
        begin
          TOBB.SetDouble('BLB_VALEURDEV',TOBB.GetDouble('BLB_VALEURDEV') + (MttvaP - SumTva));
          TOBB.SetDouble('BLB_VALEURTAXE',DeviseToPivotEx(TOBB.GetDouble('BLB_VALEURDEV'),TOBB.GetDouble('BLB_TAUXDEV'),DEV.quotite,DEV.decimale));
        end;
        // on déduit cette base au pied
        CumuleLesBases(TOBBases,TOBBASESR,Dev.decimale,'-');
        // et on déduit du document
        TOBpiece.SetDouble('GP_TOTALHTDEV',TOBpiece.GetDouble('GP_TOTALHTDEV') - TOBPorcs.detail[II].GetDouble('GPT_TOTALHTDEV')) ;
        TOBpiece.SetDouble('GP_TOTALTTCDEV',TOBpiece.GetDouble('GP_TOTALTTCDEV') - TOBPorcs.detail[II].GetDouble('GPT_TOTALTTCDEV')) ;
        TOBpiece.SetDouble('GP_TOTALHT',TOBpiece.GetDouble('GP_TOTALHT') - TOBPorcs.detail[II].GetDouble('GPT_TOTALHT')) ;
        TOBpiece.SetDouble('GP_TOTALTTC',TOBpiece.GetDouble('GP_TOTALTTC') - TOBPorcs.detail[II].GetDouble('GPT_TOTALTTC')) ;
      end;
    end;
  end;
end;

procedure GetTvaRetenuesDivHT (TOBPorcs : TOB; var RP,RD : double; CatTAxe : string);
var II,Itaxe : integer;
begin
  Rp := 0;
  RD := 0;
  if CatTaxe <> '' then ITaxe := StrToInt(Copy(CatTaxe,3,1)) else ITaxe := 1;
  //FV1 - 12/07/2016 : FS#2092 - DELABOUDINIERE - Module Contrats message violation d'accès lors validation génération de facture
  IF TobPorcs = nil then exit;

  for II := 0 to TOBporcs.detail.count -1 do
  begin
    if TOBPorcs.detail[II].GetBoolean('GPT_RETENUEDIVERSE') then
    begin
      if  (Pos(TOBPorcs.detail[II].GetString('GPT_TYPEPORT'),'HT;MI;MT')>0) then
      begin
        Rp := RP + TOBPorcs.detail[II].getDouble('GPT_TOTALTAXE'+InttOStr(ITAxe));
        RD := RD + TOBPorcs.detail[II].getDouble('GPT_TOTALTAXEDEV'+InttOStr(ITAxe));
      end;
    end;
  end;
end;


procedure GetRetenuesTTC (TOBPorcs : TOB; var RP,RD : double );
var II : integer;
begin
  Rp := 0;
  RD := 0;

  //FV1 - 12/07/2016 : FS#2092 - DELABOUDINIERE - Module Contrats message violation d'accès lors validation génération de facture
  IF TobPorcs = nil then exit;

  for II := 0 to TOBporcs.detail.count -1 do
  begin
    if TOBPorcs.detail[II].GetBoolean('GPT_RETENUEDIVERSE') then continue;
    if (Pos(TOBPorcs.detail[II].GetString('GPT_TYPEPORT'),'PT;MIC;MTC')>0) then
    begin
      Rp := RP + TOBPorcs.detail[II].getDouble('GPT_TOTALTTC');
      RD := RD + TOBPorcs.detail[II].getDouble('GPT_TOTALTTCDEV');
    end;
  end;
end;

procedure GetRetenuesCollectif (TOBPorcs : TOB; var RP,RD : double ; ZMode : string='');
var II : integer;
begin
  Rp := 0;
  RD := 0;

  //FV1 - 12/07/2016 : FS#2092 - DELABOUDINIERE - Module Contrats message violation d'accès lors validation génération de facture
  IF TobPorcs = nil then exit;

  for II := 0 to TOBporcs.detail.count -1 do
  begin
    if TOBPorcs.detail[II].GetBoolean('GPT_RETENUEDIVERSE') then
    begin
      if (ZMode = 'HT') and (Pos(TOBPorcs.detail[II].GetString('GPT_TYPEPORT'),'HT;MI;MT')>0) then
      begin
        Rp := RP + TOBPorcs.detail[II].getDouble('GPT_TOTALHT');
        RD := RD + TOBPorcs.detail[II].getDouble('GPT_TOTALHTDEV');
      end else if ((ZMode = 'TTC') and (Pos(TOBPorcs.detail[II].GetString('GPT_TYPEPORT'),'PT;MIC;MTC')>0)) or (ZMode = '') then
      begin
        Rp := RP + TOBPorcs.detail[II].getDouble('GPT_TOTALTTC');
        RD := RD + TOBPorcs.detail[II].getDouble('GPT_TOTALTTCDEV');
      end;
    end;
  end;
end;

procedure GetTVARetenuesCol (TOBporcs: TOB; var TXP,TXD : double; Zmode : string ='');
var II : integer;
begin
  TXP := 0;
  TXD := 0;
  for II := 0 to TOBporcs.detail.count -1 do
  begin
    if TOBPorcs.detail[II].GetBoolean('GPT_RETENUEDIVERSE') then
    begin
      if ((ZMode = 'HT') and (Pos(TOBPorcs.detail[II].GetString('GPT_TYPEPORT'),'HT;MI;MT')>0)) or (Zmode = '')then
      begin
        TXP := TXP + TOBPorcs.detail[II].getDouble('GPT_TOTALHT');
        TXD := TXD + TOBPorcs.detail[II].getDouble('GPT_TOTALHTDEV');
      end else if ((ZMode = 'TTC') and (Pos(TOBPorcs.detail[II].GetString('GPT_TYPEPORT'),'PT;MIC;MTC')>0)) or (ZMode = '') then
      begin
        TXP := TXP + TOBPorcs.detail[II].getDouble('GPT_TOTALTTC');
        TXD := TXD + TOBPorcs.detail[II].getDouble('GPT_TOTALTTCDEV');
      end;
    end;
  end;
end;

procedure RatioizeTVARetenue (TOBP,TTOBBases,TOBBR : TOB ; DEV : RDevise; Ratio : double);
var II : integer;
    TOBB : TOB;
    Max1,CumHt,MtRef,Diff,CumTTC,RD : double;
    EnHt : boolean;
    M1 : integer;
begin
  EnHt := (Pos(TOBP.GetString('GPT_TYPEPORT'),'HT;MI;MT')>0);
  M1 := -1;
  for II := 0 to TTOBBases.detail.count -1 do
  begin
    TOBB := TOB.create ('PIEDBASE',TOBBR,-1);
    TOBB.Dupliquer(TTOBBases.Detail[II],false,true);
    if EnHt then
    begin
      if TOBB.GetDouble('GPB_BASEDEV') > Max1 then
      begin
        Max1 := TOBB.GetDouble('GPB_BASEDEV');
        M1 := II;
      end;
      TOBB.SetDouble('GPB_BASEDEV',Arrondi(TOBB.GetDouble('GPB_BASEDEV')*ratio,DEV.Decimale));
      TOBB.SetDouble('GPB_BASETAXE',Arrondi(TOBB.GetDouble('GPB_BASETAXE')*ratio,DEV.Decimale));
      TOBB.SetDouble('GPB_VALEURDEV',Arrondi(TOBB.GetDouble('GPB_BASEDEV')*TOBB.GetDouble('GPB_TAUXTAXE')/100,DEV.Decimale));
      TOBB.SetDouble('GPB_VALEURTAXE',Arrondi(TOBB.GetDouble('GPB_BASETAXE')*TOBB.GetDouble('GPB_TAUXTAXE')/100,DEV.Decimale));
      CumHt := CumHt + TOBB.GetDouble('GPB_BASEDEV');
    end else
    begin
      if TOBB.GetDouble('GPB_BASEDEV') > Max1 then
      begin
        Max1 := TOBB.GetDouble('GPB_BASEDEV');
        M1 := II;
      end;
      TOBB.SetDouble('GPB_BASEDEV',Arrondi(TOBB.GetDouble('GPB_BASEDEV')*ratio,DEV.Decimale));
      TOBB.SetDouble('GPB_BASETAXE',Arrondi(TOBB.GetDouble('GPB_BASETAXE')*ratio,DEV.Decimale));
      TOBB.SetDouble('GPB_VALEURDEV',Arrondi(TOBB.GetDouble('GPB_BASEDEV')*TOBB.GetDouble('GPB_TAUXTAXE')/100,DEV.Decimale));
      TOBB.SetDouble('GPB_VALEURTAXE',Arrondi(TOBB.GetDouble('GPB_BASETAXE')*TOBB.GetDouble('GPB_TAUXTAXE')/100,DEV.Decimale));
      CumTTC := CumTTC + TOBB.GetDouble('GPB_BASEDEV') + TOBB.GetDouble('GPB_VALEURTAXE');
    end;
  end;
  //
  if M1 >= 0 then
  begin
    TOBB := TOBBR.detail[M1];
    if EnHt then
    begin
      Diff := TOBP.GetDouble('GPT_TOTALHTDEV') - Arrondi(Cumht,V_PGI.okdecV);
      if diff <> 0 THEN
      begin
        TOBB.SetDouble('GPB_BASEDEV',Arrondi(TOBB.GetDouble('GPB_BASEDEV')- Diff,DEV.Decimale));
        TOBB.SetDouble('GPB_BASETAXE',Arrondi(TOBB.GetDouble('GPB_BASETAXE')-diff,V_PGI.okdecV));
        TOBB.SetDouble('GPB_VALEURDEV',Arrondi(TOBB.GetDouble('GPB_BASEDEV')*TOBB.GetDouble('GPB_TAUXTAXE')/100,DEV.Decimale));
        TOBB.SetDouble('GPB_VALEURTAXE',Arrondi(TOBB.GetDouble('GPB_BASETAXE')*TOBB.GetDouble('GPB_TAUXTAXE')/100,V_PGI.okdecV));
      end;
    end else
    begin
      Diff := TOBP.GetDouble('GPT_TOTALTTCDEV') - Arrondi(CumTTC,V_PGI.okdecV);
      if diff <> 0 THEN
      begin
        CumTTC := CumTTC - (TOBB.GetDouble('GPB_BASETAXE')+TOBB.GetDouble('GPB_VALEURTAXE'));
        RD := Arrondi(Diff / (1+(TOBB.GetDouble('GPB_TAUXTAXE')/100)),V_PGI.okdecV);
        TOBB.SetDouble('GPB_BASEDEV',Arrondi(TOBB.GetDouble('GPB_BASEDEV')- RD,DEV.Decimale));
        TOBB.SetDouble('GPB_BASETAXE',Arrondi(TOBB.GetDouble('GPB_BASETAXE')- RD,V_PGI.okdecV));
        TOBB.SetDouble('GPB_VALEURDEV',Arrondi(TOBB.GetDouble('GPB_BASEDEV')*TOBB.GetDouble('GPB_TAUXTAXE')/100,DEV.Decimale));
        TOBB.SetDouble('GPB_VALEURTAXE',Arrondi(TOBB.GetDouble('GPB_BASETAXE')*TOBB.GetDouble('GPB_TAUXTAXE')/100,V_PGI.okdecV));
        CumTTC := CumTTC + (TOBB.GetDouble('GPB_BASETAXE')+TOBB.GetDouble('GPB_VALEURTAXE'));
        Diff := TOBP.GetDouble('GPT_TOTALTTCDEV') - Arrondi(CumTTC,V_PGI.okdecV);
        if Diff <> 0 then
        begin
          TOBB.SetDouble('GPB_BASEDEV',Arrondi(TOBB.GetDouble('GPB_BASEDEV')- Diff,DEV.Decimale));
          TOBB.SetDouble('GPB_BASETAXE',Arrondi(TOBB.GetDouble('GPB_BASETAXE')-diff,V_PGI.okdecV));
        end;
      end;
    end;
  end;
end;


procedure ConstitueTVARetenue (TOBPiece,TOBP,TOBBases,TOBBASER : TOB ; DEV : RDevise);
var Ratio : double;
    EnHt : boolean;
    TOBBR : TOB;
begin
  if TOBBases = nil then exit;
  EnHt := (Pos(TOBP.GetString('GPT_TYPEPORT'),'HT;MI;MT')>0);
  if EnHt then
  begin
    if TOBPiece.geTDouble('GP_TOTALHTDEV')<> 0 then Ratio := TOBP.GetDouble('GPT_TOTALHTDEV') / TOBPiece.geTDouble('GP_TOTALHTDEV') else Ratio := 1;  // % du montant global
  end else
  begin
    if TOBPiece.geTDouble('GP_TOTALTTCDEV')<> 0 then Ratio := TOBP.GetDouble('GPT_TOTALTTCDEV') / TOBPiece.geTDouble('GP_TOTALTTCDEV') else Ratio := 1;
  end;
  TOBBR := TOB.Create ('UNE RETENUE',TOBBASER,-1);
  TOBBR.AddChampSupValeur('RETENUE',TOBP.getString('GPT_CODEPORT'));
  TOBBR.AddChampSupValeur('NUMPORT',TOBP.getString('GPT_NUMPORT'));
  RatioizeTVARetenue(TOBP,TOBBases,TOBBR,DEV,ratio);
end;


procedure ConstitueBasesretenues (TOBPiece,TOBPorcs,TOBBases,TOBBasesRET : TOB; DEV : Rdevise) ;
var II : integer;
begin
  For II := 0 to TOBPorcs.detail.count -1 do
  begin
    if TOBPorcs.detail[II].GetBoolean('GPT_RETENUEDIVERSE') then
    begin
      ConstitueTVARetenue (TOBPiece,TOBporcs.detail[II],TOBbases,TOBBasesRET,DEV);
    end;
  end;
end;

procedure RecalcRetenue (TOBP,TOBBasesRetenues : TOB);
var TOBB : TOB;
    Ii : integer;
    Xp,XD : double;
begin
  Xp := 0; XD := 0;
  //
  TOBB := TOBBasesRetenues.findFirst(['RETENUE','NUMPORT'],[TOBP.GetString('GPT_CODEPORT'),TOBP.GetInteger('GPT_NUMPORT')],false);
  if TOBB <> nil then
  begin
    for II := 0 to TOBB.detail.count -1 do
    begin
      XP := XP + TOBB.detail[II].GetDouble('GPB_VALEURTAXE');
      XD := XD + TOBB.detail[II].GetDouble('GPB_VALEURDEV');
    end;
  end;
  TOBP.SetDouble('GPT_TOTALTAXEDEV1',XD);
  TOBP.SetDouble('GPT_TOTALTAXE1',XP);
  // et enfin reajustement
  if Pos(TOBP.GetString('GPT_TYPEPORT'),'HT;MI;MT')> 0 then
  begin
    TOBP.SetDouble('GPT_TOTALTTC',TOBP.GetDouble('GPT_TOTALHT')+TOBP.GetDouble('GPT_TOTALTAXE1'));
    TOBP.SetDouble('GPT_TOTALTTCDEV',TOBP.GetDouble('GPT_TOTALHTDEV')+TOBP.GetDouble('GPT_TOTALTAXEDEV1'));
  end else
  begin
    TOBP.SetDouble('GPT_TOTALHT',TOBP.GetDouble('GPT_TOTALTTC')-TOBP.GetDouble('GPT_TOTALTAXE1'));
    TOBP.SetDouble('GPT_TOTALHTDEV',TOBP.GetDouble('GPT_TOTALTTCDEV')-TOBP.GetDouble('GPT_TOTALTAXEDEV1'));
  end;

end;

procedure RecalculeRetenues (TOBPiece,TOBporcs,TOBBases: TOB; DEV: RDevise);
var II : integer;
    TOBBASER : TOB;
begin
  TOBBASER := TOB.Create ('LES BASES',nil,-1);
  TRY
    ConstitueBasesretenues (TOBPiece,TOBPorcs,TOBBases,TOBBASER,DEV) ;
    For II := 0 to TOBPorcs.detail.count -1 do
    begin
      if TOBPorcs.detail[II].GetBoolean('GPT_RETENUEDIVERSE') then
      begin
        RecalcRetenue (TOBPorcs.detail[II],TOBBASER);
      end;
    end;
  FINALLY
    TOBBASER.free;
  END;
end;

end.
