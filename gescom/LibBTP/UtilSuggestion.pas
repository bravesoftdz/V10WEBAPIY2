unit UtilSuggestion;

interface
uses SysUtils, Classes, UTOB, Hctrls, HStatus, StdCtrls, math,
     EntGC, Ent1, AglInit,NomenUtil,UtilArticle,UTofListeInv,
{$IFDEF EAGLCLIENT}
     MaineAGL,
{$ELSE}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}FE_Main,
{$ENDIF}
     Paramsoc,
     HMsgBox,
     HEnt1;

type
TTypeUniteResult = (TurStock,TurVente,TurAchat);

TGinfostarifFour = record
	TarifAch : double;
	TarifAchBrut : double;
  Remise : string;
  remisesreelle   : double;
  cascade : string;
  UniteAchat : string;
  CoefUAUs : double;
  LibelleFournisseur : string;
  Tarif : integer;
	end;

function RecupTarifAch (Fournisseur,Article : string; VAR UniteAch : string; var CoefuaUs : double; UniteResult : TTypeUniteResult = TurAchat;
                        PrixNet : boolean=true;PrxArtInclus:boolean=true;TOBA : TOB=nil;QteRef : double=1; DateRef : TdateTime=0) : double; overload;
function RecupTarifAch (TOBArt,TOBTarif,TobTiers : TOB; Naturepiece : String; var UniteAch : string ; var CoefuaUs : double; UniteResult : TTypeUniteResult = TurAchat;
                        PrixNet : boolean=true;PrxArtInclus:boolean=true;QteRef : double=1; Dateref : TdateTime=0) : double; overload;
function SelectionDocuments (LesNaturesAPrendre : string;DateDebut,DateFin: string; TOBListeDocuments : TOB) : boolean;
function GetInfosTarifAch (Fournisseur,Article : string; UniteResult : TTypeUniteResult = TurAchat;
													 PrixNet : boolean=true;PrxArtInclus:boolean=true;TOBA : TOB=nil;QteRef : double=1;
                           DateRef : TdateTime=0) : TGinfostarifFour;

//fv1 : 22/05/2014 - FS#1008 - DELABOUDINIERE : Ajout contrôle référencement fournisseur en saisie de pièces d'achats
Function RechTarifArticle(TOBA, TOBTarif, TOBTiers, TOBPIECE, TOBL : TOB) : Double;
Function RechPrixArticle (TOBA, TobTarif, Tobtiers, TobPiece, TOBL : TOB; VenteAchat : String) : Double;

implementation
uses UTOF_VideInside;


function SelectionDocuments (LesNaturesAPrendre : string;DateDebut,DateFin: string; TOBListeDocuments : TOB) : boolean;
begin
  result := false;
  TRY
    TheToB := TOBListeDocuments;
//    SaisirHorsInside('GCSELDOCREA_MUL');
    AGLLanceFiche('GC','GCSELDOCREA_MUL','','','ACTION=MODIFICATION');
    if TheTob <> nil then
    begin
      TOBListeDocuments.clearDetail;
      TOBListeDocuments.Dupliquer (TheTob,true,true);
      TheTOB.free;
      result := true;
    end;
  FINALLY
    TheTob := nil;
  END;
end;

//fv1 : 22/05/2014 - FS#1008 - DELABOUDINIERE : Ajout contrôle référencement fournisseur en saisie de pièces d'achats
Function RechTarifArticle(TOBA, TOBTarif, TOBTiers, TOBPIECE, TOBL : TOB) : Double;
Var Px        : Double;
    PxAch     : Double;
    CoefUaUS  : Double;
    UA        : String;
begin

  UA := '';
  Result := 0;

  if TOBPiece.GetValue('GP_VENTEACHAT') = 'ACH' then
  begin
    PxAch := RecupTarifAch(TOBA, TOBTarif, TOBTiers,TOBL.GetValue('GL_NATUREPIECEG'),Ua,CoefUaUs, TurAchat, false,true,TOBL.GetValue('GL_QTEFACT'),TOBL.GetValue('GL_DATEPIECE'));
    if PxAch <> 0 then
    Begin
      Px := PxAch;
      TOBL.PutValue('FROMTARIF','X');
      TOBL.PutValue('GL_DPA',Px);
      //FV1 : On gère des prix d'achat et on charge des unité de vente (????)
      //if Ua <> '' then TOBL.putValue('GL_QUALIFQTEVTE',UA);
    End;
    if Ua <> '' then TOBL.putValue('GL_QUALIFQTEACH',UA);
  end;

  Result := Px;

end;

function RecupTarifAch (TOBArt,TOBTarif,TobTiers : TOB; Naturepiece : String; var UniteAch : string ; var CoefuaUs : double; UniteResult : TTypeUniteResult = TurAchat;
                        PrixNet : boolean=true;PrxArtInclus:boolean=true;QteRef : double=1; Dateref : TdateTime=0) : double;
var TOBCatalog    : TOB;
    MTPAF         : Double;
    PrixPourQte   : Double;
    FUA           : Double;
    FUV           : double;
    FUS           : double;
    IsUniteAchat  : Boolean;
    FournCat      : String;
    TiersSaisie,RecupPrix   : String;
begin
  //
  TOBCatalog  := TOB.Create ('CATALOGU',nil,-1);
  RecupPrix := GetInfoParPiece(NaturePiece, 'GPP_APPELPRIX');
  //
	CoefUaUs    := 0;
  MTPAF       := 0;
  //
  UniteAch := 'PASCATALOGUE';
  MTPAF := 0;
  //
  TRY
    chargeTOBs (TOBArt,TOBCatalog,TOBTiers,TOBTarif,QteRef,DateRef);

    //fv1 : 22/05/2014 - FS#1008 - DELABOUDINIERE : Ajout contrôle référencement fournisseur en saisie de pièces d'achats
    //brl : 19/06/2014 - FS#1080 - DELABOUDINIERE : Traitement uniquement en saisie de commandes et proposition achats
    if (GetParamSocSecur('SO_CTRLFRSSAISIE', false)) and ((Naturepiece = 'CF') or (Naturepiece = 'DEF'))  then
    begin
      if TobCatalog.GetString('GCA_ARTICLE')= '' then
      Begin
        PGIInfo('Attention, cet article n''est pas référencé chez le fournisseur.','Récupération Tarif Catalogue');
        MTPAF    := TobArt.GetVAlue('GA_PAUA');
        UniteAch := TobArt.GetValue('GA_QUALIFUNITEACH');
        if uniteach = '' then
        begin
          FUA      := 1;
          CoefuaUs := 1;
          UniteAch := TobArt.GetValue('GA_QUALIFUNITEVTE');
        end
        else
        begin
          FUA      := RatioMesure('PIE', TobArt.GetValue('GA_QUALIFUNITEACH'));
          CoefuaUs := TobArt.getValue('GA_COEFCONVQTEACH');
        end;
        Result := MTPAF;
        Exit;
      end;
    end;

    FournCat    := TOBcatalog.GetString('GCA_TIERS');
    TiersSaisie := TobTIERS.GetString('T_TIERS');
    //
    FUV := RatioMesure('PIE', TobArt.getValue('GA_QUALIFUNITEVTE'));
    FUS := RatioMesure('PIE', TobArt.getValue('GA_QUALIFUNITESTO'));

    if (not GetParamSocSecur ('SO_BTAPLICREMFOUPRINC',false)) and (FournCat <> TiersSaisie) AND (FournCat <> '') then
    begin
      if TobArt.GetString('GA_FOURNPRINC') <> '' then
      begin
        MTPAF    := TobArt.GetVAlue('GA_PAUA');
        UniteAch := TobArt.GetValue('GA_QUALIFUNITEACH');
        FUA      := RatioMesure('PIE', TobArt.GetValue('GA_QUALIFUNITEACH'));
        CoefuaUs := TobArt.getValue('GA_COEFCONVQTEACH');
      end ;
    end
    else
    begin
      //FV1 : 18/02/2015 - Gestion du Prix d'achat dans la fiche article
      //Pourquoi l'unité d'achat du catalogue n'est pas chargée ici ??????
      Uniteach := TobCatalog.getValue('GCA_QUALIFUNITEACH');
      FUA := RatioMesure('PIE', TobCatalog.getValue('GCA_QUALIFUNITEACH'));
      CoefuaUs := TobCatalog.getValue('GCA_COEFCONVQTEACH');
      //
      if TOBTarif.GetValue('GF_PRIXUNITAIRE') <> 0 then
          MTPAF :=TOBTarif.GetValue('GF_PRIXUNITAIRE')
      else
      begin
        PrixPourQte := TOBCatalog.GetValue('GCA_PRIXPOURQTEAC');
        if PrixPourQte = 0 then PrixPourQte := 1;
        if TOBCatalog.GetValue('GCA_PRIXBASE') = 0 then
          MTPAF :=TOBCatalog.GetValue('GCA_PRIXVENTE')/PrixPourQte
        else
          MTPAF :=TOBCatalog.GetValue('GCA_PRIXBASE')/PrixPourQte;
      end;
    end;
    //
    //FV1 : 18/02/2015 - Gestion du Prix d'achat dans la fiche article
    if MTPAF = 0 then
    Begin
      if TobArt.GetString('GA_FOURNPRINC') <> '' then
      begin
        MTPAF    := TobArt.GetVAlue('GA_PAUA');
        UniteAch := TobArt.GetValue('GA_QUALIFUNITEACH');
        FUA      := RatioMesure('PIE', TobArt.GetValue('GA_QUALIFUNITEACH'));
        CoefuaUs := TobArt.getValue('GA_COEFCONVQTEACH');
      end;
      if MTPAF = 0 Then
      begin
        if RecupPrix = 'DPA' then
        begin
          result    := TOBART.GetValue('GA_DPA');
          Exit;
        end else
        begin
          MTPAF    := TOBART.GetValue('GA_PAHT');
        end;
        UniteAch := TOBART.GetValue('GA_QUALIFUNITEVTE');
        FUA      := 1.0;
        CoefuaUs := 1;
      end;
    End;

    // voila voila voila ..le seul hic c'est que ce prix est en UA..donc passage de l'UA en UV
    if (MTPAF <> 0) and (MTPAF <> TobArt.getValue('GA_PAHT')) then TobArt.putValue('GA_PAHT',MTPAF);

    if PrixNet then
      MTPAF := arrondi(MTPAF * (1-(TOBTarif.GetValue('GF_REMISE')/100)),V_PGI.OkDecP )
    else
      MTPAF := arrondi(MTPAF ,V_PGI.OkDecP );

    if UniteResult = TurStock then
    begin
      if CoefUaUs <> 0 then
        MTPAF := (MTPAF * CoefuaUs)
      else
        MTPAF := (MTPAF/FUA ) * FUS;
    end;
    //
    if uniteResult = TurVente then
    begin
      if CoefUaUs <> 0 then
        MTPAF := (MTPAF*CoefUaUs ) / FUS * FUV
      else
        MTPAF := (MTPAF/FUA ) * FUV;
    end;

    Result := MTPAF;

  FINALLY
    TOBCatalog.free;
  END;

end;

function RecupTarifAch (Fournisseur,Article : string; VAR UniteAch : string ; var CoefUaUs : double; UniteResult : TTypeUniteResult = TurAchat;
                        PrixNet : boolean=true;PrxArtInclus:boolean=true;TOBA : TOB=nil;QteRef : double=1; DateRef : TdateTime=0) : double;
var TOBTiers,TobArt,TOBTarif,TOBCatalog : TOB;
    MTPAF,PrixPourQte : double;
    FUV,FUA,FUS,CoefUSUV : double;
    QQ : TQuery;
begin

	CoefUaUs := 0;
  result := 0;

  if fournisseur = '' then exit;

  TOBTarif    := TOB.Create ('TARIF',nil,-1);
  TOBCatalog  := TOB.Create ('CATALOGU',nil,-1);
  TOBTiers    := TOB.Create ('TIERS',nil,-1);
  TOBART      := TOB.Create ('ARTICLE',nil,-1);

  TRY
    //    TOBART.PutValue('GA_ARTICLE',Article);
    QQ := OpenSql ('SELECT GA2_SOUSFAMTARART,* FROM ARTICLE LEFT JOIN ARTICLECOMPL ON GA2_ARTICLE=GA_ARTICLE WHERE GA_ARTICLE="'+Article+'"',True);
    //    TOBARt.LoadDB (true);
    TOBARt.SelectDB ('',QQ);
    ferme (QQ);
    TOBTiers.putValue ('T_TIERS',Fournisseur);
    TOBTiers.LoadDB (true);
    chargeTOBs (TOBArt,TOBCatalog,TOBTiers,TOBTarif,Qteref,Dateref);
    //
    //
    //
    FUV := RatioMesure('PIE', TobArt.getValue('GA_QUALIFUNITEVTE'));
    FUS := RatioMesure('PIE', TobArt.getValue('GA_QUALIFUNITESTO'));
    FUA := RatioMesure('PIE', TobCatalog.getValue('GCA_QUALIFUNITEACH'));
    //
		CoefUsUv := TOBART.GetDouble('GA_COEFCONVQTEVTE');
    //
    if (TOBcatalog.GetString('GCA_TIERS') <> TobTIERS.GetString('T_TIERS')) AND (TOBcatalog.GetString('GCA_TIERS')<> '') then
      begin
        MTPAF    := TobArt.GetVAlue('GA_PAUA');
        UniteAch := TobArt.GetValue('GA_QUALIFUNITEACH');
        FUA      := RatioMesure('PIE', TobArt.GetValue('GA_QUALIFUNITEACH'));
        CoefuaUs := TobArt.getValue('GA_COEFCONVQTEACH');
      end
    else
    begin
      CoefuaUs := TobCatalog.getValue('GCA_COEFCONVQTEACH');
      //Pourquoi l'unité d'achat du catalogue n'est pas chargée ici ??????
      Uniteach := TobCatalog.getValue('GCA_QUALIFUNITEACH');
      //
      if TOBTarif.GetValue('GF_PRIXUNITAIRE') <> 0 then
        MTPAF := TOBTarif.GetValue('GF_PRIXUNITAIRE')
      else
      begin
        PrixPourQte := TOBCatalog.GetValue('GCA_PRIXPOURQTEAC');
        CoefUaUs    := TOBCatalog.GetValue('GCA_COEFCONVQTEACH');
        //
        if PrixPourQte = 0 then PrixPourQte := 1;
        //
        if TOBCatalog.GetValue('GCA_PRIXBASE') = 0 then
        begin
          MTPAF :=TOBCatalog.GetValue('GCA_PRIXVENTE')/prixPourQte;
        end else
        begin
          MTPAF :=TOBCatalog.GetValue('GCA_PRIXBASE')/prixPourQte;
        end;
      end;
    end;


    //FV1 : 18/02/2015 - Gestion du Prix d'achat dans la fiche article
    if MTPAF = 0 then
    Begin
      MTPAF    := TobArt.GetVAlue('GA_PAUA');
      UniteAch := TobArt.GetValue('GA_QUALIFUNITEACH');
      FUA      := RatioMesure('PIE', TobArt.GetValue('GA_QUALIFUNITEACH'));
      CoefuaUs := TobArt.getValue('GA_COEFCONVQTEACH');
      if MTPAF = 0 Then
      begin
        MTPAF    := TOBART.GetValue('GA_PAHT');
        UniteAch := TobArt.GetValue('GA_QUALIFUNITEVTE');
        CoefuaUs := 1;
        FUA      := 1;
      end;
    End;

    if (MTPAF = 0) and (PrxArtInclus) then MTPAF := TOBART.GetValue ('GA_PAHT');

    if PrixNet then
      MTPAF := arrondi(MTPAF * (1-(TOBTarif.GetValue('GF_REMISE')/100)),V_PGI.OkDecP)
    else
      MTPAF := arrondi(MTPAF ,V_PGI.OkDecP );

    if uniteResult = TurStock then
    begin
    	if CoefUaUs <> 0 then
      	MTPAF := (MTPAF / CoefUaUS)
      else
        MTPAF := (MTPAF / FUA ) * FUS;
    end;

    if uniteResult = TurVente then
    begin
      if CoefUaUs <> 0 then
      begin
        if CoefUSUV <> 0 then
      		MTPAF := (MTPAF / CoefuaUs ) / CoefUsUv
        else
        	MTPAF := (MTPAF / CoefuaUs ) / FUS * FUV;
      end
      else
      	MTPAF := (MTPAF / FUA ) * FUV;

      if TOBA <> nil then
      begin
        if not TOBA.FieldExists ('GCA_QUALIFUNITEACH') then TOBA.AddChampSup('GCA_QUALIFUNITEACH',false);
        TOBA.PutValue('GCA_QUALIFUNITEACH',TobCatalog.getValue('GCA_QUALIFUNITEACH'));
        TOBA.PutValue('GA_PAHT',MTPAF);
        TOBA.PutValue('GA_DPA',MTPAF);
        RecalculPrPV (TOBA,TOBCataLog);
      end;
    end;

    Result := MTPAF;

  FINALLY
    TOBTArif.free;
    TOBCatalog.free;
    TOBTiers.free;
    TOBARt.free;
  END;

end;


function GetInfosTarifAch (Fournisseur,Article : string; UniteResult : TTypeUniteResult = TurAchat;
													 PrixNet : boolean=true;PrxArtInclus:boolean=true;TOBA : TOB=nil;QteRef : double=1;
                           DateRef : TdateTime=0) : TGinfostarifFour;
var TOBTiers,TobArt,TOBTarif,TOBCatalog : TOB;
    MTPAF,PrixPourQte,MTPABrut : double;
    FUV,FUA,FUS,CoefuaUS,CoefuSUV : double;
    QQ : TQuery;
    UniteAch : string;
begin
  UniteAch := '';
  CoefuaUS := 0;
  CoefuSUV := 0;
  result.TarifAch := 0;
  result.UniteAchat := '';
  result.TarifAchBrut := 0;
  result.Remise := '';
  result.remisesreelle := 0;
  result.LibelleFournisseur := '';
  result.CoefUAUs := 0;
//  if fournisseur = '' then exit;
  TOBTarif := TOB.Create ('TARIF',nil,-1);
  TOBCatalog := TOB.Create ('CATALOGU',nil,-1);
  TOBTiers := TOB.Create ('TIERS',nil,-1);
  TOBART := TOB.Create ('ARTICLE',nil,-1);
  TRY
    TOBART.PutValue('GA_ARTICLE',Article);
//    TOBARt.LoadDB (true);
    QQ := OpenSql ('SELECT A.*,AC.*,N.BNP_TYPERESSOURCE,N.BNP_LIBELLE FROM ARTICLE A '+
    							 'LEFT JOIN NATUREPREST N ON N.BNP_NATUREPRES=A.GA_NATUREPRES '+
    							 'LEFT JOIN ARTICLECOMPL AC ON AC.GA2_ARTICLE=A.GA_ARTICLE WHERE A.GA_ARTICLE="'+Article+'"',true);
    TOBArt.selectDb ('',QQ);
    ferme(QQ);
    if Fournisseur <> '' then
    begin
      TOBTiers.putValue ('T_TIERS',Fournisseur);
      TOBTiers.LoadDB (true);
      chargeTOBs (TOBArt,TOBCatalog,TOBTiers,TOBTarif,Qteref,Dateref);
    end;
//
    FUV := RatioMesure('PIE', TobArt.getValue('GA_QUALIFUNITEVTE')); if FUV = 0 then FUV := 1;
    FUS := RatioMesure('PIE', TobArt.getValue('GA_QUALIFUNITESTO'));  if FUS = 0 then FUS := 1;
    FUA := RatioMesure('PIE', TobCatalog.getValue('GCA_QUALIFUNITEACH')); if FUA = 0 then FUA := 1;
    coefUaUs := TobCatalog.getValue('GCA_COEFCONVQTEACH');
    UniteAch := TOBCatalog.GetValue('GCA_QUALIFUNITEACH');
//
    if TOBTarif.GetValue('GF_PRIXUNITAIRE') <> 0 then
    begin
        MTPAF :=TOBTarif.GetValue('GF_PRIXUNITAIRE');
    end else
    begin
        PrixPourQte := TOBCatalog.GetValue('GCA_PRIXPOURQTEAC');
        if PrixPourQte = 0 then PrixPourQte := 1;

        if TOBCatalog.GetValue('GCA_PRIXBASE') = 0 then
        begin
          MTPAF :=TOBCatalog.GetValue('GCA_PRIXVENTE')/prixPourQte;
        end else
        begin
          MTPAF :=TOBCatalog.GetValue('GCA_PRIXBASE')/prixPourQte;
        end;
    end;
    //
    coefUSUV := TOBART.GetValue('GA_COEFCONVQTEVTE');
    //
    if (MTPAF = 0) and (PrxArtInclus) then
    begin
      if TOBART.GetValue('GA_QUALIFUNITEACH') <> '' then
      begin
        MTPAF := TOBART.GetValue ('GA_PAUA');
        coefUaUs := TOBART.getValue('GA_COEFCONVQTEACH');
        UniteAch := TOBART.GetValue('GA_QUALIFUNITEACH');
      end else
      begin
        MTPAF := TOBART.GetValue ('GA_PAHT');
        coefUaUs := 0;
        UniteAch := TOBART.GetValue('GA_QUALIFUNITEVTE');
      end;
    end;

    if CoefuaUS = 0 then CoefuaUS := FUS/FUA;
    if CoefuSUV = 0 then CoefuSUV := FUS/FUV;

    if PrixNet then
    begin
      MTPABrut := arrondi(MTPAF ,V_PGI.OkDecP );
    	MTPAF := arrondi(MTPAF * (1-(TOBTarif.GetValue('GF_REMISE')/100)),V_PGI.OkDecP);
    end else
    begin
    	MTPAF := arrondi(MTPAF ,V_PGI.OkDecP );
      MTPABrut := MTPAF;
    end;

    if uniteResult = TurStock then
    begin
      MTPAF := arrondi(MTPAF / CoefUaUs,V_PGI.okdecP);
      MTPABrut := arrondi(MTPABrut /CoefuaUs,V_PGI.okdecP);
    end else
    if uniteResult = TurVente then
    begin
      MTPAF := Arrondi(MTPAF / CoefuaUs / CoefUSUV,V_PGI.okdecP);
      MTPABrut := Arrondi(MTPABrut / CoefUaUS  / COEFUSUV ,V_PGI.okdecP);
      if TOBA <> nil then
      begin
        if not TOBA.FieldExists ('GCA_QUALIFUNITEACH') then TOBA.AddChampSup('GCA_QUALIFUNITEACH',false);
        TOBA.PutValue('GCA_QUALIFUNITEACH',UniteAch);
        TOBA.PutValue('GA_PAHT',MTPAF);
        TOBA.PutValue('GA_DPA',MTPAF);
        RecalculPrPV (TOBA,TOBCataLog);
      end;
    end;
    Result.UniteAchat := UniteAch;
    Result.Remise := TOBTarif.GetValue ('GF_CALCULREMISE');
    result.remisesreelle := TOBTarif.GetValue('GF_REMISE');
    result.cascade := TOBTarif.GetValue('GF_CASCADEREMISE');
    Result.TarifAch := MTPAF;
    Result.TarifAchBrut := MTPABrut;
    result.LibelleFournisseur := TOBTiers.GetValue('T_LIBELLE');
		result.tarif := TOBTarif.GetValue('GF_TARIF');
    result.CoefUAUs := CoefuaUs;
  FINALLY
    TOBTArif.free;
    TOBCatalog.free;
    TOBTiers.free;
    TOBARt.free;
  END;
end;

Function RechPrixArticle (TOBA, TobTarif, Tobtiers, TobPiece, TOBL : TOB; VenteAchat : String) : Double;
var TOBCatalog  : TOB;
    MTPAF       : Double;
    PrixPourQte : Double;
    FUA,FUV,FUS : Double;
    CoefUaUS    : Double;
    UA          : String;
begin

  Result  := 0;

  UA      := '';

  if VenteAchat <> 'ACH' then exit;

  TOBCatalog  := TOB.Create ('CATALOGU',nil,-1);
  ChargeTOBs (TOBA, TOBCatalog, TOBTiers, TOBTarif, TOBL.GetValue('GL_QTEFACT'), TOBL.GetValue('GL_DATEPIECE'));

  FUV := RatioMesure('PIE', TobA.getValue('GA_QUALIFUNITEVTE'));
  FUS := RatioMesure('PIE', TobA.getValue('GA_QUALIFUNITESTO'));
  FUA := RatioMesure('PIE', TobCatalog.getValue('GCA_QUALIFUNITEACH'));
  CoefuaUs := TobCatalog.getValue('GCA_COEFCONVQTEACH');

  if TOBTarif.GetValue('GF_PRIXUNITAIRE') <> 0 then
    MTPAF :=TOBTarif.GetValue('GF_PRIXUNITAIRE')
  else
  begin
    PrixPourQte := TOBCatalog.GetValue('GCA_PRIXPOURQTEAC');
    if PrixPourQte = 0 then PrixPourQte := 1;
    if TOBCatalog.GetValue('GCA_PRIXBASE') = 0 then
      MTPAF :=TOBCatalog.GetValue('GCA_PRIXVENTE')/PrixPourQte
    else
      MTPAF :=TOBCatalog.GetValue('GCA_PRIXBASE')/PrixPourQte;
  end;

  Ua := TOBCatalog.GetValue('GCA_QUALIFUNITEACH');
  if (MTPAF = 0) then //and PrxArtInclus then
  begin
    if CoefUaUs <> 0 then
     MTPAF := TobA.getValue('GA_PAHT') / FUV * FUS / CoefUaUs
    else
      MTPAF := TobA.getValue('GA_PAHT') / FUV * FUA;
  end;

  //if PrixNet then
  //  MTPAF := arrondi(MTPAF * (1-(TOBTarif.GetValue('GF_REMISE')/100)),V_PGI.OkDecP )
  //else
  MTPAF := arrondi(MTPAF ,V_PGI.OkDecP );

  {if uniteResult = TurStock then
  begin
    if CoefUaUs <> 0 then
      MTPAF := (MTPAF * CoefuaUs)
    Else
      MTPAF := (MTPAF/FUA ) * FUS;
  end
  else if uniteResult = TurVente then
  begin
    if CoefUaUs <> 0 then
      MTPAF := (MTPAF*CoefUaUs ) / FUS * FUV
    else
      MTPAF := (MTPAF/FUA ) * FUV;
  end;}

  if MTPAF <> 0 then
  Begin
    TOBL.PutValue('FROMTARIF','X');
    TOBL.PutValue('GL_DPA',MTPAF);
    if Ua <> '' then TOBL.putValue('GL_QUALIFQTEVTE',UA);
  End
  else
  Begin
    MTPAF := TOBL.GetValue('GL_PUHTDEV');
  End;

  Result := MTPAF;

  FreeAndNil (TOBCatalog);

end;


end.
