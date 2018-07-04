unit UDemandeprix;

interface
uses  Classes,
      AglInit,
      forms,
      sysutils,
      UentCommun,
      HEnt1,
      {$IFNDEF EAGLCLIENT}
      db,
      {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
      Fe_Main,
      {$ENDIF}
      HCtrls,hmsgbox,
      DialogEx,
      uTOB;

Type
	TDWriteErr = (TdeOk,TdeArticle,TdeDetail);
  TEtatDemandePrix = (TedmOk,TedmDemandeNonValid);
  TTypeTraitPiece = (TttPEdition,TttPAcceptation,TttPAutre);

procedure SupprimeLigDemPrix (TOBArticles,TOBArticleDemprix,TOBDetailDemPrix,TOBFournDemPrix,TOBPiece,TOBOuvrages,TOBL : TOB);
function  DemPrixAutorise (NaturePiece : string) : boolean;
procedure ValidelesDemPrix (TOBPiece,TOBPieceDemPrix,TOBArticleDemPrix,TOBFournDemPrix,TOBDetailDemPrix : TOB; FromSaisie : boolean = false);
procedure AfficheDemprix(fForm : Tform); overload;
procedure AfficheDemprix(Cledoc : r_cledoc); overload;
procedure AfficheDemprix (TOBPiece,TOBpieceDemPrix,TOBArticleDemPrix,TOBFournDemPrix,TOBDetailDemPrix : TOB) ; overload;


procedure AppelValideDemPrix(fform : Tform);

function  FindFournDemPrix (TPADM,TOBFournDemprix : TOB) : TOB;
function  ChargeDemande (cledoc : r_cledoc ; Numero : integer; TOBPieceDemPrix,TOBArticleDemprix, TOBFournDemPrix, TOBdetailDemPrix : TOB) : boolean;
function  GetNumDemande (cledoc : r_cledoc ) : integer;
function  IsArticleInserable (cledoc : r_cledoc; Article: string; LigneNumOrdre : integer; UniqueBlo : integer) : boolean;
function  IsExisteDemandes (cledoc : r_cledoc) : boolean;
Procedure ChargementGrilleWithTob(TOBAAfficher : TOB; Grille : THGrid; LibColonne : String);

Function SupprimeDdePrix(Cledoc : R_Cledoc; Unique : integer; NomTable, Prefixe : String; ttd : T_TableDoc; FromSaisie : Boolean=false) : Integer; Overload;
function  SupprimeDdePrix(TobToDelete : TOB; Unique : Integer; FromSaisie : Boolean=false) : Integer; Overload;
//
procedure CreatTOBVide (TOBDest: TOB) ;
function  GetNewUniqueLigDemprix (TOBPieceDemPrix : TOB) : integer;
procedure EnregistreGenereDemPrix(cledoc: r_cledoc; Unique : integer);
procedure OrganiseDataFoun (TOBSource,TOBFournDemprix : TOB);
//
Procedure CreateDemandePrix(Cledoc : R_Cledoc; Unique : Integer; Action : String);  OverLoad;
Procedure CreateDemandePrix(TobDemPrix : Tob; Unique : Integer); OverLoad;
//
Procedure ExporteDocumentToExcel(Cledoc : R_cledoc;Unique : Integer; Libelle : String); Overload;
Procedure ExporteDocumentToExcel(TobPiece, TobArtFrs, TobFrsDemPx : TOB; Unique : Integer; Libelle : String);  Overload;
//
Procedure CreateTobFrsInterm(TobArtFrs, TOBFrs, TOBArt : TOB);
Procedure CreateTobArtInterm(TobArtFrs, TOBArt : TOB; Fournisseur : String);
//
function  ExportToExcel(fcledoc: r_cledoc; TobPiece, TobFrs, TobArtFrs : TOB; Unique : Integer) : boolean; Overload;
//
function  AddEntete (TOBP, TobExcel : TOB; Unique : Integer): TOB;
Procedure ChargeAdresseFrs (TOBentete, Tobfourn : TOB);
function  AddLigExcel (Fournisseur : String; Unique, NumLig : integer; TOBA, TobExcel : TOB): TOB;
Function  Recherche_dichoDsTOB(TobRech : TOB; ZoneRech : String;ValRech : Variant) : TOB;
Procedure ChargeComboWithTob(Combo : THValComboBox; TobCharge : TOB; Champ, Valeur : String);
//
procedure TraiteDemandePrix (form : Tform);
function  FindDetailDemPrix (TPADM,TOBdetailDemprix : TOB) : TOB;

function  CheckDemandesprix (cledoc : r_cledoc) : TEtatDemandePrix; overload;
function  CheckDemandesprix (TOBPieceDemPrix : TOB) : TEtatDemandePrix; overload;
function  AlertDemandePrix (Cledoc : r_cledoc; TypeTraitement : TTypeTraitPiece=TttPAutre): boolean; overload;
function  AlertDemandePrix (TOBPieceDemPrix : TOB; TypeTraitement : TTypeTraitPiece=TttPAutre):boolean; overload;
//
function  findOccurenceString(St : String; Caract : Char) : Integer;
procedure InitDemandePrix (TOBPieceDemPrix,TOBArticleDemPrix,TOBfournDemprix,TOBDetailDemprix : TOB);

//
implementation

uses UtilTOBPiece,facture,factTOB, Grids,BTPUtil,
		 FactOuvrage,FactVariante,UTofListeInv,FactFormule,
     CalcOLEGenericBTP,
     factArticle,
     FactUtil,
     Paramsoc,
     TiersUtil,
     FactCalc,
     UfactExportXLS, Math,
     PiecesRecalculs,
     SAISUTIL,
     NomenUtil,
     UtilPgi;

procedure InitDemandePrix (TOBPieceDemPrix,TOBArticleDemPrix,TOBfournDemprix,TOBDetailDemprix : TOB);
begin
  TOBPieceDemPrix.ClearDetail;
  TOBArticleDemPrix.ClearDetail;
  TOBfournDemprix.ClearDetail;
  TOBDetailDemprix.ClearDetail;
end;

function TransformeUVEnUA (QteVte : double; CoefUAUS : double; StrUV,StrUS,StrUA : string) : double;
var UA,US,UV : double;
begin

  UA := RatioMesure('PIE',StrUA);
  UV := RatioMesure('PIE',StrUV);
  US := RatioMesure('PIE',StrUS);
  if CoefUaUS<>0 then
  begin
    result := QteVte * UV / US / CoefUAUS;
  end else
  begin
    result := QteVte * UV / US * UA;
  end;
end;

function DemPrixAutorise (NaturePiece : string) : boolean;
begin
  result :=  (pos (NaturePiece,'DBT;ETU;BCE;CBT:PBT')>0);
end;

function MiseAjoutLigneDoc (TOBPiece,TOBADM,TOBPA : TOB) : Boolean;
var CodeFrs : string;
		Dpa,NbJour,CoefuAUs,CoefUA,COEFUS,CoefUv : double;
    TOBA : TOB;
    Cledoc : R_CLEDOC;
    SQL : string;
begin
  Result := true;
  CodeFrs := TOBPA.GetString('BD1_TIERS');
  TOBA := TOB.Create ('ARTICLE',nil,-1);
  TRY
  	Cledoc := TOB2CleDoc(TOBPiece);
    //chargement du DPA donné par le fournisseur et du nombre de jour
    Dpa       := TOBPA.Getdouble('BD1_PRIXACH');
    NBJour    := TOBPA.Getdouble('BD1_NBJOUR');
    CoefUAUS  := TOBPA.GetDouble('BD1_COEFCONVQTEACH');
    CoefUA    := RatioMesure('PIE', TOBPA.GetString('BD1_QUALIFUNITEACH'));
    CoefUV    := RatioMesure('PIE', TOBADM.GetString('BDP_QUALIFUNITEVTE'));
    TOBA.SelectDB('"'+TOBADM.GetString('BDP_ARTICLE')+'"',nil);
    if TOBA = nil then CoefUs := 0
    else CoefUS := RatioMesure('PIE',TOBA.GetString('GA_QUALIFUNITESTO'));
    if CoefUAUS = 0 then
      Dpa := ARRONDI((Dpa / CoefUA) * CoefUV,V_PGI.OkDecP)
    else
      Dpa := ARRONDI((Dpa / CoefUAUS) * CoefUV,V_PGI.OkDecP);
    SQL := 'UPDATE LIGNE SET GL_DPA='+STRFPOINT(dpa)+ ' WHERE '+WherePiece (Cledoc,ttdLigne,false)+' AND '+
    			 'GL_ARTICLE="'+TOBADM.GetString('BDP_ARTICLE')+'"';
    TRY
	    ExecuteSQL(SQL );
    EXCEPT
      result := false;
    END;
    if Result then
    begin
      SQL := 'UPDATE LIGNEOUV SET BLO_DPA='+STRFPOINT(dpa)+ ' WHERE '+WherePiece (Cledoc,ttdOuvrage,false) + ' AND '+
    			 'BLO_ARTICLE="'+TOBADM.GetString('BDP_ARTICLE')+'"';
      TRY
        ExecuteSQL(SQL);
      EXCEPT
        result := false;
      END;
    end;
  FINALLY
  	TOBA.Free;
  END;
end;

procedure ValidelesDemPrix (TOBPiece,TOBPieceDemPrix,TOBArticleDemPrix,TOBFournDemPrix,TOBDetailDemPrix : TOB; FromSaisie : boolean = false);
var Indice,II : integer;
		TOBB,TOBTT,TOBPA : TOB;
    ARecalc,Okok : boolean;
begin
  Arecalc := false;
  for Indice := 0 to TOBPieceDemprix.detail.count -1 do
  begin
    TOBPieceDemprix.detail[Indice].SetString('BPP_NATUREPIECEG',TOBPiece.getString('GP_NATUREPIECEG'));
    TOBPieceDemprix.detail[Indice].SetString('BPP_SOUCHE',TOBPiece.getString('GP_SOUCHE'));
    TOBPieceDemprix.detail[Indice].SetInteger('BPP_NUMERO',TOBPiece.GetInteger ('GP_NUMERO'));
    TOBPieceDemprix.detail[Indice].SetInteger('BPP_INDICEG',TOBPiece.GetInteger('GP_INDICEG'));
    //
    TOBPieceDemprix.detail[Indice].SetAllModifie(true);
    if not TOBpieceDemPrix.detail[Indice].insertDB(nil) then V_PGI.IOError := oeUnknown ;
  end;
  if V_PGI.IOError <> OeOk then Exit;

  for Indice := 0 to TOBArticleDemPrix.detail.count -1 do
  begin
    if TOBArticleDemPrix.Detail[Indice].GetInteger('BDP_UNIQUE') = 0 then continue;
    TOBArticleDemprix.detail[Indice].SetString('BDP_NATUREPIECEG',TOBPiece.getString('GP_NATUREPIECEG'));
    TOBArticleDemprix.detail[Indice].SetString('BDP_SOUCHE',TOBPiece.getString('GP_SOUCHE'));
    TOBArticleDemprix.detail[Indice].SetInteger('BDP_NUMERO',TOBPiece.GetInteger ('GP_NUMERO'));
    TOBArticleDemprix.detail[Indice].SetInteger('BDP_INDICEG',TOBPiece.GetInteger('GP_INDICEG'));
    TOBArticleDemprix.detail[Indice].SetAllModifie(true);
    if not TOBArticleDemprix.detail[Indice].InsertDB (nil) then V_PGI.Ioerror := OeUnknown;
    if not FromSaisie then
    begin
      TOBTT := FindFournDemPrix (TOBArticleDemprix.detail[Indice],TOBFournDemprix);
      if TOBTT <> nil then
      begin
        TOBPA := TOBTT.FindFirst(['BD1_SELECTED'], ['X'], True);
        if TOBPA <> nil then
        begin
          Arecalc := true;
          okok := false;
          TRY
            okok := MiseAjoutLigneDoc (TOBPiece,TOBArticleDemPrix.detail[Indice],TOBPA);
          EXCEPT
            on E: Exception do
            begin
              PgiError('Erreur SQL : ' + E.Message, 'Demande de prix');
            end;
          END;
          if not okok then
          begin
            V_PGI.IOError := oeUnknown;
            break;
          end;
        end;
      end;
    end;
  end;
  if V_PGI.IOError <> OeOk then Exit;

  for Indice := 0 to TOBFournDemprix.detail.count -1 do
  begin
    for II := 0 to TOBFournDemprix.detail[Indice].detail.count -1 do
    begin
      TOBB := TOBFournDemprix.detail[Indice].detail[II];
      if TOBB.GetInteger('BD1_UNIQUE') = 0 then continue;
      TOBB.SetString('BD1_NATUREPIECEG',TOBPiece.getString('GP_NATUREPIECEG'));
      TOBB.SetString('BD1_SOUCHE',TOBPiece.getString('GP_SOUCHE'));
      TOBB.SetInteger('BD1_NUMERO',TOBPiece.GetInteger ('GP_NUMERO'));
      TOBB.SetInteger('BD1_INDICEG',TOBPiece.GetInteger('GP_INDICEG'));
      TOBB.SetAllModifie(true);
      okok := false;
      TRY
        okok := TOBB.InsertDB (nil) ;
      EXCEPT
        on E: Exception do
        begin
          PgiError('Erreur SQL : ' + E.Message, 'Demande de prix');
        end;
      END;
      if not okok then V_PGI.IOError := oeUnknown;
    end;
  end;
  if V_PGI.IOError <> OeOk then Exit;

  for Indice := 0 to TOBDetailDemprix.detail.count -1 do
  begin
    for II := 0 to TOBDetailDemprix.detail[Indice].detail.count -1 do
    begin
      TOBB := TOBDetailDemprix.detail[Indice].detail[II];
      if TOBB.GetInteger('BD0_UNIQUE') = 0 then continue;
      TOBB.SetString('BD0_NATUREPIECEG',TOBPiece.getString('GP_NATUREPIECEG'));
      TOBB.SetString('BD0_SOUCHE',TOBPiece.getString('GP_SOUCHE'));
      TOBB.SetInteger('BD0_NUMERO',TOBPiece.GetInteger ('GP_NUMERO'));
      TOBB.SetInteger('BD0_INDICEG',TOBPiece.GetInteger('GP_INDICEG'));
      TOBB.SetAllModifie(true);
      okok := false;
      TRY
        okok := TOBB.InsertDb (nil);
      EXCEPT
        on E: Exception do
        begin
          PgiError('Erreur SQL : ' + E.Message, 'Demande de prix');
        end;
      END;
      if not okok then V_PGI.IOError := oeUnknown;
    end;
  end;
  if V_PGI.IOError <> OeOk then Exit;
  if (not FromSaisie) and (Arecalc) then
  begin
    okok := false;
    TRY
  	  okok := TraitementRecalculPiece (TOBPiece,True,True) <> TrrOk;
    EXCEPT
      on E: Exception do
      begin
        PgiError('Erreur SQL : ' + E.Message, 'Recalcul demande prix');
      end;
    END;
    if not okok  then V_PGI.IOError := oeUnknown;
  end;
end;

function LoadLesEnteteDemPrix (cledoc : r_cledoc; TOBpieceDemPrix : TOB) : boolean;
var QQ: TQuery;
begin
  result := false;
  QQ := OpenSql ('SELECT * FROM PIECEDEMPRIX WHERE '+WherePiece(cledoc,ttdPieceDemPrix,false),true,-1,'',true);
  if not QQ.eof then
  begin
    TOBPieceDemPrix.LoadDetailDB('PIECEDEMPRIX','','',QQ,false);
    result := true;
  end;
  ferme (QQ);
end;

function AlertDemandePrix (TOBPieceDemPrix : TOB; TypeTraitement : TTypeTraitPiece=TttPAutre): boolean; overload;
Var CBtnTxt         : PCustBtnText;
		InfosDocument : string;
begin

  result := true;

  if TypeTraitement = TttPAutre then exit;

  if (TypeTraitement = TttPEdition) or (TypeTraitement = TttPAcceptation) then
  begin
    //FV1 : 06/06/2014 - FS#877 - TEST BRL : en acceptation de devis avec demande de prix, message systématique de demande non validée
    if CheckDemandesprix (TOBPieceDemPrix) = TedmDemandeNonValid then
    begin
      InfosDocument := RechDom('GCNATUREPIECEG',TOBPieceDemPrix.detail[0].GetSTring('BPP_NATUREPIECEG'),false)+
      								 ' N° '+IntToStr(TOBPieceDemPrix.detail[0].Getinteger('BPP_NUMERO'));
      New(CBtnTxt);
      CBtnTxt[mbCust1] := 'Continuer';
      CBtnTxt[mbCust4] := 'Annuler';
      Case ExMessageDlg(InfosDocument+#13#10+'Certaines demandes de prix ne sont pas validées', mtxWarning, [mbCust1, mbCust4], 0, mbCust2, Application.Icon, CBtnTxt) Of
        mrCust1: result := true;
        mrCust4: result := false;
      End;
      dispose(CBtnTxt);
    end;
  end;

end;

function AlertDemandePrix (Cledoc : r_cledoc; TypeTraitement : TTypeTraitPiece=TttPAutre): boolean; overload;
var TOBPieceDemPrix : TOB;
begin
  result := true;
  if TypeTraitement = TttPAutre then exit;
  if (TypeTraitement = TttPEdition) or (TypeTraitement = TttPAcceptation) then
  begin
    TOBPieceDemPrix := TOB.Create ('LES ENTETES DEMPRIX',nil,-1);
    TRY
      if LoadLesEnteteDemPrix(cledoc,TOBPieceDemprix) then
      begin
        result := AlertDemandePrix (TOBPieceDemPrix,TypeTraitement);
      end;
    FINALLY
    	TOBPieceDemPrix.free;
    END;
  end;
end;

function CheckDemandesprix (cledoc : r_cledoc) : TEtatDemandePrix; overload;
var TOBPieceDemPrix : TOB;
begin
  result := TedmOk;
  TOBPieceDemPrix := TOB.Create ('LES ENTETES DEMPRIX',nil,-1);
  TRY
    if LoadLesEnteteDemPrix(cledoc,TOBpieceDemPrix) then
    begin
      result := CheckDemandesprix (TOBPieceDemPrix);
    end;
  FINALLY
    TOBPieceDemPrix.free;
  END;
end;

function CheckDemandesprix (TOBPieceDemPrix : TOB) : TEtatDemandePrix; overload;
var Indice : integer;
begin
  result := TedmOk;
	for Indice := 0 to TOBPieceDemPrix.detail.count -1 do
  begin
    if TOBPieceDemPrix.detail[Indice].GetString('BPP_TRAITE')='-' then
    begin
      result := TedmDemandeNonValid;
      break;
    end;
  end;
end;

procedure  EnregistreGenereDemPrix(cledoc: r_cledoc; Unique : integer);
var SQl : string;
begin

  if (cledoc.naturePiece='') OR
  	 (cledoc.souche = '') OR
     (cledoc.NumeroPiece = 0) OR
     (Unique = 0) then exit;

	Sql := 'UPDATE PIECEDEMPRIX SET BPP_ENVOYE="X" WHERE '+
  			 'BPP_NATUREPIECEG="'+cledoc.NaturePiece+'" AND '+
  			 'BPP_SOUCHE="'+cledoc.Souche+'" AND '+
  			 'BPP_NUMERO='+IntToStr(cledoc.NumeroPiece)+' AND '+
  			 'BPP_INDICEG='+IntToStr(cledoc.Indice)+' AND '+
  			 'BPP_UNIQUE='+IntToStr(Unique);
  ExecuteSql (SQl);
end;


procedure AfficheDemprix(fform : Tform); overload;
var XX        : Tffacture;
    TobPiece  : TOB;
    TOBparam  : TOB;
begin

  TOBparam := TOB.Create ('LES PARAMS',nil,-1);
  TOBParam.AddChampSupValeur ('RETOUR','');
  TOBParam.AddChampSupValeur ('FROMSAISIE','X');

  TRY
    XX := TFFacture(fform);

    TOBPiece := XX.LaPieceCourante;
    TOBPiece.data := XX.ThePieceDemPrix;
    XX.ThePieceDemPrix.data := XX.TheArticleDemPrix;
    XX.TheArticleDemPrix.Data := XX.TheFournDemPrix;
    XX.TheFournDemPrix.Data := XX.TheDetailDemPrix;
    TOBParam.data := TOBPiece;
    TheTOB := TOBParam;

  	AGLLanceFiche('BTP','BTDEMPRIX_MUL','','','ACTION=MODIFICATION');

    Thetob := Nil;

    if TOBParam.getString('RETOUR')='X'  then
    begin
    end;
  FINALLY
    TOBParam.free;
  END;

end;

procedure AfficheDemprix (TOBPiece,TOBpieceDemPrix,TOBArticleDemPrix,TOBFournDemPrix,TOBDetailDemPrix : TOB) ; overload;
var TOBParam : TOB;
begin
  TOBparam := TOB.Create ('LES PARAMS',nil,-1);
  TOBParam.AddChampSupValeur ('RETOUR','');
  TOBPiece.data := TOBpieceDemPrix;
  TOBpieceDemPrix.data := TOBArticleDemPrix;
  TOBArticleDemPrix.Data := TOBFournDemPrix;
  TOBFournDemPrix.Data := TOBDetailDemPrix;
  TOBParam.data := TOBPiece;
  TheTOB := TOBParam;
  TRY
    AGLLanceFiche('BTP','BTDEMPRIX_MUL','','','ACTION=MODIFICATION');
  FINALLY
    TOBParam.free;
  END;
end;


procedure AfficheDemprix(Cledoc : r_cledoc); overload;
begin
	AGLLanceFiche('BTP','BTPIECDDEPRIX_MUL','BPP_NATUREPIECEG='+Cledoc.NaturePiece+';'+
  																				'BPP_SOUCHE='+Cledoc.Souche +';'+
                                          'BPP_NUMERO='+IntToStr(Cledoc.NumeroPiece )+';'+
                                          'BPP_INDICEG='+IntToStr(Cledoc.indice),'','ACTION=MODIFICATION');
end;

procedure IncidenceModifPa (TOBOuvrages,TOBO : TOB; DEV : Rdevise);
var Qte : double;
    valeurs : T_valeurs;
begin
	Qte := TOBO.GetValue('BLO_QTEFACT');
	TOBO.SetDouble ('GA_PAHT',TOBO.GetDouble ('BLO_DPA'));
  if (TOBO.GetValue('BLO_TYPEARTICLE')= 'ARP') and (TOBO.detail.count >0) then
  begin
  	TOBO.detail[0].PutValue ('GA_FOURNPRINC',TOBO.getValue('GA_FOURNPRINC'));
    TOBO.detail[0].putvalue ('BLO_FOURNISSEUR',TOBO.GetValue('BLO_FOURNISSEUR'));
    TOBO.detail[0].putvalue ('GA_PAHT',TOBO.GetValue('GA_PAHT'));
    TOBO.detail[0].putvalue ('BLO_DPA',TOBO.GetValue('GA_PAHT'));
    TOBO.detail[0].PutValue('GCA_PRIXBASE',TOBO.GetValue('GCA_PRIXBASE'));
    TOBO.detail[0].PutValue('GF_CALCULREMISE','');

    CalculeLigneAcOuv (TOBO.detail[0],DEV,true,TOBO);
    TOBO.detail[0].putvalue ('GA_PVHT',TOBO.detail[0].GetValue ('BLO_PUHT'));
    CalculMontantHtDevLigOuv (TOBO.detail[0],DEV);
    CalculOuvFromDetail (TOBO,DEV,Valeurs);
    TOBO.Putvalue ('BLO_DPA',Valeurs[0]);
    TOBO.Putvalue ('BLO_DPR',Valeurs[1]);
    TOBO.putvalue ('GA_PVHT',Valeurs[2]);
    TOBO.Putvalue ('BLO_PUHT',Valeurs[2]);
    TOBO.Putvalue ('BLO_PUHTBASE',Valeurs[2]);
    TOBO.Putvalue ('BLO_PUHTDEV',pivottodevise(Valeurs[2],DEV.Taux,DEV.quotite,V_PGI.OkdecP));
    TOBO.Putvalue ('BLO_PUTTC',Valeurs[3]);
    TOBO.Putvalue ('BLO_PUTTCDEV',pivottodevise(Valeurs[3],DEV.Taux,DEV.quotite,V_PGI.OkdecP));
    TOBO.Putvalue ('BLO_PMAP',Valeurs[6]);
    TOBO.Putvalue ('BLO_PMRP',Valeurs[7]);
    //
    TOBO.PutValue('BLO_TPSUNITAIRE',Valeurs[9]);
    TOBO.PutValue('GA_HEURE',Valeurs[9]);
    TOBO.PutValue('BLO_MONTANTFG',Arrondi(Valeurs[13]*Qte,V_PGI.OkdecP));
    TOBO.PutValue('BLO_MONTANTFC',Arrondi(Valeurs[15]*Qte,V_PGI.OkdecP));
    TOBO.PutValue('BLO_MONTANTFR',Arrondi(Valeurs[14]*Qte,V_PGI.OkdecP));
    CalculeLigneAcOuvCumul (TOBO);
    //
    GetValoDetail (TOBO);
  end else
  begin
    CalculeLigneAcOuv (TOBO,DEV,true);
    TOBO.putvalue ('GA_PVHT',TOBO.GetValue ('BLO_PUHT'));
  end ;
  CalculMontantHtDevLigOuv (TOBO,DEV);
  if (TOBO.GetValue('BLO_TYPEARTICLE')= 'ARP') and (TOBO.detail.count >0) then
  begin
    TOBO.PutValue('ANCPV', TOBO.GetValue('BLO_PUHTDEV'));
    TOBO.PutValue('ANCPA', TOBO.GetValue('BLO_DPA'));
    TOBO.PutValue('ANCPR', TOBO.GetValue('BLO_DPR'));
  end;

end;

procedure MAJPrxAchatPieceDemPrix (fform : Tform);
var XX                : Tffacture;
    TOBARecalculer    : TOB;
    TobArticle        : TOB;
    TOBLigARecalculer : TOB;
    TobPiece          : Tob;
    TobTravail        : TOB;
    TobFrsDdePrix     : TOB;
    TOBPA             : TOB;
    TobDetail         : TOB;
    TobDetailOuvrage  : TOB;
    TobDetOuv         : TOB;
    TobL              : TOB;
    iInd              : Integer;
    jInd              : Integer;
    Numordre          : Integer;
    UniqueBlo         : Integer;
    CodeFrs           : String;
    Dpa               : Double;
    NBJour            : Double;
    CoefUAUS          : Double;
    CoefUA            : Double;
    CoefUS            : Double;
    CoefUV            : Double;
    CoefPaPr,CoefPrPv : double;
    AllTraite : Boolean;
begin
	AllTraite := true;
  XX := Tffacture (fform);

  //Si au moins une ligne à été sélectionner mettre la pièce à recalculer...
  TobTravail := XX.TheArticleDemPrix.FindFirst(['BDP_SELECTIONNE'],['X'], false);
  if TobTravail = nil then Exit;

  //Trier TobPiece sur le GL_NUMORDRE...
  TobPiece := XX.LaPieceCourante;
  TobPiece.Detail.sort('GL_NUMORDRE');
  //Création d'une tob pour recalcul
  TOBARecalculer := Tob.Create('A RECALCULER', nil, -1);
  TRY

    //Parcours des ligne article Demande de prix traitées
    for iInd := 0 To XX.TheArticleDemPrix.detail.count -1 do
    begin
      TobTravail := XX.TheArticleDemPrix.detail[iInd];
      //si on trouve rien on passe à l'enreg suivant !!!
      if TobTravail = nil then continue;
      if (TobTravail.getboolean('BDP_SELECTIONNE')) then
      begin
        //chargement du Prix d'Achat en fonction de la ligne fournisseur sélectionnée
        TobFrsDdePrix := FindFournDemPrix(TobTravail,XX.TheFournDemPrix);
        TOBPA := TobFrsDdePrix.FindFirst(['BD1_SELECTED'], ['X'], True);

        //si on trouve rien on passe à l'enreg suivant !!!
        if TOBPA = nil then continue;

        CodeFrs := TOBPA.GetString('BD1_TIERS');

        //chargement du DPA donné par le fournisseur et du nombre de jour
        Dpa       := TOBPA.Getdouble('BD1_PRIXACH');
        NBJour    := TOBPA.Getdouble('BD1_NBJOUR');

        //transformation du prix en unité d'achat au prix en unité de vente
        //chargement des coefs à partir des information du fournisseur
        CoefUAUS  := TOBPA.GetDouble('BD1_COEFCONVQTEACH');
        CoefUA    := RatioMesure('PIE', TOBPA.GetString('BD1_QUALIFUNITEACH'));
        CoefUV    := RatioMesure('PIE', TobTravail.GetString('BDP_QUALIFUNITEVTE'));

        //Recup du Coef de Stock dans tob Article
        TobArticle:= XX.TheTOBArticles.FindFirst(['GA_ARTICLE'], [TobTravail.GetString('BDP_ARTICLE')], True);
        if TobArticle = nil then CoefUs := 0
        else CoefUS := RatioMesure('PIE',TobArticle.GetString('GA_QUALIFUNITESTO'));

        //Conversion du dpa de prix d'Achat à prix de vente
        if CoefUAUS = 0 then
          Dpa := (Dpa / CoefUA) * CoefUV
        else
          Dpa := (Dpa / CoefUAUS) * CoefUV;

        //recherche des lignes détail demande de prix.
        TobDetail := FindDetailDemPrix(TobTravail, XX.TheDetailDemPrix);
        //si rien on passe son cheùmin et à l'enregistrement suivant
        if Tobdetail = nil then continue;
        For jInd  := 0 To TobDetail.detail.count -1 do
        begin
          Numordre  := TobDetail.Detail[jInd].GetInteger('BD0_NUMORDRE');
          UniqueBlo := TobDetail.Detail[jInd].GetInteger('BD0_UNIQUEBLO');
          TobL := Recherche_dichoDsTOB(TobPiece, 'GL_NUMORDRE', NumOrdre);
          //si on trouve rien on passe à l'enreg suivant !!!
          if TOBL = nil then continue;
          
          //on charge une table annexe pour pouvoir relancer le calcul des lignes d'ouvrage
          TobLigARecalculer := TobARecalculer.FindFirst(['NUMORDRE'],[Numordre],true);
          if TOBLigARecalculer= nil then
          begin
            TobLigARecalculer := Tob.create('LIG A RECALCULER', TobARecalculer, -1);
            TobLigARecalculer.AddChampSupValeur('NUMORDRE',NumOrdre);
            TobLigARecalculer.data := TOBL;
            TOBL.PutValue('GL_RECALCULER', 'X');
            //on réinitialise la ligne dans tout les cas (article ou ouvrage)
            XX.deduitLaLigne(TOBL);
            ZeroLigneMontant(TOBL);
          end;

          // les calculs sont gérés de manière relativement simple du fait que les elements dans Demprix sont
          // forcement des détails et jamais des ouvrages..quelquesoit le niveau
          //si article simple chargement simple
          if uniqueBlo = 0 then
          begin
            if TOBL.GetValue('GL_DPA') <> 0 then CoefPaPr := TOBL.GetValue('GL_DPR')/TOBL.GetValue('GL_DPA')
                                            else CoefPaPr := 1;
            if TOBL.GetValue('GL_DPR') <> 0 then CoefPrPv := TOBL.GetValue('GL_PUHTDEV')/TOBL.GetValue('GL_DPR')
                                            else CoefPrPv := 1;
            TOBL.SetString('GL_FOURNISSEUR', CodeFrs);
            TOBL.SetDouble('GL_DPA', Dpa);
            TOBL.SetDouble('GL_DPR', Arrondi(Dpa * Coefpapr,V_PGI.okdecP));
            if TOBL.GetValue('GL_BLOQUETARIF') = '-' then
            begin
              TOBL.SetDouble('GL_PUHTDEV', Arrondi(TOBL.GetValue('GL_DPR') * CoefPrPV,V_PGI.OkdecP));
            end else
            begin
              TOBL.SetDouble('GL_COEFMARG', 0);
            end;
            TOBL.PutValue('GL_RECALCULER', 'X');
          end else
          begin
            //si Ouvrage on cherche les détails d'ouvrage associés
            if TOBL.GetInteger('GL_INDICENOMEN') = 0 then continue;
            TobDetailOuvrage := XX.TheTOBOuvrage.Detail[TOBL.GetInteger('GL_INDICENOMEN')-1];
            if TobDetailOuvrage = nil then continue;
            //recherche dans détail d'ouvrage d'une ligne identique
            TobDetOuv := TobDetailOuvrage.findfirst(['BLO_UNIQUEBLO'], [UniqueBlo], True);
            if TobDetOuv = nil then continue;
            //si on a trouvé on charge les valeurs
            TobDetOuv.SetString('BLO_FOURNISSEUR', CodeFrs);
            TobDetOuv.SetDouble('BLO_DPA', Dpa);
  					TobDetOuv.PutValue('GA_FOURNPRINC',CodeFrs);
  					TobDetOuv.PutValue('GCA_PRIXBASE',Dpa);
  					TobDetOuv.PutValue('GF_CALCULREMISE','');
            IncidenceModifPa (XX.TheTOBOuvrage,TobDetOuv,XX.DEV);
          end;
        end;
        TobTravail.PutValue('BDP_TRAITE', 'X');
      end else
      begin
        AllTraite := false;
      end;
    end;

    //Traitement des lignes à recalculer
    for iInd := 0 to TobARecalculer.detail.count -1 do
    begin
      TobL := TOB(TobARecalculer.Detail[iInd].data);
      if TOBL.GetInteger('GL_INDICENOMEN') = 0 then continue;
      TobDetailOuvrage := XX.TheTOBOuvrage.Detail[TOBL.GetInteger('GL_INDICENOMEN')-1];
      if TobDetailOuvrage = nil then continue;
      RecalculeLigneOuv (TobDetailOuvrage,TOBL,XX.DEV,XX.LaPieceCourante.GetBoolean('GP_FACTUREHT'));
    end;
    if TOBARecalculer.detail.count > 0 then  TobPiece.SetString('GP_RECALCULER', 'X');
    if AllTraite then XX.ThePieceDemPrix.detail[0].SetBoolean('BPP_TRAITE',true) else XX.ThePieceDemPrix.detail[0].SetBoolean('BPP_TRAITE',false); 
  FINALLY
    //trie de la tob pièce sur le Numéro de Ligne et plus sur le Numéro d'Ordre
    TobPiece.Detail.sort('GL_NUMLIGNE');
    //libération mémoire
    FreeandNil(TobARecalculer);
  END;

end;


Function Recherche_dichoDsTOB(TobRech : TOB; ZoneRech : String;ValRech : Variant) : TOB;
var Milieu   : integer;
    Valeur   : Variant;
    Ok_Find  : Boolean;
    BorneInf : Integer;
    BorneSup : Integer;
begin

  Ok_Find   := false;
  result := nil;
  //
  BorneInf  := 0;
  BorneSup  := TobRech.detail.count-1;
  // recherche sur le 1er au cas ou
  Valeur := Tobrech.Detail[BorneInf].GetValue(ZoneRech);
  if Valeur = valrech then
  begin
  	result := Tobrech.Detail[BorneInf];
    exit;
  end;
  // recherche sur le dernier au cas ou
  Valeur := Tobrech.Detail[Bornesup].GetValue(ZoneRech);
  if Valeur = valrech then
  begin
  	result := Tobrech.Detail[Bornesup];
    exit;
  end;

  //Boucle de recherche
  Repeat
    Milieu := (Borneinf + Bornesup) div 2;
    Valeur := Tobrech.Detail[milieu].GetValue(ZoneRech);
    if Valeur = Valrech Then
    begin
      result := Tobrech.Detail[milieu];
      break;
    end else
    begin
    	if (BorneInf=Milieu) Or (BorneSup=Milieu) then Break;
      if Valeur > ValRech then
        BorneSup := Milieu
      else
        BorneInf := Milieu;
    end;
  until (Ok_Find);
end;

procedure AppelValideDemPrix(fform : Tform);
var XX : Tffacture;
    TOBparam  : TOB;
    TobPiece  : TOB;
begin
  TOBparam := TOB.Create ('LES PARAMS',nil,-1);
  TOBParam.AddChampSupValeur ('RETOUR','');
  TRY
    XX := TFFacture(fform);
    XX.GS.SynEnabled := false;

    TobPiece := XX.LaPieceCourante;
    XX.TheFournDemPrix.Data   := XX.TheDetailDemPrix;
    XX.TheArticleDemPrix.Data := XX.TheFournDemPrix;
    XX.ThePieceDemPrix.data   := XX.TheArticleDemPrix;
    TobPiece.data             := XX.ThePieceDemPrix;
    TobParam.data             := TobPiece;
    TheTOB                    := TOBParam;

    AGLLanceFiche('BTP','BTVALIDEDDEPRIX','','','');

    Thetob := Nil;
    if TOBParam.getString('RETOUR')='X'  then
    begin
      MAJPrxAchatPieceDemPrix (fform);
      XX.GS.BeginUpdate;
      if TOBPiece.getString('GP_RECALCULER')='X' then XX.CalculeLaSaisie(-1,-1,true);
      // LoadLesLibDetail  (TobPiece,nil,XX.TheTOBOuvrage ,XX.TheTOBTiers,XX.TheTOBAffaire,XX.DEv,false);
      XX.GS.EndUpdate;
  		XX.ThePieceDemPrix.PutValue('MODIFIED','X');
    end;
  FINALLY
    XX.GS.SynEnabled := true;
    TOBParam.free;
  END;
end;


function FindDetailDemPrix (TPADM,TOBdetailDemprix : TOB) : TOB;
begin
	result := TOBdetailDemprix.findFirst(['UNIQUE','UNIQUELIG'],[TPADM.GetInteger('BDP_UNIQUE'),TPADM.GetInteger('BDP_UNIQUELIG')],true);
end;

function FindFournDemPrix (TPADM,TOBFournDemprix : TOB) : TOB;
begin
	result := TOBFournDemprix.findFirst(['UNIQUE','UNIQUELIG'],[TPADM.GetInteger('BDP_UNIQUE'),TPADM.GetInteger('BDP_UNIQUELIG')],true);
end;

function FindLigInDetailDemPrix (TOBDDM: TOB; NumOrdre,uniqueBlo : integer): TOB;
var indice : integer;
begin
  result := nil;
  for Indice := 0 to TOBDDM.detail.count -1 do
  begin
     if (TOBDDM.detail[Indice].GetInteger ('BD0_NUMORDRE')=NumOrdre) and
        (TOBDDM.detail[Indice].GetInteger ('BD0_UNIQUEBLO')=UniqueBlo) then
     begin
       result := TOBDDM.detail[Indice];
       break;
     end;
  end;
end;

procedure OrganiseDataDetail (TOBSource,TOBDetailDemprix : TOB);
var TOBD,TPADM : TOB;
begin
  if TOBSource.detail.count = 0 then exit;
	TPADM := TOBSource.detail[0];
	repeat
  	TOBD := TOBdetailDemprix.findFirst(['UNIQUE','UNIQUELIG'],[TPADM.GetInteger('BD0_UNIQUE'),TPADM.GetInteger('BD0_UNIQUELIG')],false);
    if TOBD = nil then
    begin
      TOBD := TOB.Create ('UN DETAIL DE LIGNE',TOBDetailDemprix,-1);
      TOBD.AddChampSupValeur ('UNIQUE',TPADM.GetInteger('BD0_UNIQUE'));
      TOBD.AddChampSupValeur ('UNIQUELIG',TPADM.GetInteger('BD0_UNIQUELIG'));
    end;
    TPADM.ChangeParent(TOBD,-1);
    if TOBSource.detail.count > 0 then
    begin
			TPADM := TOBSource.detail[0];
    end;
  until TOBSource.detail.count = 0;
end;

procedure OrganiseDataFoun (TOBSource,TOBFournDemprix : TOB);
var TOBD,TPADM : TOB;
begin
  if TOBSource.detail.count = 0 then exit;
	TPADM := TOBSource.detail[0];
	repeat
  	TOBD := TOBFournDemprix.findFirst(['UNIQUE','UNIQUELIG'],[TPADM.GetInteger('BD1_UNIQUE'),TPADM.GetInteger('BD1_UNIQUELIG')],false);
    if TOBD = nil then
    begin
    	TOBD := TOB.Create ('LES FOURN D UNE LIGNE',TOBFournDemprix,-1);
      TOBD.AddChampSupValeur ('UNIQUE',TPADM.GetInteger('BD1_UNIQUE'));
      TOBD.AddChampSupValeur ('UNIQUELIG',TPADM.GetInteger('BD1_UNIQUELIG'));
    end;
    TPADM.ChangeParent(TOBD,-1);
    if TOBSource.detail.count > 0 then
    begin
			TPADM := TOBSource.detail[0];
    end;
	until TOBSource.detail.count = 0;
end;

function findArticle (TOBpieceDemprix,TOBArticleDemprix,TOBdetailDemPrix,TOBFournDemPrix,TOBL,TOBPere : TOB) : TOB;
var prefixe : string;
begin
  result := nil;
	prefixe := TableToPrefixe(TOBL.NomTable);
  if TOBPere <> nil then
  begin
    // TODO
  end else
  begin
  	result := TOBArticleDemprix.findFirst(['BDP_ARTICLE'],[TOBL.GetString(prefixe+'_ARTICLE')],false);
  end;
end;

function GetNewUniqueLigDemprix (TOBPieceDemPrix : TOB) : integer;
begin
  result := TOBPieceDemPrix.GetInteger('BPP_LASTLIG')+1;
  TOBPieceDemPrix.SetInteger('BPP_LASTLIG',result);
end;

procedure ConstitueBesoinFournisseurs(TOBArticles,TOBPFD,TOBPD,TPADM : TOB);
var TOBCata,TOBArt,TOBBF,TDD : TOB;
		QQ : TQuery;
    Indice,Ind1 : integer;
    QteVte,QteAch : double;
    //UA,US,UV : double;
begin
  // ------------------------------
  // TOBPFD = TOB Pere des fournisseurs de la ligne de demande de prix
  // TOBPD = TOB Pere du détail des lignes
  // TPADM = TOB de la ligne de demande de prix
  // --------------------------------
  TOBCata := TOB.Create ('LES CATALOGUES',nil,-1);
  TOBArt := TOB.Create ('ARTICLE',nil,-1);
  TRY
    if (TPADM.getString('NEW')='X') or (TPADM.GetString('UPDATED')='X') then
    begin
      QQ := OpenSql ('SELECT * FROM CATALOGU WHERE GCA_ARTICLE="'+TPADM.GetString('BDP_ARTICLE')+'"',true,-1,'',true);
      if not QQ.eof then
      begin
        TOBCata.LoadDetailDB('CATALOGU','','',QQ,false);
      end;
      ferme (QQ);
      TOBArt := TOBArticles.findFirst(['GA_ARTICLE'],[TPADM.GetString('BDP_ARTICLE')],true);
      if TOBARt = nil then
      begin
        QQ := OpenSql ('SELECT A.*,AC.*,N.BNP_TYPERESSOURCE,N.BNP_LIBELLE FROM ARTICLE A '+
                       'LEFT JOIN NATUREPREST N ON N.BNP_NATUREPRES=A.GA_NATUREPRES '+
                       'LEFT JOIN ARTICLECOMPL AC ON AC.GA2_ARTICLE=A.GA_ARTICLE WHERE A.GA_ARTICLE="'+
                       TPADM.GetString('BDP_ARTICLE')+'"',true,-1, '', True);

        if not QQ.eof then
        begin
      		TOBArt := CreerTOBArt(TOBArticles);
					ChargerTobArt(TOBArt,nil,'VEN','',QQ);
				end;
        ferme (QQ);
      end;
      // Constitution de la liste des fournisseurs avec leur reference et leur unite d'achat
      for Indice := 0 to TOBCata.detail.count -1 do
      begin
        TOBBF := TOBPFD.findFirst(['BD1_TIERS'],[TOBCata.detail[Indice].GetValue('GCA_TIERS')],true);
        if TOBBF = nil then
        begin
          TOBBF := TOB.Create ('FOURLIGDEMPRIX',TOBPFD,-1);
          TOBBF.SetString  ('BD1_TIERS',TOBCata.detail[Indice].GetValue('GCA_TIERS'));
          TOBBF.SetString  ('BD1_NATUREPIECEG',TPADM.GetString('BDP_NATUREPIECEG'));
          TOBBF.SetString  ('BD1_SOUCHE',TPADM.GetString('BDP_SOUCHE'));
          TOBBF.SetInteger ('BD1_NUMERO',TPADM.GetInteger('BDP_NUMERO'));
          TOBBF.SetInteger ('BD1_INDICEG',TPADM.GetInteger('BDP_INDICEG'));
          TOBBF.SetInteger ('BD1_UNIQUE',TPADM.GetInteger('BDP_UNIQUE'));
          TOBBF.SetInteger ('BD1_UNIQUELIG',TPADM.GetInteger('BDP_UNIQUELIG'));
          TOBBF.SetString  ('BD1_QUALIFUNITEACH',TOBCata.detail[Indice].GetString('GCA_QUALIFUNITEACH'));
          TOBBF.SetString  ('BD1_REFERENCE',TOBCata.detail[Indice].GetString('GCA_REFERENCE'));
        end;
        TOBBF.SetDouble  ('BD1_QTEACH',0);
        TOBBF.SetString  ('BD1_COEFCONVQTEACH',TOBCata.detail[Indice].GetString('GCA_COEFCONVQTEACH'));
      end;
      //
      if TOBArt.GetString('GA_FOURNPRINC') <> '' then
      begin
        TOBBF := TOBPFD.findFirst(['BD1_TIERS'],[TOBArt.GetString('GA_FOURNPRINC')],true);
        if TOBBF = nil then
        begin
          TOBBF := TOB.Create ('FOURLIGDEMPRIX',TOBPFD,-1);
          TOBBF.SetDouble  ('BD1_QTEACH',0);
          TOBBF.SetString  ('BD1_TIERS',TOBArt.GetString('GA_FOURNPRINC'));
          TOBBF.SetString  ('BD1_NATUREPIECEG',TPADM.GetString('BDP_NATUREPIECEG'));
          TOBBF.SetString  ('BD1_SOUCHE',TPADM.GetString('BDP_SOUCHE'));
          TOBBF.SetInteger ('BD1_NUMERO',TPADM.GetInteger('BDP_NUMERO'));
          TOBBF.SetInteger ('BD1_INDICEG',TPADM.GetInteger('BDP_INDICEG'));
          TOBBF.SetInteger ('BD1_UNIQUE',TPADM.GetInteger('BDP_UNIQUE'));
          TOBBF.SetInteger ('BD1_UNIQUELIG',TPADM.GetInteger('BDP_UNIQUELIG'));
          TOBBF.SetString ('BD1_REFERENCE',TOBART.GetString('GA_CODEARTICLE'));
          TOBBF.SetString  ('BD1_QUALIFUNITEACH',TPADM.GetString('BDP_QUALIFUNITEVTE'));
        end;
      end;
    end;//
    if TOBPFD.detail.count > 0 then
    begin
      for Indice := 0 to TOBPD.detail.count -1 do
      begin
        TDD := TOBPD.detail[Indice]; // parcours des besoin en UV
        for Ind1 := 0 to TOBPFD.detail.count -1 do
        begin
          TOBBF := TOBPFD.detail[Ind1]; // parcours par fournisseur
          // conversion unite de vente en unite d'achat
          QteVte := TDD.GetDouble ('BD0_QTEBESOIN');
          QteAch := TransformeUVEnUA (QteVte,TOBBF.GetDouble ('BD1_COEFCONVQTEACH'),
                                             TDD.GetString ('BD0_QUALIFUNITEVTE'),
                                             TOBARt.GetString ('GA_QUALIFUNITESTO'),
                                             TDD.GetString ('BD0_QUALIFUNITEVTE'));
          TOBBF.SetDouble('BD1_QTEACH',TOBBF.GetDouble('BD1_QTEACH') + QteAch);
        end;
      end;
      // Dernier tour...
      for Ind1 := 0 to TOBPFD.detail.count -1 do
      begin
        TOBBF := TOBPFD.detail[Ind1]; // parcours par fournisseur
        TOBBF.SetDouble('BD1_QTEACH',Arrondi(TOBBF.GetDouble('BD1_QTEACH'),V_PGI.okdecQ));
      end;
    end;
  FINALLY
  	TOBCata.free;
    TOBArt.free;
	END;
end;

procedure  MiseAjourInfosArticle (TOBArticles,TOBADM,TOBDetailDemprix,TOBFournDemprix : TOB);
var TOBPFD,TOBPD : TOB;
		Indice : integer;
    MtPA,MtPv : double;
begin
  TOBPD :=  FindDetailDemPrix (TOBADM,TOBDetailDemprix);
  if TOBPD = nil then exit; // --> oullaaaaa sort vite de la
  TOBADM.SetDouble ('BDP_QTEBESOIN',0); // on reinit pour le calcul via le détail
  TOBADM.SetDouble ('BDP_DPA',0); // on reinit pour le calcul via le détail
  TOBADM.SetDouble ('BDP_PUHTDEV',0); // on reinit pour le calcul via le détail
  MtPa := 0;
  MtPV := 0;
  for Indice := 0 to TOBPD.detail.count -1 do
  begin
		TOBADM.SetDouble ('BDP_QTEBESOIN',TOBADM.GetDouble ('BDP_QTEBESOIN')+TOBPD.detail[Indice].GetDouble('BD0_QTEBESOIN'));
    MtPA := MtPA + (TOBPD.detail[Indice].GetDouble('BD0_QTEBESOIN') *TOBPD.detail[Indice].GetDouble('BD0_DPA'));
    MtPV := MtPV + (TOBPD.detail[Indice].GetDouble('BD0_QTEBESOIN') *TOBPD.detail[Indice].GetDouble('BD0_PUHTDEV'));
  end;
  if TOBADM.GetDouble ('BDP_QTEBESOIN') <> 0 then
  begin
  	TOBADM.SetDouble ('BDP_DPA',Arrondi(MTPA / TOBADM.GetDouble ('BDP_QTEBESOIN'),V_PGI.okdecP));
  	TOBADM.SetDouble ('BDP_PUHTDEV',Arrondi(MTPV / TOBADM.GetDouble ('BDP_QTEBESOIN'),V_PGI.okdecP));
  end;
  // définition de FOURNDEMPRIX
  TOBPFD := FindFournDemPrix (TOBADM,TOBFournDemprix);
  if TOBPFD = nil Then
  begin
    TOBPFD := TOB.Create ('LES FOURN D UNE LIGNE',TOBFournDemprix,-1);
    TOBPFD.AddChampSupValeur ('UNIQUE',TOBADM.GetInteger('BDP_UNIQUE'));
    TOBPFD.AddChampSupValeur ('UNIQUELIG',TOBADM.GetInteger('BDP_UNIQUELIG'));
  end;
  ConstitueBesoinFournisseurs(TOBArticles,TOBPFD,TOBPD,TOBADM);
end;


function ChargeLaDemandePrix (Cledoc : r_cledoc; NumDemande : integer;TOBPieceDemPrix,TOBArticleDemprix,TOBDetailDemprix,TOBFournDemprix : TOB) : boolean;
var QQ : Tquery;
		SQL : String;
    TLocal : TOB;
    TableChoixCod,TableCommun : string;
begin
  TableChoixCod := 'CHOIXCOD';
  TableCommun := 'COMMUN';
  result := false;
  Tlocal := TOB.Create ('LES DATAS',nil,-1);
  SQL := 'SELECT * FROM PIECEDEMPRIX WHERE '+WherePiece(cledoc,ttdPieceDemPrix,false)+' AND BPP_UNIQUE='+InttoStr(NumDemande);
  QQ := OpenSql (SQL,true,1,'',true);
  if not QQ.eof then
  begin
    result := true;
    TOBPieceDemPrix.SelectDB('',QQ);
  end;
  ferme (QQ);
  // ---
  SQL := 'SELECT ARTICLEDEMPRIX.*,GA_TYPEARTICLE AS TYPEARTICLE,'+
  			 '(SELECT CO_LIBELLE FROM '+TableCommun+' WHERE CO_TYPE="TYA" AND CO_CODE=GA_TYPEARTICLE) AS LIBTYPEARTICLE,'+
  			 'GA_CODEARTICLE AS CODEARTICLE,'+
   			 'GA_FAMILLENIV1 AS FAMILLENIV1,GA_FAMILLENIV2 AS FAMILLENIV2,GA_FAMILLENIV3 AS FAMILLENIV3,'+
         '(SELECT CC_LIBELLE FROM '+TableChoixCod+' WHERE CC_TYPE="FN1" AND CC_CODE=GA_FAMILLENIV1) AS LIBFAMILLENIV1,'+
         '(SELECT CC_LIBELLE FROM '+TableChoixCod+' WHERE CC_TYPE="FN2" AND CC_CODE=GA_FAMILLENIV2) AS LIBFAMILLENIV2,'+
         '(SELECT CC_LIBELLE FROM '+TableChoixCod+' WHERE CC_TYPE="FN3" AND CC_CODE=GA_FAMILLENIV3) AS LIBFAMILLENIV3 '+
  			 'FROM ARTICLEDEMPRIX '+
  			 'LEFT JOIN ARTICLE ON GA_ARTICLE=BDP_ARTICLE '+
  			 'WHERE '+WherePiece(cledoc,TTdArticleDemPrix,false)+' AND BDP_UNIQUE='+InttoStr(NumDemande);
  QQ := OpenSql (SQL,true,-1,'',true);
  if not QQ.eof then
  begin
  	TOBArticleDemprix.LoadDetailDB('ARTICLEDEMPRIX','','',QQ,false);
  end;
  ferme (QQ);
  // ---
  SQL := 'SELECT * FROM DETAILDEMPRIX WHERE '+WherePiece(cledoc,TtdDetailDemPrix,false)+' AND BD0_UNIQUE='+InttoStr(NumDemande);
  QQ := OpenSql (SQL,true,-1,'',true);
  if not QQ.eof then
  begin
  	Tlocal.LoadDetailDB('DETAILDEMPRIX','','',QQ,false);
  end;
  ferme (QQ);
  OrganiseDataDetail (Tlocal,TOBDetailDemprix);
  Tlocal.ClearDetail;
  // ---
  SQL := 'SELECT * FROM FOURLIGDEMPRIX WHERE '+WherePiece(cledoc,TtdFournDemprix,false)+' AND BD1_UNIQUE='+InttoStr(NumDemande);
  QQ := OpenSql (SQL,true,-1,'',true);
  if not QQ.eof then
  begin
  	Tlocal.LoadDetailDB('FOURLIGDEMPRIX','','',QQ,false);
  end;
  ferme (QQ);
  OrganiseDataFoun (Tlocal,TOBFournDemprix);
  // ---
  Tlocal.free;
end;

function FindLibelleTypeArticle (TOBL : TOB) : string;
var TableCommun,Sql,prefixe : string;
		QQ : TQuery;
begin
  TableCommun := 'COMMUN';
  prefixe := TableToPrefixe(TOBL.NomTable);
  result := '';
	Sql := 'SELECT CO_LIBELLE FROM '+TableCommun+' WHERE CO_TYPE="TYA" '+
  			 'AND CO_CODE="'+TOBL.GetString(prefixe+'_TYPEARTICLE')+'"';
  QQ := OpenSql (Sql,true,1,'',true);
  if not QQ.eof then result  := QQ.findField('CO_LIBELLE').AsString;
  ferme (QQ);
end;

function FindLibelleFamille(niv : integer;TOBL : TOB) : string;
var TableChoixCod,Sql,CodeFamille,Prefixe : string;
		QQ : TQuery;
begin
  TableChoixCod := 'CHOIXCOD';
  prefixe := TableToPrefixe(TOBL.NomTable);
  CodeFamille := TOBL.GetString(prefixe+'_FAMILLENIV'+InttoStr(niv));
  result := '';
	Sql := 'SELECT CC_LIBELLE FROM '+TableChoixCod+' WHERE CC_TYPE="FN'+IntToStr(Niv)+'" '+
  			 'AND CC_CODE="'+CodeFamille+'"';
  QQ := OpenSql (Sql,true,1,'',true);
  if not QQ.eof then result  := QQ.findField('CC_LIBELLE').AsString;
  ferme (QQ);
end;

procedure InsereDemPrix (TOBLigne,TOBArticleDemprix,TOBL,TOBpere,TOBDetailDemprix : TOB; var NumLigTmp : integer);
var TPADM,TDD,TOBPD : TOB;
		prefixe : string;
begin
  Inc(NumLigTmp);
  prefixe := TableToPrefixe(TOBL.NomTable);
  // insertion de la ligne dans article demande de prix
	TPADM := TOB.Create ('ARTICLEDEMPRIX',TOBArticleDemprix,-1);
  TPADM.AddChampSupValeur('TYPEARTICLE','');
  TPADM.AddChampSupValeur('LIBTYPEARTICLE','');
  TPADM.AddChampSupValeur('CODEARTICLE','');
  TPADM.AddChampSupValeur('FAMILLENIV1','');
  TPADM.AddChampSupValeur('FAMILLENIV2','');
  TPADM.AddChampSupValeur('FAMILLENIV3','');
  TPADM.AddChampSupValeur('LIBFAMILLENIV1','');
  TPADM.AddChampSupValeur('LIBFAMILLENIV2','');
  TPADM.AddChampSupValeur('LIBFAMILLENIV3','');
  TPADM.AddChampSupValeur('UPDATED','-');
  TPADM.AddChampSupValeur('NEW','X');
  TPADM.SetString  ('BDP_NATUREPIECEG',TOBL.GetString(prefixe+'_NATUREPIECEG'));
  TPADM.SetString  ('BDP_SOUCHE',TOBL.GetString(prefixe+'_SOUCHE'));
  TPADM.SetInteger ('BDP_NUMERO',TOBL.GetInteger(prefixe+'_NUMERO'));
  TPADM.SetInteger ('BDP_INDICEG',TOBL.GetInteger(prefixe+'_INDICEG'));
  TPADM.SetInteger ('BDP_UNIQUE',0);
  TPADM.SetInteger ('BDP_UNIQUELIG',NumLigTmp);
  TPADM.SetInteger ('BDP_NUMLIGNE',NumLigTmp);
  TPADM.SetInteger ('BDP_NUMPERE',0); // par défaut
  TPADM.SetString  ('CODEARTICLE',TOBL.GetString(prefixe+'_CODEARTICLE'));
  if TOBpere<>nil then
  begin
  	TPADM.SetInteger ('BDP_NUMPERE',TOBPere.Getinteger('BPD_UNIQUELIG'));
  end;
  TPADM.SetString ('BDP_TYPELIGNE','LIG');
  TPADM.SetString ('BDP_ARTICLE',TOBL.GetString(prefixe+'_ARTICLE'));
  TPADM.SetString ('BDP_LIBELLE',TOBL.GetString(prefixe+'_LIBELLE'));
  TPADM.SetString ('BDP_TRAITE','-');
  TPADM.SetString ('BDP_SELECTIONNE','-');
  TPADM.SetDouble ('BDP_QTEBESOIN',TOBL.GetDouble(prefixe+'_QTEFACT'));
  TPADM.SetString ('BDP_QUALIFUNITEVTE',TOBL.GetString(prefixe+'_QUALIFQTEVTE'));
  TPADM.SetString ('TYPEARTICLE',TOBL.GetString(prefixe+'_TYPEARTICLE'));
  TPADM.SetString ('LIBTYPEARTICLE',FindLibelleTypeArticle(TOBL));
  TPADM.SetString ('FAMILLENIV1',TOBL.GetString(prefixe+'_FAMILLENIV1'));
  TPADM.SetString ('FAMILLENIV2',TOBL.GetString(prefixe+'_FAMILLENIV2'));
  TPADM.SetString ('FAMILLENIV3',TOBL.GetString(prefixe+'_FAMILLENIV3'));
  TPADM.SetString ('LIBFAMILLENIV1',FindLibelleFamille(1,TOBL));
  TPADM.SetString ('LIBFAMILLENIV2',FindLibelleFamille(2,TOBL));
  TPADM.SetString ('LIBFAMILLENIV3',FindLibelleFamille(3,TOBL));

  // définition de DETAILDEMPRIX
  TOBPD := FindDetailDemPrix (TPADM,TOBdetailDemprix);
  if TOBPD = nil Then
  begin
    TOBPD := TOB.Create ('UN DETAIL DE LIGNE',TOBDetailDemprix,-1);
    TOBPD.AddChampSupValeur ('UNIQUE',TPADM.GetInteger('BDP_UNIQUE'));
    TOBPD.AddChampSupValeur ('UNIQUELIG',TPADM.GetInteger('BDP_UNIQUELIG'));
  end;
	TDD := TOB.Create ('DETAILDEMPRIX',TOBPD,-1);
  TDD.SetString  ('BD0_NATUREPIECEG',TPADM.GetString('BDP_NATUREPIECEG'));
  TDD.SetString  ('BD0_SOUCHE',TPADM.GetString('BDP_SOUCHE'));
  TDD.SetInteger ('BD0_NUMERO',TPADM.GetInteger('BDP_NUMERO'));
  TDD.SetInteger ('BD0_INDICEG',TPADM.GetInteger('BDP_INDICEG'));
  TDD.SetInteger ('BD0_UNIQUE',TPADM.GetInteger('BDP_UNIQUE'));
  TDD.SetInteger ('BD0_UNIQUELIG',TPADM.GetInteger('BDP_UNIQUELIG'));
  TDD.SetInteger ('BD0_NUMORDRE',TOBLigne.GetInteger('GL_NUMORDRE'));
  if prefixe = 'BOP' then
  begin
  	TDD.SetInteger ('BD0_UNIQUEBLO',TOBL.GetInteger('UNIQUEBLO'));
  end;
  TDD.SetDouble ('BD0_QTEBESOIN',TOBL.GetDouble(prefixe+'_QTEFACT'));
  TDD.SetDouble ('BD0_DPA',TOBL.GetDouble(prefixe+'_DPA'));
  TDD.SetDouble ('BD0_PUHTDEV',TOBL.GetDouble(prefixe+'_PUHTDEV'));
  TDD.SetString ('BD0_QUALIFUNITEVTE',TOBL.GetString(prefixe+'_QUALIFQTEVTE'));
end;

procedure MajDemPrix (TOBLigne,TOBL,TOBFind,TOBDetailDemprix : TOB);
var TOBDP,TOBPPD : TOB;
		prefixe : string;
    NumOrdre,UniqueBlo : integer;
begin
  prefixe := TableToPrefixe(TOBL.NomTable);
  if not TOBFind.FieldExists('UPDATED') then
  begin
  	TOBfind.AddChampSupValeur('UPDATED','X');
  end else
  begin
  	TOBfind.SetString('UPDATED','X');
  end;
  NumOrdre := TOBLigne.GetInteger ('GL_NUMORDRE');
  if prefixe = 'BOP' then
  begin
  	UniqueBlo := TOBL.GetInteger ('UNIQUEBLO');
  end else
  begin
  	UniqueBlo := 0;
  end;
  TOBPPD := FindDetailDemPrix (TOBFind,TOBdetailDemprix);
  if TOBPPD = nil Then
  begin
    TOBPPD := TOB.Create ('UN DETAIL DE LIGNE',TOBDetailDemprix,-1);
    TOBPPD.AddChampSupValeur ('UNIQUE',TOBfind.GetInteger('BDP_UNIQUE'));
    TOBPPD.AddChampSupValeur ('UNIQUELIG',TOBfind.GetInteger('BDP_UNIQUELIG'));
  end;
  TOBFind.SetDouble ('BDP_QTEBESOIN',0);
  TOBDP := FindLigInDetailDemPrix(TOBPPD,NumOrdre,UniqueBlo);
  if TOBDP = nil then
  begin
    TOBDP := TOB.Create ('DETAILDEMPRIX',TOBPPD,-1);
    TOBDP.SetString  ('BD0_NATUREPIECEG',TOBFind.GetString('BDP_NATUREPIECEG'));
    TOBDP.SetString  ('BD0_SOUCHE',TOBFind.GetString('BDP_SOUCHE'));
    TOBDP.SetInteger ('BD0_NUMERO',TOBFind.GetInteger('BDP_NUMERO'));
    TOBDP.SetInteger ('BD0_INDICEG',TOBFind.GetInteger('BDP_INDICEG'));
    TOBDP.SetInteger ('BD0_UNIQUE',TOBFind.GetInteger('BDP_UNIQUE'));
    TOBDP.SetInteger ('BD0_UNIQUELIG',TOBFind.GetInteger('BDP_UNIQUELIG'));
    TOBDP.SetInteger ('BD0_NUMORDRE',NumOrdre);
    TOBDP.SetInteger ('BD0_UNIQUEBLO',UniqueBlo);
  end;
  TOBDP.SetDouble ('BD0_QTEBESOIN',TOBL.GetDouble(prefixe+'_QTEFACT'));
  TOBDP.SetDouble ('BD0_DPA',TOBL.GetDouble(prefixe+'_DPA'));
  TOBDP.SetDouble ('BD0_PUHTDEV',TOBL.GetDouble(prefixe+'_PUHTDEV'));
  TOBDP.SetString ('BD0_QUALIFUNITEVTE',TOBL.GetString(prefixe+'_QUALIFQTEVTE'));
end;


procedure ConstitueLigneDemPrix (TOBArticles,TOBpiece,TOBL,TOBOuvrage,TOBPieceDemPrix,TOBArticleDemprix,TOBDetailDemprix,TOBFournDemprix : TOB; var numLigTmp : integer);
var TOBpere,TOBPlat,TOBI,TOBA : TOB;
    Article,TypeArticle,TypeLigne : string;
    indice : integer;
    QQ : Tquery;
//    Decompose : boolean;
begin
	TOBPere := nil;
//  Decompose := (TOBpiece.GetString('GP_NATUREPIECEG')='PBT');
  TOBplat := TOB.Create ('LES OUVRAGES PLATS',nil,-1);
  TRY
    //
    Article     := TOBL.GetString('GL_ARTICLE');
    TypeArticle := TOBL.GetString('GL_TYPEARTICLE');
    TypeLigne   := TOBL.GetString('GL_TYPELIGNE');
    //
    if Article = '' then Exit; //ça ne sert à rien de faire des traitement avec un code article à blanc
    if pos(copy(TypeLigne, 1,2),'DP;TP')>0 then exit;
    if (copy(TypeLigne,1,2)='TP') then exit;
    if (Pos(TOBL.GetString('GLC_NATURETRAVAIL'),'001')>0) then exit; //si, cotraitance --> sortie
    if (pos(TypeArticle,'POU;COM;COV')>0) then exit;
    if (pos(TypeLigne,'SD;SDV;')>0) then exit;
    // --
		if ((TypeArticle='OUV') OR (TypeArticle='ARP')) and (TOBL.GetValue('GL_INDICENOMEN')>0) then
    BEGIN
      if (Pos(TOBL.GetString('GLC_NATURETRAVAIL'),'002')>0) then exit; // si cotraité ou ss traité --> suivant
      TRY
        MiseAplatouv (TOBPiece,TOBL,TOBOuvrage,TOBPLat,true,true,true,false);
        for indice := 0 to TOBPlat.detail.count -1 do
        begin
          // VARIANTE
    			Article := TOBPLat.detail[indice].GetString('BOP_ARTICLE');
          TOBA := TOBArticles.findfirst(['GA_ARTICLE'],[Article],true);
          if TOBA = nil then
          begin
            QQ := OpenSql ('SELECT A.*,AC.*,N.BNP_TYPERESSOURCE,N.BNP_LIBELLE FROM ARTICLE A '+
                           'LEFT JOIN NATUREPREST N ON N.BNP_NATUREPRES=A.GA_NATUREPRES '+
                           'LEFT JOIN ARTICLECOMPL AC ON AC.GA2_ARTICLE=A.GA_ARTICLE WHERE A.GA_ARTICLE="'+
                           Article+'"',true,-1, '', True);

            if not QQ.eof then
            begin
              TOBA := CreerTOBArt(TOBArticles);
              ChargerTobArt(TOBA,nil,'VEN','',QQ);
            end;
            ferme (QQ);
          end;
          //FV1 : 15/02/2018 - FS#2933 - SES_ETANCHEITE SERVICE :Problème avec les demandes de prix sur un devis
          //c'est violent comme sortie :)
    			//if (not TOBA.GetBoolean('GA_GEREDEMPRIX')) then exit; // non géré dans les demande de prix
    			if (not TOBA.GetBoolean('GA_GEREDEMPRIX')) then Continue; // non géré dans les demande de prix
          if (Pos(TOBPLAT.GetString('BOP_NATURETRAVAIL'),'001;002')>0) then exit; // si cotraité ou ss traité --> suivant
          //??????
    			if IsPrestationInterne(Article) then continue;  // on ne va pas demander des prix pour des presta interne ..

          if TOBPLat.detail[indice].GetValue('BOP_QTEFACT') = 0 then continue;
          //gestion de la non-prise en compte des lignes Cotraitantes
          if (TOBPLat.detail[indice].GetString('BOP_NATURETRAVAIL')='001') then continue;
          if (TOBPLat.detail[indice].GetValue('BOP_TYPEARTICLE')='PRE') then
          begin
            TOBPLat.detail[indice].PutValue('BNP_TYPERESSOURCE',RenvoieTypeRes(TOBPLat.detail[indice].GetValue('BOP_ARTICLE')));
          end;
          if (TOBPLat.detail[indice].getvalue('BOP_TYPEARTICLE')='POU') then continue;
          //
          TOBI := findArticle (TOBPieceDemPrix, TOBArticleDemprix,TOBDetailDemprix,TOBFournDemPrix, TOBPlat.detail[indice],TOBPere);
          if TOBI = nil then
          BEGIN
            InsereDemPrix (TOBL,TOBArticleDemprix,TOBPlat.detail[indice],TOBpere,TOBDetailDemprix,NumLigTmp)
          END ELSE
          BEGIN
            MajDemPrix (TOBL,TOBPlat.detail[indice],TOBI,TOBDetailDemprix);
          end;
        end;
      FINALLY
        TOBPlat.clearDetail;
      END;
    END
    else if (TypeArticle='MAR') OR
            (TypeArticle='PRE') OR //AND (IsPrestationInterne(Article) ) ) OR
            (TypeArticle='FRA') OR
           ((TypeArticle='ARP') and (TOBL.GetValue('GL_INDICENOMEN')=0)) then
    BEGIN
      If (TypeArticle='PRE') AND (IsPrestationInterne(Article)) AND (TOBL.GetString('GLC_NATURETRAVAIL')='') then exit; //si c'est une prestation et que sous-traité --> suivant;
      if TOBL.GetValue('GL_QTEFACT') = 0 then exit;
      //
      TOBA := TOBArticles.findfirst(['GA_ARTICLE'],[Article],true);
      if TOBA = nil then
      begin
        QQ := OpenSql ('SELECT A.*,AC.*,N.BNP_TYPERESSOURCE,N.BNP_LIBELLE FROM ARTICLE A '+
                       'LEFT JOIN NATUREPREST N ON N.BNP_NATUREPRES=A.GA_NATUREPRES '+
                       'LEFT JOIN ARTICLECOMPL AC ON AC.GA2_ARTICLE=A.GA_ARTICLE WHERE A.GA_ARTICLE="'+
                       Article+'"',true,-1, '', True);

        if not QQ.eof then
        begin
          TOBA := CreerTOBArt(TOBArticles);
          ChargerTobArt(TOBA,nil,'VEN','',QQ);
        end;
        ferme (QQ);
      end;
      if (not TOBA.GetBoolean('GA_GEREDEMPRIX')) then Exit; // non géré dans les demande de prix
      //
      TOBI := findArticle (TOBPieceDemPrix,TOBArticleDemprix,TOBDetailDemprix,TOBFournDemPrix,TOBL,TOBpere);
      if TOBI = nil then
      BEGIN
        InsereDemPrix (TOBL,TOBArticleDemprix,TOBL,TOBpere,TOBDetailDemprix,NumLigTmp);
      END ELSE
      BEGIN
        MajDemPrix (TOBL,TOBL,TOBI,TOBDetailDemprix);
      end;
    END;
  FINALLY
    TOBPlat.free;
  END;
end;

procedure NumeroteLig (TOBArticleDemPrix : TOB);
var indice : integer;
		NumLigne : integer;
    lastDemande : integer;
begin
  lastDemande := 0;
  For Indice := 0 to TOBArticleDemPrix.detail.count -1 do
  begin
    if TOBArticleDemprix.detail[Indice].GetInteger('BPP_UNIQUE')<>lastDemande then
    begin
      NuMligne := 1;
      lastDemande := TOBArticleDemprix.detail[Indice].GetInteger('BPP_UNIQUE');
    end;
    TOBArticleDemprix.detail[Indice].setInteger ('BPP_NUMLIGNE',NumLigne);
    inc(Numligne);
  end;
end;

procedure ConstitueIntermFromPiece (form : TForm;TOBArticles,TOBpieceDemprix,TOBArticleDemprix,TOBFournDemprix,TOBDetailDemprix : TOB);
var XX : Tffacture;
		indice : integer;
		TOBpiece ,TOBOuvrage,TOBL : TOB;
    NumLigTmp : integer;
begin
  NumLigTmp := 0;
	XX := form as Tffacture;
  TOBPiece := XX.LaPieceCourante;
  TOBOuvrage := XX.TheTOBOuvrage;
  for Indice := 0 to TOBPiece.detail.count -1 do
  begin
    TOBL := TOBpiece.detail[Indice];
    ConstitueLigneDemPrix (TOBArticles,TOBpiece,TOBL,TOBOuvrage,TOBPieceDemPrix,TOBArticleDemprix,TOBDetailDemprix,TOBFournDemprix,NumLigTmp);
  end;
  // Recalcul des elements modifiés ainsi que des besoins en achat
  for Indice := 0 to TOBArticleDemPrix.detail.count -1 do
  begin
    if (TOBArticleDemPrix.detail[Indice].getString('UPDATED')='X') or
	     (TOBArticleDemPrix.detail[Indice].getString('NEW')='X') then
    begin
      MiseAjourInfosArticle (TOBArticles,TOBArticleDemprix.detail[Indice],TOBDetailDemprix,TOBFournDemprix);
    end;
  end;
  //
end;

procedure ConstitueIntermFromSelected (form : TForm; TOBArticles,TOBpieceDemPrix,TOBArticleDemprix,TOBFournDemprix,TOBDetailDemprix : TOB);
var XX : Tffacture;
		indice : integer;
		TOBpiece ,TOBOuvrage,TOBL : TOB;
    NumLigTmp : integer;
begin
  NumLigTmp := 0;
	XX := form as Tffacture;
  TOBPiece := XX.LaPieceCourante;
  TOBOuvrage := XX.TheTOBOuvrage;
  for Indice := 1 to XX.GS.RowCount -1 do
  begin
    if XX.GS.IsSelected(Indice) then
    begin
      TOBL := GetTOBLigne(TOBpiece,Indice);
      ConstitueLigneDemPrix (TOBArticles,TOBpiece,TOBL,TOBOuvrage,TOBPieceDemPrix,TOBArticleDemprix,TOBDetailDemprix,TOBFournDemprix,NumLigTmp);
    end;
  end;
  // Recalcul des elements modifiés ainsi que des besoins en achat
  for Indice := 0 to TOBArticleDemPrix.detail.count -1 do
  begin
    if (TOBArticleDemPrix.detail[Indice].getString('UPDATED')='X') or
    	 (TOBArticleDemPrix.detail[Indice].getString('NEW')='X') then
    begin
      MiseAjourInfosArticle (TOBArticles,TOBArticleDemprix.detail[Indice],TOBDetailDemprix,TOBFournDemprix);
    end;
  end;
  //
end;



{***********A.G.L.***********************************************
Auteur  ...... : LS
Créé le ...... : 09/02/2012
Modifié le ... :   /  /
Description .. : Appel de l'écran d'affectation des lignes aux demande de 
Suite ........ : prix
Mots clefs ... : 
*****************************************************************}
procedure AffecteLignesDemPrix (TOBpiece,TOBPieceDemprix,TOBArticleDemPrix,TOBDetailDemPrix,TOBFournDemPrix : TOB);
var UneTOB : TOB;
begin
	UneTOB := TOB.Create ('LES PARAMS',nil,-1);
  TRY
  	UneTOB.AddChampSupValeur('NATUREPIECEG',TOBPiece.getString('GP_NATUREPIECEG'));
  	UneTOB.AddChampSupValeur('SOUCHE',TOBPiece.getString('GP_SOUCHE'));
  	UneTOB.AddChampSupValeur('NUMERO',TOBPiece.getString('GP_NUMERO'));
  	UneTOB.AddChampSupValeur('INDICEG',TOBPiece.getString('GP_INDICEG'));
    //
    UneTOB.data := TOBArticleDemPrix;
    TOBArticleDemPrix.data := TOBDetailDemPrix;
    TOBDetailDemprix.data := TOBFournDemPrix;
    TOBFournDemPrix.data := TOBPieceDemPrix;
    TOBPieceDemPrix.data := TOBPiece;
    TheTOB := UneTOB;
    AGLLanceFiche('BTP','BTARTSELDEMPRIX','','','ACTION=MODIFICATION');
  FINALLY
    TheTOB := nil;
    UneTOB.free;
  END;
end;

procedure  VireArticleDemPrix (TOBArtdemPrix,TOBDetailDemPrix,TOBFournDemPrix : TOB);
var TOBD : TOB;
begin
	TOBD := TOBDetailDemPrix.findFirst(['UNIQUE','UNIQUELIG'],
  												 [TOBArtdemPrix.getInteger('BDP_UNIQUE'),TOBArtdemPrix.getInteger('BDP_UNIQUELIG')],false);
	if TOBD <> nil then TOBD.free;
	TOBD := TOBFournDemPrix.findFirst(['UNIQUE','UNIQUELIG'],
  												 [TOBArtdemPrix.getInteger('BDP_UNIQUE'),TOBArtdemPrix.getInteger('BDP_UNIQUELIG')],false);
	if TOBD <> nil then TOBD.free;
end;

procedure NettoieLignesDemprix (TOBArticleDemprix,TOBDetailDemPrix,TOBFournDemprix : TOB);
var indice : integer;
begin
	Indice := 0;
  Repeat
    if TOBArticledemPrix.detail[Indice].GetInteger('BDP_UNIQUE')=0 then
    begin
			VireArticleDemPrix (TOBArticledemPrix.detail[Indice],TOBDetailDemPrix,TOBFournDemPrix);
      TOBArticledemPrix.detail[Indice].free;
    end else break;
  Until Indice >= TOBArticleDemprix.detail.count;
end;

procedure TraiteDemandePrix (form : TForm);
var XX : Tffacture;
    cledoc : r_cledoc;
    TOBPieceDemprix,TOBArticleDemPrix,TOBFournDemprix,TOBDetailDemPrix,TOBPiece,TOBArticles : TOB;
    Indice : integer;
    ZeronewLigne,DejaTransfered : boolean;
begin

    XX := form as Tffacture;

    if (XX.TheArticleDemPrix = nil) {or (XX.TheArticleDemPrix.detail.count = 0 )} then
    begin
//      PgiInfo ('Aucune ligne article à traiter en demandes de prix');
      exit;
    end;

    ZeronewLigne := true;
    DejaTransfered := false;
    //
    TOBPiece := XX.LaPieceCourante;
    TOBPieceDemprix  := XX.ThePieceDemPrix;
    TOBArticleDemprix := XX.TheArticleDemPrix;
    TOBFournDemprix := XX.TheFournDemPrix;
    TOBDetailDemprix := XX.TheDetailDemPrix;
    TOBArticles := XX.TheTOBArticles;
    //
    Cledoc := TOB2CleDoc(TOBPiece);
    //
    //
    TRY
      if (XX.GS.nbSelected > 0) then
      begin
        ConstitueIntermFromSelected (form,TOBArticles,TOBpieceDemprix,TOBArticleDemprix,TOBFournDemprix,TOBDetailDemprix);
      end else
      begin
        ConstitueIntermFromPiece (form,TOBArticles,TOBpieceDemprix,TOBArticleDemprix,TOBFournDemprix,TOBDetailDemprix);
      end;
      //
      TOBArticleDemPrix.detail.sort ('BDP_UNIQUE;BDP_NUMLIGNE');
      if TOBArticledemprix.detail.count = 0 then
      begin
	      PgiInfo ('Aucune ligne article à traiter en demande de prix');
  	    exit;
      end;
      // Demande d'affectation des lignes aux demandes de prix si nécéssaire
      if TOBArticleDemPrix.detail[0].getInteger('BDP_UNIQUE')=0 then
      begin
        ZeronewLigne := false;
        AffecteLignesDemPrix (TOBpiece,TOBPieceDemprix,TOBArticleDemPrix,TOBdetailDemPrix,TOBFournDemprix);
        // nettoyage des lignes non affectées à une demande de prix
        NettoieLignesDemprix (TOBArticleDemprix,TOBDetailDemPrix,TOBFournDemprix);
      end;
		  NumeroteLig (TOBArticleDemPrix);

      for Indice := 0 to TOBPieceDemprix.detail.count -1 do
      begin
        if TOBpieceDemPrix.detail[Indice].getString('BPP_ENVOYE')='X' then DejaTransfered := true;
      end;
      if DejaTransfered then
      begin
        PgiInfo ('Les demandes de prix mises à jour ont été déjà envoyées ou préparées#13#10 Vous devez les renvoyer à vos fournisseurs')
      end else if ZeroNewLigne then
      begin
        PGIInfo ('Les informations ont été mises à jour dans les demandes de prix');
      end;
    FINALLY
      XX.GS.AllSelected := false;
    END;
end;

function ChargeDemande (cledoc : r_cledoc ; Numero : integer; TOBPieceDemPrix,TOBArticleDemprix, TOBFournDemPrix, TOBdetailDemPrix : TOB) : boolean;
var QQ : TQuery;
		TableChoixCod,TableCommun : string;
    SQL : string;
    TLocal : TOB;
begin
  result := false;
  Tlocal := TOB.Create ('LES DETAILS',nil,-1);
  TableChoixCod := 'CHOIXCOD';
  TableCommun := 'COMMUN';
  TRY
    //
    SQL := 'SELECT * FROM PIECEDEMPRIX WHERE '+WherePiece(cledoc,ttdPieceDemPrix,false);
    if Numero <> 0 then SQL := SQL + ' AND BPP_UNIQUE='+InttoStr(Numero);
    QQ := OpenSql (SQL,true,-1,'',true);
    if not QQ.eof then
    begin
      TOBPieceDemPrix.LoadDetailDB('PIECEDEMPRIX','','',QQ,false);
      result := true;
    end;
    ferme (QQ);
    //
    if not result then exit; // pas la peine de continuer s'il n'y a rien a charger
    SQl := 'SELECT ARTICLEDEMPRIX.*,GA_CODEARTICLE AS CODEARTICLE,'+
           '(SELECT CO_LIBELLE FROM '+TableCommun+' WHERE CO_TYPE="TYA" AND CO_CODE=GA_TYPEARTICLE) AS LIBTYPEARTICLE,'+
           'GA_FAMILLENIV1 AS FAMILLENIV1,GA_FAMILLENIV2 AS FAMILLENIV2,GA_FAMILLENIV3 AS FAMILLENIV3, '+
           '(SELECT CC_LIBELLE FROM '+TableChoixCod+' WHERE CC_TYPE="FN1" AND CC_CODE=GA_FAMILLENIV1) AS LIBFAMILLENIV1,'+
           '(SELECT CC_LIBELLE FROM '+TableChoixCod+' WHERE CC_TYPE="FN2" AND CC_CODE=GA_FAMILLENIV2) AS LIBFAMILLENIV2,'+
           '(SELECT CC_LIBELLE FROM '+TableChoixCod+' WHERE CC_TYPE="FN3" AND CC_CODE=GA_FAMILLENIV3) AS LIBFAMILLENIV3 '+
           'FROM ARTICLEDEMPRIX '+
           'LEFT JOIN ARTICLE ON GA_ARTICLE=BDP_ARTICLE '+
           'WHERE ' + WherePiece (cledoc,TTdArticleDemPrix,true);
    if Numero <> 0 then SQl := Sql + ' AND BDP_UNIQUE='+IntToStr(Numero);
    QQ := OpenSql (Sql	,true,-1,'',true);
    if not QQ.eof then
    begin
      TOBArticleDemprix.LoadDetailDB('ARTICLEDEMPRIX','','',QQ,false);
    end else result := false;
    ferme (QQ);
    if not result then exit;
    //
    SQL := 'SELECT * FROM DETAILDEMPRIX WHERE ' + WherePiece (cledoc,TtdDetailDemPrix,true);
    if Numero <> 0 then SQL := SQL + ' AND BD0_UNIQUE='+IntToStr(Numero);
    QQ := OpenSql (SQL,true,-1,'',true);
    if not QQ.eof then
    begin
      Tlocal.LoadDetailDB('DETAILDEMPRIX','','',QQ,false);
    end else result := false;
    ferme (QQ);
    if not result then exit;
    OrganiseDataDetail (Tlocal,TOBDetailDemprix);
    Tlocal.ClearDetail;
    //
    SQL := 'SELECT * FROM FOURLIGDEMPRIX WHERE ' + WherePiece (cledoc,TtdFournDemprix,true);
    if Numero <> 0 then SQL := SQL + ' AND BD1_UNIQUE='+IntToStr(Numero);
    QQ := OpenSql (SQL,true,-1,'',true);
    if not QQ.eof then
    begin
      Tlocal.LoadDetailDB('FOURLIGDEMPRIX','','',QQ,false);
    end else result := false;
    ferme (QQ);
    OrganiseDataFoun (Tlocal,TOBFournDemPrix);
  FINALLY
    Tlocal.free;
  END;
end;


{***********A.G.L.***********************************************
Auteur  ...... : LS
Créé le ...... : 24/01/2012
Modifié le ... :   /  /
Description .. : Function permettant de récupérer un numéro de demande
Suite ........ : par sélection d'une existante (via le mul)
Mots clefs ... :
*****************************************************************}
function GetNumDemande (cledoc : r_cledoc ) : integer;
var valretour : string;
begin
	result := -1;
  valretour := AGLLanceFiche('BTP','BTPIECDDEPRIX_MUL','BPP_NATUREPIECEG='+cledoc.NaturePiece +';'+
  																				'BPP_SOUCHE='+Cledoc.Souche+';'+
                                          'BPP_NUMERO='+IntToStr(cledoc.NumeroPiece)+';'+
                                          'BPP_INDICEG='+IntToStr(cledoc.Indice),'','ACTION=MODIFICATION;SELECTION');
  if valretour <> '' then result := StrToInt(valRetour);
end;

{***********A.G.L.***********************************************
Auteur  ...... : LS
Créé le ...... : 24/01/2012
Modifié le ... :   /  /
Description .. : Function permettant de verifier que l'article ou la ligne n'est
Suite ........ : pas deja presente dans une demande de prix (insertion)
Mots clefs ... :
*****************************************************************}
function IsArticleInserable (cledoc : r_cledoc; Article: string; LigneNumOrdre : integer; UniqueBlo : integer) : boolean;
begin
	result := false;

end;

{***********A.G.L.***********************************************
Auteur  ...... : LS
Créé le ...... : 24/01/2012
Modifié le ... : 26/01/2012
Description .. : Fonction permettant de vérifier s'il existe une demande de
Suite ........ : prix sur la pièce en cours
Mots clefs ... : 
*****************************************************************}
function IsExisteDemandes (cledoc : r_cledoc) : boolean;
begin
  result := false;
  if Cledoc.NumeroPiece <> 0 then exit;
  result := ExisteSql ('SELECT BDP_UNIQUE FROM PIECEDEMPRIX WHERE '+WherePiece (cledoc,ttdPieceDemPrix,false));
end;


{***********A.G.L.***********************************************
Auteur  ...... : FV
Créé le ...... : 26/01/2012
Modifié le ... :   /  /    
Description .. :
Suite ........ : des tables associées à une demande de prix.
Mots clefs ... :
*****************************************************************}
//Procedure permettant l'automatisation de la création des demandes de prix
Procedure CreateDemandePrix(Cledoc : R_Cledoc; Unique : Integer; Action : String); OverLoad;
var Argument  : String;
begin

  if  (Cledoc.NaturePiece = '')   Or
      (Cledoc.Souche       = '')  Or
      (Cledoc.NumeroPiece  = 0)   Then Exit;;

  Argument := 'NATUREPIECEG='       + Cledoc.NaturePiece;
  Argument := Argument + ';SOUCHE=' + Cledoc.Souche;
  Argument := Argument + ';NUMERO=' + IntToStr(Cledoc.NumeroPiece);
  Argument := Argument + ';INDICEG='+ IntToStr(Cledoc.Indice);
  Argument := Argument + ';UNIQUE=' + IntToStr(Unique);
  Argument := Argument + ';ACTION=' + Action;

  AGLLanceFiche('BTP','BTCREATEDEMPRIX','','',Argument);

end;

Procedure CreateDemandePrix(TobDemPrix : Tob; Unique : Integer); OverLoad;
var argument  : String;
    Action    : String;
begin

  TRY
    if unique = 0  then
      Action := 'CREATION'
    else
      Action := 'MODIFICATION';

    Argument := 'NATUREPIECEG='       + TobDemPrix.GetString('NATUREPIECEG');
    Argument := Argument + ';SOUCHE=' + TobDemPrix.GetString('SOUCHE');
    Argument := Argument + ';NUMERO=' + TobDemPrix.GetString('NUMERO');
    Argument := Argument + ';INDICEG='+ TobDemPrix.GetString('INDICEG');
    Argument := Argument + ';UNIQUE=' + IntToStr(Unique);
    Argument := Argument + ';ACTION=' + Action;

    TheTOB := TOBDemPrix;

  	AGLLanceFiche('BTP','BTCREATEDEMPRIX','','',Argument);

    Thetob := Nil;

  FINALLY
  END;

end;

//Procedure permettant l'automatisation de la suppression
Function SupprimeDdePrix(Cledoc : R_Cledoc; Unique : integer; NomTable, Prefixe : String; ttd : T_TableDoc; FromSaisie : Boolean=false) : Integer; Overload;
var StSQL     : String;
begin
  StSQL := 'DELETE ' + NomTable +' WHERE ' + Wherepiece(Cledoc, ttd, true);
  if Unique <> 0 then StSQL := StSQL + ' AND ' + Prefixe +'_UNIQUE='+ IntToStr(Unique);

  result := Executesql(StSQL);
end;

function SupprimeDdePrix(TobToDelete : TOB; Unique : Integer; FromSaisie : Boolean=false) : Integer; Overload;
Var Maxi      : Integer;
    TobDelete : Tob;
    Prefixe,ChampToFind   : String;
    iInc      : Integer;
begin

  Result := -1;
  Prefixe := TableToPrefixe(TobToDelete.Detail[0].NomTable);
  if FromSaisie then
  begin
  	if Prefixe = '' then ChampToFind := 'UNIQUE' else ChampToFind := Prefixe + '_UNIQUE';
  	TobDelete := TobToDelete.FindFirst([ChampToFind],[Unique], True);
  	repeat
      if TobDelete <> nil then TobDelete.Free;
  		TobDelete := TobToDelete.FindNext([ChampToFind],[Unique], True);
    until TobDelete = nil;
  end else
  begin
    Maxi := TobToDelete.detail.count;

    if Maxi > 0 then
    begin
      iInc := 0;
      if prefixe = '' then Prefixe := TableToPrefixe(TobToDelete.Detail[0].NomTable);
      if prefixe <> '' then
      begin
        TobDelete := TobToDelete.FindFirst([Prefixe + '_UNIQUE'],[Unique], True);
        Repeat
          if Tobdelete = nil then
            break
          else
            if prefixe = 'BDP' then
              Tobdelete.setinteger(Prefixe+'_UNIQUE', 0)
            else
              TobDelete.free;
          TobDelete := TobToDelete.FindNext([Prefixe + '_UNIQUE'],[Unique], True);
          Result := 1;
          inc(iInc);
        until iInc > Maxi;
      end;
    end else if Maxi = 0 then
    begin
      TobToDelete.free;
      Result := 1;
    end else
    begin
      exit;
    end;
  end;
end;

Procedure ExporteDocumentToExcel(Cledoc : R_cledoc;Unique : Integer; Libelle : String); Overload;
var ToBPiece  : TOB;
    TobArtFrs : TOB;
    TobFrs    : TOB;
    TobArt    : TOB;
    //
    StSQL     : String;
    //
    QQ        : TQuery;
begin

  TobPiece := Tob.create('PIECES', Nil, -1);
  TobArtFrs:= Tob.Create('ARTFRS', Nil, -1);
  TobArt   := Tob.Create('ARTDDP', Nil, -1);
  TobFrs   := Tob.Create('FRSDDP', Nil, -1);

  //lecture des éléments propres à la pièce
  StSQL := 'SELECT * FROM PIECE WHERE ' + wherePiece(Cledoc, ttdPiece, false);

  QQ := OpenSQL(StSQl, False);

  if not QQ.eof then
  begin
    //chargement TobPiece
    TobPiece.SelectDB('', QQ);
    TobPiece.AddChampSupValeur('LIBELLE', Libelle);
    ferme(QQ);

    //chargement des informations fournisseur demande de prix...
    StSql := 'SELECT * FROM FOURLIGDEMPRIX LEFT JOIN ARTICLEDEMPRIX';
    StSQL := StSQL + '   ON BD1_NATUREPIECEG=BDP_NATUREPIECEG AND BD1_SOUCHE=BDP_SOUCHE';
    StSQL := StSQL + '  AND BD1_NUMERO=BDP_NUMERO AND BD1_UNIQUE=BDP_UNIQUE';
    StSQL := StSQL + '  AND BD1_UNIQUELIG=BDP_UNIQUELIG ';
    StSQL := StSQL + 'WHERE ' + wherePiece(Cledoc, TtdFournDemprix, false);
    StSQL := StSQL + '  AND BD1_UNIQUE=' + IntToStr(Unique) + ' ';
    StSQL := StSQL + 'ORDER BY BD1_UNIQUE, BD1_TIERS';
    QQ := OpensQL(StSQL, False);
    if Not QQ.eof then
    begin
      TobArtFrs.LoadDetailDB('ARTFOURN','','',QQ, False);
      CreateTobFrsInterm(TobArtFrs, TobFrs, TobArt);
    end;
  end;

  Ferme(QQ);

  //chargement en tob de l'entete du document à partir de Piece
  if ExportToExcel(cledoc,TobPiece, TobFrs, TobArtFrs, Unique) then
  begin
    EnregistreGenereDemPrix(cledoc,Unique);
  end;

  //Libération des TOB
  FreeAndNil(TobPiece);
  FreeAndNil(TobArtFrs);
  FreeAndNil(TOBArt);
  FreeAndNil(TobFrs);

end;

Procedure ExporteDocumentToExcel(TobPiece, TobArtFrs, TobFrsDemPx : TOB; Unique : Integer; Libelle : String); Overload;
var TobFrs    : TOB;
    TobArt    : TOB;
    fCleDoc   : R_Cledoc;
    TOBTempo  : TOB;
    TobDest   : TOB;
    TobDD     : TOB;
    iInd      : Integer;
    JInd      : Integer;

begin

  TobArt   := Tob.Create('ARTDDP', Nil, -1);
  TobFrs   := Tob.Create('FRSDDP', Nil, -1);
  TobDest  := Tob.create('DESTINATION', Nil, -1);

  fCleDoc.NaturePiece := TobPiece.GetString('GP_NATUREPIECEG');
  fCleDoc.Souche      := TobPiece.GetString('GP_SOUCHE');
  fCleDoc.NumeroPiece := TobPiece.GetInteger('GP_NUMERO');
  fCleDoc.Indice      := TobPiece.GetInteger('GP_INDICEG');

  TobPiece.AddChampSupValeur('LIBELLE', Libelle);

  for iInd := 0 to TobArtFrs.Detail.Count - 1 do
  begin
    TobTempo := FindFournDemPrix(TobArtFrs.detail[iInd],TobFrsDemPx);
    if TobTempo <> Nil then
    begin
      for Jind := 0 to TobTempo.detail.count -1 do
      begin
        TobDD := Tob.create('ARTDEMPRIX', Tobdest, -1);
        TobDD.Dupliquer(TobArtFrs.detail[iInd], False, True);
        CopieChamps (TOBDD,TOBTempo.detail[Jind]);
      end;
    end;
  end;

  if TobDest.detail.count <> 0 then
  begin
    TobDest.detail.Sort('BD1_UNIQUE;BD1_TIERS');
    CreateTobFrsInterm(TobDest, TobFrs, TobArt);
    if ExportToExcel(fCleDoc, TobPiece, TobFrs, TobDest, Unique) then
    begin
      For iInd := 0 to TobArtFrs.detail.count -1 do
      begin
        TobArtFrs.detail[iInd].PutValue('BPP_ENVOYE', 'X');
      end;
    end;
  end
  else
  begin
    PgiInfo ('Aucun fournisseur n''est associé aux articles de la demande de prix');
  end;


  FreeAndNil(TobArt);
  FreeAndNil(TobFrs);
  FreeAndNil(TobDest);

end;


Procedure CreateTobFrsInterm(TobArtFrs, TOBFrs, TOBArt : TOB);
Var Fournisseur : String;
    TobDetFrs   : Tob;
    iInd        : Integer;
begin

  For iInd := 0 to TobArtFrs.detail.count -1 do
  begin
    if Fournisseur <> TobArtFrs.Detail[iInd].Getstring('BD1_TIERS') then
    begin
      Fournisseur := TobArtFrs.Detail[iInd].Getstring('BD1_TIERS');
      TobDetFrs := Tob.create('FOURNISSEUR', TobFrs, -1);
      TobDetFrs.AddChampSupValeur('BD1_TIERS',TobArtFrs.Detail[iInd].Getstring('BD1_TIERS'));
      TobDetFrs.AddChampSupValeur('BD1_QTEACH',TobArtFrs.Detail[iInd].Getstring('BD1_QTEACH'));
      TobDetFrs.AddChampSupValeur('BD1_QUALIFUNITEACH',TobArtFrs.Detail[iInd].Getstring('BD1_QUALIFUNITEACH'));
      TobDetFrs.AddChampSupValeur('BD1_NATUREPIECEG',TobArtFrs.Detail[iInd].Getstring('BD1_NATUREPIECEG'));
      TobDetFrs.AddChampSupValeur('BD1_SOUCHE',TobArtFrs.Detail[iInd].Getstring('BD1_SOUCHE'));
      TobDetFrs.AddChampSupValeur('BD1_NUMERO',TobArtFrs.Detail[iInd].Getstring('BD1_NUMERO'));
      TobDetFrs.AddChampSupValeur('BD1_INDICEG',TobArtFrs.Detail[iInd].Getstring('BD1_INDICEG'));
      TobDetFrs.AddChampSupValeur('BD1_UNIQUE',TobArtFrs.Detail[iInd].Getstring('BD1_UNIQUE'));
      TobDetFrs.AddChampSupValeur('BD1_UNIQUELIG',TobArtFrs.Detail[iInd].Getstring('BD1_UNIQUELIG'));
    end;
    CreateTobArtInterm(TobArtFrs.Detail[iInd], TobArt, Fournisseur);
  end;

end;

Procedure CreateTobArtInterm(TobArtFrs, TOBArt : TOB; Fournisseur : String);
var iInd : Integer;
    TobDetArt : Tob;
begin

  For iInd := 0 to TobArtFrs.detail.count -1 do
  begin
    if Fournisseur <> TobArtFrs.Getstring('BD1_TIERS') then
    begin
      TobDetArt := Tob.create('ARTICLE', TobArt, -1);
      TobDetArt.AddChampSupValeur('BDP_NATUREPIECEG', TobArtFrs.GetString('BD1_NATUREPIECEG'));
      TobDetArt.AddChampSupValeur('BDP_SOUCHE', TobArtFrs.GetString('BD1_SOUCHE'));
      TobDetArt.AddChampSupValeur('BDP_NUMERO', TobArtFrs.GetString('BD1_NUMERO'));
      TobDetArt.AddChampSupValeur('BDP_INDICEG', TobArtFrs.GetString('BD1_INDICEG'));
      TobDetArt.AddChampSupValeur('BDP_UNIQUE', TobArtFrs.GetString('BD1_UNIQUE'));
      TobDetArt.AddChampSupValeur('BDP_NUMLIGNE', TobArtFrs.GetString('BD1_NUMLIG'));
      TobDetArt.AddChampSupValeur('BDP_TYPELIGNE', TobArtFrs.GetString('BDP_TYPELIGNE'));
      TobDetArt.AddChampSupValeur('BDP_ARTICLE', TobArtFrs.GetString('BDP_ARTICLE'));
      TobDetArt.AddChampSupValeur('BDP_LIBELLE', TobArtFrs.GetString('BDP_LIBELLE'));
      TobDetArt.AddChampSupValeur('BDP_QTEACH', TobArtFrs.GetString('BD1_QTEACH'));
      TobDetArt.AddChampSupValeur('BDP_QUALIFUNITEACH', TobArtFrs.GetString('BD1_QUALIFUNITEACH'));
    end;
  end;

end;

function ExportToExcel(fCleDoc: r_cledoc; TobPiece, TobFrs, TobArtFrs : TOB; Unique : Integer) : boolean;
var iInd        : Integer;
    Fournisseur : String;
    StWhere     : String;
    StChamp     : String;
    TobExcel    : Tob;
    TobLigne    : Tob;
begin
  result := true;

  //Création tobExcel entete de fichier XLS
  For iInd := 0 to TobFrs.Detail.count -1 do
  begin
    //Création Tob transfert XLS
    TobExcel := Tob.create('EXCEL', nil, -1);
    Fournisseur := TobFrs.Detail[iInd].Getstring('BD1_TIERS');
    StChamp := '*';
    StWhere := 'T_NATUREAUXI="FOU" AND T_TIERS="'+ Fournisseur + '"';
    if GetTobTiers(StChamp, StWhere, TobFrs.detail[iInd]) then
    Begin
      //création entete de transferts
      TobExcel.Detail[0] := AddEntete(TobPiece, TobExcel, Unique);
      //chargement des informations fournisseur
      ChargeAdresseFrs(TobExcel.detail[0], TobFrs.Detail[iInd]);
      //chargement des lignes pour envoie dans excel
      TobLigne := AddLigExcel(Fournisseur, Unique, iInd, TobArtFrs, TobExcel);
      //Envoi document excel
      result := (result and ExportDocumentViaTOB(TobExcel, fCledoc, Unique, Fournisseur));
    end;
    ToBExcel.free;
  end;

end;

function AddEntete (TOBP, TobExcel : TOB; Unique : Integer ): TOB;
var Infopiece : string;
begin

	result := TOB.Create ('UNDOCUMENT', TobExcel, -1);

  InfoPiece := 'Demande de Prix N°' + IntToStr(Unique);
  InfoPiece := InfoPiece + ' ' + Rechdom('GCNATUREPIECEG',TOBP.GetValue('GP_NATUREPIECEG'),false)+' N°'+inttostr(TOBP.GetValue('GP_NUMERO'));

	result.AddChampSupValeur('REFEXTERNE' ,TOBP.GetValue('GP_REFEXTERNE'));
	result.AddChampSupValeur('REFINTERNE' ,TOBP.GetValue('GP_REFINTERNE'));
	result.AddChampSupValeur('CODETIERS'  ,TOBP.GetValue('GP_TIERS'));
	result.AddChampSupValeur('AFFAIRE'    ,BTPCodeAffaireAffiche(TobP.GetValue('GP_AFFAIRE')));
	result.AddChampSupValeur('EXPEDITEUR' ,getparamSocSecur('SO_SIRET',''));
	result.AddChampSupValeur('INFOPIECE'  ,infopiece);
	result.AddChampSupValeur('IDPIECE'    ,TOBP.GetValue('GP_NATUREPIECEG')+':'+
  																		   TOBP.GetValue('GP_SOUCHE')+':'+
                                         inttostr(TOBP.GetValue('GP_NUMERO'))+':'+
                                         IntToStr(TOBP.GetValue('GP_INDICEG'))
                                         );
	result.AddChampSupValeur('DATEPIECE'  ,TOBP.GetValue('GP_DATEPIECE'));
	result.AddChampSupValeur('CODECLIENT' ,TOBP.GetValue('GP_TIERS'));
  result.AddChampSupValeur('QUALIFMAIL' , 'DDE');
  Result.AddChampSupValeur('BPP_LIBELLE',TOBP.GetValue('LIBELLE'));


end;

Procedure ChargeAdresseFrs (TOBentete, Tobfourn : TOB);
Begin

  TOBentete.AddChampSupValeur ('TYPECONTACT',TOBFourn.detail[0].GetString('T_NATUREAUXI'));
  TobEntete.AddChampSupValeur ('CODEFOURNISSEUR', TOBFourn.Getstring('BD1_TIERS'));
  TobEntete.AddChampSupValeur ('AUXILIAIRE', TOBFourn.Detail[0].Getstring('T_AUXILIAIRE'));
  TOBentete.AddChampSupValeur ('FACLIBELLE', TOBFourn.detail[0].GetString('T_LIBELLE'));
  TOBentete.AddChampSupValeur ('FACADRESSE1',TOBFourn.detail[0].GetString('T_ADRESSE1'));
  TOBentete.AddChampSupValeur ('FACADRESSE2',TOBFourn.detail[0].GetString('T_ADRESSE2'));
  TOBentete.AddChampSupValeur ('FACADRESSE3',TOBFourn.detail[0].GetString('T_ADRESSE3'));
  TOBentete.AddChampSupValeur ('FACCODEPOST',TOBFourn.detail[0].GetString('T_CODEPOSTAL'));
  TOBentete.AddChampSupValeur ('FACVILLE',   TOBFourn.detail[0].GetString('T_VILLE'));
  TOBentete.AddChampSupValeur ('FACCPVILLE', TOBFourn.detail[0].GetString('T_CODEPOSTAL')+' '+TOBFourn.GetString('T_VILLE'));
  TOBentete.AddChampSupValeur ('FACPAYS',    TOBFourn.detail[0].GetString('T_PAYS'));
  TOBentete.AddChampSupValeur ('FACNIF',     TOBFourn.detail[0].GetString('T_NIF'));
  TOBentete.AddChampSupValeur ('FACADRMAIL', TOBFourn.detail[0].GetString('T_EMAIL'));
  TOBentete.AddChampSupValeur ('REFFOURNISSEUR',TOBFourn.detail[0].GetString('T_TIERS'));

end;

function AddLigExcel (Fournisseur : String; Unique, NumLig : integer; TOBA, TobExcel : TOB): TOB;
Var iInd : Integer;
    RefUnique : String;
    indice : integer;
begin

  result := nil;

  for iInd := 0 to TOBA.Detail.count -1 do
  begin
    if Fournisseur <> TOBA.Detail[iInd].Getstring('BD1_TIERS') then continue;
    result := TOB.Create ('ONELIGNE', TobExcel.Detail[0], -1);

    if TOBA.Detail[iInd].GetString('BDP_indice') = '' then
      indice := 0
    else
      indice := TobA.Detail[iInd].GetInteger('BDP_INDICE');

    Refunique := TOBA.Detail[iInd].GetString('BDP_NATUREPIECEG') + ';';
    RefUnique := RefUnique + TOBA.Detail[iInd].GetString('BDP_SOUCHE') + ';';
    RefUnique := RefUnique + TOBA.Detail[iInd].GetString('BDP_NUMERO') + ';';
    RefUnique := RefUnique + IntToSTr(Indice) + ';';
    RefUnique := RefUnique + IntToStr(Unique);
    Refunique := RefUnique + ';' +  IntToStr(TOBA.Detail[iInd].GetValue('BDP_UNIQUELIG'));

    result.AddChampSupValeur('REFLIGNE',     Refunique);
    result.AddChampSupValeur('TYPELIGNE',    TOBA.Detail[iInd].GetString('BDP_TYPELIGNE'));
    result.AddChampSupValeur('REFARTLIGNE',  TOBA.Detail[iInd].GetString('BDP_ARTICLE'));
    result.AddChampSupValeur('LIBELLELIGNE', TOBA.Detail[iInd].GetString('BDP_LIBELLE'));
    result.AddChampSupValeur('QTELIGNE',     TOBA.Detail[iInd].GetValue('BD1_QTEACH'));
    result.AddChampSupValeur('UNITELIGNE',   TOBA.detail[iInd].GetString('BD1_QUALIFUNITEACH'));
    result.AddChampSupValeur('PULIGNE',      0);
    result.AddChampSupValeur('NBJOUR',       0);
    result.AddChampSupValeur('NIVEAUIMBRIC', 0);
  end;

end;

Procedure ChargementGrilleWithTob(TOBAAfficher : TOB; Grille : THGrid; LibColonne : String);
var indice : integer;
begin

  Grille.videPile(false);    

  if (TOBAAfficher=nil) or (TOBAAfficher.Detail.Count = 0) then
    Grille.rowcount  := 2
  else
    Grille.rowcount := TOBAAfficher.Detail.Count+1;

  if TOBAAfficher <> nil then
  begin
    for Indice := 0 to TOBAAfficher.detail.count -1 do
    begin
      TOBAAfficher.detail[Indice].PutLigneGrid (Grille, Indice+1,false, false, LibColonne);
    end;
  end;
  Grille.Row := 1;
  //

end;

Procedure ChargeComboWithTob(Combo : THValComboBox; TobCharge : TOB; Champ, Valeur : String);
var iInd    : Integer;
begin

  Combo.DataType := '';
  Combo.Clear;
  combo.values.Clear;
  combo.Items.Clear;
  For iInd := 0 to TobCharge.detail.count -1 do
  begin
    Combo.Values.Add(TobCharge.Detail[iInd].GetString(Valeur));
    Combo.Items.Add(TobCharge.Detail[iInd].GetString(Champ));
  end;

  Combo.ItemIndex := 0;

end;


{***********A.G.L.***********************************************
Auteur  ...... : FV
Créé le ...... : 02/03/2012
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure CreatTOBVide (TOBDest: TOB) ;
var TPADM : TOB;
begin
	TPADM := TOB.Create ('ARTICLEDEMPRIX',TOBDest,-1);
  TPADM.AddChampSupValeur('TYPEARTICLE','');
  TPADM.AddChampSupValeur('LIBTYPEARTICLE','');
  TPADM.AddChampSupValeur('CODEARTICLE','');
  TPADM.AddChampSupValeur('FAMILLENIV1','');
  TPADM.AddChampSupValeur('FAMILLENIV2','');
  TPADM.AddChampSupValeur('FAMILLENIV3','');
  TPADM.AddChampSupValeur('LIBFAMILLENIV1','');
  TPADM.AddChampSupValeur('LIBFAMILLENIV2','');
  TPADM.AddChampSupValeur('LIBFAMILLENIV3','');
  TPADM.AddChampSupValeur('UPDATED','X');
  TPADM.AddChampSupValeur('NEW','X');
end;

procedure SupprimeLigDemPrix (TOBArticles,TOBArticleDemprix,TOBDetailDemPrix,TOBFournDemPrix,TOBPiece,TOBOuvrages,TOBL : TOB);
var TOBA,TOBB,TOBD,TOBPlat : TOB;
		prefixe,TypeArticle : string;
    indice,UniqueBlo : integer;
    MTPA,MTPV : double;
begin
	prefixe := TableToPrefixe(TOBL.NomTable);
	TypeArticle := TOBL.GetString(prefixe+'_TYPEARTICLE');
  if ((TypeArticle='OUV') OR (TypeArticle='ARP'))  then
  begin
    TOBPlat := TOB.Create ('OUVRAGE A PLAT',nil,-1);
    TRY
      if (prefixe = 'GL') and (TOBL.GetValue('GL_INDICENOMEN')>0) then
      begin
        MiseAplatouv (TOBPiece,TOBL,TOBOuvrages,TOBPLat,true,true,true,false);
        for indice := 0 to TOBPlat.detail.count -1 do
        begin
        	SupprimeLigDemPrix (TOBArticles,TOBArticleDemprix,TOBDetailDemPrix,TOBFournDemPrix,TOBPiece,TOBOuvrages,TOBPLat.detail[Indice]);
        end;
      end;
    FINALLY
      TOBPlat.free;
    END;
  end else
  begin
  	UniqueBlo := TOBL.GetInteger('UNIQUEBLO');
    TOBA := TOBArticleDemPrix.findFirst(['BDP_ARTICLE'],[TOBL.GetString('GL_ARTICLE')],true);
    if TOBA = nil then exit;
    TOBA.SetString('UPDATED','X');
    //
    TOBB := FindFournDemPrix (TOBA,TOBFournDemPrix); // infos de prix via fournisseurs
    //
    TOBD := FindDetailDemPrix (TOBA,TOBDetailDemPrix);
    if TOBD = nil then exit;
    //
  	TOBA.SetDouble ('BDP_QTEBESOIN',0); // on reinit pour le calcul via le détail
  	TOBA.SetDouble ('BDP_DPA',0); // on reinit pour le calcul via le détail
  	TOBA.SetDouble ('BDP_PUHTDEV',0); // on reinit pour le calcul via le détail
    //
    MtPA := 0;
    MTPV := 0;
    Indice := 0;
    repeat
      if (TOBD.detail[Indice].GetInteger ('BD0_NUMORDRE')=TOBL.GetValue(prefixe+'_NUMORDRE')) and
         (TOBD.detail[Indice].GetInteger ('BD0_UNIQUEBLO')=UniqueBlo) then
      begin
        TOBD.detail[Indice].free;
      end else
      begin
        TOBA.SetDouble ('BDP_QTEBESOIN',TOBA.GetDouble ('BDP_QTEBESOIN')+TOBD.detail[Indice].GetDouble('BD0_QTEBESOIN'));
        MTPA := TOBD.detail[Indice].GetDouble('BD0_QTEBESOIN') * TOBD.detail[Indice].GetDouble('BD0_DPA');
        MTPV := TOBD.detail[Indice].GetDouble('BD0_QTEBESOIN') * TOBD.detail[Indice].GetDouble('BD0_PUHTDEV');
        inc(Indice);
      end;
    until Indice >= TOBD.detail.count;
    //
    if TOBA.GetDouble ('BDP_QTEBESOIN') <> 0 then
    begin
      TOBA.setDouble ('BDP_DPA',Arrondi(MTPA/TOBA.GetDouble ('BDP_QTEBESOIN'),V_PGI.okdecP));
      TOBA.setDouble ('BDP_PUHTDEV',Arrondi(MTPV/TOBA.GetDouble ('BDP_QTEBESOIN'),V_PGI.okdecP));
    end;
    if TOBD.detail.count = 0 then
    begin
      TOBB.free; // libération de l'info article demande de prix
      TOBA.free; // libération de l'article dans la demande de prix
    end else
    begin
      ConstitueBesoinFournisseurs (TOBArticles,TOBB,TOBD,TOBA);
    end;
	end;
end;

function findOccurenceString(St : String; Caract : char) : Integer;
var j        : Integer;
    nbOccurs :integer;
begin

  nbOccurs := 0;              // nombre d'occurence d'un caractère

  Repeat
    j := pos(Caract, St);     // position du carctère dans le texte
    if  j <> 0 Then           // s'il existe
      begin
       delete(St, J, 1);      // on efface le texte
       inc(nbOccurs);         // et on incrémente le nombre d'occurence
      end;
  Until pos(Caract, St) = 0;  // s'il le carctère n'as plus d'occurence on termine la boucle

  Result := NbOccurs;

end;

end.
