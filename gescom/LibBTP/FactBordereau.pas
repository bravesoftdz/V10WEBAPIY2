unit FactBordereau;

interface

uses sysutils,classes,windows,messages,controls,forms,hmsgbox,stdCtrls,clipbrd,nomenUtil,
     HCtrls,SaisUtil,HEnt1,Ent1,EntGC,UtilPGI,UTOB,HTB97,FactUtil,FactComm,Menus,ParamSoc,
{$IFDEF EAGLCLIENT}
  maineagl,
{$ELSE}
  fe_main,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
		BTSELLIGBORD_TOF,
    FactTOB,
    UtilArticle,
    HRichOLE,uEntCommun,UtilTOBPiece;

const MAXITEMS = 1;

type
  TUseBordereau = class
  private
  	fChanged : boolean;
  	frequeteSelect : string;
  	fTOBListBordereaux : TOB;
  	ftiers : string;
    fnatureAuxi : string;
    fAffaire : string;
    fDatePiece : TDateTime;
    fSaContexte : TModeAff;
    fNaturePiece : string;
    FF : Tform;
    POPGS : TPopupMenu;
    MesMenuItem: array[0..MAXITEMS] of TMenuItem;
    fMaxItems : integer;
    fCreatedPop : boolean;
    TheItem : TmenuItem;
    procedure SetAffaire(const Value: string);
    procedure SetDatePiece(const Value: TDateTime);
    procedure SetNatureAuxi(const Value: string);
    procedure SetTiers(const Value: string);
    function IsBordereauExists: boolean;
    procedure ReinitList;
    procedure ReCharge;
    procedure StContexte(const Value: TModeAff);
    procedure SetNaturePiece(const Value: string);
    procedure DefiniMenuPop(Parent: Tform);
		procedure GSAccesBordereau (Sender : Tobject);
    procedure DefiniMenuActif;
  public
  	constructor create (TT : TForm);
    destructor destroy; override;
    //
    property Tiers : string read ftiers write SetTiers;
    property NatureAuxi : string read fNatureAuxi write SetNatureAuxi;
    property Affaire : string read fAffaire write SetAffaire;
    property Date : TDateTime read fDatePiece write SetDatePiece;
    property BordereauExists : boolean read IsBordereauExists;
    property Requete : string read frequeteSelect;
    property Piece : string write SetNaturePiece;
    property contexte : TModeAff read fsacontexte write StContexte;
    //
    procedure AppliqueChange;
    procedure ClearAll;
    function RechArticlesFromBordereau (reference,Naturepiece : string  ) : Boolean;
    Function FindReferenceTiers (reference : string) : TOB;
    function IsExisteLigneBordereau (TOBPiece : TOB) : boolean;
  end;

procedure ReinitLigneBordereau (TOBPiece,TOBOuvrage : TOB; Ligne : integer);
procedure LoadlesSousDetailBordereaux (TOBpiece,TOBMSL,TOBArticles : TOB);
function TraiteOuvrageFromBordereau ( XX: TForm; TOBPiece,TOBArticles,TOBOuvrage,TOBRepart : TOB ; ARow : integer ; LePremier : boolean ;DEV :RDEVISE;TOBOUVB : TOB=nil) : boolean;
function IsLigneFromBordereau (TOBL : TOB) : boolean ;
function IsExisteLigneBordereau(TOBPiece: TOB): boolean;

implementation
uses factouvrage,FactArticle,
//UtilMetres,
LigNomen,FactCalc,facture;

// ---------------------------------------------------
// Partie gerant les bordereaux de prix a la création-Modif
// ---------------------------------------------------
procedure VireOuvrage (TOBL,TOBOUvrage : TOB);
var IndiceNomen : integer;
		TOBO : TOB;
begin
  IndiceNomen := TOBL.getValue('GL_INDICENOMEN');
  TOBO := TOBOUvrage.Detail[IndiceNomen-1];
  if TOBO <> nil then TOBO.PutValue('UTILISE','-');
end;

procedure ReinitLigneBordereau (TOBPiece,TOBOuvrage : TOB; Ligne : integer);
var TOBL : TOB;
begin
	TOBL := TOBPiece.detail[Ligne-1];
  if TOBL.GetvALUE('GL_INDICENOMEN') > 0 then
  begin
  	VireOuvrage (TOBL,TOBOUvrage);
  end;
  TOBL.putValue('GL_CODEARTICLE','');
  TOBL.putValue('GL_REFARTSAISIE','');
  TOBL.putValue('GL_ARTICLE','');
  TOBL.putValue('GL_TYPEARTICLE','');
  TOBL.putValue('GL_INDICENOMEN',0);
  TOBL.putValue('GL_TYPENOMENC','');
  TOBL.putValue('GL_FAMILLENIV1','');
  TOBL.putValue('GL_FAMILLENIV2','');
  TOBL.putValue('GL_FAMILLENIV3','');
  TOBL.putValue('GL_TYPENOMENC','');
  TOBL.putValue('GL_BLOCNOTE','');
  TOBL.putValue('GLC_GETCEDETAIL','-');
  TOBL.putValue('GL_DPA',0);
  TOBL.putValue('GL_DPR',0);
  TOBL.putValue('GL_PMAP',0);
  TOBL.putValue('GL_PMRP',0);
end;
// ---------------------------------------------------

function IsExisteLigneBordereau(TOBPiece: TOB): boolean;
begin
  result := (TOBPiece.findFirst(['GL_BLOQUETARIF'],['X'],true)<>nil);
end;

function IsLigneFromBordereau (TOBL : TOB) : boolean ;
begin
	result := (TOBL.GetValue('GLC_FROMBORDEREAU')='X');
end;

procedure LoadlesSousDetailBordereaux (TOBpiece,TOBMSL,TOBArticles : TOB);
var TOBGroupenomen,TOBLN,TOBNomen,TOBPere,TOBLoc,TOBArt,TOBRef : TOB;
		Q : Tquery;
    cledoc : r_cledoc;
    indice : integer;
    LigPere1,LigPere2,LigPere3,LigPere4,Lig : integer;
    LigneN1,LigneN2,LigneN3,LigneN4,LigneN5: integer;
    LaTOBArticles : TOB;
    DEv : Rdevise;
    TypeArticle,RefArticle,Sql : string;
begin

  TOBGroupeNomen:=TOB.Create('',TOBMSL,-1) ;
  TOBGroupeNomen.AddChampSup('UTILISE',False) ;
  TOBGroupeNomen.PutValue('UTILISE','-') ;
	cledoc := TOB2CleDoc (TOBMSL);
  cledoc.NumLigne  := TOBMSL.GetValue('GL_NUMLIGNE');

	TOBNomen:=TOB.Create('',Nil,-1) ;
  TRY
    Q:=OpenSQL('SELECT * FROM LIGNEOUV WHERE '+WherePiece(CleDoc,ttdOuvrage,true,false)+
               ' AND BLO_NUMLIGNE='+inttostr(Cledoc.numligne)+' ORDER BY BLO_NUMLIGNE,BLO_N1, BLO_N2, BLO_N3, BLO_N4,BLO_N5',True,-1, '', True) ;
    TOBNomen.LoadDetailDB ('LIGNEOUV','','',Q,True,true);
    ferme (Q);
//
    for indice:=0 to TOBNomen.Detail.Count-1 do
    BEGIN
      TOBLN:=TOBNomen.Detail[indice] ;
      LigPere1 := 0; LigneN1 := TOBLN.GetValue('BLO_N1');
      LigPere2 := 0; LigneN2 := TOBLN.GetValue('BLO_N2');
      LigPere3 := 0; LigneN3 := TOBLN.GetValue('BLO_N3');
      LigPere4 := 0; LigneN4 := TOBLN.GetValue('BLO_N4');
      LigneN5 := TOBLN.GetValue('BLO_N5');
      Lig := TOBLN.GetValue('BLO_NUMLIGNE');

      if LigneN5 > 0 then
      begin
        // recherche du pere au niveau 4
        TOBPere:=TOBGroupeNomen.FindFirst(['BLO_NUMLIGNE','BLO_N1','BLO_N2','BLO_N3','BLO_N4','BLO_N5'],[Lig,LigneN1,LigneN2,LigneN3,LigneN4,0],True) ;
      end else if TOBLN.GetValue('BLO_N4') > 0 then
      begin
        // recherche du pere au niveau 3
        TOBPere:=TOBGroupeNomen.FindFirst(['BLO_NUMLIGNE','BLO_N1','BLO_N2','BLO_N3','BLO_N4','BLO_N5'],[Lig,LigneN1,LigneN2,LigneN3,0,0],True) ;
      end else if TOBLN.GetValue('BLO_N3') > 0 then
      begin
        // recherche du pere au niveau 2
        TOBPere:=TOBGroupeNomen.FindFirst(['BLO_NUMLIGNE','BLO_N1','BLO_N2','BLO_N3','BLO_N4','BLO_N5'],[Lig,LigneN1,LigneN2,0,0,0],True) ;
      end else if TOBLN.GetValue('BLO_N2') > 0 then
      begin
        // recherche du pere au niveau 1
        TOBPere:=TOBGroupeNomen.FindFirst(['BLO_NUMLIGNE','BLO_N1','BLO_N2','BLO_N3','BLO_N4','BLO_N5'],[Lig,LigneN1,0,0,0,0],True) ;
      end else TOBPere:=TOBGroupeNomen;

      if TOBPere<>Nil then
      BEGIN
        TOBLoc:=TOB.Create('LIGNEOUV',TOBPere,-1) ;
        TOBLoc.Dupliquer(TOBLN,False,True) ;
        InsertionChampSupOuv (TOBLOC,false);
        //
      	TOBPiece.putValue('GP_UNIQUEBLO',TOBPiece.getValue('GP_UNIQUEBLO')+1);
        TOBLoc.putValue('BLO_UNIQUEBLO',TOBPiece.getValue('GP_UNIQUEBLO'));
        //
        DEV.Code:=TOBLOC.GetValue('BLO_DEVISE') ; GetInfosDevise(DEV) ;
        CalculMontantHtDevLigOuv (TOBLOC,DEV);
        RefArticle:=TOBLoc.GetValue('BLO_ARTICLE') ;
        TypeArticle := TOBLoc.GetValue('BLO_TYPEARTICLE');
        TOBArt:=TOBArticles.FindFirst(['GA_ARTICLE'],[RefArticle],False) ;
        if TOBArt=Nil then
        BEGIN
          TOBArt:=CreerTOBArt(TOBArticles) ;
    			SQl := 'SELECT A.*,AC.*,N.BNP_TYPERESSOURCE FROM ARTICLE A '+
    						 'LEFT JOIN NATUREPREST N ON N.BNP_NATUREPRES=A.GA_NATUREPRES '+
    						 'LEFT JOIN ARTICLECOMPL AC ON AC.GA2_ARTICLE=A.GA_ARTICLE WHERE A.GA_ARTICLE="'+RefArticle+'"';
          Q := OpenSql (Sql,true,-1, '', True);
          TOBArt.SelectDB('',Q) ;
          InitChampsSupArticle (TOBArt);
          ferme (Q);
          LoadTOBDispo(TOBArt,False) ;
        END ;
        if (TypeArticle = 'MAR') or (TypeArticle='PRE') or ((TypeArticle='ARP') and (TOBLoc.Detail.count = 0)) then
        begin
          TOBRef := TOB.Create ('ARTICLE',nil,-1);
          TOBRef.dupliquer (TOBART,false,true);
          TRY
						RecupValoDetail (TOBLOc,TOBRef);
            TOBLoc.putValue('GCA_PRIXBASE',TOBLOC.GetValue('BLO_DPA'));
            TOBLoc.putValue('GA_PAHT',TOBLOC.GetValue('BLO_PRXACHBASE'));
            //
            TOBLoc.putValue('BLO_DPA',TOBLOC.GetValue('GA_PAHT'));
            TOBLoc.putValue('BLO_PRXACHBASE',TOBLOC.GetValue('GCA_PRIXBASE'));
            //
            TOBLoc.putValue('BLO_DPR',TOBREF.GetValue('GA_DPR'));
            TOBLoc.putValue('BNP_TYPERESSOURCE',TOBREF.GetValue('BNP_TYPERESSOURCE'));
          FINALLY
            TOBREF.free;
          END;
        end;
        TOBLoc.Putvalue('ANCPV',TOBLOC.getvalue('BLO_PUHTDEV'));
        TOBLoc.Putvalue('ANCPA',TOBLOC.getvalue('BLO_DPA'));
        TOBLoc.Putvalue('ANCPR',TOBLOC.getvalue('BLO_DPR'));
      END ;
    END ;
//
  FINALLY
  	TOBNomen.free;
  END;

end;

Procedure RenseigneValoOuvBord ( TOBLN : TOB ;CodeDT: String; Qte,QteDuDetail: Double;DEV:RDEVISE;var Valeurs:T_Valeurs) ;
Var TOBPD : TOB ;
    i,IndPou      : integer ;
    TypeArtOk : string;
    QTeLoc,QteDuDet,PrixPourQte,pxht,Pxttc : Double;
    ValLoc,ValPou : T_Valeurs;
BEGIN
InitTableau (ValPou);
TypeArtOk := '';
IndPou := 0;

// fv 02032004
//if QteDudetail = 0 then QteDudetail := 1;

TOBPD:=TOBLN.findfirst(['BLO_TYPEARTICLE'],['POU'],false);
//if TOBPD<> nil  then TypeArtOk := TOBPD.getValue('LIBCOMPL');

for i:=0 to TOBLN.Detail.Count-1 do
    BEGIN
    TOBPD:=TOBLN.Detail[i] ;
    if TOBPD.GetValue('BLO_TYPEARTICLE') = 'POU' then
       begin
       IndPou := i;
       continue;
       end;
    PrixPourQte := TOBPD.GetValue('BLO_PRIXPOURQTE');
    if PrixPourQte <= 0 then PrixPourQte := 1;
    QteLoc := Qte * TOBPD.GetValue('BLO_QTEFACT');
    if TOBPD.detail.count > 0 then
       BEGIN
       InitTableau (ValLoc);
       QteDuDet:= QteDuDetail * TOBPD.GetValue('BLO_QTEDUDETAIL');
       RenseigneValoOuvBord (TOBPD,CodeDt,QTe,QteDuDet,DEV,ValLoc);
       TOBPD.Putvalue ('BLO_DPA',ValLoc[0]);
       TOBPD.Putvalue ('BLO_DPR',ValLoc[1]);
       TOBPD.Putvalue ('BLO_PMAP',ValLoc[6]);
       TOBPD.Putvalue ('BLO_PMRP',ValLoc[7]);
       END;
    Valeurs [0] := Valeurs [0] + ((QteLoc/QteDudetail) * TOBPD.GetValue ('BLO_DPA'));
    Valeurs [1] := Valeurs [1] + ((QteLoc/QteDudetail) * TOBPD.GetValue ('BLO_DPR'));
    pxHt :=(QteLoc/(QteDudetail*PrixPourQte)) * valeur(strs(TOBPD.GetValue ('BLO_PUHTDEV'),V_PGI.OkDecP));
    pxTtc :=(QteLoc/(QteDudetail*PrixPourQte)) * TOBPD.GetValue ('BLO_PUTTCDEV');
    Valeurs [2] := Valeurs [2] + arrondi(PxHt,V_PGI.okdecP);
    Valeurs [3] := Valeurs [3] + Arrondi(PxTTc,V_PGI.okdecP);
    Valeurs [6] := Valeurs [6] + ((QteLoc/QteDudetail) * TOBPD.GetValue ('BLO_PMAP'));
    Valeurs [7] := Valeurs [7] + ((QteLoc/QteDudetail) * TOBPD.GetValue ('BLO_PMRP'));
    if  ArticleOKInPOUR (TOBPD.getValue('BLO_TYPEARTICLE'),TypeartOk) then
        begin
        Valpou [0] := Valpou [0] + ((QteLoc/QteDudetail) * TOBPD.GetValue ('BLO_DPA'));
        Valpou [1] := Valpou [1] + ((QteLoc/QteDudetail) * TOBPD.GetValue ('BLO_DPR'));
        Valpou [2] := Valpou [2] + PxHt;
        Valpou [3] := Valpou [3] + PxTTc;
        Valpou [6] := Valpou [6] + ((QteLoc/QteDudetail) * TOBPD.GetValue ('BLO_PMAP'));
        Valpou [7] := Valpou [7] + ((QteLoc/QteDudetail) * TOBPD.GetValue ('BLO_PMRP'));
        end;
    END ;
    if IndPou <> 0 then
       begin
       TOBLN.detail[IndPou].Putvalue('BLO_DPA',Valpou[0]);TOBLN.detail[IndPou].Putvalue('BLO_DPR',Valpou[1]);
       TOBLN.detail[IndPou].Putvalue('BLO_PMAP',Valpou[6]);TOBLN.detail[IndPou].Putvalue('BLO_PMRP',Valpou[7]);
       TOBLN.detail[IndPou].Putvalue('BLO_PUHTDEV',Valpou[2]);TOBLN.detail[IndPou].Putvalue('BLO_PUTTCDEV',Valpou[3]);
       ValLoc := CalculSurTableau ('*',Valpou,TOBLN.detail[IndPou].getvalue('BLO_QTEFACT')/100);
       FormatageTableau (ValLoc,V_PGI.OkdecP);
       Valeurs := CalculSurTableau ('+',Valeurs,ValLoc);
       end;
    FormatageTableau (valeurs,V_PGI.OkdecP);
END ;

Procedure LigBordereauVersLigOuv ( TOBPIece,TOBL,TOBNOMEN,TOBLN,TOBA : TOB ; Niv,OrdreCompo : integer ;DEV:RDevise;EnHt:Boolean;Prixtraite : string='PUH' ) ;
Var i : integer ;
    TOBNLD,TOBLND,TOBART,TOBDetOuv : TOB ;
    prix,QTe,QteDuPv,QteDuDetail : double;
    pvinit,pvfin,PrixPr : double;
    VenteAchat : string;
    Reajuste : boolean;
BEGIN
VenteAchat := GetInfoParPiece(TOBPiece.GetValue('GP_NATUREPIECEG'), 'GPP_VENTEACHAT');
reajuste :=  isExistsArticle (trim(GetParamsoc('SO_BTECARTPMA')));

for i:=0 to TOBNOMEN.Detail.Count-1 do
    BEGIN
    TOBNLD:=TOBNOMEN.Detail[i] ;
    TOBLND:=TOB.Create('LIGNEOUV',TOBLN,-1) ;
    InsertionChampSupOuv (TOBLND,false);
    TOBLND.Dupliquer (TOBNLD,false,true);
    //
    TOBPiece.putValue('GP_UNIQUEBLO',TOBPiece.getValue('GP_UNIQUEBLO')+1);
    TOBLND.putValue('BLO_UNIQUEBLO',TOBPiece.getValue('GP_UNIQUEBLO'));
    //
    TOBLND.PutValue('BLO_NATUREPIECEG',TOBL.GetValue('GL_NATUREPIECEG')) ;
    TOBLND.PutValue('BLO_REGIMETAXE',TOBL.GetValue('GL_REGIMETAXE')) ;
    TOBLND.PutValue('BLO_SOUCHE',TOBL.GetValue('GL_SOUCHE')) ;
    TOBLND.PutValue('BLO_NUMERO',TOBL.GetValue('GL_NUMERO')) ;
    TOBLND.PutValue('BLO_INDICEG',TOBL.GetValue('GL_INDICEG')) ;
    TOBLND.PutValue('BLO_NUMLIGNE',TOBL.GetValue('GL_NUMLIGNE')) ;
    TOBLND.PutValue('BLO_NIVEAU',Niv) ;
    TOBLND.PutValue('BLO_ORDRECOMPO',OrdreCompo) ;
    //
    //
    CopieOuvFromLigne(TOBLND,TOBL); // recup des infos en provenance de la ligne de doc.
    if EnHT then
       BEGIN
       CalculeLigneHTOuv(TOBLND,TOBPiece,DEV) ;
       END else
       BEGIN
       if TOBLND.Getvalue ('BLO_PUTTCDEV') <> 0 then
          CalculeLigneTTCOuv(TOBLND,TOBPiece,DEV)
       else CalculeLigneHTOuv(TOBLND,TOBPiece,DEV) ;
       END;
		TOBLND.putvalue ('ANCPA',TOBLND.Getvalue ('BLO_DPA'));
		TOBLND.putvalue ('ANCPR',TOBLND.Getvalue ('BLO_DPR'));
    if EnHt then TOBLND.putvalue ('ANCPV',TOBLND.Getvalue ('BLO_PUHTDEV'))
    				else TOBLND.putvalue ('ANCPV',TOBLND.Getvalue ('BLO_PUTTCDEV'));
    Qte := TOBLND.GetValue('BLO_QTEFACT');
    QteDuPv := TOBLND.getValue('BLO_PRIXPOURQTE'); if QteDuPv = 0 then QteDuPV := 1;
    QteDuDetail := TOBLND.GetValue ('BLO_QTEDUDETAIL');
    // fv 02032004
    if QteDudetail = 0 then QteDuDetail := 1;
    //TOBART := TOBA.FIndfirst (['GA_ARTICLE'],[TOBNLD.GetValue('BLO_ARTICLE')],false);
    {$IFDEF BTP}
    //IF (TOBL.GetValue('GL_TYPEARTICLE')='OUV') and
    //   (TOBART.GetValue('GA_TYPEARTICLE')<>'ARP') and
    //   (VenteAchat = 'VEN') then
       //if TOBL.GetValue('GL_NATUREPIECEG') = GetParamSoc('SO_AFNATAFFAIRE') then
       //   DocMetreBiblioDetail (TOBNLD,TOBLND)
       //Else if (Pos(TOBL.GetValue('GL_NATUREPIECEG'),'FBT;FBP')>0) then
       //   DocMetreBiblioDetail (TOBNLD,TOBLND)
       //Else if TOBL.GetValue('GL_NATUREPIECEG') = GetParamSoc('SO_AFNATPROPOSITION') then
       //   DocMetreBiblioDetail (TOBNLD,TOBLND);
    (*
    A rajouter plus tard en niveau 4
    else if (TOBL.GetValue('GL_TYPEARTICLE')='ARP') and
            (TOBArt.GetValue('GA_TYPEARTICLE')<>'ARP') and
            (Qte = 0) and
            (VenteAchat = 'VEN') then DocMetreBiblioDetail (TOBNLD,TOBLND);
    *)
    {$ENDIF}
    LigBordereauVersLigOuv(TOBPiece,TOBL,TOBNLD,TOBLND,TOBA,Niv+1,TOBLND.GetValue('BLO_NUMORDRE'),DEV,EnHt) ;
    if EnHT then
       BEGIN
       CalculeLigneHTOuv(TOBLND,TOBPiece,DEV) ;
       END else
       BEGIN
       if TOBLND.Getvalue ('BLO_PUTTCDEV') <> 0 then
          CalculeLigneTTCOuv(TOBLND,TOBPiece,DEV)
       else CalculeLigneHTOuv(TOBLND,TOBPiece,DEV) ;
       END;
//    CalculMontantHtDevLigOuv (TOBLND,DEV);
    END ;
   // Controle et ajustement pour prix forfait
//
		NumeroteLigneOuv (TOBNOMEN,TOBL,1,1,0,0,0);

		TOBLN.PutValue('MONTANTHTDEV',0);
		TOBLN.PutValue('MONTANTACHAT',0);
		TOBLN.PutValue('MONTANTPR',0);
		TOBLN.PutValue('MONTANTTTCDEV',0);
//
for i:=0 to TOBLN.Detail.Count-1 do
    begin
    TOBDetOuv := TOBLN.detail[I];
		TOBLN.PutValue('MONTANTHTDEV',TOBLN.GetValue('MONTANTHTDEV')+TOBDetOUV.GetValue('MONTANTHTDEV'));
		TOBLN.PutValue('MONTANTACHAT',TOBLN.GetValue('MONTANTACHAT')+TOBDetOUV.GetValue('MONTANTACHAT'));
		TOBLN.PutValue('MONTANTPR',TOBLN.GetValue('MONTANTPR')+TOBDetOUV.GetValue('MONTANTPR'));
		TOBLN.PutValue('MONTANTTTCDEV',TOBLN.GetValue('MONTANTTTCDEV')+TOBDetOUV.GetValue('MONTANTTTCDEV'));
    end;
end;

function TraiteOuvrageFromBordereau ( XX: TForm; TOBPiece,TOBArticles,TOBOuvrage,TOBRepart : TOB ; ARow : integer ; LePremier : boolean ;DEV :RDEVISE;TOBOUVB : TOB=nil) : boolean;
Var RefUnique,TypeArt,TypeNomenc,Depot : String ;
    TOBL,TOBNomen,TOBLN,TOBTMP,TOBDEPART,TOBmetres : TOB ;
    IndiceNomen : integer ;
    conversion,arrondiOk : boolean;
    Valeurs : T_Valeurs;
    CodeDevCli : String;
    EnHt : Boolean;
    Bidon : double;
    Affichage,NiveauDepart : Integer;
    SaisieContreVal : boolean;
    Libelle : string;
    VenteAchat,RecupPrix : string;
    PrxVenteSto,ValPr : double;
BEGIN
  VenteAchat := GetInfoParPiece(TOBPIece.GetValue('GP_NATUREPIECEG'), 'GPP_VENTEACHAT');
  RecupPrix := GetInfoParPiece(TOBPIece.GetValue('GP_NATUREPIECEG'), 'GPP_APPELPRIX');

  NiveauDepart := 1;
  result := false;
  SaisieContreVal := (TOBPiece.Getvalue('GP_SAISIECONTRE')='X');

  conversion := false;
  EnHT:=(TOBPiece.GetValue('GP_FACTUREHT')='X') ;
  Affichage := DOU_AUCUN;
  RefUnique:=GetCodeArtUnique(TOBPiece,ARow) ; if RefUnique='' then Exit ;
  TOBL:=GetTOBLigne(TOBPiece,ARow) ; if TOBL=Nil then Exit ;
  // Correction LS si provenance de la recherche article
  if TOBL.GetValue('GL_INDICENOMEN') > 0 then exit;
  // --
  PrxVenteSto := TOBL.GetValue('GL_PUHTDEV');

  CodeDevCli:= TOBL.GetValue ('GL_DEVISE');
  TypeArt:=TOBL.GetValue('GL_TYPEARTICLE') ;
  if (TypeArt<>'OUV') and (TypeArt <> 'ARP') then
  BEGIN
    result := true;
    Exit ;
  END;
  TypeNomenc:=TOBL.GetValue('GL_TYPENOMENC') ;
  if (TypeNomenc<>'OUV') and (TypeNomenc <> 'OU1') and (TypeArt <> 'ARP') then Exit ;
  // Correction pour les pieces d'achat
  // On ne peut pas acheter des ouvrages ou la prestation associée au prix posée
  //if (VenteAchat = 'ACH') or (pos (TOBPiece.getValue('GP_NATUREPIECEG'),'DBT;ETU;FBT;ABT;FRC') = 0) then
  if Not(ApplicPrixPose (TOBpiece)) then
  BEGIN
    result := true;
    exit;
  END;
  TOBNomen := TOBOUVB.detail[0]; // car l'ouvrage est sur le detail de la ligne
  if TOBNomen = nil then exit;
  TOBLN:=TOB.Create('',TOBOuvrage,-1) ;
  InsertionChampSupOuv (TOBLN,false);
  if ((TOBNomen<>Nil) and (TOBNomen.Detail.Count>0)) then
  BEGIN
    InitTableau (Valeurs);
    RenseigneValoOuvBord(TOBNomen,CodeDevCli,1,1,DEV,Valeurs) ; // recup DPA,DPR,PMAP,PMRP
    if TOBL.GetValue('GL_TYPEARTICLE')='ARP' then
    begin
      TOBDEPART := TOB.Create('LIGNEOUV',TOBLN,-1) ;
      InsertionChampSupOuv (TOBDEPART,false);
      inc (NiveauDepart);
    end else
    begin
      TOBDEPART := TOBLN;
    end;

    LigBordereauVersLigOuv(TOBPiece,TOBL,TOBNomen,TOBDEPART,TOBArticles,NiveauDepart,0,DEV,EnHt) ;

    if TOBL.GetValue('GL_TYPEARTICLE')='ARP' then
    begin
      recupInfoLigneOuv (TOBDEPART,TOBDEPART.Detail[0]);
      TOBDEPART.PutValue('BLO_NIVEAU',1) ;
      TOBDEPART.PutValue('BLO_NUMORDRE',1) ;
      TOBDEPART.PutValue('BLO_QTEDUDETAIL',1) ;
      TOBDEPART.PutValue('BLO_PRIXPOURQTE',1) ;
      GetValoDetail (TOBDEPART);
    end;
    InitTableau (Valeurs);
    CalculeOuvrageDoc (TOBDEPART,1,1,true,DEV,valeurs,EnHt);
    StockeLesTypes(TOBL,Valeurs);
    result := true;
  END else
  BEGIN
    TOBTMP:=TOB.Create('LIGNEOUV',TOBLN,-1) ;
    InsertionChampSupOuv (TOBTMP,false);
    TOBDEPART := TOBLN;
    TOBTMP.PutValue('BLO_NATUREPIECEG',TOBL.GetValue('GL_NATUREPIECEG')) ;
    TOBTMP.PutValue('BLO_SOUCHE',TOBL.GetValue('GL_SOUCHE')) ;
    TOBTMP.PutValue('BLO_DATEPIECE',TOBL.GetValue('GL_DATEPIECE')) ;
    TOBTMP.PutValue('BLO_AFFAIRE',TOBL.GetValue('GL_AFFAIRE')) ;
    TOBTMP.PutValue('BLO_AFFAIRE1',TOBL.GetValue('GL_AFFAIRE1')) ;
    TOBTMP.PutValue('BLO_AFFAIRE2',TOBL.GetValue('GL_AFFAIRE2')) ;
    TOBTMP.PutValue('BLO_AFFAIRE3',TOBL.GetValue('GL_AFFAIRE3')) ;
    TOBTMP.PutValue('BLO_AVENANT',TOBL.GetValue('GL_AVENANT')) ;
    TOBTMP.PutValue('BLO_NUMERO',TOBL.GetValue('GL_NUMERO')) ;
    TOBTMP.PutValue('BLO_INDICEG',TOBL.GetValue('GL_INDICEG')) ;
    TOBTMP.PutValue('BLO_NUMLIGNE',TOBL.GetValue('GL_NUMLIGNE')) ;
    TOBTMP.PutValue('BLO_QTEDUDETAIL',1) ;
    TOBTMP.PutValue('BLO_COMPOSE',RefUnique) ;
    InitTableau (Valeurs);
    result := Entree_LigneOuv(XX,nil,TOBLN,TobArticles,TobL,TOBRepart,TobMetres,False,1,0,DEV,TOBL.Getvalue('GL_QUALIFQTEVTE'),valeurs,tamodif,EnHt,false,false,false) ;
    if not result then TOBLN.free;
  END ;
  if result then
  begin
    if TobNomen <> nil then TOBNomen.Free ;
    if result then
    begin
      //VaChercherQuantitatifDetail (TObDepart);
      InitTableau (Valeurs);
      CalculeOuvrageDoc (TOBDEPART,1,1,true,DEV,valeurs,EnHt);
      StockeLesTypes(TOBL,Valeurs);
    end;
    if TOBL.GetValue('GL_TYPEARTICLE')='ARP' then
    begin
      TOBDEPART.putvalue ('BLO_DPA',valeurs[0]);
      TOBDEPART.putvalue ('BLO_DPR',valeurs[1]);
      TOBDEPART.PutValue('BLO_PMAP',valeurs[6]);
      TOBDEPART.PutValue('BLO_PMRP',valeurs[7]);
      TOBDEPART.putvalue ('BLO_PUHTDEV',valeurs[2]);
      TOBDEPART.putvalue ('BLO_PUTTCDEV',valeurs[3]);
      TOBDEPART.putvalue ('BLO_PUHT',devisetopivotEx(TOBDEPART.Getvalue ('BLO_PUHTDEV'),DEV.Taux,DEV.quotite,V_PGI.okdecP));
      TOBDEPART.putvalue ('BLO_PUTTC',devisetopivotEx(TOBDEPART.Getvalue ('BLO_PUTTCDEV'),DEV.Taux,DEV.quotite,V_PGI.okdecP));
      TOBDEPART.putvalue ('BLO_PUHTBASE',TOBDEPART.GetValue('BLO_PUHT'));
      TOBDEPART.putvalue ('BLO_PUTTCBASE',TOBDEPART.getValue('BLO_PUTTC'));
      TOBDEPART.putvalue ('BLO_TPSUNITAIRE',valeurs[9]);
      TOBDEPART.Putvalue ('ANCPA',TOBDEPART.GetValue('BLO_DPA')) ;
      TOBDEPART.Putvalue ('ANCPR',TOBDEPART.GetValue('BLO_DPR')) ;

      if EnHt then TOBDEPART.Putvalue('ANCPV',TOBDEPART.Getvalue('BLO_PUHTDEV'))
              else TOBDEPART.Putvalue('ANCPV',TOBDEPART.Getvalue('BLO_PUTTCDEV'));
      CalculMontantHtDevLigOuv (TOBDEPART,DEV);
    end else
    begin
      GestionDetailPrixPose (TOBDepart);
    end;
    IndiceNomen:=TOBOuvrage.Detail.Count ;
    //   TOBL.putvalue('GL_LIBELLE', Libelle);
    TOBL.PutValue('GL_INDICENOMEN',IndiceNomen) ;
    if RecupPrix = 'PUH' Then
    begin
      TOBL.Putvalue('GL_PUHTDEV',valeurs[2]);
      TOBL.Putvalue('GL_PUTTCDEV',valeurs[3]);
    end else if RecupPrix = 'DPR' then
    begin
      TOBL.Putvalue('GL_PUHTDEV',valeurs[1]);
      TOBL.Putvalue('GL_PUTTCDEV',valeurs[1]);
    end else if (RecupPrix = 'DPA') or (RecupPrix = 'PAS')  Then
    begin
      TOBL.Putvalue('GL_PUHTDEV',valeurs[0]);
      TOBL.Putvalue('GL_PUTTCDEV',valeurs[0]);
    end;
    TOBL.Putvalue('GL_PUHT',DeviseToPivotEx(TOBL.GetValue('GL_PUHTDEV'),DEV.Taux,DEV.quotite,V_PGI.okdecP));
    TOBL.Putvalue('GL_PUTTC',DevisetoPivotEx(TOBL.GetValue('GL_PUTTCDEV'),DEV.taux,DEV.quotite,V_PGI.okdecP));
    TOBL.Putvalue('GL_PUHTBASE',TOBL.GetValue('GL_PUHT'));
    TOBL.Putvalue('GL_PUTTCBASE',TOBL.GetValue('GL_PUTTC'));
    TOBL.Putvalue('GL_DPA',valeurs[0]);
    TOBL.Putvalue('GL_DPR',valeurs[1]);
    TOBL.Putvalue('GL_PMAP',valeurs[6]);
    TOBL.Putvalue('GL_PMRP',valeurs[7]);
    TOBL.putvalue('GL_TPSUNITAIRE',valeurs[9]);
    if TOBL.getValue('GL_TYPEARTICLE') <> 'ARP' then TOBL.Putvalue('GL_TYPEPRESENT', Affichage)
                                                else TOBL.Putvalue('GL_TYPEPRESENT', 0);
    arrondiOk := (TOBPiece.GetValue('GP_ARRONDILIGNE')='X');
    if EnHT then CalculeLigneHTOuvMere(TOBL,nil,TOBPiece,DEV,V_PGI.okdecV,arrondiok,bidon)
            else CalculeLigneTTCOuvMere(TOBL,nil,TOBPiece,DEV,V_PGI.okdecV,arrondiok,bidon) ;
    //   SommationAchatDoc (TOBPiece,TOBL,False);

    if (PrxVenteSto <> TOBL.GetValue('GL_PUHTDEV')) then
    begin
      ValPr := TOBL.GetValue('GL_DPR');
      TOBL.PutValue('GL_BLOQUETARIF','-');
      TOBL.PutValue('GL_RECALCULER','X');
      if ReajusteMontantOuvrage (TOBARticles,TOBPiece,TOBL,TOBOUvrage.detail[IndiceNomen-1],TOBL.GetValue('GL_PUHTDEV'),ValPr,PrxVenteSto,DEV,EnHt,True) then
      begin
        TOBL.PutValue('GL_PUHTDEV',PrxVenteSto);
      end;
      TOBL.PutValue('GL_BLOQUETARIF','X');
    end;
    TOBDEPART := TOBOUvrage.detail[IndiceNomen-1];
    NumeroteLigneOuv (TOBDEPART,TOBL,1,1,0,0,0);
    positionneCoefMarge (TOBL);
    result := true;
  end;
END ;

{ TUseBordereau }

procedure TUseBordereau.AppliqueChange;
begin
	if not fchanged then exit;
  ReinitList;
  Recharge;
  fchanged := false;
end;

procedure TUseBordereau.ClearAll;
begin
  ReinitList;
  faffaire := '';
  fTiers := '';
  fnatureAuxi := '';
  fChanged :=  false;
end;

constructor TUseBordereau.create(TT : TForm);
var ThePop : Tcomponent;
begin
	FF := TT;
  if FF is TFFacture then ThePop := TFFacture(TT).Findcomponent  ('POPBTP');
	fTOBListBordereaux := TOB.Create ('LISTE DES BORDEREAUX',nil,-1);
  fDatePiece := iDate1900;
  fChanged := false;
  if ThePop = nil then
  BEGIN
    // pas de menu BTP trouve ..on le cree
    POPGS := TPopupMenu.Create(TT);
    POPGS.Name := 'POPBTP';
    fCreatedPop := true;
  END else
  BEGIN
    fCreatedPop := false;
    POPGS := TPopupMenu(thePop);
  END;
  DefiniMenuPop(TT);
end;

procedure TUseBordereau.DefiniMenuPop(Parent: Tform);
var Indice : integer;
begin
  fmaxitems := 0;
  if not fcreatedPop then
  begin
    MesMenuItem[fmaxitems] := TmenuItem.Create (parent);
    with MesMenuItem[fmaxitems] do
      begin
      Caption := '-';
      end;
    inc (fmaxitems);
  end;
  // Acces au bordereau
  MesMenuItem[fmaxitems] := TmenuItem.Create (parent);
  with MesMenuItem[fmaxitems] do
  begin
    Caption := TraduireMemoire ('Accès au bordereau de prix');  // par défaut
    Name := 'BACCESBORD';
    OnClick := GSAccesBordereau;
    enabled := false;
  end;
//  MesMenuItem[fmaxitems].ShortCut := ShortCut( Word('A'), [ssCtrl,ssAlt]);
  inc (fmaxitems);

  for Indice := 0 to fmaxitems -1 do
  begin
  	if MesMenuItem [Indice] <> nil then POPGS.Items.Add (MesMenuItem[Indice]);
  end;

  for indice := 0 to POPGS.Items.Count do
  begin
  	if POPGS.Items [indice].Name = 'BACCESBORD' then BEGIN TheItem := POPGS.Items [indice]; Break; END;
  end;

end;

destructor TUseBordereau.destroy;
var indice : integer;
begin
  inherited;
  for Indice := 0 to fmaxitems -1 do
  begin
    MesMenuItem[Indice].Free;
  end;
  if fcreatedPop then POPGS.free;
  fTOBListBordereaux.free;
end;

function TUseBordereau.IsBordereauExists: boolean;
begin
  result := (fTOBListBordereaux.detail.count > 0);
end;

function TUseBordereau.RechArticlesFromBordereau(reference,Naturepiece : string): boolean;
begin
	result := false;
	if fTobListBordereaux.detail.count = 0 then exit;
  //

  result := GetArticlesFromBordereaux (frequeteSelect,reference,NAturepiece);
end;

procedure TUseBordereau.ReCharge;
var Requete : string;
		QQ : Tquery;
    prems : boolean;
    ReqCli : string;
begin
  if tModeSaisieBordereau in fSaContexte then exit;
  if fNaturePiece = 'FRC' then exit;
	prems := true;
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
  Requete := 'SELECT * FROM BDETETUDE WHERE '+fRequeteSelect;
  QQ := OpenSql (Requete,true,-1, '', True);
  if not QQ.eof then
  begin
  	fTOBListBordereaux.LoadDetailDB ('BDETETUDE','','',QQ,false);
  end;
  DefiniMenuActif;
  ferme (QQ);
end;

procedure TUseBordereau.ReinitList;
begin
	fTOBListBordereaux.clearDetail;
  frequeteSelect := '';
end;

procedure TUseBordereau.SetAffaire(const Value: string);
begin
	if (Value <> fAffaire) then fChanged := true;
  fAffaire := Value;
end;

procedure TUseBordereau.SetDatePiece(const Value: TDateTime);
var fvalue : TDateTime;
begin
	fvalue := value;
	if fValue = iDate1900 then fValue := Date;
	if fValue <> fDatePiece then fChanged := true;
  fDatePiece := fValue;
end;

procedure TUseBordereau.SetNatureAuxi(const Value: string);
var fValue :  string;
begin
	fValue := Value;
	if fValue = '' then fValue := 'CLI';
  if fValue <> fNatureAuxi then FChanged := True;
  fNatureAuxi := fValue;
end;

procedure TUseBordereau.SetTiers(const Value: string);
begin
	if Value <> fTiers then fChanged := true;
  ftiers := Value;
end;

function TUseBordereau.FindReferenceTiers(reference: string): TOB;
var requete : string;
		QQ : Tquery;
    TOBLoc : TOB;
begin
  result := nil;
  if tModeSaisieBordereau in fSaContexte then exit;
	if reference = '' then exit;
  TOBLOc := nil;
  Requete := 'SELECT *,BDETETUDE.* FROM LIGNE LEFT JOIN BDETETUDE ON BDE_NATUREPIECEG=GL_NATUREPIECEG AND BDE_SOUCHE=GL_SOUCHE AND BDE_NUMERO=GL_NUMERO AND BDE_INDICEG=GL_INDICEG WHERE (GL_NATUREPIECEG="BBO") AND (GL_REFARTTIERS="'+reference+'") AND '+fRequeteSelect;
  TRY
    QQ := OpenSql (Requete,true,-1, '', True);
    if not QQ.eof then
    begin
      TOBLOC := TOB.Create ('LES LIGNES',nil,-1);
      TOBLOC.LoadDetailDB ('LIGNE','','',QQ,false);
    end;
  FINALLY
    ferme (QQ);
  End;
  result := TOBLoc;
end;

function TUseBordereau.IsExisteLigneBordereau(TOBPiece: TOB): boolean;
begin
  if tModeSaisieBordereau in fSaContexte then exit;
  result := (TOBPiece.findFirst(['GL_BLOQUETARIF'],['X'],true)<>nil);
end;

procedure TUseBordereau.StContexte(const Value: TModeAff);
begin
  fsacontexte := Value;
end;

procedure TUseBordereau.SetNaturePiece(const Value: string);
begin
	fNaturepiece := value;
end;

procedure TUseBordereau.GSAccesBordereau(Sender: Tobject);
var cledoc : R_CLEDOC;
		TheBordereau : string;
    TheActionSuiv : TActionFiche;
begin
	FillChar(CleDoc,Sizeof(CleDoc),#0) ;
  TheBordereau := fTOBListBordereaux.detail[0].getValue('BDE_PIECEASSOCIEE');
  if TheBordereau = '' then exit;
  DecodeRefPiece (TheBordereau,Cledoc);
  if TFFacture(FF).Action = taconsult then TheActionSuiv := TaConsult else TheActionSuiv := TaModif;
  TFFacture(FF).SauveColList;
  SaisiePiece (cledoc,TheActionSuiv,'','','',false,false,false,false,true);
  TFFacture(FF).RestoreColList;
  ReCharge;
end;

procedure TUseBordereau.DefiniMenuActif;
begin
	if fTOBListBordereaux.detail.count = 1 then
  begin
  	TheItem.visible := true;
  	TheItem.Enabled := true;
  end else
  begin
  	TheItem.Enabled := false;
  	TheItem.visible := false;
  end;
end;

end.
