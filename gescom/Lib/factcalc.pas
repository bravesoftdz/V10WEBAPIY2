unit FactCalc;

interface

Uses HEnt1, HCtrls, UTOB, Ent1, LookUp, Controls, ComCtrls, StdCtrls, ExtCtrls,
{$IFDEF EAGLCLIENT}
     Maineagl,
{$ELSE}
     DB, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} FE_Main,
{$ENDIF}
     SysUtils, FactUtil, SaisUtil, StockUtil, EntGC, FactComm, FactTOB,
     UtilPGI, UTofPiedPort,TarifUtil,UtilArticle,Forms,FactLigneBase, HMsgBox,UentCommun ;

Procedure CalculFacture ( TOBAffaire,TOBPiece,TOBPieceTrait,TOBSSTRAIT,TOBouvrages,TOBouvragesP,TOBBases,TOBBasesL,TOBTiers,TOBArticles,TOBPorcs,TOBPieceRG,TOBBasesRG,TOBVTECOLL : TOB ; DEV : RDEVISE ;SaisieTypeAvanc : boolean=false; Action: TActionFiche=TaModif; ModifAvanc : boolean=false; LigneACalcule : integer=-1; FromGenfac : boolean=false; FromZero : boolean=false) ;
// --
Procedure InverselesPieces ( TOBP : TOB ; NomTable : String ) ;
procedure  CalculTaxesLigne ( TOBL,TOBPiece,TOBTiers,TOBtaxesL : TOB ; DEV: Rdevise; NbDec : integer ; Base : Double ; EnHT : Boolean );

Procedure CalculeLigneHT ( TOBL,TOBTaxesL,TOBPiece: TOB ; DEV : RDevise; NbDec : integer ;Forcer:boolean=False ; TOBTiers : TOB = Nil ) ;
Procedure CalculeLigneTTC ( TOBL,TOBTaxesL,TOBPiece : TOB ; DEV : Rdevise; NbDec : integer; Forcer:boolean=False ; TOBTiers : TOB = Nil ) ;
// Modif BTP
Function CalculeMontantTaxe ( Base,Taux : Double ; FormuleTaxe : String ; TOBL : TOB ) : Double ;
Function CalculeMontantTaxeTTC ( Base,Taux : Double ; FormuleTaxe : String ; TOBL : TOB ) : Double ;
Procedure CumuleLesBases ( TOBBases,TOBTaxes : TOB ; NbDec : integer; Sens:string='+' ) ;
//procedure SommationAchatDoc ( TOBPiece,TOBL : TOB);
procedure GestionTaxesLignes (TOBPiece,TOBTaxes,TOBL,TOBTiers : TOB;var TotTaxesDev,TotTaxes,TOTTaxesNRDev,TOTTaxesNR : double; DEV : Rdevise; Nbdec:integer);
procedure CoordinateMont (TOBG: TOB; ChampDev,ChampPiv: string; DEV: RDevise);
procedure CalculeMontantsAssocie(X: double;var XP:double;DEV:Rdevise);
Procedure InverseX ( TOBL : TOB ; Nam : String ) ;
procedure CalculeLaLigne (TOBAffaire,TOBPiece,TOBPieceTrait,TOBSSTRAIT,TOBOUvrages,TOBOUvragesP,TOBBases,TOBBasesL,TOBTiers,TOBArticles,TOBPOrcs,TOBPieceRG,TOBBasesRG,TOBL,TOBVTECOLL : TOB;DEV : Rdevise; NbDec : integer; saisieAvanc : boolean =false; ModifAvanc : boolean=false; FF : Tform = nil);
{$IFDEF BTP}
procedure CalculTaxesLigneBTP ( TOBL,TOBPiece,TOBTiers,TOBTaxesL : TOB ; DEV : Rdevise; NbDec : integer ; Base : Double ; EnHT : Boolean;PourEnregCpta : boolean=false ) ;
{$ENDIF}
procedure DeduitLigneModifie (TOBL,TOBpiece,TOBPieceTrait,TOBOuvrages,TOBOuvragesP,TOBBases,TOBBasesL,TOBtiers: TOB; DEV : Rdevise; Action : TactionFiche);
procedure DeduitLigneCalc (TOBL,TOBpiece,TOBPieceTrait,TOBOuvrages,TOBOuvragesP,TOBBases,TOBBasesL,TOBtiers: TOB; DEV : Rdevise; Action : TactionFiche);
procedure AjoutLigneCalc (TOBL,TOBpiece,TOBPieceTrait,TOBOuvrages,TOBOuvragesP,TOBBases,TOBBasesL,TOBTiers: TOB; DEV : Rdevise; Action : TactionFiche);
procedure AddBaseLToBases (TOBTAxesL,TOBB : TOB; Nbdec : integer;Sens : string='+');
procedure AddlesPorts (TOBPiece,TOBPorcs,TOBBases,TOBbasesL,TOBTiers : TOB; DEV : Rdevise; Action : TactionFiche ; Sens : string='+');
procedure EquilibrageBases(TOBpiece,TOBBases: TOB; EnHt : Boolean; DEV : Rdevise);
procedure DeduitDeltaBases(TOBpiece,TOBBases: TOB;EnHt : boolean);
procedure GestionTaxesLignesTTCInit (TOBPiece,TOBL,TOBTiers : TOB; Base : double ; var TotTaxesDev,TotTaxes: double; DEV : Rdevise; Nbdec:integer);
Procedure ZeroFacture ( TOBPiece : TOB ) ;
procedure CalculeMontantLigneAchat (TOBL : TOB; var ValeurDpa, ValeurDpr,Valeurpmap,ValeurPmrp : double);
procedure ZeroMontantPorts (TOBPorcs : TOB);
Procedure SommeLignePiedHT ( TOBL,TOBP : TOB; Sens : string='+' ) ;
Procedure SommeLignePiedTTC ( TOBL,TOBP : TOB; Sens : string= '+' ) ;
procedure CalculLignePourcent (TOBPiece,TOBL:TOB;I:Integer);
procedure ReajusteLigneViaBase (TOBP,TOBBasesL: TOB; EnHt : boolean; DEV : RDevise);
Procedure SommeLignePlat ( TOBL,TOBP: TOB; EnHt : boolean; Sens : string='+' ) ;
function GetMontantTaxeLigne (TOBL: TOB; Base : double; DEV : Rdevise) : double;
Function RatioMesure ( Cat,Mesure : String ) : double ;
// --
implementation

uses FactRG,
  {$IFDEF BTP}
    FactOuvrage,
    BTPUtil,
    Factrgpbesoin,
  {$ENDIF}
  Facture,
  FactTarifs,
  FactVariante,
  ParamSoc,
  UCotraitance,
  BTINFOLIVRAISONS_TOF,
  ULiquidTva2014,
  factRetenues,
  FactureBTP
  ,UCumulCollectifs
  ;

(* Var MontantBaseTPF,MontantHTLignes,MontantTTCLignes : double ;*)

const PremiereTaxeSurLesAutres : boolean = True ;

{==============================================================================}
{=============================== AVOIRS =======================================}
{==============================================================================}
Procedure InverseX ( TOBL : TOB ; Nam : String ) ;
Var X : Double ;
    ii : integer ;
BEGIN
ii:=TOBL.GetNumChamp(Nam) ;
TRY
  X:=-TOBL.GetValeur(ii) ;
EXCEPT
  X :=0;
END;
if X<>0 then TOBL.PutValeur(ii,X) ;
END ;

Procedure InverselesPieces ( TOBP : TOB ; NomTable : String ) ;
Var i,k,j : integer ;
    TOBL,TOBD : TOB ;
    Nam : String ;
BEGIN
  if NomTable='PIECE' then
  BEGIN
    for i:=1 to TOBP.NbChamps do
    BEGIN
      Nam:=TOBP.GetNomChamp(i) ;
      if Copy(Nam,1,6)='GP_TOT' then InverseX(TOBP,Nam) ;
    END ;
    for k:=0 to TOBP.Detail.Count-1 do
    BEGIN
      TOBL:=TOBP.Detail[k] ;
      InverseX(TOBL,'GL_QTESTOCK') ;
      InverseX(TOBL,'GL_QTEFACT') ;
      InverseX(TOBL,'GL_QTERESTE') ; { NEWPIECE }
      // --- GUINIER ---
      if CtrlOkReliquat(TOBL, 'GL') then InverseX(TOBL,'GL_MTRESTE') ; { NEWPIECE }
      for i:=1 to TOBL.NbChamps do
      BEGIN
        Nam:=TOBL.GetNomChamp(i) ;
        if Copy(Nam,1,6)='GL_TOT'       then InverseX(TOBL,Nam) ;
        if Copy(Nam,1,10)='GL_MONTANT'  then InverseX(TOBL,Nam) ;
        if Nam='GL_VALEURREMDEV'        then InverseX(TOBL,Nam) ; //NA 20/06/03
      END ;
    END ;
  END else if NomTable='PIEDBASE' then
  BEGIN
    for k:=0 to TOBP.Detail.Count-1 do
    BEGIN
      TOBL:=TOBP.Detail[k] ;
      InverseX(TOBL,'GPB_BASEDEV')  ; InverseX(TOBL,'GPB_VALEURDEV') ;
      InverseX(TOBL,'GPB_BASETAXE') ; InverseX(TOBL,'GPB_VALEURTAXE') ;
    END ;
  END else if NomTable='PIEDPORT' then  //mcd 07/03/02 non pris en compte
  BEGIN
    for k:=0 to TOBP.Detail.Count-1 do
    BEGIN
      TOBL:=TOBP.Detail[k] ;
      InverseX(TOBL,'GPT_BASEHT')  ; InverseX(TOBL,'GPT_BASETTCDEV') ;
      InverseX(TOBL,'GPT_BASEHTDEV') ; InverseX(TOBL,'GPT_BASETTC') ;
      InverseX(TOBL,'GPT_TOTALHT');   InverseX(TOBL,'GPT_TOTALTTC');
      InverseX(TOBL,'GPT_TOTALHTDEV');   InverseX(TOBL,'GPT_TOTALTTCDEV');
      InverseX(TOBL,'GPT_TOTALTAXE1');  InverseX(TOBL,'GPT_TOTALTAXE2');InverseX(TOBL,'GPT_TOTALTAXE3');InverseX(TOBL,'GPT_TOTALTAXE4');InverseX(TOBL,'GPT_TOTALTAXE5');
      InverseX(TOBL,'GPT_TOTALTAXEDEV1');  InverseX(TOBL,'GPT_TOTALTAXEDEV2');InverseX(TOBL,'GPT_TOTALTAXEDEV3');InverseX(TOBL,'GPT_TOTALTAXEDEV4');InverseX(TOBL,'GPT_TOTALTAXEDEV5');
      InverseX(TOBL,'GPT_MONTANTMINI');
    END ;
  END else if NomTable='PIEDECHE' then
  BEGIN
    for k:=0 to TOBP.Detail.Count-1 do
    BEGIN
      TOBL:=TOBP.Detail[k] ;
      InverseX(TOBL,'GPE_MONTANTECHE') ; InverseX(TOBL,'GPE_MONTANTDEV') ;
      InverseX(TOBL,'GPE_MONTANTENCAIS') ;
    END ;
  END else if NomTable='PIEDBASERG' then
  BEGIN
    for k:=0 to TOBP.Detail.Count-1 do
    BEGIN
      TOBL:=TOBP.Detail[k] ;
      InverseX(TOBL,'PBR_BASEDEV')  ; InverseX(TOBL,'PBR_VALEURDEV') ;
      InverseX(TOBL,'PBR_BASETAXE') ; InverseX(TOBL,'PBR_VALEURTAXE') ;
    END ;
  END else if NomTable='PIECERG' then
  BEGIN
    for k:=0 to TOBP.Detail.Count-1 do
    BEGIN
      TOBL:=TOBP.Detail[k] ;
      InverseX(TOBL,'PRG_MTHTRG')  ; InverseX(TOBL,'PRG_MTHTRGDEV') ;
      InverseX(TOBL,'PRG_MTTTCRG')  ; InverseX(TOBL,'PRG_MTTTCRGDEV') ;
      InverseX(TOBL,'PRG_CAUTIONMT')  ; InverseX(TOBL,'PRG_CAUTIONMTU') ;
      InverseX(TOBL,'PRG_CAUTIONMTDEV')  ; InverseX(TOBL,'PRG_CAUTIONMTUDEV') ;
    END ;
  END ELSE if NomTable='LIGNEOUVPLAT' then
  BEGIN
    for k:=0 to TOBP.Detail.Count-1 do
    BEGIN
      TOBL:=TOBP.Detail[k] ;
      for I :=0 to TOBL.Detail.count -1 do
      begin
        TOBD := TOBL.detail[I];
        for J:=1 to TOBD.NbChamps do
        BEGIN
          Nam:=TOBD.GetNomChamp(J) ;
          if Copy(Nam,1,7)='BOP_TOT' then InverseX(TOBD,Nam) ;
          if Copy(Nam,1,11)='BOP_MONTANT' then InverseX(TOBD,Nam) ;
        END ;
      end;
    END ;
  END else if Nomtable= 'PIECETRAIT' then
  BEGIN
    for k:=0 to TOBP.Detail.Count-1 do
    BEGIN
      TOBL:=TOBP.Detail[k] ;
      for J:=1 to TOBL.NbChamps do
      BEGIN
        Nam:=TOBL.GetNomChamp(J) ;
        if Copy(Nam,1,7)='BPE_TOT' then InverseX(TOBL,Nam) ;
        if Copy(Nam,1,11)='BPE_MONTANT' then InverseX(TOBL,Nam) ;
      END ;
      if (TOBL.FieldExists('FACTURE')) then InverseX(TOBL,'FACTURE');
      if TOBL.FieldExists('FACTURAPA') then InverseX(TOBL,'FACTURAPA');
      if TOBL.FieldExists('FACTURPEV') then InverseX(TOBL,'FACTURPEV');
      if TOBL.FieldExists('REGLE') then InverseX(TOBL,'REGLE');
      if TOBL.FieldExists('TOTALFRAIS') then InverseX(TOBL,'TOTALFRAIS');
      if TOBL.FieldExists('MONTANTREGLABLE') then InverseX(TOBL,'MONTANTREGLABLE');
      if TOBL.FieldExists('MONTANTFACT') then InverseX(TOBL,'MONTANTFACT');
    END ;
  END ;
END ;

{==============================================================================}
{========================= INITIALISATIONS ====================================}
{==============================================================================}
Procedure ZeroFacture ( TOBPiece : TOB ) ;
Var i : integer ;
    Nam : String ;
BEGIN
  for i:=1 to TOBPiece.NbChamps do
  BEGIN
    Nam:=TOBPiece.GetNomChamp(i) ;
    if Copy(Nam,1,6)='GP_TOT' then TOBPiece.PutValue(Nam,0) ;
  END ;
  TobPiece.putvalue('AVANCSAISIE',0);
  TobPiece.putvalue('AVANCPREC',0);
END ;

{==============================================================================}
{========================= Calculs sur les DEVISES ============================}
{==============================================================================}
Procedure Egal2Champs ( LaTOB : TOB ; CS,CD : String ) ;
BEGIN
LaTOB.PutValue(CD,LaTOB.GetValue(CS)) ;
END ;

{==============================================================================}
{======================= TAXES, CONVERSIONS, CALCULS ==========================}
{==============================================================================}
Function CalculeMontantTaxe ( Base,Taux : Double ; FormuleTaxe : String ; TOBL : TOB ) : Double ;
BEGIN
  Result:=Base*Taux/100.0 ;
END ;

Function CalculeMontantTaxeTTC ( Base,Taux : Double ; FormuleTaxe : String ; TOBL : TOB ) : Double ;
BEGIN
//  Result:=(Base*(Taux/100.0))/(1.0+(Taux/100.0)) ;
  Result:=Base - (Base/(1.0+(Taux/100.0)));      {TTC - (TTC/(1+Taux))}
END ;

Function CalculeMontantHT ( TTC,Taux : Double ; FormuleTaxe : String ; TOBL : TOB ) : Double ;
BEGIN
  Result:=TTC/(1.0+Taux/100.0) ;
END ;

Procedure CalculeUneCategorieTaxe ( VenteAchat,FamilleTaxe : String ; NumCat,NbDec : integer ; DEV : Rdevise; MontantBase : double ; TOBL,TOBTaxes,TOBPiece : TOB ; EnHT, AvecTPF : boolean ) ;
Var Base,Taux,MontantTaxe,MontantTaxeDev,BaseDev,MontantbaseAch,baseAch,MontantTaxeAch : Double ;
    TOBA,TOBT : TOB ;
    RegimeTaxe,FormuleTaxe,Categorie,TypeInterv : String ;
    TheMillieme,TauxDev : double;
    prefixe : string;
    Fournisseur : string;
    NatureTravail,TauxSt : double;
    DatePiece : TDateTime;
BEGIN
  if TOBTaxes = nil then exit;
  prefixe := GetPrefixeTable (TOBL);
  Fournisseur := '';
  DatePiece := TOBPiece.getvalue('GP_DATEPIECE');
  if TOBL.NomTable = 'LIGNE' then
  begin
    NatureTravail := Valeur(TOBL.getvalue('GLC_NATURETRAVAIL'));
  	if TOBL.getvalue('GLC_NATURETRAVAIL')<>'' then Fournisseur := TOBL.getvalue('GL_FOURNISSEUR');
  	MontantbaseAch := TOBL.getvalue('GL_MONTANTPA');
  end else if TOBL.Nomtable = 'LIGNEOUV' then
  begin
    NatureTravail := Valeur(TOBL.getvalue('BLO_NATURETRAVAIL'));
  	if TOBL.getvalue('BLO_NATURETRAVAIL')<>'' then Fournisseur := TOBL.getvalue('BLO_FOURNISSEUR');
  	MontantbaseAch := TOBL.getvalue('BLO_MONTANTPA');
  end else if TOBL.Nomtable = 'LIGNEOUVPLAT' then
  begin
    NatureTravail := Valeur(TOBL.getvalue('BOP_NATURETRAVAIL'));
  	if TOBL.getvalue('BOP_NATURETRAVAIL')<>'' then Fournisseur := TOBL.getvalue('BOP_FOURNISSEUR');
  	MontantbaseAch := TOBL.getvalue('BOP_MONTANTPA');
  end;
  TypeInterv := '';
  if NatureTravail = 1 then TypeInterv := 'X01'
  else if NatureTravail = 2  then TypeInterv := 'Y00';

  EnHt := (TOBL.GetValue(prefixe+'_FACTUREHT')='X');
{$IFDEF BTP}
TheMillieme := 0;
  TauxDev := TOBL.GetValue(prefixe+'_TAUXDEV');
if TOBL.FieldExists('MILLIEME'+IntToStr(NumCat)) then
begin
	TheMillieme := valeur(TOBL.GetValue('MILLIEME'+IntToStr(NumCat)));
end;
{$ENDIF}
	Categorie := 'TX'+IntToStr(NumCat);
  if FamilleTaxe='' then FamilleTaxe:=TOBL.GetValue(prefixe+'_FAMILLETAXE'+IntToStr(NumCat)) ;
  RegimeTaxe:=TOBL.GetValue(prefixe+'_REGIMETAXE') ; Taux:=0 ;

// ben Voyons...
//TOBT:=VH^.LaTOBTVA.FindFirst(['TV_TVAOUTPF','TV_REGIME','TV_CODETAUX'],['TX'+IntToStr(NumCat),RegimeTaxe,FamilleTaxe],False) ;
TOBT:=VH^.LaTOBTVA.FindFirst(['TV_TVAOUTPF','TV_REGIME','TV_CODETAUX'],['TX1',RegimeTaxe,FamilleTaxe],False) ;
if TOBT<>Nil then
BEGIN
   if (RegimeTaxe = 'INT') and (venteAchat='ACH') then
   begin
     Taux := 0;
//   end else if (FamilleTaxe = VH_GC.AutoLiquiTVAST) and (DatePiece >= StrToDate('01/01/2014')) then
   end else if (TOBPiece.getBoolean('GP_AUTOLIQUID')) and (DatePiece >= StrToDate('01/01/2014')) and (TOBL.GetString('GL_TYPELIGNE')<>'GPT')  then
   begin
     // pour toute les lignes a l'exception des lignes de port ....
     Taux := 0;
   end else
   begin
     if VenteAchat='VEN' then
     begin
      Taux:=TOBT.GetValue('TV_TAUXVTE')
     end else
     begin
     	Taux:=TOBT.GetValue('TV_TAUXACH') ;
     end;
   end;
   FormuleTaxe:=TOBT.GetValue('TV_FORMULETAXE') ;
END ;
  TOBA:=TOBTaxes.FindFirst(['BLB_FOURNISSEUR','BLB_CATEGORIETAXE','BLB_FAMILLETAXE'],[Fournisseur,Categorie,FamilleTaxe],true);
  if TOBA = nil then
  begin
    TOBA:= TOB.Create('LIGNEBASE',TOBTaxes,-1);
    TOBA.putValue('BLB_FOURNISSEUR',Fournisseur);
    TOBA.putValue('BLB_TYPEINTERV',TypeInterv);
    TOBA.putValue('BLB_CATEGORIETAXE',Categorie);
    TOBA.putValue('BLB_FAMILLETAXE',FamilleTaxe);
    TOBA.putValue('BLB_TAUXTAXE',Taux);
    TOBA.putValue('BLB_TAUXDEV',TOBL.GetValue(prefixe+'_TAUXDEV'));
    TOBA.putValue('BLB_COTATION',TOBL.GetValue(prefixe+'_COTATION'));
    TOBA.putValue('BLB_NUMORDRE',TOBL.GetValue(prefixe+'_NUMORDRE'));
    TOBA.putValue('BLB_DEVISE',TOBL.GetValue(prefixe+'_DEVISE'));
  end;
  BaseDev:=TOBA.GetValue('BLB_BASEDEV') ;
  BaseAch:=TOBA.GetValue('BLB_BASEACHAT') ;
  if TheMillieme <> 0 then
  begin
    BaseDev:=BaseDev+Arrondi(MontantBase * (TheMillieme / 1000),NbDec);
  end else
  begin
    BaseDev:= BaseDev+MontantBase ;
  end;
  Base := DeviseToPivotEx(BaseDev,TauxDev,DEV.quotite,DEV.decimale);
  if EnHT then
  begin
    MontantTaxeDev:=Arrondi(CalculeMontantTaxe(BaseDev,Taux,FormuleTaxe,TOBL),NbDec);
    MontantTaxe:=Arrondi(CalculeMontantTaxe(Base,Taux,FormuleTaxe,TOBL),NbDec);
  end else
  begin
    MontantTaxeDev:=Arrondi(CalculeMontantTaxeTTC(BaseDev,Taux,FormuleTaxe,TOBL),NbDec) ;
    MontantTaxe:=Arrondi(CalculeMontantTaxeTTC(Base,Taux,FormuleTaxe,TOBL),NbDec);
  end;
  //
  TOBA.PutValue('BLB_VALEURDEV',MontantTaxeDev);
  TOBA.PutValue('BLB_VALEURTAXE',MontantTaxe);
  if EnHt then
  begin
    TOBA.PutValue('BLB_BASEDEV',BaseDev) ;
    TOBA.PutValue('BLB_BASETAXE',Base) ;
    TOBA.PutValue('BLB_BASETTCDEV',0) ;
    TOBA.PutValue('BLB_BASETAXETTC',0) ;
  end else
  begin
    TOBA.PutValue('BLB_BASEDEV',BaseDev-MontantTaxeDev) ;
    TOBA.PutValue('BLB_BASETAXE',Base-MontantTaxe) ;
    TOBA.PutValue('BLB_BASETTCDEV',BaseDev) ;
    TOBA.PutValue('BLB_BASETAXETTC',Base) ;
  end;
  // -- Partie Sous traitance
  if NatureTravail = 2 then
  begin
    TauxST := GetTauxTaxeST(Fournisseur,GetTvaST(Fournisseur));
    BaseAch := baseAch + MontantbaseAch;
    MontantTaxeAch:=Arrondi(CalculeMontantTaxe(BaseAch,TauxST,'',TOBL),NbDec);
    TOBA.PutValue('BLB_VALEURACHAT',MontantTaxeAch);
    TOBA.PutValue('BLB_BASEACHAT',BaseAch) ;
  end;
END ;

procedure  CalculTaxesLigne ( TOBL,TOBPiece,TOBTiers,TOBtaxesL : TOB ; DEV: Rdevise; NbDec : integer ; Base : Double ; EnHT : Boolean );
Var i : integer ;
    AvecTaxes : boolean ;
    TOBTaxes,TOBTaxe : TOB ;
    AvecTPF : boolean;
    TotTPF,TotTVA,LaBase : Double ;
    FamilleTaxe,RegimeTaxe,NaturePiece,VenteAchat : String ;
BEGIN
  CalculTaxesLigneBtp (TOBL,TOBPiece,TOBTiers,TOBTaxesL, DEV, NbDec,Base , EnHT );
END ;

{==============================================================================}
{======================= PIEDS, BASES, ARRONDIS ===============================}
{==============================================================================}

Procedure CumuleLesBases ( TOBBases,TOBTaxes : TOB ; NbDec : integer; Sens:string='+' ) ;
Var TOBB,TOBA : TOB ;
    i : integer ;
    FamilleTaxe,Categorie,Fournisseur,TypeInterv : String ;
BEGIN
  // Modif BTP
  if TOBBases = nil then exit;
  // --
  for i:=0 to Tobtaxes.detail.count -1 do
  BEGIN
    TOBA:=TOBTaxes.Detail[i] ;
    Fournisseur := TOBA.GetValue('BLB_FOURNISSEUR');
    Categorie := TOBA.GetValue('BLB_CATEGORIETAXE');
    FamilleTaxe:= TOBA.GetValue('BLB_FAMILLETAXE') ;
    TypeInterv := TOBA.GetValue('BLB_TYPEINTERV') ;
    if ((categorie='') or (FamilleTaxe='')) then Continue ;
    TOBB := TOBBases.findFirst(['GPB_FOURN','GPB_CATEGORIETAXE','GPB_FAMILLETAXE'],[Fournisseur,categorie,Familletaxe],true);
    //
    if (TOBB = nil) and (Sens = '-') then exit; // on ne peut pas deduire d'une base de tva inexistante...
    //
    if TOBB = nil then
    BEGIN
      TOBB:=TOB.Create('PIEDBASE',TOBBases,-1) ;
      TOBB.PutValue('GPB_FOURN',Fournisseur) ;
      TOBB.PutValue('GPB_TYPEINTERV',TypeInterv) ;
      TOBB.PutValue('GPB_CATEGORIETAXE',categorie) ;
      TOBB.PutValue('GPB_FAMILLETAXE',FamilleTaxe) ;
      TOBB.PutValue('GPB_TAUXTAXE',TOBA.GetValue('BLB_TAUXTAXE')) ;
      TOBB.PutValue('GPB_TAUXDEV',TOBA.GetValue('BLB_TAUXDEV')) ;
      TOBB.PutValue('GPB_COTATION',TOBA.GetValue('BLB_COTATION')) ;
    END ;
    AddBaseLToBases (TOBA,TOBB,Nbdec,Sens);
  END ;
END ;

{==============================================================================}
{======================= CALCULS SUR LES LIGNES ===============================}
{==============================================================================}
Function RatioMesure ( Cat,Mesure : String ) : double ;
Var TOBM : TOB ;
    X    : Double ;
BEGIN
TOBM:=VH_GC.MTOBMEA.FindFirst(['GME_QUALIFMESURE','GME_MESURE'],[Cat,Mesure],False) ;
X:=0 ; if TOBM<>Nil then X:=TOBM.GetValue('GME_QUOTITE') ; if X=0 then X:=1.0 ;
Result:=X ;
END ;

Procedure SommeLigneMesures ( TOBL,TOBP : TOB; Signe : string='+' ) ;
Var X,Qte : Double ;
    prefixe,Nprefixe : string;
BEGIN
  // Dimensions
  prefixe := GetPrefixeTable (TOBL);
  if Prefixe='GL' then NPrefixe := 'GLC' else NPrefixe := Prefixe;
  if (TOBL.GetValue(prefixe+'_TYPEREF') <> 'CAT') and (TOBL.GetValue(prefixe+'_ARTICLE') = '') then Exit ;
  if TOBL.GetValue(prefixe+'_TYPELIGNE')='GPT' then Exit ;
  if Signe = '+' then
  begin
  X:=TOBP.GetValue('GP_TOTALLINEAIRE')+TOBL.GetValue(prefixe+'_TOTALLINEAIRE') ; TOBP.PutValue('GP_TOTALLINEAIRE',X) ;
  X:=TOBP.GetValue('GP_TOTALSURFACE')+TOBL.GetValue(prefixe+'_TOTALSURFACE') ; TOBP.PutValue('GP_TOTALSURFACE',X) ;
  X:=TOBP.GetValue('GP_TOTALVOLUME')+TOBL.GetValue(prefixe+'_TOTALVOLUME') ; TOBP.PutValue('GP_TOTALVOLUME',X) ;
  X:=TOBP.GetValue('GP_TOTALHEURE')+TOBL.GetValue(prefixe+'_TOTALHEURE') ; TOBP.PutValue('GP_TOTALHEURE',X) ;
  X:=TOBP.GetValue('GP_TOTALPOIDSBRUT')+TOBL.GetValue(prefixe+'_TOTALPOIDSBRUT') ; TOBP.PutValue('GP_TOTALPOIDSBRUT',X) ;
  X:=TOBP.GetValue('GP_TOTALPOIDSNET')+TOBL.GetValue(prefixe+'_TOTALPOIDSNET') ; TOBP.PutValue('GP_TOTALPOIDSNET',X) ;
  X:=TOBP.GetValue('GP_TOTALPOIDSDOUA')+TOBL.GetValue(prefixe+'_TOTALPOIDSDOUA') ; TOBP.PutValue('GP_TOTALPOIDSDOUA',X) ;
  X:=TOBP.GetValue('GP_TOTALQTESTOCK')+TOBL.GetValue(prefixe+'_QTESTOCK') ; TOBP.PutValue('GP_TOTALQTESTOCK',X) ;
  X:=TOBP.GetValue('GP_TOTALQTEFACT')+TOBL.GetValue(prefixe+'_QTEFACT')   ; TOBP.PutValue('GP_TOTALQTEFACT',X) ;
    //
  end else
  begin
  X:=TOBP.GetValue('GP_TOTALLINEAIRE')-TOBL.GetValue(prefixe+'_TOTALLINEAIRE') ; TOBP.PutValue('GP_TOTALLINEAIRE',X) ;
  X:=TOBP.GetValue('GP_TOTALSURFACE')-TOBL.GetValue(prefixe+'_TOTALSURFACE') ; TOBP.PutValue('GP_TOTALSURFACE',X) ;
  X:=TOBP.GetValue('GP_TOTALVOLUME')-TOBL.GetValue(prefixe+'_TOTALVOLUME') ; TOBP.PutValue('GP_TOTALVOLUME',X) ;
  X:=TOBP.GetValue('GP_TOTALHEURE')-TOBL.GetValue(prefixe+'_TOTALHEURE') ; TOBP.PutValue('GP_TOTALHEURE',X) ;
  X:=TOBP.GetValue('GP_TOTALPOIDSBRUT')-TOBL.GetValue(prefixe+'_TOTALPOIDSBRUT') ; TOBP.PutValue('GP_TOTALPOIDSBRUT',X) ;
  X:=TOBP.GetValue('GP_TOTALPOIDSNET')-TOBL.GetValue(prefixe+'_TOTALPOIDSNET') ; TOBP.PutValue('GP_TOTALPOIDSNET',X) ;
  X:=TOBP.GetValue('GP_TOTALPOIDSDOUA')-TOBL.GetValue(prefixe+'_TOTALPOIDSDOUA') ; TOBP.PutValue('GP_TOTALPOIDSDOUA',X) ;
  X:=TOBP.GetValue('GP_TOTALQTESTOCK')-TOBL.GetValue(prefixe+'_QTESTOCK') ; TOBP.PutValue('GP_TOTALQTESTOCK',X) ;
  X:=TOBP.GetValue('GP_TOTALQTEFACT')-TOBL.GetValue(prefixe+'_QTEFACT')   ; TOBP.PutValue('GP_TOTALQTEFACT',X) ;
    //
  end;
END ;

Procedure CalcLigneMesure ( TOBL : TOB ) ;
Var RatioVA,QteS,X : Double ;
    prefixe : string;
BEGIN
  if TOBL=Nil then Exit ;
  prefixe := GetPrefixeTable (TOBL);
  if (TOBL.GetValue(prefixe+'_TYPEREF') <> 'CAT') and (TOBL.GetValue(prefixe+'_ARTICLE') = '') then Exit ;
  if TOBL.GetValue(prefixe+'_TYPELIGNE')='GPT' then Exit ;
  // Modif BTP
  if TOBL.GetValue(prefixe+'_RECALCULER') <> 'X' then exit;
  // --
  QteS:=TOBL.GetValue(prefixe+'_QTESTOCK') ;
  RatioVA:=GetRatio(TOBL,Nil,trsDocument) ;
  X:=RatioMesure('POI',TOBL.GetValue(prefixe+'_QUALIFPOIDS')) ;
  TOBL.PutValue(prefixe+'_TOTALPOIDSNET',X*RatioVA*QteS*TOBL.GetValue(prefixe+'_POIDSNET')) ;
  TOBL.PutValue(prefixe+'_TOTALPOIDSBRUT',X*RatioVA*QteS*TOBL.GetValue(prefixe+'_POIDSBRUT')) ;
  TOBL.PutValue(prefixe+'_TOTALPOIDSDOUA',X*RatioVA*QteS*TOBL.GetValue(prefixe+'_POIDSDOUA')) ;
  X:=RatioMesure('VOL',TOBL.GetValue(prefixe+'_QUALIFVOLUME'))   ; TOBL.PutValue(prefixe+'_TOTALVOLUME',X*RatioVA*QteS*TOBL.GetValue(prefixe+'_VOLUME')) ;
  X:=RatioMesure('SUR',TOBL.GetValue(prefixe+'_QUALIFSURFACE'))  ; TOBL.PutValue(prefixe+'_TOTALSURFACE',X*RatioVA*QteS*TOBL.GetValue(prefixe+'_SURFACE')) ;
  X:=RatioMesure('LIN',TOBL.GetValue(prefixe+'_QUALIFLINEAIRE')) ; TOBL.PutValue(prefixe+'_TOTALLINEAIRE',X*RatioVA*QteS*TOBL.GetValue(prefixe+'_LINEAIRE')) ;
  X:=RatioMesure('TEM',TOBL.GetValue(prefixe+'_QUALIFQTEFACT'))    ; TOBL.PutValue(prefixe+'_TOTALHEURE',X*QteS*TOBL.GetValue(prefixe+'_TPSUNITAIRE')) ;
END ;

procedure ReajusteLigneViaBase (TOBP,TOBBasesL: TOB; EnHt : boolean; DEV : RDevise);
var Indice : integer;
    TOBB,TOBDepart : TOB;
    IndiceLigMtTaxe,IndiceLigMtTaxeDev : integer;
    IndiceTaxe : integer;
    ValeurTaxe,ValeurTaxeDev,Valeur,ValeurDev : double;
    TOTTAxe,TOTTAxeDev : double;
    BaseHt,baseHtDev,DelTa,baseTTC,baseTTCDEV,MTTva,MtTvaDev,Oulala : double;
    TMontantTva ,TMontantTvaDev: array [1..5] of double;
begin
  // protection d'une anomalie grossière
  TOBP.putValue('GL_PUHT',DeviseToPivotEx(TOBP.GetValue('GL_PUHTDEV'),DEV.Taux,DEV.quotite,V_PGI.okdecP));
  TOBP.putValue('GL_PUTTC',DeviseToPivotEx(TOBP.GetValue('GL_PUTTCDEV'),DEV.Taux,DEV.quotite,V_PGI.okdecP));
  //
  if (TOBBasesL.detail.count > 1) then
  begin
    TOBP.putValue('GL_FAMILLETAXE1','*');
  end else
  begin
    if (TOBBasesL.detail.count > 0) then TOBP.putValue('GL_FAMILLETAXE1',TOBBasesL.detail[0].GetString('BLB_FAMILLETAXE'));
  end;
  //
  BaseHt := 0;
  for Indice := 1 to 5 do
  begin
    TMontantTva[Indice]:=0;
    TMontantTvaDev[Indice]:=0;
  end;

  for Indice := 0 to TOBBasesL.detail.count -1 do
  begin
    TOBB := TOBBasesL.detail[Indice];
    IndiceTaxe := StrToInt(copy(TOBB.GetValue('BLB_CATEGORIETAXE'),3,1));
    ValeurTaxe := TOBB.GetValue('BLB_VALEURTAXE');
    ValeurTaxeDev := TOBB.GetValue('BLB_VALEURDEV');
    //
    TMontantTva[IndiceTaxe]:=TMontantTva[IndiceTaxe]+ValeurTaxe;
    TMontantTvaDev[IndiceTaxe]:=TMontantTvaDev[IndiceTaxe]+ValeurTaxeDev;
    //
    baseHt := BaseHt + TOBB.getValue('BLB_BASETAXE');
    baseHtDev := BaseHtDev + TOBB.getValue('BLB_BASEDEV');
    //
    baseTTC := BaseTTC + TOBB.getValue('BLB_BASETAXETTC');
    baseTTCDev := BaseTTCDev + TOBB.getValue('BLB_BASETTCDEV');
    //
    MtTva := MtTva + ValeurTaxe;
    MtTvaDev := MtTvaDev + ValeurTaxeDev;
  end;

  for Indice := 1 to 5 do
  begin
    Delta := Arrondi(TMontantTva[Indice] - TOBP.GetValue('GL_TOTALTAXE'+InttOStr(Indice)),V_PGI.OkdecV) ;
    if Delta <> 0 then
    begin
      TOBP.PutValue('GL_TOTALTAXE'+InttOStr(Indice),Arrondi(TMontantTva[Indice],V_PGI.okdecV));
    end;
    Delta := Arrondi(TMontantTvaDev[Indice] - TOBP.GetValue('GL_TOTALTAXEDEV'+InttOStr(Indice)),DEV.decimale) ;
    if Delta <> 0 then
    begin
      TOBP.PutValue('GL_TOTALTAXEDEV'+InttOStr(Indice),Arrondi(TMontantTvaDev[Indice],DEV.decimale));
    end;
  end;
  //
  if EnHt then
  begin
    Delta := (BaseHt+MTTva) - TOBP.GetValue('GL_TOTALTTC');
    if DelTa <> 0 then
    begin
      TOBP.putValue('GL_TOTALTTC',ARRONDI(BaseHt+MTTva,V_PGI.okdecV ));
    end;
    //
    Delta := (BaseHtDev+MTTvaDev) - TOBP.GetValue('GL_TOTALTTCDEV');
    if DelTa <> 0 then
    begin
      TOBP.putValue('GL_TOTALTTCDEV',ARRONDI(BaseHtDev+MTTvaDev,V_PGI.okdecV ));
    end;
  end else
  begin
    DelTa := Arrondi(BaseHt - TOBP.GetValue('GL_TOTALHT'),V_PGI.okdecV);
    if DelTa <> 0 then
    begin
      TOBP.putValue('GL_TOTALHT',BaseHt);
    end;
    DelTa := Arrondi(BaseHtDev - TOBP.GetValue('GL_TOTALHTDEV'),Dev.decimale);
    if DelTa <> 0 then
    begin
      TOBP.putValue('GL_TOTALHTDEV',BaseHt);
    end;
    // reequilibrage sur devise
    Delta := Arrondi(TOBP.GetValue('GL_TOTALHTDEV')+TOBP.GetValue('GL_TOTREMPIEDDEV')+TOBP.GetValue('GL_TOTESCLIGNEDEV')-TOBP.GetValue('GL_MONTANTHTDEV'),DEV.decimale);
    if Delta <> 0 then
    begin
        TOBP.PutValue('GL_MONTANTHTDEV',TOBP.GetValue('GL_MONTANTHTDEV')+Delta)
    end;
    Delta := Arrondi(TOBP.GetValue('GL_TOTALHT')+TOBP.GetValue('GL_TOTREMPIED')+TOBP.GetValue('GL_TOTESCLIGNE') -TOBP.GetValue('GL_MONTANTHT'),V_PGI.OkdecV);
    if Delta <> 0 then
    begin
      //
      TOBP.PutValue('GL_MONTANTHT',TOBP.GetValue('GL_MONTANTHT')+Delta)
    end;
    //
  end;

  //
end;

Procedure SommeLignePlat ( TOBL,TOBP: TOB; EnHt : boolean; Sens : string='+' ) ;
Var X,TotTaxes,TotTaxesDev : Double ;
    i : integer ;
    prefixe,PrefixeD : string;
  	QteF : double;
  	PQ : double;
    PUHTDEV: double;
    PUHT: double;
BEGIN
  Prefixe := GetPrefixeTable(TOBL);
  PrefixeD := GetPrefixeTable(TOBP);
  if (TOBL.GetValue(prefixe+'_TYPEREF') <> 'CAT') and (TOBL.GetValue(prefixe+'_ARTICLE') = '') then Exit ;
  {$IFDEF BTP}
  if (not IsLigneFromCentralis(TOBL))
      and (TOBL.GetValue(prefixe+'_TYPEARTICLE') <> '$#$')
  {$ENDIF} then
  begin
    if Sens = '+' then X:=TOBP.GetValue(PrefixeD+'_MONTANTHTDEV')+TOBL.GetValue(Prefixe+'_MONTANTHTDEV')
                  else X:=TOBP.GetValue(PrefixeD+'_MONTANTHTDEV')-TOBL.GetValue(Prefixe+'_MONTANTHTDEV');
    TOBP.PutValue(PrefixeD+'_MONTANTHTDEV',X) ;
    if Sens = '+' then X:=TOBP.GetValue(PrefixeD+'_MONTANTHT')+TOBL.GetValue(Prefixe+'_MONTANTHT')
                  else X:=TOBP.GetValue(PrefixeD+'_MONTANTHT')-TOBL.GetValue(Prefixe+'_MONTANTHT');
    TOBP.PutValue(PrefixeD+'_MONTANTHT',X) ;
    if Sens = '+' then X:=TOBP.GetValue(PrefixeD+'_TOTALHTDEV')+TOBL.GetValue(prefixe+'_TOTALHTDEV')
                  else X:=TOBP.GetValue(PrefixeD+'_TOTALHTDEV')-TOBL.GetValue(prefixe+'_TOTALHTDEV');
    TOBP.PutValue(PrefixeD+'_TOTALHTDEV',X) ;
    if Sens = '+' then X:=TOBP.GetValue(PrefixeD+'_TOTALHT')+TOBL.GetValue(Prefixe+'_TOTALHT')
                  else X:=TOBP.GetValue(PrefixeD+'_TOTALHT')-TOBL.GetValue(Prefixe+'_TOTALHT');
    TOBP.PutValue(PrefixeD+'_TOTALHT',X) ;
  end;
  if Sens = '+' then X:=TOBP.GetValue(PrefixeD+'_TOTREMPIEDDEV')+TOBL.GetValue(Prefixe+'_TOTREMPIEDDEV')
                else X:=TOBP.GetValue(PrefixeD+'_TOTREMPIEDDEV')-TOBL.GetValue(Prefixe+'_TOTREMPIEDDEV');
  TOBP.PutValue(PrefixeD+'_TOTREMPIEDDEV',X) ;
  if Sens = '+' then X:=TOBP.GetValue(PrefixeD+'_TOTREMPIED')+TOBL.GetValue(Prefixe+'_TOTREMPIED')
                else X:=TOBP.GetValue(PrefixeD+'_TOTREMPIED')-TOBL.GetValue(Prefixe+'_TOTREMPIED');
  TOBP.PutValue(PrefixeD+'_TOTREMPIED',X) ;
  if Sens = '+' then X:=TOBP.GetValue(PrefixeD+'_TOTESCLIGNEDEV')+TOBL.GetValue(Prefixe+'_TOTESCLIGNEDEV')
                else X:=TOBP.GetValue(PrefixeD+'_TOTESCLIGNEDEV')+TOBL.GetValue(Prefixe+'_TOTESCLIGNEDEV');
  TOBP.PutValue(PrefixeD+'_TOTESCLIGNEDEV',X);
  if Sens = '+' then X:=TOBP.GetValue(PrefixeD+'_TOTESCLIGNE')+TOBL.GetValue(Prefixe+'_TOTESCLIGNE')
                else X:=TOBP.GetValue(PrefixeD+'_TOTESCLIGNE')+TOBL.GetValue(Prefixe+'_TOTESCLIGNE');
  TOBP.PutValue(PrefixeD+'_TOTESCLIGNE',X);
  TotTaxes:=0 ;
  for i:=1 to 5 do
  begin
    if Sens = '+' then X := TOBP.GetValue(PrefixeD+'_TOTALTAXEDEV'+IntToStr(i))+TOBL.GetValue(Prefixe+'_TOTALTAXEDEV'+IntToStr(i))
                  else X := TOBP.GetValue(PrefixeD+'_TOTALTAXEDEV'+IntToStr(i))-TOBL.GetValue(Prefixe+'_TOTALTAXEDEV'+IntToStr(i));
    TOBP.putValue(PrefixeD+'_TOTALTAXEDEV'+IntToStr(i),X);
    if Sens = '+' then X := TOBP.GetValue(PrefixeD+'_TOTALTAXE'+IntToStr(i))+TOBL.GetValue(Prefixe+'_TOTALTAXE'+IntToStr(i))
                  else X := TOBP.GetValue(PrefixeD+'_TOTALTAXE'+IntToStr(i))-TOBL.GetValue(Prefixe+'_TOTALTAXE'+IntToStr(i));
    TOBP.putValue(PrefixeD+'_TOTALTAXE'+IntToStr(i),X);
  end;

  if Sens = '+' then X:=TOBP.GetValue(prefixeD+'_MONTANTTTCDEV')+TOBL.GetValue(prefixe+'_MONTANTTTCDEV')
                else X:=TOBP.GetValue(prefixeD+'_MONTANTTTCDEV')-TOBL.GetValue(prefixe+'_MONTANTTTCDEV');
  TOBP.PutValue(prefixeD+'_MONTANTTTCDEV',X) ;
  if Sens = '+' then X:=TOBP.GetValue(prefixeD+'_MONTANTTTC')+TOBL.GetValue(prefixe+'_MONTANTTTC')
                else X:=TOBP.GetValue(prefixeD+'_MONTANTTTC')-TOBL.GetValue(prefixe+'_MONTANTTTC');
  TOBP.PutValue(prefixeD+'_MONTANTTTC',X) ;
  if Sens = '+' then X := TOBP.GetValue(PrefixeD+'_TOTALTTCDEV')+TOBL.GetValue(Prefixe+'_TOTALTTCDEV')
                else X := TOBP.GetValue(PrefixeD+'_TOTALTTCDEV')-TOBL.GetValue(Prefixe+'_TOTALTTCDEV');
  TOBP.PutValue(PrefixeD+'_TOTALTTCDEV',X);
  if Sens = '+' then X := TOBP.GetValue(PrefixeD+'_TOTALTTC')+TOBL.GetValue(Prefixe+'_TOTALTTC')
                else X := TOBP.GetValue(PrefixeD+'_TOTALTTC')-TOBL.GetValue(Prefixe+'_TOTALTTC');
  TOBP.PutValue(PrefixeD+'_TOTALTTC',X);
  if Sens = '+' then X:=TOBP.GetValue('TOTREMPIEDTTCDEV')+TOBL.GetValue('TOTREMPIEDTTCDEV')
                else X:=TOBP.GetValue('TOTREMPIEDTTCDEV')-TOBL.GetValue('TOTREMPIEDTTCDEV');
  TOBP.PutValue('TOTREMPIEDTTCDEV',X) ;
  if Sens = '+' then X:=TOBP.GetValue('TOTREMPIEDTTC')+TOBL.GetValue('TOTREMPIEDTTC')
                else X:=TOBP.GetValue('TOTREMPIEDTTC')-TOBL.GetValue('TOTREMPIEDTTC');
  TOBP.PutValue('TOTREMPIEDTTC',X) ;
  if Sens ='+' then X:=TOBP.GetValue('TOTESCLIGNETTCDEV')+TOBL.GetValue('TOTESCLIGNETTCDEV')
               else X:=TOBP.GetValue('TOTESCLIGNETTCDEV')-TOBL.GetValue('TOTESCLIGNETTCDEV');
  TOBP.PutValue('TOTESCLIGNETTCDEV',X) ;
  if Sens ='+' then X:=TOBP.GetValue('TOTESCLIGNETTC')+TOBL.GetValue('TOTESCLIGNETTC')
               else X:=TOBP.GetValue('TOTESCLIGNETTC')-TOBL.GetValue('TOTESCLIGNETTC');
  TOBP.PutValue('TOTESCLIGNETTC',X) ;
  if not EnHt then
  begin
    QteF:=TOBP.GetValue(prefixeD+'_QTEFACT') ; if QteF=0 then Exit ;
    PQ:=TOBP.GetValue(prefixeD+'_PRIXPOURQTE') ; if PQ<=0 then BEGIN PQ:=1.0 ; TOBL.PutValue(prefixe+'_PRIXPOURQTE',PQ) ; END ;
    // gestion des Pu Ht
    PUHTDEV:=arrondi(TOBP.GetValue(prefixeD+'_MONTANTHTDEV')*PQ/QteF,V_PGI.okdecP) ;
    PUHT:=arrondi(TOBP.GetValue(prefixeD+'_MONTANTHT')*PQ/QteF,V_PGI.OkdecP) ;
    TOBP.putValue(PrefixeD+'_PUHTDEV',PUHTDEV);
    TOBP.putValue(PrefixeD+'_PUHT',PUHT);
  end;
//  CumuleSuivantType (TOBP,TOBL);
END ;

Procedure SommeLignePiedHT ( TOBL,TOBP : TOB; Sens : string='+' ) ;
Var X,TotTaxes,TotTaxesDev : Double ;
    i : integer ;
    prefixe :string;
BEGIN
  prefixe := GetPrefixeTable(TOBL);
  // ----------
  // Pour POC on ne totalise pas les lignes de ST dans le cadre de la facturation
  // ----------
  if (Pos(TOBL.GetString(prefixe+'_NATUREPIECEG'),'DBT;FBP;FBT;ABP;ABT')>0) and (VH_GC.BTCODESPECIF = '001') and (IsLigneSoustraite (TOBL)) then Exit;
  // ----------
  if (TOBL.GetValue(prefixe+'_TYPEREF') <> 'CAT') and (TOBL.GetValue(prefixe+'_ARTICLE') = '') then Exit ;
  // Correction bug sur ancienne version
  if Sens = '+' then
  begin
    if (TOBL.GetValue(prefixe+'_TOTREMPIEDDEV') <>0)  and (TOBL.GetValue(prefixe+'_REMISEPIED')=0) then
    begin
      TOBL.PutValue(prefixe+'_TOTREMPIEDDEV',0);
      TOBL.PutValue(prefixe+'_TOTREMPIED',0);
    end;
    if (TOBL.GetValue(prefixe+'_TOTESCLIGNEDEV') <>0)  and (TOBL.GetValue(prefixe+'_ESCOMPTE')=0) then
    begin
      TOBL.PutValue(prefixe+'_TOTALESCLIGNEDEV',0);
      TOBL.PutValue(prefixe+'_TOTALESCLIGNE',0);
    end;
  end;
  //

  if (TOBL.GetValue(prefixe+'_REMISABLEPIED')='X') then
  BEGIN
    if Sens = '+' then X:=TOBP.GetValue('GP_TOTALBASEREMDEV')+TOBL.GetValue(prefixe+'_MONTANTHTDEV')
                  else X:=TOBP.GetValue('GP_TOTALBASEREMDEV')-TOBL.GetValue(prefixe+'_MONTANTHTDEV');
    TOBP.PutValue('GP_TOTALBASEREMDEV',X) ;
    if Sens = '+' then X:=TOBP.GetValue('GP_TOTALBASEREM')+TOBL.GetValue(prefixe+'_MONTANTHT')
                  else X:=TOBP.GetValue('GP_TOTALBASEREM')-TOBL.GetValue(prefixe+'_MONTANTHT');
    TOBP.PutValue('GP_TOTALBASEREM',X) ;
  END ;
  if (TOBL.GetValue(prefixe+'_ESCOMPTABLE')='X') then
  BEGIN
    if Sens = '+' then X:=TOBP.GetValue('GP_TOTALBASEESCDEV')+TOBL.GetValue(prefixe+'_MONTANTHTDEV')-TOBL.GetValue(prefixe+'_TOTREMPIEDDEV')
                  else X:=TOBP.GetValue('GP_TOTALBASEESCDEV')-TOBL.GetValue(prefixe+'_MONTANTHTDEV')+TOBL.GetValue(prefixe+'_TOTREMPIEDDEV');
    TOBP.PutValue('GP_TOTALBASEESCDEV',X) ;
    if Sens = '+' then X:=TOBP.GetValue('GP_TOTALBASEESC')+TOBL.GetValue(prefixe+'_MONTANTHT')-TOBL.GetValue(prefixe+'_TOTREMPIED')
                  else X:=TOBP.GetValue('GP_TOTALBASEESC')-TOBL.GetValue(prefixe+'_MONTANTHT')+TOBL.GetValue(prefixe+'_TOTREMPIED');
    TOBP.PutValue('GP_TOTALBASEESC',X) ;
  END ;

  // VARIANTES
  if {(not isVariante(TOBL))}
  {$IFDEF BTP}
      {and} (not IsLigneFromCentralis(TOBL))
      and (TOBL.GetValue(prefixe+'_TYPEARTICLE') <> '$#$')
  {$ENDIF} then
  begin
    if Sens = '+' then X:=TOBP.GetValue('GP_TOTALHTDEV')+TOBL.GetValue(prefixe+'_TOTALHTDEV')
                  else X:=TOBP.GetValue('GP_TOTALHTDEV')-TOBL.GetValue(prefixe+'_TOTALHTDEV');
    TOBP.PutValue('GP_TOTALHTDEV',X) ;
    if Sens = '+' then X:=TOBP.GetValue('GP_TOTALHT')+TOBL.GetValue(prefixe+'_TOTALHT')
                  else X:=TOBP.GetValue('GP_TOTALHT')-TOBL.GetValue(prefixe+'_TOTALHT');
    TOBP.PutValue('GP_TOTALHT',X) ;
  end;
  if Sens = '+' then X:=TOBP.GetValue('GP_TOTALREMISEDEV')+TOBL.GetValue(prefixe+'_TOTREMPIEDDEV')
                else X:=TOBP.GetValue('GP_TOTALREMISEDEV')-TOBL.GetValue(prefixe+'_TOTREMPIEDDEV');
  TOBP.PutValue('GP_TOTALREMISEDEV',X) ;
  if Sens = '+' then X:=TOBP.GetValue('GP_TOTALREMISE')+TOBL.GetValue(prefixe+'_TOTREMPIED')
                else X:=TOBP.GetValue('GP_TOTALREMISE')-TOBL.GetValue(prefixe+'_TOTREMPIED');
  TOBP.PutValue('GP_TOTALREMISE',X) ;
  if Sens = '+' then X:=TOBP.GetValue('GP_TOTALESCDEV')+TOBL.GetValue(prefixe+'_TOTESCLIGNEDEV')
                else X:=TOBP.GetValue('GP_TOTALESCDEV')-TOBL.GetValue(prefixe+'_TOTESCLIGNEDEV');
  TOBP.PutValue('GP_TOTALESCDEV',X);
  if Sens = '+' then X:=TOBP.GetValue('GP_TOTALESC')+TOBL.GetValue(prefixe+'_TOTESCLIGNE')
                else X:=TOBP.GetValue('GP_TOTALESC')-TOBL.GetValue(prefixe+'_TOTESCLIGNE');
  TOBP.PutValue('GP_TOTALESC',X);
  TotTaxes:=0 ;
  if Sens = '+' then X := TOBP.GetValue('GP_TOTALTTCDEV')+TOBL.GetValue(prefixe+'_TOTALTTCDEV')
                else X := TOBP.GetValue('GP_TOTALTTCDEV')-TOBL.GetValue(prefixe+'_TOTALTTCDEV');
  TOBP.PutValue('GP_TOTALTTCDEV',X);

  if Sens = '+' then X := TOBP.GetValue('GP_TOTALTTC')+TOBL.GetValue(prefixe+'_TOTALTTC')
                else X := TOBP.GetValue('GP_TOTALTTC')-TOBL.GetValue(prefixe+'_TOTALTTC');
  TOBP.PutValue('GP_TOTALTTC',X);
END ;

Procedure SommeLignePiedTTC ( TOBL,TOBP: TOB; Sens : string= '+' ) ;
Var X,TotTaxes,TotTaxesDev : Double ;
    i : integer ;
    prefixe : string;
BEGIN
  prefixe := GetPrefixeTable (TOBL);
  // ----------
  // Pour POC on ne totalise pas les lignes de ST dans le cadre de la facturation
  // ----------
  if (Pos(TOBL.GetString(prefixe+'_NATUREPIECEG'),'DBT;FBP;FBT;ABP;ABT')>0) and (VH_GC.BTCODESPECIF = '001') and (IsLigneSoustraite (TOBL)) then Exit;
  // ----------
  if (TOBL.GetValue(prefixe+'_TYPEREF') <> 'CAT') and (TOBL.GetValue(prefixe+'_ARTICLE') = '') then Exit ;
  //
  if ((TOBL.GetValue(prefixe+'_REMISEPIED')<>0) and (TOBL.GetValue(prefixe+'_REMISABLEPIED')='X')) then
  BEGIN
    if TOBP.getValue('GP_TOTREMISETTCDEV')<>0 then
    begin
      if Sens = '+' then X:=TOBP.GetValue('GP_TOTALBASEREMDEV')+TOBL.GetValue(prefixe+'_MONTANTTTCDEV')-TOBP.GetValue('GP_TOTALBASEESCDEV')
                    else X:=TOBP.GetValue('GP_TOTALBASEREMDEV')-TOBL.GetValue(prefixe+'_MONTANTTTCDEV')+-TOBP.GetValue('GP_TOTALBASEESCDEV');
      TOBP.PutValue('GP_TOTALBASEREMDEV',X) ;
    end;
    if TOBP.getValue('GP_TOTREMISETTC')<>0 then
    begin
      if Sens = '+' then X:=TOBP.GetValue('GP_TOTALBASEREM')+TOBL.GetValue(prefixe+'_MONTANTTTC')-TOBP.GetValue('GP_TOTALBASEESC')
                    else X:=TOBP.GetValue('GP_TOTALBASEREM')-TOBL.GetValue(prefixe+'_MONTANTTTC')+TOBP.GetValue('GP_TOTALBASEESC');
      TOBP.PutValue('GP_TOTALBASEREM',X) ;
    end;
  END ;
  if ((TOBL.GetValue(prefixe+'_ESCOMPTE')<>0) and (TOBL.GetValue(prefixe+'_ESCOMPTABLE')='X')) then
  BEGIN
    If TOBP.GetValue('GP_TOTESCTTC')<>0  then
    begin
      if Sens = '+' then X:=TOBP.GetValue('GP_TOTALBASEESCDEV')+TOBL.GetValue(prefixe+'_MONTANTTTCDEV')
                    else X:=TOBP.GetValue('GP_TOTALBASEESCDEV')-TOBL.GetValue(prefixe+'_MONTANTTTCDEV');
      TOBP.PutValue('GP_TOTALBASEESCDEV',X) ;
    end;
    If TOBP.GetValue('GP_TOTALESCDEV')<>0  then
    begin
      if Sens = '+' then X:=TOBP.GetValue('GP_TOTALBASEESC')+TOBL.GetValue(prefixe+'_MONTANTTTC')
                    else X:=TOBP.GetValue('GP_TOTALBASEESC')-TOBL.GetValue(prefixe+'_MONTANTTTC');
      TOBP.PutValue('GP_TOTALBASEESC',X) ;
    end;
  END ;
  //
  if Sens = '+' then X:=TOBP.GetValue('GP_TOTALHTDEV')+TOBL.GetValue(prefixe+'_TOTALHTDEV')
                else X:=TOBP.GetValue('GP_TOTALHTDEV')-TOBL.GetValue(prefixe+'_TOTALHTDEV');
  TOBP.PutValue('GP_TOTALHTDEV',X) ;
  if Sens = '+' then X:=TOBP.GetValue('GP_TOTALHT')+TOBL.GetValue(prefixe+'_TOTALHT')
                else X:=TOBP.GetValue('GP_TOTALHT')-TOBL.GetValue(prefixe+'_TOTALHT');
  TOBP.PutValue('GP_TOTALHT',X) ;
  //
  if Sens = '+' then X:=TOBP.GetValue('GP_TOTALREMISEDEV')+TOBL.GetValue(prefixe+'_TOTREMPIEDDEV')
                else X:=TOBP.GetValue('GP_TOTALREMISEDEV')-TOBL.GetValue(prefixe+'_TOTREMPIEDDEV');
  TOBP.PutValue('GP_TOTALREMISEDEV',X) ;
  if Sens = '+' then X:=TOBP.GetValue('GP_TOTALREMISE')+TOBL.GetValue(prefixe+'_TOTREMPIED')
                else X:=TOBP.GetValue('GP_TOTALREMISE')-TOBL.GetValue(prefixe+'_TOTREMPIED');
  TOBP.PutValue('GP_TOTALREMISE',X) ;
  //
  if Sens = '+' then X:=TOBP.GetValue('GP_TOTALESCDEV')+TOBL.GetValue(prefixe+'_TOTESCLIGNEDEV')
                else X:=TOBP.GetValue('GP_TOTALESCDEV')-TOBL.GetValue(prefixe+'_TOTESCLIGNEDEV');
  TOBP.PutValue('GP_TOTALESCDEV',X);
  if Sens = '+' then X:=TOBP.GetValue('GP_TOTALESC')+TOBL.GetValue(prefixe+'_TOTESCLIGNE')
                else X:=TOBP.GetValue('GP_TOTALESC')-TOBL.GetValue(prefixe+'_TOTESCLIGNE');
  TOBP.PutValue('GP_TOTALESC',X);
  //
  if Sens = '+' then X:=TOBP.GetValue('GP_TOTREMISETTCDEV')+TOBL.GetValue('TOTREMPIEDTTCDEV')
                else X:=TOBP.GetValue('GP_TOTREMISETTCDEV')-TOBL.GetValue('TOTREMPIEDTTCDEV');
  TOBP.PutValue('GP_TOTREMISETTCDEV',X) ;
  if Sens = '+' then X:=TOBP.GetValue('GP_TOTREMISETTC')+TOBL.GetValue('TOTREMPIEDTTC')
                else X:=TOBP.GetValue('GP_TOTREMISETTC')-TOBL.GetValue('TOTREMPIEDTTC');
  TOBP.PutValue('GP_TOTREMISETTC',X) ;

  if Sens = '+' then X:=TOBP.GetValue('GP_TOTALTTCDEV')+TOBL.GetValue(prefixe+'_TOTALTTCDEV')
                else X:=TOBP.GetValue('GP_TOTALTTCDEV')-TOBL.GetValue(prefixe+'_TOTALTTCDEV');
  TOBP.PutValue('GP_TOTALTTCDEV',X) ;
  if Sens = '+' then X:=TOBP.GetValue('GP_TOTALTTC')+TOBL.GetValue(prefixe+'_TOTALTTC')
                else X:=TOBP.GetValue('GP_TOTALTTC')-TOBL.GetValue(prefixe+'_TOTALTTC');
  TOBP.PutValue('GP_TOTALTTC',X) ;
  if Sens ='+' then X:=TOBP.GetValue('GP_TOTESCTTCDEV')+TOBL.GetValue('TOTESCLIGNETTCDEV')
               else X:=TOBP.GetValue('GP_TOTESCTTCDEV')-TOBL.GetValue('TOTESCLIGNETTCDEV');
  TOBP.PutValue('GP_TOTESCTTCDEV',X) ;
  if Sens ='+' then X:=TOBP.GetValue('GP_TOTESCTTC')+TOBL.GetValue('TOTESCLIGNETTC')
               else X:=TOBP.GetValue('GP_TOTESCTTC')-TOBL.GetValue('TOTESCLIGNETTC');
  TOBP.PutValue('GP_TOTESCTTC',X) ;
END ;

//  Modif BTP
procedure GestionTaxesLignes (TOBPiece,TOBTaxes,TOBL,TOBTiers : TOB;var TotTaxesDev,TotTaxes,TOTTaxesNRDev,TOTTaxesNR : double; DEV : Rdevise; Nbdec:integer);
var enHt : boolean;
    i_montant : integer;
    RemPied,Escompte : double;
    i : integer;
    LaTaxeDev,laTaxe,LaTaxeNR,LaTaxeNRDEV : double;
    TOBTaxeL : TOB;
    Prefixe : string;
begin
  Prefixe := GetPrefixeTable (TOBL);
  EnHT:=(TOBL.GetValue(prefixe+'_FACTUREHT')='X'); RemPied:=0 ; Escompte:=0 ;
  if EnHT then i_montant:=TOBL.GetNumChamp (prefixe+'_TOTALHTDEV')
          else i_montant:=TOBL.GetNumChamp (prefixe+'_TOTALTTCDEV') ;
  if TOBL.GetValue(prefixe+'_REMISABLEPIED')='X' then RemPied:=TOBL.GetValue(prefixe+'_REMISEPIED')/100.0 ;
  if TOBL.GetValue(prefixe+'_ESCOMPTABLE')='X' then Escompte:=TOBL.GetValue(prefixe+'_ESCOMPTE')/100.0 ;

  // Taxes
  CalculTaxesLigne(TOBL,TOBPiece,TOBTiers,TOBTaxes,DEV,NbDec,TOBL.GetValeur(i_montant),EnHt) ;
  // Autres cumuls lignes
  TotTaxes:=0 ; TotTaxesNR:=0 ;
  TotTaxesDEV:=0 ; TotTaxesNRDEV:=0 ;
  if TOBTaxes<>Nil then
  BEGIN
    for i:=0 to TOBtaxes.detail.count -1 do
    BEGIN
      LaTaxeDev:=TOBTaxes.Detail[i].GetValue('BLB_VALEURDEV') ;
      LaTaxe:=TOBTaxes.Detail[i].GetValue('BLB_VALEURTAXE') ;
      if ((RemPied=1) or (Escompte=1)) then
      begin
        LaTaxeNR:=0;
        LaTaxeNRDEV :=0;
      end else
      begin
        LaTaxeNRDEV:=Arrondi(LaTaxeDev/(1.0-RemPied)/(1.0-Escompte),NbDEC) ;
        LaTaxeNR:=Arrondi(LaTaxe/(1.0-RemPied)/(1.0-Escompte),V_PGI.okdecV) ;
      end;
      TotTaxesDEV:=TotTaxes+LaTaxeDev ;
      TotTaxesNRDEV:=TotTaxesNRDEV+LaTaxeNRDEV ;
      TotTaxes:=TotTaxes+LaTaxe;
      TotTaxesNR:=TotTaxesNR+LaTaxeNR;
      TOBL.PutValue(prefixe+'_TOTALTAXEDEV'+IntToStr(i+1),LaTaxeDev) ;
      TOBL.PutValue(prefixe+'_TOTALTAXE'+IntToStr(i+1),LaTaxe) ;
    END ;
  END ;
end;
// --
procedure GestionTaxesLignesTTCInit (TOBPiece,TOBL,TOBTiers : TOB; Base : double ; var TotTaxesDev,TotTaxes: double; DEV : Rdevise; Nbdec:integer);
var enHt : boolean;
    i_montant : integer;
    i : integer;
    LaTaxeDev,laTaxe : double;
    TOBTaxes : TOB;
    Prefixe : string;
begin
  TOBtaxes := TOB.create ('LES TAXES',nil,-1);
  Prefixe := GetPrefixeTable (TOBL);
  EnHT:=(TOBL.GetValue(prefixe+'_FACTUREHT')='X');
  // Taxes
  CalculTaxesLigne(TOBL,TOBPiece,TOBTiers,TOBTaxes,DEV,NbDec,base,EnHt) ;
  // Autres cumuls lignes
  TotTaxes:=0 ;
  TotTaxesDEV:=0 ;
  if TOBTaxes<>Nil then
  BEGIN
    for i:=0 to TOBtaxes.detail.count -1 do
    BEGIN
      LaTaxeDev:=TOBTaxes.Detail[i].GetValue('BLB_VALEURDEV') ;
      LaTaxe:=TOBTaxes.Detail[i].GetValue('BLB_VALEURTAXE') ;
      //
      TotTaxesDEV:=TotTaxes+LaTaxeDev ;
      TotTaxes:=TotTaxes+LaTaxe;
    END ;
  END ;
  TOBTaxes.free;
end;

Procedure CalculeLigneHT ( TOBL,TOBTaxesL,TOBPiece : TOB ; DEV : Rdevise; NbDec : integer ; Forcer:boolean=False ; TOBTiers : TOB = Nil ) ;
Var QteF,PQ,RemLigne,RemPied,Escompte : Double ;
    OldMontant, NewMontant, Ecart : Double;
    TTLoc : TOB ;
    i        : integer ;
    PUHTDEV,PUHT,PUHTNETDEV,PUHTNET,MontantHTDev,MontantHTBrutDev,MontantHTBrut,MontantTTCBrut,TotTaxes,TotTaxesDev,TotTaxesNR,TOTTaxesNrDev : Double ;
    PUTTCDEV,PUTTCNETDEV,TaxeLoc,TotalHT,PUTTC,PUTTCNET,TaxeLocDEV,MontantTTC : Double ;
    prefixe : string;
    RemLigneDev,EscompteDev,TotalHTDev,TotalHTRemisesDev,TotalHTRemises,RemPiedDev,MontantHt,MtRemPiedDev,MtRemPied,MtRemLigne,MtEscompte,MtescompteDev : double ;
    FactureAvanc : boolean;
BEGIN
  //
  factureAvanc := IsgestionAvanc (TOBpiece);
  prefixe := GetPrefixeTable (TOBl);
  if Pos(Prefixe,'GL;BOP')=0 then exit;
  if not Forcer then
  begin
    if (TOBL.GetValue(prefixe+'_TYPEREF') <> 'CAT') and (TOBL.GetValue(prefixe+'_ARTICLE') = '') then Exit ;
  end;
  {$IFDEF BTP}
  if (TOBL.GEtValue(prefixe+'_NATUREPIECEG')='PBT') and (TOBL.GEtValue(prefixe+'_BLOQUETARIF')='X') then
  begin
    if (TOBL.GetValue(prefixe+'_QTEFACT') <> 0) then TOBL.PutValue(prefixe+'_PUHTDEV',TOBL.GetValue(prefixe+'_TOTALHTDEV')/TOBL.GetValue(prefixe+'_QTEFACT'));
  end;
  {$ENDIF}
  ZeroLigne(TOBL) ;
  PQ:=TOBL.GetValue(prefixe+'_PRIXPOURQTE') ; if PQ<=0 then BEGIN PQ:=1.0 ; TOBL.PutValue(prefixe+'_PRIXPOURQTE',PQ) ; END ;
  QteF:=TOBL.GetValue(prefixe+'_QTEFACT') ;
  if (QteF=0) then
  begin
    if (not TOBL.FieldExists('BLF_MTSITUATION')) or  (not (TOBL.GeTDouble('BLF_MTSITUATION')<>0)) then Exit ;
  end;
  PUHTDEV:=TOBL.GetValue(prefixe+'_PUHTDEV') ; // if PUHTDEV=0 then Exit ; PA le 12/10/2001
  // calcul montant achat ligne
  if (PUHTDEV=0) then
  begin
    TOBL.PutValue(prefixe+'_PUHTNETDEV',0); TOBL.PutValue(prefixe+'_PUTTCNETDEV',0);
    TOBL.PutValue(prefixe+'_PUTTCDEV',0);
    TOBL.PutValue(prefixe+'_PUHTNET',0); TOBL.PutValue(prefixe+'_PUTTCNET',0);
    TOBL.PutValue(prefixe+'_PUTTC',0);
    Exit;
  end;
//  DEV.Code:=TOBL.GetValue('GL_DEVISE') ; GetInfosDevise(DEV) ;
  TOBL.putvalue (prefixe+'_PUHT',DEVISETOPIVOTEx(PuHTDEV,TOBL.GetValue(prefixe+'_TAUXDEV'),DEV.quotite,V_PGI.okdecP));
  PUHT:=TOBL.GetValue(prefixe+'_PUHT') ;

  RemLigne:=0 ; Rempied:=0 ; Escompte:=0 ;
  // Informations issues de l'article
  if (not isVariante(TOBL))  then
  begin
    if TOBL.GetValue(prefixe+'_REMISABLELIGNE')='X' then RemLigne:=TOBL.GetValue(prefixe+'_REMISELIGNE')/100.0 ;
    if TOBL.GetValue(prefixe+'_REMISABLEPIED')='X' then RemPied:=TOBL.GetValue(prefixe+'_REMISEPIED')/100.0 ;
    if TOBL.GetValue(prefixe+'_ESCOMPTABLE')='X' then Escompte:=TOBL.GetValue(prefixe+'_ESCOMPTE')/100.0 ;
  end;
  //
//  if (FactureAvanc) and (TOBL.FieldExists ('BLF_MTSITUATION')) and (prefixe <> 'BOP') then
//  if (TOBL.FieldExists ('BLF_NATUREPIECEG')) and (TOBL.GetValue ('BLF_NATUREPIECEG')<>'')  then
  if (TOBL.FieldExists ('BLF_NATUREPIECEG')) and (TOBL.GetValue ('BLF_NATUREPIECEG')<>'') or
  	 ((TOBL.GetDouble('BLF_MTSITUATION')<>0) AND (prefixe='BOP')) then
  begin
  	MontantHTBrutDev := TOBL.GetValue ('BLF_MTSITUATION');
    MontantHTBrut := DeviseToPivot(MontantHtBrutDev,TOBL.GetValue(prefixe+'_TAUXDEV'),DEV.quotite);
    TOBL.putValue(prefixe+'_QTEFACT',TOBL.GetValue ('BLF_QTESITUATION'));
    QteF:=TOBL.GetValue(prefixe+'_QTEFACT') ;
  end else
  begin
    MontantHTBrutDev:=Arrondi(PUHTDEV*QteF/PQ,V_PGI.okdecV) ;
    MontantHTBrut:=Arrondi(PUHT*QteF/PQ,V_PGI.okdecV) ;
  end;

  if TOBL.GetValue(prefixe+'_VALEURREMDEV')<>0 then
  BEGIN
    MtRemLigne:=TOBL.GetValue(prefixe+'_VALEURREMDEV') ;
    MontantHTDev:=MontantHTBrutDev-MtRemLigne ;
    MontantHT:=MontantHTBrut-(DeviseToPivotEx(TOBL.GetValue(prefixe+'_VALEURREMDEV'),TOBL.GetValue(prefixe+'_TAUXDEV'),DEV.quotite,DEV.decimale));
  END else
  BEGIN
    MontantHTDev:=Arrondi(MontantHTBrutDev*(1.0-RemLigne),DEV.decimale) ;
    MtRemLigne:=MontantHTBrutDev-MontantHTDev ;
    MontantHT:=Arrondi(MontantHTBrut*(1.0-RemLigne),V_PGI.OkdecV) ;
  END ;

  MontantHTDev:=Arrondi(MontantHTDev,NbDec) ;
  MontantHT:=Arrondi(MontantHT,V_PGI.OkdecV) ;

  TOBL.PutValue(prefixe+'_MONTANTHTDEV',MontantHTDEV) ;
  TOBL.PutValue(prefixe+'_MONTANTHT',MontantHt);
  // --
  TOBL.PutValue(prefixe+'_TOTREMLIGNEDEV',MtRemLigne) ;
  TOBL.PutValue(prefixe+'_TOTREMLIGNE',DeviseToPivotEx(MtRemLigne,TOBL.GetValue(prefixe+'_TAUXDEV'),DEV.quotite,DEV.decimale));
  if QteF<>0 then
  begin
    PUHTNETDEV:=Arrondi(MontantHTDev/(QteF/PQ),6);
    PUHTNET:=Arrondi(MontantHT/(QteF/PQ),6);
  end else
  begin
    PUHTNETDEV:=Arrondi(PUHTDEV*(1.0-RemLigne),6) ;
    PUHTNET:=Arrondi(PUHTDEV*(1.0-RemLigne),6) ;
  end;
  TOBL.PutValue(prefixe+'_PUHTNETDEV',PUHTNETDEV) ;
  TOBL.PutValue(prefixe+'_PUHTNET',PUHTNETDEV) ;
  if (Pos(TobL.getValue('GL_NATUREPIECEG'),GetPieceAchat (false,false,false,True))>0) and (TOBL.GetValue('BCO_TRAITEVENTE') = 'X' ) then
  begin
    EnregistrePrixLivAModifier (TOBL);
  end else if (Pos(TobL.getValue('GL_NATUREPIECEG'),GetPieceAchat (false,false,false,True))>0) and (TOBL.GetValue('NEW_LIGNE') = 'X' ) then
  begin
    EnregistrePrixLivAModifier (TOBL);
  end;
  // Avec remise pied

(*  TotalHTRemisesDev:=Arrondi(PUHTDev*QteF*(1.0-RemLigne)*(1.0-RemPied)/PQ,NbDec) ;
  TotalHTremises:=Arrondi(PUHT*QteF*(1.0-RemLigne)*(1.0-RemPied)/PQ,V_PGI.OkdecV) ; *)
  TotalHTRemisesDev:=Arrondi(MontantHTBrutDev*(1.0-RemLigne)*(1.0-RemPied),NbDec) ;
  TotalHTremises:=Arrondi(MontantHTBrut*(1.0-RemLigne)*(1.0-RemPied),V_PGI.OkdecV) ;

  MtRemPiedDev:=Arrondi(MontantHTDev-TotalHTRemisesDev,Nbdec) ;
  MtRemPied:=Arrondi(MontantHT-TotalHTRemises,V_PGI.OkdecV) ;
  TOBL.PutValue(prefixe+'_TOTREMPIEDDEV',MtRemPiedDev) ;
  TOBL.PutValue(prefixe+'_TOTREMPIED',MtRemPied) ;
  // Avec escompte
  //TotalHTDev:=Arrondi(PUHTDev*QteF*(1.0-RemLigne)*(1.0-RemPied)*(1.0-Escompte)/PQ,6) ;
  TotalHTDev:=Arrondi(MontantHTDEV*(1.0-RemPied)*(1.0-Escompte),Nbdec) ;
  TotalHT:=Arrondi(MontantHT*(1.0-RemPied)*(1.0-Escompte),V_PGI.okdecV) ;
  //
  MtEscompteDev:=Arrondi(TotalHTRemisesDev-TotalHTDev,Nbdec) ;
  MtEscompte:=Arrondi(TotalHTRemises-TotalHT,V_PGI.OkdecV) ;

  TOBL.PutValue(prefixe+'_TOTESCLIGNEDEV',MtEscompteDev) ;
  TOBL.PutValue(prefixe+'_TOTESCLIGNE',MtEscompte) ;
  TOBL.PutValue(prefixe+'_TOTALHTDEV',TotalHTDev) ;
  TOBL.PutValue(prefixe+'_TOTALHT',TotalHT) ;
  
  // Taxes
  // Modif BTP
  // VARIANTES
  if (not isVariante(TOBL))
  {$IFDEF BTP}
    and (not IsLigneFromCentralis(TOBL))
    and (TOBL.GetValue(prefixe+'_TYPEARTICLE') <> '$#$')
  {$ENDIF} then
  begin
    GestionTaxesLignes (TOBPiece,TOBTaxesL,TOBL,TOBTiers,TotTaxesDev,TotTaxes,TOTTaxesNRDEv,TOTTaxesNR,DEV,Nbdec);
  end else
  begin
    GestionTaxesLignes (TOBPiece,nil,TOBL,TOBTiers,TotTaxesDev,TotTaxes,TOTTaxesNRDEv,TOTTaxesNR,DEV,Nbdec);
  end;
  // ---
  TOBL.PutValue(prefixe+'_TOTALTTCDEV',Arrondi(TotalHTDev+TotTaxesDEV,Nbdec)) ;
  TOBL.PutValue(prefixe+'_TOTALTTC',Arrondi(TotalHT+TotTaxes,V_PGI.OkdecV)) ;
  TOBL.PutValue(prefixe+'_MONTANTTTCDEV',Arrondi(MontantHTDev+TotTaxesNRDEV,Nbdec)) ;
  TOBL.PutValue(prefixe+'_MONTANTTTC',Arrondi(MontantHT+TotTaxesNR,V_PGI.OkdecV)) ;
  if Arrondi(QteF*(1.0-RemLigne),6)<>0 then
  BEGIN
    PUTTCDEV:=TOBL.GetValue(prefixe+'_MONTANTTTCDEV')*PQ/(QteF*(1.0-RemLigne)) ;
    PUTTC:=TOBL.GetValue(prefixe+'_MONTANTTTC')*PQ/(QteF*(1.0-RemLigne)) ;
    // Correction Nb de decimale
    //TOBL.PutValue('GL_PUTTCDEV',Arrondi(PUTTCDEV,NbDec)) ;
    TOBL.PutValue(prefixe+'_PUTTCDEV',Arrondi(PUTTCDEV,V_PGI.OkdecP)) ;
    TOBL.PutValue(prefixe+'_PUTTC',Arrondi(PUTTC,V_PGI.OkdecP)) ;
    PUTTCNETDEV:=TOBL.GetValue(prefixe+'_MONTANTTTCDEV')*PQ/QteF ;
    PUTTCNET:=TOBL.GetValue(prefixe+'_MONTANTTTC')*PQ/QteF ;
    //TOBL.PutValue('GL_PUTTCNETDEV',Arrondi(PUTTCNETDEV,NbDec)) ;
    TOBL.PutValue(prefixe+'_PUTTCNETDEV',Arrondi(PUTTCNETDEV,V_PGI.okdecP)) ;
    TOBL.PutValue(prefixe+'_PUTTCNET',Arrondi(PUTTCNET,V_PGI.okdecP)) ;
  END else
  BEGIN
    TTLoc := TOB.create ('LES TAXES LIGNES',nil,-1);
    CalculTaxesLigne(TOBL,TOBPiece,TOBTiers,TTLoc,DEV,NbDec,TOBL.GetValue(prefixe+'_PUHTDEV'),True) ;
    TaxeLoc:=0 ; TaxeLocDev:=0 ;
    if TTLoc<>Nil then for i:=0 to TTloc.detail.count -1 do
    begin
      TaxeLocDev:=TaxeLocDEV+TTLoc.Detail[i].GetValue('BLB_VALEURDEV') ;
      TaxeLoc:=TaxeLoc+TTLoc.Detail[i].GetValue('BLB_VALEURTAXE') ;
    end;
    PUTTCDEV:=TOBL.GetValue(prefixe+'_PUHTDEV')+TaxeLocDEV ; PUTTCNETDEV:=PUTTCDEV*(1.0-RemLigne) ;
    PUTTC:=TOBL.GetValue(prefixe+'_PUHT')+TaxeLoc ; PUTTCNET:=PUTTC*(1.0-RemLigne) ;
    TOBL.PutValue(prefixe+'_PUTTCDEV',Arrondi(PUTTCDEV,V_PGI.OkDecP)) ;
    TOBL.PutValue(prefixe+'_PUTTC',Arrondi(PUTTC,V_PGI.OkDecP)) ;
    TOBL.PutValue(prefixe+'_PUTTCNETDEV',Arrondi(PUTTCNETDEV,V_PGI.OkDecP)) ;
    TOBL.PutValue(prefixe+'_PUTTCNET',Arrondi(PUTTCNET,V_PGI.OkDecP)) ;
    TTloc.free;
  END ;
  // total TTC Dans LIGNEFAC
  TOBL.PutValue('BLF_TOTALTTCDEV',TOBL.GetValue(prefixe+'_TOTALTTC'));
  if TOBL.getValue(prefixe+'_DPR')<> 0 then
  begin
    TOBL.PutValue(prefixe+'_COEFMARG',Arrondi(TOBL.GetValue(prefixe+'_PUHT')/TOBL.getValue(prefixe+'_DPR'),4));
    TOBL.PutValue('POURCENTMARG',Arrondi((TOBL.GetValue(prefixe+'_COEFMARG')-1)*100,2));
  end;
  //
  if TOBL.GetValue(prefixe+'_PUHT') <> 0 then
  begin
    TOBL.PutValue('POURCENTMARQ',Arrondi(((TOBL.GetValue(prefixe+'_PUHT')- TOBL.GetValue(prefixe+'_DPR'))/TOBL.GetValue(prefixe+'_PUHT'))*100,2));
  end;
  StockeMontantTypeSurLigne (TOBL);
END ;

Procedure CalculeLigneTTC ( TOBL,TOBTaxesL,TOBPiece : TOB ; DEV : Rdevise; NbDec : integer; Forcer:boolean=False ; TOBTiers : TOB = Nil ) ;
Var QteF,PQ,RemLigne,RemPied,Escompte : Double ;
    TTLoc : TOB ;
    i        : integer ;
    PUHTDEV,PUHTNETDEV,MontantTTCDev,MontantTTCBrutDev,TotTaxes,TotTAxesDev,TotTaxesNR,TotTaxesNRDEV,RemPiedTTC : Double ;
    PUTTCDEV,PUTTCNETDEV,TaxeLoc,TotalTTCDev : Double ;
    RemLigneDevTTC,EscompteTTCDev,TotalTTCRemises,RemPiedTTCDev,RemLigneTTC,PUTTC,MontantTTCBrut,MontantTTC,PUTTCNET,TotalTTCRemisesDEV : double ;
    TotalTTC,PUHT,PUHTNET,TaxeLocDev,EscompteTTC : double;
    prefixe : string;
    MontantHtDev,MontantHt,TOtalHtRemisesDev,TotalHtRemises,MtRemPiedDev,MtRemPied,TotalHtDEv,TotalHt,MtEscompteDev,MtEscompte,TTCDevEscompte,TTCEscompte : double;
BEGIN
  prefixe := GetPrefixeTable (TOBl);
  if not Forcer then
  begin
    {$IFNDEF CHR}
    if (TOBL.GetValue(prefixe+'_TYPEREF') <> 'CAT') and (TOBL.GetValue(prefixe+'_ARTICLE') = '') then Exit ;
    {$ELSE}
    if (TOBL.GetValue(prefixe+'_TYPEREF') <> 'CAT') and (TOBL.GetValue(prefixe+'_ARTICLE') = '') and (TOBL.GetValue(prefixe+'_TYPENOMENC') <> 'MAC') then Exit;
    {$ENDIF}
  end;
  // Modif BTP
  ZeroLigne(TOBL) ; InitLesSupMontants(TOBL) ;
  QteF:=TOBL.GetValue(prefixe+'_QTEFACT') ; if QteF=0 then Exit ;
  PUTTCDEV:=TOBL.GetValue(prefixe+'_PUTTCDEV') ; // if PUTTCDEV=0 then Exit ;  PA le 12/10/2001
  if (PUTTCDEV=0) then
  begin
    TOBL.PutValue(prefixe+'_PUHTDEV',0); TOBL.PutValue(prefixe+'_PUHTNETDEV',0);
    TOBL.PutValue(prefixe+'_PUTTCNETDEV',0);
    Exit;
  end;
  TOBL.putvalue (prefixe+'_PUTTC',devisetopivotEx(PuTTCDEV,TOBL.GetValue(prefixe+'_TAUXDEV'),DEV.quotite,V_PGI.okdecP));
  PUTTC := TOBL.Getvalue (prefixe+'_PUTTC');
  //
  PQ:=TOBL.GetValue(prefixe+'_PRIXPOURQTE') ; if PQ<=0 then BEGIN PQ:=1.0 ; TOBL.PutValue(prefixe+'_PRIXPOURQTE',PQ) ; END ;
  RemLigne:=0 ; RemPied:=0 ; Escompte:=0 ;
  // Informations issues de l'article
  if (not isVariante(TOBL)) {$IFDEF BTP} and (not IsLigneFromCentralis(TOBL)){$ENDIF} then
  begin
(*    if TOBL.GetValue(prefixe+'_REMISABLELIGNE')='X' then RemLigne:=TOBL.GetValue(prefixe+'_REMISELIGNE')/100.0 else TOBL.PutValue(prefixe+'_REMISELIGNE',0) ; *)
    if TOBL.GetValue(prefixe+'_REMISABLEPIED')='X' then RemPied:=TOBL.GetValue(prefixe+'_REMISEPIED')/100.0  else TOBL.PutValue(prefixe+'_REMISEPIED',0) ;
    if TOBL.GetValue(prefixe+'_ESCOMPTABLE')='X' then Escompte:=TOBL.GetValue(prefixe+'_ESCOMPTE')/100.0  else TOBL.PutValue(prefixe+'_ESCOMPTE',0) ;
  end;
  // Avec remise ligne
  MontantTTCBrutDev:=Arrondi(PUTTCDEV*QteF/PQ,Nbdec) ;
  MontantTTCBrut:=Arrondi(PUTTC*QteF/PQ,V_PGI.okdecV) ;
  //
  MontantTTCDev:=Arrondi(MontantTTCBrutDev,NbDec) ;
  MontantTTC:=Arrondi(MontantTTCBrut,V_PGI.okdecV) ;
  TOBL.PutValue(prefixe+'_MONTANTTTCDEV',MontantTTCDEV) ;
  TOBL.PutValue(prefixe+'_MONTANTTTC',MontantTTC) ;
  //
  TotalTTCDev:=Arrondi(MontantTTCBrutDev*(1-(Escompte))*(1-(RemPied)),NbDec) ;
  TotalTTC:=Arrondi(MontantTTCBrut*(1-(Escompte))*(1-(Rempied)),V_PGI.okdecV) ;
  TOBL.PutValue(prefixe+'_TOTALTTCDEV',TotalTTCDEV) ;
  TOBL.PutValue(prefixe+'_TOTALTTC',TotalTTC) ;
  //FV1 : Zone inexistante dans la tob mais GL_TOTALTTCDEV existe ==> chargement de TOTALTTCDEV !!!!
  //TOBL.PutValue('BLF_TOTALTTCDEV',TOBL.GetValue(prefixe+'_TOTALTTC'));
  TOBL.PutValue(prefixe+'_TOTALTTCDEV',totalTTC);
// Avec escompte
  EscompteTTCDev:=Arrondi(MontantTTCDEv-(MontantTTCDEV*(1.0-Escompte)),Nbdec) ;
  TOBL.PutValue('TOTESCLIGNETTCDEV',EscompteTTCDev) ;
  TTCDevEscompte := MontantTTCDev - EscompteTTCDev;
  //
  EscompteTTC:=Arrondi(MontantTTC-(MontantTTC*(1.0-Escompte)),V_PGI.okdecV) ;
  TOBL.PutValue('TOTESCLIGNETTC',EscompteTTC) ;
  TTCEscompte := MontantTTC - EscompteTTC;
// Avec remise pied
  RemPiedTTCDev:=Arrondi(TTCDevEscompte - ((TTCDevEscompte)*(1.0-RemLigne)*(1.0-RemPied))/PQ,Nbdec) ;
  TOBL.PutValue('TOTREMPIEDTTCDEV',RemPiedTTCDev) ;
  RemPiedTTC:=Arrondi(TTCEscompte - ((TTCEscompte)*(1.0-RemLigne)*(1.0-RemPied))/PQ,V_PGI.okdecV) ;
  TOBL.PutValue('TOTREMPIEDTTC',RemPiedTTC) ;
  //
  if QteF<>0 then PUTTCNETDEV:=Arrondi(MontantTTCDev/(QteF/PQ),V_PGI.okdecP) else PUTTCNETDEV:=Arrondi(PUTTCDEV*(1.0-RemLigne),V_PGI.OKdecP) ;
  if QteF<>0 then PUTTCNET:=Arrondi(MontantTTC/(QteF/PQ),V_PGI.okdecP) else PUTTCNET:=Arrondi(PUTTC*(1.0-RemLigne),V_PGI.OKdecP) ;
  TOBL.PutValue(prefixe+'_PUTTCNETDEV',PUTTCNETDEV) ;
  TOBL.PutValue(prefixe+'_PUTTCNET',PUTTCNET) ;

  // recalcul du HT via le TTC
  GestionTaxesLignesTTCInit (TOBpiece,TOBL,TOBtiers,MontantTTCDev,TaxeLocDev,TaxeLoc,DEV,Nbdec);
  // --
  MontantHtDev := MontantTTCDEV-Arrondi(TaxeLocDev,Dev.decimale);
  MontantHt := MontantTTC-Arrondi(TaxeLoc,V_PGI.okdecV);

  TOBL.PutValue(prefixe+'_MONTANTHTDEV',MontantHtDev) ;
  TOBL.PutValue(prefixe+'_MONTANTHT',MontantHt) ;

  // --- GUINIER ---
  if CtrlOkReliquat(TOBL, 'GL') then TOBL.PutValue(Prefixe+'_MTRESTE', TOBL.GetValue(Prefixe+'_MONTANTHTDEV'));

  PUHTDEV:=MontantHtDev*PQ/(QteF*(1.0-RemLigne)) ;
  PUHT:=MontantHt*PQ/(QteF*(1.0-RemLigne)) ;
  //
  TOBL.PutValue(prefixe+'_PUHTDEV',Arrondi(PUHTDEV,V_PGI.okdecP)) ;
  TOBL.PutValue(prefixe+'_PUHT',Arrondi(PUHT,V_PGI.okdecP)) ;
  if TOBL.getValue(prefixe+'_DPR')<> 0 then
  begin
    TOBL.PutValue(prefixe+'_COEFMARG',Arrondi(TOBL.GetValue(prefixe+'_PUHT')/TOBL.getValue(prefixe+'_DPR'),4));
    TOBL.PutValue('POURCENTMARG',Arrondi((TOBL.GetValue(prefixe+'_COEFMARG')-1)*100,2));
  end;
  if TOBL.GetValue(prefixe+'_PUHT') <> 0 then
  begin
    TOBL.PutValue('POURCENTMARQ',Arrondi(((TOBL.GetValue(prefixe+'_PUHT')- TOBL.GetValue(prefixe+'_DPR'))/TOBL.GetValue(prefixe+'_PUHT'))*100,2));
  end else
  begin
    TOBL.PutValue('POURCENTMARQ',0);
  end;
  //
  PUHTNETDEV:=MontantHtdev*PQ/QteF ;
  PUHTNET:=MontantHt*PQ/QteF ;
  //
  TOBL.PutValue(prefixe+'_PUHTNETDEV',Arrondi(PUHTNETDEV,V_PGI.okdecP)) ;
  TOBL.PutValue(prefixe+'_PUHTNET',Arrondi(PUHTNET,V_PGI.okdecP)) ;
  //
  TOBL.PutValue(prefixe+'_TOTREMLIGNEDEV',RemLigneDevTTC) ;    // NA : 22/02/01
  TOBL.PutValue(prefixe+'_TOTREMLIGNE',RemLigneTTC) ;    // NA : 22/02/01
  // Avec remise pied
  TotalHTRemisesDev:=Arrondi(MontantHtDev*(1.0-RemLigne)*(1.0-RemPied)/PQ,NbDec) ;
  TotalHTremises:=Arrondi(MontantHt*(1.0-RemLigne)*(1.0-RemPied)/PQ,V_PGI.OkdecV) ;
  //
  MtRemPiedDev:=Arrondi(MontantHTDev-TotalHTRemisesDev,Nbdec) ;
  MtRemPied:=Arrondi(MontantHT-TotalHTRemises,V_PGI.okdecV) ;
  //
  TOBL.PutValue(prefixe+'_TOTREMPIEDDEV',MtRemPiedDev) ;
  TOBL.PutValue(prefixe+'_TOTREMPIED',MtRemPied) ;
  //
  TotalHTDev:=Arrondi(MontantHTDEV*(1.0-RemPied)*(1.0-Escompte),Nbdec) ;
  TotalHT:=Arrondi(MontantHT*(1.0-RemPied)*(1.0-Escompte),V_PGI.okdecV) ;
  //
  TOBL.PutValue(prefixe+'_TOTALHTDEV',TotalHTDev) ;
  TOBL.PutValue(prefixe+'_TOTALHT',TotalHT) ;
  //
  MtEscompteDev:=Arrondi(TotalHTRemisesDev-TotalHTDev,Nbdec) ;
  MtEscompte:=Arrondi(TotalHTRemises-TotalHT,V_PGI.OkdecV) ;
  TOBL.PutValue(prefixe+'_TOTESCLIGNEDEV',MtEscompteDev) ;
  TOBL.PutValue(prefixe+'_TOTESCLIGNE',MtEscompte) ;

  if CtrlOkReliquat(TOBL, 'GL') then TOBL.PutValue(Prefixe+'_MTRESTE', TOBL.GetValue(Prefixe+'_MONTANTHTDEV'));

  if (not isVariante(TOBL))
  {$IFDEF BTP}
    and (not IsLigneFromCentralis(TOBL))
    and (TOBL.GetValue(prefixe+'_TYPEARTICLE') <> '$#$')
  {$ENDIF} then
  begin
    GestionTaxesLignes (TOBPiece,TOBTaxesL,TOBL,TOBTiers,TotTaxesDEV,TotTaxes,TOTTaxesNRDEV,TOTTaxesNR,DEV, Nbdec);
  end else
  begin
    GestionTaxesLignes (TOBPiece,nil,TOBL,TOBTiers,TotTaxesDEV,TotTaxes,TOTTaxesNRDEV,TOTTaxesNR,DEV, Nbdec);
  end;
  // ------------------------
  StockeMontantTypeSurLigne (TOBL);
END ;

// Modif BTP
procedure CalculBasePourcent (var TOBPiece,TOBS: TOB; R_Valeurs : R_CPercent);
var
   TOBL : TOB;
   Indice,INdL,IndA : Integer;
   TypL,TypA : String;
   dpaLigne,DprLigne,PmapLigne,PmrpLigne,Dqte,DPrixQte,DPuht,DPuTtc : double;
   IDpa,IDpr,Ipmap,Ipmrp,iqte,IPrixQte,IPuHt,IpuTTC,iMontHt,iMontTTC : integer;
   TypeArtOk : string;
begin
ZeroLignePourcent(TOBS) ;
TypeArtOk := TOBS.GetValue('GL_LIBCOMPL');
for Indice := R_Valeurs.Depart to R_Valeurs.Fin do
    begin
    TOBL := GetTOBLigne(TOBPiece,Indice+1 ) ; if TOBL = nil then break;
    if Indice=R_Valeurs.Depart then
       begin
       IndL := TOBL.GetNumChamp ('GL_TYPELIGNE');
       IndA := TOBL.GetNumChamp ('GL_TYPEARTICLE');
       Idpa := TOBL.GetNumChamp ('GL_DPA');
       Idpr := TOBL.GetNumChamp ('GL_DPR');
       Ipmap := TOBL.GetNumChamp ('GL_PMAP');
       Ipmrp := TOBL.GetNumChamp ('GL_PMRP');
       Iqte := TOBL.GetNumChamp ('GL_QTEFACT');
       IPrixQte := TOBL.GetNumChamp ('GL_PRIXPOURQTE');
       IPuHt := TOBL.GetNumChamp ('GL_PUHTDEV');
       IPuTTC := TOBL.GetNumChamp ('GL_PUTTCDEV');
       IMontHt := TOBL.GetNumChamp ('GL_MONTANTHTDEV');
       IMontTTC := TOBL.GetNumChamp ('GL_MONTANTHTDEV');
       end;
    TypL:=TOBL.GetValeur(IndL) ;
    TypA:=TOBL.GetValeur(IndA) ;
    Dqte := TOBL.GetValeur(Iqte);
    DPrixQte := TOBL.GetValeur(IPrixQte); if DprixQte = 0 then DPrixQte := 1;
    DPuHt := TOBL.GetValeur(IpuHt);
    DPuTTC := TOBL.GetValeur(IpuTTC);
    // VARIANTE
(*    if ((copy(TypL,1,2) = 'TP')OR
       (Copy(Typl,1,2)='DP') or
       (Typl = 'TOT') or
       (Typl = 'COM')) or
       (TOBL.getvalue('GL_ARTICLE') = '' ) then continue; *)

    if ((copy(TypL,1,2) = 'TP')OR
       (Copy(Typl,1,2)='DP') or
       (IsVariante(TOBL)) or
{$IFDEF BTP}
       (TOBL.GetValue('GL_TYPEARTICLE') = '$#$') or
			 (IsLigneFromCentralis(TOBL)) or
{$ENDIF}
       (Typl = 'TOT') or
       (Typl = 'COM')) or
       (TOBL.getvalue('GL_ARTICLE') = '' ) then continue;
		if IsSousDetail(TOBL) then Continue; 
    if  not ArticleOKInPOUR (Typa,TypeartOk) then continue;
    if (TypA='POU') and (Indice>=R_Valeurs.Traite ) then continue;
    TOBS.PutValeur(IPuht,TOBS.GetValeur(IpuHt)+TOBL.GetValeur(IMontHt)) ;
    TOBS.PutValeur(IPuttc,TOBS.GetValeur(IpuTTC)+TOBL.GetValeur(IMontTTC)) ;
    // Calcul des elements achats
    DpaLigne := TOBL.getValeur(idpa) * (Dqte / DPrixQte);
    DprLigne := TOBL.getValeur(idpr) * (Dqte / DPrixQte);
    PmapLigne := TOBL.getValeur(ipmap) * (Dqte / DPrixQte);
    PmrpLigne := TOBL.getValeur(ipmrp) * (Dqte / DPrixQte);
    // Cumule les elements achats
    TOBS.PutValeur(Idpa,TOBS.GetValeur(Idpa)+DpaLigne) ;
    TOBS.PutValeur(Idpr,TOBS.GetValeur(Idpr)+DprLigne) ;
    TOBS.PutValeur(ipmap,TOBS.GetValeur(Ipmap)+PmapLigne) ;
    TOBS.PutValeur(ipmrp,TOBS.GetValeur(Ipmrp)+PmrpLigne) ;
    end;
if (GetInfoParPiece(TOBPiece.GetValue('GP_NATUREPIECEG'),'GPP_ESTAVOIR')='X') then
   BEGIN
   TOBS.PutValeur(IPuht,Abs(TOBS.GetValeur(IpuHt))) ;
   TOBS.PutValeur(IPuttc,Abs(TOBS.GetValeur(IpuTTC))) ;
   END;

end;

procedure CalculLignePourcent (TOBPiece,TOBL:TOB;I:Integer);
var T_CPourcent :R_CPercent;
begin
if TOBL.getValue ('GL_TYPEARTICLE') = 'POU' then
   BEGIN
       // Calcule Pourcentage
   T_CPourcent.niveau := TOBL.getvalue ('GL_NIVEAUIMBRIC');
   T_CPourcent.Traite := I;
   T_CPourcent.Depart := RecupDebutPourcent (TOBPiece,I - 1,T_CPourcent.niveau);
   if I > 0 then
      T_CPourcent.Fin := I - 1
   else
      T_CPourcent.Fin := 0;
   TOBL.PutValue ('GL_RECALCULER','X');
   CalculBasePourcent (TOBPiece,TOBL,T_CPourcent);
   END;
end;

procedure recalculeRGLoc (TOBtiers,TOBPiece,TOBPieceTrait,TOBBasesL,TOBPieceRG,TOBBasesRG,TOBporcs,TOBBases: TOB;DEV:RDevise;FromGenFac:boolean);
var TOBL : TOB;
    I : Integer;
    T_RG :R_CPercent;
begin

  if not assigned(TOBPiece) then exit;

  TOBL := TOBPiece.findfirst(['GL_TYPELIGNE'],['RG'],true);
	repeat
    if assigned(TOBL) then
    begin
      // Calcul Retenue de garantie dans le cas de document regroup
      I:=TOBL.GetIndex;
      T_RG.Traite := I;
      T_RG.Depart := RecupDebutRG (TOBPiece,I - 1);
      if I > 0 then
        T_RG.Fin := I - 1
      else
        T_RG.Fin := 0;
      CalculBaseRG (TOBTiers,TOBPiece,TOBPieceTrait,TOBBasesL,TOBPieceRg,TOBBasesRg,TOBL,tobporcs,TOBBASES,T_RG,DEV,FromGenFac);
      TOBL := TOBPIECE.findnext(['GL_TYPELIGNE'],['RG'],true);
    end;
  until TOBL = nil;
end;
// ------------------------------

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Louis DECOSSE  -- Modifie par LS -- Puis par FV
Cr le ...... : 19/01/2000
Modifi le ... : 07/02/2007
Description .. : LE Calcul de la pice
Mots clefs ... : SAISIE;CALCUL;
*****************************************************************}
Procedure CalculFacture ( TOBAffaire,TOBPiece,TOBPieceTrait,TOBSSTRAIT,TOBouvrages,TOBouvragesP,TOBBases,TOBBasesL,TOBTiers,TOBArticles,TOBPorcs,TOBPieceRG,TOBBasesRG,TOBVTECOLL : TOB ; DEV : RDEVISE ;SaisieTypeAvanc : boolean=false; Action: TActionFiche=TaModif; ModifAvanc : boolean=false; LigneACalcule : integer=-1; FromGenfac : boolean=false; FromZero : boolean=false) ;

  procedure CalculeThisLigneOUv (TOBL : TOB; VenteAchat : string; IsDetailleCollectif : Boolean;  ImpacteTotauxPiece : boolean=true);
  var TOBPLat,TOBB,TOBOL,TOBTaxesL : TOB;
      NumOrdre,NumLigne,Indice,NbDec,IndTypeLigne : integer;
      EnHt : boolean;
  begin
    IndTypeLigne    := TOBL.GetNumChamp ('GL_TYPELIGNE');
    NbDec:=DEV.Decimale ;
    EnHT:=(TOBPiece.GetValue('GP_FACTUREHT')='X') ;
    TOBTaxesL := TOB.Create ('TAXES LIGNE',nil,-1);
    NumOrdre := TOBL.getValue('GL_NUMORDRE');
    //
    NumLigne := TOBL.getValue('GL_NUMLIGNE');

    TOBPLat := FindOuvragesPlat (TOBL,TOBOuvragesP);
    if TOBPlat = nil then
    begin
      TOBPlat := AddMereLignePlat (TOBOuvragesP,TOBL.GetValue('GL_NUMORDRE'));
    end else
    begin
      TOBPlat.ClearDetail;
    end;
    GetOuvragePlat (TOBpiece,TOBL,TOBOuvrages,TOBPlat,TOBTiers,DEV,(TOBPiece.getvalue('GP_NATUREPIECEG')<>'DBT'));
    ZeroLigneMontant (TOBL);  // reinit des montants de la ligne
    if TOBL.FieldExists('BLF_MTSITUATION') then
    begin
      if (TOBL.GetValue ('BLF_MTSITUATION') <> 0) and
         (TOBL.GetString('BLF_FOURNISSEUR')<>'') and
         (TOBL.GetString('BLF_NATURETRAVAIL')='002') and
         (GetPaiementSSTrait (TOBSSTRAIT,TOBL.GetString('BLF_FOURNISSEUR'))='001') then
      begin
        // cas du paiement direct d'un sous traitant
        TOBL.SetDouble('BLO_MONTANTPA',TOBL.Getdouble ('BLF_MTSITUATION'));
      end;
    end;
    TOBB := FindLignesbases (NumOrdre,TOBBasesL);
    if TOBB<>nil then TOBB.clearDetail;
    for Indice := 0 to TOBPLat.detail.count -1 do
    begin
      TOBOL := TOBPlat.detail[Indice];
      TOBTaxesL.clearDetail;
      ZeroLigneMontant (TOBOL);  // reinit des montants de la ligne
      if EnHT then
      BEGIN
        CalculeLigneHT(TOBOL,TOBTaxesL,TOBPiece,DEV, NbDec,False,TOBTiers) ;
        if {(VenteAchat = 'VEN') and} (IsDetailleCollectif) then CumuleCollectifs(TOBOL,TOBTaxesL,TOBVTECOLL,TOBSSTRAIT,TOBArticles,TOBTiers,TOBAffaire);
        ChangeParentLignesBases (TOBBasesL,TOBtaxesL,NBdec);
        //
        SommeLignePlat(TOBOL,TOBL,EnHT) ;
      END else
      BEGIN
        CalculeLigneTTC(TOBOL,TOBTaxesL,TOBPiece,DEV, NbDec,False,TOBTiers) ;
        if {(VenteAchat = 'VEN') and }(IsDetailleCollectif) then CumuleCollectifs(TOBOL,TOBTaxesL,TOBVTECOLL,TOBSSTRAIT,TOBArticles,TOBTiers,TOBAffaire);
        ChangeParentLignesBases (TOBBasesL,TOBtaxesL,NBdec);
        //
        SommeLignePlat(TOBOL,TOBL,EnHT) ;
      END ;
      (*
      if (not IsVariante(TOBL)) and (TOBL.GetValeur(IndTypeLigne)<>'CEN') then
      begin
        CalcLigneMesure(TOBOL) ;
        if ImpacteTotauxPiece then
        begin
          SommeLigneMesures(TOBOL,TOBPiece) ;
      end;
    end;
      *)
    end;
    // cumul final
    TOBB := FindLignesbases (NumOrdre,TOBBasesL);
    if TOBB <> nil then
    begin
      ReajusteLigneViaBase (TOBL,TOBB,EnHt,DEV);
    end;
    if (not IsVariante(TOBL)) and (TOBL.GetValeur(IndTypeLigne)<>'CEN') and (ImpacteTotauxPiece) then
    begin
      if EnHt then
      begin
        SommeLignePiedHT(TOBL,TOBPiece) ;
      end else
      begin
        SommeLignePiedTTC(TOBL,TOBPiece) ;
      end;
      if TOBB <> nil then CumuleLesBases(TOBBases,TOBB,NbDec);
    end;
    TOBTaxesL.free;
    CalcLigneMesure (TOBL);
    SommeLigneMesures(TOBL,TOBPiece) ;
    RecalculeTotauxTypes (TOBL);
  end;

  procedure SupprimeEcart (TOBOUV : TOB; ArtEcart : string) ;
  var II : integer;
  begin
    II := 0;
    repeat
      if trim(TOBOUV.detail[II].getString('BLO_CODEARTICLE')) = trim(ArtEcart) then
      begin
        TOBOUV.detail[II].free;
        break;
      end;
      inc(II);
    until II > TOBOUV.detail.count -1;
  end;


Var i,NbDec,NbLig,Indice : integer ;
    TOBL,TOBB,TOBOUV : TOB ;
    EnHT : Boolean ;
    MontantCum : double;
    TypeFacturation : string;
    TOBPlat,TOBOL,TOBTaxesL : TOB;
    IndMontantligne,IndTypeLigne,IndNumOrdre,IndIndiceNomen,IndTypeArticle,IndRecalculer : integer;
    NumOrdre,NumLigne,Ligne : integer;
    IndiceDepart,IndiceArrive,Indrecalc,IndiceNomen : integer;
    ModeAvancement : boolean;
    Qte,NewPu,NewPr,DelTa : double;
    ArtEcart : string;
    VenteAchat : string;
    IsDetailleCollectif : boolean;
BEGIN
  //
  VenteAchat:=GetInfoParPiece(TOBPiece.GetValue('GP_NATUREPIECEG'),'GPP_VENTEACHAT') ;
  IsDetailleCollectif := (getparamSocSecur('SO_BTVENTCOLLECTIF',false)) and (IsComptaPce(TOBPiece.GetValue('GP_NATUREPIECEG')));
  //
  ArtEcart := trim(GetParamsoc('SO_BTECARTPMA'));

	AssigneDocumentTva2014 (TOBpiece,TOBSSTrait);
  //
  if LigneACalcule = -1 then
  begin
    InDiceDepart := 0; IndiceArrive := TOBPiece.detail.count -1;
  end else
  begin
    InDiceDepart := LigneACalcule-1; IndiceArrive := LigneACalcule-1;
      end;
  TypeFacturation := RenvoieTypeFact(TOBPiece.GetValue('GP_AFFAIREDEVIS'));
  ModeAvancement := (Pos(TypeFacturation,'AVA;DAC')>0);
  {JLD 23/09/2003 : bricole infme mais sans risque a priori}
  if uppercase(Copy(GetParamSoc('SO_NIF'),1,2)) = 'ES' then PremiereTaxeSurLesAutres:=False ;
  {Fin modif JLD}
  {MontantBaseTPF:=0 ; MontantHTLignes:=0 ; MontantTTCLignes:=0 ; }NbDec:=DEV.Decimale ;
  EnHT:=(TOBPiece.GetValue('GP_FACTUREHT')='X') ;
  if TOBPiece=Nil then Exit ;
  // initialisation
  IndNumOrdre := -1;
  IndMontantligne := -1;
  TOBTaxesL := TOB.Create ('TAXES LIGNE',nil,-1);
  TOBTaxesL.AddChampSupValeur ('UN CHAMP',''); // pour le tob debug
  //
  if TOBPiece.detail.count > 0 then
  begin
    IndMontantLigne := TOBpiece.detail[0].GetNumChamp ('GL_MONTANTHT');
    IndTypeArticle  := TOBpiece.detail[0].GetNumChamp ('GL_TYPEARTICLE');
    IndTypeLigne    := TOBpiece.detail[0].GetNumChamp ('GL_TYPELIGNE');
    IndNumOrdre     := TOBpiece.detail[0].GetNumChamp ('GL_NUMORDRE');
    IndIndiceNomen  := TOBpiece.detail[0].GetNumChamp ('GL_INDICENOMEN');
    IndRecalculer   := TOBpiece.detail[0].GetNumChamp ('GL_RECALCULER');
  end else exit;
  //
  if (not SaisieTypeAvanc) then
  begin
    DeduitDeltaBases (TOBpiece,TOBBases,EnHt);
//    AddLesports(TOBpiece,TOBPorcs,TOBBases,TOBBasesL,TOBTiers,DEV,Action,'-');
  end;
  //
  for i:=IndiceDepart to IndiceArrive do
  BEGIN
    TOBL:=TOBPiece.Detail[i] ;

    if TOBL=Nil then Break ;
    //
    NumOrdre := TOBL.getValue('GL_NUMORDRE');
    //
    NumLigne := TOBL.getValue('GL_NUMLIGNE');
    //
    if TOBL.GetValue('GL_RECALCULER')<>'X' then continue;
    if IsSousDetail(TOBL) then continue;
    //
    if (not SaisieTypeAvanc) and (Pos(TOBL.GetValeur(IndTypeLigne),'ART;ARV;CEN')>0) and (not IsLigneFromCentralis(TOBL)) then
    begin
      Ligne := TOBL.GetValue('GL_NUMLIGNE');
      // -------------------------------------------------------
      // Deduction des montants ligne avant l'ajout de celle-ci recalcule
      // -------------------------------------------------------
      if (TOBL.GetValeur(IndTypeLigne)='CEN') OR ((TOBL.GetValeur(IndTypeLigne)='ART') AND (not IsLigneFromCentralis(TOBL))) and (not FromZero) and (not FromGenFac) then
      begin
        DeduitLigneModifie (TOBL,TOBpiece,TOBPieceTrait,TOBouvrages,TOBOuvragesP,TOBBases,TOBBasesL,TOBTiers,DEV, Action);
      end;
      // --------
      if IsOuvrage (TOBL) and (TOBL.GetValeur(IndIndiceNomen)<>0) and (TOBL.GetString('GL_NATUREPIECEG')<>'BBO') then
      begin
        if TOBL.getBoolean('GL_GESTIONPTC') and (Pos(TOBPiece.getString('GP_NATUREPIECEG'),'DBT;ETU')>0) then
        begin
          IndiceNomen := TOBL.getInteger('GL_INDICENOMEN');
          TOBOUV := nil;
          if (IndiceNomen > 0) and (IndiceNomen <= TOBOuvrages.detail.count) then
          begin
            TOBOUV := TOBOuvrages.detail[IndiceNomen-1];
            SupprimeEcart(TOBOUV,ArtEcart);
          end;
        end;
        CalculeThisLigneOUv (TOBL,VenteAchat, IsDetailleCollectif, not TOBL.getBoolean('GL_GESTIONPTC')); // on precalcule la ligne sans impatcer les montant de la piece
        if TOBL.getBoolean('GL_GESTIONPTC') and (Pos(TOBPiece.getString('GP_NATUREPIECEG'),'DBT;ETU')>0) then
        begin
          Qte := TOBL.GetDouble('GL_QTEFACT'); if QTe = 0 then Qte := 1;
          NewPu := Arrondi(TOBL.GetDouble('MONTANTPTC')/Qte,DEV.decimale);
          if NewPu <> TOBL.getDouble('GL_PUHTDEV') then
          begin
            TraitePrixOuvrage(TOBPiece,TOBL,TOBBases,TOBBasesL,TOBouvrages, EnHt, NewPu,0,DEV,false);
            // if TOBOUV <> nil then ReajusteMontantOuvrage(nil,TOBPiece, TOBL, TOBOUV, TOBL.GetDouble('MONTANTPTC'), NewPr, NewPu, DEV, EnHt,false);
            TOBL.SetDOuble('GL_PUHTDEV',newPu); TOBL.SetDouble('GL_PUHT',DEVISETOPIVOTEx (newPu,DEV.Taux,DEV.Quotite,V_PGI.okdecP));
            ReinitCoefMarg (TOBL,TOBouvrages );
          end;
          CalculeThisLigneOUv (TOBL,VenteAchat,IsDetailleCollectif);
        end;
        CumuleTypeSurPiece (TOBPiece,TOBL);
      end else
      begin
        TOBTaxesL.clearDetail;
        if TOBL.GetValeur (IndTypeArticle) = 'POU' then CalculLignePourcent (TOBPiece,TOBL,I);
        if EnHT then
        BEGIN
          if TOBL.FieldExists('BLF_MTSITUATION') then
          begin
            if (TOBL.GetValue ('BLF_MTSITUATION') <> 0) and
               (TOBL.GetString('BLF_FOURNISSEUR')<>'') and
               (TOBL.GetString('BLF_NATURETRAVAIL')='002') and 
               (GetPaiementSSTrait (TOBSSTRAIT,TOBL.GetString('BLF_FOURNISSEUR'))='001') then
            begin
              // cas du paiement direct d'un sous traitant
              TOBL.SetDouble('GL_MONTANTPA',TOBL.Getdouble ('BLF_MTSITUATION'));
              TOBL.SetDouble('GL_DPA',Arrondi(TOBL.Getdouble ('GL_MONTANTPA')/TOBL.Getdouble ('GL_QTEFACT'),V_PGI.okdecP));
            end;
          end;
          CalculeLigneHT(TOBL,TOBTaxesL,TOBPiece,DEV, NbDec,False,TOBTiers) ;
          if {(VenteAchat = 'VEN') and} (IsDetailleCollectif) then CumuleCollectifs(TOBL,TOBTaxesL,TOBVTECOLL,TOBSSTRAIT,TOBArticles,TOBTiers,TOBAffaire);
          ChangeParentLignesBases (TOBBasesL,TOBtaxesL,NBdec);
          //
          if (not IsVariante(TOBL)) and (not IsLigneFromCentralis(TOBL)) then
          begin
            SommeLignePiedHT(TOBL,TOBPiece) ;
          end;
        END else
        BEGIN
          CalculeLigneTTC(TOBL,TOBTaxesL,TOBPiece,DEV, NbDec,False,TOBTiers) ;
          if {(VenteAchat = 'VEN') and }(IsDetailleCollectif) then CumuleCollectifs(TOBL,TOBTaxesL,TOBVTECOLL,TOBSSTRAIT,TOBArticles,TOBTiers,TOBAffaire);
          ChangeParentLignesBases (TOBBasesL,TOBtaxesL,NBdec);
          //
          if (not IsVariante(TOBL)) and (not IsLigneFromCentralis(TOBL)) then
          begin
            SommeLignePiedTTC(TOBL,TOBPiece) ;
          end;
        END ;
        StockeMontantTypeSurLigne (TOBL);
        if not IsVariante(TOBL) and (not IsLigneFromCentralis(TOBL)) then
        begin
          CalcLigneMesure(TOBL) ;
          SommeLigneMesures(TOBL,TOBPiece) ;
          CumuleTypeSurPiece (TOBPiece,TOBL);
        end;
        // cumul final
        if not IsVariante(TOBL) and (not IsLigneFromCentralis(TOBL)) then
        begin
          TOBB := FindLignesbases (NumOrdre,TOBBasesL);
          if TOBB <> nil then CumuleLesBases(TOBBases,TOBB,NbDec);
        end;
      end;
      if IsVariante(TOBl) then DroplesBasesLigne (TOBL,TOBBasesL);
    end;
    if {(SaisieTypeAvanc) or }(ModifAvanc) then
    BEGIN
      if (TOBL.fieldExists('MONTANTSIT')) and (TOBPiece.fieldExists('MONTANTSIT')) then
      begin
        TOBL.PutValue('MONTANTSIT',TOBL.GetValue('BLF_MTSITUATION'));
        TOBPiece.PutValue('MONTANTSIT',TOBPiece.GetValue('MONTANTSIT')+TOBL.GetValue('MONTANTSIT'));
      end;
      CalculAvancementPiece (TobPiece,TOBL,DEV,ModifAvanc);
    END;
    TOBL.Putvalue ('GL_RECALCULER','-');
    //
  END ;

  //
  if (not SaisieTypeAvanc) then
  begin
    RecalculPiedPort (Action,TOBpiece,TOBporcs,TOBBases);
    AddLesports(TOBpiece,TOBPorcs,TOBBases,TOBBasesL,TOBTiers,DEV,Action);
    AddRetenuesHT (TOBpiece,TOBporcs,TOBBases,DEV);
  end;
  //
  if not SaisieTypeAvanc then
  begin
    EquilibrageBases(TOBpiece,TOBBases,EnHT,DEV);
  //  TOBBasesL.detail.Sort('BLB_NUMORDRE;BLB_CATEGORIETAXE;BLB_FAMILLETAXE');
  //  TOBBases.detail.Sort('GPB_CATEGORIETAXE;GPB_FAMILLETAXE');
  //FV1 : Qd on arrive des appel d'offre TOBPIECE n'est pas renseigné
  	DefiniPieceTrait (TOBBases,TOBPiece,TOBAFFaire,TOBPieceTrait,TOBSSTRAIT,DEV);
    recalculeRGloc (TOBTiers,TOBPiece,TOBPieceTrait,TOBBasesL,TOBPieceRG,TOBBasesRG,TOBporcs,TOBBases,DEV,FromGenFac);
  end else
  begin
  	DefiniPieceTrait (TOBBases,TOBPiece,TOBAFFaire,TOBPieceTrait,TOBSSTrait,DEV);
  end;
//  CalculeReglementsIntervenants(TOBSSTRAIT, TOBPiece,TOBPIECERG,TOBPiece.getValue('GP_AFFAIRE'),TOBPieceTrait,DEV);
// --
  TOBTaxesL.free;
  InitDocumentTva2014;
END ;


procedure CalculeMontantsAssocie(X: double;var XP:double;DEV:Rdevise);
begin
if DEV.code <> V_PGI.DevisePivot then
   begin
   XP:=DeviseToEuro(X,DEV.Taux,DEV.Quotite);
   end else
   begin
   XP:=X;
   end;
end;

procedure CoordinateMont (TOBG: TOB; ChampDev,ChampPiv : string; DEV: RDevise);
var X,XE : double;
begin
X:=TOBG.GetValue(ChampDev) ;
if DEV.Code<>V_PGI.DevisePivot then
   begin
   XE:=DeviseToEuro(X,DEV.Taux,DEV.Quotite) ;
   TOBG.PutValue(ChampPiv,XE) ;
   end else
   begin
   TOBG.PutValue(ChampPiv,X) ;
   end;
end;

{$IFDEF BTP}
function GetMontantTaxeLigne (TOBL: TOB; Base : double; DEV : Rdevise) : double;
var FamilleTaxe : string;
		i : Integer;
    VenteAchat : string;
    TOBTAxes : TOB;
    prefixe : string;
begin
  Prefixe := GetPrefixeTable (TOBL);
	TOBTaxes := TOB.Create ('LES TAXESL',nil,-1);
	VenteAchat:=GetInfoParPiece(TOBL.GetValue(prefixe+'_NATUREPIECEG'),'GPP_VENTEACHAT') ;
  //
  for i:=1 to 5 do
  BEGIN
    FamilleTaxe:=TOBL.GetValue(prefixe+'_FAMILLETAXE'+IntToStr(i)) ;
    if TypeTaxe (i) <> BttTVA then FamilleTaxe:='' ;
    if FamilleTaxe<>'' then
    BEGIN
      CalculeUneCategorieTaxe(VenteAchat,FamilleTaxe,i,DEV.decimale,Dev,Base,TOBL,TOBTaxes,nil,true,false) ;
    END ;
  END ;
  //
  Result := 0;
  for i := 0 to TOBTAxes.detail.count -1 do
  begin
  	Result := result + TOBTAxes.Detail[I].GetValue('BLB_VALEURDEV');
  end;
  //
  TOBTAxes.Free;
end;

procedure CalculTaxesLigneBTP ( TOBL,TOBPiece,TOBTiers,TOBTaxesL : TOB ; DEV : Rdevise; NbDec : integer ; Base : Double ; EnHT : Boolean;PourEnregCpta : boolean=false ) ;
Var i : integer ;
    AvecTaxes : boolean ;
    TOBTaxe : TOB ;
    AvecTPF : boolean;
    TotTPF,TotTVA,LaBase : Double ;
    FamilleTaxe,RegimeTaxe,NaturePiece,VenteAchat : String ;
    prefixe : string;
BEGIN
  Prefixe := GetPrefixeTable (TOBL);
  NaturePiece:=TOBL.GetValue(prefixe+'_NATUREPIECEG') ;
  if TOBTiers<>Nil then AvecTPF:=(TOBTiers.GetValue('T_SOUMISTPF')='X') else AvecTPF:=False ;
  VenteAchat:=GetInfoParPiece(NaturePiece,'GPP_VENTEACHAT') ;
  AvecTaxes:=(GetInfoParPiece(NaturePiece,'GPP_TAXE')='X') ;
  RegimeTaxe:=TOBL.GetValue(prefixe+'_REGIMETAXE') ;
  // gm les affaires sont de nature AUT , mais on doit qd m^me calculer la TVA
  if (Not AvecTaxes) or
  ((VenteAchat='AUT') and (NaturePiece <> VH_GC.AFNatAffaire) and (NaturePiece <> VH_GC.AFNatProposition) ) then Exit ;

  TotTPF:=0 ; TotTVA:=0 ;
  LaBase := Base;
  if EnHT then
  BEGIN
    {En HT, calculer les bases type TPF (sur HT) puis la TVA en dernier (sur HT+TPF)}
    if (RegimeTaxe='INT') and (VenteAchat = 'ACH') then
    begin
      // Si le fournisseur est intracommunautaire on force la tva en mode intracomnunautaire
      (*
      for I := 1 to 5 do
      begin
        FamilleTaxe:=TOBL.GetValue(prefixe+'_FAMILLETAXE'+IntToStr(i)) ;
        if TypeTaxe (i) = BttTVA then
        begin
          TOBL.SetString(prefixe+'_FAMILLETAXE'+IntToStr(i),'INT') ;
        end;
      end;
      *)
    end;
    //
    if AvecTpf then
    begin
      for i:=1 to 5 do
      BEGIN
        FamilleTaxe:=TOBL.GetValue(prefixe+'_FAMILLETAXE'+IntToStr(i)) ;
        if TypeTaxe (i) <> BttTPF then FamilleTaxe:='' ;
        if FamilleTaxe<>'' then
        BEGIN
          CalculeUneCategorieTaxe(VenteAchat,FamilleTaxe,i,NbDec,DEV,LaBase,TOBL,TOBTaxesL,TOBPiece,EnHT,AvecTPF) ;
          if i > 1 then
          BEGIN
            TOBTaxe := TOBTaxesL.Detail[i-1];
            TotTPF := TotTPF+TOBTaxe.GetValue('BLB_VALEURDEV');
          END ;
        END ;
      END ;
    end;
    LaBase := Base + TotTpf;
    //
    for i:=1 to 5 do
    BEGIN
      FamilleTaxe:=TOBL.GetValue(prefixe+'_FAMILLETAXE'+IntToStr(i)) ;
      if TypeTaxe (i) <> BttTVA then FamilleTaxe:='' ;
      if FamilleTaxe<>'' then
      BEGIN
        CalculeUneCategorieTaxe(VenteAchat,FamilleTaxe,i,NbDec,Dev,LaBase,TOBL,TOBTaxesL,TOBPiece,EnHT,AvecTPF) ;
      END ;
    END ;
  END else
  BEGIN
    {En TTC, calculer la base type TVA (sur TTC) puis les TPF en dernier (sur TTC-TVA)}
    for i:=1 to 5 do
    BEGIN
      FamilleTaxe:=TOBL.GetValue(prefixe+'_FAMILLETAXE'+IntToStr(i)) ;
      if TypeTaxe (i) <> BttTVA then FamilleTaxe:='' ;
      if FamilleTaxe<>'' then
      BEGIN
        CalculeUneCategorieTaxe(VenteAchat,FamilleTaxe,i,NbDec,DEV,LaBase,TOBL,TOBTaxesL,TOBPiece,EnHT,AvecTPF) ;
        if i > 1 then
        BEGIN
          TOBTaxe := TOBTaxesL.Detail[i-1];
          TotTVA := TotTVA+TOBTaxe.GetValue('BLB_VALEURDEV');
        END ;
      END ;
    END ;
    if AvecTPF then
    begin
      LaBase := Base - TotTva;
      //
      for i:=1 to 5 do
      BEGIN
        FamilleTaxe:=TOBL.GetValue(prefixe+'_FAMILLETAXE'+IntToStr(i)) ;
        if TypeTaxe (i) <> BttTPF then FamilleTaxe:='' ;
        if FamilleTaxe<>'' then
        BEGIN
          CalculeUneCategorieTaxe(VenteAchat,FamilleTaxe,i,NbDec,DEV,LaBase,TOBL,TOBTaxesL,TOBPiece,EnHT,AvecTPF) ;
        END ;
      END ;
    END;
  END ;
END ;
{$ENDIF}

procedure DeduitLigneModifie (TOBL,TOBpiece,TOBPieceTrait,TOBOuvrages,TOBOuvragesP,TOBBases,TOBBasesL,TOBtiers: TOB; DEV : Rdevise; Action : TactionFiche);
var EnHt : boolean;
    TOBB,TOBOL,TOBPlat,TOBtaxesL : TOB;
    NumOrdre,Indice : integer;
    prefixe : string;
    nbdec : integer;
begin
  (*
  Nbdec := Dev.Decimale;
  if (Not IsVariante(TOBL)) then
  begin
    Prefixe := GetPrefixeTable (TOBl);
		if (TOBL.GetValue(prefixe+'_MONTANTHTDEV')= 0) then Exit ; // création ou copier coller d'une ligne
		if (TOBL.GetValue(prefixe+'_TYPEREF') <> 'CAT') and (TOBL.GetValue(prefixe+'_ARTICLE') = '') then Exit ;
    NumOrdre := TOBL.getValue(prefixe+'_NUMORDRE');
    EnHt := (TOBPiece.getValue('GP_FACTUREHT')='X');
    if IsOuvrage(TOBL)and (TOBL.GetValue('GL_INDICENOMEN')<>0) then
    begin
  		TOBTaxesL := TOB.create('UNE LIGNE DE TAXE',Nil,-1);
      DroplesBasesLigne (TOBL,TOBBasesL);
      TOBPLat := FindOuvragesPlat (TOBL,TOBOuvragesP);
      if TOBplat = nil then
      begin
      	TOBPlat := AddMereLignePlat (TOBOuvragesP,TOBL.GetValue('GL_NUMORDRE'));
      end else
      begin
      	TOBPlat.clearDetail;
      end;
      GetOuvragePlat (TOBpiece,TOBL,TOBOuvrages,TOBPlat,TOBTiers,DEV);
      ZeroLigneMontant (TOBL);  // reinit des montants de la ligne
      for Indice := 0 to TOBplat.detail.count -1 do
      begin
      	//
        TOBOL := TOBPlat.detail[Indice];
        TOBTaxesL.clearDetail;
        ZeroLigneMontant (TOBOL);  // reinit des montants de la ligne
        if EnHT then
        BEGIN
          CalculeLigneHT(TOBOL,TOBTaxesL,TOBPiece,DEV, NbDec,False,TOBTiers) ;
          ChangeParentLignesBases (TOBBasesL,TOBtaxesL,NBdec);
          //
          SommeLignePlat(TOBOL,TOBL,EnHT) ;
        END else
        BEGIN
          CalculeLigneTTC(TOBOL,TOBTaxesL,TOBPiece,DEV, NbDec,False,TOBTiers) ;
          ChangeParentLignesBases (TOBBasesL,TOBtaxesL,NBdec);
          //
          SommeLignePlat(TOBOL,TOBL,EnHT) ;
        END ;
        if (not IsVariante(TOBL)) and (TOBL.GetValue('GL_TYPELIGNE')<>'CEN') then
        begin
          CalcLigneMesure(TOBOL) ;
          SommeLigneMesures(TOBOL,TOBPiece,'-') ;
        end;
        //
      end;
      TOBTaxesL.free;
    end else
    begin
    end;

    if EnHt then
    begin
      SommeLignePiedHT(TOBL,TOBPiece,'-') ;
    end else
    begin
      SommeLignePiedTTC(TOBL,TOBPiece,'-') ;
    end;
    TOBB := FindLignesbases (NumOrdre,TOBBasesL);
    if TOBB <> nil then CumuleLesBases(TOBBases,TOBB,NbDec,'-');
    DroplesBasesLigne(TOBL,TOBBasesL);
    if (TOBPiece.fieldExists('MONTANTSIT')) and (TOBL.fieldExists('MONTANTSIT')) then
    begin
      TOBPiece.PutValue('MONTANTSIT',TOBPiece.GetValue('MONTANTSIT')-TOBL.GetValue('MONTANTSIT'));
    end;
  end;
  *)
end;

procedure DeduitLigneCalc (TOBL,TOBpiece,TOBPieceTrait,TOBOuvrages,TOBOuvragesP,TOBBases,TOBBasesL,TOBtiers: TOB; DEV : Rdevise; Action : TactionFiche);
var EnHt : boolean;
    TOBB,TOBOL,TOBPlat,TOBtaxesL : TOB;
    NumOrdre,Indice : integer;
    prefixe : string;
    nbdec : integer;
begin
  (*
  Nbdec := Dev.Decimale;
  if (Not IsVariante(TOBL)) then
  begin
    Prefixe := GetPrefixeTable (TOBl);
		if (TOBL.GetValue(prefixe+'_MONTANTHTDEV')= 0) then Exit ; // création ou copier coller d'une ligne
		if (TOBL.GetValue(prefixe+'_TYPEREF') <> 'CAT') and (TOBL.GetValue(prefixe+'_ARTICLE') = '') then Exit ;
    NumOrdre := TOBL.getValue(prefixe+'_NUMORDRE');
    EnHt := (TOBPiece.getValue('GP_FACTUREHT')='X');
    if IsOuvrage(TOBL)and (TOBL.GetValue('GL_INDICENOMEN')<>0) then
    begin
  		TOBTaxesL := TOB.create('UNE LIGNE DE TAXE',Nil,-1);
      DroplesBasesLigne (TOBL,TOBBasesL);
      TOBPLat := FindOuvragesPlat (TOBL,TOBOuvragesP);
      if TOBplat = nil then
      begin
      	TOBPlat := AddMereLignePlat (TOBOuvragesP,TOBL.GetValue('GL_NUMORDRE'));
      end else
      begin
      	TOBPlat.clearDetail;
      end;
      GetOuvragePlat (TOBpiece,TOBL,TOBOuvrages,TOBPlat,TOBTiers,DEV);
      ZeroLigneMontant (TOBL);  // reinit des montants de la ligne
      for Indice := 0 to TOBplat.detail.count -1 do
      begin
      	//
        TOBOL := TOBPlat.detail[Indice];
        TOBTaxesL.clearDetail;
        ZeroLigneMontant (TOBOL);  // reinit des montants de la ligne
        if EnHT then
        BEGIN
          CalculeLigneHT(TOBOL,TOBTaxesL,TOBPiece,DEV, NbDec,False,TOBTiers) ;
          ChangeParentLignesBases (TOBBasesL,TOBtaxesL,NBdec);
          //
          SommeLignePlat(TOBOL,TOBL,EnHT) ;
        END else
        BEGIN
          CalculeLigneTTC(TOBOL,TOBTaxesL,TOBPiece,DEV, NbDec,False,TOBTiers) ;
          ChangeParentLignesBases (TOBBasesL,TOBtaxesL,NBdec);
          //
          SommeLignePlat(TOBOL,TOBL,EnHT) ;
        END ;
        if (not IsVariante(TOBL)) and (TOBL.GetValue('GL_TYPELIGNE')<>'CEN') then
        begin
          CalcLigneMesure(TOBOL) ;
          SommeLigneMesures(TOBOL,TOBPiece,'-') ;
        end;
        //
      end;
      TOBTaxesL.free;
    end else
    begin
    end;

    if EnHt then
    begin
      SommeLignePiedHT(TOBL,TOBPiece,'-') ;
    end else
    begin
      SommeLignePiedTTC(TOBL,TOBPiece,'-') ;
    end;
    TOBB := FindLignesbases (NumOrdre,TOBBasesL);
    if TOBB <> nil then CumuleLesBases(TOBBases,TOBB,NbDec,'-');
    DroplesBasesLigne(TOBL,TOBBasesL);
    if (TOBPiece.fieldExists('MONTANTSIT')) and (TOBL.fieldExists('MONTANTSIT')) then
    begin
      TOBPiece.PutValue('MONTANTSIT',TOBPiece.GetValue('MONTANTSIT')-TOBL.GetValue('MONTANTSIT'));
    end;
  end;
  *)
end;

procedure AjoutLigneCalc (TOBL,TOBpiece,TOBPieceTrait,TOBOuvrages,TOBOuvragesP,TOBBases,TOBBasesL,TOBTiers: TOB; DEV : Rdevise; Action : TactionFiche);
var EnHt : boolean;
    TOBB,TOBOL,TOBplat,TOBTaxesL : TOB;
    NumOrdre,Indice,NbDec : integer;
    prefixe : string;
begin
  (*
  NbDec := DEV.decimale;
  if (Not IsVariante(TOBL)) then
  begin
    Prefixe := GetPrefixeTable (TOBl);
		if (TOBL.GetValue(prefixe+'_TYPEREF') <> 'CAT') and (TOBL.GetValue(prefixe+'_ARTICLE') = '') then Exit ;
    NumOrdre := TOBL.getValue(prefixe+'_NUMORDRE');
    EnHt := (TOBPiece.getValue('GP_FACTUREHT')='X');
    //
    TOBTaxesL := TOB.create('UNE LIGNE DE TAXE',Nil,-1);
    if IsOuvrage(TOBL)and (TOBL.GetValue('GL_INDICENOMEN')>0) then
    begin
      DroplesBasesLigne (TOBL,TOBBasesL);
      TOBPLat := FindOuvragesPlat (TOBL,TOBOuvragesP);
      if TOBplat <> nil then
      begin
      	TOBPlat.ClearDetail;
      end else
      begin
      	TOBPlat := AddMereLignePlat (TOBOuvragesP,TOBL.GetValue('GL_NUMORDRE'));
      end;
      ZeroLigneMontant (TOBL);  // reinit des montants de la ligne
      GetOuvragePlat (TOBpiece,TOBL,TOBOuvrages,TOBPlat,TOBTiers,DEV);
      ZeroLigneMontant (TOBL);  // reinit des montants de la ligne
      for Indice := 0 to TOBplat.detail.count -1 do
      begin
      	//
        TOBOL := TOBPlat.detail[Indice];
        TOBTaxesL.clearDetail;
        ZeroLigneMontant (TOBOL);  // reinit des montants de la ligne
        if EnHT then
        BEGIN
          CalculeLigneHT(TOBOL,TOBTaxesL,TOBPiece,DEV, NbDec,False,TOBTiers) ;
          ChangeParentLignesBases (TOBBasesL,TOBtaxesL,NBdec);
          //
          SommeLignePlat(TOBOL,TOBL,EnHT) ;
        END else
        BEGIN
          CalculeLigneTTC(TOBOL,TOBTaxesL,TOBPiece,DEV, NbDec,False,TOBTiers) ;
          ChangeParentLignesBases (TOBBasesL,TOBtaxesL,NBdec);
          //
          SommeLignePlat(TOBOL,TOBL,EnHT) ;
        END ;
        if (not IsVariante(TOBL)) and (TOBL.GetValue('GL_TYPELIGNE')<>'CEN') then
        begin
          CalcLigneMesure(TOBOL) ;
          SommeLigneMesures(TOBOL,TOBPiece) ;
        end;
      end;
      // cumul final
      TOBB := FindLignesbases (NumOrdre,TOBBasesL);
      if TOBB <> nil then
      begin
        ReajusteLigneViaBase (TOBL,TOBB,EnHt,DEV);
      end;
      //
    end else
    begin
    end;
    TOBtaxesL.free;
    //
    if EnHt then
    begin
      SommeLignePiedHT(TOBL,TOBPiece,'+') ;
    end else
    begin
      SommeLignePiedTTC(TOBL,TOBPiece,'+') ;
    end;
    SommeLigneMesures(TOBL,TOBPiece,'+');
    TOBB := FindLignesbases (NumOrdre,TOBBasesL);
    if TOBB <> nil then CumuleLesBases(TOBBases,TOBB,NbDec,'+');
    if (TOBPiece.fieldExists('MONTANTSIT')) and (TOBL.fieldExists('MONTANTSIT')) then
    begin
      TOBPiece.PutValue('MONTANTSIT',TOBPiece.GetValue('MONTANTSIT')+TOBL.GetValue('MONTANTSIT'));
    end;
  end;
  *)
end;


procedure AddBaseLToBases (TOBTAxesL,TOBB : TOB; Nbdec: integer; Sens : string='+');
var delta : double;
    ValeurCalc,Taxe : double;
    Ecart : double;
begin
  if Sens = '+' then
  begin
    TOBB.putValue('GPB_BASETAXE',TOBB.GetValue('GPB_BASETAXE')+TOBTaxesL.GetValue('BLB_BASETAXE'));
    TOBB.putValue('GPB_VALEURTAXE',TOBB.GetValue('GPB_VALEURTAXE')+TOBTaxesL.GetValue('BLB_VALEURTAXE'));
    TOBB.putValue('GPB_BASEDEV',TOBB.GetValue('GPB_BASEDEV')+TOBTaxesL.GetValue('BLB_BASEDEV'));
    TOBB.putValue('GPB_VALEURDEV',TOBB.GetValue('GPB_VALEURDEV')+TOBTaxesL.GetValue('BLB_VALEURDEV'));
    TOBB.putValue('GPB_BASETAXETTC',TOBB.GetValue('GPB_BASETAXETTC')+TOBTaxesL.GetValue('BLB_BASETAXETTC'));
    TOBB.putValue('GPB_BASETTCDEV',TOBB.GetValue('GPB_BASETTCDEV')+TOBTaxesL.GetValue('BLB_BASETTCDEV'));
    TOBB.putValue('GPB_BASEACHAT',TOBB.GetValue('GPB_BASEACHAT')+TOBTaxesL.GetValue('BLB_BASEACHAT'));
    TOBB.putValue('GPB_VALEURACHAT',TOBB.GetValue('GPB_VALEURACHAT')+TOBTaxesL.GetValue('BLB_VALEURACHAT'));
  end else
  begin
    TOBB.putValue('GPB_BASETAXE',TOBB.GetValue('GPB_BASETAXE')-TOBTaxesL.GetValue('BLB_BASETAXE'));
    TOBB.putValue('GPB_VALEURTAXE',TOBB.GetValue('GPB_VALEURTAXE')-TOBTaxesL.GetValue('BLB_VALEURTAXE'));
    TOBB.putValue('GPB_BASEDEV',TOBB.GetValue('GPB_BASEDEV')-TOBTaxesL.GetValue('BLB_BASEDEV'));
    TOBB.putValue('GPB_VALEURDEV',TOBB.GetValue('GPB_VALEURDEV')-TOBTaxesL.GetValue('BLB_VALEURDEV'));
    TOBB.putValue('GPB_BASETAXETTC',TOBB.GetValue('GPB_BASETAXETTC')-TOBTaxesL.GetValue('BLB_BASETAXETTC'));
    TOBB.putValue('GPB_BASETTCDEV',TOBB.GetValue('GPB_BASETTCDEV')-TOBTaxesL.GetValue('BLB_BASETTCDEV'));
    TOBB.putValue('GPB_BASEACHAT',TOBB.GetValue('GPB_BASEACHAT')-TOBTaxesL.GetValue('BLB_BASEACHAT'));
    TOBB.putValue('GPB_VALEURACHAT',TOBB.GetValue('GPB_VALEURACHAT')-TOBTaxesL.GetValue('BLB_VALEURACHAT'));
  end;
end;

procedure ZeroMontantPorts (TOBPorcs : TOB);
var i : integer;
    TOBP : TOB;
    TypePort : string;
begin
  for i:=0 to TOBPorcs.Detail.Count-1 do
  BEGIN
    TOBP:=TOBPorcs.Detail[i] ;
    if TOBP.GetValue('GPT_FRAISREPARTIS') = 'X' then continue;
    TypePort := TOBP.GetString('GPT_TYPEPORT');
    TOBP.putValue('GPT_TOTALHTDEV',0);
    TOBP.putValue('GPT_TOTALHT',0);
    TOBP.putValue('GPT_TOTALTTCDEV',0);
    TOBP.putValue('GPT_TOTALTTC',0);
  END ;
end;

procedure DeduitlesPorts (Action : TActionFiche; TOBPiece,TOBPorcs : TOB);
var TOBP : TOB;
begin
  RecalculPiedPort (Action,TOBPiece,TOBPorcs,nil) ;
end;

Procedure PortVersLigne (TOBPiece,TOBP,TOBL : TOB ) ;
Var i : integer ;
  TypePort : string;
BEGIN

  // On ne prends pas en compte les ports&frais n'entrant pas dans le calcul
  // des totaux
  if TOBP.GetValue('GPT_FRAISREPARTIS') = 'X' then Exit;
  if TOBP.Getboolean('GPT_RETENUEDIVERSE')then Exit;
  //
  TypePort := TOBP.GetString('GPT_TYPEPORT');
  PieceVersLigne(TOBPiece,TOBL) ;
  TOBL.PutValue('GL_ARTICLE',TOBP.GetValue('GPT_CODEPORT')) ;
  TOBL.PutValue('GL_CODEARTICLE',TOBP.GetValue('GPT_CODEPORT')) ;
  TOBL.PutValue('GL_LIBELLE',TOBP.GetValue('GPT_LIBELLE')) ;
  for i:=1 to 5 do TOBL.PutValue('GL_FAMILLETAXE'+IntToStr(i),TOBP.GetValue('GPT_FAMILLETAXE'+IntToStr(i))) ;
  TOBL.PutValue('GL_QTEFACT',1) ; TOBL.PutValue('GL_QTESTOCK',1) ;
  TOBL.PutValue('GL_PUHT',TOBP.GetValue('GPT_TOTALHT')) ;
  TOBL.PutValue('GL_PUHTDEV',TOBP.GetValue('GPT_TOTALHTDEV')) ;
  TOBL.PutValue('GL_PUTTC',TOBP.GetValue('GPT_TOTALTTC')) ;
  TOBL.PutValue('GL_PUTTCDEV',TOBP.GetValue('GPT_TOTALTTCDEV')) ;
  TOBL.PutValue('GL_TYPELIGNE','GPT') ;
  TOBL.PutValue('GL_TENUESTOCK','-') ;
  TOBL.PutValue('GL_ESCOMPTABLE','-') ;
  TOBL.PutValue('GL_REMISABLELIGNE','-') ;
  TOBL.PutValue('GL_REMISABLEPIED','-') ;
  
  if TOBL.FieldExists('BLF_MTSITUATION') then
  begin
  	TOBL.PutValue('BLF_MTSITUATION',TOBP.GetValue('GPT_TOTALHTDEV')) ;
    TOBL.PutValue('BLF_QTESITUATION',1);
  end;
  if (TypePort ='HT') or (TypePort='MI') or (TypePort='MT') or (TypePort = '') then TOBL.PutValue('GL_FACTUREHT','X')
  else if (TypePort='PT') or (TypePort = 'MIC') or (TypePort ='MTC') then TOBL.PutValue('GL_FACTUREHT','-');
  // Modif BTP
  TOBL.PutValue ('GL_RECALCULER','X');
  // --
END ;

Procedure TTCTaxesLigneVersPort (TOBL,TOBP : TOB ) ;
Var i : integer ;
    sInd,TypePort : String ;
BEGIN
  if ((TOBL=Nil) or (TOBP=Nil)) then Exit ;
  TypePort := TOBP.GetValue('GPT_TYPEPORT');
  TOBP.PutValue('GPT_TOTALTTC',TOBL.GetValue('GL_PUTTC')) ;
  TOBP.PutValue('GPT_TOTALTTCDEV',TOBL.GetValue('GL_PUTTCDEV')) ;
  TOBP.PutValue('GPT_TOTALHT',TOBL.GetValue('GL_PUHT')) ;
  TOBP.PutValue('GPT_TOTALHTDEV',TOBL.GetValue('GL_PUHTDEV')) ;
  if (TypePort = 'MTC') or (TypePort = 'MT') then
  begin
    TOBP.PutValue('GPT_BASETTC',TOBL.GetValue('GL_PUTTC')) ;
    TOBP.PutValue('GPT_BASETTCDEV',TOBL.GetValue('GL_PUTTCDEV')) ;
    TOBP.PutValue('GPT_BASEHT',TOBL.GetValue('GL_PUHT')) ;
    TOBP.PutValue('GPT_BASEHTDEV',TOBL.GetValue('GL_PUHTDEV')) ;
  end;
  for i:=1 to 5 do
  BEGIN
    sInd:=IntToStr(i) ;
    TOBP.PutValue('GPT_TOTALTAXE'+sInd,TOBL.GetValue('GL_TOTALTAXE'+sInd)) ;
    TOBP.PutValue('GPT_TOTALTAXEDEV'+sInd,TOBL.GetValue('GL_TOTALTAXEDEV'+sInd)) ;
  END ;
END ;

procedure DeduitMontantPort(TOBP,TOBBases,TOBPiece : TOB; DEV : Rdevise);
var TOBBB: TOB;
    FamilleTaxe : string;
    Indice : integer;
begin
  // Deduction montant Piece
  TOBPiece.putValue('GP_TOTALHT',TOBPiece.GetValue('GP_TOTALHT')-TOBP.GetValue('GPT_TOTALHT'));
  TOBPiece.putValue('GP_TOTALHTDEV',TOBPiece.GetValue('GP_TOTALHTDEV')-TOBP.GetValue('GPT_TOTALHTDEV'));
  TOBPiece.putValue('GP_TOTALTTC',TOBPiece.GetValue('GP_TOTALTTC')-TOBP.GetValue('GPT_TOTALTTC'));
  TOBPiece.putValue('GP_TOTALTTCDEV',TOBPiece.GetValue('GP_TOTALTTCDEV')-TOBP.GetValue('GPT_TOTALTTCDEV'));
  // Deduction Base Piece
  for Indice := 1 to 5 do
  begin
    Familletaxe := TOBP.getValue('GPT_FAMILLETAXE'+IntToStr(Indice));
    TOBBB := TOBBases.findFirst(['GPB_CATEGORIETAXE','GPB_FAMILLETAXE'],['TX1',FamilleTaxe],true);
    if TOBBB <> nil then
    begin
      TOBBB.PutValue('GPB_BASETAXE',TOBBB.GetValue('GPB_BASETAXE')-TOBP.GetValue('GPT_TOTALHT'));
      TOBBB.PutValue('GPB_VALEURTAXE',TOBBB.GetValue('GPB_VALEURTAXE')-TOBP.GetValue('GPT_TOTALTAXE'+IntToStr(Indice)));
      TOBBB.PutValue('GPB_BASEDEV',TOBBB.GetValue('GPB_BASEDEV')-TOBP.GetValue('GPT_TOTALHTDEV'));
      TOBBB.PutValue('GPB_VALEURDEV',TOBBB.GetValue('GPB_VALEURDEV')-TOBP.GetValue('GPT_TOTALTAXEDEV'+IntToStr(Indice)));
    end;
  end;
end;

procedure AddlesPorts (TOBPiece,TOBPorcs,TOBBases,TOBbasesL,TOBTiers : TOB; DEV : Rdevise; Action : TactionFiche ; Sens : string='+');
Var i,Indice,NumOrdre : integer ;
    TOBP,TOBL,TOBB,TOBTaxesL : TOB ;
    EnHt : boolean;
begin
  if TOBPorcs=Nil then Exit ;
  if TOBPorcs.Detail.Count<=0 then Exit ;
  //
  TOBTaxesL := TOB.create ('TAXES DE LA LIGNE',nil,-1);
  //
  EnHt := (TOBPiece.getValue('GP_FACTUREHT')='X');
  for i:=0 to TOBPorcs.Detail.Count-1 do
  BEGIN
    TOBTaxesL.clearDetail;
    TOBP:=TOBPorcs.Detail[i] ;
    if TOBP.GetValue('GPT_FRAISREPARTIS') = 'X' then continue;
    if TOBP.GetBoolean('GPT_RETENUEDIVERSE')  then continue;
    TOBL:=NewTOBLigne(TOBPiece,-1);
    NumOrdre := TOBL.GetValue('GL_NUMORDRE');
    if Sens = '-' then
    begin
      DeduitMontantPort(TOBP,TOBBases,TOBPiece,DEV);
    end else
    begin
      PortVersLigne(TOBPiece,TOBP,TOBL) ;
      if (TOBL.getValue('GL_FACTUREHT')='X') then
      BEGIN
        CalculeLigneHT(TOBL,TOBTaxesL,TOBPiece,DEV, DEV.decimale,False,TOBTiers) ;
        ChangeParentLignesBases (TOBBasesL,TOBtaxesL,DEV.decimale);
        //
        SommeLignePiedHT(TOBL,TOBPiece) ;
        TOBB := FindLignesbases (NumOrdre,TOBBasesL);
        if TOBB <> nil then CumuleLesBases(TOBBases,TOBB,DEV.decimale);
      END else
      BEGIN
        CalculeLigneTTC(TOBL,TOBTaxesL,TOBPiece,DEV, DEV.decimale,False,TOBTiers) ;
        ChangeParentLignesBases (TOBBasesL,TOBtaxesL,DEV.decimale);
        //
        SommeLignePiedTTC(TOBL,TOBPiece) ;
        TOBB := FindLignesbases (NumOrdre,TOBBasesL);
        if TOBB <> nil then CumuleLesBases(TOBBases,TOBB,Dev.decimale);
      END ;
      TTCTaxesLigneVersPort(TOBL,TOBP) ;
      DroplesBasesLigne(TOBL,TOBBasesL); // on laisse dans les pied bases
    end;
    TOBL.Free ;
  END ;
  TOBTaxesL.free;
end;

procedure DeduitDeltaBases(TOBpiece,TOBBases: TOB;EnHt : boolean);
var Indice : integer;
    TOBB : TOB;
    delta,deltaDev : double;
begin
  for Indice := 0 to TOBBases.detail.count -1 do
  begin
    TOBB := TOBBases.detail[Indice];
    Delta := TOBB.getValue('GPB_DELTA');
    if delta <> 0 then
    begin
      TOBB.putValue('GPB_VALEURTAXE',TOBB.GetValue('GPB_VALEURTAXE')-Delta);
      if EnHt then
      begin
        TOBPiece.PutValue('GP_TOTALTTC',TOBPiece.GetValue('GP_TOTALTTC')-Delta);
      end else
      begin
        TOBPiece.PutValue('GP_TOTALHT',TOBPiece.GetValue('GP_TOTALHT')-Delta);
      end;
      TOBB.putValue('GPB_DELTA',0);
    end;
    DeltaDev := TOBB.getValue('GPB_DELTADEV');
    if deltadev <> 0 then
    begin
      TOBB.putValue('GPB_VALEURDEV',TOBB.GetValue('GPB_VALEURDEV')-Deltadev);
      if EnHt then
      begin
        TOBPiece.PutValue('GP_TOTALTTCDEV',TOBPiece.GetValue('GP_TOTALTTCDEV')-Deltadev);
      end else
      begin
        TOBPiece.PutValue('GP_TOTALHTDEV',TOBPiece.GetValue('GP_TOTALHTDEV')-Deltadev);
      end;
      TOBB.putValue('GPB_DELTADEV',0);
    end;
  end;
end;

procedure EquilibrageBases(TOBpiece,TOBBases: TOB; EnHt : Boolean; DEV : Rdevise);
var Indice : integer;
    TOBB : TOB;
    Categorie,FamilleTaxe :string;
    Taxe,Valeurcalc,delta,deltaDev,SommeTaxe,SommeTaxeDev,TaxeST : double;
begin
  SommeTaxe := 0; SommeTaxeDev := 0;
  for Indice := 0 to TOBBases.detail.count -1 do
  begin
    TOBB := TOBBases.detail[Indice];
    Taxe := (TOBB.GetValue('GPB_TAUXTAXE')/100);
    if EnHt then
    begin
      ValeurCalc := arrondi(TOBB.GetValue('GPB_BASETAXE')*Taxe,V_PGI.okdecV); // pivot
      if (VH_GC.BTCODESPECIF <> '001') or (TOBB.GetString('GPB_FOURN')='') or (TOBB.GetString('GPB_TYPEINTERV')<>'Y00') then
      begin
        Sommetaxe := SommeTaxe+ValeurCalc;
      end;
      delta := Arrondi(ValeurCalc - TOBB.GetValue('GPB_VALEURTAXE'),V_PGI.okdecV);
      if delta <> 0 then
      begin
        TOBB.putValue('GPB_VALEURTAXE',ValeurCalc);
        TOBB.putValue('GPB_DELTA',delta);
      end;
      //
      ValeurCalc := arrondi(TOBB.GetValue('GPB_BASEDEV')*Taxe,DEV.decimale); // devise
      deltadev :=Arrondi(ValeurCalc - TOBB.GetValue('GPB_VALEURDEV'),DEV.decimale);
      if (VH_GC.BTCODESPECIF <> '001') or (TOBB.GetString('GPB_FOURN')='') or (TOBB.GetString('GPB_TYPEINTERV')<>'Y00') then
      begin
        SommetaxeDEv := SommeTaxeDev+ValeurCalc;
      end;
      if deltadev <> 0 then
      begin
        TOBB.putValue('GPB_VALEURDEV',ValeurCalc);
        TOBB.putValue('GPB_DELTADEV',deltadev);
      end;
      TOBPiece.PutValue('GP_TOTALTTCDEV',TOBPiece.GetValue('GP_TOTALHTDEV')+SommeTaxeDev);
      TOBPiece.PutValue('GP_TOTALTTC',TOBPiece.GetValue('GP_TOTALHT')+SommeTaxe);
    end else
    begin
      ValeurCalc := arrondi(TOBB.GetValue('GPB_BASETAXETTC')-(TOBB.GetValue('GPB_BASETAXETTC')/(1+Taxe)),V_PGI.okdecV); // pivot
      if (VH_GC.BTCODESPECIF <> '001') or (TOBB.GetString('GPB_FOURN')='') or (TOBB.GetString('GPB_TYPEINTERV')<>'Y00') then
      begin
        Sommetaxe := SommeTaxe+ValeurCalc;
      end;
      delta :=ValeurCalc - TOBB.GetValue('GPB_VALEURTAXE');
      if delta <> 0 then
      begin
        TOBB.putValue('GPB_VALEURTAXE',ValeurCalc);
        TOBB.putValue('GPB_DELTA',delta);
      end;
      //
      ValeurCalc := arrondi(TOBB.GetValue('GPB_BASETTCDEV')- (TOBB.GetValue('GPB_BASETTCDEV')/(1+Taxe)),DEV.decimale); // devise
      deltadev :=ValeurCalc - TOBB.GetValue('GPB_VALEURDEV');
      if (VH_GC.BTCODESPECIF <> '001') or (TOBB.GetString('GPB_FOURN')='') or (TOBB.GetString('GPB_TYPEINTERV')<>'Y00') then
      begin
        SommetaxeDEv := SommeTaxeDev+ValeurCalc;
      end;
      if deltadev <> 0 then
      begin
        TOBB.putValue('GPB_VALEURDEV',ValeurCalc);
        TOBB.putValue('GPB_DELTADEV',deltadev);
      end;
    end;
    //
    if TOBB.GetValue('GPB_TYPEINTERV') = 'Y00' then
    begin
      TaxeSt := GetTauxTaxeST(TOBB.GetValue('GPB_FOURN'),GetTvaST(TOBB.GetValue('GPB_FOURN')));

    	ValeurCalc:=Arrondi(CalculeMontantTaxe(TOBB.GetValue('GPB_BASEACHAT'),TaxeST,'',nil),DEV.Decimale);
      if IsAutoLiquidationTvaST (TOBB.GetValue('GPB_FOURN')) then ValeurCalc := 0;
      TOBB.putValue('GPB_VALEURACHAT',ValeurCalc);
    end;
    //
    TOBPiece.PutValue('GP_TOTALHTDEV',TOBPiece.GetValue('GP_TOTALTTCDEV')-SommeTaxeDev);
    TOBPiece.PutValue('GP_TOTALHT',TOBPiece.GetValue('GP_TOTALTTC')-SommeTaxe);
  end;
  Indice := 0;
  if TOBBases.detail.count > 0 then
  begin
    repeat
      TOBB := TOBBases.detail[Indice];
      if (TOBB.getValue('GPB_BASEDEV')=0) and (TOBB.GetValue('GPB_BASETTCDEV')=0) then TOBB.free else Inc(Indice);
    until Indice >= TOBBases.detail.count;
  end;
end;

procedure CalculeLaLigne (TOBAffaire,TOBPiece,TOBpieceTrait,TOBSSTRAIT,TOBOUvrages,TOBOUvragesP,TOBBases,TOBBasesL,TOBTiers,TOBArticles,TOBPOrcs,TOBPieceRG,TOBBasesRG,TOBL,TOBVTECOLL : TOB;DEV : Rdevise; NbDec : integer; saisieAvanc : boolean =false; ModifAvanc : boolean=false; FF : Tform = nil);
var ValeurDpa,ValeurDpr,ValeurPmap,ValeurPmrp : double;
    Avancement : boolean;
    Typefacturation : string;
    MontantCum : double;
    LaLIgne : integer;
begin
  TypeFacturation := RenvoieTypeFact(TOBPiece.GetValue('GP_AFFAIREDEVIS'));
  Avancement := saisieAvanc or ModifAvanc;
  //
	CalculeMontantLigneAchat (TOBL,ValeurDpa,ValeurDpr,ValeurPmap,ValeurPmrp);
  //
  LaLigne := TOBL.GetValue('GL_NUMLIGNE');
  CalculFacture (TOBAffaire,TOBPiece,TOBpiecetrait,TOBSSTRAIT,TOBOUvrages,TOBOuvragesP,TOBBases,TOBBAsesL,TOBTiers,TOBArticles,TOBPOrcs,TOBPieceRG,TOBBasesRG,TOBVTECOLL,DEV,saisieAvanc,TaModif,ModifAvanc,LaLIgne);
  TOBPiece.PutValue('GP_RECALCACHAT','-');
end;


procedure CalculeMontantLigneAchat (TOBL : TOB; var ValeurDpa, ValeurDpr,Valeurpmap,ValeurPmrp : double);
var		TypeArticle : string;
      QteDuDetail : double;
begin
  if (IsArticle(TOBL)) (*and (TOBL.GEtVAlue('GL_TYPEARTICLE')<>'$$$')*) then
  begin
    TypeArticle := TOBL.GetValue('GL_TYPEARTICLE');
    QteDuDetail := TOBL.GetValue('GL_PRIXPOURQTE');

    ValeurDpa := (TOBL.Getvalue('GL_QTEFACT') * TOBL.GetValue('GL_DPA')/QteDudetail);
    ValeurDpR := (TOBL.Getvalue('GL_QTEFACT') * TOBL.GetValue('GL_DPR')/QteDudetail);
    ValeurPmap := (TOBL.Getvalue('GL_QTEFACT') * TOBL.GetValue('GL_PMAP')/QteDudetail);
    ValeurPmrp := (TOBL.Getvalue('GL_QTEFACT') * TOBL.GetValue('GL_PMRP')/QteDudetail);
     //
    ValeurDpa := Arrondi( ValeurDpa , V_PGI.okdecP);
    ValeurDpR := Arrondi( ValeurDpR , V_PGI.okdecP);
    ValeurPmap := Arrondi( ValeurPmap , V_PGI.OkdecV);
    ValeurPmrp := Arrondi( ValeurPmrp , V_PGI.OkdecV);
    //
    TOBL.Putvalue('GL_MONTANTPA',ValeurDpa); // maj montant ligne en achat
    TOBL.Putvalue('GL_MONTANTPAFG',ValeurDpa);
    TOBL.Putvalue('GL_MONTANTPAFC',ValeurDpa);
    TOBL.Putvalue('GL_MONTANTPAFR',ValeurDpa);
    TOBL.Putvalue('GL_MONTANTPR',ValeurDpr); // maj montant ligne en revient
	end;
end;

end.


