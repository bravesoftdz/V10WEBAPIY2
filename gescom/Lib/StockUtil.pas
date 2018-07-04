unit StockUtil;

interface

uses {$IFDEF VER150} variants,{$ENDIF} HEnt1, HCtrls, UTOB, Ent1, LookUp, Controls, ComCtrls, StdCtrls, ExtCtrls,
     FactContreM,
{$IFDEF EAGLCLIENT}
     Maineagl,
{$ELSE}
     DB, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} FE_Main,
{$ENDIF}
     SysUtils, Dialogs, Utiltarif, FactUtil, FactArticle, FactTOB, EntGC, UTofPrixRevient, ParamSoc,HmsgBox ;

Type T_RatioStock = (trsDocument,trsStock,trsVente) ;

Function MajQteStock(TOBL, TOBArticles, TOBOuvrage : TOB; QtePlus, QteMoins: String; Plus: boolean; RatioVA: Double; DepotEmetteur : String=''; ControleStock : Boolean=false) : Boolean;
function ChargeTOBA(TOBArticles : TOB; RefUnique : string; stDepot : string) : TOB; // DBR : Dépot unique chargé
Procedure MajPrixValo ( TOBL,TOBA : TOB ; StPrixValo : String ; Plus : boolean ; RatioVA,RatioVente : Double ) ;
Function  GetRatio ( TOBL,TOBN : TOB ; trs : T_RatioStock ) : Double ;
Procedure UpdateQteStock ( TOBDepot : TOB ; Qte,RatioVA : Double ; QtePlus,QteMoins : String ; Plus : boolean ) ;
Procedure ReaffecteDispo (TOBPiece_O,TOBPiece,TOBArticles : TOB);
Procedure ReaffecteDispoDiff(TOBArticles : TOB);
function  CalcDispo(TOBL, TobDepot: TOB): Double;
// ContreMarque
Procedure ReaffecteDispoContreM (TOBPiece_O, TOBPiece, TOBCatalogu : TOB);
Function  PrepareClauseWhereDispoContreM (TOBPiece_O,TOBPiece, TOBCatalogu : TOB) : string;
Function  MajQteStockContreM ( TOBL,TOBCatalogu : TOB ; QtePlus,QteMoins : String ; Plus : boolean ; RatioVA : Double ) : Boolean;
Procedure UpdateQteStockContreM ( TOBDCM : TOB ; Qte,RatioVA : Double ; QtePlus,QteMoins : String ; Plus : boolean ) ;
Function  ParQteToChampContreM ( LaCol : String ) : String ;
Procedure MajPrixValoContreM ( TOBL,TOBCATA : TOB ; StPrixValo : String ; Plus : boolean ; RatioVA,RatioVente : Double ) ;
// Spécifique mode
function CreerTobDispoArt(CodeArt:String):TOB ;
//$IFNDEF EAGLCLIENT    fonctionne en eAGL 06/03/2003
procedure DispoChampSupp(TOB_Dispo : TOB);
Function MakeInsertSQL(TOBA : TOB) : String ;
//$ENDIF
function  ValideArticleEtStock(TOB_Article : TOB; NaturePieceg : string): boolean;
procedure UpdateValoStock(TOBDepot,TOBL : TOB);
Function CalcRatioVA ( QualVTE,QualACH : String ) : double ;
function IsLivChantier(PiecePrecedente,Pieceorigine : string) : boolean;

Const
  // Champs traités en différentiel
  TabChampDiff : array[1..15] of string = (
    'GQ_PHYSIQUE', 'GQ_RESERVECLI', 'GQ_RESERVEFOU', 'GQ_PREPACLI',
    'GQ_LIVRECLIENT', 'GQ_LIVREFOU', 'GQ_TRANSFERT', 'GQ_QTE1', 'GQ_QTE2',
    'GQ_QTE3', 'GQ_CUMULSORTIES', 'GQ_CUMULENTREES', 'GQ_VENTEFFO',
    'GQ_ENTREESORTIES', 'GQ_ECARTINV');

implementation

Uses FactNomen
     ,FactureBtp,UtilPgi,saisutil,FactOuvrage,UentCommun
    ,CbpMCD
    ,CbpEnumerator

     ;

Function CalcRatioVA ( QualVTE,QualACH : String ) : double ;
Var TOBM : TOB ;
    XV,XA : Double ;
BEGIN
  //
  TOBM:=VH_GC.MTOBMEA.FindFirst(['GME_QUALIFMESURE','GME_MESURE'],['PIE',QualVTE],False) ;
  XV:=0 ;
  if TOBM<>Nil then XV:=TOBM.GetDouble('GME_QUOTITE') ;
  if XV=0 then XV:=1.0 ;
  //
  TOBM:=VH_GC.MTOBMEA.FindFirst(['GME_QUALIFMESURE','GME_MESURE'],['PIE',QualACH],False) ;
  XA:=0 ;
  if TOBM<>Nil then XA:=TOBM.GetDouble('GME_QUOTITE') ;
  if XA=0 then XA:=1.0 ;
  Result:=XA/XV ;
END ;

Function  GetRatio ( TOBL,TOBN : TOB ; trs : T_RatioStock ) : Double ;
Var VenteAchat : String ;
    RatioVA,FUA,FUS,FUV,CoefUAUS,COefUSUV : Double ;
BEGIN

  RatioVA := 1.0 ;

  FUA := 1.0; FUV := 1.0 ; FUS := 1.0; CoefUaUs := 0; COefUSUV := 0; Coefuaus := 0;

  VenteAchat:=GetInfoParPiece(TOBL.GetValue('GL_NATUREPIECEG'),'GPP_VENTEACHAT') ;


  Case trs of
  trsStock :
    BEGIN
      if VenteAchat='ACH' then
      BEGIN
        if TOBN<>Nil then
        begin
          Coefuaus := TOBL.getvalue('BOP_COEFCONVQTE');
          if TOBN.FieldExists('BOP_COEFCONVQTEVTE') then COefUSUV := TOBN.getDouble('BOP_COEFCONVQTEVTE');

          if CoefuaUs <> 0 then
          begin
            ratiova := 1/CoefUaUs;
          end else
          begin
            RatioVA:=CalcRatioVA(TOBN.GetValue('BOP_QUALIFQTEACH'),TOBL.GetValue('BOP_QUALIFQTESTO')) ;
          end;
        end else
        begin
          Coefuaus := Valeur(TOBL.getvalue('GL_COEFCONVQTE'));
          if TOBL.FieldExists('GL_COEFCONVQTEVTE') then COefUSUV := TOBL.GetDouble('GL_COEFCONVQTEVTE');

          if CoefuaUs <> 0 then
          begin
            ratiova := 1/CoefUaUs;
          end else
          begin
            RatioVA:=CalcRatioVA(TOBL.GetValue('GL_QUALIFQTEACH'),TOBL.GetValue('GL_QUALIFQTESTO')) ;
          end;
        end;
      END else if (VenteAchat='AUT') or (VenteAchat='TRF') then
      BEGIN
        RatioVA:=1.0 ;
      END else
      BEGIN
        if TOBN<>Nil then
        begin
          Coefuaus := TOBN.GetDouble('BOP_COEFCONVQTE');
          COefUSUV := TOBN.GetDouble('BOP_COEFCONVQTEVTE');

          if CoefUSUv <>0 then
          begin
            RatioVA:= CoefusUv;
          end else
          begin
            RatioVA:=CalcRatioVA(TOBN.GetValue('BOP_QUALIFQTEVTE'),TOBN.GetValue('BOP_QUALIFQTESTO')) ;
          end;
        end else
        begin
          Coefuaus := TOBL.GetDouble('GL_COEFCONVQTE');
          if TOBL.FieldExists('GL_COEFCONVQTEVTE') then COefUSUV := TOBL.getDouble('GL_COEFCONVQTEVTE');

          if CoefUSUv <>0 then
          begin
            RatioVA:= CoefusUv;
          end else
          begin
            RatioVA:=CalcRatioVA(TOBL.GetValue('GL_QUALIFQTEVTE'),TOBL.GetValue('GL_QUALIFQTESTO')) ;
          end;
        end;
      END ;
    END ;
  trsDocument :
    BEGIN
      if VenteAchat='ACH' then
      BEGIN
        if TOBN<>Nil then
        begin
          Coefuaus := TOBN.getDouble('BOP_COEFCONVQTE');
          if TOBN.FieldExists('BOP_COEFCONVQTEVTE') then COefUSUV := TOBN.getDouble('BOP_COEFCONVQTEVTE');
          //
          if CoefuaUs <> 0 then
          begin
            if CoefUSUV <> 0 then
            begin
              RatioVA:=CoefUSUV ;
            end else
            begin
              RatioVA:=CalcRatioVA(TOBN.GetValue('BOP_QUALIFQTEVTE'),TOBN.GetValue('BOP_QUALIFQTESTO')) ;
            end;
            RatioVA := ratioVA * CoefuaUs;
          end else
          begin
            RatioVA:=CalcRatioVA(TOBN.GetValue('BOP_QUALIFQTEVTE'),TOBN.GetValue('BOP_QUALIFQTEACH')) ;
          end;
        end else
        begin
          Coefuaus := TOBL.getDouble('GL_COEFCONVQTE');
          if TOBL.FieldExists('GL_COEFCONVQTEVTE') then COefUSUV := TOBL.getDouble('GL_COEFCONVQTEVTE');
          //
          if CoefuaUs <> 0 then
          begin
            if CoefUSUV <> 0 then
            begin
              RatioVA:=CoefUSUV ;
            end else
            begin
              RatioVA:=CalcRatioVA(TOBL.GetValue('GL_QUALIFQTEVTE'),TOBL.GetValue('GL_QUALIFQTESTO')) ;
            end;
            RatioVA := ratioVA * CoefuaUs;
          end else
          begin
            RatioVA:=CalcRatioVA(TOBL.GetValue('GL_QUALIFQTEVTE'),TOBL.GetValue('GL_QUALIFQTEACH')) ;
          end;
        end;
      END else if (VenteAchat='AUT') or (VenteAchat='TRF') then
      BEGIN
        if TOBN<>Nil then RatioVA:=CalcRatioVA(TOBN.GetValue('BOP_QUALIFQTEVTE'),TOBN.GetValue('BOP_QUALIFQTESTO'))
                     else RatioVA:=CalcRatioVA(TOBL.GetValue('GL_QUALIFQTEVTE'),TOBL.GetValue('GL_QUALIFQTESTO')) ;
      END else
      BEGIN
        RatioVA:=1.0 ;
      END ;
    END ;
  trsVente :
    BEGIN
      if VenteAchat='ACH' then
      BEGIN
        if TOBN<>Nil then
        begin
          Coefuaus := TOBN.getDouble('BOP_COEFCONVQTE');
          if TOBN.FieldExists('BOP_COEFCONVQTEVTE') then COefUSUV := TOBN.getDouble('BOP_COEFCONVQTEVTE');
          //
          if CoefUaUs <> 0 then
          begin
            if CoefUSUV <> 0 then
            begin
              RatioVA := CoefUSUv;
            end else
            begin
              RatioVA:=CalcRatioVA(TOBN.GetValue('BOP_QUALIFQTEVTE'),TOBN.GetValue('BOP_QUALIFQTESTO')) ;
            end;
            ratioVa := ratioVa * CoefUaUs;
          end else
          begin
            RatioVA:=CalcRatioVA(TOBN.GetValue('BOP_QUALIFQTEVTE'),TOBN.GetValue('BOP_QUALIFQTEACH')) ;
          end;
        end else
        begin
          Coefuaus := TOBL.getDouble('GL_COEFCONVQTE');
          if TOBL.FieldExists('GL_COEFCONVQTEVTE') then COefUSUV := TOBL.GetValue('GL_COEFCONVQTEVTE');
          //
          if CoefUaUs <> 0 then
          begin
            if CoefUSUV <> 0 then
            begin
              RatioVA := CoefUSUv;
            end else
            begin
              RatioVA:=CalcRatioVA(TOBL.GetValue('GL_QUALIFQTEVTE'),TOBL.GetValue('GL_QUALIFQTESTO')) ;
            end;
            ratioVa := ratioVa * CoefUaUs;
          end else
          begin
            RatioVA:=CalcRatioVA(TOBL.GetValue('GL_QUALIFQTEVTE'),TOBL.GetValue('GL_QUALIFQTEACH')) ;
          end;
        end;
      END else if (VenteAchat='AUT') or (VenteAchat='TRF') then
      BEGIN
        RatioVA:=1.0 ;
      END else
      BEGIN
        if TOBN<>Nil then RatioVA:=CalcRatioVA(TOBN.GetValue('BOP_QUALIFQTEVTE'),TOBN.GetValue('BOP_QUALIFQTESTO'))
                     else RatioVA:=CalcRatioVA(TOBL.GetValue('GL_QUALIFQTEVTE'),TOBL.GetValue('GL_QUALIFQTESTO')) ;
      END ;
    END ;
  END ;
  Result:=RatioVA ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Louis DECOSSE
Créé le ...... : 27/01/2000
Modifié le ... : 27/01/2000
Description .. : Calcul des prix de valorisation d'un article
Mots clefs ... : PMPA;PMAP;PMPR;PMRP;VALORISATION;
*****************************************************************}
Procedure CalculPMP ( TOBL,TOBA : TOB ; Plus : boolean ; RatioVA,RatioVente : double ) ;
Var TOBDepot{,TOBD} : TOB ;
    Depot : String ;
//    i     : integer ;
    PrixAchat,PrixRevient,Qte,Phy,PMAP,PMRP,Coef : Double ;
    PUHTNET, PUTTCNET : double ;
BEGIN
if ((TOBL=Nil) or (TOBA=Nil)) then Exit ;
if (TOBL.GetValue('GL_IDENTIFIANTWOL')=-1) and (TOBL.GetValue('GL_AFFAIRE')<>'') then Exit;  
Depot:=TOBL.GetValue('GL_DEPOT') ; Coef:=1.0 ; if Not Plus then Coef:=-1.0 ;
TOBDepot:=TOBA.FindFirst(['GQ_DEPOT'],[Depot],False) ; if TOBDepot=Nil then Exit ;
// Article
if TOBL.GetValue('GL_PRIXPOURQTE')>0 then
  begin
  PUHTNET:=TOBL.GetValue('GL_PUHTNET')/TOBL.GetValue('GL_PRIXPOURQTE');
  PUTTCNET:=TOBL.GetValue('GL_PUTTCNET')/TOBL.GetValue('GL_PRIXPOURQTE');
  end else
  begin
  PUHTNET:=TOBL.GetValue('GL_PUHTNET');
  PUTTCNET:=TOBL.GetValue('GL_PUTTCNET');
  end;

// Depot
{JLDACHATTTC}
if GetParamSoc('SO_GCACHATTTC') then PrixAchat:=Arrondi(PUTTCNET*RatioVA,V_PGI.OkDecV+2)
                                else PrixAchat:=Arrondi(PUHTNET*RatioVA,V_PGI.OkDecV+2) ;
//PrixRevient:=Arrondi(TOBA.GetValue('GA_DPR')/RatioVente,V_PGI.OkDecV+2) ;
// POURQUOI PRENAIT-ON LE GA_DPR ??????  // DBR : Dépot unique chargé
PrixRevient:=Arrondi(TOBDepot.GetValue('GQ_DPR')/RatioVente,V_PGI.OkDecV+2) ;
Qte:=Arrondi(TOBL.GetValue('GL_QTESTOCK')/RatioVA*Coef,6) ;
Phy:=TOBDepot.GetValue('GQ_PHYSIQUE') ;
{Prix Achat Pondéré}
if Phy<=0 then PMAP:=TOBDepot.GetValue('GQ_PMAP') else
   if Phy-Qte<=0 then PMAP:=PrixAchat else
      PMAP:=( (Phy-Qte)*TOBDepot.GetValue('GQ_PMAP') + Qte*PrixAchat ) / Phy ;
{Prix Revient Pondéré}
if Phy<=0 then PMRP:=TOBDepot.GetValue('GQ_PMRP') else
   if Phy-Qte<=0 then PMRP:=PrixRevient else
      PMRP:=( (Phy-Qte)*TOBDepot.GetValue('GQ_PMRP') + Qte*PrixRevient ) / Phy ;
{Mises à jour}
PMAP:=Arrondi(PMAP,V_PGI.OkDecV+2) ;
PMRP:=Arrondi(PMRP,V_PGI.OkDecV+2) ;
TOBDepot.PutValue('GQ_PMAP',PMAP) ; TOBDepot.PutValue('GQ_PMRP',PMRP) ;
// mettre a jour aussi la ligne au PMAP du document courant
TobL.PutValue ('GL_PMAP', PMAP);
END ;

Procedure MajPrixValo ( TOBL,TOBA : TOB ; StPrixValo : String ; Plus : boolean ; RatioVA,RatioVente : Double ) ;
Var TOBDepot : TOB ;
    Depot,RefUnique : String ;
    OkDPA,OkDPR,OkPMA,OkPMR,TenueStock : boolean ;
    PrixAchat,PrixRevient,PANet : Double ;
    PUHTNET, PUTTCNET : double ;
    PrixArtPCI : Double;
    DEV : RDevise;
BEGIN
if StPrixValo='' then Exit ;
if TOBL.getvalue('GL_DEVISE')<> V_PGI.DevisePivot then
begin
  DEV.Code := TOBL.GetValue('GL_DEVISE');
  GetInfosDevise(DEV);
  DEV.Taux := TOBL.GetValue('GL_TAUXDEV');
end;
if ((TOBL=Nil) or (TOBA=Nil)) then Exit ;
// Infos lignes
RefUnique:=TOBL.GetValue('GL_ARTICLE') ; if RefUnique='' then Exit ;
Depot:=TOBL.GetValue('GL_DEPOT') ;
OkDPA:=Pos('DPA;',StPrixValo)>0 ; OkDPR:=Pos('DPR;',StPrixValo)>0 ;
OkPMA:=Pos('PPA;',StPrixValo)>0 ; //OkPMR:=Pos('PPR;',StPrixValo)>0 ;
// Article
if TOBL.GetValue('GL_PRIXPOURQTE')>0 then
  begin
  PUHTNET:=TOBL.GetValue('GL_PUHTNET')/TOBL.GetValue('GL_PRIXPOURQTE');
  PUTTCNET:=TOBL.GetValue('GL_PUTTCNET')/TOBL.GetValue('GL_PRIXPOURQTE');
  end else
  begin
  PUHTNET:=TOBL.GetValue('GL_PUHTNET');
  PUTTCNET:=TOBL.GetValue('GL_PUTTCNET');
  end;
//LIGNE, +2 décimales
{JLDACHATTTC}
if GetParamSoc('SO_GCACHATTTC') then PANet:=Arrondi(PUTTCNET,V_PGI.OkDecV+2)
                                else PANet:=Arrondi(PUHTNET,V_PGI.OkDecV+2) ;
{$IFNDEF BTP}
// mais n'importe koa ca....
TOBL.PutValue('GL_DPA',PANet);
TOBL.PutValue('GL_DPR',Arrondi(CalculPrixRevient(RefUnique,TOBL.GetValue('GL_TIERS'),Depot,PANet),V_PGI.OkDecV+2)) ;
{$ENDIF}

//ARTICLE, décimales d'origine +2
PrixAchat:=Arrondi(PANet/RatioVente,V_PGI.OkDecV+2) ;
// Correction FQ  12547 --------
if TOBL.getvalue('GL_DEVISE')<> V_PGI.DevisePivot then
begin
  PrixAchat := DEVISETOPIVOTEx (prixAchat,DEV.Taux,DEV.Quotite,V_PGI.OkDecV+2);
end;
// -----------------------------
PrixRevient:=Arrondi(CalculPrixRevient(RefUnique,TOBL.GetValue('GL_TIERS'),Depot,PrixAchat),V_PGI.OkDecV+2) ;
if (Plus) then
   BEGIN
   if ((OkDPA) and (PrixAchat<>0)) then
      BEGIN
      TOBA.PutValue('GA_DPA',PrixAchat) ;
      if TOBA.FieldExists('GCA_DPA') then
         BEGIN
         TOBA.AddChampSup('DPANEW',False)    ; TOBA.PutValue('DPANEW',PANet) ;
         END ;
      END ;
   if ((OkDPR) and (PrixRevient<>0)) then
      BEGIN
      if VH_GC.GCIfDefCEGID then
         PrixArtPCI:=TOBA.GetValue('GA_DPR');
      TOBA.PutValue('GA_DPR',PrixRevient) ;
      END ;
   END ;

// Depot
{JLDACHATTTC}
if GetParamSoc('SO_GCACHATTTC') then PrixAchat:=Arrondi(PUTTCNET*RatioVA,V_PGI.OkDecV+2)
                                else PrixAchat:=Arrondi(PUHTNET*RatioVA,V_PGI.OkDecV+2) ;
// Correction FQ  12547 --------
if TOBL.getvalue('GL_DEVISE')<> V_PGI.DevisePivot then
begin
  PrixAchat := DEVISETOPIVOTEx (prixAchat,DEV.Taux,DEV.Quotite,V_PGI.OkDecV+2);
end;
// -----------------------------
PrixRevient:=Arrondi(CalculPrixRevient(RefUnique,TOBL.GetValue('GL_TIERS'),Depot,PrixAchat),V_PGI.OkDecV+2) ;
TOBDepot:=TOBA.FindFirst(['GQ_DEPOT'],[Depot],False) ;
TenueStock:=(TOBL.GetValue('GL_TENUESTOCK')='X') ;
if ((TOBDepot=Nil) and (TenueStock)) then
   BEGIN
   TOBDepot:=TOB.Create('DISPO',TOBA,-1) ;
   TOBDepot.PutValue('GQ_ARTICLE',RefUnique) ;
   TOBDepot.PutValue('GQ_DEPOT',Depot) ;
   TOBDepot.PutValue('GQ_CLOTURE','-') ;
   //$IFNDEF EAGLCLIENT fonctionne en eAGL 06/03/2003
   DispoChampSupp(TOBDepot);
   TOBDepot.AddChampSupValeur('NEW_ENREG','X');
   //$ENDIF
   END ;
if ((Plus) and (TenueStock) and (TOBDepot<>Nil)) then
   BEGIN
   if ((OkDPA) and (PrixAchat<>0)) then
      BEGIN
      TOBDepot.PutValue('GQ_DPA',PrixAchat) ;
      END ;
   if ((OkDPR) and (PrixRevient<>0)) then
      BEGIN
      TOBDepot.PutValue('GQ_DPR',PrixRevient) ;
      END ;
   END ;
if (OkPMA) then
begin
    if (TenueStock) then
    begin
        CalculPMP(TOBL,TOBA,Plus,RatioVA,RatioVente);
        if TobDepot.GetValue ('GQ_DEPOT') = VH_GC.GCDepotDefaut then
        begin
            TobA.PutValue ('GA_PMAP', TobDepot.GetValue ('GQ_PMAP'));
            TobA.PutValue ('GA_PMRP', TobDepot.GetValue ('GQ_PMRP'));
        end;
    end else // DBR : Prix et depot unique
    begin
        TobA.PutValue ('GA_PMAP', TobA.GetValue ('GA_DPA'));
        TobA.PutValue ('GA_PMRP', TobA.GetValue ('GA_DPR'));
    end;
end;

if VH_GC.GCIfDefCEGID then
   if ((Plus) and (OkDPR) and (PrixRevient<>0)) then TOBA.PutValue('GA_DPR',PrixArtPCI) ;
CalculePrixArticle(TobA, Depot); //JS FI10600
END ;

Procedure MajPrixValoContreM ( TOBL,TOBCATA : TOB ; StPrixValo : String ; Plus : boolean ; RatioVA,RatioVente : Double ) ;
Var TOBDCM : TOB ;
    Depot,RefFour,RefCata : String ;
    OkDPA,OkDPR,TenueStock : boolean ;
    PrixAchat,PrixRevient,PANet : Double ;
    PUHTNET, PUTTCNET : double ;
BEGIN
if StPrixValo='' then Exit ;
if ((TOBL=Nil) or (TOBCATA=Nil)) then Exit ;
// Infos lignes
RefCata:=TOBL.GetValue('GL_REFCATALOGUE');
RefFour:=GetCodeFourDCM(TOBL);
Depot:=TOBL.GetValue('GL_DEPOT') ;

OkDPA:=Pos('DPA;',StPrixValo)>0 ; OkDPR:=Pos('DPR;',StPrixValo)>0 ;
// Article
if TOBL.GetValue('GL_PRIXPOURQTE')>0 then
  begin
  PUHTNET:=TOBL.GetValue('GL_PUHTNET')/TOBL.GetValue('GL_PRIXPOURQTE');
  PUTTCNET:=TOBL.GetValue('GL_PUTTCNET')/TOBL.GetValue('GL_PRIXPOURQTE');
  end else
  begin
  PUHTNET:=TOBL.GetValue('GL_PUHTNET');
  PUTTCNET:=TOBL.GetValue('GL_PUTTCNET');
  end;
{JLDACHATTTC}
if GetParamSoc('SO_GCACHATTTC') then PANet:=Arrondi(PUTTCNET,V_PGI.OkDecV+2)
                                else PANet:=Arrondi(PUHTNET,V_PGI.OkDecV+2) ;
PrixAchat:=Arrondi(PANet/RatioVente,V_PGI.OkDecV+2) ;
// PrixRevient:=Arrondi(CalculPrixRevient(RefUnique,TOBL.GetValue('GL_TIERS'),Depot,PrixAchat),V_PGI.OkDecV) ;
PrixRevient:=PrixAchat;
if Plus then
   BEGIN
   if ((OkDPA) and (PrixAchat<>0)) then
      BEGIN
      TOBCATA.PutValue('GCA_DPA',PrixAchat) ;
      if TOBCATA.FieldExists('GCA_DPA') then
         BEGIN
         TOBCATA.AddChampSup('DPANEW',False)    ; TOBCATA.PutValue('DPANEW',PANet) ;
         END ;
      END ;
   END ;
// Depot
{JLDACHATTTC}
if GetParamSoc('SO_GCACHATTTC') then PrixAchat:=Arrondi(PUTTCNET*RatioVA,V_PGI.OkDecV+2)
                                else PrixAchat:=Arrondi(PUHTNET*RatioVA,V_PGI.OkDecV+2) ;
// PrixRevient:=Arrondi(CalculPrixRevient(RefUnique,TOBL.GetValue('GL_TIERS'),Depot,PrixAchat),V_PGI.OkDecV) ;
PrixRevient:=PrixAchat;
// ----- * - * - * - * - * -dd  -dz
TOBDCM:=TOBCata.FindFirst(['GQC_DEPOT'],[Depot],False) ;
if TOBDCM=Nil then
   BEGIN
   TOBDCM:=TOB.Create('DISPOCONTREM', TOBCATA,-1) ;
   TOBDCM.PutValue('GQC_REFERENCE', RefCata) ;
   TOBDCM.PutValue('GQC_FOURNISSEUR', RefFour) ;
   // A VOIR KLIENT BLANC
   TOBDCM.PutValue('GQC_CLIENT', GetCodeClientDCM(TOBL)) ;
   //
   TOBDCM.PutValue('GQC_DEPOT',Depot) ;
   END ;

if Plus and (TOBDCM<>Nil) then
   BEGIN
   if ((OkDPA) and (PrixAchat<>0)) then
      BEGIN
      TOBDCM.PutValue('GQC_DPA',PrixAchat) ;
      END ;
   END ;
END ;

Function ParQteToChamp ( LaCol : String ) : String ;
BEGIN
Result:='' ;
if LaCol='PHY' then Result:='GQ_PHYSIQUE' else
if LaCol='RC'  then Result:='GQ_RESERVECLI' else
if LaCol='PRE' then Result:='GQ_PREPACLI' else
if LaCol='RF'  then Result:='GQ_RESERVEFOU' else
if LaCol='LC'  then Result:='GQ_LIVRECLIENT' else
if LaCol='LF'  then Result:='GQ_LIVREFOU' else
if LaCol='LB1' then Result:='GQ_QTE1' else
if LaCol='LB2' then Result:='GQ_QTE2' else
if LaCol='LB3' then Result:='GQ_QTE3' else
if LaCol='TRA' then Result:='GQ_TRANSFERT' else
if LaCol='VFO' then Result:='GQ_VENTEFFO' else
if LaCol='ESE' then Result:='GQ_ENTREESORTIES' else
if LaCol='INV' then Result:='GQ_ECARTINV' ;
if LaCol='RCO' then Result:='GQ_QRUPSTOC' ; 
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Louis DECOSSE
Créé le ...... : 27/01/2000
Modifié le ... : 27/01/2000
Description .. : Mise à jour des qunatités de stock à l'enregistrement d'une pièce
Mots clefs ... : DISPO;PHYSIQUE;STOCK;QUANTITE;
*****************************************************************}
procedure UpdateValoStock(TOBDepot,TOBL : TOB);
begin
	TOBDepot.PutValue('GQ_DPA',TOBL.GetValue('GL_DPA')) ;
	TOBDepot.PutValue('GQ_PMAP',TOBL.GetValue('GL_PMAP')) ;
	TOBDepot.PutValue('GQ_DPR',TOBL.GetValue('GL_DPR')) ;
	TOBDepot.PutValue('GQ_PMRP',TOBL.GetValue('GL_PMRP')) ;
end;

Procedure UpdateQteStock ( TOBDepot : TOB ; Qte,RatioVA : Double ; QtePlus,QteMoins : String ; Plus : boolean ) ;
Var LaCol,LesCols,Champ : String ;
    Coef,X : double ;
BEGIN
Coef:=1.0 ; if Not Plus then Coef:=-1 ;
if RatioVA = 0 then RatioVA := 1; // Modif BRL 8/08/16 suite pb chez TEAM RESEAUX
LesCols:=QtePlus ;
Repeat
 LaCol:=ReadTokenSt(LesCols) ; Champ:=ParQteToChamp(LaCol) ;
 if Champ<>'' then
    BEGIN
    X:=TOBDepot.GetValue(Champ) ;
    X:=X+Coef*Qte/RatioVA ;
    TOBDepot.PutValue(Champ,X) ;
    END ;
Until ((LaCol='') or (LesCols='')) ;
LesCols:=QteMoins ;
Repeat
 LaCol:=ReadTokenSt(LesCols) ; Champ:=ParQteToChamp(LaCol) ;
 if Champ<>'' then
    BEGIN
    X:=TOBDepot.GetValue(Champ) ;
    X:=X-Coef*Qte/RatioVA ;
    if (Champ <> 'GQ_PHYSIQUE') and (X <0 ) then X := 0;
    TOBDepot.PutValue(Champ,X) ;
    END ;
Until ((LaCol='') or (LesCols='')) ;
TOBDepot.PutValue('GQ_DATEMODIF',NowH);
END ;


function IsLivChantier(Pieceprecedente,Pieceorigine : string) : boolean;
var QQ : TQuery;
    Sql  : string;
    cledoc : r_cledoc;
begin
  Result := false;
  if Pieceorigine='' then Exit; // si pas de ref origine --> Stock
  DecodeRefPiece (Pieceorigine,Cledoc);
  if (pos(cledoc.NaturePiece,'BLF;LFR;FF')>0) then
  begin
    Sql := 'SELECT GL_AFFAIRE,GL_IDENTIFIANTWOL FROM LIGNE WHERE '+
           'GL_NATUREPIECEG="'+Cledoc.NaturePiece+'" AND '+
           'GL_SOUCHE="'+Cledoc.Souche+'" AND '+
           'GL_NUMERO='+InttoStr(Cledoc.NumeroPiece)+' AND '+
           'GL_INDICEG='+InttoStr(Cledoc.Indice)+' AND '+
           'GL_NUMORDRE='+InttoStr(Cledoc.NumOrdre);
  end else if (Pos(cledoc.NaturePiece,'CBT;CC')>0) then
  begin
    Sql := 'SELECT GL_AFFAIRE,GL_IDENTIFIANTWOL FROM LIGNE WHERE GL_PIECEORIGINE="'+PieceOrigine+'" AND '+
           'GL_NATUREPIECEG IN ("BLF","LFR","FF")';
  end else if (cledoc.NaturePiece='PBT') then
  begin
    Sql := 'SELECT GL_AFFAIRE,GL_IDENTIFIANTWOL FROM LIGNE WHERE GL_PIECEORIGINE="'+PiecePrecedente+'" AND '+
           'GL_NATUREPIECEG IN ("BLF","LFR","FF")';
  end else Exit;
  QQ := OpenSql (Sql,True,1,'',True);
  if Not QQ.eof then
  begin
    if (QQ.fields[0].AsString <> '') and (QQ.fields[1].AsInteger=-1) then Result := True;
  end;
  ferme (QQ);
end;


Function MajQteStock(TOBL, TOBArticles, TOBOuvrage : TOB; QtePlus, QteMoins: String; Plus: boolean; RatioVA: Double; DepotEmetteur : String=''; ControleStock : Boolean=false) : Boolean;

  function OkGereStock (TOBL : TOB): boolean;
  var NatureP,CodeArt : string;
      PiecePrec,PieceOrigine : string;
      identifiantWol : Integer;
      Affaire : string;
  begin
    Result := false;
    //
    NatureP := TOBL.GetString('GL_NATUREPIECEG');
    PiecePrec := TOBL.GetString('GL_PIECEPRECEDENTE');
    PieceOrigine := TOBL.GetString('GL_PIECEORIGINE');
    identifiantWol := TOBL.GetInteger('GL_IDENTIFIANTWOL');
    Affaire := TOBL.GetString('GL_AFFAIRE');
    CodeArt := TOBL.GetString('GL_CODEARTICLE');
    //
    if ((Pos(NatureP ,'BLF;FF;LFR')>0) and
//       (Copy(PiecePrec,10,3) <> 'BLF') and
       ((identifiantWol<>-1) OR (Affaire='')) and
       (CodeArt<>'')) then Result := True
    else if ((Pos(NatureP ,'AFS;BFA;EEX')>0) and
       (CodeArt<>'')) then Result := True
    else if ((Pos(NatureP ,'CBT;CF;CFR')>0) and
             (CodeArt<>'') and
             ((IdentifiantWol<>-1) OR
             (Affaire=''))) then Result := True
    else if ((Pos(NatureP ,'ABC;BFC;CC')>0) and
       (CodeArt<>'')) then Result := True
    else if ((Pos(NatureP ,'LBT;FBC;BLC;')>0) and
       (not IsLivChantier(PiecePrec,Pieceorigine)) and
       (CodeArt<>'')) then Result := True
    else if ((Pos(NatureP ,'SEX;TRE;TEM;')>0) and
       (CodeArt<>'')) then Result := True;
    //
  end;

Var RefUnique,Depot : String ;
    Qte,QteN,QteDispo : Double ;
    TOBDepot,TOBPlat,TOBA,TOBPD,TOBDepotEmetteur,TOBPiece : TOB ;
    i                : integer ;
    RatN             : Double ;
    NaturePieceG,CalcRupt : string;
    ForceRupt        : boolean;
BEGIN
  Result:=True;

  // Tests préalables
  if TOBL=Nil then Exit ;
  if TOBArticles=Nil then Exit ;
  if ((QtePlus='') and (QteMoins='')) then Exit ;
  if ((TOBL.GetValue('GL_TENUESTOCK')='-') and (TOBOuvrage=Nil)) then Exit ;
  if not OkGereStock(TOBL) then Exit;

  // infos lignes
  NaturePieceG:=TOBL.GetValue('GL_NATUREPIECEG') ;

  Qte       := TOBL.GetValue('GL_QTESTOCK') ;
  if ((Qte=0) and (Naturepieceg<>'INV')) then Exit ;
  Depot     := TOBL.GetValue('GL_DEPOT') ;
  CalcRupt  := GetInfoParPiece(NaturePieceG,'GPP_CALCRUPTURE') ;

  ForceRupt:=(GetInfoParPiece(NaturePieceG,'GPP_FORCERUPTURE')='X') ;
  if TOBOuvrage=Nil then
  BEGIN
    // Chargement des infos articles
    RefUnique:=TOBL.GetValue('GL_ARTICLE') ; if RefUnique='' then Exit ;
    TOBA:=TOBArticles.FindFirst(['GA_ARTICLE'],[RefUnique],False) ;
    if TOBA=nil then TOBA:=ChargeTOBA(TOBArticles,RefUnique, Depot); // DBR : Dépot unique chargé
    // Calculs
    TOBDepot:=TOBA.FindFirst(['GQ_DEPOT'],[Depot],False) ;
    if TOBDepot=Nil then
    BEGIN
      TOBDepot:=TOB.Create('DISPO',TOBA,-1) ;
      TOBDepot.PutValue('GQ_DATEINV',StrToDate(DateToStr(V_PGI.DateEntree))) ;
      TOBDepot.PutValue('GQ_STOCKINV',0) ;
      TOBDepot.PutValue('GQ_PRIXINV',0) ;
      TOBDepot.PutValue('GQ_ARTICLE',RefUnique) ;
      TOBDepot.PutValue('GQ_DEPOT',Depot) ;
      TOBDepot.PutValue('GQ_CLOTURE','-') ;
      //$IFNDEF EAGLCLIENT fonctionne en eAGL 06/03/2003
      DispoChampSupp(TOBDepot);
      TOBDepot.AddChampSupValeur('NEW_ENREG','X');
      //$ENDIF
      if NaturePieceG='TRE' then
      begin
        TOBDepotEmetteur:=TOBA.FindFirst(['GQ_DEPOT'],[DepotEmetteur],False) ;
        if TOBDepotEmetteur<>nil then
        begin
          TOBDepot.PutValue('GQ_DPA', TOBDepotEmetteur.GetValue('GQ_DPA')) ;
          TOBDepot.PutValue('GQ_DPR', TOBDepotEmetteur.GetValue('GQ_DPR')) ;
          TOBDepot.PutValue('GQ_PMAP',TOBDepotEmetteur.GetValue('GQ_PMAP')) ;
          TOBDepot.PutValue('GQ_PMRP',TOBDepotEmetteur.GetValue('GQ_PMRP')) ;
        end;
      end;
    END ;
    QteDispo := CalcDispo(TOBL, TobDepot);
    if (ControleStock) and (CalcRupt<>'AUC') and (CalcRupt<>'') and (not ForceRupt) and (QteDispo<0) then
    BEGIN
      Result:=False;
      exit;
    END;
    UpdateQteStock(TOBDepot,Qte,RatioVA,QtePlus,QteMoins,Plus) ;
    // ----
  END else
  BEGIN
    // Modif BTP
    if TOBOuvrage.detail.count = 0 then exit;
    // --------
    TOBPiece := TOBL.Parent; if TOBPiece.NomTable <> 'PIECE' then Exit;
    if GetInfoParPiece (TOBPiece.GetString('GP_NATUREPIECEG'),'GPP_STOCKSSDETAIL')<>'X' then Exit;
    TOBPlat:=TOB.Create('',Nil,-1) ;
    //
    Qte:= TOBL.GetValue('GL_QTESTOCK');
    APlatOuvrage (TOBPiece,TOBL,TOBOuvrage,TOBPlat,Qte,1,true,true,TRUE,false);
    // Phase de controle
    for i:=0 to TOBPlat.Detail.Count-1 do
    BEGIN
      TOBPD:=TOBPlat.Detail[i] ;
      RefUnique:=TOBPD.GetValue('BOP_ARTICLE') ; if RefUnique='' then Continue ;
      if TOBPD.GetValue('BOP_TENUESTOCK') <> 'X' then Continue ;
      TOBA:=TOBArticles.FindFirst(['GA_ARTICLE'],[RefUnique],False) ; 
      RatN:=GetRatio(TOBL,TOBPD,trsStock) ; QteN:=TOBPD.GetValue('BOP_QTEFACT') ;
      // Calculs
      TOBDepot:=TOBA.FindFirst(['GQ_DEPOT'],[Depot],False) ;
      if TOBDepot=Nil then
      BEGIN
        TOBDepot:=TOB.Create('DISPO',TOBA,-1) ;
        TOBDepot.PutValue('GQ_ARTICLE',RefUnique) ;
        TOBDepot.PutValue('GQ_DEPOT',Depot) ;
        TOBDepot.PutValue('GQ_CLOTURE','-') ;
        //$IFNDEF EAGLCLIENT fonctionne en eAGL 06/03/2003
        DispoChampSupp(TOBDepot);
        TOBDepot.AddChampSupValeur('NEW_ENREG','X');
        //$ENDIF
        if NaturePieceG='TRE' then
        begin
          TOBDepotEmetteur:=TOBA.FindFirst(['GQ_DEPOT'],[DepotEmetteur],False) ;
          if TOBDepotEmetteur<>nil then
          begin
            TOBDepot.PutValue('GQ_DPA',TOBDepotEmetteur.GetValue('GQ_DPA')) ;
            TOBDepot.PutValue('GQ_DPR',TOBDepotEmetteur.GetValue('GQ_DPR')) ;
            TOBDepot.PutValue('GQ_PMAP',TOBDepotEmetteur.GetValue('GQ_PMAP')) ;
            TOBDepot.PutValue('GQ_PMRP',TOBDepotEmetteur.GetValue('GQ_PMRP')) ;
          end;
        end;
      END ;
      QteDispo := CalcDispo(TOBL, TobDepot);
      if (ControleStock) and (CalcRupt<>'AUC') and (CalcRupt<>'') and (not ForceRupt) and (QteDispo<0) then
      BEGIN
        Result:=False;  Break;
      END;
    END;

    if Result then
    begin
      // SI l'on arrive ici c'est que l'ensemble des composants de l'ouvrages peuvent etre livrés...
      for i:=0 to TOBPlat.Detail.Count-1 do
      BEGIN
        TOBPD:=TOBPlat.Detail[i] ;
        RefUnique:=TOBPD.GetValue('BOP_ARTICLE') ; if RefUnique='' then Continue ;
        if TOBPD.GetValue('BOP_TENUESTOCK') <> 'X' then Continue ;
        TOBA:=TOBArticles.FindFirst(['GA_ARTICLE'],[RefUnique],False) ; TOBA.PutValue('UTILISE','X') ;
        RatN:=GetRatio(TOBL,TOBPD,trsStock) ; QteN:=TOBPD.GetValue('BOP_QTEFACT') ;
        // Calculs
        TOBDepot:=TOBA.FindFirst(['GQ_DEPOT'],[Depot],False) ;
        if TOBDepot=Nil then
        BEGIN
          TOBDepot:=TOB.Create('DISPO',TOBA,-1) ;
          TOBDepot.PutValue('GQ_ARTICLE',RefUnique) ;
          TOBDepot.PutValue('GQ_DEPOT',Depot) ;
          TOBDepot.PutValue('GQ_CLOTURE','-') ;
          //$IFNDEF EAGLCLIENT fonctionne en eAGL 06/03/2003
          DispoChampSupp(TOBDepot);
          TOBDepot.AddChampSupValeur('NEW_ENREG','X');
          //$ENDIF
          if NaturePieceG='TRE' then
          begin
            TOBDepotEmetteur:=TOBA.FindFirst(['GQ_DEPOT'],[DepotEmetteur],False) ;
            if TOBDepotEmetteur<>nil then
            begin
              TOBDepot.PutValue('GQ_DPA',TOBDepotEmetteur.GetValue('GQ_DPA')) ;
              TOBDepot.PutValue('GQ_DPR',TOBDepotEmetteur.GetValue('GQ_DPR')) ;
              TOBDepot.PutValue('GQ_PMAP',TOBDepotEmetteur.GetValue('GQ_PMAP')) ;
              TOBDepot.PutValue('GQ_PMRP',TOBDepotEmetteur.GetValue('GQ_PMRP')) ;
            end;
          end;
        END ;
        UpdateQteStock(TOBDepot,QteN,RatN,QtePlus,QteMoins,Plus) ;
      END ;
    end;
    TOBPlat.Free ; TOBPlat:=Nil ;
  END ;
END ;

function ChargeTOBA(TOBArticles : TOB; RefUnique : string; stDepot : string) : TOB; // DBR : Dépot unique chargé
var Q : TQuery;
    SQL : String;
    TobArt : TOB;
begin
  Result:=nil;
  SQL:='Select * From ARTICLE LEFT JOIN ARTICLECOMPL ON GA2_ARTICLE=GA_ARTICLE WHERE GA_ARTICLE="'+RefUnique+'"';
  Q:=OpenSQL(SQL,True,-1,'',true) ;
  if Not Q.EOF then
  begin
    TobArt := CreerTOBArt(TOBArticles);
    TobArt.SelectDB('',Q);
    LoadTOBDispo(TobArt, True, '"' + stDepot + '"') ; // DBR : Dépot unique chargé
  end;
  Ferme(Q);
  Result:=TobArt;
end;

Function PrepareClauseWhereDispo (TOBPiece_O,TOBPiece, TOBArticles : TOB; Prefixe : string) : string;
var TOBA, TOBL, TOBDEPOT, TOBD : TOB;
    WhereArt, WhereDep, Where, RefUnique, Depot : string;
    i_ind : integer;
    TousDepot : boolean;
BEGIN
WhereArt:=''; WhereDep:=''; Where:=''; RefUnique:=''; Depot:='';
for i_ind:=0 to TOBArticles.Detail.Count - 1 do
    BEGIN
    TOBA:=TOBArticles.detail[i_ind];
    if TOBA.GetValue('GA_TENUESTOCK')='X' then
        BEGIN
        RefUnique:=TOBA.GetValue('GA_ARTICLE');
        if WhereArt='' then WhereArt:='('+Prefixe+'ARTICLE="'+RefUnique+'"'
                       else WhereArt:=WhereArt+' OR '+Prefixe+'ARTICLE="'+RefUnique+'"';
        END;
    END;
if WhereArt<>'' then
    BEGIN
    WhereArt:=WhereArt+')';
    if TOBPiece.FieldExists('REVALIDATION') then TousDepot:=(TOBPiece.GetValue('REVALIDATION')='X')
                                            else TousDepot:=False;
    if Not TousDepot then
       begin
       TOBDepot:=TOB.Create('',nil,-1);
       for i_ind:=0 to TOBPiece.Detail.Count-1 do
           BEGIN
           TOBL:=TOBPiece.detail[i_ind];
           if (TOBL.GetValue('GL_TYPELIGNE') = 'ART') and (TOBL.GetValue('GL_TYPEREF') <> 'CAT') then
              BEGIN
              Depot:=TOBL.GetValue('GL_DEPOT');
              if TOBDepot.FindFirst(['DEPOT'],[depot],False)<>Nil then continue;
              TOBD:=TOB.Create('',TOBDepot,-1);
              TOBD.AddChampSup('DEPOT',False);
              TOBD.PutValue('DEPOT', depot);
              END;
           END;
       for i_ind:=0 to TOBPiece_O.Detail.Count-1 do
           BEGIN
           TOBL:=TOBPiece_O.detail[i_ind];
           if (TOBL.GetValue('GL_TYPELIGNE')='ART') and (TOBL.GetValue('GL_TYPEREF')<>'CAT') then
              BEGIN
              Depot:=TOBL.GetValue('GL_DEPOT');
              if TOBDepot.FindFirst(['DEPOT'],[depot],False)<>Nil then continue;
              TOBD:=TOB.Create('',TOBDepot,-1);
              TOBD.AddChampSup('DEPOT',False);
              TOBD.PutValue('DEPOT', depot);
              END;
           END;
       for i_ind:=0 to TOBDepot.Detail.Count-1 do
           BEGIN
           Depot:=TOBDepot.Detail[i_ind].GetValue('DEPOT');
           if WhereDep='' then WhereDep:='('+Prefixe+'DEPOT="'+Depot+'"'
                          else WhereDep:=WhereDep+' OR '+Prefixe+'DEPOT="'+Depot+'"';
           END;
       TOBDepot.Free;
       if WhereDep<>'' then BEGIN WhereDep:=WhereDep+')'; Where:=WhereArt+' AND '+WhereDep; END;
       end else Where:=WhereArt;
    END;
Result:=Where;
END;

Procedure ReaffecteDispo (TOBPiece_O,TOBPiece, TOBArticles : TOB);
var Where,StSQL, RefUnique, Depot,NumLot : string;
    QQ : TQuery;
    TOBDispo,TOBDispoLot,TOBD,TOBA,TOBDepot,TOBDepotLot : TOB;
    i_ind,i_ind2,i_art : integer;
    Revalidation : Boolean ;
BEGIN
if TOBArticles=Nil then exit;
if TOBPiece=Nil then exit;
Where:=PrepareClauseWhereDispo(TOBPiece_O,TOBPiece, TOBArticles, 'GQ_');
if Where='' then exit; // Aucun article ou aucun depot
StSQL:='Select GQ_ARTICLE,GQ_DEPOT,GQ_PHYSIQUE,GQ_RESERVECLI,GQ_RESERVEFOU,GQ_LIVRECLIENT,GQ_LIVREFOU,GQ_PREPACLI,GQ_TRANSFERT,GQ_QTE1,GQ_QTE2,GQ_QTE3,GQ_DPA,GQ_DPR,GQ_PMAP,GQ_PMRP FROM DISPO' +
       ' WHERE (GQ_CLOTURE="-") AND '+Where;
QQ:=OpenSQL(StSQL,True,-1,'',true);
Where:=PrepareClauseWhereDispo(TOBPiece_O,TOBPiece, TOBArticles, 'GQL_');
TOBDispo:=TOB.Create('',Nil,-1);
TOBDispo.LoadDetailDB('DISPO','','',QQ,False) ;
Ferme(QQ) ;
StSQL:='Select GQL_ARTICLE,GQL_DEPOT,GQL_NUMEROLOT,GQL_PHYSIQUE FROM DISPOLOT WHERE  '+Where;
QQ:=OpenSQL(StSQL,True,-1,'',true);
TOBDispoLot:=TOB.Create('',Nil,-1);
TOBDispoLot.LoadDetailDB('DISPOLOT','','',QQ,False) ;
Ferme(QQ) ;
if TOBPiece.FieldExists('REVALIDATION') then Revalidation:=(TOBPiece.GetValue('REVALIDATION')='X')
                                        else Revalidation:=False;
if Revalidation then
   begin
   { En cas de revalidation,suppression de la TOB pour les dépôts concernés
    n'ayant aucun mouvement dans la table DISPO. }
   for i_art:=0 to TOBArticles.Detail.Count-1 do
       begin
       TOBA:=TOBArticles.Detail[i_art];
       for i_ind:=TOBA.Detail.Count-1 Downto 0 do
           begin
           TOBDepot:=TOBA.Detail[i_ind];
           RefUnique:=TOBDepot.GetValue('GQ_ARTICLE');
           Depot:=TOBDepot.GetValue('GQ_DEPOT');
           TOBD:=TOBDispo.FindFirst(['GQ_ARTICLE','GQ_DEPOT'],[RefUnique,Depot],False);
           if (TOBD=Nil)then TOBDepot.Free;
           end ;
       end;
    end;
//Réaffectation des valeurs pour les dépôts concernés.
for i_ind:=0 to TOBDispo.Detail.Count - 1 do
    BEGIN
    TOBD:=TOBDispo.Detail[i_ind];
    RefUnique:=TOBD.GetValue('GQ_ARTICLE');
    Depot:=TOBD.GetValue('GQ_DEPOT');
    TOBA:=TOBArticles.FindFirst(['GA_ARTICLE'],[RefUnique],False) ;
    TOBDepot:=TOBA.FindFirst(['GQ_DEPOT'],[Depot],False) ;
    if TOBDepot<>Nil then
        BEGIN
        TOBDepot.PutValue('GQ_PHYSIQUE',TOBD.GetValue('GQ_PHYSIQUE')) ;
        TOBDepot.PutValue('GQ_RESERVECLI',TOBD.GetValue('GQ_RESERVECLI')) ;
        TOBDepot.PutValue('GQ_RESERVEFOU',TOBD.GetValue('GQ_RESERVEFOU')) ;
        TOBDepot.PutValue('GQ_LIVRECLIENT',TOBD.GetValue('GQ_LIVRECLIENT')) ;
        TOBDepot.PutValue('GQ_LIVREFOU',TOBD.GetValue('GQ_LIVREFOU')) ;
        TOBDepot.PutValue('GQ_PREPACLI',TOBD.GetValue('GQ_PREPACLI')) ;
        TOBDepot.PutValue('GQ_TRANSFERT',TOBD.GetValue('GQ_TRANSFERT')) ;
        TOBDepot.PutValue('GQ_QTE1',TOBD.GetValue('GQ_QTE1')) ;
        TOBDepot.PutValue('GQ_QTE2',TOBD.GetValue('GQ_QTE2')) ;
        TOBDepot.PutValue('GQ_QTE3',TOBD.GetValue('GQ_QTE3')) ;
        TOBDepot.PutValue('GQ_DPA',TOBD.GetValue('GQ_DPA')) ;
        TOBDepot.PutValue('GQ_DPR',TOBD.GetValue('GQ_DPR')) ;
        TOBDepot.PutValue('GQ_PMAP',TOBD.GetValue('GQ_PMAP')) ;
        TOBDepot.PutValue('GQ_PMRP',TOBD.GetValue('GQ_PMRP')) ;
        for i_ind2 := 0 to TOBDepot.Detail.Count -1 do
            BEGIN
            NumLot:=TOBDepot.Detail[i_ind2].GetValue('GQL_NUMEROLOT');
            TOBDepotLot:=TOBDispoLot.FindFirst(['GQL_ARTICLE','GQL_DEPOT','GQL_NUMEROLOT'],[RefUnique,Depot,NumLot],False) ;
            if TOBDepotLot<>Nil then TOBDepot.Detail[i_ind2].PutValue('GQL_PHYSIQUE',TOBDepotLot.GetValue('GQL_PHYSIQUE'))
            ELSE TOBDepot.Detail[i_ind2].PutValue('GQL_PHYSIQUE',0) ;
            END;
        END ;
    END;
TOBDispo.Free; TOBDispoLot.Free;
END;

//JD 25/03/03 pour maj en différentiel
Procedure ReaffecteDispoDiff(TOBArticles : TOB);
Var iArt, iDep, iLot  : integer ;
    TOBArt, TOB_Dispo : TOB ;
begin
  for iArt := 0 to TOBArticles.Detail.Count-1 do
  begin
    TOBArt := TOBArticles.Detail[iArt];
    for iDep := 0 to TOBArt.Detail.Count -1 do
    begin
      TOB_Dispo := TOBArt.Detail[iDep];
      if (TOB_Dispo<>Nil) and (TOB_Dispo.FieldExists('Init_GQ_PHYSIQUE')) then
      begin
        With TOB_Dispo do
        begin
          PutValue('GQ_PHYSIQUE',GetValue('Init_GQ_PHYSIQUE'));
          PutValue('GQ_RESERVECLI',GetValue('Init_GQ_RESERVECLI'));
          PutValue('GQ_RESERVEFOU',GetValue('Init_GQ_RESERVEFOU'));
          PutValue('GQ_PREPACLI',GetValue('Init_GQ_PREPACLI'));
          PutValue('GQ_LIVRECLIENT',GetValue('Init_GQ_LIVRECLIENT'));
          PutValue('GQ_LIVREFOU',GetValue('Init_GQ_LIVREFOU'));
          PutValue('GQ_TRANSFERT',GetValue('Init_GQ_TRANSFERT'));
          PutValue('GQ_QTE1',GetValue('Init_GQ_QTE1'));
          PutValue('GQ_QTE2',GetValue('Init_GQ_QTE2'));
          PutValue('GQ_QTE3',GetValue('Init_GQ_QTE3'));
          PutValue('GQ_CUMULSORTIES',GetValue('Init_GQ_CUMULSORTIES'));
          PutValue('GQ_CUMULENTREES',GetValue('Init_GQ_CUMULENTREES'));
          PutValue('GQ_VENTEFFO',GetValue('Init_GQ_VENTEFFO'));
          PutValue('GQ_ENTREESORTIES',GetValue('Init_GQ_ENTREESORTIES'));
          PutValue('GQ_ECARTINV',GetValue('Init_GQ_ECARTINV'));
          if TOBArt.GetValue('GA_LOT') = 'X' then
          begin
            for iLot := 0 to Detail.Count -1 do
              Detail[iLot].PutValue('GQL_PHYSIQUE', Detail[iLot].GetValue('Init_GQL_PHYSIQUE'));
          end;
          Modifie := False;
        end;
      end;
    end;
  end;
end;

function CalcDispo(TOBL, TobDepot: TOB): Double;
var
   CalcRupt,NaturePieceG: string;
   ForceRupt: Boolean;
   ratioVA : double;
begin
 	Result := 0;
   if (TOBL = nil) or (TobDepot = nil) then EXIT;
   NaturePieceG := TOBL.GetValue('GL_NATUREPIECEG');
   RatioVA := GetRatio(TOBL, nil, trsStock);
   CalcRupt := GetInfoParPiece(NaturePieceG,'GPP_CALCRUPTURE');
	ForceRupt := (GetInfoParPiece(NaturePieceG,'GPP_FORCERUPTURE') = 'X');
   if (CalcRupt = 'DIS') and not ForceRupt then
   begin
      Result := TOBDepot.GetValue('GQ_PHYSIQUE') - TOBDepot.GetValue('GQ_RESERVECLI') - TOBDepot.GetValue('GQ_PREPACLI');
{$IFDEF BTP}
      Result := result + (TOBL.GetValue('LIVREORIGINE') * RatioVa) + GetQteResteCdeOrigine (TOBL) * RatioVA;
{$ENDIF}
   end
   else if (CalcRupt = 'PHY') and (not ForceRupt) then
   begin
      Result := TOBDepot.GetValue('GQ_PHYSIQUE');
(*
{$IFDEF BTP}
      Result := result + (TOBL.GetValue('LIVREORIGINE') * RatioVa) +
      									 GetQteResteCdeOrigine (TOBL) * RatioVA + TOBL.GetValue ('DIRECTFOU');
{$ENDIF}
*)
{$IFDEF BTP}
      Result := result + (TOBL.GetValue('LIVREORIGINE') * RatioVa);
{$ENDIF}
   end;
end;

Function PrepareClauseWhereDispoContreM (TOBPiece_O,TOBPiece, TOBCatalogu : TOB) : string;
var TOBCATA,TOBL,TOBDEPOT,TOBD : TOB;
    WhereCata, WhereDep, Where, Depot : string;
    i_ind : integer;
    RefFour,RefCata:String;

BEGIN
WhereCata:=''; WhereDep:=''; Where:=''; Depot:='';
for i_ind:=0 to TOBCatalogu.Detail.Count - 1 do
    BEGIN
    TOBCATA:=TOBCatalogu.detail[i_ind];
    if TOBCATA <> Nil then
       BEGIN
       RefCata:=TOBCATA.GetValue('GCA_REFERENCE') ;
       RefFour:=TOBCATA.GetValue('GCA_TIERS') ;
       if WhereCata <> '' then WhereCata := WhereCata + ' OR ';
       WhereCata:=WhereCata+'(GQC_REFERENCE="'+RefCata+'" AND GQC_FOURNISSEUR="'+RefFour+'")'
       END;
    END;

if WhereCata<>'' then
    BEGIN
    if TOBCatalogu.Detail.Count > 1 then WhereCata := '(' + WhereCata + ')';
    TOBDepot:=TOB.Create('',nil,-1);
    for i_ind:=0 to TOBPiece.Detail.Count-1 do
        BEGIN
        TOBL:=TOBPiece.detail[i_ind];
//        if TOBL.GetValue('GL_TYPEREF')='CAT' then
        if TOBL.GetValue('GL_ENCONTREMARQUE')='X' then  // MR
            BEGIN
            Depot:=TOBL.GetValue('GL_DEPOT');
            if TOBDepot.FindFirst(['DEPOT'],[depot],False)<>Nil then continue;
            TOBD:=TOB.Create('',TOBDepot,-1);
            TOBD.AddChampSup('DEPOT',False);
            TOBD.PutValue('DEPOT', depot);
            END;
        END;
    for i_ind:=0 to TOBPiece_O.Detail.Count-1 do
        BEGIN
        TOBL:=TOBPiece_O.detail[i_ind];
//        if TOBL.GetValue('GL_TYPEREF')='CAT' then
        if TOBL.GetValue('GL_ENCONTREMARQUE')='X' then  // MR
            BEGIN
            Depot:=TOBL.GetValue('GL_DEPOT');
            if TOBDepot.FindFirst(['DEPOT'],[depot],False)<>Nil then continue;
            TOBD:=TOB.Create('',TOBDepot,-1);
            TOBD.AddChampSup('DEPOT',False);
            TOBD.PutValue('DEPOT', depot);
            END;
        END;
    for i_ind:=0 to TOBDepot.Detail.Count-1 do
        BEGIN
        Depot:=TOBDepot.Detail[i_ind].GetValue('DEPOT');
        if WhereDep='' then WhereDep:='(GQC_DEPOT="'+Depot+'"'
                       else WhereDep:=WhereDep+' OR GQC_DEPOT="'+Depot+'"';
        END;
    TOBDepot.Free;
    if WhereDep<>'' then BEGIN WhereDep:=WhereDep+')'; Where:=WhereCata+' AND '+WhereDep; END;
    END;
Result:='(GQC_CLIENT="" OR GQC_CLIENT="' + TOBPiece.getvalue('GP_TIERS') +'") AND '+ Where;
END;

Procedure ReaffecteDispoContreM (TOBPiece_O, TOBPiece, TOBCatalogu : TOB);
var Where, StSQL, Reffour, REfCata, Depot,RefClient : string;
    QQ : TQuery;
    TOBDispoContreM,TOBDCM,TOBD,TOBCATA : TOB;
    i_ind : integer;
begin
if (TOBCatalogu=Nil) Or (TOBCatalogu.Detail.Count=0) then exit;
if (TOBPiece=Nil) Or (TOBCatalogu.Detail.Count=0) then exit;

Where:=PrepareClauseWhereDispoContreM(TOBPiece_O,TOBPiece, TOBCatalogu);
if Where='' then exit; // Aucun article ou aucun depot

StSQL:='Select GQC_REFERENCE,GQC_FOURNISSEUR,GQC_CLIENT,GQC_DEPOT,GQC_PHYSIQUE,GQC_RESERVECLI,GQC_RESERVEFOU,GQC_PREPACLI,GQC_DPA,GQC_DPR FROM DISPOCONTREM' +
       ' WHERE '+Where;
QQ:=OpenSQL(StSQL,True,-1,'',true);
TOBDispoContreM:=TOB.Create('',Nil,-1);
TOBDispoContreM.LoadDetailDB('DISPOCONTREM','','',QQ,False) ;
Ferme(QQ) ;

for i_ind:=0 to TOBDispoContreM.Detail.Count - 1 do
    BEGIN
    TOBD:=TOBDispoContreM.Detail[i_ind];
    RefFour:=TOBD.GetValue('GQC_FOURNISSEUR');
    RefCata:=TOBD.GetValue('GQC_REFERENCE');
    Depot:=TOBD.GetValue('GQC_DEPOT');
    RefClient:=TOBD.GetValue('GQC_CLIENT');
    TOBCATA:=TOBCatalogu.FindFirst(['GCA_TIERS','GCA_REFERENCE'],[RefFour,RefCata],False) ;
    TOBDCM:=TOBCATA.FindFirst(['GQC_DEPOT','GQC_CLIENT'],[Depot,RefClient],False) ;
    if TOBDCM<>Nil then
        BEGIN
        TOBDCM.PutValue('GQC_PHYSIQUE',TOBD.GetValue('GQC_PHYSIQUE')) ;
        TOBDCM.PutValue('GQC_RESERVECLI',TOBD.GetValue('GQC_RESERVECLI')) ;
        TOBDCM.PutValue('GQC_RESERVEFOU',TOBD.GetValue('GQC_RESERVEFOU')) ;
        TOBDCM.PutValue('GQC_PREPACLI',TOBD.GetValue('GQC_PREPACLI')) ;
        TOBDCM.PutValue('GQC_DPA',TOBD.GetValue('GQC_DPA')) ;
        TOBDCM.PutValue('GQC_DPR',TOBD.GetValue('GQC_DPR')) ;
        END ;
    END;
TOBDispoContreM.Free;
end;

Function MajQteStockContreM ( TOBL,TOBCatalogu : TOB ; QtePlus,QteMoins : String ; Plus : boolean ; RatioVA : Double ) : Boolean;
Var RefCata,RefFour,Depot : String ;
    QteDispo : Double ;
    TOBDCM_AvecClient, TOBDCM_SansClient, TOBCata : TOB ;
    Qte,TransPhy   : Double ;
    NaturePieceG,CalcRupt,VenteAchat : string;
    ForceRupt        : boolean;
BEGIN
Result:=True;
// Tests préalables
if TOBL=Nil then Exit ;
if TOBCatalogu=Nil then Exit ;
if ((QtePlus='') and (QteMoins='')) then Exit ;
// infos lignes
Qte:=TOBL.GetValue('GL_QTESTOCK') ; if Qte=0 then Exit ;
Depot:=TOBL.GetValue('GL_DEPOT') ;
NaturePieceG:=TOBL.GetValue('GL_NATUREPIECEG') ;
VenteAchat:=GetInfoParPiece(NaturePieceG,'GPP_VENTEACHAT') ;
if (pos('PHY',QtePlus)>0) or (pos('PHY',QteMoins)>0) then CalcRupt:='PHY' else CalcRupt:='AUC';
if VenteAchat='ACH' then CalcRupt:='AUC';
ForceRupt:=False ;
// Chargement des infos catalogue
RefCata:=TOBL.GetValue('GL_REFCATALOGUE') ;
RefFour:=GetCodeFourDCM(TOBL) ;
TOBCata:=TOBCatalogu.FindFirst(['GCA_REFERENCE','GCA_TIERS'],[RefCata,RefFour],False) ;
if TOBcata=nil then BEGIN result:=False;exit; END;
// Calculs
TOBDCM_SansClient:=TrouverTobDispoContreM(TOBL,TOBCata, False);
TOBDCM_AvecClient:=TrouverTobDispoContreM(TOBL,TOBCata, True);
if TOBDCM_AvecClient=Nil then
   begin
   ReserveCliContreM(TOBL,TOBCata, True);
   TOBDCM_AvecClient:=TrouverTobDispoContreM(TOBL,TOBCata, True);
   end;
if TOBDCM_SansClient<>nil then
   BEGIN
   if TOBDCM_SansClient.getvalue('RESERVECLI') >= TOBDCM_SansClient.getvalue('GQC_PHYSIQUE') then
      TransPhy:=TOBDCM_SansClient.getvalue('GQC_PHYSIQUE')
      ELSE
      if (VenteAchat='ACH') or (CalcRupt='PHY') then TransPhy:=TOBDCM_SansClient.getvalue('GQC_PHYSIQUE')
                                                else TransPhy:=TOBDCM_SansClient.getvalue('RESERVECLI');
   TOBDCM_SansClient.putvalue('GQC_PHYSIQUE',TOBDCM_SansClient.getvalue('GQC_PHYSIQUE')-TransPhy);
   TOBDCM_AvecClient.putvalue('GQC_PHYSIQUE',TOBDCM_AvecClient.getvalue('GQC_PHYSIQUE')+TransPhy);
   TOBDCM_AvecClient.putvalue('RESERVECLI',TOBDCM_AvecClient.getvalue('RESERVECLI')+TOBDCM_SansClient.getvalue('RESERVECLI'));
   TOBDCM_SansClient.putvalue('RESERVECLI',0);
   END;
if Qte <> TOBDCM_AvecClient.getvalue('RESERVECLI') then TOBDCM_AvecClient.putvalue('RESERVECLI',Qte);
UpdateQteStockContreM(TOBDCM_AvecClient,Qte,RatioVA,QtePlus,QteMoins,Plus);
if (CalcRupt='DIS') and not ForceRupt then
   QteDispo:=TOBDCM_AvecClient.GetValue('GQC_PHYSIQUE')-TOBDCM_AvecClient.GetValue('GQC_RESERVECLI')-TOBDCM_AvecClient.GetValue('GQC_PREPACLI')
   else if (CalcRupt='PHY') and (not ForceRupt) then QteDispo:=TOBDCM_AvecClient.GetValue('GQC_PHYSIQUE')
       else QteDispo:=0;
if (CalcRupt<>'AUC') and (CalcRupt<>'') and (not ForceRupt) and (QteDispo<0) then
   BEGIN
   Result:=False;
   END;
END ;

Procedure UpdateQteStockContreM ( TOBDCM : TOB ; Qte,RatioVA : Double ; QtePlus,QteMoins : String ; Plus : boolean ) ;
Var LaCol,LesCols,Champ : String ;
    Coef,X : double ;
BEGIN
Coef:=1.0 ; if Not Plus then Coef:=-1 ;
LesCols:=QtePlus ;
Repeat
 LaCol:=ReadTokenSt(LesCols) ; Champ:=ParQteToChampContreM(LaCol) ;
 if Champ<>'' then
    BEGIN
    X:=TOBDCM.GetValue(Champ) ; X:=X+Coef*Qte/RatioVA ;
    TOBDCM.PutValue(Champ,X) ;
    END ;
Until ((LaCol='') or (LesCols='')) ;
LesCols:=QteMoins ;
Repeat
 LaCol:=ReadTokenSt(LesCols) ; Champ:=ParQteToChampContreM(LaCol) ;
 if Champ<>'' then
    BEGIN
    X:=TOBDCM.GetValue(Champ) ; X:=X-Coef*Qte/RatioVA ;
    TOBDCM.PutValue(Champ,X) ;
    END ;
Until ((LaCol='') or (LesCols='')) ;
END ;

Function ParQteToChampContreM ( LaCol : String ) : String ;
BEGIN
Result:='' ;
if LaCol='PHY' then Result:='GQC_PHYSIQUE' else
if LaCol='RC'  then Result:='GQC_RESERVECLI' else
if LaCol='PRE' then Result:='GQC_PREPACLI' else
if LaCol='RF'  then Result:='GQC_RESERVEFOU' else
END ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Corinne Tardy
Créé le ...... : 19/09/2001
Modifié le ... : 19/09/2001
Description .. : Spécifique mode: Crée la TOB des articles avec les quantités associées
Suite ........ : Retourne une tob dispo
Mots clefs ... : DISPO;TOB
*****************************************************************}
function CreerTobDispoArt(CodeArt:String):TOB ;
var SQL: String ;
    QDispo:TQuery;
    TOBDispo:Tob ;
BEGIN
     TOBDispo:=TOB.Create('',nil,-1) ;
     TOBDispo.AddChampSup('_DISPOART',False) ;
     TOBDispo.PutValue('_DISPOART','DispoArt') ;
     SQL:='SELECT GA_CODEARTICLE,GA_ARTICLE,GQ_PHYSIQUE,GQ_DEPOT,GA_STATUTART'
     +' FROM ARTICLE left join DISPO on GA_ARTICLE = GQ_ARTICLE'
     +' WHERE GA_CODEARTICLE="'+CodeArt+'" and GQ_PHYSIQUE>0';
     QDispo:=OpenSQL(SQL,True,-1,'',true) ;
     if not QDispo.EOF then TOBDispo.LoadDetailDB('DISPO','','',QDispo,False) ;
     Ferme(QDispo) ;
     Result:=TOBDispo ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Julien DITTMAR
Créé le ...... : 05/03/2003
Modifié le ... :   /  /
Description .. : L'ajout de ses champs est indispensable pour que la
Suite ........ : validation des stocks en différentiel fonctionne
Mots clefs ... : DISPO DIFFERENTIEL
*****************************************************************}
procedure DispoChampSupp(TOB_Dispo : TOB);
var NbFilles : integer;
begin
  if (TOB_Dispo<>Nil) and (TOB_Dispo.FieldExists('GQ_PHYSIQUE')) then
  begin
    With TOB_Dispo do
    begin  //Ces champs servent à créer le delta qui mettra à jour les stocks
      AddChampSupValeur('Init_GQ_PHYSIQUE',GetValue('GQ_PHYSIQUE'));
      AddChampSupValeur('Init_GQ_RESERVECLI',GetValue('GQ_RESERVECLI'));
      AddChampSupValeur('Init_GQ_RESERVEFOU',GetValue('GQ_RESERVEFOU'));
      AddChampSupValeur('Init_GQ_PREPACLI',GetValue('GQ_PREPACLI'));
      AddChampSupValeur('Init_GQ_LIVRECLIENT',GetValue('GQ_LIVRECLIENT'));
      AddChampSupValeur('Init_GQ_LIVREFOU',GetValue('GQ_LIVREFOU'));
      AddChampSupValeur('Init_GQ_TRANSFERT',GetValue('GQ_TRANSFERT'));
      AddChampSupValeur('Init_GQ_QTE1',GetValue('GQ_QTE1'));
      AddChampSupValeur('Init_GQ_QTE2',GetValue('GQ_QTE2'));
      AddChampSupValeur('Init_GQ_QTE3',GetValue('GQ_QTE3'));
      AddChampSupValeur('Init_GQ_CUMULSORTIES',GetValue('GQ_CUMULSORTIES'));
      AddChampSupValeur('Init_GQ_CUMULENTREES',GetValue('GQ_CUMULENTREES'));
      AddChampSupValeur('Init_GQ_VENTEFFO',GetValue('GQ_VENTEFFO'));
      AddChampSupValeur('Init_GQ_ENTREESORTIES',GetValue('GQ_ENTREESORTIES'));
      AddChampSupValeur('Init_GQ_ECARTINV',GetValue('GQ_ECARTINV'));
    end;
  end else if (TOB_Dispo<>Nil) and (TOB_Dispo.FieldExists('GQL_PHYSIQUE')) then
    //Ce champ sert à créer le delta qui mettra à jour les lots
    TOB_Dispo.AddChampSupValeur('Init_GQL_PHYSIQUE',TOB_Dispo.GetValue('GQL_PHYSIQUE'));
  For NbFilles:=0 to TOB_Dispo.Detail.Count-1 do DispoChampSupp(TOB_Dispo.Detail[NbFilles]);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Julien DITTMAR
Créé le ...... : 05/03/2003
Modifié le ... :   /  /
Description .. : Retourne la valeur du champ en fonction du type de
Suite ........ : champ.
Mots clefs ... :
*****************************************************************}
Function ValeurDuChamp(TOBA : TOB; TypeChamp,NomChamp : String;
         WithInitNew : boolean=false; BInsert : boolean=false):String;
var StResult, Nom, Prefixe :string;
    Qte : Double;
    V: Variant ;
begin
  V:=TOBA.GetValue(NomChamp) ;

  { Contrôle des champs dont le suffixe est signifiant pour l'AGL }
  if WithInitNew then
    begin
    Nom:=NomChamp ;
    Prefixe:=ReadTokenPipe(Nom,'_') ;
    if (Nom='UTILISATEUR') AND (V='') then V:=V_PGI.User
    else if (Nom='DATECREATION') then V:=Date
    else if Nom='DATEMODIF' then
      BEGIN if DateWithHeure(Prefixe) then V:=NowH else V:=Date; END
    end ;

  if TypeChamp = 'BOOLEAN' then Result:='"'+V+'"'
  else if ( TypeChamp = 'INTEGER' ) OR ( TypeChamp = 'SMALLINT' ) then Result:=IntToStr (V)
  else if ( TypeChamp = 'DOUBLE' ) OR ( TypeChamp = 'RATE' ) OR (TypeChamp = 'EXTENDED') then
    begin
    if TOBA.FieldExists('Init_'+NomChamp) AND Not BInsert then
      begin
      Qte := VarAsType(V,varDouble) - TOBA.GetValue('Init_'+NomChamp);
      //StResult := FloatToStr(Qte);
      StResult := StrFPoint(Qte);
      Result := NomChamp + ' + ' + StResult ;
      end
    else Result:=StrFPoint(V);
    end
  else if TypeChamp = 'DATE' then Result:='"'+UsTime(V)+'"'
  else Result:='"'+CheckdblQuote(V)+'"' ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Julien DITTMAR
Créé le ...... : 05/03/2003
Modifié le ... :   /  /
Description .. : Construit la clause where en fonction de la clé primaire de la
Suite ........ : table
Mots clefs ... :
*****************************************************************}
function WhereCle1(TOBA : TOB; Cherche : Boolean ; LaCle1 : String) : string ;
Var Clef,Ch,Wh1,Wh2,Typ,LaCle2 : String ;
BEGIN
  With TOBA do
    begin
    Clef:=FindEtReplace(TableToCle1(NomTable),',',';',TRUE) ;
    Wh1:='' ; LaCle2:=LaCle1 ;
    While Clef<>'' do
       BEGIN
       Ch:=Uppercase(Trim(ReadTokenSt(Clef))) ;
       Typ:=ChampToType(Ch) ; Wh2:='' ;
       if (Typ='INTEGER') or (Typ='SMALLINT') then
           BEGIN
           if Cherche then
              BEGIN
              if LaCle1<>'' then Wh2:=ReadTokenSt(LaCle2)
                            else Wh2:=IntToStr(VarAsType(GetValue(Ch),VArInteger)) ;
              END else Wh2:=IntToStr(MaxInt) ;
           END else
       if (Typ='DOUBLE') or (Typ='RATE')  then
           BEGIN
           if Cherche then
              BEGIN
              if LaCle1<>'' then Wh2:=ReadTokenSt(LaCle2)
                            else Wh2:=STRFPOINT(VarAsType(GetValue(Ch),VArDouble)) ;
              END else Wh2:=IntToStr(MaxInt) ;
           END else
       if Typ='DATE'    then
           BEGIN
           if Cherche then
              BEGIN
              if LaCle1<>'' then Wh2:=ReadTokenSt(LaCle2)
                            else Wh2:='"'+USDateTime(VarAsType(GetValue(Ch),VArDate))+'"' ;
              END else Wh2:='"'+USDateTime(EncodeDate(1901,1,1))+'"' ;
           END else
       if Typ='BOOLEAN' then
          BEGIN
          if Cherche then
              BEGIN
              if LaCle1<>'' then Wh2:=ReadTokenSt(LaCle2)
                            else Wh2:='"'+VarAsType(GetValue(Ch),VArString)+'"' ;
              END else Wh2:='"D"' ;
          END else
          BEGIN
          if Cherche then
              BEGIN
              if LaCle1<>'' then Wh2:=ReadTokenSt(LaCle2)
                            else Wh2:='"'+VarAsType(GetValue(Ch),VarString)+'"' ;
              ENd else Wh2:='"'+W_W+'"' ;
          END ;
       if (Wh2<>'') or (LaCle1='')  or (Not Cherche) then
          BEGIN
          if Wh1<>'' then Wh1:=Wh1+' AND ' ;
          Wh1:=Wh1+' '+Ch+'='+Wh2 ;
          END ;
       END ;
    end;
  Result:=Wh1 ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Julien DITTMAR
Créé le ...... : 05/03/2003
Modifié le ... :   /  /
Description .. : Construit une requête Insert à partir d'une TOB
Mots clefs ... :
*****************************************************************}
Function MakeInsertSQL(TOBA : TOB) : String ;

  procedure AddS(var S:String;Const S1:String);
  begin
    S:=S+S1 ;
  end ;
var
  iCompteur: integer ;
  sChamps, sValues  : String ;
  Mcd : IMCDServiceCOM;
  Table     : ITableCOM ;
  FieldList : IEnumerator ;
  NomChamps : string;
  TypeChamps : string;
begin
  MCD := TMCD.GetMcd;
  if not mcd.loaded then mcd.WaitLoaded();
  //
  Result:='' ; sChamps := '' ; sValues:='';
  With TOBA do
  begin
    Table := mcd.getTable(TOBA.NomTable);
    if Assigned(table) then
    begin
      FieldList := Table.Fields;
      FieldList.Reset();
			While FieldList.MoveNext do
      begin
        NomChamps := (FieldList.Current as IFieldCOM).name;
        TypeChamps := (FieldList.Current as IFieldCOM).Tipe;

        if ( NomChamps = 'US_UTILISATEUR' ) then
        begin
          if sChamps='' then AddS(sChamps,NomChamps)
                        else AddS(sChamps,', '+NomChamps) ;
          if sValues='' then AddS(sValues,GetValue( NomChamps ) )
                        else AddS(sValues,', "'+GetValue( NomChamps ) + '"' ) ;
        end
        else if Not (TypeChamps='BLOB') AND Not (TypeChamps='DATA') then
        begin
          if sChamps='' then AddS(sChamps,NomChamps)
                        else AddS(sChamps,', '+NomChamps) ;
          if sValues='' then AddS(sValues,ValeurDuChamp(TOBA,TypeChamps,NomChamps,False,True))
                        else AddS(sValues,', '+ValeurDuChamp(TOBA,TypeChamps,NomChamps,True,True)) ;
        end ;
      end ;

      { Création de la requête complète }
      Result:='INSERT INTO '+NomTable+' ('+sChamps+') VALUES ('+sValues+')' ;
      Result:=FindEtReplace(Result,#0,'',True) ;
    end ;
  end;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Julien DITTMAR
Créé le ...... : 05/03/2003
Modifié le ... :   /  /
Description .. : Construit une requête Update à partir d'une TOB en ne
Suite ........ : prennant en compte que les champs modifiés
Mots clefs ... :
*****************************************************************}
Function MakeUpdateSQL(TOBA : TOB; LesChampsValo : string = '') : String ;

  function ChampsAutorise (NomChamp : string; LesChampsValo : string) : boolean;
  var LesChampsAControle,lesChampsAutorise,TheChamp,LesCodes,racine : string;
  begin
    result := true;
    lesCodes := lesChampsValo;
    racine := copy (NomChamp,1,Pos('_',NomChamp));
    lesChampsAcontrole := racine+'PAHT;'+racine+'DPA;'+Racine+'DPR;'+racine+'PMAP;'+racine+'PMRP;'+racine+'PVHT;';
    if pos(NomChamp,lesChampsAcontrole)=0 then exit;

    if LesChampsValo = '' then BEGIN result := false; exit; END; // si le champs fait bien parti des champs modifiables
    lesChampsAutorise := '';
    repeat
      TheChamp := ReadTokenSt(lesCodes);
      if TheChamp <> '' then
      begin
        if TheChamp = 'DPA' then
        begin
          if lesChampsAutorise <> '' then lesChampsAutorise := lesChampsAutorise+';';
          lesChampsAutorise := lesChampsAutorise+racine+'DPA';
        end else if TheChamp = 'PPA' then
        begin
          if lesChampsAutorise <> '' then lesChampsAutorise := lesChampsAutorise+';';
          lesChampsAutorise := lesChampsAutorise+racine+'PMAP';
        end else if TheChamp = 'DPR' then
        begin
          if lesChampsAutorise <> '' then lesChampsAutorise := lesChampsAutorise+';';
          lesChampsAutorise := lesChampsAutorise+racine+'DPR';
        end else if TheChamp = 'PPR' then
        begin
          if lesChampsAutorise <> '' then lesChampsAutorise := lesChampsAutorise+';';
          lesChampsAutorise := lesChampsAutorise+racine+'PMRP';
        end;
      end;
    until theChamp='';

    result := (Pos(NomChamp,lesChampsAutorise) > 0);
  end;

  procedure AddS(var S:String;Const S1:String);
  begin S:=S+S1 ; end ;

var
  iCompteur,nChamp: Integer ;
  TypeChamps,sChamps,NomChamps : String ;
  Mcd : IMCDServiceCOM;
  Table     : ITableCOM ;
  FieldList : IEnumerator ;
begin
  MCD := TMCD.GetMcd;
  if not mcd.loaded then mcd.WaitLoaded();

  Result:='' ; sChamps := '' ; nChamp:=0 ;
  With TOBA do
    begin
    Table := Mcd.GetTable(TOBA.NomTable);
    if Assigned(Table) then
      begin
      FieldList := Table.Fields;
      FieldList.Reset();
			While FieldList.MoveNext do
        begin
        NomChamps := (FieldList.Current as IFieldCOM).name;
        TypeChamps := (FieldList.Current as IFieldCOM).Tipe;
        if IsFieldModified(NomChamps) then
          begin
            if (ChampsAutorise(NomChamps,LesChampsValo)) then
            begin
              Inc(nChamp) ;
              if (TypeChamps <> 'BLOB') then
              begin
                if sChamps='' then sChamps:=NomChamps + '='
                  else AddS(sChamps,', '+NomChamps+'=') ;

                  AddS(sChamps,ValeurDuChamp(TOBA,TypeChamps,NomChamps)) ;
              end;
            end;
          end ;
        end ;
      // Suppression de ', ' en fin de ligne
      if nChamp>0 then
        begin
        Result:='UPDATE '+NomTable+' SET '+sChamps ;
        Result:=FindEtReplace(Result,#0,'',True) ;
        end ;
      end ;
    end;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Julien DITTMAR
Créé le ...... : 05/03/2003
Modifié le ... : 03/09/2003
Description .. : Mise à jour d'un enregistrement
Mots clefs ... :
*****************************************************************}
function UpdateMaTOB(MaTOB : TOB; LesChampsValo : string=''): Boolean;
var S,StCle : String;
    i : integer;
begin
  Result:=True;
  S:=MakeUpdateSql(MaTOB,LesChampsValo) ;

  //Ne fonctionne pas avec AGL 5.6.0m, attente de l'AGL suivant...
  //S := AglMakeUpdateTobDifferentielle(MaTOB, TabChampDiff, [], TabChampDiff);

  //Si S='' alors aucun champ à mettre à jour
  if S<>'' then
    begin
    StCle := WhereCle1(MaTOB,True,'');
    S:=S+ ' WHERE '+ StCle ;
    //Lancement de la requete UPDATE
    Try Result:=ExecuteSql(S)=1 ;
    Except Result:=False; end;
    //Si l'enreg n'existe pas, ExecuteSQL renvoi 0 donc Result=False, donc ==> Insert
    if Not Result then
      begin
      Try Result:=ExecuteSql(MakeInsertSql(MaTOB))=1 ;
      Except
      	Result:=False; end;
      end;
    end;
  MaTOB.Modifie := False ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Julien DITTMAR
Créé le ...... : 05/03/2003
Modifié le ... :   /  /
Description .. : Insertion d'un enregistrement dans la table
Mots clefs ... :
*****************************************************************}
function InsertMaTOB(MaTOB : TOB; LesChampsValo : string=''): Boolean;
begin
  Try
    Result:=ExecuteSql(MakeInsertSql(MaTOB))=1 ;
    if Result then
      begin
      MaTOB.Modifie := False ;
      MaTOB.PutValue('NEW_ENREG','-');
      end;
  Except
    Result := False;
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Julien DITTMAR
Créé le ...... : 05/03/2003
Modifié le ... :   /  /
Description .. : Validation :
Suite ........ : - des articles
Suite ........ : - des stocks en différentiel si les champs supplémentaires
Suite ........ : existent
Suite ........ : - des lots et tout ce qui est rattaché à TOB_Article
Mots clefs ... : STOCK DIFFERENTIEL
*****************************************************************}
function  ValideArticleEtStock(TOB_Article : TOB; NaturePieceg : string): boolean;
var NbArt, NbDisp, NbAutre : integer;
    TOBArt,TOBDisp,TOBDispLot,TOBAutre : TOB;
    InsertOk : Boolean;
    lesChamps : string;
    okok : boolean;
begin
  Result := True;
  lesChamps := GetInfoParPiece(NaturePieceG, 'GPP_MAJPRIXVALO');
  For NbArt:=0 to TOB_Article.Detail.Count-1 do
  begin
    if Not Result then Break;
    //Valide les articles (en général MAJ GA_DPA et GA_PMAP)
    TOBArt := TOB_Article.Detail[NbArt] ;
    if Not UpdateMaTOB(TOBArt,LesChamps) then begin Result := False; Break; end;

    For NbDisp:=0 to TOBArt.Detail.Count-1 do
    begin
      if Not Result then Break;
      //Valide les stocks
      InsertOk := False;
      TOBDisp := TOBArt.Detail[NbDisp];
      if TOBDisp.FieldExists('NEW_ENREG') then
        if TOBDisp.GetValue('NEW_ENREG')='X' then InsertOk := InsertMaTOB(TOBDisp,LesChamps);
      if Not InsertOk then
      begin
        okok := false;
        TRY
          okok := UpdateMaTOB(TOBDisp,LesChamps);
        EXCEPT
          on E: Exception do
          begin
            PgiError('Erreur SQL : ' + E.Message, 'Mise à jour Article/Stock');
      end;
        END;
        if Not okok then begin Result := False; Break; end;
      end;

      For NbAutre:=0 to TOBDisp.Detail.Count-1 do
      begin
        //Valide les lots ou tout ce qui peut-être attaché à TOBDispo.
        //if Not TOBDisp.Detail[NbAutre].InsertOrUpdateDB(False) then
        //begin Result := False; Break; end;

        InsertOk := False;
        TOBDispLot := TOBDisp.Detail[NbAutre];
        if TOBDispLot.FieldExists('NEW_ENREG') then
          if TOBDispLot.GetValue('NEW_ENREG')='X' then InsertOk := InsertMaTOB(TOBDispLot,LesChamps);
        if Not InsertOk then
        begin
          okok := false;
          TRY
            okok := UpdateMaTOB(TOBDispLot,LesChamps)
          EXCEPT
            on E: Exception do
            begin
              PgiError('Erreur SQL : ' + E.Message, 'Mise à jour Article/Stock');
        end;
          END;
          if Not okok then begin Result := False; Break; end;
      end;
    end;
  end;
  end;
end;

end.
