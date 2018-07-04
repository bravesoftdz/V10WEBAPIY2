unit FactSpec;

interface

Uses HEnt1, HCtrls, Controls, ComCtrls, StdCtrls, ExtCtrls,
     SysUtils, Classes, Graphics, Grids, Forms, Saisutil, EntGC, AGLInit, FactUtil,FactRg,ParamSoc,
{$IFDEF EAGLCLIENT}
      MaineAGL,
{$ELSE}
      DB, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} Fe_Main,
{$ENDIF}
{$IFDEF AFFAIRE}
     AffaireUtil,
{$endif}
     UtilFO, LCD_Lab, Messages, Windows,
{$IFDEF BTP}
     BTPUtil,FactOuvrage,CalcOLEGenericBTP, AppelsUtil,
{$ENDIF}
{$IFDEF GRC}
{$IFNDEF CCS3}
     UtilPerspective,
{$ENDIF}
{$ENDIF}
{$IFDEF FOS5}
     FOUtil,
{$ENDIF}
     UtilPgi,
     FactTOB,
     UTOB,
     HMsgBox,
     UtilChainage;

// Procédures et fonctions spécifiques
// Modif BTP
Function BeforeCalcul ( FF : TForm ; TOBPiece,TOBBases,TOBTiers,TOBArticles,TOBOuvrage,TOBCpta,TOBCata,TOBAffaire,TOBGCS,TOBPorcs : TOB ; DEV : RDEVISE; var Arow: integer ; CalculPv : boolean=true; RazFg : boolean=false; OnlyPv : boolean=false; Acol : Integer = -1) : boolean ;
Procedure AfterCalcul ( FF : TForm ; TOBPiece,TOBBases,TOBTiers,TOBArticles,TOBPieceRg,TOBBasesRG : TOB ; DEV : RDEVISE ) ;
// ---
Function  BeforeValide ( FF : TForm ; TOBPiece,TOBBases,TOBTiers,TOBArticles : TOB ; DEV : RDEVISE ; Old_TOTALHT : double ) : boolean ;
Procedure AfterValide ( FF : TForm ; TOBPiece,TOBBases,TOBTiers,TOBArticles : TOB ; DEV : RDEVISE ) ;
Function  BeforeReliquat ( FF : TForm ; TOBPiece,TOBBases,TOBTiers,TOBArticles : TOB ; DEV : RDEVISE ) : boolean ;
Function  BeforeGetCellCanvas ( FF : TForm ; ACol,ARow : Longint; Canvas : TCanvas ; AState: TGridDrawState) : boolean ;
Function  BeforeShow ( FF : TForm ) : boolean ;
Procedure AfterShow ( FF : TForm ) ;
Procedure AfterBloquePiece ( FF : TForm ; Action : TActionFiche ; Bloc : Boolean ) ;
Procedure BeforeGSEnter ( FF : TForm ) ;
Procedure AfterInitEnteteDefaut ( FF : TForm ) ;
Procedure AfterShowDetail ( FF : TForm; TOBPiece, TOBLig : TOB; RefUnique : String ) ;
Procedure AfterFormCreate ( FF : TForm ) ;
Procedure BeforeUpdateAncien(  TOBPiece_O,TOBPiece,TOBAcomptes_O,TobTiers : Tob);
// Spécifiques
Procedure EnvoieToucheGrid_FO ( FF : TForm; CodeArt,DimSaisie: String ; TOBPiece:TOB) ;
Procedure AfficheDetailDim ( Canvas : TCanvas ; DimSaisie : String ; ACol : integer ) ;
Procedure AfficheGenDim ( Canvas : TCanvas ; DimSaisie : String ) ;
// Modif BTP
Procedure AfterShowForLine (FF : TForm; NaturePiece : String);

procedure InitLigneValoPR (TOBL : TOB);
procedure BeforeTraitementCalculCoef (TOBArticles,TOBPiece,TOBPorcs,TOBOUvrages : TOB; DEV : Rdevise; SaisieTypeAvanc : boolean=false; NewCodeTva : string = '');
procedure BeforeApplicFrais (TOBPiece,TOBOUvrage : TOB; razFg : boolean=false);
procedure CalculeMontantsDoc (TOBPiece,TOBOuvrage : TOB ; CalculPv : boolean=true; FromImportBordereaux : boolean=false);
function IsQteZeroAutorise (TOBPiece : TOB) : boolean;
// --
implementation

Uses
   Facture
   ,AGLInitGC
   ,FactTiers
   ,FactComm
{$IFDEF GPAO}
   ,wInitChamp
   ,wCommuns
   ,hMsgBox
{$ENDIF}
   ,PiecesRecalculs
   ,Ucotraitance
   ;

// Modif BTP
{$IFDEF BTP}
function IsQteZeroAutorise (TOBPiece : TOB) : boolean;
begin
	result := false; // défaut
  if not TOBPIece.FieldExists ('AFF_OKSIZERO') then exit;
  if (TOBPiece.GEtVAlue('AFF_OKSIZERO')='X') or (GetParamSocSecur('SO_BTGESTVIDESUREDIT',false)) Then result := true;
end;
{$ENDIF}



procedure BeforeApplicFrais (TOBPiece,TOBOUvrage : TOB; razFg : boolean=false);
var Indice : Integer;
		TOBL : TOB;
    TypeLigne,TypeArticle : string;
    IndicePaFC,IndicePAFR,IndicePAFG,IndiceMT,IndiceFG,IndiceFC,IndicePR : integer;
    IndiceMtLigne,IndiceMtPAFC,IndiceMtPAFG,IndiceMTPAFR,IndiceMtFR,IndiceMtFC : integer;
    IndiceMtPR,IndicePuLig,IndiceQteLig,IndiceNomen : integer;
    IndiceTpsLig,IndiceTPS,IndiceMtFG,IndiceMtFGLig,IndiceCoefFGLIg,IndPrixPourQte,IndiceCFC,IndiceCFR : integer;
begin

	IndiceMT := TOBPiece.GetNumChamp ('GP_MONTANTPA');
	IndiceMTFG := TOBPiece.GetNumChamp ('GP_MONTANTFG');
  IndicePAFG := TOBPiece.GetNumChamp ('GP_MONTANTPAFG');
  IndicePAFC := TOBPiece.GetNumChamp ('GP_MONTANTPAFC');
  IndicePAFR := TOBPiece.GetNumChamp ('GP_MONTANTPAFR');
  IndiceTPS := TOBPiece.GetNumChamp ('GP_TPSUNITAIRE');

  // Modif BTP

  TOBPIece.putValue('GP_MONTANTPAFG',0);
  TOBPIece.putValue('GP_MONTANTPAFC',0);
  TOBPIece.putValue('GP_MONTANTPAFR',0);
  TOBPIece.putValue('GP_MONTANTPA',0);
  TOBPIece.putValue('GP_MONTANTFG',0);
  TOBPIece.putValue('GP_MONTANTFC',0);
  TOBPIece.putValue('GP_MONTANTFR',0);
  TOBPIece.putValue('GP_MONTANTPR',0);
  TOBPIece.putValue('GP_TPSUNITAIRE',0);

  // Recalcul des montant PA applicable sur Fg
	for Indice := 0 to TOBPiece.detail.count -1 do
  begin
  	TOBL := TOBPiece.detail[Indice];
    if TOBL.GetString('GL_TYPEARTICLE')='POU' then continue;

    if IsSousDetail (TOBL) then continue;
    TypeLigne := TOBL.GetValue('GL_TYPELIGNE');
    TypeArticle := TOBL.GetValue('GL_TYPEARTICLE');
    IndiceNomen := TOBL.GetValue('GL_INDICENOMEN');

  	if Indice = 0 then
    begin
    	IndiceMtLigne := TOBL.GetNumChamp ('GL_MONTANTPA');
      IndiceMtPAFG  := TOBL.GetNumChamp ('GL_MONTANTPAFG');
      IndiceMtPAFC  := TOBL.GetNumChamp ('GL_MONTANTPAFC');
      IndiceMtPAFR  := TOBL.GetNumChamp ('GL_MONTANTPAFR');
      IndiceMtFR    := TOBL.GetNumChamp ('GL_MONTANTFR');
      IndiceMtFC    := TOBL.GetNumChamp ('GL_MONTANTFC');
      IndiceMtFGLig := TOBL.GetNumChamp ('GL_MONTANTFG');
      IndiceMtPR    := TOBL.GetNumChamp ('GL_MONTANTPR');
      //
      IndiceCFC  := TOBL.GetNumChamp ('GL_COEFFC');
      IndiceCFR  := TOBL.GetNumChamp ('GL_COEFFR');
      //
      IndicePuLIg   := TOBL.GetNumChamp ('GL_DPA');
      IndiceQteLIg  := TOBL.GetNumChamp ('GL_QTEFACT');
      IndiceCoefFGLIg  := TOBL.GetNumChamp ('GL_COEFFG');
      IndiceTpsLig  := TOBL.GetNumChamp ('GL_TPSUNITAIRE');
      IndPrixPourQte := TOBL.GetNumChamp ('GL_PRIXPOURQTE');
    end;

    if (not IsArticle (TOBL)) and (TypeLigne<>'CEN') Then
    begin
      TOBL.putValue('GL_DPA',0); TOBL.putValue('GL_DPR',0); TOBL.putValue('GL_PUHTDEV',0);TOBL.putValue('GL_PUHT',0);
    end;
    if (TypeLigne <> 'ART') {or (TOBL.GEtValue('GL_ARTICLE')='')}  then continue;
    //
    TOBL.putValeur(IndiceMtPAFG,0);
    TOBL.putValeur(IndiceMtPAFC,0);
    TOBL.putValeur(IndiceMtPAFR,0);
//    TOBL.PutValeur(IndiceMtPR,0);
    //
    if ((TypeArticle = 'OUV') or (TypeArticle ='ARP')) and (IndiceNomen > 0) then
    begin
    	BeforeApplicFraisDetailOuv (TOBL,TOBOUVRAGE,razFg);
    end else
    begin
      // recalcul des montant achat
      if TOBL.GetValeur (IndPrixPourQte)=0 then TOBL.putValeur (IndPrixPourQte,1);
      TOBL.PutValeur(IndiceMtLigne, TOBL.getValeur(IndicePuLIg)*TOBL.GetValeur (IndiceQteLig)/TOBL.GetValeur (IndPrixPourQte));    // MONTANTPA
      TOBL.PutValeur(IndiceMtFGLig, TOBL.getValeur(IndiceMtLigne)*TOBL.GetValeur (IndiceCoefFGLIg));    // MONTANTFG

      if TOBL.GetVAlue('GLC_NONAPPLICFG')<>'X' then
      begin
        TOBL.putValeur(IndiceMtPAFG,TOBL.GetValeur(IndiceMtLigne));
      end;

      if TOBL.GetVAlue('GLC_NONAPPLICFC')<>'X' then
      begin
        TOBL.putValeur(IndiceMtPAFC,TOBL.GetValeur(IndiceMtLigne));
      end else
      begin
        TOBL.putValeur(IndiceMtFC,0);
        TOBL.putValeur(IndiceCFC,0);
      end;

      if TOBL.GetVAlue('GLC_NONAPPLICFRAIS')<>'X' then
      begin
        TOBL.putValeur(IndiceMtPAFR,TOBL.GetValeur(IndiceMtLigne));
      end else
      begin
        TOBL.putValeur(IndiceMtFR,0);
        TOBL.putValeur(IndiceCFR,0);
      end;

			if (Pos (TOBL.getValue('BNP_TYPERESSOURCE'),VH_GC.BTTypeMoInterne) > 0) and (not IsLigneExternalise(TOBL) ) then
      begin
      	TOBL.PutValeur(IndiceTpsLig,1);
      end else
      begin
      	TOBL.PutValeur(IndiceTpsLig,0);
      end;
    end;
    //

		TOBPiece.PutValeur (IndiceMT,TOBPiece.getValeur(IndiceMT)+TOBL.GetValeur (IndiceMtLigne));
		TOBPiece.PutValeur (IndiceMTFG,TOBPiece.getValeur(IndiceMTFG)+TOBL.GetValeur (IndiceMtFGLig));
    TOBPiece.PutValeur (IndicePAFG,TOBPiece.getValeur(IndicePAFG)+TOBL.GetValeur (IndiceMtPAFG));
    TOBPiece.PutValeur (IndicePAFC,TOBPiece.getValeur(IndicePAFC)+TOBL.GetValeur (IndiceMtPAFC));
    TOBPiece.PutValeur (IndicePAFR,TOBPiece.getValeur(IndicePAFR)+TOBL.GetValeur (IndiceMtPAFR));

    TOBPiece.PutValeur (IndiceTps,TOBPiece.getValeur(IndiceTps)+TOBL.GetValeur (IndiceTpsLig));
  end;
end;



procedure CalculeMontantsDoc (TOBPiece,TOBOuvrage : TOB ; CalculPv : boolean=true; FromImportBordereaux : boolean=false);
var Indice : integer;
		TypeArticle : string;
		TOBL : TOB;
    IndMontantPR,IndMontantPA,IndPr,IndligPr,IndLigPa,IndQte,IndPA,indLigFG,IndLigFC,IndLigFR,IndLigPAFG,IndLigPAFC,IndLigPAFR : integer;
    IndiceMTFG,IndicePAFR,IndicePAFC,IndicePAFG,IndiceTps : integer;
    IndApplicFG,IndApplicFc,IndApplicFR,IndPrixPourQte,IndPuHt,IndPuHtDev,IndLigPuHtDev : integer;
    IndLigCFC,IndLigCFR,IndLigCFG : integer;
    MontantCharge,CoefMarg,CoefDev : double;
    DEv : RDevise;
begin
  if TOBPiece.GetValue('GP_VENTEACHAT') <> 'VEN' then exit; // bicose on ne calule pas sur les pieces achats ou autres
  //
  fillchar(Dev,sizeof(dev),#0);
  DEV.code := TOBPiece.getvalue('GP_DEVISE');
  GetInfosDevise(DEV);
  DEV.Taux := TOBPiece.getvalue('GP_TAUXDEV');
  //
	IndMontantPr := TOBPiece.GetNumChamp ('GP_MONTANTPR');
	IndMontantPA := TOBPiece.GetNumChamp ('GP_MONTANTPA');

	TOBpiece.putValeur(IndMontantPA,0);
  TOBpiece.putValeur(IndMontantPr,0);

  for Indice := 0 to TOBpiece.detail.count -1 do
  begin
  	TOBL := TOBPiece.detail[Indice];
    if IsSousdetail(TOBL) then continue;
    if TOBL.GetString('GL_TYPEARTICLE')='POU' then continue;
    CoefMarg := TOBL.GetValue('GL_COEFMARG');
    Coefdev := TOBL.GetValue('GL_TAUXDEV');

    if Indice = 0 then
    begin
    	IndQte := TOBL.GetNumChamp ('GL_QTEFACT');
      //
    	IndPA := TOBL.GetNumChamp ('GL_DPA');
    	IndPr := TOBL.GetNumChamp ('GL_DPR');
    	IndPuHt := TOBL.GetNumChamp ('GL_PUHT');
    	IndPuHtDev := TOBL.GetNumChamp ('GL_PUHTDEV');
      //
    	IndLigPA := TOBL.GetNumChamp ('GL_MONTANTPA');
    	IndLigPr := TOBL.GetNumChamp ('GL_MONTANTPR');
      //
      IndLigFG := TOBL.GetNumChamp ('GL_MONTANTFG');
      IndLigFC := TOBL.GetNumChamp ('GL_MONTANTFC');
      IndLigFR := TOBL.GetNumChamp ('GL_MONTANTFR');
      //
      IndLigPAFG := TOBL.GetNumChamp ('GL_MONTANTPAFG');
      IndLigPAFC := TOBL.GetNumChamp ('GL_MONTANTPAFC');
      IndLigPAFR := TOBL.GetNumChamp ('GL_MONTANTPAFR');
      IndLigPuHtDev := TOBL.GetNumChamp ('GL_MONTANTHTDEV');
      //
      IndLigCFG := TOBL.GetNumChamp ('GL_COEFFG');
      IndLigCFC := TOBL.GetNumChamp ('GL_COEFFC');
      IndLigCFR := TOBL.GetNumChamp ('GL_COEFFR');
      //
      IndApplicFG  := TOBL.GetNumChamp ('GLC_NONAPPLICFG');
      IndApplicFC  := TOBL.GetNumChamp ('GLC_NONAPPLICFC');
      IndApplicFR  := TOBL.GetNumChamp ('GLC_NONAPPLICFRAIS');
      IndPrixPourQte := TOBL.GetNumChamp ('GL_PRIXPOURQTE');
    end;
    //
    if (not IsArticle (TOBL)) Then
    begin
      if (not FromImportBordereaux) or
        ((FromImportBordereaux) and ((TOBL.GetString('GL_LIBELLE')='') or
                                     (TOBL.GetDouble ('GL_PUHTDEV')=0) or
                                     (TOBL.GetDouble ('GL_QTEFACT')=0))) then
      begin
        TOBL.putValue('GL_DPA',0);
        TOBL.putValue('GL_DPR',0);
        if (TOBL.GetValue('GL_BLOQUETARIF')='-') and (calculPv) then
        begin
          TOBL.putValue('GL_PUHTDEV',0);
          TOBL.putValue('GL_PUHT',0);
        end;
      end;
    end;
    if (TOBL.GetValue('GL_TYPELIGNE')<>'ART') and (TOBL.GetValue('GL_NATUREPIECEG')<>'PBT') Then continue;
    if (not Isarticle(TOBL)) and (TOBL.GetValue('GL_NATUREPIECEG')='PBT') Then continue;

    //
    if TOBL.GetValue('GL_RECALCULER')='X' Then
    begin
      TypeArticle := TOBL.GetValue('GL_TYPEARTICLE');
      if ((TypeArticle='OUV') or (TypeArticle='ARP')) and (TOBL.GetValue('GL_INDICENOMEN')>0) then
      begin
        CalculMontantsDocOuv(TOBPiece,TOBL,TOBOuvrage,calculPv,FromImportBordereaux);
//        TOBL.PutValeur (IndLigPA,Arrondi(TOBL.GetValeur(indPA)*TOBL.GetValeur(indQte)/TOBL.GetValeur (IndPrixPourQte),4));
        TOBL.PutValeur (IndLigPR,Arrondi(TOBL.GetValeur(indPr)*TOBL.GetValeur(indQte)/TOBL.GetValeur (IndPrixPourQte),V_PGI.okdecP));
(*        TOBL.PutValeur (IndLigPuHtDev,Arrondi(TOBL.GetValeur(indPuHtDev)*TOBL.GetValeur(indQte)/TOBL.GetValeur (IndPrixPourQte),V_PGI.okdecP));
        TOBL.putValeur(IndPuHt,Arrondi(TOBL.getValeur(IndPuHtDev)/CoefDev,V_PGI.okdecP));
*)
        if TOBL.GetValeur(indPr) <> 0 then
        begin
          TOBL.putValue('GL_COEFMARG',Arrondi(TOBL.GetValeur(IndPuHt)/TOBL.getValeur(IndPr),4));
          TOBL.PutValue('POURCENTMARG',Arrondi((TOBL.GetValue('GL_COEFMARG')-1)*100,2));
        end;
        if TOBL.GetValue('GL_PUHT') <> 0 then
        begin
          TOBL.PutValue('POURCENTMARQ',Arrondi(((TOBL.GetValue('GL_PUHT')- TOBL.GetValue('GL_DPR'))/TOBL.GetValue('GL_PUHT'))*100,2));
        end else
        begin
          TOBL.PutValue('POURCENTMARQ',0);
        end;

      end else
      begin
        TOBL.PutValeur (IndLigPA,Arrondi(TOBL.GetValeur(indPA)*TOBL.GetValeur(indQte)/TOBL.GetValeur (IndPrixPourQte),V_PGI.okDecV));

        if TOBL.GetValeur (IndApplicFg)<>'X' then TOBL.PutValeur (IndLigFG,Arrondi(TOBL.GetValeur(indLigPA)*TOBL.GetValeur(IndLigCFG),V_PGI.okDecV));
        MontantCharge := TOBL.GetValeur(indLigPA)+TOBL.GetValeur(indLigFG);
        //
        if TOBL.GetValeur (IndApplicFC)<>'X' then TOBL.PutValeur (IndLigFC,Arrondi(MontantCharge*TOBL.GetValeur(IndLigCFC),V_PGI.okDecV));
        MontantCharge := TOBL.GetValeur(indLigPA)+TOBL.GetValeur(indLigFG)+TOBL.GetValeur(indLigFC);
        //
        if TOBL.GetValeur (IndApplicFR)<>'X' then TOBL.PutValeur (IndLigFR,Arrondi(MontantCharge*TOBL.GetValeur(IndLigCFR),V_PGI.okDecV));

        TOBL.PutValeur (IndLigPr,Arrondi(MontantCharge+TOBL.GetValeur (IndLigFR),V_PGI.okDecV));
        if TOBL.GetValeur (IndQte) <> 0 then TOBL.PutValeur (IndPr,Arrondi(TOBL.GetValeur (IndLigPR)/TOBL.GetValeur (IndQte),V_PGI.okDecP));

        if (TOBL.GetValue('GL_BLOQUETARIF')='-') and (calculPv) then
        begin
          if CoefMarg <> 0 then
          begin
            TOBL.putValeur(IndPuHt,Arrondi(TOBL.getValeur(IndPr)*CoefMarg,V_PGI.okdecP));
            TOBL.putValeur(IndPuHtDev,Arrondi(TOBL.getValeur(IndPr)*CoefMarg*CoefDev,V_PGI.okdecP));
            TOBL.putValue('GL_RECALCULER','X');
          end;
        end else
        begin
        	TOBL.putValeur(indPuHt,DeviseToPivot(TOBL.GetValeur(IndPuHtDev),DEV.Taux,DEV.quotite));
          if TOBL.GetValeur(indPr) <> 0 then
          begin
            CoefMarg := Arrondi(TOBL.GetValeur(IndPuHt)/TOBL.getValeur(IndPr),4);
            TOBL.putValue('GL_COEFMARG',CoefMarg);
            TOBL.PutValue('POURCENTMARG',Arrondi((TOBL.GetValue('GL_COEFMARG')-1)*100,2));
            if Pos(TOBL.getValue('GL_NATUREPIECEG'),'FBT;ABT;ABP;FPR;FAC;AVC;FBP;BAP')>0 then
            begin
              if Arrondi(TOBL.getValeur(IndPr)*CoefMarg,V_PGI.okdecP)<>TOBL.GetValeur(IndPuHt) then
              begin
                TOBL.PutValue('GL_BLOQUETARIF','X');
              end;
            end else
            begin
              if (not FromImportBordereaux) and (TOBL.GetValue('GL_BLOQUETARIF')='-') and (calculPv) then
              begin
                TOBL.putValeur(indPuHt,Arrondi(TOBL.getValeur(IndPr)*CoefMarg,V_PGI.okdecP));
                TOBL.putValue('GL_PUHTDEV', PIVOTTODEVISE(TOBL.getValEUR(IndPuHt),DEV.Taux ,Dev.Quotite ,V_PGI.okdecP));
              end else
              begin
                TOBL.putValue('GL_COEFMARG',0);
              end;
            end;
          end else
          begin
            TOBL.putValue('GL_COEFMARG',0);
          end;
          if TOBL.GetValue('GL_PUHT') <> 0 then
          begin
            TOBL.PutValue('POURCENTMARQ',Arrondi(((TOBL.GetValue('GL_PUHT')- TOBL.GetValue('GL_DPR'))/TOBL.GetValue('GL_PUHT'))*100,2));
          end else
          begin
            TOBL.PutValue('POURCENTMARQ',0);
          end;
        end;
      end;

    end;
    TOBPiece.PutValeur (IndMontantPr,TOBPiece.getValeur (IndMontantPr)+TOBL.GetValeur (IndligPr));
    TOBPiece.PutValeur (IndMontantPA,TOBPiece.getValeur (IndMontantPA)+TOBL.GetValeur (IndligPA));
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Louis DECOSSE

Créé le ...... : 19/01/2000

Modifié le ... : 28/08/2003
Description .. : Fonction "ouverte" avant calcul de la saisie pour code
Suite ........ : spécifique
Mots clefs ... : SAISIE;CALCUL;SPECIFIQUE;CONTEXTE;

*****************************************************************}

Function BeforeCalcul ( FF : TForm ; TOBPiece,TOBBases,TOBTiers,TOBArticles,TOBOuvrage,TOBCpta,TOBCata,TOBAffaire,TOBGCS,TOBPorcs : TOB ; DEV : RDEVISE; var Arow: integer ; CalculPv : boolean=true; RazFg : boolean=false; OnlyPv : boolean=false; Acol : Integer = -1) : boolean ;

	function VasyLeCalcul (OldTotalBFG,OldTotalBFR,OldTotalBFC,OLdTotalFG,OLdTotalFc,OldTotalFR: double ; TOBPiece : TOB) : boolean;
  var Montantchanged,CoefRepartisChanged,CoefNul : boolean;
  begin
    MontantChanged := (TOBPiece.getValue('GP_MONTANTPAFG')<>OldTotalBFG) or
  						        (TOBPiece.getValue('GP_MONTANTPAFC')<>OldTotalBFC) or
    					        (TOBPiece.getValue('GP_MONTANTPAFR')<>OldTotalBFR) or
                  		(TOBPiece.getValue('GP_MONTANTFG')<>OldTotalFG);
    CoefRepartisChanged := (TOBPiece.getValue('GP_COEFFR')<>OldTotalFR) or
    					             (TOBPiece.getValue('GP_COEFFC')<>OldTotalFC);
    CoefNul := ((TOBPiece.getValue('GP_COEFFR')=0) and (OldTotalFR = 0)) and
                  ((TOBPiece.getValue('GP_COEFFC')=0) and (OldTotalFC = 0)) ;

    result := ((CoefRepartisChanged or MontantChanged) and (not CoefNul));
  end;

  function OkCalcPUV ( acol : Integer) : boolean;
  begin
		Result := false;
    if (Acol = -1) or (Acol = SG_PXAch) or (Acol = SG_Px) or (Acol = SG_RefArt) or (Acol = SG_RL) or (Acol = SG_PxNet) then
    begin
      Result := True;
    end;
  end;

var OldTotalBFG,OldTotalBFR,OldTotalBFC,OldTotalFG,OldTotalFR,OldTotalFC : double;
BEGIN
Result:=True ;
{$IFDEF BTP}
  if (isPieceGerableFraisDetail (TOBPiece.GetValue('GP_NATUREPIECEG'))) and (not OnlyPv) then
  begin
    // Sauvegarde
    OldTotalBFG := TOBPiece.GetValue('GP_MONTANTPAFG');
    OldTotalBFC := TOBPiece.GetValue('GP_MONTANTPAFC');
    OldTotalBFR := TOBPiece.getValue('GP_MONTANTPAFR');
    OldTotalFG := TOBPiece.GetValue('GP_MONTANTFG');
    OldTotalFC := TOBPiece.GetValue('GP_COEFFC');
    OldTotalFR := TOBPiece.getValue('GP_COEFFR');
    //
  	BeforeApplicFrais (TOBPiece,TOBOUvrage,RAZFg);
  	RecalculeCoefFrais (TOBPiece,TOBPorcs);   // calcul des Taux Fc et TAux FG
    if ( VasyLeCalcul (OldTotalBFG,OldTotalBFR,OldTotalBFC,oldTotalFG,OLdTotalFc,OldTotalFR,TOBPiece)) or (TOBPiece.getValue('GP_RECALCACHAT')='X')  then
    begin
    	TOBpiece.putValue('GP_MONTANTPR',0);
    	TOBpiece.putValue('GP_MONTANTFG',0); // car deja calcule precedemment
			AppliqueFraisPiece (TOBPiece,TOBOuvrage,true,true,(TOBPiece.getVAlue('GP_FACTUREHT')='X'),DEV);
      TOBPiece.PutValue('GP_RECALCACHAT','-');
      Arow := -1; // force le calcul pour l'ensemble des lignes
    end else
    begin
  	  CalculeMontantsDoc (TOBPiece,TOBOuvrage,(calculPv and OkCalcPUV(Acol)));
    end;
  end else
  begin
  	CalculeMontantsDoc (TOBPiece,TOBOuvrage,(calculPv and OkCalcPUV(Acol)));
  end;
  TFFacture(FF).RAZFg := false;
// une fois les calculs terminé on mets en place les lignes d'ouvrages dans le document pour les calcul TVA etc..
//  AjouteLignesVirtuellesOuv (TOBPIECE,TOBOuvrage,TOBArticles,TOBCpta,TOBTiers,TOBCata,TOBAffaire,TOBGCS,DEV);
{$ENDIF}
if ctxAffaire in V_PGI.PGIContexte then ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Louis DECOSSE

Créé le ...... : 19/01/2000

Modifié le ... : 28/08/2003
Description .. : Procédure "ouverte" après calcul de la saisie pour code
Suite ........ : spécifique
Mots clefs ... : SAISIE;CALCUL;SPECIFIQUE;CONTEXTE;

*****************************************************************}

Procedure AfterCalcul ( FF : TForm ; TOBPiece,TOBBases,TOBTiers,TOBArticles,TOBPieceRg,TOBBasesRG : TOB ; DEV : RDEVISE ) ;
BEGIN
{$IFDEF BTP}
//DetruitLignesVirtuellesOuv (TOBPIece,DEV);
{$ENDIF}
if ctxAffaire in V_PGI.PGIContexte then ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Louis DECOSSE

Créé le ...... : 19/01/2000

Modifié le ... : 28/08/2003
Description .. : Fonction "ouverte" avant validation de la saisie pour code
Suite ........ : spécifique
Mots clefs ... : SAISIE;SPECIFIQUE;CONTEXTE;VALIDATION;ENREGISTREMENT;

*****************************************************************}

Function BeforeValide ( FF : TForm ; TOBPiece,TOBBases,TOBTiers,TOBArticles : TOB ; DEV : RDEVISE ; Old_TOTALHT : double ) : boolean ;
Var Ind  : Integer ;
    TOBL : TOB ;
    FFact : TFFacture;
    Requete : String;
    RefInt, s,NextPieces: String;
BEGIN
  FFact := TFFacture(FF);
  Result:=True ;
  NextPieces := GetParamSocSecur('SO_BTCHOIXPIECEPROPOS','DBT')+';'+GetParamSocSecur('SO_BTCHOIXPIECENEG','DE');

  {Affaire}
  if (ctxAffaire in V_PGI.PGIContexte) or (VH_GC.GCSeria) then
  begin
    if (FFact.SaisieTypeAffaire) And (TOBPiece.Detail.count=1) then
    begin
      // pb si une seule ligne vide sur la tob - ligne à supprimer ( pièce sans ligne d'affaires)
      if Not EstRempliGC(FFact.GS,1) then TOBPiece.Detail[0].Free ;
    end;
  end;

  if (ctxGRC in V_PGI.PGIContexte) and
    (TobPiece.getValue('GP_PERSPECTIVE')=0) and
    (Pos(TobPiece.getValue('GP_NATUREPIECEG'),NextPieces)>0) and
    (not FFact.SaisieTypeAvanc )  then
  begin
    {$IFDEF GRC}
    {$IFNDEF CCS3}
    Result:=RTAffecterPerspective (TobPiece );
    {$ENDIF}
    {$ENDIF}
  end;
  // cas de la modification d'un devis rattaché à une proposition
  if (ctxGRC in V_PGI.PGIContexte) and
      (TobPiece.getValue('GP_PERSPECTIVE')<>0) and
      (Pos(TobPiece.getValue('GP_NATUREPIECEG'),NextPieces)>0) and (Old_TOTALHT <> 0 ) then
  begin
    {$IFDEF GRC}
    {$IFNDEF CCS3}
    if (GetParamsoc('SO_RTPROPOPTCALCUL') = 'NSA') or
    (GetParamsoc('SO_RTPROPOPTCALCUL') = 'SSP') then
    begin
      Requete:='UPDATE PERSPECTIVES SET RPE_MONTANTPER=RPE_MONTANTPER-'
      +StrFPoint(Old_TOTALHT)+' + '
      +StrFPoint(TobPiece.GetValue('GP_TOTALHT'))
      +' WHERE RPE_PERSPECTIVE='+IntToStr(TobPiece.GetValue('GP_PERSPECTIVE'));
      ExecuteSQL(Requete);
    end;
    {$ENDIF}
    {$ENDIF}
  end;
  {$IFDEF BTP}
  if FFact.ModifAvanc then
  begin
    // Dans le cas de l'avancement on recalcule le GL_QTESIT afin de pouvoir ensuite reactualiser
    // le devis initial
    for ind := 0 to TOBPiece.detail.count -1 do
    begin
      TOBL:=TOBPiece.Detail[Ind];
      if TOBL.GetValue('GL_TYPELIGNE') = 'ART' then
      begin
        if (IsgestionAvanc(TOBpiece)) and (TOBL.FieldExists ('BLF_QTESITUATION')) then
        begin
          TOBL.PutValue ('GL_QTEFACT',TOBL.GetValue('BLF_QTESITUATION'));
        end;
        TOBL.PutValue ('GL_QTESIT',TOBL.GetValue('DEJAFACT')+TOBL.GetValue('GL_QTEFACT'));
      end;
    end;
  end;
  {$ENDIF}
  {Front Office}
  {$IFDEF FOS5}
  if ctxFO in V_PGI.PGIContexte then
  BEGIN
    TOBPiece.PutValue('GP_CAISSE', FOCaisseCourante) ;
    TOBPiece.PutValue('GP_NUMZCAISSE', FOGetNumZCaisse(TOBPiece.GetValue('GP_CAISSE'))) ;
    for Ind:=0 to TOBPiece.Detail.Count-1 do
    BEGIN
      TOBL:=TOBPiece.Detail[Ind] ;
      TOBL.PutValue('GL_CAISSE', TOBPiece.GetValue('GP_CAISSE')) ;
    END ;
  END ;
  {$ENDIF}

  {$IFDEF GPAO}
  { Vérif. des champ obligatoires sur l'entête de la pièce }
  s := '';
  if not wInitFields(TobPiece, 'PIECE', iaVerif, s) then
  begin
    PgiInfo(wGetLastErrorMessageFromWIL(s), '');
    Result := False;
    EXIT;
  end;
  { Vérif. des champ obligatoires sur les lignes de la pièce }
  Ind := -1;
  while (Ind < TobPiece.Detail.Count - 1) and Result do
  begin
    Inc(Ind);
    Result := wInitFields(TobPiece.Detail[Ind], 'LIGNE', iaVerif, s);
  end;
  if not Result then
  begin
    PgiInfo(wGetLastErrorMessageFromWIL(s)+ ' (' + TraduireMemoire('Sur la ligne de pièce N° ') + ' ' + IntToStr(Ind + 1) + ')', '');
    Result := False;
    EXIT;
  end;
  {$ENDIF}

  RefInt := GetInfoParPiece(TobPiece.GetValue('GP_NATUREPIECEG'), 'GPP_REFINTEXT');
  { Vérifie l'unicité de la référence interne }
  Result := ReferenceIntExtIsOk(crInterne, TobPiece, TobPiece.GetValue('GP_REFINTERNE'));
  { Et vérifie l'unicité de la référence externe }
  if Result then
    Result := ReferenceIntExtIsOk(crExterne, TobPiece, TobPiece.GetValue('GP_REFEXTERNE'));

END ;

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Louis DECOSSE
Créé le ...... : 19/01/2000
Modifié le ... : 27/01/2000
Description .. : Procédure "ouverte" après validation de la saisie pour code spécifique
Mots clefs ... : SAISIE;SPECIFIQUE;CONTEXTE;VALIDATION;ENREGISTREMENT;
*****************************************************************}
Procedure AfterValide ( FF : TForm ; TOBPiece,TOBBases,TOBTiers,TOBArticles : TOB ; DEV : RDEVISE ) ;
var stChainage  : String;
    StParametre : String;
BEGIN

if TFFacture(FF).SaisieTypeAffaire Then
   BEGIN
   if TheTob<>nil then TheTob.Free;
   TheTob:=TOB.Create('PIECE',Nil,-1) ;
   TheTOB.Dupliquer(TOBPiece,True,True) ;
      END ;

   if ((ctxGrc in V_PGI.PGIContexte ) or (VH_GC.GRFSeria)) and
      ((TFFacture(FF).Action = taCreat)      or
       (TFFacture(FF).Action = taModif)      or
       (TFFacture(FF).DuplicPiece)           or
       (TFFacture(FF).TransfoPiece)) then
      begin
      stChainage:='';
      if TobPiece.GetValue('GP_DOMAINE') <> '' then
         stChainage := GetInfoParPieceDomaine(TFFacture(FF).NewNature, TobPiece.GetValue('GP_DOMAINE'), 'GDP_CHAINAGE');
      if (stChainage='') then
         stChainage := GetInfoParPiece(TobPiece.GetValue('GP_NATUREPIECEG'), 'GPP_CHAINAGE');
      if (stChainage <> '') then
      if (TFFacture(FF).Action = taModif) then
         begin
         if (TobPiece.GetString('GP_ETATVISA') = 'ATT') then
            RTRechChainagePiece(TFFacture(FF).CleDoc,TobPiece.GetString('GP_VENTEACHAT'),false);
         end
      else
         {$IFNDEF BTP}
         RtGenerationChainage(PieceContainer,stChainage);
         {$ELSE}
         Begin
         StParametre := StParametre + TobPiece.GetString('GP_TIERS');
         StParametre := StParametre + ';' + TobTiers.GetString('T_LIBELLE');
         StParametre := StParametre + ';' + TobPiece.GetString('GP_AFFAIRE');
         StParametre := StParametre + ';' + TobPiece.GetString('GP_RESSOURCE');
         StParametre := StParametre + ';' + TobPiece.GetString('GP_VENTEACHAT');
         StParametre := StParametre + ';' + stChainage;
         BtGenerationChainage(TobPiece, StParametre);
         end;
         {$ENDIF}
      end;


END ;

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Louis DECOSSE
Créé le ...... : 19/01/2000
Modifié le ... : 01/02/2000
Description .. : Fonction "ouverte" avant appel de la création du reliquat
Mots clefs ... : RELIQUAT;SPECIFIQUE;CONTEXTE;
*****************************************************************}
Function BeforeReliquat ( FF : TForm ; TOBPiece,TOBBases,TOBTiers,TOBArticles : TOB ; DEV : RDEVISE ) : boolean ;
BEGIN
Result:=True ;
if ctxScot in V_PGI.PGIContexte then BEGIN Result:=False ; Exit ; END ;
if ctxAffaire in V_PGI.PGIContexte then ;
END ;

Procedure BeforeUpdateAncien(  TOBPiece_O,TOBPiece,TOBAcomptes_O,TobTiers : TOB);
BEGIN
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Louis DECOSSE
Créé le ...... : 19/01/2000
Modifié le ... : 01/02/2000
Description .. : Fonction "ouverte" avant appel de GetCellCanvas
Mots clefs ... : SPECIFIQUE;CONTEXTE;CANVAS;GETCELLCANVAS;
*****************************************************************}
Function BeforeGetCellCanvas ( FF : TForm ; ACol,ARow : Longint; Canvas : TCanvas ; AState: TGridDrawState) : boolean ;
BEGIN
Result:=True ;
if ctxAffaire in V_PGI.PGIContexte then ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Patrice ARANEGA
Créé le ...... : 19/01/2000
Modifié le ... : 01/02/2000
Description .. : Fonction "ouverte" avant appel de FormShow
Mots clefs ... : SPECIFIQUE;CONTEXTE;FORMSHOW;
*****************************************************************}
Function BeforeShow ( FF : TForm ) : boolean ;
Var i :integer;
    CompEntete : TComponent;
    NameCompEntete : String;
    FFact : TFFacture;
    Part0,Part1,Part2,PArt3,CodeAvenant:string;
BEGIN
Result:=True ;
FFact := TFFacture(FF);
FFact.SaisieTypeAffaire:=False;
{Module planning présent}
FFact.PlanningEntete.Visible := GerePlanningPiece(FFact.CleDoc.NaturePiece);
FFact.PlanningLigne.Visible := GerePlanningPiece(FFact.CleDoc.NaturePiece);
// Modif BTP

{Affaires et BTP}
{$IFDEF BTP}
if (ctxAffaire in V_PGI.Pgicontexte) then
   begin
      FFact.GP_TIERS__.Text := FFact.LeCodeTiers;
      FFact.GP_TIERS.Text := Ffact.LeCodeTiers;
     if FFact.LAffaire <> '' then
     begin
      BTPCodeAffaireDecoupe(FFact.LAffaire, Part0, Part1, Part2, Part3,CodeAvenant,taConsult,False);
      FFact.GP_AFFAIRE.Text := FFact.LAffaire;
      FFact.GP_AFFAIRE0.Text := Part0;
      FFact.GP_AFFAIRE1.Text := Part1;
      FFact.GP_AFFAIRE2.Text := Part2;
      FFact.GP_AFFAIRE3.Text := Part3;
      FFact.GP_AVENANT__.Text := CodeAvenant;
    end;
if FFact.LAvenant <> '' then
    FFact.Codeaffaireavenant := FFact.LAvenant
else
    FFact.Codeaffaireavenant := CodeAvenant;
if FFact.Codeaffaireavenant = '' then FFact.Codeaffaireavenant := '00';
   end;
{$ELSE}
if (ctxAffaire in V_PGI.Pgicontexte) then
   begin
   FFact.GP_TIERS__.Text := FFact.LeCodeTiers;
   FFact.GP_TIERS.Text := Ffact.LeCodeTiers;
   if FFact.LAffaire <> '' then
   begin
     CodeAffaireDecoupe(FFact.LAffaire, Part0, Part1, Part2, Part3,CodeAvenant,taConsult,False);
     FFact.GP_AFFAIRE.Text := FFact.LAffaire;
     FFact.GP_AFFAIRE0.Text := Part0;
     FFact.GP_AFFAIRE1.Text := Part1;
     FFact.GP_AFFAIRE2.Text := Part2;
     FFact.GP_AFFAIRE3.Text := Part3;
     FFact.GP_AVENANT__.Text := CodeAvenant;
   end;
if FFact.LAvenant <> '' then
    FFact.Codeaffaireavenant := FFact.LAvenant
else
    FFact.Codeaffaireavenant := CodeAvenant;
if FFact.Codeaffaireavenant = '' then FFact.Codeaffaireavenant := '00';
   end;
{$ENDIF}

{Affaires}
{sans affaire on cache les composants affaires}
if (Not(ctxAffaire in V_PGI.PGIContexte) and Not (ctxGCAFF in V_PGI.PGIcontexte)) or
   (GetInfoParPiece(FFact.NewNature,'GPP_LIENAFFAIRE')='-') or
   ((ctxGCAFF in V_PGI.PGIContexte) and (V_PGI.LaSerie=S3)) then
   BEGIN
   FFact.HGP_AFFAIRE.Visible  := False;
   FFact.GP_AVENANT.Visible   := False;
   FFact.BRECHAFFAIRE.Visible := False;
   FFact.GP_AFFAIRE1.Visible  := False;
   FFact.GP_AFFAIRE2.Visible  := False;
   FFact.GP_AFFAIRE3.Visible  := False;
   END else
{$IFDEF AFFAIRE}
if (ctxGCAFF in V_PGI.PGIContexte) then
    ChargeCleAffaire(Nil,FFact.GP_AFFAIRE1,FFact.GP_AFFAIRE2,FFact.GP_AFFAIRE3,FFact.GP_AVENANT,Nil,FFact.Action,'',False);
{$endif}
// positionnement de la variable SaisieTypeAffaire
if (ctxAffaire in V_PGI.PGIContexte) or (VH_GC.GASeria) then
   begin
   if (GetInfoParPiece(FFact.CleDoc.NaturePiece,'GPP_PREVUAFFAIRE')='X') Then FFact.SaisieTypeAffaire:= True;
   end;

if (ctxAffaire in V_PGI.PGIContexte) or (FFact.SaisieTypeAffaire) then
   BEGIN
    // Changement du panel d'entête affiché en mode affaire
    FFact.PEntete.Visible       :=False ;
    FFact.PEntete.Align         :=alNone ;
    FFact.PEnteteAffaire.Visible:=True ;
    FFact.PEnteteAffaire.Align  :=alTop ;
    if FFact.SaisieTypeAffaire then FFact.PEnteteAffaire.Enabled := False;
    // Réaffectation des noms en __ de l'entête affaire
    for i:=0 to FFact.ComponentCount-1 do
        Begin
        if (Copy (FFact.components[i].Name, length(FFact.components[i].Name)-1, 2) = '__') then
            BEGIN
            NameCompEntete := Copy (FFact.components[i].Name,0,length(FFact.components[i].Name)-2);
            CompEntete := FFact.FindComponent(NameCompEntete);
            if CompEntete <> Nil then
                BEGIN
                CompEntete.Name := CompEntete.Name + '_O';
                FFact.components[i].Name := NameCompEntete;
                END;
            END;
        End;
{$IFDEF AFFAIRE}
{$IFDEF BTP}
      ChargeCleAffaire(FFact.GP_AFFAIRE0,FFact.GP_AFFAIRE1,FFact.GP_AFFAIRE2,FFact.GP_AFFAIRE3,FFact.GP_AVENANT,Nil,FFact.Action,FFact.GP_AFFAIRE.text,False);
{$ELSE}
      ChargeCleAffaire(Nil,FFact.GP_AFFAIRE1,FFact.GP_AFFAIRE2,FFact.GP_AFFAIRE3,FFact.GP_AVENANT,Nil,FFact.Action,'',False);
{$ENDIF}
{$ENDIF}

    FFact.AdrLiv.Visible := False; FFact.AdrFac.Visible := False;
    if (ctxTempo in V_PGI.PGIContexte) And (GetInfoParPiece(FFact.NewNature,'GPP_VENTEACHAT')='ACH') then
       FFact.MBSoldeReliquat.Visible := true
    else FFact.MBSoldeReliquat.Visible := False;
    if ctxScot in v_PGI.PGIContexte then FFact.MBTarif.Visible := False;
{$IFNDEF BTP}
      FFact.MBDetailNomen.Visible := False;
      FFact.N1.Visible := False;
      FFact.N2.Visible := False;
      FFact.N3.Visible := False;
{$ELSE}
    FFact.MBSoldeReliquat.Visible := true;
    FFact.MBSoldeTousReliquat.Visible := true;
{$ENDIF}
    if FFact.SaisieTypeAffaire then
        BEGIN
        FFact.HGP_ETABLISSEMENT.Visible := False;
        FFact.GP_ETABLISSEMENT.Visible  := False;
        FFact.HGP_REPRESENTANT.Visible  := False;
        FFact.GP_REPRESENTANT.Visible   := False;
        FFact.HGP_DEPOT.Visible         := False;
        FFact.GP_DEPOT.Visible          := False;
        FFact.HGP_DATEPIECE.Visible     := False;
        FFact.GP_DATEPIECE.Visible      := False;
        FFact.HGP_NUMPIECE.Visible      := False;
        FFact.GP_NUMEROPIECE.Visible    := False;
        FFact.HGP_REFINTERNE.Visible    := False;
        FFact.HGP_REFEXTERNE.Visible    := False;
        FFact.GP_REFINTERNE.Visible     := False;
        FFact.BRazAffaire.Visible       := False;
        FFact.BRechaffaire.Visible      := False;
        FFact.Caption := 'Saisie des lignes';
        FFact.HelpContext := 120000500;
        END else
        BEGIN
        if ctxScot in V_PGI.PGIContexte then
            BEGIN
            FFact.GP_DEPOT.Visible          := False; FFact.HGP_DEPOT.Visible        := False;
            FFact.HGP_REPRESENTANT.Visible  := False; FFact.GP_REPRESENTANT.Visible  := False;
            END
        else
            BEGIN
{$IFDEF BTP}
            FFact.HGP_DEPOT.Visible := True;
            FFact.GP_DEPOT.Visible  := True;
{$ELSE}
            FFact.HGP_DEPOT.Visible := False;
            FFact.GP_DEPOT.Visible  := False;
{$ENDIF}
            END;
        END;
    END ;
{Front Office}
if ctxFO in V_PGI.PGIContexte then
    BEGIN
   if (FFact.CleDoc.NaturePiece='FCF') or (FFact.CleDoc.NaturePiece='FRF') then
      BEGIN
      FFact.GP_ETABLISSEMENT.Visible := False ;
      FFact.HGP_ETABLISSEMENT.Visible := False ;
      FFact.GP_DEVISE.Visible := False ;
      FFact.BDEVISE.Visible := False ;
      FFact.LDEVISE.Visible := False ;
      FFact.GP_DEPOT.Visible := False ;
      FFact.HGP_DEPOT.Visible := False ;
      FFact.GP_ESCOMPTE.Visible := False ;
      FFact.HGP_ESCOMPTE.Visible := False ;
      FFact.GP_REMISEPIED.Visible := False ;
      FFact.HGP_REMISEPIED.Visible := False ;
      FFact.GP_MODEREGLE.Visible := False ;
      FFact.HGP_MODEREGLE.Visible := False ;
      FFact.PTotaux.Visible := False ;
      FFact.PTotaux1.Visible := False ;
      FFact.GP_DATELIVRAISON.Left := FFact.GP_REPRESENTANT.Left ;
      FFact.GP_DATELIVRAISON.Top := FFact.GP_REPRESENTANT.Top ;
      FFact.HGP_DATELIVRAISON.Left := FFact.HGP_REPRESENTANT.Left ;
      FFact.HGP_DATELIVRAISON.Top := FFact.HGP_REPRESENTANT.Top ;
      END ;
    END ;

if GetInfoParPiece(FFact.CleDoc.NaturePiece, 'GPP_REFINTEXT') = 'INT' then
   begin
   FFact.HGP_REFINTERNE.Visible  := True;
   FFact.HGP_REFEXTERNE.Visible  := false;
   FFact.GP_REFINTERNE.Visible   := True;
   FFact.GP_REFEXTERNE.Visible   := False;
   FFact.GP_DATEREFEXTERNE.Visible := False;
   end
else if GetInfoParPiece(FFact.CleDoc.NaturePiece, 'GPP_REFINTEXT') = 'EXT' then
   begin
   FFact.HGP_REFEXTERNE.Visible := True;
   FFact.HGP_REFINTERNE.Visible := false;
   FFact.GP_REFINTERNE.Visible   := False;
   FFact.GP_REFEXTERNE.Visible   := True;
   FFact.GP_DATEREFEXTERNE.Visible := False;
   end
else if GetInfoParPiece(FFact.CleDoc.NaturePiece, 'GPP_REFINTEXT') = 'NON' then
   begin
   FFact.HGP_REFINTERNE.Visible := False;
   FFact.GP_REFINTERNE.Visible   := False;

   FFact.HGP_REFEXTERNE.Visible := False;
   FFact.GP_REFEXTERNE.Visible     := False;
   FFact.GP_DATEREFEXTERNE.Visible := False;
   end;

END;

function LibBaseValoContrat (Periodicite : string) : string;
Var LibPer : string;
begin
if Periodicite ='DMM' then libPer := 'Mensuelle' else
if Periodicite ='DME' then libPer := 'de l''échéance' else
if Periodicite ='DMA' then libPer := 'Annuelle' ;
Result := 'Valorisation '+ LibPer
end;

{***********A.G.L.***********************************************
Auteur  ...... : Patrice ARANEGA
Créé le ...... : 19/01/2000
Modifié le ... : 01/02/2000
Description .. : Procédure "ouverte" après appel de FormShow
Mots clefs ... : SPECIFIQUE;CONTEXTE;FORMSHOW;
*****************************************************************}
Procedure AfterShow ( FF : TForm ) ;
Var FFact : TFFacture;
BEGIN
FFact := TFFacture(FF);

if (ctxAffaire in V_PGI.PGIContexte) or (FFact.SaisieTypeAffaire) then
    BEGIN
    if FFact.SaisieTypeAffaire then
        BEGIN
        FFact.GS.SetFocus ;
        FFact.FTitrePiece.Caption := 'Saisie des lignes';
        FFact.BDelete.Visible := False;
        FFact.bSaisieNatAffaire := (FFact.CleDoc.NaturePiece=VH_GC.AFNatAffaire)or(FFact.CleDoc.NaturePiece=VH_GC.AFNatProposition);
        if TheTob <> Nil then
         begin
         // récup des infos de l'affaire
         if TheTob.NomTable = 'AFFAIRE' then
            begin
            if TheTob.GetValue('AFF_GENERAUTO')='CON' then
               begin
               FFact.FTitrePiece.Caption := 'CONTRAT';
               FFact.LibComplAffaire.Caption := LibBaseValocontrat(TheTob.GetValue('AFF_DETECHEANCE'));
               end;
            end;
         TheTob.Free; TheTob := Nil;
         end;
        END;
    END ;
{$IFDEF BTP}
if (FFact.Action = taConsult) and (SG_TYPA <> -1) and (FFact.GS.ColLengths[SG_TYPA]< 0) then FFact.GS.ColLengths [SG_TYPA] := 0;
FFact.ChargeTiers ;
if (tModeSaisieBordereau in FFact.SaContexte) then
begin
	// on cache les elements de valorisation qui ne servent a rien
  FFact.PTotaux.visible := false;
  FFact.PTotaux1.visible := false;
  FFact.GP_ESCOMPTE.Visible := false;
  FFact.GP_REMISEPIED.Visible := false;
  FFact.GP_MODEREGLE.Visible := false;
  FFact.HGP_ESCOMPTE.Visible := false;
  FFact.HGP_REMISEPIED.Visible := false;
  FFact.HGP_MODEREGLE.Visible := false;
  FFact.HGP_MODEREGLE.Visible := false;
  FFact.LLIBRG.Visible := false;
  FFact.LMontantRg.Visible := false;
  FFact.BRetenuGar.visible := false;
  FFact.BDelete.Visible := false;
  FFact.MBREPARTTVA.Visible := false;
  FFact.VoirFranc.Visible := false;
  //FFact.AnalyseDocHtrait.Visible := false;
  FFact.SIMULATIONRENTABILIT1.Visible := false;
  FFact.Modebordereau.Visible := false;
  FFact.BSousTotal.Visible := false;
  FFact.BPorcs.Visible := false;
  FFact.Bprixmarche.Visible := false;
  FFact.PEnteteAffaire.Enabled := false;
  FFACT.TInsParag.visible := false;
  FFACT.TFusionner.Visible := false;
  FFACT.TCommentEnt.Visible := false;
  FFACT.TCommentPied.Visible := false;
  // Sur l'action ligne
  if FFACT.OrigineEXCEL then
  begin
  	FFACT.TInsLigne.Visible := false;
  end;
end;

{$ENDIF}
{$IFDEF CHR}
FFact.ChargeTiers ;
{$ENDIF}
END;

Procedure AfterShowForLine (FF : TForm; NaturePiece : String);
Var FFact : TFFacture;
BEGIN

	FFact := TFFacture(FF);

  With FFact do
  begin
    //Modification label Affaire en lebeau Chantier
    HGP_AFFAIRE.Caption := 'Chantier';
  	// -- pannel
  	HGP_ETABLISSEMENT.visible := false;
  	GP_ETABLISSEMENT.visible := false;
    Bdevise.DropdownMenu := nil;
    Bdevise.DropdownArrow := false;
  	HGP_REPRESENTANT.visible := false;
  	GP_REPRESENTANT.visible := false;
  	HGP_DEPOT.visible := false;
		GP_DEPOT.visible := false;
    BDevise.Visible := false;
    LDevise.Visible := false;
    EUROPIVOT.Visible := false;
    GP_DEVISE.visible := false;
    ISigneEuro.Visible := False;
  	HGP_DOMAINE.visible := false;
		GP_DOMAINE.visible := false;

    // -- Options
    BZoomDevise.visible := false;
    BZoomCommercial.visible := false;
    BZoomTarif.visible := false;
    BVentil.visible := false;
    // --
    BZoomDevise.Enabled := false;
    BZoomCommercial.Enabled := false;
    BZoomTarif.Enabled := false;

    // -- Actions complémentaires
    MBREPARTTVA.visible := false;
    SepTva1000.visible := false;
    VariablesMetres.visible := false;
    MBSoldeReliquat.visible := false;
    MBSoldeTousReliquat.visible := false;
    N4.visible := false;
    FraisDetail.visible := false;
    // -- Boutons du bas de piece
    BPorcs.visible := false;
//    BCalculDocAuto.visible := false;
    BcalculDocAuto.Down := True;
    RecalculeDocument.Visible := False;
    Bminute.visible := false;
    //--- Gestion des boutons en fonction du type de document
    //-- si Prévision de chantier
    if NaturePiece = 'PBT' Then
    	 begin
       BcalculDocAuto.Down := True;
	     hgp_escompte.Visible := False;
   		 hgp_remisePied.Visible := false;
  	   hgp_Moderegle.Visible := false;
	     gp_escompte.Visible := False;
	     gp_remisePied.Visible := false;
	     gp_Moderegle.Visible := false;
       RecalculeDocument.Visible := False;
       BPorcs.Visible := false;
       BRetenuGar.Visible := false;
       end;
  end;

end;

Procedure AfterBloquePiece ( FF : TForm ; Action : TActionFiche ; Bloc : Boolean ) ;
Var FFact : TFFacture;
BEGIN
FFact := TFFacture(FF);

if ctxAffaire in V_PGI.PGIContexte then
   BEGIN
   if Action=taConsult then
      BEGIN
      FFact.PEnteteAffaire.Enabled:=Not Bloc ;
      END else
   if Action=taModif then
      BEGIN
{$IFNDEF BTP}
      if Not(FFact.DuplicPiece) then  // libre en duplication de pièce
         BEGIN
         FFact.GP_AFFAIRE1.Enabled:=Not Bloc;
         FFact.GP_AFFAIRE2.Enabled:=Not Bloc;
         FFact.GP_AFFAIRE3.Enabled:=Not Bloc;
         FFact.GP_AVENANT.Enabled:=Not Bloc;
         FFact.BRechAffaire.Enabled:=Not Bloc;
         FFact.BRazAffaire.Enabled:=Not Bloc;
         END;
{$ENDIF}
      END ;
   END ;
END;

Procedure BeforeGSEnter ( FF : TForm ) ;
// Var FFact : TFFacture ;
BEGIN
// FFact:=TFFacture(FF) ;
END ;

Function FindLCD ( FFact : TFFacture ) : TLCDLabel ;
BEGIN
Result:=TLCDLabel(FFact.FindComponent('LCD')) ;
END ;

Procedure AfterInitEnteteDefaut ( FF : TForm ) ;
// Var FFact : TFFacture ;
BEGIN
// FFact:=TFFacture(FF) ;
END ;

Procedure AfterShowDetail ( FF : TForm; TOBPiece, TOBLig : TOB; RefUnique : String ) ;
// Var FFact : TFFacture ;
BEGIN
// FFact:=TFFacture(FF) ;
END ;

Procedure AfterFormCreate ( FF : TForm ) ;
// Var FFact : TFFacture ;
BEGIN
// FFact := TFFacture(FF) ;
END ;

// Permet de se repositionner à la ligne de saisie suivante
Procedure EnvoieToucheGrid_FO ( FF : TForm; CodeArt,DimSaisie: String ; TOBPiece:TOB) ;
Var FFact : TFFacture ;
    Key   : Word  ;
    i,NbDim,Ligne : Integer ;
    TOBL: TOB ;
    TypeDim: String ;
BEGIN
if Not (ctxMode in V_PGI.PGIContexte) then Exit ;
NbDim:=0 ;
Key:=0 ;
FFact:=TFFacture(FF) ;
if ctxMode in V_PGI.PGIContexte then
   BEGIN
   if FFact.GS.Col=SG_RefArt then
   begin
     if FFact.GS.Cells[SG_Lib,FFact.GS.Row]='' then exit ;
     Ligne:=FFact.GS.Row ;
     TOBL := TOBPiece.Detail[Ligne-1];
     if TOBL <> nil then  TypeDim := TOBL.GetValue('GL_TYPEDIM');
    if (TypeDim='NOR') then
      begin
      if SG_QF<>-1 then FFact.GS.Col:=SG_Qf
       else if SG_QS<>-1 then FFact.GS.Col:=SG_QS
        else Key:=VK_TAB;
      end
     else if (DimSaisie='GEN') then Key:=VK_DOWN
     else // Si on affiche les lignes des articles dimensionnés
       Begin
         if DimSaisie='DIM' then NbDim:=1
            else NbDim:=1 ; // 1 au lieu de 0, car la ligne de commentaire n'a pas de code article de renseigné
            i:=Ligne ;
            if i=TOBPiece.Detail.Count then exit ;
            TOBL:=TOBPiece.Detail[i] ;
            While TOBL.GetValue('GL_TYPEDIM')='DIM' do
              Begin
              if TOBL.GetValue('GL_CODEARTICLE')=CodeArt then NbDim:=NbDim+1 ;
              i:=i+1 ;
              if i=TOBPiece.Detail.Count then Break ;
              TOBL:=TOBPiece.Detail[i] ;
              end ;
            end;
         For i:=1 to NbDim-1 do
            begin
            Key:=VK_DOWN ;
            SendMessage(FFact.GS.Handle, WM_KeyDown, Key, 0);
            end ;
       end ;
   end else Key:=VK_TAB ;
//   SendMessage(FFact.GS.Handle, WM_KeyDown, Key, 0);
END ;

Procedure AfficheDetailDim ( Canvas : TCanvas ; DimSaisie : String ; ACol : integer ) ;
BEGIN
// DEBUT Modif MODE 31/07/2002
Canvas.Font.Style:=Canvas.Font.Style+[fsItalic] ;
if (ACol=SG_RefArt) and (DimSaisie<>'DIM') then CanVas.Font.Color:=Canvas.Brush.Color ;
// FIN Modif MODE 31/07/2002
END ;

Procedure AfficheGenDim ( Canvas : TCanvas ; DimSaisie : String ) ;
BEGIN
if ctxMode in V_PGI.PGIContexte then
   BEGIN
   END else
   BEGIN
   END ;
END ;

procedure SauveCoefFGOUV (TOBOuvrages : TOB);
var Indice : Integer;
		TOBOUV : TOB;
    IndPR,IndPA,IndCoefFG : integer;
begin
	for Indice := 0 to TOBOUvrages.Detail.count - 1 do
  begin
  	TOBOUV := TOBOUvrages.Detail[Indice];
    if Indice = 0 then
    begin
    	IndPR := TOBOUV.GetNumChamp ('BLO_DPR');
    	IndPA := TOBOUV.GetNumChamp ('BLO_DPA');
    	IndCoefFG := TOBOUV.GetNumChamp ('BLO_COEFFG');
    end;
    if TOBOUV.detail.count > 0 then BEGIN SauveCoefFGOUV (TOBOUV); Continue; END;
    If (TOBOUV.getValeur(IndPr) <> TOBOUV.getValeur(indPa)) and (TOBOUV.getValeur(IndPa)<>0) then
    begin
    	TOBOUV.PutValeur(IndCoefFG, Arrondi((TOBOUV.getValeur(IndPR)/TOBOUV.getValeur(IndPa))-1,4));
    end;
  end;
end;

procedure SauveCoefFG (TOBPIece,TOBOuvrages : TOB);
var Indice : integer;
		TOBL : TOB;
    IndPR,IndPA,IndCoefFG,IndTypeLigne,IndTypeArticle,IndNaturePieceg : integer;
begin
	for Indice := 0 TO TOBPiece.detail.count -1 do
  begin
  	TOBL := TOBPiece.detail[indice];
    if Indice = 0 then
    begin
    	IndPR := TOBL.GetNumChamp ('GL_DPR');
    	IndPA := TOBL.GetNumChamp ('GL_DPA');
    	IndCoefFG := TOBL.GetNumChamp ('GL_COEFFG');
    	IndTypeLigne := TOBL.GetNumChamp ('GL_TYPELIGNE');
    	IndTypeArticle := TOBL.GetNumChamp ('GL_TYPEARTICLE');
    	IndNaturePieceg := TOBL.GetNumChamp ('GL_NATUREPIECEG');
    end;
    if (TOBL.GetValeur(indTypeLigne)='ART') or
       ((TOBL.GetValeur(indTypeLigne)='ARV') and (TOBL.GetValeur(IndNaturePieceg)='PBT')) then
    begin
    	if ((TOBL.GetValeur(IndTypeArticle)='OUV') or (TOBL.GetValeur(IndTypeArticle)='ARP')) and (TOBL.GetValue('GL_INDICENOMEN')<>0) then continue;
      if (TOBL.GetValeur(IndPR)<> TOBL.GetValeur(INDPA)) and (TOBL.GetValeur(INDPA)<>0) then
      begin
        TOBL.PutValeur (IndCoefFG,(TOBL.GetValeur(IndPR)/TOBL.GetValeur(INDPA)) -1);
      end;
    end;
  end;
  SauveCoefFGOUV (TOBOuvrages);
end;

procedure AffecteNewZonesLigne (TOBArticles,TOBPiece : TOB; DEV : Rdevise; SaisieTypeAvanc : boolean=false;NewCodeTva : string='');
var Indice : integer;
		TOBL : TOB;
    TOBA : TOB;
    IndTPsUnitaire,IndTypeRes,IndTypeLig,IndTypeArt,IndArt,IndFourn,IndCoefMarg,IndCoeffG,IndPvHtdev,IndCoefDev : integer;
    IndPvHt,IndPR,IndPrxbase,IndPourcentAvanc,IndTotalHtDev,IndMontantFact,IndMontantSit,IndTotHeure,IndQte,IndTva : integer;
    CoefFGLig,MontantFact : double;
begin

  for Indice := 0 to TOBPiece.detail.count -1 do
  begin
  	TOBL := TOBpiece.detail[Indice];
    if Indice = 0 then
    begin
      IndQte := TOBL.GetNumChamp('GL_QTEFACT');
      IndTotHeure := 	TOBL.GetNumChamp('GL_HEURE');
    	IndTPsUnitaire := TOBL.GetNumChamp ('GL_TPSUNITAIRE');
    	IndTypeRes := TOBL.GetNumChamp ('BNP_TYPERESSOURCE');
    	IndTypeLig := TOBL.GetNumChamp ('GL_TYPELIGNE');
    	IndTypeArt := TOBL.GetNumChamp ('GL_TYPEARTICLE');
    	IndArt := TOBL.GetNumChamp ('GL_ARTICLE');
    	IndFourn := TOBL.GetNumChamp ('GL_FOURNISSEUR');
    	IndCoefMarg := TOBL.GetNumChamp ('GL_COEFMARG');
    	IndCoeffG := TOBL.GetNumChamp ('GL_COEFFG');
    	IndPvHt := TOBL.GetNumChamp ('GL_PUHT');
    	IndPvHtDEV := TOBL.GetNumChamp ('GL_PUHTDEV');
    	IndPR := TOBL.GetNumChamp ('GL_DPR');
    	IndCoefDev := TOBL.GetNumChamp ('GL_TAUXDEV');
      IndPourcentAvanc := TOBL.GetNumChamp ('GL_POURCENTAVANC');
      IndTotalHtDev := TOBL.GetNumChamp ('GL_TOTALHTDEV');
      IndMontantFact := TOBL.GetNumChamp ('MONTANTFACT');
      IndMontantSit := TOBL.GetNumChamp ('MONTANTSIT');
      IndTVA := TOBL.GetNumChamp ('GL_FAMILLETAXE1');
    end;
    if TOBL.GetValeur (IndTypeLig) <> 'ART' Then continue;
    //
    if SaisieTypeAvanc then
    begin
      MontantFact := Arrondi((TOBL.GetValeur(IndPourcentAvanc)/100)* TOBL.GetValeur(IndTotalHtDev),DEV.Decimale );
      if IndMontantfact > 0 then
      begin
      	TOBL.PutValeur(IndMontantFact,MontantFact);
      end;
      if IndMontantSit > 0 then
      begin
      	TOBL.PutValeur(IndMontantSit,0);
      end;
    end;
    //
    if ((TOBL.GetValeur (IndTypeART) = 'OUV') or (TOBL.GetValeur (IndTypeART) = 'ARP')) and
    	 (TOBL.GetValue('GL_INDICENOMEN')<>0 ) Then continue;

    if (Pos (TOBL.GetValeur(IndTypeRes),VH_GC.BTTypeMoInterne) > 0) and (not IsLigneExternalise(TOBL)) then
    begin
      TOBL.PutValeur(IndTPsUnitaire,1);
    end else
    begin
      TOBL.PutValeur(IndTPsUnitaire,0);
    end;
    //
    TOBL.PutValeur(IndTotHeure ,Arrondi(TOBL.getValeur(IndTPsUnitaire)*TOBL.GetValeur(IndQte),V_PGI.okdecQ));
    TOBL.putvalue('GL_TOTALHEURE',TOBL.Getvalue('GL_HEURE'));

    if (TOBL.GetValeur(IndCoefDev) <> 0) then
    begin
      TOBL.putValeur (IndPvHt,TOBL.GetValeur(indPVHTDEV)/TOBL.GetValeur(IndCoefDev));
    end;
    if (TOBL.GetValeur(IndPR) <> 0) then
    begin
      TOBL.putValeur (IndCoefMarg,arrondi(TOBL.GetValeur(indPVHT)/TOBL.GetValeur(IndPR),4));
    end;

    //
    TOBA := TOBArticles.findFirst(['GA_ARTICLE'],[TOBL.GetValeur(IndArt)],true);
    if (TOBA <> nil) then
    begin
      TOBL.PutValeur(IndPrxbase,TOBA.getValue('GA_DPA'));
      if TOBL.GetValeur(IndFourn)='' then
      begin
        if TOBA.getValue('GA_FOURNPRINC')<>'' then
        begin
          TOBL.PutValeur(IndFourn,TOBA.getValue('GA_FOURNPRINC'));
        end;
      end;
    end;

    if NewCodeTva <> '' then
    begin
      TOBL.PutValeur(IndTva,NewCodeTva);
    end;
  end;
end;

procedure AffecteNewZonesOuv (TOBArticles,TOBOuvrages : TOB; NewCodeTva : string='');
var Indice : integer;
		TOBL,TOBA : TOB;
    IndTPsUnitaire,IndTypeRes,IndTypeLig,IndTypeArt,IndCoefMarg,IndFourn : integer;
    IndPuHt,IndPR,IndArt,IndPuHTdev,IndCoefDev,IndTva : integer;
    first : boolean;
    CoefFGLig : double;
begin
	first := true;
  for Indice := 0 to TOBOuvrages.detail.count -1 do
  begin
  	TOBL := TOBOuvrages.detail[Indice];
    if TOBL.detail.count > 0 then BEGIN AffecteNewZonesOuv (TOBArticles,TOBL); continue; END;
    if first then
    begin
      IndFourn := TOBL.GetNumChamp ('BLO_FOURNISSEUR');
    	IndTPsUnitaire := TOBL.GetNumChamp ('BLO_TPSUNITAIRE');
    	IndTypeRes := TOBL.GetNumChamp ('BNP_TYPERESSOURCE');
    	IndTypeLig := TOBL.GetNumChamp ('BLO_TYPELIGNE');
    	IndTypeArt := TOBL.GetNumChamp ('BLO_TYPEARTICLE');
    	IndCoefMarg := TOBL.GetNumChamp ('BLO_COEFMARG');
      IndPuHt := TOBL.GetNumChamp ('BLO_PUHT');
      IndPuHtDEV := TOBL.GetNumChamp ('BLO_PUHTDEV');
      IndCoefDEV := TOBL.GetNumChamp ('BLO_TAUXDEV');
      IndPR := TOBL.GetNumChamp ('BLO_DPR');
      IndArt := TOBL.GetNumChamp ('BLO_ARTICLE');
      IndTVa := TOBL.GetNumChamp ('BLO_FAMILLETAXE1');
      first := false;
    end;
    if TOBL.GetValeur (IndTypeLig) <> 'ART' Then continue;
//    if (TOBL.GetValeur (IndTypeART) = 'OUV') or (TOBL.GetValeur (IndTypeART) = 'ARP') Then continue;

    // Tps de Mo Interne
    if (Pos (TOBL.GetValeur(IndTypeRes),VH_GC.BTTypeMoInterne) > 0) and (not IsLigneExternalise(TOBL)) then
    begin
      TOBL.PutValeur(IndTPsUnitaire,1);
    end else
    begin
      TOBL.PutValeur(IndTPsUnitaire,0);
    end;
    if (TOBL.GetValeur(indCoefDev) <> 0) then
    begin
      TOBL.putValeur (IndPuHt,Arrondi(TOBL.GetValeur(IndPuHtDev)/TOBL.GetValeur(indCoefDev),V_PGI.okDecP));
    end;

    if (TOBL.GetValeur(IndPR) <> 0) then
    begin
      TOBL.putValeur (IndCoefMarg,Arrondi(TOBL.GetValeur(IndPuHt)/TOBL.GetValeur(indPR),4));
    end;

    if TOBL.GetValeur(IndFourn)='' then
    begin
    	TOBA := TOBArticles.findFirst(['GA_ARTICLE'],[TOBL.GetValeur(IndArt)],true);
      if (TOBA <> nil) then
      begin
      	if TOBA.getValue('GA_FOURNPRINC')<>'' then
        begin
          TOBL.PutValeur(IndFourn,TOBA.getValue('GA_FOURNPRINC'));
        end;
      end;
    end;
    //
    if NewCodeTva <> '' then
    begin
      TOBL.PutValeur(IndTva,NewCodeTva);
    end;
  end;
end;


procedure ControleApplicFraisSurStDetailOuvrage (TOBOuvrage : TOB);
var Indice : integer;
    TOBOUv : TOB;
    IndpA, IndPr , IndAPPLICFRAIS,IndApplicFC : integer;
begin
  // Pas d'application des frais sur ST si mtpa=mtpr et Frais generaux existant

  for Indice := 0 to TOBOUvrage.detail.count -1 do
  begin
    TOBOUv := TOBOuvrage.detail[Indice];
    if Indice = 0 then
    begin
      IndPR := TOBOUV.GetNumChamp ('BLO_DPR');
      IndPA := TOBOUV.GetNumChamp ('BLO_DPA');
      IndAPPLICFRAIS := TOBOUV.GetNumChamp ('BLO_NONAPPLICFRAIS');
      IndApplicFC := TOBOUV.GetNumChamp ('BLO_NONAPPLICFC');
    end;
    if IsprestationST(TOBOUV) and (TOBOUV.getValeur(IndPr) = TOBOUV.getValeur(indPa)) then
    begin
      TOBOUV.putValeur(IndAPPLICFRAIS,'X');
      TOBOUV.putValeur(IndApplicFC,'X');
    end;
    if TOBOUV.detail.count > 0 then
    begin
      ControleApplicFraisSurStDetailOuvrage (TOBOUV);
    end;
  end;
end;

procedure ControleApplicFraisSurStOuvrages (TobOuvrages : TOB);
var Indice : integer;
    TOBOUv : TOB;
begin
  for Indice := 0 to TOBOuvrages.detail.count -1 do
  begin
    TOBOUV := TOBOuvrages.detail[Indice];
    ControleApplicFraisSurStDetailOuvrage (TOBOUv);
  end;
end;

procedure ControleApplicFraisSurStLigne (TobPiece : TOB);
var Indice : integer;
    TOBL : TOB;
    IndpA, IndPr , IndAPPLICFRAIS,IndApplicFC : integer;
begin
  for Indice := 0 to TOBPiece.detail.count -1 do
  begin
    TOBL := TOBPiece.detail[Indice];
    if Indice = 0 then
    begin
      IndPR := TOBL.GetNumChamp ('GL_DPR');
      IndPA := TOBL.GetNumChamp ('GL_DPA');
      IndAPPLICFRAIS := TOBL.GetNumChamp ('GLC_NONAPPLICFRAIS');
      IndApplicFC := TOBL.GetNumChamp ('GLC_NONAPPLICFC');
    end;
    if IsprestationST(TOBL) and (TOBL.getValeur(IndPr) = TOBL.getValeur(indPa)) then
    begin
      TOBL.putValeur(IndAPPLICFRAIS,'X');
      TOBL.putValeur(IndApplicFC,'X');
    end;
  end;
end;

procedure BeforeTraitementCalculCoef (TOBArticles,TOBPiece,TOBPorcs,TOBOUvrages : TOB; DEV : Rdevise; SaisieTypeAvanc : boolean=false; NewCodeTva : string = '');
var MtFC,CoefFG : double;
		ExistFg : boolean;
begin
	Coeffg := 0;
	MtFc := GetMontantFraisDetail (TOBPiece,ExistFg);
  if TOBpiece.getValue('GP_NATUREPIECEG')<>'PBT' then
  begin
  	CoefFg := GetTauxFg (TOBPorcs);
  end;
  if (MtFc=0) and (CoefFg=0) then
  begin
    SauveCoefFG (TOBPiece,TOBOUvrages); // et pas WILLY
  end else
  begin
    // Pas d'application des frais sur ST si mtpa=mtpr et Frais generaux existant
    // correction des erreurs de la Version 2007 Edition < 9
    // Il existait dees lignes d'ouvrages ayant PA=PR et des frais existant sur document --> Anomalies
    ControleApplicFraisSurStLigne (TobPiece);
    ControleApplicFraisSurStOuvrages (TOBOuvrages);
  end;
  AffecteNewZonesLigne (TOBArticles,TOBPiece,DEV,SaisieTypeAvanc,NewCodeTva);
  AffecteNewZonesOuv (TOBArticles,TOBOuvrages,newCodeTva);
end;

procedure InitLigneValoPR (TOBL : TOB);
begin
//  TOBL.PutValue('GL_COEFFG',0) ;
  TOBL.PutValue('GL_COEFFC',0) ;
  TOBL.PutValue('GL_COEFFR',0) ;
  //
  TOBL.PutValue('GL_MONTANTPAFG',0) ;
  TOBL.PutValue('GL_MONTANTPAFC',0) ;
  TOBL.PutValue('GL_MONTANTPAFR',0) ;
  //
  TOBL.PutValue('GL_MONTANTFG',0) ;
  TOBL.PutValue('GL_MONTANTFC',0) ;
  TOBL.PutValue('GL_MONTANTFR',0) ;
  //
  TOBL.PutValue('GL_MONTANTPA',0) ;
  TOBL.PutValue('GL_MONTANTPR',0) ;
end;


end.
