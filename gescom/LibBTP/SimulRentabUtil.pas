unit SimulRentabUtil;


interface

uses Classes,UTOB,FactComm,Hctrls,HEnt1,SaisUtil,UtilPGi,
{$IFDEF EAGLCLIENT}
  maineagl,
{$ELSE}
  Doc_Parser,DBCtrls, Db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} fe_main,
{$ENDIF}
 DicoBTP,NomenUtil,FactArticle,FactUtil,FactVariante,SysUtils,BTPUtil,UentCommun;
Type T_SimulChg = (tscHrs,tscQte,tscPAH,tscPAQ,tscPA,tscCoefRenH,TscCoefRenQ,TscCoefRen,TscCoefFG,TscCoefFC,TscPRH,TscPRQ,TscPR,TscCoefMAH,TscCoefMAQ,TscCoefMA,TscPVH,TscPVTTCH,TscPVQ,TscPVTTCQ,TscPV,TscPua,TscPuR,TscPUV,TscPVTTC,TscPUVTTC) ;
		 TGestBlockage = (TgbFraisDetaille,TbgBordereau,TgbFRFixed,TgbFCFixed);
     TSGestionInterdit = set of TGestBlockage;

procedure AppliqueChangement (Change: T_SimulChg;CoefChange:double;TOBSimulPres:TOB;var ListChange : Tlist;ValeurFixe : double=0; GestionInterdit : TSGestionInterdit=[]; Global : boolean=false);
procedure AddchampsupTSimul (TOBF : TOB);
procedure AfficheDetailSimul (GS : THGrid;TOBSimulpres : TOB;NiveauMax : integer;ListeSaisie:String);
procedure AfficheLaGrille (GS : THGrid;ListeSaisie: string;TobSimulPres : TOB;NiveauMax,NbElt : integer;AvecParagraphe : boolean);
//procedure ChangeValue(Change: T_SimulChg;AncValeur,NewValeur:double;TOBSimulPres:TOB; PasTouchePv : boolean);
procedure ChangeValue(Change: T_SimulChg;AncValeur,NewValeur:double;TOBSimulPres:TOB; GestionInterdit : TSGestionInterdit; Global : Boolean=false);
procedure CreationTOBSimul(TobSimulPres,TOBPIECE,TOBOuvrage,TOBArticles : TOB; AvecParagraphe : boolean; GestionInterdit : TSGestionInterdit; AvecCotraintance : Boolean);
function CreationTSImul (Pere : TOB): TOB;
procedure DefiniOuvert (TOBL : TOB);
procedure FermeBranche (GS:THGrid; TobSimulPres: TOB;var Arow:integer;var Niveau : integer; var Found : boolean);
procedure FinaliseCalcul(TobSimulPres : TOB);
procedure FinaliseSimulation (TOBSImulPres : TOB; var Valor : T_Valeurs);
procedure IndiceGrilleInit (TOBSimulPres : TOB; NiveauMax : integer; var Numligne : integer);
function GetMontantAchatLIGNE(TOBL: TOB): double;
function GetMontantRevientLIGNE(TOBL: TOB): double;
function GetMontantVenteLIGNE(TOBL: TOB): double;
function GetMontantVenteTTCLIGNE (TOBL : TOB) : double;
function GetTOBLigne (TOBSimulPres : TOB; Arow : integer): TOB;
procedure GetValorisation (TOBS : TOB; Var Valeur : T_Valeurs; UniquementAchat : boolean=false);
procedure InitBranche (GS : THGrid; TOBSimulPres : TOB; Arow : integer);
procedure OuvreBranche (GS : THGrid; TobSimulPres: TOB;var Arow:integer;var niveau : integer; var Found : boolean; Force : boolean);
function RecupTypeGraph (TOBS : TOB) : integer;
//procedure ReinitMontantLigne (TOBS : TOB; PastouchePV : Boolean);
procedure ReinitMontantLigne (TOBS : TOB; GestionInterdit : TSGestionInterdit);
//procedure ReinitValeur (Change: T_SimulChg;TobSimulPres : TOB; PastouchePV : Boolean);
procedure ReinitValeur (Change: T_SimulChg;TobSimulPres : TOB; GestionInterdit : TSGestionInterdit);
PROCEDURE RemplitTOBSimul (TobSimulPres,TOBA,TOBL : TOB;Paragraphe,Categorie,LibCategorie,Nature,Libnature : string;prestation:boolean;Valeur :T_Valeurs ;AvecParagraphe:boolean; gestionInterdit : TSGestionInterdit ; FromBordereau : boolean; TOBOUV:TOB=nil; LibelleForce : string='');
procedure Renumerote (TOBSimulPres : TOB; var Ligne : integer; var niveau : integer);
procedure TraiteElementLig (TobSimulPres,TOBL,TOBA: TOB; Paragraphe : string;Valeur : T_Valeurs;AvecParagraphe : boolean; GestionInterdit : TSGestionInterdit; FromBordereau : Boolean; TOBOUv : TOB=nil; LIbelleForce : string='');
procedure TraiteOuvrageLIG(TobSimulPres,TOBOUvrage,TOBArticles,TOBL: TOB; Paragraphe: string; AvecParagraphe : boolean;gestionInterdit : TSGestionInterdit; FromBordereau : boolean);
Procedure TraiteDetailOuv (AvecParagraphe: boolean;Paragraphe : string;TOBArticles,TobSimulPres,TOBL,TOBOuv : TOB;Quantite,QteDuDetail : double;gestionInterdit : TSGestionInterdit; FromBordereau : boolean; LibellePrixPose : string='');
procedure SetValorisation (TOBS : TOB; Valor:T_Valeurs; UniquementAchat : boolean=false);
procedure ValorisationFinale (TOBS : TOB; UniquementSommeAchat : boolean=false);
procedure TraiteModification (TOBSimulPres,TOBPiece,TOBOuvrage,TOBMOD : TOB;DEV : Rdevise; GestionInterdit : TSGestionInterdit);
procedure AddChampsFrais (TOBPieceFrais : TOB);
procedure AppliqueCoefFC (TobSimulPres : TOB ; NewCoefFC : double);
procedure CalculePAGlobal (TOBSImulPres : TOB; var Valor : T_Valeurs);
procedure RecalcSuiteModif (TOBPiece,TOBOUvrage,TOBMod,TOBBases,TOBBasesL : TOB; DEV : Rdevise );
procedure SetArrondi (TOBS : TOB);
procedure RecalculeOuvrage (TOBPiece,TOBL,TOBOuvrage,TOBBases,TOBBasesL : TOB; DEV : Rdevise);

const	SimulMessage: array[1..2] of string 	= (
          {1}        'A défaut de pouvoir appliquer cette valeur#13#10La plus proche sera appliquée',
          {2}        '');

implementation
uses Paramsoc,FactOuvrage,FactureBtp,FactPiece;

procedure AddChampsFrais (TOBPieceFrais : TOB);
begin
	TOBPieceFrais.AddChampSupValeur ('MONTANTFC',0,false);
	TOBPieceFrais.AddChampSupValeur ('COEFFR',0,false);
end;

procedure AppliqueCoefFC (TobSimulPres : TOB ; NewCoefFC : double);
var Indice : integer;
		TOBS : TOB;
    MtCharge : double;
begin
  for Indice := 0 to TOBSimulPres.detail.count -1 do
  begin
    TOBS := TOBSimulPres.detail[Indice];
    // on fait le changement au plus fin
    if TOBS.GetValue('NIVEAU') <> 5 then
    begin
      AppliqueCoefFC(TOBS,NewCoefFc);
    end else // Niveau = 5
    begin
    	if TOBS.GetValue('TYPE')='H' then
      begin
      	if TOBS.GetValue('MONTANTPAFGH') <> 0 then
        begin
      		TOBS.putValue('COEFFG',(TOBS.GetValue('MONTANTPAFGH')+TOBS.GetValue('MONTANTFGH'))/TOBS.GetValue('MONTANTPAFGH'));
        end else
        begin
        	TOBS.putValue('COEFFG',1);
        end;
      	// Calcul du nouveau montant de Frais de chantier
      	if TOBS.GetValue('APPLICCOEFFC')='X' Then
        begin
        	MtCharge := TOBS.GetValue('MONTANTPAFGH')+TOBS.GetValue('MONTANTFGH');
          TOBS.putValue('MONTANTFCH',MtCharge*NewCoefFC);
          if MtCharge <> 0 then TOBS.putValue('COEFFCH',(MtCharge+TOBS.getValue('MONTANTFCH'))/MtCharge)
          								 else TOBS.putValue('COEFFCH',1);
        end;
      	// Calcul du nouveau montant de Frais repartis
      	if TOBS.GetValue('APPLICCOEFFR')='X' Then
        begin
        	MtCharge := TOBS.GetValue('MONTANTPAFGH')+TOBS.GetValue('MONTANTFGH')+TOBS.GetValue('MONTANTFCH');
          TOBS.putValue('MONTANTFRH',(MtCharge*TOBS.GetValue('COEFFRH')) - mTCharge);
        end;
      	// Calcul du nouveau montant de revient
        TOBS.putValue('MONTANTPRH',TOBS.GetValue('MONTANTPAFGH')+TOBS.GetValue('MONTANTFGH')+TOBS.GetValue('MONTANTFCH')+TOBS.GetValue('MONTANTFRH'));
        if TOBS.GetValue('BLOQUEPV')<>'X' then
        begin
      		// Calcul du nouveau montant de vente
          TOBS.PutValue('MONTANTPVH',TOBS.GetValue('MONTANTPRH')*TOBS.GetValue('COEFMARGH'));
          TOBS.PutValue('MONTANTPVTTCH',TOBS.GetValue('MONTANTPVH')*TOBS.GetValue('TAUXTVAH'));
        end;
      end else if TOBS.GetValue('TYPE')='F' then
      begin
      	// Calcul du nouveau montant de Frais de chantier
      	if TOBS.GetValue('APPLICCOEFFC')='X' Then
        begin
        	MtCharge := TOBS.GetValue('MONTANTPAFGQ')+TOBS.GetValue('MONTANTFGQ');
          TOBS.putValue('MONTANTFCQ',MtCharge*NewCoefFC);
          if MtCharge <> 0 then TOBS.putValue('COEFFCQ',(MtCharge+TOBS.getValue('MONTANTFCQ'))/MtCharge)
          								 else TOBS.putValue('COEFFCQ',0);
        end;
      	// Calcul du nouveau montant de Frais repartis
      	if TOBS.GetValue('APPLICCOEFFR')='X' Then
        begin
        	MtCharge := TOBS.GetValue('MONTANTPAFGQ')+TOBS.GetValue('MONTANTFGQ')+TOBS.GetValue('MONTANTFCQ');
          TOBS.putValue('MONTANTFRQ',(MtCharge*TOBS.GetValue('COEFFRQ')) - MtCharge);
        end;
      	// Calcul du nouveau montant de revient
        TOBS.putValue('MONTANTPRQ',TOBS.GetValue('MONTANTPAFGQ')+TOBS.GetValue('MONTANTFGQ')+TOBS.GetValue('MONTANTFCQ')+TOBS.GetValue('MONTANTFRQ'));
        if TOBS.GetValue('BLOQUEPV')<>'X' then
        begin
      		// Calcul du nouveau montant de vente
          TOBS.PutValue('MONTANTPVQ',TOBS.GetValue('MONTANTPRQ')*TOBS.GetValue('COEFMARGQ'));
          TOBS.PutValue('MONTANTPVTTCQ',TOBS.GetValue('MONTANTPVQ')*TOBS.GetValue('TAUXTVAQ'));
        end;
      end;
    end;
  end;
end;

procedure CalculePAGlobal (TOBSImulPres : TOB; var Valor : T_Valeurs);
var Indice : Integer;
    TOBS : TOB;
    valeur : T_Valeurs;
begin
For Indice := 0 to TOBSimulPres.detail.count -1 do
    begin
    TOBS := TOBSimulPres.detail[Indice];
    if TOBS.detail.count > 0 then
       begin
       InitTableau (Valeur);
       CalculePAGlobal (TOBS,Valeur);
       Valor := CalculSurTableau ('+',valor,valeur);
       end else
       begin
       ValorisationFinale (TOBS,True);
       InitTableau (Valeur);
       GetValorisation (TOBS,Valeur,True);
       Valor := CalculSurTableau ('+',valor,valeur);
       end;
    end;
SetValorisation (TOBSimulPres,Valor,True);
ValorisationFinale (TOBSimulPres,True);
end;

procedure RecalculeMontantsPA (TOBS : TOB);
begin
  if TOBS.GetValue('TYPE') = 'H' then
  BEGIN
    TOBS.putValue('MONTANTPA',TOBS.GetValue('MONTANTPAH'));
    TOBS.putValue('MONTANTPAFGH',TOBS.GetValue('MONTANTPAH'));
    TOBS.putValue('MONTANTPAFCH',TOBS.GetValue('MONTANTPAH'));
    TOBS.putValue('MONTANTPAFRH',TOBS.GetValue('MONTANTPAH'));
    TOBS.putValue('MONTANTFGH',(TOBS.GetValue('MONTANTPAFGH')*TOBS.GetValue('COEFFGH')) - (TOBS.GetValue('MONTANTPAFGH'))) ;
  END ELSE
  BEGIN
    TOBS.putValue('MONTANTPA',TOBS.GetValue('MONTANTPAQ'));
    TOBS.putValue('MONTANTPAFGQ',TOBS.GetValue('MONTANTPAQ'));
    TOBS.putValue('MONTANTPAFCQ',TOBS.GetValue('MONTANTPAQ'));
    TOBS.putValue('MONTANTPAFRQ',TOBS.GetValue('MONTANTPAQ'));
    TOBS.putValue('MONTANTFGQ',(TOBS.GetValue('MONTANTPAFGQ')*TOBS.GetValue('COEFFGQ')) - (TOBS.GetValue('MONTANTPAFGQ')));
  END;
end;

procedure AddchampsupTSimul (TOBF : TOB);
begin
TOBF.AddChampSupValeur ('NUMAFF',0,false);
TOBF.AddChampSupValeur ('NIVEAU',0,false);
TOBF.AddChampSupValeur ('PARAGRA','',false);
TOBF.AddChampSupValeur ('LIBPARAGRA','',false);
TOBF.AddChampSupValeur ('CATEGORIE','',false);
TOBF.AddChampSupValeur ('LIBCATEGORIE','',false);
TOBF.AddChampSupValeur ('NATURE','',false);
TOBF.AddChampSupValeur ('LIBNATURE','',false);
TOBF.AddChampSupValeur ('ARTICLE','',false);
TOBF.AddChampSupValeur ('LIBELLE','',false);
//
TOBF.AddChampSupValeur ('NUMLIGNE',0,false);
//TOBF.AddChampSupValeur ('BLO_NIVEAU',0,false);
TOBF.AddChampSupValeur ('BLO_N1',0,false);
TOBF.AddChampSupValeur ('BLO_N2',0,false);
TOBF.AddChampSupValeur ('BLO_N3',0,false);
TOBF.AddChampSupValeur ('BLO_N4',0,false);
TOBF.AddChampSupValeur ('BLO_N5',0,false);

TOBF.AddChampSupValeur ('BLO_NUMORDRE',0,false);
//
TOBF.AddChampSupValeur ('HEURES',0,false);
TOBF.AddChampSupValeur ('QTE',0,false);
//
TOBF.AddChampSupValeur ('MONTANTPAH',0,false);
TOBF.AddChampSupValeur ('PAMOYENH',0,false);
TOBF.AddChampSupValeur ('COEFFRH',0,false);
TOBF.AddChampSupValeur ('MONTANTPRH',0,false);
TOBF.AddChampSupValeur ('PRMOYENH',0,false);
TOBF.AddChampSupValeur ('COEFMARGH',0,false);
TOBF.AddChampSupValeur ('MONTANTPVH',0,false);
TOBF.AddChampSupValeur ('PVMOYENH',0,false);
TOBF.AddChampSupValeur ('MONTANTPVTTCH',0,false);
TOBF.AddChampSupValeur ('PVMOYENTTCH',0,false);
TOBF.AddChampsupValeur ('TAUXTVAH',0,false);
TOBF.AddChampsupValeur ('MONTANTFGH',0,false);
TOBF.AddChampsupValeur ('MONTANTFCH',0,false);
TOBF.AddChampsupValeur ('MONTANTFRH',0,false);
TOBF.AddChampsupValeur ('MONTANTPAFGH',0,false);
TOBF.AddChampsupValeur ('MONTANTPAFCH',0,false);
TOBF.AddChampsupValeur ('MONTANTPAFRH',0,false);
TOBF.AddChampSupValeur ('COEFFCH',0,false);
TOBF.AddChampSupValeur ('COEFFGH',0,false);
//
TOBF.AddChampSupValeur ('MONTANTPAQ',0,false);
TOBF.AddChampSupValeur ('PAMOYENQ',0,false);
TOBF.AddChampSupValeur ('COEFFRQ',0,false);
TOBF.AddChampSupValeur ('MONTANTPRQ',0,false);
TOBF.AddChampSupValeur ('PRMOYENQ',0,false);
TOBF.AddChampSupValeur ('COEFMARGQ',0,false);
TOBF.AddChampSupValeur ('MONTANTPVQ',0,false);
TOBF.AddChampSupValeur ('PVMOYENQ',0,false);
TOBF.AddChampSupValeur ('MONTANTPVTTCQ',0,false);
TOBF.AddChampSupValeur ('PVMOYENTTCQ',0,false);
TOBF.AddChampsupValeur ('TAUXTVAQ',0,false);
TOBF.AddChampsupValeur ('MONTANTFGQ',0,false);
TOBF.AddChampsupValeur ('MONTANTFCQ',0,false);
TOBF.AddChampsupValeur ('MONTANTFRQ',0,false);
TOBF.AddChampsupValeur ('MONTANTPAFGQ',0,false);
TOBF.AddChampsupValeur ('MONTANTPAFCQ',0,false);
TOBF.AddChampsupValeur ('MONTANTPAFRQ',0,false);
TOBF.AddChampSupValeur ('COEFFCQ',0,false);
TOBF.AddChampSupValeur ('COEFFGQ',0,false);

//
TOBF.AddChampSupValeur ('MONTANTPA',0,false);
TOBF.AddChampSupValeur ('PAMOYEN',0,false);
TOBF.AddChampSupValeur ('COEFFR',0,false);
TOBF.AddChampSupValeur ('MONTANTPR',0,false);
TOBF.AddChampSupValeur ('PRMOYEN',0,false);
TOBF.AddChampSupValeur ('COEFMARG',0,false);
TOBF.AddChampSupValeur ('MONTANTPV',0,false);
TOBF.AddChampSupValeur ('MONTANTPVTTC',0,false);
TOBF.AddChampSupValeur ('PVMOYEN',0,false);
TOBF.AddChampSupValeur ('PVMOYENTTC',0,false);
TOBF.AddChampsupValeur ('OPEN','-',false);
TOBF.AddChampsupValeur ('TYPE','',false);
TOBF.AddChampsupValeur ('MODIFIE','-',false);
TOBF.AddChampsupValeur ('COEFMAMODIFIE','-',false);
//
TOBF.AddChampsupValeur ('RATIOQTE',0,false);
TOBF.AddChampsupValeur ('TAUXTVA',0,false);
TOBF.AddChampsupValeur ('MONTANTFG',0,false);
TOBF.AddChampsupValeur ('MONTANTFC',0,false);
TOBF.AddChampsupValeur ('MONTANTFR',0,false);
TOBF.AddChampsupValeur ('MONTANTPAFG',0,false);
TOBF.AddChampsupValeur ('MONTANTPAFC',0,false);
TOBF.AddChampsupValeur ('MONTANTPAFR',0,false);
TOBF.AddChampSupValeur ('COEFFC',0,false);
TOBF.AddChampSupValeur ('COEFFG',0,false);
//
TOBF.AddChampsupValeur ('BLOQUEPR','-',false);
TOBF.AddChampsupValeur ('BLOQUEPV','-',false);
TOBF.AddChampsupValeur ('APPLICCOEFFC','-',false);
TOBF.AddChampsupValeur ('APPLICCOEFFR','-',false);
end;

function CreationTSImul (Pere : TOB): TOB;
var TOBF : TOB;
begin
TOBF := TOB.Create ('SIMULLIG',Pere,-1);
AddchampsupTSimul (TOBF);
result := TOBF;
end;

procedure CreationTOBSimul(TobSimulPres,TOBPIECE,TOBOuvrage,TOBArticles : TOB; AvecParagraphe : boolean; GestionInterdit : TSGestionInterdit; AvecCotraintance : boolean);
var Indice      : integer;
    TOBL        : TOB;
    TOBA        : TOB;
    QQ          : Tquery;
    Paragraphe  : string;
    Valor       : T_valeurs;
    prevision   : boolean;
    INatJob     : Integer;
begin

  prevision := (TOBPiece.getValue('GP_NATUREPIECEG')='PBT');

  for Indice := 0 to TOBPiece.detail.count -1 do
  begin
    TOBL := TOBPiece.detail[Indice];
    //
    if indice = 0 then INatJob := TobL.GetNumChamp('GLC_NATURETRAVAIL');

    if (TOBL.getvalue('GL_TYPELIGNE') <> 'ART') and ((not IsArticleVariante(TOBL)) or (not prevision)) and
       (Not IsParagrapheStd(TOBL,1)) and ((not IsParagrapheVariante(TOBL,1)) or (not prevision))
       then continue;

    if AvecParagraphe then
    begin
      if IsDebutParagraphe(TOBL,1) then
      BEGIN
        Paragraphe := TOBL.GetValue('GL_LIBELLE');
        continue;
      END;
      if IsFinParagraphe(TOBL,1) then
      BEGIN
        Paragraphe := '';
        continue;
      END;
    end else if (IsParagraphe (TOBL,1)) then continue;

    TOBA := FindTOBArtSais (TobArticles,TOBL.GetValue('GL_ARTICLE'));
    if TOBA = nil then
    begin
      TOBA := TOB.create('ARTICLE',TobArticles,-1);
      QQ := OpenSQL ('SELECT * FROM ARTICLE WHERE GA_ARTICLE="'+TOBL.GetValue('GL_ARTICLE')+'"',true,-1,'',true);
      if not QQ.eof then TOBA.selectDb ('',QQ);
      ferme (QQ);
    end;

    //si ligne cotraitant et pas de prise en compte de la cotraitance dans l'analyse...
    if ((TOBL.GetValeur (INatJob) = '001') AND  (not AvecCotraintance)) then
      Continue
    else
    begin
      if (TOBL.GetValue('GL_TYPEARTICLE') = 'PRE') or  (TOBL.GetValue('GL_TYPEARTICLE') = 'MAR') or
         (TOBL.GetValue('GL_TYPEARTICLE') = 'FRA') or
        ((TOBL.GetValue('GL_TYPEARTICLE') = 'ARP') and (TOBL.GetValue('GL_INDICENOMEN')=0))	then
      begin
        Valor[0]  := GetMontantAchatLIGNE (TOBL);
        Valor[1]  := GetMontantRevientLIGNE (TOBL);
        Valor[2]  := GetMontantVenteLIGNE (TOBL);
        Valor[3]  := GetMontantVenteTTCLIGNE (TOBL);
        //
        Valor[6]  := TOBL.GetValue('GL_QTEFACT');
        valor[7]  := TOBL.GetValue('GL_DPA');
        valor[8]  := TOBL.GetValue('GL_DPR');
        valor[9]  := TOBL.GetValue('GL_PUHTDEV');
        valor[10] := TOBL.GetValue('GL_PUTTCDEV');
        valor[11] := 1;
        // NEW CHAMPS
        Valor[12] := TOBL.GetValue('GL_MONTANTFG');
        Valor[13] := TOBL.GetValue('GL_MONTANTFC');
        Valor[14] := TOBL.GetValue('GL_MONTANTFR');
        Valor[15] := TOBL.GetValue('GL_MONTANTPAFG');
        Valor[16] := TOBL.GetValue('GL_MONTANTPAFC');
        Valor[17] := TOBL.GetValue('GL_MONTANTPAFR');
        //
        TraiteElementLig (TobSimulPres,TOBL,TOBA,Paragraphe,Valor,AvecParagraphe,GestionInterdit,(TOBL.GetValue('GLC_FROMBORDEREAU')='X'),nil,'');
      end;

      if TOBL.GetValue('GL_TYPEARTICLE') = 'POU' then
      Begin
        Valor[0]  := 0;
        Valor[1]  := 0;
        Valor[2]  := GetMontantVenteLIGNE (TOBL);
        Valor[3]  := GetMontantVenteTTCLIGNE (TOBL);
        Valor[6]  := TOBL.GetValue('GL_QTEFACT');
        Valor[7]  := 0;
        Valor[8]  := 0;
        valor[9]  := TOBL.GetValue('GL_PUHTDEV');
        valor[10] := TOBL.GetValue('GL_PUTTCDEV');
        valor[11] := 1;
        // NEW CHAMPS
        Valor[12] := 0;
        Valor[13] := 0;
        Valor[14] := 0;
        Valor[15] := TOBL.GetValue('GL_MONTANTPAFG');
        Valor[16] := TOBL.GetValue('GL_MONTANTPAFC');
        Valor[17] := TOBL.GetValue('GL_MONTANTPAFR');
        //
        TraiteElementLig (TobSimulPres,TOBL,TOBA,Paragraphe,Valor,AvecParagraphe,gestionInterdit,(TOBL.GetValue('GLC_FROMBORDEREAU')='X'));
      End;

      if (TOBL.GetValue('GL_TYPEARTICLE') = 'OUV') or //(TOBL.GetValue('GL_TYPEARTICLE') = 'ARP') then
        ((TOBL.GetValue('GL_TYPEARTICLE') = 'ARP') and
         (TOBL.GetValue('GL_INDICENOMEN')>0) )	   then
        TraiteOuvrageLIG (TobSimulPres,TOBOuvrage,TOBArticles,TOBL,Paragraphe,AvecParagraphe,GestionInterdit,(TOBL.GetValue('GLC_FROMBORDEREAU')='X'));
    end;   
  end;

  InitTableau (valor);

  FinaliseSimulation (TOBSImulPres,valor);

end;

procedure TraiteElementLig (TobSimulPres,TOBL,TOBA: TOB; Paragraphe : string;Valeur : T_Valeurs;AvecParagraphe : boolean; GestionInterdit : TSGestionInterdit; FromBordereau : Boolean; TOBOUv : TOB=nil; LIbelleForce : string='');
var QQ : TQUERY;
    TOBPRest,TOBREF : TOB;
    Categorie, libcategorie,Nature,LibNature : string;
    Prestation : boolean;
    prefixe : string;
    BTNatPrestation : String;
begin

  if TOBOUV <> nil then
  BEGIN
    TOBREF := TOBOUV;
    Prefixe := 'BLO';
  END
  else
  BEGIN
    TOBREF := TOBL;
    Prefixe := 'GL';
  END;

  if TOBREF.getValue('GLC_NATURETRAVAIL') = '002' then
    BTNatPrestation := GetParamsocSecur('SO_BTNATPRESTATION', '')
  else
    BTNatPrestation := TOBA.GetValue('GA_NATUREPRES');

  if TOBREF.GetValue(prefixe+'_TYPEARTICLE') = 'PRE' then
  begin
    Prestation := true;
    TOBPREST := TOB.Create ('NATUREPREST',nil,-1);
    QQ := OpenSQL ('SELECT BNP_LIBELLE,BNP_TYPERESSOURCE FROM NATUREPREST WHERE BNP_NATUREPRES="'+BtNatPrestation+'"',true,-1,'',true);
    if not QQ.eof then
    begin
      TOBPrest.SelectDB ('',QQ);
      Categorie := TOBPRest.GetValue('BNP_TYPERESSOURCE');
      if categorie <> '' then
        LibCategorie := RechDom ('AFTTYPERESSOURCE',TOBPREST.GetValue('BNP_TYPERESSOURCE'),false)
      else
        LibCategorie := 'Non renseigné';
      Nature := BTNatPrestation;
      LibNature := TOBPrest.GetValue('BNP_LIBELLE');
    end
    else
    Begin
      LibNature := 'Non renseigné';
    End;
    ferme (QQ);
    TOBPRest.free;
  end
  else
  begin
    if TOBRef.getValue('GLC_NATURETRAVAIL') = '002' then
    begin
      Nature    := BTNatPrestation;
      LibCategorie := rechdom('BTNATPRESTATION',BTNatPrestation,false);
      LibNature := LibCategorie;
      Categorie := BTNatPrestation;
    end
    else
    begin
      if TOBREF.GetValue(prefixe+'_TYPEARTICLE') = 'FRA' then
      begin
        Categorie := 'FRA';
        LibCategorie := 'Frais';
      end else
      begin
        Categorie := 'MAR';
        LibCategorie := 'Fourniture';
      end;
      Nature := TOBREF.GetValue(prefixe+'_FAMILLENIV1');
      if Nature <> '' then
        LibNature := RechDom ('GCFAMILLENIV1',TOBREF.GetValue(Prefixe+'_FAMILLENIV1'),false)
      else
        LibNature := 'Non renseigné';
      Prestation := False;
    end;
  end;

  RemplitTOBSimul (TobSimulPres,TOBA,TOBL,Paragraphe,Categorie,LibCategorie,Nature,Libnature,prestation,Valeur,AvecParagraphe,GestionInterdit,FromBordereau,TOBOUV,LIbelleForce);
end;

PROCEDURE RemplitTOBSimul (TobSimulPres,TOBA,TOBL : TOB;Paragraphe,Categorie,LibCategorie,Nature,Libnature : string;prestation:boolean;Valeur :T_Valeurs ;AvecParagraphe:boolean; gestionInterdit : TSGestionInterdit ; FromBordereau : boolean; TOBOUV:TOB=nil; LibelleForce : string='');
var TOBP,TOBC,TOBN,TOBCODEArt,TOBLIGNE,TOBREF,TOBPere : TOB;
    Prefixe,Article, Libart : string;
    OKApplicFR,OKApplicFC: boolean;
    fRegroupLibellediff : boolean;
BEGIN
  fRegroupLibellediff := GetParamSocSecur('SO_GRPSIMULRENTAB',false);

  if TOBOUV <> nil then
  BEGIN
    TOBREF := TOBOUV;
    Prefixe:='BLO';
    OKApplicFR := (TOBRef.GetValue('BLO_NONAPPLICFRAIS') <> 'X');
    OKApplicFC := (TOBRef.GetValue('BLO_NONAPPLICFC') <> 'X');
  END ELSE
  BEGIN
    TOBREF := TOBL;
    Prefixe := 'GL';
    OKApplicFR := (TOBRef.GetValue('GLC_NONAPPLICFRAIS') <> 'X');
    OKApplicFC := (TOBRef.GetValue('GLC_NONAPPLICFC') <> 'X');
  END;


// Niveau Paragraphe
  if AvecParagraphe then
  begin
    TOBP := TOBSimulPres.findFirst (['PARAGRA','CATEGORIE','NATURE','ARTICLE'],[Paragraphe,'','',''],false);
    if TOBP = nil then
    BEGIN
      TOBP := CreationTSImul (TOBSimulPres);
      TOBP.PutValue('NIVEAU',1);
      TOBP.PutValue('PARAGRA',Paragraphe);
      TOBP.PutValue('LIBPARAGRA',Paragraphe);
    END;
    TobPere := TOBP;
  end else
  begin
  	TOBPere := TOBSimulPres;
  end;

// Niveau Categorie
  TOBC := TOBPere.findFirst (['PARAGRA','CATEGORIE','NATURE','ARTICLE'],[Paragraphe,Categorie,'',''],false);
  if TOBC = nil then
  BEGIN
    TOBC := CreationTSImul (TOBPere);
    TOBC.PutValue('NIVEAU',2);
    TOBC.PutValue('PARAGRA',Paragraphe);
    TOBC.PutValue('CATEGORIE',Categorie);
    TOBC.PutValue('LIBCATEGORIE',LibCategorie);
  END;
// Niveau Nature
  TOBN := TOBC.findFirst (['PARAGRA','CATEGORIE','NATURE','ARTICLE'],[Paragraphe,Categorie,nature,''],false);
  if TOBN = nil then
  BEGIN
    TOBN := CreationTSImul(TOBC);
    TOBN.PutValue('NIVEAU',3);
    TOBN.PutValue('PARAGRA',Paragraphe);
    TOBN.PutValue('CATEGORIE',Categorie);
    TOBN.PutValue('NATURE',Nature);
    TOBN.PutValue('LIBNATURE',libnature) ;
  END;
// Niveau Code Article
	Article := TOBREF.GetValue(Prefixe+'_CODEARTICLE');
	Article := trim(article);
	Libart := TOBREF.GetValue(Prefixe+'_LIBELLE');
	Libart := trim(Libart);
  if not fRegroupLibellediff then
  begin
	  TOBCODEArt := TOBN.findFirst (['PARAGRA','CATEGORIE','NATURE','ARTICLE','LIBELLE'],[Paragraphe,Categorie,Nature,Article,Libart],false);
  end else
  begin
	  TOBCODEArt := TOBN.findFirst (['PARAGRA','CATEGORIE','NATURE','ARTICLE'],[Paragraphe,Categorie,Nature,Article],false);
  end;
  if TOBCODEArt = nil then
  BEGIN
    TOBCODEArt := CreationTSImul(TOBN);
    TOBCODEArt.PutValue('NIVEAU',4);
    TOBCODEArt.PutValue('PARAGRA',Paragraphe);
    TOBCODEArt.PutValue('CATEGORIE',Categorie);
    TOBCODEArt.PutValue('NATURE',NATURE);
    TOBCODEArt.PutValue('ARTICLE',Article) ;
    if LibelleForce <> '' then TOBCODEArt.PutValue('LIBELLE',LibelleForce)
                          else TOBCODEArt.PutValue('LIBELLE',TOBREF.GetValue(PRefixe+'_LIBELLE')) ;
  END;
// Niveau Ligne de document
//if TOBOUV <> nil then
	if TOBRef <> nil then
  BEGIN
    TOBLIGNE := CreationTSImul(TOBCODEArt);
    TOBLIGNE.PutValue('NIVEAU',5);
    TOBLIGNE.PutValue('PARAGRA',Paragraphe);
    TOBLIGNE.PutValue('CATEGORIE',Categorie);
    TOBLIGNE.PutValue('NATURE',NATURE);
    TOBLIGNE.PutValue('ARTICLE',Article) ;
  	if LibelleForce <> '' then TOBCODEArt.PutValue('LIBELLE',LibelleForce)
  												else TOBCODEArt.PutValue('LIBELLE',TOBREF.GetValue(PRefixe+'_LIBELLE')) ;
    TOBLIGNE.PutValue('NUMLIGNE',TOBL.GetValue('GL_NUMLIGNE'));
    //      TOBLIGNE.PutValue('BLO_NIVEAU',TOBOUV.GetValue('BLO_NIVEAU'));
    if TOBOUV <> nil then
    begin
      TOBLIGNE.PutValue('BLO_N1',TOBOUV.GetValue('BLO_N1'));
      TOBLIGNE.PutValue('BLO_N2',TOBOUV.GetValue('BLO_N2'));
      TOBLIGNE.PutValue('BLO_N3',TOBOUV.GetValue('BLO_N3'));
      TOBLIGNE.PutValue('BLO_N4',TOBOUV.GetValue('BLO_N4'));
      TOBLIGNE.PutValue('BLO_N5',TOBOUV.GetValue('BLO_N5'));
      TOBLIGNE.PutValue('BLO_NUMORDRE',TOBOUV.GetValue('BLO_NUMORDRE'));
    end;
    if Prestation then
    BEGIN
      if (Categorie = 'SAL') or (Categorie='ST') or (Categorie = 'INT') then
      BEGIN
        TOBLIGNE.PutValue('TYPE','H'); // type Horaire
        TOBLIGNE.PutValue('HEURES',Valeur[6]);
        TOBLIGNE.PutValue('MONTANTPAH',Valeur[0]);
        TOBLIGNE.PutValue('MONTANTPRH',Valeur[1]);
        TOBLIGNE.PutValue('MONTANTPVH',Valeur[2]);
        TOBLIGNE.PutValue('MONTANTPVTTCH',Valeur[3]);
        TOBLIGNE.PutValue('PAMOYENH',Valeur[7]);
        TOBLIGNE.PutValue('PRMOYENH',Valeur[8]);
        TOBLIGNE.PutValue('PVMOYENH',Valeur[9]);
        TOBLIGNE.PutValue('PVMOYENTTCH',Valeur[10]);
        TOBLIGNE.PutValue('RATIOQTE',Valeur[11]);
        // NEW CHAMPS
        TOBLIGNE.PutValue('MONTANTFGH',Valeur[12]);
        TOBLIGNE.PutValue('MONTANTFCH',Valeur[13]);
        TOBLIGNE.PutValue('MONTANTFRH',Valeur[14]);
        TOBLIGNE.PutValue('MONTANTPAFGH',Valeur[15]);
        TOBLIGNE.PutValue('MONTANTPAFCH',Valeur[16]);
        TOBLIGNE.PutValue('MONTANTPAFRH',Valeur[17]);
      	//
      END ELSE
      BEGIN
        TOBLIGNE.PutValue('QTE',Valeur[6]);
        TOBLIGNE.PutValue('MONTANTPAQ',Valeur[0]);
        TOBLIGNE.PutValue('MONTANTPRQ',Valeur[1]);
        TOBLIGNE.PutValue('MONTANTPVQ',Valeur[2]);
        TOBLIGNE.PutValue('MONTANTPVTTCQ',Valeur[3]);
        TOBLIGNE.PutValue('PAMOYENQ',Valeur[7]);
        TOBLIGNE.PutValue('PRMOYENQ',Valeur[8]);
        TOBLIGNE.PutValue('PVMOYENQ',Valeur[9]);
        TOBLIGNE.PutValue('PVMOYENTTCQ',Valeur[10]);
        TOBLIGNE.PutValue('RATIOQTE',Valeur[11]);
        TOBLIGNE.PutValue('TYPE','F'); // type fourniture
        // NEW CHAMPS
        TOBLIGNE.PutValue('MONTANTFGQ',Valeur[12]);
        TOBLIGNE.PutValue('MONTANTFCQ',Valeur[13]);
        TOBLIGNE.PutValue('MONTANTFRQ',Valeur[14]);
        TOBLIGNE.PutValue('MONTANTPAFGQ',Valeur[15]);
        TOBLIGNE.PutValue('MONTANTPAFCQ',Valeur[16]);
        TOBLIGNE.PutValue('MONTANTPAFRQ',Valeur[17]);
      	//
      END;
      //
      if OKApplicFC then
      begin
      	TOBLIGNE.PutValue('APPLICCOEFFC','X');
      end;
      if OKApplicFR then
      begin
      	TOBLIGNE.PutValue('APPLICCOEFFR','X');
      end;
      //
      if ((TOBRef.GetValue(prefixe+'_COEFFR')<>0) and (OKApplicFR))   or
      	 ((TOBRef.GetValue(prefixe+'_COEFFC')<>0) and (OKApplicFC)) then
      begin
      	TOBLIGNE.PutValue ('BLOQUEPR','X');
      	TOBCODEArt.PutValue ('BLOQUEPR','X');
      	TOBN.PutValue ('BLOQUEPR','X');
      	TOBC.PutValue ('BLOQUEPR','X');
      	TOBPERE.PutValue ('BLOQUEPR','X');
      	TOBREF.PutValue ('BLOQUEPR','X');
      end;

			if ((TbgBordereau in GestionInterdit) and (FromBordereau)) or (TOBL.GetValue('GL_BLOQUETARIF')='X') then
      begin
      	TOBLIGNE.PutValue ('BLOQUEPV','X');
      	TOBCODEArt.PutValue ('BLOQUEPV','X');
      	TOBN.PutValue ('BLOQUEPV','X');
      	TOBC.PutValue ('BLOQUEPV','X');
      	TOBPERE.PutValue ('BLOQUEPV','X');
      	TOBREF.PutValue ('BLOQUEPV','X');
      end;

    END ELSE
    BEGIN
      TOBLIGNE.PutValue('QTE',Valeur[6]);
      TOBLIGNE.PutValue('MONTANTPAQ',Valeur[0]);
      TOBLIGNE.PutValue('MONTANTPRQ',Valeur[1]);
      TOBLIGNE.PutValue('MONTANTPVQ',Valeur[2]);
      TOBLIGNE.PutValue('MONTANTPVTTCQ',Valeur[3]);
      TOBLIGNE.PutValue('PAMOYENQ',Valeur[7]);
      TOBLIGNE.PutValue('PRMOYENQ',Valeur[8]);
      TOBLIGNE.PutValue('PVMOYENQ',Valeur[9]);
      TOBLIGNE.PutValue('PVMOYENTTCQ',Valeur[10]);
      TOBLIGNE.PutValue('RATIOQTE',Valeur[11]);
      TOBLIGNE.PutValue('TYPE','F'); // type fourniture
      // NEW CHAMPS
      TOBLIGNE.PutValue('MONTANTFGQ',Valeur[12]);
      TOBLIGNE.PutValue('MONTANTFCQ',Valeur[13]);
      TOBLIGNE.PutValue('MONTANTFRQ',Valeur[14]);
      TOBLIGNE.PutValue('MONTANTPAFGQ',Valeur[15]);
      TOBLIGNE.PutValue('MONTANTPAFCQ',Valeur[16]);
      TOBLIGNE.PutValue('MONTANTPAFRQ',Valeur[17]);
      //
      if OKApplicFC then
      begin
      	TOBLIGNE.PutValue('APPLICCOEFFC','X');
      end;
      if OKApplicFR then
      begin
      	TOBLIGNE.PutValue('APPLICCOEFFR','X');
      end;

      if ((TOBRef.GetValue(prefixe+'_COEFFR')<>0) and (OKApplicFR)) or
      	 ((TOBRef.GetValue(prefixe+'_COEFFC')<>0) and (OKApplicFC)) then
      begin
      	TOBLIGNE.PutValue ('BLOQUEPR','X');
      	TOBCODEArt.PutValue ('BLOQUEPR','X');
      	TOBN.PutValue ('BLOQUEPR','X');
      	TOBC.PutValue ('BLOQUEPR','X');
      	TOBPERE.PutValue ('BLOQUEPR','X');
      	TOBREF.PutValue ('BLOQUEPR','X');
      end;

			if ((TbgBordereau in GestionInterdit) and (FromBordereau)) or (TOBL.GetValue('GL_BLOQUETARIF')='X') then
      begin
      	TOBLIGNE.PutValue ('BLOQUEPV','X');
      	TOBCODEArt.PutValue ('BLOQUEPV','X');
      	TOBN.PutValue ('BLOQUEPV','X');
      	TOBC.PutValue ('BLOQUEPV','X');
      	TOBPERE.PutValue ('BLOQUEPV','X');
      	TOBREF.PutValue ('BLOQUEPV','X');
      end;
    END;
    SetArrondi (TOBLigne);
//    TOBLIGNE.Data := TOBOUV; // lien vers TOBOUVRAGE
    TobLigne.data := TOBref;
  END;
END;

procedure TraiteOuvrageLIG(TobSimulPres,TOBOUvrage,TOBArticles,TOBL: TOB; Paragraphe: string; AvecParagraphe : boolean; gestionInterdit : TSGestionInterdit; FromBordereau : boolean);
var TOBDETOUV : TOB;
		TheLibelle : string;
begin

  //FV1 : 26/07/2016
  if TOBL.GetValue('GL_INDICENOMEN') <> 0 then
  begin
    TOBDetOUV := TOBOUvrage.detail[TOBL.GetValue('GL_INDICENOMEN')-1];
    if TOBDetOUv = nil then exit;
    if TOBL.GetValue('GL_TYPEARTICLE')='ARP' then TheLibelle := TOBL.GetValue('GL_LIBELLE') else TheLibelle := '';
    TraiteDetailOuv (AvecParagraphe,Paragraphe,TOBArticles,TobSimulPres,TOBL,TOBDetOuv,TOBL.GetValue('GL_QTEFACT'),1,GestionInterdit,FromBordereau,TheLibelle);
  end;

end;

Procedure TraiteDetailOuv (AvecParagraphe: boolean;Paragraphe : string;TOBArticles,TobSimulPres,TOBL,TOBOuv : TOB;Quantite,QteDuDetail : double;gestionInterdit : TSGestionInterdit; FromBordereau : boolean; LibellePrixPose : string='');
var QteSuiv,QteDetSuiv,PrixPourQte : double;
    TOBDetOuv,TOBA : TOB;
    Indice : integer;
    Valeur : T_Valeurs;
    QQ : Tquery;
    TheArticle,TheCodeArticle,TheLibelle : string;
begin
  if TOBOuv = nil then exit;
  if QteDuDetail = 0 then QteDuDetail := 1;
  for Indice := 0 to TOBOUV.detail.count -1 do
  begin
    TOBDetOuv := TOBOuv.detail[Indice];
    TheCodeArticle := TOBDetOuv.GetValue('BLO_CODEARTICLE');
    TheArticle := TOBDetOuv.GetValue('BLO_ARTICLE');
    if TOBDetOuv.detail.count > 0 then
    begin
      if LibelleprixPose = '' then
      begin
        if TOBDetOUV.GEtVAlue('BLO_TYPEARTICLE')='ARP' then TheLIbelle:= TOBDetOUV.GetValue('BLO_LIBELLE') else TheLibelle := '';
      end else TheLibelle := LibellePrixPose;
      QteSuiv := Quantite * TOBDETOUV.GetValue('BLO_QTEFACT');
      QteDetSuiv := QteduDetail * TOBDETOUv.GetValue('BLO_QTEDUDETAIL');
      TraiteDetailOuv (AvecParagraphe,Paragraphe,TOBArticles,TOBSimulPres,TOBL,TOBDETOUv,QteSuiv,QteDetSuiv,gestionInterdit,FromBordereau,TheLIbelle);
    end else
    BEGIN
      TOBA := FindTOBArtSais (TobArticles,TOBDetOuv.GetValue('BLO_ARTICLE'));
      if TOBA = nil then
      begin
        TOBA := TOB.create('ARTICLE',TobArticles,-1);
        QQ := OpenSQL ('SELECT A.*,N.BNP_TYPERESSOURCE,N.BNP_LIBELLE FROM ARTICLE A LEFT JOIN NATUREPREST N ON A.GA_NATUREPRES=N.BNP_NATUREPRES WHERE GA_ARTICLE="'+TOBDetOuv.GetValue('BLO_ARTICLE')+'"',true,-1,'',true);
        if not QQ.eof then TOBA.selectDb ('',QQ);
        ferme (QQ);
      end;
      if TOBDetOuv.GetValue('BLO_QTEDUDETAIL') = 0 then TOBDETOUV.PutValue('BLO_QTEDUDETAIL',1);
      if TOBDetOuv.GetValue('BLO_PRIXPOURQTE') = 0 then TOBDETOUV.PutValue('BLO_PRIXPOURQTE',1);
      PrixPourQte := TOBDetOuv.GetValue('BLO_PRIXPOURQTE');
      Valeur[2] := (TOBDetOuv.GetValue ('BLO_PUHTDEV') * TOBDetOuv.GetValue('BLO_QTEFACT') * Quantite) / (QteDuDetail* PrixPourQte*TOBDetOuv.GetValue('BLO_QTEDUDETAIL'));
      Valeur[3] := (TOBDetOuv.GetValue ('BLO_PUTTCDEV') * TOBDetOuv.GetValue('BLO_QTEFACT') * Quantite) / (QteDuDetail* PrixPourQte*TOBDetOuv.GetValue('BLO_QTEDUDETAIL'));
      valeur[6] := (TOBDetOuv.GetValue('BLO_QTEFACT') * quantite) / (QteDuDetail* TOBDetOuv.GetValue('BLO_QTEDUDETAIL'));
      valeur[11] := (quantite) / (QteDuDetail* TOBDetOuv.GetValue('BLO_QTEDUDETAIL'));
      if valeur[11] = 0 then valeur[11] := 1;
      valeur[9] := TOBDetOuv.GetValue('BLO_PUHTDEV');
      valeur[10] := TOBDetOuv.GetValue('BLO_PUTTCDEV');
      if TOBA.getvalue('GA_TYPEARTICLE')<>'POU' then
      Begin
        Valeur[0] := (TOBDetOuv.GetValue ('BLO_DPA') * TOBDetOuv.GetValue('BLO_QTEFACT') * Quantite) / (QteDuDetail* PrixPourQte*TOBDetOuv.GetValue('BLO_QTEDUDETAIL'));
        Valeur[1] := (TOBDetOuv.GetValue ('BLO_DPR') * TOBDetOuv.GetValue('BLO_QTEFACT') * Quantite) / (QteDuDetail* PrixPourQte*TOBDetOuv.GetValue('BLO_QTEDUDETAIL'));
        valeur[7] := TOBDetOuv.GetValue('BLO_DPA');
        valeur[8] := TOBDetOuv.GetValue('BLO_DPR');
        // NEW CHAMPS
        valeur [12] := (TOBDetOuv.GetValue('BLO_MONTANTFG')*Quantite) / (QteDuDetail* PrixPourQte*TOBDetOuv.GetValue('BLO_QTEDUDETAIL'));
        valeur [13] := (TOBDetOuv.GetValue('BLO_MONTANTFC')*Quantite) / (QteDuDetail* PrixPourQte*TOBDetOuv.GetValue('BLO_QTEDUDETAIL'));
        valeur [14] := (TOBDetOuv.GetValue('BLO_MONTANTFR')*Quantite) / (QteDuDetail* PrixPourQte*TOBDetOuv.GetValue('BLO_QTEDUDETAIL'));
        valeur [15] := (TOBDetOuv.GetValue('BLO_MONTANTPAFG')*Quantite) / (QteDuDetail* PrixPourQte*TOBDetOuv.GetValue('BLO_QTEDUDETAIL'));
        valeur [16] := (TOBDetOuv.GetValue('BLO_MONTANTPAFC')*Quantite) / (QteDuDetail* PrixPourQte*TOBDetOuv.GetValue('BLO_QTEDUDETAIL'));
        valeur [17] := (TOBDetOuv.GetValue('BLO_MONTANTPAFR')*Quantite) / (QteDuDetail* PrixPourQte*TOBDetOuv.GetValue('BLO_QTEDUDETAIL'));
        //

      End else
      Begin
        Valeur[0] := 0;
        Valeur[1] := 0;
        Valeur[2]:=Valeur[2]/100;
        Valeur[7] := 0;
        Valeur[8] := 0;
        // NEW CHAMPS
        valeur [12] := 0;
        valeur [13] := 0;
        valeur [14] := 0;
        valeur [15] := 0;
        valeur [16] := 0;
        valeur [17] := 0;
        //
      End;
      TraiteElementLig (TOBSimulPres,TOBL,TOBA,Paragraphe,valeur,AvecParagraphe,gestionInterdit,FromBordereau,TOBDetOUV,LibellePrixPose);
    END;
  end;
end;

function GetMontantAchatLIGNE (TOBL : TOB) : double;
BEGIN
if TOBL.GetValue('GL_PRIXPOURQTE') = 0 then TOBL.PutValue('GL_PRIXPOURQTE',1);
result := (TobL.GetValue('GL_DPA') * TOBL.GetValue('GL_QTEFACT')) / TOBL.getValue('GL_PRIXPOURQTE');
END;

function GetMontantRevientLIGNE (TOBL : TOB) : double;
BEGIN
if TOBL.GetValue('GL_PRIXPOURQTE') = 0 then TOBL.PutValue('GL_PRIXPOURQTE',1);
result := (TobL.GetValue('GL_DPR') * TOBL.GetValue('GL_QTEFACT')) / TOBL.getValue('GL_PRIXPOURQTE');
END;

function GetMontantVenteLIGNE (TOBL : TOB) : double;
BEGIN
if TOBL.GetValue('GL_PRIXPOURQTE') = 0 then TOBL.PutValue('GL_PRIXPOURQTE',1);
result := (TobL.GetValue('GL_PUHTDEV') * TOBL.GetValue('GL_QTEFACT')) / TOBL.getValue('GL_PRIXPOURQTE');
END;

function GetMontantVenteTTCLIGNE (TOBL : TOB) : double;
BEGIN
if TOBL.GetValue('GL_PRIXPOURQTE') = 0 then TOBL.PutValue('GL_PRIXPOURQTE',1);
result := (TobL.GetValue('GL_PUTTCDEV') * TOBL.GetValue('GL_QTEFACT')) / TOBL.getValue('GL_PRIXPOURQTE');
END;

procedure FinaliseSimulation (TOBSImulPres : TOB; var Valor : T_Valeurs);
var Indice : Integer;
    TOBS : TOB;
    valeur : T_Valeurs;
begin
For Indice := 0 to TOBSimulPres.detail.count -1 do
    begin
    TOBS := TOBSimulPres.detail[Indice];
    if TOBS.detail.count > 0 then
       begin
       InitTableau (Valeur);
       FinaliseSimulation (TOBS,Valeur);
       Valor := CalculSurTableau ('+',valor,valeur);
       end else
       begin
       ValorisationFinale (TOBS);
       InitTableau (Valeur);
       GetValorisation (TOBS,Valeur);
       Valor := CalculSurTableau ('+',valor,valeur);
       end;
    end;
SetValorisation (TOBSimulPres,Valor);
ValorisationFinale (TOBSimulPres);
end;

procedure SetArrondi (TOBS : TOB);
begin
TOBS.PutValue ('MONTANTPAH',Arrondi(TOBS.GetValue('MONTANTPAH'),4));
TOBS.PutValue ('MONTANTPRH',Arrondi(TOBS.GetValue('MONTANTPRH'),4));
TOBS.PutValue ('MONTANTPVH',Arrondi(TOBS.GetValue('MONTANTPVH'),V_PGI.OKDecV));
TOBS.PutValue ('MONTANTPVTTCH',Arrondi(TOBS.GetValue('MONTANTPVTTCH'),V_PGI.OKDecV));
//
TOBS.PutValue ('MONTANTPAQ',Arrondi(TOBS.GetValue('MONTANTPAQ'),4));
TOBS.PutValue ('MONTANTPRQ',Arrondi(TOBS.GetValue('MONTANTPRQ'),4));
TOBS.PutValue ('MONTANTPVQ',Arrondi(TOBS.GetValue('MONTANTPVQ'),V_PGI.OKDecV));
TOBS.PutValue ('MONTANTPVTTCQ',Arrondi(TOBS.GetValue('MONTANTPVTTCQ'),V_PGI.OKDecV));
TOBS.PutValue ('MONTANTPA',Arrondi(TOBS.GetValue('MONTANTPA'),4));
TOBS.PutValue ('MONTANTPR',Arrondi(TOBS.GetValue('MONTANTPR'),4));
TOBS.PutValue ('MONTANTPV',Arrondi(TOBS.GetValue('MONTANTPV'),V_PGI.okDecV));
TOBS.PutValue ('MONTANTPVTTC',Arrondi(TOBS.GetValue('MONTANTPVTTC'),V_PGI.okDecV));
//
TOBS.PutValue ('PAMOYENH', Arrondi(TOBS.GetValue('PAMOYENH'),V_PGI.okDecP));
TOBS.PutValue ('PRMOYENH', Arrondi(TOBS.GetValue('PRMOYENH'),V_PGI.okDecP));
TOBS.PutValue ('PVMOYENH', Arrondi(TOBS.GetValue('PVMOYENH'),V_PGI.okDecP));
TOBS.PutValue ('PVMOYENTTCH',Arrondi(TOBS.GetValue('PVMOYENTTCH'),V_PGI.okDecP));
//
TOBS.PutValue ('PAMOYENQ', Arrondi(TOBS.GetValue('PAMOYENQ'),V_PGI.okDecP));
TOBS.PutValue ('PRMOYENQ', Arrondi(TOBS.GetValue('PRMOYENQ'),V_PGI.okDecP));
TOBS.PutValue ('PVMOYENQ', Arrondi(TOBS.GetValue('PVMOYENQ'),V_PGI.okDecP));
TOBS.PutValue ('PVMOYENTTCQ',Arrondi(TOBS.GetValue('PVMOYENTTCQ'),V_PGI.okDecP));
//
TOBS.PutValue ('PAMOYEN', Arrondi(TOBS.GetValue('PAMOYEN'),V_PGI.okDecp));
TOBS.PutValue ('PRMOYEN', Arrondi(TOBS.GetValue('PRMOYEN'),V_PGI.okDecP));
TOBS.PutValue ('PVMOYEN', Arrondi(TOBS.GetValue('PVMOYEN'),V_PGI.okDecP));
TOBS.PutValue ('PVMOYENTTC',Arrondi(TOBS.GetValue('PVMOYENTTC'),V_PGI.okDecP));
end;

procedure ValorisationFinale (TOBS : TOB; UniquementSommeAchat : boolean=false);
var MontantPaH, MontantPRH, MontantPVH,MontantPVTTCH,MontantPaQ, MontantPRQ, MontantPVQ , MontantPVTTCQ : double;
		MontantPAFGQ,MontantPAFCQ,MontantPAFRQ,MontantFGQ,MontantFRQ,MontantFCQ : double;
		MontantPAFGH,MontantPAFCH,MontantPAFRH,MontantFGH,MontantFRH,MontantFCH : double;
		MontantPAFG,MontantPAFC,MontantPAFR,MontantFG,MontantFR,MontantFC : double;
    Heures , Qte: double;
begin
  MontantPAH := TOBS.GetValue('MONTANTPAH');
  MontantPAQ := TOBS.GetValue('MONTANTPAQ');
	if not UniquementSommeAchat then
  begin
    TOBS.PutValue ('HEURES',Arrondi(TOBS.GetValue('HEURES'),V_PGI.okDECQ));
    TOBS.PutValue ('QTE',Arrondi(TOBS.GetValue('QTE'),V_PGI.OKDecQ));
    Heures := TOBS.GetValue('HEURES');
    Qte := TOBS.GetValue('QTE');
    //
    MontantPRH := TOBS.GetValue('MONTANTPRH');
    MontantPVH := TOBS.GetValue('MONTANTPVH');
    MontantPVTTCH := TOBS.GetValue('MONTANTPVTTCH');
    MontantPRQ := TOBS.GetValue('MONTANTPRQ');
    MontantPVQ := TOBS.GetValue('MONTANTPVQ');
    MontantPVTTCQ := TOBS.GetValue('MONTANTPVTTCQ');
    //
    MontantFCQ := TOBS.GetValue('MONTANTFCQ');
    MontantFRQ := TOBS.GetValue('MONTANTFRQ');
  end;
  MontantFGQ := TOBS.GetValue('MONTANTFGQ');
  MontantPAFGQ := TOBS.GetValue('MONTANTPAFGQ');
  MontantPAFCQ := TOBS.GetValue('MONTANTPAFCQ');
  MontantPAFRQ:= TOBS.GetValue('MONTANTPAFRQ');
  MontantFGH := TOBS.GetValue('MONTANTFGH');
	if not UniquementSommeAchat then
  begin
    MontantFCH := TOBS.GetValue('MONTANTFCH');
    MontantFRH := TOBS.GetValue('MONTANTFRH');
  end;
  MontantPAFGH := TOBS.GetValue('MONTANTPAFGH');
  MontantPAFCH := TOBS.GetValue('MONTANTPAFCH');
  MontantPAFRH:= TOBS.GetValue('MONTANTPAFRH');
  //
  TOBS.PutValue ('MONTANTPA',MontantPAQ+MontantPaH);
  //
  if not UniquementSommeAchat then
  begin
    TOBS.PutValue ('MONTANTPR',MontantPRQ+MontantPRH);
    TOBS.PutValue ('MONTANTPV',MontantPVQ+MontantPVH);
    TOBS.PutValue ('MONTANTPVTTC',MontantPVTTCQ+MontantPVTTCH);
    //
    TOBS.PutValue ('MONTANTFC',MontantFCQ+MontantFCH);
    TOBS.PutValue ('MONTANTFR',MontantFRQ+MontantFRH);
  end;
  TOBS.PutValue ('MONTANTFG',MontantFGQ+MontantFGH);
  //
  TOBS.PutValue ('MONTANTPAFG',MontantPAFGQ+MontantPaFGH);
  TOBS.PutValue ('MONTANTPAFC',MontantPAFCQ+MontantPaFCH);
  TOBS.PutValue ('MONTANTPAFR',MontantPAFRQ+MontantPaFRH);
  //
  if not UniquementSommeAchat then
  begin
    MontantFG := TOBS.GetValue ('MONTANTFG');
    MontantFC := TOBS.GetValue ('MONTANTFC');
    MontantFR := TOBS.GetValue ('MONTANTFR');
    MontantPAFG := TOBS.GetValue ('MONTANTPAFG');
    MontantPAFC := TOBS.GetValue ('MONTANTPAFC');
    MontantPAFR := TOBS.GetValue ('MONTANTPAFR');
    //
    if MontantPaFGH <> 0 then TOBS.PutValue ('COEFFGH',Arrondi((MontantPAFGH+MontantFGH )/ MontantPAFGH,9))
                         else TOBS.PutValue ('COEFFGH',0);
    if (MontantPAH+MontantFGH) <> 0 then TOBS.PutValue ('COEFFCH',Arrondi((MontantPAH+MontantFCH+MontantFGH )/(MontantPAH+MontantFGH),9))
                         						else TOBS.PutValue ('COEFFCH',0);
    if (MontantPAH+MontantFCH+MontantFGH) <> 0 then TOBS.PutValue ('COEFFRH',Arrondi((MontantPAH+MontantFRH+MontantFCH+MontantFGH)/ (MontantPAH+MontantFCH+MontantFGH),9))
                         											 else TOBS.PutValue ('COEFFRH',0);
    //
    if TOBS.GetValue ('COEFMAMODIFIE') <> 'X' then
    begin
      if MOntantPRH <> 0 then TOBS.PutValue ('COEFMARGH',Arrondi(MontantPVH / MontantPrH,4))
                         else TOBS.PutValue ('COEFMARGH',0);
    end;
    if MontantPVH <> 0 then TOBS.PutValue ('TAUXTVAH',Arrondi(MontantPVTTCH / MontantPVH,6))
                       else TOBS.PutValue ('TAUXTVAH',0);
    //
    if MontantPaFGQ <> 0 then TOBS.PutValue ('COEFFGQ',Arrondi((MontantPAFGQ+MontantFGQ )/ MontantPAFGQ,9))
                         else TOBS.PutValue ('COEFFGQ',0);
    if (MontantPAQ+MontantFGQ) <> 0 then TOBS.PutValue ('COEFFCQ',Arrondi((MontantPAQ+MontantFCQ+MontantFGQ )/(MontantPAQ+MontantFGQ),9))
                         						else TOBS.PutValue ('COEFFCQ',0);
    if (MontantPAQ+MontantFCQ+MontantFGQ) <> 0 then TOBS.PutValue ('COEFFRQ',Arrondi((MontantPAQ+MontantFRQ+MontantFCQ+MontantFGQ)/ (MontantPAQ+MontantFCQ+MontantFGQ),9))
                         											 else TOBS.PutValue ('COEFFRQ',0);
    //
    if TOBS.GetValue ('COEFMAMODIFIE') <> 'X' then
    begin
      if MOntantPRQ <> 0 then TOBS.PutValue ('COEFMARGQ',Arrondi(MontantPVQ / MontantPrQ,4))
                        else TOBS.PutValue ('COEFMARGQ',0);
    end;
    if MontantPVQ <> 0 then TOBS.PutValue ('TAUXTVAQ',Arrondi(MontantPVTTCQ / MontantPVQ,6))
                      else TOBS.PutValue ('TAUXTVAQ',0);
    //
    if MontantPaFG <> 0 then TOBS.PutValue ('COEFFG',Arrondi((MontantPAFG+MontantFG )/ MontantPAFG,9))
                        else TOBS.PutValue ('COEFFG',0);
    if (MontantPAFG+MontantFG) <> 0 then TOBS.PutValue ('COEFFC',Arrondi((MontantPAFG+MontantFC+MontantFG )/(MontantPAFG+MontantFG),9))
                        						else TOBS.PutValue ('COEFFC',0);
    if (MontantPAFG+MontantFC+MontantFG) <> 0 then TOBS.PutValue ('COEFFR',Arrondi((MontantPAFG+MontantFR+MontantFC+MontantFG)/ (MontantPAFG+MontantFC+MontantFG),9))
                        											else TOBS.PutValue ('COEFFR',0);
    //
    if TOBS.GetValue('COEFMAMODIFIE')<>'X' then
    begin
      if MOntantPRQ+MOntantPRH <> 0 then TOBS.PutValue ('COEFMARG',Arrondi((MontantPVQ+MontantPVH) / (MontantPrQ+MontantPrH),4))
                                    else TOBS.PutValue ('COEFMARG',0);
    end;
    if MOntantPVQ+MOntantPVH <> 0 then TOBS.PutValue ('TAUXTVA',Arrondi((MontantPVTTCQ+MontantPVTTCH) / (MontantPVQ+MontantPVH),6))
                                  else TOBS.PutValue ('TAUXTVA',0);
    if Heures <> 0 then
       BEGIN
       TOBS.PutValue ('PAMOYENH', MontantPaH/Heures);
       TOBS.PutValue ('PRMOYENH', MontantPRH/Heures);
       TOBS.PutValue ('PVMOYENH', MontantPVH/Heures);
       TOBS.PutValue ('PVMOYENTTCH',MontantPVTTCH/Heures);
       end else
       BEGIN
       TOBS.PutValue ('PAMOYENH',0);
       TOBS.PutValue ('PRMOYENH',0);
       TOBS.PutValue ('PVMOYENH',0);
       TOBS.PutValue ('PVMOYENTTCH',0);
       END;
    if Qte <> 0 then
       BEGIN
       TOBS.PutValue ('PAMOYENQ', MontantPaQ/Qte);
       TOBS.PutValue ('PRMOYENQ', MontantPRQ/Qte);
       TOBS.PutValue ('PVMOYENQ', MontantPVQ/Qte);
       TOBS.PutValue ('PVMOYENTTCQ', MontantPVTTCQ/Qte);
       end else
       begin
       TOBS.PutValue ('PAMOYENQ',0);
       TOBS.PutValue ('PRMOYENQ',0);
       TOBS.PutValue ('PVMOYENQ',0);
       TOBS.PutValue ('PVMOYENTTCQ',0);
       end;
    if Heures+qte <> 0 then
       BEGIN
       TOBS.PutValue ('PAMOYEN', (MontantPAQ+MontantPaH)/(Heures+Qte));
       TOBS.PutValue ('PRMOYEN', (MontantPRQ+MontantPrH)/(Heures+Qte));
       TOBS.PutValue ('PVMOYEN', (MontantPVQ+MontantPVH)/(Heures+Qte));
       TOBS.PutValue ('PVMOYENTTC', (MontantPVTTCQ+MontantPVTTCH)/(Heures+Qte));
       END else
       BEGIN
       TOBS.PutValue ('PAMOYEN',0);
       TOBS.PutValue ('PRMOYEN',0);
       TOBS.PutValue ('PVMOYEN',0);
       TOBS.PutValue ('PVMOYENTTC',0);
       END;
  end;
  if Heures <> 0 then TOBS.PutValue('TYPE','H'); // type Horaire
  if Qte <> 0 then TOBS.PutValue('TYPE','F'); // type fourniture
  if (Heures<> 0) and (Qte <> 0 )then TOBS.PutValue('TYPE','M'); // type Mixte
  SetArrondi (TOBS);
end;

procedure GetValorisation (TOBS : TOB; Var Valeur : T_Valeurs; UniquementAchat : boolean=false);
begin
  Valeur[0] := TOBS.GetValue('MONTANTPAQ');
  if not UniquementAchat then
  begin
    Valeur[1] := TOBS.GetValue('MONTANTPRQ');
    Valeur[2] := TOBS.GetValue('MONTANTPVQ');
    Valeur[3] := TOBS.GetValue('MONTANTPVTTCQ');
  end;
  // NEW CHAMPS
  Valeur[4] := TOBS.GetValue('MONTANTFGQ');
  Valeur[5] := TOBS.GetValue('MONTANTFCQ');
  Valeur[6] := TOBS.GetValue('MONTANTFRQ');
  Valeur[7] := TOBS.GetValue('MONTANTPAFGQ');
  Valeur[8] := TOBS.GetValue('MONTANTPAFCQ');
  Valeur[9] := TOBS.GetValue('MONTANTPAFRQ');
  //
  Valeur[10] := TOBS.GetValue('MONTANTPAH');
  if not UniquementAchat then
  begin
    Valeur[11] := TOBS.GetValue('MONTANTPRH');
    Valeur[12] := TOBS.GetValue('MONTANTPVH');
    Valeur[13] := TOBS.GetValue('MONTANTPVTTCH');
  end;
  // NEW CHAMPS
  Valeur[14] := TOBS.GetValue('MONTANTFGH');
  Valeur[15] := TOBS.GetValue('MONTANTFCH');
  Valeur[16] := TOBS.GetValue('MONTANTFRH');
  Valeur[17] := TOBS.GetValue('MONTANTPAFGH');
  Valeur[18] := TOBS.GetValue('MONTANTPAFCH');
  Valeur[19] := TOBS.GetValue('MONTANTPAFRH');

  if not UniquementAchat then
  begin
    Valeur[20] := TOBS.GetValue('QTE');
    Valeur[21] := TOBS.GetValue('HEURES');
  end;
end;

procedure SetValorisation (TOBS : TOB; Valor:T_Valeurs; UniquementAchat : boolean=false);
begin
  TOBS.PutValue ('MONTANTPAQ',Valor[0]);
  if not UniquementAchat then
  begin
    TOBS.PutValue ('MONTANTPRQ',Valor[1]);
    TOBS.PutValue ('MONTANTPVQ',Valor[2]);
    TOBS.PutValue ('MONTANTPVTTCQ',Valor[3]);
  end;
  // NEW CHAMPS
  TOBS.PutValue ('MONTANTFGQ',Valor[4]);
  TOBS.PutValue ('MONTANTFCQ',Valor[5]);
  TOBS.PutValue ('MONTANTFRQ',Valor[6]);
  TOBS.PutValue ('MONTANTPAFGQ',Valor[7]);
  TOBS.PutValue ('MONTANTPAFCQ',Valor[8]);
  TOBS.PutValue ('MONTANTPAFRQ',Valor[9]);
  //
  TOBS.PutValue ('MONTANTPAH',Valor[10]);
  if not UniquementAchat then
  begin
    TOBS.PutValue ('MONTANTPRH',Valor[11]);
    TOBS.PutValue ('MONTANTPVH',Valor[12]);
    TOBS.PutValue ('MONTANTPVTTCH',Valor[13]);
  end;
  // NEW CHAMPS
  TOBS.PutValue ('MONTANTFGH',Valor[14]);
  TOBS.PutValue ('MONTANTFCH',Valor[15]);
  TOBS.PutValue ('MONTANTFRH',Valor[16]);
  TOBS.PutValue ('MONTANTPAFGH',Valor[17]);
  TOBS.PutValue ('MONTANTPAFCH',Valor[18]);
  TOBS.PutValue ('MONTANTPAFRH',Valor[19]);
  //
  if not UniquementAchat then
  begin
    TOBS.PutValue ('QTE',Valor[20]);
    TOBS.PutValue ('HEURES',Valor[21]);
  end;
  //
end;

procedure IndiceGrilleInit (TOBSimulPres : TOB; NiveauMax : integer; var Numligne : integer);
var Indice : integer;
    TOBS : TOB;
begin
for Indice := 0 to TOBSimulPres.detail.count -1 do
    begin
    TOBS := TOBSimulPres.detail[Indice];
    if TOBS.GetValue('NIVEAU')> 0 then // on saute le niveau document
       BEGIN
       inc(NumLigne);
       TOBS.putValue('NUMAFF',NumLigne);
       END;
    if NiveauMax > TOBS.GetValue('NIVEAU') then
       begin
       TOBS.putValue('OPEN','X');
       IndicegrilleInit (TOBS,NiveauMax,Numligne);
       end;
    end;
end;

procedure AfficheDetailSimul (GS : THGrid;TOBSimulpres : TOB;NiveauMax : integer;ListeSaisie:String);
var Indice : integer;
    TOBS : TOB;
    Ligne : integer;
begin
for Indice := 0 to TOBSimulPres.detail.count -1 do
    begin
    TOBS := TOBSimulPres.detail[Indice];
    if TOBS.GetValue('NIVEAU') > 0 then
       begin
       LIgne := TOBS.GetValue('NUMAFF');
       if Ligne = 0 then break;
       TOBS.PutLigneGrid (GS,Ligne,false,false,ListeSaisie);
       end;
    if (TOBS.detail.count > 0) and (NiveauMax > TOBS.GetValue('NIVEAU')) then
       AfficheDetailSimul (GS,TOBS,NiveauMax,ListeSaisie);
    end;
end;

procedure AfficheLaGrille (GS : THGrid;ListeSaisie: string;TobSimulPres : TOB;NiveauMax,NbElt : integer;AvecParagraphe : boolean);
begin
GS.rowCount := NbElt+1;
AfficheDetailSimul (GS,TOBSimulPres,NiveauMax,ListeSaisie);
if AvecParagraphe then TOBSimulPres.PutValue('LIBPARAGRA','TOTAL')
                  else TOBSimulPres.PutValue('LIBCATEGORIE','TOTAL');
TOBSimulPres.PutValue('NUMAFF',GS.rowCount -1);
TOBSimulPres.PutLigneGrid (GS,GS.rowCount-1,false,false,ListeSaisie);
end;


function GetTOBLigne (TOBSimulPres : TOB; Arow : integer): TOB;
begin
Result := TOBSimulPres.findFirst(['NUMAFF'],[Arow],true);
if Result = nil then if TobSimulPres.GetValue('NUMAFF')=Arow then result := TOBSimulPres;
end;

function RecupTypeGraph (TOBS : TOB) : integer;
begin
if TOBS.GetValue('OPEN') = 'X' then Result := 1 Else Result := 0;
end;

procedure OuvreBranche (GS : THGrid; TobSimulPres: TOB;var Arow:integer;var niveau : integer; var Found : boolean; Force : boolean);
var TOBS : TOB;
    Indice : integer;
    LigneCourante: integer;
begin
for Indice := 0 to TOBSimulPres.detail.count -1 do
    begin
    TOBS := TOBSimulPres.detail[Indice];
    LigneCourante := TOBS.GetValue('NUMAFF');
    if (found) and ((LigneCourante<> 0) or (force)) then
       BEGIN
       inc(Arow);
       TOBS.putValue('NUMAFF',Arow);
       if TOBS.GetValue('NIVEAU') > niveau then niveau := TOBS.GetValue('NIVEAU');
       if TOBS.detail.count > 0 then OuvreBranche (GS,TOBS,Arow,niveau,Found,false);
       continue;
       END;
    if (LigneCourante = 0 ) then break;
    if LigneCourante = Arow then
       BEGIN
       found := true;
       // on force la mAJ du numero de ligne dans le niveai immediatement superieur
       OuvreBranche (GS,TOBS,Arow,niveau,Found,true);
       END else if TOBS.detail.count > 0 then OuvreBranche (GS,TOBS,Arow,niveau,Found,Force);
    end;
end;

procedure FermeBranche (GS:THGrid; TobSimulPres: TOB;var Arow:integer;var Niveau : integer; var Found : boolean);
var TOBS: TOB;
    Indice : integer;
    LigneCourante: integer;
begin
for Indice := 0 to TOBSimulPres.detail.count -1 do
    begin
    TOBS := TOBSimulPres.detail[Indice];
    LigneCourante := TOBS.GetValue('NUMAFF');
    if (LigneCourante <> Arow) and (lignecourante > 0) then if TOBS.GetValue('NIVEAU') > Niveau then NIveau := TOBS.GetValue('NIVEAU');
    if (found) and (LigneCourante<> 0)  then
       BEGIN
       if TOBS.GetValue('NIVEAU') > Niveau then NIveau := TOBS.GetValue('NIVEAU');
       inc(Arow);
       TOBS.putValue('NUMAFF',Arow);
       FermeBranche (GS,TOBS,Arow,niveau,Found);
       continue;
       END;
    if (LigneCourante = 0 ) then break;
    if LigneCourante = Arow then
       BEGIN
       found := true;
       InitBranche (GS,TOBS,Arow);
       END else if TOBS.detail.count > 0 then FermeBranche (GS,TOBS,Arow,niveau,Found);
    end;
end;

procedure InitBranche (GS : THGrid; TOBSimulPres : TOB; Arow : integer);
var TOBS : TOB;
    Indice : integer;
begin
for Indice := 0 to TOBSimulPres.detail.count -1 do
    begin
    TOBS := TOBSimulPres.detail[Indice];
    if TOBS.GetValue('NUMAFF') = 0 then break;
    TOBS.PutValue('NUMAFF',0); TOBS.PuTValue('OPEN','-');
    if TOBS.detail.count > 0 then
       begin
       IniTBranche (GS,TOBS,Arow);
       end;
    end;
end;

procedure AppliqueChangement (Change: T_SimulChg;CoefChange:double;TOBSimulPres:TOB;var ListChange : Tlist;ValeurFixe : double=0; GestionInterdit : TSGestionInterdit=[]; Global : boolean=false);
var Indice : integer;
    TOBS : TOB;
begin
  for Indice := 0 to TOBSimulPres.detail.count -1 do
  begin
    TOBS := TOBSimulPres.detail[Indice];
    // on fait le changement au plus fin
    if TOBS.GetValue('NIVEAU') <> 5 then
    begin
      AppliqueChangement(Change,CoefChange,TOBS,ListChange,ValeurFixe,GestionInterdit,global);
    end;
    if TOBS.GetValue('NIVEAU') = 5 then
    begin
      if (CHange = tscHrs) and (TOBS.GetValue('TYPE') = 'F' ) then continue;
      if (CHange = tscQte) and (TOBS.GetValue('TYPE') = 'H' ) then continue;
      if (Change = TscHrs) then
      BEGIN
        if TOBS.NomTable = 'LIGNE' Then
        begin
          TOBS.PutValue('HEURES',arrondi(TOBS.GetValue('HEURES')*CoefChange,V_PGI.OKDecQ));
        end else
        begin
          if TOBS.GetValue('RATIOQTE') <> 0 then
          begin
          	TOBS.PutValue('HEURES',arrondi(TOBS.GetValue('HEURES')/TOBS.GetValue('RATIOQTE'),V_PGI.OKDecQ));
            TOBS.PutValue('HEURES',arrondi(TOBS.GetValue('HEURES')*CoefChange,V_PGI.OKDecQ));
          	TOBS.PutValue('HEURES',arrondi(TOBS.GetValue('HEURES')*TOBS.GetValue('RATIOQTE'),V_PGI.OKDecQ));
          end else
            TOBS.PutValue('HEURES',0.0);
        end;
        TOBS.PutValue('MONTANTPAH',TOBS.GetValue('HEURES')*TOBS.GetValue('PAMOYENH'));
        RecalculeMontantsPA (TOBS);
        TOBS.PutValue('MODIFIE','X');
      END;
      if (Change = TscQte)  then
      BEGIN
        if TOBS.NomTable = 'LIGNE' Then
        begin
          TOBS.PutValue('QTE',arrondi(TOBS.GetValue('QTE')*CoefChange,V_PGI.OKDecQ));
        end else
        begin
          if TOBS.GetValue('RATIOQTE') <> 0 then
          begin
             TOBS.PutValue('QTE',arrondi(TOBS.GetValue('QTE')/TOBS.GetValue('RATIOQTE'),V_PGI.OKDecQ));
             TOBS.PutValue('QTE',arrondi(TOBS.GetValue('QTE')*CoefChange,V_PGI.OKDecQ));
             TOBS.PutValue('QTE',arrondi(TOBS.GetValue('QTE')*TOBS.GetValue('RATIOQTE'),V_PGI.OKDecQ));
          end else
             TOBS.PutValue('QTE',0.0);
        end;
        TOBS.PutValue('MONTANTPAQ',TOBS.GetValue('QTE')*TOBS.GetValue('PAMOYENQ'));
        RecalculeMontantsPA (TOBS);
        TOBS.PutValue('MODIFIE','X');
      END;
      if (Change = TscPA)  then // modif Du Montant Achat
      begin
        if TOBS.GetValue('TYPE') = 'H' then
        BEGIN
          TOBS.PutValue('MONTANTPAH',TOBS.GetValue('MONTANTPAH')*CoefChange);
        END else
        BEGIN
          TOBS.PutValue('MONTANTPAQ',TOBS.GetValue('MONTANTPAQ')*CoefChange);
        END;
        RecalculeMontantsPA (TOBS);
        TOBS.PutValue('MODIFIE','X');
      end;
      if (Change = TscPR) and (TOBS.GetValue('BLOQUEPR')<>'X')  then
      begin
        if TOBS.GetValue('TYPE') = 'H' then
        BEGIN
          TOBS.PutValue('MONTANTPRH',TOBS.GetValue('MONTANTPRH')*CoefChange);
					if TOBS.GetValue('MONTANTPAH') <> 0 then
	          TOBS.PutValue('COEFFGH',TOBS.GetValue('MONTANTPRH')/TOBS.GetValue('MONTANTPAH'))
          else
          	TOBS.PutValue('COEFFGH',0);
        END ELSE
        BEGIN
          TOBS.PutValue('MONTANTPRQ',TOBS.GetValue('MONTANTPRQ')*CoefChange);
					if TOBS.GetValue('MONTANTPAQ') <> 0 then
	          TOBS.PutValue('COEFFGQ',TOBS.GetValue('MONTANTPRQ')/TOBS.GetValue('MONTANTPAQ'))
          else
          	TOBS.PutValue('COEFFGQ',0);
        END;
        TOBS.PutValue('MODIFIE','X');
      end;
      if (Change = TscPV) and (not (TbgBordereau in gestionInterdit))  then
      begin
        if TOBS.GetValue('TYPE') = 'H' then
        BEGIN
          TOBS.PutValue('MONTANTPVH',TOBS.GetValue('MONTANTPVH')*CoefChange);
          TOBS.PutValue('MONTANTPVTTCH',TOBS.GetValue('MONTANTPVH')*TOBS.GetValue('TAUXTVAH'));
        END ELSE
        BEGIN
          TOBS.PutValue('MONTANTPVQ',TOBS.GetValue('MONTANTPVQ')*CoefChange);
          TOBS.PutValue('MONTANTPVTTCQ',TOBS.GetValue('MONTANTPVQ')*TOBS.GetValue('TAUXTVAQ'));
        END;
        TOBS.PutValue('MODIFIE','X');
      end;
      if (Change = TscPVTTC) and (not (TbgBordereau in gestionInterdit)) then
      begin
        if TOBS.GetValue('TYPE') = 'H' then
        BEGIN
          TOBS.PutValue('MONTANTPVTTCH',TOBS.GetValue('MONTANTPVTTCH')*CoefChange);
					if TOBS.GetValue('TAUXTVAH') <> 0 then
          	TOBS.PutValue('MONTANTPVH',TOBS.GetValue('MONTANTPVTTCH')/TOBS.GetValue('TAUXTVAH'))
          else
          	TOBS.PutValue('MONTANTPVH',0);
        END ELSE
        BEGIN
          TOBS.PutValue('MONTANTPVTTCQ',TOBS.GetValue('MONTANTPVTTCQ')*CoefChange);
					if TOBS.GetValue('TAUXTVAQ') <> 0 then
          	TOBS.PutValue('MONTANTPVQ',TOBS.GetValue('MONTANTPVTTCQ')/TOBS.GetValue('TAUXTVAQ'))
          else
          	TOBS.PutValue('MONTANTPVQ',0);
        END;
        TOBS.PutValue('MODIFIE','X');
      end;
      //
      if (Change = TscCoefFG) then
      BEGIN
        if TOBS.GetValue('TYPE') = 'H' then
        BEGIN
          TOBS.PutValue('COEFFGH',ValeurFixe);
        END ELSE
        BEGIN
          TOBS.PutValue('COEFFGQ',ValeurFixe);
        END;
        RecalculeMontantsPA (TOBS);
        TOBS.PutValue('MODIFIE','X');
      END;
      //
      if (Change = TscCoefMA) and (not (TbgBordereau in gestionInterdit)) then
      BEGIN
        if TOBS.GetValue('TYPE') = 'H' then
        BEGIN
          TOBS.PutValue('MONTANTPVH',Arrondi(TOBS.GetValue('MONTANTPRH')*ValeurFixe,V_PGI.okdecV));
          TOBS.PutValue('MONTANTPVTTCH',Arrondi(TOBS.GetValue('MONTANTPVH')*TOBS.GetValue('TAUXTVAH'),V_PGI.okdecV));
          TOBS.PutValue('COEFMARGH',ValeurFixe);
          TOBS.PutValue('COEFMARG',ValeurFixe);
        END ELSE
        BEGIN
          TOBS.PutValue('MONTANTPVQ',Arrondi(TOBS.GetValue('MONTANTPRQ')*ValeurFixe,V_PGI.OkdecV));
          TOBS.PutValue('MONTANTPVTTCQ',Arrondi(TOBS.GetValue('MONTANTPVQ')*TOBS.GetValue('TAUXTVAQ'),V_PGI.okdecV));
          TOBS.PutValue('COEFMARGQ',ValeurFixe);
          TOBS.PutValue('COEFMARG',ValeurFixe);
        END;
        if not Global then TOBS.PutValue('COEFMAMODIFIE','X');
      END;

      if (Change = TscPUA)  then
      BEGIN
        if TOBS.GetValue('TYPE') = 'H' then
        BEGIN
          TOBS.PutValue('PAMOYENH',ValeurFixe);
          TOBS.PutValue('MONTANTPAH',arrondi(TOBS.GetValue('HEURES')*TOBS.GetValue('PAMOYENH'),V_PGI.OkdecV));
        END ELSE
        BEGIN
          TOBS.PutValue('PAMOYENQ',ValeurFixe);
          TOBS.PutValue('MONTANTPAQ',Arrondi(TOBS.GetValue('QTE')*TOBS.GetValue('PAMOYENQ'),V_PGI.okdecV));
        END;
        TOBS.PutValue('MODIFIE','X');
        RecalculeMontantsPA (TOBS);
      END;
      if (Change = TscPUR) and (TOBS.GetValue('BLOQUEPR')<>'X')  then
      BEGIN
        if TOBS.GetValue('TYPE') = 'H' then
        BEGIN
          TOBS.PutValue('PRMOYENH',ValeurFixe);
          TOBS.PutValue('MONTANTPRH',TOBS.GetValue('HEURES')*TOBS.GetValue('PRMOYENH'));
          if TOBS.GetValue('BLOQUEPV')<>'X' then
          begin
            TOBS.PutValue('MONTANTPVH',TOBS.GetValue('MONTANTPRH')*TOBS.GetValue('COEFMARGH'));
            TOBS.PutValue('MONTANTPVTTCH',TOBS.GetValue('MONTANTPVH')*TOBS.GetValue('TAUXTVAH'));
          end;
        END ELSE
        BEGIN
          TOBS.PutValue('PRMOYENQ',ValeurFixe);
          TOBS.PutValue('MONTANTPRQ',TOBS.GetValue('QTE')*TOBS.GetValue('PRMOYENQ'));
          if TOBS.GetValue('BLOQUEPV')<>'X' then
          begin
            TOBS.PutValue('MONTANTPVQ',TOBS.GetValue('MONTANTPRQ')*TOBS.GetValue('COEFMARGQ'));
            TOBS.PutValue('MONTANTPVTTCQ',TOBS.GetValue('MONTANTPVQ')*TOBS.GetValue('TAUXTVAQ'));
          end;
        END;
        TOBS.PutValue('MODIFIE','X');
      END;
      if (Change = TscPUV) and (not (TbgBordereau in gestionInterdit))  then
      BEGIN
        if TOBS.GetValue('TYPE') = 'H' then
        BEGIN
          TOBS.PutValue('PVMOYENH',ValeurFixe);
          TOBS.PutValue('MONTANTPVH',TOBS.GetValue('HEURES')*TOBS.GetValue('PVMOYENH'));
          TOBS.PutValue('MONTANTPVTTCH',TOBS.GetValue('MONTANTPVH')*TOBS.GetValue('TAUXTVAH'));
        END ELSE
        BEGIN
          TOBS.PutValue('PVMOYENQ',ValeurFixe);
          TOBS.PutValue('MONTANTPVQ',TOBS.GetValue('QTE')*TOBS.GetValue('PVMOYENQ'));
          TOBS.PutValue('MONTANTPVTTCQ',TOBS.GetValue('MONTANTPVQ')*TOBS.GetValue('TAUXTVAQ'));
        END;
        TOBS.PutValue('MODIFIE','X');
      END;
      if (Change = TscPUVTTC) and (not (TbgBordereau in gestionInterdit)) then
      BEGIN
        if TOBS.GetValue('TYPE') = 'H' then
        BEGIN
          TOBS.PutValue('PVMOYENTTCH',ValeurFixe);
          TOBS.PutValue('MONTANTPVTTCH',TOBS.GetValue('HEURES')*TOBS.GetValue('PVMOYENTTCH'));
					if TOBS.GetValue('TAUXTVAH') <> 0 then
          	TOBS.PutValue('MONTANTPVH',TOBS.GetValue('MONTANTPVTTCH')/TOBS.GetValue('TAUXTVAH'))
          else
          	TOBS.PutValue('MONTANTPVH',0);
					if TOBS.GetValue('HEURES') <> 0 then
	          TOBS.PutValue('PVMOYENH',TOBS.GetValue('MONTANTPVH')/TOBS.GetValue('HEURES'))
          else
          	TOBS.PutValue('PVMOYENH',0);
        END ELSE
        BEGIN
          TOBS.PutValue('PVMOYENQ',ValeurFixe);
          TOBS.PutValue('MONTANTPVTTCQ',TOBS.GetValue('QTE')*TOBS.GetValue('PVMOYENTTCQ'));
					if TOBS.GetValue('TAUXTVAQ') <> 0 then
          	TOBS.PutValue('MONTANTPVQ',TOBS.GetValue('MONTANTPVTTCQ')/TOBS.GetValue('TAUXTVAQ'))
          else
          	TOBS.PutValue('MONTANTPVQ',0);
					if TOBS.GetValue('QTE') <> 0 then
	          TOBS.PutValue('PVMOYENQ',TOBS.GetValue('MONTANTPVQ')/TOBS.GetValue('QTE'))
          else
          	TOBS.PutValue('PVMOYENQ',0);
        END;
        TOBS.PutValue('MODIFIE','X');
      END;
      ListChange.Add (TOBS);
    end;
  end;
end;

procedure ChangeValue(Change: T_SimulChg;AncValeur,NewValeur:double;TOBSimulPres:TOB; GestionInterdit : TSGestionInterdit; Global : Boolean=false);
var Indice : integer;
    ListChange : TList;
    CoefChange,ValeurCalc : double;
    TOBS,TOBAppliq : TOB;
    DejaAffiche : boolean;
    valor : T_Valeurs;
    NewCoefFc : double;
begin
  DejaAffiche := false;
  TOBAppliq := nil;
  ListChange := Tlist.create;
  if (Change = TscCoefFG) or (CHange = TscCoefMA) or (Change = TscPUA) or
  	 (Change = TscPUR) or (Change = TscPUV) or (Change = TscPUVTTC) then
  BEGIN
  	AppliqueChangement (Change,0,TOBSimulPres,ListChange,NewValeur,GestionInterdit,global);
  END else
  BEGIN
    if AncValeur <> 0 then
    	CoefChange := NewValeur/AncValeur
    else
    	CoefChange := 0;
    AppliqueChangement (Change,COefChange,TOBSimulPres,ListChange,0,GestionInterdit);
  END;
(*
  if (Change = TscCoefFA) or (CHange = Tschrs) or (CHange = TscQte) or (Change = TscPUA) or (change = TscPA) then
  begin
  end;
*)
  ValeurCalc := 0;
  // uniquement sur les montants
  if (Change <> TscCoefFG) AND (CHange <> TscCoefMA) AND
  	 (Change <> TscPUA) AND (Change <> TscPUR) AND (Change <> TscPUV) and (Change <> TscPUVTTC) and
  	 (Change <> TscHrs) and (Change <> TscQte) and (Change <> TscPVTTC) then
  BEGIN
    for Indice := 0 to ListChange.Count -1 do
    begin
      TOBS := ListChange.Items [Indice];
      if Change = TscHrs then ValeurCalc := ValeurCalc + TOBS.GetValue('HEURES');
      if (change = TscPA) then
      BEGIN
        ValeurCalc := ValeurCalc + TOBS.GetValue('MONTANTPA');
        if TOBS.GetValue ('TYPE') = 'H' Then if TOBS.GetValue('HEURES')=1 then TOBAppliq := TOBS;
        if TOBS.GetValue ('TYPE') = 'F' Then if TOBS.GetValue('QTE')=1 then TOBAppliq := TOBS;
      END;
      if (change = TscPR) then
      BEGIN
        ValeurCalc := ValeurCalc + TOBS.GetValue('MONTANTPR');
        if TOBS.GetValue ('TYPE') = 'H' Then if TOBS.GetValue('HEURES')=1 then TOBAppliq := TOBS;
        if TOBS.GetValue ('TYPE') = 'F' Then if TOBS.GetValue('QTE')=1 then TOBAppliq := TOBS;
      END;
    end;
    if (NewValeur - ValeurCalc) <> 0 then
    begin
      if (Change = TscPA) or (Change = TscPR) or (Change = TscPV) then
      BEGIN
        if TOBAppliq = nil then
        BEGIN
          if not DejaAffiche then BEGIN PGIBoxAf (SimulMessage[1],'Calcul'); DejaAffiche := true; END;
        END else
        BEGIN
          TOBS := TOBAppliq;
          if TOBS.GetValue('TYPE')='H' then
          BEGIN
            if (Change = TscPa) then
            BEGIN
              TOBS.PutVAlue('MONTANTPAH',TOBS.GetVAlue('MONTANTPAH')+(NewValeur-ValeurCalc))
            END ELSE IF (CHange = TscPR) and (TOBS.GetValue('BLOQUEPR')<>'X') then
            BEGIN
              TOBS.PutVAlue('MONTANTPRH',TOBS.GetVAlue('MONTANTPRH')+(NewValeur-ValeurCalc))
            END ELSE IF (CHange = TscPV) and (TOBS.GetValue('BLOQUEPV')<>'X') then
            BEGIN
              TOBS.PutVAlue('MONTANTPVH',TOBS.GetVAlue('MONTANTPVH')+(NewValeur-ValeurCalc))
            END ELSE
            IF (CHange = TscPVTTC) and not (TOBS.GetValue('BLOQUEPV')<>'X') then
            BEGIN
              TOBS.PutVAlue('MONTANTPVTTCH',TOBS.GetVAlue('MONTANTPVTTCH')+(NewValeur-ValeurCalc))
            END;
          END ELSE
          BEGIN
            if (Change = TscPa) then
            BEGIN
              TOBS.PutVAlue('MONTANTPAQ',TOBS.GetVAlue('MONTANTPAQ')+(NewValeur-ValeurCalc))
            END ELSE IF (CHange = TscPR) and (TOBS.GetValue('BLOQUEPR')<>'X')  then
            BEGIN
              TOBS.PutVAlue('MONTANTPRQ',TOBS.GetVAlue('MONTANTPRQ')+(NewValeur-ValeurCalc))
            END ELSE IF (CHange = TscPV) and (TOBS.GetValue('BLOQUEPV')<>'X') then
            BEGIN
              TOBS.PutVAlue('MONTANTPVQ',TOBS.GetVAlue('MONTANTPVQ')+(NewValeur-ValeurCalc))
            END ELSE IF (CHange = TscPVTTC) and (TOBS.GetValue('BLOQUEPV')<>'X') then
            BEGIN
              TOBS.PutVAlue('MONTANTPVTTCQ',TOBS.GetVAlue('MONTANTPVTTCQ')+(NewValeur-ValeurCalc))
            END;
          END;
        END;
      END;
    END;
  END;
  ListChange.Clear;
  ListChange.free;
end;

procedure ReinitMontantLigne (TOBS : TOB; GestionInterdit : TSGestionInterdit);
begin
TOBS.PutValue('HEURES',0);
TOBS.PutValue('QTE',0);
TOBS.PutValue('MONTANTPAH',0);
TOBS.PutValue('MONTANTPAFGH',0);
TOBS.PutValue('MONTANTPAFCH',0);
TOBS.PutValue('MONTANTPAFRH',0);
TOBS.PutValue('MONTANTFGH',0);
TOBS.PutValue('MONTANTFCH',0);
TOBS.PutValue('MONTANTFRH',0);
TOBS.PutValue('MONTANTPRH',0);
if (TOBS.GetValue('BLOQUEPV')<>'X') then
begin
  TOBS.PutValue('MONTANTPVH',0);
  TOBS.PutValue('MONTANTPVTTCH',0);
end;
//TOBS.PutValue('COEFFGH',0);
TOBS.PutValue('COEFFCH',0);
TOBS.PutValue('COEFMARGH',0);
TOBS.PutValue('TAUXTVAH',0);
TOBS.PutValue('MONTANTPAQ',0);
TOBS.PutValue('MONTANTPRQ',0);
TOBS.PutValue('MONTANTPAFGQ',0);
TOBS.PutValue('MONTANTPAFCQ',0);
TOBS.PutValue('MONTANTPAFRQ',0);
TOBS.PutValue('MONTANTFGQ',0);
TOBS.PutValue('MONTANTFCQ',0);
TOBS.PutValue('MONTANTFRQ',0);
//if not PasTouchePv then
if (TOBS.GetValue('BLOQUEPV')<>'X') then
begin
  TOBS.PutValue('MONTANTPVQ',0);
  TOBS.PutValue('MONTANTPVTTCQ',0);
end;
TOBS.PutValue('COEFFCQ',0);
TOBS.PutValue('COEFMARGQ',0);
TOBS.PutValue('TAUXTVAQ',0);

TOBS.PutValue('MONTANTPA',0);
TOBS.PutValue('MONTANTPR',0);
TOBS.PutValue('MONTANTPAFG',0);
TOBS.PutValue('MONTANTPAFC',0);
TOBS.PutValue('MONTANTPAFR',0);
TOBS.PutValue('MONTANTFG',0);
TOBS.PutValue('MONTANTFC',0);
TOBS.PutValue('MONTANTFR',0);
if (TOBS.GetValue('BLOQUEPV')<>'X') then
begin
  TOBS.PutValue('MONTANTPV',0);
  TOBS.PutValue('MONTANTPVTTC',0);
end;
TOBS.PutValue('COEFFC',0);
TOBS.PutValue('COEFMARG',0);
TOBS.PutValue('TAUXTVA',0);
end;

procedure ReinitValeur (Change: T_SimulChg;TobSimulPres : TOB; GestionInterdit : TSGestionInterdit);
var Indice : integer;
    TOBS : TOB;
begin
for Indice := 0 to TOBSimulPres.detail.count -1 do
    begin
    TOBS := TOBSimulPres.detail[Indice];
    if TOBS.GetValue('NIVEAU') = 5 then break;
    if TOBS.detail.count > 0 then ReinitValeur (Change,TOBS,GestionInterdit);
    ReinitMontantLigne (TOBS,GestionInterdit);
    end;
ReinitMontantLigne (TOBSimulPres,GestionInterdit);
end;

procedure FinaliseCalcul(TobSimulPres : TOB);
var Valor: T_valeurs;
begin
InitTableau (valor);
FinaliseSimulation (TOBSImulPres,valor);
end;

procedure TraiteModification (TOBSimulPres,TOBPiece,TOBOuvrage,TOBMOD : TOB;DEV : Rdevise; GestionInterdit : TSGestionInterdit);
var Indice,IndiceO : Integer;
    TOBS,TOBL,TOBO,TOBM : TOB;
    valeurs : T_Valeurs;
    EnHt : boolean;
    Qte : double;
    MontantCharge : double;
begin
  EnHT:=(TOBPiece.GetValue('GP_FACTUREHT')='X') ;
  for Indice := 0 To TOBSimulPres.detail.count -1 do
  begin
    TOBS := TOBSimulPres.detail[Indice];
    if TOBS.GetValue('NIVEAU') <> 5 then TraiteModification (TOBS,TOBPiece,TOBOuvrage,TOBMOD,DEV,GestionInterdit);
    if (TOBS.GetValue('NIVEAU')=5) and ((TOBS.GetValue('MODIFIE')='X') OR (TOBS.GetValue('COEFMAMODIFIE')='X')) then
    begin
      TOBL := TOB(TOBS.data);
      if TOBL.NomTable = 'LIGNE' Then
      BEGIN
        if TOBS.GetValue('TYPE')= 'H' Then
        BEGIN
          if (TOBS.GetValue('MODIFIE')='X')  then
          begin
            TOBL.PutValue('GL_MONTANTPA',TOBS.GetVAlue('MONTANTPAH'));
            TOBL.PutValue('GL_MONTANTPAFG',TOBS.GetVAlue('MONTANTPAFGH'));
            TOBL.PutValue('GL_MONTANTPAFC',TOBS.GetVAlue('MONTANTPAFCH'));
            TOBL.PutValue('GL_MONTANTPAFR',TOBS.GetVAlue('MONTANTPAFRH'));
            TOBL.PutValue('GL_MONTANTFG',TOBS.GetVAlue('MONTANTFGH'));
            TOBL.PutValue('GL_MONTANTFC',TOBS.GetVAlue('MONTANTFCH'));
            TOBL.PutValue('GL_MONTANTPR',TOBS.GetVAlue('MONTANTPRH'));
            if TOBL.GetValue('GL_MONTANTPAFG') <> 0 then TOBL.PutValue('GL_COEFFG',TOBL.GetValue('GL_MONTANTFG')/TOBL.GetValue('GL_MONTANTPAFG'))
                                                    else TOBL.PutValue('GL_COEFFG',0);
            if TOBL.GetValue('GL_MONTANTPAFC') <> 0 then TOBL.PutValue('GL_COEFFC',TOBL.GetValue('GL_MONTANTFC')/TOBL.GetValue('GL_MONTANTPAFC'))
                                                    else TOBL.PutValue('GL_COEFFC',0);
            //
            if TOBS.GetValue('HEURES') <> 0 then
            begin
              TOBL.PutValue('GL_DPA',arrondi(TOBS.GetVAlue('MONTANTPAH')/TOBS.GetValue('HEURES'),V_PGI.okDECP));
              TOBL.PutValue('GL_DPR',arrondi(TOBS.GetVAlue('MONTANTPRH')/TOBS.GetValue('HEURES'),V_PGI.okDECP));
            end else
            begin
              TOBL.PutValue('GL_DPA',0);
              TOBL.PutValue('GL_DPR',0);
            end;
          end;
          if (TOBS.GetValue('BLOQUEPV')<>'X') then
          begin
            if (TOBS.GetValue('MODIFIE')='X') then
            begin
              if (EnHT = True) then
              begin
                if TOBS.GetValue('HEURES') <> 0 then
                  TOBL.PutValue('GL_PUHTDEV',arrondi(TOBS.GetVAlue('MONTANTPVH')/TOBS.GetValue('HEURES'),V_PGI.okDECP))
                else
                  TOBL.PutValue('GL_PUHTDEV',0);
                TOBL.PutValue('GL_PUHT',DevisetoPivotEx(TOBL.GetVAlue('GL_PUHTDEV'),TOBL.GetValue('GL_TAUXDEV'),DEV.Quotite,V_PGI.okdecP));
              end else
              begin
                if TOBS.GetValue('HEURES') <> 0 then
                  TOBL.PutValue('GL_PUTTCDEV',arrondi(TOBS.GetVAlue('MONTANTPVH')/TOBS.GetValue('HEURES'),V_PGI.okDECP))
                else
                  TOBL.PutValue('GL_PUTTCDEV',0);
                TOBL.PutValue('GL_PUTTC',DevisetoPivotEx(TOBL.GetVAlue('GL_PUTTCDEV'),TOBL.GetValue('GL_TAUXDEV'),DEV.Quotite,V_PGI.OkdecP));
              end;
              TOBL.putValue('GL_COEFMARG',TOBS.GetValue('COEFMARGH'));
              TOBL.PutValue('POURCENTMARG',Arrondi((TOBL.GetValue('GL_COEFMARG')-1)*100,2));
              if TOBL.GetValue('GL_PUHT') <> 0 then
              begin
                TOBL.PutValue('POURCENTMARQ',Arrondi(((TOBL.GetValue('GL_PUHT')- TOBL.GetValue('GL_DPR'))/TOBL.GetValue('GL_PUHT'))*100,2));
              end else
              begin
                TOBL.PutValue('POURCENTMARQ',0);
              end;
            end else if (TOBS.GetValue('COEFMAMODIFIE')='X') then
            begin
          		TOBL.putValue('GL_COEFMARG',TOBS.GetValue('COEFMARGH'));
              TOBL.PutValue('POURCENTMARG',Arrondi((TOBL.GetValue('GL_COEFMARG')-1)*100,2));
              TOBL.PutValue('GL_PUHT', Arrondi(TOBL.GetValue('GL_DPR') * TOBL.GetValue('GL_COEFMARG'),V_PGI.OkdecP));
              TOBL.PutValue('GL_PUHTDEV',pivottodevise(TobL.GetValue('GL_PUHT'),DEV.Taux,DEV.quotite,V_PGI.okdecP ));
              if TOBL.GetValue('GL_PUHT') <> 0 then
              begin
                TOBL.PutValue('POURCENTMARQ',Arrondi(((TOBL.GetValue('GL_PUHT')- TOBL.GetValue('GL_DPR'))/TOBL.GetValue('GL_PUHT'))*100,2));
              end else
              begin
                TOBL.PutValue('POURCENTMARQ',0);
              end;
            end;
          end else
          begin
          	TOBL.putValue('GL_COEFMARG',0);
          end;
          if (TOBS.GetValue('MODIFIE')='X') then
          begin
            TOBL.PutValue('GL_QTEFACT',TOBS.GetVAlue('HEURES'));
            //
            if TOBL.fieldExists ('GF_PRIXUNITAIRE') then
            begin
              TOBL.Putvalue ('GF_PRIXUNITAIRE',TOBL.GetValue('GL_PUHTDEV')) ;
              TOBL.PutValue('GA_HEURE',TOBL.GetVAlue('GL_QTEFACT'));
            end;
          end;
          TOBM := TOBMod.FindFirst (['NUMLIGNE'],[TOBS.GEtVAlue('NUMLIGNE')],true);
          if TOBM = nil then
          begin
            TOBM := TOB.Create ('UNE LIGNE MODIF',TOBMOd,-1);
            TOBM.AddChampSupValeur ('NUMLIGNE',TOBL.GEtVAlue('GL_NUMLIGNE'),false);
          end;
          //
        END ELSE
        BEGIN
        	if (TOBS.GetValue('MODIFIE')='X') then
          begin
            TOBL.PutValue('GL_MONTANTPA',TOBS.GetVAlue('MONTANTPAQ'));
            TOBL.PutValue('GL_MONTANTPAFG',TOBS.GetVAlue('MONTANTPAFGQ'));
            TOBL.PutValue('GL_MONTANTPAFC',TOBS.GetVAlue('MONTANTPAFCQ'));
            TOBL.PutValue('GL_MONTANTPAFR',TOBS.GetVAlue('MONTANTPAFRQ'));
            TOBL.PutValue('GL_MONTANTFG',TOBS.GetVAlue('MONTANTFGQ'));
            TOBL.PutValue('GL_MONTANTFC',TOBS.GetVAlue('MONTANTFCQ'));
            TOBL.PutValue('GL_MONTANTPR',TOBS.GetVAlue('MONTANTPRQ'));
            //
            if TOBL.GetValue('GL_MONTANTPAFG') <> 0 then TOBL.PutValue('GL_COEFFG',TOBL.GetValue('GL_MONTANTFG')/TOBL.GetValue('GL_MONTANTPAFG'))
                                                    else TOBL.PutValue('GL_COEFFG',0);
            if TOBL.GetValue('GL_MONTANTPAFC') <> 0 then TOBL.PutValue('GL_COEFFC',TOBL.GetValue('GL_MONTANTFC')/TOBL.GetValue('GL_MONTANTPAFC'))
                                                    else TOBL.PutValue('GL_COEFFC',0);
            //
            if TOBS.GetValue('QTE') <> 0 then
            Begin
              TOBL.PutValue('GL_DPA',arrondi(TOBS.GetVAlue('MONTANTPAQ')/TOBS.GetValue('QTE'),V_PGI.okDECP));
              TOBL.PutValue('GL_DPR',arrondi(TOBS.GetVAlue('MONTANTPRQ')/TOBS.GetValue('QTE'),V_PGI.okDECP));
            end else if (TOBS.GetValue('COEFMAMODIFIE')='X') then
            begin
              TOBL.PutValue('GL_DPA',0);
              TOBL.PutValue('GL_DPR',0);
            end;
          end;
          if (TOBS.GetValue('BLOQUEPV')<>'X') then
          begin
          	if (TOBS.GetValue('MODIFIE')='X') then
            begin
              if EnHT = True then
              begin
                if TOBS.GetValue('QTE') <> 0 then
                  TOBL.PutValue('GL_PUHTDEV',arrondi(TOBS.GetVAlue('MONTANTPVQ')/TOBS.GetValue('QTE'),V_PGI.okDECP))
                else
                  TOBL.PutValue('GL_PUHTDEV',0);
                TOBL.PutValue('GL_PUHT',DevisetoPivotEx(TOBL.GetVAlue('GL_PUHTDEV'),TOBL.GetValue('GL_TAUXDEV'),DEV.Quotite,V_PGI.okdecP));
              end else
              begin
                if TOBS.GetValue('QTE') <> 0 then
                  TOBL.PutValue('GL_PUTTCDEV',arrondi(TOBS.GetVAlue('MONTANTPVTTCQ')/TOBS.GetValue('QTE'),V_PGI.okDECP))
                else
                  TOBL.PutValue('GL_PUTTCDEV',0);
                 TOBL.PutValue('GL_PUTTC',DevisetoPivotEx(TOBL.GetVAlue('GL_PUTTCDEV'),TOBL.GetValue('GL_TAUXDEV'),DEV.Quotite,V_PGI.OkdecP));
              end;
              TOBL.putValue('GL_COEFMARG',TOBS.GetValue('COEFMARGQ'));
              TOBL.PutValue('POURCENTMARG',Arrondi((TOBL.GetValue('GL_COEFMARG')-1)*100,2));
              if TOBL.GetValue('GL_PUHT') <> 0 then
              begin
                TOBL.PutValue('POURCENTMARQ',Arrondi(((TOBL.GetValue('GL_PUHT')- TOBL.GetValue('GL_DPR'))/TOBL.GetValue('GL_PUHT'))*100,2));
              end else
              begin
                TOBL.PutValue('POURCENTMARQ',0);
              end;
            end else if (TOBS.GetValue('COEFMAMODIFIE')='X') then
            begin
              TOBL.putValue('GL_COEFMARG',TOBS.GetValue('COEFMARGQ'));
              TOBL.PutValue('POURCENTMARG',Arrondi((TOBL.GetValue('GL_COEFMARG')-1)*100,2));
              TOBL.PutValue('GL_PUHT', Arrondi(TOBL.GetValue('GL_DPR') * TOBL.GetValue('GL_COEFMARG'),V_PGI.OkdecP));
              TOBL.PutValue('GL_PUHTDEV',pivottodevise(TobL.GetValue('GL_PUHT'),DEV.Taux,DEV.quotite,V_PGI.okdecP ));
              if TOBL.GetValue('GL_PUHT') <> 0 then
              begin
                TOBL.PutValue('POURCENTMARQ',Arrondi(((TOBL.GetValue('GL_PUHT')- TOBL.GetValue('GL_DPR'))/TOBL.GetValue('GL_PUHT'))*100,2));
              end else
              begin
                TOBL.PutValue('POURCENTMARQ',0);
              end;
            end;
          end else
          begin
          	TOBL.putValue('GL_COEFMARG',0);
          end;
          if (TOBS.GetValue('MODIFIE')='X') then
          begin
            TOBL.PutValue('GL_QTEFACT',TOBS.GetVAlue('QTE'));
            //
            if TOBL.FieldExists ('GA_PAHT') then
            begin
              TOBL.PutValue('GA_PAHT',TOBL.GetVAlue('GL_DPA'));
              TOBL.PutValue('GA_PVHT',TOBL.GetVAlue('GL_PUHTDEV'));
            end;
            //
          end;
          TOBM := TOBMod.FindFirst (['NUMLIGNE'],[TOBS.GEtVAlue('NUMLIGNE')],true);
          if TOBM = nil then
          begin
            TOBM := TOB.Create ('UNE LIGNE MODIF',TOBMOd,-1);
            TOBM.AddChampSupValeur ('NUMLIGNE',TOBL.GEtVAlue('GL_NUMLIGNE'),false);
          end;
        END;
        TOBL.PutValue('GL_RECALCULER','X');
      END;
      if TOBL.NomTable = 'LIGNEOUV' Then
      BEGIN
        if TOBS.GetValue('TYPE')= 'H' Then
        BEGIN
          if TOBS.GetValue('RATIOQTE') <> 0 then
          	Qte := arrondi(TOBS.GetVAlue('HEURES')/TOBS.GetValue('RATIOQTE'),V_PGI.OKDecQ)
          else
          	Qte := 0;
          if TOBS.GetVAlue('HEURES') <> 0 then
          begin
            if (TOBS.GetValue('MODIFIE')='X') then
            begin
              TOBL.PutValue('BLO_MONTANTPA',TOBS.GetVAlue('MONTANTPAH'));
              TOBL.PutValue('BLO_MONTANTPAFG',TOBS.GetVAlue('MONTANTPAFGH'));
              TOBL.PutValue('BLO_MONTANTPAFC',TOBS.GetVAlue('MONTANTPAFCH'));
              TOBL.PutValue('BLO_MONTANTPAFR',TOBS.GetVAlue('MONTANTPAFRH'));
              TOBL.PutValue('BLO_MONTANTFG',TOBS.GetVAlue('MONTANTFGH'));
              TOBL.PutValue('BLO_MONTANTFC',TOBS.GetVAlue('MONTANTFCH'));
              TOBL.PutValue('BLO_MONTANTPR',TOBS.GetVAlue('MONTANTPRH'));
              //
              if TOBL.GetValue('BLO_MONTANTPAFG') <> 0 then TOBL.PutValue('BLO_COEFFG',TOBL.GetValue('BLO_MONTANTFG')/TOBL.GetValue('BLO_MONTANTPAFG'))
                                                      else TOBL.PutValue('BLO_COEFFG',0);
              if TOBL.GetValue('BLO_MONTANTPAFC') <> 0 then TOBL.PutValue('BLO_COEFFC',TOBL.GetValue('BLO_MONTANTFC')/TOBL.GetValue('BLO_MONTANTPAFC'))
                                                      else TOBL.PutValue('BLO_COEFFC',0);
              //
              TOBL.PutValue('BLO_DPA',arrondi(TOBS.GetVAlue('MONTANTPAH')/TOBS.GetVAlue('HEURES'),V_PGI.okDECP));
              TOBL.PutValue('BLO_DPR',arrondi(TOBS.GetVAlue('MONTANTPRH')/TOBS.GetVAlue('HEURES'),V_PGI.okDECP));
            end;
            if (TOBS.GetValue('BLOQUEPV')<>'X') then
            begin
            	if (TOBS.GetValue('MODIFIE')='X') then
          		begin
                if EnHT = True then
                begin
                  TOBL.PutValue('BLO_PUHTDEV',arrondi(TOBS.GetVAlue('MONTANTPVH')/TOBS.GetVAlue('HEURES'),V_PGI.okDECP));
                  TOBL.PutValue('BLO_PUHT',DevisetoPivotEx(TOBL.GetVAlue('BLO_PUHTDEV'),TOBL.GetValue('BLO_TAUXDEV'),DEV.Quotite,V_PGI.okdecP));
                end else
                begin
                  TOBL.PutValue('BLO_PUTTCDEV',arrondi(TOBS.GetVAlue('MONTANTPVTTCH')/TOBS.GetVAlue('HEURES'),V_PGI.okDECP));
                  TOBL.PutValue('BLO_PUTTC',DevisetoPivotEx(TOBL.GetVAlue('BLO_PUTTCDEV'),TOBL.GetValue('BLO_TAUXDEV'),DEV.Quotite,V_PGI.okdecP));
                end;
                TOBL.putValue('BLO_COEFMARG',TOBS.GetValue('COEFMARGH'));
                TOBL.PutValue('POURCENTMARG',Arrondi((TOBL.GetValue('BLO_COEFMARG')-1)*100,2));
                if TOBL.GetValue('BLO_PUHT') <> 0 then
                begin
                  TOBL.PutValue('POURCENTMARQ',Arrondi(((TOBL.GetValue('BLO_PUHT')- TOBL.GetValue('BLO_DPR'))/TOBL.GetValue('BLO_PUHT'))*100,2));
                end else
                begin
                  TOBL.PutValue('POURCENTMARQ',0);
                end;
              end else if (TOBS.GetValue('COEFMAMODIFIE')='X') then
              begin
            		TOBL.putValue('BLO_COEFMARG',TOBS.GetValue('COEFMARGH'));
                TOBL.PutValue('POURCENTMARG',Arrondi((TOBL.GetValue('BLO_COEFMARG')-1)*100,2));
                TOBL.PutValue('BLO_PUHTDEV',arrondi(TOBL.GetVAlue('BLO_DPR')*TOBL.GetVAlue('BLO_COEFMARG'),V_PGI.okDECP));
                TOBL.PutValue('BLO_PUHT',DevisetoPivotEx(TOBL.GetVAlue('BLO_PUHTDEV'),TOBL.GetValue('BLO_TAUXDEV'),DEV.Quotite,V_PGI.okdecP));
              	TOBL.Putvalue('GF_PRIXUNITAIRE',TOBL.GetValue('BLO_PUHTDEV'));
                if TOBL.GetValue('BLO_PUHT') <> 0 then
                begin
                  TOBL.PutValue('POURCENTMARQ',Arrondi(((TOBL.GetValue('BLO_PUHT')- TOBL.GetValue('BLO_DPR'))/TOBL.GetValue('BLO_PUHT'))*100,2));
                end else
                begin
                  TOBL.PutValue('POURCENTMARQ',0);
                end;
              end;
						end else
            begin
            	TOBL.putValue('BLO_COEFMARG',0);
            end;
            if (TOBS.GetValue('MODIFIE')='X') then
            begin
              TOBL.PutValue('BLO_QTEFACT',Qte);
            end;
          end;
          if (TOBS.GetValue('MODIFIE')='X') then
          begin
            // Moif LS le 31/03/03
            if TOBL.fieldExists ('GF_PRIXUNITAIRE') then
            begin
              TOBL.PutValue('GA_HEURE',Qte);
              TOBL.Putvalue('GF_PRIXUNITAIRE',TOBL.GetValue('BLO_PUHTDEV'));
              TOBL.Putvalue('GA_PAHT',TOBL.GetValue('BLO_DPA'));
            END;
          end;
          //
          TOBM := TOBMod.FindFirst (['NUMLIGNE'],[TOBS.GEtVAlue('NUMLIGNE')],true);
          if TOBM = nil then
          begin
            TOBM := TOB.Create ('UNE LIGNE MODIF',TOBMOd,-1);
            TOBM.AddChampSupValeur ('NUMLIGNE',TOBS.GEtVAlue('NUMLIGNE'),false);
          end;
        END ELSE
        BEGIN
          if TOBS.GetValue('RATIOQTE') <> 0 then
	          Qte := arrondi(TOBS.GetVAlue('QTE')/TOBS.GetValue('RATIOQTE'),V_PGI.OKDecQ)
          else
          	Qte := 0;
          if TOBS.GetValue('QTE') <>  0 then
          begin
						if (TOBS.GetValue('MODIFIE')='X') then
            begin
              TOBL.PutValue('BLO_MONTANTPA',TOBS.GetVAlue('MONTANTPAQ'));
              TOBL.PutValue('BLO_MONTANTPAFG',TOBS.GetVAlue('MONTANTPAFGQ'));
              TOBL.PutValue('BLO_MONTANTPAFC',TOBS.GetVAlue('MONTANTPAFCQ'));
              TOBL.PutValue('BLO_MONTANTPAFR',TOBS.GetVAlue('MONTANTPAFRQ'));
              TOBL.PutValue('BLO_MONTANTFG',TOBS.GetVAlue('MONTANTFGQ'));
              TOBL.PutValue('BLO_MONTANTFC',TOBS.GetVAlue('MONTANTFCQ'));
              TOBL.PutValue('BLO_MONTANTPR',TOBS.GetVAlue('MONTANTPRQ'));
              //
              if TOBL.GetValue('BLO_MONTANTPAFG') <> 0 then TOBL.PutValue('BLO_COEFFG',TOBL.GetValue('BLO_MONTANTFG')/TOBL.GetValue('BLO_MONTANTPAFG'))
                                                      else TOBL.PutValue('BLO_COEFFG',0);
              if TOBL.GetValue('BLO_MONTANTPAFC') <> 0 then TOBL.PutValue('BLO_COEFFC',TOBL.GetValue('BLO_MONTANTFC')/TOBL.GetValue('BLO_MONTANTPAFC'))
                                                      else TOBL.PutValue('BLO_COEFFC',0);
              //
              TOBL.PutValue('BLO_DPA',arrondi(TOBS.GetVAlue('MONTANTPAQ')/TOBS.GetVAlue('QTE'),V_PGI.okDECP));
              TOBL.PutValue('BLO_DPR',arrondi(TOBS.GetVAlue('MONTANTPRQ')/TOBS.GetVAlue('QTE'),V_PGI.okDECP));
            end;
            if (TOBS.GetValue('BLOQUEPV')<>'X') then
            begin
              if (TOBS.GetValue('MODIFIE')='X') then
              begin
                if EnHT = True then
                begin
                  TOBL.PutValue('BLO_PUHTDEV',arrondi(TOBS.GetVAlue('MONTANTPVQ')/TOBS.GetVAlue('QTE'),V_PGI.okDECP));
                  TOBL.PutValue('BLO_PUHT',DevisetoPivotEx(TOBL.GetVAlue('BLO_PUHTDEV'),TOBL.GetValue('BLO_TAUXDEV'),DEV.Quotite,V_PGI.okdecP));
                end else
                begin
                  TOBL.PutValue('BLO_PUTTCDEV',arrondi(TOBS.GetVAlue('MONTANTPVTTCQ')/TOBS.GetVAlue('QTE'),V_PGI.okDECP));
                  TOBL.PutValue('BLO_PUTTC',DevisetoPivotEx(TOBL.GetVAlue('BLO_PUTTCDEV'),TOBL.GetValue('BLO_TAUXDEV'),DEV.Quotite,V_PGI.okdecP));
                end;
                TOBL.putValue('BLO_COEFMARG',TOBS.GetValue('COEFMARGQ'));
                TOBL.PutValue('POURCENTMARG',Arrondi((TOBL.GetValue('BLO_COEFMARG')-1)*100,2));
                if TOBL.GetValue('BLO_PUHT') <> 0 then
                begin
                  TOBL.PutValue('POURCENTMARQ',Arrondi(((TOBL.GetValue('BLO_PUHT')- TOBL.GetValue('BLO_DPR'))/TOBL.GetValue('BLO_PUHT'))*100,2));
                end else
                begin
                  TOBL.PutValue('POURCENTMARQ',0);
                end;
              end else if (TOBS.GetValue('COEFMAMODIFIE')='X') then
              begin
            		TOBL.putValue('BLO_COEFMARG',TOBS.GetValue('COEFMARGQ'));
                TOBL.PutValue('POURCENTMARG',Arrondi((TOBL.GetValue('BLO_COEFMARG')-1)*100,2));
                TOBL.PutValue('BLO_PUHTDEV',arrondi(TOBL.GetVAlue('BLO_DPR')*TOBL.GetVAlue('BLO_COEFMARG'),V_PGI.okDECP));
                TOBL.PutValue('BLO_PUHT',DevisetoPivotEx(TOBL.GetVAlue('BLO_PUHTDEV'),TOBL.GetValue('BLO_TAUXDEV'),DEV.Quotite,V_PGI.okdecP));
              	TOBL.Putvalue('GF_PRIXUNITAIRE',TOBL.GetValue('BLO_PUHTDEV'));
                if TOBL.GetValue('BLO_PUHT') <> 0 then
                begin
                  TOBL.PutValue('POURCENTMARQ',Arrondi(((TOBL.GetValue('BLO_PUHT')- TOBL.GetValue('BLO_DPR'))/TOBL.GetValue('BLO_PUHT'))*100,2));
                end else
                begin
                  TOBL.PutValue('POURCENTMARQ',0);
                end;
              end;
            end else
            begin
            	TOBL.putValue('BLO_COEFMARG',0);
            end;
            if (TOBS.GetValue('MODIFIE')='X') then
            begin
              TOBL.PutValue('BLO_QTEFACT',Qte);
            end;
          end;
          TOBM := TOBMod.FindFirst (['NUMLIGNE'],[TOBS.GEtVAlue('NUMLIGNE')],true);
          if TOBM = nil then
          begin
            TOBM := TOB.Create ('UNE LIGNE MODIF',TOBMOd,-1);
            TOBM.AddChampSupValeur ('NUMLIGNE',TOBS.GEtVAlue('NUMLIGNE'),false);
          end;
          if (TOBL.GetValue('MODIFIE')='X') then
          begin
            // Modif LS le 31/03/03
            if TOBL.FieldExists ('GA_PAHT') then
            begin
              TOBL.Putvalue('GA_PAHT',TOBL.GetValue('BLO_DPA'));
              TOBL.Putvalue('GA_PVHT',TOBL.GetValue('BLO_PUHTDEV'));
            end;
          end;
        END;
      END;
    end;
  end;
  //
end;

procedure RecalculeOuvrage (TOBPiece,TOBL,TOBOuvrage,TOBBases,TOBBasesL : TOB; DEV : Rdevise);

	procedure repriseValo (TOBOuv : TOB);
  var TOBO : TOB;
  		Indice : integer;
  begin
    for Indice := 0 to TOBOuv.detail.count -1 do
    begin
    	TOBO := TOBOUv.detail[Indice];
      if TOBO.detail.count > 0 then repriseValo(TOBO);
      GetValoDetail (TOBO);
    end;
  end;
var
		TOBM,TOBO : TOB;
    EnHt : boolean;
    valeurs : T_Valeurs;
    SavModeBloque : string;
    MontantPv : double;
begin
  EnHT:=(TOBPiece.GetValue('GP_FACTUREHT')='X') ;
  if Pos(TOBL.GetValue('GL_TYPEARTICLE'),'OUV;ARP;')=0 then exit;
  if TOBL.GetValue('GL_INDICENOMEN') = 0 then exit;
  TOBO  := TOBOuvrage.detail[TOBL.GetValue('GL_INDICENOMEN')-1];
  InitTableau (Valeurs);
  CalculeOuvrageDoc (TOBO,1,1,true,DEV,valeurs,EnHt);
  GetValoDetail (TOBO); // pour le cas des Article en prix posés
  if (TOBL.GetValue('GL_BLOQUETARIF')<>'X') and (TOBL.GetValue('GLC_FROMBORDEREAU')<>'X') then
  begin
    TOBL.Putvalue('GL_PUHT',DeviseToPivotEx(valeurs[2],DEV.Taux,DEV.quotite,V_PGI.okdecP));
    TOBL.Putvalue('GL_PUTTC',DevisetoPivotEx(valeurs[3],DEV.taux,DEV.quotite,V_PGI.okdecP));
    TOBL.Putvalue('GL_PUHTDEV',valeurs[2]);
    TOBL.Putvalue('GL_PUTTCDEV',valeurs[3]);
    TOBL.Putvalue('GL_PUHTBASE',TOBL.GetValue('GL_PUHT'));
    TOBL.Putvalue('GL_PUTTCBASE',TOBL.GetValue('GL_PUTTC'));
  end else
  begin
    if TOBL.GetString('GL_FACTUREHT')='X' then
    begin
      MontantPv := TOBL.getValue('GL_PUHTDEV');
      SavModeBloque := TOBL.GetString('GL_BLOQUETARIF');
      TOBL.PutValue('GL_BLOQUETARIF','-'); // Pour forcer le recalcul des Pv sur le sous détail
      TraitePrixOuvrage (TOBPiece,TOBL,TobBases,TOBBasesL,TOBOuvrage,EnHt, TOBL.GetValue('GL_PUHTDEV'),Valeurs[2],DEV,false,true);
      TOBL.PutValue('GL_BLOQUETARIF',SavModeBloque);
    end else
    begin
      MontantPv := TOBL.getValue('GL_PUTTCDEV');
      SavModeBloque := TOBL.GetString('GL_BLOQUETARIF');
      TOBL.PutValue('GL_BLOQUETARIF','-'); // Pour forcer le recalcul des Pv sur le sous détail
      TraitePrixOuvrage (TOBPiece,TOBL,TobBases,TOBBasesL,TOBOuvrage,EnHt,TOBL.GetValue('GL_PUTTCDEV'),Valeurs[3],DEV,False,True);
      TOBL.PutValue('GL_BLOQUETARIF',SavModeBloque);
    end;
  end;
  TOBL.Putvalue('GL_PUHTBASE',TOBL.GetValue('GL_PUHT'));
  TOBL.Putvalue('GL_PUTTCBASE',TOBL.GetValue('GL_PUTTC'));
  TOBL.Putvalue('GL_DPA',valeurs[0]);
  TOBL.Putvalue('GL_DPR',valeurs[1]);
  TOBL.Putvalue('GL_TPSUNITAIRE',valeurs[9]);
  //
  TOBL.Putvalue('GL_MONTANTPAFG',valeurs[10]);
  TOBL.Putvalue('GL_MONTANTPAFR',valeurs[11]);
  TOBL.Putvalue('GL_MONTANTPAFC',valeurs[12]);
  TOBL.Putvalue('GL_MONTANTFG',valeurs[13]);
  TOBL.Putvalue('GL_MONTANTFR',valeurs[14]);
  TOBL.Putvalue('GL_MONTANTFC',valeurs[15]);
  TOBL.Putvalue('GL_MONTANTPA',valeurs[16]);
  TOBL.Putvalue('GL_MONTANTPR',valeurs[17]);
//    if valeurs[1] <> 0 then TOBL.PutValue('GL_COEFMARG',Arrondi(Valeurs[2]/valeurs[1],4));
  StockeInfoTypeLigne (TOBL,valeurs);
  TOBL.PutValue('GL_RECALCULER','X');
end;

procedure RecalcSuiteModif (TOBPiece,TOBOUvrage,TOBMod,TOBBases,TOBBasesL : TOB; DEV : Rdevise );

var Indice : integer;
		TOBL,TOBM,TOBO : TOB;
    EnHt : boolean;
    valeurs : T_Valeurs;
    Ligne : integer;
    SavModeBloque : string;
    MontantPv : double;
begin
  EnHT:=(TOBPiece.GetValue('GP_FACTUREHT')='X') ;
  for Indice := 0 to TOBMod.detail.count -1 do
  begin
    TOBM := TOBMOD.detail[indice];
    Ligne := TOBM.GEtValue('NUMLIGNE'); if Ligne = 0 then continue;
    TOBL := TOBPiece.detail[Ligne-1];
  	if TOBL.getValue('GL_TYPELIGNE')<>'ART' Then Continue;
    RecalculeOuvrage (TOBPiece,TOBL,TOBOUvrage,TOBBases,TOBBasesL,DEV);
    TOBL.PutValue('GL_RECALCULER','X');
  end;
  TOBPiece.putValue('GP_RECALCACHAT','X'); // on force le recalcul des achats
end;

procedure DefiniOuvert (TOBL : TOB);
begin
if TOBL.GetValue('NIVEAU') < 5 then TOBL.PutValue('OPEN','X');
if (TOBL.GetValue('NIVEAU') > 0) and (TOBL.Parent <> nil)  then DefiniOuvert (TOBL.Parent);
end;

procedure Renumerote (TOBSimulPres : TOB; var Ligne : integer; var niveau : integer);
var TOBL : TOB;
    Indice : integer;
begin
for Indice := 0 to TOBSimulPres.detail.count -1 do
    begin
    TOBL := TOBSimulPres.detail[Indice];
    if TOBL.GetValue('NIVEAU') > niveau then Niveau := TOBL.GetValue('NIVEAU');
    inc (Ligne);
    TOBL.PutValue('NUMAFF',Ligne);
    if TOBL.GetValue('OPEN')='X' then Renumerote (TOBL,Ligne,Niveau);
    end;
end;

end.
