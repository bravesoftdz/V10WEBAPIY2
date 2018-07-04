unit AfPrepFact;
{
   Préparation automatique des  factures.
   Origine : UtofafPrepFact
   Input : TobRegAff : contient 1 ou n affaires donnant lieu à 1 facture.

   Principe :
   * Pour la sauce affaire j'utilise les tob   (tobpiece ,toblignes)
   ===> 1ere phase : on charge 'en vrac' dans la toblignes  tout les composants (lignes)
   de la facture  (TobPiece)
   ===> 2eme phase, traitement des lignes pour obtenir la présentation souhaitée (issue
   de PROFILGENER). dans cette phase les tobs  deviennent toblignes_OK  et tobpiece_OK
   ===> 3eme phase
   On géré une tob de cumul, tobpiece_CUM   avant d'alimenter la tob définitive
   topPiece_OK
   validation via les fcts de JLD  avec Tobpiece_OK

   DIVERS :
   Les article de % seront toujours "en fin d'affaire" (sauf si on a un Sous-total par
   affaire, dans ce cas il sera juste en dessus du Sous_total).
   Pour l'instant, on ne traite qu'un article de % par affaire, pas d'article en caascade)

   // MODIFS.........
   //04/02 gm pour les article de type contrat : N00,n10... et XPO pour Pourcentage
   // le % sera toujours à la fin
}

interface

uses  StdCtrls,Controls,Classes, forms,sysutils,ComCtrls,paramsoc,
      HCtrls,HEnt1,HMsgBox,UTOF, AffaireUtil,Grids,EntGC,
      utilarticle,FactComm, FactCalc,Facture,UtilPGI,UtilGC,
      SaisUtil,Ent1,Formule,dicobtp,tiersutil,AffEcheanceUtil,AppreciationUtil,
      UAFO_REVPRIXLECTURECOEF,factaffaire,facttiers,utilgrp,factcpta,
{$IFDEF EAGLCLIENT}

{$ELSE}
      {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} db,
{$ENDIF}
      FactUtil,UTob, UtilTarif{,voirtob}, UtilPrepFact,factpiece,facttob,factarticle,factadresse,uEntCommun,UtilTOBPiece;


Type R_GENFAC = RECORD
     zdatedeb,zdatefin,zcomfac,zacompte,zlibacompte,zcomaff,zcomaffsup,zlibstotcli: String;
     zdateact_d,zdateact_f : string;
     zgen,zprof,ztypgen,ztypech: String;
     zcumaff : boolean;
     zpres,zfrais,zfour,zctr : Integer;
     zstot,zcomlig : boolean;
     zlibpres,zlibfrais,zlibfour,zlibctr,zlibempcum,zlibempcom,zlibemppre : String;
     zlibcomsaff,zlibcumsaff : String;
     zminifrais,zmaxifrais : double;
     zminifour,zmaxifour : double;
     zreplibech : string;
     zcomsaff,zcumsaff,zcumemp,zcomemp,zemppre,zacpte,ztout : boolean;
     zcom1,zcom2,ZrefInt : Hstring;
     znat : string;
     zpour,zstotcli : boolean;

END ;

Type R_ECH = RECORD
     NumEch,NumEchBis : integer;
     MontEch,PourEch : double;
     DateEch,RepriseActiv: string;
     Libelleeche,profil : String;
End;

Const DER_RESS        : String = 'ZZZZZZZZZZZZZZZZZ' ;

Function ChargeTobProfil(TobProf : TOB) : boolean;


procedure InitToutGenfac(GenFac : R_GENFAC);
procedure  AFPrepFact_init(ZTobRegAff,ZTobProfil,ZtobFormuleVar,ZTobArticles : TOB ;  GenFac : R_GENFAC;znb_aff : Integer;
                           zdatech  : string; ZAvecTrans : boolean;
                           var znb_fac,zfac_deb,zfac_fin : Integer;
                           zRevLogFile : String;DateFacture : TDateTime);

Type TAFPrepFact = class
Public
    Constructor create(ZTobRegAff,ZTobProfil ,ZtobFormuleVar,ZTobArticles: TOB ;  ZGenFac : R_GENFAC;znb_aff : Integer;
    zdatech :  string;  ZAvecTrans : boolean ;
    var znb_fac,zfac_deb,zfac_fin  : Integer;
    zRevLogFile ,zsouche: string; zRevForcer: Integer);
    Destructor  Destroy; override;

Private
    ArtEcart : string;
{$IFDEF BTP}
    // Modif BTP
    DateFacture : TDateTime;
    // Modif FV 05/06/2007
    TobOuvrages : Tob;
    TOBpieces : TOB;
    // --
{$ENDIF}
    GenFac : R_GENFAC;
    nb_fac,nb_aff,fac_deb,fac_fin : Integer;
    avectrans : boolean;
    RevLogFile : String;
    datech,AffRef :  string;
    TobRegAff,TobRegFAff,TOBOuvragesP : TOB ;
    TobEchFact,TobEchFFact,TobEchBis ,TobEchFBis,TobAct,TobActLig : TOB;
    TobVarQte_ACT,TobVarQte_ACTbis : TOB;
    Rech :R_ECH;
    TobPiecePrepa :Tob;
		TobActivitePrepa : TOB;
    TobProfil : Tob;
    TobFormuleVar : Tob;
    TobArticles : Tob;
    TOBVTECOLL : TOB;
    titre : string;
    bSansMajBase :boolean; //AB-200610- Préparation sans maj en base
    // zone utilisées dans toutes les fct liées à la transformation
    Caff ,cemp,ddeb ,dfin,ctopart: string;
    // zone utilisées dans tGetData
    xemp,xdeb ,xfin,xlib : string;
    xclilib,xclicode,xcliville : string;
		//
    xDebFac,xFinFac : TdateTime;
    xprorata : Double;
    zsouche : string;
    Znature : string;
    ZnewNum : Integer;
    RegSurCaf : boolean;
    iCpt : Integer;   // compteur
    IndiceNomen : integer;

    Procedure AFTransaction;
    Procedure AFTrtGene;
    Function  AFTrtEntete : boolean ;
   	Procedure TrtAffaireReg ;
    Procedure MajAffaire(numdoc : String);
    Procedure MajEcheBis(numdoc : String);
    Procedure MajActivite(numdoc : String);
    Procedure MajActiviteApprec(numdoc : String);
    Procedure MajAcompteApprec;
    Procedure AFTrtLignes ;
    procedure ToutAllouer ;
    procedure ToutLiberer ;
    procedure LoadLesTOB ;
    procedure LoadTobLignes(TypGen,ztyp : String ;var nblig : integer);
    procedure TrtLignesAffaire(Typgen,ztyp : String;TobPiece_AFF,TobVarQte_AFF : TOB);
    procedure TrtActiviteSurLigneAffaire(TobDetLig,TobPiece_AFF,TobVarQte_AFF : TOB);
    procedure TrtLignesAffairesLiees(TobAct,TobDetLig,TobPiece_aff,TobVarQte_AFF: TOB);
    function  TrtLigneLiee(TobLig,TobAct,TobVarQte_AFF : TOB): boolean;
    procedure LoadTobLignesCom ;
    procedure Otesleslignes(CodeAff : String);
    procedure Transformation;
    procedure TriTobLigne(var critere : string);
    procedure ApplicProfil(deremp , RuptCli : boolean);
    function  TrtComment(Toblig : tob;pelt : string) : boolean;
    Function  RechNombreAff (var listeaff,listecli: array of string) : Integer;
    Function  RechNombreEmp (caff : string;var listeemp,listeDdeb,listeDFin,listeArt : array of string) : Integer;
    Procedure RechNombreArt (caff : string;var listeArt : array of string);
    procedure CreerComEmp (Toblig:tob;Comment : string);
    procedure TrtLibPres(Tobdet : tob;var zlib : string);

    procedure BoucleLignes(pelt : string ; deremp : boolean);
    procedure TrtCommentDebfac;
    function  RecupLig(Toblig : tob) :boolean;
    function  CumLigForfait(Toblig : tob) : boolean;
    function  RecupLigFinOuPou(Toblig : tob;pelt : string) : boolean;
    function  DetailLig(pelt1 : string;Toblig : tob) : boolean;
    function  CumLig(pelt : string;Toblig : tob) : boolean;
    procedure CumLigCode(pelt,nomcrit,valcrit : string;Toblig : tob);
    function  CumLigFamType(nomcrit,valcrit : string;Toblig : tob): boolean;
    procedure TrtTobCumul(pelt,pdetail: string ; RuptCli : boolean) ;
    procedure TrtTobCumul_Total(pelt,sv_crit : string;zmnt,zqte:double;TopTotImp : boolean;
    Tobsauv : TOB );
    procedure GereCumul(TobDet: TOB;TopDetImp :boolean;var zmnt,zqte : double;pelt : string);
    procedure AlimLibSousTot(pelt,pelt1,pelt2,zfam : string; var zlib : string;TobSauv : TOB);
    procedure InitTobLigneApresDupplic(TobLig: TOB);

    procedure MajQte(TobPiece: TOB);
    procedure RenumNumligne(TobPiece_OK : TOB);
    procedure MajLignesRess(TobP : TOB;MajRes:boolean);
    procedure ApplicTarif(TobP : TOB);


    procedure CreationLignePiece(Toblignes : TOB;zqte,zpuv : double;zlib,zori : string);
    procedure MajTobLignes;
    procedure LigneAcompte;
   // pas utlisée procedure LigneComFact;
    procedure LigneComAff(typcom : string;var comment : Hstring);
    function  GetData(s:Hstring):variant;
    procedure ChargeActivite(CAct,CodeAff : string ;TobActWork : Tob);
    Procedure DechargeActivite;
    procedure CreationLigDepuisActivite(TobAct,TobDetLig: TOB);
    procedure CreationLigDepuisActiviteLiee(TobAct,TobLigLiee,TobVarQte_AFF : TOB);
    procedure LigneApprec;


    procedure InitPieceCreation ;

    procedure ChargeTiers ;
    Procedure IncidenceTiers ;
    Procedure IncidenceAffairePiece ;
    Function  RemplirTOBAffairePrepFact ( CodeAffaire : String ) : boolean ;
    procedure GP_DEVISEChange;
    procedure ValideLaPiece ;


    function  TestSiCumulArticle(Toblig : TOB):boolean;
    function  GestionRevision(pStRevLogFile : String) : boolean;
    Procedure MajRevision;




    Function  RecupProfil(prof : string):boolean ;
    Function  RechProfilPrest(typ : string; tobprofil : TOB;var zlib : string;var zmini,zmaxi : Double) : Integer;
//    procedure MajBasePorc;
    Function  RemplirTOBRessource ( Cemp : String ) : Tob ;

		procedure RecupProrataDateFac;


//    Function  TestSetNumberAttribution ( Souche : string ) : integer ;
    procedure RechNumeroPiece;
    procedure TraitePrixLigneCourante(TOBL: TOB; Pu: double);
    procedure PositionneInfosLigneSurOuv(TOBL: TOB);
    procedure AjouteLigneEcart(TOBPiece : TOB;Ecart : double);
    procedure ReinitAffairePieceReg;
    procedure ChargelesTextes (TOBPiece,TOBOLES : TOB);
    procedure GetAdresses (TOBPiece,TOBSADRESSES,TOBAdresses : TOB);
    procedure GetAdressesContrat(TOBPiece, TOBSADRESSES, TOBAdresses: TOB);
    procedure GetMEMOS (TOBPiece,TOBSOLES,TOBLienOle : TOB);
    function RecupTiersAuxi(TypeCpt, CptTiers: String): string;

//    Procedure RechAdresseFact ( TOBTiers,TOBAdresses,TOBPiece : TOB ) ;

End;

var
    // Tob en Mémoire sur la pièce
    TOBLienOle,TOBSADRESSES,TOBSOLES : TOB;
    TOBPiece,TOBEches,TOBBases,TOBBasesL,TOBTiers,TOBTarif,TOBAdresses,TOBAffaire,TOBConds,TOBNomenclature : TOB ;
    TOBPieceRG, TOBPorcs ,TOBAnaP, TOBAnaS, TOBCpta : TOB;
    TOBPiece_OK,TOBPiece_CUM: TOB;
    TOBPiece_O,TOBBases_O : TOB ;
    CleDoc,CleDocOrigine,CleDocTemp : R_CleDoc ;
    DEV       : RDEVISE ;
    TOBRessource : TOB;
    TobVarQte_LIG : TOB;
    fTobRevision : TOB;

const
	// libellés des messages de la TOF  AFPREPFACT
	TexteMsgAffaire: array[1..6] of string 	= (
          {1}        'Paramètrage du profil facture incorrect'
          {2}        ,''
          {3}        ,''
          {4}        ,''
          {5}        ,''
          {6}        ,''
                     );

procedure AFPrepFact_SansMajBase(ZTobRegAff, ZTobArticles,ZTobPiece,ZtobActivite: TOB;Zdatedeb,ZDateFin :string);

implementation

uses FactureBTP,FactOuvrage,CalcOlegenericBtp,FactTimbres,wCommuns, BTPUtil, UCumulCollectifs;
//

Procedure  AFPrepFact_init(ZTobRegAff,ZtobProfil,ZtobFormuleVar,ztobarticles : TOB ;  GenFac : R_GENFAC;znb_aff : Integer;
                           zdatech  : string; ZAvecTrans : boolean;
                           var znb_fac,zfac_deb,zfac_fin : Integer;
                           zRevLogFile : String;DateFacture : TDateTime);
Var
  AFPrepFact : TAFPrepFact;
Begin

     AfPrepFact := TAFPrepFact.create(ZTobRegAff,ZTobProfil,ZtobFormuleVar,ZtobArticles,GenFac,znb_aff,zdatech,zavectrans,znb_fac,zfac_deb,zfac_fin, zRevLogFile,'',0);

     if (DateFacture <> 0) and (DateFacture <> iDate1900) then
    AfPrepFact.DateFacture := DateFacture
  else
				AfPrepFact.DateFacture := StrToDate(zdatech);

     AfPrepFact.AFTransaction;

     znb_fac := AfPrepFact.nb_fac;
     zfac_deb := AfPrepFact.fac_deb;
     zfac_fin := AfPrepFact.fac_fin;

     AfPrepFact.Destroy;

End;


procedure AFPrepFact_SansMajBase(ZTobRegAff, ZTobArticles,ZTobPiece,ZtobActivite: TOB;Zdatedeb,ZDateFin :string);
var
  AFPrepFact: TAFPrepFact;
  TobProfil ,TobRegAff: Tob;
  GenFac: R_GENFAC;
  nb_aff,nb_fac,fac_deb, fac_fin :integer;
  i_ind, RevForcer: Integer;
  datech,RevLogFile,psouche:string;
begin
  TobProfil := Tob.Create('les profils', nil, -1);
  TobRegAff := Tob.Create('Affaire Select', nil, -1);
  ChargeTobProfil(TobProfil);
  Genfac.zgen := 'G';
  Genfac.ztypech := 'NOR';
  Genfac.znat := 'FPR';
  GenFac.zdatedeb := Zdatedeb;
  GenFac.zdatefin := ZDateFin;
  GenFac.zdateact_d := Zdatedeb;
  GenFac.zdateact_f := ZDateFin;
  RevForcer := 0 ;nb_aff := 0;
  psouche := '';
  for i_ind := ZTobRegAff.detail.count -1 downto 0 do
  begin
    ZTobRegAff.Detail[i_ind].ChangeParent(TobRegAff, -1);
    datech := TobRegAff.detail[0].GetString('AFA_DATEECHE');
    AfPrepFact := TAFPrepFact.create(TobRegAff, TobProfil, nil, ZTobArticles, GenFac, nb_aff, datech, true, nb_fac, fac_deb, fac_fin,
      RevLogFile,psouche, RevForcer);
    try
      AfPrepFact.bSansMajBase := true;
      AfPrepFact.TobPiecePrepa := ZTobPiece;
      AfPrepFact.TobActivitePrepa := ZtobActivite;
      AfPrepFact.AFTransaction;
    finally
      AfPrepFact.Destroy;
      TobRegAff.cleardetail;
    end;
  end;
  TobProfil.free;
  TobRegAff.free;
end;



Constructor TAFPrepFact.create(ZTobRegAff ,ZTobProfil ,ZtobFormuleVar,ZtobArticles: TOB ;  ZGenFac : R_GENFAC;znb_aff : Integer;
                               zdatech  : string; Zavectrans : boolean ;
                               var znb_fac,zfac_deb,zfac_fin : Integer;
                               zRevLogFile,zsouche: string; zRevForcer: Integer);
var QArt : TQuery;
BEGIN
//
ArtEcart := '';
ArtEcart := GetParamsoc('SO_BTECARTPMA');
Qart := opensql('SELECT GA_ARTICLE FROM ARTICLE WHERE GA_CODEARTICLE="' + ArtEcart + '"', true);
if not Qart.eof then ArtEcart := Qart.fields[0].AsString else ArtEcart := '';
ferme(Qart);
//
if ArtEcart = '' then PGIInfo ('ATTENTION : l''article d''écart n''est pas défini.');
//
TobRegAff := ZTobRegAff;
TobProfil := ZTobProfil;
TobFormuleVar := ZtobFormuleVar;
TobArticles := ZTobArticles;
GenFac := ZGenFac;

datech := zdatech;
nb_aff := znb_aff;
fac_deb := zfac_deb;
fac_fin := zfac_fin;
nb_fac := znb_fac;
AvecTrans := zAvecTrans;
RevLogFile := zRevLogFile;
END;

destructor TAFPrepFact.Destroy;
begin
inherited;
end;


{***********A.G.L.***********************************************
Auteur  ...... :

G. Merieux

Créé le ...... : 17/01/2000

Modifié le ... :   /  /

Description .. : Génération automatique facture

Mots clefs ... : GENERATION;FACTURE

*****************************************************************}

Procedure TAFPrepFact.AFTransaction;

Var      io   : TIoErr ;

				znat : string;

BEGIN

  io :=oeOk ;

  if (Genfac.zgen = 'G') then

    begin

    	if (Genfac.znat = 'APR') then

      	Znat :='APR'
      else
      	if GetParamSoc('SO_FACTPROV') = False then
	       Znat :='FAC'
        Else
	       Znat :='FPR';
    end
  else
   	Znat:='FSI'; // simulation  donc factures de simulation
  Znature := znat;
  Zsouche:=GetSoucheG(znat,'','');  // attention, pour gérer une souche par établiessement

  																	// (demande mcd), il faudrait accéder à l'affaire pour

                                    // connaitre l'etablisseùment et le domaine.

                                    // attention, si on veut générer un avoir, il faudrait le savoir

                                    // ici ????


  if (AvecTrans) then

  Begin

     io:=Transactions(RechNumeroPiece,0) ;

     if io=oeOk then io := transactions(AftrtGene,1);

  End

  else

  Begin

  	 RechNumeroPiece;

  	 AftrtGene;

  End;


   Case io of
//       oeOk : G_MajEvent ;
       oeUnknown : BEGIN
                   MessageAlerte('ATTENTION : La génération ne s''est pas complètement effectuée') ;
                   END ;
       End;


END;

// *****************************************************************

//  Paramètres :

//    TOBRegAff  contenant la ou les affaires constituant 1 facture

//    Genfac : paramètres

// *****************************************************************

Procedure TAFPrepFact.AFTrtGene;

Var numdoc : string;
    zz :integer;
    topart,bPieceOk : boolean;
    Ind : integer;
    TOBPP,TOBPPCOM,TOBL : TOB;
    Indice : integer;
    TheCodeAffaire : string;
Begin

  bPieceOK := true;
  iCpt := 0;
  titre := 'Génération des factures';

  // je vide la tobarticle si elle devient trop importante
	if (tobarticles.detail.count > 200) then TobArticles.cleardetail;

  toutallouer;
  IndiceNomen := 1;

  //phase 1  (specif Affaire)
  TrtAffaireReg();

  // phase 1,5 -- :) --- Regroupement des lignes de pièces dans une seule
  TOBpiece.ClearDetail;
  TOBBases.ClearDetail;
  TOBBasesL.ClearDetail;

  if TOBpieces.detail.count > 1 then
  begin
    for Indice := 0 to TOBpieces.detail.count -1 do
    begin
      if Indice = 0 then
      begin
      	TOBPiece.Dupliquer(TOBPieces.detail[0],false,true,true);
  			ReinitAffairePieceReg; // enleve le contrat de la pièce puisqu'il en regroupe +ieurs
      end;
      TOBPP := TOBPieces.detail[Indice]; // entete
      if (copy(TOBPP.GetValue('GP_AFFAIRE'),1,1)='I') then
      begin
        // Ajoute un entete pour l'affaire --
        TheCodeAffaire := BTPCodeAffaireAffiche (TOBPP.geTvalue('GP_AFFAIRE'),'');
        //
        TOBPPCOM := NewTOBLigne( TOBPiece,0);
        TOBPPCOM.PutValue('GL_NUMORDRE',0);
        TOBPPCOM.PutValue('GL_ARTICLE','');
        // pour identifier type de commentaire par rapport au profil
        TOBPPCOM.PutValue('GL_TYPELIGNE','DP1');
        TOBPPCOM.PutValue('GL_NIVEAUIMBRIC',1);
        // formattage du libellé paramètré
        TOBPPCOM.PutValue('GL_LIBELLE','CONTRAT : '+TheCodeAffaire);

        PieceVersLigne ( TOBPiece,TOBPPCOM ) ;

        TOBPPCOM.PutValue('GL_AFFAIRE',TOBPP.GetValue('GP_AFFAIRE'));
        TOBPPCOM.PutValue('GL_AFFAIRE1',TOBPP.GetValue('GP_AFFAIRE1'));
        TOBPPCOM.PutValue('GL_AFFAIRE2',TOBPP.GetValue('GP_AFFAIRE2'));
        TOBPPCOM.PutValue('GL_AFFAIRE3',TOBPP.GetValue('GP_AFFAIRE3'));
        TOBPPCOM.PutValue('GL_AVENANT',TOBPP.GetValue('GP_AVENANT'));
        TOBPPCOM.PutValue('GL_REPRESENTANT',TOBPP.GetValue('GP_REPRESENTANT'));
        GestionChampSuppLigne(TOBPPCOM,TOBPPCOM.getValue('GL_AFFAIRE'),DatetoStr(idate2099),Tobpiece.GetValue('GP_TIERS'),TOBPieces.detail[0].getValue('GL_AFFAIRE'));
      end;
      //
      repeat
      	TOBPP.detail[0].changeparent(TOBpiece,-1);
        TOBL := TOBPiece.detail[Tobpiece.detail.count-1];
        if copy(TOBL.getValue('GL_TYPELIGNE'),1,2)='DP' then
        begin
        	TOBL.putValue('GL_TYPELIGNE','DP'+IntToStr(StrToInt(copy(TOBL.getValue('GL_TYPELIGNE'),3,1))+1));
        end;
        if copy(TOBL.getValue('GL_TYPELIGNE'),1,2)='TP' then
        begin
        	TOBL.putValue('GL_TYPELIGNE','TP'+IntToStr(StrToInt(copy(TOBL.getValue('GL_TYPELIGNE'),3,1))+1));
        end;
        TOBL.putValue('GL_NIVEAUIMBRIC',TOBL.getValue('GL_NIVEAUIMBRIC')+1);
      until TOBPP.detail.count=0;
      //
      TOBL := NewTOBLigne( TOBPiece,0);
      TOBL.PutValue('GL_NUMORDRE',0);
      TOBL.PutValue('GL_ARTICLE','');
      // pour identifier type de commentaire par rapport au profil
      TOBL.PutValue('GL_TYPELIGNE','TP1');
      TOBL.PutValue('GL_NIVEAUIMBRIC',1);
      // formattage du libellé paramètré
      TOBL.PutValue('GL_LIBELLE','TOTAL CONTRAT : '+TheCodeAffaire);

      PieceVersLigne ( TOBPiece,TOBL ) ;

      TOBL.PutValue('GL_AFFAIRE',TOBPP.GetValue('GP_AFFAIRE'));
      TOBL.PutValue('GL_AFFAIRE1',TOBPP.GetValue('GP_AFFAIRE1'));
      TOBL.PutValue('GL_AFFAIRE2',TOBPP.GetValue('GP_AFFAIRE2'));
      TOBL.PutValue('GL_AFFAIRE3',TOBPP.GetValue('GP_AFFAIRE3'));
      TOBL.PutValue('GL_AVENANT',TOBPP.GetValue('GP_AVENANT'));
      TOBL.PutValue('GL_REPRESENTANT',TOBPP.GetValue('GP_REPRESENTANT'));

      GestionChampSuppLigne(TOBL,TOBL.getValue('GL_AFFAIRE'),DatetoStr(idate2099),Tobpiece.GetValue('GP_TIERS'),TOBPieces.detail[0].getValue('GL_AFFAIRE'));

      TOBL := NewTOBLigne( TOBPiece,0);
      TOBL.PutValue('GL_NUMORDRE',0);
      TOBL.PutValue('GL_ARTICLE','');

      // pour identifier type de commentaire par rapport au profil
      TOBL.PutValue('GL_TYPELIGNE','COM');
      TOBL.PutValue('GL_NIVEAUIMBRIC',0);

      // formattage du libellé paramètré
      TOBL.PutValue('GL_LIBELLE','');

      PieceVersLigne ( TOBPiece,TOBL ) ;

      TOBL.PutValue('GL_AFFAIRE',TOBPP.GetValue('GP_AFFAIRE'));
      TOBL.PutValue('GL_AFFAIRE1',TOBPP.GetValue('GP_AFFAIRE1'));
      TOBL.PutValue('GL_AFFAIRE2',TOBPP.GetValue('GP_AFFAIRE2'));
      TOBL.PutValue('GL_AFFAIRE3',TOBPP.GetValue('GP_AFFAIRE3'));
      TOBL.PutValue('GL_AVENANT',TOBPP.GetValue('GP_AVENANT'));
      TOBL.PutValue('GL_REPRESENTANT',TOBPP.GetValue('GP_REPRESENTANT'));

      GestionChampSuppLigne(TOBL,TOBL.getValue('GL_AFFAIRE'),DatetoStr(idate2099),Tobpiece.GetValue('GP_TIERS'),TOBPieces.detail[0].getValue('GL_AFFAIRE'));

    end;
  end else
  begin
  	TOBPiece.Dupliquer(TOBPieces.detail[0],true,true,true);
    //FV1 : 12/12/2014 - FS#1350 - ACTUACOM - Non reprise de l'adresse de facturation en génération d'échéance
    //on Controle si on est sur un contrat ou non (sic)
    if copy(TOBPiece.GetString('GP_AFFAIRE'),1,1)='I' then
      GetAdressesContrat (TOBPiece,TOBSADRESSES,TOBAdresses)
    Else
    GetAdresses (TOBPiece,TOBSADRESSES,TOBAdresses);
    GetMEMOS (TOBPiece,TOBSOLES,TOBLienOle);
  end;

// phase 2 : transformation , présentation facture
  if (TobPiece.detail.count <> 0) then
    Transformation; // Renumerotation + maj clé

// phase 1 bis  : Prise en compte des révisions et des varaibles
// attention appelé cette fct avant renumleslignes car on a besoin du gl_numordre de l'affaire
  if (GetParamSoc('SO_AFREVISIONPRIX')) and (TobPiece_OK.detail.count <> 0) then
      bPieceOK := GestionRevision(RevLogFile);

  // renum des lignes de Piéces
  if bPieceOK then RenumLesLignes(TobPiece_OK);

 // phase 3 ; validation fature avec focntion JLD
  if (bPieceOK) and (TobPiece_OK.detail.count <> 0) then
  Begin
    ValideLaPeriode ( TOBPiece_OK) ;
//    MajBasePorc;
//    MajBasePort(TOBPiece_Ok, TobPorcs);
//modif BTP
    ZeroFacture (TOBPiece_Ok);
    for Ind := 0 to TOBPiece_Ok.detail.count -1 do ZeroLigneMontant (TOBPiece_Ok.detail[Ind]);
    TOBBases.ClearDetail;
    TOBBasesL.ClearDetail;
    ZeroMontantPorts (TOBPorcs);
    PutValueDetail(TOBPiece_Ok,'GP_RECALCULER','X') ;
    TOBVTECOLL.ClearDetail;
//---
    CalculFacture(TOBAffaire,TOBPiece_OK,nil,nil,TobOuvrages ,TOBOuvragesP,TOBBases,TOBBasesL,TOBTiers,TOBArticles,TobPorcs,Nil,Nil,TOBVTECOLL,DEV);
//    RecalculeSousTotaux;
    RecalculeSousTotaux(TOBPiece_OK);
  End;


// Controle de l'existence de lignes article pour ne pas générer de factures
//  vide
  if (bPieceOk) then
  begin
    topart := false;
    for zz:=0  to Tobpiece_ok.Detail.count-1 do
      Begin
      if (Tobpiece_ok.Detail[zz].getvalue('GL_TYPELIGNE') = 'ART') then
         Begin
           topart := true; break;
         End
      end;
    if not(topart) then bPieceOk := false;
  end;

  if (bPieceOk)  and (TobPiece_OK.detail.count <> 0) then    // 14/12/2000 modif en 520, avant on testait  ('GP_TOTALHTDEV') > 0)
  Begin
    ValideLaPiece;
    if (nb_fac = 0) then
    Begin
      fac_deb :=  TOBPiece_OK.getValue('GP_NUMERO');
      fac_fin :=  TOBPiece_OK.getValue('GP_NUMERO');
    End
    else
      fac_fin :=  TOBPiece_OK.getValue('GP_NUMERO');
    nb_fac := nb_fac + 1;

    Numdoc := EncodeRefPiece(TOBPIECE_OK);
  //MAJ//
    if (Genfac.zgen = 'G') then
    Begin
      UG_MajAnalPiece ( TobAnaP,TOBAnaS ,Numdoc ) ;
      if GetParamSoc('SO_AFREVISIONPRIX') then  MajRevision;
      if GetParamSoc('SO_AFVARIABLES') then MajVariable(Tobpiece_ok, TOBVarQte_LIG);

      TobEchFFact.PutValue('AFA_ECHEFACT','X');   // MAJ top echeance aff.principale
      TobEchFFact.PutValue('AFA_NUMPIECE',numdoc);
      TOBEchFact.UpdateDB(False);

      MajEcheBis(numdoc);

      MajActivite(numdoc);
      MajAffaire(Numdoc); // mcd 06/08/02 déplacer pour être fait apres MajActivite (avant était sous le begin)

    End
  End; // pieceok


  ToutLiberer;
  TobRegFAff.DelChampSup('ZMAJ',true);   // pour topage des affaires à mettre à jour

End;

// ****************************************************************************
//  Traitement de ou des affaires  regroupées.

//   Pour chaque affaire liée , on regarde si elle a une date d'échéance égale à la

//   date d'échéance de l'affaire principale

// *****************************************************************************

Procedure TAFPrepFact.TrtAffaireReg;

Var wi,Nbpiece,rang : integer;
    CodeAffaire,zprof : string;
    ret,top_ech : boolean;
    sttiers,StRep : String;
    codegenerauto : string;   // gm 08/07/02
    TOBPieceB : TOB;
BEGIN
rang := 0;
ret := true;
// boucle sur affaires regroupées ou non donnant lieu à 1 facture
for wi:=0  to TobRegAff.Detail.count-1 do
Begin

  TobRegFAff := TobRegAff.Detail[wi];
  CodeAffaire := TobRegFAff.getValue('AFF_AFFAIRE');
  CodeGenerauto := TobRegFaff.getValue('AFA_GENERAUTO');  // gm 08/07/02

  top_ech := true;
  if (wi = 0) then   // Affaire principale
  begin
    TobRegFAff.AddChampSup('ZMAJ',true);   // pour topage des affaires à mettre à jour
    AffRef := CodeAffaire;
    // init avec les données propres à l'échéance
    ChargeEcheances(CodeAffaire,CodeGenerauto,Genfac.ztypech,StrToDate(datech),StrToDate(datech),TobEchFact);
    if (TobEchFact.Detail.count <> 0) then
    Begin
      TobEchFFact  := TobEchFact.Detail[0];
      Rech.DateEch := TobEchFFact.getValue('AFA_DATEECHE');
      Rech.MontEch := TobEchFFact.getValue('AFA_MONTANTECHEDEV');
      Rech.PourEch := TobEchFFact.getValue('AFA_POURCENTAGE');
      Rech.NumEch  := TobEchFFact.getValue('AFA_NUMECHE');
      Rech.NumEchBis  := TobEchFFact.getValue('AFA_NUMECHEBIS');
      Rech.RepriseActiv  := TobEchFFact.getValue('AFA_REPRISEACTIV');
      Rech.Libelleeche  := TobEchFFact.getValue('AFA_LIBELLEECHE');
      Rech.profil  := TobEchFFact.getValue('AFA_PROFILGENER');
      CodeGenerauto := TobRegFaff.getValue('AFA_GENERAUTO');  // gm 08/07/02
    End
    else top_ech :=false;
  End;

  if (wi <> 0) and (top_ech) then   // Recherche date echeance affaires liées
  begin
    Rang := TobEchBis.detail.count;
    ChargeEcheances(CodeAffaire,CodeGenerauto,Genfac.ztypech,StrToDate(Rech.DateEch),StrToDate(Rech.DateEch),TobEchBis);
    if (TobEchBis.Detail.count <> 0) then
    Begin
      top_ech := true;
      TobEchFBis := TobEchBis.detail[rang];
      Rech.MontEch := TobEchFBis.getValue('AFA_MONTANTECHEDEV');
      Rech.PourEch := TobEchFBis.getValue('AFA_POURCENTAGE');
      Rech.NumEch  := TobEchFBis.getValue('AFA_NUMECHE');
      Rech.NumEchBis  := TobEchFBis.getValue('AFA_NUMECHEBIS');
      Rech.RepriseActiv  := TobEchFBis.getValue('AFA_REPRISEACTIV');
       //mcd 05/03/02 Rech.Libelleeche  := TobEchFFact.getValue('AFA_LIBELLEECHE');
      Rech.Libelleeche  := TobEchFBis.getValue('AFA_LIBELLEECHE');
    End
    else top_ech :=false;
  end;

  if (top_ech) then  // Verif de la commende associée à l'affaire
  Begin
{$IFDEF BTP}
    NbPiece := SelectPieceAffaireBis(CodeAffaire, 'INT', CleDocTemp,sttiers,strep);
{$ELSE}
    NbPiece := SelectPieceAffaireBis(CodeAffaire, 'AFF', CleDocTemp,sttiers,strep);
{$ENDIF}
    CleDocOrigine := CleDocTemp;
    if (wi = 0) then Cledoc := CledocTemp;
    CleDoc.DatePiece:= StrToDate(Rech.DateEch);
    if (NbPiece <> 1) then top_ech := false;  // Pas de Cde associée 12/11/01  tester <>1
  End;

  if (top_ech) then  // On traite l'affaire
  Begin
    RemplirTOBAffairePrepFact(TobRegFaff.getValue('AFF_AFFAIRE'));
    TobAffaire.AddChampSup('LEREP',false);   // pour avoir le rep
    TobAffaire.PutValue('LEREP',strep);
    if (wi = 0) then   // Affaire Principale , gestion entête
    begin
{$IFDEF BTP}
      zprof := TobRegFaff.getValue('AFF_PROFILGENER');
{$ELSE}
      if (rech.profil <> '') then
        zprof := rech.profil
      else
        zprof := TobRegFaff.getValue('AFF_PROFILGENER');
{$ENDIF}
      ret := RecupProfil(zprof);
      if not(ret) then
      Begin
       PGIInfoAf(TexteMsgAffaire[1],titre);
       exit;
       End;

      ret := AFTrtEntete;

			// Gestion des commentaire début facture (issu du mul de sélection)
      if (GenFac.zcom1 <> '') then     //Commentaire Affaire
      	LigneComAff('COM',Genfac.zcom1);

      if (GenFac.zcom2 <> '') then     //Commentaire Affaire
        LigneComAff('COM',Genfac.zcom2);

    // END Affaire Principale , gestion entête   (wi=0)
    end else
    begin
      ret := AFTrtEntete;
    end;

    // recup pour l'affaire traitée du montant de prorata et de la période REELLE de facturation
//    if (TobRegFaff.getValue('AFA_GENERAUTO') = 'CON') then  RecupProrataDateFac;
    RecupProrataDateFac;
    if VH_GC.GCIfDefCEGID then
    begin
      TobPiece.PutValue('GP_DATELIBREPIECE1',xdebfac);
      TobPiece.PutValue('GP_DATELIBREPIECE2',xfinfac);
    end;
    // traitement des lignes
    if (ret) Then AFTrtLignes;


    // Topage affaire pour maj entete Affaire
    TobRegFaff.PutValue('ZMAJ','X');
    if (ret) then
    begin
    	TOBpieceB := TOB.create ('PIECE',TOBpieces,-1);
      AddLesSupEntete (TOBPieceB);
    	TOBpieceB.Dupliquer(TOBpiece,true,true,true);
      TOBpiece.ClearDetail;
      TOBpiece.InitValeurs(false);
			TOBBases.ClearDetail;
			TOBBasesL.ClearDetail;
			TOBPorcs.ClearDetail;
    end;

  End;  // top_ech

  if (wi <> 0) then // maj échéance facturée de l'affaire secondaire
  begin
    if (top_ech) then
    begin
      TobEchFBis  := TobEchBis.Detail[rang];
      TobEchFBis.PutValue('AFA_ECHEFACT','X');   // Echeance facturée
    end;
  end;


End; // for wi:=0  to TobRegAff.Detail.count-1 do

  DechargeActivite();
END;

// ************************************************************
//  AFTrtEntete

//    Traitement des infos relatives à l'entête de la piece :

//               * Tiers

//               * Affaire

// ************************************************************
Function TAFPrepFact.AFTrtEntete : boolean;
var st : string;
		nech : integer;
BEGIN
	result := true;
	LoadLesTOB;   // TOb relatives à l'entête


	InitPieceCreation;  // Nature Piéce , Souche ...

  RegSurCaf := false;
  if (TobRegFaff.GetValue('AFF_REGSURCAF') = 'X')   then
   		RegSurCaf := true;

  if  RegSurCaf  then
    st := TobRegFaff.getValue('LECAF')
  else
    st := TobRegFaff.getValue('T_TIERS');



    TobPiece.PutValue('GP_TIERS',st);
    TobPiece.PutValue('GP_TIERSLIVRE',st);

    // Verif du payeur, s'il est différent on le garde
   	if  RegSurCaf  and (TobPiece.Getvalue('GP_TIERS')=TobPiece.Getvalue('GP_TIERSPAYEUR')) then
      TobPiece.PutValue('GP_TIERSPAYEUR',st);

     TobPiece.PutValue('GP_AFFAIRE',TobRegFaff.getValue('AFF_AFFAIRE'));
     if  (TobEchFFact.getValue('AFA_GENERAUTO') <> '') then
     	TobPiece.PutValue('GP_GENERAUTO',TobEchFFact.getValue('AFA_GENERAUTO'))
     else
     	TobPiece.PutValue('GP_GENERAUTO',TobRegFaff.getValue('AFA_GENERAUTO'));

    nech := TobEchFFact.getValue('AFA_NUMECHE');
    if TobEchFFact.getValue('AFA_GENERAUTO') = 'FOR' then
    Begin
      if (TobEchFFact.getvalue('AFA_TYPECHE') = 'APP') then
        begin
        	if (TobEchFFact.getvalue('AFA_APPRECIEE') <> 'X') then  nech := nech *(-1)
          // si on est en facture FOR en appreciation SIMPLE, on prend nech
          // sinon on prend (- nech)
        end
      else
      	nech := 1;

      TobPiece.PutValue('GP_FACREPRISE',nech);
    End;


     ChargeTiers;

     IncidenceAffairePiece;

End;

procedure TAFPrepFact.ToutAllouer ;
BEGIN
  // Pièce
  TOBVTECOLL := TOB.Create('LES VTS COLL',nil,-1);
  TOBpieces := TOB.Create ('LES PIECES REG',nil,-1);
  TOBPiece:=TOB.Create('PIECE',Nil,-1) ; TOBPiece_O:=TOB.Create('',Nil,-1) ;
  AddLesSupEntete (TOBPiece);AddLesSupEntete (TOBPiece_O); //mcd 26/02/02 suite modif BTP
  TOBBases:=TOB.Create('Les Bases',Nil,-1) ; TOBBases_O:=TOB.Create('',Nil,-1) ;
  TOBBasesL:=TOB.Create('Les Bases LIGNE',Nil,-1) ;
  TOBEches:=TOB.Create('Les Echeances',Nil,-1) ;
  TOBPorcs:=TOB.Create('Les Porcs',Nil,-1);
  TOBPieceRG:=TOB.Create('Les RG',Nil,-1);
  // Fiches
  TOBTiers:=TOB.Create('TIERS',Nil,-1) ; TOBTiers.AddChampSup('RIB',False) ;
  // TOBArticles:=TOB.Create('ARTICLES',Nil,-1) ;  déplacé dans prog appelant
  TOBTarif:=TOB.Create('TARIF',Nil,-1) ;
  TOBAffaire:=TOB.Create('AFFAIRE',Nil,-1) ;
  TOBConds:=TOB.Create('Les Conds',Nil,-1) ;
  // TobComm
  // Adresses
  TOBAdresses:=TOB.Create('LesADRESSES',Nil,-1) ;
  TOBSAdresses:=TOB.Create('LesADRESSES',Nil,-1) ;
  if GetParamSoc('SO_GCPIECEADRESSE') then
  Begin
    TOB.Create('PIECEADRESSE',TOBAdresses,-1) ; {Livraison}
    TOB.Create('PIECEADRESSE',TOBAdresses,-1) ; {Facturation}
  End
  else
  begin
    TOB.Create('ADRESSES',TOBAdresses,-1) ; {Livraison}
    TOB.Create('ADRESSES',TOBAdresses,-1) ; {Facturation}
  end;

  TOBNomenclature:=TOB.Create('Les NOMENCLATURES',Nil,-1) ;
  TobOuvrages := TOB.Create ('LES DETAIL OUVRAGES',nil,-1);
  TOBOuvragesP := TOB.create ('OUVRAGES A PLAT',nil,-1);
  //compta
  TOBCPTA := TOB.Create('', nil, -1);
  TOBANAP := TOB.Create('', nil, -1);
  TOBANAS := TOB.Create('', nil, -1);
  // Tob ddéfinitive pour MAJ
  TOBPiece_OK := TOB.Create('PIECE',Nil,-1) ;
  AddLesSupEntete (TOBPiece_OK); //gm 11/03/02 suite modif BTP

  TobAct := Tob.Create('Lignes Activite',nil,-1);       // activité liée à l'affaire
  TobActLig := Tob.Create('Lignes Activite pour lig affaires',nil,-1);  //activité liée à la ligne d'affaire


  TobEchFact :=TOB.Create('Echeances affaire principale',Nil,-1) ;
  TobEchbis := TOB.Create('Echeances affaires liées',Nil,-1) ;
  TOBRessource:=TOB.Create('Les RESSOURCE',Nil,-1) ;

  fTobRevision:= TOB.Create('les revisions',Nil,-1) ;
  if GetParamSoc('SO_AFVARIABLES') then
  begin
    TobVarQte_ACT  := Tob.Create('Les qte variables de activite',nil,-1);
    TobVarQte_ACTbis := Tob.Create('Les qte variables de activite bis',nil,-1);
    TOBVarQte_LIG:=TOB.Create('Les Qte variable de la piece finale',Nil,-1) ;
  end;

  TOBLienOle := TOB.Create('LES BLOCNOTE',nil,-1);
  TOBSOLES := TOB.Create('LES BLOCNOTE',nil,-1);
end;

procedure TAFPrepFact.ToutLiberer ;
BEGIN
	FreeAndNil(TOBpieces);
  TOBVTECOLL.Free;
  TOBPiece.Free ; TOBPiece:=Nil ; TOBPiece_O.Free ; TOBPiece_O:=Nil ;
  TOBBases.Free ; TOBBases:=Nil ; TOBBases_O.Free ; TOBBases_O:=Nil ;
  TOBBasesL.Free ; TOBBasesL:=Nil ;
  TOBEches.Free ; TOBEches:=Nil ;
  TOBPorcs.Free ; TOBPorcs:=Nil ;
  TOBPieceRG.Free ; TOBPieceRG:=Nil ;

  TOBTiers.Free ; TOBTiers:=Nil ;
  //  TOBArticles.Free ; TOBArticles:=Nil ;
  TOBTarif.Free ; TOBTarif:=Nil ;
  TOBCPTA.free; TOBAnaP.Free;  TOBAnaP := nil;
  TOBAnaS.Free;  TOBAnaS := nil;
  TOBConds.Free ; TOBConds:=Nil ;
  TOBNomenclature.Free ; TOBNomenclature:=Nil ;
  TobOuvrages.free; TobOuvrages := nil;
  TOBOuvragesP.free; TOBOuvragesP:= nil;
  TOBAdresses.Free ; TOBAdresses:=Nil ;
  TOBPiece_OK.Free ; TOBPiece_OK:=Nil ;
  TobAct.Free ; Tobact:=Nil;
  TobActLig.Free ; TobActLig:=Nil;
  TOBAffaire.Free ; TOBAffaire:=Nil ;
  TobEchFact.Free ; TobEchFact:=Nil;
  TobEchBis.Free ; TobEchBis:=Nil;
  TobRessource.Free ; TobRessource:=Nil;
  fTobRevision.Free ;  fTobRevision:=Nil;
  if GetParamSoc('SO_AFVARIABLES') then
  begin
    TOBVarQte_LIG.Free ; TOBVarQte_LIG:=Nil;
    TOBVarQte_ACT.Free ; TOBVarQte_ACT:=Nil;
    TOBVarQte_ACTbis.Free ; TOBVarQte_ACTbis:=Nil;
  end;
  TOBLienOle.free;
  TOBSOLES.free;
  TOBSADRESSES.Free;

end ;

procedure TAFPrepFact.LoadLesTOB ;
Var Q : TQuery ;
  req : string;
    TOBLAdresses  : TOB;
    TOBLOLE       : TOB;
    TOBAA         : TOB;
BEGIN

  TOBLAdresses := TOB.Create('LES ADRESSES',nil,-1);
  TOBLOLE := TOB.Create('LES LIENSOLES',nil,-1);

  // Lecture Piece  ***Création des pieces, il faut tout prendre ****
  req :=  'SELECT * FROM PIECE WHERE '+WherePiece(CleDocOrigine,ttdPiece,False);
  Q:=OpenSQL(req,True,-1,'',true) ;
  TOBPiece.SelectDB('',Q) ;
  tobpiece_o.dupliquer(TobPiece,False,True,True);
  Ferme(Q) ;

  // Lecture bases  ***Création des pieces, il faut tout prendre ****
  Q:=OpenSQL('SELECT * FROM PIEDBASE WHERE '+WherePiece(CleDocOrigine,ttdPiedBase,False),True,-1,'',true) ;
  TOBBases.LoadDetailDB('PIEDBASE','','',Q,False) ;
  Ferme(Q) ;

  // Lecture
  Q:=OpenSQL('SELECT * FROM LIGNEBASE WHERE '+WherePiece(CleDocOrigine,ttdLigneBase,False),True,-1,'',true) ;
  TOBBasesL.LoadDetailDB('LIGNEBASE','','',Q,False) ;
  Ferme(Q) ;

  // Lecture Ports   ***Création des pieces, il faut tout prendre ****
  Q:=OpenSQL('SELECT * FROM PIEDPORT WHERE '+WherePiece(CleDocOrigine,ttdPorc,False),True,-1,'',true) ;
  TOBPorcs.LoadDetailDB('PIEDPORT','','',Q,False) ;
  Ferme(Q) ;

  //LoadLesAdresses (TobPiece,TobAdresses);      non fct specifique

  //LoadLesNomen

  UG_LoadAnaPiece(TOBPiece,TOBAnaP,TOBAnaS) ;

  LoadLesAdresses(TOBPiece,TOBLAdresses);

  ChargelesTextes (TOBPiece,TOBLOLE);

  TOBAA := TOBSAdresses.FindFirst(['ORIGINE'],[CleDocOrigine.NumeroPiece],true);
  if TOBAA = nil then
  begin
    TOBAA := TOB.Create ('DES ADRESSES',TOBSADRESSES,-1);
    TOBAA.Dupliquer(TOBLAdresses,True,true);
    TOBAA.AddChampSupValeur ('ORIGINE',CleDocOrigine.NumeroPiece);
  end;
  TOBAA := TOBSOLES.FindFirst(['ORIGINE'],[CleDocOrigine.NumeroPiece],true);
  if TOBAA = nil then
  begin
    TOBAA := TOB.Create ('DES LIENSOLE',TOBSOLES,-1);
    TOBAA.Dupliquer(TOBLOLE,True,true);
    TOBAA.AddChampSupValeur ('ORIGINE',CleDocOrigine.NumeroPiece);
  end;

  TOBLAdresses.free;
  TOBLOLE.free;
END ;

procedure TAFPrepFact.InitPieceCreation ;
BEGIN

if (Genfac.zgen = 'G') then
  begin
    if (Genfac.znat = 'APR') then
      CleDoc.NaturePiece:='APR'
    else
      if GetParamSoc('SO_FACTPROV') = false then
      	CleDoc.NaturePiece := 'FAC'
      Else
	    CleDoc.NaturePiece:='FPR'
  end
else
  CleDoc.NaturePiece:='FSI'; // simulation  donc factures de simulation

//  CleDoc.Souche:=GetSoucheG(CleDoc.NaturePiece,'','');
CleDoc.Souche:= Zsouche;
Znature := CleDoc.NaturePiece;

TOBPiece.PutValue('GP_SOUCHE',CleDoc.Souche) ;
TOBPiece.PutValue('GP_NATUREPIECEG',CleDoc.NaturePiece);
TOBPiece.PutValue('GP_DATEPIECE',CleDoc.DatePiece) ;
// MODIF BTP
if DateFacture <> iDate1900 then TOBPiece.PutValue('GP_DATEPIECE',DateFacture) ;
// --
InitTOBPiece(TOBPiece) ;
TOBPiece.PutValue('GP_CREEPAR','GEN') ;
TOBPiece.PutValue('GP_VENTEACHAT',GetInfoParPiece(CleDoc.NaturePiece,'GPP_VENTEACHAT')) ;
TOBPiece.PutValue('GP_EDITEE','-') ;
if VH_GC.GCIfDefCEGID then
  TOBPiece.PutValue('GP_DOMAINE','001');
END ;


procedure TAFPrepFact.ChargeTiers ;
var ztiers  : String;
BEGIN
  if  RegSurCaf  then
    ztiers :=TobPiece.GetValue('GP_TIERSFACTURE')
  else
    ztiers :=TobPiece.GetValue('GP_TIERS');
  RemplirTOBTiers(TobTiers, zTiers, CleDoc.naturepiece,false);
  IncidenceTiers ;
END ;




Procedure TAFPrepFact.IncidenceTiers ;
Var CodeD,CodeTiers: String ;
    i  : integer ;
BEGIN

CodeTiers:=TOBTiers.GetValue('T_TIERS') ;

{
CodeLivr:=TOBPiece.GetValue('GP_TIERS');
AuxiLivr:=TOBTiers.GetValue('T_AUXILIAIRE') ;
AuxiFact:=TOBPiece.GetValue('GP_TIERSFACTURE');
AuxiPay:=TOBTiers.GetValue('GP_TIERSPAYEUR');
if ((AuxiFact='') or (AuxiFact=AuxiLivr)) then BEGIN AuxiFact:=AuxiLivr ; CodeFact:=CodeLivr ; END
                                          else BEGIN CodeFact:=TiersAuxiliaire(AuxiFact,True) ; END ;
if ((AuxiPay='') or (AuxiPay=AuxiFact))   then BEGIN AuxiPay:=AuxiFact ; CodePay:=CodeFact ; END
                                          else BEGIN CodePay:=TiersAuxiliaire(AuxiPay,True) ; END ;
TOBPiece.PutValue('GP_TIERSLIVRE',CodeLivr) ;
TOBPiece.PutValue('GP_TIERSFACTURE',CodeFact) ;
TOBPiece.PutValue('GP_TIERSPAYEUR',CodePay) ;
}
// TiersVersPiece(TOBTiers,TOBPiece)  Light , Debut;

//NaturePieceG:=TOBPiece.GetValue('GP_NATUREPIECEG') ;
//if GetInfoParPiece(NaturePieceG,'GPP_BLOBTIERS')='X' then TOBPiece.PutValue('GP_BLOCNOTE',TOBTiers.GetValue('T_BLOCNOTE')) ;
if TOBTiers.FieldExists('YTC_TABLELIBRETIERS1') then
   BEGIN
   for i:=1 to 9 do TOBPiece.PutValue('GP_LIBRETIERS'+IntToStr(i),TOBTiers.GetValue('YTC_TABLELIBRETIERS'+IntToStr(i))) ;
   TOBPiece.PutValue('GP_LIBRETIERSA',TOBTiers.GetValue('YTC_TABLELIBRETIERSA')) ;
   END ;

TOBPiece.PutValue('GP_TARIFTIERS',TOBTiers.GetValue('T_TARIFTIERS')) ;  //mcd 21/08/03 sinon pas OK si gestion tarif par catégorie
// TiersVersPiece(TOBTiers,TOBPiece) FIN ;

CodeD:=TOBTiers.GetValue('T_DEVISE') ;
if CodeD<>'' then TobPiece.PutValue('GP_DEVISE',CodeD);

END ;

{***********A.G.L.***********************************************
Auteur  ...... : G.Merieux
Créé le ...... : 31/08/2001
Modifié le ... :   /  /
Description .. : Preparation factures.
Suite ........ : Inidence de l'affaire et piece liée à l'affaire, dans la facture à
Suite ........ : emettre
Mots clefs ... : GIGA;FACTUREAUTO
*****************************************************************}
Procedure TAFPrepFact.IncidenceAffairePiece ;
Var CodeD, Fact : String ;

BEGIN
// Reprise des élément de l'affaire
CodeD:=TOBAffaire.GetValue('AFF_DEVISE') ;
if CodeD<>'' then TobPiece.PutValue('GP_DEVISE',CodeD);
TOBPiece.PutValue('GP_TIERSFACTURE', TiersAuxiliaire(TOBAffaire.GetValue('AFF_FACTURE'),true));

TOBPiece.PutValue('GP_AFFAIRE1',TOBAffaire.GetValue('AFF_AFFAIRE1'));
TOBPiece.PutValue('GP_AFFAIRE2',TOBAffaire.GetValue('AFF_AFFAIRE2'));
TOBPiece.PutValue('GP_AFFAIRE3',TOBAffaire.GetValue('AFF_AFFAIRE3'));
TOBPiece.PutValue('GP_AVENANT',TOBAffaire.GetValue('AFF_AVENANT'));

TOBPiece.PutValue('GP_FACTUREHT',TOBAffaire.GetValue('AFF_AFFAIREHT')) ;
TOBPiece.PutValue('GP_SAISIECONTRE',TOBAffaire.GetValue('AFF_SAISIECONTRE')) ;

TOBPiece.PutValue('GP_APPORTEUR', TOBAffaire.GetValue('AFF_APPORTEUR') );
TOBPiece.PutValue('GP_ETABLISSEMENT',TOBAffaire.GetValue('AFF_ETABLISSEMENT')) ;
TOBPiece.PutValue('GP_REFEXTERNE',TOBAffaire.GetValue('AFF_REFEXTERNE')) ;
   //mcd 24/10/02 ajout alimenation ref interne
xdeb := '01/01/1900'; xfin := '31/12/2099'; xemp := '';
if (Genfac.ZRefInt <> '') then  // gm 09/10/03 sinon on écrase la ref interne de l'affaire
  TobPiece.putvalue('GP_REFINTERNE',GFormule(Genfac.ZRefInt,GetData,nil,1));
   // mcd 29/01/02
TOBPiece.PutValue('GP_LIBREAFF1',TOBAffaire.GetValue('AFF_LIBREAFF1')) ;
TOBPiece.PutValue('GP_LIBREAFF2',TOBAffaire.GetValue('AFF_LIBREAFF2')) ;
TOBPiece.PutValue('GP_LIBREAFF3',TOBAffaire.GetValue('AFF_LIBREAFF3')) ;

// oter 18/07/02 car ensuite il y aun alignement GP_RESSOURCE
// sur GL_RESSOURCE et ca devient incompatible avec facture sur activité
//TOBPiece.PutValue('GP_RESSOURCE',TOBAffaire.GetValue('AFF_RESPONSABLE')) ;

// rech adresse du client livre et
// Recherche de l'adresse de facturation si GP_facture (de l'affaire) est renseigné , on le prends
// sinon on garde le t_facture du client
Fact :='';
if  TiersAuxiliaire(TOBPiece.GetValue('GP_TIERSFACTURE')) <> TobTiers.GetValue('T_FACTURE') then Fact := TOBPiece.GetValue('GP_TIERSFACTURE');
TiersVersAdresses(TobTiers,TobAdresses,Tobpiece,Fact);   //mcd suite modif pour clt fac # il faut ne mettre le champ clt fact que si # du tiers

//recup adr livraison dans l'affaire
AffaireVersAdresses(TobAffaire,TobAdresses,Tobpiece);

if (TOBPiece.GetValue('GP_ETABLISSEMENT') = '') then
  TOBPiece.PutValue('GP_ETABLISSEMENT',VH^.EtablisDefaut );
if (TOBPiece.GetValue('GP_DEPOT') = '') then
  TOBPiece.PutValue('GP_DEPOT',VH_GC.GCDepotDefaut) ;
if (TOBPiece.GetValue('GP_DEVISE') = '') then
  TOBPiece.PutValue('GP_DEVISE',V_PGI.DevisePivot);
if (TOBPiece.GetValue('GP_REGIMETAXE') = '') then
  TOBPiece.PutValue('GP_REGIMETAXE',VH^.RegimeDefaut);


  DEV.Code:=TOBPiece.GetValue('GP_DEVISE') ; GetInfosDevise(DEV) ;
  DEV.Taux:=GetTaux(DEV.Code,DEV.DateTaux,CleDoc.DatePiece) ;//mcd 11/04/2003 pb si tarification et forfait.. le taux est à 0 et calcul faux

// reprise des éléments issus de la piece rattachée à l'affaire
// recup des tiers personnalisés dans le contrat
if TOBPiece.GetValue('GP_TIERSFACTURE') = '' then
begin
	TOBPiece.PutValue('GP_TIERSFACTURE',TObPiece.GetValue('GP_TIERS'));
end;
if VH_GC.GCIfDefCEGID then
begin
  if TobPiece.GetValue('GP_LIBRETIERS1')='' then TobPiece.PutValue('GP_LIBRETIERS1','999');
  if TobPiece.GetValue('GP_LIBRETIERS2')='' then TobPiece.PutValue('GP_LIBRETIERS2','999');
  if TobPiece.GetValue('GP_LIBRETIERS3')='' then TobPiece.PutValue('GP_LIBRETIERS3','999');
end;

END ;



Function TAFPrepFact.RemplirTOBAffairePrepFact ( CodeAffaire : String ) : boolean ;
Var Q : TQuery ;
		zch : string;
BEGIN
Result:=True ;
zch := 'AFF_AFFAIRE,AFF_AFFAIRE1,AFF_AFFAIRE2,AFF_AFFAIRE3,AFF_AVENANT,AFF_DEVISE';
zch := zch + ',AFF_PROFILGENER,AFF_AFFAIREHT,AFF_SAISIECONTRE,AFF_APPORTEUR,AFF_ETABLISSEMENT';
zch := zch + ',AFF_DATEDEBGENER,AFF_DATEFIN,AFF_DATERESIL,AFF_METHECHEANCE,AFF_PERIODICITE';
zch := zch + ',AFF_INTERVALGENER,AFF_GENERAUTO,AFF_LIBELLE,AFF_COMPTAAFFAIRE,AFF_REFEXTERNE';
zch := zch + ',AFF_LIBREAFF1,AFF_LIBREAFF2,AFF_LIBREAFF3,AFF_RESPONSABLE,AFF_DATEFINGENER'; //mcd 29/01/02
zch := zch + ',AFF_REGSURCAF,AFF_TIERS,AFF_PRINCIPALE,AFF_FORCODE1,AFF_FORCODE2,AFF_FACTURE';

if CodeAffaire='' then  TOBAffaire.InitValeurs(False)
else if CodeAffaire<>TOBAffaire.GetValue('AFF_AFFAIRE') then
   BEGIN
   Q:=OpenSQL('SELECT '+zch+' FROM AFFAIRE WHERE AFF_AFFAIRE="'+CodeAffaire+'"',True,-1,'',true) ;
   If (Not Q.EOF) then
    Result:=TOBAffaire.SelectDB('',Q)
   Else
    Result:=False;
   Ferme(Q);
   END ;
END ;

Function TAFPrepFact.RemplirTOBRessource ( Cemp : String ) : Tob ;
Var Q : TQuery ;
TobRes : tob;
BEGIN
result := NIL;
if Cemp='' then  TOBRessource.InitValeurs(False)
else
	Begin
	 result  := TobRessource.FindFirst(['ARS_RESSOURCE'],[Cemp],true);
   if (result = NIL) then
   BEGIN
     Q:=OpenSQL('SELECT ARS_RESSOURCE,ARS_LIBELLE,ARS_LIBELLE2 FROM RESSOURCE WHERE ARS_RESSOURCE="'+Cemp+'"',True,-1,'',true) ;
     If (Not Q.EOF) then
     Begin
       TOBRes:=TOB.Create('RESSOURCE',TOBRessource,-1);
       TOBRes.SelectDB('',Q);
       result := Tobres;
     End
     Else Result:=NIL;
     Ferme(Q);
   END ;
  END;
END ;

procedure TAFPrepFact.GP_DEVISEChange;
begin

if (TOBPiece_OK.GetValue('GP_DEVISE')='') then BEGIN FillChar(DEV,Sizeof(DEV),#0) ; Exit ; END ;
DEV.Code:=TOBPiece_OK.GetValue('GP_DEVISE') ; GetInfosDevise(DEV) ;
DEV.Taux:=GetTaux(DEV.Code,DEV.DateTaux,CleDoc.DatePiece) ;

PutValueDetail(TOBPiece_OK,'GP_DEVISE',DEV.Code) ;
PutValueDetail(TOBPiece_OK,'GP_TAUXDEV',DEV.Taux) ;
PutValueDetail(TOBPiece_OK,'GP_DATETAUXDEV',DEV.DateTaux) ;

end;

procedure TAFPrepFact.ValideLaPiece ;
var OldEcr, OldStk: RMVT;
		NaturePieceG, St: String ;
 		i,zz: integer;
 		ztob,TOBB,TOBL,TOBA : TOB;
BEGIN
//  GestionDuVisa;
  GestionDuVisa(TobPiece_ok);
   TOBVTECOLL.SetAllModifie(true);
   TOBPiece.SetAllModifie(True) ;
   TOBEches.SetAllModifie(True) ;
   TOBBases.SetAllModifie(True) ;
   TOBBasesL.SetAllModifie(True) ;
   TOBAnaP.SetAllModifie(False);
   TOBAnaS.SetAllModifie(False);
   TOBLienOle.SetAllModifie(true);
   TOBAdresses.SetAllModifie(true);

   NaturePieceG:=TOBPiece_OK.GetValue('GP_NATUREPIECEG') ;
   if Not SetDefinitiveNumber(TOBPiece_OK,TOBBases,TOBBasesL,TOBEches,TOBNomenclature,Nil,Nil,Nil, nil, ZNewNum) then V_PGI.IoError:=oePointage ;

   if GetInfoParPiece(NaturePieceG,'GPP_ACTIONFINI')='ENR' then TOBPiece_OK.PutValue('GP_VIVANTE','-') ;
   CleDoc.NumeroPiece := TOBPiece_OK.GetValue('GP_NUMERO') ;
   for i:=0 to TOBBases.Detail.Count-1 do BEGIN TOBB:=TOBBases.Detail[i] ; MajFromCleDoc(TOBB,CleDoc) ; END ;
   for i:=0 to TOBEches.Detail.Count-1 do BEGIN TOBB:=TOBEches.Detail[i] ; MajFromCleDoc(TOBB,CleDoc) ; END ;
    // MODIF FV  05-06-07
   for i:=0 to TOBPiece_OK.Detail.Count-1 do BEGIN TOBL := TOBPiece_OK.detail[I]; PositionneInfosLigneSurOuv (TOBL); END;
    // -----
   // remettre le montant à) 1
   GereEcheancesGC(TOBPiece_OK,TOBTiers,TOBEches,NIL,Nil,nil,nil,taCreat,DEV,False);
// MAJ des prix valo et des stocks , trt nomenclature
// ValideLesLignes(TobPiece_OK, TOBArticles,TOBNomenclature,False) ;

  ValideAnalytiques(TOBPiece_OK, TOBAnaP, TOBAnaS);
  ValideLesAdresses(TOBPiece_OK, TOBPiece_OK, TOBAdresses) ;
  ValideLesLiensOle (TOBPiece_OK,TOBPiece_OK, TOBLienOle);
//gerelesreliquats
//ValideLesArticles   //maj TobArticle + Catalogue

// maj de creerpar en attendant la modif de JLD dans la fonction PieceVersLigne
for zz:=0  to Tobpiece_ok.Detail.count-1 do
    Begin
    ztob   := Tobpiece_ok.Detail[zz];
    ztob.putvalue('GL_CREERPAR','GEN');
    TOBA:=FindTOBArtRow(TOBPiece_OK,TOBArticles,zz+1) ;
    if TOBA <> Nil then ChargeAjouteCompta(TOBCpta,TOBPiece_OK,ztob,TOBA,TOBTiers,Nil,TOBaffaire,True) ;
    end;
   PassationComptable(TOBPiece_OK, TobOuvrages,TOBOuvragesP,TOBBases,TOBBasesL, TOBEches, nil,nil,TOBTiers, TOBArticles, TOBCpta, Nil, TOBPorcs, TOBPieceRG, Nil, nil,nil,nil,TOBVTECOLL,DEV, OldEcr, OldStk, True);
   LibereParamTimbres;
   TOBPiece_OK.InsertDBByNivel(False) ;  // gm 25/04/2002
   TOBBases.InsertDB(Nil) ;
   TOBEches.InsertDB(Nil) ;
   ValideLesPorcs(TOBPiece_OK,TobPorcs);
   TOBAnaP.InsertDB(nil);
   TOBAnaS.InsertDB(nil);
   TobOuvrages.setAllModifie (True);
   TobOuvrages.InsertDBByNivel (false);
   TOBOuvragesP.setAllModifie(true);
   TOBOuvragesP.InsertDBByNivel (false);
   if TOBVTECOLL.detail.count > 0 then
   begin
     PrepareInsertCollectif (TOBPiece,TOBVTECOLL);
     TOBVTECOLL.InsertDB(nil);
   end;
 // ValidelesNomen ;
   St := 'Pièce N° ' + IntToStr(TOBPiece_OK.GetValue('GP_NUMERO'))
      + ', Date ' + DateToStr(TOBPiece_OK.GetValue('GP_DATEPIECE'))
      + ', Tiers ' + TOBPiece_OK.GetValue('GP_TIERS')
      + ', Affaire ' + TOBPiece_OK.GetValue('GP_AFFAIRE')
      + ', Total HT de ' + StrfPoint(TOBPiece_OK.GetValue('GP_TOTALHTDEV')) + ' ' + RechDom('TTDEVISETOUTES', TOBPiece_OK.GetValue('GP_DEVISE'), False);
   MAJJnalEvent('GEN','OK', 'Génération de factures', St);
END ;


// ****************************************************************************
//

//  Mise à date Date et n° facture sur les affaires facturées

//

// *****************************************************************************

Procedure TAFPrepFact.MajAffaire(numdoc : string);

Var wi : integer;
    Req,CodeAffaire,stapprec : string;

BEGIN
// boucle sur affaires regroupées
for wi:=0  to TobRegAff.Detail.count-1 do
Begin
  TobRegFAff := TobRegAff.Detail[wi];
  if (TobRegFAff.getValue('ZMAJ') = 'X') Then
  Begin
    CodeAffaire := TobRegFaff.getValue('AFF_AFFAIRE');
    stapprec := '';
    if (TobEchFFact.GetValue('AFA_TYPECHE') = 'APP') then
    Begin
    // maj de la date d'appreciation seulement si facture de solde (cas Mazar avec lettrage)
    // sinon (sans lettrage) dans toutes les apprec
    	if (GetParamSoc('SO_AFAPPPOINT')=false)  or
      		((GetParamSoc('SO_AFAPPPOINT')=true)  and (TobEchFFact.GetValue('AFA_GENERAUTO')='SOL')) then
            stapprec := ',AFF_DATESITUATION ='+'"'+UsDateTime(TobEchFFact.GetValue('AFA_DATEECHE'))+'"'+
     			 ',AFF_NUMSITUATION ="'+inttostr(TobEchFFact.GetValue('AFA_NUMECHE'))+'"';
    End;
     Req := 'UPDATE AFFAIRE SET AFF_NUMDERGENER ='+'"'+Numdoc+'"'
 		 + stapprec
     +' WHERE AFF_AFFAIRE='+'"'+CodeAffaire+'"';
       ExecuteSQL(Req);
  End;
End; // for wi:=0  to TobRegAff.Detail.count-1 do

END;
// ****************************************************************************
//

//  Mise à date Date et n° facture sur les echeances des affaires regroupées

//

// *****************************************************************************

Procedure TAFPrepFact.MajEcheBis(numdoc : String);

var wi : Integer;

    TobDetEch : TOB;

BEGIN

for wi:=0  to TobEchBis.Detail.count-1 do
Begin
  TobDetEch := TobEchBis.Detail[wi];
  if (TobDetEch.getValue('AFA_ECHEFACT') = 'X') Then
  Begin
    TobDetEch.PutValue('AFA_NUMPIECE',numdoc);
  End;
End;

if (TobEchBis.detail.count<>0) then
Begin
  if (Genfac.zgen = 'G') then
    TobEchBis.UpdateDB(false);
  TobEchBis.clearDetail;
End;
END;
// ****************************************************************************
//

//  Mise à date Date et n° facture sur les lignes d'activité

// *****************************************************************************

Procedure TAFPrepFact.MajActivite(numdoc : String);

var wi : Integer;

    TobDet : TOB;

BEGIN

// cas de facturation hors apprec, avec reprise d'activite
for wi:=0  to TobAct.Detail.count-1 do
Begin
  TobDet := TobAct.Detail[wi];
  if (TobDet.getValue('ACT_ACTIVITEREPRIS') = 'FAC') Then
  Begin
    TobDet.PutValue('ACT_NUMPIECE',numdoc);
  End;
End;

if (TobAct.detail.count<>0) then
Begin
  TobAct.UpdateDB(false);
  TobAct.clearDetail;
End;

// si appreciation
  if (TobEchFFact.GetValue('AFA_TYPECHE') = 'APP') then
  Begin
   	MajActiviteApprec(numdoc);
    MajAcompteApprec;
  End;
END;

{***********A.G.L.***********************************************
Auteur  ...... : G.Merieux
Créé le ...... : 14/08/2001
Modifié le ... :   /  /
Description .. : Mise à jour de l'activte dans le cas de l'appreciation
Suite ........ :
Suite ........ : Génération des boni/mali
Mots clefs ... : ACTIVITE;APPRECIATION;GI
*****************************************************************}
Procedure TAFPrepFact.MajActiviteApprec(numdoc : String);
var req : string;
BEGIN
    // si on a lettré des lignes de temps, maj de celles-ci
	if (GetParamSoc('SO_AFAPPPOINT')=true) then
  begin
    req := 'UPDATE ACTIVITE  SET ACT_NUMPIECE ="'+numdoc+'"'
    +' WHERE ACT_AFFAIRE="'+TobRegFaff.getValue('AFF_AFFAIRE')+'"'
    +' AND ACT_NUMAPPREC="'+  inttostr(TobEchFFact.GetValue('AFA_NUMECHE')) +'"';
    ExecuteSQL(Req);
  End;

  // Gestion des boni/mali

   GenerationBoniMali ( TobEchFFact,Tobaffaire , numdoc);

END;
{***********A.G.L.***********************************************
Auteur  ...... : G.Merieux
Créé le ...... : 14/08/2001
Modifié le ... :   /  /
Description .. : Mise à jour des factures d'acompte reprise
Suite ........ :
Suite ........ : Génération des boni/mali
Mots clefs ... : ACTIVITE;APPRECIATION;GI
*****************************************************************}
Procedure TAFPrepFact.MajAcompteApprec;
var req : string;
BEGIN
// PL le 05/03/02 : INDEX 4
  req := 'UPDATE PIECE  SET gp_facreprise ='+IntToStr(TobEchFFact.GetValue('AFA_NUMECHE'))
  + ' WHERE gp_affaire="'+TobEchFFact.GetValue('AFA_AFFAIRE')+'" AND gp_facreprise='+IntToStr(50000+TobEchFFact.GetValue('AFA_NUMECHE'));

  ExecuteSQL(req);

END;
// ************************************************************
//  AFTrtLignes

//

// ************************************************************
Procedure TAFPrepFact.AFTrtLignes;
Var  CodeAff,CodeAct,TypGen,Princ : string;
			Nblig : Integer;
      Comment : Hstring;
      Ind,LigneQteUnitaire : integer;
      Ecart : double;

BEGIN

    CodeAff :=  TobRegFaff.getValue('AFF_AFFAIRE');
    TypGen  :=  TobRegFaff.getValue('AFA_GENERAUTO');
    Princ		:=  TobRegFaff.getValue('AFF_PRINCIPALE');
    CodeAct :=  rech.RepriseActiv;
    LigneQteUnitaire := -1;

    if Genfac.ztypech = 'APP' then TypGen := 'APP';


    // On charge dans la  tob TobACtLig, toute l'activité de l'affaire pour être sure de retrouver
    // les lignes  d'activité liées à la ligne d'affaire.
    // dans le chargement de la TobACT, il n'y a que les types d'articles demandées
    // dans le mode de génération de l'affaire.
    // ex on peux vouloir que les frais au niveau affaire, et une prestation depuis une
    // ligne d'affaire
    if ( Typgen <> 'APP') and ((CodeAct <> 'NON') and  (CodeAct <> '')) then
      ChargeActivite(CodeAct,CodeAff,TobAct);

    if (GetParamSoc('SO_AFGENERAUTOLIG')=true) and (TypGen = 'POT') then
      ChargeActivite('TOU',CodeAff,TobActLig);

    if GetParamSoc('SO_AFVARIABLES') then
    begin
//      LoadTobVarQte(Nil,TobVarQte_ACT,'ACT');
      LoadTobVarQte(TobVarQte_ACT, TobAffaire, 'ACT', CleDoc);
      TobVarQte_ACTbis.dupliquer(TobVarQte_ACT,true,True,True);
    end;

    if VH_GC.GCIfDefCEGID then
    begin
      Comment := '';
      LigneComAff('COM',Comment);
    end;

  if (GenFac.zcomaff <> '') then     //Commentaire Affaire
    Begin
    Comment := Genfac.zcomaff;
    LigneComAff('C10',Comment);
    End;
  if (GenFac.zcomaffsup <> '') then     //Commentaire Affaire
    Begin
    Comment := Genfac.zcomaffsup;
    LigneComAff('C10',Comment);
    End;

    // reprise des commentaires lignes dans l'affaire
    // On reprend les lignes jusqu'à la 1er prestation
    // sauf pour facturation sur affaire (dans ce cas reprise de tous les commentaires)
    if (GenFac.zcomlig) then
    Begin
      LoadTobLignesCom;
    End;

     // Traitement au Pourcentage
    if ( (TypGen = 'POU') Or (TypGen = 'POT') Or (TypGen = 'CON') ) then
    Begin
        if typgen = 'CON' then
        begin
          Comment := rech.Libelleeche;
          LigneComAff('COM',Comment);
        end;
        // Traitement des lignes de la piéce
        LoadTobLignes(TypGen,'TOU',Nblig);
        //
        ZeroFacture (TOBPiece);
        for Ind := 0 to TOBPiece.detail.count -1 do
        begin
        	ZeroLigneMontant (TOBPiece.detail[Ind]);
          if (LigneQteUnitaire < 0) then
          begin
          	if (TOBpiece.detail[ind].getValue('GL_QTEFACT')= 1) and
            	 (TOBpiece.detail[ind].getValue('GL_TYPELIGNE')= 'ART') then LigneQteUnitaire := Ind;
          end;
        end;
        TOBOuvragesP.ClearDetail;
        TOBBases.ClearDetail;
        TOBBasesL.ClearDetail;
        ZeroMontantPorts (TOBPorcs);
        TOBVTECOLL.ClearDetail;
        PutValueDetail(TOBPiece,'GP_RECALCULER','X') ;
        CalculFacture(TOBAffaire,TOBPiece,nil,nil,TobOuvrages ,TOBOuvragesP,TOBBases,TOBBasesL,TOBTiers,TOBArticles,TobPorcs,Nil,Nil,TOBVTECOLL,DEV);
        Ecart := Arrondi(rech.montEch - TOBPiece.GetValue('GP_TOTALHTDEV'),V_PGI.okdecP);
        if (Ecart <> 0) and (LigneQteUnitaire >= 0) then
        begin
        	TOBPiece.detail[LigneQteUnitaire].PutValue('GL_PUHTDEV',TOBPiece.detail[LigneQteUnitaire].GetValue('GL_PUHTDEV')+Ecart);
          ZeroFacture (TOBPiece);
          for Ind := 0 to TOBPiece.detail.count -1 do
          begin
            ZeroLigneMontant (TOBPiece.detail[Ind]);
          end;
          TOBOuvragesP.ClearDetail;
          TOBBases.ClearDetail;
          TOBBasesL.ClearDetail;
          ZeroMontantPorts (TOBPorcs);
          PutValueDetail(TOBPiece,'GP_RECALCULER','X') ;
          TOBVTECOLL.ClearDetail;
          CalculFacture(TOBAffaire, TOBPiece,nil,nil,TobOuvrages ,TOBOuvragesP,TOBBases,TOBBasesL,TOBTiers,TOBArticles,TobPorcs,Nil,Nil,TOBVTECOLL,DEV);
        end else if (Ecart <> 0) and (ArtEcart <> '') then
        begin
          AjouteLigneEcart(TOBPiece,Ecart);
        end;
        //
        // Si reg sur caf et affaire principale et pas de lignes détail, on supprime les enregs dans la tob
        // ca veut dire que l'affaire pale est là seulement pour les stats et
        // les conditions de réglement
        If (nblig = 0) and (princ = 'X') and (regsurcaf) then
        	Otesleslignes(codeaff)
    End else    // si n'est pas sutr ligne, on peut vouloir slt la ligne %
    Begin
        if (Genfac.zpour) then LoadTobLignes(TypGen,'POU',Nblig);
    end;

    // Traitement au Forfait
    if ( Typgen = 'FOR') then
        Begin
        LigneAcompte;
        End;
    if ( Typgen = 'APP') then   // Apprec
        Begin
        LigneApprec;
        End;


    if (GetParamSoc('SO_AFGENERAUTOLIG')=true) then
      TobActLig.ClearDetail;

End;


// ************************************************************
// Suppresion des lignes  déjà crées pour une  affaire

// ************************************************************

procedure TAFPrepFact.Otesleslignes(CodeAff : String);
Var TobDet : TOB;
    wj : Integer;
BEGIN
   Wj :=0;
		while (wj < TOBPiece.Detail.Count) do
        BEGIN
        TobDet:=TOBPiece.Detail[wj] ;
        if (tobDet.getValue('GL_AFFAIRE') = CodeAff) then
        	TobDet.free
        else
        	inc(wj);
        END;

END;

// *****************************************************************
//  Alimentation de la ligne par rapport aus données de ARTICLE, etc...

//  zori : A (acompte,ligne affaire)

//         T (activité)

//         C (commentaire)

// Ajout d'un champ AFFREF pour garder l'affaire principale ou de référence

// en tête lors des tris

// Ajout champ sup , DATACT  date activité pour tri

// *****************************************************************
procedure TAFPrepFact.CreationLignePiece(Toblignes : TOB;zqte,zpuv : double;zlib,zori : string);
Var TOBA,Tobart : TOB;
    refunique: string;
    zaff,zdat : String;
    ind : integer;
    QQ : Tquery;
    zcalc : double;
    SavIndice : integer;
BEGIN
    SavIndice := TOBLignes.GetValue('GL_INDICENOMEN');
    ind := TOBPiece.Detail.Count;

    RefUnique := Toblignes.GetValue('GL_ARTICLE');

    if (RefUnique <> '') then
       Begin
       TOBArt:=FindTOBArtSais(TOBArticles,RefUnique) ;
       if TOBArt<>Nil then
          BEGIN
          // CodeArticle:=TOBArt.GetValue('GA_CODEARTICLE') ;
          END
       else
          BEGIN
 					//***Création des pieces, il faut tout prendre ****
          QQ:=OpenSQL('SELECT * FROM ARTICLE WHERE GA_ARTICLE="'+RefUnique+'" ',True,-1,'',true) ;
          if Not QQ.EOF then
             BEGIN
             TOBArt:=CreerTOBArt(TOBArticles) ; TOBArt.SelectDB('',QQ) ;
             TOBArt.addchampsup('GCA_LIBELLE',false);
//           CodeArticle:=QQ.FindField('GA_CODEARTICLE').AsString ;
             END;
          Ferme(QQ) ;
          END ;
       TOBA:=FindTOBArtRow(TOBPiece,TOBArticles,ind) ;
       TOBA.putvalue('GCA_LIBELLE',zlib);
       TOBA.putvalue('GA_PVHT',zpuv); // pour reprendre notre prix et pas celui de la fiche article
       TOBA.putvalue('GA_PCB',0);  // sinon on reprend systématiquement GA_PCB et non notre Quantité
       End
    else
       Toba := NIL;

    // TobCOnd= NIL
		// PreAffecteLigne (TOBPiece,TOBLignes,TOBA,TOBTiers,TOBTarif,Nil,TOBAffaire,ind, DEV ,zqte );
    InitLigneVide(TOBPiece,TOBLignes,TOBTiers,TobAffaire,ind,zQte) ;
    TobLIgnes.putvalue ('GL_REFARTSAISIE',''); // mcd 26/02/2003 pour OK sur fct ArticleVrs lignes
    ArticleVersLigne(TOBPiece,TOBA,TOBConds,TOBLignes,TOBTiers) ;
    TobLignes.PutValue ('GL_REFARTSAISIE',TobLIgnes.getvalue('Gl_CODEARTICLE')); // mcd 26/02/03
    TobLignes.PutValue ('GL_INDICENOMEN',SavIndice);

    if (zori = 'A') then
       Begin
       TOBLignes.PutValue('GL_PUHTDEV',zpuv) ;
       TOBLignes.PutValue('GL_PUHT',DeviseToPivot(zpuv,DEV.Taux,DEV.Quotite));
       end
    else
       Begin
       TOBLignes.PutValue('GL_PUHT',zpuv) ;
       TOBLignes.PutValue('GL_PUHTDEV',PivotToDevise(zpuv,DEV.Taux,DEV.Quotite,DEV.Decimale)) ;
       end;

    zcalc:= Arrondi(Toblignes.GetValue('GL_QTEFACT') *  Toblignes.GetValue('GL_PUHTDEV'),V_PGI.OkDecV);
    Toblignes.PutValue('GL_TOTALHTDEV',zcalc);
    // MODIF FV 05-06-07
    //TOBLignes.PutValue('GL_INDICENOMEN',SavIndice ); // restauration de lindice ouvrage
//    TOBLignes.PutValue('GL_INDICENOMEN',0 );
    TraitePrixLigneCourante (TOBLIgnes,zPuv);
    // ---
    TobLignes.PutValue('GL_AFFAIRE',TOBAffaire.GetValue('AFF_AFFAIRE'));
    TobLignes.PutValue('GL_AFFAIRE1',TOBAffaire.GetValue('AFF_AFFAIRE1'));
    TobLignes.PutValue('GL_AFFAIRE2',TOBAffaire.GetValue('AFF_AFFAIRE2'));
    TobLignes.PutValue('GL_AFFAIRE3',TOBAffaire.GetValue('AFF_AFFAIRE3'));
	  TobLignes.PutValue('GL_AVENANT',TOBAffaire.GetValue('AFF_AVENANT'));
    TobLignes.PutValue('GL_REPRESENTANT',TOBAffaire.GetValue('LEREP'));
    zaff := Toblignes.getValue('GL_AFFAIRE');
    zdat := DatetoStr(idate2099);
    GestionChampSuppLigne(Toblignes,zaff,zdat,TobAffaire.GetValue('AFF_TIERS'), AffRef);

END;

// ************************************************************
//  Préparation de la ligne d'acompte pour facture au Forfait

// ************************************************************

procedure TAFPrepFact.LigneAcompte;
Var TobLignes : TOB;
    s,rlib,dim : string;
    zqte : integer;
    zpuv :double;
Begin
    TobLignes := NewTOBLigne( TOBPiece,0);

    TobLignes.PutValue('GL_ARTICLE',Genfac.zacompte);
    TobLignes.PutValue('GL_TYPEARTICLE','PRE');
    TobLignes.PutValue('GL_CODEARTICLE',trim(CodeArticleGenerique(Genfac.zacompte,dim,dim,dim,dim,dim)));
    TobLignes.PutValue('GL_REFARTSAISIE',trim(CodeArticleGenerique(Genfac.zacompte,dim,dim,dim,dim,dim)));

    // formattage du libellé paramètré
    if (Genfac.zreplibech = 'X') then
      Begin
      rlib := Rech.libelleeche;
      End
    else
      Begin
      xdeb := '01/01/1900'; xfin := '31/12/2099'; xemp := '';
      s := '{'+Genfac.zlibacompte+'}';
      rlib := GFormule(s,GetData,nil,1);
      rlib := trim(copy(rlib,1,70)); // gm 19/11/02
      End;
    zqte := 1;
    zpuv := Rech.MontEch;
    CreationLignePiece(Toblignes,zqte,zpuv,rlib,'A');
End;
// Création des lignes appréciation
//  plusieurs cas possibles
//
procedure TAFPrepFact.LigneApprec;
Var TobLignes : TOB;
    rlib,comment : Hstring;
    codeart,dim : string;
    zqte : integer;
    zpuv :double;
Begin

 	if (TobEchFFact.GetValue('AFA_LIBELLE1') <> '') then
    Begin
    	Comment := TobEchFFact.GetValue('AFA_LIBELLE1');
    	LigneComAff('C00',Comment);
    End;
    if (TobEchFFact.GetValue('AFA_LIBELLE2') <> '') then
    Begin
    	Comment := TobEchFFact.GetValue('AFA_LIBELLE2');
    	LigneComAff('C00',Comment);
    End;
    if (TobEchFFact.GetValue('AFA_LIBELLE3') <> '') then
    Begin
    	Comment := TobEchFFact.GetValue('AFA_LIBELLE3');
    	LigneComAff('C00',Comment);
    End;

    if (TobEchFFact.GetValue('AFA_AFACTURERDEV') <> 0) then
    Begin
      TobLignes := NewTOBLigne( TOBPiece,0);
      codeart := CodeArticleUnique2(GetParamSoc ('SO_AFAPPPRES'),'');
      TobLignes.PutValue('GL_ARTICLE',codeart);
      TobLignes.PutValue('GL_TYPEARTICLE','PRE');
      TobLignes.PutValue('GL_CODEARTICLE',trim(CodeArticleGenerique(codeart,dim,dim,dim,dim,dim)));
      TobLignes.PutValue('GL_REFARTSAISIE',trim(CodeArticleGenerique(codeart,dim,dim,dim,dim,dim)));

      rlib := 'Honoraires';

      zqte := 1;
      zpuv := Tobechffact.GetValue('AFA_AFACTURERDEV');
      if (Genfac.znat = 'APR') then begin zpuv := zpuv * (-1);zqte := (-1);end;
      CreationLignePiece(Toblignes,zqte,zpuv,rlib,'A');
    End;

    if (TobEchFFact.GetValue('AFA_AFACTFODEV') <> 0) then
    Begin
      TobLignes := NewTOBLigne( TOBPiece,0);
      codeart := CodeArticleUnique2(GetParamSoc ('SO_AFAPPFOUR'),'');
      TobLignes.PutValue('GL_ARTICLE',codeart);
      TobLignes.PutValue('GL_TYPEARTICLE','MAR');
      TobLignes.PutValue('GL_CODEARTICLE',trim(CodeArticleGenerique(codeart,dim,dim,dim,dim,dim)));
      TobLignes.PutValue('GL_REFARTSAISIE',trim(CodeArticleGenerique(codeart,dim,dim,dim,dim,dim)));

      rlib := 'Fournitures';

      zqte := 1;
      zpuv := Tobechffact.GetValue('AFA_AFACTFODEV');
      if (Genfac.znat = 'APR') then begin zpuv := zpuv * (-1);zqte := (-1);end ;
      CreationLignePiece(Toblignes,zqte,zpuv,rlib,'A');
    End;

    if (TobEchFFact.GetValue('AFA_AFACTFRDEV') <> 0) then
    Begin
      TobLignes := NewTOBLigne( TOBPiece,0);
      codeart := CodeArticleUnique2(GetParamSoc ('SO_AFAPPFRAIS'),'');
      TobLignes.PutValue('GL_ARTICLE',codeart);
      TobLignes.PutValue('GL_TYPEARTICLE','FRA');
      TobLignes.PutValue('GL_CODEARTICLE',trim(CodeArticleGenerique(codeart,dim,dim,dim,dim,dim)));
      TobLignes.PutValue('GL_REFARTSAISIE',trim(CodeArticleGenerique(codeart,dim,dim,dim,dim,dim)));

      rlib := 'Frais';

      zqte := 1;
      zpuv := Tobechffact.GetValue('AFA_AFACTFRDEV');
      if (Genfac.znat = 'APR') then begin zpuv := zpuv * (-1);zqte := (-1);end;
      CreationLignePiece(Toblignes,zqte,zpuv,rlib,'A');
    End;

    zpuv := Tobechffact.GetValue('AFA_ACPTEPRDEV');
    if (zpuv <> 0) then
    Begin
      TobLignes := NewTOBLigne( TOBPiece,0);
      codeart := CodeArticleUnique2(GetParamSoc ('SO_AFAPPPRES'),'');
      TobLignes.PutValue('GL_ARTICLE',codeart);
      TobLignes.PutValue('GL_TYPEARTICLE','PRE');
      TobLignes.PutValue('GL_CODEARTICLE',trim(CodeArticleGenerique(codeart,dim,dim,dim,dim,dim)));
      TobLignes.PutValue('GL_REFARTSAISIE',trim(CodeArticleGenerique(codeart,dim,dim,dim,dim,dim)));

      rlib := 'Reprise d''acompte prestations';

      zqte := -1;
      CreationLignePiece(Toblignes,zqte,zpuv,rlib,'A');
    End;

    zpuv := Tobechffact.GetValue('AFA_ACPTEFRDEV');
    if (zpuv <> 0) then
    Begin
      TobLignes := NewTOBLigne( TOBPiece,0);
      codeart := CodeArticleUnique2(GetParamSoc ('SO_AFAPPFRAIS'),'');
      TobLignes.PutValue('GL_ARTICLE',codeart);
      TobLignes.PutValue('GL_TYPEARTICLE','FRA');
      TobLignes.PutValue('GL_CODEARTICLE',trim(CodeArticleGenerique(codeart,dim,dim,dim,dim,dim)));
      TobLignes.PutValue('GL_REFARTSAISIE',trim(CodeArticleGenerique(codeart,dim,dim,dim,dim,dim)));

      rlib := 'Reprise d''acompte frais';

      zqte := -1;
      CreationLignePiece(Toblignes,zqte,zpuv,rlib,'A');
    End;

    zpuv := Tobechffact.GetValue('AFA_ACPTEFODEV');

    if (zpuv <> 0) then
    Begin
      TobLignes := NewTOBLigne( TOBPiece,0);
      codeart := CodeArticleUnique2(GetParamSoc ('SO_AFAPPFOUR'),'');
      TobLignes.PutValue('GL_ARTICLE',codeart);
      TobLignes.PutValue('GL_TYPEARTICLE','MAR');
      TobLignes.PutValue('GL_CODEARTICLE',trim(CodeArticleGenerique(codeart,dim,dim,dim,dim,dim)));
      TobLignes.PutValue('GL_REFARTSAISIE',trim(CodeArticleGenerique(codeart,dim,dim,dim,dim,dim)));

      rlib := 'Reprise d''acompte fournitures';

      zqte := -1;
      CreationLignePiece(Toblignes,zqte,zpuv,rlib,'A');
    End;

End;


// ************************************************************
//  Préparation de la ligne Commentaire début affaire

// ************************************************************

procedure TAFPrepFact.LigneComAff(Typcom : string;var Comment : Hstring);

Var TobLignes : TOB;
    s,rlib : string;
    zaff,zdat : string;
Begin
     TobLignes := NewTOBLigne( TOBPiece,0);

     TobLignes.PutValue('GL_ARTICLE','');
     // pour identifier type de commentaire par rapport au profil
     TobLignes.PutValue('GL_TYPELIGNE',Typcom);
     // formattage du libellé paramètré
     xdeb := '01/01/1900'; xfin := '31/12/2099'; xemp := '';
     s := '{'+Comment+'}';

     if comment = '.' then rlib := '' else rlib := GFormule(s,GetData,nil,1);
     rlib := trim(copy(rlib,1,70)); // gm 19/11/02
     TobLignes.PutValue('GL_LIBELLE',rlib);

     PieceVersLigne ( TOBPiece,TobLignes ) ;

     TobLignes.PutValue('GL_AFFAIRE',TOBAffaire.GetValue('AFF_AFFAIRE'));
     TobLignes.PutValue('GL_AFFAIRE1',TOBAffaire.GetValue('AFF_AFFAIRE1'));
     TobLignes.PutValue('GL_AFFAIRE2',TOBAffaire.GetValue('AFF_AFFAIRE2'));
     TobLignes.PutValue('GL_AFFAIRE3',TOBAffaire.GetValue('AFF_AFFAIRE3'));
     TobLignes.PutValue('GL_AVENANT',TOBAffaire.GetValue('AFF_AVENANT'));
     TobLignes.PutValue('GL_REPRESENTANT',TOBAffaire.GetValue('LEREP'));

     zaff := Toblignes.getValue('GL_AFFAIRE');
     zdat := DatetoStr(idate2099);
     GestionChampSuppLigne(Toblignes,zaff,zdat,TobAffaire.GetValue('AFF_TIERS'),AffRef);

End;



// ************************************************************************
//  Chargement des lignes et extraction des lignes Article
// ************************************************************************
procedure TAFPrepFact.LoadTobLignes(Typgen,ztyp : String;var nblig : integer) ;
Var Q : TQuery ;
    TobPiece_AFF,TobVarQte_AFF : TOB;
    req: string;
    StFact : String;

BEGIN
// Lecture Lignes de l'affaire, application du % si necessaire

  TobPiece_AFF:=TOB.Create('PIECE',Nil,-1) ;
  if GetParamSoc('SO_AFVARIABLES') then
    TobVarQte_AFF:=TOB.Create('Les qtés variables de affaire',Nil,-1) ;

  StFact :=  ' and GL_FACTURABLE <> "N"';    //GMONYX

  if (ztyp = 'POU') then
    begin
            //***Création des pieces, il faut tout prendre ****
        req := 'SELECT *,GP_TOTALHTDEV FROM LIGNE LEFT JOIN PIECE ON GP_NATUREPIECEG=GL_NATUREPIECEG AND '+
               'GP_SOUCHE=GL_SOUCHE AND GP_NUMERO=GL_NUMERO AND GP_INDICEG=GL_INDICEG WHERE '+
               WherePiece(CleDocOrigine,ttdLigne,False)+
        'and GL_TYPELIGNE="ART"  and GL_TYPEARTICLE="POU"'+ StFact +
        ' ORDER BY GL_NUMLIGNE'
    End
  else
    Begin
      if  not(Genfac.ztout) then
              //***Création des pieces, il faut tout prendre ****
        req := 'SELECT *,GP_TOTALHTDEV FROM LIGNE LEFT JOIN PIECE ON GP_NATUREPIECEG=GL_NATUREPIECEG AND '+
               'GP_SOUCHE=GL_SOUCHE AND GP_NUMERO=GL_NUMERO AND GP_INDICEG=GL_INDICEG WHERE '+
        WherePiece(CleDocOrigine,ttdLigne,False)+
        'and GL_TYPELIGNE="ART" '+  Stfact +
        ' ORDER BY GL_NUMLIGNE'
      else
        req := 'SELECT *,GP_TOTALHTDEV FROM LIGNE LEFT JOIN PIECE ON GP_NATUREPIECEG=GL_NATUREPIECEG AND '+
               'GP_SOUCHE=GL_SOUCHE AND GP_NUMERO=GL_NUMERO AND GP_INDICEG=GL_INDICEG WHERE '+
        WherePiece(CleDocOrigine,ttdLigne,False)+
        'and (GL_TYPELIGNE="ART" or GL_TYPELIGNE="COM" or GL_TYPELIGNE="TOT")'+   StFact +
        ' ORDER BY GL_NUMLIGNE';
    End;

  Q:=OpenSQL(Req,True,-1,'',true) ;
  TobPiece_AFF.LoadDetailDB('LIGNE','','',Q,False) ;  // ajout dans la tob
  Ferme(Q) ;

  if GetParamSoc('SO_AFVARIABLES') then
//    LoadTobVarQte(TobVarQte_AFF,Nil,'LIG');
    LoadTobVarQte(TobVarQte_AFF, TobAffaire, 'LIG', CleDocOrigine);

  if GetParamSoc('SO_AFVARIABLES') then
    TrtLignesAffaire(TypGen,ztyp ,TobPiece_AFF,TobVarQte_AFF)
  else
    TrtLignesAffaire(TypGen,ztyp ,TobPiece_AFF,Nil);

  TobPiece_AFF.free;

  if GetParamSoc('SO_AFVARIABLES') then
    TobVarQte_AFF.free;

END ;


procedure TAFPrepFact.TrtLignesAffaire(Typgen,ztyp : String;TobPiece_AFF,TobVarQte_AFF : TOB);
var wi,INbLigneValo : Integer;
    zqte,zpuv,Zprorata,ZmtEche: double;
    rlib,zaff,zdat: String;
    TobDetLig,TobLignes,TobQte: TOB;
begin
  Zqte :=0;
  //
  INbligneValo := 0;
  for wi:=0 to  TobPiece_AFF.detail.count-1  do
  Begin
    if TobPiece_AFF.detail[wi].GetValue('GL_TYPELIGNE')='ART' then Inc(INbLigneValo);
  end;
  //
  for wi:=0 to  TobPiece_AFF.detail.count-1  do
  Begin
    TobDetLig := TobPiece_AFF.detail[wi];

    if (TobdetLig.GetValue('GL_LIGNELIEE') <> 0) then continue;

    if (GetParamSoc('SO_AFGENERAUTOLIG')=true) and
        (TobdetLig.GetValue('GL_TYPELIGNE') = 'ART') and
        (TobDetLig.GetValue('GL_GENERAUTO') = 'ACT') then
    begin
      TrtActiviteSurLigneAffaire(TobDetLig,TobPiece_AFF,TobVarQte_AFF);
      Continue;
    End;

    TobLignes := NewTOBLigne( TOBPiece,0);
    TOBCopyFieldValues(TobDetLig,TobLignes,['GL_ARTICLE','GL_CODEARTICLE',
    'GL_TYPELIGNE','GL_TYPEARTICLE','GL_REFARTSAISIE','GL_RESSOURCE','GL_REMISELIGNE',
    'GL_NONIMPRIMABLE','GL_AFFAIRE','GL_AFFAIRE1','GL_AFFAIRE2','GL_AFFAIRE3','GL_AVENANT',
    'GL_PRIXPOURQTE','GL_FORMULEVAR','GL_GENERAUTO','GL_FACTURABLE','GL_FORCODE1','GL_FORCODE2'
    ,'GL_NUMORDRE','GL_VALEURREMDEV','GL_NUMLIGNE','GL_FAMILLETAXE1']);

    if (TobDetLig.GetValue('GL_TYPELIGNE') = 'ART') then
    Begin
      // MODIF FV 05-06-07
      TobLignes.PutValue('GL_NATUREPIECEG',TOBDETLIG.GetValue('GL_NATUREPIECEG'));
      TobLignes.PutValue('GL_SOUCHE',TOBDETLIG.GetValue('GL_SOUCHE'));
      TobLignes.PutValue('GL_NUMERO',TOBDETLIG.GetValue('GL_NUMERO'));
      TobLignes.PutValue('GL_INDICEG',TOBDETLIG.GetValue('GL_INDICEG'));
      TobLignes.PutValue('GL_DATEPIECE',TOBDETLIG.GetValue('GL_DATEPIECE'));
      UG_LoadOuvrageLigne (TobLignes,TobOuvrages,TobArticles,IndiceNomen,false);
      //
      UG_LoadAnaLigne ( tobpiece_o,TOBLignes ) ;

      rlib := TobDetLig.GetValue('GL_LIBELLE');
      if ( Typgen = 'POT' ) or (ztyp = 'POU') then zqte := TobDetLig.GetValue('GL_QTEFACT');

      if GetParamSoc('SO_AFVARIABLES') then
      begin
        TobQte := RechVariable(TobVarQte_AFF,TobDetLig,Nil,'LIG');
        if (TobQte <> Nil) and (TobQte.getvalue('AVV_QTECALCUL')<>0)  then
          zqte := TobQte.getvalue('AVV_QTECALCUL');
      end;

      if  TobDetLig.getvalue('GL_TYPEARTICLE')='POU' then zqte:=  TobDetLig.GetValue('GL_QTEFACT'); // mcd 21/10/02
      // mcd 21/10/02 il ne faut pas affecté le % sur article POU if ( Typgen = 'POU' ) then
      if ( Typgen = 'POU' )and (TobDetLig.getvalue('GL_TYPEARTICLE')<>'POU') then
        Begin
        zqte := TobDetLig.GetValue('GL_QTEFACT');
        zqte := Arrondi(zqte*rech.pourech/100.0,V_PGI.OkDecQ) ;
        End;
      if (INbLigneValo = 1) then
      begin
        zpuv := Rech.MontEch ;
        zqte := Arrondi(TobDetLig.GetValue('GL_QTEFACT')* xprorata,V_PGI.OkDecQ);
        if ( Typgen = 'CON' ) then
        Begin
        	zQte := 1;
          zPuv := Rech.MontEch ;
        End;
      end else
      begin
        zqte := Arrondi(TobDetLig.GetValue('GL_QTEFACT'),V_PGI.OkDecQ);
        if ( Typgen = 'CON' ) then
        Begin
          ZmtEche := rech.montEch;
          Zprorata := TobDetLig.GetValue('GL_PUHTDEV')/TobDetLig.GetDouble ('GP_TOTALHTDEV');
          zpuv := Arrondi(ZmtEche * Zprorata,V_PGI.okdecP);
        End;
      end;
      CreationLignePiece(Toblignes,zqte,zpuv,rlib,'A');
      // reprise des infos personnalisées dans la ligne
      TOBCopyFieldValues(TobDetLig,TobLignes,['GL_LIBREART1','GL_LIBREART2','GL_LIBREART3','GL_LIBREART4','GL_LIBREART5',
      'GL_LIBREART6','GL_LIBREART7','GL_LIBREART8','GL_LIBREART9','GL_LIBREARTA',
      'GL_FAMILLENIV1','GL_FAMILLENIV2','GL_FAMILLENIV3','GL_DPR','GL_LIBCOMPL','GL_FAMILLETAXE1']);

      if GetParamSoc('SO_AFVARIABLES') then
        AlimTobVarQte(TobVarQte_LIG, Toblignes, nil, TobVarQte_AFF, 'LIG', iCpt);

     End  // fin  ART
  else
    Begin   // commnentaires
    TOBCopyFieldValues(TobDetLig,TobLignes,['GL_ARTICLE','GL_CODEARTICLE','GL_TYPELIGNE',
    'GL_LIBELLE','GL_AFFAIRE','GL_AFFAIRE1','GL_AFFAIRE2','GL_AFFAIRE3','GL_AVENANT','GL_LIBCOMPL']);
    PieceVersLigne ( TOBPiece,TobLignes ) ;
     TobLignes.PutValue('GL_AFFAIRE',TobDetLig.GetValue('GL_AFFAIRE'));
     TobLignes.PutValue('GL_AFFAIRE1',TobDetLig.GetValue('GL_AFFAIRE1'));
     TobLignes.PutValue('GL_AFFAIRE2',TobDetLig.GetValue('GL_AFFAIRE2'));
     TobLignes.PutValue('GL_AFFAIRE3',TobDetLig.GetValue('GL_AFFAIRE3'));
     TobLignes.PutValue('GL_AVENANT',TobDetLig.GetValue('GL_AVENANT'));
     TobLignes.PutValue('GL_REPRESENTANT',TOBAffaire.GetValue('LEREP')); //1911
     zaff := TobDetLig.getValue('GL_AFFAIRE');
     zdat := DatetoStr(idate2099);
     GestionChampSuppLigne(Toblignes,zaff,zdat ,TobAffaire.GetValue('AFF_TIERS'),AffRef);
    End;
  End;   // for
End;


{***********A.G.L.***********************************************
Auteur  ...... : Merieux
Créé le ...... : 24/06/2003
Modifié le ... :   /  /
Description .. : Génération de la tob des variables pour les articles liés
Mots clefs ... : GIGA;FACTURATION;
*****************************************************************}
function TAFPrepFact.TrtLigneLiee(TobLig,TobAct,TobVarQte_AFF : TOB): boolean;
var stAff,stFormule : string;
    iNumligUnique,iNumOrdre : integer;
    TobVarDetAct,TobVarDetLig,TobVarNew : TOB;
    Qte : double;
    TobVarMixte: TOB; // tob composée des tob ACT et LIG pour appel fonction evaluationafformule
                      // la fonction me rend  un 3eme détail GEN
begin
    result := true;
    TobVarMixte  := Tob.Create('Les variables de LIG etACT ',nil,-1);

    stAff         :=  TobAct.getvalue('ACT_AFFAIRE');
    iNumligUnique :=  TobAct.getvalue('ACT_NUMLIGNEUNIQUE');
    iNumordre     :=  TobLig.GetValue('GL_NUMORDRE');

    TobVarDetAct  := TobVarQte_ACTbis.FindFirst(['AVV_AFFAIRE','AVV_NUMLIGNEUNIQUE'],[StAff,iNumligUnique],true);
    TobVarDetLig  := TobVarQte_AFF.FindFirst(['AVV_AFFAIRE','AVV_NUMORDRE'],[StAff,iNumOrdre],true);
    if (TobVarDetLig = Nil) or (TobVarDetAct = Nil) then result := false;

    if result then
    begin
      TOB.Create('',TobVarMixte,-1) ;
      TOB.Create('',TobVarMixte,-1) ;

      TobVarMixte.detail[0].dupliquer(TobVarDetLig,false,True,true);
      TobVarMixte.detail[1].dupliquer(TobVarDetAct,false,True,true);

      // mise à jour de la tob var lig à partir des éléments variables de l'activité :

      stFormule :=  TobVarDetAct.GetValue('AVV_FORMULEVAR');
      Qte := EvaluationAFFormuleVar ('GEN',StFormule,'',TOBFormuleVar,TOBVarMixte,TobLig);
      Toblig.PutValue('GL_QTEFACT',Qte);

      TobVarNew  := TOBVarMixte.detail[0];
      FinalisationTobVar(TobVarQte_LIG, TobVarNew, Toblig, iCpt);

    end;

    TobVarMixte.free;
end;



procedure TAFPrepFact.TrtActiviteSurLigneAffaire(TobDetLig ,TobPiece_AFF,TobVarQte_AFF: TOB);
Var TobActLigRes,TobActRes : TOB;
    StAff : String;
    iNumOrdre,INumLigAct : Integer;
    ok : boolean;
begin
  // recherche dans TobACTLIG pour savoir si des lignes d'activité sont rattachées
  // à cette ligne d'affaire
  StAff := TobDetLig.GetValue('GL_AFFAIRE');
  iNumOrdre := TobDetLig.GetValue('GL_NUMORDRE');

  ok := true;
  while (ok) do
  begin
    TobActLigRes := TobActLig.FindFirst(['ACT_AFFAIRE','ACT_NUMORDRE'],[StAff,iNumOrdre],true);
    if (tobActLigRes = Nil)  then  exit;

    // création ligne depuis ligne d'activité
    CreationLigDepuisActivite(TobActLigRes,TobDetLig);

    TrtLignesAffairesliees(TobActLigRes,TobDetLig,TobPiece_aff,TobVarQte_AFF);

    // Traitement de cette ligne d'actvité , si elle n'existe pas dans TobACT, on la crée
    // afin de la mettre à jour à la  fin de lma prépa
    // sinon on mets à jour ACTIVITEREPRIS = FAC pour ne pas la traiter une 2eme fois.
    // dans decharge activite
    iNumLigAct := TobActLigRes.GetValue('ACT_NUMLIGNEUNIQUE');
    TobActRes := TobAct.FindFirst(['ACT_AFFAIRE','ACT_NUMLIGNEUNIQUE'],[StAff,iNumLigAct],true);
    if (TobActRes <> NIL) then
    begin
      TobActRes.PutValue('ACT_ACTIVITEREPRIS','FAC');
      TobActLigRes.free;
    end
    else
    begin
      TobActLigRes.PutValue('ACT_ACTIVITEREPRIS','FAC');
      TobActLigRes.ChangeParent(TobAct,-1);
    end
  end;

End;

procedure TAFPrepFact.TrtLignesAffairesliees(TobAct,TobDetLig,TobPiece_aff,TobVarQte_AFF: TOB);
Var TobLigLiee : TOB;
    StAff : String;
    iNumLiee : Integer;
begin
  // recherche dans TobPiece_aff des lignes d'affaires liées
  StAff := TobDetLig.GetValue('GL_AFFAIRE');
  iNumLiee := TobDetLig.GetValue('GL_NUMORDRE');
  if (TobPiece_aff.FindFirst(['GL_AFFAIRE','GL_LIGNELIEE'],[StAff,iNumLiee],true)) = Nil then exit;

  TobLigLiee := TobPiece_aff.FindFirst(['GL_AFFAIRE','GL_LIGNELIEE'],[StAff,iNumLiee],true);
  while (TobLigLiee <> Nil ) do
  begin
    CreationLigDepuisActiviteLiee(TobAct,TobLigLiee,TobVarQte_AFF);
    TobLigLiee := TobPiece_aff.FindNext(['GL_AFFAIRE','GL_LIGNELIEE'],[StAff,iNumLiee],true);
  end;

End;

// ************************************************************************
//  Chargement des lignes et extraction des commentaires de début d'affaire
// ************************************************************************
procedure TAFPrepFact.LoadTobLignesCom ;
Var Q : TQuery ;
    TobPiece_AFF ,TobDet,TobLignes: TOB;
    req: string;
    wi : Integer;
    zaff,zdat : string;
BEGIN
TOBPiece_AFF:=TOB.Create('PIECE',Nil,-1) ;
      //***Création des pieces, il faut tout prendre ****
req := 'SELECT * FROM LIGNE WHERE '+WherePiece(CleDocOrigine,ttdLigne,False)+' ORDER BY GL_NUMLIGNE';
Q:=OpenSQL(Req,True,-1,'',true) ;
TobPiece_AFF.LoadDetailDB('LIGNE','','',Q,False) ;  // ajout dans la tob

Ferme(Q) ;


for wi:=0 to  TobPiece_AFF.detail.count-1  do
Begin
  TobDet := TobPiece_AFF.detail[wi];
  if (TobDet.GetValue('GL_TYPELIGNE') <> 'COM') then exit;
  if (TobDet.GetValue('GL_TYPELIGNE') = 'COM') then
    Begin
    TobLignes := NewTOBLigne( TOBPiece,0);

    TOBCopyFieldValues(TobDet,TobLignes,['GL_ARTICLE','GL_CODEARTICLE','GL_TYPELIGNE',
    'GL_LIBELLE','GL_AFFAIRE','GL_AFFAIRE1','GL_AFFAIRE2','GL_AFFAIRE3','GL_AVENANT']);
    TobLignes.PutValue('GL_TYPELIGNE','C00'); //pour reconnaitre le type commentaire
    PieceVersLigne ( TOBPiece,TobLignes ) ;

    TobLignes.PutValue('GL_AFFAIRE',TOBDet.GetValue('GL_AFFAIRE'));
    TobLignes.PutValue('GL_AFFAIRE1',TOBDet.GetValue('GL_AFFAIRE1'));
    TobLignes.PutValue('GL_AFFAIRE2',TOBDet.GetValue('GL_AFFAIRE2'));
    TobLignes.PutValue('GL_AFFAIRE3',TOBDet.GetValue('GL_AFFAIRE3'));
    TobLignes.PutValue('GL_AVENANT',TOBDet.GetValue('GL_AVENANT'));

    zaff := TOBLignes.GetValue('GL_AFFAIRE');
    zdat := DatetoStr(idate2099);
    GestionChampSuppLigne(Toblignes,zaff,zdat,TobAffaire.GetValue('AFF_TIERS'),AffRef);
    End;
End;

TobPiece_AFF.free;

END ;

function TAFPrepFact.GetData(s:Hstring):variant;
Var tdate,EDtDeb,EDtFin : TDateTime;
    aa,mm,jj , Edaa,Edmm,Edjj,Efaa,Efmm,Efjj : word;
    mois, Edmois,Efmois : string;
    aflib,codeaf,codeaf1,codeaf2,codeaf3,codeavnt : string;
    tobres : tob;
    zpre,znom : string;
    nomtab : string; // 2911
    toplib  :boolean;
    nlib : integer;
Begin

  //Result := '';
  s:=AnsiUppercase(s);
  nlib:=0; toplib:=false;     //2911

  if copy(s,1,1) = '$' then
  begin
    toplib := true;
    s := copy(s,2,strlen(Pchar(s)));
  end;

  tdate :=  StrToDate(Rech.DateEch);
  DecodeDate(tdate,aa,mm,jj);
  mois := FirstMajuscule(LongMonthNames[mm]);
  aflib := TOBAffaire.GetValue('AFF_LIBELLE');
  codeaf := TOBAffaire.GetValue('AFF_AFFAIRE');
  codeaf1 := TOBAffaire.GetValue('AFF_AFFAIRE1');
  codeaf2 := TOBAffaire.GetValue('AFF_AFFAIRE2');
  codeaf3 := TOBAffaire.GetValue('AFF_AFFAIRE3');
  codeavnt := TOBAffaire.GetValue('AFF_AVENANT');
  // employe
  EDtDeb :=  StrToDate(xdeb);
  DecodeDate(EDtDeb,Edaa,Edmm,Edjj);
  Edmois := FirstMajuscule(LongMonthNames[Edmm]);
  EDtFin :=  StrToDate(xfin);
  DecodeDate(EDtFin,Efaa,Efmm,Efjj);
  Efmois := FirstMajuscule(LongMonthNames[Efmm]);

  zpre:='';znom:='';
  TobRes := RemplirTOBRessource(xemp);
  if  TobRes <> NIL then
  Begin
    zpre := TobREs.GetValue('ARS_LIBELLE2');
    znom := TobREs.GetValue('ARS_LIBELLE');
  End;



  //pop up  Acompte +ref interne
  if (s= 'NUMACPTE') then
  Begin
     // mcd 15/05/03 result := Rech.numech ;
    result := Rech.numechbis ;
    if (result > 5000) then result := result - 5000; // cas des echeances d'appréciation *
  End;
  if (s= 'MOISMMACPTE') then result := mm;
  if (s= 'MOISTEXTACPTE' ) then result := mois;

  //pop up Commentaire début Facture
  if (s= 'FACDATEMM') then result := mm;
  if (s= 'FACDATEMOIS' ) then result := mois;
  if (s= 'FACDATE' ) then result := Rech.DateEch;
  if (s= 'FACAFF' ) then result := codeaf;
  if (s= 'FACAFF1' ) then begin result := codeaf1; nlib:=1;end;   //2911
  if (s= 'FACAFF2' ) then begin result := codeaf2; nlib:=2;end;   //2911
  if (s= 'FACAFF3' ) then result := codeaf3;

  //pop up Commentaire début affaire  et sous-affaire
  if (s= 'COMLIB') or (s= 'COM2LIB') or (s ='SAFFLIB') then result := aflib;
  if (s= 'COMAFF') or (s= 'COM2AFF') or (s ='SAFFCODE') then result := codeaf;

  if (s= 'COMAFF1') or (s= 'COM2AFF1') or (s ='SAFFCODE1') then
  begin result := codeaf1; nlib:=1; end;   //2911
  if (s= 'COMAFF2') or (s= 'COM2AFF2') or (s ='SAFFCODE2') then
  begin result := codeaf2; nlib:=2; end;   //2911
  if (s= 'COMAFF3') or (s= 'COM2AFF3') or (s ='SAFFCODE3') then
  begin result := codeaf3; nlib:=3; end;   //2911


  if (s= 'COMAVNT') or (s= 'COM2AVNT') then result := codeavnt;
  if (s= 'COMDFAC') or (s= 'COM2DFAC') then result := DateToStr(xDebFac);
  if (s= 'COMFFAC') or (s= 'COM2FFAC') then result := DateToStr(xFinFac);

  //pop up  employé
  if (s= 'PEMPCODE') or (s= 'PREEMPCODE')then  result := xemp ;
  if (s= 'PEMPPRE')  or (s= 'PREEMPPRE') then  result := zpre ;
  if (s= 'PEMPNOM')  or (s= 'PREEMPNOM')  then  result := znom ;

  //pop up  sur client
  if (s= 'SCLICODE') 	then  result := xclicode ;
  if (s= 'SCLILIB')   then  result := xclilib ;
  if (s= 'SCLIVILLE') then  result := xcliville ;

  if (StrToDate(xdeb) <> idate1900) and ( StrToDate(xdeb) <> idate2099) then
  Begin
    if (s= 'PEMPDTDEBJMA') or (s='PREDTDEBJMA') then result := xdeb;
    if (s= 'PEMPDTDEBMA')  or (s='PREDTDEBMA')then result := copy(Edmois,1,3) +' '+copy(xdeb,7,4);
    if (s= 'PEMPDTDEBMOIS')or (s= 'PREDTDEBMOIS') then result := Edmois;
  End;

  if (StrToDate(xfin) <> idate1900) and ( StrToDate(xfin) <> idate2099) then
  Begin
    if (s= 'PEMPDTFINJMA') or (s= 'PREDTFINJMA')then result := xfin;
    if (s= 'PEMPDTFINMA')	 or (s= 'PREDTFINMA')  then result  := copy(Efmois,1,3) +' '+copy(xfin,7,4);
    if (s= 'PEMPDTFINMOIS')or (s= 'PREDTFINMOIS')then result := Efmois;
  End;

  if (s= 'PRELIB') then result := xlib;

  // Recherche libellé (utile si liste dans code affaire , pour avoir le libellé)
  if (toplib) and (nlib<>0) then
  Begin
    nomtab := 'AFFAIREPART'+inttostr(nlib);
    result := RechDom(nomtab,result,False);  //2911
  End;
End;

// *****************************************************************
//  Chargement des lignes d'activité,

//  Le déversement dans la tobligne se fera à la fin pour la gestion

//  des seuils minis

// *****************************************************************
procedure TAFPrepFact.ChargeActivite(CAct,CodeAff : string;TobActWork : Tob);
Var Req,whereplus,wheredate : string;
    QQ : Tquery;


BEGIN
// Lecture Activité
Req := '';
Whereplus := '';
Wheredate := '';
if (CAct = 'ART') then   Whereplus := 'and ACT_TYPEARTICLE = "MAR"'
else
if (CAct = 'PRE') then   Whereplus := 'and ACT_TYPEARTICLE = "PRE"'
else
if (CAct = 'FRA') then   Whereplus := 'and ACT_TYPEARTICLE = "FRA"'
else
if (CAct = 'ARF') then   Whereplus := 'and (ACT_TYPEARTICLE = "MAR" or ACT_TYPEARTICLE = "FRA")'
else
if (CAct = 'ARP') then   Whereplus := 'and (ACT_TYPEARTICLE = "MAR" or ACT_TYPEARTICLE = "PRE")'
else
if (CAct = 'FRP') then   Whereplus := 'and (ACT_TYPEARTICLE = "PRE" or ACT_TYPEARTICLE = "FRA")'
else
if (Cact <> 'TOU') then exit;

wheredate := ' AND ACT_DATEACTIVITE <= "'+ usdatetime(strtodate(Genfac.zdateact_f))+'"';
wheredate := wheredate + ' AND ACT_DATEACTIVITE >= "'+ usdatetime(strtodate(Genfac.zdateact_d))+'"';


//   Ligne 1 et 2 = CLE (Obligatoire pour que l'update fonctionne
Req := 'SELECT  ACT_TYPEACTIVITE,ACT_AFFAIRE,ACT_AFFAIRE1,ACT_AFFAIRE2,ACT_AFFAIRE3,ACT_AVENANT'
+',ACT_RESSOURCE, ACT_DATEACTIVITE'
+ ',ACT_FOLIO, ACT_TYPEARTICLE,ACT_NUMLIGNE'
+ ',ACT_NUMLIGNEUNIQUE,ACT_NUMORDRE'
+ ',ACT_ARTICLE,ACT_CODEARTICLE'
+ ',ACT_LIBELLE,ACT_QTEFAC,ACT_PUVENTE,ACT_TOTVENTE,ACT_NUMPIECE,ACT_ACTIVITEREPRIS'
+ ',ACT_MNTREMISE'
+ ' FROM ACTIVITE WHERE ACT_AFFAIRE = "'+codeAff+'" and ACT_ACTIVITEREPRIS="F" '
+ ' and ACT_TYPEACTIVITE ="REA" and ACT_ETATVISAFAC="VIS" and ACT_ETATVISA="VIS" '
+Whereplus+WHeredate+ ' ORDER BY ACT_DATEACTIVITE';   // mcd 04/04/03 ajout order by

QQ:=OpenSQL(Req,True,-1,'',true) ;
If Not QQ.EOF then
  begin
  TobActWork.LoadDetailDB('ACTIVITE','','',QQ,True) ;
  end;
Ferme(QQ) ;
END;
// *****************************************************************
// DeChargement des lignes d'activité,  avec prise en compte des seuils

// *****************************************************************

Procedure TAFPrepFact.DechargeActivite;
Var TobDetAct: TOB;
    typart : string;
    wi : integer;
    topfour,topfrais,maxfour,maxfrais : boolean;
    zmnt,cum_four,cum_frais : double;
BEGIN
if (TobAct.Detail.Count <= 0) Then  exit;
cum_four := 0;cum_frais := 0;
// 1er passage pour Cumul des montants pour gérer les seuils Minimum
for wi:=0 To TobAct.Detail.Count-1 do
Begin
  TobDetAct := TobAct.Detail[wi];
  typart := TobDetAct.GetValue('ACT_TYPEARTICLE');
  zmnt := TobDetAct.GetValue('ACT_TOTVENTE');
  if (typart = 'MAR') then  cum_four := cum_four + zmnt;
  if (typart = 'FRA') then  cum_frais := cum_frais + zmnt;
End;


topfour := true; topfrais := true;
if (cum_four < Genfac.zminifour) then topfour := false;
if (cum_frais < Genfac.zminifrais) then topfrais := false;

cum_four := 0; cum_frais := 0;
for wi:=0 To TobAct.Detail.Count-1 do
Begin
  maxfour := true; maxfrais := true;
  TobDetAct := TobAct.Detail[wi];
  typart := TobDetAct.GetValue('ACT_TYPEARTICLE');
  zmnt := TobDetAct.GetValue('ACT_TOTVENTE');

  if ((typart = 'MAR') and not(topfour)) then continue;
  if ((typart = 'FRA') and not(topfrais)) then continue;

  // la ligne sera a FAC, si on a repris la ligne d'activité rattachée
  // à la ligne d'affaire
  if (TobDetAct.GetValue('ACT_ACTIVITEREPRIS') = 'FAC') then continue;

  if (typart = 'MAR') then
    Begin
      cum_four := cum_four + zmnt;
      if ((genfac.zmaxifour <> 0) and  (cum_four > genfac.zmaxifour)) then
      Begin
        maxfour := false;
        cum_four := cum_four - zmnt;
      End;
    End;
  if not(maxfour) then continue;


  if (typart = 'FRA')  then
    Begin
      cum_frais := cum_frais + zmnt;
      if ((genfac.zmaxifrais <> 0) and (cum_frais > genfac.zmaxifrais)) then
      Begin
        maxfrais := false;
        cum_frais := cum_frais - zmnt;
      End;
    End;
  if not(maxfrais) then continue ;

  TobDetAct.PutValue('ACT_ACTIVITEREPRIS','FAC');

  CreationLigDepuisActivite(TobDetAct,Nil)

End;  // for

END;

procedure TAFPrepFact.CreationLigDepuisActivite(TobAct,TobDetLig : TOB);
Var zqte,zpuv : double;
    TobLignes,TobQte : TOB;
    rlib,zaff,zdat : string;
begin
  TobLignes := NewTOBLigne( TOBPiece,0);

  TobLignes.PutValue('GL_ARTICLE',TobAct.GetValue('ACT_ARTICLE'));
  TobLignes.PutValue('GL_CODEARTICLE',TobAct.GetValue('ACT_CODEARTICLE'));
  TobLignes.PutValue('GL_TYPEARTICLE',TobAct.GetValue('ACT_TYPEARTICLE'));
  TobLignes.PutValue('GL_REFARTSAISIE',TobAct.GetValue('ACT_CODEARTICLE'));
  TobLignes.PutValue('GL_RESSOURCE',TobAct.GetValue('ACT_RESSOURCE'));
  TobLignes.PutValue('GL_VALEURREMDEV',TobAct.GetValue('ACT_MNTREMISE'));

  rlib := TobAct.GetValue('ACT_LIBELLE');
  zqte := TobAct.GetValue('ACT_QTEFAC');
  zpuv := TobAct.GetValue('ACT_PUVENTE');

  if GetParamSoc('SO_AFVARIABLES') then
  begin
    TobQte := RechVariable(TobVarQte_ACT,Nil,TobAct,'ACT');
    if (TobQte <> Nil) and (TobQte.getvalue('AVV_QTECALCUL')<>0)  then
      zqte := TobQte.getvalue('AVV_QTECALCUL');
  end;


  CreationLignePiece(Toblignes,zqte,zpuv,rlib,'T');

  TobLignes.PutValue('GL_AFFAIRE',TobAct.GetValue('ACT_AFFAIRE'));

  TobLignes.PutValue('GL_AFFAIRE1',TobAct.GetValue('ACT_AFFAIRE1'));
  TobLignes.PutValue('GL_AFFAIRE2',TobAct.GetValue('ACT_AFFAIRE2'));
  TobLignes.PutValue('GL_AFFAIRE3',TobAct.GetValue('ACT_AFFAIRE3'));
  TobLignes.PutValue('GL_AVENANT',TobAct.GetValue('ACT_AVENANT'));

  zaff := TobAct.GetValue('ACT_AFFAIRE');
  zdat := TobAct.GetValue('ACT_DATEACTIVITE');
  GestionChampSuppLigne(Toblignes,zaff,zdat,TobAffaire.GetValue('AFF_TIERS'),AffRef);
  //GMONYX ???  champsupp  :on l'a déjà dans Creation ligne piece

  if (TobDetLig<>Nil) then
  begin
      TOBCopyFieldValues(TobDetLig,TobLignes,['GL_FORMULEVAR','GL_FACTURABLE',
      'GL_GENERAUTO','GL_FORCODE1','GL_FORCODE2','GL_NUMORDRE','GL_NUMLIGNE','GL_FAMILLETAXE1']);
  end;

  UG_LoadAnaLigne ( tobpiece_o,TOBLignes ) ;

  if GetParamSoc('SO_AFVARIABLES') then
    AlimTobVarQte(TobVarQte_LIG, Toblignes, TobAct, TobVarQte_ACT, 'ACT', iCpt);
end;

procedure TAFPrepFact.CreationLigDepuisActiviteLiee(TobAct,TobLigLiee,TobVarQte_AFF : TOB);
Var zqte,zpuv : double;
    TobLignes : TOB;
    rlib,zaff,zdat : string;
begin
  zqte := 0;
  TobLignes := NewTOBLigne( TOBPiece,0);
  TOBCopyFieldValues(TobLigLiee,TobLignes,['GL_ARTICLE','GL_CODEARTICLE',
  'GL_TYPELIGNE','GL_TYPEARTICLE','GL_REFARTSAISIE','GL_REMISELIGNE',
  'GL_NONIMPRIMABLE','GL_AFFAIRE','GL_AFFAIRE1','GL_AFFAIRE2','GL_AFFAIRE3','GL_AVENANT',
  'GL_PRIXPOURQTE','GL_FORMULEVAR','GL_GENERAUTO','GL_FACTURABLE','GL_FORCODE1','GL_FORCODE2'
  ,'GL_NUMORDRE','GL_NUMLIGNE','GL_FAMILLETAXE1']);

  UG_LoadAnaLigne ( tobpiece_o,TOBLignes ) ;

  TobLignes.PutValue('GL_RESSOURCE',TobAct.GetValue('ACT_RESSOURCE'));

  rlib := TobLigLiee.GetValue('GL_LIBELLE');
 // zqte := TobAct.GetValue('ACT_QTEFAC');
  zpuv := TobAct.GetValue('ACT_PUVENTE');

  CreationLignePiece(Toblignes,zqte,zpuv,rlib,'T');
      // reprise des infos personnalisées dans la ligne
  TOBCopyFieldValues(TobLigLiee,TobLignes,['GL_LIBREART1','GL_LIBREART2','GL_LIBREART3','GL_LIBREART4','GL_LIBREART5',
  'GL_LIBREART6','GL_LIBREART7','GL_LIBREART8','GL_LIBREART9','GL_LIBREARTA',
  'GL_FAMILLENIV1','GL_FAMILLENIV2','GL_FAMILLENIV3','GL_DPR','GL_LIBCOMPL','GL_FAMILLETAXE1']);

  zaff := TobAct.GetValue('ACT_AFFAIRE');
  zdat := TobAct.GetValue('ACT_DATEACTIVITE');
  GestionChampSuppLigne(Toblignes,zaff,zdat,TobAffaire.GetValue('AFF_TIERS'),AffRef);

  TrtLigneLiee(TobLignes,TobAct,TobVarQte_AFF);


end;

// ************************************************************
//  Phase 3 : Transformation pour la présentation

// ************************************************************

procedure TAFPrepFact.Transformation;
var TobLigne : TOB;
  wa,nb_aff : Integer;
  we,nb_emp : Integer;
  ListeAff : array[0..99] of string;
  ListeCli : array[0..99] of string;
  ListeEmp,ListeDDeb,ListeDFin ,ListeArt: array[0..99] of string;

  critere: string;
  wi,Indice:integer;
  deremp,ruptcli : boolean;
BEGIN
	for Indice := 0 to TOBPiece.detail.count -1 do
  begin
		PieceVersLigne ( TOBPiece,TOBPiece.detail[Indice],false,false) ;
  end;

  MajLignesRess(TobPiece,true);

  TobPiece_OK.dupliquer(TobPiece,False,True,True);// dupplic slt  de l'entete


  RenumNumligne(TobPiece);    // pour bien trier ensuite
  TriTobLigne(critere);


  TobPiece.detail.sort(critere);


  // ajout champ supp si cumul sur famille prest
  TobLigne := TobPiece.detail[0];
  Tobligne.AddChampSup('ZZ_FAMARTICLE',true);

  // Si Affaire Cumulé 1 seul appel, sinon autant d'appel que d'affaires
  nb_aff := 1;
  fillchar(ListeAff,sizeof(ListeAff),#0);
  fillchar(ListeCli,sizeof(ListeCli),#0);
  if not(Genfac.zcumaff) or (Genfac.zcumsaff) or (Genfac.zcomsaff) then
    nb_aff := RechNombreAff(ListeAff,ListeCli);

  for wa:=0 to nb_aff-1 do
  Begin
    Caff := ListeAff[wa];
    //comment ss aff ??
    nb_emp :=1;
    fillchar(ListeEmp,sizeof(ListeEmp),#0);
    fillchar(ListeArt,sizeof(ListeArt),#0);
    for wi:=0 to 99 do
    Begin
   		ListeDdeb[wi] := '01/01/1900';
    	ListeDfin[wi] := '31/12/2099';
    End;

    if (Genfac.zcumemp) or (Genfac.zcomemp) then
    	nb_emp := RechNombreEmp(caff,ListeEmp,ListeDDeb,ListeDFin,ListeArt)
    else
    	RechNombreArt( caff,ListeArt);

    for we:= 0 to nb_emp-1 do
    Begin
   	cemp := ListeEmp[we];
    ddeb := ListeDDeb[we];
    dfin := ListeDFin[we];
    ctopart := ListeArt[we];
    // applic profil pour caff, cemp,ddeb,dfin
    if (we = nb_emp-1) then deremp := true else deremp := false;
    RuptCLi:=false;
//if VH_GC.GCIfDefCEGID then    why ??
    if   (wa = nb_aff-1) or (ListeCli[wa]<>ListeCli[wa+1]) then
    begin
      RuptCli := true;
    end;

    ApplicProfil(deremp,RuptCli);
    End  //fin for we (employe)
    //ici
  End; // fin for wa (affaire)

  MajTobLignes();
END;
// **********************************************************
//  Phase 3 : Application du profil de présentation

// Tri de la tob ligne en fonction des ruptures demandées

//***********************************************************

Procedure TAFPrepFact.TriTobLigne(var critere : string);
BEGIN
	critere := '';     

{$IFNDEF BTP}
  if not(Genfac.zcumaff) or (Genfac.zcomsaff) or (Genfac.zcumsaff) then
  	critere := critere + 'GLLAFFREF;GL_AFFAIRE';

  If (Genfac.zcumemp)or (Genfac.zcomemp) then
  	critere := critere + ';GL_RESSOURCE;GLL_DATEACT';

  If (Genfac.zemppre) then
  	critere := critere + ';GL_ARTICLE;GLL_DATEACT';

  if (Genfac.zstotcli) then
  	critere := ';GLLTIERSORI;GL_AFFAIRE';
{$ENDIF}
	if (critere = '') then
   	critere := ';GL_AFFAIRE';
//   	critere := ';GLLAFFREF;GL_AFFAIRE';

  critere := critere + ';GL_NUMLIGNE';

 	if (critere[1] = ';') then
  Begin
  	trim(critere);
    critere := copy(critere,2,length(critere)-1);
   ENd;


END;
// **********************************************************
//  Phase 3 : Application du profil de présentation

//  Recherche du nombre d'affaire à traiter

//***********************************************************

Function TAFPrepFact.RechNombreAff (var listeaff ,listecli: array of string) : Integer;
var wj : integer;
    sv_aff: string;
    toblig : TOB;
BEGIN
result := 0; // nombre d'affaire cumulées sur la facture
sv_aff := '';

for wj:=0 to  TOBPiece.Detail.Count-1  do
Begin
  Toblig := TobPiece.detail[wj];
  if ((sv_aff = '')  or
  ((sv_aff <> '') AND (sv_aff <> TobLig.getValue('GL_AFFAIRE'))))then
  Begin
    ListeAff[result] := TobLig.getValue('GL_AFFAIRE');
    ListeCli[result] := TobLig.getValue('GLLTIERSORI');
    inc(result);
  End;

  sv_aff := TobLig.getValue('GL_AFFAIRE');
End;

END;
// **********************************************************
//  Phase 3 : Application du profil de présentation

//  Recherche du nombre d'employé à traiter

//***********************************************************

Function TAFPrepFact.RechNombreEmp (caff : string;var listeemp,listeDdeb,listeDFin,listeArt : array of string) : Integer;
var wj : integer;
    sv_emp,zdate: string;
    toblig : TOB;
BEGIN
result := 0; // nombre d'employé
sv_emp := '';
zdate := '';

for wj:=0 to  TOBPiece.Detail.Count-1  do
Begin
  Toblig := TobPiece.detail[wj];
  if (caff <> '') AND (caff <> TobLig.getValue('GL_AFFAIRE')) then
  	break;

  if ( TobLig.getValue('GL_RESSOURCE')  = DER_RESS) then continue; //3004

  if ((wj =0 )  or
  		(sv_emp <> TobLig.getValue('GL_RESSOURCE')))  then
  Begin
    ListeEmp[result] := TobLig.getValue('GL_RESSOURCE');
    inc(result);
  End;
  // pour savoir si on a des lignes article pour l'employé
  // Si on n'en a pas, ca veut dire que on est sur un commentaire et on
  // n'aura pas besoin d'editer les ECO et EST (comment et stot employé)
  if Toblig.getvalue('GL_TYPELIGNE') = 'ART' then
  		ListeArt[result-1] := 'X';

  // trt des dates debut et fin d'intervention
  zdate := TobLig.getValue('GLL_DATEACT');
  if (ListeDDEB[result-1] = '') then   ListeDDeb[result-1] := zdate;
  if (ListeDFin[result-1] = '') then   ListeDFin[result-1] := zdate;
  if (ListeDDeb[result-1] = '01/01/1900') or (zdate < ListeDDeb[result-1]) then
  		ListeDDeb[result-1] := zdate;
  if (ListeDFin[result-1] = '31/12/2099') or (zdate > ListeDFin[result-1]) then
  		ListeDFin[result-1] := zdate;

  sv_emp := TobLig.getValue('GL_RESSOURCE');

End;  // for


END;
// **********************************************************
//  Phase 3 : Application du profil de présentation

// pour savoir si on a des lignes article pour la sous-affaire

// Si on n'en a pas, ca veut dire que on est sur un commentaire et on
// n'aura pas besoin d'editer les ACO et AST (comment et stot ss affaire)

//***********************************************************

Procedure TAFPrepFact.RechNombreArt (caff : string;var listeArt : array of string);
var wj : integer;
    toblig : TOB;
BEGIN

for wj:=0 to  TOBPiece.Detail.Count-1  do
Begin
  Toblig := TobPiece.detail[wj];
  if (caff <> '') AND (caff <> TobLig.getValue('GL_AFFAIRE')) then
  	break;

  if Toblig.getvalue('GL_TYPELIGNE') = 'ART' then
  		ListeArt[0] := 'X';

End;  // for

END;
// ************************************************************
//  Phase 3 : Application du profil de présentation,

//  Lecture du profil pour connaitre l'ORDRE de présentation

// ************************************************************

procedure TAFPrepFact.ApplicProfil(deremp , RuptCli: boolean);
Var wj : Integer;
    pelt,pelt1,pdetail : string;
    TobDet : Tob;
BEGIN
  // independemment du profil reprise des lignes de commentaires début facture;
  // ils sont stocké en typeligne = 'COM'

  TrtCommentDebFac;


for wj:=0 to TOBProfil.Detail.Count-1 do
Begin
  Tobdet := TobProfil.detail[wj];
  If  (TobDet.GetValue('APG_PROFILGENER') = Genfac.zprof) then
  Begin
  	TOBPiece_CUM:=TOB.Create('PIECE',Nil,-1) ;

    pelt 		:= 	Tobdet.getValue('APG_ELEMENT');
    pdetail :=  Tobdet.getValue('APG_DETAIL');
    pelt1 	:=  copy(pelt,1,1);

    if (pelt <> 'A00')  Then
    Begin
      BoucleLignes(pelt,deremp);
      TrtTobCumul(pelt,pdetail,RuptCLi);
    End;

    TOBPiece_CUM.free; TOBPiece_CUM:=Nil ;
  End;
End;
// A la fin, on regarde s'il reste des lignes non traités
// exemple : cas où le profil n'est pas correct
	pelt := 'FIN';
	if (Genfac.zcumemp) then    pelt := 'EST1';
	if (Genfac.zcumsaff) then  pelt := 'AST1';

  TOBPiece_CUM:=TOB.Create('PIECE',Nil,-1) ;
  BoucleLignes(pelt,deremp);
  TrtTobCumul(pelt,' ',false);
  TOBPiece_CUM.free; TOBPiece_CUM:=Nil ;
// Si on a un article de %
// on le met systématiquement à la FIN.
	if ((Genfac.zpour) and (pelt <> 'EST1')) or
       ((Genfac.zpour) and (pelt = 'EST1') and (deremp))  then
    Begin
    pelt := 'FPO';  // fin %
    TOBPiece_CUM:=TOB.Create('PIECE',Nil,-1) ;
    BoucleLignes(pelt,deremp);
    TrtTobCumul(pelt,' ',false);
    TOBPiece_CUM.free; TOBPiece_CUM:=Nil ;
    End;

END;
// ************************************************************
//  Phase 3 : Application du profil de présentation,

//  Lecture des lignes et  transfert dans TobCUM

// ************************************************************

procedure TAFPrepFact.BoucleLignes (pelt : string;deremp : boolean);
Var wj : Integer;
    pelt1 : string;
    TobLig : Tob;
    ChgTob : boolean;  // Gestion de l'indice de la tobPiece en fct du chg de parent
BEGIN
wj:=0;
pelt1 := copy(pelt,1,1);
while (wj < TOBPiece.Detail.Count) do
Begin

  chgtob:=false;
  Toblig := TobPiece.detail[wj];

  if (pelt = 'XPO') or (pelt = 'XCL') then exit;
  //GMPOU , l'article de % sera toujours en fin de profil donc on le saute qd
  // on le trouve dans le Profil
  // idem pour XCL , sous total sur Client à facturer

  if ((Genfac.zcumaff = false) AND
     (TobLig.getValue('GL_AFFAIRE') <> Caff) )then exit;

  if ((Genfac.zcumsaff) or (Genfac.zcomsaff)) AND
     (TobLig.getValue('GL_AFFAIRE') <> Caff )then exit;

 if ((Genfac.zcumemp) or (Genfac.zcomemp)) AND
     (TobLig.getValue('GL_RESSOURCE') <> Cemp) and
     	not(deremp) then exit;     // il fait qu'on puisse tariter 'FPO' apres le dernier emp
     // attention car on n'a pas l"employé sur chaque ligne (ex : commentaire )


// trt special FIN , gestion des restes
  if ((pelt = 'FIN') or (pelt = 'FPO')) then
  Begin
    ChgTob := RecupLigFinOuPou(Toblig,pelt);
  End;

// trt TOU , reprises de toutes les lignes de l'affaire
  if (pelt = 'TOU') then
  Begin
    ChgTob := RecupLig(Toblig);
  End;

// trt commentaire
  if (pelt1 = 'C') or (pelt = 'ACO') then
  Begin
    Chgtob := TrtComment(Toblig,pelt);
  End;

// trt commentaire  employe
  if (pelt = 'ECO') then
  Begin
  	if (ctopart = 'X') then
    Begin
	    CreerComEmp(Toblig,Genfac.zlibempcom);
	   	exit;
    End;
  End;

// trt montant forfaitaire : acompte + prest/frais/four si non détaillés
  if (pelt = 'G00') then
  Begin
    ChgTob := CumLigForfait(Toblig);
  End;

//  Detail des lignes Prest/Frais/Fourn
  if ((pelt = 'P00') or  (pelt = 'S00') or (pelt = 'V00') or (pelt = 'N00')) then
  Begin
     pelt1 := copy(pelt,1,1);
     ChgTob := DetailLig(pelt1,Toblig);
  End;

// trt cumul sur type/famille ou code
   if ((pelt = 'P10') or  (pelt = 'S10') or (pelt = 'V10') or (pelt = 'N10') or
       (pelt = 'P20') or  (pelt = 'S20') or (pelt = 'V20') or (pelt = 'N20') or
       (pelt = 'P30') or  (pelt = 'S30') or (pelt = 'V30') or (pelt = 'N30') or
       (pelt = 'EPR') or  (pelt = 'EST1') or (pelt = 'AST1'))  then
  Begin
    ChgTob := CumLig(pelt,Toblig);
  End;

  if not(ChgTob) then inc(wj);  // pas de changement de parent  inc

End;
END;
// ************************************************************
//  Phase 3 : Application du profil de présentation,

//  Recup des lignes de commentaire  début de facture (issu de la saisie des

//  ligne de commentaire dans le mul )

// ************************************************************

procedure TAFPrepFact.TrtCommentDebfac;
Var wj : Integer;
    TobLig : Tob;
    ChgTob : boolean;  // Gestion de l'indice de la tobPiece en fct du chg de parent
BEGIN
  TOBPiece_CUM:=TOB.Create('PIECE',Nil,-1) ;

  wj:=0;
  while (wj < TOBPiece.Detail.Count) do
  Begin
    chgtob:=false;
    Toblig := TobPiece.detail[wj];
    if (Toblig.GetValue('GL_TYPELIGNE') = 'COM') then
       ChgTob := RecupLig(Toblig)
    else
      wj := 999999;

    if not(ChgTob) then inc(wj);  // pas de changement de parent  inc

  end;

  if (TOBPiece_CUM <> Nil) then
  Begin
    TrtTobCumul('COM','',false);
    TOBPiece_CUM.free; TOBPiece_CUM:=Nil ;
  End;
END;
// ************************************************************************
//  Phase 3 : Application du profil de présentation,

//  Reprise de TOUTES les lignes affaires

//  transfert dans TobCUM

// ************************************************************************

function TAFPrepFact.RecupLig(Toblig : tob) :boolean;
Begin
    Toblig.ChangeParent(TobPiece_CUM,-1); result := true;
End;
// **********************************************************************
//  Phase 3 : Application du profil de présentation,

//  trt des commentaires

// **********************************************************************
function TAFPrepFact.TrtComment(Toblig : tob;pelt : string) : boolean;
Var  pelt1 : string;
BEGIN
  result := false;
  pelt1:=copy(pelt,1,1);
  if (toblig.getvalue('GL_TYPELIGNE')= pelt) then
  Begin
    result := true;
    Toblig.ChangeParent(TobPiece_CUM,-1);
    toblig.PutValue('GL_TYPELIGNE','COM');
  End;
END;
// **********************************************************************
//  Phase 3 : Application du profil de présentation,

//  trt des commentaires debut employe

// **********************************************************************
procedure TAFPrepFact.CreerComEmp(Toblig :tob;Comment : string);
var Tobcum :tob;
		zlib,s : string;
BEGIN
//  TobCum:=TOB.Create('LIGNE',TobPiece_OK,-1);
  TobCum := NewTOBLigne( TobPiece_OK,0);
  TobCum.dupliquer(TobLig,false,True,True);
  NewTOBLigneFille(TobCum); // pour bien récreer les filles analytiques  à VIDE

  InitTobLigneApresDupplic(Tobcum);

  Tobcum.PutValue('GL_TYPELIGNE','COM');


  xdeb := ddeb; xfin := dfin; xemp := cemp;
  s := '{'+comment+'}';
  zlib := GFormule(s,GetData,nil,1);
  zlib := trim(copy(zlib,1,70)); // gm 19/11/02
  Tobcum.PutValue('GL_LIBELLE',zlib);
END;
// **********************************************************************
//  Phase 3 : Application du profil de présentation,

//  trt des commentaires ligne détail prestation

// **********************************************************************
procedure TAFPrepFact.TrtLibPres(Tobdet : tob;var zlib : string);
var s : string;
BEGIN

	xdeb := TobDet.GetValue('GLL_DATEACT');
  xfin := TobDet.GetValue('GLL_DATEACT');
  xemp := TobDet.GetValue('GL_RESSOURCE');
  xlib := TobDet.GetValue('GL_LIBELLE');

  if (Genfac.zlibpres <> '') then
  Begin
    s := '{'+Genfac.zlibpres+'}';
    zlib := GFormule(s,GetData,nil,1);
  End
  else
  	zlib :=  TobDet.GetValue('GL_LIBELLE');
  zlib := trim(copy(zlib,1,70)); // gm 19/11/02
END;
// **********************************************************************
//  Phase 3 : Application du profil de présentation,

//  trt montant forfaitaire : acompte + prest/frais/four si non détaillés

// **********************************************************************

function TAFPrepFact.CumLigForfait(Toblig : tob) : boolean;
Var chg : boolean;
  st : string;
BEGIN
  result:= false;
  chg := false;
  st := toblig.GetValue('GL_ARTICLE');
  st := toblig.GetValue('GL_TYPEARTICLE');
  if (Toblig.GetValue('GL_ARTICLE') = Genfac.zacompte) then  chg := true;

  if (chg) then
  Begin
    Toblig.ChangeParent(TobPiece_CUM,-1); result := true;
  End;

END;
// **********************************************************************
//  Phase 3 : Application du profil de présentation,

//  trt du reste  et  gestion article Pourcentage

// **********************************************************************

function TAFPrepFact.RecupLigFinOuPou(Toblig : tob;pelt : String) : boolean;
Var chg : boolean;
  st,typ : string;
BEGIN
  result:= false;
  chg := false;
  st := toblig.GetValue('GL_ARTICLE');
  typ := toblig.GetValue('GL_TYPEARTICLE');

  if (pelt = 'FIN') then
  Begin
  	if ((typ='PRE') or (typ='FRA') or (typ='MAR'))then chg := true;
  End;

  if (pelt = 'FPO') then
  Begin
  	if  (typ='POU') then chg := true;
  End;

  if (chg) then
  Begin
    Toblig.ChangeParent(TobPiece_CUM,-1); result := true;
  End;

END;
// ************************************************************
//  Phase 3 : Application du profil de présentation,

//  trt du détail de lignes

// ************************************************************

function TAFPrepFact.DetailLig(pelt1 : string;Toblig : tob) : boolean;
Var chg: boolean;
  st,codart : string;
BEGIN
  result:= false;
  chg := false;
  codart := toblig.GetValue('GL_ARTICLE');
  st := toblig.GetValue('GL_TYPEARTICLE');

  if (Genfac.zacpte) and (codart = GenFac.zacompte) then exit;
  if ((Toblig.GetValue('GL_TYPEARTICLE')='PRE') and (pelt1 = 'P')) then chg := true;
  if ((Toblig.GetValue('GL_TYPEARTICLE')='FRA') and (pelt1 = 'V')) then chg := true;
  if ((Toblig.GetValue('GL_TYPEARTICLE')='MAR') and (pelt1 = 'S')) then chg := true;
  if ((Toblig.GetValue('GL_TYPEARTICLE')='CTR') and (pelt1 = 'N')) then chg := true;

  if (chg) then
  Begin
    Toblig.ChangeParent(TobPiece_CUM,-1); result := true;
  End;

END;
// ************************************************************************
//  Phase 3 : Application du profil de présentation,

//  trt des cumuls Prest.frais;four.

//  On doit gérer une tob intermédiaire pour gérer les cumul : TobPiece_CUM

// ************************************************************************

function TAFPrepFact.CumLig(pelt : string;Toblig : tob) : boolean;
Var chg : boolean;
  pelt1,pelt2,nomcrit,valcrit,codart,TypArt : string;
  QQ : Tquery;
BEGIN
  result:= false;
  chg := false;
  pelt1 := copy(pelt,1,1);
  pelt2 := copy(pelt,2,1);
  codart := toblig.GetValue('GL_ARTICLE');

  // on prend aussi les lignes d'acomptes dans cumul employé
  if (Genfac.zacpte) and (codart = GenFac.zacompte) and
  	(pelt <> 'EST1') and (pelt <> 'AST1') and (pelt <> 'EPR') then exit;

  TypArt := Toblig.GetValue('GL_TYPEARTICLE');
  if ((TypArt ='PRE') and (pelt1 = 'P')) then chg := true;
  if ((TypArt ='FRA') and (pelt1 = 'V')) then chg := true;
  if ((TypArt='MAR') and  (pelt1 = 'S')) then chg := true;
  if ((TypArt='CTR') and  (pelt1 = 'N')) then chg := true;
  if (pelt = 'EST1') or (pelt = 'AST1') or (pelt = 'EPR') then chg := true;

  if (typart = 'POU') and (pelt<> 'AST1') then chg := false; //3004

  if not(chg) then exit;

  if (pelt2 = '1') or (pelt = 'EPR') then
  Begin
    nomcrit := 'GL_ARTICLE';
    valcrit := Toblig.GetValue('GL_ARTICLE');
  End;

  if (pelt2 = '2') then
  Begin
    nomcrit := 'GL_TYPEARTICLE';
    valcrit := Toblig.GetValue('GL_TYPEARTICLE');
  End;

  if (pelt2 = '3') then
  Begin
    nomcrit := 'ZZ_FAMARTICLE';
    valcrit := ' ';
    QQ:=OpenSQL('SELECT GA_FAMILLENIV1 FROM ARTICLE WHERE GA_ARTICLE="'+ CodArt +'"',true,-1,'',true);
    if not QQ.EOF then
        valcrit:=QQ.findField('GA_FAMILLENIV1').asString ;
    ferme(QQ);
    Toblig.PutValue('ZZ_FAMARTICLE',valcrit)
  End;


// Cas de cumul sur le Code
  if ((pelt = 'P10') or (pelt = 'V10') or (pelt = 'S10') or (pelt = 'N10') or (pelt = 'EPR')) then
  Begin
    CumLigCode(pelt,nomcrit,valcrit ,toblig);
    result := true;  // gm 16/01/03 pour dire que enreg traite (bug suite à optimisation)
  End;
// Cas de cumul sur Type ou Famille
// Cas de cumul pour sous total Affaire ou employé
  if ((pelt = 'P20') or (pelt = 'V20') or (pelt = 'S20') or  (pelt = 'N20') or
      (pelt = 'P30') or (pelt = 'V30') or (pelt = 'S30') or  (pelt = 'N30') or
      (pelt = 'EST1') or (pelt = 'AST1') ) then
  Begin
    result := CumLigFamType(nomcrit,valcrit ,toblig);
  End;


END;

// ************************************************************************
//  Phase 3 : Application du profil de présentation,

//  trt des cumuls Prest;frais;four.

// CUMUL sur CODE Prest

// ************************************************************************

procedure TAFPrepFact.CumLigCode(pelt,nomcrit,valcrit : string;Toblig : tob);
Var zmnt,zqte : double;
    TobRes : TOB;
Begin
    TobRes := TobPiece_CUM.FindFirst(nomcrit,[valcrit],false);
    if (TobRes <> NIL) then
    Begin
             // mcd 02/12/02  if (TobRes.GetValue('GL_PUHTDEV') =  Toblig.GetValue('GL_PUHTDEV')) then
      if (GetParamSoc('So_AfRecalPvFact')) or ((TobRes.GetValue('GL_PUHTDEV') =  Toblig.GetValue('GL_PUHTDEV'))) then
      Begin
      // si même pv , on garde les notion de qte et pu, sauf si recalcul pv, on conserve la qte même si prix #
        zqte := TobLig.GetValue('GL_QTEFACT') + TobRes.GetValue('GL_QTEFACT');
        TobRes.PutValue('GL_QTEFACT',zqte);
        zmnt := Arrondi(TobRes.GetValue('GL_QTEFACT') *  TobRes.GetValue('GL_PUHTDEV'),V_PGI.OkDecV);
        TobRes.PutValue('GL_TOTALHTDEV',zmnt);
      End
      else
      Begin // cas ou pas le même mtt unitaire .. on est obligé de forcer la qte à 1  et de conserver le mtt total
        zmnt := TobLig.GetValue('GL_TOTALHTDEV') + TobRes.GetValue('GL_TOTALHTDEV');
        TobRes.PutValue('GL_TOTALHT',zmnt);
        TobRes.PutValue('GL_TOTALHTDEV',zmnt);
        TobRes.PutValue('GL_QTEFACT',1);
        TobRes.PutValue('GL_PUHT',zmnt);
        TobRes.PutValue('GL_PUHTDEV',zmnt);
      End;
      if (pelt = 'EPR') then
      	Toblig.ChangeParent(TobPiece_CUM,-1)  //on garde les elts pour commnetaires emp
      else
      	Toblig.free;
    End
    else
    Begin
      Toblig.ChangeParent(TobPiece_CUM,-1);
    End
  End;
// ************************************************************************
//  Phase 3 : Application du profil de présentation,

//  trt des cumuls Prest.frais;four.

// CUMUL sur Famille ou Type Prestation : transfert dans TobCUM

// ************************************************************************

function TAFPrepFact.CumLigFamtype(nomcrit,valcrit : string;Toblig : tob) :boolean;
Begin
    Toblig.ChangeParent(TobPiece_CUM,-1); result := true;
End;
// ************************************************************************
//  Phase 3 : Application du profil de présentation,

//  trt des cumuls Prest.frais;four.

// Exploitation de la TOb Cumul

// ************************************************************************

procedure TAFPrepFact.TrtTobCumul(pelt,pdetail: string ; RuptCli : boolean);
Var wi,zindex : integer;
    TobDet,TobSauv,TobRes,TobWork : TOB;
    pelt1,pelt2,sv_crit,nomcrit,zaff : string;
    zmnt,zqte : double;
    TopTotImp,TopDetImp,acompte_exist,fin,topart : boolean;
    sv_pres,svcour,zlib : string;
Begin

  TobSauv:=Nil;
//  si toutes les lignes sont déjà mises , il faut générer le stot  sauf si on
// n'avait que des commentaires 	avant
	if  (TobPiece_CUM.Detail.count = 0) and ((pelt = 'EST1') or (pelt = 'AST1'))
   and (ctopart = 'X')  then
  Begin
  	tobsauv := tobpiece_ok.detail[0]; // ???
//    TopCum := False;   // avert vérifier valeur
    TopTotImp := False;
    TrtTobCumul_Total(pelt,'',zmnt,zqte,TopTotImp,Tobsauv);
  	exit;
  End;

  if ((TobPiece_CUM = NIL) or (TobPiece_CUM.Detail.count = 0)) then exit;

  // si FPO , génération de la ligne de Pourcentage
	if (pelt = 'FPO') then
  Begin     // mettre boucle
        TobDet := TobPiece_CUM.detail[0];
  			Tobdet.ChangeParent(TobPiece_OK,-1);
//        topcum  := false; //2905
        TopDetImp := true;
        TopTotImp := false;
        GereCumul(TobDet,TopDetImp,zmnt,zqte,pelt);
        tobsauv := TobDet; // 2905
    		TrtTobCumul_Total(pelt,'',zmnt,zqte,TopTotImp,Tobsauv); //2905
  End;

  zmnt := 0; zqte := 0;
  pelt2 := copy(pelt,2,1);
  pelt1 := copy(pelt,1,1);
  if (pelt2 = '1') then  nomcrit := 'GL_ARTICLE';
  if (pelt2 = '2') then  nomcrit := 'GL_TYPEARTICLE';
  if (pelt2 = '3') then  nomcrit := 'ZZ_FAMARTICLE';


  if ((pelt = 'P10') or (pelt = 'V10') or (pelt = 'S10') or (pelt = 'N10') or
      (pelt = 'P00') or (pelt = 'V00') or (pelt = 'S00') or (pelt = 'N00') or
      (pelt = 'G00') or  (pelt1 = 'C') or  (pelt = 'ACO')) then
  Begin
 //   topcum := false;  // pas de sous-total demandé par le profil
    TopDetImp := true;  // pas de sous-total demandé par le profil
    TopTotImp := false;
    //    if ((pelt = 'G00') and (Genfac.zcumaff) )  then  topcum := true; //GMGMGM 15/12/2000
    if ((pelt = 'G00') and (Genfac.zcumaff) )  then
    begin
      TopDetImp := false; //GMGMGM 15/12/2000
      TopTotImp := true;
    end;
    wi := 0;
    while (TobPiece_CUM.Detail.count <> 0) do
      Begin
      TobDet := TobPiece_CUM.detail[wi];
      TobSauv:= TobPiece_CUM.detail[wi];
      if (pelt = 'P00') then
      Begin
      	TrtLibPres(TobDet,zlib);
        Tobdet.putvalue('GL_LIBELLE',zlib);
      End;
      Tobdet.ChangeParent(TobPiece_OK,-1);
      GereCumul(TobDet,TopDetImp,zmnt,zqte,pelt);
     End;
     TrtTobCumul_Total(pelt,'',zmnt,zqte,TopTotImp,Tobsauv)
  End;


  if ((pelt = 'P20') or (pelt = 'V20') or (pelt = 'S20') or (pelt = 'N20') or
      (pelt = 'P30') or (pelt = 'V30') or (pelt = 'S30') or (pelt = 'N30')) then
  Begin
//    topcum := true;  // sous-total demandé par le profil
    TopDetImp := false;  // sous-total demandé par le profil
    TopTotImp := true;
    TobPiece_CUM.detail.sort(nomcrit);
    sv_crit := '';
    wi := 0;
    while TobPiece_CUM.Detail.count <> 0 do
    Begin
      TobDet := TobPiece_CUM.detail[wi];
      TobSauv:= TobPiece_CUM.detail[wi];
      if (sv_crit = '') or (sv_crit = TobDet.GetValue(nomcrit)) then
      Begin
         Tobdet.ChangeParent(TobPiece_OK,-1);
         //gm si avec édition détail , alors topcum := false  (pour editer ligne et sous_total)
         //GereCumul(TobDet,topcum,zmnt,zqte,pelt);
//         if (pdetail = 'X') then topcum := false;    // pbm CC sur non eiditon du détail
         if (pdetail = 'X') then
         begin
          TopDetImp := true;
          TopTotImp := true;
         end;
         GereCumul(TobDet,TopDetImp,zmnt,zqte,pelt);
      End
      else
      Begin
        TrtTobCumul_Total(pelt,sv_crit,zmnt,zqte,TopTotImp,Tobsauv);
        zmnt:=0; zqte:=0;
        End;
      sv_crit := TobDet.GetValue(nomcrit);
    End; // fin while

    TrtTobCumul_Total(pelt,sv_crit,zmnt,zqte,TopTotImp,Tobsauv);
  End;

  if (pelt = 'EPR') then
  Begin
    sv_pres := '';
    wi := 0;
    while TobPiece_CUM.Detail.count <> 0 do
    Begin
      TobDet := TobPiece_CUM.detail[wi];
      svcour := TobDet.GetValue('GL_ARTICLE');
      if (sv_pres <> svcour) then
      	Tobdet.ChangeParent(TobPiece_OK,-1);

      ddeb := TobDet.GetValue('GLL_DATEACT');
      dfin := TobDet.GetValue('GLL_DATEACT');
      cemp := TobDet.GetValue('GL_RESSOURCE');
      CreerComEmp(TobDet,Genfac.zlibemppre);


      if (sv_pres = svcour) then
          TobDet.free;

       sv_pres := svcour;

    End; // fin while
  End;

  if ((pelt = 'EST1') or (pelt = 'AST1')) then
  Begin
    wi := 0;
    topart := false;
    // je vide les lignes en trop
    while TobPiece_CUM.Detail.count <> 0 do
    Begin
 //     topcum := false; //3004
      TopDetImp := true; //3004
      TobDet := TobPiece_CUM.detail[wi];
      TobSauv:= TobPiece_CUM.detail[wi];

      if (TobDet.GetValue('GL_TYPELIGNE') = 'ART') then Topart := true;
      Tobdet.ChangeParent(TobPiece_OK,-1);
      GereCumul(TobDet,TopDetImp,zmnt,zqte,pelt);
    End; // fin while
//   topcum := true; //3004
    TopTotImp := true;
    if (topart) then   //Pas de sous-total si on n'a pas de ligne article
    	TrtTobCumul_Total(pelt,'',zmnt,zqte,TopTotImp,Tobsauv);
  End;

  if (pelt = 'TOU') then
  Begin
  TobSauv := TobPiece_CUM.detail[0];      //SIC
    while TobPiece_CUM.Detail.count <> 0 do
    Begin
    	wi := 0;
    	TobDet := TobPiece_CUM.detail[wi];
      Tobdet.ChangeParent(TobPiece_OK,-1);
    End; // fin while
// Si REgroupement sur CAF et rupture client et sous-total par client
		if  (RegSurCaf) and (RuptCli) and (Genfac.zstotcli) then
    Begin
//    	TopCum := true;
      TopTotImp := true;
			TrtTobCumul_Total('XCL','',zmnt,zqte,TopTotImp,Tobsauv);
    end;

  End;

  if (pelt = 'FIN') then
  Begin
    acompte_exist := false;

    // on regarde s'il y a déja une ligne "acompte"
    TobDet := TobPiece_CUM.detail[0];
    zaff := TobDet.GetValue('GL_AFFAIRE');
    if (Genfac.zcumaff) then
      TobRes := TobPiece_OK.FindFirst(['GL_ARTICLE'],[Genfac.zacompte],false)
    else
      TobRes := TobPiece_OK.FindFirst(['GL_ARTICLE','GL_AFFAIRE'],[Genfac.zacompte,zaff],false);

    zindex := -1;
    if (tobres <> NIL) then
    Begin
      acompte_exist := true;
      zindex := TobRes.GetIndex + 1;
      // On va avoir un cumul avec l'compte, donc il faut inverser les notions
      // "Imprimable" de la ligne acompte et du Sous Total associé
      Tobres.PutValue('GL_NONIMPRIMABLE','X');

      wi := TobRes.GetIndex;
      fin := false;
      while ((wi < TobPiece_OK.detail.count) and not(fin)) do
      Begin
        TobWork := TobPiece_OK.detail[wi];
        if (TobWork.GetValue('GL_ARTICLE')= GenFac.zacompte) then
         TobWork.PutValue('GL_NONIMPRIMABLE','X');
        if (TobWork.GetValue('GL_TYPELIGNE')= 'TOT') then
        Begin
          TobWork.PutValue('GL_NONIMPRIMABLE','-');
          fin := true;
        End;
        inc(wi);
      End;

    End;
    // s'il y a des restes ca veut dire qu'on aura des lignes NON IMPRIMABLES
//    topcum := true;
    TopTotImp := true;

    wi := 0;
    while (TobPiece_CUM.Detail.count <> 0) do
      Begin
      TobDet := TobPiece_CUM.detail[wi];
      TobSauv:= TobPiece_CUM.detail[wi];
      Tobdet.ChangeParent(TobPiece_OK,zindex);
      TobDet.PutValue('GL_NONIMPRIMABLE','X');
      if (acompte_exist) then inc(zindex);

     End;


     if ((GenFac.zstot) and not(acompte_exist)) then
     Begin
        TrtTobCumul_Total (pelt,'',zmnt,zqte,TopTotImp,Tobsauv);
     End
  End;

End;
// ************************************************************************
//  Phase 3 : Application du profil de présentation,

//  trt des cumuls Prest.frais;four.

//  Cumul des lignes et trt de la notion IMPRIMABLE / non IMPRIMABLE (LIGNE)

//      TopDetImp = false (cumul demandé donc ligne détail non imprimable) sinon cumul technique

//        (ligne deétail imprimable)

// ************************************************************************

Procedure TAFPrepFact.GereCumul(TobDet: TOB;TopDetImp :boolean;var zmnt,zqte : double;pelt : string);
var pelt1 : string;
BEGIN
	pelt1 := copy(pelt,1,1);
  if (pelt1 = 'C') or (pelt = 'ACO') Then exit;

//  if (topcum)  then
  if TopDetImp  then
    TobDet.PutValue('GL_NONIMPRIMABLE','-')     // ligne imprimable
  else
    TobDet.PutValue('GL_NONIMPRIMABLE','X');      // ligne non imprimable

  zmnt := zmnt + TobDet.getValue('GL_TOTALHTDEV');
  zqte := zqte + TobDet.getValue('GL_QTEFACT');

END;
// ************************************************************************
//  Phase 3 : Application du profil de présentation,

//  trt des cumuls Prest.frais;four.

//  Création de la ligne de sous-total

// ************************************************************************

procedure TAFPrepFact.TrtTobCumul_Total(pelt,sv_crit : string; zmnt,zqte:double;

TopTotImp : boolean;Tobsauv : TOB);

Var TobCum : TOB;
  pelt1,zlib,zfam,pelt2: string;
BEGIN
  pelt1 := copy(pelt,1,1);
  pelt2 := copy(pelt,2,1);

  if not(GenFac.zstot) and (pelt <>'EST1') and (pelt <>'AST1') and (pelt <>'XCL')then exit;

  if (pelt1 = 'C') or (pelt = 'ACO') then exit;

 // TobCum:=TOB.Create('LIGNE',TobPiece_OK,-1);
  TobCum := NewTOBLigne( TobPiece_OK,0);
  if (TobSauv<>NIL) then
  begin
 		 TobCum.dupliquer(TobSauv,false,True,True);
     NewTOBLigneFille(tobcum); // pour bien récreer les filles analytiques VIDE
  end;


  InitTobLigneApresDupplic(Tobcum);
  Tobcum.PutValue('GL_TYPELIGNE','TOT');

  zlib := '';
  zfam := '';
  if (pelt2 = '3') then
    zfam := sv_crit;

  AlimLibSousTot(pelt,pelt1,pelt2,zfam,zlib,TobSauv);

  Tobcum.PutValue('GL_LIBELLE',zlib);


//  trt de la notion IMPRIMABLE / non IMPRIMABLE     sur le Sous total
//    TopTotImp = true (cumul demandé donc cumul imprimable), sinon cumul technique

  if (TopTotImp) or (Genfac.zcumemp) or (Genfac.zcumsaff) then
    TobCum.PutValue('GL_NONIMPRIMABLE','-')
  else
    TobCum.PutValue('GL_NONIMPRIMABLE','X');

END;

// ************************************************************************
//  Phase 3 : Application du profil de présentation,

//  trt des cumuls Prest.frais;four.

//  Recherche du libellé de la ligne Sous Total

// ************************************************************************

procedure TAFPrepFact.AlimLibSousTot(pelt,pelt1,pelt2,zfam : string; var zlib : string;TobSauv : TOB);
Var zaff,s  : String;
    TobRes : TOB;
    QQ : Tquery;
BEGIN
zlib := '';

	if (pelt = 'FPO') then     //2904
     Zlib := TobSauv.GetValue('GL_LIBELLE');

  if (pelt = 'XCL') then
  Begin
  	xclilib := 'Client inconnu';
    xclicode :='';
    xcliville :='';
  	QQ:=OpenSQL('SELECT T_LIBELLE,T_TIERS,T_VILLE FROM TIERS WHERE T_TIERS="'+ TobSauv.GetValue('GLLTIERSORI') +'"',true,-1,'',true);
    if not QQ.EOF then
    Begin
        xclilib 	:=	QQ.findField('T_LIBELLE').asString ;
        xclicode 	:=	QQ.findField('T_TIERS').asString ;
        xcliville :=	QQ.findField('T_VILLE').asString ;
    End;
    ferme(QQ);

    s := '{'+Genfac.zlibstotcli+'}';
    zlib := GFormule(s,GetData,nil,1);
  End;

  if (pelt = 'AST1') then
  Begin
  	RemplirTOBAffairePrepFact(caff);
  	xdeb := ddeb; xfin := dfin; xemp := cemp;
    s := '{'+Genfac.zlibcumsaff+'}';
    zlib := GFormule(s,GetData,nil,1);
  End;


  if (pelt = 'EST1') then
  Begin
  	xdeb := ddeb; xfin := dfin; xemp := cemp;
    s := '{'+Genfac.zlibempcum+'}';
    zlib := GFormule(s,GetData,nil,1);
  End;


  if (pelt2 = '3')  then// cumul par famille
  Begin
    zlib := RechDom('GCFAMILLENIV1',zfam,False);
  End
  else
  Begin
  if (pelt1 = 'P') then
    Begin
       zlib := GenFac.zlibpres;
       if (zlib = '') then zlib := 'Cumul Prestation';
    End;
    if (pelt1 = 'S') then
    Begin
      zlib := GenFac.zlibfour;
      if (zlib = '') then zlib := 'Cumul Fourniture';
    End;
    if (pelt1 = 'V') then
    Begin
      zlib := GenFac.zlibfrais;
      if (zlib = '') then zlib := 'Cumul Frais';
    End;
    if (pelt1 = 'N') then
    Begin
      zlib := GenFac.zlibctr;
      if (zlib = '') then zlib := 'Cumul Contrat';
    End;

    if ((pelt1 = 'G') or (pelt = 'FIN')) then
    Begin
    // On récupéré le libellé de l'acompte
       Zaff := TobSauv.GetValue('GL_AFFAIRE');
       TobRes := TobPiece_OK.FindFirst(['GL_ARTICLE','GL_AFFAIRE'],[Genfac.zacompte,zaff],false);
       if (tobres <> NIL) then
          zlib := TobRes.GetValue('GL_LIBELLE')
       else
        Begin
        // formattage du libellé paramètré
        if (Genfac.zreplibech = 'X') then
          Begin
          zlib := Rech.libelleeche;
          End
        else
          Begin
          xdeb := '01/01/1900'; xfin := '31/12/2099'; xemp := '';
          s := '{'+Genfac.zlibacompte+'}';
          zlib := GFormule(s,GetData,nil,1);
          if (zlib ='') then  zlib := 'Cumul Forfaitaire';
          End;
        End;
    End;
  End;
  zlib := trim(copy(zlib,1,70)); // gm 19/11/02
END;

{ a suppr
// ************************************************************
//  Maj des lignes : Compteur + cle

// ************************************************************
Procedure  TAFPrepFact.CalcLigne(var TobLig : Tob);
var qte,puv : double;
BEGIN
  qte := TobLig.GetValue('GL_QTEFACT');
  puv := TobLig.GetValue('GL_PUHTDEV');

 // TobLig.PutValue('GL_PUHT',puv);

  TobLig.PutValue('GL_TOTALHTDEV',qte*puv);
//  TobLig.PutValue('GL_TOTALHTDEV',qte*puv);

  Toblig.PutValue('GL_TYPELIGNE','ART');

  // attention ces infos devraient être reprise depuis Tobligne,
  // usine à gaz dan Factutil...
  TobLig.PutValue('GL_FAMILLETAXE1','NOR');
  TobLig.PutValue('GL_FACTUREHT','X');
  TobLig.PutValue('GL_REGIMETAXE','FRA');
  TobLig.PutValue('GL_ESCOMPTABLE','X');
  TobLig.PutValue('GL_REMISABLEPIED','X');
END;
}
// ************************************************************
//  Maj des lignes : Compteur + cle

// ************************************************************

// ************************************************************
//  Maj des lignes : Compteur + cle

// ************************************************************

procedure TAFPrepFact.MajTobLignes;
BEGIN

     MajLignesRess(TobPiece_OK,false);
     ApplicTarif(TobPiece_OK);
     MajQte  (TobPiece_OK);

     GP_DEVISEChange;

END;

{procedure TAFPrepFact.RenumLesLignes(TobPiece_OK : TOB);
Var wj : Integer;
begin
  TobPiece_OK.putvalue('GP_CODEORDRE',0);
  // je réinitialise car sinon la numérotation repart du mac du numordre de la piece d'origine
  for wj:=0 to TOBPiece_OK.Detail.Count-1 do
    TOBPiece_OK.Detail[wj].putvalue('GL_NUMORDRE',0);

  NumeroteLignesGC(Nil,TobPiece_OK) ;   // jld43
end;
}

procedure TAFPrepFact.RenumNumligne(TobPiece_OK : TOB);
Var wj,IndiceOuv : Integer;
begin
  // je rénumérote numligne car le tri suivant en tient compte
  for wj:=0 to TOBPiece_OK.Detail.Count-1 do
  begin
    TOBPiece_OK.Detail[wj].putvalue('GL_NUMLIGNE',wj+1);
    if IsOuvrage (TOBPiece_OK.detail[wj]) then
    begin
      IndiceOuv := TOBPiece_OK.detail[wj].getValue('GL_INDICENOMEN');
      if (IndiceOuv > 0) and (TobOuvrages  <> nil) then
      begin
        ReAffecteLigOuv(IndiceOuv, TOBPiece_OK.detail[wj], TobOuvrages);
      end;
    end;
  end;

end;

procedure TAFPrepFact.MajQte(TobPiece: TOB);
Var wj : Integer;
    TobDet : TOB;
begin
// Modif pour synchroniser les qtés car en 5.0 il est fait ceci dans factgrp
// avec certains profils (avec cumul sur code article)  on était déphasé
//   NewL.PutValue('GL_QTEFACT', Tobl.GetValue('GL_QTERESTE'));
//   NewL.PutValue('GL_QTESTOCK', Tobl.GetValue('GL_QTERESTE'));
  for wj:=0 to TOBPiece.Detail.Count-1 do
    begin
      TobDet :=  TOBPiece.Detail[wj];
      if TobDet.Getvalue ('GL_TYPELIGNE') <> 'COM'   then
      begin
        TobDet.Putvalue('GL_QTESTOCK',TobDet.Getvalue ('GL_QTEFACT'));
        TobDet.Putvalue('GL_QTERESTE',TobDet.Getvalue ('GL_QTEFACT'));
        //
        //--- GUINIER ---
        //
        if CtrlOkReliquat(TobDet, 'GL') then TobDet.Putvalue('GL_MTRESTE',TobDet.Getvalue ('GL_MONTANTHTDEV'));
      end;
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : merieux
Créé le ...... : 06/06/2003
Modifié le ... :   /  /
Description .. : Mise à jour dans la tob lignes des ressources
Mots clefs ... : GIGA;FACTURATION
*****************************************************************}
procedure TAFPrepFact.MajLignesRess(TobP : TOB;MajRes:boolean);
Var TobLignes : TOB;
    wj : Integer;
BEGIN

 for wj:=0 to TOBP.Detail.Count-1 do
 begin
    TobLignes:=TOBP.Detail[wj] ;
    // Ajout pour mettre article de % en fin (utilisé si Sous-totaux par ressources
    if (toblignes.getValue('GL_TYPEARTICLE') = 'POU')
        then begin    // mcd 02/12/02 pour ne pas ecrire la ressource ZZZZ sur %
        if MajRes then Toblignes.PutValue('GL_RESSOURCE',DER_RESS);
        end;
    if (Not Majres) and ( toblignes.getValue('GL_RESSOURCE')=DER_RESS) then Toblignes.PutValue('GL_RESSOURCE','') ;

 end;

END;

{***********A.G.L.***********************************************
Auteur  ...... : Desseignet
Créé le ...... : 12/06/2003
Modifié le ... :   /  /
Description .. : Application des tarifs sur les lignes d'activités
Mots clefs ... : GIGA;FACTURE;TARIF
*****************************************************************}
procedure TAFPrepFact.ApplicTarif(TobP : TOB);
Var TobLignes ,TobTarif, tobart: TOB;
    wj : Integer;
    EnHt, QueQte : boolean;
    QQ: tquery;
    Zpuv, Zcalc : double;
BEGIN

 for wj:=0 to TOBP.Detail.Count-1 do
   BEGIN
   TobLignes:=TOBP.Detail[wj] ;

   if (GetParamSoc('So_AfRecalPvFact')) and (TobLignes.GetValue('GL_TYPELIGNE')='ART') then
     begin  // mcd 02/12/02 permet d'aller rechercher le tarif article si existe
     TOBTarif:=TOB.Create('TARIF',Nil,-1) ;
     EnHT:=(TOBPiece.GetValue('GP_FACTUREHT')='X') ;
     if (VH_GC.GCOuvreTarifQte) then QueQte :=True else QueQte :=False;
       // il faut renseigner la ton article de l'article en cours (fait sinon dans la fct suivante   creationLignePiece
      // mcd 11/04/03  if (Toblignes.GetValue('GL_ARTICLE')<> '') then
     if (Toblignes.GetValue('GL_ARTICLE')<> '') and (Toblignes.GetValue('GL_ARTICLE')<> GenFac.Zacompte) then
       Begin
       tobArt := Nil;
       zpuv := TobLignes.getvalue('GL_PUHT');
					//***Création des pieces, il faut tout prendre ****
          // on va rechercher le prix dans la table article ... il a pu être changer dans la tob dans le traitement avant.
       QQ:=OpenSQL('SELECT * FROM ARTICLE WHERE GA_ARTICLE="'+Toblignes.GetValue('GL_ARTICLE')+'" ',True,-1,'',true) ;
       if Not QQ.EOF then
         BEGIN
         TOBArt:=CreerTOBArt(TOBArticles) ;
         TOBArt.SelectDB('',QQ) ;
         TOBArt.addchampsup('GCA_LIBELLE',false);
         if (QQ.FindField('GA_PVHT').asfloat > 0) then Zpuv :=  QQ.FindField('GA_PVHT').asfloat; // on prend le prix article
         END;
       Ferme(QQ) ;
       if ChercheTarif (TobArt,TobTiers,TobLignes,TobP,TobTarif,QueQte,EnHT) then
         begin
           (* mcd 13/02/03if TOBP.GetValue('GP_SAISIECONTRE')='X' then Zpuv:=TOBTarif.GetValue('GF_PRIXCON')
                                                 else*)
            //mcd 21/08/03 modif pour traiter cas tarif en remise
           if TOBTarif.GetValue('GF_CALCULREMISE')<>'' then
           begin
             TobLIgnes.PutValue('GL_REMISELIGNE',TOBTarif.GetValue('GF_CALCULREMISE')); // mcd 21/08/03
             Zpuv:=TOBTarif.GetValue('GF_PRIXUNITAIRE') ;// pour OK prix + remise dans tarif
           end
           else
             Zpuv:=TOBTarif.GetValue('GF_PRIXUNITAIRE');
         end;
       TobTarif.free;
       TOBLignes.PutValue('GL_PUHT',zpuv) ;
       TOBLignes.PutValue('GL_PUHTDEV',PivotToDevise(zpuv,DEV.Taux,DEV.Quotite,DEV.Decimale)) ;
       zcalc:= Arrondi(Toblignes.GetValue('GL_QTEFACT') *  Toblignes.GetValue('GL_PUHTDEV'),V_PGI.OkDecV);
       Toblignes.PutValue('GL_TOTALHTDEV',zcalc);
       end;
    end;
    END;

END;

function TAFPrepFact.RecupProfil(prof : string):boolean ;
var TobDet : TOB;
    dim : string;
    zmini,zmaxi: double;
BEGIN

  result := true;

  if (prof = '') Then prof := 'DEF';

  //init Genfac  lié au profil
  Genfac.zacompte := '';
  Genfac.zlibacompte := '';
  Genfac.zlibpres := '';
  Genfac.zlibfrais := '';
  Genfac.zlibfour := '';
  Genfac.zlibctr := '';
  Genfac.zlibempcom := '';
  Genfac.zlibempcum := '';
  Genfac.zlibemppre := '';
  Genfac.zpres := 0;
  Genfac.zfrais := 0;
  Genfac.zfour := 0;
  Genfac.zctr := 0;
  Genfac.zcumaff := False;
  Genfac.zminifrais := 0;
  Genfac.zmaxifrais := 0;
  Genfac.zminifour := 0;
  Genfac.zmaxifour := 0;
  Genfac.zreplibech := '-';
  Genfac.zcomaff := '';
  Genfac.zcomaffsup := '';
  Genfac.zlibstotcli := '';
  Genfac.zlibcumsaff :='';
  Genfac.zlibcomsaff :='';
  Genfac.zcomsaff := false;
  Genfac.zcumsaff := false;
  Genfac.zcomemp := false;
  Genfac.zcumemp := false;
  Genfac.zemppre := false;
  Genfac.zacpte :=false;
  Genfac.ztout :=false;
  Genfac.zpour :=false;
  Genfac.zstotcli :=false;

  Genfac.zprof := prof;

  TobDet := TobProfil.FindFirst(['APG_PROFILGENER','APG_ELEMENT'],[Genfac.zprof,'G00'],False);
  if TobDet <> NIL Then
    Begin
      Genfac.zacompte := TobDet.GetValue('APG_ARTICLE');
      Genfac.zreplibech := TobDet.GetValue('APG_REPLIBECH');
      if (Genfac.zreplibech = '') then Genfac.zreplibech := '-';
      Genfac.zlibacompte := TobDet.GetValue('APG_LIBART');
      Genfac.zacpte := true;
    End;

  if (Genfac.zacompte = '') then
     Genfac.zacompte := CodeArticleUnique(VH_GC.AfAcompte,dim,dim,dim,dim,dim);
   // je suis sure qu'il n'auront pas besoin d'une ligen fourre-tout
   // car ils reprennent les lignes intégrales de l'affaire (évite une lecture)
  if not VH_GC.GCIfDefCEGID then
  if ((Genfac.zlibacompte = '') and (Genfac.zreplibech = '-')) then
      Genfac.zlibacompte := LibelleArticleGenerique(trim(CodeArticleGenerique(Genfac.zacompte,dim,dim,dim,dim,dim)));
  //*****************************

  TobDet := TobProfil.FindFirst(['APG_PROFILGENER','APG_ELEMENT'],[Genfac.zprof,'TOU'],False);
  if Tobdet <> NIL Then
    Begin
      Genfac.ztout := true;
    End;

//*****************************

  TobDet := TobProfil.FindFirst(['APG_PROFILGENER','APG_ELEMENT'],[Genfac.zprof,'A00'],False);
  if Tobdet <> NIL Then
    Begin
      Genfac.zcumaff := (TobDet.GetValue('APG_CUMUL') = 'X');
    End;
//***************************
  TobDet :=TobProfil.FindFirst(['APG_PROFILGENER','APG_ELEMENT'],[Genfac.zprof,'ACO'],False);
  if TobDet <> NIL Then
  Begin
       Genfac.zcomsaff := true;
       Genfac.zlibcomsaff := TobDet.GetValue('APG_LIBART');
  End;
  TobDet := TobProfil.FindFirst(['APG_PROFILGENER','APG_ELEMENT'],[Genfac.zprof,'AST'],False);
  if TobDet  <> NIL Then
   Begin
       Genfac.zcumsaff := true;
       Genfac.zlibcumsaff:= TobDet.GetValue('APG_LIBART');
  End;


//*********************
  TobDet :=TobProfil.FindFirst(['APG_PROFILGENER','APG_ELEMENT'],[Genfac.zprof,'ECO'],False);
  if TobDet <> NIL Then
  Begin
       Genfac.zcomemp := true;
       Genfac.zlibempcom := TobDet.GetValue('APG_LIBART');
  End;

  TobDet := TobProfil.FindFirst(['APG_PROFILGENER','APG_ELEMENT'],[Genfac.zprof,'EST'],False);
  if TobDet  <> NIL Then
   Begin
       Genfac.zcumemp := true;
       Genfac.zlibempcum := TobDet.GetValue('APG_LIBART');
  End;
//*********************
  TobDet :=TobProfil.FindFirst(['APG_PROFILGENER','APG_ELEMENT'],[Genfac.zprof,'EPR'],False);
  if TobDet <> NIL Then
  Begin
   	  Genfac.zemppre := true;
      Genfac.zlibemppre := TobDet.GetValue('APG_LIBART');
  End;




  TobDet := TobProfil.FindFirst(['APG_PROFILGENER','APG_ELEMENT'],[Genfac.zprof,'C10'],False);
  if TobDet <> NIL Then
    Begin
      Genfac.zcomaff := TobDet.GetValue('APG_LIBART');
      Genfac.zcomaffsup := TobDet.GetValue('APG_LIBSUP');
    End;



  Genfac.zcomlig := false;
  TobDet := TobProfil.FindFirst(['APG_PROFILGENER','APG_ELEMENT'],[Genfac.zprof,'C00'],False);
  if TobDet <> NIL Then
    Begin
      Genfac.zcomlig := true;
    End;

  Genfac.zpour := false;
  TobDet := TobProfil.FindFirst(['APG_PROFILGENER','APG_ELEMENT'],[Genfac.zprof,'XPO'],False);
  if TobDet <> NIL Then
    Begin
      Genfac.zpour := true;
    End;

    Genfac.zstotcli := false;
    TobDet := TobProfil.FindFirst(['APG_PROFILGENER','APG_ELEMENT'],[Genfac.zprof,'XCL'],False);
    if TobDet <> NIL Then
      Begin
        Genfac.zstotcli := true;
        Genfac.zlibstotcli := TobDet.GetValue('APG_LIBART');
      End;

    Genfac.zpres := RechProfilPrest('P',tobProfil,Genfac.zlibpres,zmini,zmaxi);
    Genfac.zctr := RechProfilPrest('N',tobProfil,Genfac.zlibctr,zmini,zmaxi);
    Genfac.zfour := RechProfilPrest('S',tobProfil,Genfac.zlibfour,Genfac.zminifour,Genfac.zmaxifour);
    Genfac.zfrais := RechProfilPrest('V',tobProfil,Genfac.zlibfrais,Genfac.zminifrais,Genfac.zmaxifrais);

    // On n'aura pas besoin de Sous total seulement si on a :
    // Detail ou cumul sur code Prestation ET idem sur frais ET idem sur Four
    //  ET si affaire non regroupée (car dans le cas d'une affaire regroupée, on peut
    //  avoir 2 lignes "forfait acompte" issue des 2 affaires et donc cumulées sur un sous
    //  total qui sera édité
    Genfac.zstot := true ;
    if (((Genfac.zpres = 0) or (Genfac.zpres = 1))   AND
       ((Genfac.zfour = 0) or (Genfac.zfour = 1))   AND
       ((Genfac.zfrais = 0) or (Genfac.zfrais = 1)) AND
       ((Genfac.zctr = 0) or (Genfac.zctr = 1)) AND
       not(Genfac.zcumaff) ) then Genfac.zstot := false;   // GMGMGM 15/12/2000

   if (Genfac.zcumemp) or (Genfac.zcumsaff) or (Genfac.zcomsaff) then Genfac.zstot := false;



END;

// Recup  dans la table PROFILGENER des elements du profil Prest/Frais/Fourn
Function TAFPrepFact.RechProfilPrest(typ : string; tobprofil : TOB;var zlib : string;var zmini,zmaxi : Double) : Integer;
Var st0,st1,st2,st3 : string;
  TobDet : Tob;
BEGIN
      zlib:='';
      zmini := 0; zmaxi := 0;
      st0:=typ+'00';  st1:=typ+'10';
      st2:=typ+'20';  st3:=typ+'30';
      result := 9;
      Tobdet :=  TobProfil.FindFirst(['APG_PROFILGENER','APG_ELEMENT'],[Genfac.zprof,st0],False);
      if ( Tobdet <> NIL) Then
        Begin
        Result := 0;
        zlib := TobDet.GetValue('APG_LIBART');
      End
      else
        if (TobProfil.FindFirst(['APG_PROFILGENER','APG_ELEMENT'],[Genfac.zprof,st1],False) <> NIL) Then
          Result := 1
        else
          Begin
            TobDet := TobProfil.FindFirst(['APG_PROFILGENER','APG_ELEMENT'],[Genfac.zprof,st2],False);
            if (TobDet <> NIL) Then
            Begin
              Result := 2;
              zlib := TobDet.GetValue('APG_LIBART');
              zmaxi := TobDet.GetValue('APG_SEUILMAXI');
              zmini:= TobDet.GetValue('APG_SEUILMINI');
            End
            else
              if (TobProfil.FindFirst(['APG_PROFILGENER','APG_ELEMENT'],[Genfac.zprof,st3],False) <> NIL) Then
              Result := 3;
          End;
END;

{procedure TAFPrepFact.RecalculeSousTotaux ;
Var i : integer ;
    TOBL : TOB ;
BEGIN
for i:=0 to TOBPiece_OK.Detail.Count-1 do
    BEGIN
    TOBL:=TOBPiece_OK.Detail[i] ;
    if TOBL.GetValue('GL_TYPELIGNE')='TOT' then
       BEGIN
       SommerLignes(TOBPiece_OK,i+1,0) ;
       END ;
    END ;
END ;
}

{procedure TAFPrepFact.MajBasePorc;
var wi,wz : Integer;
    base : Double;
    tobDet,tobdet1 : TOB;
BEGIN
    Base:=0;
    for wi:=0 to TobPorcs.detail.count-1 do
     begin
          TobDET := Tobporcs.detail[wi];
          if (TOBDet.GetValue('GPT_TYPEPORT') <> 'MT') then
          begin
          for wz:=0 to TobPiece_OK.detail.count-1 do
           begin
                TobDET1 := Tobpiece_ok.detail[wz];
                if (TobDet1.getvalue('GL_TYPELIGNE')='ART') and (TobDet1.getvalue('GL_TYPEARTICLE')<>'POU') then base:=base +TobDet1.getvalue('GL_TOTALHTDEV');
                if (TobDet1.getvalue('GL_TYPELIGNE')='ART') and (TobDet1.getvalue('GL_TYPEARTICLE')='POU') then base:=base +(TobDet1.getvalue('GL_TOTALHTDEV')/100);// prix par 100 par encore appliqué
                end;
          TOBdet.PutValue('GPT_BASEHTDEV',base) ;
          end;
     end;
END;
}

procedure InitToutGenfac(GenFac : R_GENFAC);
BEGIN
  Genfac.zdatedeb:= '';
  Genfac.zdatefin:= '';
  Genfac.zcomfac:= '';
  Genfac.zacompte:= '';
  Genfac.zlibacompte:= '';
  Genfac.zcomaff := '';
  Genfac.zcomaffsup := '';
  Genfac.zlibstotcli := '';
  Genfac.zdateact_d:= '';
  Genfac.zdateact_f:= '';
  Genfac.zgen:= '';
  Genfac.zprof:= '';
  Genfac.ztypgen:= '';
  Genfac.ztypech:= '';
  Genfac.zcumaff := False;
  Genfac.zpres := 0;
  Genfac.zfrais := 0;
  Genfac.zfour := 0;
  Genfac.zstot := False;
  Genfac.zcomlig := False;
  Genfac.zlibpres := '';
  Genfac.zlibfrais := '';
  Genfac.zlibfour := '';
  Genfac.zlibctr :='';
  Genfac.zlibempcom := '';
  Genfac.zlibempcum := '';
  Genfac.zlibemppre := '';
  Genfac.zlibcumsaff :='';
  Genfac.zlibcomsaff :='';
  Genfac.zminifrais := 0;
  Genfac.zmaxifrais := 0;
  Genfac.zminifour := 0;
  Genfac.zmaxifour := 0;
  Genfac.zreplibech := '-';
  Genfac.zcomsaff := false;
  Genfac.zcumsaff := false;
  Genfac.zcomemp := false;
  Genfac.zcumemp := false;
  Genfac.zemppre := false;
  Genfac.zacpte :=false;
  Genfac.ztout :=false;
  Genfac.zcom1:= '';
  Genfac.zcom2:= '';
  Genfac.znat := '';
  Genfac.zpour := False;
  Genfac.zstotcli := False;
END;

procedure TAFPrepFact.RecupProrataDateFac;
var    DebEch,DebPer,FinPer : TDateTime;
BEGIN
  DebEch :=  StrToDate(TobEchFFact.getValue('AFA_DATEECHE'));

  DebPer :=  StrToDate(TOBAffaire.GetValue('AFF_DATEDEBGENER'));

  FinPer :=  StrToDate(TOBAffaire.GetValue('AFF_DATEFINGENER'));
  if ( StrToDate(TOBAffaire.GetValue('AFF_DATERESIL')) < FinPer) then
  	FinPer := StrToDate(TOBAffaire.GetValue('AFF_DATERESIL'));

  xDebFac := idate1900;
  xFinFac := idate2099;

  // modif BRL 160707 (climatech) : on ne passe plus le nombre d'inter car le montant est déjà
  // multiplié dans le montant d'échéance
  //     	xprorata :=  CalculProrata(TOBAffaire.GetValue('AFF_METHECHEANCE'),TOBAffaire.GetValue('AFF_PERIODICITE'),
  //      Integer(TOBAffaire.GetValue('AFF_INTERVALGENER')),DebEch,DebPer,FinPer,xDebFac,xFinFac);
  xprorata :=  CalculProrata(TOBAffaire.GetValue('AFF_METHECHEANCE'),TOBAffaire.GetValue('AFF_PERIODICITE'),
  														1,DebEch,DebPer,FinPer,xDebFac,xFinFac);
END;

procedure TAFPrepFact.RechNumeroPiece ;
Var NewNum : integer ;
BEGIN
   NewNum := SetNumberAttribution (Znature,zsouche,DateFacture,1);
//   NewNum:=TestSetNumberAttribution (ZSouche);
   if NewNum > 0 then ZNewNum:=NewNum;
//   if NewNum > 0 then CleDoc.NumeroPiece:=NewNum;
END ;
(*
Function TAFPrepFact.TestSetNumberAttribution ( Souche : string ) : integer ;
Var SoucheG : String3 ;
    QQ      : TQuery ;
    NumDef  : Longint ;
    Nb,ii : integer ;
    Okok : boolean;
BEGIN
Result:=0 ;
//SoucheG:=TOBPiece.GetValue('GP_SOUCHE') ; if SoucheG='' then Exit ;
SoucheG:=Souche ; if SoucheG='' then Exit ;
ii:=0 ;
// Okok:=False ;
Repeat
 NumDef:=-1 ; inc(ii) ;
 QQ:=OpenSQL('SELECT SH_NUMDEPART FROM SOUCHE WHERE SH_TYPE="GES" AND SH_SOUCHE="'+SoucheG+'"',True,-1,'',true) ;
 if Not QQ.EOF then NumDef:=QQ.Fields[0].AsInteger ;
 Ferme(QQ) ;
 if NumDef<=0 then exit ;
 Nb:=ExecuteSQL('UPDATE SOUCHE SET SH_NUMDEPART='+IntToStr(NumDef+1)+' WHERE SH_TYPE="GES" AND SH_SOUCHE="'+SoucheG+'" AND SH_NUMDEPART='+IntToStr(NumDef)) ;
 Okok:=(Nb=1) ;
Until ((Okok) or (ii>20)) ;
if Not Okok then BEGIN V_PGI.IoError:=oeUnknown ; Exit ; END ;
Result:=NumDef ;
END;
*)
{***********A.G.L.***********************************************
Auteur  ...... : G.Merieux
Créé le ...... : 24/08/2001
Modifié le ... :   /  /
Description .. : Chargement des profils facture dans une Tob
Mots clefs ... :
*****************************************************************}
Function ChargeTobProfil(TobProf : TOB) : boolean;
var QQ : Tquery;
		wi : integer;
    TobDet: TOB;
    Codeart,art,dim : string;
BEGIN

//chargement de tous les profils, afin de ne pas avoir à les relire a chaque affaire
// on passera la tobprofil en paramètre
  result := true;
  QQ := nil;
  Try;
  QQ := OpenSQL('SELECT * FROM PROFILGENER order by APG_PROFILGENER,APG_NUMORDRE',True,-1,'',true) ;
  If Not QQ.EOF then TOBProf.LoadDetailDB('PROFILGENER','','',QQ,True);
  Finally
   Ferme(QQ);
   End;

	for wi := 0 to TobProf.detail.count-1  do
  begin
  	TobDet := TobProf.detail[wi];
    if (TobDet.GetValue('APG_ELEMENT') = 'G00') then
		begin
    dim := '';
    art := TobDet.GetValue('APG_ARTICLE');
    Codeart := trim(CodeArticleGenerique(art,dim,dim,dim,dim,dim));
    QQ:=OpenSQL('SELECT GA_LIBELLE FROM ARTICLE WHERE GA_CODEARTICLE="'+ codeart +'"',true,-1,'',true);

    if QQ.EOF then
        result := false;
    ferme(QQ);
		end;
  end;

END;


{***********A.G.L.***********************************************
Auteur  ...... : merieux
Créé le ...... : 26/05/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... : GIGA;FACTURE;REVISION;VARIABLES
*****************************************************************}
function TAFPrepFact.GestionRevision(pStRevLogFile : String) : boolean;
Var   pDtEcheance : TdateTime;
      iInd : Integer;
      TobLig : TOB;
Begin
// Traitement des révisions de prix

  for iInd:=0  to Tobpiece_ok.Detail.count-1 do
  Begin
    TobLig     :=  TobPiece_ok.Detail[iInd];
    if (toblig.GetValue('GL_TYPELIGNE') = 'ART') and (toblig.GetValue('GL_TYPEARTICLE') <> 'POU') then
    begin
    // pas de révision si on gére un cumul sur code
    // car dans ce cas on a perdu le détail des lignes artilces, et le prix correspond au montant cumulé
      if not(TestSiCumulArticle(Toblig)) then
      begin
       toblig.PutValue('GL_FORCODE1','');
       toblig.PutValue('GL_FORCODE2','');
      end;
    end;
  end;
  //gcvoirtob(TobPiece_OK);
  pDtEcheance := StrToDate(Rech.DateEch);
  result := LectureCoefFacture(TobPiece_OK, fTobRevision ,pDtEcheance, pStRevLogFile, True);
  //gcvoirtob(TobPiece_OK);


End;

{***********A.G.L.***********************************************
Auteur  ...... : merieux
Créé le ...... : 26/05/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... : GIGA;FACTURE;VARIABLES
*****************************************************************}
{Procedure TAFPrepFact.MajVariable;
Var   TobLig,TobVar: TOB;
      iInd,wj : Integer;
Begin
for iInd:=0  to Tobpiece_ok.Detail.count-1 do
  Begin
    TobLig     :=  TobPiece_ok.Detail[iInd];
    if (toblig.GetValue('GL_TYPELIGNE') = 'ART') and (toblig.GetValue('GL_TYPEARTICLE') <> 'POU') then
    begin
      if not(TestSiCumulArticle(Toblig)) then
        continue
      else
      begin
        TobVar     :=  TobVarQte_LIG.FindFirst(['AVVLIENVAR'],[TobLig.GetValue('GLLLIENVAR')],true);
         if (TobVar <> NIL) then
         begin
            TobVar.PutValue('AVV_NATUREPIECEG',TobLig.getvalue('GL_NATUREPIECEG'));
            TobVar.PutValue('AVV_ORIGVAR',TobLig.getvalue('GL_NATUREPIECEG'));
            TobVar.PutValue('AVV_SOUCHE',TobLig.getvalue('GL_SOUCHE'));
            TobVar.PutValue('AVV_NUMERO',TobLig.getvalue('GL_NUMERO'));
            TobVar.PutValue('AVV_NUMORDRE',TobLig.getvalue('GL_NUMORDRE'));
         end;
      end;
    end;
  end;

  wj := 0; Iind := 0;
  while  (wj < TobVarQte_LIG.Detail.count) do
  begin
    TobVar := TobVarQte_LIG.Detail[wj];
    if tobvar.getValue('AVV_NUMERO') = 0 then
      TobVar.free
    else
      Inc(wj);
    TobVar.PutValue('AVV_RANGVAR',iInd+1);
    Inc(Iind);
  end;

  if (TobVarQte_LIG.detail.count <> 0) then
    TobVarQte_LIG.InsertDB(Nil);
End;
}

{***********A.G.L.***********************************************
Auteur  ...... : Merieux
Créé le ...... : 27/06/2003
Modifié le ... : 27/06/2003
Description .. : contrôle du mode de cumul.
Suite ........ : Si on a un cumul sur code artilce, on perd la notion de
Suite ........ : variable
Mots clefs ... : GIGA;FACTURATION;VARIABLE
*****************************************************************}
function TAFPrepFact.TestSiCumulArticle(Toblig : TOB):boolean;
begin
  result := true;
  if (Genfac.zemppre)  then     // profil ressource par prestation
    result := false
  else
    if (toblig.GetValue('GL_TYPEARTICLE') = 'PRE') and (Genfac.zpres=1) then
      result := false
    else
      if (toblig.GetValue('GL_TYPEARTICLE') = 'MAR') and (Genfac.zfour=1) then
        result := false
      else
        if (toblig.GetValue('GL_TYPEARTICLE') = 'FRA') and (Genfac.zfrais=1) then
          result := false
        else
        if (toblig.GetValue('GL_TYPEARTICLE') = 'CTR') and (Genfac.zctr=1) then
          result := false;

  if not(result) then toblig.PutValue('GL_FORMULEVAR','');
end;


{***********A.G.L.***********************************************
Auteur  ...... : merieux
Créé le ...... : 26/05/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... : GIGA;FACTURE;REVISION
*****************************************************************}
Procedure TAFPrepFact.MajRevision;
Var   TobLig: TOB;
      iInd : Integer;
Begin
  for iInd:=0  to Tobpiece_ok.Detail.count-1 do
  Begin
    TobLig     :=  TobPiece_ok.Detail[iInd];
    if (toblig.GetValue('GL_TYPELIGNE') = 'ART') and (toblig.GetValue('GL_TYPEARTICLE') <> 'POU') then
    begin
//      MajTobRev(TobLig,1);
//      MajTobRev(TobLig,2);
      MajTobRev(fTobRevision, TobLig, 1, False);
      MajTobRev(fTobRevision, TobLig, 2, False);
    end;
  end;

  if (fTobRevision.detail.count <> 0) then
    fTobRevision.InsertDB(Nil);

End;

procedure TAFPrepFact.InitTobLigneApresDupplic(TobLig: TOB);
begin

  TobLig.PutValue('GL_QTEFACT',0);
  TobLig.PutValue('GL_QTESTOCK',0);
  TobLig.PutValue('GL_QTETARIF',0);
  TobLig.PutValue('GL_QTERESTE',0);
  //--- GUINIER
  TobLig.PutValue('GL_MTRESTE',0);
  TobLig.PutValue('GL_PUHTDEV',0);
  TobLig.PutValue('GL_TOTALHTDEV',0);
  TobLig.PutValue('GL_PMRP',0);
  TobLig.PutValue('GL_PMAPACTU',0);
  TobLig.PutValue('GL_PMRPACTU',0);
  TobLig.PutValue('GL_ARTICLE','');
  TobLig.PutValue('GL_CODEARTICLE','');
  TobLig.PutValue('GL_REFARTSAISIE','');
  TobLig.PutValue('GL_TYPEARTICLE','');
  TobLig.PutValue('GL_LIBCOMPL','');
  TobLig.PutValue('GL_FACTURABLE','');
  TobLig.PutValue('GL_GENERAUTO','');
  TobLig.PutValue('GL_LIGNELIEE',0);
  TobLig.PutValue('GL_FORMULEVAR','');
  TobLig.PutValue('GL_FORCODE1','');
  TobLig.PutValue('GL_FORCODE2','');
  TobLig.PutValue('GL_QUALIFQTEVTE','');
end;


procedure TAFPrepFact.TraitePrixLigneCourante (TOBL : TOB;Pu : double);
var TypeArticle : string;
		EnHt : boolean;
begin
//
	TypeArticle := TOBL.getValue('GL_TYPEARTICLE');
  EnHt := (TOBL.getValue('GL_FACTUREHT')='X');
  if (TypeArticle = 'OUV') or (TypeArticle = 'ARP') then
  begin
  	TraitePrixOuvrage(TOBPiece,TOBL,Nil,nil,TobOuvrages, EnHt, Pu,0,DEV,false);
  end;
//
  if EnHt then TOBL.PutValue('GL_PUHTDEV',Pu)
          else TOBL.PutValue('GL_PUTTCDEV',Pu);
//
  TOBL.Putvalue('GL_COEFMARG', 0);
  TOBL.Putvalue('POURCENTMARG', 0);
  TOBL.Putvalue('POURCENTMARQ', 0);
	TOBL.putValue('GL_BLOQUETARIF','X');
end;

procedure TAFPrepFact.PositionneInfosLigneSurOuv (TOBL :TOB)
;
  procedure  PosInfosLigOuv (TOBL,TOBO : TOB);
  begin
    TOBO.putValue('BLO_AFFAIRE',TOBL.GetValue('GL_AFFAIRE'));
    TOBO.putValue('BLO_AFFAIRE1',TOBL.GetValue('GL_AFFAIRE1'));
    TOBO.putValue('BLO_AFFAIRE2',TOBL.GetValue('GL_AFFAIRE2'));
    TOBO.putValue('BLO_AFFAIRE3',TOBL.GetValue('GL_AFFAIRE3'));
    TOBO.putValue('BLO_AVENANT',TOBL.GetValue('GL_AVENANT'));
    TOBO.putValue('BLO_NUMLIGNE',TOBL.GetValue('GL_NUMLIGNE'));
    TOBO.putValue('BLO_NATUREPIECEG',TOBL.GetValue('GL_NATUREPIECEG'));
    TOBO.putValue('BLO_SOUCHE',TOBL.GetValue('GL_SOUCHE'));
    TOBO.putValue('BLO_NUMERO',TOBL.GetValue('GL_NUMERO'));
    TOBO.putValue('BLO_INDICEG',TOBL.GetValue('GL_INDICEG'));
    TOBO.putValue('BLO_DATEPIECE',TOBL.GetValue('GL_DATEPIECE'));
  end;

  procedure TraiteSousDetail (TOBOUvrage,TOBL : TOB);
  var Indice : integer;
      TOBO : TOB;
  begin
    for Indice := 0 To TOBOUvrage.detail.count -1 do
    begin
      TOBO := TOBOuvrage.detail[Indice];
      PosInfosLigOuv (TOBL,TOBO);
      if TOBO.detail.count > 0 then
      begin
        TraiteSousDetail (TOBO,TOBL);
      end;
    end;
  end;

var ZIndiceNomen: integer;
    TOBO,TOBOUvrage : TOB;
begin
  if TOBL.GetValue('GL_TYPELIGNE') <> 'ART' Then exit;
  if Pos(TOBL.GetValue('GL_TYPEARTICLE'),'OUV;ARP;')=0 then exit;
  if TOBL.GetValue('GL_INDICENOMEN')=0 then exit;

  ZIndiceNomen := TOBL.GetValue('GL_INDICENOMEN')-1;
  TOBOuvrage := TobOuvrages.detail[ZIndiceNomen];
  TraiteSOusDetail (TOBOuvrage,TOBL);
end;

procedure TAFPrepFact.AjouteLigneEcart(TOBPiece: TOB; Ecart: double);
var TOBL,TOBA : TOB;
    QQ : Tquery;
    NumL,j : integer;
begin
  TOBA := TOB.Create ('ARTICLE',nil,-1);
  TOBL := NewTOBLigne(TOBPiece, 0);
  QQ := OpenSql ('SELECT * FROM ARTICLE WHERE GA_ARTICLE="'+ArtEcart+'"',true,-1,'',true);
  if not QQ.eof then
  begin
    TOBA.SelectDB ('',QQ);
    PieceVersLigne ( TOBPiece,TobL) ;
    TOBL.putValue('GL_TYPELIGNE','ART');
    TOBL.putValue('GL_QTEFACT', 1);
    TOBL.putValue('GL_QTESTOCK',1);
    TOBL.putValue('GL_QTERESTE',1);
    //--- GUINIER ---
    TOBL.putValue('GL_MTRESTE',1);
    TOBL.putValue('GL_ARTICLE',TOBA.getValue('GA_ARTICLE'));
    TOBL.putValue('GL_LIBELLE',TOBA.getValue('GA_LIBELLE'));
    TOBL.putValue('GL_CODEARTICLE',TOBA.getValue('GA_CODEARTICLE'));
    TOBL.PutValue('GL_TYPEARTICLE',TOBA.getValue('GA_TYPEARTICLE'));
    TOBL.PutValue('GL_REFARTSAISIE',TOBA.getValue('GA_CODEARTICLE'));
    TOBL.putValue('GL_PUHTDEV',Ecart);
    TOBL.PutValue('GL_AFFAIRE',TOBAffaire.GetValue('AFF_AFFAIRE'));
    TOBL.PutValue('GL_AFFAIRE1',TOBAffaire.GetValue('AFF_AFFAIRE1'));
    TOBL.PutValue('GL_AFFAIRE2',TOBAffaire.GetValue('AFF_AFFAIRE2'));
    TOBL.PutValue('GL_AFFAIRE3',TOBAffaire.GetValue('AFF_AFFAIRE3'));
    TOBL.PutValue('GL_AVENANT',TOBAffaire.GetValue('AFF_AVENANT'));
    TOBL.PutValue('GL_REPRESENTANT',TOBAffaire.GetValue('LEREP'));
//
    TOBL.PutValue('GL_LIBCOMPL', TOBA.GetValue('GA_LIBCOMPL'));
    TOBL.PutValue('GL_REFARTBARRE', TOBA.GetValue('GA_CODEBARRE'));
    TOBL.PutValue('GL_ESCOMPTABLE', TOBA.GetValue('GA_ESCOMPTABLE'));
    TOBL.PutValue('GL_REMISABLEPIED', TOBA.GetValue('GA_REMISEPIED'));
    TOBL.PutValue('GL_REMISABLELIGNE', TOBA.GetValue('GA_REMISELIGNE'));
    TOBL.PutValue('GL_NUMEROSERIE', TOBA.GetValue('GA_NUMEROSERIE'));
    TOBL.PutValue('GL_PAYSORIGINE', TOBA.GetValue('GA_PAYSORIGINE'));  //JS FI10226
    TOBL.PutValue('GL_CODEDOUANE', TOBA.GetValue('GA_CODEDOUANIER'));
    {Familles, collection, domaine}
    TOBL.PutValue('GL_FAMILLENIV1', TOBA.GetValue('GA_FAMILLENIV1'));
    TOBL.PutValue('GL_FAMILLENIV2', TOBA.GetValue('GA_FAMILLENIV2'));
    TOBL.PutValue('GL_FAMILLENIV3', TOBA.GetValue('GA_FAMILLENIV3'));
    TOBL.PutValue('GL_COLLECTION', TOBA.GetValue('GA_COLLECTION'));
    TOBL.PutValue('GL_DOMAINE', TOBPiece.GetValue('GP_DOMAINE'));
    if TOBPiece.GetValue('GP_DOMAINE')='' then
    begin
    	TOBL.PutValue('GL_DOMAINE', TOBA.GetValue('GA_DOMAINE'));
    end;
    {Nomenclature}
    TOBL.PutValue('GL_TYPEARTICLE', TOBA.GetValue('GA_TYPEARTICLE'));
    TOBL.PutValue('GL_TYPENOMENC', TOBA.GetValue('GA_TYPENOMENC'));
    for j := 1 to 5 do TOBL.PutValue('GL_FAMILLETAXE' + IntToStr(j), TOBA.GetValue('GA_FAMILLETAXE' + IntToStr(j)));
    {Unités de mesure}
    TOBL.PutValue('GL_QUALIFSURFACE', TOBA.GetValue('GA_QUALIFSURFACE'));
    TOBL.PutValue('GL_QUALIFVOLUME', TOBA.GetValue('GA_QUALIFVOLUME'));
    TOBL.PutValue('GL_QUALIFPOIDS', TOBA.GetValue('GA_QUALIFPOIDS'));
    TOBL.PutValue('GL_QUALIFLINEAIRE', TOBA.GetValue('GA_QUALIFLINEAIRE'));
    TOBL.PutValue('GL_QUALIFHEURE', TOBA.GetValue('GA_QUALIFHEURE'));
    TOBL.PutValue('GL_SURFACE', TOBA.GetValue('GA_SURFACE'));
    TOBL.PutValue('GL_VOLUME', TOBA.GetValue('GA_VOLUME'));
    TOBL.PutValue('GL_POIDSBRUT', TOBA.GetValue('GA_POIDSBRUT'));
    TOBL.PutValue('GL_POIDSNET', TOBA.GetValue('GA_POIDSNET'));
    TOBL.PutValue('GL_POIDSDOUA', TOBA.GetValue('GA_POIDSDOUA'));
    TOBL.PutValue('GL_LINEAIRE', TOBA.GetValue('GA_LINEAIRE'));
//    TOBL.PutValue('GL_HEURE', TOBA.GetValue('GA_HEURE'));
    TOBL.PutValue('GL_QUALIFQTESTO', TOBA.GetValue('GA_QUALIFUNITESTO'));
    TOBL.PutValue('GL_QUALIFQTEVTE', TOBA.GetValue('GA_QUALIFUNITEVTE'));
    TOBL.PutValue('GL_COEFCONVQTEVTE', TOBA.GetValue('GA_COEFCONVQTEVTE'));  // coef US/UV
    {Tables libres}
    for j := 1 to 9 do TOBL.PutValue('GL_LIBREART' + IntToStr(j), TOBA.GetValue('GA_LIBREART' + IntToStr(j)));
    TOBL.PutValue('GL_LIBREARTA', TOBA.GetValue('GA_LIBREARTA'));
    TOBL.PutValue('GL_RECALCULER', 'X');
//
  end;
  ferme (QQ);
  TOBA.free;
end;

procedure TAFPrepFact.ReinitAffairePieceReg;
begin
	TOBpiece.putValue('GP_AFFAIRE','');
	TOBpiece.putValue('GP_AFFAIRE1','');
	TOBpiece.putValue('GP_AFFAIRE2','');
	TOBpiece.putValue('GP_AFFAIRE3','');
	TOBpiece.putValue('GP_AVENANT','');
end;

procedure TAFPrepFact.ChargelesTextes(TOBPiece,TOBOLES: TOB);
var Q : TQuery;
		Cledoc : R_CLEDOC;
begin
	TOBOLES.ClearDetail;
  Cledoc := TOB2CleDoc(TOBPiece);
  Q := OpenSQL('SELECT * FROM LIENSOLE WHERE ' + WherePiece(CleDoc, ttdLienOle, False), True,-1, '', True);
  TOBOLES.LoadDetailDB('LIENSOLE', '', '', Q, False);
  Ferme(Q);
end;

procedure GetAdrFromAuxi(TOBAdr: TOB; Auxi: string);

  procedure GetAdrFromTOB(TOBAdr, TOBTiers: TOB; ForceTobAdresse: Boolean = False);

    procedure PutInTobAdresse;
    begin
      TOBAdr.PutValue('ADR_LIBELLE',    TOBTiers.GetValue('T_LIBELLE'));
      TOBAdr.PutValue('ADR_LIBELLE2',   TOBTiers.GetValue('T_PRENOM'));
      TOBAdr.PutValue('ADR_JURIDIQUE',  TOBTiers.GetValue('T_JURIDIQUE'));
      TOBAdr.PutValue('ADR_ADRESSE1',   TOBTiers.GetValue('T_ADRESSE1'));
      TOBAdr.PutValue('ADR_ADRESSE2',   TOBTiers.GetValue('T_ADRESSE2'));
      TOBAdr.PutValue('ADR_ADRESSE3',   TOBTiers.GetValue('T_ADRESSE3'));
      TOBAdr.PutValue('ADR_CODEPOSTAL', TOBTiers.GetValue('T_CODEPOSTAL'));
      TOBAdr.PutValue('ADR_VILLE',      TOBTiers.GetValue('T_VILLE'));
      TOBAdr.PutValue('ADR_PAYS',       TOBTiers.GetValue('T_PAYS'));
      TOBAdr.PutValue('ADR_TELEPHONE',  TOBTiers.GetValue('T_TELEPHONE'));
      TobAdr.PutValue('ADR_INCOTERM',   TobTiers.GetValue('YTC_INCOTERM'));
      TobAdr.PutValue('ADR_MODEEXP',    TobTiers.GetValue('YTC_MODEEXP'));
      TobAdr.PutValue('ADR_LIEUDISPO',  TobTiers.GetValue('YTC_LIEUDISPO'));
      TobAdr.PutValue('ADR_EAN',        TobTiers.GetValue('T_EAN'));
      TobAdr.PutValue('ADR_NIF',        TobTiers.GetValue('T_NIF'));
      TobAdr.PutValue('ADR_REGION',     TobTiers.GetValue('T_REGION'));
      //
      TobAdr.PutValue('ADR_NUMEROCONTACT', wGetSqlFieldValue('C_NUMEROCONTACT', 'CONTACT', 'C_AUXILIAIRE = "' + TOBTiers.GetValue('T_AUXILIAIRE') + '" AND C_PRINCIPAL = "' + wTrue + '"'));
    end;

  begin
    if ForceTobAdresse then
      PutInTobAdresse
    else
    begin
      if GetParamSoc('SO_GCPIECEADRESSE') then
      begin
        TOBAdr.PutValue('GPA_LIBELLE',    TOBTiers.GetValue('T_LIBELLE'));
        TOBAdr.PutValue('GPA_LIBELLE2',   TOBTiers.GetValue('T_PRENOM'));
        TOBAdr.PutValue('GPA_JURIDIQUE',  TOBTiers.GetValue('T_JURIDIQUE'));
        TOBAdr.PutValue('GPA_ADRESSE1',   TOBTiers.GetValue('T_ADRESSE1'));
        TOBAdr.PutValue('GPA_ADRESSE2',   TOBTiers.GetValue('T_ADRESSE2'));
        TOBAdr.PutValue('GPA_ADRESSE3',   TOBTiers.GetValue('T_ADRESSE3'));
        TOBAdr.PutValue('GPA_CODEPOSTAL', TOBTiers.GetValue('T_CODEPOSTAL'));
        TOBAdr.PutValue('GPA_VILLE',      TOBTiers.GetValue('T_VILLE'));
        TOBAdr.PutValue('GPA_PAYS',       TOBTiers.GetValue('T_PAYS'));
        TobAdr.PutValue('GPA_INCOTERM',   TobTiers.GetValue('YTC_INCOTERM'));
        TobAdr.PutValue('GPA_MODEEXP',    TobTiers.GetValue('YTC_MODEEXP'));
        TobAdr.PutValue('GPA_LIEUDISPO',  TobTiers.GetValue('YTC_LIEUDISPO'));
        TobAdr.PutValue('GPA_EAN',        TobTiers.GetValue('T_EAN'));
        TobAdr.PutValue('GPA_NIF',        TobTiers.GetValue('T_NIF'));
        TobAdr.PutValue('GPA_REGION',     TobTiers.GetValue('T_REGION'));
        //
        TobAdr.PutValue('GPA_NUMEROCONTACT', wGetSqlFieldValue('C_NUMEROCONTACT', 'CONTACT', 'C_AUXILIAIRE = "' + TOBTiers.GetValue('T_AUXILIAIRE') + '" AND C_PRINCIPAL = "' + wTrue + '"'));
      end
      else
        PutInTobAdresse;
    end;
  end;

var TOBT: TOB;
  Q: TQuery;
  Req, chp: string;
begin
  if Auxi = '' then Exit;
  TOBT := TOB.Create('TIERS', nil, -1);
  chp := 'T_AUXILIAIRE, T_LIBELLE,T_PRENOM,T_JURIDIQUE,T_ADRESSE1,T_ADRESSE2,T_ADRESSE3,T_CODEPOSTAL,T_VILLE,T_PAYS,T_TELEPHONE,T_EAN,T_NIF,T_REGION'
    + ',YTC_INCOTERM,YTC_MODEEXP,YTC_LIEUDISPO';
  Req := 'SELECT ' + chp + ' FROM TIERS LEFT JOIN TIERSCOMPL ON T_TIERS=YTC_TIERS WHERE T_AUXILIAIRE="' + Auxi + '"';

  Q := OpenSQL(Req, True,-1, '', True);
  TOBT.SelectDB('', Q);
  Ferme(Q);
  GetAdrFromTOB(TOBAdr, TOBT);
  TOBT.Free;
end;


//FV1 : 12/12/2014 - FS#1350 - ACTUACOM - Non reprise de l'adresse de facturation en génération d'échéance
procedure TAFPrepFact.GetAdressesContrat(TOBPiece, TOBSADRESSES,TOBAdresses: TOB);
Var TOBT        : TOB;
    TOBADr      : TOB;
    //
    TiersSaisie : string;
    TiersFile   : string;
    TiersFact   : String;
    //
    AuxiSaisie  : string;
    AuxiFile    : string;
    //
    CodeTiers   : string;
    NumAdrFac   : Integer;
    //
    QQ          : TQuery;
    StSQl       : string;
    //
    I           : Integer;
begin

  TOBT    := TOB.Create('TIERS', nil, -1);

  CodeTiers := TOBPiece.GetString('GP_TIERS');

  TiersSaisie := TOBPiece.GetString('GP_TIERSFACTURE');

  //Lecture de la fiche tiers à partir du code tiers saisit
  StSQl := 'SELECT T_LIBELLE, T_PRENOM, T_JURIDIQUE ';
  StSQL := StSQL + ', T_ADRESSE1,T_ADRESSE2,T_ADRESSE3,T_CODEPOSTAL,T_VILLE,T_PAYS ';
  StSQl := StSQl + ', T_AUXILIAIRE,T_FACTURE, T_TELEPHONE, T_EAN, T_NIF, T_REGION ';
  StSQl := StSQl + ', YTC_INCOTERM,YTC_MODEEXP,YTC_LIEUDISPO, YTC_NADRESSEFAC ';
  StSQl := StSQL + 'FROM TIERS LEFT JOIN TIERSCOMPL ON T_TIERS=YTC_TIERS ';
  StSQL := StSQl + 'WHERE T_TIERS="' + CodeTiers + '"';

  QQ := OpenSQL(StSQl, False);

  if QQ.Eof then
  begin
    Ferme(QQ);
    Exit;
  end;
  //
  TOBT.SelectDB('', QQ);
  //
  AuxiFile  := TOBT.GetString('T_FACTURE');
  NumAdrFac := TOBT.GetInteger('YTC_NADRESSEFAC');

  //On recherche le compte auxiliaire du tiers Facturé de la saisie
  AuxiSaisie := RecupTiersAuxi('A', TiersSaisie);

  if AuxiSaisie <> AuxiFile then
  begin
    //Si le tiers de la saisie est égal au tiers du fichier
    If NumAdrFac = 0 then
    begin
      //Lecture de la fiche tiers à partir du code tiers saisit
      StSQl := 'SELECT T_LIBELLE, T_PRENOM, T_JURIDIQUE ';
      StSQL := StSQL + ', T_ADRESSE1,T_ADRESSE2,T_ADRESSE3,T_CODEPOSTAL,T_VILLE,T_PAYS ';
      StSQl := StSQl + ', T_AUXILIAIRE,T_FACTURE, T_TELEPHONE, T_EAN, T_NIF, T_REGION ';
      StSQl := StSQl + ', YTC_INCOTERM,YTC_MODEEXP,YTC_LIEUDISPO, YTC_NADRESSEFAC ';
      StSQl := StSQL + 'FROM TIERS LEFT JOIN TIERSCOMPL ON T_TIERS=YTC_TIERS ';
      StSQL := StSQl + 'WHERE T_AUXILIAIRE="' + AuxiSaisie + '"';
      //
      QQ := OpenSQL(StSQl, False);
      //
      if not QQ.Eof then
      begin
        TOBT.SelectDB('', QQ);
        //A partir de TOBAdresse on charge l'adresse de Facturation (????)
        //chargement de l'adresse de facturation si client différents de celui saisit
        for I := 0 To TOBAdresses.Detail.Count -1 do
        begin
          TOBADr := TOBAdresses.Detail[I];
          GetAdrFromTOB(TOBADr, TOBT);
        end;
      end;
      //
      Ferme(QQ);
      //
    end
    else
    begin
      TiersFile := RecupTiersAuxi('', AuxiFile);
      //Lecture de la fiche tiers à partir du code tiers saisit
      StSQl := 'SELECT ADR_TYPEADRESSE,ADR_LIBELLE, ADR_LIBELLE, ADR_LIBELLE2 ';
      StSQL := StSQL + ', ADR_ADRESSE1, ADR_ADRESSE2, ADR_ADRESSE3, ADR_CODEPOSTAL, ADR_VILLE, ADR_PAYS ';
      StSQl := StSQl + ', ADR_TELEPHONE, ADR_EAN, ADR_NIF, ADR_REGION, ADR_NUMEROCONTACT ';
      StSQl := StSQL + 'FROM ADRESSES WHERE ADR_NATUREAUXI="CLI" ';
      StSQl := StSQl + ' AND ADR_REFCODE="' + TiersFile + '"';
      StSQL := StSQL + ' AND ADR_NADRESSE=' + IntToStr(NumAdrFac);
      StSQl := StSQl + ' AND ADR_FACT="X"';
      //
      QQ := OpenSQL(StSQl, False);
      //
      if not QQ.Eof then
      begin
        TOBT.SelectDB('', QQ);
        //A partir de TOBAdresse on charge l'adresse de Facturation (????)
        //chargement de l'adresse de facturation si client différents de celui saisit
        for I := 0 To TOBAdresses.Detail.Count -1 do
        begin
          TOBADr := TOBAdresses.Detail[I];
          GetTobPieceAdresseFromTobAdresses(TOBADr, TOBT);
        end;
      end;
      //
      Ferme(QQ);
      //
    End;
  end;

end;

function TAFPrepFact.RecupTiersAuxi(TypeCpt, CptTiers : String) : string;
Var StSQL : string;
    QQ    : TQuery;
begin

  Result := '';

  If TypeCpt = 'A' then
    StSQl := 'SELECT  T_AUXILIAIRE FROM TIERS WHERE T_TIERS="' + CptTiers + '"'
  else
    StSQl := 'SELECT  T_TIERS FROM TIERS WHERE T_AUXILIAIRE="' + CptTiers + '"';

  QQ := OpenSQL(StSQl, False);

  if Not QQ.Eof then
  begin
    Result := QQ.Fields[0].AsString;
  end;

  Ferme (QQ);

end;


procedure TAFPrepFact.GetAdresses(TOBPiece, TOBSADRESSES,TOBAdresses: TOB);

  function GetTiersFacAffaire(TOBPiece: TOB) : string;
  var QQ : TQuery;
  begin
    Result := '';
    QQ := OpenSQL('SELECT AFF_FACTURE FROM AFFAIRE WHERE AFF_AFFAIRE="'+TOBPiece.GetString('GP_AFFAIRE')+'"',True,1,'',True);
    if not QQ.Eof then
    begin
      Result := QQ.Fields [0].AsString;
    end;
    ferme (QQ);
  end;

var TOBAA     : TOB;
    TOBBB     : TOB;
    CodeTiers : string;
    TiersCpta : string;
    ii : Integer;
begin

  TOBAdresses.ClearDetail;

  CodeTiers:=TOBPiece.GetValue('GP_TIERS') ;
  TiersCpta:=GetTiersFacAffaire(TOBPiece) ;

	TOBAA := TOBSADRESSES.FindFirst(['ORIGINE'],[TOBPiece.GetInteger('GP_NUMERO')],true);

  if TOBAA <> nil then
  begin
    for ii := 0 to TOBAA.detail.count -1 do
    begin
      TOBBB := TOB.Create('PIECEADRESSE', TOBAdresses, -1);
      if (TOBAA.detail[ii].GetString('GPA_TYPEPIECEADR')='002') and
         ((CodeTiers<>TiersCpta) and (TiersCpta<>'')) then
      begin
        GetAdrFromAuxi (TOBBB,TiersCpta);
        TOBBB.SetString('GPA_TYPEPIECEADR','002');
      end else
      begin
        TOBBB.Dupliquer(TOBAA.detail[ii],True,true);
  end;
    end;
  end;

end;

procedure TAFPrepFact.GetMEMOS(TOBPiece, TOBSOLES, TOBLienOle: TOB);
var TOBAA : TOB;
begin
  TOBLienOLE.ClearDetail;
	TOBAA := TOBSOLES.FindFirst(['ORIGINE'],[TOBPiece.GetInteger('GP_NUMERO')],true);
  if TOBAA <> nil then
  begin
    TOBLienOle.Dupliquer(TOBAA,True,true);
  end;

end;

end.



