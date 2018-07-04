unit Tarifs;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Marc MORRETTON
Créé le ...... : 18/06/2002
Modifié le ... : 18/06/2002
Description .. : Système Tarifaire
Mots clefs ... :
*****************************************************************}

interface

uses {$IFDEF VER150} variants,{$ENDIF}
  StdCtrls,
  Classes,
  SysUtils,
  Dialogs,
  Graphics,
  Windows,
  extctrls,
  Controls, //pour mrYes
  {$IFNDEF EAGLCLIENT}
    {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
    db,
    fe_main,
  {$ELSE}
    mainEagl,
  {$ENDIF}
  hCtrls,
  hRichOle,
  Forms,
  hEnt1,
  hMsgBox,
  formule,
  uTob,
  uTom,
  Menus,
  EntGC,
  SaisUtil
  ;
  { Chargement des paramètres de la table YTARIFSFONCT }
  function  GetTobYFO(sFonctionnalite: string; TheTob: Tob): Boolean;
  function  GetParametresFromYFO(sFonctionnalite: string): string; overload;
	procedure GetParametresFromYFO(sFonctionnalite, sCodePort, sOrientation: string; var sParametreYFOActif, sParametreYFOSelection, sParametreYFOEntete, sParametreYFOCLef: string); overload

  { Chargement des SurQuoi d'un Aqui !!!}
  function  GetTobYTQ(iIdentifiant:Integer; TheTob: Tob): Boolean;
  procedure DuplicationYTQ(sFonctionnalite, sTiers, sFamilleTiers, sAffaire, sFamilleAffaire : string; iDeIdentifiantYTA, iAIdentifiantYTA : integer; sDeDeviseYTA, sADeviseYTA : string);

  { Chargement des Fourchettes d'un SurQuoi !!!}
  function  GetTobYTF(iIdentifiantYTA, iIdentifiantYTQ :Integer; TheTob: Tob): Boolean;
  procedure DuplicationYTF(sFonctionnalite : string; iDeIdentifiantYTA, iDeIdentifiantYTQ, iAIdentifiantYTA, iAIdentifiantYTQ : integer; nTauxMonnaie : double; sDevise : string);

  { Controle d'existance de la clef naturelle pour message d'alerte }
  function  ExistYTA(sWhere : string):Boolean;
  function  ExistYTQ(sWhere : string):Boolean;

  { Calcul du poids du contexte de AQui }
  function  CalculPoidsAQui(TobYTA : tob) : integer;
  { Calcul du poids du contexte de AQui }
  function  CalculPoidsSurQuoi(TobYTQ : tob; iPartiel : integer) : integer;

  { Test Existance du Tarif Spécial }
  function  ExistTarifSpecial(sFonctionnalite, sCodeTarifSpecial : string):Boolean;
  function  WhereTarifSpecial(sFonctionnalite, sCodeTarifSpecial : string): string;

  { Test Existance des fourchettes quantitatives sur un Tarif }
  function  ExistTarifsFourchette(iIdentifiantYTA, iIdentifiantYTQ : integer):Boolean;

  { Affichage Résultat de la recherche tarifaire : Valeur + Origine }
  procedure AffichageOrigineSystemeTarifaire(sAppel: string; TobContexte, TobTarifs, TobAffichage: tob; GAffichage : tHGrid; MBLOCNOTE : tMemo);
  procedure AffichageCommentaireTarif(sParametre : string; TobContexte, TobAffichage : tob; GAffichage: thGrid;  MBLOCNOTE : tMemo);

  { Donne les droits pour Action}
  function  DroitAction(sDroit :string) : string;

  { Donne la nature d'auxiliaire en fonction de la fonctionnalité demandée }
  function  NatureAuxiliaire(sFonctionnalite :string) : string;
  { Recherche code fonctionnalité en fonction de la nature de pièce }
  function  RechercheFonctionnalite(sNaturePiece: string): string;
  { Construction en automatique du libellé de yTarifsSurQuoi }
  function  SetTarifLibelle(sFonctionnalite, sTiers, sFamilleTiers, sAffaire, sFamilleAffaire, sArticle, sFamilleArticle, sRessource, sFamilleRessource :string) : string;
  { Verification des conditions particulières d'application tarifaire }
  function  VerifConditionsApplications(sAppel, sConditionsApplications : string; tobtiers, TobArticle, TobPiece, TobLigne : Tob) : Boolean ;
  { Renvoie le libellé de la colonne en fonction de la fonctionnalité }
  function  LibelleParFonctionnalite(sFonctionnalite, sQuoi : string) : string;
  { Renvoie la donnée à cumuler en foncion du paramètre YTQ_FOURCHETTEDE }
  function  Quoi(sParam: string; TobContexte: tob): Double;
  { Construction du WHERE selon une combinaison des informations du contexte }
  function  ContexteAqui(sRecherche, sParametre: string; sDepot, sTarifDepot, sDevise, sTarifSpecial, sCommercial, sTypeCommercial, sTiers, sTarifTiers, sAffaire, sTarifAffaire, sPole, sAtelier, sGroupe, sJourSemaine, sHeureJour, sLieuConso : string) : string;
  { Construction du WHERE selon une combinaison des informations du contexte }
  function  ContexteSurQuoi(sRecherche, sParametre : string; TobContexte: tob; iPartiel : integer) : string;

  { Calcul du Prix Net }
  function  CalculPrixNet(nPrixBrut, nRemise1, nRemise2, nRemise3 : double; iSensTaux: integer; sMonnaie, sModule, sQuoi, sMethode : string) : double;
  { Calcul de la Valeur Totale d'un contexte }
  function  CalculValeurLigne(nForfait, nCoutLancement, nRemiseEnMontant, nPrixNet, nQte : double; sMonnaie, sModule, sQuoi, sMethode : string) : double;
  { Donne la nature d'auxiliaire en fonction de la fonctionnalité demandée }
  function  SensTauxOuRemise(sFonctionnalite :string) : integer;

  { Charge les données d'une devise }
  function  wGetTobDevise(sChamp, sWhere: String; TobDevise: Tob; SelectDB:Boolean = False): Boolean;
  { Convertion d'un montant exprimé dans une monnaie en une autre monnaie }
  function  ConvMonnaie(sMonnaieOrigine, sMonnaieDestination :String; nValeur: Double; lFixing: boolean; sModule, sQuoi, sMethode: String): Double;
  { Arrondir un Prix en focntion de la monnaie }
  function  ArrondirPrixMontant(nValeur : double; sMonnaie, sModule, sQuoi, sMethode: string): double;

  { Appel du système de recherche tarifaire depuis l'écran de recherche }
  procedure AppelTarifsRecherche(sFonctionnalite, sRang, sAppel: String);

  { Fonction de recherche du système tarifaire : renvoi le résultat dans TobTarifs }
  function  RechercheTarifs(sAppel : string; TobContexte, TobTarifs, TobTiers, TobArticle, TobPiece, TobLigne : Tob; lComplement : boolean) : Tob;

  { Fonction de recherche du système tarifaire dans un contexte groupé commande }
  function  CalculTarifGroupePiece(TobTiers, TobArticles, TobPiece, TobLigneTarif : tob) : boolean;

  { }
  procedure TobTarifsAddChampsSupValeur( TobTarifs: Tob);
  function  CreationTobRemises( TobTarifs: Tob): Tob;

  { Fonctions appelées depuis les pièces }

  { Appel du système de recherche des tarifs et des commissionnements depuis les pièces }
  procedure RechercheTarifsCommissions(sAppel: string; Action: TActionFiche; TobTiers, TobArticle, TobPiece, TobLigne, TobLigneTarif, TobTarif : tob; Devise: rDevise);
  { Appel du système de recherche tarifaire depuis les pièces }
  function  RechercheEnregistreTarifs(sFonctionnalite, sAppel, sRang: string; TobTiers, TobArticle, TobPiece, TobLigne, TobLigneTarif : tob; var nMontantFixe: double; var nPrixLigne : double; var nRemiseTotale : double; var nRemiseEnMontant : double; var sMonnaiePiece: string): boolean;
  { Mise à jour de la TobContexte, pour affichage des origines des tarifs }
  Procedure GetTobContexteFromPiece(sFonctionnalite, sRang: string; TobContexte, TobPiece, TobLigne: Tob);
  { enregistrement du système tarifiare dans les lignes de pièces }
  procedure EnregistreTarifs(sFonctionnalite, sAppel, sRang: string; TobContexte, TobTarifs, TobLigne, TobLigneTarif: tob; var nMontantFixe : double; var nPrixLigne : double; var nRemiseTotale : double; var nRemiseEnMontant : double; var sMonnaiePiece : string);
  { Appel de la justification tarifaire}
  procedure ConsultationTarifsRecherche(sFonctionnalite, sRang: string; TobUneLigneTarif, TobContexte: Tob);

  { Création de la TobLigneTarif pour une ligne et pour une donnée }
  function  CreationTobLigneTarif( TobLigne, TobLigneTarif:tob; sFonctionnalite: string; sRang:string; sQuoi:string; nValeur:double; sOu:string): Tob;
  { Annulation de la TobLigneTarif pour une ligne et pour une donnée }
  function  AnnulationTobLigneTarif(TobLigne, TobLigneTarif:tob; sFonctionnalite: string; sRang: string) :boolean;
  { Mise à jour de la tobLigneTarif en fonction des mofidifications manuelles sur TobLigne }
  Procedure MajTobLigneTarifFromTobLigne( TobLigne, TobLigneTarif:tob; sFonctionnalite, sRang, sQuoi:string; nValeur:double);
  { Calcul de la remise totale système sur TobLigneTarif en % }
  Function  CalculRemiseTotaleSurTobLigneTarif( TobLigne, TobLigneTarif: tob; sRang : string) : double;

  { Test l'existence d'une affaire  à intégrer ou utiliser une fonction équivalente  }
  function  ExistAffaire(sAffaire : string; WithAlert: Boolean = false): Boolean;

  function GetNomChampsFromPopMenu(lTestActif: boolean; sParametreYFOActif: string): string;
	function Remplace(sOrigine, sCode: string; iPos: integer) : string;
	function GetParametresFromYFOPParParametre(TobYFO : tob; sQuelParametre : string) : string;

Var
	TransfertTobs: TList;

Const
  sFournisseur  = 'FOU'; //Nature Auxiliaire
  sClient       = 'CLI'; //Nature Auxiliaire

  sOu_L = 'L';
  sOu_P = 'P';

  sTarifFournisseur       = '101';   //Fonctionnalité : tarif fournisseur
  sCommissionFournisseur  = '102';   //Fonctionnalité : commissionnement fournisseur
  sIndirectFournisseur    = '103';   //Fonctionnalité : indirect fournisseur

  sTarifClient            = '201';   //Fonctionnalité : tarif client
  sCommissionClient       = '202';   //Fonctionnalité : commissionnement client
  sIndirectClient         = '203';   //Fonctionnalité : indirect client

  sTarifSousTraitantAchat = '301';   //Fonctionnalité : tarif sous-traitant d'achat

  sTarifSousTraitantPhase = '401';   //Fonctionnalité : tarif sous-traitant de phase

  iPosOrigineDepot             =  1;
  iPosOrigineTarifDepot        =  2;
  iPosOrigineDevise            =  3;
  iPosOrigineSpecial           =  4;
  iPosOrigineCommercial        =  5;
  iPosOrigineTypeCommercial    =  6;
  iPosOrigineTiers             =  7;
  iPosOrigineTarifTiers        =  8;
  iPosOrigineAffaire           =  9;
  iPosOrigineTarifAffaire      = 10;
  iPosOriginePole              = 11;
  iPosOrigineAtelier           = 12;
  iPosOrigineGroupe            = 13;
  iPosOrigineJourSemaine       = 14;
  iPosOrigineHeureJour         = 15;
  iPosOrigineLieuConso         = 16;

  iPosOrigineArticle           = 20;
  iPosOrigineTarifArticle      = 21;
  iPosOrigineRessource         = 22;
  iPosOrigineTarifRessource    = 23;
  iPosOrigineMarque            = 24;
  iPosOrigineCircuit           = 25;
  iPosOriginePhase             = 26;
  iPosOrigineTraitement        = 27;

  iPosOrigineDateDeReference   = 30;
  iPosOrigineTTCouHT           = 31;
  iPosOrigineRechercherOu      = 32;
  iPosOrigineSommerOu          = 33;
  iPosOrigineSommerQuoi        = 34;

  iTailleOu                    = 35;


//

  iPosDepot               =  1;
  iPosTarifDepot          =  2;
  iPosDevise              =  3;
  iPosSpecial             =  4;
  iPosModeExp             =  5;
  iPosSecteurGeo          =  6;
  iPosTiers               =  7;
  iPosTarifTiers          =  8;
  iPosAffaire             =  9;
  iPosTarifAffaire        = 10;
  iPosAtelier             = 12;
  iPosCodePostal          = 13;
  iPosPays                = 14;
  iPosRegion              = 15;
  iPosDatesDeValidite     = 17;
  iPosCodePort            = 18;
  iPosNatureTra           = 19;

  iPosArticle             = 20;
  iPosTarifArticle        = 21;
  iPosFamilleNiv1         = 22;
  iPosFamilleNiv2         = 23;
  iPosFamilleNiv3         = 24;
  iPosRessource           = 25;
  iPosTarifRessource      = 26;
  iPosCircuit             = 28;
  iPosPhase               = 29;
  iPosTraitement          = 30;

  iPosDateDeReference     = 33;
  iPosTTCouHT             = 34;
  iPosRechercherOu        = 35;
  iPosSommerOu            = 36;
  iPosSommerQuoi          = 37;
  iPosFamilleHierarchique = 39;
  iPosCascadeParEntite    = 40;
  iPosLibelle             = 41;
  iPosRemiseTiers         = 42;
  iPosCascContexte        = 43;
  iPosParametres          = 44;
  iPosCondReelle          = 45;
  iPosActif               = 46;
  iPosPrixVenteMarge      = 47;
  iPosTarifMarche         = 48;

  iPosForfait             = 50;
  iPosCoutLancement       = 51;
  iPosRemiseEnMontant     = 52;
  iPosPrixBrut            = 53;
  iPosRemise1             = 54;
  iPosRemise2             = 55;
  iPosRemise3             = 56;
  iPosPrixNet             = 57;

  iPosMTarif              = 60;
  iPosMRemiseLibre        = 61;
  iPosMremiseLigne        = 62;
  iPosMRemisePied         = 63;
  iPosMFixe               = 64;
  iPosMRemiseMt           = 65;
//
  sOrigineMasque               = '................---........--.....-';

  sImageVide                = 'IMAGE_VIDE';
  sImageMain                = 'IMAGE_MAIN';
  sImageSommePleine         = 'IMAGE_SOMMEPLEINE';
  sImageSommeVide           = 'IMAGE_SOMMEVIDE';
  sImageJumellePleine       = 'IMAGE_JUMELLEPLEINE';
  sImageJumelleVide         = 'IMAGE_JUMELLEVIDE';
  sImageSommeJumellePleine  = 'IMAGE_SOMMEJUMELLEPLEINE';
  sImageSommeJumelleVide    = 'IMAGE_SOMMEJUMELLEVIDE';

implementation

uses
//  EntGP,
	wCommuns,
  Math,
  wParam,
  UtilArticle,
  wJetons,
  ParamSoc,
  FactTarifs,
  FactUtil,
  ED_TOOLS
//  wTobDebug_Tof
  ;


function GetNomChampsFromPopMenu(lTestActif: boolean; sParametreYFOActif: string): string;
begin {GetNomChampsFromPopMenu}
  if (not lTestActif) or (sParametreYFOActif[iPosCodePort]      =wTrue) then Result := Result + 'YTS_CODEPORT;'      ;
  if (not lTestActif) or (sParametreYFOActif[iPosSpecial]       =wTrue) then Result := Result + 'YTS_TARIFSPECIAL;'  ;
  if (not lTestActif) or (sParametreYFOActif[iPosTiers]         =wTrue) then Result := Result + 'YTS_TIERS;'         ;
  if (not lTestActif) or (sParametreYFOActif[iPosTarifTiers]    =wTrue) then Result := Result + 'YTS_TARIFTIERS;'    ;
  if (not lTestActif) or (sParametreYFOActif[iPosSecteurGeo]    =wTrue) then Result := Result + 'YTS_SECTEURGEO;'    ;
  if (not lTestActif) or (sParametreYFOActif[iPosCodePostal]    =wTrue) then Result := Result + 'YTS_CODEPOSTAL;'    ;
  if (not lTestActif) or (sParametreYFOActif[iPosPays]          =wTrue) then Result := Result + 'YTS_PAYS;'          ;
  if (not lTestActif) or (sParametreYFOActif[iPosRegion]        =wTrue) then Result := Result + 'YTS_REGION;'        ;
  if (not lTestActif) or (sParametreYFOActif[iPosModeExp]       =wTrue) then Result := Result + 'YTS_MODEEXP;'       ;
  if (not lTestActif) or (sParametreYFOActif[iPosNatureTra]     =wTrue) then Result := Result + 'YTS_NATURETRAVAIL;' ;
  if (not lTestActif) or (sParametreYFOActif[iPosArticle]       =wTrue) then Result := Result + 'YTS_ARTICLE;'       ;
  if (not lTestActif) or (sParametreYFOActif[iPosTarifArticle]  =wTrue) then Result := Result + 'YTS_TARIFARTICLE;'  ;
  if (not lTestActif) or (sParametreYFOActif[iPosFamilleNiv1]   =wTrue) then Result := Result + 'YTS_FAMILLENIV1;'   ;
  if (not lTestActif) or (sParametreYFOActif[iPosFamilleNiv2]   =wTrue) then Result := Result + 'YTS_FAMILLENIV2;'   ;
  if (not lTestActif) or (sParametreYFOActif[iPosFamilleNiv3]   =wTrue) then Result := Result + 'YTS_FAMILLENIV3;'   ;
  if (not lTestActif) or (sParametreYFOActif[iPosRessource]     =wTrue) then Result := Result + 'YTS_RESSOURCE;'     ;
  if (not lTestActif) or (sParametreYFOActif[iPosTarifRessource]=wTrue) then Result := Result + 'YTS_TARIFRESSOURCE;';
end; {GetNomChampsFromPopMenu}

{-------------------------------------------------------------------------------
   Chargement de la Tob des Tarifs d'un AQui par rapport au N° d'identifiant
-------------------------------------------------------------------------------}
function GetTobYTQ(iIdentifiant:Integer; TheTob: Tob): Boolean;
var
	sRequete: String;
begin
  Result := False;
  if TheTob <> nil then
  begin
    sRequete := 'SELECT * FROM YTARIFSSURQUOI WHERE YTQ_IDENTIFIANTYTA=' + IntToStr(iIdentifiant);
		Result   := TheTob.LoadDetailDBFromSql('YTARIFSSURQUOI', sRequete);
  end;
end;

{-------------------------------------------------------------------------------
   Duplication des SurQuoi pour un AQui
-------------------------------------------------------------------------------}
procedure DuplicationYTQ(sFonctionnalite, sTiers, sFamilleTiers, sAffaire, sFamilleAffaire : string; iDeIdentifiantYTA, iAIdentifiantYTA : integer; sDeDeviseYTA, sADeviseYTA : string);
var
  TobYTQ  : tob;
  TomYTQ  : Tom;
  iCpt    : integer;
  iDeIdentifiantYTQ, iAIdentifiantYTQ : integer;
  sNatureAuxi   : string;
  iSensTaux     : integer;
  nTauxMonnaie  : double;
begin
  TobYTQ := Tob.Create('YTQ', nil, -1);
  TomYTQ := CreateTOM('YTARIFSSURQUOI', nil, false, true);
  try

    { Si changement de devise alors conversion en duplication de tarif }
    sNatureAuxi   := NatureAuxiliaire(sFonctionnalite);
    iSensTaux     := SensTauxOuRemise(sFonctionnalite);
    if (sDeDeviseYTA<>sADeviseYTA) then
      nTauxMonnaie  := ConvMonnaie(sDeDeviseYTA, sADeviseYTA, 1, True, sNatureAuxi, 'T', ' ')
    else
      nTauxMonnaie := 1;

    //Chargement de la Tob sur les lignes de Sur Quoi
    if GetTobYTQ(iDeIdentifiantYTA, TobYTQ) then
    begin
      InitMoveProgressForm(nil, TraduireMemoire('Duplication des tarifs'), '', TobYTQ.Detail.Count, False, True);
      for iCpt:=0 to (TobYTQ.Detail.Count-1) do
      begin
        if (TobYTQ.Detail[iCpt].GetValue('YTQ_CODEARTICLE')<>'') then
    	    MoveCurProgressForm(TraduireMemoire('Article')+' : '+TobYTQ.Detail[iCpt].GetValue('YTQ_CODEARTICLE'))
        else
    	    MoveCurProgressForm(TraduireMemoire('Famille')+' : '+TobYTQ.Detail[iCpt].GetValue('YTQ_TARIFARTICLE'));
        iDeIdentifiantYTQ := TobYTQ.Detail[iCpt].GetValue('YTQ_IDENTIFIANT');

        TobYTQ.Detail[iCpt].PutValue('YTQ_IDENTIFIANTYTA',iAIdentifiantYTA);
        TobYTQ.Detail[iCpt].PutValue('YTQ_IDENTIFIANT',wSetJeton('YTQ'));
        iAIdentifiantYTQ  := TobYTQ.Detail[iCpt].GetValue('YTQ_IDENTIFIANT');

        TobYTQ.Detail[iCpt].PutValue('YTQ_LIBELLETARIF', SetTarifLibelle(sFonctionnalite, sTiers, sFamilleTiers, sAffaire, sFamilleAffaire, TobYTQ.Detail[iCpt].GetValue('YTQ_CODEARTICLE'), TobYTQ.Detail[iCpt].GetValue('YTQ_TARIFARTICLE'), TobYTQ.Detail[iCpt].GetValue('YTQ_RESSOURCE'), TobYTQ.Detail[iCpt].GetValue('YTQ_TARIFRESSOURCE')));

        { Si changement de devise alors conversion en duplication de tarif }
        if (nTauxMonnaie<>1) then
        begin
          TobYTQ.Detail[iCpt].PutValue('YTQ_FORFAIT'      ,ArrondirPrixMontant(TobYTQ.Detail[iCpt].GetValue('YTQ_FORFAIT')      *nTauxMonnaie, sADeviseYTA ,sNatureAuxi ,'M' ,'P'));
          TobYTQ.Detail[iCpt].PutValue('YTQ_COUTLANCEMENT',ArrondirPrixMontant(TobYTQ.Detail[iCpt].GetValue('YTQ_COUTLANCEMENT')*nTauxMonnaie, sADeviseYTA ,sNatureAuxi ,'M' ,'P'));
          TobYTQ.Detail[iCpt].PutValue('YTQ_PRIXBRUT'     ,ArrondirPrixMontant(TobYTQ.Detail[iCpt].GetValue('YTQ_PRIXBRUT')     *nTauxMonnaie, sADeviseYTA ,sNatureAuxi ,'P' ,'P'));
          if (TobYTQ.Detail[iCpt].GetValue('YTQ_PRIXBRUT')<>0) then
            TobYTQ.Detail[iCpt].PutValue('YTQ_PRIXNET',CalculPrixNet(TobYTQ.Detail[iCpt].GetValue('YTQ_PRIXBRUT'), TobYTQ.Detail[iCpt].GetValue('YTQ_REMISE1'), TobYTQ.Detail[iCpt].GetValue('YTQ_REMISE2'), TobYTQ.Detail[iCpt].GetValue('YTQ_REMISE3'), iSensTaux, sADeviseYTA, sNatureAuxi, 'P', 'T'))
          else
            TobYTQ.Detail[iCpt].PutValue('YTQ_PRIXNET',ArrondirPrixMontant(TobYTQ.Detail[iCpt].GetValue('YTQ_PRIXNET')*nTauxMonnaie, sADeviseYTA ,sNatureAuxi ,'P' ,'P'));
        end;

        { UpDateRecord }
        TobYTQ.Detail[iCpt].AddChampSupValeur('IKC', 'C', false);
        if TomYTQ.VerifTOB(TobYTQ.Detail[iCpt]) then
          DuplicationYTF(sFonctionnalite, iDeIdentifiantYTA, iDeIdentifiantYTQ, iAIdentifiantYTA, iAIdentifiantYTQ, nTauxMonnaie, sADeviseYTA);
      end;
      FiniMoveProgressForm;
//      wShowMeTheTob(TobYTQ, True);

      //Enregistrement de cette Tob
      TobYTQ.SetAllModifie(True);    // Obligatoire pour que le Memo se duplique
      TobYTQ.InsertDB(nil, true);
    end;
  finally
    TobYTQ.Free;
    TomYTQ.Free;
  end;
end;

{-------------------------------------------------------------------------------
   Chargement de la Tob des Tarifs d'un SurQuoi par rapport au N° d'identifiant
-------------------------------------------------------------------------------}
function GetTobYTF(iIdentifiantYTA, iIdentifiantYTQ :Integer; TheTob: Tob): Boolean;
var
	sRequete: String;
begin
  Result := False;
  if TheTob <> nil then
  begin
    sRequete := 'SELECT * FROM YTARIFSFOURCHETTE WHERE (YTF_IDENTIFIANTYTA=' + IntToStr(iIdentifiantYTA) + ') and (YTF_IDENTIFIANTYTQ=' + IntToStr(iIdentifiantYTQ)+')';
    Result := TheTob.LoadDetailDBFromSql('YTARIFSFOURCHETTE', sRequete);
  end;
end;

{-------------------------------------------------------------------------------
   Duplication des Fourchettes pour un SurQuoi
-------------------------------------------------------------------------------}
procedure DuplicationYTF(sFonctionnalite : string; iDeIdentifiantYTA, iDeIdentifiantYTQ, iAIdentifiantYTA, iAIdentifiantYTQ : integer; nTauxMonnaie : double; sDevise : string);
var
  TobYTF   : tob;
  TomYTF   : Tom;
  iCpt     : integer;
  sNatureAuxi   : string;
  iSensTaux     : integer;
begin
  TobYTF := Tob.Create('YTF', nil, -1);
  TomYTF := CreateTOM('YTARIFSFOURCHETTE', nil, false, true);
  try
    sNatureAuxi   := NatureAuxiliaire(sFonctionnalite);
    iSensTaux     := SensTauxOuRemise(sFonctionnalite);

    //Chargement de la Tob sur les lignes de Fourchette
    if GetTobYTF(iDeIdentifiantYTA, iDeIdentifiantYTQ, TobYTF) then
    begin
      InitMoveProgressForm(nil, TraduireMemoire('Duplication des tarifs'), '', TobYTF.Detail.Count, False, True);
      for iCpt:=0 to (TobYTF.Detail.Count-1) do
      begin
  	    MoveCurProgressForm(TraduireMemoire('De')+' : '+FloatToStr(TobYTF.Detail[iCpt].GetValue('YTF_FOURCHETTEMINI'))+'    ----    '+TraduireMemoire('A')+' : '+FloatToStr(TobYTF.Detail[iCpt].GetValue('YTF_FOURCHETTEMAXI')));

        TobYTF.Detail[iCpt].PutValue('YTF_IDENTIFIANTYTA',iAIdentifiantYTA);
        TobYTF.Detail[iCpt].PutValue('YTF_IDENTIFIANTYTQ',iAIdentifiantYTQ);
        TobYTF.Detail[iCpt].PutValue('YTF_IDENTIFIANT'   ,wSetJeton('YTF'));

        { Si changement de devise alors conversion en duplication de tarif }
        if (nTauxMonnaie<>1) then
        begin
          TobYTF.Detail[iCpt].PutValue('YTF_FORFAIT'      ,ArrondirPrixMontant(TobYTF.Detail[iCpt].GetValue('YTF_FORFAIT')      *nTauxMonnaie, sDevise, sNatureAuxi ,'M' ,'P'));
          TobYTF.Detail[iCpt].PutValue('YTF_COUTLANCEMENT',ArrondirPrixMontant(TobYTF.Detail[iCpt].GetValue('YTF_COUTLANCEMENT')*nTauxMonnaie, sDevise, sNatureAuxi ,'M' ,'P'));
          TobYTF.Detail[iCpt].PutValue('YTF_PRIXBRUT'     ,ArrondirPrixMontant(TobYTF.Detail[iCpt].GetValue('YTF_PRIXBRUT')     *nTauxMonnaie, sDevise, sNatureAuxi ,'P' ,'P'));
          if (TobYTF.Detail[iCpt].GetValue('YTF_PRIXBRUT')<>0) then
            TobYTF.Detail[iCpt].PutValue('YTF_PRIXNET',CalculPrixNet(TobYTF.Detail[iCpt].GetValue('YTF_PRIXBRUT'), TobYTF.Detail[iCpt].GetValue('YTF_REMISE1'), TobYTF.Detail[iCpt].GetValue('YTF_REMISE2'), TobYTF.Detail[iCpt].GetValue('YTF_REMISE3'), iSensTaux, sDevise, sNatureAuxi, 'P', 'T'))
          else
            TobYTF.Detail[iCpt].PutValue('YTF_PRIXNET',ArrondirPrixMontant(TobYTF.Detail[iCpt].GetValue('YTF_PRIXNET')*nTauxMonnaie, sDevise ,sNatureAuxi ,'P' ,'P'));
        end;

        { UpDateRecord }
        TobYTF.Detail[iCpt].AddChampSupValeur('IKC', 'C', false);
        TomYTF.VerifTOB(TobYTF.Detail[iCpt]);
      end;
      FiniMoveProgressForm;
//      wShowMeTheTob(TobYTF, True);

      //Enregistrement de cette Tob
      TobYTF.SetAllModifie(True);    // Obligatoire pour que le Memo se duplique
      TobYTF.InsertDB(nil, true);
    end;
  finally
    TobYTF.Free;
    TomYTF.Free;
  end;
end;

{-------------------------------------------------------------------------------
   Calcul du poids du contexte de AQui
-------------------------------------------------------------------------------}
function CalculPoidsAQui(TobYTA : tob) : integer;
var
  iOffSet : integer;
begin
  iOffSet:= 1;
  Result := 0;
  Result := Result + iif(TobYTA.GetValue('YTA_LIEUCONSO')      <>''                , iOffSet, 0); iOffSet := iOffSet*2;
  Result := Result + iif(TobYTA.GetValue('YTA_HEUREJOUR')      <>''                , iOffSet, 0); iOffSet := iOffSet*2;
  Result := Result + iif(TobYTA.GetValue('YTA_JOURSEMAINE')    <>''                , iOffSet, 0); iOffSet := iOffSet*2;
  Result := Result + iif(TobYTA.GetValue('YTA_GRP')            <>''                , iOffSet, 0); iOffSet := iOffSet*2;
  Result := Result + iif(TobYTA.GetValue('YTA_SITE')           <>''                , iOffSet, 0); iOffSet := iOffSet*2;
  Result := Result + iif(TobYTA.GetValue('YTA_POLE')           <>''                , iOffSet, 0); iOffSet := iOffSet*2;
  Result := Result + iif(TobYTA.GetValue('YTA_TARIFAFFAIRE')   <>''                , iOffSet, 0); iOffSet := iOffSet*2;
  Result := Result + iif(TobYTA.GetValue('YTA_AFFAIRE')        <>''                , iOffSet, 0); iOffSet := iOffSet*2;
  Result := Result + iif(TobYTA.GetValue('YTA_TARIFTIERS')     <>''                , iOffSet, 0); iOffSet := iOffSet*2;
  Result := Result + iif(TobYTA.GetValue('YTA_TIERS')          <>''                , iOffSet, 0); iOffSet := iOffSet*2;
  Result := Result + iif(TobYTA.GetValue('YTA_TYPECOMMERCIAL') <> ''               , iOffSet, 0); iOffSet := iOffSet*2;
  Result := Result + iif(TobYTA.GetValue('YTA_COMMERCIAL')     <> ''               , iOffSet, 0); iOffSet := iOffSet*2;
  Result := Result + iif(TobYTA.GetValue('YTA_TARIFSPECIAL')   <> ''               , iOffSet, 0); iOffSet := iOffSet*2;
  Result := Result + iif(TobYTA.GetValue('YTA_DEVISE')         <> V_PGI.DevisePivot, iOffSet, 0); iOffSet := iOffSet*2;
  Result := Result + iif(TobYTA.GetValue('YTA_TARIFDEPOT')     <> ''               , iOffSet, 0); iOffSet := iOffSet*2;
  Result := Result + iif(TobYTA.GetValue('YTA_DEPOT')          <> ''               , iOffSet, 0);
end;

{-------------------------------------------------------------------------------
   Calcul du poids du contexte de AQui
-------------------------------------------------------------------------------}
function CalculPoidsSurQuoi(TobYTQ : tob; iPartiel : integer) : integer;
var
  iOffSet : integer;
begin
  iOffSet:= 1;
  Result := 0;
  Result := Result + iif(TobYTQ.GetValue('YTQ_TRAITEMENT')    <>''                                                             , iOffSet, 0); iOffSet := iOffSet*2;
  Result := Result + iif(TobYTQ.GetValue('YTQ_PHASE')         <>''                                                             , iOffSet, 0); iOffSet := iOffSet*2;
  Result := Result + iif(TobYTQ.GetValue('YTQ_CIRCUIT')       <>''                                                             , iOffSet, 0); iOffSet := iOffSet*2;
  Result := Result + iif(TobYTQ.GetValue('YTQ_MARQUE')        <>''                                                             , iOffSet, 0); iOffSet := iOffSet*2;
  Result := Result + iif(TobYTQ.GetValue('YTQ_TARIFARTICLE')  <>''                                                             , iOffSet, 0); iOffSet := iOffSet*2;
  Result := Result + iif((TobYTQ.GetValue('YTQ_CODEARTICLE')  <>'') and (Length(TobYTQ.GetValue('YTQ_CODEARTICLE'))=iPartiel)  , iOffSet, 0); iOffSet := iOffSet*2;
  Result := Result + iif((TobYTQ.GetValue('YTQ_CODEARTICLE')  <>'') and (Length(TobYTQ.GetValue('YTQ_CODEARTICLE'))<>iPartiel) , iOffSet, 0); iOffSet := iOffSet*2;
  Result := Result + iif(TobYTQ.GetValue('YTQ_TARIFRESSOURCE')<>''                                                             , iOffSet, 0); iOffSet := iOffSet*2;
  Result := Result + iif(TobYTQ.GetValue('YTQ_RESSOURCE')     <>''                                                             , iOffSet, 0);
end;

{-------------------------------------------------------------------------------
   Création de l'enregistrement dans WPARAM si manquant
   Lancement Fiche du Recherche d'un tarif selon un contexte prédéfini
--------------------------------------------------------------------------------}
Procedure AppelTarifsRecherche(sFonctionnalite, sRang, sAppel: String);
var
  sCodeFonction: String;
  Params       : String;
begin
  sCodeFonction:='TARIFSRECH : ' + sFonctionnalite;

  { La 1ère fois initialisation de certains paramètres par défaut }
  wIniWPA(sCodeFonction);

  Params := 'ACTION=MODIFICATION;MONOFICHE;YTA_FONCTIONNALITE='+sFonctionnalite+';RANG='+ sRang+ ';APPEL='+sAppel+';DROIT=...MV';

  AglLanceFiche('Y', 'YTARIFSRECH_FIC','', wGetLequelWPA(sCodeFonction), Params);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thierry Petetin
Créé le ...... : 07/10/2002
Description .. : Appel de la fiche de visualisation de l'origine des tarifs
Suite ........ : (Depuis la saisie des pièces)
*****************************************************************}
Procedure ConsultationTarifsRecherche(sFonctionnalite, sRang: string; TobUneLigneTarif, TobContexte: Tob);
var
	TobUneLigneTarifBis, TobContexteBis: Tob;
begin
  if (TobUneLigneTarif <> nil) and (TobContexte <> nil) then
  begin
    TobUneLigneTarifBis  := Tob.Create('TobUneLigneTarifBis', nil, -1);
    TobContexteBis       := Tob.Create('TobContexteBis', nil, -1);
    try
      { Copie les tobs }
      TobUneLigneTarifBis.Dupliquer(TobUneLigneTarif, True, True);
      TobContexteBis.Dupliquer(TobContexte, True, True);
      { Passe les Tobs dans la wParamTom }
      TransfertTobs.Clear;
      TransfertTobs.Add(TobContexteBis);
      TransfertTobs.Add(TobUneLigneTarifBis);
      { Appel de la fiche }
      AppelTarifsRecherche(sFonctionnalite, sRang, 'PIECE');
    finally
      TobUneLigneTarifBis.Free;
      TobContexteBis.Free;
      TransfertTobs.Clear;
    end;
  end;
end;

{--------------------------------------------------------------------------------*
   Teste existance du tarif spécial
---------------------------------------------------------------------------------}
function ExistTarifSpecial(sFonctionnalite, sCodeTarifSpecial : string):Boolean;
var
	sSql  : string;
begin
	sSql := ' SELECT YTP_TARIFSPECIAL'
			+ ' FROM YTARIFSPECIAL'
			+ ' WHERE ' + WhereTarifSpecial(sFonctionnalite,sCodeTarifSpecial)
         ;
	Result := ExisteSQL(sSql);
end;

{--------------------------------------------------------------------------------*
   Teste existance des fourchettes sur un tarif
---------------------------------------------------------------------------------}
function ExistTarifsFourchette(iIdentifiantYTA, iIdentifiantYTQ : integer):Boolean;
var
	sSql  : string;
begin
	sSql := ' SELECT YTF_IDENTIFIANT'
			+ ' FROM YTARIFSFOURCHETTE'
			+ ' WHERE YTF_IDENTIFIANTYTA = "' + IntToStr(iIdentifiantYTA) + '" and YTF_IDENTIFIANTYTQ = "' + IntToStr(iIdentifiantYTQ) + '"'
         ;
	Result := ExisteSQL(sSql);
end;

{--------------------------------------------------------------------------------*
   Construction du WHERE pour du Tarif Spécial
---------------------------------------------------------------------------------}
function WhereTarifSpecial(sFonctionnalite, sCodeTarifSpecial : string): string;
begin
	Result := 'YTP_FONCTIONNALITE = "' + sFonctionnalite + '" and '+'YTP_TARIFSPECIAL = "' + sCodeTarifSpecial + '"';
end;

{-------------------------------------------------------------------------------
   Chargement de la Tob des fonctionnalites
-------------------------------------------------------------------------------}
function GetTobYFO(sFonctionnalite:string; TheTob: Tob): Boolean;
var
	sRequete: String;
begin
  Result := False;
  if TheTob <> nil then
  begin
    sRequete := 'SELECT * FROM YTARIFSFONCTIONNALITE WHERE YFO_FONCTIONNALITE=' + sFonctionnalite;
    Result   := TheTob.LoadDetailDBFromSql('_TarifsFonctionnalite_', sRequete);
  end;
end;

{-------------------------------------------------------------------------------
   Va rechercher les paramètres de gestions des différentes fonctionnalites
-------------------------------------------------------------------------------}

procedure GetParametresFromYFO(sFonctionnalite, sCodePort, sOrientation: string; var sParametreYFOActif, sParametreYFOSelection, sParametreYFOEntete, sParametreYFOCLef: string);
var
  TobYFO     : tob;
  Sql        : string;
begin {GetParametresFromYFO}
  TobYFO := Tob.Create('_Fonctionnalite_' ,nil,-1);
  try
    Sql:= 'SELECT * FROM YTARIFSPARAMETRES WHERE (YFO_FONCTIONNALITE="'+sFonctionnalite+'"'
        + ' AND YFO_ORIENTATION="'+sOrientation+'"'
        + iif(sCodePort<>'', ' AND YFO_CODEPORT="'+ sCodePort+'"','')
        +')';
    if TobYFO.LoadDetailDBFromSQL('_TarifsParametres_', Sql) then
    begin
      sParametreYFOActif     := GetParametresFromYFOPParParametre (TobYFO, 'EntitéActive');
      sParametreYFOSelection := GetParametresFromYFOPParParametre(TobYFO, 'MulSélection');
      sParametreYFOEntete    := GetParametresFromYFOPParParametre(TobYFO, 'SaisieEntête');
      sParametreYFOClef      := GetParametresFromYFOPParParametre(TobYFO, 'ClefEntête'  );
    end;
  finally
    TobYFO.Free;
  end;
end; {GetParametresFromYFO}

function GetParametresFromYFO(sFonctionnalite:string): String;
var
  TobYFO     : tob;
  sParametre : string;
begin
  Result := sOrigineMasque;
  TobYFO := Tob.Create('_Fonctionnalite_' ,nil,-1);
  try
    if TobYFO.LoadDetailDBFromSql('_TarifsFonctionnalite_', 'SELECT * FROM YTARIFSFONCTIONNALITE WHERE YFO_FONCTIONNALITE=' + sFonctionnalite) then
    begin
      sParametre := sOrigineMasque;
      sParametre := Remplace(sParametre, TobYFO.Detail[0].GetValue('YFO_OKDEPOT')      , iPosOrigineDepot          );
      sParametre := Remplace(sParametre, TobYFO.Detail[0].GetValue('YFO_OKDEPOT')      , iPosOrigineTarifDepot     );
      sParametre := Remplace(sParametre, TobYFO.Detail[0].GetValue('YFO_OKDEVISE')     , iPosOrigineDevise         );
      sParametre := Remplace(sParametre, TobYFO.Detail[0].GetValue('YFO_OKSPECIAL')    , iPosOrigineSpecial        );
      sParametre := Remplace(sParametre, TobYFO.Detail[0].GetValue('YFO_OKCOMMERCIAL') , iPosOrigineCommercial     );
      sParametre := Remplace(sParametre, TobYFO.Detail[0].GetValue('YFO_OKCOMMERCIAL') , iPosOrigineTypeCommercial );
      sParametre := Remplace(sParametre, TobYFO.Detail[0].GetValue('YFO_OKTIERS')      , iPosOrigineTiers          );
      sParametre := Remplace(sParametre, TobYFO.Detail[0].GetValue('YFO_OKTIERS')      , iPosOrigineTarifTiers     );
      sParametre := Remplace(sParametre, TobYFO.Detail[0].GetValue('YFO_OKAFFAIRE')    , iPosOrigineAffaire        );
      sParametre := Remplace(sParametre, TobYFO.Detail[0].GetValue('YFO_OKAFFAIRE')    , iPosOrigineTarifAffaire   );
      sParametre := Remplace(sParametre, TobYFO.Detail[0].GetValue('YFO_OKPOLESITEGRP'), iPosOriginePole           );
      sParametre := Remplace(sParametre, TobYFO.Detail[0].GetValue('YFO_OKPOLESITEGRP'), iPosOrigineAtelier        );
      sParametre := Remplace(sParametre, TobYFO.Detail[0].GetValue('YFO_OKPOLESITEGRP'), iPosOrigineGroupe         );
      sParametre := Remplace(sParametre, TobYFO.Detail[0].GetValue('YFO_OKJHL')        , iPosOrigineJourSemaine    );
      sParametre := Remplace(sParametre, TobYFO.Detail[0].GetValue('YFO_OKJHL')        , iPosOrigineHeureJour      );
      sParametre := Remplace(sParametre, TobYFO.Detail[0].GetValue('YFO_OKJHL')        , iPosOrigineLieuConso      );

      sParametre := Remplace(sParametre, TobYFO.Detail[0].GetValue('YFO_OKARTICLE')    , iPosOrigineArticle        );
      sParametre := Remplace(sParametre, TobYFO.Detail[0].GetValue('YFO_OKARTICLE')    , iPosOrigineTarifArticle   );
      sParametre := Remplace(sParametre, TobYFO.Detail[0].GetValue('YFO_OKRESSOURCE')  , iPosOrigineRessource      );
      sParametre := Remplace(sParametre, TobYFO.Detail[0].GetValue('YFO_OKRESSOURCE')  , iPosOrigineTarifRessource );
      sParametre := Remplace(sParametre, TobYFO.Detail[0].GetValue('YFO_OKMARQUE')     , iPosOrigineMarque         );
      sParametre := Remplace(sParametre, TobYFO.Detail[0].GetValue('YFO_OKCPT')        , iPosOrigineCircuit        );
      sParametre := Remplace(sParametre, TobYFO.Detail[0].GetValue('YFO_OKCPT')        , iPosOriginePhase          );
      sParametre := Remplace(sParametre, TobYFO.Detail[0].GetValue('YFO_OKCPT')        , iPosOrigineTraitement     );
    end;
  finally
    TobYFO.Free;
  end;
  Result := sParametre;
end;

{-------------------------------------------------------------------------------
   Recherche Tarifaire
--------------------------------------------------------------------------------}
Function RechercheTarifs(sAppel : string; TobContexte, TobTarifs, TobTiers, TobArticle, TobPiece, TobLigne : Tob; lComplement : boolean) : Tob;
var
  iPartiel                            : integer;
  lConditionsApplications             : boolean;
  TobAQui, TobSurQuoi, TobFourchette  : Tob;
  sFonctionnalite, sLigneOuCommande   : string;
  sParametre                          : string;

  {----------------------------------------------------------------------------
    Ajout dans la TobContexte des informations provenant de la fiche Commercial
  -----------------------------------------------------------------------------}
  procedure AjoutComplementCommercial;
  var
    aCommercial : MyArrayValue;
  begin
    aCommercial := nil;

    if (TobContexte.GetValue('COMMERCIAL')<>'') then
    begin
      // Recherche des données de la fiche tiers
      aCommercial := WGetSqlFieldsValues(['GCL_TYPECOMMERCIAL'], 'COMMERCIAL', 'GCL_COMMERCIAL="'+TobContexte.GetValue('COMMERCIAL')+'"');
      if (aCommercial<>nil) then
        TobContexte.AddChampSupValeur('TYPECOMMERCIAL', aCommercial[0], false)
      else
        PgiInfo('Le code commercial est inconnu','La recherche sera incomplète');
    end
    else
      TobContexte.AddChampSupValeur('TYPECOMMERCIAL', '', false);
  end;

  {----------------------------------------------------------------------------
    Ajout dans la TobContexte des informations provenant de la fiche tiers
  -----------------------------------------------------------------------------}
  procedure AjoutComplementTiers;
  var
    aTiers : MyArrayValue;
  begin
    // Recherche des données de la fiche tiers
    aTiers := WGetSqlFieldsValues(['T_TARIFTIERS', 'T_FACTUREHT'], 'TIERS', 'T_NATUREAUXI="'+TobContexte.GetValue('NATUREAUXI')+'" and T_TIERS="'+TobContexte.GetValue('TIERS')+'"');
    if (aTiers<>nil) then
    begin
      TobContexte.AddChampSupValeur('TARIFTIERS', aTiers[0], false);
      TobContexte.AddChampSupValeur('FACTUREHT' , aTiers[1], false);
    end
    else
      PgiInfo('Le code tiers est inconnu','La recherche sera incomplète');
  end;

  {----------------------------------------------------------------------------
    Ajout dans la TobContexte des informations provenant de la fiche article
  -----------------------------------------------------------------------------}
  procedure AjoutComplementArticle;
  var
    aArticle : MyArrayValue;
  begin
    // Recherche des données de la fiche article
    aArticle := WGetSqlFieldsValues(['GA_TARIFARTICLE', 'GA_POIDSNET', 'GA_POIDSBRUT', 'GA_LINEAIRE', 'GA_SURFACE', 'GA_VOLUME', 'GA_PRIXPOURQTE', 'GA_QUALIFUNITEVTE'], 'ARTICLE', 'GA_ARTICLE="'+TobContexte.GetValue('ARTICLE')+'"');
    if (aArticle<>nil) then
    begin
      TobContexte.AddChampSupValeur('TARIFARTICLE'    , aArticle[0], false);
      TobContexte.AddChampSupValeur('POIDSNET'        , aArticle[1], false);
      TobContexte.AddChampSupValeur('POIDSBRUT'       , aArticle[2], false);
      TobContexte.AddChampSupValeur('LINEAIRE'        , aArticle[3], false);
      TobContexte.AddChampSupValeur('SURFACE'         , aArticle[4], false);
      TobContexte.AddChampSupValeur('VOLUME'          , aArticle[5], false);
      TobContexte.AddChampSupValeur('PRIXPOURQTE'     , aArticle[6], false);
      TobContexte.AddChampSupValeur('QUALIFUNITEVTE'  , aArticle[7], false);
    end
    else
      PgiInfo('Le code article est inconnu','La recherche sera incomplète');
  end;

  {----------------------------------------------------------------------------
    Ajout dans la TobContexte des informations provenant de la fiche dépôt
  -----------------------------------------------------------------------------}
  procedure AjoutComplementDepot;
  begin
    TobContexte.AddChampSupValeur('TARIFDEPOT', '', false);
  end;

  {----------------------------------------------------------------------------
    Ajout dans la TobContexte des informations provenant de la fiche affaire
  -----------------------------------------------------------------------------}
  procedure AjoutComplementAffaire;
  begin
    TobContexte.AddChampSupValeur('TARIFAFFAIRE', '', false);
  end;

  {----------------------------------------------------------------------------
    Ajout dans la TobContexte des informations provenant de la fiche ressource
  -----------------------------------------------------------------------------}
  procedure AjoutComplementRessource;
  begin
    TobContexte.AddChampSupValeur('TARIFRESSOURCE', '', false);
  end;

  {----------------------------------------------------------------------------
    Charge la TobAqui de toutes les tarifs A Qui répondant au contexte
  -----------------------------------------------------------------------------}
  function GetTobAQui : boolean;
  var
    sContexte, sChamps, sRequete : String;

    {----------------------------------------------------------------------------
      Construit la liste des champs à ramener de la table YTARIFSAQUI en fonction du tarif et de ses paramètres
    -----------------------------------------------------------------------------}
    function ConstruitChampsAQui : string;
    var
      sConditionsApplications : string;
    begin

      Result := 'YTA_IDENTIFIANT, '+
                'YTA_FONCTIONNALITE, '+
                'YTA_POIDSRECHERCHE, '
                ;

      if (wTrue=wTrue) or (sParametre[iPosOrigineDepot]          =wTrue) then Result := Result + 'YTA_DEPOT, YTA_CASCDEPOT, ';
      if (wTrue=wTrue) or (sParametre[iPosOrigineTarifDepot]     =wTrue) then Result := Result + 'YTA_TARIFDEPOT, YTA_CASCTARIFDEPOT, ';
      if (wTrue=wTrue) or (sParametre[iPosOrigineDevise]         =wTrue) then Result := Result + 'YTA_DEVISE, ';
      if (wTrue=wTrue) or (sParametre[iPosOrigineSpecial]        =wTrue) then Result := Result + 'YTA_TARIFSPECIAL, YTA_CASCSPECIAL, ';
      if (wTrue=wTrue) or (sParametre[iPosOrigineCommercial]     =wTrue) then Result := Result + 'YTA_COMMERCIAL, YTA_CASCCOMMERCIAL, ';
      if (wTrue=wTrue) or (sParametre[iPosOrigineTypeCommercial] =wTrue) then Result := Result + 'YTA_TYPECOMMERCIAL, YTA_CASCTYPECOMM, ';
      if (wTrue=wTrue) or (sParametre[iPosOrigineTiers]          =wTrue) then Result := Result + 'YTA_TIERS, YTA_CASCTIERS, ';
      if (wTrue=wTrue) or (sParametre[iPosOrigineTarifTiers]     =wTrue) then Result := Result + 'YTA_TARIFTIERS, YTA_CASCTARIFTIERS, ';
      if (wTrue=wTrue) or (sParametre[iPosOrigineAffaire]        =wTrue) then Result := Result + 'YTA_AFFAIRE, YTA_CASCAFFAIRE, ';
      if (wTrue=wTrue) or (sParametre[iPosOrigineTarifAffaire]   =wTrue) then Result := Result + 'YTA_TARIFAFFAIRE, YTA_CASCTARIFAFF, ';
      if (wTrue=wTrue) or (sParametre[iPosOriginePole]           =wTrue) then Result := Result + 'YTA_POLE, YTA_CASCPOLE, ';
      if (wTrue=wTrue) or (sParametre[iPosOrigineAtelier]        =wTrue) then Result := Result + 'YTA_SITE, YTA_CASCATELIER, ';
      if (wTrue=wTrue) or (sParametre[iPosOrigineGroupe]         =wTrue) then Result := Result + 'YTA_GRP, YTA_CASCGROUPE, ';
      if (wTrue=wTrue) or (sParametre[iPosOrigineJourSemaine]    =wTrue) then Result := Result + 'YTA_JOURSEMAINE, YTA_CASCJOUR, ';
      if (wTrue=wTrue) or (sParametre[iPosOrigineHeureJour]      =wTrue) then Result := Result + 'YTA_HEUREJOUR, YTA_CASCHEURE, ';
      if (wTrue=wTrue) or (sParametre[iPosOrigineLieuConso]      =wTrue) then Result := Result + 'YTA_LIEUCONSO, YTA_CASCLIEU, ';

      sConditionsApplications := iif(lConditionsApplications,'YTA_BLOCNOTE , ','');

      Result := Result+
                'YTA_DATEDEBUT, '+
                'YTA_DATEFIN, '+
                'YTA_DATEREFERENCE, '+
                'YTA_TTCOUHT, '+
                sConditionsApplications+
                '(YTA_DATEFIN - YTA_DATEDEBUT) PLAGE ';
    end;

    {----------------------------------------------------------------------------
      construit la requête sur la table YTARIFSAQUI en fonction du tarif et de ses paramètres
    -----------------------------------------------------------------------------}
    function ConstruitContexteAQui(sDepot, sTarifDepot, sDevise, sTarifSpecial, sCommercial, sTypeCommercial, sTiers, sTarifTiers, sAffaire, sTarifAffaire, sPole, sAtelier, sGroupe, sJourSemaine, sHeureJour, sLieuConso : string) : string;
    var
      iCptTob            : integer;
      sContexte, sFormat : string;
      TheTob, TobRequete : tob;

      iCptDepot, iCptDevise, iCptSpecial, iCptCommercial, iCptTiers, iCptAffaire, iCptPole, iCptAtelier, iCptGroupe, iCptJourSemaine, iCptHeureJour, iCptLieuConso : integer;
      iMaxDepot, iMaxDevise, iMaxSpecial, iMaxCommercial, iMaxTiers, iMaxAffaire, iMaxPole, iMaxAtelier, iMaxGroupe, iMaxJourSemaine, iMaxHeureJour, iMaxLieuConso : integer;
    begin
      Result := '';
      TobRequete := Tob.Create('_Requete_', nil, -1);
      TobRequete.AddChampSup('CONTEXTE', False);

      iMaxDepot       := iif(sParametre[iPosOrigineDepot      ]=wTrue,2,1);
      iMaxDevise      := iif(sParametre[iPosOrigineDevise     ]=wTrue,2,1);
      iMaxSpecial     := iif(sParametre[iPosOrigineSpecial    ]=wTrue,2,1);
      iMaxCommercial  := iif(sParametre[iPosOrigineCommercial ]=wTrue,3,1);
      iMaxTiers       := iif(sParametre[iPosOrigineTiers      ]=wTrue,3,1);
      iMaxAffaire     := iif(sParametre[iPosOrigineAffaire    ]=wTrue,3,1);
      iMaxPole        := iif(sParametre[iPosOriginePole       ]=wTrue,2,1);
      iMaxAtelier     := iif(sParametre[iPosOrigineAtelier    ]=wTrue,2,1);
      iMaxGroupe      := iif(sParametre[iPosOrigineGroupe     ]=wTrue,2,1);
      iMaxJourSemaine := iif(sParametre[iPosOrigineJourSemaine]=wTrue,2,1);
      iMaxHeureJour   := iif(sParametre[iPosOrigineHeureJour  ]=wTrue,2,1);
      iMaxLieuConso   := iif(sParametre[iPosOrigineLieuConso  ]=wTrue,2,1);

      try
        for iCptDepot      :=1 to iMaxDepot       do
        begin
          for iCptDevise     :=1 to iMaxDevise      do
          begin
            for icptSpecial    :=1 to iMaxSpecial     do
            begin
              for iCptCommercial :=1 to iMaxCommercial  do
              begin
                for iCptTiers      :=1 to iMaxTiers       do
                begin
                  for iCptAffaire    :=1 to iMaxAffaire     do
                  begin
                    for iCptPole       :=1 to iMaxPole        do
                    begin
                      for iCptAtelier    :=1 to iMaxAtelier     do
                      begin
                        for iCptGroupe     :=1 to iMaxGroupe      do
                        begin
                          for icptJourSemaine:=1 to iMaxJourSemaine do
                          begin
                            for iCptHeureJour  :=1 to iMaxHeureJour   do
                            begin
                              for iCptLieuConso  :=1 to iMaxLieuConso   do
                              begin
                                sFormat := sOrigineMasque;
                                sFormat := Remplace(sFormat, IntToStr(iCptDepot)      , iPosOrigineDepot      );
                                sFormat := Remplace(sFormat, IntToStr(iCptDevise)     , iPosOrigineDevise     );
                                sFormat := Remplace(sFormat, IntToStr(iCptSpecial)    , iPosOrigineSpecial    );
                                sFormat := Remplace(sFormat, IntToStr(iCptCommercial) , iPosOrigineCommercial );
                                sFormat := Remplace(sFormat, IntToStr(iCptTiers)      , iPosOrigineTiers      );
                                sFormat := Remplace(sFormat, IntToStr(iCptAffaire)    , iPosOrigineAffaire    );
                                sFormat := Remplace(sFormat, IntToStr(iCptPole)       , iPosOriginePole       );
                                sFormat := Remplace(sFormat, IntToStr(iCptAtelier)    , iPosOrigineAtelier    );
                                sFormat := Remplace(sFormat, IntToStr(iCptGroupe)     , iPosOrigineGroupe     );
                                sFormat := Remplace(sFormat, IntToStr(iCptJoursemaine), iPosOrigineJourSemaine);
                                sFormat := Remplace(sFormat, IntToStr(iCptHeureJour)  , iPosOrigineHeureJour  );
                                sFormat := Remplace(sFormat, IntToStr(iCptLieuConso)  , iPosOrigineLieuConso  );

                                sContexte := ContexteAqui(sFormat, sParametre, sDepot, sTarifDepot, sDevise, sTarifSpecial, sCommercial, sTypeCommercial, sTiers, sTarifTiers, sAffaire, sTarifAffaire, sPole, sAtelier, sGroupe, sJourSemaine, sHeureJour, sLieuConso);

                                TheTob := TobRequete.FindFirst(['CONTEXTE'],[sContexte], False);
                                if (not Assigned(TheTob)) then
                                begin
                                  TheTob := Tob.Create('_Requete_', TobRequete, -1);
                                  TheTob.AddChampSup('CONTEXTE', False);
                                  TheTob.PutValue('CONTEXTE',sContexte);
                                end;
                              end;
                            end;
                          end;
                        end;
                      end;
                    end;
                  end;
                end;
              end;
            end;
          end;
        end;
        for iCptTob:=0 to (TobRequete.Detail.Count-1) do
        begin
          Result := Result + TobRequete.Detail[iCptTob].GetValue('CONTEXTE');
        end;
        Result := copy(Result,1,length(Result)-4);
      finally
        TobRequete.Free;
      end;
    end;

  { Début GetTobAqui}
  begin
    sChamps   := ConstruitChampsAQui;
    sContexte := ConstruitContexteAQui(TobContexte.GetValue('DEPOT'),
                                       TobContexte.GetValue('TARIFDEPOT'),
                                       TobContexte.GetValue('DEVISE'),
                                       TobContexte.GetValue('TARIFSPECIAL'),
                                       TobContexte.GetValue('COMMERCIAL'),
                                       TobContexte.GetValue('TYPECOMMERCIAL'),
                                       TobContexte.GetValue('TIERS'),
                                       TobContexte.GetValue('TARIFTIERS'),
                                       TobContexte.GetValue('AFFAIRE'),
                                       TobContexte.GetValue('TARIFAFFAIRE'),
                                       TobContexte.GetValue('POLE'),
                                       TobContexte.GetValue('ATELIER'),
                                       TobContexte.GetValue('GROUPE'),
                                       TobContexte.GetValue('JOURSEMAINE'),
                                       TobContexte.GetValue('HEUREJOUR'),
                                       TobContexte.GetValue('LIEUCONSO')
                                      );
    sRequete := 'SELECT '+
                   sChamps +
                'FROM '+
                   'YTARIFSAQUI '+
                'WHERE '+
                   ' ( '+ 'YTA_FONCTIONNALITE="'+TobContexte.GetValue('FONCTIONNALITE')+'"'+' ) '+
                   ' and '+
                   ' ( '
                        +'((YTA_DATEREFERENCE="CDE") and (YTA_DATEDEBUT<="'+USDateTime(TobContexte.GetValue('DATECOMMANDE' ))+'") and ('+'YTA_DATEFIN > "'+USDateTime(TobContexte.GetValue('DATECOMMANDE' ))+'"))'
                        +' or '
                        +'((YTA_DATEREFERENCE="DEL") and (YTA_DATEDEBUT<="'+USDateTime(TobContexte.GetValue('DATELIVRAISON'))+'") and ('+'YTA_DATEFIN > "'+USDateTime(TobContexte.GetValue('DATELIVRAISON'))+'"))'+
                   ' ) '+
                   ' and '+
                   ' ( '+ sContexte +' ) '+
                   ' and '+
                   ' ( '+ 'YTA_TTCOUHT="'+iif(TobContexte.GetValue('FACTUREHT')=wTrue,'H','T')+'"'+' ) '+
                ' ORDER BY YTA_POIDSRECHERCHE DESC, PLAGE ';
    Result := TobAqui.LoadDetailDBFromSql('_TarifsAQui_', sRequete);
  end;
  { Fin GetTobAqui}


  {----------------------------------------------------------------------------
    Charge la Tobsurquoi de toutes les tarifs Sur Quoi répondant au contexte
  -----------------------------------------------------------------------------}
  function GetTobSurQuoi : boolean;
  var
    sChamps, sContexte, sRequete : String;

    {----------------------------------------------------------------------------
      Construit la liste des champs à ramener de la table YTARIFSSURQUOI en fonction du tarif et de ses paramètres
    -----------------------------------------------------------------------------}
    function ConstruitChampsSurQuoi : string;
    var
      sConditionsApplications : string;
    begin
      Result := 'YTQ_IDENTIFIANTYTA, '+
                'YTQ_IDENTIFIANT, '+
                'YTQ_POIDSRECHERCHE, ';

      if (wTrue=wTrue) or (sParametre[iPosOrigineArticle]       =wTrue) then Result := Result + 'YTQ_ARTICLE, YTQ_CASCARTICLE, ';
      if (wTrue=wTrue) or (sParametre[iPosOrigineTarifArticle]  =wTrue) then Result := Result + 'YTQ_TARIFARTICLE, YTQ_CASCTARIFARTIC, ';
      if (wTrue=wTrue) or (sParametre[iPosOrigineRessource]     =wTrue) then Result := Result + 'YTQ_RESSOURCE, YTQ_CASCRESSOURCE, ';
      if (wTrue=wTrue) or (sParametre[iPosOrigineTarifRessource]=wTrue) then Result := Result + 'YTQ_TARIFRESSOURCE, YTQ_CASCTARIFRESSO, ';
      if (wTrue=wTrue) or (sParametre[iPosOrigineMarque]        =wTrue) then Result := Result + 'YTQ_MARQUE, YTQ_CASCMARQUE, ';
      if (wTrue=wTrue) or (sParametre[iPosOrigineCircuit]       =wTrue) then Result := Result + 'YTQ_CIRCUIT, YTQ_CASCCIRCUIT, ';
      if (wTrue=wTrue) or (sParametre[iPosOriginePhase]         =wTrue) then Result := Result + 'YTQ_PHASE, YTQ_CASCPHASE, ';
      if (wTrue=wTrue) or (sParametre[iPosOrigineTraitement]    =wTrue) then Result := Result + 'YTQ_TRAITEMENT, YTQ_CASCTRAITEMENT, ';

      sConditionsApplications := iif(lConditionsApplications,'YTQ_BLOCNOTE , ','');

      Result := Result +
                sConditionsApplications+
                'YTQ_PARTRANCHE, '+
                'YTQ_FORFAIT, '+
                'YTQ_COUTLANCEMENT, '+
                'YTQ_REMISEMONTANT, '+
                'YTQ_PRIXBRUT, '+
                'YTQ_REMISE1, '+
                'YTQ_REMISE2, '+
                'YTQ_REMISE3, '+
                'YTQ_PRIXNET, '+
                'YTQ_CONDFORFAIT, '+
                'YTQ_CONDCOUTLANC, '+
                'YTQ_CONDREMISEM, '+
                'YTQ_CONDPRIXBRUT, '+
                'YTQ_CONDREMISE1, '+
                'YTQ_CONDREMISE2, '+
                'YTQ_CONDREMISE3, '+
                'YTQ_CONDPRIXNET, '+
                'YTQ_FOURCHETTEDE, '+
                'YTQ_SOMMEROU, '+
                'YTQ_RECHERCHEROU, '+
                'YTQ_LIBELLETARIF ';
    end;

    {----------------------------------------------------------------------------
      Construit la requête sur la table YTARIFSSURQUOI en fonction du tarif et de ses paramètres
    -----------------------------------------------------------------------------}
    function ConstruitContexteSurQuoi : string;
    var
      iCptTob             : integer;
      TheTob, TobRequete  : tob;
      sContexte, sFormat  : string;
      iCptArticle, iCptRessource, iCptMarque, iCptCircuit, iCptPhase, iCptTraitement : integer;
      iMaxArticle, iMaxRessource, iMaxMarque, iMaxCircuit, iMaxPhase, iMaxTraitement : integer;
    begin
      Result := '';
      TobRequete := Tob.Create('_Requete_', nil, -1);
      TobRequete.AddChampSup('CONTEXTE', False);

      iMaxArticle    := iif(sParametre[iPosOrigineArticle   ]=wTrue,4,1);
      iMaxRessource  := iif(sParametre[iPosOrigineRessource ]=wTrue,3,1);
      iMaxMarque     := iif(sParametre[iPosOrigineMarque    ]=wTrue,2,1);
      iMaxCircuit    := iif(sParametre[iPosOrigineCircuit   ]=wTrue,2,1);
      iMaxPhase      := iif(sParametre[iPosOriginePhase     ]=wTrue,2,1);
      iMaxTraitement := iif(sParametre[iPosOrigineTraitement]=wTrue,2,1);

      try
        for iCptArticle    :=1 to iMaxArticle    do
        begin
          for iCptRessource  :=1 to iMaxRessource  do
          begin
            for iCptMarque     :=1 to iMaxMarque     do
            begin
              for iCptCircuit    :=1 to iMaxCircuit    do
              begin
                for iCptPhase      :=1 to iMaxPhase      do
                begin
                  for iCptTraitement :=1 to iMaxTraitement do
                  begin
                    sFormat := sOrigineMasque;
                    sFormat := Remplace(sFormat, IntToStr(iCptArticle)   , iPosOrigineArticle   );
                    sFormat := Remplace(sFormat, IntToStr(iCptRessource) , iPosOrigineRessource );
                    sFormat := Remplace(sFormat, IntToStr(iCptMarque)    , iPosOrigineMarque    );
                    sFormat := Remplace(sFormat, IntToStr(iCptCircuit)   , iPosOrigineCircuit   );
                    sFormat := Remplace(sFormat, IntToStr(iCptPhase)     , iPosOriginePhase     );
                    sFormat := Remplace(sFormat, IntToStr(iCptTraitement), iPosOrigineTraitement);

                    sContexte := ContexteSurQuoi(sFormat, sParametre, TobContexte, iPartiel);

                    TheTob := TobRequete.FindFirst(['CONTEXTE'],[sContexte], False);
                    if (not Assigned(TheTob)) then
                    begin
                      TheTob := Tob.Create('_Requete_', TobRequete, -1);
                      TheTob.AddChampSup('CONTEXTE', False);
                      TheTob.PutValue('CONTEXTE',sContexte);
                    end;
                  end;
                end;
              end;
            end;
          end;
        end;
        for iCptTob:=0 to (TobRequete.Detail.Count-1) do
        begin
          Result := Result + TobRequete.Detail[iCptTob].GetValue('CONTEXTE');
        end;
        Result := copy(Result,1,length(Result)-4);
      finally
        TobRequete.Free;
      end;
    end;

    {-------------------------------------------------------------------------
    Construction du WHERE pour la TobSurQuoi à partir du contenu de la TobAQui
    --------------------------------------------------------------------------}
    function DeContexteAQui : string;
    var
       iCptAQui : integer;
    begin
      Result := ' ( ';
      if (TobAQui.Detail.Count<>0) then
        begin
        for iCptAQui:=0 to (TobAQui.Detail.Count-1) do
        begin
          Result := Result + '(YTQ_IDENTIFIANTYTA='+IntToStr(TobAQui.Detail[iCptAqui].GetValue('YTA_IDENTIFIANT'))+')'+' or ';
        end;
      end
      else
        Result := Result + '(YTQ_IDENTIFIANTYTA='+IntToStr(0)+')'+' or ';

      Result := copy(Result,1,Length(Result)-4)+' ) ';
    end;

  { Début GetTobSurquoi}
  begin
    sChamps   := ConstruitChampsSurQuoi;
    sContexte := ConstruitContexteSurQuoi;
    sRequete := 'SELECT '+
                    sChamps+
                'FROM '+
                   'YTARIFSSURQUOI '+
                'WHERE '+
                   '( '+DeContexteAQui+' )'+' and '+
                   '( '+
                   sContexte+
                   ' ) '+
                'ORDER BY  YTQ_IDENTIFIANTYTA, YTQ_POIDSRECHERCHE DESC';
    Result := TobSurQuoi.LoadDetailDBFromSql('_TarifsSurQuoi_', sRequete);
  end;
  { Fin GetTobSurQuoi }


  {----------------------------------------------------------------------------
    Charge la TobTarifs et le TobRemises
  -----------------------------------------------------------------------------}
  procedure GetTobTarifs;
  var
    iCptAQui, iCptSurQuoi      : integer;

    {---------------------------------------------------------------------------------
       Permet de codifier l'origine de la recherche qui a aboutie en fonction du contexte
    ----------------------------------------------------------------------------------}
    function CodifieOrigine : string;
    var
      sArticle, sOrigine : string;

    begin
      sOrigine := sOrigineMasque;
      sOrigine := Remplace(sOrigine, iif((sParametre[iPosOrigineDepot]         =wTrue) and (TobAQui.Detail[iCptAQui].GetValue('YTA_DEPOT')         = TobContexte.GetValue('DEPOT')         ) and (   TobAQui.Detail[iCptAQui   ].GetValue('YTA_DEPOT')         <>'' )               ,'D','.'), iPosOrigineDepot         );
      sOrigine := Remplace(sOrigine, iif((sParametre[iPosOrigineTarifDepot]    =wTrue) and (TobAQui.Detail[iCptAQui].GetValue('YTA_TARIFDEPOT')    = TobContexte.GetValue('DEPOT')         ) and (   TobAQui.Detail[iCptAQui   ].GetValue('YTA_TARIFDEPOT')    <>'' )               ,'F','.'), iPosOrigineTarifDepot    );
      sOrigine := Remplace(sOrigine, iif((sParametre[iPosOrigineDevise]        =wTrue) and (TobAQui.Detail[iCptAQui].GetValue('YTA_DEVISE')        = TobContexte.GetValue('DEVISE')        ) and (   TobAQui.Detail[iCptAQui   ].GetValue('YTA_DEVISE')        <>V_PGI.DevisePivot) ,'D','.'), iPosOrigineDevise        );
      sOrigine := Remplace(sOrigine, iif((sParametre[iPosOrigineSpecial]       =wTrue) and (TobAQui.Detail[iCptAQui].GetValue('YTA_TARIFSPECIAL')  = TobContexte.GetValue('TARIFSPECIAL')  ) and (   TobAQui.Detail[iCptAQui   ].GetValue('YTA_TARIFSPECIAL')  <>'' )               ,'S','.'), iPosOrigineSpecial       );
      sOrigine := Remplace(sOrigine, iif((sParametre[iPosOrigineCommercial]    =wTrue) and (TobAQui.Detail[iCptAQui].GetValue('YTA_COMMERCIAL')    = TobContexte.GetValue('COMMERCIAL')    ) and (   TobAQui.Detail[iCptAQui   ].GetValue('YTA_COMMERCIAL')    <>'' )               ,'C','.'), iPosOrigineCommercial    );
      sOrigine := Remplace(sOrigine, iif((sParametre[iPosOrigineTypeCommercial]=wTrue) and (TobAQui.Detail[iCptAQui].GetValue('YTA_TYPECOMMERCIAL')= TobContexte.GetValue('TYPECOMMERCIAL')) and (   TobAQui.Detail[iCptAQui   ].GetValue('YTA_TYPECOMMERCIAL')<>'' )               ,'F','.'), iPosOrigineTypeCommercial);
      sOrigine := Remplace(sOrigine, iif((sParametre[iPosOrigineTiers]         =wTrue) and (TobAQui.Detail[iCptAQui].GetValue('YTA_TIERS')         = TobContexte.GetValue('TIERS')         ) and (   TobAQui.Detail[iCptAQui   ].GetValue('YTA_TIERS')         <>'' )               ,'T','.'), iPosOrigineTiers         );
      sOrigine := Remplace(sOrigine, iif((sParametre[iPosOrigineTarifTiers]    =wTrue) and (TobAQui.Detail[iCptAQui].GetValue('YTA_TARIFTIERS')    = TobContexte.GetValue('TARIFTIERS')    ) and (   TobAQui.Detail[iCptAQui   ].GetValue('YTA_TARIFTIERS')    <>'' )               ,'F','.'), iPosOrigineTarifTiers    );

      sArticle := iif((sParametre[iPosOrigineArticle]=wTrue) and (TobSurQuoi.Detail[iCptSurQuoi].GetValue('YTQ_ARTICLE')=TobContexte.GetValue('ARTICLE'))and (TobSurQuoi.Detail[iCptSurQuoi].GetValue('YTQ_ARTICLE')<>''),'A','.');
      if (sArticle='.') and (sParametre[iPosOrigineArticle]=wTrue) and (iPartiel<>0) then
        sArticle := iif((TobSurQuoi.Detail[iCptSurQuoi].GetValue('YTQ_ARTICLE')=Copy(TobContexte.GetValue('ARTICLE'),1,iPartiel)) and (TobSurQuoi.Detail[iCptSurQuoi].GetValue('YTQ_ARTICLE')<>''),'P','.');

      sOrigine := Remplace(sOrigine, sArticle                                                                                                                                                                                                                                            , iPosOrigineArticle     );
      sOrigine := Remplace(sOrigine, iif((sParametre[iPosOrigineTarifArticle]=wTrue) and (TobSurQuoi.Detail[iCptSurQuoi].GetValue('YTQ_TARIFARTICLE') = TobContexte.GetValue('TARIFARTICLE'))             and (TobSurQuoi.Detail[iCptSurQuoi].GetValue('YTQ_TARIFARTICLE')<>'' ),'F','.'), iPosOrigineTarifArticle);
      sOrigine := Remplace(sOrigine, iif((sParametre[iPosOrigineMarque]      =wTrue) and (TobSurQuoi.Detail[iCptSurQuoi].GetValue('YTQ_MARQUE')       = TobContexte.GetValue('MARQUE')      )             and (TobSurQuoi.Detail[iCptSurQuoi].GetValue('YTQ_MARQUE')      <>'' ),'M','.'), iPosOrigineMarque      );
      sOrigine := Remplace(sOrigine, iif((sParametre[iPosOrigineCircuit]     =wTrue) and (TobSurQuoi.Detail[iCptSurQuoi].GetValue('YTQ_CIRCUIT')      = TobContexte.GetValue('CIRCUIT')     )             and (TobSurQuoi.Detail[iCptSurQuoi].GetValue('YTQ_CIRCUIT')     <>'' ),'C','.'), iPosOrigineCircuit     );
      sOrigine := Remplace(sOrigine, iif((sParametre[iPosOriginePhase]       =wTrue) and (TobSurQuoi.Detail[iCptSurQuoi].GetValue('YTQ_PHASE')        = TobContexte.GetValue('PHASE')       )             and (TobSurQuoi.Detail[iCptSurQuoi].GetValue('YTQ_PHASE')       <>'' ),'P','.'), iPosOriginePhase       );
      sOrigine := Remplace(sOrigine, iif((sParametre[iPosOrigineTraitement]  =wTrue) and (TobSurQuoi.Detail[iCptSurQuoi].GetValue('YTQ_TRAITEMENT')   = TobContexte.GetValue('TRAITEMENT')  )             and (TobSurQuoi.Detail[iCptSurQuoi].GetValue('YTQ_TRAITEMENT')  <>'' ),'T','.'), iPosOrigineTraitement  );

      sOrigine := Remplace(sOrigine, Copy(TobAqui.Detail[iCptAQui].GetValue('YTA_DATEREFERENCE'),1,1), iPosOrigineDateDeReference);
      sOrigine := Remplace(sOrigine, TobAqui.Detail[iCptAQui].GetValue('YTA_TTCOUHT')                , iPosOrigineTTCouHT        );

      if (TobSurQuoi.Detail[iCptSurQuoi].GetValue('YTQ_PARTRANCHE')=wTrue) then
      begin
        sOrigine := Remplace(sOrigine, TobSurQuoi.Detail[iCptSurQuoi].GetValue('YTQ_RECHERCHEROU'), iPosOrigineRechercherOu );
        sOrigine := Remplace(sOrigine, TobSurQuoi.Detail[iCptSurQuoi].GetValue('YTQ_SOMMEROU')    , iPosOrigineSommerOu     );
        sOrigine := Remplace(sOrigine, TobSurQuoi.Detail[iCptSurQuoi].GetValue('YTQ_FOURCHETTEDE'), iPosOrigineSommerQuoi   );
      end;
      Result := sOrigine;
    end;

    {----------------------------------------------------------------------------
       Mise à Jour des paramètres de gestion des remises en cascade
    -----------------------------------------------------------------------------}
    procedure ParametreCascadeRemises(sOriginePrecedente:string);
    begin
      TobTarifs.PutValue('YTA_CASCDEPOT'     , TobAQui.Detail[iCptAqui].GetValue('YTA_CASCDEPOT'));
      TobTarifs.PutValue('YTA_CASCSPECIAL'   , TobAQui.Detail[iCptAqui].GetValue('YTA_CASCSPECIAL'));
      TobTarifs.PutValue('YTA_CASCCOMMERCIAL', TobAQui.Detail[iCptAqui].GetValue('YTA_CASCCOMMERCIAL'));
      TobTarifs.PutValue('YTA_CASCTYPECOMM'  , TobAQui.Detail[iCptAqui].GetValue('YTA_CASCTYPECOMM'));
      TobTarifs.PutValue('YTA_CASCTIERS'     , TobAQui.Detail[iCptAqui].GetValue('YTA_CASCTIERS'));
      TobTarifs.PutValue('YTA_CASCTARIFTIERS', TobAQui.Detail[iCptAqui].GetValue('YTA_CASCTARIFTIERS'));
      TobTarifs.PutValue('YTQ_CASCARTICLE'   , TobSurQuoi.Detail[iCptSurQuoi].GetValue('YTQ_CASCARTICLE'));
      TobTarifs.PutValue('YTQ_CASCTARIFARTIC', TobSurQuoi.Detail[iCptSurQuoi].GetValue('YTQ_CASCTARIFARTIC'));
      TobTarifs.PutValue('YTQ_CASCMARQUE'    , TobSurQuoi.Detail[iCptSurQuoi].GetValue('YTQ_CASCMARQUE'));
      TobTarifs.PutValue('YTQ_CASCCIRCUIT'   , TobSurQuoi.Detail[iCptSurQuoi].GetValue('YTQ_CASCCIRCUIT'));
      TobTarifs.PutValue('YTQ_CASCPHASE'     , TobSurQuoi.Detail[iCptSurQuoi].GetValue('YTQ_CASCPHASE'));
      TobTarifs.PutValue('YTQ_CASCTRAITEMENT', TobSurQuoi.Detail[iCptSurQuoi].GetValue('YTQ_CASCTRAITEMENT'));
      TobTarifs.PutValue('ORIGINEPRECEDENTE' , sOriginePrecedente);
    end;

    {----------------------------------------------------------------------------
       Cascade Autorise sur un contexte moins précis
    -----------------------------------------------------------------------------}
    function CascadeRemisesAutorisee(sOrigineSuivante : string) : boolean;
    var
      sOriginePrecedente, sOrigineAutorisee : string;

    begin
      sOriginePrecedente := TobTarifs.GetValue('ORIGINEPRECEDENTE');
      if (trim(sOriginePrecedente)<>'') then
      begin
        sOriginePrecedente := copy(TobTarifs.GetValue('ORIGINEPRECEDENTE'), 1, 35);
        sOrigineSuivante   := copy(sOrigineSuivante                       , 1, 35);

        // Construction du masque de l'origine autorisée selon les cascades paramétrées
        //
        //  Pour chaque position de l'origine
        //  Si la cascade est autorise alors la construction du masque se fait avec le caractère de l'origine suivante
        //                             sinon la construction du masque se fait avec le caractère de l'origine précédente

        sOrigineAutorisee  := sOrigineSuivante;

        // dépôt
        if (TobTarifs.GetValue('YTA_CASCDEPOT'    )=wTrue) then
          sOrigineAutorisee  := Remplace(sOrigineAutorisee, sOrigineSuivante[iPosOrigineDepot]  , iPosOrigineDepot)
        else
          sOrigineAutorisee  := Remplace(sOrigineAutorisee, sOriginePrecedente[iPosOrigineDepot], iPosOrigineDepot);

        // Devise
        sOrigineAutorisee  := Remplace(sOrigineAutorisee, sOriginePrecedente[iPosOrigineDevise], iPosOrigineDevise);

        // Conditions Spéciales
        if (TobTarifs.GetValue('YTA_CASCSPECIAL'   )=wTrue) then
          sOrigineAutorisee  := Remplace(sOrigineAutorisee, sOrigineSuivante[iPosOrigineSpecial]  , iPosOrigineSpecial)
        else
          sOrigineAutorisee  := Remplace(sOrigineAutorisee, sOriginePrecedente[iPosOrigineSpecial], iPosOrigineSpecial);

        // Commercial et Type de Commercial
        if (TobTarifs.GetValue('YTA_CASCCOMMERCIAL')=wTrue) or (TobTarifs.GetValue('YTA_CASCTYPECOMM')=wTrue) then
          sOrigineAutorisee  := Remplace(sOrigineAutorisee, sOrigineSuivante[iPosOrigineCommercial]  , iPosOrigineCommercial)
        else
          sOrigineAutorisee  := Remplace(sOrigineAutorisee, sOriginePrecedente[iPosOrigineCommercial], iPosOrigineCommercial);

        if (TobTarifs.GetValue('YTA_CASCCOMMERCIAL')=wTrue) then
        begin
          if (TobTarifs.GetValue('YTA_CASCTYPECOMM')=wTrue)  or (sOrigineAutorisee[iPosOrigineCommercial]='C') then
            sOrigineAutorisee  := Remplace(sOrigineAutorisee, sOrigineSuivante[iPosOrigineTypeCommercial], iPosOrigineTypeCommercial)
          else
            sOrigineAutorisee  := Remplace(sOrigineAutorisee, 'T'                                        , iPosOrigineTypeCommercial);
        end
        else
          sOrigineAutorisee  := Remplace(sOrigineAutorisee, sOriginePrecedente[iPosOrigineTypeCommercial], iPosOrigineTypeCommercial);

        // Tiers et Famille de Tiers
        if (TobTarifs.GetValue('YTA_CASCTIERS')=wTrue) or (TobTarifs.GetValue('YTA_CASCTARIFTIERS')=wTrue) then
          sOrigineAutorisee  := Remplace(sOrigineAutorisee, sOrigineSuivante[iPosOrigineTiers]  , iPosOrigineTiers)
        else
          sOrigineAutorisee  := Remplace(sOrigineAutorisee, sOriginePrecedente[iPosOrigineTiers], iPosOrigineTiers);

        if (TobTarifs.GetValue('YTA_CASCTIERS')=wTrue) then
        begin
          if (TobTarifs.GetValue('YTA_CASCTARIFTIERS')=wTrue) or (sOrigineAutorisee[iPosOrigineTiers]='T') then
            sOrigineAutorisee  := Remplace(sOrigineAutorisee, sOrigineSuivante[iPosOrigineTarifTiers], iPosOrigineTarifTiers)
          else
            sOrigineAutorisee  := Remplace(sOrigineAutorisee, 'F'                                    , iPosOrigineTarifTiers);
        end
        else
          sOrigineAutorisee  := Remplace(sOrigineAutorisee, sOriginePrecedente[iPosOrigineTarifTiers], iPosOrigineTarifTiers);

        // Article et Famille d'article
        if (TobTarifs.GetValue('YTQ_CASCARTICLE')=wTrue) or (TobTarifs.GetValue('YTQ_CASCTARIFARTIC')=wTrue) then
          sOrigineAutorisee  := Remplace(sOrigineAutorisee, sOrigineSuivante  [iPosOrigineArticle], iPosOrigineArticle)
        else
          sOrigineAutorisee  := Remplace(sOrigineAutorisee, sOriginePrecedente[iPosOrigineArticle], iPosOrigineArticle);

        if (TobTarifs.GetValue('YTQ_CASCARTICLE')=wTrue) then
        begin
          if (TobTarifs.GetValue('YTQ_CASCTARIFARTIC')=wTrue) or (sOrigineAutorisee[iPosOrigineArticle]='T') then
            sOrigineAutorisee  := Remplace(sOrigineAutorisee, sOrigineSuivante[iPosOrigineTarifArticle], iPosOrigineTarifArticle)
          else
            sOrigineAutorisee  := Remplace(sOrigineAutorisee, 'F'                                      , iPosOrigineTarifArticle);
        end
        else
          sOrigineAutorisee  := Remplace(sOrigineAutorisee, sOriginePrecedente[iPosOrigineTarifArticle], iPosOrigineTarifArticle);

        sOrigineAutorisee  := Remplace(sOrigineAutorisee, iif(TobTarifs.GetValue('YTQ_CASCMARQUE')    =wTrue,sOrigineSuivante[iPosOrigineMarque    ],sOriginePrecedente[iPosOrigineMarque    ]),iPosOrigineMarque);
        sOrigineAutorisee  := Remplace(sOrigineAutorisee, iif(TobTarifs.GetValue('YTQ_CASCCIRCUIT')   =wTrue,sOrigineSuivante[iPosOrigineCircuit   ],sOriginePrecedente[iPosOrigineCircuit   ]),iPosOrigineCircuit);
        sOrigineAutorisee  := Remplace(sOrigineAutorisee, iif(TobTarifs.GetValue('YTQ_CASCPHASE')     =wTrue,sOrigineSuivante[iPosOriginePhase     ],sOriginePrecedente[iPosOriginePhase     ]),iPosOriginePhase);
        sOrigineAutorisee  := Remplace(sOrigineAutorisee, iif(TobTarifs.GetValue('YTQ_CASCTRAITEMENT')=wTrue,sOrigineSuivante[iPosOrigineTraitement],sOriginePrecedente[iPosOrigineTraitement]),iPosOrigineTraitement);
      end
      else
      begin
        sOrigineAutorisee:=sOrigineSuivante;
      end;

      Result := (sOrigineAutorisee=sOrigineSuivante);
    end;

    {---------------------------------------------------------------------------------
       Charge la TobTarifs et le TobRemises à partir des Sur Quoi (sans fourchette)
    ---------------------------------------------------------------------------------}
    procedure GetTobTarifsSansFourchette;
    var
      TobRemises   : Tob;
      sOrigine, sMonnaie, sNatureAuxi : string;
      nTauxMonnaie : double;
    begin
      sOrigine     := CodifieOrigine;
      sMonnaie     := TobContexte.GetValue('DEVISE');
      sNatureAuxi  := TobContexte.GetValue('NATUREAUXI');
      nTauxMonnaie := ConvMonnaie(TobAQui.Detail[iCptAqui].GetValue('YTA_DEVISE'), sMonnaie, 1, True, TobContexte.GetValue('NATUREAUXI'), 'T', ' ');

      //Forfait
      if (TobTarifs.GetValue('FORFAIT_OU')='') and (TobSurQuoi.Detail[iCptSurQuoi].GetValue('YTQ_CONDFORFAIT')=wTrue) then
      begin
        TobTarifs.PutValue('FORFAIT_VAL',ArrondirPrixMontant(TobSurQuoi.Detail[iCptSurQuoi].GetValue('YTQ_FORFAIT')*nTauxMonnaie,sMonnaie,sNatureAuxi,'M','P'));
        TobTarifs.PutValue('FORFAIT_OU' ,sOrigine);
      end;

      //Cout Lancement
      if (TobTarifs.GetValue('COUTLANCEMENT_OU')='') and (TobSurQuoi.Detail[iCptSurQuoi].GetValue('YTQ_CONDCOUTLANC')=wTrue) then
      begin
        TobTarifs.PutValue('COUTLANCEMENT_VAL',ArrondirPrixMontant(TobSurQuoi.Detail[iCptSurQuoi].GetValue('YTQ_COUTLANCEMENT')*nTauxMonnaie,sMonnaie,sNatureAuxi,'M','P'));
        TobTarifs.PutValue('COUTLANCEMENT_OU' ,sOrigine);
      end;

      //Remise en montant
      if (TobTarifs.GetValue('REMISEMONTANT_OU')='') and (TobSurQuoi.Detail[iCptSurQuoi].GetValue('YTQ_CONDREMISEM')=wTrue) then
      begin
        TobTarifs.PutValue('REMISEMONTANT_VAL',ArrondirPrixMontant(TobSurQuoi.Detail[iCptSurQuoi].GetValue('YTQ_REMISEMONTANT')*nTauxMonnaie,sMonnaie,sNatureAuxi,'M','P'));
        TobTarifs.PutValue('REMISEMONTANT_OU' ,sOrigine);
      end;

      //Prix Brut  : Si aucune recherche a abouti sur un prix net ou un prix brut et que la recherche concerne un prix brut alors affectation
      if ((TobTarifs.GetValue('PRIXBRUT_OU')='') and (TobTarifs.GetValue('PRIXNET_OU')='')) and (TobSurQuoi.Detail[iCptSurQuoi].GetValue('YTQ_CONDPRIXBRUT')=wTrue) then
      begin
        TobTarifs.PutValue('PRIXBRUT_VAL',ArrondirPrixMontant(TobSurQuoi.Detail[iCptSurQuoi].GetValue('YTQ_PRIXBRUT')*nTauxMonnaie,sMonnaie,sNatureAuxi,'P','P'));
        TobTarifs.PutValue('PRIXBRUT_OU' ,sOrigine);
      end;

      //Prix Net : Si aucune recherche a abouti sur un prix brut ou un prix net et que la recherche concerne un prix net alors affectation
      if ((TobTarifs.GetValue('PRIXNET_OU')='') and (TobTarifs.GetValue('PRIXBRUT_OU')='')) and (TobSurQuoi.Detail[iCptSurQuoi].GetValue('YTQ_CONDPRIXNET')=wTrue) then
      begin
        TobTarifs.PutValue('PRIXNET_VAL',ArrondirPrixMontant(TobSurQuoi.Detail[iCptSurQuoi].GetValue('YTQ_PRIXNET')*nTauxMonnaie,sMonnaie,sNatureAuxi,'P','P'));
        TobTarifs.PutValue('PRIXNET_OU' ,sOrigine);
      end;

      //Création des les lignes Remises si la 1ère fois ou si la cascade remise est autorisée
      if     (    (TobSurQuoi.Detail[iCptSurQuoi].GetValue('YTQ_CONDREMISE1')=wTrue)
               or (TobSurQuoi.Detail[iCptSurQuoi].GetValue('YTQ_CONDREMISE2')=wTrue)
               or (TobSurQuoi.Detail[iCptSurQuoi].GetValue('YTQ_CONDREMISE3')=wTrue)
             )
         and ( CascadeRemisesAutorisee(sOrigine) ) then
      begin

        //Sauvegarde des paramètres de Remises en Cascade
        ParametreCascadeRemises( sOrigine );

        //Remise 1/2/3
        TobRemises := CreationTobRemises( TobTarifs );

        //Remise 1/2/3
        if (TobRemises.GetValue('REMISES_OU')='') and ((TobSurQuoi.Detail[iCptSurQuoi].GetValue('YTQ_CONDREMISE1')=wTrue)  or (TobSurQuoi.Detail[iCptSurQuoi].GetValue('YTQ_CONDREMISE2')=wTrue) or (TobSurQuoi.Detail[iCptSurQuoi].GetValue('YTQ_CONDREMISE3')=wTrue)) then
        begin
          TobRemises.PutValue('REMISE1_VAL'      ,TobSurQuoi.Detail[iCptSurQuoi].GetValue('YTQ_REMISE1'));
          TobRemises.PutValue('REMISE2_VAL'      ,TobSurQuoi.Detail[iCptSurQuoi].GetValue('YTQ_REMISE2'));
          TobRemises.PutValue('REMISE3_VAL'      ,TobSurQuoi.Detail[iCptSurQuoi].GetValue('YTQ_REMISE3'));
          TobRemises.PutValue('REMISES_OU'       ,sOrigine);
          TobRemises.PutValue('REMISES_INTITULE' ,TobSurQuoi.Detail[iCptSurQuoi].GetValue('YTQ_LIBELLETARIF'));
        end;
      end;
    end;
    { Fin GetTobTarifsSansFourchette }

    {---------------------------------------------------------------------------------
       Charge la TobTarifs et le TobRemises à partir des Sur Quoi (avec fourchette)
    ---------------------------------------------------------------------------------}
    procedure GetTobTarifsAvecFourchette;
    var
      TobRemises     : Tob;
      sOrigine, sMonnaie, sNatureAuxi : string;
      nTauxMonnaie : double;

      {----------------------------------------------------------------------------
        Charge la TobFourchette à partir des conditions tarifaires avec fourchette
      -----------------------------------------------------------------------------}
      function GetTobFourchette : boolean;
      var
        sRequete: String;

      { Début GetTobFourchette}
      begin
        sRequete := 'SELECT '+
                       'YTF_IDENTIFIANTYTA, '+
                       'YTF_IDENTIFIANTYTQ, '+
                       'YTF_IDENTIFIANT, '+
                       'YTF_FORFAIT, '+
                       'YTF_COUTLANCEMENT, '+
                       'YTF_REMISEMONTANT, '+
                       'YTF_PRIXBRUT, '+
                       'YTF_REMISE1, '+
                       'YTF_REMISE2, '+
                       'YTF_REMISE3, '+
                       'YTF_PRIXNET, '+
                       'YTF_CONDFORFAIT, '+
                       'YTF_CONDCOUTLANC, '+
                       'YTF_CONDREMISEM, '+
                       'YTF_CONDPRIXBRUT, '+
                       'YTF_CONDREMISE1, '+
                       'YTF_CONDREMISE2, '+
                       'YTF_CONDREMISE3, '+
                       'YTF_CONDPRIXNET '+
                    'FROM '+
                       'YTARIFSFOURCHETTE '+
                    'WHERE '+
                       '(YTF_IDENTIFIANTYTA='+IntToStr(TobAQui.Detail[iCptAqui].GetValue('YTA_IDENTIFIANT'))      +')'+' and '+
                       '(YTF_IDENTIFIANTYTQ='+IntToStr(TobSurquoi.Detail[iCptSurquoi].GetValue('YTQ_IDENTIFIANT'))+')'+' and '+
                       '(YTF_FOURCHETTEMINI<='+StrFPoint(Quoi(TobSurquoi.Detail[iCptSurquoi].GetValue('YTQ_FOURCHETTEDE'), TobContexte))+')'+ 'and' +
                       '(YTF_FOURCHETTEMAXI> '+StrFPoint(Quoi(TobSurquoi.Detail[iCptSurquoi].GetValue('YTQ_FOURCHETTEDE'), TobContexte))+')';
        Result := TobFourchette.LoadDetailDBFromSql('_TarifsFourchette_', sRequete);
      end;
      { Fin GetTobFourchette}


      { ------------------------------------------------------------------------------------------------------------
          Mise à jour du champ valeur et du champ Ou pour chaque donnée
      --------------------------------------------------------------------------------------------------------------}
      Procedure MaJ_Val_Ou( sChampval, sChampOu : string; nValeur : double);
      begin
          if (sLigneOuCommande=sOu_L) and (sOrigine[iPosOrigineSommerOu]=sOu_L) and (sOrigine[iPosOrigineRechercherOu]=sOu_L) then
          begin
            TobTarifs.PutValue(sChampVal ,nValeur);
            TobTarifs.PutValue(sChampOu  ,sOrigine);
          end
          else
            TobTarifs.PutValue(sChampOu, sOrigineMasque+sOrigine);
      end;

    { Début GetTobTarifsAvecFourchette}
    begin
      if GetTobFourchette then
      begin
//        wShowMeTheTob(TobFourchette, True);
        sOrigine     := CodifieOrigine;
        sMonnaie     := TobContexte.GetValue('DEVISE');
        sNatureAuxi  := TobContexte.GetValue('NATUREAUXI');
        nTauxMonnaie := ConvMonnaie(TobAQui.Detail[iCptAqui].GetValue('YTA_DEVISE'), sMonnaie, 1, True, TobContexte.GetValue('NATUREAUXI'), 'T', ' ');

        //Forfait
        if (TobTarifs.GetValue('FORFAIT_OU')='') and (TobFourchette.Detail[0].GetValue('YTF_CONDFORFAIT')=wTrue) then
          MaJ_Val_Ou('FORFAIT_VAL', 'FORFAIT_OU', ArrondirPrixMontant(TobFourchette.Detail[0].GetValue('YTF_FORFAIT')*nTauxMonnaie,sMonnaie,sNatureAuxi,'M','P'));

        //Cout Lancement
        if (TobTarifs.GetValue('COUTLANCEMENT_OU')='') and (TobFourchette.Detail[0].GetValue('YTF_CONDCOUTLANC')=wTrue) then
          MaJ_Val_Ou('COUTLANCEMENT_VAL', 'COUTLANCEMENT_OU', ArrondirPrixMontant(TobFourchette.Detail[0].GetValue('YTF_COUTLANCEMENT')*nTauxMonnaie,sMonnaie,sNatureAuxi,'M','P'));

        //Remise en montant
        if (TobTarifs.GetValue('REMISEMONTANT_OU')='') and (TobFourchette.Detail[0].GetValue('YTF_CONDREMISEM')=wTrue) then
          MaJ_Val_Ou('REMISEMONTANT_VAL', 'REMISEMONTANT_OU', ArrondirPrixMontant(TobFourchette.Detail[0].GetValue('YTF_REMISEMONTANT')*nTauxMonnaie,sMonnaie,sNatureAuxi,'M','P'));

        //Prix Brut  : Si aucune recherche a abouti sur un prix net ou un prix brut et que la recherche concerne un prix brut alors affectation
        if ((TobTarifs.GetValue('PRIXBRUT_OU')='') and (TobTarifs.GetValue('PRIXNET_OU')='')) and (TobFourchette.Detail[0].GetValue('YTF_CONDPRIXBRUT')=wTrue) then
          MaJ_Val_Ou('PRIXBRUT_VAL', 'PRIXBRUT_OU', ArrondirPrixMontant(TobFourchette.Detail[0].GetValue('YTF_PRIXBRUT')*nTauxMonnaie,sMonnaie,sNatureAuxi,'P','P'));

        //Prix Net : Si aucune recherche a abouti sur un prix brut ou un prix net et que la recherche concerne un prix net alors affectation
        if ((TobTarifs.GetValue('PRIXNET_OU')='') and (TobTarifs.GetValue('PRIXBRUT_OU')='')) and (TobFourchette.Detail[0].GetValue('YTF_CONDPRIXNET')=wTrue) then
          MaJ_Val_Ou('PRIXNET_VAL', 'PRIXNET_OU', ArrondirPrixMontant(TobFourchette.Detail[0].GetValue('YTF_PRIXNET')*nTauxMonnaie,sMonnaie,sNatureAuxi,'P','P'));

        //Création des les lignes Remises si la 1ère fois ou si la cascade remise est autorisée
        if    (
                    (TobFourchette.Detail[0].GetValue('YTF_CONDREMISE1')=wTrue)
                or  (TobFourchette.Detail[0].GetValue('YTF_CONDREMISE2')=wTrue)
                or  (TobFourchette.Detail[0].GetValue('YTF_CONDREMISE3')=wTrue)
              )
          and (CascadeRemisesAutorisee(sOrigine)) then
        begin

          //Sauvegarde des paramètres de Remises en Cascade
          ParametreCascadeRemises(sOrigine);

          //Remise 1/2/3
          TobRemises := CreationTobRemises( TobTarifs );

          //Remise N°1
          if (TobRemises.GetValue('REMISES_OU')='') and ((TobFourchette.Detail[0].GetValue('YTF_CONDREMISE1')=wTrue) or (TobFourchette.Detail[0].GetValue('YTF_CONDREMISE2')=wTrue) or (TobFourchette.Detail[0].GetValue('YTF_CONDREMISE3')=wTrue)) then
          begin
            if (sLigneOuCommande=sOu_L) and (sOrigine[iPosOrigineSommerOu]=sOu_L) and (sOrigine[iPosOrigineRechercherOu]=sOu_L) then
            begin
              TobRemises.PutValue('REMISE1_VAL', TobFourchette.Detail[0].GetValue('YTF_REMISE1'));
              TobRemises.PutValue('REMISE2_VAL', TobFourchette.Detail[0].GetValue('YTF_REMISE2'));
              TobRemises.PutValue('REMISE3_VAL', TobFourchette.Detail[0].GetValue('YTF_REMISE3'));
              TobRemises.PutValue('REMISES_OU' , sOrigine);
            end
            else
              TobRemises.PutValue('REMISES_OU' , sOrigineMasque+sOrigine);

            TobRemises.PutValue('REMISES_INTITULE', TobSurQuoi.Detail[iCptSurQuoi].GetValue('YTQ_LIBELLETARIF'));
          end;
        end;
      end;
    end;
    { Fin GetTobTarifsAvecFourchette }

  { Début GetTobTarifs}
  begin
    for iCptAQui:=0 to (TobAQui.Detail.Count-1) do
    begin
      for iCptSurQuoi:=0 to (TobSurQuoi.Detail.Count-1) do
      begin
        if (TobSurQuoi.Detail[iCptSurQuoi].GetValue('YTQ_IDENTIFIANTYTA')=TobAQui.Detail[iCptAQui].GetValue('YTA_IDENTIFIANT')) then
        begin
          // Conditions par fourchette
          if (TobSurQuoi.Detail[iCptSurquoi].GetValue('YTQ_PARTRANCHE')=wTrue) then
            GetTobTarifsAvecFourchette
          else
            GetTobTarifsSansFourchette;
        end;
      end;
    end;
  end;
  { Fin GetTobTarifs}

  {----------------------------------------------------------------------------
    Sur la TobAQui
    suppression si les conditions particuliaires d'application ne sont pas remplies
  -----------------------------------------------------------------------------}
  procedure wFreeTobAqui;
  var
    iCptAQui : integer;
  begin
    for iCptAqui:=(TobAQui.Detail.Count-1) downto 0 do
    begin
      if (not VerifConditionsApplications(sAppel, TobAqui.Detail[iCptAqui].GetValue('YTA_BLOCNOTE'), TobTiers, TobArticle, TobPiece, TobLigne)) then
        TobAQui.Detail[iCptAqui].Free;
    end;
  end;

  {----------------------------------------------------------------------------
    Sur la TobSurQuoi
    suppression si les conditions particuliaires d'application ne sont pas remplies
  -----------------------------------------------------------------------------}
  procedure wFreeTobSurQuoi;
  var
    iCptSurQuoi : integer;
  begin
    for iCptsurQuoi:=(TobSurQuoi.Detail.Count-1) downto 0 do
    begin
      if (not VerifConditionsApplications(sAppel, Tobsurquoi.Detail[iCptsurQuoi].GetValue('YTQ_BLOCNOTE'), TobTiers, TobArticle, TobPiece, TobLigne)) then
        TobSurQuoi.Detail[iCptsurQuoi].Free;
    end;
  end;

{ Début RechercheTarifs}
begin
  PutToDebugLog('RechercheTarifs', true);

  TobAqui           := Tob.Create('_TarifsAQui_'           ,nil ,-1);
  TobSurQuoi        := Tob.Create('_TarifsSurQuoi_'        ,nil ,-1);
  TobFourchette     := Tob.Create('_TarifsFourchette_'     ,nil ,-1);

  sFonctionnalite         := TobContexte.GetValue('FONCTIONNALITE');
  iPartiel                := GetParamSoc('SO_TARIFPARTIEL');
  lConditionsApplications := GetParamSoc('SO_TARIFCONDAPPL');

  sParametre := GetParametresFromYFO(sFonctionnalite);

  if (sAppel='PIECE') then
    sLigneOuCommande := sOu_P
  else
    sLigneOuCommande := sOu_L;

  try
    // Ajout des champs Supplémentaires dans Tobtarifs
    TobTarifsAddChampsSupValeur( TobTarifs);

    //Si appel depuis le menu et non depuis les pièces ajout des informations nécessaire dans la TobContexte
    if (lComplement) then
    begin
      AjoutComplementTiers;
      AjoutComplementArticle;
    end;
    AjoutComplementCommercial;

    AjoutComplementDepot;
    AjoutComplementAffaire;
    AjoutComplementRessource;

//    if (TobContexte <>nil) then  wShowMeTheTob(TobContexte, True);

    PutToDebugLog('RechercheTarifs', false, 'Avant GetTobAqui');
    //Contexte de A qui
    if GetTobAqui then
    begin
      PutToDebugLog('RechercheTarifs', false, 'Après GetTobAqui');

      //Vérification des conditions particulières d'application
      if lConditionsApplications then wFreeTobAQui;
//      if (TobAQui<>nil) then wShowMeTheTob(TobAQui, True);

      //Contexte de Sur Quoi
      PutToDebugLog('RechercheTarifs', false, 'Avant GetTobSurQuoi');
      if (TobAqui.Detail.Count<>0) and GetTobSurQuoi then
      begin
        PutToDebugLog('RechercheTarifs', false, 'Après GetTobSurQuoi');
        //Vérification des conditions particulières d'application
        if lConditionsApplications then wFreeTobSurQuoi;
//        if (TobsurQuoi  <>nil) then   wShowMeTheTob(TobSurQuoi, True);

        if (TobSurQuoi.Detail.Count<>0) then
        begin
          //Contexte de Fourchette
          PutToDebugLog('RechercheTarifs', false, 'Avant GetTobTarifs');
          GetTobTarifs;
          PutToDebugLog('RechercheTarifs', false, 'Après GetTobTarifs');
//          if (TobTarifs   <>nil) then   wShowMeTheTob(TobTarifs, True);
        end;
      end;
    end;
    Result := TobTarifs;
  finally
    TobFourchette.Free;
    TobSurQuoi.Free;
    TobAQui.Free;
  end;
  PutToDebugLog('RechercheTarifs', false, 'Fin');
end;

{------------------------------------------------------------------------------------
   Affichage Résultat de la recherche tarifaire : Valeur + Origine
-------------------------------------------------------------------------------------}
procedure AffichageOrigineSystemeTarifaire(sAppel: string; TobContexte, TobTarifs, TobAffichage: tob; GAffichage : tHGrid; MBLOCNOTE : tMemo);
//sAppel = RECHERCHE    si menu de Recherche  (Sont renseignées à l'appel la TobContexte et la TobTarifs : Génération de la TobAffichage)
//       = PIECE/LIGNE  si menu consultation origine tarifaire depuis les pièces (Sont Renseignées la TobContexte et la TobAffichage : Pas besoin de générer la Tobaffichage)
var
  GAffichageColonnes, sListeChamps, sChamps, sFonctionnalite, sParametre, sColonnes : string;

  sLibForfait, sLibCoutLancement, sLibRemiseEnMontant, sLibPrixBrut, sLibRemises, sLibPrixNet, sLibRemiseTotale, sLibValeurLigne : string;

  iCpt, iSensTaux, iColonnes : integer;
  nRemiseTotale, nPrixNet    : double;

  iLargeur1ereCellule, iLargeurCellules   : integer;

  {----------------------------------------------------------------------------
    Création une Tob Flle  TobAffichage
  -----------------------------------------------------------------------------}
  function CreationTobAffichage( sQuoi:string; nValeur:double; sOu:string): Tob;

    {---------------------------------------------------------------------------
      Afficher l'image en fonction
          * Recherche ou Somme
          * Manuellement
          * Avec Code complet
          * Avec code Partiel
          * Sans : n'a pas participé à la recher
    ----------------------------------------------------------------------------}
    Function ChoixImage( sOu : string; iPosOrigine: integer) : string;
    var
      sOuSomme, sOuJumelle : string;
      lSomme  , lJumelle   : boolean;

      {---------------------------------------------------------------------------
        Afficher l'image :
          * Manuellement
          * Avec Code complet
          * Avec code Partiel
          * Sans : n'a pas participé à la recher
      ----------------------------------------------------------------------------}
      function ChoixImageMAPS(sOu, sImageMain, sImagePartielle, sImagePleine, sImageVide: string) : string;
      begin
        if      (sOu[iPosOrigine]='M') then
          Result := sImageMain
        else if (sOu[iPosOrigine]='P') then
          Result := sImagePartielle
        else if (sOu[iPosOrigine]<>'.') then
          Result := sImagePleine
        else
          Result := sImageVide;
      end;

    begin {ChoixImage}
      sOuJumelle := copy(sOu,           1, iTailleOu);
      sOuSomme   := copy(sOu, iTailleOu+1, iTailleOu);

      lSomme     := (sOuSomme  <>sOrigineMasque) and (sOuSomme  <>'');
      lJumelle   := (sOuJumelle<>sOrigineMasque) and (sOuJumelle<>'');

      if      (not lJumelle) and (not lSomme) then
        Result := sImageVide
      else if (    lJumelle) and (not lSomme) then
        Result := ChoixImageMAPS(sOuJumelle, sImageMain, sImageJumelleVide, sImageJumellePleine, sImageVide)
      else if (not lJumelle) and (    lSomme) then
        Result := ChoixImageMAPS(sOuSomme  , sImageMain, sImageSommeVide  , sImageSommePleine  , sImageVide)
      else if (    lJumelle) and (    lSomme) then
      begin
        if      (sOuJumelle[iPosOrigine]=sOuSomme[iPosOrigine]) then
          Result := ChoixImageMAPS(sOuSomme  , sImageMain, sImageSommeJumelleVide, sImageSommeJumellePleine, sImageVide)
        else if (sOuJumelle[iPosOrigine]<>'.') then
          Result := ChoixImageMAPS(sOuSomme  , sImageMain, sImageJumelleVide     , sImageJumellePleine     , sImageVide)
        else if (sOuJumelle[iPosOrigine]<>'.') then
          Result := ChoixImageMAPS(sOuSomme  , sImageMain, sImageSommeVide       , sImageSommePleine       , sImageVide)
        else
          Result := sImageVide;
      end;
    end;  {ChoixImage}

  begin
    Result   := Tob.Create('_Affichage_',TobAffichage,-1);
    Result.AddChampSupValeur('QUOI'          ,sQuoi                                ,False);
    Result.AddChampSupValeur('VALEUR'        ,nValeur                              ,False);
    Result.AddChampSupValeur('OU'            ,sOu                                  ,False);

    { si l'origine = 'M' : Saisie manuelle donc aucun point ne sera visualisé }
    if (sOu='M') then   sOu := StringReplace(sOrigineMasque, '.', 'M', [rfReplaceAll]);

    Result.AddChampSupValeur('DEPOT'         ,ChoixImage(sOu, iPosOrigineDepot         ),False);
    Result.AddChampSupValeur('TARIFDEPOT'    ,ChoixImage(sOu, iPosOrigineTarifDepot    ),False);
    Result.AddChampSupValeur('DEVISE'        ,ChoixImage(sOu, iPosOrigineDevise        ),False);
    Result.AddChampSupValeur('TARIFSPECIAL'  ,ChoixImage(sOu, iPosOrigineSpecial       ),False);
    Result.AddChampSupValeur('COMMERCIAL'    ,ChoixImage(sOu, iPosOrigineCommercial    ),False);
    Result.AddChampSupValeur('TYPECOMMERCIAL',ChoixImage(sOu, iPosOrigineTypeCommercial),False);
    Result.AddChampSupValeur('TIERS'         ,ChoixImage(sOu, iPosOrigineTiers         ),False);
    Result.AddChampSupValeur('TARIFTIERS'    ,ChoixImage(sOu, iPosOrigineTarifTiers    ),False);
    Result.AddChampSupValeur('AFFAIRE'       ,ChoixImage(sOu, iPosOrigineAffaire       ),False);
    Result.AddChampSupValeur('TARIFAFFAIRE'  ,ChoixImage(sOu, iPosOrigineTarifAffaire  ),False);
    Result.AddChampSupValeur('POLE'          ,ChoixImage(sOu, iPosOriginePole          ),False);
    Result.AddChampSupValeur('ATELIER'       ,ChoixImage(sOu, iPosOrigineAtelier       ),False);
    Result.AddChampSupValeur('GROUPE'        ,ChoixImage(sOu, iPosOrigineGroupe        ),False);
    Result.AddChampSupValeur('JOURSEMAINE'   ,ChoixImage(sOu, iPosOrigineJoursemaine   ),False);
    Result.AddChampSupValeur('HEUREJOUR'     ,ChoixImage(sOu, iPosOrigineHeureJour     ),False);
    Result.AddChampSupValeur('LIEUCONSO'     ,ChoixImage(sOu, iPosOrigineLieuConso     ),False);

    Result.AddChampSupValeur('ARTICLE'       ,ChoixImage(sOu, iPosOrigineArticle       ),False);
    Result.AddChampSupValeur('TARIFARTICLE'  ,ChoixImage(sOu, iPosOrigineTarifArticle  ),False);
    Result.AddChampSupValeur('RESSOURCE'     ,ChoixImage(sOu, iPosOrigineRessource     ),False);
    Result.AddChampSupValeur('TARIFRESSOURCE',ChoixImage(sOu, iPosOrigineTarifRessource),False);
    Result.AddChampSupValeur('MARQUE'        ,ChoixImage(sOu, iPosOrigineMarque        ),False);
    Result.AddChampSupValeur('CIRCUIT'       ,ChoixImage(sOu, iPosOrigineCircuit       ),False);
    Result.AddChampSupValeur('PHASE'         ,ChoixImage(sOu, iPosOriginePhase         ),False);
    Result.AddChampSupValeur('TRAITEMENT'    ,ChoixImage(sOu, iPosOrigineTraitement    ),False);
  end;
begin
  sFonctionnalite := TobContexte.GetValue('FONCTIONNALITE');
  iSensTaux := SensTauxOuRemise(sFonctionnalite);

  sParametre       := GetParametresFromYFO(sFonctionnalite);

  sLibForfait         := LibelleParFonctionnalite(sFonctionnalite, 'FORFAIT');
  sLibCoutLancement   := LibelleParFonctionnalite(sFonctionnalite, 'COUTLANCEMENT');
  sLibRemiseEnMontant := LibelleParFonctionnalite(sFonctionnalite, 'REMISEMONTANT');
  sLibPrixBrut        := LibelleParFonctionnalite(sFonctionnalite, 'PRIXBRUT');
  sLibRemises         := LibelleParFonctionnalite(sFonctionnalite, 'REMISE');
  sLibPrixNet         := LibelleParFonctionnalite(sFonctionnalite, 'PRIXNET');
  sLibRemiseTotale    := LibelleParFonctionnalite(sFonctionnalite, 'REMISETOTALE');
  sLibValeurLigne     := LibelleParFonctionnalite(sFonctionnalite, 'MONTANTTOTAL');

  //Création de la Tob pour l'affichage de l'origine  de chaques composantes du système tarifaire
  CreationTobAffichage(sLibForfait        ,TobTarifs.GetValue('FORFAIT_VAL'      ),TobTarifs.GetValue('FORFAIT_OU'      ));
  CreationTobAffichage(sLibCoutLancement  ,TobTarifs.GetValue('COUTLANCEMENT_VAL'),TobTarifs.GetValue('COUTLANCEMENT_OU'));
  CreationTobAffichage(sLibRemiseEnMontant,TobTarifs.GetValue('REMISEMONTANT_VAL'),TobTarifs.GetValue('REMISEMONTANT_OU'));
  CreationTobAffichage(sLibPrixBrut       ,TobTarifs.GetValue('PRIXBRUT_VAL'     ),TobTarifs.GetValue('PRIXBRUT_OU'     ));

  //Calcul de la remise globale en coefficient
  nRemiseTotale := 1;
  for iCpt:=0 to (TobTarifs.Detail.Count-1) do
  begin
    CreationTobAffichage(iif(TobTarifs.Detail[iCpt].GetValue('REMISES_INTITULE')='',sLibRemises,String(TobTarifs.Detail[iCpt].GetValue('REMISES_INTITULE'))),TobTarifs.Detail[iCpt].GetValue('REMISE1_VAL'),TobTarifs.Detail[iCpt].GetValue('REMISES_OU'));
    if (TobTarifs.Detail[iCpt].GetValue('REMISE2_VAL')<>0) then
      CreationTobAffichage(iif(TobTarifs.Detail[iCpt].GetValue('REMISES_INTITULE')='',sLibRemises,String(TobTarifs.Detail[iCpt].GetValue('REMISES_INTITULE'))),TobTarifs.Detail[iCpt].GetValue('REMISE2_VAL'),TobTarifs.Detail[iCpt].GetValue('REMISES_OU'));
    if (TobTarifs.Detail[iCpt].GetValue('REMISE3_VAL')<>0) then
      CreationTobAffichage(iif(TobTarifs.Detail[iCpt].GetValue('REMISES_INTITULE')='',sLibRemises,String(TobTarifs.Detail[iCpt].GetValue('REMISES_INTITULE'))),TobTarifs.Detail[iCpt].GetValue('REMISE3_VAL'),TobTarifs.Detail[iCpt].GetValue('REMISES_OU'));
    nRemiseTotale := nRemiseTotale*(1-TobTarifs.Detail[iCpt].GetValue('REMISE1_VAL')/100)*(1-TobTarifs.Detail[iCpt].GetValue('REMISE2_VAL')/100)*(1-TobTarifs.Detail[iCpt].GetValue('REMISE3_VAL')/100);
  end;

  nRemiseTotale:=(1-nRemiseTotale)*100;   //Transformation du Coefficient de remise en pourcentage de remise
  CreationTobAffichage(sLibRemiseTotale,nRemiseTotale,'');
  if (TobTarifs.GetValue('PRIXNET_OU')<>'') then
  begin
    nPrixNet := TobTarifs.GetValue('PRIXNET_VAL');
    CreationTobAffichage(sLibPrixNet,TobTarifs.GetValue('PRIXNET_VAL'),TobTarifs.GetValue('PRIXNET_OU'))
  end
  else
  begin
    nPrixNet := CalculPrixNet(TobTarifs.GetValue('PRIXBRUT_VAL'), nRemiseTotale, 0, 0, iSensTaux, TobContexte.GetValue('DEVISE'), TobContexte.GetValue('NATUREAUXI'), 'P', 'T');
    CreationTobAffichage(sLibPrixNet,nPrixNet,TobTarifs.GetValue('PRIXNET_OU'));
  end;
  CreationTobAffichage(sLibValeurLigne,CalculValeurLigne(TobTarifs.GetValue('FORFAIT_VAL'), TobTarifs.GetValue('COUTLANCEMENT_VAL'), TobTarifs.GetValue('REMISEMONTANT_VAL'), nPrixNet, TobContexte.GetValue('QUANTITE'), TobContexte.GetValue('DEVISE'), TobContexte.GetValue('NATUREAUXI'), 'M', 'P'),'');
//  if (TobAffichage<>nil) then wShowMeTheTob(TobAffichage, True);

  sColonnes := 'Quoi;Valeur;';
  iColonnes := 0;
  Gaffichage.Cells[ iColonnes, 0] := RechDom('YFONCTIONNALITES', sFonctionnalite, false); iColonnes := iColonnes + 1;
  Gaffichage.Cells[ iColonnes, 0] := TraduireMemoire('Valeur')                          ; iColonnes := iColonnes + 1;

  if (sParametre[iPosOrigineDepot         ]=wTrue) then begin sColonnes := sColonnes + 'Depot;'         ; Gaffichage.Cells[ iColonnes, 0] := TraduireMemoire('dépôt')     ; iColonnes := iColonnes + 1; end;
  if (sParametre[iPosOrigineTarifDepot    ]=wTrue) then begin sColonnes := sColonnes + 'TarifDepot;'    ; Gaffichage.Cells[ iColonnes, 0] := TraduireMemoire('Famille')   ; iColonnes := iColonnes + 1; end;
  if (sParametre[iPosOrigineDevise        ]=wTrue) then begin sColonnes := sColonnes + 'Devise;'        ; Gaffichage.Cells[ iColonnes, 0] := TraduireMemoire('Devise')    ; iColonnes := iColonnes + 1; end;
  if (sParametre[iPosOrigineSpecial       ]=wTrue) then begin sColonnes := sColonnes + 'TarifSpecial;'  ; Gaffichage.Cells[ iColonnes, 0] := TraduireMemoire('Spécial')   ; iColonnes := iColonnes + 1; end;
  if (sParametre[iPosOrigineCommercial    ]=wTrue) then begin sColonnes := sColonnes + 'Commercial;'    ; Gaffichage.Cells[ iColonnes, 0] := TraduireMemoire('Commercial'); iColonnes := iColonnes + 1; end;
  if (sParametre[iPosOrigineTypeCommercial]=wTrue) then begin sColonnes := sColonnes + 'TypeCommercial;'; Gaffichage.Cells[ iColonnes, 0] := TraduireMemoire('Famille')   ; iColonnes := iColonnes + 1; end;
  if (sParametre[iPosOrigineTiers         ]=wTrue) then begin sColonnes := sColonnes + 'Tiers;'         ; Gaffichage.Cells[ iColonnes, 0] := TraduireMemoire('Tiers')     ; iColonnes := iColonnes + 1; end;
  if (sParametre[iPosOrigineTarifTiers    ]=wTrue) then begin sColonnes := sColonnes + 'TarifTiers;'    ; Gaffichage.Cells[ iColonnes, 0] := TraduireMemoire('Famille')   ; iColonnes := iColonnes + 1; end;
  if (sParametre[iPosOrigineAffaire       ]=wTrue) then begin sColonnes := sColonnes + 'Affaire;'       ; Gaffichage.Cells[ iColonnes, 0] := TraduireMemoire('Affaire')   ; iColonnes := iColonnes + 1; end;
  if (sParametre[iPosOrigineTarifAffaire  ]=wTrue) then begin sColonnes := sColonnes + 'TarifAffaire;'  ; Gaffichage.Cells[ iColonnes, 0] := TraduireMemoire('Famille')   ; iColonnes := iColonnes + 1; end;
  if (sParametre[iPosOriginePole          ]=wTrue) then begin sColonnes := sColonnes + 'Pole;'          ; Gaffichage.Cells[ iColonnes, 0] := TraduireMemoire('Pôle')      ; iColonnes := iColonnes + 1; end;
  if (sParametre[iPosOrigineAtelier       ]=wTrue) then begin sColonnes := sColonnes + 'Atelier;'       ; Gaffichage.Cells[ iColonnes, 0] := TraduireMemoire('Atelier')   ; iColonnes := iColonnes + 1; end;
  if (sParametre[iPosOrigineGroupe        ]=wTrue) then begin sColonnes := sColonnes + 'Groupe;'        ; Gaffichage.Cells[ iColonnes, 0] := TraduireMemoire('Groupe')    ; iColonnes := iColonnes + 1; end;
  if (sParametre[iPosOrigineJourSemaine   ]=wTrue) then begin sColonnes := sColonnes + 'Jour;'          ; Gaffichage.Cells[ iColonnes, 0] := TraduireMemoire('Jour')      ; iColonnes := iColonnes + 1; end;
  if (sParametre[iPosOrigineHeureJour     ]=wTrue) then begin sColonnes := sColonnes + 'Horaire;'       ; Gaffichage.Cells[ iColonnes, 0] := TraduireMemoire('Horaire')   ; iColonnes := iColonnes + 1; end;
  if (sParametre[iPosOrigineLieuConso     ]=wTrue) then begin sColonnes := sColonnes + 'Lieu;'          ; Gaffichage.Cells[ iColonnes, 0] := TraduireMemoire('Lieu')      ; iColonnes := iColonnes + 1; end;
  if (sParametre[iPosOrigineArticle       ]=wTrue) then begin sColonnes := sColonnes + 'Article;'       ; Gaffichage.Cells[ iColonnes, 0] := TraduireMemoire('Article')   ; iColonnes := iColonnes + 1; end;
  if (sParametre[iPosOrigineTarifArticle  ]=wTrue) then begin sColonnes := sColonnes + 'TarifArticle;'  ; Gaffichage.Cells[ iColonnes, 0] := TraduireMemoire('Famille')   ; iColonnes := iColonnes + 1; end;
  if (sParametre[iPosOrigineRessource     ]=wTrue) then begin sColonnes := sColonnes + 'Ressource;'     ; Gaffichage.Cells[ iColonnes, 0] := TraduireMemoire('Ressource') ; iColonnes := iColonnes + 1; end;
  if (sParametre[iPosOrigineTarifRessource]=wTrue) then begin sColonnes := sColonnes + 'TarifRessource;'; Gaffichage.Cells[ iColonnes, 0] := TraduireMemoire('Famille')   ; iColonnes := iColonnes + 1; end;
  if (sParametre[iPosOrigineMarque        ]=wTrue) then begin sColonnes := sColonnes + 'Marque;'        ; Gaffichage.Cells[ iColonnes, 0] := TraduireMemoire('Marque')    ; iColonnes := iColonnes + 1; end;
  if (sParametre[iPosOrigineCircuit       ]=wTrue) then begin sColonnes := sColonnes + 'Circuit;'       ; Gaffichage.Cells[ iColonnes, 0] := TraduireMemoire('Circuit')   ; iColonnes := iColonnes + 1; end;
  if (sParametre[iPosOriginePhase         ]=wTrue) then begin sColonnes := sColonnes + 'Phase;'         ; Gaffichage.Cells[ iColonnes, 0] := TraduireMemoire('Phase')     ; iColonnes := iColonnes + 1; end;
  if (sParametre[iPosOrigineTraitement    ]=wTrue) then begin sColonnes := sColonnes + 'Traitement;'    ; Gaffichage.Cells[ iColonnes, 0] := TraduireMemoire('Traitement'); iColonnes := iColonnes + 1; end;

  GAffichage.ColCount := iColonnes;
  GAffichageColonnes  := sColonnes;

  sListeChamps := GAffichageColonnes;

  iLargeurCellules := Trunc(wDivise(GAffichage.ClientWidth, GAffichage.ColCount + 1))-Gaffichage.GridLineWidth;
  if (iLargeurCellules<100) then
  begin
    iLargeurCellules    := 50;
    iLargeur1ereCellule := 100;
  end
  else
  begin
    iLargeur1ereCellule := (GAffichage.ClientWidth - ((iLargeurCellules+1)*(GAffichage.ColCount-1)));
  end;

  for iCpt := 0 to GAffichage.ColCount-1 do
  begin
    sChamps := ReadTokenSt(sListeChamps) ;
    if (sChamps = 'Quoi') then
    begin
      Gaffichage.ColWidths[iCpt] := iLargeur1ereCellule;
    end
    else if      (sChamps = 'Valeur') then
    begin
      GAffichage.ColAligns[iCpt]    := taRightJustify;
      GAffichage.ColFormats[iCpt]   := '##0.000';
      Gaffichage.ColWidths[iCpt] := iLargeurCellules;
    end
    else
    begin
      Gaffichage.ColWidths[iCpt] := iLargeurCellules;
    end;
  end;

  //Chargement de la grille
  TobAffichage.PutGridDetail(GAffichage, False, True, GAffichageColonnes, True);

  //Ajout des autres commentaires
  AffichageCommentaireTarif(sParametre, TobContexte, TobAffichage, GAffichage, MBLOCNOTE);
end;

{-----------------------------------------------------------------------------------
   Affiche le commentaire du tarif contenant le contexte et ...
------------------------------------------------------------------------------------}
procedure AffichageCommentaireTarif(sParametre : string; TobContexte, TobAffichage : tob; GAffichage: thGrid;  MBLOCNOTE : tMemo);
var
  sOu : string;
begin
  if (GAffichage.Row<=TobAffichage.Detail.Count) then
  begin
    sOu :=TobAffichage.Detail[GAffichage.Row-1].GetValue('OU');
    MBLOCNOTE.Clear;
    MBLOCNOTE.Lines.Add(TraduireMemoire('Recherche des ')+RechDom('YFONCTIONNALITES',TobContexte.GetValue('FONCTIONNALITE'),False)+TraduireMemoire(' dans le contexte suivant :'));

    MBLOCNOTE.Lines.Add('');
    if (sParametre[iPosOrigineDepot]=wTrue) and (TobContexte.GetValue('DEPOT')<>'') then
      MBLOCNOTE.Lines.Add('    · '+TraduireMemoire('pour le dépôt ')+String(TobContexte.GetValue('DEPOT')));

    if (sParametre[iPosOrigineSpecial]=wTrue) and (TobContexte.GetValue('TARIFSPECIAL')<>'') then
      MBLOCNOTE.Lines.Add('    · '+TraduireMemoire('pour les conditions spéciales ')+string(TobContexte.GetValue('TARIFSPECIAL')));

    if (sParametre[iPosOrigineCommercial]=wTrue) and (TobContexte.GetValue('COMMERCIAL')<>'') then
      MBLOCNOTE.Lines.Add('    · '+TraduireMemoire('pour le commercial ') +WGetSqlFieldValue('GCL_LIBELLE', 'COMMERCIAL', 'GCL_COMMERCIAL="'+TobContexte.GetValue('COMMERCIAL')+'"')+' ( '+TobContexte.GetValue('COMMERCIAL')+' )'+iif(TobContexte.GetValue('TYPECOMMERCIAL')<>'',TraduireMemoire(' appartenant au type commercial ')+RechDom('GCTYPECOMMERCIAL',TobContexte.GetValue('TYPECOMMERCIAL'),False),''));

    if (sParametre[iPosOrigineTiers]=wTrue) and (TobContexte.GetValue('TIERS')<>'') and (Pos(TobContexte.GetValue('FONCTIONNALITE'),sTarifFournisseur+sTarifSousTraitantAchat+sTarifSousTraitantPhase)>0) then
      MBLOCNOTE.Lines.Add('    · '+TraduireMemoire('pour le fournisseur ')+WGetSqlFieldValue('T_LIBELLE', 'TIERS', 'T_NATUREAUXI="'+TobContexte.GetValue('NATUREAUXI')+'" and T_TIERS="'+TobContexte.GetValue('TIERS')+'"')+' ( '+TobContexte.GetValue('TIERS')+' )'+iif(TobContexte.GetValue('TARIFTIERS')='','',TraduireMemoire(' appartenant à la famille tarif ')+RechDom('TTTARIFFOURNISSEUR',TobContexte.GetValue('TARIFTIERS'),False)));

    if (sParametre[iPosOrigineTiers]=wTrue) and (TobContexte.GetValue('TIERS')<>'') and (Pos(TobContexte.GetValue('FONCTIONNALITE'),sTarifClient+sCommissionClient)>0) then
      MBLOCNOTE.Lines.Add('    · '+TraduireMemoire('pour le client ')+WGetSqlFieldValue('T_LIBELLE', 'TIERS', 'T_NATUREAUXI="'+TobContexte.GetValue('NATUREAUXI')+'" and T_TIERS="'+TobContexte.GetValue('TIERS')+'"')+' ( '+TobContexte.GetValue('TIERS')+' )'+iif((TobContexte.GetValue('TARIFTIERS')='') or (Pos(TobContexte.GetValue('FONCTIONNALITE'),sCommissionClient)<>0),'',TraduireMemoire(' appartenant à la famille tarif ')+RechDom('TTTARIFCLIENT',TobContexte.GetValue('TARIFTIERS'),False)));

    if (sParametre[iPosOrigineAffaire]=wTrue) and (TobContexte.GetValue('AFFAIRE')<>'') then
      MBLOCNOTE.Lines.Add('    · '+TraduireMemoire('pour l''affaire ')+WGetSqlFieldValue('AFF_LIBELLE', 'AFFAIRE', 'AFF_AFFAIRE="'+TobContexte.GetValue('AFFAIRE')+'"')+' ( '+TobContexte.GetValue('AFFAIRE')+' )');

    if (sParametre[iPosOriginePole]=wTrue) and (TobContexte.GetValue('POLE')<>'') then
      MBLOCNOTE.Lines.Add('    · '+TraduireMemoire('pour le pôle ')+wGetSqlFieldValue('QPO_POLELIB', 'QPOLE', 'QPO_CTX="0" and QPO_POLE="'+TobContexte.GetValue('YTA_POLE')+'"')+' ( '+TobContexte.GetValue('POLE')+' )');

    if (sParametre[iPosOrigineAtelier]=wTrue) and (TobContexte.GetValue('PHASE')<>'') then
      MBLOCNOTE.Lines.Add('    · '+TraduireMemoire('pour l''atelier ')+wGetSqlFieldValue('QSI_SITELIB', 'QSITE', 'QSI_CTX="0" and QSI_POLE="'+TobContexte.GetValue('YTA_POLE')+'"'+' and QSI_SITE="'+TobContexte.GetValue('YTA_SITE')+'"')+' ( '+TobContexte.GetValue('ATELIER')+' )');

    if (sParametre[iPosOrigineGroupe]=wTrue) and (TobContexte.GetValue('GROUPE')<>'') then
      MBLOCNOTE.Lines.Add('    · '+TraduireMemoire('pour le groupe ')+wGetSqlFieldValue('QGR_GROUPELIB', 'QGROUPE', 'QGR_CTX="0" and QGR_SITE="'+TobContexte.GetValue('YTA_SITE')+'"'+' and QGR_GRP="' +TobContexte.GetValue('YTA_GRP')+'"')+' ( '+TobContexte.GetValue('GROUPE')+' )');

    if (sParametre[iPosOrigineJourSemaine]=wTrue) and (TobContexte.GetValue('JOURSEMAINE')<>'') then
      MBLOCNOTE.Lines.Add('    · '+TraduireMemoire('pour la période en jour ')+RechDom('YTARIFJOURSEMAINE', TobContexte.GetValue('JOURSEMAINE'),False)+' ( '+TobContexte.GetValue('JOURSEMAINE')+' )');

    if (sParametre[iPosOrigineHeureJour]=wTrue) and (TobContexte.GetValue('HEUREJOUR')<>'') then
      MBLOCNOTE.Lines.Add('    · '+TraduireMemoire('pour la plage horaire ')+RechDom('YTARIFHEUREJOUR', TobContexte.GetValue('HEUREJOUR'),False)+' ( '+TobContexte.GetValue('HEUREJOUR')+' )');

    if (sParametre[iPosOrigineLieuConso]=wTrue) and (TobContexte.GetValue('LIEUCONSO')<>'') then
      MBLOCNOTE.Lines.Add('    · '+TraduireMemoire('pour le lieu de consommation ')+RechDom('YTARIFLIEUCONSO', TobContexte.GetValue('LIEUCONSO'),False)+' ( '+TobContexte.GetValue('LIEUCONSO')+' )');

    if (sParametre[iPosOrigineArticle]=wTrue) and (TobContexte.GetValue('ARTICLE')<>'') then
      MBLOCNOTE.Lines.Add('    · '+TraduireMemoire('pour l''article ')+wGetSqlFieldValue('GA_LIBELLE', 'ARTICLE', 'GA_ARTICLE="'+TobContexte.GetValue('ARTICLE')+'"')+' ( '+TobContexte.GetValue('ARTICLE')+' )'+iif((TobContexte.GetValue('TARIFARTICLE')='') or (Pos(TobContexte.GetValue('FONCTIONNALITE'),sCommissionClient)<>0) ,'',TraduireMemoire(' appartenant à la famille tarif ')+RechDom('GCTARIFARTICLE',TobContexte.GetValue('TARIFARTICLE'),False)));

    if (sParametre[iPosOrigineRessource]=wTrue) and (TobContexte.GetValue('RESSOURCE')<>'') then
      MBLOCNOTE.Lines.Add('    · '+TraduireMemoire('pour la ressource ')+wGetSqlFieldValue('ARS_LIBELLE', 'RESSOURCE', 'ARS_RESSOURCE="'+TobContexte.GetValue('RESSOURCE')+'"')+' ( '+TobContexte.GetValue('RESSOURCE')+' )'+TraduireMemoire(' appartenant à la famille tarif ')+RechDom('AFTTYPERESSOURCE',TobContexte.GetValue('TARIFRESSOURCE'),False));

    if (sParametre[iPosOrigineCircuit]=wTrue) and (TobContexte.GetValue('CIRCUIT')<>'') then
      MBLOCNOTE.Lines.Add('    · '+TraduireMemoire('pour le circuit ')+TobContexte.GetValue('CIRCUIT'));

    if (sParametre[iPosOriginePhase]=wTrue) and (TobContexte.GetValue('PHASE')<>'') then
      MBLOCNOTE.Lines.Add('    · '+TraduireMemoire('pour la phase ')+TobContexte.GetValue('PHASE'));

    if (sParametre[iPosOriginePhase]=wTrue) and (TobContexte.GetValue('TRAITEMENT')<>'') then
      MBLOCNOTE.Lines.Add('    · '+TraduireMemoire('pour le traitement ')+TobContexte.GetValue('TRAITEMENT'));


    MBLOCNOTE.Lines.Add('');
    if (sOu<>'') and (pos(sOu[iPosOrigineDateDeReference],'. ')=0) then MBLOCNOTE.Lines.Add(TraduireMemoire('Le contrôle des dates de validité se fait sur ')+iif(copy(sOu,iPosOrigineDateDeReference,1)='C',TraduireMemoire('la date de commande'),TraduireMemoire('le délai de livraison prévu')));
    if (sOu<>'') and (pos(sOu[iPosOrigineTTCouHT        ],'. ')=0) then MBLOCNOTE.Lines.Add(TraduireMemoire('Les conditions sont exprimées en ')+iif(sOu[iPosOrigineTTCouHT]='T','T.T.C.','H.T.')+TraduireMemoire(' dans l''unité ')+TobContexte.GetValue('QUALIFUNITEVTE')+TraduireMemoire(' pour une quantité de ')+IntToStr(TobContexte.GetValue('PRIXPOURQTE'))+TraduireMemoire(' en ')+iif(sOu[iPosOrigineDevise]='D',String(TobContexte.GetValue('DEVISE')),V_PGI.DEvisePivot));

    MBLOCNOTE.Lines.Add('');
    if (sOu<>'') and (pos(sOu[iPosOrigineRechercherOu ],'. ')=0) then MBLOCNOTE.Lines.Add(TraduireMemoire('Rechercher ')            +RechDom('YTARIFSRECHERCHE',sOu[iPosOrigineRechercherOu], False));
    if (sOu<>'') and (pos(sOu[iPosOrigineSommerOu     ],'. ')=0) then MBLOCNOTE.Lines.Add(TraduireMemoire('Sommer ')                +RechDom('YTARIFSOU'       ,sOu[iPosOrigineSommerOu    ], False));
    if (sOu<>'') and (pos(sOu[iPosOrigineSommerQuoi   ],'. ')=0) then MBLOCNOTE.Lines.Add(TraduireMemoire('Sur des fourchettes de ')+RechDom('YTARIFSQUOI'     ,sOu[iPosOrigineSommerQuoi  ], True ));

    MBLOCNOTE.Lines.Add('');
    MBLOCNOTE.Lines.Add(TraduireMemoire('Recherche demandée pour une quantité de ')+IntToStr(TobContexte.GetValue('QUANTITE'))+TraduireMemoire(' pour le ')+DateToStr(TobContexte.GetValue('DATECOMMANDE')));
  end;
end;

{-----------------------------------------------------------------------------------
   Calcul du Prix Net à partir d'un prix brut et d'une remise
------------------------------------------------------------------------------------}
function CalculPrixNet(nPrixBrut, nRemise1, nRemise2, nRemise3 : double; iSensTaux: integer; sMonnaie, sModule, sQuoi, sMethode : string) : double;
begin
  Result := ArrondirPrixMontant(nPrixBrut*(1+iSensTaux*nRemise1/100)*(1+iSensTaux*nRemise2/100)*(1+iSensTaux*nRemise3/100),sMonnaie, sModule, sQuoi, sMethode);
end;

{-----------------------------------------------------------------------------------
   Calcul de la valeur d'une ligne
------------------------------------------------------------------------------------}
function CalculValeurLigne(nForfait, nCoutLancement, nRemiseEnMontant, nPrixNet, nQte : double; sMonnaie, sModule, sQuoi, sMethode : string) : double;
begin
  Result := ArrondirPrixMontant(Max(nCoutLancement - nRemiseEnMontant + nPrixNet*nQte , nForfait),sMonnaie, sModule, sQuoi, sMethode);
end;

{-----------------------------------------------------------------------------------
   Donne les droits pour Action
------------------------------------------------------------------------------------}
function DroitAction(sDroit :string) : string;
begin
  if       (pos('C',sDroit)>0) then Result:='CREATION'
  else if  (pos('D',sDroit)>0) then Result:='CREATION'
  else if  (pos('M',sDroit)>0) then Result:='MODIFICATION'
  else if  (pos('V',sDroit)>0) then Result:='CONSULTATION'
  else                              Result:='CONSULTATION';
end;

{-----------------------------------------------------------------------------------
   Donne la nature d'auxiliaire en fonction de la fonctionnalité demandée
------------------------------------------------------------------------------------}
function NatureAuxiliaire(sFonctionnalite :string) : string;
begin
  Result := '';
  if      (sFonctionnalite=sTarifFournisseur        ) then Result := 'FOU'  // Tarif achat                Fournisseur
  else if (sFonctionnalite=sTarifSousTraitantachat  ) then Result := 'FOU'  // Tarif sous-traitance Achat Fournisseur
  else if (sFonctionnalite=sTarifSousTraitantPhase  ) then Result := 'FOU'  // Tarif sous-traitance Phase Fournisseur
  else if (sFonctionnalite=sTarifClient             ) then Result := 'CLI'  // Tarif vente      Client
  else if (sFonctionnalite=sCommissionClient        ) then Result := 'CLI'  // Commissionnement Client
  else if (sFonctionnalite=sIndirectFournisseur     ) then Result := 'FOU'  // Indirect         Fournisseur
  else PGIInfo('Contactez le S.A.T.','La fonctionnalité demandée est inconnue par le système');
end;

{-----------------------------------------------------------------------------------
   Donne le sens Taux ou Remise en fonction de la fonctionnalité demandée
------------------------------------------------------------------------------------}
function SensTauxOuRemise(sFonctionnalite :string) : integer;
begin
  if (sFonctionnalite=sCommissionClient) or (sFonctionnalite=sCommissionFournisseur) then
    Result := +1
  else
    Result := -1;
end;

{--------------------------------------------------------------------------------------
	Conversion des monnaies suivant la monnaie de présentation et la monnaie de tenue
---------------------------------------------------------------------------------------}
function  ConvMonnaie(sMonnaieOrigine, sMonnaieDestination :String; nValeur: Double; lFixing: boolean; sModule, sQuoi, sMethode: String): Double;
	// sMonnaieOrigine : Monnaie d'origine
	// sMonnaieDestination : Monnaie de destination
	// nValeur   : Prix ou Montant à convertir
	// lFixing: Utilisation du taux de fixing ou du taux du jour
	// sOrigine de la conversion	'VTE':Vente
	//										'ACH':Achat
	//										'   ':Pas d'arrondi
	//  									A renseigner seulement si cQUOI <> 'T'
	// sQuoi = 'P' on arrondi avec l'arrondi du prix
	// sQuoi = 'T' on n'arrondi pas le taux de monnaie converti
	// sQuoi = 'M' on n'arrondi pas le montant converti (a voir Windows)

var
  nTaux    : double;
  TobDevise : tob;
  sNewMethode : tArrondiMeth;
begin

  nTaux   := 1;

  //Conversion en EURO : monnaie pivot
  TobDevise := VH_GC.TobDevise.FindFirst(['D_DEVISE'], [sMonnaieOrigine], false);
  if (TobDevise<>nil) then
  begin
    if lFixing and (TobDevise.GetValue('D_PARITEEUROFIXING')<>0) then
      nTaux := TobDevise.GetValue('D_PARITEEUROFIXING')
    else
      nTaux := TobDevise.GetValue('D_PARITEEURO');
    if (nTaux=0) then  PGIInfo('Le taux conversion pour la monnaie '+sMonnaieOrigine+' n''est pas renseigné.','ATTENTION : Conversion de monnaie impossible');
  end
  else
    PGIInfo('Le code monnaie '+sMonnaieOrigine+' est inconnu.','ATTENTION : Conversion de monnaie impossible');
  nValeur  := wDivise(nValeur,nTaux);  //Valeur en EURO non arrondie

  // Conversion en monnaie de destination
  TobDevise := VH_GC.TobDevise.FindFirst(['D_DEVISE'], [sMonnaieDestination], false);
  if (TobDevise<>nil) then
  begin
    if lFixing and (TobDevise.GetValue('D_PARITEEUROFIXING')<>0) then
      nTaux := TobDevise.GetValue('D_PARITEEUROFIXING')
    else
      nTaux := TobDevise.GetValue('D_PARITEEURO');
    if (nTaux = 0) then PGIInfo('Le taux conversion pour la monnaie '+sMonnaieDestination+' n''est pas renseigné.','ATTENTION : Conversion de monnaie impossible');
  end
  else
    PGIInfo('Le code monnaie '+sMonnaieDestination+' est inconnu.','ATTENTION : Conversion de monnaie impossible');
  sNewMethode :=MethodeToArrondiMeth(sMethode) ;
  // Gestion de l'arrondi
  if 	  (sQuoi='T') then   // Pas d'arrondi
    nValeur := nValeur * nTaux
  else if (sQuoi='P') and (sModule='FOU') then   //Arrondi avec l'arrondi du prix d'achat
//    nValeur := wArrondir(nValeur*nTaux, TobDevise.GetValue('D_ARRONDIPRIXACHAT'),sMethode)
    nValeur := wArrondir(nValeur*nTaux, PoidsToArrondiPrec(TobDevise.GetValue('D_ARRONDIPRIXACHAT')),sNewMethode)
  else if (sQuoi='P') and (sModule='CLI') then   //Arrondi avec l'arrondi du prix de vente
    nValeur := wArrondir(nValeur*nTaux, PoidsToArrondiPrec(TobDevise.GetValue('D_ARRONDIPRIXVENTE')),sNewMethode)
  else if (sQuoi='M') then   //Arrondi avec l'arrondi du montant
    nValeur := wArrondir(nValeur*nTaux,PoidsToArrondiPrec( 0.01), sNewMethode)
  else
    nValeur := nValeur*nTaux;

  Result := nValeur;
end;

{-------------------------------------------------------------------------------
   Arrondir un Prix ou un Montant en fonction de la monnaie
--------------------------------------------------------------------------------}
function ArrondirPrixMontant(nValeur: double; sMonnaie, sModule, sQuoi, sMethode: string): double;
var
  TobDevise : tob;
begin
  TobDevise := VH_GC.TobDevise.FindFirst(['D_DEVISE'], [sMonnaie], false);
  if (TobDevise<>nil) then
  begin
    // Gestion de l'arrondi
    if 	  (sQuoi='T') then   // Pas d'arrondi
      nValeur := nValeur
    else if (sQuoi='P') and (sModule='FOU') then   //Arrondi avec l'arrondi du prix d'achat
      nValeur := wArrondir(nValeur, PoidsToArrondiPrec(TobDevise.GetValue('D_ARRONDIPRIXACHAT')),MethodeToArrondiMeth(sMethode))
    else if (sQuoi='P') and (sModule='CLI') then   //Arrondi avec l'arrondi du prix de vente
      nValeur := wArrondir(nValeur, PoidsToArrondiPrec(TobDevise.GetValue('D_ARRONDIPRIXVENTE')),MethodeToArrondiMeth(sMethode))
    else if (sQuoi='M') then   //Arrondi avec l'arrondi du montant
      nValeur := wArrondir(nValeur, PoidsToArrondiPrec(0.01), MethodeToArrondiMeth(sMethode));
  end
  else
    PGIInfo('Le code monnaie '+sMonnaie+' est inconnu.','ATTENTION : L''arrondi de monnaie est impossible');

  Result := nValeur;
end;

{-------------------------------------------------------------------------------
   Charge de la Tob avec une ou plusieurs devise suivant le where
--------------------------------------------------------------------------------}
function wGetTobDevise(sChamp, sWhere: String; TobDevise: Tob; SelectDB:Boolean = False): Boolean;
var
	sRequete : String;
  Q        : tQuery;
begin
  Result := False;
  if (TobDevise <> nil) then
  begin
    sRequete := 'SELECT '+sChamp+' FROM DEVISE ';
    if (sWhere<>'') then sRequete := sRequete +'WHERE '+sWhere;
    if ExisteSQL(sRequete) then
    begin
      Q := OpenSQL(sRequete, True,-1,'',true);
      try
        if (not SelectDB) then
        begin
           TobDevise.LoadDetailDB('DEVISE', '', '', Q, True, True);
           Result := (TobDevise.Detail.Count > 0);
        end
        else
        begin
           TobDevise.SelectDB('', Q);
           Result := True;
        end;
      finally
        Ferme(Q);
      end;
    end;
  end;
end;

{-------------------------------------------------------------------------------
   Création de la TobContexte pour affichage de l'origine du système tarifaire
   Appel depuis les pièces
--------------------------------------------------------------------------------}
Procedure GetTobContexteFromPiece(sFonctionnalite, sRang: string; TobContexte, TobPiece, TobLigne: Tob);
begin
  TobContexte.AddChampSupValeur('FONCTIONNALITE', sFonctionnalite                       , false);
  TobContexte.AddChampSupValeur('NATUREAUXI'    , NatureAuxiliaire(sFonctionnalite)    , false);
  TobContexte.AddChampSupValeur('FACTUREHT'     , TobLigne.GetValue('GL_FACTUREHT')     , false);
  TobContexte.AddChampSupValeur('DEPOT'         , TobLigne.GetValue('GL_DEPOT')         , false);
  TobContexte.AddChampSupValeur('TARIFDEPOT'    , ''                                    , false);
  TobContexte.AddChampSupValeur('DEVISE'        , TobPiece.GetValue('GP_DEVISE')        , false);
  TobContexte.AddChampSupValeur('TARIFSPECIAL'  , TobLigne.GetValue('GL_TARIFSPECIAL')  , false);

  if ( (sFonctionnalite=sCommissionClient) or (sFonctionnalite=sCommissionFournisseur) ) then
    TobContexte.AddChampSupValeur('TARIFSPECIAL', TobLigne.GetValue('GL_TARIFSPECIAL') , false); //en attendant que GP_COMMSPECIAL n'est pas accepté

  if      (sRang='2') then TobContexte.AddChampSupValeur('COMMERCIAL', TobLigne.GetValue('GL_REPRESENTANT') , false)
  else if (sRang='3') then TobContexte.AddChampSupValeur('COMMERCIAL', TobLigne.GetValue('GL_REPRESENTANT2'), false)
  else if (sRang='4') then TobContexte.AddChampSupValeur('COMMERCIAL', TobLigne.GetValue('GL_REPRESENTANT3'), false)
  else                     TobContexte.AddChampSupValeur('COMMERCIAL', ''                                   , false);
  TobContexte.AddChampSupValeur('TYPECOMMERCIAL', ''                                    , false);

  TobContexte.AddChampSupValeur('TIERS'         , TobPiece.GetValue('GP_TIERS')         , false);
  TobContexte.AddChampSupValeur('TARIFTIERS'    , TobPiece.GetValue('GP_TARIFTIERS')    , false);

  TobContexte.AddChampSupValeur('AFFAIRE'       , TobPiece.GetValue('GP_AFFAIRE')       , false);
  TobContexte.AddChampSupValeur('TARIFAFFAIRE'  , ''                                    , false);

  TobContexte.AddChampSupValeur('POLE'          , ''                                    , false);
  TobContexte.AddChampSupValeur('ATELIER'       , ''                                    , false);
  TobContexte.AddChampSupValeur('GROUPE'        , ''                                    , false);

  TobContexte.AddChampSupValeur('JOURSEMAINE'   , ''                                    , false);
  TobContexte.AddChampSupValeur('HEUREJOUR'     , ''                                    , false);
  TobContexte.AddChampSupValeur('LIEUCONSO'     , ''                                    , false);

  TobContexte.AddChampSupValeur('ARTICLE'       , TobLigne.GetValue('GL_ARTICLE')       , false);
  TobContexte.AddChampSupValeur('TARIFARTICLE'  , TobLigne.GetValue('GL_TARIFARTICLE')  , false);

  TobContexte.AddChampSupValeur('RESSOURCE'     , ''                                    , false);
  TobContexte.AddChampSupValeur('TARIFRESSOURCE', ''                                    , false);

  TobContexte.AddChampSupValeur('POIDSNET'      , TobLigne.GetValue('GL_POIDSNET')      , false);
  TobContexte.AddChampSupValeur('POIDSBRUT'     , TobLigne.GetValue('GL_POIDSBRUT')     , false);
  TobContexte.AddChampSupValeur('LINEAIRE'      , TobLigne.GetValue('GL_LINEAIRE')      , false);
  TobContexte.AddChampSupValeur('SURFACE'       , TobLigne.GetValue('GL_SURFACE')       , false);
  TobContexte.AddChampSupValeur('VOLUME'        , TobLigne.GetValue('GL_VOLUME')        , false);

  TobContexte.AddChampSupValeur('PRIXPOURQTE'   , TobLigne.GetValue('GL_PRIXPOURQTE')   , false);
  TobContexte.AddChampSupValeur('QUALIFUNITEVTE', TobLigne.GetValue('GL_QUALIFQTEVTE')  , false);

  TobContexte.AddChampSupValeur('MARQUE'        , ''                                    , false);
  TobContexte.AddChampSupValeur('CIRCUIT'       , TobLigne.GetValue('GLC_CIRCUIT')      , false);
  TobContexte.AddChampSupValeur('PHASE'         , ''                                    , false);
  TobContexte.AddChampSupValeur('TRAITEMENT'    , ''                                    , false);

  TobContexte.AddChampSupValeur('DATECOMMANDE'  , TobPiece.GetValue('GP_DATEREFEXTERNE'), false);
  TobContexte.AddChampSupValeur('DATELIVRAISON' , TobLigne.GetValue('GL_DATELIVRAISON') , false);
  TobContexte.AddChampSupValeur('QUANTITE'      , TobLigne.GetValue('GL_QTEFACT')       , false);
end;


{-------------------------------------------------------------------------------
   Création de la TobTarifs
--------------------------------------------------------------------------------}
procedure TobTarifsAddChampsSupValeur( TobTarifs: Tob);
begin
  { Cette Tob contient la valeur et le contexte dans lequel le système a trouvé, le forfait, le cout de lancement, le prix brut ou le prix net}
  TobTarifs.AddChampSupValeur('FORFAIT_VAL'          , 0    , False); TobTarifs.AddChampSupValeur('FORFAIT_OU'      , ''   , False);
  TobTarifs.AddChampSupValeur('COUTLANCEMENT_VAL'    , 0    , False); TobTarifs.AddChampSupValeur('COUTLANCEMENT_OU', ''   , False);
  TobTarifs.AddChampSupValeur('REMISEMONTANT_VAL'    , 0    , False); TobTarifs.AddChampSupValeur('REMISEMONTANT_OU', ''   , False);
  TobTarifs.AddChampSupValeur('PRIXBRUT_VAL'         , 0    , False); TobTarifs.AddChampSupValeur('PRIXBRUT_OU'     , ''   , False);
  TobTarifs.AddChampSupValeur('PRIXNET_VAL'          , 0    , False); TobTarifs.AddChampSupValeur('PRIXNET_OU'      , ''   , False);
  { Cette Tob contient également les possibilités de cascades sur les remises }
  TobTarifs.AddChampSupValeur('YTA_CASCDEPOT'        , False, False);
  TobTarifs.AddChampSupValeur('YTA_CASCSPECIAL'      , False, False);
  TobTarifs.AddChampSupValeur('YTA_CASCCOMMERCIAL'   , False, False); TobTarifs.AddChampSupValeur('YTA_CASCTYPECOMM'  , False, False);
  TobTarifs.AddChampSupValeur('YTA_CASCTIERS'        , False, False); TobTarifs.AddChampSupValeur('YTA_CASCTARIFTIERS', False, False);
  TobTarifs.AddChampSupValeur('YTQ_CASCARTICLE'      , False, False); TobTarifs.AddChampSupValeur('YTQ_CASCTARIFARTIC', False, False);
  TobTarifs.AddChampSupValeur('YTQ_CASCMARQUE'       , False, False);
  TobTarifs.AddChampSupValeur('YTQ_CASCCIRCUIT'      , False, False);
  TobTarifs.AddChampSupValeur('YTQ_CASCPHASE'        , False, False);
  TobTarifs.AddChampSupValeur('YTQ_CASCTRAITEMENT'   , False, False);
  TobTarifs.AddChampSupValeur('ORIGINEPRECEDENTE'    , ''   , False);
end;

{----------------------------------------------------------------------------
   Création une Tob Flle  TobTarifs : TobRemises
-----------------------------------------------------------------------------}
function CreationTobRemises( TobTarifs: Tob): Tob;
begin
  Result   := Tob.Create('_TarifsRemises_',TobTarifs,-1);
  { Cette Tob contient la valeur et le contexte dans lequel le système a trouvé, la remise 1, la remise 2 ou le remise 3}
  Result.AddChampSupValeur('REMISE1_VAL'       ,0,False);
  Result.AddChampSupValeur('REMISE2_VAL'       ,0,False);
  Result.AddChampSupValeur('REMISE3_VAL'       ,0,False);
  Result.AddChampSupValeur('REMISES_OU'        ,'',False);
  Result.AddChampSupValeur('REMISES_INTITULE'  ,'',False);
end;

{----------------------------------------------------------------------------
   Recherche code fonctionnalité en fonction de
      * la nature de pièce
      * du type de flux
-----------------------------------------------------------------------------}
function RechercheFonctionnalite(sNaturePiece: string): string;
begin
  Result := GetInfoParPiece(sNaturePiece, 'GPP_TARIFMODULE');
  if (Result='') then
  begin
    PGIInfo(TraduireMemoire('Aucun tarif défini pour la nature de pièce')+sNaturePiece, TraduireMemoire('Recherche tarifaire abandonnée'));
    Result := '';
  end;
end;

{----------------------------------------------------------------------------
   Recherche des tarifs
   Appel depuis les pièces
-----------------------------------------------------------------------------}
function RechercheEnregistreTarifs(sFonctionnalite, sAppel, sRang: string; TobTiers, TobArticle, TobPiece, TobLigne, TobLigneTarif : tob; var nMontantFixe : double ; var nPrixLigne : double; var nRemiseTotale : double; var nRemiseEnMontant : double; var sMonnaiePiece : string): boolean;
var
  TobContexte, TobTarifs : Tob;
begin
  TobContexte  := Tob.Create('_TobContexte_', nil, -1);
  TobTarifs    := Tob.Create('_TobTarifs_'  , nil, -1);
  try
    if (sFonctionnalite<>'') then
    begin
      GetTobContexteFromPiece(sFonctionnalite, sRang, TobContexte, TobPiece, TobLigne);
      RechercheTarifs(sAppel, TobContexte, TobTarifs, Tobtiers, TobArticle, TobPiece, TobLigne, False);
      EnregistreTarifs(sFonctionnalite, sAppel, sRang, TobContexte, TobTarifs, TobLigne, TobLigneTarif, nMontantFixe, nPrixLigne, nRemiseTotale, nRemiseEnMontant, sMonnaiePiece);
    end;  
    result:=(TobTarifs <> nil);
  finally
    TobContexte.Free;
    TobTarifs.Free;
  end;
end;

{----------------------------------------------------------------------------
   Enregistrement des tarifs et de ses origines
-----------------------------------------------------------------------------}
procedure EnregistreTarifs(sFonctionnalite, sAppel, sRang: string; TobContexte, TobTarifs, TobLigne, TobLigneTarif: tob; var nMontantFixe : double; var nPrixLigne : double; var nRemiseTotale : double; var nRemiseEnMontant : double; var sMonnaiePiece : string);
var
  sLibRemise, sLigneOuCommande   : string;
  iCpt, iSensTaux : integer;
  TheTob     : tob;
begin

  if (sAppel='PIECE') then
    sLigneOuCommande := sOu_P
  else
    sLigneOuCommande := sOu_L;

  if (TobTarifs<>nil) then
  begin
    sMonnaiePiece := TobContexte.GetValue('DEVISE');
    iSensTaux := SensTauxOuRemise(sFonctionnalite);

    //Suppression de la fille ligne et des filles détail ligne tarif
    TheTob := TobLigneTarif.FindFirst(['GL_NATUREPIECEG', 'GL_SOUCHE', 'GL_NUMERO', 'GL_INDICEG', 'GL_NUMLIGNE'],[TobLigne.GetValue('GL_NATUREPIECEG'), TobLigne.GetValue('GL_SOUCHE'), TobLigne.GetValue('GL_NUMERO'), TobLigne.GetValue('GL_INDICEG'), TobLigne.GetValue('GL_NUMLIGNE')], False);
    if (TheTob=nil) then
    begin
      //Création de la fille ligne pour laquelle nous allons créer les filles détail ligne tarif
      CreerLigneTarifMere(TobLigneTarif, -1, TobLigne, TobLigne.GetValue('GL_NUMLIGNE'));
    end;

    //Si le système tarifaire a trouvé des prix net ou brut alors on calcule avec ces prix sinon on prend le prix saisi
    if       (TobTarifs.GetValue('PRIXNET_OU') <>'') then
      nPrixLigne    := TobTarifs.GetValue('PRIXNET_VAL')
    else if  (TobTarifs.GetValue('PRIXBRUT_OU')<>'') then
      nPrixLigne := CalculPrixNet(TobTarifs.GetValue('PRIXBRUT_VAL'), nRemiseTotale, 0, 0, iSensTaux, TobContexte.GetValue('DEVISE'), TobContexte.GetValue('NATUREAUXI'), 'P', 'T')
    else if (sFonctionnalite=sTarifFournisseur) or (sFonctionnalite=sTarifSousTraitantAchat) or (sFonctionnalite=sTarifSousTraitantPhase) or (sFonctionnalite=sTarifClient) then
      if (TobContexte.GetValue('FACTUREHT')=wTrue) then nPrixLigne := TobLigne.GetValue('GL_PUHTDEV') else nPrixLigne := TobLigne.GetValue('GL_PUTTCDEV')
    else
      nPrixLigne := 0;

    //Si le système tarifaire a trouvé des remises alors on calcule avec ces remises systèmes sinon on prend la remise saisie
    if (TobTarifs.Detail.Count<>0) then
    begin
      nRemiseTotale := 1;
      for iCpt:=0 to (TobTarifs.Detail.Count-1) do
         nRemiseTotale := nRemiseTotale*(1-TobTarifs.Detail[iCpt].GetValue('REMISE1_VAL')/100)*(1-TobTarifs.Detail[iCpt].GetValue('REMISE2_VAL')/100)*(1-TobTarifs.Detail[iCpt].GetValue('REMISE3_VAL')/100);
      nRemiseTotale :=(1-nRemiseTotale)*100;   //Transformation du Coefficient de remise en pourcentage de remise
    end
    else if (sFonctionnalite=sTarifFournisseur) or (sFonctionnalite=sTarifSousTraitantAchat) or (sFonctionnalite=sTarifSousTraitantPhase) or (sFonctionnalite=sTarifClient) then
      nRemiseTotale := TobLigne.GetValue('GL_REMISELIGNE')
    else
      nRemiseTotale := 0;

    //Mise à jour ou création de la Tob des tarifs sur une ligne de pièce
    //iSensTaux := +1;     //Commissionnenement ou Tarifs

    TobLigne.PutValue('GL_VALEURFIXEDEV', 0);
    TobLigne.PutValue('GL_VALEURREMDEV' , 0);
    TobLigne.PutValue('GL_REMISELIGNE'  , 0); 

    //Annulation de tout le groupe
    AnnulationTobLigneTarif( TobLigne, TobLigneTarif, sFonctionnalite, sRang);
    //Recréation des trois remises libres de la ligne pièce
    MajTobLigneTarifFromTobLigne(TobLigne, TobLigneTarif, sFonctionnalite, '1501', 'REMISELIBRE1' , TobLigne.GetValue('GL_REMISELIBRE1'));
    MajTobLigneTarifFromTobLigne(TobLigne, TobLigneTarif, sFonctionnalite, '1502', 'REMISELIBRE2' , TobLigne.GetValue('GL_REMISELIBRE2'));
    MajTobLigneTarifFromTobLigne(TobLigne, TobLigneTarif, sFonctionnalite, '1503', 'REMISELIBRE3' , TobLigne.GetValue('GL_REMISELIBRE3'));   

    //Si forfait <> 0 et que la valeur de ligne est égal au forfait alors on crée la ligne forfait
    if (TobTarifs.GetValue('FORFAIT_VAL')<>0) and (TobTarifs.GetValue('FORFAIT_VAL')=CalculValeurLigne(TobTarifs.GetValue('FORFAIT_VAL'), TobTarifs.GetValue('COUTLANCEMENT_VAL'), TobTarifs.GetValue('REMISEMONTANT_VAL'), nPrixLigne, TobContexte.GetValue('QUANTITE'), TobContexte.GetValue('DEVISE'), TobContexte.GetValue('NATUREAUXI'), 'M', 'P')) then
    begin
      nMontantFixe  := TobTarifs.GetValue('FORFAIT_VAL');
      nPrixLigne    := 0;
      nRemiseTotale := 0;
      CreationTobLigneTarif(TobLigne, TobLigneTarif, sFonctionnalite, sRang+'1', LibelleParFonctionnalite(sFonctionnalite, 'FORFAIT'), TobTarifs.GetValue('FORFAIT_VAL'), TobTarifs.GetValue('FORFAIT_OU'));
    end
    else
    begin
      nMontantFixe := TobTarifs.GetValue('COUTLANCEMENT_VAL');
      if (TobTarifs.GetValue('COUTLANCEMENT_VAL')<>0) or (copy(TobTarifs.GetValue('COUTLANCEMENT_OU'),1, iTailleOu)=sOrigineMasque) then
        CreationTobLigneTarif(TobLigne, TobLigneTarif, sFonctionnalite, sRang+'2', LibelleParFonctionnalite(sFonctionnalite, 'COUTLANCEMENT'), TobTarifs.GetValue('COUTLANCEMENT_VAL'), TobTarifs.GetValue('COUTLANCEMENT_OU'));

      nRemiseEnMontant :=  TobTarifs.GetValue('REMISEMONTANT_VAL');
      if (TobTarifs.GetValue('REMISEMONTANT_VAL')<>0) or (copy(TobTarifs.GetValue('REMISEMONTANT_OU'),1, iTailleOu)=sOrigineMasque) then
        CreationTobLigneTarif(TobLigne, TobLigneTarif, sFonctionnalite, sRang+'3', LibelleParFonctionnalite(sFonctionnalite, 'REMISEMONTANT'), TobTarifs.GetValue('REMISEMONTANT_VAL'), TobTarifs.GetValue('REMISEMONTANT_OU'));

      nRemiseTotale := (1 - (CalculRemiseTotaleSurTobLigneTarif( TobLigne, TobLigneTarif, sRang+'5')/100));

      if (TobTarifs.GetValue('PRIXNET_VAL')<>0) or (copy(TobTarifs.GetValue('PRIXNET_OU'),1, iTailleOu)=sOrigineMasque) then
      begin
        CreationTobLigneTarif(TobLigne, TobLigneTarif, sFonctionnalite, sRang+'7', LibelleParFonctionnalite(sFonctionnalite, 'PRIXNET'), TobTarifs.GetValue('PRIXNET_VAL') , TobTarifs.GetValue('PRIXNET_OU'));
        nRemiseTotale := 0;
      end
      else
      begin
        if      (TobTarifs.GetValue('PRIXBRUT_VAL')<>0) or (copy(TobTarifs.GetValue('PRIXBRUT_OU'),1, iTailleOu)=sOrigineMasque) then
          CreationTobLigneTarif(TobLigne, TobLigneTarif, sFonctionnalite, sRang+'4', LibelleParFonctionnalite(sFonctionnalite, 'PRIXBRUT')   , TobTarifs.GetValue('PRIXBRUT_VAL'), TobTarifs.GetValue('PRIXBRUT_OU'))
        else if (nPrixLigne                        <>0) then
          CreationTobLigneTarif(TobLigne, TobLigneTarif, sFonctionnalite, sRang+'4', LibelleParFonctionnalite(sFonctionnalite, 'PRIXSAISIE') , nPrixLigne                        , 'M'                              );

        //Calcul de la remise globale en coefficient
        for iCpt:=0 to (TobTarifs.Detail.Count-1) do
        begin
          sLibRemise := TobTarifs.Detail[iCpt].GetValue('REMISES_INTITULE');
          if (TobTarifs.Detail[iCpt].GetValue('REMISE1_VAL')<>0) or (copy(TobTarifs.Detail[iCpt].GetValue('REMISES_OU'),1, iTailleOu)=sOrigineMasque) then
          begin
            if (sLibRemise='') then
              sLibRemise := LibelleParFonctionnalite(sFonctionnalite, 'REMISE')+' '+ TraduireMemoire('1 Niveau')+IntToStr(iCpt);
            CreationTobLigneTarif(TobLigne, TobLigneTarif, sFonctionnalite, sRang+'6'+IntToStr((iCpt+1)*10+1), sLibRemise, TobTarifs.Detail[iCpt].GetValue('REMISE1_VAL'),TobTarifs.Detail[iCpt].GetValue('REMISES_OU'));
          end;
          if (TobTarifs.Detail[iCpt].GetValue('REMISE2_VAL')<>0) or (copy(TobTarifs.Detail[iCpt].GetValue('REMISES_OU'),1, iTailleOu)=sOrigineMasque) then
          begin
            if (sLibRemise='') then
              sLibRemise := LibelleParFonctionnalite(sFonctionnalite, 'REMISE')+' '+ TraduireMemoire('2 Niveau')+IntToStr(iCpt);
            CreationTobLigneTarif(TobLigne, TobLigneTarif, sFonctionnalite, sRang+'6'+IntToStr((iCpt+1)*10+2), sLibRemise, TobTarifs.Detail[iCpt].GetValue('REMISE2_VAL'),TobTarifs.Detail[iCpt].GetValue('REMISES_OU'));
          end;
          if (TobTarifs.Detail[iCpt].GetValue('REMISE3_VAL')<>0) or (copy(TobTarifs.Detail[iCpt].GetValue('REMISES_OU'),1, iTailleOu)=sOrigineMasque) then
          begin
            if (sLibRemise='') then
              sLibRemise := LibelleParFonctionnalite(sFonctionnalite, 'REMISE')+' '+ TraduireMemoire('3 Niveau')+IntToStr(iCpt);
            CreationTobLigneTarif(TobLigne, TobLigneTarif, sFonctionnalite, sRang+'6'+IntToStr((iCpt+1)*10+3), sLibRemise, TobTarifs.Detail[iCpt].GetValue('REMISE3_VAL'),TobTarifs.Detail[iCpt].GetValue('REMISES_OU'));
          end;
          nRemiseTotale := nRemiseTotale*(1-TobTarifs.Detail[iCpt].GetValue('REMISE1_VAL')/100)*(1-TobTarifs.Detail[iCpt].GetValue('REMISE2_VAL')/100)*(1-TobTarifs.Detail[iCpt].GetValue('REMISE3_VAL')/100);
        end;
        nRemiseTotale:=(1-nRemiseTotale)*100;   //Transformation du Coefficient de remise en pourcentage de remise
      end;
    end;
  end
  else
  begin
    nRemiseTotale := (1 - (CalculRemiseTotaleSurTobLigneTarif( TobLigne, TobLigneTarif, '15')/100));
  end;
end;

{-----------------------------------------------------------------------------------
   alimentation en automatique du libellé du tarif
------------------------------------------------------------------------------------}
function  SetTarifLibelle(sFonctionnalite, sTiers, sFamilleTiers, sAffaire, sFamilleAffaire, sArticle, sFamilleArticle, sRessource, sFamilleRessource :string) : string;
var
  sParametre : string;
begin
  Result := LibelleParFonctionnalite(sFonctionnalite, 'REMISE');

  sParametre := GetParametresFromYFO(sFonctionnalite);

  if      (sTiers            <>'') then Result := Result +' '+ TraduireMemoire(' : tiers')
  else if (sFamilleTiers     <>'') then Result := Result +' '+ TraduireMemoire(' : famille de tiers')
  else                                  Result := Result +' '+ TraduireMemoire(' : tous les tiers');

//  if      (sAffaire          <>'') then Result := Result +' '+ TraduireMemoire('affaire')
//  else if (sFamilleAffaire   <>'') then Result := Result +' '+ TraduireMemoire('famille d''affaire')
//  else                                  Result := Result +' '+ TraduireMemoire('toutes les affaires');

  if      (sArticle          <>'') then Result := Result +' '+ TraduireMemoire(' - article')
  else if (sFamilleArticle   <>'') then Result := Result +' '+ TraduireMemoire(' - famille d''article')
  else                                  Result := Result +' '+ TraduireMemoire(' - tous les articles');

//  if      (sRessource        <>'') then Result := Result +' '+ TraduireMemoire('ressource')
//  else if (sFamilleRessource <>'') then Result := Result +' '+ TraduireMemoire('famille de ressource')
//  else                                  Result := Result +' '+ TraduireMemoire('toutes les ressources');
end;

{-----------------------------------------------------------------------------------
   Controle de la clef naturelle sur la table yTarifsAQui
------------------------------------------------------------------------------------}
function ExistYTA(sWhere : string):Boolean;
var
	sSql  : string;
begin
	sSql := ' SELECT null'
			+ ' FROM YTARIFSAQUI'
			+ ' WHERE ' + sWhere
         ;
	Result := ExisteSQL(sSql);
end;

{-----------------------------------------------------------------------------------
   Controle de la clef naturelle sur la table yTarifsSurQuoi
------------------------------------------------------------------------------------}
function ExistYTQ(sWhere : string):Boolean;
var
	sSql  : string;
begin
	sSql := ' SELECT null'
			+ ' FROM YTARIFSSURQUOI'
			+ ' WHERE ' + sWhere
         ;
	Result := ExisteSQL(sSql);
end;

{-----------------------------------------------------------------------------------
   Vérification des conditions particulières d'application
------------------------------------------------------------------------------------}
function VerifConditionsApplications(sAppel, sConditionsApplications : string; tobtiers, TobArticle, TobPiece, TobLigne : Tob) : Boolean ;
var
  sConditions : TStrings ;
  iLigne      : Integer ;
  sLigneConditions : String ;
  sTable, sChampNom, sOperateur, sValeur : String ;
  vChampVal   : Variant ;
  vValeurVal  : Variant ;
  sChampVal   : String ;
begin
  Result:=True ;
  sConditions      := TStringList.Create ;
  sConditions.Text := sConditionsApplications;

  if (sConditions<>nil) then
  begin
    for iLigne:=0 to sConditions.Count-1 do
    begin
      sLigneConditions  := sConditions.Strings[iLigne];
      sTable            := UpperCase(ReadTokenSt(sLigneConditions));
      sChampNom         := UpperCase(ReadTokenSt(sLigneConditions));
      sOperateur        := UpperCase(ReadTokenSt(sLigneConditions));
      sValeur           := UpperCase(ReadTokenSt(sLigneConditions));

      // Si appel depuis les pièces ou Si la conditions ne concerne pas la table PIECE ou la table LIGNE
      // alors on teste de la condition
      // sinon on considère la condition est respectée à défaut car on ne peut évaluer la condition
      if (Pos(sAppel,'PIECE/LIGNE')>0) or (Pos(sTable,'PIECE/LIGNE')=0) then
      begin
        { Valeur du Champ à tester }
        if (sTable='TIERS'  ) and (assigned(TobTiers))    then vChampVal := TobTiers.GetValue(sChampNom);
        if (sTable='ARTICLE') and (assigned(TobArticle))  then vChampVal := TobArticle.GetValue(sChampNom);
        if (sTable='PIECE'  ) and (assigned(TobPiece))    then vChampVal := TobPiece.GetValue(sChampNom);
        if (sTable='LIGNE'  ) and (assigned(TobLigne))    then vChampVal := TobLigne.GetValue(sChampNom);

        sChampVal := Uppercase(VarAsType(vChampVal,VarString));
        if (sChampVal<>'') then
        begin
          {Valeur de test}
          try
            vValeurVal  := VarAsType(sValeur,VarType(vChampVal));

            {Test de la condition}
            if      (sOperateur='=' ) then Result := (vChampVal =  vValeurVal)
            else if (sOperateur='<>') then Result := (vChampVal <> vValeurVal)
            else if (sOperateur='>' ) then Result := (vChampVal >  vValeurVal)
            else if (sOperateur='<' ) then Result := (vChampVal <  vValeurVal)
            else if (sOperateur='>=') then Result := (vChampVal >= vValeurVal)
            else if (sOperateur='<=') then Result := (vChampVal <= vValeurVal)
            else if (sOperateur='C' ) then Result := (Copy(sChampVal,1,Length(sValeur))=sValeur)
            else if (sOperateur='D' ) then Result := (Copy(sChampVal,1,Length(sValeur))<>sValeur)
            else if (sOperateur='L' ) then Result := (Pos(sValeur,sChampVal)>0)
            else if (sOperateur='M' ) then Result := (Pos(sValeur,sChampVal)=0);
          except
          end;
        end;
      end;
    end;
  end;
  sConditions.Free ;
end ;

{-------------------------------------------------------------------------------
  Renvoir le libellé en fonction du champ et de la fonctionnalité
--------------------------------------------------------------------------------}
function LibelleParFonctionnalite(sFonctionnalite, sQuoi : string) : string;
begin
  Result := TraduireMemoire('Libellé par fonctionnalité non traduite');
  if (sFonctionnalite=sCommissionClient) then
  begin
    if      (sQuoi='FORFAIT')       then Result := TraduireMemoire('Forfait')
    else if (sQuoi='COUTLANCEMENT') then Result := TraduireMemoire('Base + ')
    else if (sQuoi='REMISEMONTANT') then Result := TraduireMemoire('Base - ')
    else if (sQuoi='PRIXSAISIE')    then Result := TraduireMemoire('Unitaire saisie')
    else if (sQuoi='PRIXBRUT')      then Result := TraduireMemoire('Unitaire brute')
    else if (sQuoi='REMISE')        then Result := TraduireMemoire('Taux')
    else if (sQuoi='REMISE123')     then Result := TraduireMemoire('Taux 1/2/3')
    else if (sQuoi='REMISETOTALE')  then Result := TraduireMemoire('Taux total')
    else if (sQuoi='PRIXNET')       then Result := TraduireMemoire('Unitaire nette')
    else if (sQuoi='MONTANTTOTAL')  then Result := TraduireMemoire('Montant total')
    ;
  end
  else
  begin
    if      (sQuoi='FORFAIT')       then Result := TraduireMemoire('Forfait')
    else if (sQuoi='COUTLANCEMENT') then Result := TraduireMemoire('Coût de lancement')
    else if (sQuoi='REMISEMONTANT') then Result := TraduireMemoire('Remise en montant')
    else if (sQuoi='PRIXSAISIE')    then Result := TraduireMemoire('Prix saisie')
    else if (sQuoi='PRIXBRUT')      then Result := TraduireMemoire('Prix brut')
    else if (sQuoi='REMISE')        then Result := TraduireMemoire('Remise')
    else if (sQuoi='REMISE123')     then Result := TraduireMemoire('Remises 1/2/3')
    else if (sQuoi='REMISETOTALE')  then Result := TraduireMemoire('Remise totale')
    else if (sQuoi='REMISELIBRE1')  then Result := TraduireMemoire('Remise libre N°1')
    else if (sQuoi='REMISELIBRE2')  then Result := TraduireMemoire('Remise libre N°2')
    else if (sQuoi='REMISELIBRE3')  then Result := TraduireMemoire('Remise libre N°3')
    else if (sQuoi='PRIXNET')       then Result := TraduireMemoire('Prix net')
    else if (sQuoi='MONTANTTOTAL')  then Result := TraduireMemoire('Montant total')
    ;
  end;
end;

{-------------------------------------------------------------------------------
  Mise à jour de la tobLigneTarif
  en fonction des mofidifications manuelles sur TobLigne
--------------------------------------------------------------------------------}
Procedure MajTobLigneTarifFromTobLigne( TobLigne, TobLigneTarif:tob; sFonctionnalite, sRang, sQuoi:string; nValeur:double);
var
  TobUneLigneTarif  : tob;
  iCpt              : integer;
  lTrouve           : boolean;
begin
  { Mise à Jour de la TobLigne Tarif }
  TobUneLigneTarif := TobLigneTarif.FindFirst(['GL_NATUREPIECEG', 'GL_SOUCHE', 'GL_NUMERO', 'GL_INDICEG', 'GL_NUMLIGNE'],[TobLigne.GetValue('GL_NATUREPIECEG'), TobLigne.GetValue('GL_SOUCHE'), TobLigne.GetValue('GL_NUMERO'), TobLigne.GetValue('GL_INDICEG'), TobLigne.GetValue('GL_NUMLIGNE')], False);
  if (assigned(TobUneLigneTarif)) then
  begin
    lTrouve := False;
    for iCpt := 0 to TobUneLigneTarif.Detail.Count - 1 do
    begin
      if (     (TobUneLigneTarif.Detail[iCpt].GetValue('GLT_FONCTIONNALITE')=sFonctionnalite)
           and (sRang =TobUneLigneTarif.Detail[iCpt].GetValue('GLT_RANG'))
         ) then
      begin
        if (nValeur<>TobUneLigneTarif.Detail[iCpt].GetValue('GLT_VALEUR')) then
        begin
          TobUneLigneTarif.Detail[iCpt].PutValue('GLT_VALEUR', nValeur);
          TobUneLigneTarif.Detail[iCpt].PutValue('GLT_OU'    , 'M');
        end;
        lTrouve := True;
      end;
    end;
    if (not lTrouve) and (nValeur<>0) then
    begin
      CreationTobLigneTarif(TobLigne, TobLigneTarif, sFonctionnalite, sRang, LibelleParFonctionnalite(sFonctionnalite, sQuoi), nValeur, 'M');
    end;
  end;
end;

{----------------------------------------------------------------------------
  Création une Tob Fille  TobLigneTarif
-----------------------------------------------------------------------------}
function CreationTobLigneTarif( TobLigne, TobLigneTarif:tob; sFonctionnalite, sRang, sQuoi:string; nValeur:double; sOu:string): Tob;
var
  TobUneLigneTarif  : tob;
begin

  TobUneLigneTarif := GetLigneTarif(TobLigneTarif, TobLigne);

  AnnulationTobLigneTarif(TobLigne, TobLigneTarif, sFonctionnalite, sRang);

  Result := Tob.Create('LigneTarif',TobUneLigneTarif,-1);
  Result.PutValue('GLT_NATUREPIECEG'  ,TobLigne.GetValue('GL_NATUREPIECEG'));
  Result.PutValue('GLT_SOUCHE'        ,TobLigne.GetValue('GL_SOUCHE'      ));
  Result.PutValue('GLT_NUMERO'        ,TobLigne.GetValue('GL_NUMERO'      ));
  Result.PutValue('GLT_INDICEG'       ,TobLigne.GetValue('GL_INDICEG'     ));
  Result.PutValue('GLT_NUMLIGNE'      ,TobLigne.GetValue('GL_NUMLIGNE'    ));
  Result.PutValue('GLT_FONCTIONNALITE',sFonctionnalite );
  Result.PutValue('GLT_RANG'          ,wPadRight(sRang,4,'0'));
  Result.PutValue('GLT_QUOI'          ,sQuoi  );
  Result.PutValue('GLT_VALEUR'        ,nValeur);
  Result.PutValue('GLT_OU'            ,sOu    );
end;

{----------------------------------------------------------------------------
  Annulation d'une Tob Fille  TobLigneTarif
-----------------------------------------------------------------------------}
function AnnulationTobLigneTarif(TobLigne, TobLigneTarif:tob; sFonctionnalite, sRang : string) :boolean;
var
  TobUneLigneTarif  : tob;
  iCpt : integer;
begin
  Result := False;
  TobUneLigneTarif := GetLigneTarif(TobLigneTarif, TobLigne);

  if (assigned(TobUneLigneTarif)) then
  begin

    //Suppression de la fille ligne et des filles détail ligne tarif
    for iCpt:=TobUneLigneTarif.Detail.count - 1 downto 0 do
    begin
      if  (    (TobUneLigneTarif.Detail[iCpt].GetValue('GLT_FONCTIONNALITE')=sFonctionnalite)
           and (copy(TobUneLigneTarif.Detail[iCpt].GetValue('GLT_RANG'),1,Length(sRang)) = sRang)
          ) then
      begin
        TobUneLigneTarif.Detail[iCpt].Free;
        Result := True;
      end;
    end;
  end;
end;

{-------------------------------------------------------------------------------
  Calcul de la remise totale système sur TobLigneTarif en %
--------------------------------------------------------------------------------}
Function CalculRemiseTotaleSurTobLigneTarif( TobLigne, TobLigneTarif: tob; sRang : string) : double;
var
  sFonctionnalite   : string;
  TobUneLigneTarif  : tob;
  iCpt : integer;
begin
  sFonctionnalite := RechercheFonctionnalite(TobLigne.GetValue('GL_NATUREPIECEG'));
  TobUneLigneTarif := GetLigneTarif(TobLigneTarif, TobLigne);
  Result := 1;
  if (assigned(TobUneLigneTarif)) then
  begin
    for iCpt := 0 to TobUneLigneTarif.Detail.Count - 1 do
    begin
      if (     (TobUneLigneTarif.Detail[iCpt].GetValue('GLT_FONCTIONNALITE')=sFonctionnalite)
           and (copy(TobUneLigneTarif.Detail[iCpt].GetValue('GLT_RANG'),1,Length(sRang))=sRang)
         ) then
      begin
        Result := Result * (1 - TobUneLigneTarif.Detail[iCpt].GetValue('GLT_VALEUR')/100 );
      end;
    end;
  end;
  Result := (1-Result)*100;
end;

{ --------------------------------------------------------------------------------------------------
  Appel du système de recherche des tarifs et des commissionnements depuis les pièces
  en sortie de cellule de la grille de saisie
  en sortie de complément ligne
----------------------------------------------------------------------------------------------------  }
procedure RechercheTarifsCommissions(sAppel: string; Action : TActionFiche; TobTiers, TobArticle, TobPiece, TobLigne, TobLigneTarif, TobTarif : tob; Devise : rDevise);
var
  nFixeComm, nTauxComm, nPrixComm, nTauxEnMontant: double;
  sNatureDePiece, sTypeDeFlux : string;
  sDevise : String ;
begin

  { Recherche des tarifs }
  if (     (TobLigne.GetValue('RECALCULTARIF')='X')
        and (
              (TobLigne.GetValue('GL_BLOQUETARIF')=wFalse )  //Système non bloqué
              or
              (Action=taCreat)                               //système bloqué mais en création on fait quand même une première recherche
            )
{       and (
                 ( (TobLigne.GetValue('GL_BLOQUETARIF')='NON') or (TobLigne.GetValue('GL_BLOQUETARIF')='') )
              or ( (TobLigne.GetValue('GL_BLOQUETARIF')='VEM') and (Action=taModif)                        )
           )
}     ) then
  begin
    TarifVersLigne(TobArticle, TobTiers, TobLigne, TobligneTarif, TobPiece, TobTarif, True, True, Devise);
    TobPiece.PutValue('GP_RECALCULER','X');
    TobLigne.PutValue('GL_RECALCULER','X');
    TobLigne.PutValue('RECALCULTARIF','-');
  end;

  { Recherche des commissionnements }
  if (     ( TobLigne.GetValue('RECALCULCOMM')='X' )
        //En attendant que GL_BLOQUECOMM soit accepté
        and (
              (TobLigne.GetValue('GL_BLOQUETARIF')=wFalse )  //Système non bloqué
              or
              (Action=taCreat)                               //système bloqué mais en création on fait quand même une première recherche
            )
{       and (
                 ( (TobLigne.GetValue('GL_BLOQUECOMM')='NON') or (TobLigne.GetValue('GL_BLOQUECOMM') ='') )
              or ( (TobLigne.GetValue('GL_BLOQUECOMM')='VEM') and (Action=taModif)                        )
           )
}     ) then
  begin
    { Paramétrage Par Pièce }
    sNatureDePiece := TobLigne.GetValue('GL_NATUREPIECEG');
    sTypeDeFlux    := GetInfoParPiece(sNatureDePiece, 'GPP_VENTEACHAT');
    sDevise        := V_PGI.DevisePivot;
    { Si Flux Vente et Représentant renseigné alors Recherche du commissionnement }
    if (sTypeDeFlux='VEN') and (TobLigne.GetValue('GL_REPRESENTANT')<>'') then
    begin                      //MM
       RechercheEnregistreTarifs('202', sAppel, '2', TobTiers, TobArticle, TobPiece, TobLigne, TobLigneTarif, nFixeComm, nPrixComm, nTauxComm, nTauxEnMontant, sDevise );
    end;
{
    // Si Flux Vente et Représentant renseigné alors Recherche du commissionnement
    if (sTypeDeFlux='VEN') and (TobLigne.GetValue('GL_REPRESENTANT2')<>'') then
    begin                      //MM
       RechercheEnregistreTarifs('202', sAppel, '3', TobTiers, TobArticle, TobPiece, TobLigne, TobLigneTarif, nFixeComm, nPrixComm, nTauxComm, nTauxEnMontant, sDevise );
    end;
    // Si Flux Vente et Représentant renseigné alors Recherche du commissionnement
    if (sTypeDeFlux='VEN') and (TobLigne.GetValue('GL_REPRESENTANT3')<>'') then
    begin                      //MM
       RechercheEnregistreTarifs('202', sAppel, '4', TobTiers, TobArticle, TobPiece, TobLigne, TobLigneTarif, nFixeComm, nPrixComm, nTauxComm, nTauxEnMontant, sDevise );
    end;
}
    TobLigne.PutValue('RECALCULCOMM','-');
  end;
end;

{ ------------------------------------------------------------------------------------------------------------
  Renvoi la valeur (Quantité , Valeur ,Poids, Volume ou Surface) pour rechercher dans la table des fourchettes
--------------------------------------------------------------------------------------------------------------}
Function Quoi(sParam: string; TobContexte: tob): Double;
begin
  if      (sParam='Q') then Result := TobContexte.GetValue('QUANTITE')
  else if (sParam='M') then Result := TobContexte.GetValue('VALEUR')
  else if (sParam='N') then Result := TobContexte.GetValue('QUANTITE')*TobContexte.GetValue('POIDSNET')
  else if (sParam='B') then Result := TobContexte.GetValue('QUANTITE')*TobContexte.GetValue('POIDSBRUT')
  else if (sParam='V') then Result := TobContexte.GetValue('QUANTITE')*TobContexte.GetValue('VOLUME')
  else if (sParam='S') then Result := TobContexte.GetValue('QUANTITE')*TobContexte.GetValue('SURFACE')
  else if (sParam='L') then Result := TobContexte.GetValue('QUANTITE')*TobContexte.GetValue('LINEAIRE')
  else
  begin
    PGIInfo(' On ne sait pas quoi comparer avec les bornes inférieures et supérieures ' , '');
    Result := 0 ;
  end;
end;

{-------------------------------------------------------------------------
  Construction du WHERE selon une combinaison des informations du contexte
--------------------------------------------------------------------------}
function ContexteAqui(sRecherche, sParametre: string; sDepot, sTarifDepot, sDevise, sTarifSpecial, sCommercial, sTypeCommercial, sTiers, sTarifTiers, sAffaire, sTarifAffaire, sPole, sAtelier, sGroupe, sJourSemaine, sHeureJour, sLieuConso : string) : string;
//  1er  caractère : dépôt           / Famille tarif dépôt
//  2ème caractère : Devise          / Monnaie de tenue
//  3ème caractère : Tarif spécial
//  4ème caractère : Commercial      / Type commercial
//  5ème caractère : Tiers           / Famille tarif tiers
//  6ème caractère : Affaire         / Famille tarif affaire
//  7ème caractère : Pole
//  8ème caractère : Atelier
//  9ème caractère : Groupe
// 10ème caractère : Jour de la Semaine
// 11ème caractère : Heure du Jour
// 12ème caractère : Lieu de consommation
  {----------------------------------------------------------------------------
  -----------------------------------------------------------------------------}
  function LimiteContexteAQui : boolean;
  begin
    Result :=     (sParametre[iPosOrigineDepot]      =wTrue)
               or (sParametre[iPosOrigineDevise]     =wTrue)
               or (sParametre[iPosOrigineSpecial]    =wTrue)
               or (sParametre[iPosOrigineCommercial] =wTrue)
               or (sParametre[iPosOrigineTiers]      =wTrue)
               or (sParametre[iPosOrigineAffaire]    =wTrue)
               or (sParametre[iPosOriginePole]       =wTrue)
               or (sParametre[iPosOrigineAtelier]    =wTrue)
               or (sParametre[iPosOrigineGroupe]     =wTrue)
               or (sParametre[iPosOrigineJourSemaine]=wTrue)
               or (sParametre[iPosOrigineHeureJour]  =wTrue)
               or (sParametre[iPosOrigineLieuConso]  =wTrue)
              ;
  end;
begin
  Result := '';

  if LimiteContexteAQui then
  begin
    Result := '( ';

    //dépôt
    if (sParametre[iPosOrigineDepot]=wTrue) then
    begin
      case sRecherche[iPosOrigineDepot] of
        '1' : Result := Result + '(YTA_DEPOT="'  +sDepot+'")'+' and '+'(YTA_TARIFDEPOT="")'                 +' and ';
        '2' : Result := Result + '(YTA_DEPOT="")'            +' and '+'(YTA_TARIFDEPOT="'  +sTarifDepot+'")'+' and ';
        '3' : Result := Result + '(YTA_DEPOT="")'            +' and '+'(YTA_TARIFDEPOT="")'                 +' and ';
      end;
    end;

    //Devise
    if (sParametre[iPosOrigineDevise]=wTrue) then
    begin
      case sRecherche[iPosOrigineDevise] of
        '1' : Result := Result + '(YTA_DEVISE="'+sDevise          +'")'+' and ';
        '2' : Result := Result + '(YTA_DEVISE="'+V_PGI.DevisePivot+'")'+' and ';
      end;
    end;

    //Condition spéciale
    if (sParametre[iPosOrigineSpecial]=wTrue) then
    begin
      case sRecherche[iPosOrigineSpecial] of
        '1' : Result := Result + '(YTA_TARIFSPECIAL="'+sTarifSpecial+'")'+' and ';
        '2' : Result := Result + '(YTA_TARIFSPECIAL="")'                 +' and ';
      end;
    end;

    //Commercial
    if (sParametre[iPosOrigineCommercial]=wTrue) then
    begin
      case sRecherche[iPosOrigineCommercial] of
        '1' : Result := Result + '(YTA_COMMERCIAL="'  +sCommercial+'")'+' and '+'(YTA_TYPECOMMERCIAL="")'                     +' and ';
        '2' : Result := Result + '(YTA_COMMERCIAL="")'                 +' and '+'(YTA_TYPECOMMERCIAL="'  +sTypeCommercial+'")'+' and ';
        '3' : Result := Result + '(YTA_COMMERCIAL="")'                  +' and '+'(YTA_TYPECOMMERCIAL="")'                    +' and ';
      end;
    end;

    //Tiers
    if (sParametre[iPosOrigineTiers]=wTrue) then
    begin
      case sRecherche[iPosOrigineTiers] of
        '1' : Result := Result + '(YTA_TIERS="'  +sTiers+'")'+' and '+'(YTA_TARIFTIERS="")'                 +' and ';
        '2' : Result := Result + '(YTA_TIERS="")'            +' and '+'(YTA_TARIFTIERS="'  +sTarifTiers+'")'+' and ';
        '3' : Result := Result + '(YTA_TIERS="")'            +' and '+'(YTA_TARIFTIERS="")'                 +' and ';
      end;
    end;

    //Affaire
    if (sParametre[iPosOrigineAffaire]=wTrue) then
    begin
      case sRecherche[iPosOrigineAffaire] of
        '1' : Result := Result + '(YTA_AFFAIRE="'  +sAffaire+'")'+' and '+'(YTA_TARIFAFFAIRE="")'                   +' and ';
        '2' : Result := Result + '(YTA_AFFAIRE="")'              +' and '+'(YTA_TARIFAFFAIRE="'  +sTarifAffaire+'")'+' and ';
        '3' : Result := Result + '(YTA_AFFAIRE="")'              +' and '+'(YTA_TARIFAFFAIRE="")'                   +' and ';
      end;
    end;

    //Pole, l'atelier et Groupe
    if (sParametre[iPosOriginePole]=wTrue) then
    begin
      //Pole
      case sRecherche[iPosOriginePole] of
        '1' : Result := Result + '(YTA_POLE="'  +sPole+'")'+' and ';
        '2' : Result := Result + '(YTA_POLE="")'           +' and ';
      end;

      //Atelier
      case sRecherche[iPosOrigineAtelier] of
        '1' : Result := Result + '(YTA_SITE="'  +sAtelier+'")'+' and ';
        '2' : Result := Result + '(YTA_SITE="")'              +' and ';
      end;

      //Groupe
      case sRecherche[iPosOrigineGroupe] of
        '1' : Result := Result + '(YTA_GRP="'  +sGroupe+'")'+' and ';
        '2' : Result := Result + '(YTA_GRP="")'             +' and ';
      end;
    end;

    //Jour de la semaine, l'heure du jour et lieu de consommation
    if (sParametre[iPosOrigineJourSemaine]=wTrue) then
    begin
      //Jour de la semaine
      case sRecherche[iPosOrigineJourSemaine] of
        '1' : Result := Result + '(YTA_JOURSEMAINE="'  +sJourSemaine+'")'+' and ';
        '2' : Result := Result + '(YTA_JOURSEMAINE="")'                  +' and ';
      end;

      //Heur dans la journée
      case sRecherche[iPosOrigineHeureJour] of
        '1' : Result := Result + '(YTA_HEUREJOUR="'  +sHeureJour+'")'+' and ';
        '2' : Result := Result + '(YTA_HEUREJOUR="")'                +' and ';
      end;

      //Lieu de consommation
      case sRecherche[iPosOrigineLieuConso] of
        '1' : Result := Result + '(YTA_LIEUCONSO="'  +sLieuConso+'")'+' and ';
        '2' : Result := Result + '(YTA_LIEUCONSO="")'                +' and ';
      end;
    end;

    Result := copy(Result,1,length(Result)-5)+' )'+' or ';
  end;
end;

{-------------------------------------------------------------------------
  Construction du WHERE selon une combinaison des informations du contexte
--------------------------------------------------------------------------}
function ContexteSurQuoi(sRecherche, sParametre: string; TobContexte: tob; iPartiel: integer) : string;
  {----------------------------------------------------------------------------
  -----------------------------------------------------------------------------}
  function LimiteContexteSurQuoi : boolean;
  begin
    Result :=     (sParametre[iPosOrigineArticle]   =wTrue)
               or (sParametre[iPosOrigineRessource] =wTrue)
               or (sParametre[iPosOrigineMarque]    =wTrue)
               or (sParametre[iPosOrigineCircuit]   =wTrue)
               or (sParametre[iPosOriginePhase]     =wTrue)
               or (sParametre[iPosOrigineTraitement]=wTrue)
              ;
  end;
begin
  Result := '';
  if LimiteContexteSurQuoi then
  begin
    Result := '( ';

    //Article / Article partiel / Famillle tarif article
    if (sParametre[iPosOrigineArticle]=wTrue) then
    begin
      case sRecherche[iPosOrigineArticle] of
        '1' : Result := Result + '(YTQ_ARTICLE="'  +TobContexte.GetValue('ARTICLE')+'")                 '+' and '+'(YTQ_TARIFARTICLE="")'                                          +' and ';
        '2' : Result := Result + '(YTQ_ARTICLE="'  +copy(TobContexte.GetValue('ARTICLE'),1,iPartiel)+'")'+' and '+'(YTQ_TARIFARTICLE="")'                                          +' and ';
        '3' : Result := Result + '(YTQ_ARTICLE="")'                                                      +' and '+'(YTQ_TARIFARTICLE="'  +TobContexte.GetValue('TARIFARTICLE')+'")'+' and ';
        '4' : Result := Result + '(YTQ_ARTICLE="")'                                                      +' and '+'(YTQ_TARIFARTICLE="")'                                          +' and ';
      end;
    end;

    //Ressource / Famillle tarif ressource
    if (sParametre[iPosOrigineRessource]=wTrue) then
    begin
      case sRecherche[iPosOrigineRessource] of
        '1' : Result := Result + '(YTQ_RESSOURCE="'  +TobContexte.GetValue('RESSOURCE')+'")'+' and '+'(YTQ_TARIFRESSOURCE="")'                                            +' and ';
        '2' : Result := Result + '(YTQ_RESSOURCE="")'                                       +' and '+'(YTQ_TARIFRESSOURCE="'  +TobContexte.GetValue('TARIFRESSOURCE')+'")'+' and ';
        '3' : Result := Result + '(YTQ_RESSOURCE="")'                                       +' and '+'(YTQ_TARIFRESSOURCE="")'                                            +' and ';
      end;
    end;

    //Marque
    if (sParametre[iPosOrigineMarque]=wTrue) then
    begin
      case sRecherche[iPosOrigineMarque] of
        '1' : Result := Result + '(YTQ_MARQUE="'  +TobContexte.GetValue('MARQUE')+'")'+' and ';
        '2' : Result := Result + '(YTQ_MARQUE="")'                                    +' and ';
      end;
    end;

    //Circuit / Phase / Traitement
    if (sParametre[iPosOrigineCircuit]=wTrue) then
    begin

      //Circuit
      case sRecherche[iPosOrigineCircuit] of
        '1' : Result := Result + '(YTQ_CIRCUIT="'     +TobContexte.GetValue('CIRCUIT')   +'")'+' and ';
        '2' : Result := Result + '(YTQ_CIRCUIT="")'                                           +' and ';
      end;

      //Phase
      case sRecherche[iPosOriginePhase] of
        '1' : Result := Result + '(YTQ_PHASE="'       +TobContexte.GetValue('PHASE')     +'")'+' and ';
        '2' : Result := Result + '(YTQ_PHASE="")'                                             +' and ';
      end;

      //Traitement
      case sRecherche[iPosOrigineTraitement] of
        '1' : Result := Result + '(YTQ_TRAITEMENT="'  +TobContexte.GetValue('TRAITEMENT')+'")'+' and ';
        '2' : Result := Result + '(YTQ_TRAITEMENT="")'                                        +' and ';
      end;
    end;

    Result := copy(Result,1,length(Result)-5)+' )'+' or ';
  end;
end;


{-------------------------------------------------------------------------------
  En fin de pièce
    * Regrouper toutes les lignes qui ont le même contexte de système tarifaire pièce
    * Calcul du cumul de ces lignes
    * Recherche des données du système tarifaire pour ce cumul
    * Calcul de la ligne
--------------------------------------------------------------------------------}
function CalculTarifGroupePiece(TobTiers, TobArticles, TobPiece, TobLigneTarif : tob) : boolean;
var
  sFonctionnalite, sRang, sOu, sParametre : string;

  iPartiel : integer;
  iCptLigne          , iCptLigneTarif           : integer;
  iCptLigneATotaliser, iCptLigneTarifATotaliser : integer;
  iCptLigneAAffecter , iCptLigneTarifAAffecter  : integer;

  TobContextePiece, TobContexteLigne : tob;
  TobLigne, TobUneligneTarif         : tob;
  TobContexteLigneATraiter  , TobUneLigneTarifATraiter                       : tob;
  TobContexteLigneATotaliser, TobUneLigneTarifATotaliser, TobLigneATotaliser : tob;
  TobContexteLigneAAffecter , TobUneLigneTarifAAffecter , TobLigneAAffecter  : tob;

  nCumul, nValeur : double;

  TobFourchette   : tob;

  {------------------------------------------------------------------------------------------
    Mémorisation des informations du contexte
    en supprimant les information n'ayant pas participé à la recherche
  -------------------------------------------------------------------------------------------}
  procedure GetTobContexteLigneATraiter;
  begin
    //Mémorisation du contexte de regroupement
    TobContexteLigneATraiter :=  TobContextePiece.FindFirst(['NATUREPIECEG', 'SOUCHE', 'NUMERO', 'INDICEG', 'NUMLIGNE'],[TobUneLigneTarifATraiter.GetValue('GLT_NATUREPIECEG'), TobUneLigneTarifATraiter.GetValue('GLT_SOUCHE'), TobUneLigneTarifATraiter.GetValue('GLT_NUMERO'), TobUneLigneTarifATraiter.GetValue('GLT_INDICEG'), TobUneLigneTarifATraiter.GetValue('GLT_NUMLIGNE')], False);
  end;

  {------------------------------------------------------------------------------------------
  -------------------------------------------------------------------------------------------}
  function MemeInformationsContexte : boolean;
  var
    sOuSomme : string;
  begin
    sOuSomme := copy(TobUneLigneTarifATraiter.GetValue('GLT_OU'),iTailleOu+1, iTailleOu);
    Result := True;
    Result := Result and ((sOuSomme[iPosOrigineDepot         ]='.') or (TobContexteLigneATraiter.GetValue('DEPOT'         )=TobContexteLigneATotaliser.GetValue('DEPOT'         )));
    Result := Result and ((sOuSomme[iPosOrigineTarifDepot    ]='.') or (TobContexteLigneATraiter.GetValue('TARIFDEPOT'    )=TobContexteLigneATotaliser.GetValue('TARIFDEPOT'    )));
    Result := Result and ((sOuSomme[iPosOrigineDevise        ]='.') or (TobContexteLigneATraiter.GetValue('DEVISE'        )=TobContexteLigneATotaliser.GetValue('DEVISE'        )));
    Result := Result and ((sOuSomme[iPosOrigineSpecial       ]='.') or (TobContexteLigneATraiter.GetValue('TARIFSPECIAL'  )=TobContexteLigneATotaliser.GetValue('TARIFSPECIAL'  )));
    Result := Result and ((sOuSomme[iPosOrigineCommercial    ]='.') or (TobContexteLigneATraiter.GetValue('COMMERCIAL'    )=TobContexteLigneATotaliser.GetValue('COMMERCIAL'    )));
    Result := Result and ((sOuSomme[iPosOrigineTypeCommercial]='.') or (TobContexteLigneATraiter.GetValue('TYPECOMMERCIAL')=TobContexteLigneATotaliser.GetValue('TYPECOMMERCIAL')));
    Result := Result and ((sOuSomme[iPosOrigineTiers         ]='.') or (TobContexteLigneATraiter.GetValue('TIERS'         )=TobContexteLigneATotaliser.GetValue('TIERS'         )));
    Result := Result and ((sOuSomme[iPosOrigineTarifTiers    ]='.') or (TobContexteLigneATraiter.GetValue('TARIFTIERS'    )=TobContexteLigneATotaliser.GetValue('TARIFTIERS'    )));
    Result := Result and ((sOuSomme[iPosOrigineAffaire       ]='.') or (TobContexteLigneATraiter.GetValue('AFFAIRE'       )=TobContexteLigneATotaliser.GetValue('AFFAIRE'       )));
    Result := Result and ((sOuSomme[iPosOrigineTarifAffaire  ]='.') or (TobContexteLigneATraiter.GetValue('TARIFAFFAIRE'  )=TobContexteLigneATotaliser.GetValue('TARIFAFFAIRE'  )));
    Result := Result and ((sOuSomme[iPosOriginePole          ]='.') or (TobContexteLigneATraiter.GetValue('POLE'          )=TobContexteLigneATotaliser.GetValue('POLE'          )));
    Result := Result and ((sOuSomme[iPosOrigineAtelier       ]='.') or (TobContexteLigneATraiter.GetValue('ATELIER'       )=TobContexteLigneATotaliser.GetValue('ATELIER'       )));
    Result := Result and ((sOuSomme[iPosOrigineGroupe        ]='.') or (TobContexteLigneATraiter.GetValue('GROUPE'        )=TobContexteLigneATotaliser.GetValue('GROUPE'        )));

    Result := Result and ((sOuSomme[iPosOrigineArticle       ]='.') or (TobContexteLigneATraiter.GetValue('ARTICLE'       )=TobContexteLigneATotaliser.GetValue('ARTICLE'       )));
    Result := Result and ((sOuSomme[iPosOrigineArticle       ]='.') or (TobContexteLigneATraiter.GetValue('ARTICLE'       )=TobContexteLigneATotaliser.GetValue('ARTICLE'       )));
    Result := Result and ((sOuSomme[iPosOrigineTarifArticle  ]='.') or (TobContexteLigneATraiter.GetValue('TARIFARTICLE'  )=TobContexteLigneATotaliser.GetValue('TARIFARTICLE'  )));
    Result := Result and ((sOuSomme[iPosOrigineMarque        ]='.') or (TobContexteLigneATraiter.GetValue('MARQUE'        )=TobContexteLigneATotaliser.GetValue('MARQUE'        )));
    Result := Result and ((sOuSomme[iPosOrigineCircuit       ]='.') or (TobContexteLigneATraiter.GetValue('CIRCUIT'       )=TobContexteLigneATotaliser.GetValue('CIRCUIT'       )));
    Result := Result and ((sOuSomme[iPosOriginePhase         ]='.') or (TobContexteLigneATraiter.GetValue('PHASE'         )=TobContexteLigneATotaliser.GetValue('PHASE'         )));
    Result := Result and ((sOuSomme[iPosOrigineTraitement    ]='.') or (TobContexteLigneATraiter.GetValue('TRAITEMENT'    )=TobContexteLigneATotaliser.GetValue('TRAITEMENT'    )));
  end;

  {------------------------------------------------------------------------------------------
    A partir d'une origine
    Construction de la requete pour recherche une donnée
    Retourne l'origne avec laquelle le système a trouver la donnée
  -------------------------------------------------------------------------------------------}
  function VaChercherValeur(sRang: string; var nValeur: double) : string;
  var
    sOuSomme              : string;
    sRequeteYTA, sAQui    : string;
    sRequeteYTQ, sSurQuoi : string;
    sRequeteYTF           : string;
    sChampsValeur         : string;
    sChampsCondition      : string;
    sQuelleDonnee         : string;
    sQuelleRemise         : string;
  begin

    Result := '';

    sQuelleDonnee := sRang[2];
    sQuelleRemise := sRang[4];
    if      (sQuelleDonnee='1') then
      begin sChampsValeur := 'YTF_FORFAIT'      ; sChampsCondition := 'YTF_CONDFORFAIT' ; end
    else if (sQuelleDonnee='2') then
      begin sChampsValeur := 'YTF_COUTLANCEMENT'; sChampsCondition := 'YTF_CONDCOUTLANC'; end
    else if (sQuelleDonnee='3') then
      begin sChampsValeur := 'YTF_REMISEMONTANT'; sChampsCondition := 'YTF_CONDREMISEM' ; end
    else if (sQuelleDonnee='4') then
      begin sChampsValeur := 'YTF_PRIXBRUT'     ; sChampsCondition := 'YTF_CONDPRIXBRUT'; end
    else if (sQuelleDonnee='6') then
      begin
        if      (sQuelleRemise='1') then
          begin sChampsValeur := 'YTF_REMISE1'      ; sChampsCondition := 'YTF_CONDREMISE1' ; end
        else if (sQuelleRemise='2') then
          begin sChampsValeur := 'YTF_REMISE2'      ; sChampsCondition := 'YTF_CONDREMISE2' ; end
        else if (sQuelleRemise='3') then
          begin sChampsValeur := 'YTF_REMISE3'      ; sChampsCondition := 'YTF_CONDREMISE3' ; end
        else
          begin sChampsValeur := 'YTF_REMISE1'      ; sChampsCondition := 'YTF_CONDREMISE1' ; end
        ;
      end
    else if (sQuelleDonnee='7') then
      begin sChampsValeur := 'YTF_PRIXNET'      ; sChampsCondition := 'YTF_CONDPRIXNET' ; end
    ;

    sOuSomme := copy(TobUneLigneTarifATraiter.GetValue('GLT_OU'),iTailleOu+1, iTailleOu);

    sRequeteYTA := '( '+ 'YTA_FONCTIONNALITE="'+TobContexteLigneATraiter.GetValue('FONCTIONNALITE')+'"'+' ) and ';

    sAqui := sOrigineMasque;
    sAqui := Remplace(sAqui, iif(sOuSomme[iPosOrigineDepot]      <>'.', '1', iif(sOuSomme[iPosOrigineTypeCommercial]<>'.', '2', '3')), iPosOrigineDepot      );
    sAqui := Remplace(sAqui, iif(sOuSomme[iPosOrigineDevise]     <>'.', '1', '2')                                                    , iPosOrigineDevise     );
    sAqui := Remplace(sAqui, iif(sOuSomme[iPosOrigineSpecial]    <>'.', '1', '2')                                                    , iPosOrigineSpecial    );
    sAqui := Remplace(sAqui, iif(sOuSomme[iPosOrigineCommercial] <>'.', '1', iif(sOuSomme[iPosOrigineTypeCommercial]<>'.', '2', '3')), iPosOrigineCommercial );
    sAqui := Remplace(sAqui, iif(sOuSomme[iPosOrigineTiers     ] <>'.', '1', iif(sOuSomme[iPosOrigineTarifTiers    ]<>'.', '2', '3')), iPosOrigineTiers      );
    sAqui := Remplace(sAqui, iif(sOuSomme[iPosOrigineAffaire   ] <>'.', '1', iif(sOuSomme[iPosOrigineTarifAffaire  ]<>'.', '2', '3')), iPosOrigineAffaire    );
    sAqui := Remplace(sAqui, iif(sOuSomme[iPosOriginePole]       <>'.', '1', '2')                                                    , iPosOriginePole       );
    sAqui := Remplace(sAqui, iif(sOuSomme[iPosOrigineAtelier]    <>'.', '1', '2')                                                    , iPosOrigineAtelier    );
    sAqui := Remplace(sAqui, iif(sOuSomme[iPosOrigineGroupe]     <>'.', '1', '2')                                                    , iPosOrigineGroupe     );
    sAqui := Remplace(sAqui, iif(sOuSomme[iPosOrigineJourSemaine]<>'.', '1', '2')                                                    , iPosOrigineJourSemaine);
    sAqui := Remplace(sAqui, iif(sOuSomme[iPosOrigineHeureJour]  <>'.', '1', '2')                                                    , iPosOrigineHeureJour  );
    sAqui := Remplace(sAqui, iif(sOuSomme[iPosOrigineLieuConso]  <>'.', '1', '2')                                                    , iPosOrigineLieuConso  );

    sRequeteYTA := sRequeteYTA + ContexteAqui(sAQui,
                                              sParametre,
                                              TobContexteLigneATraiter.GetValue('DEPOT'),
                                              TobContexteLigneATraiter.GetValue('TARIFDEPOT'),
                                              TobContexteLigneATraiter.GetValue('DEVISE'),
                                              TobContexteLigneATraiter.GetValue('TARIFSPECIAL'),
                                              TobContexteLigneATraiter.GetValue('COMMERCIAL'),
                                              TobContexteLigneATraiter.GetValue('TYPECOMMERCIAL'),
                                              TobContexteLigneATraiter.GetValue('TIERS'),
                                              TobContexteLigneATraiter.GetValue('TARIFTIERS'),
                                              TobContexteLigneATraiter.GetValue('AFFAIRE'),
                                              TobContexteLigneATraiter.GetValue('TARIFAFFAIRE'),
                                              TobContexteLigneATraiter.GetValue('POLE'),
                                              TobContexteLigneATraiter.GetValue('ATELIER'),
                                              TobContexteLigneATraiter.GetValue('GROUPE'),
                                              TobContexteLigneATraiter.GetValue('JOURSEMAINE'),
                                              TobContexteLigneATraiter.GetValue('HEUREJOUR'),
                                              TobContexteLigneATraiter.GetValue('LIEUCONSO')
                                             );

    sRequeteYTA := copy(sRequeteYTA,1,length(sRequeteYTA)-4);

    sRequeteYTA := sRequeteYTA +' and ';
    if (sOuSomme[iPosOrigineDateDeReference]='C') then
      sRequeteYTA := sRequeteYTA + '((YTA_DATEREFERENCE="CDE") and (YTA_DATEDEBUT<="'+USDateTime(TobContexteLigneATraiter.GetValue('DATECOMMANDE' ))+'") and ('+'YTA_DATEFIN > "'+USDateTime(TobContexteLigneATraiter.GetValue('DATECOMMANDE' ))+'"))'
    else
      sRequeteYTA := sRequeteYTA + '((YTA_DATEREFERENCE="DEL") and (YTA_DATEDEBUT<="'+USDateTime(TobContexteLigneATraiter.GetValue('DATELIVRAISON'))+'") and ('+'YTA_DATEFIN > "'+USDateTime(TobContexteLigneATraiter.GetValue('DATELIVRAISON'))+'"))';

    sRequeteYTA := sRequeteYTA +' and ';
    if (sOuSomme[iPosOrigineTTCouHT]='H') then
      sRequeteYTA := sRequeteYTA + ' ( YTA_TTCOUHT="H" ) '
    else
      sRequeteYTA := sRequeteYTA + ' ( YTA_TTCOUHT<>"H" ) ';

    if (sRequeteYTA<>'') then
    begin
      sRequeteYTQ := '';
      sSurQuoi := sOrigineMasque;
      sSurQuoi := Remplace(sSurQuoi, iif(sOuSomme[iPosOrigineArticle   ] ='A', '1', iif(sOuSomme[iPosOrigineTarifArticle]   ='P', '2', iif(sOuSomme[iPosOrigineTarifArticle]<>'.', '3', '4'))), iPosOrigineArticle   );
      sSurQuoi := Remplace(sSurQuoi, iif(sOuSomme[iPosOrigineRessource ]<>'.', '1', iif(sOuSomme[iPosOrigineTarifRessource]<>'.', '2', '3')), iPosOrigineRessource );
      sSurQuoi := Remplace(sSurQuoi, iif(sOuSomme[iPosOrigineMarque    ]<>'.', '1', '2'), iPosOrigineMarque    );
      sSurQuoi := Remplace(sSurQuoi, iif(sOuSomme[iPosOrigineCircuit   ]<>'.', '1', '2'), iPosOrigineCircuit   );
      sSurQuoi := Remplace(sSurQuoi, iif(sOuSomme[iPosOriginePhase     ]<>'.', '1', '2'), iPosOriginePhase     );
      sSurQuoi := Remplace(sSurQuoi, iif(sOuSomme[iPosOrigineTraitement]<>'.', '1', '2'), iPosOrigineTraitement);

      sRequeteYTQ := sRequeteYTQ + ContexteSurquoi(sSurQuoi, sParametre, TobContexteLigneATraiter, iPartiel);

      sRequeteYTQ := copy(sRequeteYTQ,1,length(sRequeteYTQ)-4);

      if (sRequeteYTA<>'') then
      begin
        sRequeteYTF := 'SELECT '+sChampsValeur+', '+sChampsCondition+' '+
                       'FROM YTARIFSAQUI '+
                       'INNER JOIN YTARIFSSURQUOI    ON (YTQ_IDENTIFIANTYTA=YTA_IDENTIFIANT) '+
                       'INNER JOIN YTARIFSFOURCHETTE ON (YTF_IDENTIFIANTYTQ=YTQ_IDENTIFIANT) '+
                       'WHERE '+
                          '('+sRequeteYTA+')'+' and '+
                          '('+sRequeteYTQ+')'+' and '+
                          '(YTF_FOURCHETTEMINI<='+StrFPoint(nCumul)+')'+ 'and' +
                          '(YTF_FOURCHETTEMAXI> '+StrFPoint(nCumul)+')';

        TobFourchette  := Tob.Create('_TarifsFourchette_' ,nil ,-1);
        Try
          if TobFourchette.LoadDetailDBFromSql('_TarifsFourchette_', sRequeteYTF) then
          begin
            if (TobFourchette.Detail[0].GetValue(sChampsCondition)=wTrue) then
            begin
              Result  := sOuSomme;
              nValeur := TobFourchette.Detail[0].GetValue(sChampsValeur);
            end;
          end;
        finally
          TobFourchette.Free;
        end;
      end;
    end;
  end;

begin {CalculTarifGroupePiece}
  Result := False;
  iPartiel := GetParamSoc('SO_TARIFPARTIEL');

  if (PGIAsk('Confirmez vous le calcul groupé des tarifs ?','Système tarifaire')=mrYes) then
  begin
    InitMoveProgressForm(nil, TraduireMemoire('Calcul groupé des tarifs'), '', TobPiece.Detail.Count*3, False, True);

    sFonctionnalite  := RechercheFonctionnalite(TobPiece.GetValue('GP_NATUREPIECEG'));
    sParametre       := GetParametresFromYFO(sFonctionnalite);

    TobContextePiece  := Tob.Create('_TobContextePiece_', nil, -1);
    sRang := '1';

    //Pour chaque ligne de commande création d'une TobContexteLigne pour pouvoir lancer wRechercheTarif
    //Pour chaque ligne de commande annulation de l'origine de la recherche de tarif dans le cas d'un contexte groupé pièce
    for iCptLigne:=0 to TobPiece.Detail.Count-1 do
    begin
      MoveCurProgressForm('');
      TobLigne := TobPiece.Detail[iCptLigne];

      //Pour chaque ligne de commande création d'une TobContexteLigne pour pouvoir lancer wRechercheTarif
      TobContexteLigne := Tob.Create('_TobContexteLigne_', TobContextePiece, -1);
      TobContexteLigne.AddChampSupValeur('NATUREPIECEG', TobLigne.GetValue('GL_NATUREPIECEG'), false);
      TobContexteLigne.AddChampSupValeur('SOUCHE'      , TobLigne.GetValue('GL_SOUCHE'      ), false);
      TobContexteLigne.AddChampSupValeur('NUMERO'      , TobLigne.GetValue('GL_NUMERO'      ), false);
      TobContexteLigne.AddChampSupValeur('INDICEG'     , TobLigne.GetValue('GL_INDICEG'     ), false);
      TobContexteLigne.AddChampSupValeur('NUMLIGNE'    , TobLigne.GetValue('GL_NUMLIGNE'    ), false);
      TobContexteLigne.AddChampSupValeur('TOTALISEE'   , wFalse                              , false);
      TobContexteLigne.AddChampSupValeur('AFFECTEE'    , wFalse                              , false);
      GetTobContexteFromPiece(sFonctionnalite, sRang, TobContexteLigne, TobPiece, TobLigne);

      //Pour chaque ligne de commande annulation de l'origine de la recherche de tarif dans le cas d'un contexte groupé pièce
      TobUneLigneTarif := GetLigneTarif(TobLigneTarif, TobLigne);
      if assigned(TobUneLigneTarif) then
      begin
        for iCptLigneTarif:=0 to TobUneLigneTarif.Detail.Count-1 do
        begin
            if  (    (TobUneLigneTarif.Detail[iCptLigneTarif].GetValue('GLT_FONCTIONNALITE')=sFonctionnalite)
                 and (copy(TobUneLigneTarif.Detail[iCptLigneTarif].GetValue('GLT_RANG'),1, length(sRang)) = sRang)
                 and (copy(TobUneLigneTarif.Detail[iCptLigneTarif].GetValue('GLT_OU'),iTailleOu+1, iTailleOu)<>'')
                 )
            then
              TobUneLigneTarif.Detail[iCptLigneTarif].PutValue('GLT_OU', sOrigineMasque+copy(TobUneLigneTarif.Detail[iCptLigneTarif].GetValue('GLT_OU'),iTailleOu+1, iTailleOu));
        end;
      end;
    end;

    //Recherche d'une donnée d'une ligne à traiter
    for iCptLigne:=0 to TobPiece.Detail.Count-1 do
    begin
      MoveCurProgressForm('');
      TobLigne := TobPiece.Detail[iCptLigne];
      TobUneLigneTarif := GetLigneTarif(TobLigneTarif, TobLigne);
      if Assigned(TobUneLigneTarif) then
      begin
        for iCptLigneTarif:=0 to TobUneLigneTarif.Detail.Count-1 do
        begin
          //Pour la fonctionnalite demandée
          //Pour le groupe demandé
          //Si origine de la recherche est non renseignée (vide !) alors on a trouver une ligne à traiter
          if  (    (TobUneLigneTarif.Detail[iCptLigneTarif].GetValue('GLT_FONCTIONNALITE')=sFonctionnalite)
               and (copy(TobUneLigneTarif.Detail[iCptLigneTarif].GetValue('GLT_RANG'),1 ,length(sRang)) = sRang         )
               and (copy(TobUneLigneTarif.Detail[iCptLigneTarif].GetValue('GLT_OU')  ,1 ,iTailleOu    ) = sOrigineMasque)
               )
          then
          begin
            //Mémorisation d'une ligne représentant une donnée à traiter
            TobUneLigneTarifATraiter := TobUneLigneTarif.Detail[iCptLigneTarif];

            //Mémorisation des informations du contexte en supprimé les informations n'ayant pas participées à la recherche
            GetTobContexteLigneATraiter;

            //Avant totalisation RaZ du cumul et du flag de ligne ayant participé à la totalisation
            nCumul := 0;
            for iCptLigneATotaliser:=0 to TobContextePiece.Detail.Count-1 do
            begin
              TobContextePiece.Detail[iCptLigneATotaliser].PutValue('TOTALISEE', wFalse);
            end;

            //Recherche et totalisation de toutes les autres lignes correspondantes au contexte de regroupement
            for iCptLigneATotaliser:=0 to TobPiece.Detail.Count-1 do
            begin
              TobLigneATotaliser         := TobPiece.Detail[iCptLigneATotaliser];
              TobUneLigneTarifATotaliser := GetLigneTarif(TobLigneTarif, TobLigneATotaliser);
              TobContexteLigneATotaliser := TobContextePiece.FindFirst(['NATUREPIECEG', 'SOUCHE', 'NUMERO', 'INDICEG', 'NUMLIGNE'],[TobLigneATotaliser.GetValue('GL_NATUREPIECEG'), TobLigneATotaliser.GetValue('GL_SOUCHE'), TobLigneATotaliser.GetValue('GL_NUMERO'), TobLigneATotaliser.GetValue('GL_INDICEG'), TobLigneATotaliser.GetValue('GL_NUMLIGNE')], False);
              if Assigned(TobUneLigneTarifATotaliser) then
              begin
                for iCptLigneTarifATotaliser := 0 to TobUneLigneTarifATotaliser.Detail.Count-1 do
                begin
                  //   Si même donnée (Forfait, Cout lancement, remise en montant, prix brut, remise, prix net ...)
                  //et Si même contexte de totalisation groupé pièce
                  if  (    (TobUneLigneTarifATraiter.GetValue('GLT_RANG')=TobUneLigneTarifATotaliser.Detail[iCptLigneTarifATotaliser].GetValue('GLT_RANG'))
                       and (copy(TobUneLigneTarifATraiter.GetValue('GLT_OU'),iTailleOu+1,iTailleOu)=Copy(TobUneLigneTarifATotaliser.Detail[iCptLigneTarifATotaliser].GetValue('GLT_OU'),iTailleOu+1,iTailleOu))
                      ) then
                  begin
                    //Si même information de contexte alors on totalise
                    if MemeInformationsContexte then
                    begin
                      TobContexteLigneAtotaliser.PutValue('TOTALISEE', wTrue);
                      nCumul := nCumul+Quoi(copy(copy(TobUneLigneTarifATotaliser.Detail[iCptLigneTarifATotaliser].GetValue('GLT_OU'),iTailleOu+1, iTailleOu), iPosOrigineSommerQuoi, 1),TobContexteLigneATotaliser);
                    end;
                  end;
                end;
              end;
            end;

            //Avec le cumul calculé
            //Recherche de la donnée concerné
            sOu:=VaChercherValeur( TobUneLigneTarifATraiter.GetValue('GLT_RANG'), nValeur);
            if (sOu<>'') then
            begin
              //Reprise de toutes les lignes ayant participées à la totalisation pour affectation
              for iCptLigneAAffecter:=0 to TobPiece.Detail.Count-1 do
              begin
                if (TobContextePiece.Detail[iCptLigneAAffecter].GetValue('TOTALISEE')=wTrue) then
                begin
                  TobLigneAAffecter         := TobPiece.Detail[iCptLigneAAffecter];
                  TobUneLigneTarifAAffecter := GetLigneTarif(TobLigneTarif, TobLigneAAffecter);
                  TobContexteLigneAAffecter := TobContextePiece.FindFirst(['NATUREPIECEG', 'SOUCHE', 'NUMERO', 'INDICEG', 'NUMLIGNE'],[TobLigneAAffecter.GetValue('GL_NATUREPIECEG'), TobLigneAAffecter.GetValue('GL_SOUCHE'), TobLigneAAffecter.GetValue('GL_NUMERO'), TobLigneAAffecter.GetValue('GL_INDICEG'), TobLigneAAffecter.GetValue('GL_NUMLIGNE')], False);
                  if Assigned(TobUneligneTarifAAffecter) then
                  begin
                    for iCptLigneTarifAAffecter := 0 to TobUneLigneTarifAAffecter.Detail.Count-1 do
                    begin
                      //   Si même donnée (Forfait, Cout lancement, remise en montant, prix brut, remise, prix net ...)
                      //et Si même contexte de totalisation groupé pièce
                      if  (    (TobUneLigneTarifATraiter.GetValue('GLT_RANG')=TobUneLigneTarifAAffecter.Detail[iCptLigneTarifAAffecter].GetValue('GLT_RANG'))
                           and (copy(TobUneLigneTarifATraiter.GetValue('GLT_OU'),iTailleOu+1,iTailleOu)=Copy(TobUneLigneTarifAAffecter.Detail[iCptLigneTarifAAffecter].GetValue('GLT_OU'),iTailleOu+1,iTailleOu))
                          ) then
                      begin
                        if (copy(TobUneLigneTarifATraiter.GetValue('GLT_OU'),iTailleOu+iPosOrigineRechercherOu,1)=sOu_P) then
                        begin
                          //Avec le cumul calculé en pièce groupée
                          //Affectation de la valeur de ce contexte
                          TobContexteLigneAAffecter.PutValue('AFFECTEE', wTrue);
                          TobUneLigneTarifAAffecter.Detail[iCptLigneTarifAAffecter].PutValue('GLT_VALEUR',nValeur);
                          TobUneLigneTarifAAffecter.Detail[iCptLigneTarifAAffecter].PutValue('GLT_OU'    ,sOu+copy(TobUneLigneTarifAAffecter.Detail[iCptLigneTarifAAffecter].GetValue('GLT_OU'),iTailleOu+1,iTailleOu)+FloatToStr(nCumul));
                          TobUneLigneTarifAAffecter.Detail[iCptLigneTarifAAffecter].PutValue('GLT_CUMUL' ,nCumul);
                        end
                        else
                        begin
                          //Avec le cumul calculer en groupé pièce
                          //Recherche ligne à ligne
                          TobUneLigneTarifAAffecter.Detail[iCptLigneTarifAAffecter].PutValue('GLT_CUMUL',nCumul);
                        end;
                      end;
                    end;
                  end;
                end;
              end;
            end;
          end;
        end;
      end;
    end;

    //Une fois que la TobLigneTarif est actualisée en fonction du groupée commande
    //MaJ de la TobLigne à partir de la TobLigneTarif
    for iCptLigne:=0 to TobPiece.Detail.Count-1 do
    begin
      MoveCurProgressForm('');
      TobLigne := TobPiece.Detail[iCptLigne];
      TobUneLigneTarif := GetLigneTarif(TobLigneTarif, TobLigne);
      TobContexteLigne := TobContextePiece.FindFirst(['NATUREPIECEG', 'SOUCHE', 'NUMERO', 'INDICEG', 'NUMLIGNE'],[TobLigne.GetValue('GL_NATUREPIECEG'), TobLigne.GetValue('GL_SOUCHE'), TobLigne.GetValue('GL_NUMERO'), TobLigne.GetValue('GL_INDICEG'), TobLigne.GetValue('GL_NUMLIGNE')], False);
      //Si ligne avec un changelent tarifiare du au groupé commande
      if (TobContexteLigne.GetValue('AFFECTEE')=wTrue) and Assigned(TobUneLigneTarif) then
      begin
        //Mise à Jour de TobLigne
        //Recalcul de la ligne
        for iCptLigneTarif:=0 to TobUneLigneTarif.Detail.Count-1 do
        begin
          //Si le système tarifaire a trouvé des prix net ou brut alors on calcule avec ces prix sinon on prend le prix saisi
          nValeur := TobUneLigneTarif.Detail[iCptLigneTarif].GetValue('GLT_VALEUR');
          if      (copy(TobUneLigneTarif.Detail[iCptLigneTarif].GetValue('GLT_RANG'),1,2) = sRang+'1') then
            TobLigne.PutValue('GL_VALEURFIXEDEV', nValeur)
          else if (copy(TobUneLigneTarif.Detail[iCptLigneTarif].GetValue('GLT_RANG'),1,2) = sRang+'2') then
            TobLigne.PutValue('GL_VALEURFIXEDEV', nValeur)
          else if (copy(TobUneLigneTarif.Detail[iCptLigneTarif].GetValue('GLT_RANG'),1,2) = sRang+'3') then
            TobLigne.PutValue('GL_VALEURREMDEV' , nValeur)
          else if (copy(TobUneLigneTarif.Detail[iCptLigneTarif].GetValue('GLT_RANG'),1,2) = sRang+'4') then
          begin
            if (TobContexteLigne.GetValue('FACTUREHT')=wTrue) then
              TobLigne.PutValue('GL_PUHTDEV', nValeur)
            else
              TobLigne.PutValue('GL_PUTTCDEV', nValeur);
          end
          else if (copy(TobUneLigneTarif.Detail[iCptLigneTarif].GetValue('GLT_RANG'),1,2) = sRang+'5') then
          else if (copy(TobUneLigneTarif.Detail[iCptLigneTarif].GetValue('GLT_RANG'),1,2) = sRang+'6') then
            TobLigne.PutValue('GL_REMISELIGNE',(CalculRemiseTotaleSurTobLigneTarif(TobLigne, TobLigneTarif, sRang+'6')))
          else if (copy(TobUneLigneTarif.Detail[iCptLigneTarif].GetValue('GLT_RANG'),1,2) = sRang+'7') then
          begin
            if (TobContexteLigne.GetValue('FACTUREHT')=wTrue) then
              TobLigne.PutValue('GL_PUHTDEV', nValeur)
            else
              TobLigne.PutValue('GL_PUTTCDEV', nValeur);
          end;
        end;
        TobLigne.PutValue('GL_RECALCULER',wTrue);
      end;
    end;
    TobContextePiece.Free;
    FiniMoveProgressForm;
  end;
end; {CalculTarifGroupePiece}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Marc Morretton
Créé le ...... : 11/06/2003
Description .. : Teste existence d'une affaire
*****************************************************************}
function ExistAffaire(sAffaire : string; WithAlert: Boolean = false):Boolean;
var
	sSql  : string;
begin
	sSql := ' SELECT AFF_AFFAIRE'
			+ ' FROM AFFAIRE'
			+ ' WHERE AFF_AFFAIRE="' + sAffaire +'"';
	Result := ExisteSQL(sSql);

  if WithAlert and (not Result) then
  begin
    PgiError(TraduireMemoire('L''affaire ') + sAffaire + TraduireMemoire(' n''existe pas.'), 'Affaire');
  end;
end;


{***********A.G.L.***********************************************
Auteur  ...... : Marc MORRETTON
Créé le ...... : 07/11/2003
Modifié le ... : 18/03/2004
Description .. : Renvoie la série de paramètre de la fonctionnalité
Suite ........ : demandée et du type de paramètre demandé
Mots clefs ... :
*****************************************************************}
function GetParametresFromYFOPParParametre(TobYFO : tob; sQuelParametre : string) : string;
var
  iParametre : integer;
begin {GetParametresFromYFOPParParametre}
  if      (sQuelParametre='EntitéActive') then iParametre:=1
  else if (sQuelParametre='MulSélection') then iParametre:=2
  else if (sQuelParametre='SaisieEntête') then iParametre:=3
  else if (sQuelParametre='ClefEntête'  ) then iParametre:=4
  else                                         iParametre:=1;
  Result := sOrigineMasque;
  Result := Remplace(Result, copy(TobYFO.Detail[0].GetString('YFO_OKCODEPORT')    ,iParametre, 1) , iPosCodePort       );
  Result := Remplace(Result, copy(TobYFO.Detail[0].GetString('YFO_OKDEPOT')       ,iParametre, 1) , iPosDepot          );
  Result := Remplace(Result, copy(TobYFO.Detail[0].GetString('YFO_OKDEPOT')       ,iParametre, 1) , iPosTarifDepot     );
  Result := Remplace(Result, copy(TobYFO.Detail[0].GetString('YFO_OKDEVISE')      ,iParametre, 1) , iPosDevise         );
  Result := Remplace(Result, copy(TobYFO.Detail[0].GetString('YFO_OKSPECIAL')     ,iParametre, 1) , iPosSpecial        );
  Result := Remplace(Result, copy(TobYFO.Detail[0].GetString('YFO_OKTIERS')       ,iParametre, 1) , iPosTiers          );
  Result := Remplace(Result, copy(TobYFO.Detail[0].GetString('YFO_OKTIERS')       ,iParametre, 1) , iPosTarifTiers     );
  Result := Remplace(Result, copy(TobYFO.Detail[0].GetString('YFO_OKSECTEURGEO')  ,iParametre, 1) , iPosSecteurGeo     );
  Result := Remplace(Result, copy(TobYFO.Detail[0].GetString('YFO_OKADRESSE')     ,iParametre, 1) , iPosCodePostal     );
  Result := Remplace(Result, copy(TobYFO.Detail[0].GetString('YFO_OKADRESSE')     ,iParametre, 1) , iPosPays           );
  Result := Remplace(Result, copy(TobYFO.Detail[0].GetString('YFO_OKADRESSE')     ,iParametre, 1) , iPosRegion         );
  Result := Remplace(Result, copy(TobYFO.Detail[0].GetString('YFO_OKMODEEXP')     ,iParametre, 1) , iPosModeExp        );
  Result := Remplace(Result, copy(TobYFO.Detail[0].GetString('YFO_OKAFFAIRE')     ,iParametre, 1) , iPosAffaire        );
  Result := Remplace(Result, copy(TobYFO.Detail[0].GetString('YFO_OKAFFAIRE')     ,iParametre, 1) , iPosTarifAffaire   );
  Result := Remplace(Result, copy(TobYFO.Detail[0].GetString('YFO_OKATELIER')     ,iParametre, 1) , iPosAtelier        );
  Result := Remplace(Result, copy(TobYFO.Detail[0].GetString('YFO_OKDATE')        ,iParametre, 1) , iPosDatesDeValidite);

  Result := Remplace(Result, copy(TobYFO.Detail[0].GetString('YFO_OKNATURETRA')   ,iParametre, 1) , iPosNatureTra      );
  Result := Remplace(Result, copy(TobYFO.Detail[0].GetString('YFO_OKARTICLE')     ,iParametre, 1) , iPosArticle        );
  Result := Remplace(Result, copy(TobYFO.Detail[0].GetString('YFO_OKARTICLE')     ,iParametre, 1) , iPosTarifArticle   );
  Result := Remplace(Result, copy(TobYFO.Detail[0].GetString('YFO_OKFAMILLENIV')  ,iParametre, 1) , iPosFamilleNiv1    );
  Result := Remplace(Result, copy(TobYFO.Detail[0].GetString('YFO_OKFAMILLENIV')  ,iParametre, 1) , iPosFamilleNiv2    );
  Result := Remplace(Result, copy(TobYFO.Detail[0].GetString('YFO_OKFAMILLENIV')  ,iParametre, 1) , iPosFamilleNiv3    );
  Result := Remplace(Result, copy(TobYFO.Detail[0].GetString('YFO_OKRESSOURCE')   ,iParametre, 1) , iPosRessource      );
  Result := Remplace(Result, copy(TobYFO.Detail[0].GetString('YFO_OKRESSOURCE')   ,iParametre, 1) , iPosTarifRessource );
  Result := Remplace(Result, copy(TobYFO.Detail[0].GetString('YFO_OKCPT')         ,iParametre, 1) , iPosCircuit        );
  Result := Remplace(Result, copy(TobYFO.Detail[0].GetString('YFO_OKCPT')         ,iParametre, 1) , iPosPhase          );
  Result := Remplace(Result, copy(TobYFO.Detail[0].GetString('YFO_OKCPT')         ,iParametre, 1) , iPosTraitement     );

  Result := Remplace(Result, copy(TobYFO.Detail[0].GetString('YFO_OKHTOUTTC')     ,iParametre, 1) , iPosTTCouHT        );
  Result := Remplace(Result, copy(TobYFO.Detail[0].GetString('YFO_OKLIBELLE')     ,iParametre, 1) , iPosLibelle        );
  Result := Remplace(Result, copy(TobYFO.Detail[0].GetString('YFO_OKREMISETIERS') ,iParametre, 1) , iPosRemiseTiers    );
  Result := Remplace(Result, copy(TobYFO.Detail[0].GetString('YFO_OKCASCREMISE')  ,iParametre, 1) , iPosCascContexte   );
  Result := Remplace(Result, copy(TobYFO.Detail[0].GetString('YFO_OKPARAMETRES')  ,iParametre, 1) , iPosParametres     );
  Result := Remplace(Result, copy(TobYFO.Detail[0].GetString('YFO_OKCONDREELLE')  ,iParametre, 1) , iPosCondReelle     );
  Result := Remplace(Result, copy(TobYFO.Detail[0].GetString('YFO_OKACTIF')       ,iParametre, 1) , iPosActif          );
  Result := Remplace(Result, copy(TobYFO.Detail[0].GetString('YFO_OKPRIXMARGE')   ,iParametre, 1) , iPosPrixVenteMarge );
  Result := Remplace(Result, copy(TobYFO.Detail[0].GetString('YFO_OKTARIFMARCHE') ,iParametre, 1) , iPosTarifMarche    );

  Result := Remplace(Result, copy(TobYFO.Detail[0].GetString('YFO_OKFORFAIT')     ,iParametre, 1) , iPosForfait        );
  Result := Remplace(Result, copy(TobYFO.Detail[0].GetString('YFO_OKCOUTLANC')    ,iParametre, 1) , iPosCoutLancement  );
  Result := Remplace(Result, copy(TobYFO.Detail[0].GetString('YFO_OKREMISEMT')    ,iParametre, 1) , iPosRemiseEnMontant);
  Result := Remplace(Result, copy(TobYFO.Detail[0].GetString('YFO_OKPRIXBRUT')    ,iParametre, 1) , iPosPrixBrut       );
  Result := Remplace(Result, copy(TobYFO.Detail[0].GetString('YFO_OKREMISE1')     ,iParametre, 1) , iPosRemise1        );
  Result := Remplace(Result, copy(TobYFO.Detail[0].GetString('YFO_OKREMISE2')     ,iParametre, 1) , iPosRemise2        );
  Result := Remplace(Result, copy(TobYFO.Detail[0].GetString('YFO_OKREMISE3')     ,iParametre, 1) , iPosRemise3        );
  Result := Remplace(Result, copy(TobYFO.Detail[0].GetString('YFO_OKPRIXNET')     ,iParametre, 1) , iPosPrixNet        );

  Result := Remplace(Result, copy(TobYFO.Detail[0].GetString('YFO_CASCPARENTITE') ,iParametre, 1) , iPosCascadeParEntite   );
  Result := Remplace(Result, copy(TobYFO.Detail[0].GetString('YFO_HIERARCHIQUE')  ,iParametre, 1) , iPosFamilleHierarchique);
end; {GetParametresFromYFOPParParametre}

{***********A.G.L.***********************************************
Auteur  ...... : Marc MORRETTON
Créé le ...... : 07/11/2003
Modifié le ... : 07/11/2003
Description .. : Remplace le code d'une position sur le champs origine de la
Suite ........ : recherche
Mots clefs ... :
*****************************************************************}
function Remplace(sOrigine, sCode: string; iPos: integer) : string;
begin {Remplace}
  Result := copy(sOrigine,1,iPos-1)+sCode+copy(sOrigine,iPos+length(sCode),Length(sOrigine));
end; {Remplace}

Initialization
	TransfertTobs := TList.Create;

Finalization
	TransfertTobs.Free;

end.
