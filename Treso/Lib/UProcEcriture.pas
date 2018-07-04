{-------------------------------------------------------------------------------------
    Version  |   Date | Qui | Commentaires
--------------------------------------------------------------------------------------
6.00.001.001  28/06/04  JP   Création de l'unité : contient les routines de manipulation des écritures,
                             qu'il s'agisse des écritures de comptabilité ou de trésorerie
7.00.001.001  12/01/06  JP   FQ 10322 : Gestion du champ E_ANA
7.00.001.001  12/01/06  JP   FQ 10323 : Correction de la gestion de la TVA
7.00.001.001  12/01/06  JP   FQ 10324 : Gestion de la TVA pour les TIC / TID
7.00.001.001  12/01/06  JP   FQ 10325 : Gestion de la TVA pour les divers lettrable
7.09.001.001  04/08/06  JP   Gestion du Multi sociétés :
                             1/ Gestion des champs _GENERAL dont la valeur est maintanant BQ_CODE
                             2/ Réécriture CreeTrEcritureFromVirement pour gérer les comptes courants
7.09.001.001  17/08/06  SBO  Nouvelle intégration en comptabilité
8.01.001.010  29/03/04  JP   Nouvelle fonction de création des écritures de Tréso CreerTrEcriture
8.00.001.019  19/06/07  JP   FQ 10478 : gestion des exercices en mono dossier
8.00.001.026  20/07/07  JP   Uniformisation du format des RIB pour la compta
8.10.001.013  11/10/07  JP   FQ 10530 : on peut écrire sur des clôture provisoires  
--------------------------------------------------------------------------------------}
unit UProcEcriture;

interface

uses
  {$IFNDEF EAGLCLIENT}
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  {$ENDIF}
  {$IFDEf VER150}
  Variants,
  {$ENDIF}
  SysUtils, Ent1, HCtrls , Hmsgbox, Hent1, Constantes,
  UTOB, Classes,  HQry,
  uLibECriture,           // TInfoEcriture
  {$IFNDEF NOVH}
  uLibPieceCompta,         // TPieceCompta
  {$ENDIF NOVH}
  {$IFDEF MODENT1}
  CPTypeCons,
  {$ENDIF MODENT1}
  ULibExercice, UObjGen
  ;

{25/06/04 : création des deux écritures de trésorerie à partir d'un virement}
function  CreeTrEcritureFromVirement(TobVir : TOB; ObjDev : TObjDevise; AvecMaj : Boolean;
                                     var SClef, DClef : string) : Boolean;
{12/10/06 : Création des écritures d'ICC}
function  CreeTrEcritureIcc(TobIcc : TOB) : Boolean;
{Retourne le numéro de ligne pour l'écriture comptable à générer}
function  GetNumLigne          (NumTransac, NumPiece : string) : Integer;
{Traite-t-on, en intégration en comptabilité, une nouvelle pièce ?}
function  IsNouvellePiece      (NumTransac, NumPiece : string) : Boolean;
{Réinitialise la variable globale RecordCompta}
procedure ReInitRecordCompta   (Transac, Piece : string);
{06/12/04 : Initialisation d'une nouvelle écriture avec une Tob : Il est temps que je remplace
            progressivement le Record TrEcriture par une Tob réèlle}
procedure InitNlleEcritureTob  (var TobEcriture : TOB; Gene : string; Dossier : string = '');
{05/01/05 : Termine l'écriture initialisée dans InitNlleEcritureTob l'écrit dans la base et
            génère les commissions afférentes}
function  TermineEcritureTob   (var T : TOB; ObjDev : TObjDevise; ComplOk : Boolean) : Boolean;
{08/12/04 : Remplit les champs standards d'une ligne ECRITURE à partir d'une ligne TRECRITURE}
procedure RemplitChpStdTob      (var TobCompta : TOB; TobEcr : TOB);
{23/03/05 : Teste si les écritures de trésorerie sont valides}
function  TesteEcrTresorerie    (T : TOB) : Boolean;
{17/03/05 : FQ 10223 : Génération du message d'erreur lors de la créatio de TrEcritures ou d'Ecritures}
function  AfficheMessageErreur(Titre : string = ''; Entete : string = '') : Boolean;
{21/03/05 : FQ 10223 : Réinitialise les ensembles d'erreur}
procedure InitGestionErreur(CatCur : TCatErreur = CatErr_None);
{31/03/05 : FQ 10223 : Renseigne la structure d'erreur}
procedure SetErreurTreso(Erreur : TNatErreur);
{26/07/05 : FQ 10158 : Mise à jour d'une écriture de compta à partir d'une écriture de Treso}
procedure MajEcritureParClef(Chp, Valeur : string; Clef : TClefCompta);
{20/09/06 : l'écriture est-elle intégrable en comptabilité ?}
function  IsIntegrable(ligneTEcr : TOB) : Boolean;
{29/03/07 : Nouvelle fonction de création des écritures de Tréso CreerTrEcriture}
function  CreerTrEcriture(Cpte, Flux, Libelle, Nature  : string; MntDev : Double;DateC : TDateTime;
                          DateV : TDateTime = 0; Cib : string = ''; Ref : string = ''; Dev : string = '') : TOB;

{ ========================================================
  ========== NOUVELLES FONCTIONS D'INTERGRATION ==========
  ======================================================== }
{$IFNDEF NOVH}
{17/08/2005 : Création d'un TPieceCompta pour intégration}
function  TRCreerPieceCompta ( vInfoEcr : TInfoEcriture ; vParent : Tob ; vRegle : TReglesInteg ) : TPieceCompta ;
{17/08/2005 : Génération des pièces comptables depuis la liste des ecritures de tréso}
function  TRGenererPieceCompta (var TobGen : TobPieceCompta; vTobOrig : Tob) : Boolean ;
{17/08/2005 : Génération de la ligne d'ecriture compta principale depuis la ligne de trésorerie }
procedure TRGenereEcrCompta(var vPiece : TPieceCompta ; vTobTEcr : Tob ; vRegle : TReglesInteg ) ;
{17/08/2005 : Génération de la ligne de contrepartie comptable depuis la ligne de trésorerie }
procedure TRGenereEcrContrepartie(var vPiece : TPieceCompta ; vTobTEcr : Tob ; vRegle : TReglesInteg ) ;
{17/08/2005 : Génération de la ligne de TVA compta depuis la ligne de trésorerie }
procedure TRGenereEcrTVA(var vPiece : TPieceCompta ; vTobTEcr : Tob ; vRegle : TReglesInteg ) ;
{17/08/2005 : Enregistrement des Pieces comptables }
function  TREnregistrePieces ( var TobCompta : TobPieceCompta ; MajCTerme : Boolean ) : Boolean;
{21/08/2005 : Concentre les différentes étapes de l'intégration une fois les pièces comptable générées }
function  TRIntegrationPieces (var TobCompta : TobPieceCompta; MajCTerme : Boolean ) : Boolean;
{17/08/2005 : Remplit les champs standards d'une ligne ECRITURE à partir d'une ligne TRECRITURE}
procedure TRRemplitPieceStd  (var vPiece : TPieceCompta; vNumL : integer ; vTobTEcr : TOB ; vRegle : TReglesInteg );
{$ENDIF NOVH}
{ ======================================================== }

type
  TObjetExercice = class(TPersistent)
  private
    FListeExo : TStringList;
    FExercice : TZExercice;
    FDossier  : string;

    function    GetEnCours : TExoDate;
    function    GetExoV8 : TExoDate;
    function    GetTZExercice(aDossier : string) : TZExercice;
    function    GetExercices(n : Integer) : TExoDate;
  public
    class function GetCurrent : TObjetExercice;
    constructor Create;
    destructor  Destroy; override;
    procedure   ChargeObjet;
    procedure   PositionneExercice(aDossier : string);
    function    GetExoNodos(DateOpe : TDateTime; aNoDossier : string = ''; PourCompta : Boolean = False) : string;
    function    GetExercice(DateOpe : TDateTime; aDossier : string = ''; PourCompta : Boolean = False) : string;
    function    IsExoOuvert(DateOpe : TDateTime; aDossier : string = '') : Boolean;
    function    GetDebExo  (aExo : string; aDossier : string = '') : TDateTime;
    function    GetFinExo  (aExo : string; aDossier : string = '') : TDateTime;

    property    Exercice[adossier : string] : TZExercice read GetTZExercice;
    property    Exercices[index : Integer]  : TExoDate   read GetExercices;
    
    property    Encours  : TExoDate read GetEncours;
    property    ExoV8    : TExoDate read GetExoV8;
    property    Dossier  : string   read FDossier;
  end;

{$IFNDEF EAGLSERVER}
var
  VG_ObjetExo : TObjetExercice;
{$ENDIF EAGLSERVER}

implementation

uses
  Commun, UProcGen, SaisUtil, UProcCommission, ULibAnalytique,
  UProcSolde, ParamSoc, UtilPgi, Controls,
  {$IFDEF EAGLSERVER}
   eSession,
   uHTTP,
  {$ELSE}
  uTobDebug, AglInit,
  {$ENDIF EAGLSERVER}
  ed_tools
  {$IFDEF MODENT1}
  , CPProcGen
  {$ENDIF MODENT1}
  {$IFDEF TRESO}
  {Je le met en directive pour éviter de tout tirer en comptabilité}
  , TRLISTEECRCOMPTA_TOF
  {$ENDIF}
  ;

{---------------------------------------------------------------------------------------}
function CreeTrEcritureFromVirement(TobVir : TOB; ObjDev : TObjDevise; AvecMaj : Boolean;
                                    var SClef, DClef : string) : Boolean;
{---------------------------------------------------------------------------------------}
var
  TobEqui : TOB;
  aCurTob : TOB;
  Obj     : TobjComptesTreso;
  NumTra  : string;
  ObjCom  : TObjCommissionTob;
  CodeUnq : string;
  SQL     : string;

    {------------------------------------------------------------------}
    function Termine(Gen, CtrPt, Dos : string) : Boolean;
    {------------------------------------------------------------------}
    var
      S : string;
      n : Integer;
    begin
      {17/03/05 : FQ 10223 : Refonte de la gestion des erreurs}
      Result := True;

      aCurTob.SetString('TE_NODOSSIER', Dos); {29/05/06 : FQ 10360}
      aCurTob.SetString('TE_DEVISE'   ,	ObjDev.GetDeviseCpt(Gen));
      aCurTob.SetString('TE_GENERAL'  , Gen);
      {TE_CONTREPARTIETR contient des comptes généraux et non des BQ_CODE}
      CtrPt := Obj.GetGeneFromCode(CtrPt);
      aCurTob.SetString('TE_CONTREPARTIETR', CtrPt);
      Obj.SetLeRib(aCurTob);

      {Récupération du libellé du code AFB}
      S := RechDom('TRCIB', aCurTob.GetString('TE_CODECIB'), False);
      aCurTob.SetString('TE_REFINTERNE', Copy(S, 1, 35 - Length(CtrPt) - 1) + ' ' + CtrPt);

      aCurTob.SetString('TE_SOCIETE', GetInfosFromDossier('DOS_NODOSSIER', Dos, 'DOS_SOCIETE'));
      {Crée une valeur unique (varchar(17))}
//      aCurTob.SetString('TE_NUMTRANSAC', CodeModule + aCurTob.GetString('TE_SOCIETE') + TRANSACEQUI + NumTra);
      aCurTob.SetString('TE_NUMTRANSAC', CodeModule + TobVir.GetString('TEQ_SOCIETE') + TRANSACEQUI + NumTra);

      aCurTob.SetString('TE_LIBELLE', 'Equilibrage');
      aCurTob.SetString('TE_USERCREATION', TobVir.GetString('TEQ_USERCREATION'));
      aCurTob.SetString('TE_USERVALID', V_PGI.User);
      aCurTob.SetDateTime('TE_DATECREATION' , V_PGI.DateEntree);
      aCurTob.SetDateTime('TE_DATEVALID'    , V_PGI.DateEntree);
      {26/12/06 : FQ 10387 : Paramétrage de la date d'opération}
      aCurTob.SetDateTime('TE_DATECOMPTABLE', TobVir.GetDateTime('TEQ_DATECREATION'));
      aCurTob.SetDateTime('TE_DATEMODIF'    , iDate1900);
      aCurTob.SetDateTime('TE_DATERAPPRO'   , iDate1900);

      aCurTob.SetDateTime('TE_DATEVALEUR',  CalcDateValeur(aCurTob.GetString('TE_CODECIB'),
                                                           aCurTob.GetString('TE_GENERAL'),
                                                           aCurTob.GetDateTime('TE_DATECOMPTABLE')));

      aCurTob.SetString('TE_EXERCICE', TObjetExercice.GetCurrent.GetExoNodos(aCurTob.GetDateTime('TE_DATECOMPTABLE'),
                                                               aCurTob.GetString('TE_NODOSSIER')));

      aCurTob.SetDouble('TE_SOLDEDEV', 0) ;
      aCurTob.SetDouble('TE_SOLDEDEVVALEUR', 0);
      {Les écritures liées aux virements seront réalisées lors de la génération du virement}
      aCurTob.SetString('TE_NATURE', na_Prevision);

      {Génération des clés}
      CodeUnq := Commun.GetNum(CODEUNIQUE, CODEUNIQUE, CODEUNIQUE);
      aCurTob.SetString('TE_CLEVALEUR'   , RetourneCleEcriture(aCurTob.GetDateTime('TE_DATEVALEUR'), StrToInt(CodeUnq)));
      aCurTob.SetString('TE_CLEOPERATION', RetourneCleEcriture(aCurTob.GetDateTime('TE_DATECOMPTABLE'), StrToInt(CodeUnq)));
      Commun.SetNum(CODEUNIQUE, CODEUNIQUE, CODEUNIQUE, CodeUnq);
      //TobDebug(aCurTob);
      {Vérification de la validité de l'écriture saisie}
      if not TesteEcrTresorerie(aCurTob) then begin
        Result := False;
        Exit;
      end;

      CategorieCurrent := CatErr_COM;
      ObjCom := TObjCommissionTob.Create(aCurTob.GetString('TE_GENERAL'),
                                         aCurTob.GetString('TE_CODEFLUX'),
                                         aCurTob.GetDateTime('TE_DATECOMPTABLE'));
      try
        {16/07/04 : gestion des commissions avant l'écriture de l'opération financière car le numéro
                    de transaction peut être modifié dans ObjCom.GenererCommissions}
        ObjCom.GenererCommissions(aCurTob);
        {07/01/08 : FQ 10544 : la destruction ObjCom détruisant TobEcriture, on ne peut pas changer directement
                    de parent de TobEcriture mais seulement de ses filles}
        for n := 0 to ObjCom.TobEcriture.Detail.Count - 1 do
          ObjCom.TobEcriture.Detail[n].ChangeParent(aCurTob, -1);

        {... écriture du flux}
        CategorieCurrent := CatErr_TRE;
      finally
        if Assigned(ObjCom) then FreeAndNil(ObjCom);
      end;
    end;

    {------------------------------------------------------------------}
    function MajEcrCourantes : Boolean;
    {------------------------------------------------------------------}
    begin
      Result := True;

      {Gestion de la devise : pour les comptes courantes la devises est la devise pivot}
      if aCurTob.GetString('TE_DEVISE') <> V_PGI.DevisePivot then begin
        aCurTob.SetString('TE_DEVISE', V_PGI.DevisePivot);
        {On récupère le montant en devise pivot}
        aCurTob.SetDouble('TE_MONTANTDEV', aCurTob.GetDouble('TE_MONTANT'));
        aCurTob.SetDouble('TE_COTATION', 1); 
      end;

      aCurTob.SetDateTime('TE_DATEVALEUR', aCurTob.GetDateTime('TE_DATECOMPTABLE'));
      aCurTob.SetInteger('TE_NUMECHE', 0); {Pas d'échéance ?}
      {Les montants sont à l'opposé de l'écriture bancaire}
      aCurTob.SetDouble('TE_MONTANT'   , aCurTob.GetDouble('TE_MONTANT') * -1);
      aCurTob.SetDouble('TE_MONTANTDEV', aCurTob.GetDouble('TE_MONTANTDEV') * -1);
      aCurTob.SetString('TE_COMMISSION', suc_SansCom); {Pas de commission}

      {Pour éviter que l'écriture apparaise dans le mul d'intégration, budgétaire, gestion des dates de valeur ...}
      aCurTob.SetString('TE_CODECIB', CODECIBCOURANT); {Moyen le plus simple de distinguer les écritures courantes}
      aCurTob.SetString('TE_NATURE', na_Realise); {Déjà réalisé}
      aCurTob.SetString('TE_USERCOMPTABLE', V_PGI.User); {Déjà intégrée}
      aCurTob.SetDateTime('TE_DATEVALID', Date); {Déjà validée}
      aCurTob.SetDateTime('TE_DATERAPPRO', Date); {Déjà pointée}

      {Gestion des clefs : les deux écritures ont les mêmes clefs, mais ce n'est pas un problème
       car le compte bancaire est différent et surtout la nature donc la banque et l'agence sont
       différentes. Par contre, sur les comptes courants, il n'y a pas de distinction entre la
       date de valeur et la date d'opération}
      aCurTob.SetString('TE_CLEVALEUR', aCurTob.GetString('TE_CLEOPERATION'));
      {Mise à jour du RIB}
      Obj.SetLeRib(aCurTob);
      {Vérification de la validité de l'écriture saisie}
      if not TesteEcrTresorerie(aCurTob) then begin
        Result := False;
        RollBackDiscret;
      end;
    end;

var
  g, c, s : string;
  Mnt, MntD : Double;
  CptC : string;
  n : Integer;
begin
  Result := True;
  Mnt  := 0;
  MntD := 0;

  {Création de l'objet de gestion des contrepartie}
  Obj := TobjComptesTreso.Create(TobVir.GetString('TEQ_SNODOSSIER'), TobVir.GetString('TEQ_DNODOSSIER'));

  TobEqui := TOB.Create('£TRECRIT', nil, -1);
  try
    BeginTrans;
    try
      NumTra := Commun.GetNum(CodeModule, v_pgi.CodeSociete, TRANSACEQUI);

      aCurTob :=  TOB.Create('TRECRITURE', TobEqui, -1);
      aCurTob.SetString('TE_NODOSSIER', TobVir.GetString('TEQ_SNODOSSIER'));
      aCurTob.SetString('TE_SOCIETE', TobVir.GetString('TEQ_SOCIETE'));

      InitNlleEcritureTob(aCurTob, TobVir.GetString('TEQ_SGENERAL'));

      Mnt  := - Arrondi(TobVir.GetDouble('TEQ_MONTANT'), V_PGI.OkDecV);
      MntD := 0;
      ObjDev.ConvertitMnt(Mnt, MntD, V_PGI.DevisePivot, TobVir.GetString('TEQ_SGENERAL'));
      aCurTob.SetDouble('TE_MONTANT'   , Mnt);
      aCurTob.SetDouble('TE_MONTANTDEV', MntD);
      if Mnt <> 0 then aCurTob.SetDouble('TE_COTATION', Arrondi(MntD / Mnt, NBDECIMALTAUX))
                  else aCurTob.SetDouble('TE_COTATION', 1);

      GetCibSensGeneral(g, s, c, CODETRANSACDEP);
      aCurTob.SetString('TE_CODECIB', c);
      aCurTob.SetString('TE_CODEFLUX', CODETRANSACDEP);

      {17/03/05 : FQ 10223 : Refonte de la gestion des erreurs}
      if not Termine(TobVir.GetString('TEQ_SGENERAL'),
                     Obj.GetCouFromDoss(TobVir.GetString('TEQ_SNODOSSIER'), TobVir.GetString('TEQ_DNODOSSIER')),
                     TobVir.GetString('TEQ_SNODOSSIER')) then begin
        Result := False;
        RollBackDiscret;
        Exit;
      end;

      {On récupère la clé d'opération de l'écriture d'émission}
      SClef := aCurTob.GetString('TE_CLEOPERATION');

      aCurTob :=  TOB.Create('TRECRITURE', TobEqui, -1);
      aCurTob.SetString('TE_NODOSSIER', TobVir.GetString('TEQ_DNODOSSIER'));
      aCurTob.SetString('TE_SOCIETE', GetInfosFromDossier('DOS_NODOSSIER', TobVir.GetString('TEQ_DNODOSSIER'), 'DOS_SOCIETE'));
      InitNlleEcritureTob(aCurTob, TobVir.GetString('TEQ_DGENERAL'));

      Mnt  := Arrondi(TobVir.GetDouble('TEQ_MONTANT'), V_PGI.OkDecV);
      MntD := 0;
      ObjDev.ConvertitMnt(Mnt, MntD, V_PGI.DevisePivot, TobVir.GetString('TEQ_DGENERAL'));
      aCurTob.SetDouble('TE_MONTANT'   , Mnt);
      aCurTob.SetDouble('TE_MONTANTDEV', MntD);
      if Mnt <> 0 then aCurTob.SetDouble('TE_COTATION', Arrondi(MntD / Mnt, NBDECIMALTAUX))
                  else aCurTob.SetDouble('TE_COTATION', 1);

      GetCibSensGeneral(g, s, c, CODETRANSACREC);
      aCurTob.SetString('TE_CODECIB', c);
      aCurTob.SetString('TE_CODEFLUX', CODETRANSACREC);

      {17/03/05 : FQ 10223 : Refonte de la gestion des erreurs}
      if not Termine(TobVir.GetString('TEQ_DGENERAL'),
                     Obj.GetCouFromDoss(TobVir.GetString('TEQ_DNODOSSIER'), TobVir.GetString('TEQ_SNODOSSIER')),
                     TobVir.GetString('TEQ_DNODOSSIER')) then begin
        RollBackDiscret;
        Result := False;
        Exit;
      end;

      {On récupère la clé d'opération de l'écriture de destination}
      DClef := aCurTob.GetString('TE_CLEOPERATION');

      {Création de la contrepartie sur comptes courants}
      if (TobVir.GetString('TEQ_DNODOSSIER') <> TobVir.GetString('TEQ_SNODOSSIER'))
         and (TobEqui.Detail.Count = 2) then begin
        {PREMIERE CONTREPARTIE}
        aCurTob :=  TOB.Create('TRECRITURE', TobEqui, -1);
        {Duplication de la Tob sans ses filles (il n'y a pas de commissions ??)}
        aCurTob.Dupliquer(TobEqui.Detail[0], False, True);

        aCurTob.SetInteger('TE_NUMEROPIECE', TrGetNewNumPiece(TrGetSoucheDuJal(aCurTob.GetString('TE_JOURNAL')).RE,
                                             'TRE', aCurTob.GetString('TE_SOCIETE'), aCurTob.GetString('TE_JOURNAL'), True));

        aCurTob.SetString('TE_NUMTRANSAC', CODEMODULECOU + aCurTob.GetString('TE_SOCIETE') + TRANSACEQUI + NumTra);
        CptC := Obj.GetCouFromDoss(TobVir.GetString('TEQ_SNODOSSIER'), TobVir.GetString('TEQ_DNODOSSIER'));
        aCurTob.SetString('TE_GENERAL', CptC);
        {Le champ de contrepartie contenant un compte général et non un BQ_CODE}
        aCurTob.SetString('TE_CONTREPARTIETR', Obj.GetGeneFromCode(TobEqui.Detail[0].GetString('TE_GENERAL')));

        {Termine et contrôle l'ecriture}
        if not MajEcrCourantes then begin
          Result := False;
          Exit;
        end;

        {SECONDE CONTREPARTIE}
        aCurTob :=  TOB.Create('TRECRITURE', TobEqui, -1);
        {Duplication de la Tob sans ses filles (il n'y a pas de commissions ??)}
        aCurTob.Dupliquer(TobEqui.Detail[1], False, True);

        aCurTob.SetInteger('TE_NUMEROPIECE', TrGetNewNumPiece(TrGetSoucheDuJal(aCurTob.GetString('TE_JOURNAL')).RE,
                                             'TRE', aCurTob.GetString('TE_SOCIETE'), aCurTob.GetString('TE_JOURNAL'), True));

        aCurTob.SetString('TE_NUMTRANSAC', CODEMODULECOU + aCurTob.GetString('TE_SOCIETE') + TRANSACEQUI + NumTra);
        CptC := Obj.GetCouFromDoss(TobVir.GetString('TEQ_DNODOSSIER'), TobVir.GetString('TEQ_SNODOSSIER'));
        aCurTob.SetString('TE_GENERAL', CptC);
        {Le champ de contrepartie contient un compte général}
        aCurTob.SetString('TE_CONTREPARTIETR', Obj.GetGeneFromCode(TobEqui.Detail[1].GetString('TE_GENERAL')));

        {Termine et contrôle l'ecriture}
        if not MajEcrCourantes then begin
          Result := False;
          Exit;
        end;
      end;

      //TobDebug(TobEqui);
      TobEqui.InsertDb(nil);

      {Mise à jour du compteur de tréso}
      Commun.SetNum(CodeModule, v_pgi.CodeSociete, TRANSACEQUI, NumTra);

      {Recalcul des soldes}
      for n := 0 to TobEqui.Detail.Count -1 do begin
        aCurTob := TobEqui.Detail[n];
        RecalculSolde(aCurTob.GetString('TE_GENERAL'), aCurTob.GetString('TE_DATECOMPTABLE'), 0, True);
      end;

      if AvecMaj then begin
        {Mise à jour de la table équilibrage}
        SQL := 'UPDATE EQUILIBRAGE SET TEQ_CLESOURCE = "' + SClef + '", ' +
               'TEQ_CLEDESTINATION = "' + DClef + '", TEQ_VALIDBO = "X" ' +
               'WHERE TEQ_NUMEQUI = ' + TobVir.GetString('TEQ_NUMEQUI');
        ExecuteSQL(SQL);
      end;
      CommitTrans;
    except
      on E : Exception do begin
        RollBack;
        PgiError(TraduireMemoire('Erreur lors de la création des écritures avec le message :') + #13#13 + E.Message); 
        Result := False;
      end;
    end;
  finally
    if Assigned(TobEqui) then FreeAndNil(TobEqui);
    if Assigned(Obj    ) then FreeAndNil(Obj);
  end;
end;

{12/10/06 : Création des écritures d'ICC}
{---------------------------------------------------------------------------------------}
function  CreeTrEcritureIcc(TobIcc : TOB) : Boolean;
{---------------------------------------------------------------------------------------}
var
  TobMere : TOB;
  TobEcr  : TOB;
  NomBase : string;
  Gene    : string;
  Q       : TQuery;
  n       : Integer;

    {-------------------------------------------------------------------}
    function GenereEcritureIcc(CtrPartie : Boolean) : Boolean;
    {-------------------------------------------------------------------}
    var
      g, c    : string;
      TypeF   : string;
      Num     : string;
      CodeUnq : string;
      Sens    : Integer;
    begin
      Result := True;
      TobEcr.SetString('TE_GENERAL', Gene);
      Q := OpenSQL('SELECT BQ_GENERAL, BQ_NODOSSIER, DOS_SOCIETE, DOS_NOMBASE FROM BANQUECP ' +
                   'LEFT JOIN DOSSIER ON DOS_NODOSSIER = BQ_NODOSSIER ' +
                   'WHERE BQ_CODE = "' + TobEcr.GetString('TE_GENERAL') + '"', True);
      if not Q.EOF then begin
        TobEcr.SetString('TE_SOCIETE', Q.FindField('DOS_SOCIETE').AsString);
        TobEcr.SetString('TE_NODOSSIER', Q.FindField('BQ_NODOSSIER').AsString);
        NomBase := Q.FindField('DOS_NOMBASE').AsString;
        Gene    := Q.FindField('BQ_GENERAL').AsString;
      end
      else begin
        SetErreurTreso(NatErr_Gen);
        Result := False;
      end;
      Ferme(Q);

      if not Result then Exit;

      {Récupération du journal de substitution}
      TobEcr.SetString('TE_JOURNAL', GetParamSocDossierSecur('SO_CPJALENCADECA', '', NomBase));
      {Initialisation du record}
      InitNlleEcritureTob(TobEcr, TobEcr.GetString('TE_GENERAL'), TobEcr.GetString('TE_NODOSSIER'));

      if CtrPartie then TobEcr.SetString('TE_CODEFLUX', TobIcc.GetString('TE_CODEFLUX1'))
                   else TobEcr.SetString('TE_CODEFLUX', TobIcc.GetString('TE_CODEFLUX'));
      TobEcr.SetString('TE_NATURE',   na_Prevision);

      {Récupération du compte de contrepartie et du code cib dans la table FLUXTRESO}
      GetCibSensGeneral(g, TypeF, c, TobEcr.GetString('TE_CODEFLUX'));
      TobEcr.SetString('TE_CONTREPARTIETR', g);
      TobEcr.SetString('TE_CODECIB', CODECIBCOURANT);

      {Maj du sens}
      if GetSensFromTypeFlux(TypeF) = 'D' then Sens := -1
                                          else Sens := 1;

      {Récupération du RIB}
      GetRibTob(TobEcr);

      {27/05/05 : nouvelle gestion des erreurs}
      if TobEcr.GetString('TE_CODEBANQUE') = '' then SetErreurTreso(NatErr_Bqe);

      {Génération du numéro de transaction}
      Num := Commun.GetNum(CODEMODULE, TobEcr.GetString('TE_SOCIETE'), CODEICC);
      TobEcr.SetString('TE_NUMTRANSAC', CODEMODULE + TobEcr.GetString('TE_SOCIETE') + TRANSACICC + Num);

      TobEcr.SetString('TE_LIBELLE', TobIcc.GetString('TE_LIBELLE'));
      TobEcr.SetString('TE_USERCREATION', V_PGI.User);
      TobEcr.SetDateTime('TE_DATECREATION', V_PGI.DateEntree);
      {Les ICC sont calculés sur une période allant de TE_DATEVALID à TE_DATECOMPTABLE}
      TobEcr.SetDateTime('TE_DATECOMPTABLE', TobIcc.GetDateTime('TE_DATECOMPTABLE'));
      TobEcr.SetDateTime('TE_DATEVALID', TobIcc.GetDateTime('TE_DATEVALID'));

      TobEcr.SetString('TE_EXERCICE', TObjetExercice.GetCurrent.GetExoNodos(TobIcc.GetDateTime('TE_DATECOMPTABLE'),
                                                              TobEcr.GetString('TE_NODOSSIER')));
      TobEcr.SetString('TE_REFINTERNE', 'ICC ' + TobEcr.GetString('TE_GENERAL') + ' / Exercice : ' + TobEcr.GetString('TE_EXERCICE'));

      {Ecritures sur comptes courants : la devise est la devise pivot}
      TobEcr.SetString('TE_DEVISE', V_PGI.DevisePivot);
      TobEcr.SetDouble('TE_COTATION', 1);
      TobEcr.SetDouble('TE_MONTANTDEV', Arrondi(TobIcc.GetDouble('TE_MONTANT') * Sens, V_PGI.OkDecV));
      TobEcr.SetDouble('TE_MONTANT'   , Arrondi(TobIcc.GetDouble('TE_MONTANT') * Sens, V_PGI.OkDecV));

      if (TobEcr.GetString('TE_CODECIB') > '') and (TobEcr.GetString('TE_CODEFLUX') > '') then begin
        {Dans le cas d'une écriture de simulation, les dates de valeur et comptable sont équivalentes}
        TobEcr.SetDateTime('TE_DATEVALEUR', TobEcr.GetDateTime('TE_DATECOMPTABLE'));
        {On considère l'écriture comme pointée}
        TobEcr.SetDateTime('TE_DATERAPPRO', TobEcr.GetDateTime('TE_DATECOMPTABLE'));

        {... Génération des clés}
        CodeUnq := Commun.GetNum(CODEUNIQUE, CODEUNIQUE, CODEUNIQUE);
        TobEcr.SetString('TE_CLEVALEUR',    RetourneCleEcriture(TobEcr.GetDateTime('TE_DATEVALEUR'), StrToInt(CodeUnq)));
        TobEcr.SetString('TE_CLEOPERATION', RetourneCleEcriture(TobEcr.GetDateTime('TE_DATECOMPTABLE'), StrToInt(CodeUnq)));
        Commun.SetNum(CODEUNIQUE, CODEUNIQUE, CODEUNIQUE, CodeUnq);
        {20/11/06 : Avertissement sur les dates de réinitialisation}
        if TestDateEtMillesime(TobEcr.GetDateTime('TE_DATEVALEUR'), Date) or
           TestDateEtMillesime(TobEcr.GetDateTime('TE_DATECOMPTABLE'), Date) then begin
          if HShowmessage('0;Intérêts comptes courants;ATTENTION !'#13'La nouveau flux n''appartient pas au millésime courant.'#13 +
             'Le montant saisi peut avoir un impact sur les montants de réinitialisation.'#13#13 +
             'Souhaitez-vous poursuivre ?;W;YN;N;N', '', '') = mrNo then begin
            Result := False;
            Exit;
          end;
        end;

        if not TesteEcrTresorerie(TobEcr) then begin
          {17/03/05 : FQ 10223 : Refonte de la gestion des erreurs}
          AfficheMessageErreur('Intérêts comptes courants', 'L''écriture ou sa contre-partie n''a pu être créée');
          Result := False;
        end;
        {... mise à jour du compteur}
        Commun.SetNum(CODEMODULE, TobEcr.GetString('TE_SOCIETE'), CODEICC, Num)
      end
      else
        Result := False
    end;

begin
  Result := True;
  Gene := TobIcc.GetString('TE_GENERAL');

  TobMere := TOB.Create('AATRECRITURE', nil, -1);
  try
    {1/ Création de l'écriture saisie}
    TobEcr  := TOB.Create('TRECRITURE', TobMere, -1);
    if not GenereEcritureIcc(False) then begin
      Result := False;
      Exit;
    end;

    {2/ Création de l'écriture de contrepartie}
    {Recherche du dossier de contre partie}
    Q := OpenSQL('SELECT CLS_DOSSIER FROM ' + GetTableDossier(NomBase, 'CLIENSSOC') + ' WHERE CLS_GENERAL = "' + Gene+ '"', True);
    if not Q.EOF  then Gene := Q.FindField('CLS_DOSSIER').AsString
                  else Result := False;
    Ferme(Q);
    if not Result then Exit;

    {Recherche de la contrepartie "générale" dans le dossier de contre partie}
    Q := OpenSQL('SELECT CLS_GENERAL FROM ' + GetTableDossier(Gene, 'CLIENSSOC') + ' WHERE CLS_DOSSIER = "' + NomBase + '"', True);
    {NomBase contenait le nom de base du flux d'origine, et gene celui du flux de Contrepartie : on le mémorise}
    NomBase := Gene;
    if not Q.EOF  then Gene := Q.FindField('CLS_GENERAL').AsString
                  else Result := False;
    Ferme(Q);
    if not Result then Exit;

    {Recherche de la contrepartie Bancaire}
    Q := OpenSQL('SELECT BQ_CODE FROM BANQUECP LEFT JOIN DOSSIER ON DOS_NODOSSIER = BQ_NODOSSIER ' +
                 'WHERE BQ_GENERAL = "' + Gene + '" AND DOS_NOMBASE = "' + NomBase + '"', True);
    if not Q.EOF  then Gene := Q.FindField('BQ_CODE').AsString
                  else Result := False;
    Ferme(Q);
    if not Result then Exit;

    {Génération de la contrepartie}
    TobEcr  := TOB.Create('TRECRITURE', TobMere, -1);
    if not GenereEcritureIcc(True) then begin
      Result := False;
      Exit;
    end;

    Result := False;
    TobMere.InsertDb(nil);
    Result := True;

    {... recalcul des soldes}
    for n := 0 to TobMere.Detail.Count - 1 do
      RecalculSolde(TobMere.Detail[n].GetString('TE_GENERAL'), TobMere.Detail[n].GetString('TE_DATECOMPTABLE'), 0, True);
  finally
    FreeAndNil(TobMere);
  end;
end;

{06/12/04 : Initialisation d'une nouvelle écriture avec une Tob : Il est temps que je remplace
            progressivement le Record TrEcriture par une Tob réèlle}
{---------------------------------------------------------------------------------------}
procedure InitNlleEcritureTob(var TobEcriture : TOB; Gene : string; Dossier : string = '');
{---------------------------------------------------------------------------------------}
var
  Res  : TDblResult;
  Base : string;
  Jal  : string;
begin
  if Dossier = '' then Dossier := V_PGI.NoDossier
                  else Base := GetInfosFromDossier('DOS_NODOSSIER', Dossier, 'DOS_NOMBASE');
  {12/10/06 : Pour les Intérêts de C / C, le journal est stocké dans un ParamSoc}
  Jal := TobEcriture.GetString('TE_JOURNAL');

  if Jal <> '' then
    Res := TrGetSoucheDuJal(Jal, '')
  else
    Res := TrGetSoucheDuJal('', Gene);

  if Res.RT = '' then begin
    {17/03/05 : FQ 10223 : Refonte de la gestion des erreurs}
    if V_PGI.SAV then
      HShowMessage('0;Initialisation d''une nouvelle écriture;Le compte bancaire "' + Gene + '" n''est rattaché à aucun journal.;W;O;O;O;', '', '');

    ErreurCategorie.TypeErreur := ErreurCategorie.TypeErreur + [CatErr_TRE];
    ErreurCategorie.Ecriture   := ErreurCategorie.Ecriture  + [NatErr_Jal];
  end;

  TobEcriture.SetInteger('TE_NUMEROPIECE'  , TrGetNewNumPiece(Res.RE, CODEMODULE, V_PGI.CodeSociete, Res.RT, True));
  TobEcriture.SetInteger('TE_NUMLIGNE' 	   , 1);
  TobEcriture.SetInteger('TE_CPNUMLIGNE'   , 1);
  TobEcriture.SetInteger('TE_NUMECHE' 	    , 1);
  TobEcriture.SetString('TE_JOURNAL'       , Res.RT);
  TobEcriture.SetString('TE_EXERCICE'      , TObjetExercice.GetCurrent.GetExoNodos(V_PGI.DateEntree, Dossier));
  TobEcriture.SetString('TE_GENERAL' 	     , Gene);
  TobEcriture.SetString('TE_CODECIB' 	     , '');
  TobEcriture.SetString('TE_CODEFLUX' 	    , '');
  TobEcriture.SetString('TE_REFINTERNE'    , '');

  if TobEcriture.GetString('TE_SOCIETE') = '' then
    TobEcriture.SetString('TE_SOCIETE' 	   , V_PGI.CodeSociete);
  if TobEcriture.GetString('TE_NODOSSIER') = '' then
    TobEcriture.SetString('TE_NODOSSIER'   , Dossier);

  TobEcriture.SetString('TE_QUALIFORIGINE' , QUALIFTRESO);
  TobEcriture.SetString('TE_NUMTRANSAC'	   , '');
  TobEcriture.SetString('TE_LIBELLE'	   , '');                    
  TobEcriture.SetString('TE_USERCREATION'  , V_PGI.User);
  TobEcriture.SetString('TE_USERMODIF'	   , '');
  TobEcriture.SetString('TE_USERVALID'	   , '');
  TobEcriture.SetString('TE_USERCOMPTABLE' , '');
  TobEcriture.SetString('TE_DEVISE'	   , '');
  TobEcriture.SetString('TE_CODEBANQUE'    , '');
  TobEcriture.SetString('TE_CODEGUICHET'   , '');
  TobEcriture.SetString('TE_NUMCOMPTE' 	   , '');
  TobEcriture.SetString('TE_CLERIB' 	   , '');       
  TobEcriture.SetString('TE_ETABLISSEMENT' , GetParamsocDossierSecur('SO_ETABLISDEFAUT','', Base));
  {Par défaut, les écritures sont prévisionnelles}
  TobEcriture.SetString('TE_NATURE'        , na_Prevision);
  {Par défaut, une écriture est considérée sans commission}
  TobEcriture.SetDouble('TE_MONTANTDEV'	   , 0);
  TobEcriture.SetDouble('TE_MONTANT'	   , 0);
  {La cotation est mise à zéro, même si dans tous les traitements la valeur des par défaut est 1 :
   Cela permet ainsi de voir s'il y a une anomalie et de voir à qu'elle niveau elle se situe}
  TobEcriture.SetDouble('TE_COTATION'	   , 0);
  TobEcriture.SetDouble('TE_SOLDEDEVVALEUR', 0);
  TobEcriture.SetDouble('TE_SOLDEDEV' 	   , 0);
  TobEcriture.SetString('TE_COMMISSION'    , suc_SansCom);
  TobEcriture.SetDateTime('TE_DATEVALEUR'  , iDate1900);
  TobEcriture.SetDateTime('TE_DATECREATION', Date);
  TobEcriture.SetDateTime('TE_DATEMODIF'   , iDate1900);
  TobEcriture.SetDateTime('TE_DATEVALID'   , iDate1900);
  TobEcriture.SetDateTime('TE_DATERAPPRO'  , iDate1900);
  TobEcriture.SetDateTime('TE_DATECOMPTABLE', iDate1900);
end;

{Termine l'écriture contenue dans T et l'écrit dans la base et génère les commissions
{---------------------------------------------------------------------------------------}
function TermineEcritureTob(var T : TOB; ObjDev : TObjDevise; ComplOk : Boolean) : Boolean;
{---------------------------------------------------------------------------------------}
var
  CodeUnq : string;
  ObjComm : TObjCommissionTob;
begin
  Result := True;
  try

    if ComplOk then begin
      {11/08/06 : Remplit les infos multi sociétés}
      InitInfosDossier(T);

      {Gestion des dates}
      T.SetString('TE_EXERCICE', TObjetExercice.GetCurrent.GetExoNodos(T.GetDateTime('TE_DATECOMPTABLE'),
                                                         T.GetString('TE_NODOSSIER')));

      {Gestion de l'utilisateur / date de validation : la gestion des la création est faite dans
       InitNlleEcritureTob et on ne renseigne pas le modificateur}
      T.SetString('TE_USERVALID', V_PGI.User);
      T.SetDateTime('TE_DATEVALID', Date);

      {Gestion de la devise : La devise du compte n'est pas nécessairement celle du compte !!!}
      T.SetString('TE_DEVISE', ObjDev.GetDeviseCpt(T.GetString('TE_GENERAL')));
      T.SetDouble('TE_MONTANTDEV', ObjDev.GetMntDevFromEur(T.GetDouble('TE_MONTANT'), T.GetString('TE_GENERAL'), T.GetDateTime('TE_DATECOMPTABLE')));
      T.SetDouble('TE_COTATION', ObjDev.Get_Cotation(T.GetString('TE_GENERAL'), T.GetDateTime('TE_DATECOMPTABLE')));

      T.SetDateTime('TE_DATEVALEUR', CalcDateValeur(T.GetString('TE_CODECIB'), T.GetString('TE_GENERAL'), T.GetDateTime('TE_DATECOMPTABLE')));
    end
    else
      {La fonction n'est jamais appelée avec ComplOk à False, mais ....!}
      T.SetString('TE_EXERCICE', TObjetExercice.GetCurrent.GetExoNodos(T.GetDateTime('TE_DATECOMPTABLE'),
                                                         T.GetString('TE_NODOSSIER')));

    {Récupération des informations bancaires}
    GetRIBTob(T);

    {23/03/05 : FQ 10223 : Nouvelle gestion des erreurs sur les écritures}
    if not TesteEcrTresorerie(T) then begin
      Result := False;
      Exit;
    end;

    {Génération des clés}
    CodeUnq := Commun.GetNum(CODEUNIQUE, CODEUNIQUE, CODEUNIQUE);
    T.SetString('TE_CLEVALEUR'   , RetourneCleEcriture(T.GetDateTime('TE_DATEVALEUR'), StrToInt(CodeUnq)));
    T.SetString('TE_CLEOPERATION', RetourneCleEcriture(T.GetDateTime('TE_DATECOMPTABLE'), StrToInt(CodeUnq)));
    Commun.SetNum(CODEUNIQUE, CODEUNIQUE, CODEUNIQUE, CodeUnq);

    {création des Frais et Commissions}
    ObjComm := TObjCommissionTob.Create(T.GetString('TE_GENERAL'), T.GetString('TE_CODEFLUX'), T.GetDateTime('TE_DATECOMPTABLE'));
    try
      ObjComm.GenererCommissions(T);

      {Ecriture de l'enregistrement dans la table}
      ObjComm.TobEcriture.InsertDB(nil);
      T.InsertDB(nil);
    finally
      if Assigned(ObjComm) then FreeAndNil(ObjComm);
    end;
  except
    raise;
  end;
end;
      
{07/12/04 : Remplit les champs standards d'une ligne ECRITURE à partir d'une ligne TRECRITURE}
{---------------------------------------------------------------------------------------}
procedure RemplitChpStdTob(var TobCompta : TOB; TobEcr : TOB);
{---------------------------------------------------------------------------------------}
begin
  {Standard compta}
  CPutDefautEcr(TobCompta);
  CSupprimerInfoLettrage(TobCompta);
  TobCompta.SetString('E_QUALIFORIGINE', QUALIFTRESO);
  TobCompta.SetDateTime('E_DATECREATION', V_PGI.DateEntree);
  TobCompta.SetString('E_AUXILIAIRE','');
  TobCompta.SetString('E_VALIDE','X');
  TobCompta.SetDateTime('E_DATEREFEXTERNE', iDate1900);
  TobCompta.SetDateTime('E_DATERELANCE', iDate1900);
  TobCompta.SetDateTime('E_DATEPOINTAGE', iDate1900);
  TobCompta.SetDateTime('E_DATERELANCE', iDate1900);
  TobCompta.SetString('E_IO'          , 'X');
  TobCompta.SetString('E_ETATREVISION', '-');
  TobCompta.SetDateTime('E_DATETAUXDEV' , V_PGI.DateEntree);
  TobCompta.SetString('E_TRESOSYNCHRO', 'SYN');
  TobCompta.SetString('E_CREERPAR'    , 'IMP');
  TobCompta.SetString('E_ETATREVISION', '-');
  TobCompta.SetString('E_RIB', '');

  TobCompta.SetDateTime('E_DATECOMPTABLE', TobEcr.GetDateTime('TE_DATECOMPTABLE'));
  TobCompta.SetString('E_EXERCICE', TobEcr.GetString('TE_EXERCICE'));
  TobCompta.SetInteger('E_SEMAINE', NumSemaine(TobEcr.GetDateTime('TE_DATECOMPTABLE')));
  TobCompta.SetInteger('E_PERIODE', GetPeriode(TobEcr.GetDateTime('TE_DATECOMPTABLE')));
  TobCompta.SetDateTime('E_DATEVALEUR', TobEcr.GetDateTime('TE_DATEVALEUR'));
  TobCompta.SetString('E_JOURNAL', TobEcr.GetString('TE_JOURNAL'));
  TobCompta.SetString('E_REFINTERNE', TobEcr.GetString('TE_NUMTRANSAC') + '-' +
                                      TobEcr.GetString('TE_NUMEROPIECE') + '-' +
                                      TobEcr.GetString('TE_NUMLIGNE'));
  TobCompta.SetInteger('E_NUMEROPIECE', TobEcr.GetInteger('TE_NUMEROPIECE'));
  TobCompta.SetString('E_SOCIETE', TobEcr.GetString('TE_SOCIETE'));
  TobCompta.SetString('E_ETABLISSEMENT', TobEcr.GetString('TE_ETABLISSEMENT'));
  TobCompta.SetString('E_LIBELLE', TobEcr.GetString('TE_LIBELLE'));
  TobCompta.SetString('E_DEVISE', TobEcr.GetString('TE_DEVISE'));
  TobCompta.SetDouble('E_COTATION', TobEcr.GetDouble('TE_COTATION'));
  if TobEcr.GetDouble('TE_COTATION') <> 0 then TobCompta.SetDouble('E_TAUXDEV', Arrondi(1 / TobEcr.GetDouble('TE_COTATION'),NBDECIMALTAUX))
                                          else TobCompta.SetDouble('E_TAUXDEV', 1);
  TobCompta.SetDateTime('E_ORIGINEPAIEMENT', iDate1900);
end;

{Retourne le numéro de ligne pour l'écriture comptable à générer
{---------------------------------------------------------------------------------------}
function GetNumLigne(NumTransac, NumPiece : string) : Integer;
{---------------------------------------------------------------------------------------}
begin
  Result := 1;
  if IsNouvellePiece(NumTransac, NumPiece) then
    ReInitRecordCompta(NumTransac, NumPiece)
  else begin
    Result := RecordCompta.NumLigne;
    Inc(RecordCompta.NumLigne);
  end;
end;

{Traite-t-on, en intégration en comptabilité, une nouvelle pièce ?
{---------------------------------------------------------------------------------------}
function IsNouvellePiece(NumTransac, NumPiece : string) : Boolean;
{---------------------------------------------------------------------------------------}
begin
  Result := (RecordCompta.NumTransac <> NumTransac) or (RecordCompta.NumPiece <> NumPiece);
end;

{Réinitialise la variable globale RecordCompta
{---------------------------------------------------------------------------------------}
procedure ReInitRecordCompta(Transac, Piece : string);
{---------------------------------------------------------------------------------------}
begin
  with RecordCompta do begin
    NumTransac := Transac;
    NumPiece   := Piece;
    NumLigne   := 2;
  end;
end;

{17/03/05 : FQ 10223 : Teste si les écritures de trésorerie sont valides
{---------------------------------------------------------------------------------------}
function TesteEcrTresorerie(T : TOB) : Boolean;
{---------------------------------------------------------------------------------------}
begin
  Result := True;
  if (T.GetString('TE_CODEFLUX') = '') then begin
    ErreurCategorie.TypeErreur := ErreurCategorie.TypeErreur + [CatErr_TRE];
    ErreurCategorie.TrEcriture := ErreurCategorie.TrEcriture + [NatErr_Flx];
    Result := False;
  end;
  if (T.GetString('TE_CODECIB') = '') then begin
    ErreurCategorie.TypeErreur := ErreurCategorie.TypeErreur + [CatErr_TRE];
    ErreurCategorie.TrEcriture := ErreurCategorie.TrEcriture + [NatErr_Cib];
    Result := False;
  end;
  if (T.GetString('TE_CODEBANQUE') = '') then begin
    ErreurCategorie.TypeErreur := ErreurCategorie.TypeErreur + [CatErr_TRE];
    ErreurCategorie.TrEcriture := ErreurCategorie.TrEcriture + [NatErr_Bqe];
    Result := False;
  end;
  if (T.GetString('TE_GENERAL') = '') then begin
    ErreurCategorie.TypeErreur := ErreurCategorie.TypeErreur + [CatErr_TRE];
    ErreurCategorie.TrEcriture := ErreurCategorie.TrEcriture + [NatErr_Gen];
    Result := False;
  end;
  if (T.GetString('TE_CONTREPARTIETR') = '') then begin
    ErreurCategorie.TypeErreur := ErreurCategorie.TypeErreur + [CatErr_TRE];
    ErreurCategorie.TrEcriture := ErreurCategorie.TrEcriture + [NatErr_CGn];
    Result := False;
  end;
  if (T.GetString('TE_EXERCICE') = '') then begin
    ErreurCategorie.TypeErreur := ErreurCategorie.TypeErreur + [CatErr_TRE];
    ErreurCategorie.TrEcriture := ErreurCategorie.TrEcriture + [NatErr_Exo];
    Result := False;
  end;
  if (T.GetString('TE_JOURNAL') = '') then begin
    ErreurCategorie.TypeErreur := ErreurCategorie.TypeErreur + [CatErr_TRE];
    ErreurCategorie.TrEcriture := ErreurCategorie.TrEcriture + [NatErr_Jal];
    Result := False;
  end;
end;

{17/03/05 : FQ 10223 : Génération du message d'erreur lors de la créatio de TrEcritures ou d'Ecritures
                       Renvoie True s'il y a des messages à afficher
{---------------------------------------------------------------------------------------}
function AfficheMessageErreur(Titre : string = ''; Entete : string = '') : Boolean;
{---------------------------------------------------------------------------------------}
var
  Msg : string;
begin
  {Il n'y a pas d'erreur, on sort}
  if ErreurCategorie.TypeErreur = [] then Result := False
                                     else Result := True;

  if not Result then Exit;
  if Entete <> '' then Msg := Entete
                  else Msg := 'Certaines écritures n''ont pu être créés :';

  {Gestion des erreurs lors de l'intégration en Comptabilité}
  if CatErr_CPT in ErreurCategorie.TypeErreur then begin
    if NatErr_Cpt in ErreurCategorie.Ecriture then
      Msg := Msg + #13' - Certaines pièces générées n''etaient pas correctes pour la comptabilité :'
                 + #13'     ' +  ErreurCategorie.MsgCompta;
    if NatErr_Tva in ErreurCategorie.Ecriture then
      Msg := Msg + #13' - Compte général absent dans le paramétrage de la TVA';
    if NatErr_Bqe in ErreurCategorie.Ecriture then
      Msg := Msg + #13' - Le code banque de certains établissements bancaires n''est pas renseigné.';
    if NatErr_Jal in ErreurCategorie.Ecriture then
      Msg := Msg + #13' - Certains comptes ne sont pas en contrepartie d''un journal bancaire.';
    if NatErr_Exo in ErreurCategorie.Ecriture then
      Msg := Msg + #13' - Certaines écritures ne correspondent à aucun exercice ouvert en comptabilité.';
    if NatErr_Gen in ErreurCategorie.Ecriture then
      Msg := Msg + #13' - Impossible de récupérer le compte général de certaines écritures.'#13 +
                      '   Veuillez notamment vérifier votre paramétrage de la TVA.';
    if NatErr_CIB in ErreurCategorie.Ecriture then
      Msg := Msg + #13' - Impossible de récupérer le mode de règlement de certaines écritures.'#13 +
                      '   Veuillez vérifier le paramétrage vos codes InterBancaires.';
  end;

  {Gestion des erreurs lors de la création d'écriture de trésorerie}
  if CatErr_TRE in ErreurCategorie.TypeErreur then begin
    if NatErr_Int in ErreurCategorie.TrEcriture then
      Msg := Msg + #13' - Certaines opérations ont été intégrées en comptabilité.'#13 +
                      '   Il n''est pas possible de supprimer les écritures concernées.';
    if NatErr_VBO in ErreurCategorie.TrEcriture then
      Msg := Msg + #13' - Certaines opérations n''ont pas été validées BO.'#13 +
                      '   Il n''est pas possible de les traiter.';
    if NatErr_Bqe in ErreurCategorie.TrEcriture then
      Msg := Msg + #13' - Le code banque de certains établissements bancaires n''est pas renseigné.';
    if NatErr_Jal in ErreurCategorie.TrEcriture then
      Msg := Msg + #13' - Certains comptes ne sont pas en contrepartie d''un journal bancaire.';
    if NatErr_Exo in ErreurCategorie.TrEcriture then
      Msg := Msg + #13' - Certaines écritures ne correspondent à aucun exercice ouvert en comptabilité.';
    if NatErr_Flx in ErreurCategorie.TrEcriture then
      Msg := Msg + #13' - Impossible de recupérer les codes flux de certaines écritures.';
    if NatErr_CGn in ErreurCategorie.TrEcriture then
      Msg := Msg + #13' - Impossible de récupérer le compte de contrepartie de certaines écritures.'#13 +
                      '   Veuillez notamment vérifier votre paramétrage de vos flux.';
    if NatErr_CIB in ErreurCategorie.TrEcriture then
      Msg := Msg + #13' - Impossible de récupérer le code CIB de certaines écritures.'#13 +
                      '   Veuillez vérifier le paramétrage vos codes Flux.';
    if NatErr_Gen in ErreurCategorie.TrEcriture then
      Msg := Msg + #13' - Impossible de récupérer certains comptes bancaires.';
  end;

  {Gestion des erreurs lors de la création d'écriture de trésorerie}
  if CatErr_COM in ErreurCategorie.TypeErreur then begin
    if NatErr_Jal in ErreurCategorie.Commission then
      Msg := Msg + #13' - Certains comptes ne sont pas en contrepartie d''un journal bancaire.';
    if NatErr_CoE in ErreurCategorie.Commission then
      Msg := Msg + #13' - Erreur lors de l''écriture de commissions dans la base.';
    if NatErr_Exo in ErreurCategorie.Commission then
      Msg := Msg + #13' - Certaines commissions ne correspondent à aucun exercice ouvert en comptabilité.';
    if NatErr_Flx in ErreurCategorie.Commission then
      Msg := Msg + #13' - Impossible de recupérer les codes flux de certaines commissions.';
    if NatErr_CGn in ErreurCategorie.Commission then
      Msg := Msg + #13' - Impossible de récupérer le compte de contrepartie de certaines commissions.'#13 +
                      '   Veuillez notamment vérifier votre paramétrage de vos flux de commissions.';
    if NatErr_CIB in ErreurCategorie.Commission then
      Msg := Msg + #13' - Impossible de récupérer le code CIB de certaines commissions.'#13 +
                      '   Veuillez vérifier le paramétrage vos Flux de commission.';
    if NatErr_Gen in ErreurCategorie.Commission then
      Msg := Msg + #13' - Impossible de récupérer certains comptes bancaires.';
  end;

  if Titre <> '' then
    HShowMessage('1;' + Titre + ';' + Msg + ';W;O;O;O;', '', '')

  else
    PGIError(Msg);

  {Le message ayant été affiché, on réinitialise les ensembles}
  InitGestionErreur;
end;

{21/03/05 : FQ 10223 : Réinitialise les ensembles d'erreur
{---------------------------------------------------------------------------------------}
procedure InitGestionErreur(CatCur : TCatErreur = CatErr_None);
{---------------------------------------------------------------------------------------}
begin
  CategorieCurrent := CatCur;
  with ErreurCategorie do begin
    TypeErreur := [];
    Commission := [];
    Ecriture   := [];
    TrEcriture := [];
    MsgCompta  := '';
  end;
end;

{---------------------------------------------------------------------------------------}
procedure SetErreurTreso(Erreur : TNatErreur);
{---------------------------------------------------------------------------------------}
begin
       if CategorieCurrent = CatErr_CPT then ErreurCategorie.Ecriture   := ErreurCategorie.Ecriture   + [Erreur]
  else if CategorieCurrent = CatErr_TRE then ErreurCategorie.TrEcriture := ErreurCategorie.TrEcriture + [Erreur]
  else if CategorieCurrent = CatErr_COM then ErreurCategorie.Commission := ErreurCategorie.Commission + [Erreur];
  if CategorieCurrent <> CatErr_None then
    ErreurCategorie.TypeErreur := ErreurCategorie.TypeErreur + [CategorieCurrent];
end;

{26/07/05 : FQ 10158 : Mise à jour d'une écriture de compta à partir d'une écriture de Treso
            Si Chp est un VARCHAR, valeur doit être de la forme "VALUE"
 10/08/06 : Gestion du multi sociétés
{---------------------------------------------------------------------------------------}
procedure MajEcritureParClef(Chp, Valeur : string; Clef : TClefCompta);
{---------------------------------------------------------------------------------------}
var
  Tab : string;
  SQl : string;
begin
  if (Clef.NoD = '') or (Clef.NoD = V_PGI.NoDossier) then
    Tab := 'ECRITURE'
  else
    Tab := GetTableDossier(GetInfosFromDossier('DOS_NODOSSIER', Clef.NoD, 'DOS_NOMBASE'), 'ECRITURE');

  SQL := 'UPDATE ' + Tab + ' SET ' + Chp + ' = ' + Valeur + ' WHERE ' +
         'E_EXERCICE = "' + Clef.Exo + '" AND ' +
         'E_JOURNAL = "'  + Clef.Jal + '" AND ' +
         {26/08/05 : On ne prend plus la date comptable dans la clef, mais cela peut
                     soulever quelque petit problème si on modifie les souches !!!
         'E_DATECOMPTABLE = "' + UsDateTime(Clef.dtC) + '" AND ' +}
         'E_NUMEROPIECE = ' + IntToStr(Clef.Pce) + ' AND ' +
         'E_NUMLIGNE = '    + IntToStr(Clef.Lig) + ' AND ' +
         'E_NUMECHE = '     + IntToStr(Clef.Ech) + ' AND ';
  if Clef.Per > -1 then {17/11/06 : Gestion des bordereaux}
    SQL := SQL + 'E_PERIODE = '   + IntToStr(Clef.Per) + ' AND ';
  SQL := SQL + 'E_QUALIFPIECE = "' + Clef.Qlf + '"';
  ExecuteSQL(SQL);
end;

{20/09/06 : l'écriture est-elle intégrable en comptabilité ?
{---------------------------------------------------------------------------------------}
function IsIntegrable(ligneTEcr : TOB) : Boolean;
{---------------------------------------------------------------------------------------}
begin
  Result := False;
  {Déjà intégrée}
  if ligneTEcr.GetString('TE_USERCOMPTABLE') <> '' then Exit;
  {Ecriture courante non intégrable}
  if Copy(ligneTEcr.GetString('TE_NUMTRANSAC'), 1, 3) = CODEMODULECOU then Exit;

  if ligneTEcr.GetString('TE_QUALIFORIGINE') <> QUALIFTRESO then Exit;
  Result := True;
end;

{29/03/07 : Nouvelle fonction de création des écritures de Tréso CreerTrEcriture
{---------------------------------------------------------------------------------------}
function  CreerTrEcriture(Cpte, Flux, Libelle, Nature  : string; MntDev : Double;DateC : TDateTime;
                          DateV : TDateTime = 0; Cib : string = ''; Ref : string = ''; Dev : string = '') : TOB;
{---------------------------------------------------------------------------------------}
var
  Num      : string;
  TypeFlux : string;
  Sens     : Shortint;
  CodeUnq  : string;
  Devise   : string;
  ObjCom   : TObjCommissionTob;
  ObjDev   : TObjDetailDevise;
  Q        : TQuery;
  g, c     : string;
begin
  Devise := Dev;
  InitGestionErreur;
  {27/05/05 : on précise que l'on travaille sur une éciture de trésorerie}
  CategorieCurrent := CatErr_TRE;

  Result := TOB.Create('TRECRITURE', nil, -1);
  Result.SetString('TE_GENERAL', Cpte);

  if IsTresoMultiSoc then begin
    Q := OpenSQL('SELECT BQ_CODE, BQ_DEVISE, BQ_NODOSSIER, DOS_SOCIETE FROM BANQUECP ' +
                 'LEFT JOIN DOSSIER ON DOS_NODOSSIER = BQ_NODOSSIER ' +
                 'WHERE BQ_CODE = "' + Cpte + '"', True);
    if not Q.EOF then begin
      Devise := Q.FindField('BQ_DEVISE').AsString;
      Result.SetString('TE_SOCIETE', Q.FindField('DOS_SOCIETE').AsString);
      Result.SetString('TE_NODOSSIER', Q.FindField('BQ_NODOSSIER').AsString);
    end
    else
      SetErreurTreso(NatErr_Gen);
    Ferme(Q);
  end
  else begin
    if Devise = '' then Devise := RetDeviseCompte(Cpte);
    Result.SetString('TE_SOCIETE', V_PGI.CodeSociete);
    Result.SetString('TE_NODOSSIER', V_PGI.NoDossier);
  end;

  {Initialisation du record}
  InitNlleEcritureTob(Result, Cpte, Result.GetString('TE_NODOSSIER'));

  Result.SetString('TE_CODEFLUX', Flux);
  Result.SetString('TE_NATURE',   Nature);

  {S'il s'agit d'une écriture prévisionnelle, donc gérée avec un code flux ...}
  if StrToChr(Nature) in [na_Prevision, na_Realise] then begin
    {Récupération du compte de contrepartie et du code cib dans la table FLUXTRESO}
    GetCibSensGeneral(g, TypeFlux, c, Flux);
    Result.SetString('TE_CONTREPARTIETR', g);
    {on force éventuellement le CIB fourni}
    if Cib <> '' then Result.SetString('TE_CODECIB', Cib)
                 else Result.SetString('TE_CODECIB', c);

    {Maj du sens}
    if GetSensFromTypeFlux(TypeFlux) = 'D' then Sens := -1
                                           else Sens := 1;
  end

  {... sinon, il s'agit d'une écriture de simulation (S) gérée avec un code rubrique}
  else begin
    {On récupère le sens à partir des rubriques}
    if GetSensFromRubrique(Flux) = 'D' then Sens := -1
                                       else Sens := 1;
    {... En attendant la gestion des cib !!}
    Result.SetString('TE_CODECIB', CODECIBSIMUL);
    {28/12/05 : FQ 10321 créée sur le sujet.
     21/12/05 : Maintenant avec la nouvel gestion des erreurs, la contrepartie est obligatoire,
                en attendant quelque chose de mieux, car il n'est pas facile de trouver la
                contrepartie dans la fiche des rubriques}
    Result.SetString('TE_CONTREPARTIETR', 'SIMUL');
  end;

  {Récupération du RIB}
  GetRibTob(Result);

  {27/05/05 : nouvelle gestion des erreurs}
  if Result.GetString('TE_CODEBANQUE') = '' then SetErreurTreso(NatErr_Bqe);

  {Génération du numéro de transaction}
  Num := Commun.GetNum(CODEMODULE, Result.GetString('TE_SOCIETE'), CODESAISIE);
  Result.SetString('TE_NUMTRANSAC', CODEMODULE + Result.GetString('TE_SOCIETE') + TRANSACSAISIE + Num);

  Result.SetString('TE_LIBELLE', Libelle);
  Result.SetString('TE_USERCREATION', V_PGI.User);
  Result.SetDateTime('TE_DATECREATION', V_PGI.DateEntree);
  Result.SetDateTime('TE_DATECOMPTABLE', DateC);
  
  if Ref = '' then Result.SetString('TE_REFINTERNE', 'Saisie le ' + DateToStr(DateC) + ' par ' + V_PGI.User)
              else Result.SetString('TE_REFINTERNE', Ref);

  Result.SetString('TE_EXERCICE', TObjetExercice.GetCurrent.GetExoNodos(DateC, Result.GetString('TE_NODOSSIER')));

  {Conversion du montant en Euros}
  Result.SetString('TE_DEVISE', Devise);

  {24/05/05 : Gestion du format des montants}
  ObjDev := TObjDetailDevise.Create;
  try
    {$IFNDEF EAGLSERVER}
    TheData := ObjDev;
    {$ENDIF EAGLSERVER}

    Result.SetDouble('TE_COTATION', RetPariteEuro(Result.GetString('TE_DEVISE'), Result.GetDateTime('TE_DATECREATION')));
    Result.SetDouble('TE_MONTANTDEV', Arrondi(MntDev * Sens, ObjDev.NbDecimal));
    {RetPariteEuro renvoie le taux (1 Dev = x.xx €) => on transforme en cotation (1€ = x.xx Dev)}
    if Result.GetDouble('TE_COTATION') = 0 then
      Result.SetDouble('TE_COTATION', 1)
    else
      Result.SetDouble('TE_COTATION', Arrondi(1 / Result.GetDouble('TE_COTATION'), NBDECIMALTAUX));
    Result.SetDouble('TE_MONTANT', Arrondi(MntDev * Sens/ Result.GetDouble('TE_COTATION'), V_PGI.OkDecV));
  finally
    if Assigned(ObjDev) then FreeAndNil(ObjDev);
    {$IFNDEF EAGLSERVER}
    TheData := nil;
    {$ENDIF EAGLSERVER}
  end;

  {Si l'on bien récupérer le code CIB et le flux ..}
  if (Result.GetString('TE_CODECIB') > '') and (Result.GetString('TE_CODEFLUX') > '') then begin
    {... Calcul de la date de valeur}
    if DateV > iDate1900 then
      {On récupère la date fournie en paramètre du OnArgument}
      Result.SetDateTime('TE_DATEVALEUR', DateV)
    else if (Nature = na_Simulation) then
      {Dans le cas d'une écriture de simulation, les dates de valeur et comptable sont équivalentes}
      Result.SetDateTime('TE_DATEVALEUR', Result.GetDateTime('TE_DATECOMPTABLE'))
    else
      {Sinon, on la calcul à partir du code CIB}
      Result.SetDateTime('TE_DATEVALEUR', CalcDateValeur(Result.GetString('TE_CODECIB'),
                                                         Result.GetString('TE_GENERAL'),
                                                         Result.GetDateTime('TE_DATECOMPTABLE')));
    if not TesteEcrTresorerie(Result) then begin
      FreeAndNil(Result);
      Exit;
    end;

    {... Génération des clés}
    CodeUnq := Commun.GetNum(CODEUNIQUE, CODEUNIQUE, CODEUNIQUE);
    Result.SetString('TE_CLEVALEUR',    RetourneCleEcriture(Result.GetDateTime('TE_DATEVALEUR'), StrToInt(CodeUnq)));
    Result.SetString('TE_CLEOPERATION', RetourneCleEcriture(Result.GetDateTime('TE_DATECOMPTABLE'), StrToInt(CodeUnq)));
    Commun.SetNum(CODEUNIQUE, CODEUNIQUE, CODEUNIQUE, CodeUnq);

    ObjCom := TObjCommissionTob.Create(Result.GetString('TE_GENERAL'), Result.GetString('TE_CODEFLUX'), Result.GetDateTime('TE_DATECOMPTABLE'));
    try
      {16/07/04 : gestion des commissions avant l'écriture de l'opération financière car le numéro
                  de transaction peut être modifié dans ObjCom.GenererCommissions}
      ObjCom.GenererCommissions(Result);
      ObjCom.TobEcriture.ChangeParent(Result, -1);
      {... mise à jour du compteur}
      Commun.SetNum(CODEMODULE, Result.GetString('TE_SOCIETE'), CODESAISIE, Num);
    finally
      if Assigned(ObjCom) then FreeAndNil(ObjCom);
    end;
  end
  else
    FreeAndNil(Result);
end;

{$IFNDEF NOVH}
{ 17/08/2006 : Fonction retourant un TPieceCompta initialisé depuis le TReglesInteg }
{---------------------------------------------------------------------------------------}
function  TRCreerPieceCompta ( vInfoEcr : TInfoEcriture ; vParent : Tob ; vRegle : TReglesInteg ) : TPieceCompta ;
{---------------------------------------------------------------------------------------}
begin

  result   := TPieceCompta.CreerPiece( vInfoEcr ) ;
  if Assigned( vParent ) then
    result.ChangeParent( vParent, -1 ) ;

  // Paramétrage
  result.InitSaisie ;
{  if not Assigned( GErrEncaDeca ) then
    GErrEncaDeca  := TErrorEncaDeca.CreerMsgErreur ;
  result.OnError := GErrEncaDeca.OnError ;
}
  result.SetMultiEcheOff ; // pas de multi-échéance
  result.Contexte.AttribRIBOff := True ; // pas d'attribution auto du rib :

  // Entête
  result.PutEntete('E_QUALIFPIECE',   'N' ) ;
  result.PutEntete('E_ECRANOUVEAU',   'N' ) ;
  result.PutEntete('E_JOURNAL',       vRegle.Jal ) ;
  result.PutEntete('E_DATECOMPTABLE', vRegle.DtC ) ;
  result.PutEntete('E_DEVISE',        vRegle.Dev ) ;
  result.PutEntete('E_ETABLISSEMENT', result.Contexte.EtablisDefaut );

  // NumeroTemp
  result.AttribNumeroTemp ;

end ;

{17/08/2006 : Génération des pièces comptables depuis la liste des ecritures de tréso}
{---------------------------------------------------------------------------------------}
function TRGenererPieceCompta(var TobGen : TobPieceCompta; vTobOrig : Tob) : Boolean ;
{---------------------------------------------------------------------------------------}
var
  lTobTEcr : Tob ;
  i        : Integer ;
  Regle    : TReglesInteg ;
  ObjTva   : TObjTVA;
  ObjV     : TObjDetailTVA;
  lPieceCP : TPieceCompta ;
  lDossier : string ;
  lBase    : string ;
  lInfo    : TInfoEcriture ;
  lStModeP : string ;
  Q        : TQuery;
  MsgCpta  : TMessageCompta; {08/11/07 : Récupération du message d'erreur}
begin

  Result := False;
  if vTobOrig.Detail.count = 0 then Exit ;

  // Initialisation du TInfoEcriture
  lDossier := vTobOrig.Detail[0].GetString('TE_NODOSSIER') ;
  lBase    := GetInfosFromDossier('DOS_NODOSSIER', lDossier, 'DOS_NOMBASE') ;
  {On recherche si le TInfoEcriture existe}
  lInfo    := TobGen.CreateInfoEcr(lBase);

  // Instanciation de l'objet qui va permettre une éventuelle gestion de la TVA}
  ObjTva := TObjTVA.Create;

  {08/11/07 : Récupération du message d'erreur}
  MsgCpta  := TMessageCompta.Create('');
  try

    {Récupération de la devise de la compta}
    Regle.Dev := V_PGI.DevisePivot;

    {On balaie la liste des écritures à intégrer}
    for i := 0 to vTobOrig.Detail.count - 1 do begin
      {Récupération des infos de la ligne d'écriture}
      lTobTEcr := vTobOrig.Detail[ i ] ;

      {JP 20/09/06 : Ajout de tests pour éviter des intégration non voulues}
      if not IsIntegrable(lTobTEcr) then Continue;

      // Gestion du dossier de génération :
      // - tout passe par le TInfoEcriture : 1 instance par dossier
      if lTobTEcr.GetString('TE_NODOSSIER') <> lDossier then
        begin
        lDossier := lTobTEcr.GetString('TE_NODOSSIER') ;
        lBase    := GetInfosFromDossier('DOS_NODOSSIER', lDossier, 'DOS_NOMBASE') ;
        lInfo    := TobGen.CreateInfoEcr( lBase ) ;
        end ;

      // Récupération du compte depuis le BQCode utilisé en tréso
      Regle.BQGene := GetGeneFromBqCode( lTobTEcr.GetString('TE_GENERAL') ) ;

      {Initialisation de la date comptable}
      Regle.dtC := lTobTEcr.GetDateTime( 'TE_DATECOMPTABLE' ) ;
      Regle.Exo := VG_ObjetExo.GetExercice(Regle.dtC, lBase, True) ;
      Regle.dtV := lTobTEcr.GetDateTime('TE_DATEVALEUR') ;
      {25/09/06 : Maj de la devise qui est fonction du compte bancaire}
      Regle.Dev := lTobTEcr.GetString('TE_DEVISE') ;

      {Si la date comptable ne correspond à un exercice ouvert}
      if Regle.Exo = '' then begin
        if V_PGi.SAV then
          HShowMessage('0;Intégration en comptabilité;La date d''opération ne correspond pas à un exercice ouvert.;W;O;O;O;', '', '');
        ErreurCategorie.TypeErreur := ErreurCategorie.TypeErreur + [CatErr_CPT];
        ErreurCategorie.Ecriture   := ErreurCategorie.Ecriture   + [NatErr_Exo];
        Continue;
      end;

      Result := True;

      {Constitution du rib
       JP 20/07/07 : Uniformisation avec la structure de la Table ECRITURE :
                     ajout de l'* et des / et de la domiciliation}
      if (lTobTEcr.GetString('TE_CODEBANQUE' ) = '') or
         (lTobTEcr.GetString('TE_CODEGUICHET') = '') or
         (lTobTEcr.GetString('TE_NUMCOMPTE'  ) = '') or
         (lTobTEcr.GetString('TE_CLERIB')      = '') then
        Regle.RIB := '*' + lTobTEcr.GetString('TE_IBAN')
      else begin
        Regle.RIB := lTobTEcr.GetString('TE_CODEBANQUE') + '/'
                   + lTobTEcr.GetString('TE_CODEGUICHET') + '/'
                   + lTobTEcr.GetString('TE_NUMCOMPTE') + '/'
                   + lTobTEcr.GetString('TE_CLERIB') + '/';
        Q := OpenSQL('SELECT BQ_DOMICILIATION FROM BANQUECP WHERE BQ_CODE = "' + lTobTEcr.GetString('TE_GENERAL') + '"', True);
        if not Q.EOF then Regle.RIB := Regle.RIB + Q.Fields[0].AsString;
        Ferme(Q);
      end;
      {Initialisation du journal
       11/01/06 : FQ 10323 : déplacé dans la boucle, car il faut chercher le journal pour chaque compte ...}
      //lQJal := OpenSQL('SELECT J_JOURNAL FROM JOURNAL WHERE J_CONTREPARTIE = "' + Regle.BQGene + '"', True);
      //if not lQJal.EOF then
        //Regle.Jal := lQJal.FindField('J_JOURNAL').AsString;
      //Ferme(lQJal);
      Regle.Jal := lTobTEcr.GetString('TE_JOURNAL');

      {16/03/05 : FQ 10223 : Si on n'a pas trouvé de journal pour ce compte, on arrête le processus}
      if Regle.Jal = '' then begin
        if V_PGi.SAV then
          HShowMessage('4;Intégration en comptabilité;Impossible de trouver le journal de banque correspondant au compte "' +
                       Regle.BQGene + '.;W;O;O;O;', '', '');

        ErreurCategorie.TypeErreur := ErreurCategorie.TypeErreur + [CatErr_CPT];
        ErreurCategorie.Ecriture   := ErreurCategorie.Ecriture   + [NatErr_Jal];
        Result := False;
        Exit;
      end;

     {On regarde s'il faut gérer la TVA}
      ObjTva.GetCorrespondance(ObjV, lTobTEcr.GetString('TE_GENERAL'), lTobTEcr.GetString('TE_CODEFLUX') );
      if ObjV <> nil then begin
        Regle.TVA := True;
        Regle.Tau := ObjV.Taux;
        Regle.Cpt := ObjV.Cpte;
        {11/01/06 : Gestion du régime et du code TVA}
        Regle.Cod := ObjV.Code;
        Regle.Reg := ObjV.Regm;

        {16/03/05 : FQ 10223 : si le compte est vide, on risque de se retrouver avec une pièce déséquilibrée}
        if ObjV.Cpte = '' then begin
          if V_PGi.SAV then
            HShowMessage('0;Intégration en comptabilité;Compte général absent dans le paramétrage de la TVA.;W;O;O;O;', '', '');
          ErreurCategorie.TypeErreur := ErreurCategorie.TypeErreur + [CatErr_CPT];
          ErreurCategorie.Ecriture   := ErreurCategorie.Ecriture   + [NatErr_Tva];
        end;
      end else begin
        Regle.TVA := False;
        Regle.Tau := 0;
        Regle.Cpt := '';
      end;

      // Calcul du mode de paiement ET Vérification du Code CIB
      if lTobTEcr.GetDouble('TE_MONTANT') > 0
        then lStModeP := 'ENC'
        else lStModeP := 'DEC';
      lStModeP := GetModeReglement(lStModeP, lTobTEcr.GetString('TE_CODECIB'), lTobTEcr.GetString('TE_GENERAL'), lBase);
      if lStModeP = '@@@' then begin
        // 24/03/05 : FQ 10223 : Nouvelle gestion des erreurs
        if V_PGi.SAV then
          HShowMessage('0;Intégration en comptabilité;Impossible de récupérer le mode de paiement.'#13 +
                       '("' + lTobTEcr.GetString('TE_CODECIB') + '", "' + Regle.BQGene + '");W;O;O;O;', '', '');
        ErreurCategorie.TypeErreur := ErreurCategorie.TypeErreur + [CatErr_CPT];
        ErreurCategorie.Ecriture   := ErreurCategorie.Ecriture  + [NatErr_Cib];
        Exit;
      end else
        Regle.ModeP := lStModeP ;

      // Creation d'une nouvelle piece
      // - sauf si la ligne de tréso à traiter est une commission rattachée à la ligne précédente
      if not ( ( lTobTEcr.GetString('TE_COMMISSION') = suc_Commission ) and                            // Commission ?
               Assigned( lPieceCP ) and ( lPieceCP.Detail.count > 1 ) and                              // Piece précédente valide ?
               ( lPieceCP.Detail[0].GetString('TE_NUMTRANSAC') = lTobTEcr.GetString('TE_NUMTRANSAC') ) // Même numéro de transaction ?
              ) then
        lPieceCP := TRCreerPieceCompta( lInfo, TobGen, Regle ) ;

      // Génération de la ligne de trésorerie
      TRGenereEcrCompta( lPieceCP, lTobTEcr, Regle );

      // Génération de la ligne de contrepartie
      TRGenereEcrContrepartie( lPieceCP, lTobTEcr, Regle );

      // Génération de la TVA si besoin
      if Regle.TVA then
        begin
        TRGenereEcrTVA( lPieceCP, lTobTEcr, Regle );
        lPieceCP.SetInfosTva( 1, Regle.Reg ,Regle.Cod, '') ;
        end ;

      // test de la pièce
      if not lPieceCP.IsValidPiece then begin
        {18/12/06 : récupération de l'erreur}
        ErreurCategorie.TypeErreur := ErreurCategorie.TypeErreur + [CatErr_CPT];
        ErreurCategorie.Ecriture   := ErreurCategorie.Ecriture  + [NatErr_Cpt];
        {08/11/07 : Récupération du message d'erreur}
        ErreurCategorie.MsgCompta  := MsgCpta.GetMessage(lPieceCP.LastError.RC_Error);
        FreeAndNil( lPieceCP ) ;
      end;

    end; // Fin de la boucle principale ( For i )

  finally
    if Assigned(ObjTva) then
      FreeAndNil(ObjTva);

    if Assigned(MsgCpta) then
      FreeAndNil(MsgCpta);
  end;

end;


{17/08/2005 : Génération de la ligne d'ecriture compta principale depuis la ligne de trésorerie }
{---------------------------------------------------------------------------------------}
procedure TRGenereEcrCompta(var vPiece : TPieceCompta ; vTobTEcr : Tob ; vRegle : TReglesInteg ) ;
{---------------------------------------------------------------------------------------}
var lMontantDev : Double;
    lNumL       : Integer ;
    lTobL       : Tob ;
begin

  // nouvelle ligne :
  lTobL := vPiece.NewRecord ;
  lNumL := vPiece.Detail.count ;

  {13/06/07 : En attendant de voir avec Stéphane quoi faire dans le TPieceCompta}
  vPiece.PutValue( lNumL, 'E_EXERCICE', vRegle.Exo);

  // Compte général
  vPiece.PutValue( lNumL, 'E_GENERAL',         vRegle.BQGene );
  vPiece.PutValue( lNumL, 'E_CONTREPARTIEGEN', vTobTEcr.GetString('TE_CONTREPARTIETR') );
  {30/03/07 : Gestion du pointage}
  vPiece.PutValue( lNumL, 'E_DATEPOINTAGE',    vTobTEcr.GetDateTime('TE_DATERAPPRO'));
  vPiece.PutValue( lNumL, 'E_NATURETRESO',     vTobTEcr.GetString('TE_CODERAPPRO'));
  vPiece.PutValue( lNumL, 'E_REFPOINTAGE',     vTobTEcr.GetString('TE_REFPOINTAGE'));

  // Montant
  lMontantDev := abs( vTobTEcr.GetDouble('TE_MONTANTDEV') ) ;
  {27/03/07 : FQ 10426 : gestion de l'arrondi comme dans GetMntHTFromTTC}
  lMontantDev := Arrondi(lMontantDev, vPiece.Devise.Decimale);
  if vTobTEcr.GetDouble('TE_MONTANTDEV') < 0
    then vPiece.PutValue( lNumL, 'E_CREDITDEV', lMontantDev )
    else vPiece.PutValue( lNumL, 'E_DEBITDEV',  lMontantDev ) ;

  // Autres informations
  TRRemplitPieceStd( vPiece, lNumL, vTobTEcr, vRegle ) ;
  vPiece.PutValue( lNumL, 'E_RIB' , vRegle.RIB);


  // Cas spécifique du mode de paiement
  if VPiece.GetValue( lNumL, 'E_ECHE' ) = 'X' then begin
    vPiece.PutValue( lNumL, 'E_MODEPAIE',        vRegle.ModeP);
    vPiece.PutValue( lNumL, 'E_DATEVALEUR',      vRegle.dtV);
    vPiece.PutValue( lNumL, 'E_DATEECHEANCE',    vRegle.dtV);
    vPiece.PutValue( lNumL, 'E_ORIGINEPAIEMENT', vRegle.dtV);
  end ;

  // Champ supplémentaire pour référence à la ligne de tréso
  lTobL.AddChampSupValeur('TE_MAJTRESO',    'X' ) ;
  lTobL.AddChampSupValeur('TE_NUMTRANSAC',  vTobTEcr.GetValue('TE_NUMTRANSAC') ) ;
  lTobL.AddChampSupValeur('TE_NUMEROPIECE', vTobTEcr.GetValue('TE_NUMEROPIECE') ) ;
  lTobL.AddChampSupValeur('TE_NUMLIGNE',    vTobTEcr.GetValue('TE_NUMLIGNE') ) ;

end ;

{17/08/2005 : Génération de la ligne de contrepartie comptable depuis la ligne de trésorerie }
{---------------------------------------------------------------------------------------}
procedure TRGenereEcrContrepartie(var vPiece : TPieceCompta ; vTobTEcr : Tob ; vRegle : TReglesInteg ) ;
{---------------------------------------------------------------------------------------}
var lMontantDev : Double;
    lNumL       : Integer ;
begin

  // nouvelle ligne :
  vPiece.NewRecord ;
  lNumL := vPiece.Detail.count ;
  {13/06/07 : En attendant de voir avec Stéphane quoi faire dans le TPieceCompta}
  vPiece.PutValue( lNumL, 'E_EXERCICE', vRegle.Exo);

  // Compte général et contrepartie
  vPiece.PutValue( lNumL, 'E_GENERAL', vTobTEcr.GetString('TE_CONTREPARTIETR') );
  vPiece.PutValue( lNumL, 'E_CONTREPARTIEGEN'  , vPiece.GetValue( lNumL - 1, 'E_GENERAL' )     );

  // Montant
  lMontantDev := Abs( vTobTEcr.GetDouble('TE_MONTANTDEV') ) ;
  {27/03/07 : FQ 10426 : gestion de l'arrondi comme dans GetMntHTFromTTC}
  if vRegle.Tva then // s'il y a une gestion de la tva sur le flux en cours
    lMontantDev := GetMntHTFromTTC(lMontantDev, vRegle.Tau, vPiece.Devise.Decimale);
  if vTobTEcr.GetDouble('TE_MONTANTDEV') > 0
    then vPiece.PutValue( lNumL, 'E_CREDITDEV', Arrondi(lMontantDev, vPiece.Devise.Decimale) )
    else vPiece.PutValue( lNumL, 'E_DEBITDEV',  Arrondi(lMontantDev, vPiece.Devise.Decimale) ) ;

  // Autres informations
  TRRemplitPieceStd( vPiece, lNumL, vTobTEcr, vRegle ) ;

  // Cas spécifique du mode de paiement
  if vPiece.GetValue( lNumL, 'E_ECHE' ) = 'X' then begin
    vPiece.PutValue( lNumL, 'E_MODEPAIE',        vRegle.ModeP );
    vPiece.PutValue( lNumL, 'E_DATEVALEUR',      vRegle.dtV);
    vPiece.PutValue( lNumL, 'E_DATEECHEANCE',    vRegle.dtV);
    vPiece.PutValue( lNumL, 'E_ORIGINEPAIEMENT', vRegle.dtV);
    end ;

end ;


{17/08/2005 : Génération de la ligne de TVA compta depuis la ligne de trésorerie }
{---------------------------------------------------------------------------------------}
procedure TRGenereEcrTVA(var vPiece : TPieceCompta ; vTobTEcr : Tob ; vRegle : TReglesInteg ) ;
{---------------------------------------------------------------------------------------}
var lNumL       : Integer ;
begin

  // nouvelle ligne :
  vPiece.NewRecord ;
  lNumL := vPiece.Detail.count ;

  {13/06/07 : En attendant de voir avec Stéphane quoi faire dans le TPieceCompta}
  vPiece.PutValue( lNumL, 'E_EXERCICE', vRegle.Exo);

  // Compte général et contrepartie
  vPiece.PutValue( lNumL, 'E_GENERAL',           vRegle.Cpt );
  vPiece.PutValue( lNumL, 'E_CONTREPARTIEGEN'  , vPiece.GetValue( lNumL - 2, 'E_GENERAL' )     );

  // Montant : on solde la pièce
  vPiece.AttribSolde( lNumL ) ;

  // Autres informations
  TRRemplitPieceStd( vPiece, lNumL, vTobTEcr, vRegle ) ;

  // Cas spécifique du mode de paiement
  if vPiece.GetValue( lNumL, 'E_ECHE' ) = 'X' then begin
    vPiece.PutValue( lNumL, 'E_MODEPAIE',        vRegle.ModeP );
    vPiece.PutValue( lNumL, 'E_DATEVALEUR',      vRegle.dtV);
    vPiece.PutValue( lNumL, 'E_DATEECHEANCE',    vRegle.dtV);
    vPiece.PutValue( lNumL, 'E_ORIGINEPAIEMENT', vRegle.dtV);
    end ;

end ;

{17/08/2005 : Enregistrement des Pieces comptables }
{---------------------------------------------------------------------------------------}
function  TREnregistrePieces ( var TobCompta : TobPieceCompta ; MajCTerme : Boolean ) : Boolean;
{---------------------------------------------------------------------------------------}
var lPiece    : TPieceCompta ;
    i         : integer ;
    j         : integer ;
    lBoOk     : boolean ;
    lTobE     : Tob ;
    lTransac  : string ;
    lStReq    : string ;
    lInTCT    : integer ;
    lInTR     : integer ;
begin

  result := False ;

  // On balaie la liste des pièces générées
  for i := 0 to TobCompta.Detail.count - 1 do begin

    lPiece := TPieceCompta( TobCompta.Detail[ i ] ) ;
    lInTR  := 0 ;
    lInTCT := 1 ;

    try
      BeginTrans;

      // Enregistrement de la pièce
      lBoOk  := lPiece.Save ;
      // Ajout des champs résultats
      lPiece.AddChampSup('RESULTAT',    True);
      lPiece.AddChampSup('COMMENTAIRE', True);

      if lBoOk then begin

        // Parcours des lignes de la pièce ( car il peut y avoir des commissions sur la pièce )
        for j := 0 to lPiece.Detail.count - 1 do begin

          lTobE := lPiece.Detail[ j ] ;
          if (lTobE.GetNumChamp('TE_MAJTRESO') < 0) or ( lTobE.GetValue('TE_MAJTRESO') <> 'X' ) then
            continue ;

          lTransac   := lTobE.GetString('TE_NUMTRANSAC') ;
          lInTCT     := 1 ;

          // MAJ de TRECRITURE
          lStReq := 'UPDATE TRECRITURE SET TE_USERCOMPTABLE = "' + V_PGI.User + '" ,' +
                                         ' TE_NUMEROPIECE = ' + IntToStr( lPiece.GetEntete('E_NUMEROPIECE') ) + ', ' +
                                         ' TE_NUMECHE = 1, ' +
                                         ' TE_CPNUMLIGNE = '  + IntToStr( lTobE.GetInteger('E_NUMLIGNE') ) + ', ' +
                                         ' TE_DATEMODIF = "' + USDateTime( NowH ) + '" '
                  + 'WHERE TE_NUMLIGNE = '    + lTobE.GetString('TE_NUMLIGNE') + ' AND '
                        + 'TE_NUMTRANSAC = "' + lTransac + '" AND '
                        + 'TE_NUMEROPIECE = ' + lTobE.GetString('TE_NUMEROPIECE') + ' AND '
                        + 'TE_NODOSSIER = "'  + lTobE.GetString('TE_NODOSSIER') + '"';
          // Execute
          lInTR := ExecuteSQL(lStReq);

          // Maj de COURTTERME (sauf pour les écriture saisies ou d'équilibrage)
          if (Pos(TRANSACEQUI, UpperCase(lTransac)) < 1 ) and (Pos(TRANSACSAISIE, UpperCase(lTransac)) < 1 ) and MajCTerme then begin
            lStReq := 'UPDATE COURTSTERMES SET TCT_STATUT = "X" WHERE TCT_NUMTRANSAC = "' + lTransac + '" AND '
                                                              + 'TCT_NODOSSIER = "'  + lTobE.GetString('TE_NODOSSIER') + '"';
            // Execute
            lInTCT := ExecuteSQL(lStReq);
          end ;

          // MAJ des messages
          if lInTR = 0 then begin
            lPiece.PutValue('RESULTAT',    'Echec' ) ;
            lPiece.PutValue('COMMENTAIRE', 'Impossible de mettre à jour l''écriture de trésorerie !' ) ;
            break ;
          end else if lInTCT = 0 then begin
            lPiece.PutValue('RESULTAT',    'Echec' ) ;
            lPiece.PutValue('COMMENTAIRE', 'Impossible de mettre à jour les courts termes !' ) ;
            break ;
          end else begin
            lPiece.PutValue('RESULTAT',    'OK' ) ;
            lPiece.PutValue('COMMENTAIRE', '' ) ;
          end ;

        end // fin boucle lignes

      end else begin
        lPiece.PutValue('RESULTAT',    'Echec' ) ;
        lPiece.PutValue('COMMENTAIRE', 'Impossible d''enregistrer la pièce comptable !' ) ;
      end ;

      // Validation des transactions
      if not lBoOk or ( lInTR = 0 ) or ( lInTCT = 0 ) // Cas de rejet
        then RollBackDiscret
        else begin
             CommitTrans ;
             result := True ;
      end ;


    except // Try...
      on E : Exception do begin
        lPiece.PutValue('RESULTAT',    'Echec' ) ;
        lPiece.PutValue('COMMENTAIRE', E.Message ) ;
        RollBack;
      end;
    end ;

  end ; // fin boucle pièces

end ;


{17/08/2005 : Remplit les champs standards d'une ligne ECRITURE à partir d'une ligne TRECRITURE}
{---------------------------------------------------------------------------------------}
procedure TRRemplitPieceStd  (var vPiece : TPieceCompta; vNumL : integer ; vTobTEcr : TOB ; vRegle : TReglesInteg );
{---------------------------------------------------------------------------------------}
begin
  vPiece.PutValue( vNumL, 'E_QUALIFORIGINE'    , QUALIFTRESO );
  vPiece.PutValue( vNumL, 'E_VALIDE'           , 'X'         );
  vPiece.PutValue( vNumL, 'E_IO'               , 'X'         );
  vPiece.PutValue( vNumL, 'E_TRESOSYNCHRO'     , ets_Synchro );
  vPiece.PutValue( vNumL, 'E_CREERPAR'         , 'IMP'       );
  vPiece.PutValue( vNumL, 'E_LIBELLE'          , vTobTEcr.GetString('TE_LIBELLE') );
  vPiece.PutValue( vNumL, 'E_REFINTERNE'       , vTobTEcr.GetString('TE_NUMTRANSAC')
                                             + '-' + vTobTEcr.GetString('TE_NUMEROPIECE')
                                             + '-' + vTobTEcr.GetString('TE_NUMLIGNE')    );

  if vRegle.Tva then begin
    vPiece.PutValue( vNumL, 'E_TVA',       vRegle.Cod ) ;
    vPiece.PutValue( vNumL, 'E_REGIMETVA', vRegle.Reg ) ;
  end ;

  vPiece.GetTob( vNumL ).AddChampSupValeur('TE_NODOSSIER',   vTobTEcr.GetValue('TE_NODOSSIER') ) ;

end ;


{21/08/2005 : Intégration des Pieces comptables }
{---------------------------------------------------------------------------------------}
function TRIntegrationPieces (var TobCompta : TobPieceCompta; MajCTerme : Boolean ) : Boolean ;
{---------------------------------------------------------------------------------------}
begin
  Result := False;
  // Directive pour éviter de ramener la Treso dans la Compta et CCMP
  {$IFDEF TRESO}

  // Affichage en liste des futures écritures comptables
  if TRAfficheListePieces( TobCompta ) then begin

    // Si au moins une écriture a été intégrée ...
    if TREnregistrePieces( TobCompta, MajCTerme ) then begin
      // ... et on affiche le résultat de l'intégration
      TRAffichePiecesInteg( TobCompta );
      Result := True;
    end;

  end;
  {$ENDIF}
end ;

{$ENDIF NOVH}

{ TObjetExercice }

{---------------------------------------------------------------------------------------}
constructor TObjetExercice.Create;
{---------------------------------------------------------------------------------------}
begin
  FListeExo := TStringList.Create;
end;

{---------------------------------------------------------------------------------------}
procedure TObjetExercice.ChargeObjet;
{---------------------------------------------------------------------------------------}
var
  lDossier  : string;
  UnDossier : string;
  UnExo     : TZExercice;
begin
  {$IFDEF TRESO}
  ChargeVarTreso;
  lDossier := VarTreso.lNomBase;
  {$ELSE}
  lDossier := GetBasesMS;
  {$ENDIF TRESO}

  if Assigned(FListeExo) then LibereListe(FListeExo, False);
  {19/06/07 : FQ 10478 : gestion des exercices en mono dossier}
  if lDossier = '' then begin
    UnDossier := V_PGI.SchemaName;
    UnExo := TZExercice.Create(False, UnDossier);
    FListeExo.AddObject(UnDossier, UnExo);
  end
  else begin
    UnDossier := ReadTokenSt(lDossier);
    while UnDossier <> '' do begin
      UnExo := TZExercice.Create(False, UnDossier);
      FListeExo.AddObject(UnDossier, UnExo);
      UnDossier := ReadTokenSt(lDossier);
    end;
  end;
end;

{---------------------------------------------------------------------------------------}
destructor TObjetExercice.Destroy;
{---------------------------------------------------------------------------------------}
begin
  if Assigned(FListeExo) then LibereListe(FListeExo, True);
  inherited;
end;

{---------------------------------------------------------------------------------------}
function TObjetExercice.GetTZExercice(aDossier : string) : TZExercice;
{---------------------------------------------------------------------------------------}
var
  n : Integer;
begin
  Result := nil;
  if aDossier = '' then aDossier := V_PGI.SchemaName;
  FDossier := aDossier;

  for n := 0 to FListeExo.Count - 1 do begin
    if FListeExo[n] = aDossier then begin
      Result := TZExercice(FListeExo.Objects[n]);
      Break;
    end;
  end;
end;

{---------------------------------------------------------------------------------------}
function TObjetExercice.GetExoNodos(DateOpe : TDateTime; aNoDossier : string = ''; PourCompta : Boolean = False) : string;
{---------------------------------------------------------------------------------------}
var
  aDossier : string;
  Q : TQuery;
begin
  Result := '';
  {On n'est pas en multi Dossiers ou sur le dossier courant}
  if (aNoDossier = '') or (aNoDossier = CODEDOSSIERDEF) or (aNoDossier = V_PGI.NoDossier) then
    Result := GetExercice(DateOpe, V_PGI.SchemaName, PourCompta)
  else begin
    {... Récupération du dossier correspondant au NoDossier}
    Q := OpenSQL('SELECT DOS_NOMBASE FROM DOSSIER WHERE DOS_NODOSSIER = "' + aNoDossier + '"', True);
    try
      if not Q.EOF then begin
        aDossier := Q.FindField('DOS_NOMBASE').AsString;
        Result := GetExercice(DateOpe, aDossier, PourCompta);
      end
      else
        if V_PGI.SAV then PGIBox(TraduireMemoire('Impossible de récupérer de code dossier'), NomHalley);
    finally
      Ferme(Q);
    end;
  end;
end;

{---------------------------------------------------------------------------------------}
function TObjetExercice.GetExercice(DateOpe : TDateTime; aDossier : string = ''; PourCompta : Boolean = False) : string;
{---------------------------------------------------------------------------------------}
var
  oExercice : TZExercice;
  n : Integer;
  B : Boolean;
begin
  Result := '';
  {Recherche de l'objet exercice correspondant au dossier}
  oExercice := GetTZExercice(aDossier);

  if not Assigned(oExercice) then begin
    if V_PGI.SAV then PGIInfo(TraduireMemoire('Il n''y a pas d''exercice correspondant à la date du ' + DateToStr(DateOpe) + '.'));
    Exit;
  end;

  {Booléen qui permet de savoir s'il existe un exercice (Ouvert ou non) pour la date passé en paramètre}
  B := False;
  for n := 0 to oExercice.NbExercices - 1 do begin
    if (oExercice.Exercices[n].Deb <= DateOpe) and (oExercice.Exercices[n].Fin >= DateOpe) then begin
      B := True;
      {Si l'exercice n'est pas ouvert ...}
      if (oExercice.Exercices[n].EtatCpta <> 'OUV') and
         {11/10/07 : FQ 10530 : on peut écrire sur des clôture provisoires}
         (oExercice.Exercices[n].EtatCpta <> 'CPR') then begin
        {... si c'est pour une écriture de trésorerie, on avertit, mais on ne bloque pas en renvoyant quand
             même l'exercice trouvé}
        if not PourCompta then begin
          if V_PGI.SAV then
            PGIInfo(TraduireMemoire('L''exercice correspondant à la date du ' + DateToStr(DateOpe) + ' n''est pas ouvert.'));
          Result := oExercice.Exercices[n].Code;
        end;
        {... Par contre pour la compta on renvoie vide}
      end
      else
        Result := oExercice.Exercices[n].Code;
      Break;
    end;
  end;

  {Pour la Trésorerie, on avertit si l'exercice n'existe pas}
  if (Result = '') and not B and not PourCompta and V_PGI.SAV then
    PGIInfo(TraduireMemoire('Il n''y a pas d''exercice correspondant à la date du ' + DateToStr(DateOpe) + '.'));
end;

{---------------------------------------------------------------------------------------}
function TObjetExercice.IsExoOuvert(DateOpe : TDateTime; aDossier : string = '') : Boolean;
{---------------------------------------------------------------------------------------}
var
  oExercice : TZExercice;
  n : Integer;
begin
  Result := False;
  {Recherche de l'objet exercice correspondant au dossier}
  oExercice := GetTZExercice(aDossier);
  if not Assigned(oExercice) then Exit;

  for n := 0 to oExercice.NbExercices - 1 do begin
    if (oExercice.Exercices[n].Deb <= DateOpe) and
       (oExercice.Exercices[n].Fin >= DateOpe) and
       ((oExercice.Exercices[n].EtatCpta = 'OUV') or
       {11/10/07 : FQ 10530 : on peut écrire sur des clôture provisoires ?!}
        (oExercice.Exercices[n].EtatCpta = 'CPR')) then begin
      Result := True;
      Break;
    end;
  end;
end;

{---------------------------------------------------------------------------------------}
function TObjetExercice.GetFinExo(aExo : string; aDossier : string = '') : TDateTime;
{---------------------------------------------------------------------------------------}
var
  oExercice : TZExercice;
  n : Integer;
begin
  Result := iDate1900;
  {Recherche de l'objet exercice correspondant au dossier}
  oExercice := GetTZExercice(aDossier);
  if not Assigned(oExercice) then Exit;

  for n := 0 to oExercice.NbExercices - 1 do
    if oExercice.Exercices[n].Code = aExo then begin
      Result := oExercice.Exercices[n].Fin;
      Break;
    end;
end;

{---------------------------------------------------------------------------------------}
function TObjetExercice.GetDebExo(aExo : string; aDossier : string = '') : TDateTime;
{---------------------------------------------------------------------------------------}
var
  oExercice : TZExercice;
  n : Integer;
begin
  Result := iDate1900;
  {Recherche de l'objet exercice correspondant au dossier}
  oExercice := GetTZExercice(aDossier);
  if not Assigned(oExercice) then Exit;

  for n := 0 to oExercice.NbExercices - 1 do
    if oExercice.Exercices[n].Code = aExo then begin
      Result := oExercice.Exercices[n].Deb;
      Break;
    end;
end;

{---------------------------------------------------------------------------------------}
function TObjetExercice.GetEnCours : TExoDate;
{---------------------------------------------------------------------------------------}
begin
  Result := FExercice.EnCours;
end;

{---------------------------------------------------------------------------------------}
function TObjetExercice.GetExoV8 : TExoDate;
{---------------------------------------------------------------------------------------}
begin
  Result := FExercice.ExoV8;
end;

{---------------------------------------------------------------------------------------}
function TObjetExercice.GetExercices(n : Integer) : TExoDate;
{---------------------------------------------------------------------------------------}
begin
  Result := FExercice.Exercices[n];
end;

{---------------------------------------------------------------------------------------}
procedure TObjetExercice.PositionneExercice(aDossier : string);
{---------------------------------------------------------------------------------------}
begin
  FExercice := GetTZExercice(aDossier);
end;

{---------------------------------------------------------------------------------------}
class function TObjetExercice.GetCurrent : TObjetExercice;
{---------------------------------------------------------------------------------------}
{$IFDEF EAGLSERVER}
var
 MySession      : TISession ;
 lIndex         : integer ;
{$ENDIF}
begin
  {$IFDEF EAGLSERVER}
  {Récupération de la session courante}
  MySession := LookupCurrentSession ;
  {On regarde s'il existe un Objet Exercice dans la session}
  lIndex    := MySession.UserObjects.IndexOf(ClassName);

  if lIndex > -1 then
    Result := TObjetExercice(MySession.UserObjects.Objects[lIndex])
  else begin
    {Création d'une instance de l'objet exercice}
    Result := TObjetExercice.Create;
    {Ajout de l'instance à la session}
    MySession.UserObjects.AddObject(ClassName, Result);
  end;
  {$ELSE}
  {Dans les autres modes, on renvoie la variable globale VG_ObjetExo}
  if VG_ObjetExo = nil then 
    VG_ObjetExo := TObjetExercice.Create;
  Result := VG_ObjetExo;
  {$ENDIF}
end;

end.
