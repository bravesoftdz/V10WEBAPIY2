{***********UNITE*************************************************
Auteur  ...... : TGA
Créé le ...... : 02/01/2006
Modifié le ... :   /  /
Description .. : Externalisation du Traitement de l'éclatement
                 appelée par imeclate et Amchgtmethode
Mots clefs ... :
Suite..........: TGA - 26/01/2006 Ajout champs sur table IMMO
                 TGA - 01/02/2006 FQ 17401
                 TGA - 04/04/2006 Ajout i_reprisedepcedee
                 BTY - 04/06 FQ 17516 Ajout i_opereg
                 TGA 07/04/06 Initialisation de champs IMMO, IMMOLOG
                 TGA 18/056/2006 suite Agl 179
Suite......... : BTY - 06/06 - FQ 18394 - En série, reprendre le no de serie pour qu'en annulation
//               on puisse annuler les autres opérations de la série
Suite......... : XVI 28/02/2007 S/FQ Nouveaux types des variants pour D7 (préconisation CBP)
Suite......... : MBO - 09/2006 - Gestion de la prime d'équipement
Suite......... : MBO - FQ 19851 - 19.03.2007 - correction anomalie si plan fiscal plus long que l'éco
                                               dernière dotation fiscale de la fille non proratisée
                                             + pb arrondi dernière dotation fiscale de la mère
Suite......... : MBO - FQ 17512 - 19.03.2007 - Prise en compte de la gestion fiscale
Suite......... : BTY - FQ 19968 En éclatement, reporter l'analytique de l'immo mère sur l'immo fille
*****************************************************************}
unit imeclate2;

interface

uses  SysUtils,
      HCtrls,
      HEnt1,
      Utob,
      AmType,
      {$IFDEF eAGLClient}
      db,
      {$ELSE}
      {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
      {$ENDIF eAGLClient}

      {$IFDEF VER150}
      Variants,
      {$ENDIF}

      implan,
      OpEnCour,
      Classes,
      immo_tom,
      outils,
{$IFNDEF CMPGIS35}
      AMSYNTHESEDPI_TOF,
{$ENDIF}
      ImEnt ;

type

    TAmPlan = class (TOB)
    public
      fPlanActif  : integer;
      fbFiscal    : boolean;
      fBaseEco    : double;
      fBaseFisc   : double;
      fBasePri    : double;
      fBaseSBV    : double;
      fRepriseEco : double;
      fRepriseFisc:double;
      fRepriseUO :double;
      fRepriseSBV : double;
      fPRI : boolean;
      fSBV : boolean;
      fReprisePrem : boolean;
      fRepriseDR :double;      //ajout pour fq 17512
      fRepriseFEC:double;
      fGestionFiscale :boolean;


      constructor Create (T: Tob);  reintroduce;  //YCP 21/10/05
      destructor  Destroy; override;
      function Enregistre : boolean;
      function Enregistre_bis : boolean;
      function Enregistre_PRISVB(NumPlanFille:integer) : boolean;
      procedure CopieDetail ( PSource : TAmPlan );
      procedure CalculDerogatoire;
    published
      property PlanActif : integer read fPlanActif;
      property Fiscal : boolean read fbFiscal;
      property BaseEco : double read fBaseEco;
      property BaseFisc : double read fBaseFisc;
      property RepriseEco : double read fRepriseEco;
      property RepriseFisc : double read fRepriseFisc;

      // ajout mbo
      property RepriseUO : double read fRepriseUO;
      property RepriseSBV :double read fRepriseSBV;
      property BasePRI : double read fBasePri;
      property BaseSBV : double read fbaseSBV;
      property PRI : boolean read fPRI;
      property SBV : boolean read fSBV;
      property ReprisePrem : boolean read fReprisePrem;

      // ajout mbo pour fq 17512
      property RepriseDR : double read fRepriseDR;
      property RepriseFEC : double read fRepriseFEC;
      property GestionFiscale : boolean read fGestionFiscale;

    end;


procedure enregistreEclatement_2(FFcode:String;FFMttHt:double;FFQuantiteAvEcl:double;
   FFnouveaucode:String;FFValeur:double;FFnew_Valeur:double;FFLibelle:String;FFQuantite:double;
   FFDateope:String;Boldecl:Boolean;FFBlocnote:String;FFOrdreSerie:integer );
// FQ 18394 FFDateope:String;Boldecl:Boolean;FFBlocnote:String );

procedure TraiteElementExceptionnel(CodeOrig,CodeEclate:string;PartEclatee:double;FFDateope:string);
procedure CalculSurChampsTob(T1: TOB;LeChamp,Operande: string;X: Variant;ValArrondi:integer) ;
procedure EclatePlan ( PIni, PDest : TAmPlan; dProrata : double ; ori_dProrata : double ;bElemExcept:double;
                       var TabPriFille, TabSbvFille : array of double);
// ajout mbo pour traitement de la prime d'équipement
procedure EnregLogPri(FFCode:string; NewCode:string; NewPrime:double; FFDateope: string;
                      TabPriFille, TabSbvFille :array of double; var NumPlanFille:integer; TrtSbv:boolean);
// ajout mbo pour traitement de la subvention d'équipement
procedure EnregLogSBV(FFCode:string; NewCode:string; NewSBV:double; FFDateope: string;
                      TabPriFille,TabSbvFille :array of double; var NumPlanFille:integer; TrtPri:boolean);


const
  // fq 17512 ajout des champs i_reprisedr, et i_reprisefec
  LesChamps1: array[1..17] of string=('I_MONTANTHT','I_VALEURACHAT','I_BASETAXEPRO',
  'I_BASEECO','I_BASEFISC','I_TVARECUPERABLE','I_TVARECUPEREE','I_MONTANTBASEAMORT',
  'I_BASEAMORDEBEXO','I_BASEAMORFINEXO','I_MONTANTEXC','I_REPRISEECO','I_REPRISEFISCAL',
  'I_REINTEGRATION','I_UNITEOEUVRE','I_REPRISEDR', 'I_REPRISEFEC'
  ) ;
  LesChamps2:array [1..5] of string =('I_MONTANTPREMECHE','I_MONTANTSUIVECHE','I_FRAISECHE',
  'I_DEPOTGARANTIE','I_RESIDUEL') ;

  LesChamps3:array [1..2] of string = ('I_SBVPRI', 'I_REPRISEUO');
  LesChamps4:array [1..2] of string = ('I_SBVMT', 'I_CORRECTIONVR');

implementation


{***********A.G.L.***********************************************
Auteur  ...... : Thierry GAUZAN
Créé le ...... : 22/11/2005
Modifié le ... :   /  /
Description .. : enregistrement de l'éclatement
Suite ........ :
Mots clefs ... :

 en entrée FFcode = Fcode = code immo mère
           FFMttHt = FMttHt = montant ht immo mère
           FFQuantiteAvEcl = FQuantiteAvEcl = Quantité immo mère
           FFnouveaucode = Fnouveaucode = code immo fille
           FFValeur = FValeur = montant ht immo fille
           FFnew_valeur = nouvelle valorisation immo fille
           FFLibelle = FLibelle = libelle immo fille
           FFQuantite = FQuantite = Quantité immo fille
           FFDateOpe = DateOpe = Date opération éclatement
           Boldecl = Boldeclate = type de traitement
           FFblocnote = il_blocnote = bloc-note
*****************************************************************}
procedure enregistreEclatement_2(FFcode: String;
                                 FFMttHt : double;
                                 FFQuantiteAvEcl : double;
                                 FFnouveaucode: String;
                                 FFValeur : double;
                                 FFnew_Valeur : double;
                                 FFLibelle : String;
                                 FFQuantite : double;
                                 FFDateope : String;
                                 Boldecl : Boolean;
                                 FFblocnote:String;
                                 FFOrdreSerie: integer );    // FQ 18394

var New_PartEclatee, PartEclatee, TQteO: double ; i: integer ;
    T1,T2,T3,T4: TOB ; OkEche: boolean ; Q1: TQuery ;
    T1UO,T2UO : TOB;
    infos : TInfoLog;
    P1,P2 : TPlanAmort;
    P1Detail : TOB;
    QPlan : TQuery;
    PlanIni , PlanDest : TAmPlan;
    ffNumPlanAvant, ffNumPlanApres : integer;
    x, ffOrdre, ffOrdreS : integer;
    bElemExcept : double;
    revalorisation : Boolean;
    GestPrime : boolean;
    GestSBV : boolean;
    NewPrime :double;
    NewSBV :double;
    SAVreprisePRI :double;
    SAVrepriseSBV :double;
    SAVmntPRI :double;
    SAVmntSBV :double;

    Qlog : TQuery;   // ajout mbo
    PlanPRI : integer;
    PlanSBV : integer;
    FirstPRI : boolean;

    TabPriFille : array[0..MAX_LIGNEDOT] of double;
    TabSbvFille : array[0..MAX_LIGNEDOT] of double;
    NumPlanFille : integer;

Begin
  revalorisation:=False;

  for i := 0 to MAX_LIGNEDOT do
  begin
    TabPriFille[i] := 0;
    TabSbvFille[i] := 0;
  end;

  IF (FFNew_Valeur<>0) AND (FFNew_Valeur<>FFValeur)THEN
    revalorisation:=True;

  ffOrdre := TrouveNumeroOrdreLogSuivant(ffCode);
  // FQ 18394
  //ffOrdreS := TrouveNumeroOrdreSerieLogSuivant;
  ffOrdreS := FFOrdreSerie;


  T1:=TOB.Create('IMMO',nil,-1) ;
  T1.SelectDB('"'+FFCode+'"',nil) ;

  SAVreprisePRI := (T1.GetValue('I_REPRISEUO'));
  SAVrepriseSBV := (T1.GetValue('I_CORRECTIONVR'));
  SAVmntPRI := (T1.GetValue('I_SBVPRI'));
  SAVmntSBV := (T1.GetValue('I_SBVMT'));


  T2:=TOB.Create('IMMO',nil,-1) ;
  T2.Dupliquer(T1,true,true) ;

  OkEche:=(T1.GetValue('I_NATUREIMMO')='CB') or (T1.GetValue('I_NATUREIMMO')='LOC') ;

  GestPrime := (T1.GetValue('I_SBVPRI')<> 0);
  GestSBV := (T1.GetValue('I_SBVMT')<> 0);


  // TGA 24/11/2005 bElemExcept:=(T1.GetValue('I_MONTANTEXC')<>0);
  bElemExcept:=T1.GetValue('I_MONTANTEXC');

  // Chargement des plan avant éclatement - FQ 14789 - CA 11/10/2004
  QPlan:=OpenSQL('SELECT * FROM IMMO WHERE I_IMMO="'+FFCode+'"', FALSE) ;
  P1:=TPlanAmort.Create(true) ;
  P1.Charge(QPlan);
  P1.Recupere(FFCode,QPlan.FindField('I_PLANACTIF').AsString);
  P2:=TPlanAmort.Create(true) ;
  P2.Charge(QPlan);
  P2.Recupere(FFCode,QPlan.FindField('I_PLANACTIF').AsString);
  Ferme (QPlan);

  //Gestion du prorata de la nouvelle Immo

  if FFMttHt<>0 then PartEclatee:=FFValeur/FFMttHt else PartEclatee:=0;

  //Gestion du prorata de la nouvelle Immo si revalorisation
  if ((FFMttHt<>0) AND (revalorisation=True))then
     New_PartEclatee:=FFNew_Valeur/FFMttHt
  else
     New_PartEclatee := PartEclatee;


  //Recalcul de l'antériorité sur immo origine
  P1.CalculProrataDotationOrigine(P1.AmortEco,(1-PartEclatee));

  //Recalcul de l'antériorité sur immo fille
  P2.CalculProrataDotationOrigine(P2.AmortEco,New_PartEclatee);

  //Gestion de l'Immo Fille
  T2.PutValue('I_IMMO',FFNouveauCode) ;
  T2.PutValue('I_IMMOORIGINEECL',FFCode) ;
  T2.PutValue('I_OPEMUTATION','-') ;
  T2.PutValue('I_OPECESSION','-') ;
  T2.PutValue('I_OPECHANGEPLAN','-') ;
  T2.PutValue('I_OPELIEUGEO','-') ;
  T2.PutValue('I_OPEETABLISSEMENT','-') ;
  T2.PutValue('I_OPELEVEEOPTION','-') ;
  T2.PutValue('I_OPEMODIFBASES','-') ;
  // BTY FQ 17259 01/06 Nouveau top de dépréciation
  T2.PutValue('I_OPEDEPREC','-') ;

  T2.PutValue('I_OPERATION','-') ;
  // ajout mbo 19.10.06
  if (T1.GetValue('I_SBVPRI') <> 0) or (T1.GetValue('I_SBVMT') <> 0) then
    T2.PutValue('I_OPERATION','X') ;


  T2.PutValue('I_CHANGECODE',FFCode) ;
  T2.PutValue('I_LIBELLE',FFLibelle) ;
  T2.PutValue('I_QUANTITE',FFQuantite) ;
  // TGA 26/01/2006
  T2.PutValue('I_REPRISEDEP',0);
  T2.PutValue('I_ECCLEECR','');
  T2.PutValue('I_DOCGUID','');
  T2.PutValue('I_DATEFINCB',iDate1900);
  // TGA 04/04/2006
  T2.PutValue('I_REPRISEDEPCEDEE',0);
  // TGA 01/02/2006 FQ 17401
  T2.PutValue('I_OPEECLATEMENT','-');
  // BTY 04/06 FQ 17516
  T2.PutValue('I_OPEREG','-') ;
  // ajout tga 07.04.06
  T2.PutValue('I_SUSDEF','A');
  T2.PutValue('I_REGLECESSION','NOR');
  // fq 17512 mbo  doit être comme la mère T2.PutValue('I_NONDED','-');
  T2.PutValue('I_REPRISEINT',0);
  T2.PutValue('I_REPRISEINTCEDEE',0);
  T2.PutValue('I_DPI','-');
  // mbo - on ne réinitialise pas car doit être comme la mère - T2.PutValue('I_DPIEC','-');
  // T2.PutValue('I_CORRECTIONVR',0);
  // MBO - on ne réinitialise pas le champ T2.PutValue('I_CORVRCEDDE',0);
  T2.PutValue('I_SUBVENTION','NON');

  //T2.PutValue('I_SBVPRI',0);
  //T2.PutValue('I_SBVMT,0);
  //T2.PutValue('I_SBVMTC',0);
  //T2.PutValue('I_SBVPRIC',0);

  T2.PutValue('I_SBVEC','C');
  // mbo - on n'initialise pas car il peut y avoir une subvention - T2.PutValue('I_SBVDATE',iDate1900);
  // T2.PutValue('I_CPTSBVR','');
  // T2.PutValue('I_CPTSBVB','');
  T2.PutValue('I_PFR','-');
  T2.PutValue('I_COEFDEG',0);
  T2.PutValue('I_AMTFOR',0);
  T2.PutValue('I_AMTFORC',0);
  T2.PutValue('I_ACHFOR',iDate1900);
  T2.PutValue('I_PRIXACFORC',0);
  T2.PutValue('I_VNCFOR',0);
  T2.PutValue('I_DURRESTFOR',0);

  { Modif mbo fq 18149
  T2.PutValue('I_DATEDEBECO',T2.GetValue('I_DATEAMORT'));
  T2.PutValue('I_DATEDEBFISC',T2.GetValue('I_DATEAMORT')); }
  T2.PutValue('I_DATEDEBECO',T2.GetValue('I_DATEDEBECO'));
  T2.PutValue('I_DATEDEBFIS',T2.GetValue('I_DATEDEBFIS'));

  // FQ 19968 appliquer l'analytique de la mère à la fille
  if T1.GetValue('I_VENTILABLE')='X' then
     ImDupliqueVentil(FFcode, FFNouveaucode) ;


  // TGA 25/08/2006 gestion DPI sur fille
  IF T1.GetValue('I_DPI')='X' Then
    T2.PutValue('I_DPI','X');

  for i:=low(LesChamps1) to high(LesChamps1) do CalculSurChampsTob(T2,LesChamps1[i],'*',New_PartEclatee,V_PGI.OkDecV) ;

  for i:=low(LesChamps3) to high(LesChamps3) do CalculSurChampsTob(T2,LesChamps3[i],'*',New_PartEclatee,V_PGI.OkDecV) ;
  for i:=low(LesChamps4) to high(LesChamps4) do CalculSurChampsTob(T2,LesChamps4[i],'*',New_PartEclatee,V_PGI.OkDecV) ;

  NewPrime := T2.GetValue('I_SBVPRI');   // sert pour la mise à jour de Immolog
  NewSBV := T2.GetValue('I_SBVMT');

  if OkEche then
    for i:=low(LesChamps2) to high(LesChamps2) do CalculSurChampsTob(T2,LesChamps2[i],'*',New_PartEclatee,V_PGI.OkDecV) ;

  //Gestion de l'ancienne Immo
  if FFQuantiteAvEcl<>1 then  TQteO:=FFQuantiteAvEcl-FFQuantite else TQteO:=1;

  T1.PutValue('I_CHANGECODE',FFNouveauCode) ;
  T1.PutValue('I_QUANTITE',TQteO) ;
  T1.PutValue('I_OPEECLATEMENT','X');
  T1.PutValue('I_OPERATION','X') ;
  //ajout mbo pour prime et subvention
  //T1.PutValue('I_SBVPRI',(T1.GetValue('I_SBVPRI')- T2.GetValue('I_SBVPRI')));
  //T1.PutValue('I_SBVMT',(T1.GetValue('I_SBVMT')-NewSBV));

  IF revalorisation=True Then
    Begin
     for i:=low(LesChamps1) to high(LesChamps1) do CalculSurChampsTob(T1,LesChamps1[i],'*', (1-PartEclatee),V_PGI.OkDecV) ;

     if (gestprime) then
       for i:=low(LesChamps3) to high(LesChamps3) do CalculSurChampsTob(T1,LesChamps3[i],'*', (1-PartEclatee),V_PGI.OkDecV) ;

     if (GestSBV) then
       for i:=low(LesChamps4) to high(LesChamps4) do CalculSurChampsTob(T1,LesChamps4[i],'*', (1-PartEclatee),V_PGI.OkDecV) ;

     if OkEche then
       for i:=low(LesChamps2) to high(LesChamps2) do CalculSurChampsTob(T1,LesChamps2[i],'*',(1-PartEclatee),V_PGI.OkDecV) ;
    end
  Else
    Begin
     for i:=low(LesChamps1) to high(LesChamps1) do CalculSurChampsTob(T1,LesChamps1[i],'-',T2.GetValue(LesChamps1[i]),V_PGI.OkDecV) ;

     if (gestprime) then
        for i:=low(LesChamps3) to high(LesChamps3) do CalculSurChampsTob(T1,LesChamps3[i],'-',T2.GetValue(LesChamps3[i]),V_PGI.OkDecV) ;

     if (gestSBV) then
        for i:=low(LesChamps4) to high(LesChamps4) do CalculSurChampsTob(T1,LesChamps4[i],'-',T2.GetValue(LesChamps4[i]),V_PGI.OkDecV) ;

     if OkEche then
      for i:=low(LesChamps2) to high(LesChamps2) do CalculSurChampsTob(T1,LesChamps2[i],'-',T2.GetValue(LesChamps2[i]),V_PGI.OkDecV) ;
  End;

  { Gestion de la table IMMOUO }
  Q1:=OpenSQL('SELECT * FROM IMMOUO WHERE IUO_IMMO="'+FFCode+'" ORDER BY IUO_DATE',false);
  if not Q1.Eof then
  begin
    T1UO:=Tob.Create('IMMOUO',nil,-1) ;
    T1UO.LoadDetailDB('IMMOUO','','',Q1,true) ;
    T2UO:=Tob.Create('IMMOUO',nil,-1) ;
    Ferme(Q1) ;
    for i:=0 to T1UO.Detail.Count-1 do
    begin
      T3:=T1UO.Detail[i] ;  T4:=TOB.Create('IMMOUO',T2UO,-1) ;
      T4.Dupliquer(T3,false,true) ;
      T4.PutValue('IUO_IMMO',FFNouveauCode);

      // TGA Gestion Revalorisation
      IF revalorisation=True Then
        Begin
         CalculSurChampsTob(T4,'IUO_UNITEOEUVRE','*',New_PartEclatee,V_PGI.OkDecV) ;
         CalculSurChampsTob(T3,'IUO_UNITEOEUVRE','*',(1-PartEclatee),V_PGI.OkDecV) ;
        End
      Else
        Begin
         CalculSurChampsTob(T4,'IUO_UNITEOEUVRE','*',PartEclatee,V_PGI.OkDecV) ;
         CalculSurChampsTob(T3,'IUO_UNITEOEUVRE','-',T4.GetValue('IUO_UNITEOEUVRE'),V_PGI.OkDecV) ;
        End;
    end ;
    T1UO.InsertOrUpdateDB ; T2UO.InsertOrUpdateDB ;
    T1UO.Free ; T2UO.Free ;
  end;

  if not BOLDECL then
  begin
   PlanIni := TAmPlan.Create ( T1 );
   PlanDest := TAmPlan.Create ( T2 );
   ffNumPlanAvant := PlanIni.PlanActif;

   EclatePlan ( PlanIni, PlanDest, New_PartEclatee , PartEclatee , bElemExcept, TabPriFille, TabSbvFille);

   PlanIni.Enregistre;
   PlanDest.Enregistre_bis;
   T1.PutValue('I_PLANACTIF',PlanIni.PlanActif) ;
   //T2.PutValue('I_PLANACTIF',PlanDest.PlanActif) ;
   T2.PutValue('I_PLANACTIF', 1) ;
   ffNumPlanApres := PlanIni.PlanActif;
   NumPlanFille := 1;
   PlanIni.Free;
   PlanDest.Free;
  end else
  begin
    { Calcul des plans d'amortissement }
     ffNumPlanAvant:=T1.GetValue('I_PLANACTIF') ;
     P1.CalculTOB(T1, iDate1900);
     P1Detail := TOB.Create ('UNKNOWN',T1,-1);
     P1.SauveTOB (P1Detail );
     T1.PutValue('I_PLANACTIF',P1.NumSeq) ;
     ffNumPlanApres:= P1.NumSeq;
     P1.Free;

     T2.PutValue('I_PLANACTIF',0) ;
     P2.CalculTOB(T2, iDate1900);
     P2.SauveTOB (P1Detail );
     T2.PutValue('I_PLANACTIF',P2.NumSeq) ;
     P2.Free;
  end;

  //ajout mbo pour savoir dans quel ordre ont été saisies la prime et la sbv
  // pour pouvoir les recréer sur la fille dans le bon ordre

  if (GestPrime) and (GestSBV) then
  begin
     QLog:=OpenSQL('SELECT IL_PLANACTIFAP FROM IMMOLOG WHERE IL_IMMO="'+FFCode+
        '" AND IL_TYPEOP ="PRI" ', FALSE) ;
     PlanPRI := Qlog.FindField('IL_PLANACTIFAP').AsInteger;

     QLog:=OpenSQL('SELECT IL_PLANACTIFAP FROM IMMOLOG WHERE IL_IMMO="'+FFCode+
        '" AND IL_TYPEOP ="SBV" ', FALSE) ;
     PlanSBV := Qlog.FindField('IL_PLANACTIFAP').AsInteger;

     if PlanPRI > PlanSBV then FirstPRI := false else FirstPRI := true;

     ferme (Qlog);
  end;


  //Enregistrement du log d'acquisition de la nouvelle Immo
  Infos.TVARecuperable:= T2.GetValue('I_TVARECUPERABLE') ;
  Infos.TVARecuperee  := T2.GetValue('I_TVARECUPEREE') ;

  // Création enreg immolog de l'immo fille
  EnregLogAcquisition(T2.GetValue('I_IMMO'),T2.GetValue('I_DATEPIECEA'),T2.GetValue('I_PLANACTIF'),Infos);

  // 15/12/2005 Ajout maj il_montantecl sur immo fille
  //ExecuteSQL ('UPDATE IMMOLOG SET IL_DATEOPREELLE="'+USDateTime(StrToDate(FFDATEOPE))+
  //     '", IL_MONTANTECL="'+FloatToStr(FFvaleur)+'" WHERE IL_IMMO="'+T2.GetValue('I_IMMO')+'"');

  // FQ 17291 TGA 10/01/2006
  ExecuteSQL ('UPDATE IMMOLOG SET IL_DATEOPREELLE="'+USDateTime(StrToDate(FFDATEOPE))+
       '", IL_MONTANTECL="'+Strfpoint(FFvaleur)+'" WHERE IL_IMMO="'+T2.GetValue('I_IMMO')+'"');


  // TGA 24/11/2005 maj et fermeture remontées sinon plantage ds traitement exceptionnel

  T1.InsertOrUpdateDB ;
  T2.InsertOrUpdateDB ;
  T1.Free ;
  T2.Free ;

  //Tga 28/06/2006 Maj Immomvtd, création de l'éclatement
{$IFNDEF CMPGIS35}
  IF FFCode<>FFNouveauCode Then
    AM_MAJ_IMMOMVTD('E',FFCode,FFNouveauCode,PartEclatee);
{$ENDIF}

  // Pour corriger pb agl 179  (pb corriger pat agl 182)
  //T2.InsertOrUpdateDB ;
  //T2.Free ;
  //T1.InsertOrUpdateDB ;
  //T1.Free ;



  if bElemExcept<>0 then
    begin
    //TGA maj mise dans TraiteElementExceptionnel( T1.PutValue('I_OPECHANGEPLAN','X');
    // TGA 23/11/2005 paramêtres inversés TraiteElementExceptionnel(FFNouveauCode,FFCode,PartEclatee);
    TraiteElementExceptionnel(FFCode,FFNouveauCode,PartEclatee,FFDateOPE);
    end;

  //T1.InsertOrUpdateDB ;
  //T2.InsertOrUpdateDB ;
  //T1.Free ;
  //T2.Free ;

  //Gestion des échéances pour les immo CB et LOC
  if OkEche then
  begin
    Q1:=OpenSQL('SELECT * FROM IMMOECHE WHERE IH_IMMO="'+FFCode+'" ORDER BY IH_DATE',false);
    T1:=Tob.Create('IMMOECHE',nil,-1) ; T1.LoadDetailDB('IMMOECHE','','',Q1,true) ;
    T2:=Tob.Create('IMMOECHE',nil,-1) ;
    Ferme(Q1) ;
    for i:=0 to T1.Detail.Count-1 do
    begin
      T3:=T1.Detail[i] ;  T4:=TOB.Create('IMMOECHE',T2,-1) ;
      T4.Dupliquer(T3,false,true) ;

      IF revalorisation=True Then
      Begin
         CalculSurChampsTob(T4,'IH_MONTANT','*',New_PartEclatee,V_PGI.OkDecV) ;
         CalculSurChampsTob(T4,'IH_FRAIS','*',New_PartEclatee,V_PGI.OkDecV) ;
         CalculSurChampsTob(T3,'IH_MONTANT','*',(1-PartEclatee),V_PGI.OkDecV) ;
         CalculSurChampsTob(T3,'IH_FRAIS','*',(1-PartEclatee),V_PGI.OkDecV) ;
      End ELSE
      Begin
         CalculSurChampsTob(T4,'IH_MONTANT','*',PartEclatee,V_PGI.OkDecV) ;
         CalculSurChampsTob(T4,'IH_FRAIS','*',PartEclatee,V_PGI.OkDecV) ;
         CalculSurChampsTob(T3,'IH_MONTANT','-',T4.GetValue('IH_MONTANT'),V_PGI.OkDecV) ;
         CalculSurChampsTob(T3,'IH_FRAIS','-',T4.GetValue('IH_FRAIS'),V_PGI.OkDecV) ;
       End;
    end ;
    T1.InsertOrUpdateDB ; T2.InsertOrUpdateDB ;
    T1.Free ; T2.Free ;
  end ;

  // Création nouvel enregistrement ImmoLog
  T1:=TOB.Create('IMMOLOG',nil,-1) ;
  T1.PutValue('IL_TYPEOP','ECL') ;
  //T1.PutValue('IL_BLOCNOTE', RichToString (Il_Blocnote));
  T1.PutValue('IL_BLOCNOTE', FFBlocNote);

  // Test sur la valorisation
  IF revalorisation=True Then
    Begin
     T1.PutValue('IL_MONTANTECL',FFNEW_Valeur);
     // Sauvegarde du prorata initiale (en cas de suppression de l'éclatement)
     //T1.PutValue('IL_MONTANTDOT',PartEclatee);

     // TGA 20/01/2006
     // si PartEclatee = 0.99998 on retrouve stocké dans IL_MONTANTDOT = 1
     // ce qui provoque 1 plantage en suppression de l'éclatement
     // pour sauvegarder les décimales > 4 on sauvegarde en partie entière
     // ex: si PartEclatee = 0.99998 => on sauvegarde 99998
     x:=length(strfpoint(PartEclatee))-2;
     For i:=1 To x DO PartEclatee:=PartEclatee*10 ;

     T1.PutValue('IL_MONTANTDOT',PartEclatee);

    End
  else
    T1.PutValue('IL_MONTANTECL',FFValeur) ;

  if FFQuantiteAvEcl=1 then T1.PutValue('IL_QTEECLAT',FFQuantiteAvEcl)
                      else T1.PutValue('IL_QTEECLAT',FFQuantite) ;

  T1.PutValue('IL_CODEECLAT',FFNouveauCode) ;
  T1.PutValue('IL_IMMO',FFCode) ;
  T1.PutValue('IL_DATEOP',StrToDate(FFDATEOPE));
  T1.PutValue('IL_ORDRE',ffOrdre);
  T1.PutValue('IL_ORDRESERIE',ffOrdreS);
  T1.PutValue('IL_PLANACTIFAV',ffNumPlanAvant) ;
  T1.PutValue('IL_PLANACTIFAP',ffNumPlanApres) ;
  // Correction mbo - 05.10.06 - le code opération était faux
  //T1.PutValue('IL_LIBELLE',RechDom('TIOPEAMOR', T1.GetValue('IL_TYPEOP'), FALSE)+' '+DateToStr(StrToDate(FFDATEOPE))) ;
  T1.PutValue('IL_LIBELLE',RechDom('TIOPEAMOR', 'ECL', FALSE)+' '+DateToStr(StrToDate(FFDATEOPE))) ;
  T1.PutValue('IL_TYPEMODIF',AffecteCommentaireOperation('ECL'));
  // TGA 07.04.2006
  T1.PutValue('IL_TAUX',0);

  // On stocke sur l'immolog de la mère le montant de la prime et de la subvention avant éclatement
  // pour pouvoir les restituer en cas d'annulation de l'éclatement
  if (GestPrime) then
  begin
    T1.PutValue('IL_REVISIONDOTECO', SAVmntPRI);
    T1.PutValue('IL_REVISIONREPECO', SAVreprisePRI);
  end;
  if (GestSBV) then
  begin
    T1.PutValue('IL_REVISIONDOTFISC', SAVmntSBV);
    T1.PutValue('IL_REVISIONREPFISC', SAVrepriseSBV);
  end;

  T1.InsertOrUpdateDB ;
  T1.Free ;

 // si présence d'une prime sur la mère, nécessité de créer un enreg dans immolog avec opération PRI
 // si présence d'une SBV sur la mère, nécessité de créer un enreg dans immolog avec opération SBV
 // si les deux : il faut les recréer dans l'ordre de la mère
  if (GestPrime) and (GestSBV) then
  begin
     if (FirstPri) then
     begin
        EnregLogPri(FFCode,FFNouveauCOde,NewPrime, FFDateope, TabPriFille,TabSbvFille, NumPlanFille, false);
        EnregLogSBV(FFCode,FFNouveauCOde,NewSBV, FFDateope, TabPriFille, TabSbvFille, NumPlanFille, true);
     end else
     begin
        EnregLogSBV(FFCode,FFNouveauCOde,NewSBV, FFDateope, TabPriFille, TabSbvFille,NumPlanFille, false);
        EnregLogPri(FFCode,FFNouveauCOde,NewPrime, FFDateope, TabPriFille,TabSbvFille, NumPlanFille, true);
     end;
  end else
  begin
     if (GestSBV) then
        EnregLogSBV(FFCode,FFNouveauCOde,NewSBV, FFDateope, TabSbvFille, TabSbvFille,NumPlanFille, false);

     if (GestPrime) then
        EnregLogPri(FFCode,FFNouveauCOde,NewPrime, FFDateope, TabPriFille, TabSbvFille,NumPlanFille, false);
  end;

end;


procedure CalculSurChampsTob(T1: TOB;LeChamp,Operande: string;X: Variant;ValArrondi:integer) ;
var w: double ; j:integer ; Va: variant ;
begin
  case Operande[1] of
    '+': w:=T1.GetValue(LeChamp)+VarAsType(X,vardouble) ;
    '-': w:=T1.GetValue(LeChamp)-VarAsType(X,vardouble) ;
    '*': w:=T1.GetValue(LeChamp)*VarAsType(X,vardouble) ;
    '/': w:=T1.GetValue(LeChamp)/VarAsType(X,vardouble) ;
    else exit ;
    end ;
  j:=T1.GetNumChamp(LeChamp) ;
  //if (varType(T1.Valeurs[j]) in [varByte,varSmallint,varInteger]) then Va:=ValeurI(FloatToStr(w)) //XVI 28/02/2007 S/FQ Nouveaux types des variants pour D7 (préconisation CBP)
  if (varType(T1.Valeurs[j]) in [varSmallint, varInteger, varByte, varShortInt, varWord, varLongWord, varInt64]) then Va:=ValeurI(FloatToStr(w))
  else  if (varType(T1.Valeurs[j]) in [varSingle,varDouble,varCurrency]) then Va:=Valeur(FloatToStr(Arrondi(w,ValArrondi))) ;
  T1.PutValue(LeChamp,Va) ;
end ;


procedure TraiteElementExceptionnel(CodeOrig,CodeEclate : string;PartEclatee : double;FFdateope:String);
var
  ListeLog : TList;
  EnrListeLog : PLogChPlan;
  sTypeExc : string;
  dMontantDot,dMontantExc : double;
  Ind,iPlanAvant,iPlanApres : integer;
  QImmo,QLog : TQuery;
  PlanEclate : TPlanAmort ;
begin

  // Mise à jour immo d'origine
  QLog:=OpenSQL('SELECT * FROM IMMOLOG WHERE IL_IMMO="'+CodeOrig+'" AND (IL_TYPEOP="ELE" OR IL_TYPEOP="CDM") ORDER BY IL_DATEOP ASC', FALSE) ;
  ListeLog := TList.Create;
  while not QLog.EOF do
  begin
    QLog.Edit;
    dMontantDot := QLog.FindField('IL_MONTANTDOT').AsFloat;
    QLog.FindField('IL_MONTANTDOT').AsFloat:=(1-PartEclatee)*dMontantDot;
    QLog.FindField('IL_TYPEOP').AsString:='EEC';
    QLog.FindField('IL_LIBELLE').AsString:=RechDom('TIOPEAMOR', 'EEC', FALSE);
    QLog.Post;
    New(EnrListeLog);
    TraiteEnregLogChPlan(QLog,EnrListeLog^);
    EnrListeLog^.TypeExc:=QLog.FindField('IL_TYPEDOT').AsString;
    EnrListeLog^.MontantExc:=dMontantDot*PartEclatee;
    ListeLog.Add(EnrListeLog);
    QLog.Next;
  end;
  Ferme(QLog);
  sTypeExc:=''; dMontantExc:=0.00;
  for Ind := 0 to (ListeLog.Count - 1) do
  begin
    QLog:=OpenSQL('SELECT * FROM IMMOLOG WHERE IL_IMMO="'+W_W+'"', FALSE) ;
    QLog.Insert ;
    EnrListeLog := ListeLog.Items[Ind];
    PlanEclate:=TPlanAmort.Create(true) ;// := CreePlan (True);
    try
      QImmo:=OpenSQL('SELECT * FROM IMMO WHERE I_IMMO="'+CodeEclate+'"', false) ;
      QImmo.Edit;
      iPlanAvant := QImmo.FindField('I_PLANACTIF').AsInteger;
      QImmo.FindField('I_TYPEEXC').AsString := EnrListeLog^.TypeExc;
      QImmo.FindField('I_MONTANTEXC').AsFloat:=EnrListeLog^.MontantExc;
      // Ajout TGA 24/11/2005
      QImmo.FindField('I_OPECHANGEPLAN').AsString:='X';
      PlanEclate.typeope :='ELE';

      PlanEclate.Charge(QImmo);
      // TGA 23/11/2005 PlanEclate.Calcul(QImmo, StrToDate(DATEOPE.Text));
      PlanEclate.Calcul(QImmo, StrToDate(FFDATEOPE));
      PlanEclate.Sauve;
      QImmo.FindField('I_PLANACTIF').AsInteger := PlanEclate.NumSeq;
      iPlanApres := QImmo.FindField('I_PLANACTIF').AsInteger;
    finally
      PlanEclate.free ; //Detruit;
    end ;
    QImmo.Post;
    Ferme(QImmo);

    QLog.FindField('IL_IMMO').AsString:=CodeEclate;
    QLog.FindField('IL_LIBELLE').AsString:=RechDom('TIOPEAMOR', EnrListeLog^.TypeOpe, FALSE);
    QLog.FindField('IL_TYPEMODIF').AsString:=AffecteCommentaireOperation('CHP');
    QLog.FindField('IL_ORDRE').AsInteger := EnrListeLog^.Ordre;
    QLog.FindField('IL_ORDRESERIE').AsInteger := EnrListeLog^.OrdreSerie;
    QLog.FindField('IL_TYPEOP').AsString := EnrListeLog^.TypeOpe;
    QLog.FindField('IL_DATEOP').AsDateTime := EnrListeLog^.DateOpe;
    QLog.FindField('IL_PLANACTIFAV').AsInteger := iPlanAvant;//EnrListeLog^.PlanActifAv;
    QLog.FindField('IL_PLANACTIFAP').AsInteger := iPlanApres;//EnrListeLog^.PlanActifAp;
    QLog.FindField('IL_TYPEEXC').AsString := sTypeExc;//???EnrListeLog^.TypeExc;
    QLog.FindField('IL_MONTANTEXC').AsFloat := dMontantExc;//???EnrListeLog^.MontantExc;
    QLog.FindField('IL_TYPEDOT').AsString := EnrListeLog^.TypeExc;
    QLog.FindField('IL_MONTANTDOT').AsFloat := EnrListeLog^.MontantExc;
    QLog.FindField('IL_REVISIONECO').AsFloat := EnrListeLog^.RevisionEco;
    QLog.FindField('IL_REVISIONFISC').AsFloat := EnrListeLog^.RevisionFisc;
    QLog.FindField('IL_DUREEECO').AsInteger := EnrListeLog^.DureeEco;
    QLog.FindField('IL_DUREEFISC').AsInteger :=EnrListeLog^.DureeFisc;
    QLog.FindField('IL_METHODEECO').AsString :=EnrListeLog^.MethodeEco;
    QLog.FindField('IL_METHODEFISC').AsString := EnrListeLog^.MethodeFisc;
    QLog.Post;
    sTypeExc := QLog.FindField('IL_TYPEEXC').AsString;
    dMontantExc := QLog.FindField('IL_MONTANTEXC').AsFloat;
    Ferme(QLog);
    Dispose(EnrListeLog);
  end;
  ListeLog.Free
end;


procedure EclatePlan ( PIni, PDest : TAmPlan; dProrata : double; ori_dProrata:double; bElemExcept:double;
                       var TabPriFille, TabSbvFille:array of double);
var i, nbAnnuite : integer;
    CumulBE, CumulBF : double;
    CumulPRI, CumulSBV : double;

    EcartBE, EcartBF : double;
    EcartPRI, EcartSBV : double;

    BaseEco, BaseFisc : double;
    BasePRI, BaseSBV : double;

    TDest, TIni : TOB;
    dDotationEco, dDotationFisc,ini_dDotationEco : double;
    DotationPRI, DotationSBV : double;
    CumulBEmere, CumulBFmere, EcartBEmere, EcartBFmere : double;  // ajout mbo fq 19851
begin
  PDest.CopieDetail ( PIni );
  CumulBE := PDest.RepriseEco;
  CumulBF := PDest.RepriseFisc;

  // ajout mbo fq 19851
  CumulBEmere := PIni.RepriseEco;
  CumulBFmere := PIni.RepriseFisc;


  CumulPRI := PDest.RepriseUO;

  DotationPRI := 0;
  DotationSBV := 0;

  if (PDest.SBV = true) and (PDest.ReprisePrem = true) then
     CumulSBV := 0  // les antérieurs sont sur la première dotation
  else
     CumulSBV := PDest.RepriseSBV;

  BaseEco := PDest.BaseEco - CumulBE;   // sert pour la base éco début d'exercice
  BaseFisc := PDest.BaseFisc - CumulBF; // sert pour la base fiscale début d'exercice
  baseSBV  := PDest.BaseSBV ; //- CumulSbv ;            // pas de zone base début exercice pour la subvention
  BasePRI := (PDest.BasePRI/2); //- CumulPRI;         // pas de zone base début exercice pour la prime
  nbAnnuite := PIni.Detail.Count - 1;

  for i := 0 to PIni.Detail.Count - 1 do
  begin
    TDest := PDest.Detail[i];
    TIni  := PIni.Detail[i];

    { Dotation économique }

    // TGA 24/11/2005
    // on force pour la nouvelle immo la séquence à 1 et le type à CRE
    TDest.PutValue ('IA_CHANGEAMOR', 'CRE');

    // mbo 05.10.06 on crée le plan de  l'immo d'origine avec type ECL
    TIni.PutValue ('IA_CHANGEAMOR', 'ECL');

    // Calcul dotation pour immo de destination
    IF bElemExcept<>0  Then
       Begin
        IF i=0 Then
          Begin
           // première annuité on retire l'exceptionnel pour calculer l'annuité
           dDotationEco := TIni.GetValue('IA_MONTANTECO');
           dDotationEco := Arrondi(dDotationEco-bElemExcept,V_PGI.OkDecV);
           dDotationEco := Arrondi(dDotationEco*dProrata,V_PGI.OkDecV);
          end
        Else
          dDotationEco := Arrondi(TIni.GetValue('IA_MONTANTECO')*dProrata,V_PGI.OkDecV);
       end
    else
      dDotationEco := Arrondi(TIni.GetValue('IA_MONTANTECO')*dProrata,V_PGI.OkDecV);

    TDest.PutValue ('IA_MONTANTECO', dDotationEco);

    // Calcul dotation pour immo d'origine
    if ( i < nbAnnuite ) then
      Begin
       // tga 25/11/2005 test si positif car peut être négatif si exceptionnel
       ini_dDotationEco := Arrondi(TIni.GetValue('IA_MONTANTECO')*ori_dProrata,V_PGI.OkDecV);
       IF (TIni.GetValue('IA_MONTANTECO')-ini_dDotationEco) >=0 THEN
         TIni.PutValue ('IA_MONTANTECO', TIni.GetValue('IA_MONTANTECO')-ini_dDotationEco)
       else
         TIni.PutValue ('IA_MONTANTECO',0);
      end;

    { Base début exercice  }
    TDest.PutValue ('IA_BASEDEBEXOECO', BaseEco);
    TIni.PutValue ('IA_BASEDEBEXOECO', Arrondi(TIni.GetValue('IA_BASEDEBEXOECO')*(1-ori_dProrata),V_PGI.OkDecV));
    //TIni.PutValue ('IA_BASEDEBEXOECO', TIni.GetValue('IA_BASEDEBEXOECO')-BaseEco);

    if PIni.Fiscal then
    begin
      { Dotation fiscale }
      dDotationFisc := Arrondi(TIni.GetValue('IA_MONTANTFISCAL')*dProrata,V_PGI.OkDecV);
      TDest.PutValue ('IA_MONTANTFISCAL', dDotationFisc);

      // mbo fq 19851 if ( i < nbAnnuite ) then
        //TIni.PutValue ('IA_MONTANTFISCAL', TIni.GetValue('IA_MONTANTFISCAL')-dDotationFisc);
        TIni.PutValue ('IA_MONTANTFISCAL', Arrondi(TIni.GetValue('IA_MONTANTFISCAL')*(1-ori_dProrata),V_PGI.OkDecV));

      { Dérogatoire }
      PIni.CalculDerogatoire;
      PDest.CalculDerogatoire;

      { Base début exercice }
      TDest.PutValue ('IA_BASEDEBEXOFISC', BaseFisc);
      //TIni.PutValue ('IA_BASEDEBEXOFISC', TIni.GetValue('IA_BASEDEBEXOFISC')-BaseFisc);
      TIni.PutValue ('IA_BASEDEBEXOFISC',Arrondi(TIni.GetValue('IA_BASEDEBEXOFISC')*(1-ori_dProrata),V_PGI.OkDecV));

    end;

    //ajout pour prime d'équipement
    if PIni.PRI then
    begin
      { Dotation PRIME }
      DotationPRI := Arrondi(TIni.GetValue('IA_MONTANTSBV')*dProrata,V_PGI.OkDecV);
      TabPriFille[i] := DotationPri;
      // sur l'enreg ACQ de la fille on initialise la dotation PRI car on va recréer les enreg PRI après
      TDest.PutValue ('IA_MONTANTSBV', 0);

      if ( i < nbAnnuite ) then
        TIni.PutValue ('IA_MONTANTSBV', Arrondi(TIni.GetValue('IA_MONTANTSBV')*(1-ori_dProrata),V_PGI.OkDecV));
    end;

    if PIni.SBV then
    begin
      { Dotation SUBVENTION }
      DotationSBV := Arrondi(TIni.GetValue('IA_MONTANTDPI')*dProrata,V_PGI.OkDecV);
      TabSbvFille[i] := DotationSBV;
      // sur l'enreg ACQ de la fille on initialise la dotation PRI car on va recréer les enreg SBV après
      TDest.PutValue ('IA_MONTANTDPI', 0);

      if ( i < nbAnnuite ) then
        TIni.PutValue ('IA_MONTANTDPI', Arrondi(TIni.GetValue('IA_MONTANTDPI')*(1-ori_dProrata),V_PGI.OkDecV));
    end;

    { Réintégration }
    TDest.PutValue ('IA_REINTEGRATION', Arrondi(TIni.GetValue('IA_REINTEGRATION')*dProrata,V_PGI.OkDecV));
    //TIni.PutValue ('IA_REINTEGRATION', TIni.GetValue('IA_REINTEGRATION')-TDest.GetValue('IA_REINTEGRATION'));
    TIni.PutValue ('IA_REINTEGRATION', Arrondi(TIni.GetValue('IA_REINTEGRATION')*(1-ori_dProrata),V_PGI.OkDecV));

    { Quote-part }
    TDest.PutValue ('IA_QUOTEPART', Arrondi(TIni.GetValue('IA_QUOTEPART')*dProrata,V_PGI.OkDecV));
    //TIni.PutValue ('IA_QUOTEPART', TIni.GetValue('IA_QUOTEPART')-TDest.GetValue('IA_QUOTEPART'));
    TIni.PutValue ('IA_QUOTEPART', Arrondi(TIni.GetValue('IA_QUOTEPART')*(1-ori_dProrata),V_PGI.OkDecV));

    { Cumuls }
    CumulBE := CumulBE + TDest.GetValue ('IA_MONTANTECO');
    CumulBF := CumulBF + TDest.GetValue ('IA_MONTANTFISCAL');

    // ajout pour fq 19851 mbo 19.03.2007
    CumulBEmere := CumulBEmere + TIni.GetValue ('IA_MONTANTECO');
    CumulBFmere := CumulBFmere + TIni.GetValue ('IA_MONTANTFISCAL');

    CumulPRI := CumulPRI + DotationPri;
    CumulSBV := CumulSBV + DotationSbv;

    BaseEco := BaseEco - TDest.GetValue ('IA_MONTANTECO');
    BaseFisc := BaseFisc - TDest.GetValue ('IA_MONTANTFISCAL');
    { On ne traite pas les champs de sortie car éclatement impossible après sortie }
  end;

  if ( PIni.Detail.Count > 0 ) then
  begin
    { Régularisation sur la dernière annuité pour une VNC à 0 }
    EcartBE := CumulBE - PDest.BaseEco;
    EcartBF := 0;

    EcartBEmere := CumulBEmere-PIni.BaseEco;
    // CONSEIL COMPIL EcartBFmere := 0;

    if PIni.Fiscal then
      EcartBF := CumulBF - PDest.BaseFisc;
      EcartBFmere := CumulBFmere - PIni.BaseFisc; // ajout fq 19851

    // dernière annuité pour nouvelle immo
    PDest.Detail[nbAnnuite].PutValue ('IA_MONTANTECO', PDest.Detail[nbAnnuite].GetValue ('IA_MONTANTECO')-EcartBE);

    // dernière annuité pour immo origine
    // tga 25/11/2005 test si positif car peut être négatif si exceptionnel
    ini_dDotationEco := Arrondi(PIni.Detail[nbAnnuite].GetValue('IA_MONTANTECO')*ori_dProrata,V_PGI.OkDecV);

    IF (PIni.Detail[nbAnnuite].GetValue('IA_MONTANTECO')-ini_dDotationEco) >=0 THEN
       PIni.Detail[nbAnnuite].PutValue('IA_MONTANTECO', PIni.Detail[nbAnnuite].GetValue('IA_MONTANTECO')-ini_dDotationEco)
    else
      PIni.Detail[nbAnnuite].PutValue('IA_MONTANTECO', 0);

    if PIni.Fiscal then
    begin
      { correction mbo 19.03.07 FQ 19851
      //PDest.Detail[nbAnnuite].PutValue ('IA_MONTANTFISCAL', PIni.Detail[nbAnnuite].GetValue('IA_MONTANTFISCAL')-EcartBF);
      }
      PDest.Detail[nbAnnuite].PutValue ('IA_MONTANTFISCAL', PDest.Detail[nbAnnuite].GetValue('IA_MONTANTFISCAL')-EcartBF);

      { correction mbo 19.03.07 FQ 19851
      //PIni.Detail[nbAnnuite].PutValue ('IA_MONTANTFISCAL', Arrondi(PIni.Detail[nbAnnuite].GetValue('IA_MONTANTFISCAL')*(1-ori_dProrata),V_PGI.OkDecV));
      }
      PIni.Detail[nbAnnuite].PutValue ('IA_MONTANTFISCAL', PIni.Detail[nbAnnuite].GetValue('IA_MONTANTFISCAL')-EcartBFmere);

      { ajout mbo 19.03.07 FQ 19851 - Dérogatoire : il faut le recalculer car
        les dotations fiscales ont pu changer de quelques pouyèmes dus aux arrondis sur dernière dotations}
      PIni.CalculDerogatoire;
      PDest.CalculDerogatoire;

      //PIni.Detail[nbAnnuite].PutValue ('IA_MONTANTFISCAL', PIni.Detail[nbAnnuite].GetValue('IA_MONTANTFISCAL')-PDest.Detail[nbAnnuite].GetValue('IA_MONTANTFISCAL'));
    end;

    if PIni.PRI then
    begin
      EcartPRI := CumulPRI - BasePRI;
      // sur la fille on initialise la dotation PRI car on recrée après les enreg PRI
      TabPriFille[nbAnnuite]:=  (TabPriFille[nbAnnuite]-EcartPRI);
      PDest.Detail[nbAnnuite].PutValue ('IA_MONTANTSBV', 0);
      PIni.Detail[nbAnnuite].PutValue ('IA_MONTANTSBV', Arrondi(PIni.Detail[nbAnnuite].GetValue('IA_MONTANTSBV')*(1-ori_dProrata),V_PGI.OkDecV));
    end;

    if PIni.SBV then
    begin
      EcartSBV := CumulSBV - BaseSBV;
      // Sur la fille on initialise la dotation SBV car on recrée après les enreg SBV
      TabSbvFille[nbAnnuite] := TabSbvFille[nbAnnuite]-EcartSBV;
      PDest.Detail[nbAnnuite].PutValue ('IA_MONTANTDPI', 0);
      PIni.Detail[nbAnnuite].PutValue ('IA_MONTANTDPI', Arrondi(PIni.Detail[nbAnnuite].GetValue('IA_MONTANTDPI')*(1-ori_dProrata),V_PGI.OkDecV));
    end;

  end;

end;


{***********A.G.L.***********************************************
Auteur  ...... : Maryse Boudin
Créé le ...... : 19/03/2007
Modifié le ... :   /  /
Description .. : FQ 17512 -
Suite ........ : Nouveau calcul du dérogatoire qui tient compte de la
Suite ........ : gestion fiscale
Mots clefs ... :
*****************************************************************}
procedure tamplan.CalculDerogatoire;
Var
  DotDerog, DotEco, DotFisc: double;
  i : integer;
  //Dérogatoire
  CE, CF, AE, AF, EC, EA : double;
  DotReint : double;
begin

  CE := fRepriseEco;
  CF := fRepriseFisc;
  AE := CE;
  AF := CF;

  for i := 0 to Detail.Count - 1 do
  begin

    DotEco := Detail[i].GetValue('IA_MONTANTECO');
    DotFisc := Detail[i].GetValue('IA_MONTANTFISCAL');
    DotDerog := 0.00;
    DotReint := 0.00;

    // Dérogatoire
    CE := arrondi((CE + DotEco),V_PGI.OkDecV);
    CF := arrondi((CF + DotFisc),V_PGI.OkDecV);

    EC := arrondi((CF-CE),V_PGI.OkDecV); //écart sur (cumul antérieur + dotation de l'exercice)
    EA := arrondi((AF-AE),V_PGI.OkDecV); //écart sur cumuls antérieurs

    // le plan d'amortissement éco n'est pas variable
    if not (fGestionFiscale) then
    begin
       // on fait les calculs comme avant
       if (EC >= 0) then     // écart sur antérieurs
       begin
          if EA < 0 then EA := 0;
          if EC > EA then DotDerog := arrondi((EC-EA),V_PGI.OkDecV)
          else if EA > EC then DotDerog := arrondi((EC-EA),V_PGI.OkDecV);
       end else
           if (EC < 0) then if (EA>0) then DotDerog := (-1)*EA;

    end else  // traitement de la gestion fiscale
    begin
       // ajout chantier fiscal : on génére une réintégration ou une déduction
       // qd c'est positif = réintégration - qd c'est négatif = déduction
       DotReint := arrondi((DotEco - DotFisc),V_PGI.OkDecV);
       DotDerog := 0;
    end;

    AE := CE;
    AF := CF;

    //Dotation Dérogatoire
    Detail[i].PutValue('IA_MONTANTDEROG',DotDerog);
      // réintégration extra comptable
    Detail[i].PutValue('IA_NONDEDUCT',DotReint);
  end;
end;

//================================================================
{ANCIEN CALCUL DU DEROGATOIRE  fq 17512
procedure tamplan.CalculDerogatoire;
var i : integer;
    CE, CF, AE, AF, EC, EA : double;
    DotDerog : double;
begin
  CE := fRepriseEco;
  CF := fRepriseFisc;
  AE := CE;
  AF := CF;

  for i := 0 to Detail.Count - 1 do
  begin
    CE := CE + Detail[i].GetValue('IA_MONTANTECO');
    CF := CF + Detail[i].GetValue('IA_MONTANTFISCAL');
    EC := CF-CE;
    EA := AF-AE;
    DotDerog := 0;
    if (EC >= 0) then
    begin
      if EA < 0 then DotDerog := 0
      else if EC > EA then DotDerog := EC-EA
      else if EA > EC then DotDerog := EC-EA;
    end else
      if (EC < 0) then if (EA>0) then DotDerog := (-1)*EA;
    AE := CE;
    AF := CF;
    Detail[i].PutValue('IA_MONTANTDEROG',DotDerog);
  end;
end;
}
//=========================================================================

procedure tamplan.CopieDetail(PSource: tamplan);
var i : integer;
    T : TOB;
begin
  ClearDetail;
  for i:=0 to PSource.Detail.Count - 1 do
  begin
    T := TOB.Create ('IMMOAMOR',Self,-1);
    T.Dupliquer ( PSource.Detail[i],False,True);
    T.PutValue('IA_IMMO',GetValue('CODEIMMO'));
  end;
end;

constructor tamplan.Create( T : TOB );
var i : integer;
begin
  inherited Create ('',nil,-1);
  fRepriseEco := T.GetValue('I_REPRISEECO');
  fRepriseFisc := T.GetValue('I_REPRISEFISCAL');
  fBaseEco := T.GetValue('I_BASEECO');
  fBaseFisc := T.GetValue('I_BASEFISC');

  //ajout mbo FQ 17512 ajout pour dérogatoire/réintégration
  if T.GetValue('I_NONDED') = '-' then
     fGestionFiscale := false
  else
     fGestionFiscale := true;

  fRepriseDR := T.GetValue('I_REPRISEDR');  // antérieur dérogatoire
  fRepriseFEC := T.GetValue('I_REPRISEFEC'); // antérieur réintégration fiscale

  // ajout mbo pour prime
  fRepriseUO := T.GetValue('I_REPRISEUO');
  fRepriseSBV := T.GetValue('I_CORRECTIONVR');

  fBasePRI := T.GetValue('I_SBVPRI');
  fBaseSBV := T.GetValue('I_SBVMT');

  AddChampSupValeur ('CODEIMMO',T.GetValue('I_IMMO'));
  fPlanActif := T.GetValue('I_PLANACTIF');
  fbFiscal := (T.GetValue('I_METHODEFISC')<>'');
  fPRI := (T.GetValue('I_SBVPRI')<>0);
  fSBV := (T.GetValue('I_SBVMT')<>0);
  fReprisePrem := (T.GetValue('I_DPIEC') = 'X');

  { Chargement du plan d'amortissement }
  LoadDetailFromSQL ('SELECT * FROM IMMOAMOR WHERE IA_IMMO="'+T.GetValue('I_IMMO')+'" AND IA_NUMEROSEQ='+T.GetString('I_PLANACTIF'));
  for i:=0 to Detail.Count - 1 do
    Detail[i].VirtuelleToReelle ('IMMOAMOR');
end;

destructor tamplan.Destroy;
begin
  inherited;
end;

function tamplan.Enregistre: boolean;
var i : integer;
begin
  for i := 0 to Detail.Count - 1 do
    Detail[i].PutValue('IA_NUMEROSEQ', fPlanActif+1);

  Result := InsertDB(nil);
  if Result then Inc (fPlanActif);
end;

function tamplan.Enregistre_bis: boolean;
var i : integer;
begin
  for i := 0 to Detail.Count - 1 do
    Detail[i].PutValue('IA_NUMEROSEQ',1);
  Result := InsertDB(nil);
  if Result then Inc (fPlanActif);
end;

function tamplan.Enregistre_PRISVB(NumPlanFille:integer) : boolean;
var i : integer;
begin
  for i := 0 to Detail.Count - 1 do
    Detail[i].PutValue('IA_NUMEROSEQ', NumPlanFille);

  Result := InsertDB(nil);
end;


{***********A.G.L.***********************************************
Auteur  ...... : Maryse Boudin
Créé le ...... : 20/10/2006
Modifié le ... :   /  /
Description .. : On crée pour la fille les enreg PRI dans IMMOAMOR
Suite ........ : avec le code PRI et l'enreg dans IMMOLOG
Mots clefs ... :
*****************************************************************}
Procedure EnregLogPri(FFCode:string; NewCode:string; NewPrime:double; FFDateope:string;
                      TabPriFille, TabSbvFille:array of double; var NumPlanFille:integer; trtSbv :boolean);
var
   Ordre : integer;
   //QPlan : TQuery;
   Qlog : TQuery;
   TLog2 : TOB;
   //PlanNew : TPlanAmort;
   fPlanActifAv: integer;
   fPlanActifAp: integer;
   PlanDest : TAmPlan;
   T1, TDest : Tob;
   i : integer;
begin

 { QPlan:=OpenSQL('SELECT * FROM IMMO WHERE I_IMMO="'+NewCode+'"', FALSE) ;
  PlanNew:=TPlanAmort.Create(true) ;
  QPlan.Edit;

  try
     PlanNew.Charge(QPlan);
     PlanNew.Recupere(NewCode,QPlan.FindField('I_PLANACTIF').AsString);
     fPlanActifAv := PlanNew.NumSeq;
     PlanNew.SetTypeOpe('PRI');
     PlanNew.Calcul(Qplan, idate1900);
     PlanNew.Sauve;

     // Attribuer à l'immo le n° du nouveau plan d'amortissement
     Qplan.FindField('I_PLANACTIF').AsInteger := PlanNew.NumSeq;

     Qplan.Post;
     fPlanActifAp := PlanNew.NumSeq;

   finally
      PlanNew.Free;
      Ferme(Qplan);
   end;
 }

  T1:=TOB.Create('IMMO',nil,-1) ;
  T1.SelectDB('"'+NewCode+'"',nil) ;
  PlanDest := TAmPlan.Create ( T1 );
  fPlanActifAv := NumPlanFille;
  fPlanActifAp := NumPlanFille +1;
  NumPlanFille := NumPlanFille + 1;

  for i := 0 to PlanDest.Detail.Count - 1 do
  begin
    TDest := PlanDest.Detail[i];

    { Dotation économique }

    // TGA 24/11/2005
    // on force pour la nouvelle immo la séquence à 1 et le type à CRE
    TDest.PutValue ('IA_CHANGEAMOR', 'PRI');
    TDest.PutValue ('IA_MONTANTSBV', TabPriFille[i]);
    if trtSbv then
       TDest.PutValue ('IA_MONTANTDPI', TabSbvFille[i]);
  end;
  PlanDest.Enregistre_PRISVB(NumPlanFille);
  T1.PutValue('I_PLANACTIF',NumPlanFille) ;
  T1.InsertOrUpdateDB ;
  T1.Free ;

  PlanDest.Free;

   // Mise à jour de IMMOLOG

   Ordre := TrouveNumeroOrdreLogSuivant(NewCode);

   QLog := OpenSQL('SELECT * FROM IMMOLOG WHERE IL_IMMO="'+FFCode+'" AND IL_TYPEOP="PRI"', FALSE) ;
   TLog2 := TOB.Create ('IMMOLOG', nil, -1);
   try
     TLog2.SelectDB('',QLog,false);
     TLog2.PutValue('IL_IMMO',NewCode);
     TLog2.PutValue('IL_ORDRE', Ordre);
     TLog2.PutValue('IL_MONTANTEXC', NewPrime);
     TLog2.PutValue('IL_PLANACTIFAV', fPlanActifAv);
     TLog2.PutValue('IL_PLANACTIFAP', fPlanActifAp);
     TLog2.PutValue('IL_DATEOP',StrToDate(FFDATEOPE));
     TLog2.InsertDB(nil);
   finally
     TLog2.Free;
     Ferme(QLog);
   end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Maryse Boudin
Créé le ...... : 20/10/2006
Modifié le ... :   /  /
Description .. : On crée pour la fille les enreg SBV dans IMMOAMOR
Suite ........ : avec l'opération SBV et l'enreg dans immolog
Mots clefs ... :
*****************************************************************}
Procedure EnregLogsbv(FFCode:string; NewCode:string; NewSBV:double; FFDateope:string ;
                      TabPriFille,TabSbvFille : array of double; var NumPlanFille:integer; trtPri:boolean);
var
   Ordre : integer;
   //QPlan : TQuery;
   Qlog : TQuery;
   TLog2 : TOB;
   //PlanNew : TPlanAmort;
   fPlanActifAv: integer;
   fPlanActifAp: integer;
   PlanDest : TAmPlan;
   T1, TDest : Tob;
   i : integer;

begin
{
  QPlan:=OpenSQL('SELECT * FROM IMMO WHERE I_IMMO="'+NewCode+'"', FALSE) ;
  PlanNew:=TPlanAmort.Create(true) ;
  QPlan.Edit;

  try
     PlanNew.Charge(QPlan);
     PlanNew.Recupere(NewCode,QPlan.FindField('I_PLANACTIF').AsString);
     fPlanActifAv := PlanNew.NumSeq;
     PlanNew.SetTypeOpe('SBV');
     PlanNew.Calcul(Qplan, idate1900);
     PlanNew.Sauve;

     // Attribuer à l'immo le n° du nouveau plan d'amortissement
     Qplan.FindField('I_PLANACTIF').AsInteger := PlanNew.NumSeq;

     Qplan.Post;
     fPlanActifAp := PlanNew.NumSeq;

   finally
      PlanNew.Free;
      Ferme(Qplan);
   end;
}

  T1:=TOB.Create('IMMO',nil,-1) ;
  T1.SelectDB('"'+NewCode+'"',nil) ;
  PlanDest := TAmPlan.Create ( T1 );
  fPlanActifAv := NumPlanFille;
  fPlanActifAp := NumPlanFille + 1;
  NumPlanFille:=  NumPlanFille + 1;

  for i := 0 to PlanDest.Detail.Count - 1 do
  begin
    TDest := PlanDest.Detail[i];

    { Dotation économique }

    // TGA 24/11/2005
    // on force pour la nouvelle immo la séquence à 1 et le type à CRE
    TDest.PutValue ('IA_CHANGEAMOR', 'SBV');
    TDest.PutValue ('IA_MONTANTDPI', TabSbvFille[i]);
    if trtPri then
       TDest.PutValue ('IA_MONTANTSBV', TabPriFille[i]);
  end;
  PlanDest.Enregistre_PRISVB(NumPlanFille);
  T1.PutValue('I_PLANACTIF',NumPlanFille) ;
  T1.InsertOrUpdateDB ;
  T1.Free ;

  PlanDest.Free;

   // Mise à jour de IMMOLOG

   Ordre := TrouveNumeroOrdreLogSuivant(NewCode);

   QLog := OpenSQL('SELECT * FROM IMMOLOG WHERE IL_IMMO="'+FFCode+'" AND IL_TYPEOP="SBV"', FALSE) ;
   TLog2 := TOB.Create ('IMMOLOG', nil, -1);
   try
     TLog2.SelectDB('',QLog,false);
     TLog2.PutValue('IL_IMMO',NewCode);
     TLog2.PutValue('IL_ORDRE', Ordre);
     TLog2.PutValue('IL_MONTANTEXC', NewSBV);
     TLog2.PutValue('IL_PLANACTIFAV', fPlanActifAv);
     TLog2.PutValue('IL_PLANACTIFAP', fPlanActifAp);
     TLog2.PutValue('IL_DATEOP',StrToDate(FFDATEOPE));
     TLog2.InsertDB(nil);
   finally
     TLog2.Free;
     Ferme(QLog);
   end;
end;

end.
