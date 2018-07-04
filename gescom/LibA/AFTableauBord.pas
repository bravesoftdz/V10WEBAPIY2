unit AFTableauBord;

interface
uses {$IFDEF VER150} variants,{$ENDIF} SysUtils,ComCtrls,Controls,Classes,StdCtrls,Hent1,UTOB,HCtrls,EntGC,DicoBTP,
     FactUtil,Affaireutil,Ent1,Formule,HStatus,ActiviteUtil,
{$IFDEF BTP}
	 CalcOleGenericBTP,
{$ENDIF}

{$IFDEF EAGLCLIENT}
{$ELSE}
      {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}                   
     CalcOLEAffaire,ParamSoc,UtilRessource,uafo_ressource, UtilArticle;

Const PrefixeTB = 'ATB_';
      MaxCritTB = 20; MaxCorrespChamps = 70;  // mcd 18/11/02

      // Emplacement dans le dico de chaque table / colonnes du tableau
      nTB = 1; nACT = 2; nGL = 3; nGLPR = 4; nCutOff = 5; nBud = 6; nPlanning = 7;

Type    T_NiveauTB    = (tniTous,tniTiers,tniAffaire,tniTypeArt,tniArt,tniTypeArtRes,tniArtRes,tniRes,tniTypeRes,tniFonction,tniFamilleArt1,tniFamilleArt2,tniFamilleArt3,tniStatArt1,tniStatArt2,tniStatArt3) ;
        T_TypeTB      = (ttbTiers,ttbAffaire,ttbRessource,ttbArticle);
        T_TypeDateTB  = (tdaDetail,tdaSemaine,tdaFolio,tdaMois,tdaCumul);
        T_TableTB     = (ttaTableTB, ttaTBTable);
        T_TypePRFA    = (tligne,tactivite,Tboni,TCutOff,TfactEclat,Tbud,TPlanning);
        T_MtActivite  = (ActPR,ActPV,ActPRPV);
        T_TypeEnr     = (Tan,Tnor,Tmois);
                       
Type    TCritereTBAffaire = RECORD
            Champ, BorneDebut, BorneFin,ChampTable,Table : string;
            bTypeDate,bReel,bMultiVal,WhereTiers,WhereTiersCompl,WhereAffaire,WhereAct,
            WhereBud,WhereCutOff,WhereLigne,WhereArt,WhereRes, WherePlanning : Boolean;
            End;
Type    T_ParamTb =Record    // pour appel fct alimentation table AftableuaBord avec Critères
            CltDeb, CltFin, AffDeb, AffFin : string; // code client et affaire début et fin
            DateDeb, DateFin : TdateTime; // date début et fin prise en compte activité
            Regroup : T_TYPEDATETB; // niveau de regroupemet Mois, total, date ...
            Valoris : T_MTACTIVITE; // niveau de valorisation
            NiveauAna : T_NIVEAUTB ; // niveau analyse en dsssous de l'affaire : article, ressource
            Unite: string; // H/J unité de restitution des qtés
            GroupConf : string; //MULTI géré par prog, zone pour groupe de confidentialité
            ActiviteReprise : string ; // type actiovite reprise
            PieceAchat, PieceAchatEng : string ; // nature des pièces d'acaht à traiter (si alimachat=true)
            AlimPrevu, AlimFacture, AlimReal, AlimAchat, AlimBoni,AlimAn,
            AffSurPeriode,AlimVisa, AlimCutOff,CutOffMois,
            AlimMois, AvecSSAff,ALimSolde: Boolean; // initialisation de ce que l'on doit traiter
            Libre1Clt, Libre2Clt, Libre3Clt, Libre4Clt, Libre5Clt,Libre6Clt, Libre7Clt, Libre8Clt, Libre9Clt, LibreAClt : string; //table libre client
            Ress1Clt, Ress2clt, Ress3Clt : string; //ressource libre client
            MoisClotDeb, MoisClotFin : integer;  //mois de cloture client
            GroupeClt : string;  //groupe client
            EtatAff, StatutAff, DptAff, EtbAff :string; //MULTI valeurs de sélections sur affaire
            RespAff :string; //valeurs de sélections sur affaire
            Libre1aff, Libre2aff, Libre3aff,Libre4aff,Libre5aff,Libre6aff,Libre7aff,Libre8aff,Libre9aff,LibreAaff : string; // table libre affaire
            Ress1Aff, Ress2Aff, Ress3Aff : string; //ressource libre affaire
            DateDebAffDeb, DateFinAffDeb, DateDebAffFin, DateFinAffFin,
            DateDebClot, DateFinClot,DateDebgar, DateFinGar, DateDebExerDeb,DateFinExerDeb,
            DateDebExerFin, DateFinExerFin, DateDebInter, DateFinInter :Tdatetime; // date début et fin sélection de l'affaire
           end;

Type    TSyntheseTBAffaire = RECORD
            TotPrevuPV,TotPrevuPR, TotFact, TotRealPR, TotRealPV ,TotBoniPr,TotBoniPv,
            TotFAE,TotAAe,totPCA,TotFAR,TotAAR,totCCA: Double;
            END;

Type    TParamTBAffaire = Class
          Public   // Paramétres de pilotage du TB
            AlimPrevu, AlimFacture, AlimReal, AlimAchat,AlimBoni,AlimAn,AffSurPeriode,AlimVisa,AlimCutOff,CutOffMois,AlimMois,
            GereFacturable,DetailParAffaire, AvecInitMove, CentrFam,CentrStatRes, CentrStatArt,ALimValoRess_Art,AlimSolde: Boolean;
            MtActivite : T_MtActivite;
            FTOBTB, TobDetEnc : TOB;
            SyntheseTB : TSyntheseTBAffaire;
            Constructor Create(xTypeTB: T_TypeTB; xTypeDateTB: T_TypeDateTB; xNiveauTB: T_NiveauTB;xUniteTemps : string; xNivAffRef : Boolean);
            Destructor Destroy;override;
            Function AlimTableauBordAffaire (var ListeChampsSel,stSQL,NatPieceAchat,NatPieceAchatEngage: string;
                     bChargeTobTB,bRetourneSQL,bChargementSelectif,bComplRes,bComplArt,bAvecSsAff:Boolean) : TOB;
            Procedure ChargeCriteres  (Contain : TComponent );
            //*Fonct d'init appelée depuis la fiche de lancement
               // MajMono Affaire et tiers
            Procedure MajMonoAffaire (CodeAffaire,DateDeb,DateFin: string; VisuSousAffaire : Boolean);
            Procedure MajMonoTiers   (CodeTiers,DateDeb,DateFin: string);

            procedure MajMultiTiers  (strTiersDeb, strTiersFin, strDateDeb, strDateFin:string); // $$$jp 20/12/2002
            procedure MajSelect  (Zone : T_ParamTb);
            procedure MajUneZone (NomCHamp: string; Deb,fin :variant; var posit : integer;Date, Multi:boolean);

            Procedure MajEcart (NumEcart : integer; Formule : hstring);  // Gestion des écarts

            Private
            NiveauTB  : T_NiveauTB;
            TypeTB    : T_TypeTB;
            TypeDateTB: T_TypeDateTB;
            NivAffRef,AlimPrevuAff,AlimPrevuBudg : Boolean;
            // C.B 11/09/02
            AlimPrevuPlanning : Boolean;
            ListeChampsAct,ListeChampsCutOff, ListeChampsligne,ListeChampsBud,
            // C.B 11/09/02
            listeChampsPlanning,
            WhereNaturePiece, WhereLigne,WhereNatAchatsurAct,
            ListeChampsTot,FormuleEcart1,FormuleEcart2,FormuleEcart3,FormuleEcart4,FormuleEcart5, UniteTemps,NatAchat,NatAchatEngage : string;
            MonoAffaire, MonoTiers,GereEcart1,GereEcart2,GereEcart3,GereEcart4,GereEcart5,ChargeTobTB,RetourneSQL,bJoinAffaire,
            bJoinTiers,bJoinTiersCompl,ComplRes,ComplArt,
            AvecssAff : Boolean;
            // mcd 18/10/02 Compteur  : Word;
            Compteur  : Integer;  //permet d'aller à 2 147 483 648 enrgt !....  Passé de LOngWord à Integer pour Avertissement
            DatePrevu,DateCutOff,DateSituation : TDateTime;  // date pour le prévu de l'affaire. alimenter à partir dates mission
            TobArticles:TOB ;
            AFOAssistants:TAFO_Ressources;
            FListeCriteres : array[1..MaxCritTB] of TCritereTBAffaire;
            XTab : array[1..7,1..MaxCorrespChamps] of string ;

            // Criteres
            Function ChampsTableVersTB (NomChamps : String; NumTable : integer ; Sens: T_TableTB): String;
            Function TestCompleteCriteresTB : Boolean;
            Procedure InitCriteresTB ;
            Procedure CompleteCriteresTB ;
            // liaison entre les tables
            Function InitLiaisonTables : Boolean;
            // Traitement + Alimentation du TB
            Procedure InitTobTB (TobDetTB, TobTBModele : TOB);

            function LoadTob(pStRequete, pStTableName : String) : Tob;
            procedure TraiteTob(TobTB, TobTraite, TobDet, TobTBModele : Tob;
                                    TypeAlim : T_TypeTB; TypeRequete : T_TypePRFA;
                                    CodeDetail, Prefixe : String;
                                    var LigneTraitee, MajTobExistante : Boolean);
            Procedure TraiteDetail(CodeDetail : String; TobDetail : TOB;Enr : T_TypeEnr);
            function PreparePlanning(pTob : Tob; pStPrefixe : String) : Tob;

            Procedure TraiteDetailBis(CodeDetail : String; TobDetail : TOB);
            Procedure AlimTobTBModele(TobTBModele ,TobDetail : Tob;Enr : T_TypeEnr);
            Procedure AlimTobTB (TobTBDet, pTob : TOB; MajTobExistante: Boolean; TypeAlim : T_TypePRFA );
            Function  VerifCleTobTB (TobTB : TOB; TypeAlim : T_TypePRFA ; pTob : Tob) : TOB;
            Procedure TraitePrevuAffaire (TobAffaire,TobTBModele,TobTB : TOB; var LigneTraitee : Boolean;Enr : T_TypeEnr);
            Procedure TraiteComplLigneTB(TobTB:TOB;Enr : T_TypeEnr);
            Procedure AlimComplLigneTB(TypeT : string; TobDetTB : TOB; Q: TQuery);

            // Gestion des requêtes
            Function FabriqueWhere(TypeTable: String; Var bJoinArt,bJoinRes: Boolean;Cutoff : Boolean;Enr : T_TypeEnr) : String ;
            Function FabriqueWherePRFA  (TypeTable,Code : String;TypeAlim : T_TypePRFA; Var bJoinArt,bJoinRes : Boolean;Enr : T_TypeEnr) : String ;
            Function FabriqueWhereNaturePiece : string;
            Function FabriqueWhereCutOff (Code,TypeEnr : String; Var bJoinArt,bJoinRes : Boolean;Enr : T_TypeEnr) : String ;
            Function FabriqueWhereBud (Code,TypeEnr : String; Var bJoinArt,bJoinRes : Boolean;Enr : T_TypeEnr) : String ;
            Function FabriqueWherePlanning (Code : String; Var bJoinArt,bJoinRes : Boolean;Enr : T_TypeEnr) : String ;
            Function FabriqueWhereLigne : string;
            Function FabriqueWhereNatAchatsurAct : string;
            Function ChampsTBversWhere(ChampWhere,TypeTable : String; NumCrit:Integer; Jointure : string): String;

            Function FabriqueRequete( TypeTable,CodeEnr : String;Enr : T_TypeEnr ) : String;
            Function FabriqueRequetePRFA (TypeRequete : T_TypePRFA; CodeEnr : string;Enr : T_TypeEnr)  : String;
            Function FabriqueRequeteCut ( CodeEnr,TypeEnr : string;Enr : T_TypeEnr)  : String;
            Function FabriqueRequeteBud ( CodeEnr,TypeEnr : string;Enr : T_TypeEnr)  : String;
            Function FabriqueRequetePlanning ( CodeEnr : string; Enr : T_TypeEnr)  : String;

            Procedure Centralisation(var ListeChamps : String; var bJoinArt, bJoinRes : Boolean);
            Procedure AlimListeChampsSt (TypeRequete : T_TypePRFA ; Var ChampsCle, ListeChamps : String; Var bJoinArt,bJoinRes : Boolean);
            Procedure AlimListeChampsStCut ( Var ChampsCle, ListeChamps : String; Var bJoinArt,bJoinRes : Boolean);
            Procedure AlimListeChampsStBud ( Var ChampsCle, ListeChamps : String; Var bJoinArt,bJoinRes : Boolean);
            Procedure AlimListeChampsStPlanning ( Var ChampsCle, ListeChamps : String; Var bJoinArt, bJoinRes : Boolean);
            Procedure AlimSuiteChamps(Var ListeChamps: string;Var bJoinArt, bJoinRes:boolean);

            // Gestion de la sélection sur la tob TB
            Function ListeChampsTBSelectif (AvecVirgule : Boolean): String;
            // Gestion des ecarts
            Procedure CalculLesEcarts(TobaTraiter : TOB);
            function GetDataEcart(ss:Hstring):variant;
            // Retraitement sur la tobaffaires
            Procedure TobAffavecSsAff(TobAffaires:TOB);
            END ;

// Fonction d'interphase avec les combos
Function GetNiveauTB  (ComboValue : String) : T_NiveauTB ;
Function GetRegroupeTB  (ComboValue : String) : T_TypeDateTB ;
function GetTypeTB (NiveauTiers, AvecDetailAff, bMulti : Boolean) : T_TypeTB;
// Fonction de recup des dates de debut de periode et de semaine
Function GetDateDebutPeriode ( Periode : Integer) : TDateTime;
Function GetDateDebutSemaine ( Semaine, Periode : Integer) : TDateTime;
Procedure InitRecordParamTb (Var zone : T_PARAMTB);

implementation

Constructor TParamTBAffaire.Create(xTypeTB: T_TypeTB; xTypeDateTB: T_TypeDateTB; xNiveauTB: T_NiveauTB; xUniteTemps : string; xNivAffRef : Boolean);
BEGIN
// Initialisation
Compteur := 1;
CentrFam  := false;  CentrStatArt  := false;CentrStatRes  := false;
AlimPrevu  := false; AlimFacture := false;  AlimReal := True; AlimAchat := False;
AlimSolde :=False;
AlimPrevuAff:=False; AlimPrevuPlanning := False; AlimPrevuBudg := False;

AlimBoni:=False; AlimCutOff:=False; AlimAn:=False;AffSurPeriode:=True;AlimVisa:=False; AlimMois:=False;
AlimValoRess_Art:=False;
MonoAffaire := False; MonoTiers := False; ChargeTobTB := False; RetourneSQL :=false; bJoinAffaire:=False; bJoinTiers:=False; bJoinTiersCompl:=False;
GereFacturable := False; DetailParAffaire := false; GereEcart1 := false; GereEcart2 := false; GereEcart3 := false; GereEcart4 := false; GereEcart5 := false; ComplRes:=False; ComplArt:=False;
AvecssAff := false;
ListeChampsTot := ''; ListeChampsAct := '';ListeChampsCutOff := '';
ListeChampsBud:='';  ListeChampsLigne := '';

//C.B 11/09/02
ListeChampsPlanning := '';

WhereNaturePiece := ''; WhereLigne := ''; WhereNatAchatsurAct:='';
MtActivite := ActPR;
SyntheseTB.TotPrevuPV := 0;SyntheseTB.TotPrevuPR := 0; SyntheseTB.TotFact := 0; SyntheseTB.TotRealPR := 0; SyntheseTB.TotRealPV := 0;
SyntheseTB.TotBoniPr := 0;SyntheseTB.TotBoniPV := 0;SyntheseTB.TotFAE := 0;SyntheseTB.TotAAE := 0;SyntheseTB.TotPCA := 0;
SyntheseTB.TotFAR := 0;SyntheseTB.TotAAR := 0;SyntheseTB.TotCCA := 0;
// recup des propriétés de la classe
TypeTB     := xTypeTB;
TypeDateTB := xTypeDateTB;
NiveauTB   := xNiveauTB;
NivAffRef  := xNivAffRef;

if xUnitetemps <> '' then UniteTemps := xUniteTemps else UniteTemps := VH_GC.AFMesureActivite;
// correspondance entre les champs
if not(InitLiaisonTables) then exit;
AvecInitMove := True;
END;

destructor TParamTBAffaire.Destroy;
begin
inherited;
if FTOBTB <> Nil then FTOBTB.Free; FTOBTB := Nil;
end;

Function TParamTBAffaire.AlimTableauBordAffaire (var ListeChampsSel,stSQL,NatPieceAchat,NatPieceAchatEngage: string; bChargeTobTB,bRetourneSQL,bChargementSelectif,bComplRes,bComplArt,bAvecSsAff:Boolean) : TOB;
Var TobAffaires, TobTiers, TobRessources, TobDet : TOB;
    stRequete, CodeTraite : string;
    QQ : TQuery ;
    i : integer;
Begin
  Result := Nil; ListeChampsSel := ''; stSQL := '';
  ChargeTobTB := bChargeTobTB;
  RetourneSQL := bRetourneSQL;                               
  ComplRes    := bComplRes; ComplArt :=bComplArt;
  NatAchat := NatPieceAchat;
  NatAchatEngage := NatPieceAchatEngage;
  AvecSsAff:=bAvecSsAff;
  TOBArticles:=TOB.Create('Les Articles',Nil,-1) ;  //mcd 28/01/03 change nom tob
  AFOAssistants:=TAFO_Ressources.Create;

  If AlimPrevu then
    begin
      // en fct des paramètres on alimente le type de prévu pris en compte
      If (GetParamSoc('SO_AFTYPEPREVUPR')='AFF') Or (GetParamSoc('SO_AFTYPEPREVUPV')='AFF') then AlimPrevuAff:=True;
      If (GetParamSoc('SO_AFTYPEPREVUPR')='PLA') Or (GetParamSoc('SO_AFTYPEPREVUPV')='PLA') then AlimPrevuPlanning := True;
      If (GetParamSoc('SO_AFTYPEPREVUPR')='BUD') Or (GetParamSoc('SO_AFTYPEPREVUPV')='BUD') then AlimPrevuBudg:=True;
    end;

  if AvecInitMove then InitMove (10,'');
  // Remise à blanc du tableau de bord pour cet utilisateur
  if Not(ChargeTobTB) then
     ExecuteSQL('DELETE FROM AFTABLEAUBORD WHERE ATB_UTILISATEUR="'+V_PGI.User+'"');
  // Chargement de la correspondance entre les champs des tables (ACT,LIGNE...) et celles du TB
  TobAffaires := Nil; TobRessources := Nil; TobTiers:=nil;
  if Not (TestCompleteCriteresTB) then
      BEGIN PGIBoxAF ('Critères incorrects' , 'Tableau de Bord'); Exit; END;
  // Contrôle que si chargement sélectif (compléments art + res non lus)
  if bChargementSelectif then begin ComplRes:=False; ComplArt:=false; end;

  // init des wheres
  WhereNaturePiece := FabriqueWhereNaturePiece;
  WhereLigne := FabriqueWhereLigne;
  WhereNatAchatsurAct := FabriqueWhereNatAchatsurAct;

  // Gestion de la TOB TB
  if ChargeTOBTB then FTobTB := TOB.Create('detail TB',Nil,-1) ;
  if AvecInitMove then MoveCur(False);
  // Chargement des affaires + jointure tiers
  If (TypeTB  = ttbAffaire) then
      BEGIN
      stRequete := FabriqueRequete( 'AFF','',Tnor);
      if (stRequete <> '') then
          Begin
          QQ := OpenSQL(stRequete,True,-1,'',true);
          if Not QQ.EOF then
              Begin
              TobAffaires:=TOB.Create('liste des affaires',Nil,-1) ;
              TobAffaires.LoadDetailDB('ChampsTB AFFAIRE','','',QQ,False,False);
              End;
         Ferme(QQ);
          End;
      if AvecSsAff then TobAffavecSsAff(TobAffaires);
      END;

  // Chargement des tiers
  If (TypeTB  = ttbTiers) then
      BEGIN
      stRequete := FabriqueRequete('T','',Tnor);
      if (stRequete <> '') then
          Begin
          QQ := OpenSQL(stRequete,True,-1,'',true);
          if Not QQ.EOF then
              Begin
              TobTiers:=TOB.Create('liste des tiers',Nil,-1) ;
              TobTiers.LoadDetailDB('ChampsTB TIERS','','',QQ,False,False);
              End;
         Ferme(QQ);
          End;
      END;

  // Chargement des ressources
  If (TypeTB  = ttbRessource) then
      BEGIN
      stRequete := FabriqueRequete( 'ARS','',Tnor);
      if (stRequete <> '') then
          Begin
          QQ := OpenSQL(stRequete,True,-1,'',true);
          if Not QQ.EOF then
              BEGIN
              TobRessources:=TOB.Create('liste des ressources',Nil,-1) ;
              TobRessources.LoadDetailDB('ChampsTB RESSOURCE','','',QQ,False,False);
              END;
          Ferme(QQ);
          End;
      END;
  if AvecInitMove then MoveCur(False);

  // Gestion de la TOB TB
  if ChargeTOBTB then FTobTB := TOB.Create('detail TB',Nil,-1)
                 else FTOBTB := nil;
  // obligation d'avoir le même ordre de tri par affaire, ressource entre les différentes requêtes...
  If (TypeTB  = ttbAffaire) then // Niveau Affaire
      BEGIN
      if TobAffaires <> Nil then
        BEGIN
        if AvecInitMove then InitMove(TOBAffaires.Detail.Count,'');
        for i:=0 to TOBAffaires.Detail.Count-1 do
            BEGIN
            TOBDet:=TOBAffaires.Detail[i] ;
            CodeTraite := TobDet.GetValue('AFF_AFFAIRE');
            DateCutOff:=TobDet.getvalue('AFF_DATECUTOFF');         // mcd 24/07/02
            DateSituation:=TobDet.getvalue('AFF_DATESITUATION');         // mcd 24/07/02
                       // mcd 25/03/02 déplacé depuis TraitePevuAffaire
            DatePrevu := StrToDate(TobDet.getValue('AFF_DATEDEBGENER'));
           if (Dateprevu = iDate1900) then
              DatePrevu := StrToDate(TobDet.getValue('AFF_DATEDEBUT'));
           if (TypeDateTb =TdaCumul) then DatePrevu:=Idate1900; //mcd 18/03/02

            TraiteDetailBis(CodeTraite, TobDet);
            if AvecInitMove then MoveCur (False);
            End;
        END;
      END else
  if (TypeTB  = ttbTiers) then   // Niveau tiers
      BEGIN
      if TOBTiers <> Nil then
        BEGIN
        if AvecInitMove then InitMove(TOBTiers.Detail.Count,'');
        for i:=0 to TOBTiers.Detail.Count-1 do
            BEGIN
            TOBDet:=TOBTiers.Detail[i] ;
            CodeTraite := TobDet.GetValue('T_TIERS');
            TraiteDetailBis(CodeTraite, TobDet);
            if AvecInitMove then MoveCur (False);
            End;
        END;
      END else
  If (TypeTB  = ttbRessource) then
      BEGIN
      if TOBRessources <> Nil then
        BEGIN
        if AvecInitMove then InitMove(TOBRessources.Detail.Count,'');
        for i:=0 to TOBRessources.Detail.Count-1 do
            BEGIN
            TOBDet:=TOBRessources.Detail[i] ;
            CodeTraite := TobDet.GetValue('ARS_RESSOURCE');
            TraiteDetailBis(CodeTraite, TobDet);
            if AvecInitMove then MoveCur (False);
            End;
        END;
      END else
      ;
  if ChargeTobTB then
     BEGIN
     if bChargementSelectif then ListeChampsSel := ListeChampsTBSelectif(false);
     if FTobTB.Detail.count < 1 then result := Nil else Result := FTobTB;
     END;
  if bRetourneSQL then
     BEGIN
     if bChargementSelectif then stSQL := 'SELECT ATB_TIERS,' + ListeChampsTBSelectif(true)
                            else stSQL := 'SELECT * ';  // normal ... on veut tout
     stSQL := stSQL + ' from AFTABLEAUBORD where ATB_UTILISATEUR="'+V_PGI.User+'"';
     END;
  if AvecInitMove then FiniMove ();
                               
  TobAffaires.Free; TobTiers.Free; TobRessources.Free;
  AFOAssistants.Free;
  TOBArticles.Free;
End;

Procedure TParamTBAffaire.TraitePrevuAffaire (TobAffaire,TobTBModele,TobTB : TOB; var LigneTraitee : Boolean;Enr : T_TypeEnr);
Var PrevuBaseEche,Trouve : Boolean;
    TotHTLigne,TotPrevu,TotQteLigne : Double;
    NbEcheFact,i : integer;
    TobDet,TobDet2 : TOB;
    QQ : Tquery;
BEGIN
  TotHTLigne := 0;
  PrevuBaseEche:=True;
  If TobAffaire.getValue('AFF_CALCTOTHTGLO')<>'X' then begin
     Totprevu:= TobAffaire.getValue('AFF_TOTALHTGLO'); //mcd 11/02/02 + 18/03/02
     for i := 0 to TobTB.Detail.count-1 do
        BEGIN   // il faut remettre à 0 les mtt des lignes éventuelles, si pas de calcul auto
        TobTB.Detail[i].PutValue('ATB_PREVUPV',0);
        TobTB.Detail[i].PutValue('ATB_PREVUPR',0); // mcd 23/04/02 ajoute des 3 lignes de RAZ
        TobTB.Detail[i].PutValue('ATB_PREVUQTE',0);
        TobTB.Detail[i].PutValue('ATB_PREQTEUNITEREF',0);
        END;
     end
  else begin
    for i := 0 to TobTB.Detail.count-1 do
        BEGIN
        TotHTLigne := TotHTLigne + TobTB.Detail[i].GetValue('ATB_PREVUPV');
        END;
    TotPrevu := CalculTotalAffaire(TobAffaire.getValue('AFF_AFFAIRE'), TobAffaire.getValue('AFF_GENERAUTO'),
      TobAffaire.getValue('AFF_TYPEPREVU'),TotHTLigne,Nil, NbEcheFact, PrevubaseEche,TobAffaire.getValue('AFF_DATEDEBGENER') );
    End;
  if PrevuBaseEche then
     BEGIN // Ajout d'une montant correspondant aux écheances sur le code acompte
     if TotPrevu = 0 then Exit;
     If Enr<>Tnor then exit;   // mcd 23/04/2000 sur enrgt gloal mis uniquement sur NOR
     TotQteLigne:= TobAffaire.getValue('AFF_ACTIVITEGLOBAL'); //mcd 23/04/02

    // TobDet:= Tob.Create('AFTABLEAUBORD',TobTB,-1);
     TobDet:= Tob.Create('AFTABLEAUBORD',Nil,-1);
     InitTobTB (TobDet, TobTBModele);
     TobDet.PutValue('ATB_DATE', DatePrevu);
     if (TypeDateTB =tdaSemaine) then
       begin
       TobDet.PutValue('ATB_SEMAINE', NumSemaine(DatePrevu));
       end;
     if (TypeDateTB =tdaMois) or (TypeDateTB =tdaSemaine) then
        begin
        TobDet.PutValue('ATB_PERIODE', GetPeriode(DatePrevu));
        {$IFDEF BTP}
        TobDet.PutValue ('ATB_LIBPERIODE', BTPOLEFormatPeriode (TobDet.GetValue('ATB_PERIODE'),True));
        {$ELSE}
        TobDet.PutValue ('ATB_LIBPERIODE', OLEFormatPeriode (TobDet.GetValue('ATB_PERIODE'),True));
        {$ENDIF}
        end;

     if (NiveauTB <> tniRes) and (NiveauTB <> tniTypeRes)
        and(NiveauTB <> tniFamilleArt1) and(NiveauTB <> tniFamilleArt2)and(NiveauTB <> tniFamilleArt3)
        and(NiveauTB <> tniStatArt1) and(NiveauTB <> tniStatArt2) and(NiveauTB <> tniStatArt3)
        then begin
        TobDet.PutValue('ATB_TYPEARTICLE','PRE');
        end;
        //mcd 18/03/02 ajout test pour ne pas avoir plusieurs lignes
     if (NiveauTB <> tniTypeArt) and (NiveauTB <> tniRes) and(NiveauTB <> tniTypeRes)
        then begin
        TobDet.PutValue('ATB_CODEARTICLE',VH_GC.AFACOMPTE);
        TobDet.PutValue('ATB_ARTICLE',CodeArticleUnique2(VH_GC.AFACOMPTE,''));
        TobDet.PutValue('ATB_ACTIVITEREPRIS','F');
        end;
     if (NiveauTB = tniFamilleARt1) or (NiveauTB = tniFamilleArt2) or(NiveauTB = tniFamilleARt3)
        or (NiveauTB = tniStatARt1) or (NiveauTB = tniStatArt2) or(NiveauTB = tniStatARt3)
        then begin
        TobDet.PutValue('ATB_CODEARTICLE','');
        TobDet.PutValue('ATB_ARTICLE','');
        TobDet.PutValue('ATB_ACTIVITEREPRIS','');
        QQ:=OpenSQL('SELECT GA_FAMILLENIV1,GA_FAMILLENIV2,GA_FAMILLENIV3,GA_LIBREART1,GA_LIBREART2,GA_LIBREART3 FROM ARTICLE WHERE GA_ARTICLE="'
                            +CodeArticleUnique2(VH_GC.AFACOMPTE,'')+'"',TRUE,-1,'',true) ;
        if Not QQ.EOF then
           begin
           If NiveauTB =TniFamilleArt1 then TobDet.PutValue('ATB_FAMILLENIV1',QQ.FindField('GA_FAMILLENIV1').asString)
           else if NiveauTB =TniFamilleArt2 then TobDet.PutValue('ATB_FAMILLENIV2',QQ.FindField('GA_FAMILLENIV2').asString)
           else if NiveauTB =TniFamilleArt3 then TobDet.PutValue('ATB_FAMILLENIV3',QQ.FindField('GA_FAMILLENIV3').asString)
           else if NiveauTB =TniStatArt1 then TobDet.PutValue('ATB_LIBREART1',QQ.FindField('GA_LIBREART1').asString)
           else if NiveauTB =TniStatArt2 then TobDet.PutValue('ATB_LIBREART2',QQ.FindField('GA_LIBREART2').asString)
           else if NiveauTB =TniStatArt3 then TobDet.PutValue('ATB_LIBREART3',QQ.FindField('GA_LIBREART3').asString);
           end;
        Ferme(QQ);

        end;
     if (NiveauTB = tniTypeRes) then TobDet.PutValue('ATB_TYPERESSOURCE','SAL');

     // on regarde si l'enrgt que l'on veut créer, existe déja en tob
     // fait uniquement sur les champs renseignés
     // mcd 18/03/02
     trouve := false;
     TOBDet2 := nil;
     For i:=0 to TobTB.Detail.count-1 do
      BEGIN
         TOBDet2:=TobTB.Detail[i];
         if (TobDet.GetValue('ATB_DATE')=TobDEt2.Getvalue('ATB_DATE'))
          and (TobDet.GetValue('ATB_TYPERESSOURCE')=TobDEt2.Getvalue('ATB_TYPERESSOURCE'))
          and (TobDet.GetValue('ATB_RESSOURCE')=TobDEt2.Getvalue('ATB_RESSOURCE'))
          and (TobDet.GetValue('ATB_TYPEARTICLE')=TobDEt2.Getvalue('ATB_TYPEARTICLE'))
          and (TobDet.GetValue('ATB_CODEARTICLE')=TobDEt2.Getvalue('ATB_CODEARTICLE'))
          and (TobDet.GetValue('ATB_ARTICLE')=TobDEt2.Getvalue('ATB_ARTICLE'))
          and (TobDet.GetValue('ATB_UTILISATEUR')=TobDEt2.Getvalue('ATB_UTILISATEUR'))
          and (TobDet.GetValue('ATB_TYPEENR')=TobDEt2.Getvalue('ATB_TYPEENR'))
          and (TobDet.GetValue('ATB_AFFAIRE')=TobDEt2.Getvalue('ATB_AFFAIRE'))
          and (TobDet.GetValue('ATB_TIERS')=TobDEt2.Getvalue('ATB_TIERS'))
          and (TobDet.GetValue('ATB_FAMILLENIV1')=TobDEt2.Getvalue('ATB_FAMILLENIV1'))
          and (TobDet.GetValue('ATB_LIBREART1')=TobDEt2.Getvalue('ATB_LIBREART1'))
          and (TobDet.GetValue('ATB_FAMILLENIV2')=TobDEt2.Getvalue('ATB_FAMILLENIV2'))
          and (TobDet.GetValue('ATB_LIBREART2')=TobDEt2.Getvalue('ATB_LIBREART2'))
          and (TobDet.GetValue('ATB_FAMILLENIV3')=TobDEt2.Getvalue('ATB_FAMILLENIV3'))
          and (TobDet.GetValue('ATB_LIBREART3')=TobDEt2.Getvalue('ATB_LIBREART3'))
          then Trouve:=True else Trouve:=False;
         if trouve then break;
       end;        

     if not (trouve) then
       begin
         TobDet2:= Tob.Create('AFTABLEAUBORD',TobTB,-1);
         TobDEt2.dupliquer (Tobdet,False,True);
       end;

     TobDet2.PutValue('ATB_PREVUPV',TotPrevu);
     TobDet2.PutValue('ATB_PREVUQTE',TotQteLigne); //mcd 23/04/02
     TobDet2.PutValue('ATB_PREQTEUNITEREF',ConversionUnite('H',UniteTemps,TotQteLigne)); //mcd 23/04/02 NB on considère que saisie en heure ..
     TobDet2.PutValue('ATB_LIBELLE',TraduitGA('Cumul des échéances de l''affaire'));
     LigneTraitee := true;
     SyntheseTB.TotPrevuPV := SyntheseTB.TotPrevuPV + TotPrevu;
     TobDEt.free;
     //Tobdet:=Nil;
     END
  else
     BEGIN
     if NbEcheFact > 1 then
        BEGIN
         // mcd 12/07/02 ???? SyntheseTB.TotPrevu := SyntheseTB.TotPrevu + (TotPrevu - TotHTLigne);
        SyntheseTB.TotPrevuPV := SyntheseTB.TotPrevuPV + TotPrevu ;
        for i := 0 to TobTB.Detail.count-1 do
          BEGIN
          TotHTLigne := TobTB.Detail[i].GetValue('ATB_PREVUPV');
          TobTB.Detail[i].putValue('ATB_PREVUPV',TotHTLigne * NbEcheFact);
          TotQteLigne := TobTB.Detail[i].GetValue('ATB_PREVUQTE');
          TobTB.Detail[i].putValue('ATB_PREVUQTE',TotQteLigne * NbEcheFact);
          TotQteLigne := TobTB.Detail[i].GetValue('ATB_PREQTEUNITEREF');
          TobTB.Detail[i].putValue('ATB_PREQTEUNITEREF',TotQteLigne * NbEcheFact);
          END;
        END
     else    SyntheseTB.TotPrevuPV := SyntheseTB.TotPrevuPV + TotPrevu; // mcd 12/07/02
     END;
END;
                             

Procedure TParamTBAffaire.InitTobTB (TobDetTB, TobTBModele : TOB);
BEGIN
   TobDetTB.Dupliquer (TobTBModele, False,True);
   TobDetTB.PutValue( 'ATB_NUMERO' , Compteur);
   Inc( Compteur);
END;
 
   // fct qui appelle traite detail, avec Option AN ou pas ...
Procedure TParamTBAffaire.TraiteDetailBis(CodeDetail : String; TobDetail : TOB);
var       wTypeDateTB: T_TypeDateTB;
 AffSurPer : boolean;
Begin
     // appel sans gestion AN
TraiteDetail(CodeDetail, TobDetail,Tnor);
If AlimAN then begin
          // mcd 06/01/03 si rien en période, on en traite pas  , en fct de l'option écran
   If not(AffSurPeriode) then AffSurPer:= ExisteSql ('SELECT ATB_AFFAIRE FROM AFTABLEAUBORD WHERE ATB_AFFAIRE="'+ CodeDetail+'"')
    else AffSurPer:=True; // on fait comme avant... on prend tout
   If AffSurPer then
    begin
          // on gère les AN
          // on stock ancien regroupement par date
    wTypeDateTb :=TypeDateTb;
    TypeDateTb:=TdaCumul;
    TraiteDetail(CodeDetail, TobDetail,Tan);
    TypeDateTb:=WtypeDateTb;
    end;
   end;
If AlimMois then begin
          // on gère un cumul pour le mois en cours
          // on stock ancien regroupement par date
   wTypeDateTb :=TypeDateTb;
   TypeDateTb:=TdaCumul;
   TraiteDetail(CodeDetail, TobDetail,TMois);
   TypeDateTb:=WtypeDateTb;
   end;
end;

function TParamTBAffaire.LoadTob(pStRequete, pStTableName : String) : Tob;
var
  vQr : Tquery;
begin

  result := nil;
  vQr := nil;
  Try
    vQr := OpenSQL(pStRequete,True,-1,'',true);
    if Not vQr.Eof then
      begin
        result := Tob.create(pStTableName, nil, -1);
        //mcd ...n'est pas fait sur la table, car on peut avoir une jointure sur artcile,ressource result.LoadDetailDB(pStTableName, '', '', vQr , False, True);
        result.LoadDetailDB('', '', '', vQr , False, True);
      end
  finally
    ferme(vQr);
  end;
end;

procedure TParamTBAffaire.TraiteTob(TobTB, TobTraite, TobDet, TobTBModele : Tob;
                                    TypeAlim : T_TypeTB; TypeRequete : T_TypePRFA;
                                    CodeDetail, Prefixe : String;
                                    var LigneTraitee, MajTobExistante : Boolean);
var
  i         : Integer;
  ChampTest : String;
  StoDate:tdateTime; // mcd 15/10/2002
  TypeArt,TypeRes,REs,Art,ArtUni:string ; //mcd 15/10/02

begin
  stodate:=Idate1900;
  if TobTraite <> nil then
  for i := 0 to TobTraite.detail.count - 1 do
    begin
      if TypeAlim = ttbAffaire        then ChampTest := TobTraite.detail[i].GetValue(Prefixe + 'AFFAIRE')
      else if TypeAlim = ttbTiers     then ChampTest := TobTraite.detail[i].GetValue(Prefixe + 'TIERS')
      else if TypeAlim = ttbRessource then ChampTest := TobTraite.detail[i].GetValue(Prefixe + 'RESSOURCE')
      else ChampTest := '';
      If (CodeDetail = ChampTest) then
        Begin
        //  TobDet := nil;
          if LigneTraitee then
           if ((typeRequete=tactivite) OR (typeRequete=tBoni)) and (AlimValoREss_Art) then
              begin  // mcd 15/10/2002 si recalul PV , on lit sur détail l'activite, mais il faut écrire en fct détail choisi
              If TypedateTb=TdaCumul then begin
                 StoDate:=TobTraite.detail[I].GetValue('ACT_DATEACTIVITE');
                 TobTraite.detail[I].PutValue('ACT_DATEACTIVITE',Idate1900);
                 end;
                 // mcd 30/10/02 ajout tnitous
              if (NiveauTb<>TniTypeart) and (NiveauTb<>TniARt) and (NiveauTb<>TniTypeArtREs) and (NiveauTb<>TniTous) then begin
                 TypeArt:=TobTraite.detail[I].GetValue('ACT_TYPEARTICLE');
                 TobTraite.detail[I].PutValue('ACT_TYPEARTICLE','');
                 end;
              if (NiveauTb<>Tniart) and (NiveauTb<>TniARtRes) and (NiveauTb<>TniTous)  then begin
                 Art:=TobTraite.detail[I].GetValue('ACT_CODEARTICLE');
                 TobTraite.detail[I].PutValue('ACT_CODEARTICLE','');
                 ArtUni:=TobTraite.detail[I].GetValue('ACT_ARTICLE');
                 TobTraite.detail[I].PutValue('ACT_ARTICLE','');
                 end;
              if (NiveauTb<>TniartREs) and (NiveauTb<>TniRes) and (NiveauTb<>TniTypeRes) and (NiveauTb<>TniTous) then begin
                 TypeRes:=TobTraite.detail[I].GetValue('ACT_TYPERESSOURCE');
                 TobTraite.detail[I].PutValue('ACT_TYPERESSOURCE','');
                 end;
              if (NiveauTb<>TniRes) and (NiveauTb<>TniArtREs) and (NiveauTb<>TniTous) then begin
                 Res:=TobTraite.detail[I].GetValue('ACT_RESSOURCE');
                 TobTraite.detail[I].PutValue('ACT_RESSOURCE','');
                 end;
              end;
            //TobTB Déjà alimentée vérif si nouvelle tob ou maj tob existante
           TobDet := VerifCleTobTB (TobTB, TypeRequete, TobTraite.detail[i]);
           if ((typeRequete=tactivite) OR (typeRequete=tBoni)) and (AlimValoREss_Art) then
              begin  // mcd 15/10/2002 si recalul PV , on remet vaelur origine pour calcul prix
              If TypedateTb=TdaCumul then TobTraite.detail[I].PutValue('ACT_DATEACTIVITE',StoDate);
              if (NiveauTb<>TniTypeart) and (NiveauTb<>TniARt) and (NiveauTb<>TniTypeArtREs)  and (NiveauTb<>TniTous) then TobTraite.detail[I].PutValue('ACT_TYPEARTICLE',TypeArt);
              if (NiveauTb<>Tniart) and (NiveauTb<>TniARtRes)  and (NiveauTb<>TniTous) then
                begin
                TobTraite.detail[I].PutValue('ACT_CODEARTICLE',Art);
                TobTraite.detail[I].PutValue('ACT_ARTICLE',ArtUni);
                end;
              if (NiveauTb<>TniartREs) and (NiveauTb<>TniRes) and (NiveauTb<>TniTypeRes)  and (NiveauTb<>TniTous) then TobTraite.detail[I].PutValue('ACT_TYPERESSOURCE',TypeRes);
              if (NiveauTb<>TniRes) and (NiveauTb<>TniArtREs)  and (NiveauTb<>TniTous) then TobTraite.detail[I].PutValue('ACT_RESSOURCE',Res);
              end;

          if TobDet = nil then
            Begin
              TobDet:= Tob.Create('AFTABLEAUBORD',TobTB,-1);
              InitTobTB (TobDet, TobTBModele);
              MajTobExistante := False;
            End
          else
            MajTobExistante := True;
          AlimTobTB (TobDet, TobTraite.detail[i], MajTobExistante, TypeRequete);
          LigneTraitee := True;
        End;

      If (CodeDetail < ChampTest) then Break;
    End;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 16/09/2002
Modifié le ... :
Description .. : repartition des quantités et des montants avec
                 un ligne de planning par jour
Mots clefs ... :
*****************************************************************}
function TParamTBAffaire.PreparePlanning(pTob : Tob; pStPrefixe : String) : Tob;
var
  i,j           : Integer;
  vInQte        : Integer;
  vRdMntPR      : Double;
  vRdMntPV      : DOuble;
  vRdLastQte    : Double;
  vRdLastMntPR  : Double;
  vRdLastMntPV  : Double;
  vTobTypeRes   : Tob;
  vSt           : String;
  vQr           : TQuery;

  procedure createLignePlanning(var pTobMere : tob; pTob, pTobTypeRessource : tob; pRdQte, pRdMntPR, pRdMntPv : Double; pInJour : Integer; pStPrefixe : String);
  var
    vTob        : Tob;
    vTobTypeRes : Tob;

  begin
    vTob := Tob.create('AFPLANNING', pTobMere, -1);
    vTob.Dupliquer(pTob,true, true);
    vTob.putValue(pStPrefixe + '_QTEPLANIFIEE', pRdQte);
    vTob.putValue(pStPrefixe + '_INITPTPR', pRdMntPR);
    vTob.putValue(pStPrefixe + '_INITPTVENTEHT', pRdMntPv);
    vTob.putValue('APL_DATEDEBPLA', PlusDate(varAsType(pTob.GetValue('APL_DATEDEBPLA'), varDate), j, 'J'));
    vTob.putValue('APL_DATEFINPLA', pTob.GetValue('APL_DATEDEBPLA'));
    vTob.AddChampSupValeur('APL_SEMAINE', NumSemaine(pTob.GetValue('APL_DATEDEBPLA')));
    vTob.AddChampSupValeur('APL_PERIODE', GetPeriode(pTob.GetValue('APL_DATEDEBPLA')));
    vTobTypeRes := pTobTypeRessource.FindFirst(['ARS_RESSOURCE'], [pTob.GetValue('APL_RESSOURCE')], false);
    if vTobTypeRes <> nil then
      vTob.AddChampSupValeur('ARS_TYPERESSOURCE', vTobTypeRes.GetValue('ARS_TYPERESSOURCE'))
    else
      vTob.AddChampSup('ARS_TYPERESSOURCE', false);
  end;

begin
  result := nil;
  if assigned(pTob) then
    begin
      result := Tob.create('AFPLANNING', nil, -1);
      vTobTypeRes := Tob.create('RESSOURCE', nil, -1);
      try
        // C.B 19/11/02 deplacement de la requete
        vSt := 'SELECT ARS_TYPERESSOURCE, ARS_RESSOURCE FROM RESSOURCE';
        vQr := OpenSQL(vSt, True,-1,'',true);
        if not vQr.eof then vTobTypeRes.LoadDetailDB('RESSOURCE', '', '', vQr, False, False);

        for i := 0 to pTob.detail.count - 1 do
          begin
            vInQte := round(int(pTob.detail[i].GetValue(pStPrefixe + '_QTEPLANIFIEE')));
            // calcul de la derniere quantité
            if frac(pTob.detail[i].GetValue(pStPrefixe + '_QTEPLANIFIEE')) <> 0 then
              vRdLastQte := frac(pTob.detail[i].GetValue(pStPrefixe + '_QTEPLANIFIEE'))
            else
              vRdLastQte := 0;
                                                  
            //Calcul des derniers montants
            if vInQte <> 0 then
              begin
                vRdMntPR :=  arrondi(pTob.detail[i].GetValue(pStPrefixe + '_INITPTPR') / vInQte, V_PGI.OkDecV);
                vRdLastMntPR := arrondi(pTob.detail[i].GetValue(pStPrefixe + '_INITPTPR') - (vRdMntPR * vInQte), V_PGI.OkDecV);
                vRdMntPv :=  arrondi(pTob.detail[i].GetValue(pStPrefixe + '_INITPTVENTEHT') / vInQte, V_PGI.OkDecV);
                vRdLastMntPv := arrondi(pTob.detail[i].GetValue(pStPrefixe + '_INITPTVENTEHT') - (vRdMntPv * vInQte), V_PGI.OkDecV);
              end
            else
              begin
                vRdMntPR := 0; vRdLastMntPR := 0;
                vRdMntPv := 0; vRdLastMntPv := 0;
              end;

            for j := 0 to pTob.detail[i].GetValue(pStPrefixe + '_QTEPLANIFIEE') - 2 do
              createLignePlanning(result, pTob.detail[i], vTobTypeRes, 1, vRdMntPR, vRdMntPv, j, pStPrefixe);

            // on met le dernier montant sur le dernier element de la boucle
            // si la quantité est entière
            // sinon, on met le dernier montant si la ligne supplémentaire
            if vRdLastQte <> 0 then
              createLignePlanning(result, pTob.detail[i], vTobTypeRes, 1,
                                  vRdMntPR + vRdLastMntPR, vRdMntPv + vRdLastMntPv,
                                  pTob.detail[i].GetValue(pStPrefixe + '_QTEPLANIFIEE') - 1, pStPrefixe)
            else
              begin
                // gestion de la derniere quantite et de la quantité supplementaire
                createLignePlanning(result, pTob.detail[i], vTobTypeRes, 1,
                                  vRdMntPR, vRdMntPv,  vInQte - 1, pStPrefixe);

                createLignePlanning(result, pTob.detail[i], vTobTypeRes, vRdLastQte,
                              vRdLastMntPR, vRdLastMntPv, vInQte, pStPrefixe);
              end;
          end;
      finally
        vTobTypeRes.Free;
      end;
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... :
Modifié le ... :
Description .. : En fonction des types d'enregistrements traités :
                 Réalisé, Boni, cutoff, facturé, prévu,
                 charge les tob des tables correspondantes,
                 puis boucle sur ces tob pour alimenter la tob TableauBord
                 avec ces informations.
                 Puis en fin de fonction, écrit la tob tableau de bord dans
                 la table AfTableauBord

Mots clefs ... :
*****************************************************************}
Procedure TParamTBAffaire.TraiteDetail(CodeDetail : String; TobDetail : TOB;Enr : T_TypeEnr);
Var
  PrefixeAct,PrefixeFac, PrefixeCutOff,  PrefixeBud, PrefixePlanning : String;
  vStRequete : string;
  LigneTraitee, MajTobExistante: Boolean;
  TobActivite, TobLignes, TobBoni, TobCutOff, TobBudget, TobQfactEclat,
  TobPlanning, TobPlanningBrut : Tob;
  TobDet, TobTB, TobTBModele  : TOB;
  i : integer;

Begin

  LigneTraitee := False;

  PrefixeAct:= 'ACT_'; PrefixeFac:= 'GL_'; PrefixeCutOff := 'ACU_';
  PrefixeBud := 'ABU_'; PrefixePlanning := 'APL_';

  TobTB := TOB.Create('detail TB',Nil,-1) ;
  TobTBModele := Tob.Create('AFTABLEAUBORD',Nil,-1);
  Try
    AlimTobTBModele(TobTBModele, TobDetail, Enr);
    TobDet := nil;
    TobLignes := nil;
    TobQfactEclat := nil;

    // chargement du réalisé  ACTIVITE
    if AlimReal Then
      begin
        vStRequete  := FabriqueRequetePRFA (tActivite,CodeDetail,Enr);
        TobActivite := loadTob(vStRequete, 'ZACTIVITE');
        try
          TraiteTob(TobTB, TobActivite, TobDet, TobTBModele, TypeTB, tActivite,
                    CodeDetail, PrefixeAct, LigneTraitee, MajTobExistante);
        finally
           if TobActivite <> nil then TobActivite.free;
        end;
      end;

    // chargement du réalisé  Boni/mali
    if AlimBoni Then
      begin
        vStRequete := FabriqueRequetePRFA (tBoni,CodeDetail,Enr);
        TobBoni := loadTob(vStRequete, 'ZACTIVITE'); //BONI
        try
          TraiteTob(TobTB, TobBoni, TobDet, TobTBModele, TypeTB, tBoni,
                    CodeDetail, PrefixeAct, LigneTraitee, MajTobExistante);
        finally
          if TobBoni <> nil then TobBoni.free;
        end;
      end;

    // chargement du réalisé  CutOff
    if AlimCutOff Then
      begin
        vStRequete := FabriqueRequeteCut(CodeDetail,'CVE',Enr);
        TobCutOff := LoadTob(vStRequete, 'ZAFCUMUL'); //CUTOFF
        try
          TraiteTob(TobTB, TobCutOff, TobDet, TobTBModele, TypeTB, tCutOff,
                    CodeDetail, PrefixeCutOff, LigneTraitee, MajTobExistante);
        finally
          if TobCutOff <> nil then TobCutOff.free;
        end;
      end;

    // chargement du prévu pour le Budget
    if AlimPrevuBudg Then
      begin
        vStRequete := FabriqueRequeteBud(CodeDetail,'PVT',Enr);
        TobBudget := LoadTob(vStRequete, 'ZAFBUDGET');
        try
          TraiteTob(TobTB, TobBudget, TobDet, TobTBModele, TypeTB, tBud,
                    CodeDetail, PrefixeBud, LigneTraitee, MajTobExistante);
        finally
          if TobBudget <> nil then TobBudget.free;
        end;
      end;

    // C.B 11/09/02
    // chargement du prévu pour le planning
    if AlimPrevuPlanning Then
      begin                                      
        vStRequete := FabriqueRequetePlanning(CodeDetail,Enr);
        TobPlanningBrut := LoadTob(vStRequete, 'AFPLANNING');
        // C.B 19/11/02 ajout des prefixes 
        if pos('SUM', vStRequete) = 0 then
          TobPlanning := PreparePlanning(TobPlanningBrut, 'APL')
        else
          TobPlanning := PreparePlanning(TobPlanningBrut, 'SUM');
        try
          TraiteTob(TobTB, TobPlanning, TobDet, TobTBModele, TypeTB, tPlanning,
                    CodeDetail, PrefixePlanning, LigneTraitee, MajTobExistante);
        finally
          if TobPlanning <> nil then TobPlanning.free;
          if TobPlanningBrut <> nil then TobPlanningBrut.free;
        end;
      end;
 
     // Chargement des pieces (PREVU , FACTURE )
    if AlimFacture or AlimPrevuAff Then      //alimPrevu : prévu de l'affaire ....
      BEGIN
        // attention pour les requêtes sur le slignes, fait sur datepiece compris dans période,
        // Y compris pour le prévu si fait à partir des lignes!!!!
        If (GetParamSoc('SO_AFFACTPARRES') <> 'SAN') then // mcd 03/06/2002 ajout traitement fact eclat
          begin
            if AlimFacture then
              begin
                vStRequete := FabriqueRequeteCut (CodeDetail,'FAC',Enr);
                TobQfactEclat := LoadTob(vStRequete, 'ZAFCUMUL'); // FACTECLAT
              end;

            if AlimPrevuAff then
              begin
                vStRequete := FabriqueRequetePRFA (tligne,CodeDetail,Enr);
                TobLignes := LoadTob(vStRequete, 'ZLIGNE'); // LIGNES
              end;
          end
        else
          begin
            vStRequete := FabriqueRequetePRFA (tligne,CodeDetail,Enr);
            TobLignes := LoadTob(vStRequete, 'ZLIGNES');
          end;

          try
            TraiteTob(TobTB, TobLignes, TobDet, TobTBModele, TypeTB, tLigne,
                      CodeDetail, PrefixeFac, LigneTraitee, MajTobExistante);

            TraiteTob(TobTB, TobQfactEclat, TobDet, TobTBModele, TypeTB, tFActEclat,
                      CodeDetail, PrefixeCutOff, LigneTraitee, MajTobExistante);
          finally
            if TobLignes <> nil then TobLignes.free;
            if TobQfactEclat <> nil then TobQfactEclat.free;
          end;
      END;

    // retraitement sur le prévu ( ajuot des écheances si pas de lignes ou * des lignes / nb ech
    if  (TypeTB = ttbAffaire) and (AlimPrevuAff)  then
        TraitePrevuAffaire(TobDetail, TobTBModele, TobTB,LigneTraitee,Enr); // A voir pour le prévu / niveau client ???

    // retraitement sur les compléments de la ligne Article / ressources
    if  (TypeTB = ttbAffaire) and ((ComplRes) or (ComplArt)) then
        TraiteComplLigneTB(TobTB,Enr);

    If LigneTraitee Then
       BEGIN
       CalculLesEcarts(TobTB);
       if (ChargeTobTB) then
          BEGIN
          if (TOBTB.Detail.count > 0) And (FTOBTB <> Nil) then
            BEGIN
            for i := TobTB.detail.count-1 downto 0 do
                BEGIN
                TOBDet := TOBTB.detail[i];     // attention filles supprimées dans tobtb au fur et à mesure
                TOBDet.ChangeParent ( FTobTB,-1);
                END;
            END;
          END
       else   // maj de la table tableau de bord
          BEGIN
               // mcd 25/11/02 ne plus utiliser TobTB.InsertDBTable(Nil);
          TobTB.InsertDB(Nil);
          END;
       END;
  finally
    TobTB.Free;
    TobTBModele.free;
  end;
End;


Procedure TParamTBAffaire.AlimTobTBModele(TobTBModele ,TobDetail : Tob;Enr : T_TypeEnr);
Var i : integer;
    Part0, Part1, Part2, Part3, Avenant : string;
Begin
TobTBModele.PutValue('ATB_UTILISATEUR',V_PGI.User);
if Enr=Tan then TobTBModele.PutValue('ATB_TYPEENR', 'AN')
 else if Enr=TMois then TobTBModele.PutValue('ATB_TYPEENR', 'MOI')
 else TobTBModele.PutValue('ATB_TYPEENR', 'NOR');
if (TypeTB = ttbAffaire) or (TypeTB = ttbTiers) then
    BEGIN
    // Champs affaires
    if (TypeTB = ttbAffaire) then
      BEGIN
      // code affaire
      TobTBModele.PutValue('ATB_AFFAIRE', TobDetail.GetValue('AFF_AFFAIRE'));
      TobTBModele.PutValue('ATB_AFFAIRE0', TobDetail.GetValue('AFF_AFFAIRE0'));
      TobTBModele.PutValue('ATB_AFFAIRE1', TobDetail.GetValue('AFF_AFFAIRE1'));
      TobTBModele.PutValue('ATB_AFFAIRE2', TobDetail.GetValue('AFF_AFFAIRE2'));
      TobTBModele.PutValue('ATB_AFFAIRE3', TobDetail.GetValue('AFF_AFFAIRE3'));
      TobTBModele.PutValue('ATB_AVENANT', TobDetail.GetValue('AFF_AVENANT'));
          // mcd 06/03/03
      TobTBModele.PutValue('ATB_DATEDEBUT', TobDetail.GetValue('AFF_DATEDEBUT'));
      TobTBModele.PutValue('ATB_DATEFIN', TobDetail.GetValue('AFF_DATEFIN'));
      TobTBModele.PutValue('ATB_DATELIMITE', TobDetail.GetValue('AFF_DATELIMITE'));
      TobTBModele.PutValue('ATB_DATEGARANTIE', TobDetail.GetValue('AFF_DATEGARANTIE'));
      TobTBModele.PutValue('ATB_DATECLOTTECH', TobDetail.GetValue('AFF_DATECLOTTECH'));
      TobTBModele.PutValue('ATB_DATEDEBEXER', TobDetail.GetValue('AFF_DATEDEBEXER'));
      TobTBModele.PutValue('ATB_DATEFINEXER', TobDetail.GetValue('AFF_DATEFINEXER'));
      TobTBModele.PutValue('ATB_REGROUPEFACT', TobDetail.GetValue('AFF_REGROUPEFACT'));
      // affaire de référence
      if (TobDetail.GetValue('AFF_AFFAIRE') = TobDetail.GetValue('AFF_AFFAIREREF')) then
         begin
         TobTBModele.PutValue('ATB_AFFAIREREF', TobDetail.GetValue('AFF_AFFAIRE'));
         TobTBModele.PutValue('ATB_AFFAIREREF0', TobDetail.GetValue('AFF_AFFAIRE0'));
         TobTBModele.PutValue('ATB_AFFAIREREF1', TobDetail.GetValue('AFF_AFFAIRE1'));
         TobTBModele.PutValue('ATB_AFFAIREREF2', TobDetail.GetValue('AFF_AFFAIRE2'));
         TobTBModele.PutValue('ATB_AFFAIREREF3', TobDetail.GetValue('AFF_AFFAIRE3'));
         TobTBModele.PutValue('ATB_AVENANTREF', TobDetail.GetValue('AFF_AVENANT'));
         end
      else
         begin
         {$IFDEF BTP}
         BTPCodeAffaireDecoupe(TobDetail.GetValue('AFF_AFFAIREREF'), Part0, Part1, Part2, Part3, Avenant, taModif, false);
         {$ELSE}
         CodeAffaireDecoupe(TobDetail.GetValue('AFF_AFFAIREREF'), Part0, Part1, Part2, Part3, Avenant, taModif, false);
         {$ENDIF}
         TobTBModele.PutValue('ATB_AFFAIREREF', TobDetail.GetValue('AFF_AFFAIREREF'));
         TobTBModele.PutValue('ATB_AFFAIREREF0', Part0); TobTBModele.PutValue('ATB_AFFAIREREF1', Part1);
         TobTBModele.PutValue('ATB_AFFAIREREF2', Part2); TobTBModele.PutValue('ATB_AFFAIREREF3', Part3);
         TobTBModele.PutValue('ATB_AVENANTREF', Avenant);
         end;
      // stats
      for i := 1 to 9 do begin TobTBModele.PutValue('ATB_LIBREAFF'+IntTostr(i), TobDetail.GetValue('AFF_LIBREAFF'+IntTostr(i))); end;
      TobTBModele.PutValue('ATB_LIBREAFFA', TobDetail.GetValue('AFF_LIBREAFFA'));
      for i:=1 to 3 do begin TobTBModele.PutValue('ATB_AFFRESSOURCE'+IntTostr(i), TobDetail.GetValue('AFF_RESSOURCE'+IntTostr(i))); end;
      // divers
      TobTBModele.PutValue('ATB_LIBELLEAFF', TobDetail.GetValue('AFF_LIBELLE'));
      TobTBModele.PutValue('ATB_RESPONSABLE', TobDetail.GetValue('AFF_RESPONSABLE'));
      TobTBModele.PutValue('ATB_ADMINISTRATIF', TobDetail.GetValue('AFF_ADMINISTRATIF'));
      TobTBModele.PutValue('ATB_DEPARTEMENT', TobDetail.GetValue('AFF_DEPARTEMENT'));
      TobTBModele.PutValue('ATB_ETABLISSEMENT', TobDetail.GetValue('AFF_ETABLISSEMENT'));
      TobTBModele.PutValue('ATB_TIERS', TobDetail.GetValue('AFF_TIERS'));
      END
    else
      BEGIN      // Client
      TobTBModele.PutValue('ATB_TIERS', TobDetail.GetValue('T_TIERS'));
      END;
    // Champs tiers repris dans les deux cas
    TobTBModele.PutValue('ATB_SOCIETEGROUPE', TobDetail.GetValue('T_SOCIETEGROUPE'));
    TobTBModele.PutValue('ATB_LIBELLETIERS', TobDetail.GetValue('T_LIBELLE'));
    TobTBModele.PutValue('ATB_MOISCLOTURE', TobDetail.GetValue('T_MOISCLOTURE'));
    for i := 1 to 9 do begin TobTBModele.PutValue('ATB_LIBRETIERS'+IntTostr(i), TobDetail.GetValue('YTC_TABLELIBRETIERS'+IntTostr(i))); end;
    TobTBModele.PutValue('ATB_LIBRETIERSA', TobDetail.GetValue('YTC_TABLELIBRETIERSA'));
    TobTBModele.PutValue('ATB_RES1TIERS', TobDetail.GetValue('YTC_RESSOURCE1'));
    TobTBModele.PutValue('ATB_RES2TIERS', TobDetail.GetValue('YTC_RESSOURCE2'));
    TobTBModele.PutValue('ATB_RES3TIERS', TobDetail.GetValue('YTC_RESSOURCE3'));
    END;
End;

//Regarde si l'enrgt que l'on est en train de traiter existe déjà dans la tob
//(mise à jour enrgt ou traitement). Rend la tob trouvée
Function  TParamTBAffaire.VerifCleTobTB (TobTB : TOB; TypeAlim : T_TypePRFA; pTob : Tob) : TOB;
Var i, j,ColDico : integer;
    TobDet : TOB;
    St, St2, Tmp, Prefixe,PrefixeEnc, ChampUnite : string;
    Trouve, bTopPrevu : Boolean;

BEGIN
  Result := Nil;
  Trouve := False;

  // mcd 09/07 if (TypeAlim <> tActivite) then
  if (TypeAlim = TCutOff) or (TypeAlim = TFactEclat) then
    BEGIN
      Prefixe:= 'ACU_';
      st := ListeChampsCutOff; ColDico := nCutOff; ChampUnite := 'UNITE';
    END
  else if (TypeAlim =TBud) then
    BEGIN
      Prefixe:= 'ABU_';
      st := ListeChampsBud; ColDico := nBud; ChampUnite := '';
    END

  // C.B 11/09/02
  // pas d'unite pour le planning
  else if (TypeAlim =TPlanning) then
    BEGIN
      Prefixe:= 'APL_';
      st := ListeChampsPlanning; ColDico := nPlanning; ChampUnite := '';
    END

  else if (TypeAlim =Tligne) then
    BEGIN
      Prefixe:= 'GL_';
      st := ListeChampsLigne; ColDico := nGL; ChampUnite := 'QUALIFQTEVTE';
    END
  else
    BEGIN
      Prefixe:= 'ACT_';
      st := ListeChampsAct;  ColDico := nACT; ChampUnite := 'UNITE';
    END;

  For i:=0 to TobTB.Detail.count-1 do
    BEGIN
      TOBDet := TobTB.Detail[i];
      bTopPRevu := False;

      // mcd : obligation de le refaire, car sinon on a perdu l'info d'origine lors du readtoken st
      if (TypeAlim = TCutOff) or (TypeAlim = TFActEclat) then st := ListeChampsCutOff
      else if (TypeAlim = TBud) then st:=ListeChampsBud
      else if (TypeAlim = Tligne) then
        begin
          st := ListeChampsLigne ;
          // mcd 25/03/02, si cumul par mois ou semaine, il faudra prendre en compte la date de la mission
          // pour l'affectation en date dans le TB et non pas la date de la pièce
          if (TypeDateTb <>TdaCumul) then
            begin
              // if (Q.FindField('GL_NATUREPIECEG').AsString =GetParamSoc('SO_AFNATAFFAIRE'))
              //or (Q.FindField('GL_NATUREPIECEG').AsString =GetParamSoc('SO_AFNATPROPOSITION'))
              if (pTob.GetValue('GL_NATUREPIECEG') = GetParamSoc('SO_AFNATAFFAIRE')) or
                 (pTob.GetValue('GL_NATUREPIECEG') = GetParamSoc('SO_AFNATPROPOSITION')) then
                bTopPrevu:=true;
            end;
        end
      else st := ListeChampsAct;
      While St<>'' do
        BEGIN
          PrefixeEnc := Prefixe;
          Tmp:= ReadTokenSTV(St);
          j:=Pos(Prefixe,tmp) ;
          // Traitement articles ils suivent la logique activité
          if (j = 0) then
           BEGIN
           j:= Pos('GA_',tmp);
           if j > 0 then begin
              PrefixeEnc := 'GA_';
                          //mcd 27/12/02    cas particulie du champ libelle qui a le même nom dans X tables
              if Pos('LIBELLE',tmp)>0 then
                begin
                PrefixeEnc:='';
                j:=0;
                end;
              end
           else begin   // mcd 18/11/02
                j:= Pos('ARS_',tmp);
                if j > 0 then begin
                   PrefixeEnc := 'ARS_';
                         //mcd 27/12/02  cas particulier libelle ressource qui va dans un autre champ de TB
                   if (Pos('LIBELLE2',tmp)>0) or (Pos('LIBELLE',tmp)>0) then
                     begin
                     PrefixeEnc:='';
                     j:=0;
                     end;
                   end;
                end;
           END;

          if j>0 then tmp := Copy (tmp, pos('_',tmp)+1,length(tmp)-pos('_',tmp));
          tmp := Trim(tmp);
          if tmp = 'NATUREPIECEG' then tmp :=''; // mcd 27/12/02 On a besoin si piece achat géré de l'avoir dans le groupe by, mais il ne faut pas l'écrire dans la table
          St2 := ChampsTableVersTB(tmp,ColDico,ttaTableTB);
          if St2 <> '' then
              BEGIN
              if Pos('DATE',tmp) > 0 then  // Champs de type Date
                  BEGIN
                  if btopprevu then
                    if (TobDet.GetValue(PrefixeTB + St2)=DatePrevu ) then Trouve:=True else Trouve:=False
                  else
                    //if (TobDet.GetValue(PrefixeTB + St2)=Q.FindField(Prefixe + tmp).AsDateTime)then Trouve:=True else Trouve:=False;
                    if (TobDet.GetValue(PrefixeTB + St2) = StrToDate(pTob.GetValue(Prefixe + tmp))) then Trouve:=True else Trouve:=False;
                  END
              else  if Pos('PERIODE',tmp) > 0 then
                  BEGIN
                  if btopprevu then
                    if (TobDet.GetValue(PrefixeTB + St2)=GetPeriode(DatePrevu) ) then Trouve:=True else Trouve:=False
                  else
                    //if (TobDet.GetValue(PrefixeTB + St2)=Q.FindField(Prefixe + tmp).AsString)then Trouve:=True else Trouve:=False;
                    if (TobDet.GetValue(PrefixeTB + St2) = varAsType(pTob.GetValue(Prefixe + tmp), varString)) then Trouve:=True else Trouve:=False;
                  END
              else  if Pos('SEMAINE',tmp) > 0 then
                  BEGIN
                  if btopprevu then
                    if (TobDet.GetValue(PrefixeTB + St2)=NumSemaine(DatePrevu) ) then Trouve:=True else Trouve:=False
                  else
                    //if (TobDet.GetValue(PrefixeTB + St2)=Q.FindField(Prefixe + tmp).AsString)then Trouve:=True else Trouve:=False;
                    if (TobDet.GetValue(PrefixeTB + St2) = varAsType(pTob.GetValue(Prefixe + tmp), varInteger)) then Trouve:=True else Trouve:=False;
                  END
              else if Pos(ChampUnite,tmp) > 0 then   // le champs unite n'est pas un critère de distinction
                  Trouve := True
              else // Champs string à traiter par défaut
                BEGIN
                  //if (TobDet.GetValue (PrefixeTB + St2)=Q.FindField(PrefixeEnc + tmp).AsString) then Trouve:=True else Trouve:=False;
                  // on test <> Null pour le cas ou la jointure sur article ou ressource n'existe pas ...
                  If pTob.GetValue(PrefixeEnc + tmp)<>Null then if (TobDet.GetValue (PrefixeTB + St2) = varAsType(pTob.GetValue(PrefixeEnc + tmp), VarString)) then Trouve:=True else Trouve:=False;
                END;
              if Not Trouve then Break;
              END;
          END;
      if Trouve then BEGIN Result := TobDet; Exit; END;
    END;
END;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... :
Modifié le ... :
Description .. : Alimentation dans la TOB TB des informations concernant la
                 table en cours cut off, activite, ligne ….).
                 Regarde si l'enrgt existe déjà pour mise à jour ou création,
                 En fonction de la table en cours de traitement, met à jours
                 les bons champs qté et valeurs dans la table AftableauBord.

Mots clefs ... :
*****************************************************************}
Procedure TParamTBAffaire.AlimTobTB (TobTBDet, pTob : TOB;
                                     MajTobExistante: Boolean; TypeAlim : T_TypePRFA );
Var
  Prefixe, PrefixeQte, st, st2, tmp, ChampUnite, AffectLigne,PrefixeEnc,
  ValoPr,ValoPV,UniteValo,Ress: string;
  i, ColDico :integer;
  PRCharge, PV, Qte, QteUniteRef, AddQte, AddQteUniteRef,
  QteFac,QteFacUniteRef, PVFac, QtePrevu, PVPrevu,
  QtePrevuUniteRef ,DCoeffCOnvert,PR,PV1 : Double;
  EstAvoir,TopPrevu,TopEngage : Boolean;
  DateTB : TDatetime;
  TobValo : Tob;
  bPROK, bPVOK : boolean;
  TypeArt,TypeRes,REs,Art,ARtUni:string; //mcd 15/10/02
  Stodate:Tdatetime; //mcd 15/10/02
Begin

  Qte := 0; QteUniteRef := 0;  PRCharge := 0; PV := 0; AddQte := 0;
  QteFac := 0; QteFacUniteRef := 0; PVFac := 0; AddQteUniteRef := 0;
  QtePrevu := 0; QtePrevuUniteRef := 0; PVPrevu := 0;
  TopPrevu := false;
  stodate := iDate1900;
  if (TypeAlim =TCutOff) or (TypeAlim=TFactEclat) then
    BEGIN
      Prefixe:= 'ACU_';
      st := ListeChampsCutOff; ColDico := nCutOff;  ChampUnite := 'UNITE';
      if (TypeDateTB = tdaDetail ) then PrefixeQte := 'ACU_' else PrefixeQte := 'SUM_';
    END

  else if (TypeAlim=TBud) then
    begin
      Prefixe:= 'ABU_';
      st := ListeChampsBud; ColDico := nBud;  ChampUnite := '';
      if (TypeDateTB = tdaDetail ) then PrefixeQte := 'ABU_' else PrefixeQte := 'SUM_';
    end

  // C.B 11/09/02
  else if (TypeAlim = TPLanning) then
    begin
      Prefixe:= 'APL_';
      st := ListeChampsPLanning; ColDico := nPLanning;  ChampUnite := '';
      if (TypeDateTB = tdaDetail ) then PrefixeQte := 'APL_' else PrefixeQte := 'SUM_';
    end

  else if (TypeAlim =Tligne) then
    BEGIN
      Prefixe:= 'GL_';
      st := ListeChampsLigne; ColDico := nGL;  ChampUnite := 'QUALIFQTEVTE';
      if (TypeDateTB = tdaDetail ) then PrefixeQte := 'GL_' else PrefixeQte := 'SUM_';
      // mcd 25/03/02, si cumul par mois ou semaine, il faudra prendre en compte la date de la mission
      // et non pas date de la pièce pour affectation à la bonne date
      if (TypeDateTb <>TdaCumul) then begin
        //if (Q.FindField('GL_NATUREPIECEG').AsString =GetParamSoc('SO_AFNATAFFAIRE'))
        //or (Q.FindField('GL_NATUREPIECEG').AsString =GetParamSoc('SO_AFNATPROPOSITION'))
        if (pTob.GetValue('GL_NATUREPIECEG') = GetParamSoc('SO_AFNATAFFAIRE')) or
           (pTob.GetValue('GL_NATUREPIECEG') = GetParamSoc('SO_AFNATPROPOSITION')) then
          TopPrevu := true;
       end;

    END

  else
    BEGIN
      Prefixe:= 'ACT_';
      st := ListeChampsAct;  ColDico := nACT;  ChampUnite := 'UNITE';
      if (TypeDateTB = tdaDetail ) then PrefixeQte := 'ACT_' else PrefixeQte := 'SUM_';
    END;
if ((TypeAlim=tactivite)  OR (typeAlim=tBoni)) and (AlimValoREss_Art) then
  begin  // mcd 15/10/2002 si recalul PV , on lit sur détail l'activite, mais il faut écrire en fct détail choisi
  If TypedateTb=TdaCumul then begin
     StoDate:=Ptob.GetValue('ACT_DATEACTIVITE');
     Ptob.PutValue('ACT_DATEACTIVITE',Idate1900);
     end;
     // mcd 30/10/02 ajout tniTous
  if (NiveauTb<>TniTypeart) and (NiveauTb<>TniARt) and (NiveauTb<>TniTypeArtREs) and (NiveauTb<>TniTous) then begin
   TypeArt:=Ptob.GetValue('ACT_TYPEARTICLE');
   Ptob.PutValue('ACT_TYPEARTICLE','');
   end;
  if (NiveauTb<>Tniart) and (NiveauTb<>TniARtRes) and (NiveauTb<>TniTous)  then begin
   Art:=Ptob.GetValue('ACT_CODEARTICLE');
   Ptob.PutValue('ACT_CODEARTICLE','');
   ArtUni:=Ptob.GetValue('ACT_ARTICLE');
   Ptob.PutValue('ACT_ARTICLE','');
   end;
  if (NiveauTb<>TniartREs) and (NiveauTb<>TniRes) and (NiveauTb<>TniTypeRes) And( NiveauTb<>TniTous)then begin
   TypeRes:=Ptob.GetValue('ACT_TYPERESSOURCE');
   Ptob.PutValue('ACT_TYPERESSOURCE','');
   end;
  if (NiveauTb<>TniRes) and (NiveauTb<>TniArtREs) and( NiveauTb<>TniTous) then begin
   Res:=Ptob.GetValue('ACT_RESSOURCE');
   Ptob.PutValue('ACT_RESSOURCE','');
   end;
  end;

  if Not(MajTobExistante) then
      BEGIN
      While St<>'' do
          BEGIN
          Tmp:= ReadTokenSTV(St);
          i:=Pos(Prefixe,tmp) ;
          if i > 0 then
              BEGIN
              tmp := Copy (tmp, pos('_',tmp)+1,length(tmp)-pos('_',tmp));
              PrefixeEnc := Prefixe;
              END
          else
              BEGIN
              // possible par activité + champs articles
              i:=Pos('GA',tmp);
              if i > 0 then begin
                     PrefixeEnc := 'GA_';
                     if Pos('LIBELLE',tmp)>0 then
                        begin //mcd 18/11/02    cas particulie du champ libelle qui a le même nom dans X tables
                        PrefixeEnc:='';
                        i:=0;
                        end
                     end
               else  begin
                      i:=Pos('ARS',tmp);
                      if i >0 then begin
                         PrefixeEnc:='ARS_';
                         //mcd 18/11/02
                         if (Pos('LIBELLE2',tmp)>0) or (Pos('LIBELLE',tmp)>0) then
                           begin
                           PrefixeEnc:='';
                           i:=0;
                           end;
                         end;
                      end;
              if i > 0 then tmp := Copy (tmp, pos('_',tmp)+1,length(tmp)-pos('_',tmp)); 
              END;
          tmp := Trim(tmp);      
          if (tmp = 'NATUREPIECEG') and (TypeDateTB <> Tdadetail) then tmp :=''; // mcd 27/12/02 On a besoin si piece achat géré de l'avoir dans le groupe by, mais il ne faut pas l'écrire dans la table , sauf si détail, pour N° facture dans GL
          st2 := ChampsTableVersTB(tmp,ColDico,ttaTableTB);
          if st2 <> '' then
              BEGIN      //mcd 25/03/02 ajout test sur prévu pour chnager date
              if Pos('DATE',tmp) > 0  then
                 begin
                 if topprevu then
                     TobTbDet.PutValue(PrefixeTB + St2,DatePrevu )
                else
                   //TobTBDet.PutValue (PrefixeTB + St2, Q.FindField(PrefixeEnc + tmp).AsDateTime )
                   TobTBDet.PutValue (PrefixeTB + St2, varAsType(pTob.GetValue(PrefixeEnc + tmp), varDate))
                 end
              else if Pos('SEMAINE',tmp) > 0  then
                 begin
                 if topprevu then
                   TobTbDet.PutValue(PrefixeTB + St2,NumSemaine(DatePrevu ))
                 else
                   TobTBDet.PutValue (PrefixeTB + St2, varAsType(pTob.GetValue(PrefixeEnc + tmp), varInteger))
                 end
              else if Pos('PERIODE',tmp) > 0  then
                 begin
                 if topprevu then
                   begin
                   TobTbDet.PutValue(PrefixeTB + St2,GetPeriode(DatePrevu ));
                   pTob.PutValue(PrefixeEnc + St2,GetPeriode(DatePrevu )); //mcd 18/06/03 sinon libperiode faux sur ligne de prévu
                   end
                 else
                   //TobTBDet.PutValue (PrefixeTB + St2, Q.FindField(PrefixeEnc + tmp).AsSTRING )
                   TobTBDet.PutValue (PrefixeTB + St2, pTob.GetValue(PrefixeEnc + tmp));
                 end
              else
                begin
                  // on test <> Null pour le cas ou la jointure sur article ou ressource n'existe pas ...
                  If pTob.GetValue(PrefixeEnc + tmp)<>Null then TobTBDet.PutValue (PrefixeTB + St2,  varAsType(pTob.GetValue(PrefixeEnc + tmp), VarString));
                end;

              if Pos('PERIODE',tmp)>0 then
                //TobTBDet.PutValue (PrefixeTB + 'LIBPERIODE', OLEFormatPeriode (Q.FindField(PrefixeEnc + tmp).AsString,True));
                {$IFDEF BTP}
                TobTBDet.PutValue (PrefixeTB + 'LIBPERIODE', BTPOLEFormatPeriode (varAsType(pTob.GetValue(PrefixeEnc + tmp), VarString),True));
                {$ELSE}
                TobTBDet.PutValue (PrefixeTB + 'LIBPERIODE', OLEFormatPeriode (varAsType(pTob.GetValue(PrefixeEnc + tmp), VarString),True));
                {$ENDIF}
              END;
          End;
      if (TypeAlim =Tligne)or (TypeAlim=TFactEclat) then
        BEGIN
          if ((NiveauTB=tniTypeArt) or (NiveauTB=tniArt) or (NiveauTB=tniTypeArtRes) or (NiveauTB=tniArtRes)) then
            BEGIN  // test car dans ces modes de regroupement pas de type article
              //if (Q.FindField(Prefixe + 'TYPEARTICLE').AsString= 'PRE') and (TypeDateTB <> tdaDetail) then
              if (pTob.GetValue(Prefixe + 'TYPEARTICLE') = 'PRE') and (TypeDateTB <> tdaDetail) then
                TobTBDet.PutValue('ATB_UNITE' , UniteTemps );   // par defaut
            END;
         END;
      if GereFacturable then // traitement facturable
          BEGIN
          if (TypeAlim =Tligne) then
              TobTBDet.PutValue('ATB_ACTIVITEREPRIS','F');  // par defaut sur les lignes
 (* mcd 02/04/2003  pour ne plus faire un cumul sur F si facturable ..
  on distingue tous les codes différents
        else if (TobTBDet.GetValue('ATB_ACTIVITEREPRIS') <> 'N') then
              TobTBDet.PutValue('ATB_ACTIVITEREPRIS','F')*) ;
          END;

      if (TypeDateTB =tdaSemaine) or (TypeDateTB = tdaMois ) then
          BEGIN// Champs date alimentée dans tous les cas au premier jour semaine ou mois
          DateTB := iDate1900;
          if (TypeDateTB =tdaSemaine) then DateTB := GetDateDebutSemaine ( TobTBDet.GetValue('ATB_SEMAINE'),TobTBDet.GetValue('ATB_PERIODE')) else
          if (TypeDateTB =tdaMois)    then DateTB := GetDateDebutPeriode ( TobTBDet.GetValue('ATB_PERIODE'));
          TobTBDet.PutValue('ATB_DATE', DateTB);
          END;

      END;
if ((TypeAlim=tactivite) or (typeAlim=tBoni)) and (AlimValoREss_Art) then
  begin  // mcd 15/10/2002 il faut remettre la bonne date pour calcul prix
  If TypedateTb=TdaCumul then Ptob.PutValue('ACT_DATEACTIVITE',StoDate);
  if (NiveauTb<>TniTypeart) and (NiveauTb<>TniARt) and (NiveauTb<>TniTypeArtREs)  and (NiveauTb<>TniTous) then Ptob.PutValue('ACT_TYPEARTICLE',TypeArt);
  if (NiveauTb<>Tniart) and (NiveauTb<>TniARtRes)  and (NiveauTb<>TniTous) then
    begin
    ptob.PutValue('ACT_CODEARTICLE',Art);
    ptob.PutValue('ACT_ARTICLE',ArtUni);
    end;
  if (NiveauTb<>TniartREs) and (NiveauTb<>TniRes) and (NiveauTb<>TniTypeRes)  and (NiveauTb<>TniTous)then Ptob.PutValue('ACT_TYPERESSOURCE',TypeRes);
  if (NiveauTb<>TniRes) and (NiveauTb<>TniArtREs)  and (NiveauTb<>TniTous) then Ptob.PutValue('ACT_RESSOURCE',Res);
  end ;

  // Traitement des quantités avec conversion
  // mcd ??? if (TypeAlim <=Tligne) then AddQte := Q.FindField(PrefixeQte + 'QTEFACT').AsFloat
  If (TypeAlim=TActivite) and (NatAchatEngage <>'') then
    begin
      TopEngage:=False;
      //If pos (Q.FindField('ACT_NATUREPIECEG').AsString, NatAchatEngage)>0 then TopEngage:=True
      If pos (pTob.GetValue('ACT_NATUREPIECEG'), NatAchatEngage)>0 then TopEngage := True
    end
  else
    TopEngage:=False;

  if ((TypeAlim =Tligne) or (TypeAlim=TFactEclat)) then
    //AddQte := Q.FindField(PrefixeQte + 'QTEFACT').AsFloat
    AddQte := valeur(pTob.GetValue(PrefixeQte + 'QTEFACT'))

  // C.B 19/11/02
  else if (TypeAlim = TPlanning) then
    AddQte := valeur(pTob.GetValue(PrefixeQte + 'QTEPLANIFIEE'))

  else if (TypeAlim <>tCutOff) then
    //AddQte := Q.FindField(PrefixeQte + 'QTE').AsFloat;
    AddQte := valeur(pTob.GetValue(PrefixeQte + 'QTE'));

  if ((TypeAlim<>TCutOff) and (TypeAlim<>TfactEclat)and (TypeAlim<>TBud)
      and (TypeAlim <> TPlanning))  then // C.B 19/11/02
    BEGIN            // pas de qte pour cutofff  // pas d'unité pour budget // pas d'unite pour planning

      //if (Q.FindField(Prefixe + ChampUnite).AsString <> UniteTemps) then
      if (pTob.GetValue(Prefixe + ChampUnite) <> UniteTemps) then
        BEGIN
          //if IsConvertible(Q.FindField(Prefixe + ChampUnite).AsString, UniteTemps) then
          if IsConvertible(pTob.GetValue(Prefixe + ChampUnite), UniteTemps) then
            BEGIN
              if ((NiveauTB <> tniTypeArt) And (NiveauTB <> tniArt) And (NiveauTB <> tniTypeArtRes) And (NiveauTB <> tniArtRes)) then
                 //AddQteUniteRef:=ConversionUnite(Q.FindField(Prefixe + ChampUnite).AsString,UniteTemps,AddQte)
                 AddQteUniteRef := ConversionUnite(pTob.GetValue(Prefixe + ChampUnite), UniteTemps, AddQte)
              else
                BEGIN
                  //if (Q.FindField(Prefixe + 'TYPEARTICLE').AsString= 'PRE') then
                  //  AddQteUniteRef:=ConversionUnite(Q.FindField(Prefixe + ChampUnite).AsString,UniteTemps,AddQte);
                  if (pTob.GetValue(Prefixe + 'TYPEARTICLE') = 'PRE') then
                    AddQteUniteRef := ConversionUnite(pTob.GetValue(Prefixe + ChampUnite), UniteTemps, AddQte);
                END;

              if (TypeDateTB <> tdaDetail) then
                BEGIN
                  AddQte := AddQteUniteRef ; // Si pas de détail Qté toujours dans l'unité de réf.
                  if not (TopEngage) then TobTBDet.PutValue('ATB_UNITE' , UniteTemps );
                END;
            END

          // si pas de convertion possible, on ne fait rien dans qteref
          // mais en +, si niveau regroupement par ressource, on ajoute pas dans zone qte (risque
          // d'ajouter des Fou+frais+Pres
          //else if  (Q.FindField(Prefixe + ChampUnite).AsString='') and
          //         ((NiveauTb=Tnires) or (NiveauTb=tnityperes) or (niveautb=tnifonction)) then
          else if  (pTob.GetValue(Prefixe + ChampUnite) = '') and
                   ((NiveauTb=Tnires) or (NiveauTb=tnityperes) or (niveautb=tnifonction)) then
            addqte:=0                
          else if TopEngage then
            AddQteUniteRef := AddQte;

        END
      else AddQteUniteRef := AddQte; // a l'identique

    END;// fin traitement des zones QTE

  //if (TypeAlim =Tligne) then AffectLigne := GetInfoParPiece(Q.FindField('GL_NATUREPIECEG').AsString,'GPP_AFAFFECTTB');
  if (TypeAlim =Tligne) then AffectLigne := GetInfoParPiece(pTob.GetValue('GL_NATUREPIECEG'),'GPP_AFAFFECTTB');

  if MajTobExistante then                                  
      BEGIN
      if (TypeAlim =Tligne) then  // traitement requete ligne
          BEGIN
          if AffectLigne = 'FAC' then
             BEGIN
             QteFac:= TobTBDet.GetValue('ATB_FACQTE'); PVFac := TobTBDet.GetValue('ATB_FACPV');
             QteFacUniteRef := TobTBDet.GetValue('ATB_FACQTEUNITEREF');
             END else
          if AffectLigne = 'PRE' then
             BEGIN
             QtePrevu := TobTBDet.GetValue('ATB_PREVUQTE'); PVPrevu := TobTBDet.GetValue('ATB_PREVUPV');
             QtePrevuUniteRef := TobTBDet.GetValue('ATB_PREQTEUNITEREF');
             END else
      (* mcd 07/08/02 pas possible en saisie de pièce ... à voir si mis en place ...    if AffectLigne = 'ACH' then
             BEGIN
             QteAch := TobTBDet.GetValue('ATB_PREVUQTE'); PRAch := TobTBDet.GetValue('ATB_FACPV');
             END;  *)

          END
      else
          BEGIN
          if (TypeAlim =TCutOff) then begin
             QteUniteRef := 0;
             if pTob.GetValue('ACU_TYPEAC')='CVE' then begin
               Pv := TobTBDet.GetValue('ATB_FAE');
               Qte := TobTBDet.GetValue('ATB_PCA');
               PRCharge    := TobTBDet.GetValue('ATB_AAE');
               end
             else begin
               Pv := TobTBDet.GetValue('ATB_FAR');
               Qte := TobTBDet.GetValue('ATB_CCA');
               PRCharge    := TobTBDet.GetValue('ATB_AAR');
               end;
             end

         // C.B 11/09/02
         else if (TypeAlim =TBud) or (TypeAlim =TPLanning) then
           begin
             QteUniteRef := TobTBDet.GetValue('ATB_PREVUQTE');
             PRCharge    := TobTBDet.GetValue('ATB_PREVUPR');
             Pv := TobTBDet.GetValue('ATB_PREVUPV');
             Qte := TobTBDet.GetValue('ATB_PREVUQTE');
           end

          Else if (TypeAlim =TFactEclat) then begin
             QteUniteRef := 0;
             PRCharge    := 0;
             PvFac := TobTBDet.GetValue('ATB_FACPV');
             Qte := 0;
             QteFac:= TobTBDet.GetValue('ATB_FACQTe');
             end
          else if (TypeAlim =Tboni) then begin
             Qte:= TobTBDet.GetValue('ATB_BONIQTE');   QteUniteRef := TobTBDet.GetValue('ATB_BONIQTEUNITERE');
             PRCharge    := TobTBDet.GetValue('ATB_BONIPR');
             PV := TobTBDet.GetValue('ATB_BONIVENTE');
             end
          else begin
             if TopEngage then begin
                  Qte:= TobTBDet.GetValue('ATB_ENGAGEQTE');
                  QteUniteRef := TobTBDet.GetValue('ATB_ENGAGEQTE');
                  PRCharge    := TobTBDet.GetValue('ATB_ENGAGEPR');
                  PV := TobTBDet.GetValue('ATB_ENGAGEPV');
                end
             else begin
                  Qte:= TobTBDet.GetValue('ATB_QTE');   QteUniteRef := TobTBDet.GetValue('ATB_QTEUNITEREF');
                  PRCharge    := TobTBDet.GetValue('ATB_TOTPRCHARGE');
                  PV := TobTBDet.GetValue('ATB_TOTVENTEACT');
                  end;
             end;
          END
      END;

  if (TypeAlim =Tligne) then
    BEGIN

      //if (GetInfoParPiece(Q.FindField('GL_NATUREPIECEG').AsString,'GPP_ESTAVOIR')='X') then
      if (GetInfoParPiece(pTob.GetValue('GL_NATUREPIECEG'),'GPP_ESTAVOIR')='X') then
        Estavoir:= True
      else
        EstAvoir := False;

      if AffectLigne = 'FAC' then
        BEGIN
          QteFac := QteFac + AddQte; QteFacUniteRef:= QteFacUniteRef+ AddQteUniteRef;

          if EstAvoir then
            BEGIN   // les avoirs sont en négatifs
              //PVFac  := PVFac  + Q.FindField(PrefixeQte +'TOTALHT').AsFloat;
              //SyntheseTB.TotFact := SyntheseTB.TotFact + Q.FindField(PrefixeQte +'TOTALHT').AsFloat;
              PVFac  := PVFac  + valeur(pTob.GetValue(PrefixeQte +'TOTALHT'));
              SyntheseTB.TotFact := SyntheseTB.TotFact + valeur(pTob.GetValue(PrefixeQte +'TOTALHT'));
            END
          else
            BEGIN
              PVFac  := PVFac  + valeur(pTob.GetValue(PrefixeQte +'TOTALHT'));
              SyntheseTB.TotFact := SyntheseTB.TotFact + valeur(pTob.GetValue(PrefixeQte +'TOTALHT'));
            END;

          TobTBDet.PutValue('ATB_FACQTE', QteFac   ); TobTBDet.PutValue('ATB_FACQTEUNITEREF' , QteFacUniteref );
          TobTBDet.PutValue('ATB_FACPV' , PVFac );
        END

      else if AffectLigne = 'PRE' then
        BEGIN
          QtePrevu := QtePrevu + AddQte;  QtePrevuUniteRef:= QtePrevuUniteref+ AddQteUniteRef;
          PVPrevu  := PVPrevu  + valeur(pTob.GetValue(PrefixeQte +'TOTALHT'));
          SyntheseTB.TotPrevuPV := SyntheseTB.TotPrevuPV + valeur(pTob.GetValue(PrefixeQte +'TOTALHT'));
          TobTBDet.PutValue('ATB_PREVUQTE', QtePrevu   );
          TobTBDet.PutValue('ATB_PREQTEUNITEREF' , QtePrevuUniteRef );
          TobTBDet.PutValue('ATB_PREVUPV' , PVPrevu );
        END
    END

  else

    BEGIN

      if (TypeAlim =TCutOff) then
        begin
          if pTob.GetValue('ACU_TYPEAC')='CVE' then begin  // traitement cut off vente
            PRCharge   := PRCharge   + valeur(pTob.GetValue(PrefixeQte +'AAE'));
            PV  := PV + valeur(pTob.GetValue(PrefixeQte +'FAE'));
            Qte:= Qte + valeur(pTob.GetValue(PrefixeQte +'PCA'));
            SyntheseTB.TotFAE:= SyntheseTB.TotFAE + valeur(pTob.GetValue(PrefixeQte +'FAE'));
            SyntheseTB.TotAAE:= SyntheseTB.TotAAE + valeur(pTob.GetValue(PrefixeQte +'AAE'));
            SyntheseTB.TotPCA:= SyntheseTB.TotPCA + valeur(pTob.GetValue(PrefixeQte +'PCA'));

            TobTBDet.PutValue('ATB_AAE' ,PRCharge);
            TobTBDet.PutValue('ATB_FAE' ,PV);
            TobTBDet.PutValue('ATB_PCA' ,Qte);
            end
          else begin     // triatement cut off achat
            PRCharge   := PRCharge   + valeur(pTob.GetValue(PrefixeQte +'AAE'));
            PV  := PV + valeur(pTob.GetValue(PrefixeQte +'FAE'));
            Qte:= Qte + valeur(pTob.GetValue(PrefixeQte +'PCA'));
            SyntheseTB.TotFAR:= SyntheseTB.TotFAR + valeur(pTob.GetValue(PrefixeQte +'FAE'));
            SyntheseTB.TotAAR:= SyntheseTB.TotAAR + valeur(pTob.GetValue(PrefixeQte +'AAE'));
            SyntheseTB.TotCCA:= SyntheseTB.TotCCA + valeur(pTob.GetValue(PrefixeQte +'PCA'));
            TobTBDet.PutValue('ATB_AAR' ,PRCharge);
            TobTBDet.PutValue('ATB_FAR' ,PV);
            TobTBDet.PutValue('ATB_CCA' ,Qte);
            end;
        end //fin traitement cut off

      else if (TypeAlim =TBud) then
        begin
          PRCharge   := PRCharge + valeur(pTob.GetValue(PrefixeQte +'MTPRBUD'));
          PV  := PV + valeur(pTob.GetValue(PrefixeQte +'MTPVBUD'));
          Qte:= Qte + valeur(pTob.GetValue(PrefixeQte +'QTE'));
            // mcd 14/04/03 ajout des tests pour ne pas cumulé si pas option param
          If (GetParamSoc('SO_AFTYPEPREVUPR')='BUD') then SyntheseTB.TotPrevuPr := SyntheseTB.TotPrevuPR + valeur(pTob.GetValue(PrefixeQte +'MTPRBUD'));
          If (GetParamSoc('SO_AFTYPEPREVUPV')='BUD') then SyntheseTB.TotPrevuPV := SyntheseTB.TotPrevuPV + valeur(pTob.GetValue(PrefixeQte +'MTPVBUD'));
          If (GetParamSoc('SO_AFTYPEPREVUPR')='BUD') then TobTBDet.PutValue('ATB_PREVUPR', PRCharge);
          If (GetParamSoc('SO_AFTYPEPREVUPV')='BUD') then TobTBDet.PutValue('ATB_PREVUPV', PV);
          TobTBDet.PutValue('ATB_PREVUQTE', Qte);
          TobTBDet.PutValue('ATB_PREQTEUNITEREF', Qte);
        end //fin traitement budget

      else if (TypeAlim =TPlanning) then
        begin
          PRCharge   := PRCharge + valeur(pTob.GetValue(PrefixeQte +'INITPTPR'));
          PV  := PV + valeur(pTob.GetValue(PrefixeQte +'INITPTVENTEHT'));
          Qte:= Qte + valeur(pTob.GetValue(PrefixeQte +'QTEPLANIFIEE'));
            //mcd 14/04/03 ajout des test pour ne pas cumulé si pas otpion param
          If (GetParamSoc('SO_AFTYPEPREVUPR')='PLA') then SyntheseTB.TotPrevuPr := SyntheseTB.TotPrevuPR + valeur(pTob.GetValue(PrefixeQte +'INITPTPR'));
          If (GetParamSoc('SO_AFTYPEPREVUPV')='PLA') then SyntheseTB.TotPrevuPV := SyntheseTB.TotPrevuPV + valeur(pTob.GetValue(PrefixeQte +'INITPTVENTEHT'));
          If (GetParamSoc('SO_AFTYPEPREVUPR')='PLA') then TobTBDet.PutValue('ATB_PREVUPR', PRCharge);
          If (GetParamSoc('SO_AFTYPEPREVUPV')='PLA') then TobTBDet.PutValue('ATB_PREVUPV', PV);
          TobTBDet.PutValue('ATB_PREVUQTE', Qte);
          TobTBDet.PutValue('ATB_PREQTEUNITEREF', Qte);
        end //fin traitement planning

      else if (TypeAlim =TFactEclat) then
        begin
          QteFac  := PVFac  + valeur(pTob.GetValue(PrefixeQte +'QTEFACT'));
          PVFac  := PVFac  + valeur(pTob.GetValue(PrefixeQte +'PVFACT'));
          SyntheseTB.TotFact := SyntheseTB.TotFact + valeur(pTob.GetValue(PrefixeQte +'PVFACT'));

          TobTBDet.PutValue('ATB_FACQTEUNITEREF' , QteFac );
          TobTBDet.PutValue('ATB_FACPV' ,PVFAC);
          TobTBDet.PutValue('ATB_FACQTE' ,QteFac);
        end //fin traitement Fact eclatée

    else if (TypeAlim =TBoni) then
      begin
        Qte        := Qte        + AddQte;
        QteUniteRef:= QteUniteref+ AddQteUniteRef;
        PRCharge   := PRCharge   + valeur(pTob.GetValue(PrefixeQte +'TOTPRCHARGE'));
        PV         := PV         + valeur(pTob.GetValue(PrefixeQte +'TOTVENTE'   ));
        if PrefixeQte ='SUM_' then
          begin
            SyntheseTB.totBonipr := SyntheseTB.TotRealPR + valeur(pTob.GetValue(PrefixeQte +'TOTPRCHARGE'));
            SyntheseTB.TotBoniPV := SyntheseTB.TotRealPV + valeur(pTob.GEtValue(PrefixeQte +'TOTVENTE'));
          end

        else
          begin
            SyntheseTB.TotBoniPR := SyntheseTB.TotRealPR + valeur(pTob.GetValue(PrefixeQte +'TOTPRCHARGE'));
            SyntheseTB.TotBoniPV := SyntheseTB.TotRealPV + valeur(pTob.GetValue(PrefixeQte +'TOTVENTE'));
          end;

        TobTBDet.PutValue('ATB_BONIQTE'         , Qte   );
        TobTBDet.PutValue('ATB_BONIQTEUNITERE' , QteUniteref );
        TobTBDet.PutValue('ATB_BONIPR' ,PRCharge);
        TobTBDet.PutValue('ATB_BONIVENTE' ,PV );
      end //fin traitement boni

    // ligne activité normale
    else
      begin
        Qte         := Qte         + AddQte;
        QteUniteRef := QteUniteref + AddQteUniteRef;
        PRCharge   := PRCharge   + valeur(pTob.GetValue(PrefixeQte +'TOTPRCHARGE'));
        PV         := PV         + valeur(pTob.GetValue(PrefixeQte +'TOTVENTE'   ));
        if TopEngage then
          begin
            TobTBDet.PutValue('ATB_ENGAGEQTE' , QteUniteref );
            TobTBDet.PutValue('ATB_ENGAGEPR' ,PRCharge);
            TobTBDet.PutValue('ATB_ENGAGEPV' ,PV );
          end
        else
          begin
            SyntheseTB.TotRealPR := SyntheseTB.TotRealPR + valeur(pTob.GetValue(PrefixeQte +'TOTPRCHARGE'));
            SyntheseTB.TotRealPV := SyntheseTB.TotRealPV + valeur(pTob.GetValue(PrefixeQte +'TOTVENTE'));
            TobTBDet.PutValue('ATB_QTE'         , Qte   ); TobTBDet.PutValue('ATB_QTEUNITEREF' , QteUniteref );
            TobTBDet.PutValue('ATB_TOTPRCHARGE' ,PRCharge);
            TobTBDet.PutValue('ATB_TOTVENTEACT' ,PV );
          end;
         // mcd 05/08/2002 ajout de l'alimentation autre type de valorisation
        If ALimValoRess_ART then
          begin
            TobValo:=Nil;
            If GetParamSoc('SO_AFVALOACTPR')='RES' then
              TobTBDet.putValue('ATB_PRRESS', TobTBDet.GetValue('ATB_TOTPRCHARGE'))
            else
              // on calcul le PR RESSOURCE
              ValoPr:='RES';
            If GetParamSoc('SO_AFVALOACTPR')='ART' then
              TobTBDet.putValue('ATB_PRART', TobTBDet.GetValue('ATB_TOTPRCHARGE'))
            else   ValoPr:='ART';   // on calcul le PR Article
            If GetParamSoc('SO_AFVALOACTPV')='RES' then
              TobTBDet.putValue('ATB_PVRESS', TobTBDet.GetValue('ATB_TOTVENTEACT'))
            else   ValoPV:='RES';     // on calcul le PV RESSOURCE
            If GetParamSoc('SO_AFVALOACTPV')='ART' then
              TobTBDet.putValue('ATB_PVART', TobTBDet.GetValue('ATB_TOTVENTEACT'))
            else   ValoPV:='ART'; // on calcul le PV ARticle
            // on calcul les nouveaux prix unitaire
            If TobTBdet.GetValue('ATB_QTE')<>0  then
              begin
                Ress:=  pTob.GetValue('ACT_RESSOURCE');
                AFOAssistants.AddRessource(Ress);
                // PL le 19/09/02 : ajout des variables de sortie : bPROK, bPVOK pour savoir si les valeurs rendues dans
                // la TOBValo sont fiables ou non.
                bPROK:=true; bPVOK:=true;
                TOBValo:=MajTOBValo(varAsType(pTob.GetValue('ACT_DATEACTIVITE'), varDate), tacGlobal,
                           TobTBDet.GetValue('ATB_AFFAIRE'),  Ress,
                           varAsType(pTob.GetValue('ACT_CODEARTICLE'), varString), TobTBDet.GetValue('ATB_UNITE'),
                           TOBArticles, nil, nil, AFOAssistants, false, ValoPR, ValoPV, bPROK, bPVOK);
                // VALORISATION PAR AFFAIRE : ATTENTION : le jour où l'on implémente cette valorisation, il faudra
                // tester les booleens bPROK, bPVOK en sortie de la fonction MajTOBValo. En effet, en valorisation par
                // affaire il se peut que plusieurs lignes de cette l'affaire en cours aient le même article.
                // On ne peut faire un choix pour savoir sur quel article valoriser (on a choisi l'option 'sans zoom de choix'
                // dans l'appel à MajTOBValo), on renvoie donc bPROK et/ou bPVOK = false pour le signifier...
          if (TOBValo<>nil)  then
                  begin
                    // L'unité de valorisation est l'unité dans laquelle on a valorisé les prix dans la fonction MajTOBValo
                  UniteValo:=TOBValo.GetValue('GA_QUALIFUNITEACT');
                  If UniteValo='' then UniteValo:=TobTBDET.GetValue('ATB_UNITE');
                  if (TobTBDET.GetValue('ATB_UNITE')='') or  (IsConvertible(TobTBDET.GetValue('ATB_UNITE'), UniteValo)) then
                      begin
                      dCoeffConvert:=1;
                       // On calcule le coefficient de conversion pour passer de l'unité de saisie en unité de valorisation
                      if (UniteValo<>TobTbDEt.GetValue('ATB_UNITE')) and (UniteValo<>'') then
                          begin
                            // Les prix fournis sont en unité de saisie, il faut les convertir en unité de valorisation
                          dCoeffConvert := ConversionUnite( UniteValo, TobTbDEt.GetValue('ATB_UNITE'), 1);
                          if dCoeffConvert=0 then dCoeffConvert:=1;
                          end;
                      PR:=(TobTBDET.GetValue('ATB_QTE')/dCoeffConvert)*TOBValo.GetValue('GA_PMRP');
                      PV1:=(TobTBdet.GetValue('ATB_QTE')/dCoeffConvert)*TOBValo.GetValue('GA_PVHT');
                      if ValoPr ='RES' then
                        begin
                          If Ress <> '' then
                              TobTBDet.putValue('ATB_PRRESS',PR)
                          else
                              TobTBDet.putValue('ATB_PRRESS',TobTBDet.GetValue('ATB_TOTPRCHARGE'));    // si pas ressource, il faut conserver prix d'origine.
                          end
                        else
                          TobTBDet.putValue('ATB_PRART',PR);

                          if ValoPV ='RES' then
                            begin
                            If Ress <> '' then
                              TobTBDet.putValue('ATB_PVRESS',PV1)
                            else
                              TobTBDet.putValue('ATB_PVRESS',TobTBDet.GetValue('ATB_TOTVENTEACT'));    // si pas ressource, il faut conserver prix d'origine.
                             end
                           else TobTBDet.putValue('ATB_PVART',PV1);
                        end;
                     end
                  else
                     begin // cas ou qte à 0, on conserve le PRix d'origine
                     If valoPR='RES' then  TobTBDet.putValue('ATB_PRRESS',TobTBDet.GetValue('ATB_TOTPRCHARGE'))
                       else TobTBDet.putValue('ATB_PRART',TobTBDet.GetValue('ATB_TOTPRCHARGE'));
                     If ValoPV='RES' then   TobTBDet.putValue('ATB_PVRESS',TobTBDet.GetValue('ATB_TOTVENTEACT') )
                       else  TobTBDet.putValue('ATB_PVART',TobTBDet.GetValue('ATB_TOTVENTEACT'));
                     end;
                 end;
             if (TOBValo<>nil) then begin TOBValo.Free; end; //mcd 30/10/02
             end;
          end; // fin ALimValoRess_ART
       END;
    End;

Procedure TParamTBAffaire.TraiteComplLigneTB(TobTB:TOB;Enr : T_TypeEnr);
Var i : integer;
    TobDet : TOB;
    stSQL : String;
    Q : TQuery;
begin
for i := 0 to TobTB.Detail.count-1 do
    BEGIN
    TobDet := TobTB.Detail[i];
    if (ComplArt) And (TobDet.GetValue('ATB_CODEARTICLE')<> '') then
      begin
      stSQL := FabriqueRequete ('GA',TobDet.GetValue('ATB_ARTICLE'),Enr);
      Q := OpenSQL (stSQL,True,-1,'',true);
      if Not Q.EOF then AlimComplLigneTB('GA',TobDet,Q);
      Ferme(Q);
      end;
    if (ComplRes) And (TobDet.GetValue('ATB_RESSOURCE')<> '') then
      begin
      stSQL := FabriqueRequete ('ARS',TobDet.GetValue('ATB_RESSOURCE'),Enr);
      Q := OpenSQL (stSQL,True,-1,'',true);
      if Not Q.EOF then AlimComplLigneTB('ARS',TobDet,Q);
      Ferme(Q);
      end;
    END;
end;

Procedure TParamTBAffaire.AlimComplLigneTB(TypeT : string; TobDetTB : TOB; Q: TQuery);
begin
if typeT ='ARS' then
   begin
   TobDetTB.PutValue('ATB_TYPERESSOURCE', Q.FindField('ARS_TYPERESSOURCE').AsString);
   TobDetTB.PutValue('ATB_LIBELLERES', Q.FindField('ARS_LIBELLE').AsString);
   TobDetTB.PutValue('ATB_LIBELLE2RES', Q.FindField('ARS_LIBELLE2').AsString);
   TobDetTB.PutValue('ATB_LIBRERES1', Q.FindField('ARS_LIBRERES1').AsString);
   TobDetTB.PutValue('ATB_LIBRERES2', Q.FindField('ARS_LIBRERES2').AsString);
   TobDetTB.PutValue('ATB_LIBRERES3', Q.FindField('ARS_LIBRERES3').AsString);
   TobDetTB.PutValue('ATB_LIBRERES4', Q.FindField('ARS_LIBRERES4').AsString);
   TobDetTB.PutValue('ATB_LIBRERES5', Q.FindField('ARS_LIBRERES5').AsString);
   TobDetTB.PutValue('ATB_LIBRERES6', Q.FindField('ARS_LIBRERES6').AsString);
   TobDetTB.PutValue('ATB_LIBRERES7', Q.FindField('ARS_LIBRERES7').AsString);
   TobDetTB.PutValue('ATB_LIBRERES8', Q.FindField('ARS_LIBRERES8').AsString);
   TobDetTB.PutValue('ATB_LIBRERES9', Q.FindField('ARS_LIBRERES9').AsString);
   TobDetTB.PutValue('ATB_LIBRERESA', Q.FindField('ARS_LIBRERESA').AsString);
   end else
if typeT ='GA' then
   begin
   TobDetTB.PutValue('ATB_FAMILLENIV1', Q.FindField('GA_FAMILLENIV1').AsString);
   TobDetTB.PutValue('ATB_FAMILLENIV2', Q.FindField('GA_FAMILLENIV2').AsString);
   TobDetTB.PutValue('ATB_FAMILLENIV3', Q.FindField('GA_FAMILLENIV3').AsString);
   TobDetTB.PutValue('ATB_LIBREART1',   Q.FindField('GA_LIBREART1').AsString);
   TobDetTB.PutValue('ATB_LIBREART2',   Q.FindField('GA_LIBREART2').AsString);
   TobDetTB.PutValue('ATB_LIBREART3',   Q.FindField('GA_LIBREART3').AsString);
   TobDetTB.PutValue('ATB_LIBREART4',   Q.FindField('GA_LIBREART4').AsString);
   TobDetTB.PutValue('ATB_LIBREART5',   Q.FindField('GA_LIBREART5').AsString);
   TobDetTB.PutValue('ATB_LIBREART6',   Q.FindField('GA_LIBREART6').AsString);
   TobDetTB.PutValue('ATB_LIBREART7',   Q.FindField('GA_LIBREART7').AsString);
   TobDetTB.PutValue('ATB_LIBREART8',   Q.FindField('GA_LIBREART8').AsString);
   TobDetTB.PutValue('ATB_LIBREART9',   Q.FindField('GA_LIBREART9').AsString);
   TobDetTB.PutValue('ATB_LIBREARTA',   Q.FindField('GA_LIBREARTA').AsString);
   end;
end;

(*******************************************************************************
 *********************** Gestion des requestes SQL *****************************
 *******************************************************************************)
Function TParamTBAffaire.FabriqueRequete( TypeTable, CodeEnr : String;Enr : T_TypeEnr ) : String;
Var st : string;
    bArt,bRes : Boolean;
//    TypeAlim    : T_TypePRFA;
Begin

  Result := '';
  //TypeAlim := Tactivite;

  if (TypeTable = 'AFF') then
    BEGIN
      if codeEnr <> '' then st := ' WHERE AFF_AFFAIREREF="'+ CodeEnr + '" AND AFF_ISAFFAIREREF="-"'
      else st := FabriqueWhere (TypeTable, bArt, bRes, False, Enr) + ' Order By AFF_AFFAIRE';
      Result:='SELECT AFF_AFFAIRE, AFF_TIERS, AFF_AFFAIRE0, AFF_AFFAIRE1, AFF_AFFAIRE2, AFF_AFFAIRE3, AFF_AVENANT, AFF_AFFAIREREF,AFF_ISAFFAIREREF, AFF_LIBELLE, '+
              'AFF_LIBREAFF1, AFF_LIBREAFF2, AFF_LIBREAFF3,AFF_LIBREAFF4, AFF_LIBREAFF5, AFF_LIBREAFF6,AFF_LIBREAFF7, AFF_LIBREAFF8, AFF_LIBREAFF9,AFF_LIBREAFFA,'+
              'AFF_RESSOURCE1,AFF_RESSOURCE2,AFF_RESSOURCE3,AFF_RESPONSABLE, AFF_DEPARTEMENT,AFF_ETABLISSEMENT, AFF_ADMINISTRATIF, '+
              'AFF_GENERAUTO, AFF_TYPEPREVU, AFF_DATEDEBGENER, AFF_DATEDEBUT, AFF_ETATAFFAIRE,AFF_DATECUTOFF,' +
              'AFF_DATEDEBUT, AFF_DATEFIN, AFF_DATELIMITE, AFF_DATEGARANTIE, AFF_DATECLOTTECH, AFF_DATEDEBEXER, AFF_DATEFINEXER, AFF_REGROUPEFACT,' +     // mcd 06/03/03
              'AFF_TOTALHTGLO, AFF_CALCTOTHTGLO,AFF_ACTIVITEGLOBAL,'+             // mcd 11/02/02
              'AFF_DATESITUATION,'+ //mcd 20/08/03
              'T_LIBELLE, T_MOISCLOTURE, T_SOCIETEGROUPE, YTC_TABLELIBRETIERS1, YTC_TABLELIBRETIERS2, YTC_TABLELIBRETIERS3, YTC_TABLELIBRETIERS4, '+
              'YTC_TABLELIBRETIERS5, YTC_TABLELIBRETIERS6, YTC_TABLELIBRETIERS7, YTC_TABLELIBRETIERS8, '+
              'YTC_TABLELIBRETIERS9, YTC_TABLELIBRETIERSA,'+
              'YTC_RESSOURCE1,YTC_RESSOURCE2,YTC_RESSOURCE3 ' +
              'FROM AFFAIRE LEFT OUTER JOIN TIERS ON AFF_TIERS = T_TIERS LEFT OUTER JOIN TIERSCOMPL ON AFF_TIERS=YTC_TIERS ' + st;
    END
  else if (TypeTable = 'T') then
    BEGIN
      Result:='SELECT T_TIERS, T_LIBELLE, T_MOISCLOTURE,T_SOCIETEGROUPE, ' +
                      'YTC_TABLELIBRETIERS1, YTC_TABLELIBRETIERS2, YTC_TABLELIBRETIERS3, YTC_TABLELIBRETIERS4, '+
                      'YTC_TABLELIBRETIERS5, YTC_TABLELIBRETIERS6, YTC_TABLELIBRETIERS7, YTC_TABLELIBRETIERS8, '+
                      'YTC_TABLELIBRETIERS9, YTC_TABLELIBRETIERSA,'+
                      'YTC_RESSOURCE1,YTC_RESSOURCE2,YTC_RESSOURCE3 ' +
                      'FROM TIERS LEFT OUTER JOIN TIERSCOMPL ON T_TIERS=YTC_TIERS ' +
                      FabriqueWhere (TypeTable, bArt, bRes, False, Enr) + ' Order By T_TIERS';
    END
  else if (TypeTable = 'ARS') then
    BEGIN
      if CodeEnr <> '' then st := ' WHERE ARS_RESSOURCE="'+ CodeEnr +'"'
      else st := FabriqueWhere (TypeTable, bArt, bRes, False, Enr) + ' Order By ARS_RESSOURCE';
      Result:='SELECT ARS_RESSOURCE, ARS_TYPERESSOURCE, ARS_LIBELLE, ARS_LIBELLE2, ARS_LIBRERES1, ARS_LIBRERES2, ARS_LIBRERES3, '+
                      'ARS_LIBRERES4, ARS_LIBRERES5, ARS_LIBRERES6, ARS_LIBRERES7, ARS_LIBRERES8, ARS_LIBRERES9, ARS_LIBRERESA ' ;
      Result:=Result+'FROM RESSOURCE ' + st;

    END
  else if (TypeTable = 'GA') then
    BEGIN
      if CodeEnr <> '' then st := ' WHERE GA_ARTICLE="'+ CodeEnr +'"'
      else st:= FabriqueWhere (TypeTable, bArt, bRes, False, Enr) + ' Order By GA_ARTICLE';
      Result:='SELECT GA_ARTICLE, GA_TYPEARTICLE, GA_LIBELLE, GA_FAMILLENIV1, GA_FAMILLENIV2, GA_FAMILLENIV3, '+
                      'GA_LIBREART1, GA_LIBREART2, GA_LIBREART3, GA_LIBREART4, GA_LIBREART5, GA_LIBREART6, '+
                      'GA_LIBREART7, GA_LIBREART8, GA_LIBREART9, GA_LIBREARTA ';
      Result:=Result+ 'FROM ARTICLE ' + st;

    END;
End;

(*********** Gestion du where ************)
Function TParamTBAffaire.FabriqueWhere (TypeTable : String; Var bJoinArt,bJoinRes : Boolean;CutOff: Boolean;Enr : T_TypeEnr) : String ;
Var ChampWhere, stWhere, stTypeTable, Jointure : String;
    i,j, CptWhere, Posit, MaxBoucle :integer;
    dDateDebut, dDateFin : TDateTime;
    CritTraite : Boolean;
Begin

  CptWhere := 0; stWhere := 'WHERE ';
  stTypeTable := TypeTable; Jointure := '';
  MaxBoucle := 1;
                            
  // par défaut pas de jointure
  bJoinTiers := False; bJoinTiersCompl:=False;
  if (TypeTable = 'AFF') then MaxBoucle := 2;
  if (TypeTable = 'ACU') then MaxBoucle := 3;
  if (TypeTable = 'ACT') or (TypeTable = 'GL') then MaxBoucle :=3;

  dDateDebut := iDate1900; dDateFin := iDate2099;

  for j := 1 to MaxBoucle do
    BEGIN

    if (j=2) And (TypeTable = 'AFF') then
      BEGIN stTypeTable:='T'; Jointure := TypeTable; END;
    if (j >1) And ((TypeTable = 'ACT') or (TypeTable = 'GL')or (TypeTable = 'ACU')) then
      BEGIN
      Jointure := TypeTable;
      if (j=2) then stTypeTable:='GA' else
      if (j=3) then stTypeTable:='ARS';
      END;

    for i:=1 to MaxCritTB do
        BEGIN
        if (FListeCriteres[i].Champ <> '') then  // Champs critéres utilisé
            BEGIN
            ChampWhere := FListeCriteres[i].Champ;
            Posit := Pos('_' , ChampWhere);
            if (Posit > 0) then ChampWhere := Copy (ChampWhere, Posit+1, length(ChampWhere)- Posit);
            ChampWhere := ChampsTBversWhere(ChampWhere,stTypeTable,i,Jointure);
            if (ChampWhere <> '') then
                BEGIN
                if Pos('ARS_' , ChampWhere)>0 then bJoinRes :=True;
                if Pos('GA_' , ChampWhere)>0  then bJoinArt :=True;
                CritTraite := True;
                if FListeCriteres[i].bTypeDate then
                    BEGIN
                    If CutOff = True then
                       Begin
                            // AN : cut_off = cut off généré en fin exercice précédent ==> mois debut moins 1 mois
                        if Enr=Tan then begin
                             dDateFin   := PlusDate(StrToDate (FListeCriteres[i].BorneDebut),-1,'J');
                             dDateDebut := DebutDeMois(dDateFin);
                             end
                        else if Enr=TMois then begin
                             dDateDebut := DebutDeMois(StrToDate (FListeCriteres[i].BorneFin));
                             dDateFin   :=StrToDate (FListeCriteres[i].BorneFin);
                             end
                        else begin
                                    // MCD mettre dates dernier cutoff paramètre
                                    // si fait parti de la sélection (cas cut off provisoire)
                                    // sinon, pas de cut off, mise date à 2099
                                    // mcd 03/06/2002 : si option écran cutoff par mois
                                    // prend tous les cutoff de la période ...
                                    // mcd  24/07/02 prend la date dernier cut off pas dans les paramètres mais dans l'affaire
                             If CutoffMois then begin
                                 dDateDebut := StrToDate (FListeCriteres[i].BorneDebut);
                                 dDateFin   := StrToDate (FListeCriteres[i].BorneFin);
                                 end
                             else begin
                                      // mcd 24/07/02 if (GetParamSoc('So_DATECUTOFF')<>idate1900) and (dateToSTr(GetParamSoc('So_DATECUTOFF'))<>'30/12/1899')then
                                      // mcd 24/07 if (GetParamSoc ('SO_DATECUTOFF') >=StrToDate (FListeCriteres[i].BorneDebut) )
                                      // mcd 24/07/02  and (GetParamSoc ('SO_DATECUTOFF') <= StrToDate (FListeCriteres[i].BorneFin)) then
                                  If ( DateCutoff<= StrToDate (FListeCriteres[i].BorneFin))
                                        and (DateCutoff >=StrToDate (FListeCriteres[i].BorneDebut) )
                                     then begin
                                     // mcd 24/07/02 dDateFin   := GetParamSoc('So_DATECUTOFF');
                                        dDateFin   := DateCutOff; // génération cut off par affaire ..
                                        dDateDebut := DebutDeMois(dDateFin);
                                        end
                                  else begin // ajout mcd 14/10/2002  si le dernier cutoof pas ds sélection, on regarde si un cut off sur mois fin période
                                        dDateFin   := StrToDate (FListeCriteres[i].BorneFin);
                                        dDateDebut := DebutDeMois(dDateFin);
                                        end;
                                  end;
                              end;
                       end
                    else begin // cas ou pas traitement table cut off
                        If AlimSolde then
                          begin  //mcd 20/08/03
                          if Enr=Tan then begin
                             dDateDebut := idate1900;
                             dDateFin   := DateSituation;
                             end
                          else if Enr=TMois then begin
                             dDateDebut := DebutDeMois(StrToDate (FListeCriteres[i].BorneFin));
                             dDateFin   :=StrToDate (FListeCriteres[i].BorneFin);
                             end
                          else begin
                             dDateDebut := PlusDate(DateSituation,1,'J');
                             dDateFin   := StrToDate (FListeCriteres[i].BorneFin);
                             end;
                          end
                       else begin
                         if Enr=Tan then begin
                             dDateDebut := idate1900;
                             dDateFin   := PlusDate(StrToDate (FListeCriteres[i].BorneDebut),-1,'J');
                             end
                         else if Enr=TMois then begin
                             dDateDebut := DebutDeMois(StrToDate (FListeCriteres[i].BorneFin));
                             dDateFin   :=StrToDate (FListeCriteres[i].BorneFin);
                             end
                         else begin
                             dDateDebut := StrToDate (FListeCriteres[i].BorneDebut);
                             dDateFin   := StrToDate (FListeCriteres[i].BorneFin);
                             end;
                         end;
                        end;
                   if (dDateDebut = idate1900) And (dDateFin = iDate2099) then CritTraite := False;
                   END;
                if FListeCriteres[i].bReel then
                    BEGIN
                    if  (Valeur(FListeCriteres[i].BorneDebut) = 0) And (Valeur(FListeCriteres[i].BorneFin)=0) then
                       CritTraite := False;
                    END;
                if CritTraite then
                    BEGIN
                    Inc(CptWhere);
                    if (CptWhere > 1) And not(FListeCriteres[i].bMultiVal) then stWhere:= stWhere + ' AND ';
                    if FListeCriteres[i].bTypeDate then // type date
                        stWhere:= stWhere + '(('+ChampWhere+' >= "'+ UsDateTime(dDateDebut)+ '" ) AND ('+ChampWhere +' <= "'+ UsDateTime(dDateFin) +'" ))'
                    else if FListeCriteres[i].bReel then // type réel
                        BEGIN
                        if FListeCriteres[i].BorneDebut = FListeCriteres[i].BorneFin then
                            stWhere:= stWhere + '('+ChampWhere+' = '+FloatToStr(Valeur(FListeCriteres[i].BorneDebut))+ ')'
                        else
                            stWhere:= stWhere + '(('+ChampWhere+' >= '+FloatToStr(Valeur(FListeCriteres[i].BorneDebut))+ ' ) AND ('+ChampWhere +' <= '+ FloatToStr(Valeur(FListeCriteres[i].BorneFin))+' ))';
                        END
                    else if FListeCriteres[i].bMultiVal then // Multivalcombo
                        BEGIN
                            if FListeCriteres[i].BorneDebut <>'' then
                             stWhere:= stWhere + FindEtReplace(FListeCriteres[i].BorneDebut,FListeCriteres[i].Champ,ChampWhere,True);
                        END
                    else // Type string *****
                        BEGIN
                        if FListeCriteres[i].BorneDebut = FListeCriteres[i].BorneFin then
                            stWhere:= stWhere + '('+ChampWhere+' = "'+FListeCriteres[i].BorneDebut+ '")'
                        else
                            stWhere:= stWhere + '(('+ChampWhere+' >= "'+FListeCriteres[i].BorneDebut+ '" ) AND ('+ChampWhere +' <= "'+ FListeCriteres[i].BorneFin+'" ))';
                        END;
                    END;
                END;
            END;
        END;
    END;
  if CptWhere > 0 then Result := stWhere else Result := '';
End;

Function TParamTBAffaire.FabriqueWherePRFA (TypeTable,Code : String;TypeAlim: T_TypePRFA; Var bJoinArt,bJoinRes : Boolean;Enr : T_TypeEnr) : String ;
Var stWhere, ChampsCle,stWhereCrit : String;
Begin
stWhere := 'WHERE ';
stWhereCrit := FabriqueWhere(TypeTable, bJoinArt, bJoinRes, False, Enr);
if (stWhereCrit <> '') then stWhere :=stWhereCrit+ ' AND';

if (TypeTable = 'GL') then
   begin
   if TypeTB = ttbTiers then ChampsCle:='GL_TIERS' else ChampsCle:='GL_AFFAIRE';
   stWhere := stWhere + ' ('+ ChampsCle + '="'+Code + '")';
   end;
// mcd 09/07 if (TypeTable = 'ACT') then
if (TypeTable <>'GL') then
   begin
   if TypeTB = ttbTiers then ChampsCle:='ACT_TIERS' else ChampsCle:='ACT_AFFAIRE';
   stWhere := stWhere + ' ('+ ChampsCle + '="'+Code + '") AND ';
   if (TypeAlim = TBoni) then stWhere := stWhere + ' (ACT_TYPEACTIVITE = "BON")'
    else stWhere := stWhere + ' (ACT_TYPEACTIVITE = "REA")';

   if GetParamSoc('SO_AFCLIENT') = cInClientAlgoe then
   		stWhere := stWhere + ' AND (ACT_ACTIVITEREPRIS <> "A") ';
   end;

// Traitement des natures de pièces + type de ligne à reprendre
Result := stWhere;
if (TypeTable = 'GL') then
   begin
   if WhereNaturePiece <> '' then Result := stWhere + ' AND ' + WhereNaturePiece;
   if WhereLigne <> '' then Result := Result + ' AND ' + WhereLigne;
   end;
if (TypeTable = 'ACT') then
   begin
   if WhereNatAchatsurAct <> '' then Result := Result + ' AND ' + WhereNatAchatsurAct;
      // on ne prend pas en compte la notion de visa pour les BONI. A modifier si besoin
   if ((AlimVisa) and (TypeAlim <>TBoni)) then Result := Result + ' AND (ACT_ETATVISA="VIS")';
   end;
End;


{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... :
Modifié le ... :
Description .. :
Mots clefs ... :
*****************************************************************}
Function TParamTBAffaire.ChampsTBversWhere(ChampWhere,TypeTable : String; NumCrit:Integer; Jointure : string): String;
BEGIN
  Result := '';
  if (TypeTable = 'T')  and (FListeCriteres[NumCrit].WhereTiers = False) And
     (FListeCriteres[NumCrit].WhereTiersCompl = False) then Exit;
  if (TypeTable = 'AFF')and (FListeCriteres[NumCrit].WhereAffaire = False) then Exit;
  if (TypeTable = 'GL')  And (FListeCriteres[NumCrit].WhereLigne=False) then Exit;
  if (TypeTable = 'ACT') And (FListeCriteres[NumCrit].WhereAct=False) then Exit;
  if (TypeTable = 'GA') And (FListeCriteres[NumCrit].WhereArt=False) then Exit;
  if (TypeTable = 'ARS') And (FListeCriteres[NumCrit].WhereRes=False) then Exit;
  if (TypeTable = 'ACU') And (FListeCriteres[NumCrit].WhereCutOff=False) then Exit;
  if (TypeTable = 'ABU') And (FListeCriteres[NumCrit].WhereBud=False) then Exit;
  if (TypeTable = 'APL') And (FListeCriteres[NumCrit].WherePlanning=False) then Exit;

  if (Jointure ='ACT') And (FListeCriteres[NumCrit].WhereAct=true) then exit;   // déjà traité
  if (Jointure='GL') And (FListeCriteres[NumCrit].WhereLigne=true) then exit;  // déja traité
  if (Jointure='AFF') And (FListeCriteres[NumCrit].WhereAffaire=true) then exit;  // déja traité
  if ChampWhere = 'DATE' then
    BEGIN
      if TypeTable ='ACT' then Result := 'ACT_DATEACTIVITE'
      else if TypeTable ='ACU' then Result := 'ACU_DATE'
      else if TypeTable ='ABU' then Result := 'ABU_DATEBUD'
      else if TypeTable ='GL'  then Result := 'GL_DATEPIECE'
      else if TypeTable ='APL' then Result := 'APL_DATEDEBPLA'
      else Result := TypeTable + '_' + ChampWhere;
    END
  else if (FListeCriteres[NumCrit].WhereTiersCompl=True) And (TypeTable='T') then
    BEGIN  // complément du tiers
      if (Pos('LIBRETIERS',ChampWhere)>0) then ChampWhere := 'TABLE'+ ChampWhere;
      Result := 'YTC' + '_' + ChampWhere;
      bJoinTiersCompl := True;
    END
  else // par defaut même nom de champs
    Result := TypeTable + '_' + ChampWhere;
  if (TypeTable = 'T'  ) And (result <> '') And (Pos('YTC_',Result)=0) then bJoinTiers  := True;
  if (TypeTable = 'AFF') And (result <> '') then bJoinAffaire:= True;
END;

Function TParamTBAffaire.FabriqueWhereNatAchatsurAct : string;
Var Nat,tmpNatachat : string;
BEGIN
Result :='';
// mcd 06/0_/02 il faut tout prendre .... if not(AlimAchat) then begin Result := ' ACT_NATUREPIECEG="" '; exit; end;
if (NatAchat = '') and (NatAchatEngage = '') then Exit;
tmpNatachat := NatAchat;
Nat:=Trim(ReadTokenSt(tmpNatAchat));
if Nat <> '' then Result := '(ACT_NATUREPIECEG=""'; // lignes d'activités par défaut + ...
While (Nat <>'') do
    BEGIN
    Result := Result + ' OR ACT_NATUREPIECEG="' + Nat + '"';
    Nat:=Trim(ReadTokenSt(tmpNatAchat));
    END;
tmpNatachat := NatAchatEngage;
Nat:=Trim(ReadTokenSt(tmpNatAchat));
if (Nat <> '') and (NatAchat ='') then Result := '(ACT_NATUREPIECEG=""'; // lignes d'activités par défaut + ...
While (Nat <>'') do
    BEGIN
    Result := Result + ' OR ACT_NATUREPIECEG="' + Nat + '"';
    Nat:=Trim(ReadTokenSt(tmpNatAchat));
    END;

Result := result+')'
END;

Function TParamTBAffaire.FabriqueWhereCutOff (Code,TypeEnr : String; Var bJoinArt,bJoinRes : Boolean;Enr : T_TypeEnr) : String ;
Var stWhere, ChampsCle,stWhereCrit : String;
BEGIN
stWhere := 'WHERE ';
if Typeenr = 'CVE' then stWhereCrit := FabriqueWhere('ACU', bJoinArt, bJoinRes, True, Enr)
else stWhereCrit := FabriqueWhere('ACU', bJoinArt, bJoinRes, False, Enr);  // pour dates idem activite et piece
if (stWhereCrit <> '') then stWhere :=stWhereCrit+ ' AND';
if TypeTB = ttbTiers then ChampsCle:='ACU_TIERS' else ChampsCle:='ACU_AFFAIRE';
stWhere := stWhere + ' ('+ ChampsCle + '="'+Code + '") AND (ACU_TYPEAC="'+TypeEnr+'"';
if (VH_GC.GAAchatSeria) and (typeEnr='CVE') then stwhere:=stwhere+' OR ACU_TYPEAC="CAC"';
stWhere:=Stwhere+')' ;
Result := stWhere;
END;

Function TParamTBAffaire.FabriqueWhereBud (Code,TypeEnr : String; Var bJoinArt,bJoinRes : Boolean;Enr : T_TypeEnr) : String ;
Var stWhere, ChampsCle,stWhereCrit : String;
BEGIN
stWhere := 'WHERE ';
stWhereCrit := FabriqueWhere('ABU', bJoinArt, bJoinRes, False, Enr);  // pour dates idem activite et piece
if (stWhereCrit <> '') then stWhere :=stWhereCrit+ ' AND';
if TypeTB = ttbTiers then ChampsCle:='ABU_TIERS' else ChampsCle:='ABU_AFFAIRE';
stWhere := stWhere + ' ('+ ChampsCle + '="'+Code + '") AND (ABU_TYPEAFBUDGET="'+TypeEnr+'")';
Result := stWhere;
END;

Function TParamTBAffaire.FabriqueWherePlanning (Code : String; Var bJoinArt,bJoinRes : Boolean;Enr : T_TypeEnr) : String ;
var
  vStWhereCrit : String;
  vStChampsCle : String;
begin
  result := 'WHERE ';
  vStWhereCrit := FabriqueWhere('APL', bJoinArt, bJoinRes, False, Enr);  // pour dates idem activite et piece
  if (vStWhereCrit <> '') then result := vStWhereCrit+ ' AND';
  if TypeTB = ttbTiers then vStChampsCle := 'APL_TIERS' else vStChampsCle:='APL_AFFAIRE';
  result := result + ' ('+ vStChampsCle + ' = "' + Code + '") ';
  // on supprime les entetes de planning
  // on ne prend en compte que la planification des ressources
  result := result + ' AND APL_RESSOURCE <> "" ';
end;

Function TParamTBAffaire.FabriqueWhereNaturePiece : string;
Var i,CptWhere : integer;
    stAffect, stWhere : string;
    TOBDet : TOB;
    TraiteNature : Boolean;
BEGIN
CptWhere := 0;
if VH_GC.TOBParPiece = Nil then Exit;
for i := 0 To VH_GC.TOBParPiece.Detail.Count-1 do
    BEGIN
    TraiteNature := False;
    TobDet := VH_GC.TOBParPiece.Detail[i];
    stAffect := TobDet.GetValue ('GPP_AFAFFECTTB');
    //if (stAffect = 'ACH') And AlimAchat then TraiteNature := true;  pièces d'achats non reprises
    if (stAffect = 'PRE') And AlimPrevuAff then TraiteNature := true;
      // mcd 03/06/02 if (stAffect = 'FAC') And AlimFacture then TraiteNature := true;
    if (stAffect = 'FAC') And AlimFacture and (GetParamSoc('SO_AFFACTPARRES')='SAN') then TraiteNature := true;
    if TraiteNature then
       BEGIN
       if cptWhere > 0 then stWhere := stWhere + ' OR GL_NATUREPIECEG="' + TobDet.GetValue('GPP_NATUREPIECEG')+  '" '
                        else stWhere := '( GL_NATUREPIECEG="' + TobDet.GetValue('GPP_NATUREPIECEG')+ '" ';
       inc(CptWhere);
       END;
    END;
stWhere := stWhere + ' ) ';
stWhere := stWhere+ ' AND GL_INDICEG=0 ';  // mcd 17/09/02 on ne prend pas que les pièces principale, et pas les reliquat .. (utile sur piece achat)
if CptWhere > 0 then Result := stWhere else Result := '';
END;

Function TParamTBAffaire.FabriqueWhereLigne : string;
BEGIN
// les lignes de TOT sous totaux non repris .
Result := '(GL_TYPELIGNE="ART") '
END;

Procedure TParamTBAffaire.AlimListeChampsSt (TypeRequete: T_TypePRFA ; Var ChampsCle, ListeChamps : String; Var bJoinArt, bJoinRes : Boolean);
BEGIN
ChampsCle := ''; ListeChamps := '';

//mcd 09/07 if (TypeRequete = tActivite) then
if (TypeRequete<>Tligne) then
    BEGIN
    // Traitement des champs texte non regroupés
    if TypeTB = ttbTiers then ChampsCle := 'ACT_TIERS'
                         else ChampsCle := 'ACT_AFFAIRE,ACT_TIERS';

    if (NiveauTB = tniTypeArt) then
        ListeChamps := ListeChamps+ ',ACT_TYPEARTICLE' else
    if (NiveauTB = tniArt) then
        ListeChamps := ListeChamps+ ',ACT_TYPEARTICLE,ACT_CODEARTICLE,ACT_ARTICLE' else
    if (NiveauTB = tniRes) then
        ListeChamps := ListeChamps+ ',ACT_TYPERESSOURCE,ACT_RESSOURCE,ACT_FONCTIONRES' else
    if (NiveauTB = tniTypeRes) then
        ListeChamps := ListeChamps+ ',ACT_TYPERESSOURCE' else
    if (NiveauTB = tniFonction) then
        ListeChamps := ListeChamps+ ',ACT_FONCTIONRES' else
    if (NiveauTB = tniTypeArtRes) then
        ListeChamps := ListeChamps+ ',ACT_TYPEARTICLE,ACT_TYPERESSOURCE,ACT_RESSOURCE' else
    if (NiveauTB = tniArtRes) then
        ListeChamps := ListeChamps+ ',ACT_TYPEARTICLE,ACT_CODEARTICLE,ACT_ARTICLE,ACT_TYPERESSOURCE,ACT_RESSOURCE' else
    if (NiveauTB <> tniFamilleArt1)    //mcd 18/11/02
       and (NiveauTB <> tniFamilleArt2)
       and  (NiveauTB <> tniFamilleArt3)
       and  (NiveauTB <> tniStatArt1)
       and (NiveauTB <> tniStatArt2)
       and (NiveauTB <> tniStatArt3)
       then
         // Chargement de tous les champs par défaut (tniTous = Détail)
         ListeChamps := ListeChamps+ ',ACT_TYPEARTICLE,ACT_CODEARTICLE,ACT_ARTICLE,ACT_TYPERESSOURCE,ACT_RESSOURCE,ACT_FONCTIONRES';
    // Gestion en plus du champs : ...  facturable, libelle, affaire ref ...
    if GereFacturable         then ListeChamps := ListeChamps + ',ACT_ACTIVITEREPRIS';
    if TypeDateTB = tdaDetail then
       begin
       ListeChamps := ListeChamps + ',ACT_LIBELLE,ACT_TYPEHEURE';
       if (AlimAchat) And (NatAchat <> '') then ListeChamps := ListeChamps + ',ACT_NATUREPIECEG,ACT_SOUCHE,ACT_NUMERO,ACT_INDICEG,ACT_FOURNISSEUR';
       end;
    AlimSuiteChamps(ListeChamps,bJoinArt, bJoinRes); //mcd 18/11/02
    Centralisation(ListeChamps, bJoinArt, bJoinRes); //mcd 18/11/02
    END
else
    BEGIN
    // Traitement des champs texte non regroupés
    If CentrFam then begin
        ListeChamps := ListeChamps+ ',GL_FAMILLENIV1,GL_FAMILLENIV2,GL_FAMILLENIV3';
       end;
    If CentrStatArt then begin
        ListeChamps := ListeChamps+ ',GL_LIBREART1,GL_LIBREART2,GL_LIBREART3,GL_LIBREART4,GL_LIBREART5,GL_LIBREART6,GL_LIBREART7,GL_LIBREART8,GL_LIBREART9,GL_LIBREARTA';
       end;
    If CentrStatRes then begin
        ListeChamps := ListeChamps+ ',ARS_LIBRERES1,ARS_LIBRERES2,ARS_LIBRERES3,ARS_LIBRERES4,ARS_LIBRERES5,ARS_LIBRERES6,ARS_LIBRERES7,ARS_LIBRERES8,ARS_LIBRERES9,ARS_LIBRERESA';
        bJoinRes:= True;
       end;
    if TypeTB = ttbTiers then ChampsCle := 'GL_TIERS'
                         else ChampsCle := 'GL_AFFAIRE,GL_TIERS';

    if (NiveauTB = tniTypeArt) then
        ListeChamps := ListeChamps+ ',GL_TYPEARTICLE' else
    if (NiveauTB = tniArt) then
        ListeChamps :=ListeChamps+  ',GL_TYPEARTICLE,GL_CODEARTICLE,GL_ARTICLE' else
    if (NiveauTB = tniRes) then
        ListeChamps :=ListeChamps+  ',GL_RESSOURCE' else
    if (NiveauTB = tniTypeRes) then
        ListeChamps := ListeChamps+ ',GL_RESSOURCE' else
    if (NiveauTB = tniFonction) then
        ListeChamps :=ListeChamps+  ',GL_RESSOURCE' else
    if (NiveauTB = tniTypeArtRes) then
        ListeChamps :=ListeChamps+  ',GL_TYPEARTICLE,GL_RESSOURCE' else
    if (NiveauTB = tniArtRes) then
        ListeChamps := ListeChamps+ ',GL_TYPEARTICLE,GL_CODEARTICLE,GL_ARTICLE,GL_RESSOURCE' else
    if (NiveauTB = tniFamilleArt1) then
        ListeChamps := ListeChamps+ ',GL_FAMILLENIV1' else
    if (NiveauTB = tniFamilleArt2) then
        ListeChamps := ListeChamps+ ',GL_FAMILLENIV2' else
    if (NiveauTB = tniFamilleArt3) then
        ListeChamps := ListeChamps+ ',GL_FAMILLENIV3' else
    if (NiveauTB = tniStatArt1) then
        ListeChamps := ListeChamps+ ',GL_LIBREART1' else
    if (NiveauTB = tniStatArt2) then
        ListeChamps := ListeChamps+ ',GL_LIBREART2' else
    if (NiveauTB = tniStatArt3) then
        ListeChamps := ListeChamps+ ',GL_LIBREART3'
    else ListeChamps := ListeChamps+ ',GL_TYPEARTICLE,GL_CODEARTICLE,GL_ARTICLE,GL_RESSOURCE';
                     //mcd 13/03/02 ajout n° et nature
    if TypeDateTB = tdaDetail then ListeChamps := ListeChamps + ',GL_LIBELLE,GL_NUMERO,GL_NATUREPIECEG,GL_SOUCHE';
(*    if (NiveauTb=TniTous) or (NiveauTb=TniRes) or (NiveauTB=TniArtRes)
      or (NiveauTb=TniTypeArtRes) then
      begin ListeChamps  :=ListeChamps +',ARS_LIBELLE, ARS_LIBELLE2';bJoinRes:= True; END ;
*)
    if (NiveauTb=TniTous) or (NiveauTb=TniArt)
      or (NiveauTb=TniArtRes) then
      begin ListeChamps := ListeChamps  +',GA_LIBELLE';bJoinArt:= True; END ;
    END;
END;

Procedure TParamTBAffaire.alimSuiteChamps(Var ListeChamps: string;Var bJoinArt, bJoinRes:boolean) ;
begin    // fct contenant tous les champs sur article et ressource obligatoire en fct paramétrage
  if (NiveauTB = tniFamilleArt1) then
      BEGIN ListeChamps := ListeChamps+ ',GA_FAMILLENIV1 '; bJoinArt:= True; END
  else if (NiveauTB = tniFamilleArt2) then
      BEGIN ListeChamps := ListeChamps+ ',GA_FAMILLENIV2 '; bJoinArt:= True; END
  else if (NiveauTB = tniFamilleArt3) then
      BEGIN ListeChamps := ListeChamps+ ',GA_FAMILLENIV3 '; bJoinArt:= True; END
  else if (NiveauTB = tniStatArt1) then
      BEGIN ListeChamps := ListeChamps+ ',GA_LIBREART1 '; bJoinArt:= True; END
  else if (NiveauTB = tniStatArt2) then
      BEGIN ListeChamps := ListeChamps+ ',GA_LIBREART2 '; bJoinArt:= True; END
  else if (NiveauTB = tniStatArt3) then
      BEGIN ListeChamps := ListeChamps+ ',GA_LIBREART3 '; bJoinArt:= True; END ;
  if (NiveauTb=TniTous) or (NiveauTb=TniRes) or (NiveauTB=TniArtRes)
      or (NiveauTb=TniTypeArtRes) then
      begin ListeChamps  :=ListeChamps +',ARS_LIBELLE , ARS_LIBELLE2  '; bJoinRes:= True; END ;
  if (NiveauTb=TniTous) or (NiveauTb=TniArt)
      or (NiveauTb=TniArtRes) then
      begin ListeChamps := ListeChamps  +',GA_LIBELLE ';bJoinArt:= True; END ;
end;

Procedure TParamTBAffaire.AlimListeChampsStCut ( Var ChampsCle, ListeChamps : String; Var bJoinArt,bJoinRes : Boolean);
BEGIN
  ChampsCle := ''; ListeChamps := '';

      // Traitement des champs texte non regroupés
  if TypeTB = ttbTiers then ChampsCle := 'ACU_TIERS'
                       else ChampsCle := 'ACU_AFFAIRE,ACU_TIERS';

  Centralisation(ListeChamps, bJoinArt, bJoinRes);

  if (NiveauTB = tniTypeArt) then
      ListeChamps :=ListeChamps+  ',ACU_TYPEARTICLE' else
  if (NiveauTB = tniArt) then
      ListeChamps :=ListeChamps+  ',ACU_TYPEARTICLE,ACU_CODEARTICLE,ACU_ARTICLE' else
  if (NiveauTB = tniRes) then
      ListeChamps :=ListeChamps+  ',ACU_TYPERESSOURCE,ACU_RESSOURCE' else
  if (NiveauTB = tniTypeRes) then
      ListeChamps :=ListeChamps+  ',ACU_TYPERESSOURCE' else
  if (NiveauTB = tniFonction) then
      ListeChamps := ListeChamps+ ',ACU_RESSOURCE' else
  if (NiveauTB = tniTypeArtRes) then
      ListeChamps := ListeChamps+ ',ACU_TYPEARTICLE,ACU_TYPERESSOURCE,ACU_RESSOURCE' else
  if (NiveauTB = tniArtRes) then
      ListeChamps := ListeChamps+ ',ACU_TYPEARTICLE,ACU_CODEARTICLE,ACU_ARTICLE,ACU_TYPERESSOURCE,ACU_RESSOURCE'
  else   // Chargement de tous les champs par défaut (tniTous = Détail)
      ListeChamps := ListeChamps+ ',ACU_TYPEARTICLE,ACU_CODEARTICLE,ACU_ARTICLE,ACU_TYPERESSOURCE,ACU_RESSOURCE';
  if TypeDateTB = tdaDetail then ListeChamps := ListeChamps + ',ACU_NUMERO,ACU_NATUREPIECEG,ACU_SOUCHE';
  AlimSuiteChamps(ListeChamps,bJoinArt, bJoinRes); //mcd 18/11/02
END;

Procedure TParamTBAffaire.Centralisation(var ListeChamps : String; var bJoinArt, bJoinRes : Boolean);
begin
  If CentrFam then
    begin
      ListeChamps := ListeChamps +  ',GA_FAMILLENIV1,GA_FAMILLENIV2,GA_FAMILLENIV3';
      bJoinArt := True;
    end;
  If CentrStatArt then
    begin
      ListeChamps := ListeChamps + ',GA_LIBREART1,GA_LIBREART2,GA_LIBREART3,GA_LIBREART4,GA_LIBREART5,GA_LIBREART6,GA_LIBREART7,GA_LIBREART8,GA_LIBREART9,GA_LIBREARTA';
      bJoinArt := True;
    end;
  If CentrStatRes then
    begin
      ListeChamps := ListeChamps + ',ARS_LIBRERES1,ARS_LIBRERES2,ARS_LIBRERES3,ARS_LIBRERES4,ARS_LIBRERES5,ARS_LIBRERES6,ARS_LIBRERES7,ARS_LIBRERES8,ARS_LIBRERES9,ARS_LIBRERESA';
      bJoinRes := True;
    end;
end;

Procedure TParamTBAffaire.AlimListeChampsStBud ( Var ChampsCle, ListeChamps : String; Var bJoinArt,bJoinRes : Boolean);
BEGIN
  ChampsCle := ''; ListeChamps := '';

      // Traitement des champs texte non regroupés
  if TypeTB = ttbTiers then ChampsCle := 'ABU_TIERS'
                       else ChampsCle := 'ABU_AFFAIRE,ABU_TIERS';

  Centralisation(ListeChamps, bJoinArt, bJoinRes);

  if (NiveauTB = tniTypeArt) then
      ListeChamps :=ListeChamps+  ',ABU_TYPEARTICLE' else
  if (NiveauTB = tniArt) then
      ListeChamps :=ListeChamps+  ',ABU_TYPEARTICLE,ABU_CODEARTICLE,ABU_ARTICLE' else
  if (NiveauTB = tniRes) then
      ListeChamps :=ListeChamps+  ',ABU_TYPERESSOURCE,ABU_RESSOURCE' else
  if (NiveauTB = tniTypeRes) then
      ListeChamps :=ListeChamps+  ',ABU_TYPERESSOURCE' else
  if (NiveauTB = tniFonction) then
      ListeChamps := ListeChamps+ ',ABU_RESSOURCE' else
  if (NiveauTB = tniTypeArtRes) then
      ListeChamps := ListeChamps+ ',ABU_TYPEARTICLE,ABU_TYPERESSOURCE,ABU_RESSOURCE' else
  if (NiveauTB = tniArtRes) then
      ListeChamps := ListeChamps+ ',ABU_TYPEARTICLE,ABU_CODEARTICLE,ABU_ARTICLE,ABU_TYPERESSOURCE,ABU_RESSOURCE'
  else   // Chargement de tous les champs par défaut (tniTous = Détail)
      ListeChamps := ListeChamps+ ',ABU_TYPEARTICLE,ABU_CODEARTICLE,ABU_ARTICLE,ABU_TYPERESSOURCE,ABU_RESSOURCE';
  if TypeDateTB = tdaDetail then ListeChamps := ListeChamps + ',ABU_NUMERO';
  AlimSuiteChamps(ListeChamps,bJoinArt, bJoinRes);//mcd 18/11/02
END;

Procedure TParamTBAffaire.AlimListeChampsStPlanning ( Var ChampsCle, ListeChamps : String; Var bJoinArt, bJoinRes : Boolean);
begin

  ChampsCle := '';
  ListeChamps := '';

  // Traitement des champs texte non regroupés
  if TypeTB = ttbTiers then ChampsCle := 'APL_TIERS'
                       else ChampsCle := 'APL_AFFAIRE, APL_TIERS';

  Centralisation(ListeChamps, bJoinArt, bJoinRes);

  if (NiveauTB = tniTypeArt)          then ListeChamps := ListeChamps + ', APL_TYPEARTICLE'
  else if (NiveauTB = tniArt)         then ListeChamps := ListeChamps + ', APL_TYPEARTICLE, APL_CODEARTICLE,APL_ARTICLE'

  else if (NiveauTB = tniRes)         then
    begin ListeChamps := ListeChamps + ', ARS_TYPERESSOURCE'; bJoinRes := true; end //, APL_RESSOURCE

  else if (NiveauTB = tniTypeRes)     then
    begin ListeChamps := ListeChamps + ', ARS_TYPERESSOURCE'; bJoinRes := true; end

  //else if (NiveauTB = tniFonction)    then ListeChamps := ListeChamps + ', APL_RESSOURCE'

  else if (NiveauTB = tniTypeArtRes)  then
    begin ListeChamps := ListeChamps + ', APL_TYPEARTICLE, ARS_TYPERESSOURCE'; bJoinRes := true; end //, APL_RESSOURCE

  else if (NiveauTB = tniArtRes)      then
    begin ListeChamps := ListeChamps + ', APL_TYPEARTICLE, APL_CODEARTICLE,APL_ARTICLE, ARS_TYPERESSOURCE'; bJoinRes := true; end //, APL_RESSOURCE

  // Chargement de tous les champs par défaut (tniTous = Détail)
  else
    begin
      ListeChamps := ListeChamps + ', APL_TYPEARTICLE, APL_CODEARTICLE,APL_ARTICLE, ARS_TYPERESSOURCE'; //, APL_RESSOURCE
      bJoinRes := true;
    end;

  // pour le découpage des lignes                                  
  ListeChamps := ListeChamps + ', APL_DATEDEBPLA, APL_DATEFINPLA, APL_RESSOURCE';

  if TypeDateTB = tdaDetail then ListeChamps := ListeChamps + ', APL_NUMEROLIGNE';
  AlimSuiteChamps(ListeChamps,bJoinArt, bJoinRes);//mcd 18/11/02

end;

Function TParamTBAffaire.FabriqueRequetePRFA (TypeRequete: T_TypePRFA;CodeEnr : string;Enr : T_TypeEnr)  : String;
Var ChampsRegroupe, ChampsCle,ListeChamps, ListeChampsNum, ListeChampsCompl,
    ListeChampsNumSum, ChampsUnite, ChampsDate, stWhere,
    PrefixeTable, NomTable, JoinAffaire, JoinTiers, JoinArticle, JoinTiersCompl,
    JoinRessource,SQL,StOrderBy : String;
    bJoinArt,bJoinRes : Boolean;
BEGIN
ChampsRegroupe := ''; ListeChamps := '' ; ChampsCle := '';
SQL := ''; stOrderBy := '';
ListeChampsNum := ''; ListeChampsNumSum := ''; ListeChampsCompl := '';
Result := ''; StWhere :='';
JoinAffaire := '';  JoinTiers:=''; JoinArticle:=''; JoinTiersCompl:=''; JoinRessource:='';
bJoinArt:=false; bJoinRes :=false;
(**Initialisation + définition du champs unité champs géré à part car placé en fin de group by ou d'order by
   car les lignes sur une même clé différenciées uniquement sur l'unité sont regroupées
   classification :
   - ChampsCle champs cle repris en premier dans le group by non repris ds champs selectifs
   - Listechamps : liste des champs texte repris
   - ListechampsNum : champs numérique direct repris en détail
   - ListechampsNumSum : champs numérique sommés pour group by ( cas du cumul semaine, mois, global)
   - ChampsRegroupe : Champs repris dans le group by.
   - listechampsCompl : non repris dans la listechampstot car non affichés dans la tob **)
if (TypeRequete <>Tligne) then
    BEGIN
    PrefixeTable:= 'ACT_';         NomTable := 'ACTIVITE';
    ChampsDate := 'ACT_DATEACTIVITE';
    ChampsUnite:= 'ACT_UNITE';
    END
else
    BEGIN
    PrefixeTable := 'GL_';          NomTable := 'LIGNE';
    ChampsDate := 'GL_DATEPIECE';
    ChampsUnite:= 'GL_QUALIFQTEVTE';
    END;

// Traitement des champs texte non regroupés en fonction du paramétrage
AlimListeChampsSt (TypeRequete,ChampsCle,ListeChamps,bJoinArt,bJoinRes);

if (TypeDateTB = tdaDetail ) then ListeChamps:= ListeChamps + ','+ ChampsDate;

// Champs numérique géré à part / fonction de sum nécessaire en fonction du contexte
if (TypeRequete <>Tligne) then
    BEGIN
    ListeChampsNum := ',ACT_QTE,ACT_TOTPRCHARGE,ACT_TOTVENTE ';      // ACT_TOTPR
    ListeChampsNumSum := ',SUM(ACT_QTE) as SUM_QTE,SUM(ACT_TOTPRCHARGE) as SUM_TOTPRCHARGE,SUM(ACT_TOTVENTE) as SUM_TOTVENTE '; //SUM(ACT_TOTPR) as SUM_TOTPR
    If ALimValoRess_Art then ListeChamps:=ListeChamps + ',ACT_TYPEARTICLE,ACT_CODEARTICLE,ACT_ARTICLE,ACT_TYPERESSOURCE,ACT_RESSOURCE,ACT_DATEACTIVITE';
    if (AlimAchat) and (NatAchatEngage<>'') then ListeChamps:=ListeChamps+',ACT_NATUREPIECEG,ACT_SOUCHE ';
    END
else
    BEGIN
    ListeChampsNum := ',GL_QTEFACT,GL_TOTALHT ';
    ListeChampsNumSum := ', SUM(GL_QTEFACT) as SUM_QTEFACT, SUM(GL_TOTALHT) as SUM_TOTALHT ';
    ListeChampsCompl := ',GL_NATUREPIECEG,GL_SOUCHE';
    END;

// Ajout de champs cumul + unité en DERNIER
if (TypeDateTB = tdaSemaine ) Or (TypeDateTB = tdaMois ) then
    BEGIN
    if (TypeDateTB = tdaSemaine ) then ChampsRegroupe := ',' + PrefixeTable + 'SEMAINE,'+PrefixeTable + 'PERIODE';
    if (TypeDateTB = tdaMois)     then ChampsRegroupe := ',' + PrefixeTable + 'PERIODE';
    ChampsRegroupe := ChampsRegroupe + ', ' + ChampsUnite;
    END
else ChampsRegroupe := ', '+ ChampsUnite;

// Gestion du Where
if (TypeRequete <>Tligne) then stWhere := FabriqueWherePRFA ('ACT',CodeEnr,TypeRequete,bJoinArt,bJoinRes,Enr) else stWhere := FabriqueWherePRFA ('GL',CodeEnr,Typerequete,bJoinArt,bJoinRes,Enr);
// intégration des jointures : Attention apès le Fabrique where car détermine si besoin de jointures
if (bJoinArt) then JoinArticle  := ' LEFT OUTER JOIN ARTICLE ON '+ PrefixeTable + 'ARTICLE = GA_ARTICLE';
if (bJoinRes) then JoinRessource:= ' LEFT OUTER JOIN RESSOURCE ON '+ PrefixeTable + 'RESSOURCE = ARS_RESSOURCE';

// Requête de base
SQL := 'SELECT '+ ChampsCle + ListeChamps + ChampsRegroupe + ListeChampsCompl+
        ListeChampsNum + 'FROM '+ NomTable +' '+ JoinAffaire + JoinTiers + JoinTiersCompl + JoinArticle + JoinRessource + ' '+ stWhere;
if (TypeDateTB = tdaDetail ) then
    BEGIN
    // Gestion de l'order by
    if (TypeTB  = ttbAffaire)     then stOrderBy := ' Order By '+ PrefixeTable+ 'AFFAIRE'   else
    if (TypeTB  = ttbTiers)       then stOrderBy := ' Order By '+ PrefixeTable+ 'TIERS'     else
    if (TypeTB  = ttbRessource)   then stOrderBy := ' Order By '+ PrefixeTable+ 'RESSOURCE' ;
    Result := SQL + stOrderBy ;
    END
else
    BEGIN
    // Traitement de type cumul global ou sur la période ou sur la semaine
    // Group By direct sur la table ACTIVITE sur les champs regroupés
    stOrderBy := ' GROUP BY ' + ChampsCle + ListeChamps + ChampsRegroupe + ListeChampsCompl;
    SQL := 'SELECT '+ ChampsCle + ListeChamps + ChampsRegroupe + ListeChampsCompl + ListeChampsNumSum
            + 'FROM '+ NomTable + JoinAffaire + JoinTiers + JoinTiersCompl + JoinArticle + JoinRessource+ ' '+ stWhere + stOrderBy;
    Result := SQL;
    END;
// Stockage de la liste des champs pour maj auto de ces champs dans la table TB
//mcd 09/07 if (TypeRequete = tActivite) then
if (TypeRequete <>Tligne) then
   BEGIN
   ListeChampsAct  := ChampsCle + ListeChamps + ChampsRegroupe;
   ListeChampsTot := ListeChamps + ChampsRegroupe + ListeChampsNum ;
   if (TypeDateTB = tdaDetail ) then ListeChampsTot := ChampsDate + ListeChampsTot;
   END;
if (TypeRequete = tLigne) then
   BEGIN
   ListeChampsLigne := ChampsCle + ListeChamps + ChampsRegroupe;
   ListeChampsTot := ListeChampsTot + ListeChampsNum ;
   END;
END;

Function TParamTBAffaire.FabriqueRequeteCut(CodeEnr,TypeEnr : string;Enr : T_TypeEnr)  : String;
Var ChampsRegroupe,  ChampsCle,ListeChamps, ListeChampsNum, ListeChampsCompl,
    ListeChampsNumSum, ChampsUnite, ChampsDate, stWhere,
    PrefixeTable, NomTable, JoinAffaire, JoinTiers, JoinArticle, JoinTiersCompl,
    JoinRessource,SQL,StOrderBy : String;
    bJoinArt,bJoinRes : Boolean;
BEGIN
ChampsRegroupe := ''; ListeChamps := '' ; ChampsCle := '';
SQL := ''; stOrderBy := '';
ListeChampsNum := ''; ListeChampsNumSum := ''; ListeChampsCompl := '';
Result := ''; StWhere :='';
JoinAffaire := '';  JoinTiers:=''; JoinArticle:=''; JoinTiersCompl:=''; JoinRessource:='';
bJoinArt:=false; bJoinRes :=false;
(**Initialisation + définition du champs unité champs géré à part car placé en fin de group by ou d'order by
   car les lignes sur une même clé différenciées uniquement sur l'unité sont regroupées
   classification :
   - ChampsCle champs cle repris en premier dans le group by non repris ds champs selectifs
   - Listechamps : liste des champs texte repris
   - ListechampsNum : champs numérique direct repris en détail
   - ListechampsNumSum : champs numérique sommés pour group by ( cas du cumul semaine, mois, global)
   - ChampsRegroupe : Champs repris dans le group by. *)
PrefixeTable:= 'ACU_';         NomTable := 'AFCUMUL';
ChampsDate := 'ACU_DATE';
ChampsUnite:= '';
// Traitement des champs texte non regroupés en fonction du paramétrage
AlimListeChampsStCut (ChampsCle,ListeChamps,bJoinArt,bJoinRes);
if (TypeDateTB = tdaDetail ) then ListeChamps:= ListeChamps + ','+ ChampsDate;
// Champs numérique géré à part / fonction de sum nécessaire en fonction du contexte
if TypeEnr='CVE' then begin
   ListeChampsNum := ',ACU_AAE,ACU_FAE,ACU_PCA ';      // ACT_TOTPR
   ListeChampsNumSum := ',SUM(ACU_AAE) as SUM_AAE,SUM(ACU_FAE) as SUM_FAE ,SUM(ACU_PCA) as SUM_PCA  ';
   end
else begin
   ListeChampsNum := ',ACU_QTEFACT,ACU_PVFACT ';
   ListeChampsNumSum := ',SUM(ACU_QTEFACT) as SUM_QTEFACT,SUM(ACU_PVFACT) as SUM_PVFACT  ';
     end;
// Gestion du Where    // Mcd 03/06/02 ajout type d'enrgt
stWhere := FabriqueWhereCutOff(CodeEnr,TypeEnr,bJoinArt,bJoinRes,Enr);
// intégration des jointures : Attention apès le Fabrique where car détermine si besoin de jointures
if (bJoinArt) then JoinArticle  := ' LEFT OUTER JOIN ARTICLE ON '+ PrefixeTable + 'ARTICLE = GA_ARTICLE';
if (bJoinRes) then JoinRessource:= ' LEFT OUTER JOIN RESSOURCE ON '+ PrefixeTable + 'RESSOURCE = ARS_RESSOURCE';
// Ajout de champs cumul par mois mcd 03/06/2002
if ((TypeDateTB = tdaSemaine ) Or (TypeDateTB = tdaMois ))
   and ((CutOffMois) or (TypeEnr<>'CVE')) and (Enr<>TAn)then
    BEGIN
    if (TypeDateTB = tdaSemaine ) then ChampsRegroupe := ',' + PrefixeTable + 'SEMAINE,'+PrefixeTable + 'PERIODE';
    if (TypeDateTB = tdaMois)     then ChampsRegroupe := ',' + PrefixeTable + 'PERIODE';
    // pas d'unité dans la table ChampsRegroupe := ChampsRegroupe + ', ' + ChampsUnite;
    END ;
// pas d'unité dans la table ...else ChampsRegroupe := ', '+ ChampsUnite;
// Requête de base
if (TypeDateTB = tdaDetail ) then ListeChamps:=ListeCHamps+ ',ACU_TYPEAC';  //mcd 22/01/03 sinon plante si détail
SQL := 'SELECT '+ ChampsCle + ListeChamps + ChampsRegroupe + ListeChampsCompl+
        ListeChampsNum + 'FROM '+ NomTable +' '+ JoinAffaire + JoinTiers + JoinTiersCompl + JoinArticle + JoinRessource + ' '+ stWhere;
if (TypeDateTB = tdaDetail ) then
    BEGIN
    // Gestion de l'order by
    if (TypeTB  = ttbAffaire)     then stOrderBy := ' Order By '+ PrefixeTable+ 'AFFAIRE'   else
    if (TypeTB  = ttbTiers)       then stOrderBy := ' Order By '+ PrefixeTable+ 'TIERS'     else
    if (TypeTB  = ttbRessource)   then stOrderBy := ' Order By '+ PrefixeTable+ 'RESSOURCE' ;
    Result := SQL + stOrderBy ;
    END
else
    BEGIN
    // Traitement de type cumul global ou sur la période ou sur la semaine
    // Group By direct sur la table ACTIVITE sur les champs regroupés
    stOrderBy := ' GROUP BY ' + ChampsCle + ListeChamps + ChampsRegroupe + ListeChampsCompl;
    if (VH_GC.GAAchatSeria) and (TypeEnr='CVE') then  begin
       stOrderBy:=StOrderby+',ACU_TYPEAC';    // si traitement cut off, et fact fournissuer, on ajoute traitement enrg CAC
       ListeChamps:=ListeCHamps+ ',ACU_TYPEAC';
       end;
    SQL := 'SELECT '+ ChampsCle + ListeChamps + ChampsRegroupe + ListeChampsCompl + ListeChampsNumSum
            + 'FROM '+ NomTable + JoinAffaire + JoinTiers + JoinTiersCompl + JoinArticle + JoinRessource+ ' '+ stWhere + stOrderBy;
    Result := SQL;
    END;
ListeChampsCutOff  := ChampsCle + ListeChamps + ChampsRegroupe;
//mcd 1ListeChampsTot := ListeChampsTot+ ListeChampsNum ;
//mcd 18/10 if (TypeDateTB = tdaDetail ) then ListeChampsTot := ChampsDate + ListeChampsTot;
END;


Function TParamTBAffaire.FabriqueRequeteBud(CodeEnr,TypeEnr : string;Enr : T_TypeEnr)  : String;
Var ChampsRegroupe, ChampsCle,ListeChamps, ListeChampsNum, ListeChampsCompl,
    ListeChampsNumSum, ChampsUnite, ChampsDate, stWhere,
    PrefixeTable, NomTable, JoinAffaire, JoinTiers, JoinArticle, JoinTiersCompl,
    JoinRessource,SQL,StOrderBy : String;
    bJoinArt,bJoinRes : Boolean;
BEGIN
ChampsRegroupe := ''; ListeChamps := '' ; ChampsCle := '';
SQL := ''; stOrderBy := '';
ListeChampsNum := ''; ListeChampsNumSum := ''; ListeChampsCompl := '';
Result := ''; StWhere :='';
JoinAffaire := '';  JoinTiers:=''; JoinArticle:=''; JoinTiersCompl:=''; JoinRessource:='';
bJoinArt:=false; bJoinRes :=false;
(**Initialisation + définition du champs unité champs géré à part car placé en fin de group by ou d'order by
   car les lignes sur une même clé différenciées uniquement sur l'unité sont regroupées
   classification :
   - ChampsCle champs cle repris en premier dans le group by non repris ds champs selectifs
   - Listechamps : liste des champs texte repris
   - ListechampsNum : champs numérique direct repris en détail
   - ListechampsNumSum : champs numérique sommés pour group by ( cas du cumul semaine, mois, global)
   - ChampsRegroupe : Champs repris dans le group by. *)
PrefixeTable:= 'ABU_';         NomTable := 'AFBUDGET';
ChampsDate := 'ABU_DATEBUD';
ChampsUnite:= '';
// Traitement des champs texte non regroupés en fonction du paramétrage
AlimListeChampsStBud (ChampsCle,ListeChamps,bJoinArt,bJoinRes);
if (TypeDateTB = tdaDetail ) then ListeChamps:= ListeChamps + ','+ ChampsDate;
ListeChampsNum := ',ABU_QTE,ABU_MTPVBUD,ABU_MTPRBUD ';
ListeChampsNumSum := ',SUM(ABU_QTE) as SUM_QTE,SUM(ABU_MTPVBUD) as SUM_MTPVBUD,SUM(ABU_MTPRBUD) as SUM_MTPRBUD ';

// Gestion du Where
stWhere := FabriqueWhereBud(CodeEnr,TypeEnr,bJoinArt,bJoinRes,Enr);
// intégration des jointures : Attention apès le Fabrique where car détermine si besoin de jointures
if (bJoinArt) then JoinArticle  := ' LEFT OUTER JOIN ARTICLE ON '+ PrefixeTable + 'ARTICLE = GA_ARTICLE';
if (bJoinRes) then JoinRessource:= ' LEFT OUTER JOIN RESSOURCE ON '+ PrefixeTable + 'RESSOURCE = ARS_RESSOURCE';
if ((TypeDateTB = tdaSemaine ) Or (TypeDateTB = tdaMois )) then
    BEGIN
    if (TypeDateTB = tdaSemaine ) then ChampsRegroupe := ',' + PrefixeTable + 'SEMAINE,'+PrefixeTable + 'PERIODE';
    if (TypeDateTB = tdaMois)     then ChampsRegroupe := ',' + PrefixeTable + 'PERIODE';
    END ;
SQL := 'SELECT '+ ChampsCle + ListeChamps + ChampsRegroupe + ListeChampsCompl+
        ListeChampsNum + 'FROM '+ NomTable +' '+ JoinAffaire + JoinTiers + JoinTiersCompl + JoinArticle + JoinRessource + ' '+ stWhere;
if (TypeDateTB = tdaDetail ) then
    BEGIN
    // Gestion de l'order by
    if (TypeTB  = ttbAffaire)     then stOrderBy := ' Order By '+ PrefixeTable+ 'AFFAIRE'   else
    if (TypeTB  = ttbTiers)       then stOrderBy := ' Order By '+ PrefixeTable+ 'TIERS'     else
    if (TypeTB  = ttbRessource)   then stOrderBy := ' Order By '+ PrefixeTable+ 'RESSOURCE' ;
    Result := SQL + stOrderBy ;
    END
else
    BEGIN
    // Traitement de type cumul global ou sur la période ou sur la semaine
    // Group By direct sur la table ACTIVITE sur les champs regroupés
    stOrderBy := ' GROUP BY ' + ChampsCle + ListeChamps + ChampsRegroupe + ListeChampsCompl;
    SQL := 'SELECT '+ ChampsCle + ListeChamps + ChampsRegroupe + ListeChampsCompl + ListeChampsNumSum
            + 'FROM '+ NomTable + JoinAffaire + JoinTiers + JoinTiersCompl + JoinArticle + JoinRessource+ ' '+ stWhere + stOrderBy;
    Result := SQL;
    END;
ListeChampsBud := ChampsCle + ListeChamps + ChampsRegroupe;
//mcd 18/10/02ListeChampsTot := ListeChamps + ChampsRegroupe + ListeChampsNum ;
//mcd 18/10 if (TypeDateTB = tdaDetail ) then ListeChampsTot := ChampsDate + ListeChampsTot;
END;

// retourne la requete sql du planning
Function TParamTBAffaire.FabriqueRequetePlanning ( CodeEnr : string; Enr : T_TypeEnr)  : String;
var
  vStPrefixeTable, vStNomTable      : String;
  vStListeChamps, vStChampsCle      : String;
  vStChampsDate, vStListeChampsNum  : String;
  vStOrderBy, vStListeChampsNumSum  : String;
  vStJoinAffaire, vStJoinTiers      : String;
  vStJoinArticle, vStJoinTiersCompl : String;
  vStJoinRessource, vStWhere        : String;
  vStChampsRegroupe                 : String;
  vBoJoinArt, vBoJoinRes            : Boolean;

  procedure initVar;
  begin
    vStOrderBy := ''; vStListeChamps := '' ; vStChampsCle := '';
    vStListeChampsNum := ''; vSTListeChampsNumSum := ''; vStWhere :='';
    vStChampsCle := ''; vStListeChamps := '';
    vBoJoinArt := false; vBoJoinRes := false;
  end;

BEGIN

  InitVar;

  vStPrefixeTable := 'APL_';
  vStNomTable := 'AFPLANNING';
  vStChampsDate := 'APL_DATEDEBPLA';

  // Traitement des champs texte non regroupés en fonction du paramétrage
  AlimListeChampsStPlanning (vStChampsCle, vStListeChamps, vBoJoinArt, vBoJoinRes);

  if (TypeDateTB = tdaDetail ) then vStListeChamps := vStListeChamps + ','+ vStChampsDate;
  vStListeChampsNum := ', APL_QTEPLANIFIEE, APL_INITPTPR, APL_INITPTVENTEHT ';
  vStListeChampsNumSum := ',SUM(APL_QTEPLANIFIEE) as SUM_QTEPLANIFIEE, SUM(APL_INITPTPR) AS SUM_INITPTPR, SUM(APL_INITPTVENTEHT) AS SUM_INITPTVENTEHT ';

  vStWhere := FabriqueWherePlanning(CodeEnr, vBoJoinArt, vBoJoinRes,Enr);

  // intégration des jointures : Attention apès le Fabrique where car détermine si besoin de jointures
  if (vBoJoinArt) then vStJoinArticle  := ' LEFT OUTER JOIN ARTICLE ON '+ vStPrefixeTable + 'ARTICLE = GA_ARTICLE';
  if (vBoJoinRes) then vStJoinRessource:= ' LEFT OUTER JOIN RESSOURCE ON '+ vStPrefixeTable + 'RESSOURCE = ARS_RESSOURCE';

  result := 'SELECT '+ vStChampsCle + vStListeChamps + vStChampsRegroupe +
            vStListeChampsNum + 'FROM '+ vStNomTable +' '+ vStJoinAffaire + vStJoinTiers + vStJoinTiersCompl + vStJoinArticle + vStJoinRessource + ' '+ vStWhere;

  if (TypeDateTB = tdaDetail ) then
    BEGIN
      // Gestion de l'order by
      if (TypeTB  = ttbAffaire)         then vStOrderBy := ' Order By '+ vStPrefixeTable+ 'AFFAIRE'
      else if (TypeTB  = ttbTiers)      then vStOrderBy := ' Order By '+ vStPrefixeTable+ 'TIERS'
      else if (TypeTB  = ttbRessource)  then vStOrderBy := ' Order By '+ vStPrefixeTable+ 'RESSOURCE'
      else vStOrderBy := '';
      Result := result + vStOrderBy ;
    END
  else
      BEGIN
      // Traitement de type cumul global ou sur la période ou sur la semaine
      // Group By direct sur la table ACTIVITE sur les champs regroupés
      vStOrderBy := ' GROUP BY ' + vStChampsCle + vStListeChamps + vStChampsRegroupe;
      result := 'SELECT '+ vStChampsCle + vStListeChamps + vStChampsRegroupe + vStListeChampsNumSum
              + 'FROM '+ vStNomTable + vStJoinAffaire + vStJoinTiers + vStJoinArticle + vStJoinRessource+ ' '+ vStWhere + vStOrderBy;
      END;

  ListeChampsPLanning := vStChampsCle + vSTListeChamps + vStChampsRegroupe;
 // mcd 18/10/02 ListeChampsTot := vStListeChamps + vStChampsRegroupe + vStListeChampsNum ;
// mcd 18/10   if (TypeDateTB = tdaDetail ) then ListeChampsTot := vStChampsDate + ListeChampsTot;

end;

Function TParamTBAffaire.TestCompleteCriteresTB : Boolean;
Var i: integer;
    DateDeb, Datefin : TDateTime;
    bOK{,bCritAffaire} :  Boolean;

BEGIN

  Result := True; //bCritAffaire := False;
  CompleteCriteresTB; // alim des champs where pour fabrique requete

  // Test des bornes début supérieures aux bornes de fin
  for i := 1 to MaxCritTB do
    BEGIN
      if (FListeCriteres[i].Champ <> '') then  // Champs critéres utilisé
        BEGIN
          bOK := True;
          if FListeCriteres[i].bTypeDate then
            BEGIN
              DateDeb := StrToDate (FListeCriteres[i].BorneDebut);
              DateFin := StrToDate (FListeCriteres[i].BorneFin);
              if DateDeb > DateFin then bOk := False;
            END
          else if (FListeCriteres[i].BorneDebut > FListeCriteres[i].BorneFin ) then bOK := False;
          if Not(bOK) then
            BEGIN
              if (FListeCriteres[i].BorneFin = '' ) then
                 FListeCriteres[i].BorneFin := FListeCriteres[i].BorneDebut
              else
                BEGIN
                  if FListeCriteres[i].bTypeDate then FListeCriteres[i].BorneDebut := '01/01/1900'
                                                else FListeCriteres[i].BorneDebut :='';
                END;
            END;
          if (FListeCriteres[i].BorneFin = '' ) then
            BEGIN
              if FListeCriteres[i].bTypeDate then FListeCriteres[i].BorneFin := '31/12/2099'
                                             else FListeCriteres[i].BorneFin :='ZZZ';
            END;
        END;
    END;

  if (TypeDateTB = tdaDetail) then NiveauTB := tniTous;    // Valeur forcée en détail
  if (TypeDateTB <> tdaDetail) And (NiveauTB = tniTous ) then Result := false;
END;

Procedure TParamTBAffaire.InitCriteresTB;
Var i: integer;
BEGIN
for i:=1 to MaxCritTB do
    BEGIN
    FListeCriteres[i].Champ := '';
    FListeCriteres[i].BorneDebut := ''; FListeCriteres[i].BorneFin :='';
    FListeCriteres[i].bTypeDate := False ; FListeCriteres[i].bReel := False ;
    FListeCriteres[i].bMultiVal := False ;
    FListeCriteres[i].WhereTiers:= False;FListeCriteres[i].WhereAffaire:= False;
    FListeCriteres[i].WhereAct  := False;FListeCriteres[i].WhereLigne:= False;
    FListeCriteres[i].WhereRes  := False;FListeCriteres[i].WhereArt:= False;
    FListeCriteres[i].WhereCutOff  := False;
    FListeCriteres[i].WhereBud  := False;
    FListeCriteres[i].WherePlanning  := False;
    END;
END;

Procedure TParamTBAffaire.CompleteCriteresTB;
Var i: integer;
BEGIN
  for i:=1 to MaxCritTB do
    BEGIN
      With  FListeCriteres[i] do
        BEGIN
          if (Champ = '') then exit;
          WhereTiers:=False;  WhereAffaire:=False;  WhereAct:=False; WhereCutOff:=False;
          WhereLigne:=False; WhereRes:=false; WhereArt:=False; WhereBud:=False;
          WherePlanning:=False;
          WhereTiersCompl := False; // par défaut non repris
          if Champ = 'ATB_AFFAIRE'     then BEGIN WhereAffaire:=True;  END else
          if Champ = 'ATB_STATUTAFFAIRE'    then  BEGIN WhereAffaire:=True;  END else
          if (Champ='ATB_AFFAIRE1') or (Champ='ATB_AFFAIRE2') or (Champ='ATB_AFFAIRE3') then
              BEGIN WhereAffaire:=True; END else
          if Champ = 'ATB_AFFAIREREF'  then BEGIN WhereAffaire:=True;  END else
          if Champ = 'ATB_TIERS'       then BEGIN WhereTiers:=True;  WhereAffaire:=True;  END else
          if Champ = 'ATB_SOCIETEGROUPE' then BEGIN WhereTiers:=True;    END else    //mcd 06/03/03
          if Champ = 'ATB_DATE'        then BEGIN WhereAct:=True;  WhereLigne:=True;WhereCutOff:=True; WhereBud:=True; WherePLanning:=True; END else
          if Champ = 'ATB_MOISCLOTURE' then BEGIN WhereTiers:=True;  END else
          if Copy(Champ,1,12) = 'ATB_LIBREAFF' then BEGIN WhereAffaire:=True; END else
          if (Champ='ATB_AFFRESSOURCE1') or (Champ='ATB_AFFRESSOURCE2') or (Champ='ATB_AFFRESSOURCE3') then
              BEGIN WhereAffaire:=True; Champ:='ATB_RESSOURCE'+Copy(Champ,17,1); END else
          if (Copy(Champ,1,11)='ATB_AFFDATE')  then
              BEGIN
              WhereAffaire:=True;
              Champ:='ATB_'+Copy(Champ,8,StrLen(Pchar(champ)));
              END else
          if Champ = 'ATB_DEPARTEMENT' then BEGIN WhereAffaire:=True;  END else
          if Champ = 'ATB_ETABLISSEMENT' then BEGIN WhereAffaire:=True;  END else
          if Champ = 'ATB_ETATAFFAIRE' then BEGIN WhereAffaire:=True;  END else
          if Champ = 'ATB_ADMINISTRATIF' then BEGIN WhereAffaire:=True;  END else
          if Champ = 'ATB_RESPONSABLE' then BEGIN WhereAffaire:=True;  END else
          if Champ = 'ATB_GROUPECONF' then BEGIN WhereAffaire:=True;  END else
          if Champ = 'ATB_TYPEARTICLE' then BEGIN WhereAct:=True; WhereCutOff:=True; WhereLigne:=True; WhereBud:=True; WherePLanning:=True; END else
          if Champ = 'ATB_CODEARTICLE' then BEGIN WhereAct:=True; WhereCutOff:=True; WhereLigne:=True; WhereBud:=True; WherePlanning:=True; END else
          if Champ = 'ATB_ARTICLE' then BEGIN WhereAct:=True; WhereCutOff:=True; WhereLigne:=True; WhereBud:=True; WherePlanning:=True; END else
          if (Champ='ATB_FAMILLENIV1') or (Champ='ATB_FAMILLENIV2') or (Champ='ATB_FAMILLENIV3') then
              BEGIN WhereArt:=True; END else
          if Champ = 'ATB_ACTIVITEREPRIS' then BEGIN WhereAct:= True; END else
          if Champ = 'ATB_TYPERESSOURCE' then BEGIN WhereAct:= True; WhereCutOff:=True;WhereBud:=True;END else
          if Champ = 'ATB_RESSOURCE'   then BEGIN WhereAct:= True;WhereCutOff:=True;WhereBud:=True; WherePlanning:=True; END else
          if Champ = 'ATB_LIBELLETIERS'then BEGIN WhereTiers:=True; Champ := 'ATB_LIBELLE';  END else
          if Champ = 'ATB_MOISCLOTURE' then BEGIN WhereTiers:=True; END else
          if Champ = 'ATB_FERME' then BEGIN WhereTiers:=True; END else  //mcd 05/02/03 pour avoir la possibilité de faire une sélection sur ferme
          if Copy(Champ,1,14) = 'ATB_LIBRETIERS' then BEGIN WhereTiersCompl:=True; END else
          if (Champ='ATB_RES1TIERS') or (Champ='ATB_RES2TIERS') or (Champ='ATB_RES3TIERS') then
              BEGIN WhereTiersCompl:=True; Champ:='ATB_RESSOURCE'+Copy(Champ,8,1); END;
        END;
    END;
END;

// ********* Création du tabeau de correspondance entre les zones des différentes tables *********
Function TParamTBAffaire.InitLiaisonTables : Boolean ;
Var i : integer;
BEGIN
  i := 1; Result := True;
  (**Col 1=Tableau de bord  Col 2=Table Activité  **** Col 3 Lignes Facture *** Col 4=Lignes de prévu  **COL 5 =afcumul *   **col 6 Budget **col 7 Planning *)
  // Correspondances Données Activitées
  XTab[1,i]:='TYPEHEURE';     XTab[2,i]:='TYPEHEURE';     XTab[3,i]:='';          XTab[4,i]:='';       XTab[5,i]:='';       XTab[6,i]:='';              XTab[7,i]:='';                   Inc(i);
  XTab[1,i]:='ACTIVITEREPRIS';XTab[2,i]:='ACTIVITEREPRIS';XTab[3,i]:='';          XTab[4,i]:='';       XTab[5,i]:='';       XTab[6,i]:='';              XTab[7,i]:='';                   Inc(i);
  XTab[1,i]:='UNITE';         XTab[2,i]:='UNITE';         XTab[3,i]:='';          XTab[4,i]:='';       XTab[5,i]:='';       XTab[6,i]:='';              XTab[7,i]:='';                   Inc(i);
  XTab[1,i]:='QTE';           XTab[2,i]:='QTE';           XTab[3,i]:='';          XTab[4,i]:='';       XTab[5,i]:='';       XTab[6,i]:='';              XTab[7,i]:='';                   Inc(i);
  XTab[1,i]:='TOTPR';         XTab[2,i]:='TOTPR';         XTab[3,i]:='';          XTab[4,i]:='';       XTab[5,i]:='';       XTab[6,i]:='';              XTab[7,i]:='';                   Inc(i);
  XTab[1,i]:='TOTPRCHARGE';   XTab[2,i]:='TOTPRCHARGE';   XTab[3,i]:='';          XTab[4,i]:='';       XTab[5,i]:='';       XTab[6,i]:='';              XTab[7,i]:='';                   Inc(i);
  XTab[1,i]:='TOTPRCHINDI';   XTab[2,i]:='TOTPRCHINDI';   XTab[3,i]:='';          XTab[4,i]:='';       XTab[5,i]:='';       XTab[6,i]:='';              XTab[7,i]:='';                   Inc(i);
  XTab[1,i]:='TOTVENTEACT';   XTab[2,i]:='TOTVENTE';      XTab[3,i]:='';          XTab[4,i]:='';       XTab[5,i]:='';       XTab[6,i]:='';              XTab[7,i]:='';                   Inc(i);
  XTab[1,i]:='DATE';          XTab[2,i]:='DATEACTIVITE';  XTab[3,i]:='DATEPIECE'; XTab[4,i]:='';       XTab[5,i]:='DATE';   XTab[6,i]:='DATEBUD';       XTab[7,i]:='DATEDEBPLA';     Inc(i);
  XTab[1,i]:='FOLIO';         XTab[2,i]:='FOLIO';         XTab[3,i]:='';          XTab[4,i]:='';       XTab[5,i]:='';       XTab[6,i]:='';              XTab[7,i]:='';                   Inc(i);
  XTab[1,i]:='LIBELLE';       XTab[2,i]:='LIBELLE';       XTab[3,i]:='LIBELLE';   XTab[4,i]:='';       XTab[5,i]:='';       XTab[6,i]:='';              XTab[7,i]:='';                   Inc(i);
  XTab[1,i]:='TYPEHEURE';     XTab[2,i]:='TYPEHEURE';     XTab[3,i]:='';          XTab[4,i]:='';       XTab[5,i]:='';       XTab[6,i]:='';              XTab[7,i]:='';                   Inc(i);
  XTab[1,i]:='NATUREPIECEG';  XTab[2,i]:='NATUREPIECEG';  XTab[3,i]:='NATUREPIECEG';    XTab[4,i]:='';       XTab[5,i]:='NATUREPIECEG';       XTab[6,i]:='';              XTab[7,i]:='';                   Inc(i);
                                  //mcd 28/07/03 mis enaturepieceg dans naturepeiceg et souche, dans souche !.....
            // pour ligne et afcumul, on met la nature de pièce dans la souche pour impression FAC ou AVC ... si plus tard modif pour impresion nature+souche, modifier ici plus requête
  XTab[1,i]:='SOUCHE';        XTab[2,i]:='SOUCHE';        XTab[3,i]:='SOUCHE';          XTab[4,i]:='';       XTab[5,i]:='SOUCHE';       XTab[6,i]:=''; XTab[7,i]:='';  Inc(i);
  XTab[1,i]:='NUMEROACHAT';   XTab[2,i]:='NUMERO';        XTab[3,i]:='NUMERO';          XTab[4,i]:='';       XTab[5,i]:='NUMERO';        XTab[6,i]:=''; XTab[7,i]:='';  Inc(i);
  XTab[1,i]:='INDICEG';       XTab[2,i]:='INDICEG';       XTab[3,i]:='';                XTab[4,i]:='';       XTab[5,i]:='INDICEG';       XTab[6,i]:=''; XTab[7,i]:='';  Inc(i);
  XTab[1,i]:='FOURNISSEUR';   XTab[2,i]:='FOURNISSEUR';   XTab[3,i]:='';                XTab[4,i]:='';       XTab[5,i]:='';              XTab[6,i]:=''; XTab[7,i]:='';  Inc(i);
  // Correspondances Données Lignes Facture
  XTab[1,i]:='FACQTE';        XTab[2,i]:='';              XTab[3,i]:='QTEFACT';   XTab[4,i]:='';       XTab[5,i]:='QTEFACT';   XTab[6,i]:='';  XTab[7,i]:='';     Inc(i);
  XTab[1,i]:='FACPV';         XTab[2,i]:='';              XTab[3,i]:='TOTALHT';   XTab[4,i]:='';       XTab[5,i]:='PVFACT';    XTab[6,i]:='';  XTab[7,i]:='';     Inc(i);
  XTab[1,i]:='FACUNITE';      XTab[2,i]:='';              XTab[3,i]:='QUALIFQTEVTE';XTab[4,i]:='';     XTab[5,i]:='';          XTab[6,i]:='';  XTab[7,i]:='';     Inc(i);
  // Correspondances Données Lignes prévu
  XTab[1,i]:='PREVUQTE';      XTab[2,i]:='';              XTab[3,i]:='';          XTab[4,i]:='QTEFACT';XTab[5,i]:='';       XTab[6,i]:='QTE';      XTab[7,i]:='QTEPLANIFIEE';     Inc(i);
  XTab[1,i]:='PREVUPV';       XTab[2,i]:='';              XTab[3,i]:='';          XTab[4,i]:='TOTALHT';XTab[5,i]:='';       XTab[6,i]:='MTPVBUD';  XTab[7,i]:='INITPTVENTEHT';    Inc(i);
  XTab[1,i]:='PREVUPR';       XTab[2,i]:='';              XTab[3,i]:='';          XTab[4,i]:='DPR';    XTab[5,i]:='';       XTab[6,i]:='MTPRBUD';  XTab[7,i]:='INITPTPR';         Inc(i);
  // Correspondances Données Affaires tiers
  XTab[1,i]:='TIERS';         XTab[2,i]:='TIERS';         XTab[3,i]:='TIERS';     XTab[4,i]:='TIERS';    XTab[5,i]:='TIERS';            XTab[6,i]:='TIERS';    XTab[7,i]:='TIERS';     Inc(i);
  XTab[1,i]:='AFFAIRE';       XTab[2,i]:='AFFAIRE';       XTab[3,i]:='AFFAIRE';   XTab[4,i]:='AFFAIRE';  XTab[5,i]:='AFFAIRE';          XTab[6,i]:='AFFAIRE';  XTab[7,i]:='AFFAIRE';   Inc(i);
  XTab[1,i]:='AFFAIRE1';      XTab[2,i]:='AFFAIRE1';      XTab[3,i]:='AFFAIRE1';  XTab[4,i]:='AFFAIRE1'; XTab[5,i]:='AFFAIRE1';         XTab[6,i]:='AFFAIRE1'; XTab[7,i]:='AFFAIRE1';  Inc(i);
  XTab[1,i]:='AFFAIRE2';      XTab[2,i]:='AFFAIRE2';      XTab[3,i]:='AFFAIRE2';  XTab[4,i]:='AFFAIRE2'; XTab[5,i]:='AFFAIRE2';         XTab[6,i]:='AFFAIRE2'; XTab[7,i]:='AFFAIRE2';  Inc(i);
  XTab[1,i]:='AFFAIRE3';      XTab[2,i]:='AFFAIRE3';      XTab[3,i]:='AFFAIRE3';  XTab[4,i]:='AFFAIRE3'; XTab[5,i]:='AFFAIRE3';         XTab[6,i]:='AFFAIRE3'; XTab[7,i]:='AFFAIRE3';  Inc(i);
  XTab[1,i]:='GROUPECONF';    XTab[2,i]:='GROUPECONF';    XTab[3,i]:='';          XTab[4,i]:='';         XTab[5,i]:='';                 XTab[6,i]:='';         XTab[7,i]:='';          Inc(i);
  XTab[1,i]:='LIBELLEAFF';    XTab[2,i]:='LIBELLE';       XTab[3,i]:='';          XTab[4,i]:='';         XTab[5,i]:='';                 XTab[6,i]:='';         XTab[7,i]:='';          Inc(i);  //mcd 18/11/02
  XTab[1,i]:='LIBELLETIERS';  XTab[2,i]:='LIBELLE';       XTab[3,i]:='';          XTab[4,i]:='';         XTab[5,i]:='';                 XTab[6,i]:='';         XTab[7,i]:='';          Inc(i);  // mcd 18/11/02

  // Correspondances Données ressources
  XTab[1,i]:='RESSOURCE';     XTab[2,i]:='RESSOURCE';     XTab[3,i]:='RESSOURCE';     XTab[4,i]:='RESSOURCE';    XTab[5,i]:='RESSOURCE';      XTab[6,i]:='RESSOURCE';       XTab[7,i]:='RESSOURCE';     Inc(i);
  XTab[1,i]:='TYPERESSOURCE'; XTab[2,i]:='TYPERESSOURCE'; XTab[3,i]:='';              XTab[4,i]:='';             XTab[5,i]:='TYPERESSOURCE';  XTab[6,i]:='TYPERESSOURCE';   XTab[7,i]:='TYPERESSOURCE'; Inc(i);
  XTab[1,i]:='FONCTION';      XTab[2,i]:='FONCTION';      XTab[3,i]:='';              XTab[4,i]:='';             XTab[5,i]:='';               XTab[6,i]:='';                XTab[7,i]:='FONCTION';      Inc(i);
  XTab[1,i]:='LIBRERES1';     XTab[2,i]:='LIBRERES1';     XTab[3,i]:='LIBRERES1';     XTab[4,i]:='LIBRERES1';    XTab[5,i]:='LIBRERES1';      XTab[6,i]:='';                XTab[7,i]:='';              Inc(i);
  XTab[1,i]:='LIBRERES2';     XTab[2,i]:='LIBRERES2';     XTab[3,i]:='LIBRERES2';     XTab[4,i]:='LIBRERES2';    XTab[5,i]:='LIBRERES2';      XTab[6,i]:='';                XTab[7,i]:='';              Inc(i);
  XTab[1,i]:='LIBRERES3';     XTab[2,i]:='LIBRERES3';     XTab[3,i]:='LIBRERES3';     XTab[4,i]:='LIBRERES3';    XTab[5,i]:='LIBRERES3';      XTab[6,i]:='';                XTab[7,i]:='';              Inc(i);
  XTab[1,i]:='LIBRERES4';     XTab[2,i]:='LIBRERES4';     XTab[3,i]:='LIBRERES4';     XTab[4,i]:='LIBRERES4';    XTab[5,i]:='LIBRERES4';      XTab[6,i]:='';                XTab[7,i]:='';              Inc(i);
  XTab[1,i]:='LIBRERES5';     XTab[2,i]:='LIBRERES5';     XTab[3,i]:='LIBRERES5';     XTab[4,i]:='LIBRERES5';    XTab[5,i]:='LIBRERES5';      XTab[6,i]:='';                XTab[7,i]:='';              Inc(i);
  XTab[1,i]:='LIBRERES6';     XTab[2,i]:='LIBRERES6';     XTab[3,i]:='LIBRERES6';     XTab[4,i]:='LIBRERES6';    XTab[5,i]:='LIBRERES6';      XTab[6,i]:='';                XTab[7,i]:='';              Inc(i);
  XTab[1,i]:='LIBRERES7';     XTab[2,i]:='LIBRERES7';     XTab[3,i]:='LIBRERES7';     XTab[4,i]:='LIBRERES7';    XTab[5,i]:='LIBRERES7';      XTab[6,i]:='';                XTab[7,i]:='';              Inc(i);
  XTab[1,i]:='LIBRERES8';     XTab[2,i]:='LIBRERES8';     XTab[3,i]:='LIBRERES8';     XTab[4,i]:='LIBRERES8';    XTab[5,i]:='LIBRERES8';      XTab[6,i]:='';                XTab[7,i]:='';              Inc(i);
  XTab[1,i]:='LIBRERES9';     XTab[2,i]:='LIBRERES9';     XTab[3,i]:='LIBRERES9';     XTab[4,i]:='LIBRERES9';    XTab[5,i]:='LIBRERES9';      XTab[6,i]:='';                XTab[7,i]:='';              Inc(i);
  XTab[1,i]:='LIBRERESA';     XTab[2,i]:='LIBRERESA';     XTab[3,i]:='LIBRERESA';     XTab[4,i]:='LIBRERESA';    XTab[5,i]:='LIBRERESA';      XTab[6,i]:='';                XTab[7,i]:='';              Inc(i);
  XTab[1,i]:='LIBELLERES';    XTab[2,i]:='ARS_LIBELLE';   XTab[3,i]:='ARS_LIBELLE';   XTab[4,i]:='ARS_LIBELLE';  XTab[5,i]:='ARS_LIBELLE';    XTab[6,i]:='ARS_LIBELLE';     XTab[7,i]:='ARS_LIBELLE';   Inc(i); //mcd 18/11/02
  XTab[1,i]:='LIBELLE2RES';   XTab[2,i]:='ARS_LIBELLE2';  XTab[3,i]:='ARS_LIBELLE2';  XTab[4,i]:='ARS_LIBELLE2'; XTab[5,i]:='ARS_LIBELLE2';   XTab[6,i]:='ARS_LIBELLE2';    XTab[7,i]:='ARS_LIBELLE2';  Inc(i); // mcd 18/11/02

  // Correspondances Données articles
  XTab[1,i]:='ARTICLE';       XTab[2,i]:='ARTICLE';       XTab[3,i]:='ARTICLE';   XTab[4,i]:='ARTICLE';      XTab[5,i]:='ARTICLE';         XTab[6,i]:='ARTICLE';      XTab[7,i]:='ARTICLE';     Inc(i);
  XTab[1,i]:='CODEARTICLE';   XTab[2,i]:='CODEARTICLE';   XTab[3,i]:='CODEARTICLE';XTab[4,i]:='CODEARTICLE'; XTab[5,i]:='CODEARTICLE';     XTab[6,i]:='CODEARTICLE';  XTab[7,i]:='CODEARTICLE'; Inc(i);
  XTab[1,i]:='ARTICLE';       XTab[2,i]:='ARTICLE';       XTab[3,i]:='ARTICLE';   XTab[4,i]:='ARTICLE';       XTab[5,i]:='ARTICLE';        XTab[6,i]:='ARTICLE';      XTab[7,i]:='ARTICLE'; Inc(i);
  XTab[1,i]:='TYPEARTICLE';   XTab[2,i]:='TYPEARTICLE';   XTab[3,i]:='TYPEARTICLE';XTab[4,i]:='TYPEARTICLE'; XTab[5,i]:='TYPEARTICLE';     XTab[6,i]:='TYPEARTICLE';  XTab[7,i]:='TYPEARTICLE'; Inc(i);
  XTab[1,i]:='FAMILLENIV1';   XTab[2,i]:='FAMILLENIV1';   XTab[3,i]:='FAMILLENIV1';XTab[4,i]:='FAMILLENIV1'; XTab[5,i]:='FAMILLENIV1';     XTab[6,i]:='FAMILLENIV1';  XTab[7,i]:='FAMILLENIV1';            Inc(i);
  XTab[1,i]:='FAMILLENIV2';   XTab[2,i]:='FAMILLENIV2';   XTab[3,i]:='FAMILLENIV2';XTab[4,i]:='FAMILLENIV2'; XTab[5,i]:='FAMILLENIV2';     XTab[6,i]:='FAMILLENIV2';  XTab[7,i]:='FAMILLENIV2';            Inc(i);
  XTab[1,i]:='FAMILLENIV3';   XTab[2,i]:='FAMILLENIV3';   XTab[3,i]:='FAMILLENIV3';XTab[4,i]:='FAMILLENIV3'; XTab[5,i]:='FAMILLENIV3';     XTab[6,i]:='FAMILLENIV3';  XTab[7,i]:='FAMILLENIV3';            Inc(i);
  XTab[1,i]:='LIBREART1';     XTab[2,i]:='LIBREART1';     XTab[3,i]:='LIBREART1';  XTab[4,i]:='LIBREART1';   XTab[5,i]:='LIBREART1';       XTab[6,i]:='LIBREART1';  XTab[7,i]:='LIBREART1';            Inc(i);
  XTab[1,i]:='LIBREART2';     XTab[2,i]:='LIBREART2';     XTab[3,i]:='LIBREART2';  XTab[4,i]:='LIBREART2';   XTab[5,i]:='LIBREART2';       XTab[6,i]:='LIBREART2';  XTab[7,i]:='LIBREART2';            Inc(i);
  XTab[1,i]:='LIBREART3';     XTab[2,i]:='LIBREART3';     XTab[3,i]:='LIBREART3';  XTab[4,i]:='LIBREART3';   XTab[5,i]:='LIBREART3';       XTab[6,i]:='LIBREART3';  XTab[7,i]:='LIBREART3';            Inc(i);
  XTab[1,i]:='LIBREART4';     XTab[2,i]:='LIBREART4';     XTab[3,i]:='LIBREART4';  XTab[4,i]:='LIBREART4';   XTab[5,i]:='LIBREART4';       XTab[6,i]:='LIBREART4';  XTab[7,i]:='LIBREART4';            Inc(i);
  XTab[1,i]:='LIBREART5';     XTab[2,i]:='LIBREART5';     XTab[3,i]:='LIBREART5';  XTab[4,i]:='LIBREART5';   XTab[5,i]:='LIBREART5';       XTab[6,i]:='LIBREART5';  XTab[7,i]:='LIBREART5';            Inc(i);
  XTab[1,i]:='LIBREART6';     XTab[2,i]:='LIBREART6';     XTab[3,i]:='LIBREART6';  XTab[4,i]:='LIBREART6';   XTab[5,i]:='LIBREART6';       XTab[6,i]:='LIBREART6';  XTab[7,i]:='LIBREART6';            Inc(i);
  XTab[1,i]:='LIBREART7';     XTab[2,i]:='LIBREART7';     XTab[3,i]:='LIBREART7';  XTab[4,i]:='LIBREART7';   XTab[5,i]:='LIBREART7';       XTab[6,i]:='LIBREART7';  XTab[7,i]:='LIBREART7';            Inc(i);
  XTab[1,i]:='LIBREART8';     XTab[2,i]:='LIBREART8';     XTab[3,i]:='LIBREART8';  XTab[4,i]:='LIBREART8';   XTab[5,i]:='LIBREART8';       XTab[6,i]:='LIBREART8';  XTab[7,i]:='LIBREART8';            Inc(i);
  XTab[1,i]:='LIBREART9';     XTab[2,i]:='LIBREART9';     XTab[3,i]:='LIBREART9';  XTab[4,i]:='LIBREART9';   XTab[5,i]:='LIBREART9';       XTab[6,i]:='LIBREART9';  XTab[7,i]:='LIBREART9';            Inc(i);
  XTab[1,i]:='LIBREARTA';     XTab[2,i]:='LIBREARTA';     XTab[3,i]:='LIBREARTA';  XTab[4,i]:='LIBREARTA';   XTab[5,i]:='LIBREARTA';       XTab[6,i]:='LIBREARTA';  XTab[7,i]:='LIBREARTA';            Inc(i);
  XTab[1,i]:='LIBELLEART';    XTab[2,i]:='GA_LIBELLE';    XTab[3,i]:='GA_LIBELLE'; XTab[4,i]:='GA_LIBELLE';  XTab[5,i]:='GA_LIBELLE';      XTab[6,i]:='GA_LIBELLE'; XTab[7,i]:='GA_LIBELLE';           Inc(i); //mcd 18/11/02
  // Correspondances Calcul de dates
  XTab[1,i]:='SEMAINE';       XTab[2,i]:='SEMAINE';       XTab[3,i]:='SEMAINE';   XTab[4,i]:=' ';              XTab[5,i]:='SEMAINE';        XTab[6,i]:='SEMAINE';  XTab[7,i]:='SEMAINE';    Inc(i);
  XTab[1,i]:='PERIODE';       XTab[2,i]:='PERIODE';       XTab[3,i]:='PERIODE';   XTab[4,i]:=' ';              XTab[5,i]:='PERIODE';        XTab[6,i]:='PERIODE';  XTab[7,i]:='PERIODE';    Inc(i);
  // Correspondances CutOff
  XTab[1,i]:='AAE';       XTab[2,i]:='';                  XTab[3,i]:='';          XTab[4,i]:='';              XTab[5,i]:='AAE';            XTab[6,i]:='';  XTab[7,i]:=''; Inc(i);
  XTab[1,i]:='FAE';       XTab[2,i]:='';                  XTab[3,i]:='';          XTab[4,i]:='';              XTab[5,i]:='FAE';            XTab[6,i]:='';  XTab[7,i]:=''; Inc(i);
  XTab[1,i]:='PCA';       XTab[2,i]:='';                  XTab[3,i]:='';          XTab[4,i]:='';              XTab[5,i]:='PCA';            XTab[6,i]:='';  XTab[7,i]:=''; Inc(i);
  if (i > MaxCorrespChamps ) then BEGIN PGIBoxAF ('Limite Liaison inter-tables dépassée','Tableau de Bord'); Result := false; END;
END;

Function TParamTBAffaire.ChampsTableVersTB (NomChamps : String; NumTable : integer ; Sens: T_TableTB  ): String;
Var     i,RefDico, TraducDico : integer;
BEGIN
Result := '';
// C.B 11/09/02
if (NumTable < 2) or (NumTable > 7) then Exit;
if NomChamps ='' then Exit;
if Sens = ttaTableTB then BEGIN RefDico := NumTable; TraducDico := 1; END
                     else BEGIN RefDico := 1; TraducDico := NumTable; END;
for i:=1 to MaxCorrespChamps do if XTab[RefDico,i]=AnsiUppercase(NomChamps) then
  BEGIN result:=XTab[TraducDico,i] ; Break; END;
END;

Procedure TParamTBAffaire.ChargeCriteres (Contain : TComponent);
Var i,j, PosPrefixe ,PosCrit , CptCrit : integer;
    Vr: Variant;
    Name : String;
    bFin, ChampTrouve, TypeDate, bReel,MultiVal : Boolean;
BEGIN
CptCrit := 0;
InitCriteresTB;

for i:=0 to Contain.ComponentCount-1 do
    BEGIN
    Name := Contain.Components[i].Name;
    Name := FindEtReplace(Name,'ATBX','ATB_',False);
    PosPrefixe := Pos ('ATB_', Name);
    if (PosPrefixe > 0) And Not(Contain.Components[i] is THLabel) then
        BEGIN
        TypeDate := False; bReel := false; MultiVal := false;
        Vr := '';
        if Copy(Name,Length(Name),1)='_' then BEGIN bFin:= True; Name:= Copy(Name,1,Length(Name)-1); END
                                         else bFin := False;
        if (Contain.Components[i] is TControl) and (TControl(Contain.Components[i]).Parent is TTabSheet) then
            BEGIN
            if Contain.Components[i] is THEdit         then
                BEGIN
                Case THEdit(Contain.Components[i]).OpeType of
                    otDate   : BEGIN Vr := THEdit(Contain.Components[i]).Text ; TypeDate := True; END;
                    otReel   : BEGIN Vr := Valeur(THEdit(Contain.Components[i]).Text) ; bReel := True; END;
                    otString : Vr := THEdit(Contain.Components[i]).Text ;
                    end ;
                END else
            if Contain.Components[i] is THValComboBox  then
                BEGIN
                Vr:= THValComboBox(Contain.Components[i]).Value;
                END else
            if Contain.Components[i] is THMultiValComboBox  then
                BEGIN
                Vr:= THMultiValComboBox(Contain.Components[i]).GetSQLValue;
                Vr := FindEtReplace(Vr,'ATBX','ATB_',True);
                MultiVal := true;
                END else
            if Contain.Components[i] is TCheckBox      then
                BEGIN
                Case TCheckBox(Contain.Components[i]).State of
                   //mcd 05/02/03 il ne faut pas mettre cette valeur, car pas en fichier cbGrayed   : Vr := '|' ;
                    cbGrayed   : Vr := '' ;  //mcd 05/02/03
                    cbChecked  : Vr := 'X' ;
                    cbUnChecked: Vr := '-' ;
                    End;
                END else
            if Contain.Components[i] is THCritMaskEdit then
                BEGIN
                Case THCritMaskEdit(Contain.Components[i]).OpeType of
                    otDate   : BEGIN Vr := THCritMaskEdit(Contain.Components[i]).Text ; TypeDate := True; END;
                    otReel   : Vr := Valeur(THCritMaskEdit(Contain.Components[i]).Text) ;
                    otString : Vr := THCritMaskEdit(Contain.Components[i]).Text ;
                    end ;
                END;
          //Mcd 11/04/01  if (vr <> '') then
            if (VartoStr(vr) <>'')  then
                BEGIN
                ChampTrouve := False;
                for j:=1 to MaxCritTB do
                    if FListeCriteres[j].Champ=Name then BEGIN ChampTrouve := True; break; END;
                if ChampTrouve then PosCrit:= j else BEGIN Inc(CptCrit); PosCrit:= CptCrit; END;
                if (PosCrit > MaxCritTB ) then BEGIN PGIBoxAF ('Limite des critères de sélection dépassée','Tableau de Bord'); Exit; END;
                // Alimentation de la liste des critères
                if Not(ChampTrouve) then FListeCriteres[PosCrit].Champ := Name;
                if bFin then FListeCriteres[PosCrit].BorneFin := Vr else FListeCriteres[PosCrit].BorneDebut := Vr;
                FListeCriteres[PosCrit].bTypeDate := TypeDate;
                FListeCriteres[PosCrit].bReel := bReel;
                FListeCriteres[PosCrit].bMultiVal := MultiVal;
                END;
            END ;
        END;
    END;
END;

//******************************************************************************
//*****************Fonct d'init appelée depuis la fiche de lancement ***********
//******************************************************************************
Procedure TParamTBAffaire.MajMonoAffaire (CodeAffaire,DateDeb,DateFin : string; VisuSousAffaire : Boolean);
Var PosCrit : integer;
BEGIN
if codeAffaire = '' then BEGIN MonoAffaire := False ; Exit; END;
PosCrit := 1;
InitCriteresTB;
// Intégration du critère affaire ou affaire de référence
if Not(VisuSousAffaire) then
   begin
   FListeCriteres[PosCrit].Champ := 'ATB_AFFAIRE';
   FListeCriteres[PosCrit].BorneFin := CodeAffaire;
   FListeCriteres[PosCrit].BorneDebut := CodeAffaire;
   FListeCriteres[PosCrit].bTypeDate := False;
   MonoAffaire := True; MonoTiers := False;
   end
else
   begin
   FListeCriteres[PosCrit].Champ := 'ATB_AFFAIREREF';
   FListeCriteres[PosCrit].BorneFin := CodeAffaire;
   FListeCriteres[PosCrit].BorneDebut := CodeAffaire;
   FListeCriteres[PosCrit].bTypeDate := False;
   MonoAffaire := False; MonoTiers := False;
   end;
// Intégration du critère dates
PosCrit := 2;
FListeCriteres[PosCrit].Champ := 'ATB_DATE';
FListeCriteres[PosCrit].BorneDebut := DateDeb;
FListeCriteres[PosCrit].BorneFin := Datefin;
FListeCriteres[PosCrit].bTypeDate := True;
END;

procedure TParamTBAffaire.MajMonoTiers   (CodeTiers,DateDeb,DateFin : string);
var
   PosCrit : integer;
BEGIN
     if codeTiers = '' then BEGIN MonoTiers := False ; Exit; END;

     PosCrit := 1;      // Intégration du critère client
     InitCriteresTB;
     FListeCriteres[PosCrit].Champ := 'ATB_TIERS';
     FListeCriteres[PosCrit].BorneFin := CodeTiers;
     FListeCriteres[PosCrit].BorneDebut := CodeTiers;
     FListeCriteres[PosCrit].bTypeDate := False;
     MonoTiers := True; MonoAffaire := False;

     PosCrit := 2;     // // Intégration du critère dates
     FListeCriteres[PosCrit].Champ := 'ATB_DATE';
     FListeCriteres[PosCrit].BorneDebut := DateDeb;
     FListeCriteres[PosCrit].BorneFin := Datefin;
     FListeCriteres[PosCrit].bTypeDate := True;
END;

procedure TParamTBAffaire.MajEcart(NumEcart: integer; Formule: hstring);
BEGIN
     if (NumEcart < 1) or (NumEcart > 5) then Exit;
     if Formule = '' then Exit;
     if Not(TestFormule(formule)) then PgiBoxAF ('Formule de l''écart '+ IntToStr(NumEcart) + ' invalide', 'calcul des écarts');
     if NumEcart = 1 then BEGIN FormuleEcart1 := '{'+Formule+'}'; GereEcart1 := true; END else
     if NumEcart = 2 then BEGIN FormuleEcart2 := '{'+Formule+'}'; GereEcart2 := true; END;
     if NumEcart = 3 then BEGIN FormuleEcart3 := '{'+Formule+'}'; GereEcart3 := true; END;
     if NumEcart = 4 then BEGIN FormuleEcart4 := '{'+Formule+'}'; GereEcart4 := true; END;
     if NumEcart = 5 then BEGIN FormuleEcart5 := '{'+Formule+'}'; GereEcart5 := true; END;
END;

//$$$jp 20/12/2002: pour fixer les tiers (pour lien OLE)
procedure TParamTBAffaire.MajMultiTiers (strTiersDeb, strTiersFin, strDateDeb, strDateFin:string);
var
   PosCrit : integer;
BEGIN
     MonoTiers   := FALSE;
     MonoAffaire := FALSE;
     if (strTiersDeb = '') AND (strTiersFin = '') then
        exit;

     PosCrit := 1;
     InitCriteresTB;
     FListeCriteres[PosCrit].Champ      := 'ATB_TIERS';
     FListeCriteres[PosCrit].BorneDebut := strTiersDeb;
     FListeCriteres[PosCrit].BorneFin   := strTiersFin;
     FListeCriteres[PosCrit].bTypeDate  := FALSE;

     PosCrit := 2;
     FListeCriteres[PosCrit].Champ      := 'ATB_DATE';
     FListeCriteres[PosCrit].BorneDebut := strDateDeb;
     FListeCriteres[PosCrit].BorneFin   := strDateFin;
     FListeCriteres[PosCrit].bTypeDate  := TRUE;
end;
//$$$jp fin

{***********A.G.L.Privé.*****************************************
Auteur  ...... : MC DESSEIGNET
Créé le ...... : 13/06/2003
Modifié le ... :   /  /    
Description .. : fct qui boucle sur tous les paramêtre de T_Paramtb et qui 
Suite ........ : les stocke dans les critères à gérér
Mots clefs ... : 
*****************************************************************}
procedure TParamTBAffaire.MajSelect (Zone : T_ParamTb);
var
   PosCrit : integer;
BEGIN
     MonoTiers   := FALSE;
     MonoAffaire := FALSE;
     InitCriteresTB;
     PosCrit :=1;
     If (zone.cltdeb <>'') or (ZOne.cltfin <>'') then MajUneZone('ATB_TIERS',Zone.cltdeb,zone.cltfin,PosCrit,False,False);
     If (zone.MoisCLotDeb >0) or (Zone.MoisclotFin <13) then MajUneZone('ATB_MOISCLOTURE',Zone.MoisCLotDeb,zone.MoisCLotFIn,PosCrit,False,false);
     If (zone.Affdeb <>'') or (ZOne.Afffin <>'') then MajUneZone('ATB_AFFAIRE',Zone.affdeb,zone.afffin,PosCrit,False,False);
     If (zone.ActiviteReprise <>'')  then MajUneZone('ATB_ACTIVITEREPRIS',Zone.ActiviteReprise,'',PosCrit,False,False);
     If (zone.GroupCOnf <>'')  then MajUneZone('ATB_GROUPECONF',Zone.GroupConf,'',PosCrit,False,True);
     If (zone.GroupeClt <>'')  then MajUneZone('ATB_SOCIETEGROUPE',Zone.GroupeClt,'',PosCrit,False,False);
     If (zone.Libre1clt <>'')  then MajUneZone('ATB_LIBRETIERS1',Zone.Libre1clt,'',PosCrit,False,False);
     If (zone.Libre2clt <>'')  then MajUneZone('ATB_LIBRETIERS2',Zone.Libre1clt,'',PosCrit,False,False);
     If (zone.Libre3clt <>'')  then MajUneZone('ATB_LIBRETIERS3',Zone.Libre3clt,'',PosCrit,False,False);
     If (zone.Libre4clt <>'')  then MajUneZone('ATB_LIBRETIERS4',Zone.Libre4clt,'',PosCrit,False,False);
     If (zone.Libre5clt <>'')  then MajUneZone('ATB_LIBRETIERS5',Zone.Libre5clt,'',PosCrit,False,False);
     If (zone.Libre6clt <>'')  then MajUneZone('ATB_LIBRETIERS6',Zone.Libre6clt,'',PosCrit,False,False);
     If (zone.Libre7clt <>'')  then MajUneZone('ATB_LIBRETIERS7',Zone.Libre7clt,'',PosCrit,False,False);
     If (zone.Libre8clt <>'')  then MajUneZone('ATB_LIBRETIERS8',Zone.Libre8clt,'',PosCrit,False,False);
     If (zone.Libre9clt <>'')  then MajUneZone('ATB_LIBRETIERS9',Zone.Libre9clt,'',PosCrit,False,False);
     If (zone.LibreAclt <>'')  then MajUneZone('ATB_LIBRETIERSA',Zone.LibreAclt,'',PosCrit,False,False);
     If (zone.REss1clt <>'')  then MajUneZone('ATB_RES1TIERS',Zone.Ress1Clt,'',PosCrit,False,False);
     If (zone.REss2clt <>'')  then MajUneZone('ATB_RES2TIERS',Zone.Ress2Clt,'',PosCrit,False,False);
     If (zone.REss3clt <>'')  then MajUneZone('ATB_RES3TIERS',Zone.Ress3Clt,'',PosCrit,False,False);
     If (zone.EtatAff <>'')  then MajUneZone('ATB_ETATAFFAIRE',Zone.EtatAff,'',PosCrit,False,true);
     If (zone.StatutAff <>'')  then MajUneZone('ATB_STATUTAFFAIRE',Zone.StatutAff,'',PosCrit,False,True);
     If (zone.DptAff <>'')  then MajUneZone('ATB_DEPARTEMENT',Zone.DptAff,'',PosCrit,False,true);
     If (zone.ETbAff <>'')  then MajUneZone('ATB_ETABLISSEMENT',Zone.EtbAff,'',PosCrit,False,True);
     If (zone.REspAff <>'')  then MajUneZone('ATB_Responsable',Zone.RespAff,'',PosCrit,False,False);
     If (zone.REss1AFF <>'')  then MajUneZone('ATB_AFFRESSOURCE1',Zone.Ress1Aff,'',PosCrit,False,False);
     If (zone.REss2AFF <>'')  then MajUneZone('ATB_AFFRESSOURCE2',Zone.Ress2Aff,'',PosCrit,False,False);
     If (zone.REss3AFF <>'')  then MajUneZone('ATB_AFFRESSOURCE3',Zone.Ress3Aff,'',PosCrit,False,False);
     If (zone.Libre1Aff <>'')  then MajUneZone('ATB_LIBREAFF1',Zone.Libre1Aff,'',PosCrit,False,False);
     If (zone.Libre2Aff <>'')  then MajUneZone('ATB_LIBREAFF2',Zone.Libre2Aff,'',PosCrit,False,False);
     If (zone.Libre3Aff <>'')  then MajUneZone('ATB_LIBREAFF3',Zone.Libre3Aff,'',PosCrit,False,False);
     If (zone.Libre4Aff <>'')  then MajUneZone('ATB_LIBREAFF4',Zone.Libre4Aff,'',PosCrit,False,False);
     If (zone.Libre5Aff <>'')  then MajUneZone('ATB_LIBREAFF5',Zone.Libre5Aff,'',PosCrit,False,False);
     If (zone.Libre6Aff <>'')  then MajUneZone('ATB_LIBREAFF6',Zone.Libre6Aff,'',PosCrit,False,False);
     If (zone.Libre7Aff <>'')  then MajUneZone('ATB_LIBREAFF7',Zone.Libre7Aff,'',PosCrit,False,False);
     If (zone.Libre8Aff <>'')  then MajUneZone('ATB_LIBREAFF8',Zone.Libre8Aff,'',PosCrit,False,False);
     If (zone.Libre9Aff <>'')  then MajUneZone('ATB_LIBREAFF9',Zone.Libre9Aff,'',PosCrit,False,False);
     If (zone.LibreAAff <>'')  then MajUneZone('ATB_LIBREAFFA',Zone.LibreAAff,'',PosCrit,False,False);
     If (DateTostr(zone.Datedeb) <>'01/01/1900') or (DatetoStr(ZOne.Datefin) <>'31/12/2099') then MajUneZone('ATB_DATE',DateToStr(Zone.DAtedeb),DateToStr(zone.Datefin),PosCrit,True,False);
     If (DateTostr(zone.DatedebAffDEb) <>'01/01/1900') or (DatetoStr(ZOne.DatefinAffDEb) <>'31/12/2099') then MajUneZone('ATB_AFFDATEDEBUT',DateToStr(Zone.DAtedebAffDEb),DateToStr(zone.DatefinAffDeb),PosCrit,True,False);
     If (DateTostr(zone.DatedebAffFin) <>'01/01/1900') or (DatetoStr(ZOne.DatefinAffFin) <>'31/12/2099') then MajUneZone('ATB_AFFDATEFIN',DateToStr(Zone.DAteDEbAffDEB),DateToStr(zone.DatefinAffFin),PosCrit,True,False);
     If (DateTostr(zone.DatedebClot) <>'01/01/1900') or (DatetoStr(ZOne.DatefinClot) <>'31/12/2099') then MajUneZone('ATB_AFFDATECLOTTECH',DateToStr(Zone.DAtedebClot),DateToStr(zone.DatefinClot),PosCrit,True,False);
     If (DateTostr(zone.Datedebgar) <>'01/01/1900') or (DatetoStr(ZOne.DatefinGar) <>'31/12/2099') then MajUneZone('ATB_AFFDATEGARANTIE',DateToStr(Zone.DAtedebGar),DateToStr(zone.DatefinGar),PosCrit,True,False);
     If (DateTostr(zone.DatedebExerDEb) <>'01/01/1900') or (DatetoStr(ZOne.DatefinExerDeb) <>'31/12/2099') then MajUneZone('ATB_AFFDATEDEBEXER',DateToStr(Zone.DAtedebExerDeb),DateToStr(zone.DatefinExerDEb),PosCrit,True,False);
     If (DateTostr(zone.DatedebExerFin) <>'01/01/1900') or (DatetoStr(ZOne.DatefinExerFin) <>'31/12/2099') then MajUneZone('ATB_AFFDATEFINEXER',DateToStr(Zone.DAtedebExerFin),DateToStr(zone.DatefinExerFin),PosCrit,True,False);
     If (DateTostr(zone.DatedebInter) <>'01/01/1900') or (DatetoStr(ZOne.DatefinInter) <>'31/12/2099') then MajUneZone('ATB_AFFDATELIMITE',DateToStr(Zone.DAtedebInter),DateToStr(zone.DatefinInter),PosCrit,True,False);
end;

procedure TParamTBAffaire.MajUneZone (NomCHamp: string; Deb,fin :variant; var posit : integer;Date, Multi:boolean);
Var Vr,Critere, Ddeb: string;
begin
     FListeCriteres[Posit].Champ      :=NomChamp;
     FListeCriteres[Posit].BorneDebut := Deb;
     FListeCriteres[Posit].BorneFin   := Fin;
     FListeCriteres[Posit].bTypeDate  := Date;
     If Multi then
       begin   // cas zone en multi sélection, les valeurs sont séparés par ;, on fabrique requête
       Vr :='AND ( ' ;
       Ddeb :=deb;
       Critere :=ReadTokenSt(Ddeb);
       if (critere ='') and (Length(deb) <>0) then
         begin  // on prévoit si on accepte de géré "blanc aussi " cas groupe conf
         Vr :=VR + Nomchamp+ '="" OR ' ;
         Critere :=ReadTokenSt(Ddeb);
         end;
       While (Critere <>'') do
         BEGIN
         if Critere<>'' then Vr := Vr + Nomchamp+ '="' + critere+'" OR ';
         Critere :=ReadTokenSt(Ddeb);
         end;
          // on enlève le dernier OR qui a été mis dans la requête
       Vr := Copy(VR, 1,Length(vr) -3) + ')';
       FListeCriteres[Posit].BorneDebut :=Vr ;
       FListeCriteres[Posit].bMultiVal := true;
       end ;
     inc (posit);
end;

Function TParamTBAffaire.ListeChampsTBSelectif (AvecVirgule : Boolean) : String;
Var tmp, st, st2,ChampsAffaire: string;
    i, ColDico :integer;
    TraiteLigne : Boolean;
Begin

  Result := '';
  ColDico := 1;

  // les champs de référence sont repris de l'activité si alimentée
  st := ListeChampsTot;
  While St<>'' do
    BEGIN
      TraiteLigne := False;

      Tmp:= ReadTokenSTV(St);
      // Traitement activité
      i:=Pos('ACT_',tmp) ; if (i > 0) then ColDico := nACT;
      // Traitement Ligne
      if (i = 0) then
        BEGIN  i:=Pos('GL_',tmp); if (i > 0) then ColDico := nGL; TraiteLigne := True; END;
      if (i = 0) then
        BEGIN  i:=Pos('APL_',tmp); if (i > 0) then ColDico := nPlanning;  END;
      // Traitement articles ils suivent la logique activité
      if (i = 0) then
        BEGIN i:= Pos('GA_',tmp); if (i > 0) then ColDico := nAct; END;

      if i > 0 then
        BEGIN
          tmp := Copy (tmp, pos('_',tmp)+1,length(tmp)-pos('_',tmp));  tmp := Trim(tmp);
          tmp := Trim(tmp);
          st2 := ChampsTableVersTB(tmp,ColDico,ttaTableTB);

          if st2 <> '' then
            BEGIN
              if Result <> '' then Result := Result+ ';' + PrefixeTB + St2
                              else Result :=  PrefixeTB + St2 ;
            END;

          if (TraiteLigne) then //Traitement des lignes plusieures zones différentes peuvent être alim
           BEGIN
           if AlimPrevuAff then
               BEGIN   // recup des champs du prevu pour l'affichage de la tob
               ColDico := nGLPR;
               st2 := ChampsTableVersTB(tmp,ColDico,ttaTableTB);
               if st2 <> '' then
                   BEGIN
                   if Result <> '' then Result := Result+ ';' + PrefixeTB + St2
                                   else Result :=  PrefixeTB + St2 ;
                   END;
               END;
           END;
         END;
    END;

  // Champs LibPeriode au lieu du champs période
 // mcd 18/10/2002  Result := FindEtReplace(Result,';ATB_PERIODE',';ATB_LIBPERIODE',False);
// ajout MCD 18/10/02 pour traiter les champs à retournées si pas sélect all, en fct des tables demandés
//  If (AlimPrevu) and (not AlimPrevuAff) then Result:=Result+  ';ATB_PREVUQTE;ATB_PREVUPV;ATB_PREVUPR';
  // C.B 03/04/2003 on met les champs systématiquement
  // mcd 12/09/03 ?? suite modif CB on a parfois les champs en double??? je change
  If (AlimPrevu) then
   begin
      //Result:=Result+  ';ATB_PREVUQTE;ATB_PREVUPV;ATB_PREVUPR;ATB_PREQTEUNITEREF';
   Result:=Result+  ';ATB_PREVUPR;ATB_PREQTEUNITEREF';
   If Pos('ATB_PREVUPV',Result)=0 then Result:=Result+  ';ATB_PREVUQTE;ATB_PREVUPV';  //mcd
   end;
  if AlimBoni then Result:=Result+';ATB_BONIQTE;ATB_BONIPR;ATB_BONIVENTE;ATB_BONIQTEUNITERE';
  If ALimValoRess_Art then Result:=Result +';ATB_PRRESS;ATB_PVRESS;ATB_PRART;ATB_PVART';
    //mcd 12/09/03 ajout test POS,car parfois en double
 If (AlimFacture)
     and (GetParamSOc('SO_AFFACTPARRES')<>'SAN')
     and (Pos('ATB_FACPV',Result)=0) then Result:=Result+';ATB_FACPV';
  if NatAchatEngage<>'' then Result:=Result+';ATB_ENGAGEQTE;ATB_ENGAGEPV;ATB_ENGAGEPR';
// fin mcd
  // Retraitement pour supprimer le PR ou PV de l'activité en fonction du choix
  if MtActivite = ActPR then Result := FindEtReplace(Result,';ATB_TOTVENTEACT','',False) else
  if MtActivite = ActPV then Result := FindEtReplace(Result,';ATB_TOTPRCHARGE','',False);
     //mcd 18/11/02
  if (NiveauTb=TniTous) or (NiveauTb=TniRes) or (NiveauTB=TniArtRes)
      or (NiveauTb=TniTypeArtRes) then Result  :=Result+';ATB_LIBELLERES;ATB_LIBELLE2RES' ;
  if (NiveauTb=TniTous) or (NiveauTb=TniArt)
      or (NiveauTb=TniArtRes) then Result := Result  +';ATB_LIBELLEART';

  // Ajout en entête
  if (DetailParAffaire) then
    BEGIN
      { mcd V575 17/01/02 ChampsAffaire := 'ATB_AFFAIRE;ATB_AFFAIRE1;';
      if (VH_GC.CleAffaire.Co2Visible) then ChampsAffaire := ChampsAffaire + 'ATB_AFFAIRE2;';
      if (VH_GC.CleAffaire.Co3Visible) then ChampsAffaire := ChampsAffaire + 'ATB_AFFAIRE3;';
      Result := ChampsAffaire + Result;  }
      ChampsAffaire := 'ATB_AFFAIRE;ATB_AFFAIRE1';
      if (VH_GC.CleAffaire.Co2Visible) then ChampsAffaire := ChampsAffaire + ';ATB_AFFAIRE2';
      if (VH_GC.CleAffaire.Co3Visible) then ChampsAffaire := ChampsAffaire + ';ATB_AFFAIRE3';
      if (VH_GC.CleAffaire.GestionAvenant) then   ChampsAffaire := ChampsAffaire + ';ATB_AVENANT'; //mcd 08/04/2003 il faut le passer si gestion ..
      if result<>'' then
        begin
             // mcd 13/06/03 ??? il y a parfois déjà le ';'
        If Copy(Result,1,1) =';' then  Result := ChampsAffaire + Result
         else Result := ChampsAffaire +';'+ Result ;
        end
      else Result := ChampsAffaire ;
     END;
  if AvecssAff then
    BEGIN
      ChampsAffaire := 'ATB_AFFAIREREF;ATB_AFFAIREREF1;';
      if (VH_GC.CleAffaire.Co2Visible) then ChampsAffaire := ChampsAffaire + 'ATB_AFFAIREREF2;';
      if (VH_GC.CleAffaire.Co3Visible) then ChampsAffaire := ChampsAffaire + 'ATB_AFFAIREREF3;';
      Result := ChampsAffaire + Result;
    END;
  // Ajout à la fin des champs
  if GereEcart1 then Result := Result+ ';ATB_ECART1';
  if GereEcart2 then Result := Result+ ';ATB_ECART2';
  if GereEcart3 then Result := Result+ ';ATB_ECART3';
  if GereEcart4 then Result := Result+ ';ATB_ECART4';
  if GereEcart5 then Result := Result+ ';ATB_ECART5';
  if ( ChargeTobTB) then Result:=Result+';XXX1;XXX2;XXX3';  // mcd 20/03/02 pour avoir champ formule même si pas reprise intégrale des champs
  if AvecVirgule then Result :=FindEtReplace(Result,';',',',True);
  END;

//******************************************************************************
//********************** Gestions des écarts du TB *****************************
//******************************************************************************
Procedure TParamTBAffaire.CalculLesEcarts(TobaTraiter : TOB);
Var Ecart : String;
    i : integer;
BEGIN
if Not(GereEcart1) and Not(GereEcart2)and Not(GereEcart3)and Not(GereEcart4)and Not(GereEcart5) then Exit;
for i := 0 to TobaTraiter.Detail.Count-1 do
    BEGIN
    TobDetEnc := TobaTraiter.Detail[i];
    if GereEcart1 then
       BEGIN
       FormuleEcart1:='{"#.###,0"'+FormuleEcart1+'}';     //mcd 09/12/02
       Ecart := GFormule(FormuleEcart1,GetDataEcart,nil,1);
       If Ecart='Erreur' then Ecart:='0'; // mcd 23/09/02
         // mcd 23/09/02 TobDetEnc.PutValue ('ATB_ECART1',Format('%n',[StrtoFloat(Ecart)]));
       TobDetEnc.PutValue ('ATB_ECART1',Format('%n',[Valeur(Ecart)]));
       END;
    if GereEcart2 then
       BEGIN
       FormuleEcart2:='{"#.###,0"'+FormuleEcart2+'}';     //mcd 09/12/02
       Ecart := GFormule(FormuleEcart2,GetDataEcart,nil,1);
       If Ecart='Erreur' then Ecart:='0';
       TobDetEnc.PutValue ('ATB_ECART2',Format('%n',[Valeur(Ecart)]));
       END;
    if GereEcart3 then
       BEGIN
       FormuleEcart3:='{"#.###,0"'+FormuleEcart3+'}';     //mcd 09/12/02
       Ecart := GFormule(FormuleEcart3,GetDataEcart,nil,1);
       If Ecart='Erreur' then Ecart:='0';
       TobDetEnc.PutValue ('ATB_ECART3',Format('%n',[Valeur(Ecart)]));
       END;
    if GereEcart4 then
       BEGIN
       FormuleEcart4:='{"#.###,0"'+FormuleEcart4+'}';     //mcd 09/12/02
       Ecart := GFormule(FormuleEcart4,GetDataEcart,nil,1);
       If Ecart='Erreur' then Ecart:='0';
       TobDetEnc.PutValue ('ATB_ECART4',Format('%n',[Valeur(Ecart)]));
       END;
    if GereEcart5 then
       BEGIN
       FormuleEcart5:='{"#.###,0"'+FormuleEcart5+'}';     //mcd 09/12/02
       Ecart := GFormule(FormuleEcart5,GetDataEcart,nil,1);
       If Ecart='Erreur' then Ecart:='0';
       TobDetEnc.PutValue ('ATB_ECART5',Format('%n',[Valeur(Ecart)]));
       END;
    END;
END;

function TParamTBAffaire.GetDataEcart(sS:Hstring):variant;
BEGIN
(*mcd 23/09/02 if TobDetEnc.FieldExists(s) then Result := StrtoFloat(TobDetEnc.GetValue(ss))
                            else Result := '0';  *)
if TobDetEnc.FieldExists(ss) then Result := Valeur(TobDetEnc.GetValue(ss))
                            else Result := 0; 
END;


//******************************************************************************
//********************** Retraitement TOB Affaires ******************************
//******************************************************************************
Procedure TParamTBAffaire.TobAffavecSsAff(TobAffaires:TOB);
var i,j : integer;
    TobDet,TobSsAff: TOB;
    stSQL,CodeSsAff : string;
    QQ : TQuery;

BEGIN
if TobAffaires = Nil then Exit;
For i:=0 to TobAffaires.Detail.Count-1 do
   begin
   TobDet := TobAffaires.Detail[i];
   // boucle sur les affaires de ref pour savoir si toutes les sous affaires sont bien présentes...
   if (Tobdet.GetValue('AFF_ISAFFAIREREF') ='X') then
      begin
      stSQL :=FabriqueRequete ('AFF',TobDet.GetValue('AFF_AFFAIREREF'),TNor);
      if (stSQL <> '') then
         Begin
         QQ := OpenSQL(stSQL,True,-1,'',true);
         if Not QQ.EOF then
            Begin
              TobSsAff := TOB.Create('liste sous affaires',Nil,-1) ;
              try
                TobSsAff.LoadDetailDB('ChampsTB AFFAIRE','','',QQ,False,False);
                for j := TobSsAff.detail.count-1 downto 0 do
                   begin
                   CodeSsAff := TobSsAff.Detail[j].GetValue('AFF_AFFAIRE');
                   if TobAffaires.FindFirst(['AFF_AFFAIRE'],[codeSsAff],false) = Nil then
                      TobSsAff.Detail[j].ChangeParent (Tobaffaires,-1);
                   end;
              finally
                TobSsAff.Free;
              end;
            End;
         Ferme(QQ);
         End;
      end;
   end;
END;


//******************************************************************************
//********************** fonctions utiles / interface ecran ********************
//******************************************************************************
// Interface avec les combo pour recup du bon code (niveau + type TB)
Function GetNiveauTB  (ComboValue : String) : T_NiveauTB ;
BEGIN
if (ComboValue = '1') then Result := tniTiers      else
if (ComboValue = '2') then Result := tniAffaire    else
if (ComboValue = '3') then Result := tniTypeArt    else
if (ComboValue = '4') then Result := tniArt        else
if (ComboValue = '5') then Result := tniTypeArtRes else
if (ComboValue = '6') then Result := tniArtRes     else
if (ComboValue = '7') then Result := tniRes        else
if (ComboValue = '8') then Result := tniTypeRes    else
if (ComboValue = '9') then Result := tniFonction   else
if (Combovalue = '10')then Result := tniFamilleArt1 else
if (Combovalue = '11')then Result := tniFamilleArt2 else
if (Combovalue = '12')then Result := tniFamilleArt3 else
if (Combovalue = '21')then Result := tniStatArt1  else
if (Combovalue = '22')then Result := tniStatArt2  else
if (Combovalue = '23')then Result := tniStatArt3  else
Result := tniArt ;    // Par défaut
END;

Function GetRegroupeTB  (ComboValue : String) : T_TypeDateTB ;
BEGIN
if ComboValue = 'DET' then Result := tdaDetail  else
if ComboValue = 'SEM' then Result := tdaSemaine else
if ComboValue = 'FOL' then Result := tdaFolio   else
if ComboValue = 'MOI' then Result := tdaMois    else
Result := tdaCumul;   // cumul par defaut
END;

function GetTypeTB (NiveauTiers, AvecDetailAff, bMulti : Boolean) : T_TypeTB;
BEGIN
if NiveauTiers then
   BEGIN
   if AvecDetailAff then Result := ttbAffaire else Result := ttbTiers ;
   if bMulti then Result := ttbTiers ;
   END
else Result := ttbAffaire ;
END;

{***********A.G.L.***********************************************
Auteur  ...... : MC DESSEIGNET
Créé le ...... : 13/06/2003
Modifié le ... :   /  /    
Description .. : Fct qui permet de mettre des valeurs par défaut dans le 
Suite ........ : record T_ParamTb
Mots clefs ... : PARAMTB;INIT
*****************************************************************}
Procedure InitRecordParamTb (Var zone : T_PARAMTB);
begin
Zone.cltDEb :='';   Zone.cltFin :='';
Zone.AffDeb :='';   Zone.AffFin :='';
Zone.DateDeb :=Idate1900;     Zone.dateFin :=idate2099;
Zone.regroup :=TdaCumul;    Zone.valoris:= ActprPV;
Zone.NiveauAna :=TniTypeRes;  Zone.Unite := GetParamSoc('SO_AFMESUREACTIVITE');
Zone.GroupConf :='';
Zone.ActiviteReprise :='';    Zone.PieceAchat :='';   Zone.PieceAchatEng:='';
Zone.AlimSolde :=False;
Zone.AlimPrevu :=False;       Zone.AlimFacture:=True; Zone.Alimreal :=true;
Zone.AlimAchat :=False;       Zone.alimBoni :=False;  Zone.AlimAn :=False;
Zone.AffSurPeriode := True;   Zone.AlimVisa :=False;  Zone.AlimCutoff :=False;
Zone.CutOffMois :=False;      Zone.alimMois :=False;  Zone.AvecSSAff :=False;
Zone.libre1clt :='';          Zone.libre2Clt :='';    Zone.libre3Clt :='';
Zone.libre4clt :='';          Zone.libre5Clt :='';    Zone.libre6Clt :='';
Zone.libre7Clt :='';          Zone.libre8Clt :='';    Zone.libre9Clt :='';
Zone.libre4clt :='';          Zone.moisClotDEb:=0;    Zone.MoisClotFin :=12;
Zone.Ress1Clt :='';           Zone.Ress2Clt := '';    Zone.Ress3Clt :='';
Zone.GroupeClt:='';           Zone.EtatAff:='';       Zone.StatutAff := '';
Zone.DptAff :='';             Zone.EtbAff :='';       Zone.RespAff := '';
Zone.libre1Aff :='';          Zone.libre2Aff :='';    Zone.libre3Aff :='';
Zone.libre4Aff :='';          Zone.libre5Aff :='';    Zone.libre6Aff :='';
Zone.libre7Aff :='';          Zone.libre8Aff :='';    Zone.libre9Aff :='';
Zone.libre4Aff :='';
Zone.Ress1Aff :='';           Zone.Ress2Aff := '';    Zone.Ress3Aff :='';
Zone.DateDebAffDeb := idate1900;  Zone.DateFinAffDeb :=Idate2099;
Zone.DateDebAffFIn := idate1900;  Zone.DateFinAffFin :=Idate2099;
Zone.DateDebClot := idate1900;    Zone.DateFinClot :=Idate2099;
Zone.DateDebGar := idate1900;     Zone.DateFinGar :=Idate2099;
Zone.DateDebExerDeb := idate1900; Zone.DateFinExerDeb :=Idate2099;
Zone.DateDebExerFin := idate1900; Zone.DateFinExerFin :=Idate2099;
Zone.DateDebInter := idate1900;   Zone.DateFinInter :=Idate2099;
end;

Function GetDateDebutSemaine ( Semaine, Periode : Integer) : TDateTime;
Var YY,MM : integer ;
    StPeriode : String;
BEGIN
Result := iDate1900;
if (Semaine = 0) or (periode < 111111) then exit;
StPeriode := IntToSTr(Periode);
YY := StrToInt (Copy (StPeriode,1,4));
MM := StrToInt (Copy (StPeriode,5,2));
{$IFDEF AGLINF545}
Result := PremierJourSemaineTempo(Semaine,YY);
{$ELSE}
// PL le 23/05/02 : réparation de la fonction AGL545
if (MM=12) and (Semaine=1) then YY := YY + 1;
if (MM=1) and ((Semaine=52) or (Semaine=53)) then YY := YY - 1;
Result := PremierJourSemaine(Semaine,YY);
////////////////////////////////// AGL545
{$ENDIF}
END;

Function GetDateDebutPeriode ( Periode : Integer) : TDateTime;
Var StPeriode : String;
    YY,MM : Word ;
BEGIN
Result := iDate1900;
if (periode < 111111) then exit;
StPeriode := IntToSTr(Periode);
YY := StrToInt (Copy (StPeriode,1,4));
MM := StrToInt (Copy (StPeriode,5,2));
Result := EncodeDate(YY,MM,01);
END;

end.
