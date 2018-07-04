unit UIMPORTCOM;

interface

uses
classes,
{$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
forms, UtilTrans, UTOB, RecordCom, HEnt1, HCTrls, stdctrls,
Paramsoc, Controls, Windows,
{$IFNDEF EAGLSERVER}
 hmsgbox, HStatus,
{$ELSE}
 esession, UWA, ULibCpContexte, UTOZ, FileCtrl, galMainSauveDossier,
{$ENDIF}
 Sysutils, UtilPGI, ULibAnalytique, ed_tools,
 ULibEcriture, SoldeCpt, RecupUtil, ImSaiCoef,
 UlibExercice, ImpFicU, ControlParam, MajTable,
 Ent1, MesgErrCom, Saisutil, uLibStdCpta, TimpFic, 
 SiscoVersPGI, HashTable, SaisComm, Echeance,
 ULibBonapayer,
 UBob, UProcGen,
 ReCalcTva,
 ugedfiles,
{$IFDEF COMSX}
 TiersPayeur,
{$IFNDEF EAGLSERVER}
 UTOFPGImportFic,
{$ENDIF}
{$ENDIF}
{$IFDEF MODENT1}
  CPTypeCons,
{$ENDIF MODENT1}
 UtilGed;

const
  MaxEnreg = 100;

type

  TInfoEvent = function(Chaine: string; Listecom : TListBox; pourcentage : double=0) : Boolean of object;

  // pointeur sur fonction
  Procimport = procedure (Ligne : string) of object;
  TTypeCar = (tcGene, tcblanc, tcnonblanc, tcAna, tcGed, tcBap, tcnone, tcpiece);


  TFExecimport = record
    Code: string;
    ToDo: Procimport;
  end;

  TFTransfertInfo = record
   iCompte, NatureCpte, Pointable, Regimetva,
   Collectif, TPayeur, RibAuto, Etva : string ;
   OKLettrage : Boolean;
  end;

  TEvtSup = class
    Journal,Codeexo,Datec : string;
    Datecomptable         : TDateTime;
    vperiode              : integer;
    npiece,Etablissement  : string;
    refrevision           : string;
    parpiece              : boolean;
    TWhereEtabY,TWhereEtablissement : string;
    procedure DegageEcriture;
    procedure DegageExercice;
  end;
 Type T_DM = Class
            Cpte   : String17 ;
            DateP : tDateTime ;
            end ;


  TImportCom = class

  private
             FAffInfo                            : TInfoEvent ;
             GrandNbGene,GrandNbAux              : Boolean;
             {$IFDEF EAGLSERVER}
             FicIE                               : TextFile;
             {$ENDIF}
             LISTEIMPORT                         : TListBox;
             OrigineFichier,NatureFichier        : string;
             OrigineEnvoie                       : string;
             VersionFichier                      : string;
             Paramgene                           : Boolean;
             TGen, TGenupdate, TGenInsert        : TOB;
             TTiers, TTiersupdate, TTiersInsert  : TOB;
             TTiersCAU, TRibCAU, TContactCAU     : TOB;
             TEcrParJournal,TBUDEcr,TPieceCon    : TOB;
             TOBAno,TOBPIECE,TOBLettre,TOBJrl    : TOB;
             TDevise                             : TOB;
             AutoriseEcheck,BCtrP,BCtrS          : Boolean;
             BRecalculS,BSoldeLigne,BBlanc       : Boolean;
             stArg, RepriseComplement, CStat     : string;
             NbEcriture,NbEcriturepiece,NbEcrAna : integer;
             ListeEtab                           : TStringList;
             ModeSaisie,numeropiece              : string;
             RupturePiece,Gestiondoublon,BIntegre: Boolean;
             CalculPiece,Modiftiers              : string;
             Savperiode,FCorresp                 : string;
             OKexercice,BControlEtab, GerBAP      : Boolean;
             BInterdireCreationTiers             : Boolean;
             OkCreatCompte,BEcartChange          : Boolean;
             MAJCPEXOREF,IgnorePointage          : Boolean;
             VentAttent                          : Array[1..5] of boolean;
             GestPointage,FichierEntree          : string;
             Changementgene, Changementaux       : Boolean;
             PresenceExercice,BMAJCPTE           : Boolean;
             FiTOBG_LETTRABLE                    : integer;
             FiTOBG_REGIMETVA, FiTOBG_NATUREGENE : integer;
             FiTOBG_POINTABLE, FiTOBT_LETTRABLE  : integer;
             FiTOBT_REGIMETVA, FiTOBT_NATUREAUXI : integer;
             FiTOBG_VENTILABLE1,FiTOBG_VENTILABLE2: integer;
             FiTOBG_VENTILABLE3,FiTOBG_VENTILABLE4: integer;
             FiTOBG_VENTILABLE5                   : integer;
             TiersAutre,OkModePaieDefaut,OkVerAxe : Boolean;
             TypeEcrR,TypeEcrI,QualifOrigine      : string;
             PInfoImp                             : PtTInfoImport ;
             Sansaux,ChargeChoixCod,OkMajExoV8    : Boolean;
             fbGeneLg, fbAuxLg, NbSection         : integer; // fiche 10568
             fbGenecb, fbAuxcb, fbAxe1cb, fbAxe2cb: string;
             fbAxe3cb, fbAxe4cb, fbAxe5cb         : string;
             MontantNegatif,ModeCroisaxe          : Boolean;
             UpdateLettreDiv, BRejet,OkRejet      : Boolean;
             TenueEuro,AttribRIBAuto,BValider     : Boolean ;
             TotalMouvDebit, TotalMouvCredit      :  double;
             ListeRecep                           : array[0..MaxEnreg] of TFExecimport;
             {$IFDEF EAGLSERVER}
             FichierResultat                      : string;
             {$ENDIF}
             RootPath,NoDossier,TypeVisa,AxeTva   : string; // fiche 10568
             DecoupLon                            : integer;
             InterdireLaCreation,VentilType       : string;
             STiers,SGeneraux,SJournal,SSection   : TStringList;
             lpTi, lpG, lpJ, lpS, lpGM            : TList;
             SExercice,SSectionInbase,STlibre     : TStringList;
             SModePaie,STva,SSectionFerme         : TStringList;
             SPays                                : TStringList;
             CpteGene, CpteClient, CpteFour,
             CpteSalaire,CpteDivers,
             RAxe1,RAxe2,RAxe3,RAxe4,RAxe5,
             SubCollCli, SubCollFou,
             Mrdefaut, Regdefaut, MdpDefault      : string;
             fHGene, fHGeneInsert, fHLettre       : THashTable;     // Objet de stockage des tiers et généraux en mémoire
             fHTiers, fHTiersInsert               : THashTable;     // Objet de stockage des tiers et généraux en mémoire
             vInfo                                : TInfoEcriture;
             ZDevise                              : TZDevise;
             TSAnv                                : TList ;
             OptionEchange, OptionTP, ExisteTP    : Boolean;
             EnregBap                             : TBonAPayer;
             SaveLigne, GarderJour                : string;
             AppliquerRest,MajDate,BEquilibre     : Boolean;
             DateEcrChange                        : TDateTime;
             YCotation                            : double;
             BCalculModeregle,ControlExercClos    : Boolean;
             GereVentilType, ModeEVentilType      : Boolean;
             TailleSource, TailleBloc             : double;
             NbLec                                : integer;
             ScanBanck                            : Boolean;
             TModeRegle                           : TOB;
             FRejet                               : TextFile;
             ExerciceSuperVise                    : string;
             SDevise, SEtablissement              : string;
             procedure ChargeToDo;
             {$IFDEF EAGLSERVER}
             Function  AfficheComImport(Chaine: string; Listecom : TListBox; pourcentage : double=0) : Boolean;
             {$ENDIF}
             procedure ImporteEcritureBud(Ligne: string);
             procedure ImporteEcriture(Ligne: string);
             procedure LibereTob;
             procedure Genere_ecratEuro;
             procedure DelettrageSynchro (TPiece : TOB);
             procedure InsertTiers;
             procedure Remplirgeneraux (var T : TOB; Ligne,compte, Nature : string; Casupdate : Boolean=FALSE);
             procedure RemplissageTiers (var T : Tob; compteauxi, Ligne : string; update : Boolean=TRUE; Casupdate : Boolean=FALSE);
             procedure InsertUpdateGeneraux;
             procedure MiseajourSouche;
             procedure MajExoV8;
             procedure ImportParame1(Ligne: string);
             procedure ImportParame2(Ligne: string);
             procedure ImportParame3(Ligne: string);
             procedure ImportParame4(Ligne: string);
             procedure ImportParame5(Ligne: string);
             procedure ImportExercice(Ligne: string);
             procedure ImportTableLibre(Ligne: string);
             procedure ImportSection(Ligne: string);
             procedure ImporteEtablissement(Ligne: string);
             procedure ImporteModepaiement(Ligne: string);
             procedure ImporteModereglement(Ligne: string);
             procedure remplissageMDR (indice: integer; var Tm : TOB);
             procedure ImporteDevise(Ligne: string);
             procedure ImporteRegimeTVA(Ligne: string);
             procedure ImportSectionAnalytique(Ligne: string);
             procedure ImportSouche(Ligne: string);
             procedure InitGeneraux;
             function  OKUpdategeneraux (TG : TOB; Ligne, compte : string; var Nature : string) : Boolean;
             procedure ImporteGeneraux(Ligne: string);
             procedure remplir_compteattente;
             procedure InitTiers (Affiche : Boolean=TRUE); 
             procedure ImporteTiers(Ligne: string);
             procedure ImportRib(Ligne: string);
             procedure ImportJournaux(Ligne: string);
             procedure ImportBanqueCp(Ligne: string);
             procedure ImportBanque(Ligne: string);
             procedure ImportReleveBanque(Ligne: string);
             procedure ImportLigneReleveBanque(Ligne: string);
             procedure ImportContact(Ligne: string);
             procedure ImportTiersComp(Ligne : string);
             function  EnleveBlanc(Chaine: string; x, y: integer ; ValNumeric : String = ''): string;
             procedure InitEcriture;
             function  AlimPiece(Ligne: string; var PieceS1 : string) : Boolean;
             function  RechercheOptimise (const TobSrc: tob ; const Field: string ; const ValeurR: string; RechOpti : Boolean=TRUE; THbl : THashTable=nil) : TOB ;
             function  Okdoublon (Ligne : string; var TT : TOB; CJ : string; FindSurTob : Boolean=FALSE): boolean;
             Function  remplir_ecriture(Ext, Ligne, Modesaisie: string; periode: integer; T: TOB; Codejournal : string; DateEcriture : TDateTime; Exerc : TExoDate; CpteRemplace :string='') : string;
             function  OKVentilable (Compte : string) : Boolean;
             procedure MetVentilable (Compte,Axe : string);
             procedure DecoupageEcriture;
             procedure AttenteAnaTOB ( TOBEcr : TOB ; NbDec : integer );
             procedure MiseajourDetail (var Tlec : TOB; okc,oke,okgrp : boolean);
             Function  NombreExercice(WW : string; Max : Boolean=FALSE): integer;
             Function  RenExoClo (Where : string) : Boolean;
             Procedure ChargeTobTiersGene (sSQl : string; Quoi : Char);
             function  OKUpdateTiers (T : Tob; compteauxi, Ligne : string) : Boolean;
             procedure Insert_journal;
             procedure VerifAxeSection;
             procedure AffectationMontantdoublon(St: string; var debit,credit : double);
             procedure AffectationMontant(St: string; var T: TOB; ext: string);
             procedure AffectationMontantcouverture(St: string; var T: TOB; ext: string);
             function  ExisteCompte (Ext : string; var Tftr : TFTransfertInfo; ForceLettrage : Boolean=TRUE) : Boolean;
             function  RetournCompteattente (nature : string) :string;
             function  MajDernLettrage (Ext, Compte, lettrage : string) : Boolean;
             procedure InsertCompte (Ext, Compte, libelle :string; Var nature : string; Collectif : string=''; EtatLettrage : string='');
             function  RendModePaiementParDefaut (Modereg : string= ''): string ;
             procedure DecoupageJournal(TLec,TLecpiece,TLecPeriode : TOB; var numpiece : integer);
             procedure MiseNumVentil (var Tlec : TOB);
             procedure GenereCompteDattente (TLec : TOB; ie : integer; Soldedeb, SoldeCre : double);
             procedure RemplirNatureSens ( Gene,Auxi : String ; Var Nature,Sens : String ) ;
             Function  TransfertDeSISCO : Boolean;
             Procedure AffMessage(Mess : string);
             Function  TraiteCorrespCpt(Quoi : Integer; CptInput : string; Var CptOut : string) : Boolean ;
             procedure CreationModePaieInexistant;
             procedure ImportCpteBudget(Ligne : string);
             procedure ImportJalBudget(Ligne : string);
             procedure ImportSectionBudget(Ligne : string);
             procedure ImportChoixCode(Ligne : string);
             procedure ImportChoixExt(Ligne : string);
             procedure ImportRessource(Ligne : string);
             procedure ImportParamlib(Ligne : string);
             procedure CreatJournal (CodeJournal : string);
             procedure SuppressionDesdoublons(Var TLecpiece,TlecNew :TOB; Journal : string; ii : integer);
             // Import Tiers CCIMPEX  fiche 10134
             procedure ImporteTiersAncienFormat(Ligne: string);

             procedure ImporteRelance (Ligne : string);
             Function  ImportBobIMMO : Boolean;
             function  AjouterFichierDansGed (Fichier, CodeTiers, CodeGed, Description, Auteur, Annee, Mois : String; var DocId: string): Variant;
             procedure AffectDocGuid ( TOBEcr : TOB; DocGuid : string='');
             procedure AJouterListeSoldes (Compte : string; var ST : TStringList);
             Function  RendMRDefaut : string;
             Function  RendRegTvadefaut : string;
             Function  AutoriseLaCreation (Code, MMessage : string) : Boolean;
             Function  Veriftablelibre (TypeNat,Valeur : string) : string;
             procedure LibereHTable;
             procedure ChargeSection;
             procedure GenereEcartDechange(Var TLecpiece, TlecNew :TOB);
             procedure GestiondesBAP (Tlec : TOB);
             procedure ImporteBapVisa(Ligne: string);
             procedure ImporteBapCircuit(Ligne: string);
             procedure ImporteBonapayer(Ligne: string; exercice, Docid : string; numpiece : integer);
             function  ResSectionExiste(Strc, StAxe, FCompte : string) : Boolean;
             function  TrouveRubriqueDuCpteGene(vStCpte: string; Liasse : Boolean=FALSE): string;
             procedure ImporteCorresp(Ligne: string);
             procedure ImporteCorrespImp(Ligne: string);
             Function  TraiteJockerCorrespCpt(CRTYPE, CptInput : string; Var CptOut : string) : Boolean ;
             procedure ImporteVentiltype (Ligne : string);
             procedure MajChamp (var T: TOB; Champ, Valeur : string; dd : double=0; Numerique : Boolean=FALSE);  // fiche 10441
             procedure SuppressionDesSoldesAZero(Var TLecpiece, TlecNew :TOB; Journal : string; ii : integer);
             // BVE 28.08.07
{$IFDEF CERTIFNF}
             procedure ImporteSuiviValidation (Ligne : string);
{$ENDIF}
             procedure AjoutListeDatedernMvt ( TS : TList; Cpte : string; DateP : tDateTime);
             procedure MAJDateDernMvt;
             procedure RemplirTOBDevise; // BVE 04.07.07
             function  RendCodeAcceptation (MP : string) : string;
             procedure RemplirDateEchance (var T_Mode : T_ModeRegl; DateEcriture : TDateTime; T : TOB);
{$IFDEF EAGLSERVER}
             //function  ImportBobEagl (Fich : string; var TB : TOB) : Boolean;
             function  LanceScruteRepertoire(var RTOB : TOB; sttArg : string) : Boolean;
             Procedure LanceDeZippfile (Filezip : string);
{$ENDIF}
             procedure EcrireFicRejet(Tl : TOB);
             function TrouveCodeIsoPays (Pays : string) : string; // fiche 10586

  public
             constructor Create (TE : Tob; sttArg : string; AffInfo    : TInfoEvent); overload;
             function  ImportComSx (TImp : TOB; SArg : string) : Boolean;
             Property  OnAfficheListeCom  : TInfoEvent read FAffInfo  write FAffInfo ;
             procedure CalcSoldeList (var LJB : TList; fb : TFichierBase);
             Function ReChRibPrincipal(TA : TOB; Nature : string) : String ;
  end;


Function LanceImport (var RequestTOB : TOB; sttArg : string; AffInfo    : TInfoEvent; var PImport : TImportCom) : Boolean;

implementation


uses
  {$IFDEF MODENT1}
  CPProcGen,
  CPProcMetier,
  {$ENDIF MODENT1}
  Constantes;

Function LanceImport (var RequestTOB : TOB; sttArg : string; AffInfo : TInfoEvent; var PImport : TImportCom) : Boolean;
var
Nr   : string;
{$IFNDEF EAGLSERVER}
Q    : TQuery;
TImp : TOB;
{$ENDIF}
begin
  PImport := TImportCom.Create(RequestTOB, sttArg, AffInfo);
  With PImport do
  begin
       RemplirTOBDevise;
       if OkExercice then
       begin
{$IFDEF EAGLSERVER}
        if UpperCase(ExtractFileExt(FichierEntree)) = '.ZIP' then
           LanceScruteRepertoire(RequestTOB, sttArg)
        else
{$ENDIF}
           ImportComSx(RequestTOB, sttArg);
       end;
       OnAfficheListeCom('*** Fin d''import :'  + FormatDateTime(Traduitdateformat('dd/mm/yyyy à hh:nn'),NowH),LISTEIMPORT);
       if (not OKexercice) or (not OkCreatCompte) then
       begin
            if not RequestTOB.FieldExists ('ERROR') then
              RequestTOB.AddChampSupValeur ('ERROR', '');
            if (not OkCreatCompte) then RequestTOB.PutValue('ERROR', ERR_CREATE_CPTE);
            if (not OKexercice) then
               RequestTOB.PutValue('ERROR', 'ERREUR');
            OnAfficheListeCom(RequestTOB.GetValue('ERROR'),LISTEIMPORT);
            Result := FALSE;   Nr := 'NRE';
       end
       else
       begin
            Result := TRUE; Nr := 'ROK';
       end;
{$IFNDEF EAGLSERVER}
       if (sttArg <> '') and (ctxPCL in V_PGI.PGIContexte)  then // pour TrfS5
       begin
            Q := OpenSQl ('SELECT * FROM TRFFICHIER Where TRF_FICHIER Like "%' +ExtractFileName(PImport.FichierEntree)+ '%" AND '+
            'TRF_REPERTOIRE="'+ ExtractFileDir(PImport.FichierEntree) +'\"', TRUE);
            if not Q.EOF then
            begin
                 TImp :=TOB.Create('',Nil,-1) ;
                 TImp.LoadDetailDB('TRFFICHIER', '', '', Q, TRUE, FALSE);
                 TImp.detail[0].PutValue ('TRF_ETATREST', Nr);
                 TImp.InsertOrUpdateDB(True);
                 TImp.free;
            end;
            Ferme(Q);
       end;
{$ELSE}
       CloseFile(FicIE);
       if not BRejet then
       LanceRemoveDirectory (RequestTOB.GetValue('APPLICATION') );
{$ENDIF}
       if assigned(TDevise) then
          FreeAndNil(TDevise);
       free;
  end;
end;

{$IFDEF EAGLSERVER}

function TImportCom.LanceScruteRepertoire(var RTOB : TOB; sttArg : string) : Boolean;
var SearchRec   : TSearchRec;
    strDir, Ext : string;
begin
     Result := False;
     strDir := RootPath;
     if not DirectoryExists(strDir) then exit;
     LanceDeZippfile (ExtractFileName(FichierEntree));
     FindFirst(strDir+'\*.*', faAnyFile, SearchRec);
     repeat
       Ext := Uppercase(ExtractFileExt(SearchRec.Name));
       if (Ext <> '.TRA') and (Ext <> '.IMM') and (Ext <> '.TOB')
       and (Ext <> '.CPTA') then continue;
       if (SearchRec.Name<>'') then
       begin
           strDir := RootPath +  '\'+ SearchRec.Name;
           if ((SearchRec.Attr and faDirectory) <= 0) then
            begin
                 if (SearchRec.Name<>'.') and (SearchRec.Name<>'..') then
                 begin
                      RTOB.Detail.Items[0].putValue('FichierSortie', SearchRec.Name);
                      FichierEntree       := strDir;
                      if Ext = '.IMM' then
                         OrigineFichier := 'IMMO'
                      else
                         OrigineFichier := RTOB.Detail.Items[0].GetValue ('OrigineFichier');
                      if Ext = '.TOB' then
                         LanceMainSauveDossier (nil, RootPath, Nodossier, 'R')
                      else
                         ImportComSx(RTOB, sttArg);

                 end;
            end
      end;
  until FindNext(SearchRec)<>0 ;
  sysutils.FindClose(SearchRec);
  Result := True;
end;


procedure TImportCom.LanceDeZippfile (Filezip : string);
var
zip                       : TOZ;
RepertoireRoot            : string;
begin
      RepertoireRoot := RootPath ;
      ForceDirectories(RepertoireRoot);
      // ouverture du zip
      zip := TOZ.Create;
      zip.OpenZipFile(RepertoireRoot+'\'+Filezip, moOpen);
      if zip.OpenSession(osExtAll) then
      begin
        if zip.SetDirOut(RepertoireRoot) then
          begin
           zip.UnZipFileWithPath := FALSE; // pour extraire dans le reprtoire même
           zip.CloseSession;
           DeleteFile(RepertoireRoot+'\'+Filezip);
          end
        // ou incorrecte !
        else
          begin
          zip.CancelSession;
          zip.Free;
          exit;
          end;
      end;
      zip.Free;
end;

{$ENDIF}


Procedure TImportCom.AffMessage(Mess : string);
begin
{$IFNDEF EAGLSERVER}
        PgiBox(Mess+'#10','IMPORT') ;
{$ELSE}
        //ddWriteLN(Mess);
        cWA.MessagesAuClient('COMSX.IMPORT','',Mess) ;
{$ENDIF}
end;

constructor TImportCom.Create  (TE : Tob; sttArg : string; AffInfo    : TInfoEvent);
var
L1                      : TOB;
Exerc                   : TExoDate;
Q                       : TQuery;
{$IFNDEF EAGLSERVER}
Source                  : TFileStream;
{$ENDIF}
begin
     OnAfficheListeCom := Affinfo;
{$IFDEF EAGLSERVER}
     TCPContexte.GetCurrent.InfoCpta.ChargeLgDossier ;
     TCPContexte.GetCurrent.Exercice.ChargeMagExo ;
     RootPath := RenseigneWrootPath(TE.GetValue('APPLICATION'));
     if not DirectoryExists(RootPath) then LanceCreatDir(TE.GetValue('APPLICATION'));

     if not assigned(OnAfficheListeCom) then    // appel par process serveur
     begin
       OnAfficheListeCom := AfficheComImport;
       FichierResultat :=  RootPath + '\'+'ListeCom'+FormatDateTime(Traduitdateformat('yyyymmddhhnn'),NowH)+'.txt';
       AssignFile(FicIE, FichierResultat) ;
       Rewrite(FicIE) ;
     end;
{$ENDIF}
     stArg := sttArg;
     OnAfficheListeCom('*** Début d''import :' + FormatDateTime(Traduitdateformat('dd/mm/yyyy à hh:nn'),NowH),LISTEIMPORT);

     InitDossierEnNombre(GrandNbGene, GrandNbAux); // AJOUT ME 14-01-2005
     if (OrigineFichier = 'S5') and (NatureFichier <> 'BAL') then
{$IFDEF EAGLSERVER}
        GerBAP := (not (ctxPCL in V_PGI.PGIContexte)) and (ExisteTypeVisa (''))
{$ELSE}
        GerBAP := (not (V_PGI.ModePCL='1')) and (ExisteTypeVisa (''))
{$ENDIF}
     else
        GerBAP := FALSE;

     NoDossier           := TE.Getvalue ('DOSSIER');
     if NoDossier = '' then NoDossier := '000000';
     L1                  := TE.Detail.Items[0];
     FichierEntree       := L1.getValue('FichierSortie') ;

// A PORTER SUR V800
     TailleSource := 0;
{$IFNDEF EAGLSERVER}
     Source := TFileStream.Create(FichierEntree, fmOpenRead or fmShareDenyNone);
     TailleSource := Source.size;
     Source.Free;
{$ENDIF}
// fin A PORTER SUR V800
     NbLec := 0; TailleBloc := 0;

     AutoriseEcheck      := L1.getValue('AutoriseEcheck');
     // Interdire La Creation IGE=Généraux IAU=auxiliaire IJA=Journal IET=Etablissement
     InterdireLaCreation := L1.GetValue ('INTERDIRECREAT');
{$IFNDEF EAGLSERVER}
     RootPath            := ExtractFileDir (FichierEntree);
{$ENDIF}
     RepriseComplement   := L1.getValue('RepriseComplement');
     Paramgene           := L1.getValue ('Paramgene');
     CStat               := L1.GetValue ('CPSTAT');
     RupturePiece        := L1.GetValue ('RupturePiece');
     CalculPiece         := L1.GetValue ('CalculPiece');
     Gestiondoublon      := L1.GetValue ('Gestiondoublon');
     BMAJCPTE            := L1.Getvalue ('BMAJCPTE');
     TypeEcrR            := L1.GetValue ('TypeEcrR');
     TypeEcrI            := L1.GetValue ('TypeEcrI');
     QualifOrigine       := L1.GetValue ('QualifOrigine');
     Sansaux             := L1.GetValue ('Sansaux');
     FCorresp            := L1.GetValue ('FCorresp');
     OrigineEnvoie       := L1.GetValue ('OrigineEnvoie');
     fbGeneLg            := GetInfoCpta (fbGene).lg ;
     fbAuxLg             := GetInfoCpta (fbAux).lg;
     fbGeneCb            := GetInfoCpta (fbGene).Cb;
     fbAuxCb             := GetInfoCpta (fbAux).Cb;
     fbAxe1Cb            := GetInfoCpta (fbAxe1).Cb;
     fbAxe2Cb            := GetInfoCpta (fbAxe2).Cb;
     fbAxe3Cb            := GetInfoCpta (fbAxe3).Cb;
     fbAxe4Cb            := GetInfoCpta (fbAxe4).Cb;
     fbAxe5Cb            := GetInfoCpta (fbAxe5).Cb;
     MontantNegatif      := GetParamSocSecur('SO_MONTANTNEGATIF',False) ;
     TenueEuro           := GetParamSocSecur ('SO_TENUEEURO', 'X');
     AttribRIBAuto       := GetParamSocSecur('SO_ATTRIBRIBAUTO',False);
     OrigineFichier      := L1.GetValue ('OrigineFichier');
     DecoupLon           := L1.GetValue ('DecoupLon');
     BIntegre            := L1.GetValue ('BIntegre');
     BControlEtab        := L1.GetValue ('BControlEtab');
     BInterdireCreationTiers      := L1.Getvalue ('CREATIONTIERS');
     BRejet              := L1.GetValue ('BRejet');
     BEquilibre          := L1.GetValue ('BDESEQUILIBRE');

     BValider            := L1.GetValue ('BValider');
     BEcartChange        := L1.GetValue ('EcartChange');
     OptionTP            := L1.GetValue ('BTPayeur');
    // Mode avancée compte de remplacement
    CpteGene        := L1.GetValue ('CPTEGENE');
    CpteClient      := L1.GetValue ('CPTECLIENT');
    CpteFour        := L1.GetValue ('CPTEFOUR');
    CpteSalaire     := L1.GetValue ('CPTESALAIRE');
    CpteDivers      := L1.GetValue ('CPTEDIVERS');
    RAxe1           := L1.GetValue ('RAXE1');
    RAxe2           := L1.GetValue ('RAXE2');
    RAxe3           := L1.GetValue ('RAXE3');
    RAxe4           := L1.GetValue ('RAXE4');
    RAxe5           := L1.GetValue ('RAXE5');
    SubCollCli      := L1.GetValue ('SUBCOLLCLI');
    SubCollFou      := L1.GetValue ('SUBCOLLFOU');
    OptionEchange   := L1.GetValue ('OptionEchange');
    BSoldeLigne     := L1.GetValue ('BSoldeLigne');

    BBlanc          := L1.GetValue ('BBLANC');
    ModifTiers      := L1.GetValue ('MODIFTIERS');
    MajDate         := L1.GetValue ('MajDateDernMvt');
    // fiche 10558
    ControlExercClos := L1.GetValue ('ControlExercClos');
    BCalculModeregle := L1.GetValue ('CalculdateEche');

    ScanBanck        := L1.GetValue ('ScanBanck');

    // Import Ventilation Type BVE 25.07.07
    GereVentilType  := ExisteSQL('SELECT ##TOP 1##RB_CLASSERUB  FROM RUBRIQUE WHERE RB_CLASSERUB = "TRA"');
    ModeEVentilType := FALSE;

    OkExercice := TRUE; ExisteTP := FALSE;
    if stArg <> ''  then  // en ligne de commande vérification de l'existance des comptes
    begin
           if (CpteGene <> '') and (not ExisteSQL ('SELECT * FROM GENERAUX Where G_GENERAL="'+BourreOuTronque(CpteGene,FbGene)+'"')) then
           begin
              OnAfficheListeCom('Compte de remplaçement ' + CpteGene + ' inexistant, veuillez corriger le fichier INI' ,LISTEIMPORT); OKexercice := FALSE;
           end;
           if (SubCollCli <> '') and (not ExisteSQL ('SELECT * FROM GENERAUX Where G_GENERAL="'+BourreOuTronque(SubCollCli, FbGene)+'"')) then
           begin
              OnAfficheListeCom('Compte de substitution ' + SubCollCli + ' inexistant, veuillez corriger le fichier INI' ,LISTEIMPORT); OKexercice := FALSE;
           end;
           if (SubCollFou <> '') and (not ExisteSQL ('SELECT * FROM GENERAUX Where G_GENERAL="'+BourreOuTronque(SubCollFou, FbGene)+'"')) then
           begin
              OnAfficheListeCom('Compte de substitution ' + SubCollFou + ' inexistant, veuillez corriger le fichier INI' ,LISTEIMPORT);  OKexercice := FALSE;
           end;
           if (CpteFour <> '') and (not ExisteSQL (' Select T_AUXILIAIRE from TIERS Where T_AUXILIAIRE="'+BourreOuTronque(CpteFour, FbAux)+'"')) then
           begin
              OnAfficheListeCom('Compte de remplaçement ' + CpteFour + ' inexistant, veuillez corriger le fichier INI' ,LISTEIMPORT); OKexercice := FALSE;
           end;
           if (CpteClient <> '') and (not ExisteSQL (' Select T_AUXILIAIRE from TIERS Where T_AUXILIAIRE="'+BourreOuTronque(CpteClient, FbAux)+'"')) then
           begin
              OnAfficheListeCom('Compte de remplaçement ' + CpteClient + ' inexistant, veuillez corriger le fichier INI' ,LISTEIMPORT); OKexercice := FALSE;
           end;
           if (CpteSalaire <> '') and (not ExisteSQL (' Select T_AUXILIAIRE from TIERS Where T_AUXILIAIRE="'+BourreOuTronque(CpteSalaire, FbAux)+'"')) then
           begin
              OnAfficheListeCom('Compte de remplaçement ' + CpteSalaire + ' inexistant, veuillez corriger le fichier INI' ,LISTEIMPORT); OKexercice := FALSE;
           end;
           if (CpteDivers <> '') and (not ExisteSQL (' Select T_AUXILIAIRE from TIERS Where T_AUXILIAIRE="'+BourreOuTronque(CpteDivers, FbAux)+'"')) then
           begin
              OnAfficheListeCom('Compte de remplaçement ' + CpteDivers + ' inexistant, veuillez corriger le fichier INI' ,LISTEIMPORT); OKexercice := FALSE;
           end;
    end;
    Mrdefaut        := L1.GetValue ('MRDEFAUT');
    Regdefaut       := L1.GetValue ('REGDEFAUT');
    BCtrP           := L1.GetValue ('BCtrP');
    BCtrS           := L1.GetValue ('BCtrS');
    BRecalculS      := L1.GetValue ('BRecalculS');

    TypeVisa        := L1.GetValue ('TYPEVISA');

{$IFDEF EAGLSERVER}
     if ((L1.GetValue('ModePCL')='1')) then V_PGI.PGIContexte:=V_PGI.PGIContexte+[ctxPCL] ;
{$ENDIF}

     OkCreatCompte       := TRUE;
     MAJCPEXOREF         := FALSE;
     PresenceExercice    := FALSE;
     TiersAutre          := FALSE;
     OkModePaieDefaut    := FALSE;
     OkMajExoV8          := FALSE;
     OkVerAxe            := FALSE;
     TotalMouvDebit := 0.0; TotalMouvCredit := 0.0; NbEcrAna := 0;
     ModeCroisaxe    := FALSE;
     ChargeChoixCod  := FALSE;
     UpdateLettreDiv := FALSE;
     STiers := nil; SGeneraux := nil; SJournal := nil; SSection := nil; SSectionInbase := nil;
     STLibre := nil; SSectionFerme := nil; SPays := nil;
     vInfo:= nil; ZDevise := nil; SModePaie := nil; STva := nil;

     AppliquerRest  := ExisteSQl('SELECT CLA_GENERAL  FROM CLIENGENEMODELA');
     DateEcrChange := iDate1900;  GarderJour := '';
     if (L1.GetValue ('DATEECR') <> '') and (L1.GetValue ('DATEECR') <> Stdate1900) then
     begin
          if Copy(L1.GetValue ('DATEECR'), 0, 2) <> '99' then
          begin
              DateEcrChange  := StrToDate(L1.GetValue ('DATEECR'));
              if not CQuelExercice (DateEcrChange, Exerc) then
              begin
                   OnAfficheListeCom(ERR_DATEECRITURE + ': '+FormatDateTime(Traduitdateformat('mm/dd/yyyy'),DateEcrChange),LISTEIMPORT);
                   OKexercice := FALSE; exit;
              end;
          end
          else GarderJour := L1.GetValue ('DATEECR');
     end;
     YCotation := 0;
///////////////////////// MAJ des soldes
     if BSoldeLigne then
     begin
          lpTi := TList.Create;
          lpG  := TList.Create;
          lpJ  := TList.Create;
          lpS  := TList.Create;
     end;
     // mise à jour g_datedernmvt
     if MajDate Then
          lpGM  := TList.Create;

     AxeTva := '';  // fiche 10568
     NbSection := 0;
     if GetParamSocSecur('SO_CPPCLAXETVA', '') <> '' then
        AxeTva := GetParamSocSecur('SO_CPPCLAXETVA', '');

  ExerciceSuperVise := ''; //Fiche 10587
  try
  Q := OpenSQL('SELECT CIR_EXERCICE FROM CREVINFODOSSIER WHERE ' +
               'CIR_NODOSSIER = "' + V_Pgi.NoDossier + '" AND ' +
               'CIR_EXERCICE = "' + GetEnCours.Code + '" AND ' +
               'CIR_SUPERVISE="X" ' , True );
  if not Q.Eof then
    ExerciceSuperVise := Q.FindField('CIR_EXERCICE').AsString;
  Ferme(Q);
  except;
  end;
end;

Function  TImportCom.TransfertDeSISCO : Boolean;
var
Fichier      : string;
begin
        Result := TRUE;
        if OrigineFichier = 'SI' then
        begin
          system.New(PInfoImp) ; FillChar(PInfoImp^,SizeOf(PInfoImp^),#0) ;
          PInfoImp.LCS       := HtStringList.Create ;
          PInfoImp.LAnaLu    := HtStringList.Create ;
          PInfoImp.TOBStat   := TOB.Create ('', nil, -1);
          Fichier := TransfertSiscoVersPgi(FichierEntree, FALSE, Sansaux,'', PInfoImp, stArg);
          if PInfoImp.PbEnr[0] then Result := FALSE;
          if NatureFichier = 'EXE' then
          begin
               NatureFichier := 'DOS';
               Paramgene := TRUE;
          end;
          if Fichier = '' then OKexercice := FALSE;
{$IFDEF EAGLSERVER}
          DeleteFile (FichierEntree);
{$ENDIF}
          FichierEntree := Fichier;
        end;
        TModeRegle := nil;
end;

function TImportCom.ImportComSx (TImp : TOB; SArg : string) : Boolean;
var
  FichierIE    : TextFile;
  SLect        : string;
  CodeEnreg    : string;
  i,ic,j       : integer;
  monn,Tmp     : string;
  inf          : integer;
  D1,D2        : TDatetime;
  TA           : TOB;
  OKExerc,ANNUL: Boolean;
  {$IFDEF EAGLSERVER}
  Source       : string;
  SListFile    : TStringList;
  TDataStream  : TMemoryStream;
  {$ELSE}
  Source       : TFileStream;
  {$ENDIF}
  FichierRejet : string;
  Gamme        : string;
  WhereJ       : string;
  Tail,TailSav : double;
begin

  OkExercice := TRUE;  ANNUL := FALSE; Tail := 0;
{$IFDEF EAGLSERVER}
   SListFile := nil;
{$ENDIF}
  stArg := SArg;
  try
    BeginTrans ;
    ChargeToDo;
    VersionFichier := versionexport;
    OKExerc := TRUE;
    if (OrigineFichier = 'IMMO') then
    begin
         Result := ImportBobIMMO;
         CommitTrans ;
         exit;
    end;
{$IFDEF COMSX}
{$IFNDEF EAGLSERVER}
    // ajout me pour le fichier nsv
    if (OrigineFichier = 'PAIE') then
    begin
         Result := PgImportDef(FichierEntree,'AUTO',';') ;
         if Result then
         begin
            Result := ControleFichierImport(nil);
            if Result then CommitTrans
            else OKexercice := FALSE;
         end;
         exit;
    end;
{$ENDIF}
{$ENDIF}

{$IFDEF EAGLSERVER}
   if TImp.FieldExists ('SOURCE') then
   begin
          Source         := TImp.GetValue('SOURCE') ;
          SListFile      := TStringList.create;

          TDataStream := TMemoryStream.Create;
          TDataStream.Write(PChar(Source)^, Length(Source));
          TDataStream.Seek(0, 0);
          TailleSource := TDataStream.size;  // A PORTER SUR V800

      // Pour format SISCOII
          FichierEntree := RootPath + '\'+ ExtractFileName(FichierEntree);

         //SListFile.LoadFromStream (TDataStream);

          TDataStream.SaveToFile (RootPath + '\'+ ExtractFileName(FichierEntree));
          TDataStream.free;
   end;
   if (OrigineFichier = 'SI') and (not TransfertDeSISCO) then begin OKexercice := FALSE; Result := FALSE; exit; end;
    if BRejet then
    begin
           tmp := FichierEntree;
           FichierRejet := ReadTokenPipe (tmp, '.') +'.REJET';
           if FileExists (FichierRejet) then DeleteFile (FichierRejet);
           AssignFile(FRejet, FichierRejet);
           try Rewrite(FRejet) ; except end;
           Writeln(FRejet, '!');
    end;


(*    for  ic := 0 to SListFile.count-1 do
    begin
      SLect := SListFile.Strings[ic];
*)
    AssignFile(FichierIE, FichierEntree);
    Reset(FichierIE);
    if EOF(FichierIE) then
    begin Result := FALSE; exit; end;
    while not EOF(FichierIE) do
    begin
      Readln(FichierIE, SLect);

{$ELSE}
    if not TransfertDeSISCO then
    begin
      Result := FALSE;  OkExercice := FALSE;
      RollBackDiscret;
      VideListeInfoImp(PInfoImp^, TRUE);
      Dispose(PInfoImp) ;
      exit;
    end;
    if BRejet then
    begin
           tmp := FichierEntree;
           FichierRejet := ReadTokenPipe (tmp, '.') +'.REJET';
           if FileExists (FichierRejet) then DeleteFile (FichierRejet);
           AssignFile(FRejet, FichierRejet);
           try Rewrite(FRejet) ; except end;
           Writeln(FRejet, '!');
    end;
{$IFNDEF EAGLSERVER}
    if OrigineFichier = 'SI' then
    begin
     TailleSource := 0;
     Source := TFileStream.Create(FichierEntree, fmOpenRead or fmShareDenyNone);
     TailleSource := Source.size;
     Source.Free;
    end;
{$ENDIF}
    AssignFile(FichierIE, FichierEntree);
    {$I-}  Reset(FichierIE);  {$I+}
    if EOF(FichierIE) then
    begin Result := FALSE; exit; end;
    while not EOF(FichierIE) do
    begin
      Readln(FichierIE, SLect);
{$ENDIF}

      if (not OKexercice) or (not OKExerc) then
      begin
           OnAfficheListeCom('Import Annulé' ,LISTEIMPORT);
           break;
      end;
      Application.ProcessMessages;
      sleep(1);  // pour libérer le temps cpu

      Gamme := Copy (SLect,0, 5);
      if (Gamme = '***S1') or (Gamme = '***S3') or
         (Gamme = '***S2') or (Gamme = '***S5') or
         (Gamme = '***S7') or (Gamme = '***SI') or
         (Gamme = '***QU') or (Gamme = '***WT') or (Gamme = '***WS') then
      begin
           OrigineFichier := Copy (SLect,4, 2);
           NatureFichier := Copy (SLect,9, 3);
           VersionFichier := Copy (SLect,34, 3);
           OnAfficheListeCom('Origine du fichier : ' + Copy (SLect,4, 2),LISTEIMPORT);
           OnAfficheListeCom('Fichier généré par : ' + Copy (SLect,54, 35),LISTEIMPORT);
      end;
      if (Copy (SLect,0, 1) = '!') or (NatureFichier='') then // ajout me pour Fichier CCIMPEX
      begin
               NatureFichier  := 'JRL';
               OrigineFichier := 'S5';
               QualifOrigine  := 'ND0' ;
               Paramgene := FALSE;
               Readln(FichierIE, SLect);
       end;
      TailleBloc := TailleBloc + length (SLect);
      if (TailleSource <> 0) and (NbLec > 10) then
      begin
          TailSav := Tail;
          Tail := (TailleBloc / TailleSource);
          Tail := arrondi(Tail, 2);
          if Tail <> TailSav then
          OnAfficheListeCom('', LISTEIMPORT, (Tail * 100));
          NbLec := 0;
      end
      else inc(NbLec);
      if EstUneLigneCpt(SLect) then
      begin
        CodeEnreg := Copy(SLect, 4, 3);
        for inf := low(ListeRecep) to high(ListeRecep) do
          if (ListeRecep[inf].Code = CodeEnreg) or (ListeRecep[inf].Code =
            Copy(SLect, 4, 2)) then
            ListeRecep[inf].ToDo(SLect);
      end
      else
      begin
          if (SLect <> '') then
          begin
            if Copy(SLect, 31, 1) = 'B' then
               ImporteEcritureBud(SLect) // import des écritures budgétaires
            else
            begin
               OkRejet := FALSE;
               ImporteEcriture(SLect); // import des écritures
               if BRejet and OkRejet then
               begin
                      // if not FileExists (FichierRejet) then Rewrite(FRejet) else Append(FRejet);
                       Writeln(FRejet, SLect);
                       Flush(FRejet);
                       if (not BIntegre) then Annul := TRUE;
               end;
            end;
          end;
      end;
    end;
{$IFDEF EAGLSERVER}
        SListFile.free;
{$ENDIF}
        CloseFile(FichierIE);
    // création d'un établissement si pas d'établissement dans le fichier
    OnAfficheListeCom('Vérification des établissements ',LISTEIMPORT);
    if ListeEtab <> nil then
    begin
           // fiche 10296
        if BControlEtab or ((not AutoriseEcheck) and (pos('IET', InterdireLaCreation) <> 0)) then
        begin
                For i:=0 To ListeEtab.Count-1 Do
                begin
                     if not ExisteSql ('SELECT ET_ETABLISSEMENT from ETABLISS WHERE ET_ETABLISSEMENT="'+ListeEtab[i]+'"') then
                     begin
                        OnAfficheListeCom('Etablissement '+ ListeEtab[i] +' est manquant',LISTEIMPORT);
                        OkCreatCompte := FALSE
                     end;
                end;
        end
        else
        begin
              if not ExisteSql ('SELECT ET_ETABLISSEMENT from ETABLISS') then
                 ImporteEtablissement('***ETB'+GetParamSocSecur ('SO_ETABLISDEFAUT', '001')+GetParamSocSecur ('SO_LIBELLE', ''));
              if ListeEtab <> nil then
              begin
                For i:=0 To ListeEtab.Count-1 Do
                begin
                     if not ExisteSql ('SELECT ET_ETABLISSEMENT from ETABLISS WHERE ET_ETABLISSEMENT="'+ListeEtab[i]+'"') then
                     begin
                        ExecuteSQL('Insert Into ETABLISS (ET_ETABLISSEMENT,ET_LIBELLE) ' +
                          'Values("' + ListeEtab[i] + '","Etablissement ' + ListeEtab[i]+ '")');
                     end;
                end;
              end;
       end;
       ListeEtab.free; ListeEtab := nil;
    end;

    if (not OKexercice) or (not OkCreatCompte) or (Annul) then
    begin
         Result := FALSE;
         RollBackDiscret;
         LibereTob;
         LibereHTable;
         if TTiers <> nil then begin TTiers.Free;  TTiers := nil; end;
         if TTiersInsert <> nil then begin TTiersInsert.Free; TTiersInsert := nil; end;
         if TTiersupdate <> nil then begin TTiersupdate.Free;  TTiersupdate := nil; end;
         if TGen <> nil then begin TGen.Free; TGen := nil; end;
         if TGenInsert <> nil then begin TGenInsert.Free; TGenInsert := nil; end;
         if TGenupdate <> nil then begin TGenupdate.Free;  TGenupdate := nil; end;
         if TBUDEcr <> nil then begin TBUDEcr.Free;  TBUDEcr := nil; end;
         if TTiersCAU <> nil then begin TTiersCAU.free; TTiersCAU := nil; end;
         if BRejet  then  CloseFile(FRejet);
         exit;
    end;

    //TEcrParJournal.SaveToFile('journal.txt',True,True,True);

    if TEcrParJournal <> nil then
    begin
           OnAfficheListeCom('Enregistrement des écritures',LISTEIMPORT);
           Genere_ecratEuro;
           TEcrParJournal.InsertDBByNivel(FALSE);
           if (not Paramgene) and ((NatureFichier = 'SYN')  or (OrigineFichier = 'S1')) then
           begin
                     if TPieceCon <> nil then
                          DelettrageSynchro (nil);
           end;
         // ANouveau dynamique
         if GetParamSocSecur('SO_CPANODYNA',false) then
         begin
              OnAfficheListeCom('Enregistement des a-nouveaux dynamiques',LISTEIMPORT);
              if not ExecReqMAJAno(TSAnv) then
               OnAfficheListeCom('Erreur enregistrement des a-nouveaux dynamiques',LISTEIMPORT);

              for i:=0 to TSAnv.Count-1 do
               if assigned(TSAnv[i]) then Dispose(TSAnv[i]);
              TSAnv.Free ;

         end;
         // Visa
         if assigned(EnregBap) then
         begin
              // fiche 10416
              if  (EnregBap.TobMasse <> nil) and (EnregBap.TobMasse.detail.count > 0) then
              begin
                   OnAfficheListeCom('Nombre de bons à payer : '+ IntToStr(EnregBap.TobMasse.detail.count), LISTEIMPORT);
                   EnregBap.TobMasse.InsertOrUpdateDB(TRUE);
              end;
         end;
    end
    else
    begin
          if OptionEchange then // Fiche 16078
           OnAfficheListeCom('Aucune écriture comptable n''est importée',LISTEIMPORT);
    end;
    { Fermeture du fichier de rejet après l'enregistrement des écriture}
    if BRejet  then  CloseFile(FRejet);

    // fiche 10396
    if ((NatureFichier = 'DOS') or ((OrigineFichier = 'S1') and (NatureFichier = 'SYN'))) and (TOBLettre <> nil) then
    begin
{$IFNDEF EAGLSERVER}
         if TOBLettre.detail.count > 0 then
            InitMove(TOBLettre.detail.count-1,'Lettrage');
{$ENDIF}
        for ic := 0 to TOBLettre.detail.count-1 do
        begin
{$IFNDEF EAGLSERVER}
              movecur(false);
{$ENDIF}
              D1 := StrToDate( TOBLettre.detail[ic].getvalue('DATEPAQUETMAX'));
              D2 := StrToDate( TOBLettre.detail[ic].getvalue('DATEPAQUETMIN'));
              ExecuteSQL('UPDATE ECRITURE SET E_DATEPAQUETMAX="' + UsDateTime(D1) +
                '",E_DATEPAQUETMIN="'+UsDateTime(D2)+'" ' +
                ' WHERE E_GENERAL="'+ TOBLettre.detail[ic].getvalue('COMPTEGENE') + '" AND E_AUXILIAIRE="' + TOBLettre.detail[ic].getvalue('COMPTEAUX') + '" and E_LETTRAGE="'+TOBLettre.detail[ic].getvalue('LETTRAGE')
                + '" ' //and E_EXERCICE="'+TOBLettre.detail[ic].getvalue('EXERCICE')+'"'
                + ' and E_ETATLETTRAGE="'+TOBLettre.detail[ic].getvalue('ETATLETTRAGE')+'"'
                );
        end;

{$IFNDEF EAGLSERVER}
         FiniMove;
{$ENDIF}
    end;
    if (OrigineFichier = 'S1') and (TOBJrl <> nil)then
    begin
         for i := 0 to TOBJrl.detail.Count - 1 do
         begin
              TA := TOBJrl.detail[i];
              if TA.GetValue ('OK') then
              begin
                  ExecuteSQL('UPDATE ECRITURE SET E_MODESAISIE="LIB"' +
                    ' WHERE E_JOURNAL="' + TA.GetValue ('JOURNAL')+'"');
                  ExecuteSQL('UPDATE JOURNAL SET J_MODESAISIE="LIB"' +
                    ' WHERE J_JOURNAL="' + TA.GetValue ('JOURNAL')+'"');
              end;
         end;
    end;
    if (OrigineFichier = 'S1') and (GetParamSocSecur('SO_CPLIENGAMME','') = '') and (ctxPCL in V_PGI.PGIContexte) then
        SetParamsoc ('SO_CPLIENGAMME', 'S1');
    if (OrigineFichier = 'S3') and (GetParamSocSecur('SO_CPLIENGAMME','') = '')  and (ctxPCL in V_PGI.PGIContexte) then
            SetParamsoc ('SO_CPLIENGAMME', 'S3');
     OnAfficheListeCom('Enregistement des comptes tiers',LISTEIMPORT);
     InsertTiers;
     OnAfficheListeCom('Enregistement des comptes généraux',LISTEIMPORT);
     InsertUpdateGeneraux;


    if TEcrParJournal <> nil then
    begin
      OnAfficheListeCom('Nombre d''écritures : ' + IntToStr(NbEcriture),LISTEIMPORT);
      if NbEcrAna  > 0 then
         OnAfficheListeCom('Nombre d''écritures analytique : ' + IntToStr(NbEcrAna),LISTEIMPORT);

      MiseajourSouche;
    end;

    // ajout me 02-02-2005
     if TBUDEcr <> nil then
     begin
           TBUDEcr.InsertOrUpdateDB(TRUE);
           TBUDEcr.Free;  TBUDEcr := nil;
     end;
     if BRecalculS then
     begin
          if not BSoldeLigne then
          begin
               OnAfficheListeCom('Mise à jour des soldes ',LISTEIMPORT);
               if (OrigineFichier = 'SI') or (NatureFichier = 'DOS') then // fiche 10234
                  MajTotTousComptes(False, '')
               else
               begin
                  //MajTotComptes([fbsect,fbJal], FALSE, FALSE , '');
                  OnAfficheListeCom('Mise à jour des soldes Généraux',LISTEIMPORT);
                  if SGeneraux<> Nil then
                  begin
                       for i := 0 to SGeneraux.count-1 do
                       begin
                            For j := 0 to SExercice.count-1 do
                                MajTotComptes([fbGene], FALSE, FALSE , SExercice[j], SGeneraux[i], FALSE);
                       end;
                       for i:=0 to SGeneraux.Count-1 do If SGeneraux.Objects[i]<>NIL Then SGeneraux.Objects[i].Free ;
                       SGeneraux.free;  SGeneraux := nil;
                  end;
                  OnAfficheListeCom('Mise à jour des soldes Tiers',LISTEIMPORT);
                  if STiers <> nil then
                  begin
                        for i := 0 to STiers.Count-1 do
                        begin
                            For j := 0 to SExercice.count-1 do
                                MajTotComptes([fbaux], FALSE, FALSE , SExercice[j], STiers[i], FALSE);
                        end;
                        for i:=0 to STiers.Count-1 do If STiers.Objects[i]<>NIL Then STiers.Objects[i].Free ;
                        STiers.free;  STiers := nil;
                   end;
                  OnAfficheListeCom('Mise à jour des soldes Section',LISTEIMPORT);
                  if SSection <> nil then
                  begin
                        for i := 0 to SSection.Count-1 do
                        begin
                            For j := 0 to SExercice.count-1 do
                                MajTotComptes([fbsect], FALSE, FALSE , SExercice[j], SSection[i], FALSE);
                        end;
                        for i:=0 to SSection.Count-1 do If SSection.Objects[i]<>NIL Then SSection.Objects[i].Free ;
                       SSection.free;  SSection := nil;
                   end;
                  if  SJournal <> nil then
                  begin
                       for i := 0 to SJournal.Count-1 do
                       begin
                            For j := 0 to SExercice.count-1 do
                                MajTotComptes([fbJal], FALSE, FALSE , SExercice[j], SJournal[i], FALSE);
                       end;
(* fiche 10490                       if not GetParamSocSecur('SO_OUITVAENC',FALSE) then
                       begin
                            for i:=0 to SJournal.Count-1 do If SJournal.Objects[i]<>NIL Then SJournal.Objects[i].Free ;
                            SJournal.free;  SJournal := nil;
                       end;
*)
                  end;
                  if SExercice <> nil then
                  begin
                       for i:=0 to SExercice.Count-1 do If SExercice.Objects[i]<>NIL Then SExercice.Objects[i].Free ;
                       SExercice.free;  SExercice := nil;
                  end;
               end;
          end
          else      ///////////////////////// MAJ des soldes  à la ligne
          begin
                      CalcSoldeList (lpTi, FbAux);
                      CalcSoldeList (lpG, FbGene);
                      CalcSoldeList (lpJ, FbJal);
                      CalcSoldeList (lpS, FbSect);
                      VideListe (lpTi); lpTi.free;
                      VideListe (lpG); lpG.free;
                      VideListe (lpJ); lpJ.free;
                      VideListe (lpS); lpS.free;
          end;
          // fiche 10160
          if MajDate then
          begin
           MAJDateDernMvt;
           VideListe (lpGM); lpGM.free;
          end;
     end;

     CpStatutDossier;
     if SSectionInbase <> nil then
     begin
              for i:=0 to SSectionInbase.Count-1 do If SSectionInbase.Objects[i]<>NIL Then SSectionInbase.Objects[i].Free ;
             SSectionInbase.free;  SSectionInbase := nil;
     end;
     if SSectionFerme <> nil then
     begin
              for i:=0 to SSectionFerme.Count-1 do If SSectionFerme.Objects[i]<>NIL Then SSectionFerme.Objects[i].Free ;
             SSectionFerme.free;  SSectionFerme := nil;
     end;
     if STLibre <> nil then
     begin
              for i:=0 to STLibre.Count-1 do If STLibre.Objects[i]<>NIL Then STLibre.Objects[i].Free ;
             STLibre.free;  STLibre := nil;
     end;

     if STva <> nil then
     begin
         for i:=0 to STva.Count-1 do If STva.Objects[i]<>NIL Then STva.Objects[i].Free ;
         STva.free;  STva := nil;
     end;
     
     if SPays <> nil then
     begin
              for i:=0 to SPays.Count-1 do If SPays.Objects[i]<>NIL Then SPays.Objects[i].Free ;
             SPays.free;  SPays := nil;
     end;

     // A PORTER sur 7XX
     if GetParamSocSecur('SO_OUITVAENC',FALSE) and
     ((SJournal <> nil) or ((OrigineFichier <> 'SI') and (NatureFichier = 'DOS'))) then
     begin
         OnAfficheListeCom('Calcul tva sur encaissement',LISTEIMPORT);
         if (SJournal <> nil) then
         begin
           for i := 0 to SJournal.Count-1 do
           begin
            if WhereJ = '' then WhereJ := 'AND (J_JOURNAL="'+ SJournal[i] +'"'
            else WhereJ := WhereJ + 'OR J_JOURNAL="'+ SJournal[i] +'"';
           end;
         end else WhereJ := '';
         if WhereJ <> '' then WhereJ := WhereJ + ')';
         CCalcTvaSimplifiee('', ' E_DATECREATION="'+ USDATETIME(DATE)+'"', WhereJ);
     end;
     // gestion des tiers payeurs fiche 10490 deplacement
{$IFDEF COMSX}
     if OptionTP and ExisteTP and GetParamSocSecur('SO_OUITP', FALSE) then
     begin
          GenerePiecesPayeurCom ;
          WhereJ := '';
          if SJournal <> nil then
          begin
                for i := 0 to SJournal.Count-1 do
                begin
                  if WhereJ = '' then WhereJ := 'AND (E_JOURNAL="'+ SJournal[i] +'"'
                  else WhereJ := WhereJ + 'OR E_JOURNAL="'+ SJournal[i] +'"';
                end;
                if WhereJ <>'' then WhereJ := WhereJ +')';
          end;
          ExecuteSQL ('UPDATE ECRITURE SET E_FLAGECR="" '
          +'Where E_FLAGECR="TP" AND E_AUXILIAIRE<>""'
          +'AND ((E_NATUREPIECE<>"AC" AND E_NATUREPIECE<>"AF"))'+
          ' AND E_DATECREATION="'+ USDATETIME(DATE)+'"' + WhereJ);

     end;
{$ENDIF}
     if  SJournal <> nil then
     begin
         for i:=0 to SJournal.Count-1 do If SJournal.Objects[i]<>NIL Then SJournal.Objects[i].Free ;
          SJournal.free;  SJournal := nil;
     end;

    if (AxeTva <> '') and (NbSection <> 0) then   // fiche 10568
     OnAfficheListeCom(IntToStr(NbSection) + ' sections présentes dans le fichier n''ont pas été créée car elles concernent l''axe de TVA' ,LISTEIMPORT);

    OnAfficheListeCom('Mise à jour des paramètres sociétés ',LISTEIMPORT);
    if ExisteSQL('SELECT Y_EXERCICE FROM ANALYTIQ') then
    begin
         SetParamsoc ('SO_ZGEREANAL', TRUE);
         SetParamsoc ('SO_ZSAISIEANAL', TRUE);
    end;
    SetParamsoc ('SO_SUIVILOG', TRUE);
    SetParamsoc ('SO_CPCONVERTAUTO', TRUE);
    SetParamsoc ('SO_BOUCLERSAISIECREAT', TRUE);
    SetParamsoc ('SO_ZACTIVEPFU', TRUE);
    SetParamsoc ('SO_ZSAUVEFOLIOLOCAL', TRUE);
    SetParamsoc ('SO_ZFOLIOTEMPSREEL', TRUE);

    OnAfficheListeCom('Mise à jour des devises ',LISTEIMPORT);

    if (GetParamSocSecur('SO_DEVISEPRINC', 'EUR') = 'EUR') or (GetParamSocSecur('SO_TENUEEURO', TRUE) = FALSE) then
    begin
        ExecuteSQL('UPDATE DEVISE SET D_DEVISE="' + GetParamSocSecur('SO_DEVISEPRINC', 'EUR') +
          '",D_FERME="-"'+
          ' WHERE D_DEVISE="' + GetParamSocSecur('SO_DEVISEPRINC', 'EUR') + '"');
          monn := 'FRF';
    end
    else
    monn := GetParamSocSecur('SO_DEVISEPRINC', 'EUR');
    if ExisteSQl ('SELECT * from DEVISE Where D_DEVISE="'+monn+'"') then
    begin
        ExecuteSQL('UPDATE DEVISE SET D_DEVISE="' + monn +
          '",D_FONGIBLE="X",D_MONNAIEIN="X" ' +
          ' WHERE D_DEVISE="' + monn + '"');
    end;
(*XXXXXXX dans CPTESAV A REECRIRE    if (NatureFichier = 'DOS') then // ajout me
       RecalculTotPoint;
*)

   //--- Suppression des monnaie 'In'
    ExecuteSQL ('DELETE FROM DEVISE WHERE D_MONNAIEIN="X" AND D_FONGIBLE="-"');

//  finally
    SetParamsoc ('SO_CPRDDATERECEPTION', Date);

    MajExoV8;
    LibereTob;

    // création des comptes dans paramètres sociétés
    FiniRecupInfoDossier;
    // Fiche 19744
    if (OrigineFichier = 'SI') and (NatureFichier = 'DOS') then
    begin
          if  GetParamSocSecur('SO_REGIMEDEFAUT', False) = '' then
              SetParamsoc('SO_REGIMEDEFAUT', 'FRA');
          if GetParamSocSecur('SO_GCMODEREGLEDEFAUT', False) = '' then
              SetParamsoc('SO_GCMODEREGLEDEFAUT', '001');
          if GetParamSocSecur('SO_CODETVADEFAUT', False) = '' then
              SetParamsoc('SO_CODETVADEFAUT', 'TD');
          if GetParamSocSecur('SO_CODETVAGENEDEFAULT', False) = '' then
              SetParamsoc('SO_CODETVAGENEDEFAULT', 'NOR');
          if not GetParamSocSecur('SO_LETMVTOD', False) then
              SetParamsoc('SO_LETMVTOD', TRUE);
          if not GetParamSocSecur('SO_LETTOTAL', False) then
              SetParamsoc('SO_LETTOTAL', TRUE);
          if not GetParamSocSecur('SO_LETDC', False) then
              SetParamsoc('SO_LETDC', TRUE);
    end;

    // ajout me 19-08-2005 pour la création souche CPT par défaut
    if (NatureFichier = 'DOS') then
    begin
       // fiche 10529
       LoadStandardCompta (7, 'JOURNAL','JALREF', ' Where J_JOURNAL="ODR"', ' AND JR_JOURNAL="ODR"');
       VerifUneSouche('CPT',FALSE,FALSE) ;
    end;

    // Création de modepaiement s'il n'existe pas
    CreationModePaieInexistant;

    if (NatureFichier = 'SYN') then
       SetParamsoc ('SO_CPSYNCHROSX', FALSE);
    SetParamsoc ('SO_FLAGSYNCHRO', NatureFichier);

    CommitTrans ;


  if TotalMouvDebit <> 0 then
     OnAfficheListeCom(TraduireMemoire('Total mouvements débiteurs :'+ StrfMontant(TotalMouvDebit,20,V_PGI.OkDecV,'',False)),LISTEIMPORT);
  if TotalMouvCredit <> 0 then
     OnAfficheListeCom(TraduireMemoire('Total mouvements créditeurs :'+ StrfMontant(TotalMouvCredit,20,V_PGI.OkDecV,'',False)),LISTEIMPORT);

  // installation des coef imo2
   if Paramgene then
     InstalleLesCoefficientsDegressifs ;
  if (OrigineFichier = 'SI') and ((CStat = 'TL') or (CStat = 'ANA'))  then
  begin
            if TraiteCodeStatSISCO(CStat, PInfoImp^, True) then
               TraiteCodeStatSISCO(CStat, PInfoImp^, False) ;
  end;

  if OrigineFichier = 'SI' then
  begin
            VideListeInfoImp(PInfoImp^, TRUE);
            Dispose(PInfoImp) ;
  end;

    //OkFinImport := TRUE;
 EXCEPT

    On E: Exception do
      begin
      {$IFDEF EAGLSERVER}
        OnAfficheListeCom(E.Message,LISTEIMPORT);
      {$ENDIF}
        OnAfficheListeCom('Attention, aucune donnée n''a été importée',LISTEIMPORT);
        OnAfficheListeCom(SLect,LISTEIMPORT);
        LibereTob;
{$IFDEF EAGLSERVER}
        SListFile.free;
{$ENDIF}
         CloseFile(FichierIE);
         if BRejet  then  CloseFile(FRejet);
         LibereHTable;
         if TTiers <> nil then begin TTiers.Free;  TTiers := nil; end;
         if TTiersInsert <> nil then begin TTiersInsert.Free; TTiersInsert := nil; end;
         if TTiersupdate <> nil then begin TTiersupdate.Free;  TTiersupdate := nil; end;
         if TGen <> nil then begin TGen.Free; TGen := nil; end;
         if TGenInsert <> nil then begin TGenInsert.Free; TGenInsert := nil; end;
         if TGenupdate <> nil then begin TGenupdate.Free;  TGenupdate := nil; end;
         if TBUDEcr <> nil then begin TBUDEcr.Free;  TBUDEcr := nil; end;
         if TTiersCAU <> nil then begin TTiersCAU.free; TTiersCAU := nil; end;
         RollBackDiscret ;
      if OrigineFichier = 'SI' then
      begin
                VideListeInfoImp(PInfoImp^, TRUE);
                Dispose(PInfoImp) ;
      end;

      end ;
  end;

   Result := TRUE;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : M.ENTRESSANGLE
Créé le ...... : 30/04/2002
Modifié le ... :   /  /
Description .. : Procedure de remonter des paramètres généraux
Mots clefs ... :
*****************************************************************}
procedure TImportCom.ChargeToDo;
begin

  ListeRecep[0].Code := 'PS1';
  ListeRecep[0].ToDo := ImportParame1;
  ListeRecep[1].Code := 'PS2';
  ListeRecep[1].ToDo := ImportParame2;
  ListeRecep[2].Code := 'PS3';
  ListeRecep[2].ToDo := ImportParame3;
  ListeRecep[3].Code := 'PS4';
  ListeRecep[3].ToDo := ImportParame4;
  ListeRecep[4].Code := 'PS5';
  ListeRecep[4].ToDo := ImportParame5;
  ListeRecep[5].Code := 'EXO';
  ListeRecep[5].ToDo := ImportExercice;
  ListeRecep[6].Code := 'TL';
  ListeRecep[6].ToDo := ImportTableLibre;
  ListeRecep[7].Code := 'SSA';
  ListeRecep[7].ToDo := ImportSection;
  ListeRecep[8].Code := 'ETB';
  ListeRecep[8].ToDo := ImporteEtablissement;
  ListeRecep[9].Code := 'MDP';
  ListeRecep[9].ToDo := ImporteModepaiement;
  ListeRecep[10].Code := 'MDR';
  ListeRecep[10].ToDo := ImporteModereglement;
  ListeRecep[11].Code := 'DEV';
  ListeRecep[11].ToDo := ImporteDevise;
  ListeRecep[12].Code := 'REG';
  ListeRecep[12].ToDo := ImporteRegimeTVA;
  ListeRecep[13].Code := 'CGE';
  ListeRecep[13].ToDo := ImporteGeneraux;
  ListeRecep[14].Code := 'CAE';
  ListeRecep[14].ToDo := ImporteTiers;

  ListeRecep[15].Code := 'JAL';
  ListeRecep[15].ToDo := ImportJournaux;
  ListeRecep[16].Code := 'SAT';
  ListeRecep[16].ToDo := ImportSectionAnalytique;
  ListeRecep[17].Code := 'SOU';
  ListeRecep[17].ToDo := ImportSouche;
// ajout me 13-05-2002 format à partir de la version 002 des généraux
  ListeRecep[18].Code := 'CGN';
  ListeRecep[18].ToDo := ImporteGeneraux;

  ListeRecep[19].Code := 'RIB';
  ListeRecep[19].ToDo := ImportRib;
  ListeRecep[20].Code := 'BQC';
  ListeRecep[20].ToDo := ImportBanqueCp;

  ListeRecep[21].Code := 'BQE';
  ListeRecep[21].ToDo := ImportBanque;

  ListeRecep[22].Code := 'EXB';
  ListeRecep[22].ToDo := ImportReleveBanque;

  ListeRecep[23].Code := 'EXL';
  ListeRecep[23].ToDo := ImportLigneReleveBanque;

  ListeRecep[24].Code := 'CON';
  ListeRecep[24].ToDo := ImportContact;

  ListeRecep[25].Code := 'YTC';
  ListeRecep[25].ToDo := ImportTiersComp;

  // Cptes budgétaires
  ListeRecep[26].Code := 'FBG';
  ListeRecep[26].ToDo := ImportCpteBudget;

  // journal budgetaire
  ListeRecep[27].Code := 'FBJ';
  ListeRecep[27].ToDo := ImportJalBudget;

  // section budgetaire budgetaire
  ListeRecep[28].Code := 'FBS';
  ListeRecep[28].ToDo := ImportSectionBudget;

  // Table libre personnalisée
  ListeRecep[29].Code := 'CTL';
  ListeRecep[29].ToDo := ImportChoixCode;

  // Paramétrage Table libre personnalisée
  ListeRecep[30].Code := 'PTL';
  ListeRecep[30].ToDo := ImportParamlib;

  // ChoixExt
  ListeRecep[31].Code := 'XTL';
  ListeRecep[31].ToDo := ImportChoixExt;

  // ressource
  ListeRecep[32].Code := 'ARS';
  ListeRecep[32].ToDo := ImportRessource;

  // fiche 10134 format ccimpex des tiers
  ListeRecep[33].Code := 'CAU';
  ListeRecep[33].ToDo := ImporteTiersAncienFormat;


  ListeRecep[34].Code := 'RRL';
  ListeRecep[34].ToDo := ImporteRelance;

  // BAP
  ListeRecep[35].Code := 'CTI';
  ListeRecep[35].ToDo := ImporteBapVisa;

  ListeRecep[36].Code := 'CCI';
  ListeRecep[36].ToDo := ImporteBapCircuit;

  // table de correspondance
  ListeRecep[37].Code := 'CRR';
  ListeRecep[37].ToDo := ImporteCorresp;

  ListeRecep[38].Code := 'CRI';
  ListeRecep[38].ToDo := ImporteCorrespImp;

  // import ventilation type
  ListeRecep[39].Code := 'VEN';
  ListeRecep[39].ToDo := ImporteVentiltype;
                                                   
  // BVE 28.08.07 : import suivi validation
{$IFDEF CERTIFNF}
  ListeRecep[40].Code := 'CPJ';
  ListeRecep[40].ToDo := ImporteSuiviValidation;
{$ENDIF}

end;

{$IFDEF EAGLSERVER}
Function TImportCom.AfficheComImport(Chaine: string; Listecom : TListBox; pourcentage : double=0) : Boolean;
begin
     //DDWriteLn (Chaine);
     cWA.MessagesAuClient('COMSX.IMPORT','',Chaine) ;
     if Listecom = nil then
        writeln(FicIE, Chaine);
     Result := TRUE;
end;
{$ENDIF}

procedure TImportCom.ImporteEcritureBud(Ligne: string);
var
  T,T1                  : TOB;
  Section               : string;
  Compte                : string;
  Montant               : array[1..3] of Double;
  SMontant              : array[1..3] of string;
  Datecomptable         : TDateTime;
  Journal,Axe           : string;
begin
  Compte := BourreOuTronque(EnleveBlanc(Ligne, 14, 17), fbGene);
  Journal := EnleveBlanc(Ligne, 1, 3);
  // vérification des montants à zéro
  Montant[1] := 0.0; Montant[2] := 0.0; Montant[3] := 0.0 ;
  SMontant[1] := StPoint(EnleveBlanc(Ligne, 131, 20));
  SMontant[2] := StPoint(EnleveBlanc(Ligne, 176, 20));
  SMontant[3] := StPoint(EnleveBlanc(Ligne, 196, 20));
  if SMontant[1] <> '' then
    Montant[1] := Valeur(SMontant[1]);
  if SMontant[2] <> '' then
    Montant[2] := Valeur(SMontant[2]);
  if SMontant[3] <> '' then
    Montant[3] := Valeur(SMontant[3]);

  if (Montant[1] = 0.0) and (Montant[2] = 0.0) and (Montant[3] = 0.0) then exit;
  if (Montant[1] = 0.0) and (StrToint(VersionFichier) <= 4) then exit;
  Datecomptable := Format_Date_HAL(EnleveBlanc(Ligne, 4, 8))  ;
  Section := EnleveBlanc(Ligne, 32, 17);

  T1 := nil ;
  Axe := EnleveBlanc(Ligne, 219, 2, 'A1');
  if Axe = 'A1' then Section := BourreOuTronque(Section, fbAxe1) else
  if Axe = 'A2' then Section := BourreOuTronque(Section, fbAxe2) else
  if Axe = 'A3' then Section := BourreOuTronque(Section, fbAxe3) else
  if Axe = 'A4' then Section := BourreOuTronque(Section, fbAxe4) else
  if Axe = 'A5' then Section := BourreOuTronque(Section, fbAxe5);

  if TBUDEcr <> nil then
  begin
      T1 := TBUDEcr.FindFirst(['JOURNAL', 'PIECE','ETABLISSEMENT', 'SECTION', 'AXE', 'NATURE'],
      [Journal, EnleveBlanc(Ligne, 152, 8), EnleveBlanc(Ligne, 216, 3), Section, Axe, EnleveBlanc(Ligne, 302, 3)], FALSE);
      if T1 = nil then
      begin
           TBUDEcr.InsertOrUpdateDB(TRUE);
           TBUDEcr.Free;  TBUDEcr := nil;
      end;
  end;
  if TBUDEcr = nil then
  begin
     TBUDEcr := TOB.Create('Mere', nil, -1);
     T1 := TOB.Create ('CLEBUD',TBUDEcr,-1);
     T1.AddChampSupValeur('JOURNAL', EnleveBlanc(Ligne, 1, 3));
     T1.AddChampSupValeur('PIECE', EnleveBlanc(Ligne, 152, 8));
     T1.AddChampSupValeur('ETABLISSEMENT', EnleveBlanc(Ligne, 216, 3));
     T1.AddChampSupValeur('SECTION', Section);
     T1.AddChampSupValeur('AXE', Axe);
     T1.AddChampSupValeur('NATURE', EnleveBlanc(Ligne, 302, 3));
     OnAfficheListeCom(TraduireMemoire('Journal '+ EnleveBlanc(Ligne, 1, 3)),LISTEIMPORT);
     OnAfficheListeCom('         Période : ' + FormatDateTime (Traduitdateformat('mmmm yyyy'),Format_Date_HAL(EnleveBlanc(Ligne, 4, 8))) + ' Pièce :'+EnleveBlanc(Ligne, 152, 8), LISTEIMPORT);
  end;

  T := TOB.Create ('BUDECR', T1, -1);
  T.putvalue ('BE_EXERCICE', QUELEXODTBUD(Datecomptable));
  T.putvalue ('BE_BUDJAL', EnleveBlanc(Ligne, 1, 3));
  T.putvalue ('BE_DATECOMPTABLE', DateComptable);
  T.putvalue ('BE_RESOLUTION', EnleveBlanc(Ligne, 12, 2));
  T.putvalue ('BE_BUDGENE', Compte);
  T.putvalue ('BE_BUDSECT', Section);

  T.PutValue('BE_REFINTERNE', EnleveBlanc(Ligne, 49, 35));
  T.PutValue('BE_LIBELLE', VireTouteSub(EnleveBlanc(Ligne, 84, 35),'"'));
  if EnleveBlanc(Ligne, 130, 1) = 'D' then
     T.PutValue('BE_DEBIT', Montant[1])
  else
     T.PutValue('BE_CREDIT', Montant[1]);
  T.PutValue('BE_QUALIFPIECE', EnleveBlanc(Ligne, 151, 1));
  T.PutValue('BE_NUMEROPIECE', StrToint(EnleveBlanc(Ligne, 152, 8, '1')));
  T.PutValue('BE_ETABLISSEMENT', EnleveBlanc(Ligne, 216, 3));
  T.PutValue('BE_AXE', Axe);
  T.PutValue('BE_TYPESAISIE', EnleveBlanc(Ligne, 221, 2));
  T.PutValue('BE_DATECREATION', Date);
  T.PutValue('BE_AFFAIRE', EnleveBlanc(Ligne, 277, 17));
  T.PutValue('BE_NATUREBUD', EnleveBlanc(Ligne, 302, 3));
  if EnleveBlanc(Ligne, 305, 20) <> '' then
    T.PutValue('BE_QTE1', Valeur(StPoint(EnleveBlanc(Ligne, 305, 20))));
  if EnleveBlanc(Ligne, 325, 20) <> '' then
    T.PutValue('BE_QTE2', Valeur(StPoint(EnleveBlanc(Ligne, 325, 20))));
  if EnleveBlanc(Ligne, 345, 3) <> '' then
    T.PutValue('BE_QUALIFQTE1', EnleveBlanc(Ligne, 345, 3));
  if EnleveBlanc(Ligne, 348, 3) <> '' then
    T.PutValue('BE_QUALIFQTE2', EnleveBlanc(Ligne, 348, 3));

  T.PutValue('BE_LIBRETEXTE1', EnleveBlanc(Ligne, 351, 35));
  T.PutValue('BE_LIBRETEXTE2', EnleveBlanc(Ligne, 386, 35));
  T.PutValue('BE_LIBRETEXTE3', EnleveBlanc(Ligne, 421, 35));
  T.PutValue('BE_LIBRETEXTE4', EnleveBlanc(Ligne, 456, 35));
  T.PutValue('BE_LIBRETEXTE5', EnleveBlanc(Ligne, 491, 35));

  T.PutValue('BE_TABLE0', EnleveBlanc(Ligne, 526, 17));
  T.PutValue('BE_TABLE1', EnleveBlanc(Ligne, 543, 17));
  T.PutValue('BE_TABLE2', EnleveBlanc(Ligne, 560, 17));
  T.PutValue('BE_TABLE3', EnleveBlanc(Ligne, 577, 17));

end;

procedure ChargeTypeMvt (TypeMvt : string; var fstypes : array of TTypeCar);
var
i : integer;
begin
     for i := 0 to 5 do fstypes[i] := tcnone;
     // ligne avec des montants analytique et ligne gene
     if (TypeMvt[1] <> 'G') and (TypeMvt[1] <> 'P') then
        fstypes[0] := tcGene;
     if (TypeMvt[1] <> 'A') and (TypeMvt[1] <> 'G') and (TypeMvt[1] <> 'P')then
        fstypes[1] := tcpiece;
     // tous type de ligne sauf gene afin de rattaché à la ligne pièce
     if (TypeMvt[1] = 'A') or (TypeMvt[1] = 'H') or (TypeMvt[1] = 'O') or (TypeMvt[1] = 'G')
     or (TypeMvt[1] = 'P')  then
        fstypes[2] := tcnonblanc;
     if (TypeMvt[1] <> 'A') and (TypeMvt[1] <> 'O') and (TypeMvt[1] <> 'G') and (TypeMvt[1] <> 'P') then
        fstypes[3] := tcblanc;
     if (TypeMvt[1] = 'A') or (TypeMvt[1] = 'H') or (TypeMvt[1] = 'O') then
        fstypes[4] := tcAna;
end;

// Restriction analytique
function TImportCom.ResSectionExiste(Strc, StAxe, FCompte : string) : Boolean;
var
  CompteAna     : array[1..MaxAxe] of String;
  FRestriction  : TRestrictionAnalytique;        // Modèle de restriction ana
  Q             : TQuery;
  SR            : string;
begin
FillChar(CompteAna, sizeof(CompteAna), #0);
FRestriction := TRestrictionAnalytique.Create;
SR := FRestriction.GetClauseCompteAutorise(FCompte, StAxe, 'SECTION', CompteAna);
if SR = ' (1=1)' then Result := TRUE
else
begin
    Q:=OpenSql('Select S_SECTION From SECTION Where S_SECTION="'+Strc+'" And S_AXE="'+StAxe+'"'
      +' AND ' + SR
    ,True) ;
    Result:=Not Q.Eof ;
    Ferme(Q);
end;
FRestriction.free;
end;

procedure TImportCom.ImporteEcriture(Ligne: string);
var
  T, T1, T2, T3 : TOB;
  T4, T5        : TOB;
  TypeMvt       : string;
  periode       : integer;
  periodepiece  : integer;
  Q1            : TQuery;
  Ext           : string;
  EvtSup        : TEvtSup;
  NewNum        : integer;
  Compte        : string;
  Montant: array[1..3] of Double;
  SMontant: array[1..3] of string;
  JNatureJal    : string;
  TA1           : TOB;
  OkAlim        : Boolean;
  Jsouche       : string;
  TAJ,TS1       : TOB;
  PieceS1       : string;
  SFperiode     : string;
  Mtsolde       : double;
  MtsoldeDev    : double;
  DocCID        : string;
  ReqSql        : string;
  LL            : TOB;
  Okmajp        : Boolean;
  CompteCree    : string;
  totana        : double;
  CodeExercice  : string;
  Tcmpl         : TOB;
  Nb            : integer;
  OkDouble      : Boolean;
  DateEcriture  : TDateTime;
  Description   : string;
  OKCreatJ      : Boolean;
  Exerc         : TExoDate;
  Stypes        : array[0..5] of TTypeCar;
  Axe           : string;
  MvtImport     : TMouvement;
  procedure MajEcriture (num : integer; Majsouche: Boolean; md :string=''; nt: string=''; cpt : string=''; axe : string=''; sim : string=''); // fiche 10564
  var
  SAVModesaisie : string;
  begin
      remplir_compteattente;
      Genere_ecratEuro;
      TEcrParJournal.InsertDBByNivel(FALSE);
//      if (stArg = '') then
//              CBloqueurJournal (FALSE, CodeJournal, TRUE);
      if Majsouche then MiseajourSouche;

      TEcrParJournal.free; TEcrParJournal := nil;

      SAVModesaisie := Modesaisie;
      InitEcriture;
      if not Majsouche then Modesaisie := SAVModesaisie;
      T1 := TOB.Create('JRL', TEcrParJournal, -1);
      T1.AddChampSupValeur('J_JOURNAL', MvtImport.Journal);
      T1.AddChampSupValeur('J_DATEDERNMVT', iDate1900);
      T1.AddChampSupValeur('J_NUMDERNMVT', num);
      T1.AddChampSupValeur('J_MODESAISIE', md);
      T1.AddChampSupValeur('J_NATUREJAL', nt);
      T1.AddChampSupValeur('J_COMPTEURNORMAL', cpt);
      T1.AddChampSupValeur('J_COMPTEURSIMUL', sim);  //fiche 10564
      T1.AddChampSupValeur('J_Axe', axe);
      NbEcriturepiece := 0;
  end;
       // cas des pièces non renseignées  ou à zéro
  procedure AlimNumpiece (OkCalcul : Boolean=FALSE);
  var
  Ne : integer;
  begin
       Ne := 0;  NewNum := 0;
       // fiche 10207 rupture numéropiece en mode libre
       if (Modesaisie = '-') or (OrigineFichier = 'S1') or OkCalcul  then
            Ne := T1.GetValue('J_NUMDERNMVT')+1
       else
           if (DecoupLon <> 0) and (T2.FieldExists ('FPIECE')) then
              Ne := T2.GetValue('FPIECE')+1;

       if (Modesaisie <> '-') then
       begin
           if T1.Getvalue('J_COMPTEURNORMAL') <> '' then
              NewNum := GetNewNumJal(MvtImport.Journal, True, DateEcriture, T1.Getvalue('J_COMPTEURNORMAL'), Modesaisie)
           else
           if T1.Getvalue('J_COMPTEURSIMUL') <> '' then
              NewNum := GetNewNumJal(MvtImport.Journal, False, DateEcriture, T1.Getvalue('J_COMPTEURSIMUL'), Modesaisie)
           else
               NewNum := GetNewNumJal(MvtImport.Journal, False, DateEcriture);
       end;
       // fiche 10207 rupture numéropiece en mode libre
       if ((DecoupLon <> 0) or (Modesaisie = '-') or (OrigineFichier = 'S1') or OkCalcul) and (Ne > NewNum) then
          NewNum := Ne;
       numeropiece := IntToStr(NewNum);
  end;

begin
  T := nil;
  DocCID  := '';
  JNatureJal := '';
  Fillchar(MvtImport,SizeOf(MvtImport), #0) ;

  MvtImport.General := EnleveBlanc(Ligne, 14, 17);
  Compte := BourreOuTronque(MvtImport.General, fbGene);
  MvtImport.Journal := EnleveBlanc(Ligne, 1, 3);
  MvtImport.Typepiece := EnleveBlanc(Ligne, 12, 2);
  if (DateEcrChange <> iDate1900)  or (GarderJour <> '') then
  begin
     if (GarderJour <> '') then  // pour signifier qu'il faut garder le jour du fichier
     begin
      DateEcriture := Evaluedate(EnleveBlanc(Ligne, 4, 2)+EnleveBlanc(GarderJour, 4, 2)+EnleveBlanc(GarderJour, 7, 4));
     end else
      DateEcriture := DateEcrChange;
  end
  else
     DateEcriture := Format_Date_HAL(EnleveBlanc(Ligne, 4, 8))  ;

  if Gestiondoublon then OkDouble := TRUE
  else OKDouble := FALSE;
  OKCreatJ := TRUE;

  // application de la Table de correspondance
  TraiteCorrespCpt(8, EnleveBlanc(Ligne, 1, 3), MvtImport.Journal);
  TypeMvt := Copy(Ligne, 31, 1);

  if (TypeMvt[1] = 'G') and (not ScanBanck) then exit;

  // vérification des montants à zéro
  Montant[1] := 0.0; Montant[2] := 0.0; Montant[3] := 0.0 ;
  SMontant[1] := StPoint(EnleveBlanc(Ligne, 131, 20));
  SMontant[2] := StPoint(EnleveBlanc(Ligne, 176, 20));
  SMontant[3] := StPoint(EnleveBlanc(Ligne, 196, 20));
  if SMontant[1] <> '' then
    Montant[1] := Valeur(SMontant[1]);
  if SMontant[2] <> '' then
    Montant[2] := Valeur(SMontant[2]);
  if SMontant[3] <> '' then
    Montant[3] := Valeur(SMontant[3]);

  ChargeTypeMvt (TypeMvt, stypes);

 // if (TypeMvt[1] <> 'G') then
  if (Stypes[0] = tcGene) then
  begin
       if (Montant[1] = 0.0) and (Montant[2] = 0.0) and (Montant[3] = 0.0) then exit;
       if (Montant[1] = 0.0) and (StrToint(VersionFichier) <= 4) then exit;
       if not MontantNegatif and ( (Montant[1] < 0.0) or (Montant[2] < 0.0) and (Montant[3] < 0.0)) then
        begin
                      OnAfficheListeCom('Montant négatif, l''option "Montants négatifs autorisés" est modifiée dans les paramètres sociétés', LISTEIMPORT);
                      SetParamsoc('SO_MONTANTNEGATIF', TRUE);  MontantNegatif := TRUE;
        end;
  end;
  InitEcriture;
  // après initecriture car chargemagexo est dans ini pour format dos
  if CQuelExercice (DateEcriture, Exerc) then
             CodeExercice := Exerc.Code
  else
  begin
             OnAfficheListeCom(ERR_DATEECRITURE,LISTEIMPORT);
             OnAfficheListeCom(EnleveBlanc(Ligne, 0, 20),LISTEIMPORT);
             OKexercice := FALSE; exit;
  end;

  PieceS1 := '';
  OkAlim := FALSE;
  // Fiche 10299
  if (Stypes[0] = tcGene) then // on cherche les informations de la ligne piece sauf pour enreg bap et GEd
  begin
      AJouterListeSoldes (CodeExercice, SExercice);
      OkAlim := AlimPiece(Ligne, PieceS1);
      AJouterListeSoldes (MvtImport.Journal, SJournal);
      if (BSoldeLigne) and GetParamSocSecur('SO_OUITVAENC',FALSE) then
      begin
           if SJournal = nil then
           begin
                  SJournal :=  TStringList.Create;
                  SJournal.Sorted:=TRUE ;
                  SJournal.Duplicates := dupIgnore;
           end;
           SJournal.Add (MvtImport.Journal);
      end;
  end;

  T1 := RechercheOptimise(TEcrParJournal, 'J_JOURNAL', MvtImport.Journal);
  if T1 = nil then
  begin
        MajEcriture (0, TRUE);
        Q1 := OpenSql(' Select J_FERME,J_MODESAISIE,J_NATUREJAL,J_COMPTEURNORMAL,J_AXE,J_COMPTEURSIMUL from JOURNAL Where J_JOURNAL="' +
          MvtImport.Journal + '"', True);
        if not Q1.Eof then
        begin
          // journal est fermé fiche 10413
          if Q1.FindField ('J_FERME').asstring = 'X' then
          begin
              OnAfficheListeCom('Attention, le journal ' + MvtImport.Journal + ' est fermé' ,LISTEIMPORT);
              OKexercice := FALSE;  ferme(Q1); exit;
          end;
          Modesaisie := Q1.FindField('J_MODESAISIE').asstring;
          JNatureJal := Q1.FindField('J_NATUREJAL').asstring;

          if (TypeEcrI = 'N') or (TypeEcrI = 'I') or
          ((EnleveBlanc(Ligne, 151, 1)='N') and (TypeEcrI= '')) or
          ((EnleveBlanc(Ligne, 151, 1)='') and (TypeEcrI= '')) then  // fiche 10197
                  JSouche    := Q1.FindField('J_COMPTEURNORMAL').asstring
          else
          begin
                  JSouche    := Q1.FindField('J_COMPTEURSIMUL').asstring;
                  T1.PutValue('J_COMPTEURSIMUL', JSouche);
          end;
          if T1 <> nil then
          begin
               T1.putvalue('J_MODESAISIE', Modesaisie);
               T1.putvalue('J_NATUREJAL', JNatureJal);
               T1.putvalue('J_COMPTEURNORMAL', JSouche);
               T1.putvalue('J_AXE', Q1.FindField('J_AXE').asstring);
          end;
          if JNatureJal = 'ANO' then
          begin
              if TOBAno = nil then // ajout me cas journal sans parametres et si l'enregitrement exercice manque dans le fichier
              begin
                   TOBAno := TOB.Create('', nil, -1);
                   TA1 := TOB.Create ('',TOBAno,-1);
                   TA1.AddChampSupValeur('EXERCICE', CodeExercice);
                   TA1.AddChampSupValeur('ANO', 'H');
                   TA1.AddChampSupValeur('JOURNAL', MvtImport.Journal);
              end
              else
              begin
                   TA1 := RechercheOptimise(TOBAno, 'EXERCICE', CodeExercice);
                   if TA1 <> nil then
                   TA1.PutValue('JOURNAL', MvtImport.Journal);
              end;
          end;
        end
        else
        begin
              if AutoriseLaCreation ('IJA', 'Le journal '+ MvtImport.Journal + ' est manquant.') then
              begin
                 CreatJournal (MvtImport.Journal);   Modesaisie := 'LIB';
              end
              else if BIntegre then OKCreatJ := FALSE;
        end;
        ferme(Q1);

        if (OrigineFichier = 'S1') and (TOBJrl <> nil)then
        begin
          TAJ := RechercheOptimise(TOBJrl, 'JOURNAL', MvtImport.Journal);
          if TAJ <> nil then
          begin
             TAJ.PutValue ('MODESAISIE', 'LIB');
             TAJ.PutValue ('OK', TRUE);
             Modesaisie := 'LIB';
          end;
        end;                        // fiche 10340
        if (Modesaisie = '-') and ((OrigineFichier = 'SI') and (NatureFichier = 'BAL') or (NatureFichier = 'JRL') or (NatureFichier = 'SYN'))then // mode saisie
        begin
             Q1 := OpenSQl ('SELECT SH_NUMDEPART FROM SOUCHE Where SH_TYPE="CPT" and SH_SOUCHE="'+ Jsouche + '"', TRUE);
             if not Q1.EOF then
              T1.Putvalue('J_NUMDERNMVT', Q1.FindField ('SH_NUMDEPART').asinteger-1);
             ferme (Q1);
        end;
        // fiche 10148
         if (stArg = '') and (Modesaisie <> '-') and ( CEstBloqueJournal(MvtImport.Journal, false) ) then
             OnAfficheListeCom('Attention, il y a un autre utilisateur sur le journal : '+ MvtImport.Journal,LISTEIMPORT);

       // cas des pièces non renseignées  ou à zéro
        OKexercice := OnAfficheListeCom('Ecritures du journal : '+ MvtImport.Journal, LISTEIMPORT);
        if (stArg = '') and (JNatureJal = 'ANO') and (NatureFichier = 'JRL') and (EnleveBlanc(Ligne, 302, 3) = 'OAN') and (OrigineFichier <> 'S1')then
                 OnAfficheListeCom(TraduireMemoire('Journal '+ MvtImport.Journal + ' non traité, les mouvements sont de type PGI'),LISTEIMPORT);

  end;

  periode := GetPeriode(DateEcriture);
  if (Modesaisie = '-') then   // saisie pièce
  begin
     if (OrigineFichier = 'SI') and (EnleveBlanc(Ligne, 1, 2)='AA') then
       periodepiece := GetPeriodeJour(GetEncours.Deb)
     else
       periodepiece := GetPeriodeJour(DateEcriture);
     SFperiode := IntToStr(periodepiece);
     if RupturePiece then
        SFperiode := IntToStr(periodepiece)+EnleveBlanc(Ligne, 152, 8);
     if (Copy(Ligne, 31, 1) = 'O') then      // AJOUT ME ODA ANALYTIQUE rupture sur période et compte
        SFperiode := SFperiode+ MvtImport.General;
  end
  else
     SFperiode := IntToStr(periode);

 // sur la ligne ged il n'y a pas les devises et etablissement on prend la ligne prédédent
  if (TypeMvt[1] <> 'G') then
  begin
      SDevise :=  EnleveBlanc(Ligne, 160, 3, V_PGI.DevisePivot);
      SEtablissement := EnleveBlanc(Ligne, 216, 3);
      if SEtablissement = '' then
      begin
         if OptionEchange and (GetParamSocSecur ('SO_CPEDIETABLISSEMENTDEF', '') <> '') then
                SEtablissement := GetParamSocSecur ('SO_CPEDIETABLISSEMENTDEF', GetParamSocSecur ('SO_ETABLISDEFAUT', '001'))
         else
                SEtablissement := GetParamSocSecur ('SO_ETABLISDEFAUT', '001');
      end;
  end;
  // Pour la prise en compte des devises dans les ruptures des pièces
  if (Modesaisie <> 'BOR') then SFperiode := SFperiode + SDevise;
  // Pour la prise en compte des établissements dans la rupture
  SFperiode := SFperiode + SEtablissement;

  T2 := T1.FindFirst(['FPERIODE'], [SFperiode], FALSE);
  // fiche 10576  // A PORTER sur 7XX
  if (OrigineFichier <> 'SI') and ((Stypes[3] = tcblanc) or (TypeMvt[1] = 'H')) and (T2 <> nil) and (T2.detail.count <> 0 ) and (NatureFichier <> 'SYN') then
  begin
       { pb de perte des écritures en format dossier si position 152 sur 8  le n° pièce est renseigné
        on ne calcul pas le numéro de pièce }
       if (NatureFichier = 'JRL') or ((NatureFichier = 'DOS') and OkAlim) then
       begin
           T3 := T2.detail[T2.detail.count-1];
           if (T3 <> nil) and (T3.detail.count > 200) and (T3.Getdouble('SOLDE') = 0.0) then
           begin
               OnAfficheListeCom('Enregistrement des écritures ',LISTEIMPORT);
               MajEcriture (T1.Getvalue('J_NUMDERNMVT'), FALSE, T1.Getvalue('J_MODESAISIE'), T1.Getvalue('J_NATUREJAL'), T1.Getvalue('J_COMPTEURNORMAL'), T1.Getvalue('J_AXE'), T1.Getvalue('J_COMPTEURSIMUL'));  // fiche 10564
               T2 := nil;
           end;
       end;
  end;

  if T2 = nil then
  begin
    if T1.getvalue('J_NATUREJAL') = 'ANO' then
    begin
      TA1 := RechercheOptimise(TOBAno, 'EXERCICE', CodeExercice);
      if (TA1 <> nil) and (TA1.GetValue('JOURNAL') = '') then
         TA1.PutValue('JOURNAL', MvtImport.Journal);
    end;
    // ajout me  dans le cas des gros dossier
    Okmajp :=  (NbEcriturepiece > 1000);
    if Okmajp then
    begin
       OKexercice := OnAfficheListeCom('Enregistrement des écritures :'+ IntToStr(NbEcriturepiece), LISTEIMPORT);
       MajEcriture (T1.Getvalue('J_NUMDERNMVT'), FALSE, T1.Getvalue('J_MODESAISIE'), T1.Getvalue('J_NATUREJAL'), T1.Getvalue('J_COMPTEURNORMAL'), T1.Getvalue('J_AXE'), T1.Getvalue('J_COMPTEURSIMUL'));  // fiche 10564
    end;

    T2 := TOB.Create('PERIODE'+SFperiode, T1, -1);
    T2.AddChampSupValeur('FPERIODE', SFperiode);
// Pb des écritures dans le même journal et même période ayant des devises mélangées ajout TRUE
    if OkAlim then
       AlimNumpiece (TRUE);
    T2.AddChampSupValeur('FPIECE', numeropiece);
    T2.AddChampSupValeur('NPIECE', MvtImport.Typepiece);
    T2.AddChampSupValeur('NBDECOUP', '0');

    T1.putValue('J_DATEDERNMVT', DateEcriture);

    // suppression ecriture
    if (not Okmajp) and (((OrigineFichier <> 'S1') and (not Paramgene) and (NatureFichier = 'DOS'))) then
    begin
        EvtSup := TEvtSup.Create;
        With (EvtSup) do
        begin
             Journal  := MvtImport.Journal;
             DateComptable := DateEcriture;
             Codeexo  :=  CodeExercice;
             vperiode := periode;
             Datec    := '';
             if (Modesaisie = '-') then Datec :=USDATETIME(DateComptable);
             parpiece := FALSE;
             if Transactions(DegageEcriture, 1) <> oeOK then
                        OnAfficheListeCom('Suppression Ecriture impossible',LISTEIMPORT);
             Free;
        end;
    end;


    if (Modesaisie = '-') then   // saisie pièce
    begin
       if Savperiode <> FormatDateTime (Traduitdateformat('mmmm yyyy'), DateEcriture) then
          OKexercice := OnAfficheListeCom('         Période : ' + FormatDateTime (Traduitdateformat('mmmm yyyy'), DateEcriture), LISTEIMPORT);
       Savperiode  := FormatDateTime (Traduitdateformat('mmmm yyyy'), DateEcriture);
    end
    else
       OKexercice := OnAfficheListeCom('         Période : ' + FormatDateTime (Traduitdateformat('mmmm yyyy'), DateEcriture), LISTEIMPORT);
  end
  else
  begin
       if OkAlim or (DecoupLon <> 0) then numeropiece := T2.Getvalue('FPIECE');
  end;

  if (PieceS1 <> '') then
  begin
       TS1     := T2.FindFirst(['E_NUMEROPIECES1'], [PieceS1], FALSE);
       if TS1 <> nil then
       begin
                     numeropiece := TS1.getvalue ('E_NUMEROPIECE');
                     DocCID      := TS1.getvalue ('E_DOCIDS1');
       end
       else
       begin
            if (NatureFichier = 'DOS') then
            begin
                             AlimNumpiece (TRUE);
                             T2.PutValue('FPIECE', numeropiece);
            end
            else
            begin
                     ReqSql :=  'select E_NUMEROPIECE, E_REFREVISION, EC_DOCGUID from ecriture left outer join ecrcompl on ec_exercice=e_exercice and and e_journal=ec_journal '+
                     ' and ec_datecomptable=e_datecomptable and ec_numeropiece=e_numeropiece and ec_numligne=e_numligne where E_JOURNAL="'+ MvtImport.Journal + '" AND E_REFREVISION='+
                           PieceS1;
                     if (OrigineFichier <> 'S1') then
                     ReqSql := ReqSql + ' and E_PERIODE='+ IntToStr(periode);

                     Q1 := Opensql (ReqSql, TRUE);
                     if not Q1.EOF then
                     begin
                         numeropiece := IntToStr(Q1.FindField ('E_NUMEROPIECE').asinteger);
                         DocCID      := Q1.FindField ('EC_DOCGUID').asstring;
                         ferme (Q1);
                     end
                     else
                     begin
                         ferme (Q1);
                         AlimNumpiece (TRUE);
                         T2.PutValue('FPIECE', numeropiece);
                     end;
            end;
       end;
  end;
  if (numeropiece = '') or (numeropiece = '0') then
  begin
        AlimNumpiece (TRUE);
        T2.PutValue('FPIECE', numeropiece);
  end;

  if Gestiondoublon  and (TypeMvt[1] <> 'A') then
  begin
       OkDouble := Okdoublon(Ligne, LL, MvtImport.Journal);
       if OkDouble then
       begin
             if BRejet then OkRejet := TRUE;
             if (not BIntegre)
             then
             begin
                if not BRejet then OKexercice := FALSE; // permet de sortir de suite
                exit;
             end
       end;
  end;

  // pour ne pas remonter le journal type ANO (anouveau) en mode journal
  if (stArg = '') and (NatureFichier = 'JRL')and (OrigineFichier <> 'S1')then
  begin
     if TOBAno <> nil then
     begin
          TA1 := TOBAno.FindFirst(['JOURNAL','ANO'], [MvtImport.Journal,'OAN'], FALSE);
          if (TA1 <> nil) and (EnleveBlanc(Ligne, 302, 3) = 'OAN') then
          exit;
     end;
  end;

  Mtsolde:= 0.0;  MtsoldeDev := 0.0;                                              //fiche 10164
  if ((DecoupLon <> 0) or ((Modesaisie = '-') and OkAlim and (not RupturePiece)))
  and (Stypes[1] = tcpiece) then  // saisie pièce
  begin
                     T3     := T2.FindFirst(['E_NUMEROPIECE', 'NATUREPIECE'], [numeropiece, MvtImport.Typepiece], FALSE);
                     if ((Modesaisie = '-') and (T3 <> nil)  and (Arrondi(T3.Getdouble('SOLDE'), V_PGI.OkDecV)= 0.0))
                     or ((Modesaisie = '-') and (T3 = nil) and (MvtImport.Typepiece <>  (T2.GetValue('NPIECE'))))
                     or (T2.GetValue('NBDECOUP') >= DecoupLon+5)
                     or (((T2.GetValue('NBDECOUP') >= DecoupLon) and (T3 <> nil)  and (Arrondi(T3.GetDouble('SOLDE'), V_PGI.OkDecV)= 0.0))) then
                     begin
                            AlimNumpiece;
                            T2.PutValue('FPIECE', numeropiece);
                            T3 := nil;
                            if DecoupLon <> 0 then T2.PutValue('NBDECOUP', '0');
                     end;
                     T2.PutValue('NPIECE', MvtImport.Typepiece);
                     if DecoupLon <> 0 then //Découpage de folio
                     begin
                        Nb := T2.GetValue('NBDECOUP');
                        inc(Nb);
                        T2.PutValue('NBDECOUP', IntTostr(Nb));
                     end;
  end
  else
  begin                                                                // si conserver les numéro de piece n'est pas coché
       if RupturePiece and (TypeMvt[1] <> 'G') and (TypeMvt[1] <> 'A') and (Calculpiece <> 'N') then // fiche 10207 rupture numéropiece en mode libre
       begin
                     T3     := T2.FindFirst(['E_NUMEROPIECE', 'NUMFOLIO'], [numeropiece, EnleveBlanc(Ligne, 152, 8)], FALSE);
                     if (T3 = nil) then
                     begin
                            AlimNumpiece(TRUE);
                            T2.PutValue('FPIECE', numeropiece);
                     end;
       end
       else
                     T3     := T2.FindFirst(['E_NUMEROPIECE'], [numeropiece], FALSE);
                     // cas changement de date dans BL
                     // date1/pieceS1N°1/pieceS5N°1 et date2/pieceS1N°2/pieceS5N°1
                     // changement date1/pieceS1N°1/pieceS5N°1 en date2/pieceS1N°1/pieceS5N°1
                     // on se trouve avec  date2/pieceS1N°1/pieceS5N°1 et date2/pieceS1N°2/pieceS5N°1
                     // perte de  date2/pieceS1N°1/pieceS5N°1
                     if (NatureFichier = 'SYN') and (PieceS1 <> '') and (T3 <> nil) then
                     begin
                                       T3     := T2.FindFirst(['E_NUMEROPIECE', 'E_NUMEROPIECES1'], [numeropiece, PieceS1], FALSE);
                                       if T3 = nil then
                                       begin
                                              AlimNumpiece (TRUE);
                                              T2.PutValue('FPIECE', numeropiece);
                                       end;
                     end;
  end;
  if T3 = nil then
  begin
    T3 := TOB.Create('NUMEROPIECE', T2, -1);
    T3.AddChampSupValeur('E_NUMEROPIECE', numeropiece);
    T3.AddChampSupValeur('E_NUMEROPIECES1', PieceS1);

    // Traitement Piece S1 et récupération DOCID
    if (OrigineFichier = 'S1') and (NatureFichier = 'SYN') and (PieceS1 = '') then
    begin
                           ReqSql :=  'select E_NUMEROPIECE, E_REFREVISION, EC_DOCGUID from ecriture left outer join ecrcompl on ec_exercice=e_exercice and and e_journal=ec_journal '+
                           ' and ec_datecomptable=e_datecomptable and ec_numeropiece=e_numeropiece and ec_numligne=e_numligne ';

                           ReqSql := ReqSql +  ' where E_EXERCICE="' + CodeExercice+ '"'+ 'AND E_JOURNAL="'+ MvtImport.Journal + '" AND E_NUMEROPIECE='+ numeropiece+
                           ' and E_PERIODE='+ IntToStr(periode);

                           Q1 := Opensql (ReqSql, TRUE);
                           if not Q1.EOF then
                               DocCID      := Q1.FindField ('EC_DOCGUID').asstring;

                           ferme (Q1);

    end;
    T3.AddChampSupValeur('E_DOCIDS1', DocCID);

    T3.AddChampSupValeur('E_DOCIDGED', '');  // ajout me pour GED
    T3.AddChampSupValeur('E_DOCIDFIC', '');
    T3.AddChampSupValeur('NATUREPIECE', MvtImport.Typepiece);
    T3.AddChampSupValeur('SOLDE', Arrondi(Mtsolde, V_PGI.OkDecV));
    // Fiche  10361
    T3.AddChampSupValeur('SOLDEDEV', MtsoldeDev);

    T3.AddChampSupValeur('COMPTEINSERT', '');
    T3.AddChampSupValeur ('NUMFOLIO',EnleveBlanc(Ligne, 152, 8));
    if not Gestiondoublon then OkRejet := (not OKCreatJ);
    if OkRejet then
       T3.AddChampSupValeur('OKREJET', 'X')
    else
       T3.AddChampSupValeur('OKREJET', '-');

    if OkDouble then
       T3.AddChampSupValeur('OKDOUBLON', 'X')
    else
       T3.AddChampSupValeur('OKDOUBLON', '-');

    // cas synchro S1 pb numero de piece existant est inférieur à numdernmvt
    if  (StrtoInt(numeropiece) > StrToInt(T1.GetValue('J_NUMDERNMVT'))) then
       T1.putValue('J_NUMDERNMVT', StrtoInt(numeropiece));

    // suppression ecriture
    if ((NatureFichier = 'JRL') and (Calculpiece = 'N')) or (not Paramgene) and ((NatureFichier = 'SYN')  or (OrigineFichier = 'S1')) then
    begin
        EvtSup := TEvtSup.Create;
        with (EvtSup) do
        begin
             Journal       := MvtImport.Journal;
             DateComptable := DateEcriture;
             Codeexo       := CodeExercice;
             vperiode      := periode;
             npiece        := numeropiece;
             refrevision   := '';
             // pour mode saisie pièce suppression sur la référence S1
             if(NatureFichier = 'SYN') and (OrigineFichier = 'S1') and (PieceS1 <> '') then
                              refrevision   := PieceS1;
             // cas piece venant de BL, modifier dans S5(ex ajout des lignes), envoi dans BL, puis changement de date dans BL + renvoie dans S5
             if(NatureFichier = 'SYN') and (OrigineFichier = 'S1') and (PieceS1 = '') and (EnleveBlanc(Ligne, 803, 6) <> '') then
                              refrevision   := EnleveBlanc(Ligne, 803, 6);

             parpiece      := TRUE;
            // en synchro pour ne pas écraser le lettrage
              if ((VersionFichier <> '') and (StrToint(VersionFichier) >= 6)) then
              begin
                   if (TPieceCon <> nil) then
                   begin
                        DelettrageSynchro(T2);
                        TPieceCon.free;  TPieceCon := nil;
                   end;
                 Q1 := OpenSql('SELECT E_JOURNAL,E_EXERCICE,E_DATECOMPTABLE,E_NUMEROPIECE,E_GENERAL,E_AUXILIAIRE,E_DEBITDEV, E_CREDITDEV,E_DEVISE, '+
                   ' E_LETTRAGE, E_LETTRAGEDEV, E_ECHE, E_ETATLETTRAGE, E_DATEPAQUETMIN, E_DATEPAQUETMIN, E_DATEPAQUETMAX, '+                  // fiche 10531
                   ' E_COUVERTURE, E_COUVERTUREDEV, E_DATEECHEANCE, E_NUMECHE, E_REFPOINTAGE, E_DATEPOINTAGE,E_REFREVISION,E_PAQUETREVISION,E_IO,E_BLOCNOTE,E_TABLE0,E_TABLE1,E_TABLE2,E_TABLE3 '+
                   ' FROM Ecriture WHERE E_EXERCICE="' + Codeexo+ '"'

  //                 Q1 := OpenSql('SELECT * FROM Ecriture WHERE E_EXERCICE="' + Codeexo+ '"'
                      + ' AND E_JOURNAL="' + Journal + '" AND E_PERIODE=' +
                      IntToStr(VPeriode) + ' AND E_NUMEROPIECE=' + npiece, TRUE);
                   if (TPieceCon = nil) then TPieceCon := TOB.Create('', nil, -1);
                   TPieceCon.LoadDetailDB('Ecriture', '', '', Q1, TRUE, FALSE);
                   Ferme(Q1);
              end;
              if Transactions(DegageEcriture, 1) <> oeOK then
                 OnAfficheListeCom('Suppression Ecriture impossible',LISTEIMPORT);
              Free;
        end;
    end;
  end
  else
  begin
      if Gestiondoublon and (not OkDouble) and (TypeMvt[1] <> 'A') then
         if (T3.GetValue('OKDOUBLON') = 'X') then T3.PutValue('OKDOUBLON', '-');
      OkRejet := (T3.GetValue('OKREJET')='X');
  end;

  Ext := 'E_';

  if (Stypes[2] = tcnonblanc) then
  T := T3.FindFirst(['E_NUMEROPIECE', 'E_NUMLIGNE'], [numeropiece, T3.detail.Count], FALSE);

  if T = nil then
  begin
      if TypeMvt[1] = 'O' then
      begin
        T := TOB.Create('ODA', T3, -1); // ODA
        T.AddChampSupValeur('E_NUMEROPIECE', Valeur(numeropiece));
        T.AddChampSupValeur('E_NUMLIGNE', T3.detail.Count);
        Ext := 'Y_';
      end
      else
        T := TOB.Create('ECRITURE', T3, -1);
        if Ext = 'E_' then
             CPutDefautEcr (T);
  end
  else
  begin
        if (TypeMvt[1] = 'H') then
        begin
              T := TOB.Create('ECRITURE', T3, -1);
              if Ext = 'E_' then
                   CPutDefautEcr (T);
        end;
  end;

  // fiche 10148
  if (stArg = '') and (Modesaisie <> '-') and  (CEstBloqueBor(MvtImport.Journal,DateEcriture,StrtoInt(numeropiece),TRUE)) then
  begin
             OnAfficheListeCom('Il y a un autre utilisateur sur  le journal :'+ MvtImport.Journal  +
             ' de la date ' + Datetostr(DateEcriture),LISTEIMPORT);
             OnAfficheListeCom(EnleveBlanc(Ligne, 0, 20),LISTEIMPORT);
             OKexercice := FALSE; exit;
  end;

  if (Stypes[3] = tcblanc) or (TypeMvt[1] = 'H') then 
  begin
    T.PutValue(Ext + 'NUMLIGNE', T3.detail.Count);
    T.PutValue(Ext + 'NUMEROPIECE', Valeur(numeropiece));
     if ((T3.GetValue('OKDOUBLON') <> 'X') and Gestiondoublon) or (not Gestiondoublon)
     then inc (NbEcriture);
    inc (NbEcriturepiece);
    // cas d'un ajout journal pour prise en compte dans synchro
    if (OrigineFichier <> 'S1') and (NatureFichier = 'JRL') then
    if Ext = 'E_' then T.PutValue(Ext + 'IO', 'X');

    CompteCree := remplir_ecriture(Ext, Ligne, Modesaisie, periode, T, MvtImport.Journal, DateEcriture, Exerc);
    if CompteCree  <> '' then
       T3.PutValue('COMPTEINSERT', CompteCree)
    else  T3.PutValue('COMPTEINSERT', '');

    if (not AutoriseEcheck) (*and (pos('IGE;IAU', InterdireLaCreation) <> 0)*) then
    begin
       if (T.GetValue(Ext + 'AUXILIAIRE')= 'INCOHERENT') or (BIntegre and (not OKCreatJ)) then
          T3.PutValue('OKDOUBLON', 'I')
    end;

    if OkRejet then
    begin
        T3.PutValue('OKREJET', 'X');
        if BRejet then   // fiche 10551
        begin
            T3.PutValue('OKDOUBLON', 'I');
            if AutoriseEcheck then AutoriseEcheck := FALSE;
        end;
    end;

      // pour synchro si pas de mondification recherche si
    if(NatureFichier = 'SYN') then
            Okdoublon(Ligne, T, MvtImport.Journal, TRUE);
    Mtsolde := T.GetDouble (ext + 'DEBIT')- T.GetDouble (ext + 'CREDIT');
    T3.PutValue('SOLDE', (Arrondi(T3.Getdouble('SOLDE'), V_PGI.OkDecV) + Mtsolde));
    // fiche 10361
    MtsoldeDev := T.GetDouble (ext + 'DEBITDEV')- T.GetDouble (ext + 'CREDITDEV');
    T3.PutValue('SOLDEDEV', (Arrondi(T3.GetDouble('SOLDEDEV'), V_PGI.OkDecV) + MtsoldeDev));
  end;

  // si axe  renseigné pour analytique
  if Stypes[4] = tcAna then
  begin
    Axe := EnleveBlanc(Ligne, 219, 2, 'A1');
    // sur un compte non ventilable on ne remonte pas l'analytique
    TraiteCorrespCpt(0, Compte, MvtImport.General); // fiche 10617
    if not OKVentilable (MvtImport.General) and (T3.GetValue('COMPTEINSERT') = '') then exit
    else    // en cas de création d'un compte par la comsx le rendre ventilable s'il y a de l'analytique
    begin
        if T3.GetValue('COMPTEINSERT') <> '' then
             MetVentilable (T3.GetValue('COMPTEINSERT'), Axe)
         else
         begin
              if Axe[2] in  ['1'..'5'] then // fiche 10443
              begin
                   if  not VentAttent[StrToInt(Axe[2])] then
                   begin
                       OnAfficheListeCom('Attention, Le compte ' + compte + ' n''est pas ventilable sur l''axe ' + Axe,LISTEIMPORT);
                       exit;
                   end;
              end;
         end;
    end;
    T4 := T.FindFirst(['Y_AXE'], [Axe], FALSE);
    if T4 = nil then
    begin
             T4 := TOB.Create(Axe, T, -1);
             T4.AddChampSupValeur('Y_AXE', Axe);
    end;
    Ext := 'Y_';
    T5 := TOB.Create('ANALYTIQ', T4, -1);
    // numéro de la ligne d'écriture
    T5.PutValue(Ext + 'NUMLIGNE', T3.detail.Count);
    // numéro de la ligne analytique
    T5.PutValue(Ext + 'NUMVENTIL', T4.detail.Count);
    T5.PutValue(Ext + 'NUMEROPIECE', Valeur(numeropiece));

    // Restriction analytique
    if AppliquerRest and (not ResSectionExiste(EnleveBlanc(Ligne, 32, 17), Axe, MvtImport.General)) then
    begin
              OnAfficheListeCom('Il existe des incompatibilités entre le modèle de ventilations à appliquer et ',LISTEIMPORT);
              OnAfficheListeCom('les modèles de restrictions analytiques attachés aux comptes généraux sélectionnés',LISTEIMPORT);
              OnAfficheListeCom(EnleveBlanc(Ligne, 1, 50),LISTEIMPORT);
    end
    else
    begin
              remplir_ecriture(Ext, Ligne, Modesaisie, periode, T5, MvtImport.Journal, DateEcriture, Exerc, T3.GetValue('COMPTEINSERT'));
              if T.GetValue ('E_AUXILIAIRE') <> '' then  // fiche 21811
               T5.PutValue ('Y_AUXILIAIRE', T.GetValue ('E_AUXILIAIRE'));
               // fiche 21951
              if (TypeMvt[1] <> 'O') and OKexercice and GetParamSocSecur('SO_CROISAXE', False)  and (T5.Getvalue ('Y_POURCENTAGE') = 0) then
              begin
                      Mtsolde := T.GetDouble ('E_DEBIT') + T.GetDouble ('E_CREDIT');
                      totana := T5.GetDouble ('Y_DEBIT') + T5.GetDouble ('Y_CREDIT');
                      if Mtsolde <> 0 then
                      T5.Putvalue ('Y_POURCENTAGE', Arrondi((totana/Mtsolde)*100, ADecimP)) ;
              end;
              if (TypeMvt[1] = 'O') and (Modesaisie = '-') and OkAlim then // ODA ANALYTIQ
              begin
                   if (T1.Getvalue('J_AXE') <> '') and (Axe <>  T1.Getvalue('J_AXE')) then
                   begin
                        OnAfficheListeCom('L''axe analytique est différent de l''axe du journal',LISTEIMPORT);
                        if stArg = '' then
                           AffMessage ('L''axe analytique est différent de l''axe du journal');
                         OKexercice := FALSE; exit;
                   end;
                   Mtsolde:= T5.GetDouble (ext + 'DEBIT')- T5.GetDouble (ext + 'CREDIT');
                   T3.PutValue('SOLDE', (T3.Getdouble('SOLDE')+Mtsolde));
              end;
    end;
  end;

  if TypeMvt[1] = 'G' then // document GED         Format_Date_HAL(EnleveBlanc(Ligne, 4, 8))
  begin
       Description := FormatDateTime (Traduitdateformat('mmmm yyyy'), DateEcriture)
                                        + ' Jrl :' + MvtImport.Journal
                                        + ' Folio :' + IntToStr(T.Getvalue(Ext + 'NUMEROPIECE'));
       if (Modesaisie = 'BOR')  then
       begin
             if (T3.GetValue('E_DOCIDFIC') <> TrimRight(Copy(Ligne, 84, 255))) then
             begin
               AjouterFichierDansGed (TrimRight(Copy(Ligne, 84, 255)), T.GetValue(Ext + 'AUXILIAIRE'), '###', Description, 'COMSX', EnleveBlanc(Ligne, 8, 4), EnleveBlanc(Ligne, 6, 2), DocCID);
               if DocCID  <> '' then
                    T3.Putvalue('E_DOCIDGED', DocCID);
               T3.PutValue('E_DOCIDFIC', TrimRight(Copy(Ligne, 84, 255)));
             end
             else DocCID := T3.GetValue('E_DOCIDGED');
       end
       else
       begin
           if T3.detail.count = 1 then // ajout document GED uniquement la première ligne
           begin
             AjouterFichierDansGed (TrimRight(Copy(Ligne, 84, 255)), T.GetValue(Ext + 'AUXILIAIRE'), '###', Description, 'COMSX', EnleveBlanc(Ligne, 8, 4), EnleveBlanc(Ligne, 6, 2), DocCID);
             if DocCID  <> '' then
                  T3.Putvalue('E_DOCIDGED', DocCID);
           end else DocCID := '';
       end;
  end;

  // ECRITURES COMPLEMENTAIRES
  if (TypeMvt[1] <> 'A') and (TypeMvt[1] <> 'O') and ((EnleveBlanc(Ligne, 1178, 8) <> '') or (DocCID <> '')) then
  begin
                    // Fiche 10299
                     Tcmpl := T.FindFirst(['EC_NUMEROPIECE', 'EC_NUMLIGNE'], [T.Getvalue(Ext + 'NUMEROPIECE'), T.Getvalue(Ext + 'NUMLIGNE')], FALSE);
                     if Tcmpl = nil then
                     begin
                           Tcmpl := TOB.Create('ECRCOMPL', T, -1);
                           Tcmpl.PutValue ('EC_EXERCICE', T.Getvalue(Ext + 'EXERCICE'));
                           Tcmpl.PutValue ('EC_JOURNAL', T.Getvalue(Ext + 'JOURNAL'));
                           Tcmpl.PutValue ('EC_DATECOMPTABLE', T.Getvalue(Ext + 'DATECOMPTABLE'));
                           Tcmpl.PutValue ('EC_NUMEROPIECE', T.Getvalue(Ext + 'NUMEROPIECE'));
                           Tcmpl.PutValue ('EC_NUMLIGNE', T.Getvalue(Ext + 'NUMLIGNE'));
                           Tcmpl.PutValue ('EC_QUALIFPIECE', T.Getvalue(Ext + 'QUALIFPIECE'));
                           Tcmpl.PutValue ('EC_NUMECHE', T.Getvalue(Ext + 'NUMECHE'));
                           Tcmpl.PutValue ('EC_CUTOFFDEB', Format_Date_HAL(EnleveBlanc(Ligne, 1178, 8, stDate1900)));
                           Tcmpl.PutValue ('EC_CUTOFFFIN', Format_Date_HAL(EnleveBlanc(Ligne, 1186, 8, stDate1900)));
                           Tcmpl.PutValue ('EC_CUTOFFDATECALC', Format_Date_HAL(EnleveBlanc(Ligne, 1194, 8)));
                           Tcmpl.PutValue ('EC_CLEECR', EnleveBlanc(Ligne, 1202, 35, '0'));  // fiche 32913
                           Tcmpl.PutValue ('EC_DOCGUID', DocCID);
                     end
                     else
                     begin
                          if (TypeMvt[1] = 'G') then
                           Tcmpl.PutValue ('EC_DOCGUID', DocCID);
                     end;
  end;
  // Import des con à payer
  if (TypeMvt[1] = 'P') then
     ImporteBonapayer(Ligne, T.Getvalue(Ext + 'EXERCICE'), DocCID, T.Getvalue(Ext + 'NUMEROPIECE'));
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : M.ENTRESSANGLE
Créé le ...... : 22/03/2002
Modifié le ... :   /  /
Description .. : Fonction d'import ComSx
Mots clefs ... : IMPORTCOMSX
*****************************************************************}
procedure TImportCom.LibereTob;
begin
  if ListeEtab <> nil then begin ListeEtab.free; ListeEtab := nil end;
  if TEcrParJournal <> nil then begin TEcrParJournal.Free; TEcrParJournal := nil end;
  if TOBAno <> nil then begin TOBAno.Free; TOBAno := nil; end;
  IF TOBPIECE <> nil then begin TOBPIECE.Free; TOBPIECE := nil; end;
  if TOBLettre <> nil then
  begin
         if assigned(fHLettre) then
         begin
              fHLettre.Clear(True);
              FreeAndNil (fHLettre);
         end;
         TOBLettre.free; TOBLettre := nil;
  end;
  if TOBJrl <> nil then begin TOBJrl.Free; TOBJrl := nil; end;
  if TPieceCon <> nil then begin TPieceCon.free; TPieceCon := nil; end;
  if vInfo <> nil then begin vInfo.Free;  vInfo:= nil; end;
  if ZDevise <> nil then begin ZDevise.free; ZDevise := nil; end;
end;

procedure TImportCom.Genere_ecratEuro;
var
i,ii,iii,kk,tv,NumPiece,ij          : integer;
TLec,TLecPeriode,TLecpiece          : TOB;
TOBAxe,TOBana                       : TOB;
fb                                  : TFichierBase;
Section, Str                        : string;
NumVentil,NumAxe,DernMvt            : integer ;
Pourcentage                         : double ;
TVentil,T, Ta, TlecEcc              : TOB;
Q                                   : TQuery;
TCroisaxe                           : TTraitementCroisaxe;
TobResult,TlecNew,TLecJournal       : TOB;
SaveCompte                          : string;

         procedure ChargeVentil (Cpte, NumAxe : string);
         var
         CompteAna     : array[1..MaxAxe] of String;
         Sql           : string;
         Restriction   : TRestrictionAnalytique;
         begin
                    FillChar(CompteAna, sizeof(CompteAna), #0);
                    Restriction := TRestrictionAnalytique.Create;

                    TVentil := TOB.Create('', nil, - 1);
                    Sql := 'SELECT * FROM VENTIL WHERE V_NATURE="GE'+NumAxe+'" AND V_COMPTE="'+ Cpte +'"';
                   if AppliquerRest then
                   Sql := Sql +' AND ' + Restriction.GetClauseCompteAutorise(
                                Cpte, 'A'+NumAxe, 'VENTIL', CompteAna);
                    Q := OpenSQL (Sql, True);
                    TVentil.LoadDetailDB('VENTIL','','', Q, TRUE, FALSE);
                    Ferme (Q);
                    Restriction.free;
         end;
         procedure AffectVentil;
         begin
               TOBAxe := Tlec.detail[kk].FindFirst(['Y_AXE'], ['A'+inttostr(NumAxe)], FALSE);
               if TOBAxe = nil then   // pb de plusieurs ventilation sur même axe
               begin
                     TOBAxe := TOB.Create('A'+inttostr(NumAxe), Tlec.detail[kk], -1);
                     TOBAxe.AddChampSupValeur('Y_AXE', 'A'+inttostr(NumAxe));
               end;
               fb:=AxeToFb('A'+inttostr(NumAxe)) ;
               TOBAna:=TOB.Create('ANALYTIQ',TOBAxe,-1) ;
               EcrVersAna(Tlec.detail[kk],TOBAna) ;
               TOBAna.PutValue('Y_AXE','A'+inttostr(NumAxe));
               if T <> nil then
               begin
                     Section := T.GetValue('V_SECTION');
                     Pourcentage := T.GetValue('V_TAUXMONTANT');
                     NumVentil := T.GetValue('V_NUMEROVENTIL');
               end
               else
               begin
                     Section := GetInfoCpta(fb).Attente;
                     NumVentil:=1 ;
                     Pourcentage:=100.0 ;
               end;
               inc(NbEcrAna);
               VentilLigneTOB(TOBAna,Section,NumVentil,V_PGI.OkDecV,Pourcentage,Tlec.detail[kk].GetValue ('E_DEBIT')<>0) ;
         end;
begin
     TlecNew := nil;
     DecoupageEcriture;
     for i := 0 to TEcrParJournal.detail.Count - 1 do
     begin
          TLecPeriode := TEcrParJournal.detail[i];
          for ii := 0 to TLecPeriode.detail.Count - 1 do
          begin
               TLecpiece := TLecPeriode.detail[ii];

               // si pas de doublon mais intégration d'autres pièces
               if (BIntegre) or ((not AutoriseEcheck) and (pos('IGE;IAU', InterdireLaCreation) <> 0)) then
                  SuppressionDesdoublons(TLecpiece, TlecNew, TLecPeriode.GetValue('J_JOURNAL'), ii);
               if (BEquilibre  and BIntegre) or  (BEquilibre  and BRejet)then
                  SuppressionDesSoldesAZero(TLecpiece, TlecNew, TLecPeriode.GetValue('J_JOURNAL'), ii);

               // génération d'écart de change
               GenereEcartDechange (TLecpiece, TlecEcc);

               for iii := 0 to TLecpiece.detail.Count - 1 do
               begin
                    TLec := TLecpiece.detail[iii];
                    if Tlec.detail[0].Nomtable = 'ODA' then exit;
                    // on n'utilise pas la fonction d'équilibrage dans
                    // le cas des écritures d'anouveau
                    if (TLec.detail[0].GetValue('E_ECRANOUVEAU') <> 'OAN') and
                    (TLec.detail[0].GetValue('E_ECRANOUVEAU') <> 'H') // fiche 10361
                    and ((TLec.getdouble('SOLDE') <> 0) or (TLec.Getdouble ('SOLDEDEV') <> 0))then
                    begin
                        if BEquilibre then // Fiche 10473
                        begin
                                OnAfficheListeCom('Fichier déséquilibré en montant :' ,LISTEIMPORT);
                                OnAfficheListeCom( FormatDateTime (Traduitdateformat('dd/mm/yyyy'),Tlec.detail[0].GetValue ('E_DATECOMPTABLE'))
                                        + ' Journal : ' + Tlec.detail[0].GetValue ('E_JOURNAL')
                                        + ' Référence : ' + Tlec.detail[0].GetValue ('E_REFINTERNE')
                                        + ' Libellé :' + Tlec.detail[0].GetValue ('E_LIBELLE') ,LISTEIMPORT);
                                OKexercice := FALSE; OkCreatCompte := FALSE;
                        end;
                       CEquilibrePiece(TLec);
                    end;
                    if BCtrP then
                    begin
                      if vInfo = nil then
                         vInfo := TInfoEcriture.create;
                       CAffectCompteContrePartie (TLec, vInfo);
                    end;
                    if GerBAP then   // gestion des visa
                      GestiondesBAP (Tlec);
                    for kk:=0 to TLec.Detail.Count-1 do
                    begin
                       //AJOUT ME
                       if (OrigineFichier = 'SI') and (CStat = 'ANA')  then
                            break;

                               // mise à jour de docid sur toute la piece
                       if (Tlec.detail[kk].GetValue('E_MODESAISIE') = '-')
                       and (TLec.GetValue('E_DOCIDGED') <> '') then
                                AffectDocGuid (Tlec.detail[kk], TLec.GetValue('E_DOCIDGED')) ;

                                  // ajout ventilationa
                        if OKVentilable (Tlec.detail[kk].GetValue ('E_GENERAL')) then
                        begin
                               Tlec.detail[kk].PutValue('E_ANA', 'X');
                               // fiche 10525
                               if BCtrP then
                               CSynchroVentil (TLec.detail[kk]);

                               TOBAxe := TLec.Detail[kk];
                               if TOBAXE.Detail.Count = 0 then
                               begin
                                    for NumAxe:=1 to 5 do if VentAttent[NumAxe] then
                                    begin
                                         ChargeVentil (Tlec.detail[kk].GetValue ('E_GENERAL'), inttostr(NumAxe));
                                         T := TVentil.FindFirst(['V_NATURE','V_COMPTE'],['GE'+inttostr(NumAxe),Tlec.detail[kk].GetValue ('E_GENERAL')],False);
                                         if T <> nil then // SI VENTILATION TYPE EXISTE
                                         begin
                                                for tv:=0 to TVentil.Detail.Count-1 do
                                                begin
                                                    T := TVentil.Detail[tv];
                                                    if T.GetValue('V_NATURE') = ('GE'+inttostr(NumAxe)) then AffectVentil;
                                                end;
                                         end
                                         else
                                                    AffectVentil;
                                         TVentil.free;
                                    end;
                                    AttenteAnaTOB ( Tlec.detail[kk] , V_PGI.OkDecV );
                               end
                               else
                               begin
                                 // si je le dossier est en mode croisé et le fichier non croisé
                                 if GetParamSocSecur('SO_CROISAXE', False) and not ModeCroisaxe then
                                 begin
                                              TCroisaxe := TTraitementCroisaxe.Create(GetParamSocSecur('SO_VENTILA1', False), GetParamSocSecur('SO_VENTILA2', False), GetParamSocSecur('SO_VENTILA3', False), GetParamSocSecur('SO_VENTILA4', False), GetParamSocSecur('SO_VENTILA5', False), TRUE);
                                              TCroisaxe.SetTobEcr(Tlec.detail[kk]);
                                              TobResult := TCroisaxe.PassageVentAnaClassiqueCroisaxe;
                                              Tlec.detail[kk].free;
                                              TobResult.ChangeParent(Tlec, kk);
                                              TCroisaxe.free;
                                 end
                                 else   // si je le dossier est en mode non croisé et le fichier en croisé
                                 if ModeCroisaxe and not GetParamSocSecur('SO_CROISAXE', False) then
                                 begin
//                                              TCroisaxe := TTraitementCroisaxe.Create(TRUE);
                                              TCroisaxe := TTraitementCroisaxe.Create(VentAttent[1], VentAttent[2], VentAttent[3], VentAttent[4], VentAttent[5], TRUE);
                                              TCroisaxe.SetTobEcr(Tlec.detail[kk]);
                                              TobResult := TCroisaxe.PassageVentCroisaxeAnaClassique;
                                              Tlec.detail[kk].free;
                                              TobResult.ChangeParent(Tlec, kk);
                                              TCroisaxe.free;
                                 end
                                 else
                                 begin
                                              // fiche 10095  cas B si dans le fichier un seul axe et au niveau du dossier plusieurs axes
                                              for NumAxe:=1 to 5 do if VentAttent[NumAxe] then
                                              begin
                                                    Ta := Tlec.detail[kk].FindFirst(['Y_AXE'], ['A'+inttostr(NumAxe)], FALSE);
                                                    if Ta = nil then
                                                    begin
                                                          ChargeVentil (Tlec.detail[kk].GetValue ('E_GENERAL'), inttostr(NumAxe));
                                                           T := TVentil.FindFirst(['V_NATURE','V_COMPTE'],['GE'+inttostr(NumAxe),Tlec.detail[kk].GetValue ('E_GENERAL')],False);
                                                           if T <> nil then // SI VENTILATION TYPE EXISTE
                                                           begin
                                                                  for tv:=0 to TVentil.Detail.Count-1 do
                                                                  begin
                                                                      T := TVentil.Detail[tv];
                                                                      if T.GetValue('V_NATURE') = ('GE'+inttostr(NumAxe)) then AffectVentil;
                                                                  end;
                                                           end
                                                           else
                                                                      AffectVentil;
                                                            TVentil.free;
                                                            if SaveCompte <> Tlec.detail[kk].GetValue ('E_GENERAL') then
                                                            begin
                                                                 OnAfficheListeCom('Paramétrage analytique du compte '+ Tlec.detail[kk].GetValue ('E_GENERAL')+ ' différent entre le fichier et le dossier ',LISTEIMPORT);
                                                                 SaveCompte := Tlec.detail[kk].GetValue ('E_GENERAL');
                                                            end;
                                                    end;
                                               end;
                                             // section d'attente en cas de problème
                                             // fiche 10367 refonte de la fonction
                                             // reprendre 6.5
                                             if not GetParamSocSecur('SO_CROISAXE', False)  then AttenteAnaTOB ( Tlec.detail[kk] , V_PGI.OkDecV );
                                 end;  //else ModeCroisaxe
                               end;  //else TOBAXE.Detail.Count
                        end;
                    end;
                    MiseajourDetail (Tlec,TRUE,FALSE,FALSE );
               end;  //TLec.Detail.Count
          end;
     end;
     if TlecNew <> nil then   // REnumérotaion des pièces si suppression des doublons
     begin
         NumPiece := 1;
         for i := 0 to TEcrParJournal.detail.Count - 1 do
         begin
           TLecJournal := TEcrParJournal.detail[i];
           if TLecJournal.Getvalue('J_MODESAISIE') <> '-' then break;
           Q := OpenSQl ('Select SH_NUMDEPART,J_MODESAISIE,J_NATUREJAL,J_COMPTEURNORMAL,J_AXE,J_COMPTEURSIMUL from JOURNAL,souche Where J_COMPTEURNORMAL=SH_SOUCHE and J_JOURNAL="'+ TLecJournal.Getvalue('J_JOURNAL') + '"', TRUE);
           if not Q.EOF then
                   NumPiece := Q.FindField ('SH_NUMDEPART').asinteger;
           ferme (Q);
           DernMvt := TLecJournal.Getvalue('J_NUMDERNMVT')-TLecNew.detail.Count;
           TLecJournal.putValue('J_NUMDERNMVT', DernMvt);
           for ij := 0 to TLecJournal.detail.Count - 1 do
           begin
                TLecPeriode := TLecJournal.detail[ij];
                for ii := 0 to TLecPeriode.detail.Count - 1 do
                begin
                     TLecpiece := TLecPeriode.detail[ii];
                     TLecpiece.PutValue('E_NUMEROPIECE', NumPiece);
                     for iii := 0 to TLecpiece.detail.Count - 1 do
                     begin
                          TLec := TLecpiece.detail[iii];
                          TLec.PutValue('E_NUMEROPIECE', NumPiece);
                          for kk:=0 to TLec.Detail.Count-1 do
                          begin
                              TOBAxe := TLec.Detail[kk];
                              for tv:=0 to TOBAxe.Detail.Count-1 do
                                  TOBAxe.detail[tv].PutValue('Y_NUMEROPIECE', NumPiece);
                          end; //TLec
                     end; //TLecpiece
                     inc (NumPiece);
                end; //TLecPeriode
            end; // TLecJournal
           end; //TEcrParJournal
          (*******************)

          if BEquilibre  and BRejet then  // Création du fichier de rejet
          begin
              for iii := 0 to TLecNew.detail.Count - 1 do
              begin
                   TLecpiece := TLecNew.detail[iii];
                    for ij := 0 to TLecpiece.detail.Count - 1 do
                    begin
                        TLec := TLecpiece.detail[ij];
                        EcrireFicRejet(TLec);
                    end;
              end;
          end;

          TLecNew.free;  TLecNew := nil;
     end;
     if TlecEcc <> nil then // Ecart de change
     begin
           OnAfficheListeCom('Enregistrement des écritures d''écarts de change',LISTEIMPORT);
           TlecEcc.InsertDBByNivel(FALSE);
           TlecEcc.free;  TlecEcc := nil;
     end;
end;

procedure TImportCom.DelettrageSynchro (TPiece : TOB);
var
i         : integer;
lStSQL    : string;
TL        : TOB;
          procedure Delletre(anclettre : string);  // déléttrage en mémoire
          var
          sp,se        : integer;
          TEcr         : TOB;
          begin
                    if TPiece = nil then exit;
                    for se := 0 to TPiece.detail.count-1 do     //piece
                    begin
                         TEcr := TPiece.detail[se];
                         for sp := 0 to TEcr.detail.count-1 do     //piece
                              if
                              ((TEcr.detail[sp].getvalue('E_LETTRAGE') = anclettre) and
                              (TEcr.detail[sp].getvalue('E_GENERAL') = TL.Getvalue ('E_GENERAL')) and
                              (TEcr.detail[sp].getvalue('E_AUXILIAIRE') = TL.Getvalue ('E_AUXILIAIRE')) and
                              (TEcr.detail[sp].getvalue('E_DATECOMPTABLE') = TL.Getvalue ('E_DATECOMPTABLE')))
                              then
                              begin
                                 CRemplirInfoLettrage (TEcr.detail[sp]);
                                 // voir si E_PASUETREVISION mettre à zéro cas 8
                                 TEcr.detail[sp].PutValue('E_PAQUETREVISION', 0);
                              end;
                    end;
          end;

begin
       for i := 0 to TPieceCon.detail.Count - 1 do
       begin
            TL := TPieceCon.detail[i];
            AJouterListeSoldes (TL.Getvalue ('E_GENERAL'), SGeneraux);
            if TL.Getvalue ('E_AUXILIAIRE') <> '' then
                AJouterListeSoldes (TL.Getvalue ('E_AUXILIAIRE'), STiers);

            if (TL.Getvalue ('E_PAQUETREVISION') = 0) and  (TL.Getvalue ('E_LETTRAGE') = '') then continue
            else // délettrage
            begin
               lStSQL := 'update ecriture set e_datepaquetmax=e_datecomptable,e_datepaquetmin=e_datecomptable,e_lettrage="",'+
                         'e_etatlettrage="AL",e_lettragedev="-",e_couverture=0,e_couverturedev=0, '+
                         'e_reflettrage="",e_etat="0000000000", e_datemodif="'+UsTime(NowH)+'" , E_TRESOSYNCHRO = "LET", E_PAQUETREVISION=0 ';

                lStSQL := lStSQL + ' WHERE E_GENERAL="' + TL.Getvalue ('E_GENERAL') + '" AND E_AUXILIAIRE="' + TL.Getvalue ('E_AUXILIAIRE') + '" ' +
                          ' AND E_LETTRAGE="' + TL.Getvalue ('E_LETTRAGE') + '" AND E_DATEPAQUETMIN = "' + UsDateTime(TL.Getvalue ('E_DATEPAQUETMIN'))
                          + '" AND E_DATEPAQUETMAX = "' + UsDateTime(TL.Getvalue ('E_DATEPAQUETMAX'))+'"';
                ExecuteSQL(lStSQL);
                Delletre (TL.Getvalue ('E_LETTRAGE'));
            end;
       end;
end;

procedure TImportCom.InsertTiers;
begin

  if TTiers <> nil then
  begin
     // Objet de stockage en mémoire
    if assigned(fHTiers) then
    begin
          fHTiers.Clear(True);
          FreeAndNil (fHTiers);
    end;
    TTiers.Free;  TTiers := nil;
  end;
  if TTiersInsert <> nil then
  begin
        if (NatureFichier = 'BAL') and BInterdireCreationTiers then
        begin
             if TTiersInsert.detail.count > 0 then
             begin
                OnAfficheListeCom('L''option "Interdire la création des tiers" est cochée : ',LISTEIMPORT);
                OnAfficheListeCom('    Les tiers non créés sont au nombre de '+
                IntTostr(TTiersInsert.detail.count) ,LISTEIMPORT);
             end;
        end
        else
        begin
             TTiersInsert.InsertDB(nil, TRUE);
               // Objet de stockage en mémoire
             if assigned(fHTiersInsert) then
             begin
                    fHTiersInsert.Clear(True);
                    FreeAndNil (fHTiersInsert);
             end;
             TTiersInsert.Free; TTiersInsert := nil;
        end;
  end;
  if TTiersupdate <> nil then
  begin
    TTiersUpdate.SetAllModifie(True);
    TTiersupdate.UpdateDB(TRUE);
    TTiersupdate.Free;  TTiersupdate := nil;
  end;

  // Tiers CCIMPEX
  if TTiersCAU <> nil then
  begin
       TTiersCAU.InsertOrUpdateDB(TRUE);
       TTiersCAU.Free;  TTiersCAU := nil;
       if TRibCAU <> nil then   // rib ccimpex
       begin
            TRibCAU.InsertOrUpdateDB(TRUE);
            TRibCAU.Free;  TRibCAU := nil;
       end;
       if TContactCAU <> nil then  // contact ccimpex
       begin
           TContactCAU.InsertOrUpdateDB(TRUE);
           TContactCAU.Free;  TContactCAU := nil;
       end;
  end;

end;

procedure TImportCom.InsertUpdateGeneraux;
var
TA, Ti      :TOB;
i,j         : integer;
Mdp,ModPaie : string;
lQMR        : Tquery;
SectAtt     : array [1..5] of Boolean;
lAxeVent    : array [1..5] of Boolean; // FQ 21259 SBO 27/08/2007
lCpt        : TStringList;
indc        : integer;
begin
  lCpt := nil;
  remplir_compteattente;

  if TGen <> nil then
  begin
     if assigned(fHGene) then
     begin
          fHGene.Clear(True);
          FreeAndNil (fHGene);
     end;
     TGen.Free; TGen := nil;
  end;
  if TGenInsert <> nil then
  begin
         { BVE 25.07.07 Rajout Ventilation Type }
         for j := 1 to 5 do
              SectAtt[j] := true;
         if GereVentilType and ModeEVentilType then
         begin
            lCpt  := TStringList.Create ;
            for i:=0 to TGenInsert.detail.count-1 do
            begin
                    Ti := TGenInsert.detail[i];
                    InitialiseVentilDefaut(Ti.GetValue ('G_GENERAL'), SectAtt, lAxeVent ); // FQ 21259 SBO 27/08/2007
                    for j:= 1 to 5 do
                    begin
                        if lAxeVent[j] = TRUE then  // fiche 20869
                        begin
                          Ti.PutValue ('G_VENTILABLE', 'X');
                          Ti.PutValue ('G_VENTILABLE'+Inttostr(j), 'X');
                              // fiche 21789
                          if lCpt.IndexOf (Ti.GetValue ('G_GENERAL')) < 0 then
                            lCpt.Add (Ti.GetValue ('G_GENERAL'));
                        end;
                    end;
            end;
         end;
         { END BVE 25.07.07 }

    TGenInsert.InsertDB(nil, TRUE);
    if assigned(fHGeneInsert) then
    begin
          fHGeneInsert.Clear(True);
          FreeAndNil (fHGeneInsert);
    end;
    TGenInsert.Free; TGenInsert := nil;
    // fiche 21789
    if GereVentilType and ModeEVentilType then
    begin
         for indc := 0 to lCpt.Count-1 do
                        MajGVentil (FALSE, lCpt[indc]);
         if lCpt <> nil then
         begin
                  for indc :=0 to lCpt.Count-1 do If lCpt.Objects[indc]<>NIL Then lCpt.Objects[indc].Free ;
                 lCpt.free;
         end;
    end;
  end;
  if TGenupdate <> nil then
  begin
    if (OrigineFichier = 'S1') and UpdateLettreDiv then // fiche 13221
    begin
        Mdp := GetParamSocSecur('SO_GCMODEREGLEDEFAUT', '') ;
        ModPaie  := '' ;

        if Mdp <> '' then
        begin
          lQMR := OpenSQL('SELECT MR_MP1 FROM MODEREGL WHERE MR_MODEREGLE="'+Mdp+'" ',TRUE) ;
          if Not lQMR.Eof Then
            ModPaie := lQMR.FindField('MR_MP1').AsString ;
          Ferme(lQMR) ;
          if (ModPaie <> '') then
          begin
              for i := 0 to TGenupdate.detail.Count - 1 do
              begin
                  TA := TGenupdate.detail[i];
                  MajLettrageEcriture( ModPaie, TA.GetValue ('G_GENERAL'), fbGene)
              end;
          end;
        end ;
    end;
    TGenUpdate.SetAllModifie (True);
    TGenupdate.UpdateDB(TRUE);
    TGenupdate.Free;  TGenupdate := nil;
  end;
end;

procedure TImportCom.MiseajourSouche;
var
Q2                        : TQuery;
dmvt,NewNum               : integer;
Typesouche, ana           : string;
i,IndiceDern              : integer;
Datedernmvt               : TDatetime;
dernmvt,natjrl            : string;
TSou                      : TOB;
SIM                       : string;
Q                         : Tquery;
begin
  NewNum := 0;
  if TEcrParJournal <> nil then
  begin
    // maj dans le journal de la date dernier mouvement et numdernmvt
    for i := 0 to TEcrParJournal.detail.Count - 1 do
    begin
      Datedernmvt :=
        StrToDate(TEcrParJournal.detail[i].GetValue('J_DATEDERNMVT'));
      dernmvt := TEcrParJournal.detail[i].GetValue('J_NUMDERNMVT');
      if dernmvt <> '' then  // calcul des souches
      begin
           Typesouche := '';
           ana := '-';  // dmvt := 1;
           Typesouche := TEcrParJournal.detail[i].GetValue ('J_COMPTEURNORMAL');
           modesaisie := TEcrParJournal.detail[i].GetValue ('J_MODESAISIE');
           natjrl     := TEcrParJournal.detail[i].GetValue ('J_NATUREJAL');
           SIM        := TEcrParJournal.detail[i].GetValue('J_COMPTEURSIMUL');
           // if modesaisie <> '-' then dmvt := 0;
           if natjrl = 'ODA' then ana := 'X';
           if Typesouche = '' then Typesouche := TEcrParJournal.detail[i].GetValue('J_JOURNAL');

           dmvt := StrToint (dernmvt); // pour enlever les 000049
           if (Typesouche <> '') then
           begin
               if (not ExisteSQl ('SELECT SH_SOUCHE FROM SOUCHE Where SH_TYPE="CPT" AND SH_SOUCHE="'+ Typesouche + '"')) then
               begin
                         TSou := TOB.Create('SOUCHE', nil, -1);
                         TSou.Putvalue ('SH_TYPE', 'CPT');
                         TSou.Putvalue ('SH_SOUCHE', Typesouche);
                         TSou.Putvalue ('SH_LIBELLE', 'SOUCHE :'+ Typesouche);
                         TSou.Putvalue ('SH_NUMDEPART', dmvt+1);
                         if SIM <> '' then // fiche 10197
                            TSou.Putvalue ('SH_SIMULATION', 'X')
                         else
                            TSou.Putvalue ('SH_SIMULATION', '-');
                         TSou.Putvalue ('SH_ANALYTIQUE', ana);
                         TSou.InsertDB(nil, TRUE);
                         TSou.Free;
               end
               else
               begin
                    if (Modesaisie = '-') then
                    begin
                      Q2 := OpenSQl('SELECT SH_NUMDEPART FROM SOUCHE Where SH_TYPE="CPT" AND SH_SOUCHE="'+ Typesouche + '"',TRUE);
                      if not Q2.EOF then
                         NewNum := Q2.FindField ('SH_NUMDEPART').asinteger-1;
                      if NewNum > dmvt then dmvt := NewNum;
                      ferme (Q2);
                      ExecuteSQL('UPDATE SOUCHE SET SH_NUMDEPART=' + IntToStr(dmvt+1) +' WHERE SH_TYPE="CPT" AND SH_SOUCHE="' + Typesouche + '"');
                    end;
               end;
             end;
           if (SIM = '') then // fiche 10197
           begin
              if (NatureFichier = 'SYN') then  // fiche 10073 pour ne pas écraser J_NUMDERNMVT à chaque synchro
              begin
                    IndiceDern := 0;
                    Q := OpenSql ('SELECT J_DATEDERNMVT, J_NUMDERNMVT from journal '
                    + ' WHERE J_JOURNAL="' + TEcrParJournal.detail[i].GetValue('J_JOURNAL') + '"', TRUE);
                    if not Q.EOF then
                    IndiceDern := Q.FindField ('J_NUMDERNMVT').asinteger;
                    Ferme(Q);
                    if IndiceDern < dmvt then
                    begin
                              ExecuteSQL('UPDATE JOURNAL SET J_DATEDERNMVT="' + USDATETIME(Datedernmvt)
                              +'",J_NUMDERNMVT=' + inttostr(dmvt) +',J_COMPTEURNORMAL="'+Typesouche+'" '
                              + ' WHERE J_JOURNAL="' + TEcrParJournal.detail[i].GetValue('J_JOURNAL') + '"');
                    end;
              end
              else
                    ExecuteSQL('UPDATE JOURNAL SET J_DATEDERNMVT="' + USDATETIME(Datedernmvt)
                    +'",J_NUMDERNMVT=' + inttostr(dmvt) +',J_COMPTEURNORMAL="'+Typesouche+'" '
                    + ' WHERE J_JOURNAL="' + TEcrParJournal.detail[i].GetValue('J_JOURNAL') + '"');
           end;
      end;
    end;
  end;
end;



procedure TImportCom.MajExoV8;
var Qa  : Tquery;
EXERC   : string;
TEX     : TOB;
d,ii    : integer;
TheTob  : TOB;
begin
  EXERC := '';
  if OkMajExoV8 then
  begin
     TheTob := TOB.Create ('', nil,-1);
     TheTob.LoadDetailDBFromSql('JOURNAL', 'SELECT J_JOURNAL FROM JOURNAL Where J_NATUREJAL="ANO"');
     for ii := 0 to TheTob.detail.count-1 do
     begin
         Qa := OpenSql ('select e_exercice from ecriture where e_journal="'+TheTob.detail[ii].GetValue('J_JOURNAL') +'" and e_ecranouveau="H" ORDER BY E_DATECOMPTABLE DESC ',TRUE);
         if not Qa.EOF then Exerc := Qa.Fields[0].asstring;
         ferme(Qa);
         if (EXERC <> '') then
          SetParamsoc ('SO_EXOV8', EXERC);
     end;
     Thetob.free;
  end;

  // pour les dossiers TRA venant de Winner tenue ou surveillance
  if (NatureFichier = 'DOS') then
  begin
      EXERC := '';   // fiche 21995
      Qa := OpenSQl ('select ex_exercice from exercice where ex_etatcpta="OUV" order by ex_datedebut asc', TRUE);
      if not Qa.EOF then
        EXERC := Qa.Fields[0].asstring;
      ferme(Qa);
      if (EXERC <> '') and ((OrigineFichier = 'WT') or (OrigineFichier='WS')) then SetParamsoc ('SO_EXOV8', EXERC);
      if (EXERC <> '') and (GetParamsocSecur ('SO_CPEXOREF', '') = '') then SetParamsoc ('SO_CPEXOREF', EXERC); // fiche 10526
      if (GetParamsocSecur ('SO_LETMODE', '') = '') then
      begin
            if (ctxPCL in V_PGI.PGIContexte) then SetParamsoc ('SO_LETMODE', 'PCL')
            else SetParamsoc ('SO_LETMODE', 'PME');
      end;
      if (GetParamsocSecur ('SO_LETCHOIXDEFVALID', '') = '') then
            SetParamsoc ('SO_LETCHOIXDEFVALID', 'AL1')

  end;

  if MAJCPEXOREF then // en synchro mise à jour de exercice de référence
  begin
      EXERC := '';
      Qa := OpenSQl ('select ex_exercice from exercice where ex_etatcpta="OUV"', TRUE);
      if not Qa.EOF then
        EXERC := Qa.Fields[0].asstring;
      ferme(Qa);
      if EXERC <> '' then SetParamsoc ('SO_CPEXOREF', EXERC);
  end;
  d := NombreExercice('WHERE EX_ETATCPTA="OUV"');
  if (d > 2) then
  begin
       Qa := OpenSql ('select * from exercice where ex_etatcpta="OUV" ', TRUE);
       TEX := TOB.Create('', nil, -1);
       TEX.LoadDetailDB('EXERCICE', '', '', Qa, TRUE, FALSE);
       Ferme(Qa);
       for ii := TEX.detail.Count -1 downto 0 do
       begin
          TEX.detail[ii].PutValue ('EX_ETATCPTA', 'NON');
          dec(d);
          if d = 2 then break;
       end;
       TEX.SetAllModifie(True);
       TEX.UpdateDB(TRUE);
       TEX.free;
  end;

end;

procedure TImportCom.ImportParame1(Ligne: string);
begin

  if (not Paramgene) then  exit;

  SetParamsoc('SO_LIBELLE', VireTouteSub(EnleveBlanc(Ligne, 7, 35,''),'"'));
  SetParamsoc('SO_ADRESSE1', EnleveBlanc(Ligne, 42, 35,''));
  SetParamsoc('SO_ADRESSE2', EnleveBlanc(Ligne, 77, 35,''));
  SetParamsoc('SO_ADRESSE3', EnleveBlanc(Ligne, 112, 35,''));
  SetParamsoc('SO_CODEPOSTAL', EnleveBlanc(Ligne, 147, 9,''));
  SetParamsoc('SO_VILLE', EnleveBlanc(Ligne, 156, 35,''));
  SetParamsoc('SO_PAYS', EnleveBlanc(Ligne, 191, 3,''));
  SetParamsoc('SO_TELEPHONE', EnleveBlanc(Ligne, 194, 25,''));
  SetParamsoc('SO_FAX', EnleveBlanc(Ligne, 219, 25,''));
  SetParamsoc('SO_TELEX', EnleveBlanc(Ligne, 244, 25,''));
  SetParamsoc('SO_MAIL', EnleveBlanc(Ligne, 269, 35,''));
  SetParamsoc('SO_RVA', EnleveBlanc(Ligne, 304, 35,''));
  SetParamsoc('SO_CONTACT', EnleveBlanc(Ligne, 339, 35,''));
  SetParamsoc('SO_NIF', EnleveBlanc(Ligne, 374, 17,''));
  SetParamsoc('SO_SIRET', EnleveBlanc(Ligne, 391, 35,''));
  SetParamsoc('SO_RC', EnleveBlanc(Ligne, 426, 17,''));
  SetParamsoc('SO_APE', EnleveBlanc(Ligne, 443, 35,''));
  SetParamsoc('SO_CAPITAL', EnleveBlanc(Ligne, 478, 20,'0'));
  SetParamsoc('SO_TXTJURIDIQUE', EnleveBlanc(Ligne, 498, 35,''));
  OnAfficheListeCom(EnleveBlanc(Ligne, 1, 6,'') + ' : '+ EnleveBlanc(Ligne, 7, 35,''),LISTEIMPORT);

end;

procedure TImportCom.ImportParame2(Ligne: string);
var
   lg1         : integer;
   lg2         : integer;
   Bourre      : string;
   TA          : TOB;
   pointj      : Boolean;
   POINTX      : string;
   procedure InsertAxe (Axe, Section, Br : string; Lgr : integer; LeType : TFichierBase);
   begin
          if Section = '' then exit; // fiche 10514
          TA := TOB.Create('AXE', nil, -1);
          TA.putvalue ('X_AXE', Axe);
          TA.putvalue ('X_LONGSECTION', Lgr);

          if (NatureFichier = 'DOS') and (GetInfoCpta(fbaxe1).Lg <> Lgr) then
            TA.putvalue ('X_SECTIONATTENTE', BourreEtLess(Section,LeType, Lgr))
          else
            TA.putvalue ('X_SECTIONATTENTE', BourreOuTronque(Section,LeType));
          TA.putvalue ('X_BOURREANA', Br);
          TA.InsertOrUpdateDB(TRUE);
          TA.Free;
 end;
begin
  // ajout ME pour la gestion du pointage
  GestPointage := EnleveBlanc(Ligne, 198, 3,'');
  POINTX := GetParamSocSecur ('SO_CPPOINTAGESX', 'EXP');
  if (stArg = '') and (NatureFichier = 'SYN')  and (OrigineFichier <> 'S1') and
  (((GestPointage = 'CLI') and (POINTX <> 'CLI')) or
  ((GestPointage = 'EXP') and (POINTX <> 'EXP'))) then
  begin
       AffMessage ('L''import est impossible.#10#13 Les paramètres de gestion du pointage sont incohérents,'+
       ' veuillez les mettre en conformité.');
       OKexercice := FALSE;
       exit;
  end;
  if not (ctxPCL in V_PGI.PGIContexte) then
  begin
    if (OrigineEnvoie = 'EXP') and (NatureFichier <> 'DOS') and (GestPointage = 'CLI') then IgnorePointage := TRUE;
    if not ExisteSQL('SELECT DOS_NODOSSIER FROM DOSSIER') then
    begin
         if  (GestPointage = 'CLI') or (GestPointage = 'EXP') then
             SetParamsoc ('SO_CPPOINTAGESX', GestPointage)
         else
             SetParamsoc ('SO_CPPOINTAGESX', 'CLI');
    end
    else
    begin
         if  (GestPointage = 'CLI') or (GestPointage = 'EXP') then
             SetParamsoc ('SO_CPPOINTAGESX', GestPointage);
    end;
  end
  else   // cas expert
  begin
         if (OrigineEnvoie = 'CLI') and (NatureFichier <> 'DOS') and (GetParamsoc ('SO_CPPOINTAGESX') = 'EXP') then IgnorePointage := TRUE ;
         // fiche 10405
         if (OrigineEnvoie = 'CLI') and (NatureFichier = 'DOS') and (GestPointage <> '') then
             SetParamsoc ('SO_CPPOINTAGESX', GestPointage);
  end;
  if (not Paramgene) then  exit;

  lg1 :=  StrToInt(EnleveBlanc(Ligne, 7, 2,'10'));
  lg2 :=  StrToInt(EnleveBlanc(Ligne, 10, 2,'10'));
  if (fbGeneLg <> lg1) and (fbGeneLg <> 0)  then
  begin
       if  fbGeneLg > lg1 then lg1 := fbGeneLg
       else
       begin
            if fbGeneLg < lg1 then
            begin
               fbGeneLg := lg1;
               Changementgene := TRUE;
            end;
       end;
  end;
  if (fbAuxLg <> lg2) and (fbAuxLg <> 0) then
  begin
       if fbAuxLg > lg2 then lg2 :=  fbAuxLg
       else
       begin
            if fbAuxLg < lg2 then
            begin
               fbAuxLg := lg2;
               Changementaux := TRUE;
            end;
       end;
  end;

  SetParamsoc('SO_LGCPTEGEN', lg1);
  fbGeneLg := lg1 ;

  SetParamsoc('SO_BOURREGEN', EnleveBlanc(Ligne, 9, 1,''));
  if EnleveBlanc(Ligne, 9, 1,'') = '' then
     SetParamsoc('SO_BOURREGEN', '0');

  Bourre := GetParamSocSecur('SO_BOURREGEN', '0');
{$IFDEF EAGLSERVER}
  fbGeneCb:= Bourre[1];
{$ELSE}
  VH^.Cpta[fbGene].Cb:= Bourre[1];
  VH^.Cpta[fbGene].lg := lg1;
{$ENDIF}

  SetParamsoc('SO_LGCPTEAUX', lg2);
  fbAuxLg := lg2 ;


  SetParamsoc('SO_BOURREAUX', EnleveBlanc(Ligne, 12, 1,''));
  if EnleveBlanc(Ligne, 12, 1,'') = '' then
     SetParamsoc('SO_BOURREAUX', '0');

  Bourre := GetParamSocSecur('SO_BOURREAUX', '0');

{$IFDEF EAGLSERVER}
  fbAuxcb := Bourre[1];
{$ELSE}
  VH^.Cpta[fbAux].Cb := Bourre[1];
  VH^.Cpta[fbAux].lg := lg2;
{$ENDIF}

  SetParamsoc('SO_LGMAXBUDGET', EnleveBlanc(Ligne, 7, 2,'10'));


  InsertAxe ('A1', EnleveBlanc(Ligne, 113, 17,''), EnleveBlanc(Ligne, 15, 1,'0'), StrToInt(EnleveBlanc(Ligne, 13, 2,'0')), fbaxe1);
  InsertAxe ('A2', EnleveBlanc(Ligne, 130, 17,''), EnleveBlanc(Ligne, 18, 1,'0'), StrToInt(EnleveBlanc(Ligne, 16, 2,'0')), fbaxe2);
  InsertAxe ('A3', EnleveBlanc(Ligne, 147, 17,''), EnleveBlanc(Ligne, 21, 1,'0'), StrToInt(EnleveBlanc(Ligne, 19, 2,'0')), fbaxe3);
  InsertAxe ('A4', EnleveBlanc(Ligne, 164, 17,''), EnleveBlanc(Ligne, 24, 1,'0'), StrToInt(EnleveBlanc(Ligne, 22, 2,'0')), fbaxe4);
  InsertAxe ('A5', EnleveBlanc(Ligne, 181, 17,''), EnleveBlanc(Ligne, 27, 1,'0'), StrToInt(EnleveBlanc(Ligne, 25, 2,'0')), fbaxe5);

  if EnleveBlanc(Ligne, 28, 17,'') <> '' then
     SetParamsoc('SO_GENATTEND', BourreOuTronque(EnleveBlanc(Ligne, 28, 17,''),fbgene));
  if EnleveBlanc(Ligne, 45, 17,'') <> '' then
     SetParamsoc('SO_CLIATTEND', UpperCase(BourreOuTronque(EnleveBlanc(Ligne, 45, 17,''),fbAux)));
  if EnleveBlanc(Ligne, 62, 17,'') <> '' then
     SetParamsoc('SO_FOUATTEND', UpperCase(BourreOuTronque(EnleveBlanc(Ligne, 62, 17,''),fbAux)));
  if EnleveBlanc(Ligne, 79, 17,'') <> '' then
     SetParamsoc('SO_SALATTEND', BourreOuTronque(EnleveBlanc(Ligne, 79, 17,''),fbAux));
  if EnleveBlanc(Ligne, 96, 17,'') <> '' then
     SetParamsoc('SO_DIVATTEND', BourreOuTronque(EnleveBlanc(Ligne, 96, 17,''),fbAux));

  pointj := (EnleveBlanc(Ligne, 201, 1,'') = 'X');
  SetParamsoc('SO_POINTAGEJAL', pointj);
  pointj := (EnleveBlanc(Ligne, 202, 1,'') = 'X');
  SetParamsoc('SO_CPPOINTAGESECU', pointj);
  pointj := (EnleveBlanc(Ligne, 203, 1,'') = 'X');
  SetParamsoc('SO_CPPCLSAISIEQTE', pointj);
{$IFDEF EAGLSERVER}     //chargement quand le dossier existe et lg différent du fichier
     TCPContexte.GetCurrent.InfoCpta.ChargeLgDossier ;
{$ENDIF}

  OnAfficheListeCom(EnleveBlanc(Ligne, 1, 6,''),LISTEIMPORT);
end;

procedure TImportCom.ImportParame3(Ligne: string);
begin
  if (not Paramgene) then  exit;

  if EnleveBlanc(Ligne, 7, 17,'') <> '' then
  begin
     SetParamsoc('SO_OUVREBIL', BourreOuTronque(EnleveBlanc(Ligne, 7, 17,''),fbgene));
     if GetParamSocSecur ('SO_FERMEBIL', '') = '' then
        SetParamsoc('SO_FERMEBIL', BourreOuTronque(EnleveBlanc(Ligne, 7, 17,''),fbgene));
  end;
  if EnleveBlanc(Ligne, 24, 17,'') <> '' then
     SetParamsoc('SO_RESULTAT', BourreOuTronque(EnleveBlanc(Ligne, 24, 17,''),fbgene));
  if EnleveBlanc(Ligne, 41, 17,'') <> '' then
  begin
     SetParamsoc('SO_OUVREBEN', BourreOuTronque(EnleveBlanc(Ligne, 41, 17,''),fbgene));
     if GetParamSocSecur ('SO_FERMEBEN', '') = '' then
        SetParamsoc('SO_FERMEBEN', BourreOuTronque(EnleveBlanc(Ligne, 41, 17,''),fbgene));
  end;
  if EnleveBlanc(Ligne, 58, 17,'') <> '' then
  begin
     SetParamsoc('SO_OUVREPERTE', BourreOuTronque(EnleveBlanc(Ligne, 58, 17,''),fbgene));
     if GetParamSocSecur ('SO_FERMEPERTE', '') = '' then
        SetParamsoc('SO_FERMEPERTE', BourreOuTronque(EnleveBlanc(Ligne, 58, 17,''),fbgene));
  end;
  if EnleveBlanc(Ligne, 75, 3,'') <> '' then
     SetParamsoc('SO_JALOUVRE', EnleveBlanc(Ligne, 75, 3,''));
  if EnleveBlanc(Ligne, 78, 3,'') <> '' then
     SetParamsoc('SO_JALFERME', EnleveBlanc(Ligne, 78, 3,''));
  if EnleveBlanc(Ligne, 81, 3,'') <> '' then
     SetParamsoc('SO_JALREPBALAN', EnleveBlanc(Ligne, 81, 3,''));
  if EnleveBlanc(Ligne, 84, 17,'') <> '' then
     SetParamsoc('SO_DEFCOLCLI', BourreOuTronque(EnleveBlanc(Ligne, 84, 17,''),fbgene));
  if EnleveBlanc(Ligne, 101, 17,'') <> '' then
     SetParamsoc('SO_DEFCOLFOU', BourreOuTronque(EnleveBlanc(Ligne, 101, 17,''),fbgene));
  if EnleveBlanc(Ligne, 118, 17,'') <> '' then
     SetParamsoc('SO_DEFCOLSAL', BourreOuTronque(EnleveBlanc(Ligne, 118, 17,''),fbgene));
  if EnleveBlanc(Ligne, 135, 17,'') <> '' then
     SetParamsoc('SO_DEFCOLDDIV', BourreOuTronque(EnleveBlanc(Ligne, 135, 17,''),fbgene));
  if EnleveBlanc(Ligne, 152, 17,'') <> '' then
     SetParamsoc('SO_DEFCOLCDIV', BourreOuTronque(EnleveBlanc(Ligne, 152, 17,''),fbgene));
  if EnleveBlanc(Ligne, 169, 17,'') <> '' then
     SetParamsoc('SO_DEFCOLDIV', BourreOuTronque(EnleveBlanc(Ligne, 169, 17,''),fbgene));
  OnAfficheListeCom(EnleveBlanc(Ligne, 1, 6,''),LISTEIMPORT);
end;

procedure TImportCom.ImportParame4(Ligne: string);
begin
  if (not Paramgene) then  exit;

  SetParamsoc('SO_BILDEB1', EnleveBlanc(Ligne, 7, 17,''));
  SetParamsoc('SO_BILFIN1', EnleveBlanc(Ligne, 24, 17,''));
  SetParamsoc('SO_BILDEB2', EnleveBlanc(Ligne, 41, 17,''));
  SetParamsoc('SO_BILFIN2', EnleveBlanc(Ligne, 58, 17,''));
  SetParamsoc('SO_BILDEB3', EnleveBlanc(Ligne, 75, 17,''));
  SetParamsoc('SO_BILFIN3', EnleveBlanc(Ligne, 92, 17,''));
  SetParamsoc('SO_BILDEB4', EnleveBlanc(Ligne, 109, 17,''));
  SetParamsoc('SO_BILFIN4', EnleveBlanc(Ligne, 126, 17,''));
  SetParamsoc('SO_BILDEB5', EnleveBlanc(Ligne, 143, 17,''));
  SetParamsoc('SO_BILFIN5', EnleveBlanc(Ligne, 160, 17,''));
  SetParamsoc('SO_CHADEB1', EnleveBlanc(Ligne, 177, 17,''));
  SetParamsoc('SO_CHAFIN1', EnleveBlanc(Ligne, 194, 17,''));
  SetParamsoc('SO_CHADEB2', EnleveBlanc(Ligne, 211, 17,''));
  SetParamsoc('SO_CHAFIN2', EnleveBlanc(Ligne, 228, 17,''));
  SetParamsoc('SO_CHADEB3', EnleveBlanc(Ligne, 245, 17,''));
  SetParamsoc('SO_CHAFIN3', EnleveBlanc(Ligne, 262, 17,''));
  SetParamsoc('SO_CHADEB4', EnleveBlanc(Ligne, 279, 17,''));
  SetParamsoc('SO_CHAFIN4', EnleveBlanc(Ligne, 296, 17,''));
  SetParamsoc('SO_CHADEB5', EnleveBlanc(Ligne, 313, 17,''));
  SetParamsoc('SO_CHAFIN5', EnleveBlanc(Ligne, 330, 17,''));
  SetParamsoc('SO_PRODEB1', EnleveBlanc(Ligne, 347, 17,''));
  SetParamsoc('SO_PROFIN1', EnleveBlanc(Ligne, 364, 17,''));
  SetParamsoc('SO_PRODEB2', EnleveBlanc(Ligne, 381, 17,''));
  SetParamsoc('SO_PROFIN2', EnleveBlanc(Ligne, 398, 17,''));
  SetParamsoc('SO_PRODEB3', EnleveBlanc(Ligne, 415, 17,''));
  SetParamsoc('SO_PROFIN3', EnleveBlanc(Ligne, 432, 17,''));
  SetParamsoc('SO_PRODEB4', EnleveBlanc(Ligne, 449, 17,''));
  SetParamsoc('SO_PROFIN4', EnleveBlanc(Ligne, 466, 17,''));
  SetParamsoc('SO_PRODEB5', EnleveBlanc(Ligne, 483, 17,''));
  SetParamsoc('SO_PROFIN5', EnleveBlanc(Ligne, 500, 17,''));
  OnAfficheListeCom(EnleveBlanc(Ligne, 1, 6,''),LISTEIMPORT);

end;

procedure TImportCom.ImportParame5(Ligne: string);
var
tauxeuro        : double;
monn,Planref    : string;
Oknum           : Boolean;
j               : integer;
GeneAna         : Boolean;
AnvDynamique    : Boolean;
begin
  if (EnleveBlanc(Ligne, 77, 1,'') <> '')then  // fiche
  begin
       if EnleveBlanc(Ligne, 77, 1,'') = 'X' then ModeCroisaxe := TRUE
       else ModeCroisaxe := FALSE;
  end;
  if (not Paramgene) then  exit;

  SetParamsoc('SO_DEVISEPRINC', EnleveBlanc(Ligne, 7, 3,''));
  SetParamsoc('SO_DECVALEUR', EnleveBlanc(Ligne, 10, 1,'2'));
  SetParamsoc('SO_TENUEEURO', EnleveBlanc(Ligne, 11, 1,''));

// voir si blanc
  if (EnleveBlanc(Ligne, 7, 3,'')='FRF') and (EnleveBlanc(Ligne, 11, 1,'') = '') then
     SetParamsoc('SO_TENUEEURO', FALSE);

  TenueEuro  := GetParamSocSecur ('SO_TENUEEURO', TRUE);
  if (Paramgene) and (OrigineFichier = 'SI') then
  begin
    if (EnleveBlanc(Ligne, 7, 3,'') = 'EUR') or (GetParamSocSecur('SO_TENUEEURO', TRUE) = FALSE) then
    begin                             // fiche 10577
        ExecuteSQL('UPDATE DEVISE SET D_DECIMALE='+ EnleveBlanc(Ligne, 10, 1,'2') +',D_DEVISE="' + GetParamSocSecur('SO_DEVISEPRINC', 'EUR') +
          '",D_FERME="-"'+
          ' WHERE D_DEVISE="' + GetParamSocSecur('SO_DEVISEPRINC', 'EUR') + '"');
          monn := 'FRF';
    end
    else
    monn := GetParamSocSecur('SO_DEVISEPRINC', 'EUR');
    if ExisteSQl ('SELECT * from DEVISE Where D_DEVISE="'+monn+'"') then
    begin
        ExecuteSQL('UPDATE DEVISE SET D_DEVISE="' + monn +
          '",D_FONGIBLE="X",D_MONNAIEIN="X" ' +
          ' WHERE D_DEVISE="' + monn + '"');
    end;
  end;

  tauxeuro :=  Valeur(StPoint((EnleveBlanc(Ligne, 49, 10, '1'))));
  SetParamsoc('SO_TAUXEURO',tauxeuro);
  if EnleveBlanc(Ligne, 49, 10,'') = '' then
    SetParamSoc ('SO_TAUXEURO', 6.55957);
  SetParamsoc('SO_REGLEEQUILSAIS', 'CPT');
  if EnleveBlanc(Ligne, 59, 3,'') <> '' then
     SetParamsoc('SO_REGIMEDEFAUT', EnleveBlanc(Ligne, 59, 3,''));
  if EnleveBlanc(Ligne, 62, 3,'') <> '' then
     SetParamsoc('SO_GCMODEREGLEDEFAUT', EnleveBlanc(Ligne, 62, 3,''));
  if EnleveBlanc(Ligne, 65, 3,'') <> '' then
     SetParamsoc('SO_CODETVADEFAUT', EnleveBlanc(Ligne, 65, 3,''));
  if EnleveBlanc(Ligne, 68, 3,'') <> '' then
     SetParamsoc('SO_CODETVAGENEDEFAULT', EnleveBlanc(Ligne, 68, 3,''));
  SetParamsoc('SO_ETABLISDEFAUT', EnleveBlanc(Ligne, 71, 3,''));
  planref :=  EnleveBlanc(Ligne, 74, 2,'0');
  Oknum := TRUE;
  for j:=1 To  length(planref) do
      if planref[j] in Alpha then begin Oknum := FALSE; break; end;
  if Oknum then
     SetParamsoc('SO_NUMPLANREF', StrToint(EnleveBlanc(Ligne, 74, 2,'0')));
  if (EnleveBlanc(Ligne, 74, 2,'') = '') or (not Oknum) then
     SetParamsoc('SO_NUMPLANREF', 7);
  // Fiche 10582
  if (OrigineFichier = 'S1') and (OrigineEnvoie = 'CLI') and ((NatureFichier = 'SYN') or (NatureFichier = 'DOS')) then
      SetParamsoc('SO_ETABLISCPTA', '-')
  else
      SetParamsoc('SO_ETABLISCPTA', 'X');
  SetParamsoc('SO_NBJECRAVANT', '999');
  SetParamsoc('SO_NBJECRAPRES', '999');
  SetParamsoc('SO_NBJECHAVANT', '999');
  SetParamsoc('SO_NBJECHAPRES', '999');
  SetParamsoc ('SO_ZACTIVEPFU', TRUE);
  OnAfficheListeCom(EnleveBlanc(Ligne, 1, 6,'') + ' FIN paramètres généraux',LISTEIMPORT);
  if (EnleveBlanc(Ligne, 76, 1,'') <> '') then
  begin
     GeneAna := (EnleveBlanc(Ligne, 76, 1,'')='X');
     SetParamsoc('SO_ZGEREANAL', GeneAna);
  end;
  if (EnleveBlanc(Ligne, 77, 1,'') <> '')then  // fiche
  begin
       if EnleveBlanc(Ligne, 77, 1,'') = 'X' then ModeCroisaxe := TRUE
       else ModeCroisaxe := FALSE;
       SetParamsoc('SO_CROISAXE', ModeCroisaxe);
  end;
  if (EnleveBlanc(Ligne, 78, 1,'') <> '')then  // fiche 10393
  begin
       AnvDynamique := (EnleveBlanc(Ligne, 78, 1,'') = 'X');
       SetParamsoc('SO_CPANODYNA', AnvDynamique);
  end;

{$IFNDEF EAGLSERVER}
  ChargeMagHalley ;
{$ENDIF}
end;

procedure TImportCom.ImportExercice(Ligne: string);
var
  D1, D2,Ddeb       : TDateTime;
  sLibelleExo       : string;
  St,Etatcomptable  : string;
  cdd               : string;
  TA                : TOB;
  TEX               : TOB;
  Q1                : TQuery;
  Exerc             : TExoDate;
  MAJEXO            : Boolean;
begin
  St                := EnleveBlanc(Ligne, 10, 8,'');
  D1                := Format_Date_HAL(St);
  St                := EnleveBlanc(Ligne, 18, 8,'');
  D2                := Format_Date_HAL(St);
  MAJEXO            := FALSE;

  if (OrigineFichier = 'SI') then
  begin
{$IFNDEF EAGLSERVER}
       VH^.Encours.Deb := D1;
       VH^.Encours.Code := EnleveBlanc(Ligne, 7, 3,'');
{$ENDIF}
  end;
  DDeb := 0;
  sLibelleExo := VireTouteSub(EnleveBlanc(Ligne, 32, 35,''),'"');
  Etatcomptable := EnleveBlanc(Ligne, 26, 3,'');
  if Etatcomptable = 'CLO' then Etatcomptable := 'CDE';
  if TOBAno = nil then
       TOBAno := TOB.Create('', nil, -1);
  TA := TOB.Create ('',TOBAno,-1);
  CQuelExercice (D1, Exerc);
  if (Exerc.Code <> '') and (Exerc.Code <> EnleveBlanc(Ligne, 7, 3,'')) then
     TA.AddChampSupValeur('EXERCICE', Exerc.code)
  else
     TA.AddChampSupValeur('EXERCICE', EnleveBlanc(Ligne, 7, 3,''));
  TA.AddChampSupValeur('ANO', EnleveBlanc(Ligne, 67, 3,''));
  TA.AddChampSupValeur('JOURNAL', '');

 // si dossier complet on supprime l'exercice.
 // teste de existesql sur la date plutot que le code
 TEX := TOB.Create('EXERCICE', nil, -1);
 TEX.PutValue ('EX_EXERCICE', EnleveBlanc(Ligne, 7, 3,''));
 TEX.PutValue ('EX_LIBELLE', sLibelleExo);
 TEX.PutValue ('EX_ABREGE', EnleveBlanc(sLibelleExo, 1, 14,''));
 TEX.PutValue ('EX_DATEDEBUT', D1);
 TEX.PutValue ('EX_DATEFIN', D2);
 TEX.PutValue ('EX_ETATCPTA', Etatcomptable);
 TEX.PutValue ('EX_ETATBUDGET', EnleveBlanc(Ligne, 29, 3,''));
 TEX.PutValue ('EX_ETATADV', 'NON');
 TEX.PutValue ('EX_ETATAPPRO', 'NON');
 TEX.PutValue ('EX_ETATPROD', 'NON');
 TEX.PutValue ('EX_DATECUM', DDeb);
 TEX.PutValue ('EX_DATECUMRUB', DDeb);


 if (((NatureFichier = 'DOS') and (stArg = '')) or (NatureFichier = '')) and (OrigineFichier <> 'SI')then
   TEX.InsertOrUpdateDB(TRUE)
 else
  begin
    if not ExisteSQL('SELECT EX_EXERCICE from EXERCICE Where EX_DATEDEBUT="' + UsDateTime(D1) + '" AND EX_DATEFIN="' + UsDateTime(D2)+'"') then
    begin
      if ExisteSQL('SELECT EX_EXERCICE from EXERCICE Where EX_EXERCICE="' +
        EnleveBlanc(Ligne, 7, 3,'') + '"' +
        ' AND EX_DATEDEBUT="' + UsDateTime(D1) + '" AND EX_DATEFIN="' + UsDateTime(D2)+'"')
        then
         TEX.InsertOrUpdateDB(TRUE)
      else
      begin
        if ExisteSQL('SELECT EX_EXERCICE from EXERCICE Where EX_EXERCICE="' +
        EnleveBlanc(Ligne, 7, 3,'')+ '"')  then
        begin
          Cdd := Format ('%.03d',[NombreExercice('', TRUE)+1]);
          TEX.PutValue ('EX_EXERCICE', Cdd);
          TEX.InsertOrUpdateDB(TRUE);
          if TA <> nil then
          TA.PutValue('EXERCICE', Cdd);
        end
        else
          TEX.InsertOrUpdateDB(TRUE);
      end;
    end
    else  (* AJOUT ME 05/02/2004 PB décalage des exercices pour les anouveaux*)
    begin
     // fiche 10508
     if (NatureFichier = 'BAL') or (NatureFichier = 'JRL') then
     begin
           Q1 := OpenSQL('SELECT EX_EXERCICE,EX_ETATCPTA from EXERCICE Where EX_DATEDEBUT="' + UsDateTime(D1) + '" AND EX_DATEFIN="' + UsDateTime(D2)+'"', TRUE);
           if not Q1.EOF then
           begin
                 if (Q1.FindField ('EX_ETATCPTA').asstring = 'NON') and (Etatcomptable <> 'NON') then
                 begin
                        if stArg <> '' then
                        begin
                          OnAfficheListeCom('Exercice ' + sLibelleExo +' est non ouvert dans le dossier '+
                          ' alors que l''état est à '+Etatcomptable+ ' dans le fichier.',LISTEIMPORT);
                        end
                        else
                             AffMessage ('Exercice ' + sLibelleExo +'  est non ouvert dans le dossier' +
                             ' alors que l''état est à '+ Etatcomptable + ' dans le fichier.');
                        OKexercice := FALSE; Ferme (Q1);
                        TEX.Free; exit;
                 end;
           end;
           Ferme (Q1);
     end
     else
     if (NatureFichier = 'SYN') then
     begin
           Q1 := OpenSQL('SELECT EX_EXERCICE,EX_ETATCPTA from EXERCICE Where EX_DATEDEBUT="' + UsDateTime(D1) + '" AND EX_DATEFIN="' + UsDateTime(D2)+'"', TRUE);
           if not Q1.EOF then
           begin
            Cdd := Q1.FindField ('EX_EXERCICE').asstring;
            if (TA <> nil) and (Q1.FindField ('EX_EXERCICE').asstring <> EnleveBlanc(Ligne, 7, 3,'')) then
               TA.PutValue('EXERCICE', Cdd);
               // fiche 10508
            if (Q1.FindField ('EX_ETATCPTA').asstring = 'NON') and (Etatcomptable <> 'NON') then
            begin
                  if stArg <> '' then
                  begin
                    OnAfficheListeCom('Exercice ' + sLibelleExo +' est non ouvert dans le dossier '+
                    ' alors que l''état est à '+Etatcomptable+ ' dans le fichier.',LISTEIMPORT);
                  end
                  else
                       AffMessage ('Exercice ' + sLibelleExo +' est non ouvert dans le dossier' +
                       ' alors que l''état est à '+ Etatcomptable + ' dans le fichier.');
                  OKexercice := FALSE; Ferme (Q1);
                  TEX.Free; exit;
            end
            else
            if (Q1.FindField ('EX_ETATCPTA').asstring = 'OUV') and (Etatcomptable = 'CDE') then
            begin
                             MAJEXO := TRUE;  MAJCPEXOREF := TRUE;
                             // fiche 10387 car cloture provisoire n'existe pas dans s1
                             if (OrigineFichier = 'S1') then
                             begin
                                  OnAfficheListeCom('Exercice ' + DateToStr(D1) +' est ouvert dans le dossier '+
                                  ' alors qu''il est clôturé dans le fichier. ',LISTEIMPORT);
                                  OKexercice := FALSE; Ferme (Q1);
                                  TEX.Free; exit;
                             end;
            end
            else
            begin
                 if (Q1.FindField ('EX_ETATCPTA').asstring = 'CDE') and (Etatcomptable = 'OUV') then
                 begin
                        if stArg <> '' then
                        begin
                          OnAfficheListeCom('Exercice ' + sLibelleExo +' est clôturé dans le dossier '+
                          ' alors qu''il est ouvert dans le fichier. Pour intégrer le fichier déclôturez l''exercice.',LISTEIMPORT);
                        end
                        else
                             AffMessage ('Exercice ' + sLibelleExo +' est clôturé dans le dossier' +
                             ' alors qu''il est ouvert dans le fichier.'+
                             '#10#13 Pour intégrer le fichier déclôturez l''exercice.');
                        OKexercice := FALSE; Ferme (Q1);
                        TEX.Free; exit;
                 end;
                 if (Q1.FindField ('EX_ETATCPTA').asstring = 'CPR') and ((Etatcomptable = 'CDE') or
                 (Etatcomptable = 'OUV'))then
                 begin
                                TEX.PutValue ('EX_ETATCPTA', 'CPR');
                                if (Etatcomptable = 'OUV') then
                                  OnAfficheListeCom('Exercice ' + DateToStr(D1) +' est clôturé provisoirement dans le dossier '+
                                  ' alors qu''il est ouvert dans le fichier. ',LISTEIMPORT);
                                if (Etatcomptable = 'OUV') then
                                  OnAfficheListeCom('Exercice ' + EnleveBlanc(sLibelleExo, 1, 14,'') +' est clôturé provisoirement dans le dossier '+
                                  ' alors qu''il est clôturé dans le fichier. ',LISTEIMPORT);
                 end;
                 if (Q1.FindField ('EX_ETATCPTA').asstring = 'OUV') and (Etatcomptable = 'CPR') then
                 begin
                                TEX.PutValue ('EX_ETATCPTA', 'OUV');
                                OnAfficheListeCom('Exercice ' + DateToStr(D1) +' est ouvert dans le dossier '+
                                  ' alors qu''il est clôturé provisoirement dans le fichier. ',LISTEIMPORT);
                 end;

            end;
           end;
           Ferme (Q1);
           if MAJEXO then // Modification EtatCpta
           begin
               if AuMoinsUneImmo and  (GetParamSocSecur('SO_EXOCLOIMMO', '') <> Cdd) then
               begin
                        if stArg <> '' then
                        begin
                             OnAfficheListeCom('Import impossible. La réception du fichier va générer une clôture comptable. Vous devez au préalable clôturer les immobilisations.',LISTEIMPORT);
                        end
                        else
                             AffMessage ('Import impossible. La réception du fichier va générer une clôture comptable #10#13Vous devez au préalable clôturer les immobilisations.');
                        OKexercice := FALSE;
                        TEX.Free; exit;
               end
               else
               begin
                   If ExisteSQL('SELECT E_IO FROM ECRITURE WHERE E_IO="X" and E_EXERCICE="'+Cdd+'"') then
                   begin

                        if stArg <> '' then
                          OnAfficheListeCom('La réception du fichier va générer une clôture comptable. Vous avez modifié des écritures sur cet exercice , la cohérence des A Nouveaux ne sera pas assurée .',LISTEIMPORT)
                        else
                        begin
{$IFNDEF EAGLSERVER}
                             if PGIAsk('La réception du fichier va générer une clôture comptable.'
                             + #10#13+'Vous avez modifié des écritures sur cet exercice , la cohérence des A Nouveaux ne sera pas assurée .'
                             + #10#13+'Il est conseillé de repartir d''un dossier complet en vue d''une première synchronisation.'
                             + #10#13+ 'Voulez vous continuer ?', 'Import') <> mryes Then
                             begin
                                  OKexercice := FALSE;
                                  TEX.Free; exit;
                             end;
{$ELSE}
                             OnAfficheListeCom('La réception du fichier va générer une clôture comptable. Vous avez modifié des écritures sur cet exercice , la cohérence des A Nouveaux ne sera pas assurée .',LISTEIMPORT);
{$ENDIF}
                        end;
                   end;
               end;
               TEX.SetAllModifie (TRUE); TEX.UpdateDB(TRUE);
           end;
     end;    (* fin AJOUT ME 05/02/2004 PB décalage des exercices pour les anouveaux*)
    end;
  end;

  if not RenExoClo (' Where EX_DATEDEBUT="' + UsDateTime(D1) + '" AND EX_DATEFIN="' + UsDateTime(D2)+'"'+
                     ' AND EX_ETATCPTA <>"OUV" AND EX_ETATCPTA <>"CPR" AND EX_ETATCPTA <>"NON"') then begin TEX.Free; exit; end;

  OKexercice       := TRUE;
  PresenceExercice := TRUE;
  TEX.Free;
{$IFDEF EAGLSERVER}
   TCPContexte.GetCurrent.Exercice.ChargeMagExo ;
{$ENDIF}

  OKexercice := OnAfficheListeCom('Exercice : '+  DateToStr(D1) + ' au ' +  DateToStr(D2),LISTEIMPORT);
end;

procedure TImportCom.ImportTableLibre(Ligne: string);
var
  Code, Nature: string;
  TL          : TOB;
begin

  Code := EnleveBlanc(Ligne, 59, 1,'');
  // voir nature s'il ne faut enlever les blancs ??
  Nature := EnleveBlanc(Ligne, 7, 17,'');
  // si aucune nature sortie ?
  if Code = '' then
    exit;

  TL := TOB.Create('NATCPTE', nil, -1);

  Code := Code + '0' + EnleveBlanc(Ligne, 6, 1,'');
  TL.PutValue ('NT_TYPECPTE', Code);
  TL.PutValue ('NT_NATURE', Nature);
  TL.PutValue ('NT_LIBELLE', VireTouteSub(EnleveBlanc(Ligne, 24, 35,''),'"'));
  TL.PutValue ('NT_TEXTE0', EnleveBlanc(Ligne, 62, 35,''));
  TL.PutValue ('NT_TEXTE1', EnleveBlanc(Ligne, 97, 35,''));
  TL.PutValue ('NT_TEXTE2', EnleveBlanc(Ligne, 132, 35,''));
  TL.PutValue ('NT_TEXTE3', EnleveBlanc(Ligne, 167, 35,''));
  TL.PutValue ('NT_TEXTE4', EnleveBlanc(Ligne, 202, 35,''));
  TL.PutValue ('NT_TEXTE5', EnleveBlanc(Ligne, 237, 35,''));
  TL.PutValue ('NT_TEXTE6', EnleveBlanc(Ligne, 272, 35,''));
  TL.PutValue ('NT_TEXTE7', EnleveBlanc(Ligne, 307, 35,''));
  TL.PutValue ('NT_TEXTE8', EnleveBlanc(Ligne, 342, 35,''));
  TL.PutValue ('NT_TEXTE9', EnleveBlanc(Ligne, 377, 35,''));
  TL.InsertOrUpdateDB(TRUE);
  TL.Free;

  OKexercice := OnAfficheListeCom('Table Libre nature : '+ Nature,LISTEIMPORT);

end;

procedure TImportCom.ImportSection(Ligne: string);
var
TS : TOB;
begin
  // contient la définition des sous plans
    TS := TOB.Create('STRUCRSE', nil, -1);
    TS.PutValue ('SS_AXE',EnleveBlanc(Ligne,59, 3,''));
    TS.PutValue ('SS_SOUSSECTION', EnleveBlanc(Ligne, 62, 3,''));
    TS.PutValue ('SS_LIBELLE', VireTouteSub(EnleveBlanc(Ligne, 69, 35,''),'"'));
    TS.PutValue ('SS_DEBUT', EnleveBlanc(Ligne, 65, 2,'0'));
    TS.PutValue ('SS_LONGUEUR', EnleveBlanc(Ligne, 67, 2,'10'));
    TS.PutValue ('SS_AXE', EnleveBlanc(Ligne, 59, 3,''));
    TS.InsertOrUpdateDB(TRUE);
    TS.Free;


  // Contient le contenu des sous plan

    TS := TOB.Create('SSSTRUCR', nil, -1);
    TS.PutValue ('PS_AXE',EnleveBlanc(Ligne,59, 3,''));
    TS.PutValue ('PS_SOUSSECTION', EnleveBlanc(Ligne, 62, 3,''));
    TS.PutValue ('PS_AXE', EnleveBlanc(Ligne, 59, 3,''));
    TS.PutValue ('PS_LIBELLE', VireTouteSub(EnleveBlanc(Ligne, 24, 35,''),'"'));
    TS.PutValue ('PS_CODE', EnleveBlanc(Ligne, 7, 17,''));
    TS.InsertOrUpdateDB(TRUE);
    TS.Free;

end;

procedure TImportCom.ImporteEtablissement(Ligne: string);
var
TEta : TOB;
Etab : string;
begin
    Etab := EnleveBlanc(Ligne, 7, 3,'');
    if BControlEtab then  // Il faut contrôler l'existance de l'établissement
    begin
        if not ExisteSql ('SELECT ET_ETABLISSEMENT from ETABLISS WHERE ET_ETABLISSEMENT="'+Etab+'"') then
        begin
                    OnAfficheListeCom('Etablissement '+ Etab +' est manquant',LISTEIMPORT);
                    OkCreatCompte := FALSE;  exit;
        end;
    end;
    TEta := TOB.Create('ETABLISS', nil, -1);
    TEta.PutValue ('ET_ETABLISSEMENT', Etab);
    TEta.PutValue ('ET_LIBELLE', VireTouteSub(EnleveBlanc(Ligne, 10, 35,''),'"'));
    if (VersionFichier <> '') and (StrToint(VersionFichier) >= 8) then
    begin
        TEta.PutValue ('ET_ADRESSE1', EnleveBlanc(Ligne, 45, 35,''));
        TEta.PutValue ('ET_ADRESSE2', EnleveBlanc(Ligne, 80, 35,''));
        TEta.PutValue ('ET_ADRESSE3', EnleveBlanc(Ligne, 115, 35,''));
        TEta.PutValue ('ET_CODEPOSTAL', EnleveBlanc(Ligne, 150, 9,''));
        TEta.PutValue ('ET_VILLE', EnleveBlanc(Ligne, 159, 35,''));
        TEta.PutValue ('ET_PAYS', EnleveBlanc(Ligne, 194, 3,''));
        TEta.PutValue ('ET_TELEPHONE', EnleveBlanc(Ligne, 197, 25,''));
        TEta.PutValue ('ET_TELEX', EnleveBlanc(Ligne, 222, 25,''));
        TEta.PutValue ('ET_SIRET', EnleveBlanc(Ligne, 247, 17,''));
        TEta.PutValue ('ET_APE', EnleveBlanc(Ligne, 264, 5,''));
    end;

    TEta.InsertOrUpdateDB(TRUE);
    TEta.Free;
    OKexercice := OnAfficheListeCom('Etablisssement : ' + EnleveBlanc(Ligne, 7, 3,'') + ' ' + EnleveBlanc(Ligne, 10, 35,''),LISTEIMPORT);
end;

procedure TImportCom.ImporteModepaiement(Ligne: string);
var
TMdp      : TOB;
MDP       : string;
begin
    TMdp := TOB.Create('MODEPAIE', nil, -1);
    MDP := EnleveBlanc(Ligne, 7, 3, '');
    TraiteCorrespCpt(7, EnleveBlanc(Ligne, 7, 3, ''), MDP);

    TMdp.PutValue ('MP_MODEPAIE', MDP);
    TMdp.PutValue ('MP_LIBELLE', VireTouteSub(EnleveBlanc(Ligne, 10, 35, ''),'"'));
    TMdp.PutValue ('MP_CATEGORIE', EnleveBlanc(Ligne, 45, 3, ''));
    TMdp.PutValue ('MP_CODEACCEPT', EnleveBlanc(Ligne, 48, 3, ''));
    TMdp.PutValue ('MP_LETTRECHEQUE', EnleveBlanc(Ligne, 51, 1, ''));
    TMdp.PutValue ('MP_LETTRETRAITE', EnleveBlanc(Ligne, 52, 1, ''));
    TMdp.PutValue ('MP_CONDITION', EnleveBlanc(Ligne, 53, 1, ''));
    TMdp.PutValue ('MP_MONTANTMAX', Valeur(StPoint(EnleveBlanc(Ligne, 54, 20,'0'))));
    TMdp.PutValue ('MP_REMPLACEMAX', EnleveBlanc(Ligne, 74, 3, ''));
    if EnleveBlanc(Ligne, 77, 17, '') <> '' then  // fiche 10604
    TMdp.PutValue ('MP_GENERAL', BourreOuTronque(EnleveBlanc(Ligne, 77, 17, ''), FbGene));
    TMdp.InsertOrUpdateDB(TRUE);
    TMdp.Free;

  OnAfficheListeCom('Mode de paiement : ' + EnleveBlanc(Ligne, 7, 3, ''),LISTEIMPORT);
end;

procedure TImportCom.ImporteModereglement(Ligne: string);
var
TMdr      : TOB;
Apart     : string;
nbeche    : integer;
begin
    TMdr := TOB.Create('MODEREGL', nil, -1);

    TMdr.PutValue ('MR_MODEREGLE', EnleveBlanc(Ligne, 7, 3, ''));
    TMdr.PutValue ('MR_LIBELLE', VireTouteSub(EnleveBlanc(Ligne, 10, 35, ''),'"'));
    Apart := EnleveBlanc(Ligne, 45, 3, '');
    if (Apart <> 'BL') and (Apart <> 'DEB') and (Apart <> 'ECR')
    and (Apart <> 'FAC') and (Apart <> 'FAF') and (Apart <> 'FIN') then
    Apart := 'FAC';
    TMdr.PutValue ('MR_APARTIRDE', Apart);
    TMdr.PutValue ('MR_PLUSJOUR', EnleveBlanc(Ligne, 48, 3, '30'));
    TMdr.PutValue ('MR_ARRONDIJOUR', EnleveBlanc(Ligne, 51, 3, ''));

    nbeche := StrToint(EnleveBlanc(Ligne, 54, 2,'0'));
    if (nbeche >=1) and (nbeche <= 12) then
       TMdr.PutValue ('MR_NOMBREECHEANCE', nbeche)
    else
    begin
       nbeche := 1;
       TMdr.PutValue ('MR_NOMBREECHEANCE', nbeche);
    end;
    TMdr.PutValue ('MR_SEPAREPAR', EnleveBlanc(Ligne, 56, 3, ''));
// modif V590
    TMdr.PutValue ('MR_MONTANTMIN', Valeur(StPoint(EnleveBlanc(Ligne, 59, 20,'0'))));

    TMdr.PutValue ('MR_REMPLACEMIN', EnleveBlanc(Ligne, 79, 3, ''));
    TMdr.PutValue ('MR_MP1', EnleveBlanc(Ligne, 82, 3, ''));
    TMdr.PutValue ('MR_TAUX1', Valeur(StPoint(EnleveBlanc(Ligne, 85, 6, '0'))));

    TMdr.PutValue ('MR_MP2', EnleveBlanc(Ligne, 91, 3, ''));
    TMdr.PutValue ('MR_TAUX2', Valeur(StPoint(EnleveBlanc(Ligne, 94, 6, '0'))));

    TMdr.PutValue ('MR_MP3', EnleveBlanc(Ligne, 100, 3, ''));
    TMdr.PutValue ('MR_TAUX3', Valeur(StPoint(EnleveBlanc(Ligne, 103, 6, '0'))));

    TMdr.PutValue  ('MR_MP4', EnleveBlanc(Ligne, 109, 3, ''));
    TMdr.PutValue ('MR_TAUX4', Valeur(StPoint(EnleveBlanc(Ligne, 112, 6, '0'))));

    TMdr.PutValue ('MR_MP5', EnleveBlanc(Ligne, 118, 3, ''));
    TMdr.PutValue ('MR_TAUX5', Valeur(StPoint(EnleveBlanc(Ligne, 121, 6, '0'))));

    TMdr.PutValue ('MR_MP6', EnleveBlanc(Ligne, 127, 3, ''));
    TMdr.PutValue ('MR_TAUX6', Valeur(StPoint(EnleveBlanc(Ligne, 130, 6, '0'))));

    TMdr.PutValue ('MR_MP7', EnleveBlanc(Ligne, 136, 3, ''));
    TMdr.PutValue ('MR_TAUX7', Valeur(StPoint(EnleveBlanc(Ligne, 139, 6, '0'))));

    TMdr.PutValue ('MR_MP8', EnleveBlanc(Ligne, 145, 3, ''));
    TMdr.PutValue ('MR_TAUX8', Valeur(StPoint(EnleveBlanc(Ligne, 148, 6, '0'))));

    TMdr.PutValue ('MR_MP9', EnleveBlanc(Ligne, 154, 3, ''));
    TMdr.PutValue ('MR_TAUX9', Valeur(StPoint(EnleveBlanc(Ligne, 157, 6, '0'))));

    TMdr.PutValue ('MR_MP10', EnleveBlanc(Ligne, 163, 3, ''));
    TMdr.PutValue ('MR_TAUX10', Valeur(StPoint(EnleveBlanc(Ligne, 166, 6, '0'))));

    TMdr.PutValue ('MR_MP11', EnleveBlanc(Ligne, 172, 3, ''));
    TMdr.PutValue ('MR_TAUX11', Valeur(StPoint(EnleveBlanc(Ligne, 175, 6, '0'))));

    TMdr.PutValue ('MR_MP12', EnleveBlanc(Ligne, 181, 3, ''));
    TMdr.PutValue ('MR_TAUX12', Valeur(StPoint(EnleveBlanc(Ligne, 184, 6, '0'))));

    remplissageMDR (nbeche, TMdr);

    TMdr.InsertOrUpdateDB(TRUE);
    TMdr.Free;

  OnAfficheListeCom('Mode de réglement  : ' + EnleveBlanc(Ligne, 7, 3, ''),LISTEIMPORT);
end;


procedure TImportCom.remplissageMDR (indice: integer; var Tm : TOB);
var ii        : integer;
Total         : double;
OkMR          : Boolean;
begin
    Total := 0.0;
    OKMr := TRUE;
    for ii := 1 to indice do
    begin
         if Tm.GetValue ('MR_MP'+IntTostr(ii)) = '' then
            Tm.PutValue ('MR_MP'+IntTostr(ii), 'CHQ');
         if Tm.GetValue ('MR_TAUX'+IntTostr(ii)) = 0 then
         begin OKMR := FALSE; break; end
         else
            Total := Total + Tm.GetValue ('MR_TAUX'+IntTostr(ii));
    end;
    if OKMR and (Total <> 0) and (Arrondi(Total,V_PGI.OkDecV) <> 100.0) then
    begin
              if Total < 100 then
                 Tm.PutValue ('MR_TAUX'+IntTostr(indice), (100-Total))
              else
                  OKMR := FALSE;
    end;

    if not OKMR then
    begin
         Tm.PutValue ('MR_TAUX1', 100.0);
         Tm.PutValue ('MR_NOMBREECHEANCE', 1);
         for ii := 2 to 12 do
         begin
              Tm.PutValue ('MR_MP'+IntTostr(ii),'');
              Tm.PutValue ('MR_TAUX'+IntTostr(ii), 0);
         end;
    end
    else
    begin
         for ii := indice+1 to 12 do
         begin
              if Tm.GetValue ('MR_MP'+IntTostr(ii)) <> '' then
                 Tm.PutValue ('MR_MP'+IntTostr(ii),'');
              if Tm.GetValue ('MR_TAUX'+IntTostr(ii)) <>  0 then
                 Tm.PutValue ('MR_TAUX'+IntTostr(ii), 0);
         end;
    end;
end;

procedure TImportCom.ImporteDevise(Ligne: string);
var
  Pariete, quotite : double;
  TDv : TOB;
begin
    TDv := TOB.Create('DEVISE', nil, -1);

    Pariete := Valeur(StPoint((EnleveBlanc(Ligne, 57, 20, '1'))));
    Quotite := Valeur(StPoint((EnleveBlanc(Ligne, 50, 6, '0'))));

    TDv.putValue ('D_DEVISE', EnleveBlanc(Ligne, 7, 3));
    TDv.putValue ('D_LIBELLE', VireTouteSub(EnleveBlanc(Ligne, 10, 35),'"'));
    TDv.putValue ('D_SYMBOLE', EnleveBlanc(Ligne, 45, 3));
    TDv.putValue ('D_FERME', EnleveBlanc(Ligne, 48, 1));
    TDv.putValue ('D_DECIMALE', EnleveBlanc(Ligne, 49, 1));
    TDv.putValue ('D_QUOTITE', Quotite);
    TDv.putValue ('D_MONNAIEIN', EnleveBlanc(Ligne, 56, 1));
    TDv.putValue ('D_PARITEEURO', Pariete);
    TDv.putValue ('D_CODEISO', EnleveBlanc(Ligne, 77, 3));
    TDv.putValue ('D_FONGIBLE', EnleveBlanc(Ligne, 80, 1));
    TDv.putValue ('D_CPTLETTRDEBIT', EnleveBlanc(Ligne, 81, 17));
    TDv.putValue ('D_CPTLETTRCREDIT', EnleveBlanc(Ligne, 98, 17));
    TDv.putValue ('D_CPTPROVDEBIT', EnleveBlanc(Ligne, 115, 17));
    TDv.putValue ('D_CPTPROVCREDIT', EnleveBlanc(Ligne, 132, 17));
    TDv.putValue ('D_MAXDEBIT', Valeur(StPoint(EnleveBlanc(Ligne, 149, 20, '0'))));
    TDv.putValue ('D_MAXCREDIT', Valeur(StPoint(EnleveBlanc(Ligne, 169, 20, '0'))));
    TDv.InsertOrUpdateDB(TRUE);
    TDv.Free;

  OnAfficheListeCom('Devise : ' + EnleveBlanc(Ligne, 7, 3),LISTEIMPORT);
end;

procedure TImportCom.ImporteRegimeTVA(Ligne: string);
var
  Ttva : TOB;
begin
    Ttva := TOB.Create('CHOIXCOD', nil, -1);
    Ttva.putvalue ('CC_TYPE', 'RTV');
    Ttva.putvalue ('CC_CODE', EnleveBlanc(Ligne, 7, 3));
    Ttva.putvalue ('CC_LIBELLE', VireTouteSub(EnleveBlanc(Ligne, 10, 35),'"'));
    Ttva.InsertOrUpdateDB(TRUE);
    Ttva.Free;

  OnAfficheListeCom('Régime TVA  : ' + EnleveBlanc(Ligne, 7, 3),LISTEIMPORT);
end;


procedure TImportCom.ImportSectionAnalytique(Ligne: string);
var
  Code, axe: string;
  Q1       : Tquery;
  attente  : string;
  TSec     : TOB;
begin
  // voir nature s'il ne faut enlever les blancs ??
  axe := EnleveBlanc(Ligne, 59, 3);
  if axe = 'A1' then
  Code := BourreOuTronque(EnleveBlanc(Ligne, 7, 17), fbaxe1);
  if axe = 'A2' then
  Code := BourreOuTronque(EnleveBlanc(Ligne, 7, 17), fbaxe2);
  if axe = 'A3' then
  Code := BourreOuTronque(EnleveBlanc(Ligne, 7, 17), fbaxe3);
  if axe = 'A4' then
  Code := BourreOuTronque(EnleveBlanc(Ligne, 7, 17), fbaxe4);
  if axe = 'A5' then
  Code := BourreOuTronque(EnleveBlanc(Ligne, 7, 17), fbaxe5);

  // fiche 10568
  if (OrigineEnvoie <> 'EXP') then
  begin
     if (NatureFichier <> 'DOS') and (AxeTva <> '') and (Axe = AxeTva) then
     begin
              inc (NbSection);
              exit;
     end;
  end;

  TSec := TOB.Create('SECTION', nil, -1);
  TSec.Putvalue ('S_AXE', axe);
  TSec.Putvalue ('S_SECTION', Code);
  TSec.Putvalue ('S_LIBELLE', VireTouteSub(EnleveBlanc(Ligne, 24, 35),'"'));
  TSec.Putvalue ('S_TABLE0',  Veriftablelibre('S00', EnleveBlanc(Ligne, 62, 17)));
  TSec.Putvalue ('S_TABLE1',  Veriftablelibre('S01', EnleveBlanc(Ligne, 79, 17)));
  TSec.Putvalue ('S_TABLE2',  Veriftablelibre('S02', EnleveBlanc(Ligne, 96, 17)));
  TSec.Putvalue ('S_TABLE3',  Veriftablelibre('S03', EnleveBlanc(Ligne, 113, 17)));
  TSec.Putvalue ('S_TABLE4',  Veriftablelibre('S04', EnleveBlanc(Ligne, 130, 17)));
  TSec.Putvalue ('S_TABLE5',  Veriftablelibre('S05', EnleveBlanc(Ligne, 147, 17)));
  TSec.Putvalue ('S_TABLE6',  Veriftablelibre('S06', EnleveBlanc(Ligne, 164, 17)));
  TSec.Putvalue ('S_TABLE7',  Veriftablelibre('S07', EnleveBlanc(Ligne, 181, 17)));
  TSec.Putvalue ('S_TABLE8',  Veriftablelibre('S08', EnleveBlanc(Ligne, 198, 17)));
  TSec.Putvalue ('S_TABLE9',  Veriftablelibre('S09', EnleveBlanc(Ligne, 215, 17)));
  
  if EnleveBlanc(Ligne, 232, 17) <> '' then
     TSec.Putvalue ('S_ABREGE', VireTouteSub(EnleveBlanc(Ligne, 232, 17),'"'))
  else
     TSec.Putvalue ('S_ABREGE', VireTouteSub(EnleveBlanc(Ligne, 24, 17),'"'));
  if EnleveBlanc(Ligne, 249, 3) <> '' then
     TSec.Putvalue ('S_SENS', EnleveBlanc(Ligne, 249, 3))
  else
     TSec.Putvalue ('S_SENS', 'M');
  TSec.PutValue ('S_CONFIDENTIEL', '0');
  if SSectionInbase <> nil then  // fiche 10583 modification pour les sections qui sont avant chaque écriture  A PORTER sur V7 ed9
  begin
      if SSectionInbase.IndexOf (Code+':'+Axe) < 0 then
         SSectionInbase.Add (Code + ':'+ axe);
  end;
  TSec.InsertOrUpdateDB(TRUE);
  TSec.Free;

  OnAfficheListeCom('Sections analytiques  : ' + Code, LISTEIMPORT );
  Q1 := OpenSql ('SELECT X_SECTIONATTENTE from axe where X_AXE="'+ axe+'"', True);
  if  not Q1.EOF then
          attente := Q1.findField ('X_SECTIONATTENTE').asstring;
  ferme (Q1);

  if (attente = '') and (Code <> '') then
  begin
               ExecuteSQL('UPDATE AXE SET X_AXE="' + axe +'",X_SECTIONATTENTE="' + code +
               '" Where X_AXE="'+axe+'"');
  end;
end;

procedure TImportCom.ImportSouche(Ligne: string);
var
  Typesouche: string;
  TSou      : TOB;
begin

  if EnleveBlanc(Ligne, 62, 1) = 'O' then
    Typesouche := 'REL'
  else
    Typesouche := 'CPT';

  TSou := TOB.Create('SOUCHE', nil, -1);
  TSou.Putvalue ('SH_TYPE', Typesouche);
  TSou.Putvalue ('SH_SOUCHE', EnleveBlanc(Ligne, 7, 3));
  TSou.Putvalue ('SH_LIBELLE', VireTouteSub(EnleveBlanc(Ligne, 10, 35),'"'));
  TSou.Putvalue ('SH_NUMDEPART', EnleveBlanc(Ligne, 45, 15));
  TSou.Putvalue ('SH_SIMULATION', EnleveBlanc(Ligne, 60, 1));
  TSou.Putvalue ('SH_ANALYTIQUE', EnleveBlanc(Ligne, 61, 1));
  if Not TSou.SelectDB('"'+Typesouche+'";"'+EnleveBlanc(Ligne, 7, 3)+'"', Nil) then
  begin
       // pour fiche 10073
       if NatureFichier = 'SYN' then TSou.Putvalue ('SH_NUMDEPART', 1);

       TSou.InsertDB(nil, TRUE)
  end
  else
  begin
       if (NatureFichier <> 'SYN') then
       begin
          TSou.Putvalue ('SH_TYPE', Typesouche);
          TSou.Putvalue ('SH_SOUCHE', EnleveBlanc(Ligne, 7, 3));
          TSou.Putvalue ('SH_LIBELLE', VireTouteSub(EnleveBlanc(Ligne, 10, 35),'"'));
          TSou.Putvalue ('SH_NUMDEPART', EnleveBlanc(Ligne, 45, 15));
          TSou.Putvalue ('SH_SIMULATION', EnleveBlanc(Ligne, 60, 1));
          TSou.Putvalue ('SH_ANALYTIQUE', EnleveBlanc(Ligne, 61, 1));
          if Paramgene then begin TSou.SetAllModifie (TRUE); TSou.UpdateDB(TRUE); end;
       end;
  end;
  TSou.Free;
  OKexercice := OnAfficheListeCom('Souches : ' + EnleveBlanc(Ligne, 7, 3),LISTEIMPORT );
end;

procedure TImportCom.InitGeneraux;
var
  sSQL       : string;
  comptegene : string;
  ii         : integer;
  Tg         : TOB;
  Val        : string;
begin
  if TGen = nil then
  begin
      // Récupération des indices significatifs de la TOB des comptes généraux
      Tg := TOB.Create ('GENERAUX',nil,-1);
      FiTOBG_LETTRABLE := Tg.GetNumChamp ('G_LETTRABLE');
      FiTOBG_NATUREGENE := Tg.GetNumChamp ('G_NATUREGENE');
      FiTOBG_POINTABLE := Tg.GetNumChamp ('G_POINTABLE');
      FiTOBG_VENTILABLE1 := Tg.GetNumChamp ('G_VENTILABLE1');
      FiTOBG_VENTILABLE2 := Tg.GetNumChamp ('G_VENTILABLE2');
      FiTOBG_VENTILABLE3 := Tg.GetNumChamp ('G_VENTILABLE3');
      FiTOBG_VENTILABLE4 := Tg.GetNumChamp ('G_VENTILABLE4');
      FiTOBG_VENTILABLE5 := Tg.GetNumChamp ('G_VENTILABLE5');
      FiTOBG_REGIMETVA := Tg.GetNumChamp ('G_REGIMETVA');
      Tg.free;

  // si les natures des comptes différents par rapport à paramsoc
    if (Paramgene) and (OrigineFichier <> 'SI') then
    begin
        for ii:=0 to SizeRecControl do
        begin
             comptegene := BourreOuTronque(GetParamSocSecur (RecControl[ii].NomChamp, ''),RecControl[ii].LeFb);
             if (RecControl[ii].champ  = 'CPT')  and (RecControl[ii].LeFb = fbGene) then
             begin
              Val := BourreOuTronque(RecControl[ii].valeur, fbgene);
              if (comptegene <> Val) and (Val <> '') then
               begin
                    SetParamSoc (RecControl[ii].NomChamp, Val);
                    ExecuteSQL('DELETE FROM GENERAUX WHERE G_GENERAL="'+ comptegene+'"');
               end;
             end;
        end;
    end;
//    i:= gettickcount;

// pb chargement gros dossier + 28000 lignes
//XXXX   if (NatureFichier = 'JRL') then GrandNbGene := TRUE;

    TGen := TOB.Create('', nil, -1);

     // Objet de stockage en mémoire
     fHGene := THashTable.Create();
     fHGene.useElfHash := true;
     fHGene.singleThread := true;

    if not GrandNbGene then // AJOUT ME 14-01-2005
    begin
          sSQL := 'SELECT * FROM GENERAUX ORDER BY G_GENERAL';
          ChargeTobTiersGene (sSQl, 'G');
    end;
//    j:= gettickcount;
//    showmessage(inttostr(j-i));

    // si changement longueur compte
    if Changementgene then
    begin
          for ii := TGen.detail.Count -1 downto 0 do
          begin
            comptegene := BourreOuTronque( TGen.detail[ii].GetValue('G_GENERAL'), fbgene);
            if (comptegene) <> (TGen.detail[ii].GetValue('G_GENERAL')) then
            begin
                 ExecuteSQL('DELETE FROM GENERAUX WHERE G_GENERAL="'+ TGen.detail[ii].GetValue('G_GENERAL')+'"');
                 TGen.detail[ii].PutValue('G_GENERAL', comptegene);
            end;
          end;
          TGen.SetAllModifie (TRUE);
          TGen.InsertOrUpdateDB(TRUE);
    end;
    TGenupdate := TOB.Create('', nil, -1);
    TGenInsert := TOB.Create('', nil, -1);

     fHGeneInsert := THashTable.Create();
     fHGeneInsert.useElfHash := true;
     fHGeneInsert.singleThread := true;

    OKexercice := OnAfficheListeCom('Comptes généraux',LISTEIMPORT);
  end;
end;

procedure TImportCom.Remplirgeneraux (var T : TOB; Ligne,compte, Nature : string; Casupdate : Boolean=FALSE);
Var
OkMaj : Boolean;
begin
    OkMaj := TRUE;

    // si compte avec minuscule dedans
    compte := UpperCase (compte);
    if CasUpdate then  //pour initialiser les champ
             T.SelectDB('"' + compte + '"', nil, false);

    T.PutValue('G_GENERAL', compte);

    if EnleveBlanc(Ligne, 24, 35) = '' then
        T.PutValue('G_LIBELLE', '.')
    else
        T.PutValue('G_LIBELLE', VireTouteSub(EnleveBlanc(Ligne, 24, 35),'"'));

    if Casupdate then
         if ExisteSQL('SELECT * FROM ECRITURE WHERE E_GENERAL ="'+compte+'"') then OkMaj := FALSE
    else
            T.PutValue('G_CREERPAR', 'COM');


    if OkMaj then
    begin
        T.PutValue('G_NATUREGENE', Nature);
        T.PutValue('G_POINTABLE', EnleveBlanc(Ligne, 63, 1));
        if (Nature = 'COC') or (Nature = 'COF') or (Nature = 'COS') or (Nature =
          'COD') then
          T.PutValue('G_COLLECTIF', 'X');
        if Nature = 'BQE' then   //pour ne pas perdre les infos de pointage, dans le cas où il y a des informations du pointage sur l'écriture
          T.PutValue('G_POINTABLE', 'X');
        T.PutValue('G_LETTRABLE', EnleveBlanc(Ligne, 62, 1));
        T.PutValue('G_VENTILABLE1', EnleveBlanc(Ligne, 64, 1));
        T.PutValue('G_VENTILABLE2', EnleveBlanc(Ligne, 65, 1));
        T.PutValue('G_VENTILABLE3', EnleveBlanc(Ligne, 66, 1));
        T.PutValue('G_VENTILABLE4', EnleveBlanc(Ligne, 67, 1));
        T.PutValue('G_VENTILABLE5', EnleveBlanc(Ligne, 68, 1));
        if (EnleveBlanc(Ligne, 64, 1) = 'X') or  (EnleveBlanc(Ligne, 65, 1) = 'X')
        or (EnleveBlanc(Ligne, 66, 1) = 'X') or  (EnleveBlanc(Ligne, 67, 1) = 'X')
        or (EnleveBlanc(Ligne, 68, 1) = 'X') then
           T.PutValue('G_VENTILABLE', 'X')
        else
           T.PutValue('G_VENTILABLE', '-');

    end;

    // fiche 13221
    if (Nature <> 'DIV') and (Nature <> 'TID') and (Nature <> 'TIC') and (EnleveBlanc(Ligne, 62, 1) = 'X') then
    begin
                    if (stArg = '') then
                    begin
                            OnAfficheListeCom('Attention compte :'+ compte + ' ne peut être lettrable', LISTEIMPORT);
                    end;
                    T.PutValue('G_LETTRABLE', '-');
    end;

                                        // fiche 13221
    if (Nature = 'TID') or (Nature = 'TIC') or ((Nature = 'DIV') and (EnleveBlanc(Ligne, 62, 1) = 'X')) then
    begin
      T.PutValue('G_LETTRABLE', 'X');
      if T.GetValue('G_MODEREGLE') = '' then
           T.PutValue('G_MODEREGLE', RendMRDefaut);
      if T.GetValue('G_REGIMETVA') = '' then
           T.PutValue('G_REGIMETVA', RendRegTvadefaut);
      if T.GetValue('G_TVAENCAISSEMENT') = '' then
           T.PutValue('G_TVAENCAISSEMENT', GetParamSocSecur('SO_CODETVADEFAUT', ''));
    end;


    if Casupdate and  (Nature = 'DIV') and (EnleveBlanc(Ligne, 62, 1) = 'X') then
       UpdateLettreDiv := TRUE;

    if (OrigineFichier = 'S1') then // à vérifier
    begin                                              // fiche 13221
         if (Nature <> 'TID') and (Nature <> 'TIC') and (Nature <> 'DIV') and (EnleveBlanc(Ligne, 62, 1) = 'X') then
         begin
                 T.PutValue('G_LETTRABLE', '-');
                 if  (copy (compte, 1, 1) = '6') then T.PutValue('G_NATUREGENE', 'CHA');
                 if  (copy (compte, 1, 1) = '7') then T.PutValue('G_NATUREGENE', 'PRO');
                 if  (copy (compte, 1, 3) = '512') then T.PutValue('G_NATUREGENE', 'BQE');
                 if  (copy (compte, 1, 3) = '514') then T.PutValue('G_NATUREGENE', 'BQE');
                 if  (copy (compte, 1, 2) = '53') then T.PutValue('G_NATUREGENE', 'CAI');
                 if  (copy (compte, 1, 2) = '21') then T.PutValue('G_NATUREGENE', 'IMO');
                 if  (copy (compte, 1, 2) = '20') then T.PutValue('G_NATUREGENE', 'IMO');
                 // CA - 16/04/2002 - Affichage message information correction
                 OnAfficheListeCom('Correction compte général  : ' + compte, LISTEIMPORT );
         end
         else
         begin
                 if  (copy (compte, 1, 1) = '6') and (Nature <> 'CHA') then T.PutValue('G_NATUREGENE', 'CHA');
                 if  (copy (compte, 1, 1) = '7') and (Nature <> 'PRO') then T.PutValue('G_NATUREGENE', 'PRO');
                 if  (copy (compte, 1, 3) = '512') and (Nature <> 'BQE') then T.PutValue('G_NATUREGENE', 'BQE');
                 if  (copy (compte, 1, 3) = '514') and (Nature <> 'BQE') then T.PutValue('G_NATUREGENE', 'BQE');
                 if  (copy (compte, 1, 2) = '53') and (Nature <> 'CAI') then T.PutValue('G_NATUREGENE', 'CAI');
                 if  (copy (compte, 1, 2) = '21') and (Nature <> 'IMO') then T.PutValue('G_NATUREGENE', 'IMO');
                 if  (copy (compte, 1, 2) = '20') and (Nature <> 'IMO') then T.PutValue('G_NATUREGENE', 'IMO');
                 // CA - 16/04/2002 - Affichage message information correction
                 if (Nature <> T.GetValue('G_NATUREGENE')) then
                    OnAfficheListeCom('Correction nature compte général  : ' + compte, LISTEIMPORT );
         end;
         // si le compte est rendu ventilable dans S1 fiche 10292
         // pour rendre ventilable tous les écritures
         if (T.GetValue('G_VENTILABLE1') <> 'X') and  (EnleveBlanc(Ligne, 64, 1) = 'X') then
         begin
              T.PutValue ('G_VENTILABLE1','X');
              T.PutValue ('G_VENTILABLE', 'X');
              ExecuteSQL('UPDATE  GENERAUX SET G_VENTILABLE="X", G_VENTILABLE1="X"  Where G_GENERAL="'+compte+'"');
              MajGVentil (FALSE, compte);
         end;
    end;
    if VersionFichier = '001' then
    begin
        T.PutValue('G_TABLE0', Veriftablelibre('G00', EnleveBlanc(Ligne, 69, 3)));
        T.PutValue('G_TABLE1', Veriftablelibre('G01', EnleveBlanc(Ligne, 72, 3)));
        T.PutValue('G_TABLE2', Veriftablelibre('G02', EnleveBlanc(Ligne, 75, 3)));
        T.PutValue('G_TABLE3', Veriftablelibre('G03', EnleveBlanc(Ligne, 78, 3)));
        T.PutValue('G_TABLE4', Veriftablelibre('G04', EnleveBlanc(Ligne, 81, 3)));
        T.PutValue('G_TABLE5', Veriftablelibre('G05', EnleveBlanc(Ligne, 84, 3)));
        T.PutValue('G_TABLE6', Veriftablelibre('G06', EnleveBlanc(Ligne, 87, 3)));
        T.PutValue('G_TABLE7', Veriftablelibre('G07', EnleveBlanc(Ligne, 90, 3)));
        T.PutValue('G_TABLE8', Veriftablelibre('G08', EnleveBlanc(Ligne, 93, 3)));
        T.PutValue('G_TABLE9', Veriftablelibre('G09', EnleveBlanc(Ligne, 96, 3)));
        T.PutValue('G_ABREGE', VireTouteSub(EnleveBlanc(Ligne, 99, 17),'"')) ;
        T.PutValue('G_SENS', EnleveBlanc(Ligne, 116, 3));
    end
    else
    begin
        T.PutValue('G_TABLE0',  Veriftablelibre('G00', EnleveBlanc(Ligne, 69, 17)));
        T.PutValue('G_TABLE1',  Veriftablelibre('G01', EnleveBlanc(Ligne, 86, 17)));
        T.PutValue('G_TABLE2',  Veriftablelibre('G02', EnleveBlanc(Ligne, 103, 17)));
        T.PutValue('G_TABLE3',  Veriftablelibre('G03', EnleveBlanc(Ligne, 120, 17)));
        T.PutValue('G_TABLE4',  Veriftablelibre('G04', EnleveBlanc(Ligne, 137, 17)));
        T.PutValue('G_TABLE5',  Veriftablelibre('G05', EnleveBlanc(Ligne, 154, 17)));
        T.PutValue('G_TABLE6',  Veriftablelibre('G06', EnleveBlanc(Ligne, 171, 17)));
        T.PutValue('G_TABLE7',  Veriftablelibre('G07', EnleveBlanc(Ligne, 188, 17)));
        T.PutValue('G_TABLE8',  Veriftablelibre('G08', EnleveBlanc(Ligne, 205, 17)));
        T.PutValue('G_TABLE9',  Veriftablelibre('G09', EnleveBlanc(Ligne, 222, 17)));
        if EnleveBlanc(Ligne, 239, 17) <> '' then
           T.PutValue('G_ABREGE', VireTouteSub(EnleveBlanc(Ligne, 239, 17),'"'))
        else
           T.PutValue('G_ABREGE', VireTouteSub(EnleveBlanc(Ligne, 24, 17),'"'));
        if EnleveBlanc(Ligne, 256, 3) <> '' then
           T.PutValue('G_SENS', EnleveBlanc(Ligne, 256, 3))
        else
           T.PutValue('G_SENS', 'M');
        T.PutValue('G_CORRESP1', EnleveBlanc(Ligne, 259, 17));
        T.PutValue('G_CORRESP2', EnleveBlanc(Ligne, 276, 17));

        // ajout me 12-01-2005
        if EnleveBlanc(Ligne, 293, 3) <> '' then T.PutValue('G_TVA', EnleveBlanc(Ligne, 293, 3));
        if EnleveBlanc(Ligne, 296, 3) <> '' then T.PutValue('G_TVAENCAISSEMENT', EnleveBlanc(Ligne, 296, 3));
        if EnleveBlanc(Ligne, 299, 3) <> '' then T.PutValue('G_TPF', EnleveBlanc(Ligne, 299, 3));
    end;
     // pour synchro on met la date de création à une date - 1 jour
    T.PutValue('G_DATECREATION', Date-1);

    T.PutValue('G_CONFIDENTIEL', '0');
    if (VersionFichier <> '') and (StrToint(VersionFichier) >= 7) then
    begin
        T.PutValue('G_CUTOFF', EnleveBlanc(Ligne, 302, 1, '-'));
        T.PutValue('G_CUTOFFPERIODE', EnleveBlanc(Ligne, 303, 3));
        T.PutValue('G_CUTOFFECHUE', EnleveBlanc(Ligne, 306, 1, '-'));
        T.PutValue('G_VISAREVISION', EnleveBlanc(Ligne, 307, 1, '-'));
        T.PutValue('G_CYCLEREVISION', EnleveBlanc(Ligne, 308, 17));
        T.PutValue('G_CUTOFFCOMPTE', EnleveBlanc(Ligne, 325, 17));
        T.PutValue('G_CONFIDENTIEL', EnleveBlanc(Ligne, 342, 1, '0'));
        //Sauvegarde de qualifiant qte  
        T.PutValue('G_QUALIFQTE1', EnleveBlanc(Ligne, 343, 3, ''));
        T.PutValue('G_QUALIFQTE2', EnleveBlanc(Ligne, 346, 3, ''));
    end;
end;


function TImportCom.OKUpdategeneraux (TG : TOB; Ligne, compte : string; var Nature : string) : Boolean;
begin
    Result := TRUE;
    if BMAJCPTE then exit;
    if (NatureFichier ='BAL') or (NatureFichier ='JRL') or
    ((OrigineFichier = 'SI') and (NatureFichier ='DOS')) then
    begin Result := FALSE; exit; end;

    if TG.GetValue('G_NATUREGENE') <> Nature then
    begin
     if (OrigineFichier = 'S1') and  (TG.GetValue('G_NATUREGENE') = 'COS') and  (Nature = 'COD') then
         Nature := 'COS'
     else
         exit;
    end;
    if (EnleveBlanc(Ligne, 63, 1) <> '') and (TG.GetValue('G_POINTABLE') <> EnleveBlanc(Ligne, 63, 1)) then exit;

    if (EnleveBlanc(Ligne, 62, 1) <> '') and (TG.GetValue('G_LETTRABLE') <> EnleveBlanc(Ligne, 62, 1)) then exit;
    if (EnleveBlanc(Ligne, 64, 1) <> '') and (TG.GetValue('G_VENTILABLE1') <> EnleveBlanc(Ligne, 64, 1)) then exit;
    if (EnleveBlanc(Ligne, 65, 1) <> '') and (TG.GetValue('G_VENTILABLE2') <> EnleveBlanc(Ligne, 65, 1)) then exit;
    if (EnleveBlanc(Ligne, 66, 1) <> '') and (TG.GetValue('G_VENTILABLE3') <> EnleveBlanc(Ligne, 66, 1)) then exit;
    if (EnleveBlanc(Ligne, 67, 1) <> '') and (TG.GetValue('G_VENTILABLE4') <> EnleveBlanc(Ligne, 67, 1)) then exit;
    if (EnleveBlanc(Ligne, 68, 1) <> '') and (TG.GetValue('G_VENTILABLE5') <> EnleveBlanc(Ligne, 68, 1)) then exit;

    if (OrigineFichier = 'S1') then
    begin
         if (Nature <> 'TID') and (Nature <> 'TIC') and (EnleveBlanc(Ligne, 62, 1) = 'X') then exit
         else
         begin
                 if  (copy (compte, 1, 1) = '6') and (Nature <> 'CHA') then exit;
                 if  (copy (compte, 1, 1) = '7') and (Nature <> 'PRO') then exit;
                 if  (copy (compte, 1, 3) = '512') and (Nature <> 'BQE') then exit;
                 if  (copy (compte, 1, 3) = '514') and (Nature <> 'BQE') then exit;
                 if  (copy (compte, 1, 2) = '53') and (Nature <> 'CAI') then exit;
                 if  (copy (compte, 1, 2) = '21') and (Nature <> 'IMO') then exit;
                 if  (copy (compte, 1, 2) = '20') and (Nature <> 'IMO') then exit;
                 if (Nature <> TG.GetValue('G_NATUREGENE')) then exit;

         end;
    end;

    if (EnleveBlanc(Ligne, 69, 3) <> '') and (TG.GetValue('G_TABLE0') <> EnleveBlanc(Ligne, 69, 3)) then exit;
    if (EnleveBlanc(Ligne, 72, 3) <> '') and (TG.GetValue('G_TABLE1') <> EnleveBlanc(Ligne, 72, 3)) then exit;
    if (EnleveBlanc(Ligne, 75, 3) <> '') and (TG.GetValue('G_TABLE2') <> EnleveBlanc(Ligne, 75, 3)) then exit;
    if (EnleveBlanc(Ligne, 78, 3) <> '') and (TG.GetValue('G_TABLE3') <> EnleveBlanc(Ligne, 78, 3)) then exit;
    if (EnleveBlanc(Ligne, 81, 3) <> '') and (TG.GetValue('G_TABLE4') <> EnleveBlanc(Ligne, 81, 3)) then exit;
    if (EnleveBlanc(Ligne, 84, 3) <> '') and (TG.GetValue('G_TABLE5') <> EnleveBlanc(Ligne, 84, 3)) then exit;
    if (EnleveBlanc(Ligne, 87, 3) <> '') and (TG.GetValue('G_TABLE6') <> EnleveBlanc(Ligne, 87, 3)) then exit;
    if (EnleveBlanc(Ligne, 90, 3) <> '') and (TG.GetValue('G_TABLE7') <> EnleveBlanc(Ligne, 90, 3)) then exit;
    if (EnleveBlanc(Ligne, 93, 3) <> '') and (TG.GetValue('G_TABLE8') <> EnleveBlanc(Ligne, 93, 3)) then exit;
    if (EnleveBlanc(Ligne, 96, 3) <> '') and (TG.GetValue('G_TABLE9') <> EnleveBlanc(Ligne, 96, 3)) then exit;
    if (EnleveBlanc(Ligne, 116, 3) <> '') and (TG.GetValue('G_SENS') <> EnleveBlanc(Ligne, 116, 3)) then exit;

    if (VersionFichier <> '') and (StrToint(VersionFichier) >= 7) then
    begin
         if (EnleveBlanc(Ligne, 343, 3, '') <> '') and (TG.GetValue('G_QUALIFQTE1') <> EnleveBlanc(Ligne, 343, 3, '')) then exit;
         if (EnleveBlanc(Ligne, 346, 3, '') <> '') and (TG.GetValue('G_QUALIFQTE2') <> EnleveBlanc(Ligne, 346, 3, '')) then exit;
    end;
       Result := FALSE;

end;

procedure TImportCom.ImporteGeneraux(Ligne: string);
var
  T, T1, T2      : TOB;
  compte,Cpte    : string;
  Nature         : string;
  j              : integer;
  SectAtt        : array [1..5] of Boolean;
  lAxeVent       : array [1..5] of Boolean; // FQ 21259 SBO 27/08/2007
begin
  T := nil;
  T2 := nil;
  InitGeneraux;
  compte := BourreOuTronque(EnleveBlanc(Ligne, 7, 17), fbGene);
  // si compte avec minuscule dedans
  Compte := UpperCase (compte);

  // application de la Table de correspondance
  Cpte := Compte;
  TraiteCorrespCpt(0, Cpte, Compte);

  for j:=1 To  length(Compte) do
      if Compte[j] in Alpha then
                OnAfficheListeCom(ERR_COMPTEALPHA + Compte,LISTEIMPORT);

  if GrandNbGene then // AJOUT ME 14-01-2005
  begin
        ChargeTobTiersGene ('SELECT * FROM GENERAUX WHERE G_GENERAL="'+Compte+'"', 'G');
        T1 := RechercheOptimise(TGen, 'G_GENERAL', compte, FALSE, fHgene);
  end
  else
        T1 := RechercheOptimise(TGen, 'G_GENERAL', compte, FALSE, fHgene);

  Nature := EnleveBlanc(Ligne, 59, 3);
  if T1 = nil then
  begin
    //T := TGenInsert.FindFirst(['G_GENERAL'], [compte], FALSE);
    T := RechercheOptimise(TGenInsert, 'G_GENERAL', compte, FALSE, fHgeneInsert);
    if T = nil then
    begin
         T := TOB.Create('GENERAUX', TGenInsert, -1);
         Remplirgeneraux (T, Ligne, compte, Nature);
         fHGeneInsert.Put(compte, TObject(T));
         { BVE 25.07.07 Rajout Ventilation Type }
         if GereVentilType then
         begin
            for j := 1 to 5 do
              SectAtt[j] := true;
            InitialiseVentilDefaut(compte,SectAtt,lAxeVent); // FQ 21259 SBO 27/08/2007
            for j:= 1 to 5 do
            begin
                if lAxeVent[j] = TRUE then // fiche 20869
                begin
                  T.PutValue ('G_VENTILABLE', 'X');
                  T.PutValue ('G_VENTILABLE'+Inttostr(j), 'X');
                end;
            end;
         end;
         { END BVE 25.07.07 }
    end
    else
    begin
       if OKUpdategeneraux (T, Ligne, compte, Nature) then
       begin
         T2 := TOB.Create('GENERAUX', TGenupdate, -1);
         Remplirgeneraux (T2, Ligne, compte, Nature);
       end;
    end;
  end
  else
  begin
       // Si réellement il y a un changement d'une info
       // on fait update
       if OKUpdategeneraux (T1, Ligne, compte, Nature) then
       begin
         T2 := TOB.Create('GENERAUX', TGenupdate, -1);
         Remplirgeneraux (T2, Ligne, compte, Nature, TRUE);
       end;
  end;

end;

procedure TImportCom.remplir_compteattente;
var
  T, T1   : TOB;
  compte  : string;
begin
  if TGen = nil then exit;
  compte := GetParamSocSecur('SO_GENATTEND', '');
  compte := BourreOuTronque(compte, fbGene);
//  T1 := TGen.FindFirst(['G_GENERAL'], [compte], FALSE);
  T1 := RechercheOptimise(TGen, 'G_GENERAL', compte, FALSE, fHgene);
  if GrandNbGene and (T1 = nil) then // AJOUT ME 14-01-2005
  begin
        ChargeTobTiersGene ('SELECT * FROM GENERAUX WHERE G_GENERAL="'+Compte+'"', 'G');
        T1 := RechercheOptimise(TGen, 'G_GENERAL', compte, FALSE, fHgene);
  end;

  if T1 = nil then
  begin
//    T := TGenInsert.FindFirst(['G_GENERAL'], [compte], FALSE);
    T := RechercheOptimise(TGenInsert, 'G_GENERAL', compte, FALSE, fHgeneInsert);
    if T = nil then
    begin
         T := TOB.Create('GENERAUX', TGenInsert, -1);
         T.PutValue('G_GENERAL', compte);
         T.PutValue('G_LIBELLE', 'COMPTE D''ATTENTE');
         T.PutValue('G_NATUREGENE', 'DIV');
         T.PutValue('G_ABREGE','COMPTE D''ATTENTE') ;
         T.PutValue('G_SENS', 'M') ;
         T.PutValue('G_VENTILABLE1', '-');
         T.PutValue('G_VENTILABLE2', '-');
         T.PutValue('G_VENTILABLE3', '-');
         T.PutValue('G_VENTILABLE4', '-');
         T.PutValue('G_VENTILABLE5', '-');
         T.PutValue('G_VENTILABLE', '-');
         T.PutValue('G_CONFIDENTIEL', '0');
         T.PutValue('G_DATECREATION', Date-1);
         fHGeneInsert.Put(Compte, TObject(T));

    end;
  end;
end;


procedure TImportCom.InitTiers (Affiche : Boolean=TRUE);  
var
  sSQL      : string;
  compteaux : string;
  ii        : integer;
  Tt        : TOB;
  Val       : string;
begin

  if TTiers = nil then
  begin
      // Récupération des indices significatifs de la TOB des comptes généraux
    Tt := TOB.Create ('TIERS',nil,-1);
    FiTOBT_LETTRABLE := Tt.GetNumChamp ('T_LETTRABLE');
    FiTOBT_NATUREAUXI := Tt.GetNumChamp ('T_NATUREAUXI');
    FiTOBT_REGIMETVA := Tt.GetNumChamp ('T_REGIMETVA');

    Tt.free;

    // si les natures des comptes différents par rapport à paramsoc
    if (Paramgene) and (OrigineFichier <> 'SI') then
    begin
      for ii:=0 to SizeRecControl do
      begin
           compteaux := BourreOuTronque(GetParamSocSecur (RecControl[ii].NomChamp, ''),RecControl[ii].LeFb);
           if (RecControl[ii].champ  = 'CPT')  and (RecControl[ii].LeFb = fbAux) then
           begin
                Val := BourreOuTronque(RecControl[ii].valeur, fbaux);
                if (BourreOuTronque(RecControl[ii].valeur, fbaux) <> compteaux) and
                (Val <> '') then
                begin
                     SetParamSoc (RecControl[ii].NomChamp, UpperCase(Val));
                     // suppression car déjà creer par planref
                     ExecuteSQL('DELETE FROM TIERS WHERE T_AUXILIAIRE="'+ compteaux+'"');
                end;
           end;
      end;
    end;

//XXXXX   if (NatureFichier = 'JRL') then GrandNbAux := TRUE;

    TTiers := TOB.Create('', nil, -1);
     // Objet de stockage en mémoire
    fHTiers := THashTable.Create();
    fHTiers.useElfHash := true;
    fHTiers.singleThread := true;

// pb chargement gros dossier + 28000 lignes    sSQL := 'SELECT * FROM TIERS ORDER BY T_AUXILIAIRE';
    if not GrandNbAux then // AJOUT ME 14-01-2005
    begin
          sSQL := 'SELECT T_AUXILIAIRE,T_TIERS,T_LETTRABLE,T_NATUREAUXI,T_MULTIDEVISE,T_LIBELLE,T_COLLECTIF, '+
          'T_EAN,T_TABLE0,T_TABLE1,T_TABLE2,T_TABLE3,T_TABLE4,T_TABLE5,T_TABLE6,T_TABLE7,T_TABLE8,T_TABLE9,'+
          'T_ADRESSE1,T_ADRESSE2,T_ADRESSE3,T_CODEPOSTAL,T_VILLE,T_PAYS,T_ABREGE,T_LANGUE,T_TELEPHONE,T_FAX,'+
          'T_REGIMETVA,T_MODEREGLE,T_COMMENTAIRE,T_NIF,T_SIRET,T_APE,T_JURIDIQUE,T_FORMEJURIDIQUE,T_TVAENCAISSEMENT,'+
          'T_DEVISE,T_CONFIDENTIEL,T_DATECREATION, T_PAYEUR, T_ISPAYEUR, T_AVOIRRBT, T_RELANCEREGLEMENT, T_RELANCETRAITE '+
          ',T_FERME, T_CORRESP1, T_CORRESP2, T_ESCOMPTE, T_REMISE, T_FACTURE, T_CREDITDEMANDE,T_CREDITACCORDE,T_CREDITPLAFOND '+
          ',R_PAYS, R_ETABBQ, R_GUICHET,R_NUMEROCOMPTE,R_CLERIB,R_DOMICILIATION,R_CODEIBAN ' +
          'FROM TIERS left outer join rib on ( t_auxiliaire=r_auxiliaire AND R_PRINCIPAL="X") '
          + ' Where  (t_natureauxi="FOU" or t_natureauxi="CLI" OR t_natureauxi="SAL" or t_natureauxi="DIV"'
          + ' or t_natureauxi="AUC" or t_natureauxi="AUD")'
          + ' ORDER BY T_AUXILIAIRE';
          ChargeTobTiersGene (sSQL, 'T');
    end;
    // si changement longueur compte
    if Changementaux then
    begin
    for ii := TTiers.detail.Count -1 downto 0 do
      begin
      compteaux := BourreOuTronque( TTiers.detail[ii].GetValue('T_AUXILIAIRE'), fbAux);
          if compteaux <> TTiers.detail[ii].GetValue('T_AUXILIAIRE') then
          begin
              ExecuteSQL('DELETE FROM TIERS WHERE T_AUXILIAIRE="'+ TTiers.detail[ii].GetValue('T_AUXILIAIRE')+'"');
              TTiers.detail[ii].PutValue('T_AUXILIAIRE', compteaux);
          end;
      end;
       TTiers.SetAllModifie(TRUE);
       TTiers.InsertOrUpdateDB(TRUE);
    end;
    TTiersupdate := TOB.Create('', nil, -1);
    TTiersInsert := TOB.Create('', nil, -1);

     fHTiersInsert := THashTable.Create();
     fHTiersInsert.useElfHash := true;
     fHTiersInsert.singleThread := true;
    if Affiche then OKexercice := OnAfficheListeCom('Comptes Tiers',LISTEIMPORT);

  end;
end;

procedure TImportCom.MajChamp (var T: TOB; Champ, Valeur : string; dd : double=0; Numerique : Boolean=FALSE);  // fiche 10441
begin
 if not Numerique then
 begin
// GP       if (not BBlanc) or (BBlanc and ((Pos (Champ, ModifTiers) = 0) and (ModifTiers <> '<<Tous>>'))) then
       if (not BBlanc) or (BBlanc and ((Pos (Champ, ModifTiers) = 0) and (ModifTiers <> Traduirememoire('<<Tous>>')))) then
       begin
           if Valeur <> '' then T.PutValue(Champ, Valeur);
       end
       else
           T.PutValue(Champ, Valeur);
 end
 else
 begin
// GP       if (not BBlanc) or (BBlanc and ((Pos (Champ, ModifTiers) = 0) and (ModifTiers <> '<<Tous>>'))) then
       if (not BBlanc) or (BBlanc and ((Pos (Champ, ModifTiers) = 0) and (ModifTiers <> Traduirememoire('<<Tous>>')))) then
       begin
           if dd <> 0 then T.PutValue(Champ, dd);
       end
       else
           T.PutValue(Champ, dd);
 end;
end;

procedure TImportCom.RemplissageTiers (var T : Tob; compteauxi, Ligne : string; update : Boolean=TRUE; Casupdate : Boolean=FALSE);
var
Collectif, CodeTva : string;
begin
    if (not update) and ((NatureFichier ='BAL') or (NatureFichier ='JRL')) then exit;

        // si compte avec minuscule dedans
    compteauxi := UpperCase (compteauxi);
    Collectif := BourreOuTronque(EnleveBlanc(Ligne, 63, 17), fbGene);
    if Casupdate then
    begin
            if (T.GetValue('T_AUXILIAIRE') = '') then T.PutValue('T_AUXILIAIRE', compteauxi);
            if (T.GetValue('T_TIERS') = '') then T.PutValue('T_TIERS', EnleveBlanc(Ligne, 7, 17));
            if BMAJCPTE then
            begin
                    MajChamp(T, 'T_COLLECTIF', Collectif);
                    MajChamp(T, 'T_LIBELLE', VireTouteSub(EnleveBlanc(Ligne, 24, 35, '.'),'"'));
                    MajChamp(T, 'T_EAN', EnleveBlanc(Ligne, 80, 17));
                    MajChamp(T, 'T_TABLE0', Veriftablelibre('T00', EnleveBlanc(Ligne, 97, 17)));
                    MajChamp(T, 'T_TABLE1', Veriftablelibre('T01', EnleveBlanc(Ligne, 114, 17)));
                    MajChamp(T, 'T_TABLE2', Veriftablelibre('T02', EnleveBlanc(Ligne, 131, 17)));
                    MajChamp(T, 'T_TABLE3', Veriftablelibre('T03', EnleveBlanc(Ligne, 148, 17)));
                    MajChamp(T, 'T_TABLE4', Veriftablelibre('T04', EnleveBlanc(Ligne, 165, 17)));
                    MajChamp(T, 'T_TABLE5', Veriftablelibre('T05', EnleveBlanc(Ligne, 182, 17)));
                    MajChamp(T, 'T_TABLE6', Veriftablelibre('T06', EnleveBlanc(Ligne, 199, 17)));
                    MajChamp(T, 'T_TABLE7', Veriftablelibre('T07', EnleveBlanc(Ligne, 216, 17)));
                    MajChamp(T, 'T_TABLE8', Veriftablelibre('T08', EnleveBlanc(Ligne, 233, 17)));
                    MajChamp(T, 'T_TABLE9', Veriftablelibre('T09', EnleveBlanc(Ligne, 250, 17)));
                    MajChamp(T, 'T_ADRESSE1', VireTouteSub(EnleveBlanc(Ligne, 267, 35),'"'));
                    MajChamp(T, 'T_ADRESSE2', VireTouteSub(EnleveBlanc(Ligne, 302, 35),'"'));
                    MajChamp(T, 'T_ADRESSE3', VireTouteSub(EnleveBlanc(Ligne, 337, 35),'"'));
                    MajChamp(T, 'T_CODEPOSTAL', EnleveBlanc(Ligne, 372, 9));
                    MajChamp(T, 'T_VILLE', EnleveBlanc(Ligne, 381, 35));
                    MajChamp(T, 'T_PAYS', EnleveBlanc(Ligne, 463, 3));
                    MajChamp(T, 'T_ABREGE', VireTouteSub(EnleveBlanc(Ligne, 466, 17),'"'));
                    MajChamp(T, 'T_LANGUE', EnleveBlanc(Ligne, 483, 3));
                    MajChamp(T, 'T_MULTIDEVISE', EnleveBlanc(Ligne, 486, 1));
                    MajChamp(T, 'T_DEVISE', EnleveBlanc(Ligne, 487, 3));
                    MajChamp(T, 'T_TELEPHONE', EnleveBlanc(Ligne, 490, 25));
                    MajChamp(T, 'T_FAX', EnleveBlanc(Ligne, 515, 25));
                    MajChamp(T, 'T_REGIMETVA', EnleveBlanc(Ligne, 540, 3, ''));
                    MajChamp(T, 'T_MODEREGLE', EnleveBlanc(Ligne, 543, 3, ''));
                    MajChamp(T, 'T_COMMENTAIRE', EnleveBlanc(Ligne, 546, 35));
                    MajChamp(T, 'T_NIF', EnleveBlanc(Ligne, 581, 17));
                    MajChamp(T, 'T_SIRET', EnleveBlanc(Ligne, 598, 17));
                    MajChamp(T, 'T_APE', EnleveBlanc(Ligne, 615, 5));
                    MajChamp(T, 'T_FORMEJURIDIQUE', EnleveBlanc(Ligne, 854, 3));
                    if (StrToint(VersionFichier) >= 7) then
                    begin
                        MajChamp(T, 'T_PRENOM', EnleveBlanc(Ligne, 620, 35));
                        MajChamp(T, 'T_JURIDIQUE', EnleveBlanc(Ligne, 978, 3));
                    end;
                    MajChamp(T, 'T_RELANCEREGLEMENT', EnleveBlanc(Ligne, 880, 3));
                    MajChamp(T, 'T_RELANCETRAITE', EnleveBlanc(Ligne, 883, 3));
                    CodeTva :=  EnleveBlanc(Ligne, 858, 3); // Fiche 10425
                    if (CodeTva <> '') and ((CodeTva = 'TD') or (CodeTva = 'TE') or (CodeTva = 'TM'))then
                     MajChamp(T, 'T_TVAENCAISSEMENT', CodeTva);
                    MajChamp(T, 'T_PAYEUR', EnleveBlanc(Ligne, 861, 17));
                    MajChamp(T, 'T_ISPAYEUR', EnleveBlanc(Ligne, 878, 1, '')); // fiche 10570
                    MajChamp(T, 'T_AVOIRRBT', EnleveBlanc(Ligne, 879, 1, ''));
                    MajChamp(T, 'T_CORRESP1', EnleveBlanc(Ligne, 887, 17, ''));
                    MajChamp(T, 'T_CORRESP2', EnleveBlanc(Ligne, 904, 17, ''));
                    MajChamp(T, 'T_ESCOMPTE', '', Valeur(StPoint(EnleveBlanc(Ligne, 921, 20, '0'))), TRUE);
                    MajChamp(T, 'T_REMISE', '', Valeur(StPoint(EnleveBlanc(Ligne, 941, 20, '0'))), TRUE);
                    MajChamp(T, 'T_FACTURE', EnleveBlanc(Ligne, 961, 17, compteauxi));
                    // fiche com 31027
                    MajChamp(T, 'T_CREDITDEMANDE', '', Valeur(StPoint(EnleveBlanc(Ligne, 981, 20, '0'))), TRUE);
                    MajChamp(T, 'T_CREDITACCORDE', '', Valeur(StPoint(EnleveBlanc(Ligne, 1001, 20, '0'))), TRUE);
                    MajChamp(T, 'T_CREDITPLAFOND', '', Valeur(StPoint(EnleveBlanc(Ligne, 1021, 20, '0'))), TRUE);
                                   // Fiche 10638
                    MajChamp(T, 'T_CODEIMPORT', EnleveBlanc(Ligne, 7, 17));
            end;
    end
    else
    begin
            T.PutValue('T_AUXILIAIRE', compteauxi);
            T.PutValue('T_TIERS', EnleveBlanc(Ligne, 7, 17));
            if EnleveBlanc(Ligne, 24, 35) = '' then
               T.PutValue('T_LIBELLE', '.')
            else
               T.PutValue('T_LIBELLE', VireTouteSub(EnleveBlanc(Ligne, 24, 35),'"'));
            Collectif := BourreOuTronque(EnleveBlanc(Ligne, 63, 17), fbGene);
            T.PutValue('T_COLLECTIF', Collectif);
            T.PutValue('T_EAN', EnleveBlanc(Ligne, 80, 17));
            // Controle de l'existance des tables libres  fiche 10330
            T.PutValue('T_TABLE0', Veriftablelibre('T00', EnleveBlanc(Ligne, 97, 17)));
            T.PutValue('T_TABLE1', Veriftablelibre('T01', EnleveBlanc(Ligne, 114, 17)));
            T.PutValue('T_TABLE2', Veriftablelibre('T02', EnleveBlanc(Ligne, 131, 17)));
            T.PutValue('T_TABLE3', Veriftablelibre('T03', EnleveBlanc(Ligne, 148, 17)));
            T.PutValue('T_TABLE4', Veriftablelibre('T04', EnleveBlanc(Ligne, 165, 17)));
            T.PutValue('T_TABLE5', Veriftablelibre('T05', EnleveBlanc(Ligne, 182, 17)));
            T.PutValue('T_TABLE6', Veriftablelibre('T06', EnleveBlanc(Ligne, 199, 17)));
            T.PutValue('T_TABLE7', Veriftablelibre('T07', EnleveBlanc(Ligne, 216, 17)));
            T.PutValue('T_TABLE8', Veriftablelibre('T08', EnleveBlanc(Ligne, 233, 17)));
            T.PutValue('T_TABLE9', Veriftablelibre('T09', EnleveBlanc(Ligne, 250, 17)));

            T.PutValue('T_ADRESSE1', VireTouteSub(EnleveBlanc(Ligne, 267, 35),'"'));
            T.PutValue('T_ADRESSE2', VireTouteSub(EnleveBlanc(Ligne, 302, 35),'"'));
            T.PutValue('T_ADRESSE3', VireTouteSub(EnleveBlanc(Ligne, 337, 35),'"'));
            T.PutValue('T_CODEPOSTAL', EnleveBlanc(Ligne, 372, 9));
            T.PutValue('T_VILLE', EnleveBlanc(Ligne, 381, 35));
            T.PutValue('T_PAYS', EnleveBlanc(Ligne, 463, 3));
            if EnleveBlanc(Ligne, 466, 17) <> '' then
               T.PutValue('T_ABREGE', VireTouteSub(EnleveBlanc(Ligne, 466, 17),'"'))
            else
               T.PutValue('T_ABREGE', VireTouteSub(EnleveBlanc(Ligne, 24, 17),'"'));
            T.PutValue('T_LANGUE', EnleveBlanc(Ligne, 483, 3));
            T.PutValue('T_TELEPHONE', EnleveBlanc(Ligne, 490, 25));
            T.PutValue('T_FAX', EnleveBlanc(Ligne, 515, 25));
            if (Regdefaut = '') and (EnleveBlanc(Ligne, 540, 3) <> '') then
                 T.PutValue('T_REGIMETVA', EnleveBlanc(Ligne, 540, 3))
            else
                 T.PutValue('T_REGIMETVA', RendRegTvadefaut);
            if (Mrdefaut = '') and (EnleveBlanc(Ligne, 543, 3) <> '') then
                 T.PutValue('T_MODEREGLE', EnleveBlanc(Ligne, 543, 3))
            else
                T.PutValue('T_MODEREGLE', RendMRDefaut);

            T.PutValue('T_COMMENTAIRE', EnleveBlanc(Ligne, 546, 35));
            T.PutValue('T_NIF', EnleveBlanc(Ligne, 581, 17));
            T.PutValue('T_SIRET', EnleveBlanc(Ligne, 598, 17));
            T.PutValue('T_APE', EnleveBlanc(Ligne, 615, 5));
            if (StrToint(VersionFichier) >= 7)  and ( EnleveBlanc(Ligne, 620, 35) <> '') then
                        T.PutValue('T_PRENOM', EnleveBlanc(Ligne, 620, 35));

            if (StrToint(VersionFichier) >= 7)  and ( EnleveBlanc(Ligne, 978, 3) <> '') then
                        T.PutValue('T_JURIDIQUE', EnleveBlanc(Ligne, 978, 3));

            T.PutValue('T_FORMEJURIDIQUE', EnleveBlanc(Ligne, 854, 3));
            T.PutValue('T_CREERPAR', 'COM');

            CodeTva :=  EnleveBlanc(Ligne, 858, 3);
            if (CodeTva <> '') and ((CodeTva = 'TD') or (CodeTva = 'TE') or (CodeTva = 'TM'))then
               T.PutValue('T_TVAENCAISSEMENT', CodeTva)
            else
               T.PutValue('T_TVAENCAISSEMENT', GetParamSocSecur('SO_CODETVADEFAUT', ''));

             // pour synchro on met la date de création à une date - 1 jour
            T.PutValue('T_DATECREATION', Date-1);
            T.PutValue('T_NATUREAUXI', EnleveBlanc(Ligne, 59, 3));
            T.PutValue('T_LETTRABLE', EnleveBlanc(Ligne, 62, 1));
            if (OrigineFichier <> 'S1') then
            begin
               T.PutValue('T_MULTIDEVISE', EnleveBlanc(Ligne, 486, 1));
               T.PutValue('T_DEVISE', EnleveBlanc(Ligne, 487, 3));
               if (EnleveBlanc(Ligne, 486, 1) = '') and (EnleveBlanc(Ligne, 487, 3) = '')  then
                     T.PutValue('T_MULTIDEVISE', 'X');
            end;

             // ajout me 18-08-2005 fiche 10149
            if (EnleveBlanc(Ligne, 861, 17) <> '') then
               T.PutValue('T_PAYEUR', EnleveBlanc(Ligne, 861, 17));
            if (EnleveBlanc(Ligne, 878, 1) <> '') then
               T.PutValue('T_ISPAYEUR', EnleveBlanc(Ligne, 878, 1, '-'));
            if (EnleveBlanc(Ligne, 879, 1) <> '') then
               T.PutValue('T_AVOIRRBT', EnleveBlanc(Ligne, 879, 1, '-'));
            // ajout me 26/08/2005
            if (EnleveBlanc(Ligne, 880, 3) <> '') then
               T.PutValue('T_RELANCEREGLEMENT', EnleveBlanc(Ligne, 880, 3));
            if (EnleveBlanc(Ligne, 883, 3) <> '') then
               T.PutValue('T_RELANCETRAITE', EnleveBlanc(Ligne, 883, 3));

            T.PutValue('T_CONFIDENTIEL', EnleveBlanc(Ligne, 886, 1, '0'));
            if (StrToint(VersionFichier) >= 8) then
            begin
                T.PutValue('T_CORRESP1', EnleveBlanc(Ligne, 887, 17, ''));
                T.PutValue('T_CORRESP2', EnleveBlanc(Ligne, 904, 17, ''));
                T.PutValue('T_ESCOMPTE', Valeur(StPoint(EnleveBlanc(Ligne, 921, 20, '0'))));
                T.PutValue('T_REMISE', Valeur(StPoint(EnleveBlanc(Ligne, 941, 20, '0'))));
                T.PutValue('T_FACTURE', EnleveBlanc(Ligne, 961, 17, compteauxi));
                // fiche com 31027
                T.PutValue('T_CREDITDEMANDE', Valeur(StPoint(EnleveBlanc(Ligne, 981, 20, '0'))));
                T.PutValue('T_CREDITACCORDE', Valeur(StPoint(EnleveBlanc(Ligne, 1001, 20, '0'))));
                T.PutValue('T_CREDITPLAFOND', Valeur(StPoint(EnleveBlanc(Ligne, 1021, 20, '0'))));
            end
            else
                T.PutValue('T_FACTURE',compteauxi);
                // Fiche 10638
            T.PutValue('T_CODEIMPORT', EnleveBlanc(Ligne, 7, 17));
    end;
end;

function TImportCom.OKUpdateTiers (T : Tob; compteauxi, Ligne : string) : Boolean;
begin
    Result := TRUE;
    if BMAJCPTE then exit;

    if (NatureFichier ='BAL') or (NatureFichier ='JRL') or
    ((OrigineFichier = 'SI') and (NatureFichier ='DOS')) then
    begin Result := FALSE; exit; end;

    if T.GetValue('T_EAN') <> EnleveBlanc(Ligne, 80, 17) then exit;
    if T.GetValue('T_TABLE0') <> EnleveBlanc(Ligne, 97, 17) then exit;
    if T.GetValue('T_TABLE1') <> EnleveBlanc(Ligne, 114, 17) then exit;
    if T.GetValue('T_TABLE2') <> EnleveBlanc(Ligne, 131, 17) then exit;
    if T.GetValue('T_TABLE3') <> EnleveBlanc(Ligne, 148, 17) then exit;
    if T.GetValue('T_TABLE4') <> EnleveBlanc(Ligne, 165, 17) then exit;
    if T.GetValue('T_TABLE5') <> EnleveBlanc(Ligne, 182, 17) then exit;
    if T.GetValue('T_TABLE6') <> EnleveBlanc(Ligne, 199, 17) then exit;
    if T.GetValue('T_TABLE7') <> EnleveBlanc(Ligne, 216, 17) then exit;
    if T.GetValue('T_TABLE8') <> EnleveBlanc(Ligne, 233, 17) then exit;
    if T.GetValue('T_TABLE9') <> EnleveBlanc(Ligne, 250, 17) then exit;
    if T.GetValue('T_ADRESSE1') <> EnleveBlanc(Ligne, 267, 35) then exit;
    if T.GetValue('T_ADRESSE2') <> EnleveBlanc(Ligne, 302, 35) then exit;
    if T.GetValue('T_ADRESSE3') <> EnleveBlanc(Ligne, 337, 35) then exit;
    if T.GetValue('T_CODEPOSTAL') <> EnleveBlanc(Ligne, 372, 9) then exit;
    if T.GetValue('T_VILLE') <> EnleveBlanc(Ligne, 381, 35) then exit;
    if T.GetValue('T_PAYS') <> EnleveBlanc(Ligne, 463, 3) then exit;
    if T.GetValue('T_ABREGE') <> VireTouteSub(EnleveBlanc(Ligne, 466, 17),'"') then exit;
    if T.GetValue('T_LANGUE') <> EnleveBlanc(Ligne, 483, 3) then exit;
    if T.GetValue('T_TELEPHONE') <> EnleveBlanc(Ligne, 490, 25) then exit;
    if T.GetValue('T_FAX') <> EnleveBlanc(Ligne, 515, 25) then exit;
    if T.GetValue('T_REGIMETVA') <> EnleveBlanc(Ligne, 540, 3) then exit;
    if T.GetValue('T_MODEREGLE') <> EnleveBlanc(Ligne, 543, 3) then exit;
    if T.GetValue('T_COMMENTAIRE') <> EnleveBlanc(Ligne, 546, 35) then exit;
    if T.GetValue('T_NIF') <> EnleveBlanc(Ligne, 581, 17) then exit;
    if T.GetValue('T_SIRET') <> EnleveBlanc(Ligne, 598, 17) then exit;
    if T.GetValue('T_APE') <> EnleveBlanc(Ligne, 615, 5) then exit;
    if T.GetValue('T_FORMEJURIDIQUE') <> EnleveBlanc(Ligne, 854, 3) then exit;
    if T.GetValue('T_TVAENCAISSEMENT') <> EnleveBlanc(Ligne, 858, 3) then exit;

    if T.GetValue('T_NATUREAUXI') <> EnleveBlanc(Ligne, 59, 3) then exit;
    if T.GetValue('T_LETTRABLE') <> EnleveBlanc(Ligne, 62, 1) then exit;
    if T.GetValue('T_MULTIDEVISE') <> EnleveBlanc(Ligne, 486, 1) then exit;
    if T.GetValue('T_DEVISE') <> EnleveBlanc(Ligne, 487, 3) then exit;
    Result := FALSE;
end;


procedure TImportCom.ImporteTiers(Ligne: string);
var
  T, T1, T2              : TOB;
  compteauxi,NatAux,sSQL : string;
  Cpte                   : string;
  QDos                   : TQuery;
begin
  T := nil;
  T2 := nil;
  InitTiers;

  compteauxi := BourreOuTronque(EnleveBlanc(Ligne, 7, 17), fbAux);
  compteauxi :=FindEtReplace(compteauxi,'''','0',TRUE) ;
    // si compte avec minuscule dedans
  Compteauxi := UpperCase (compteauxi);

  // application de la Table de correspondance
  Cpte := Compteauxi;
  TraiteCorrespCpt(1, Cpte, Compteauxi);

//  T1 := TTiers.FindFirst(['T_AUXILIAIRE'], [compteauxi], FALSE);
  T1 := RechercheOptimise(TTiers, 'T_AUXILIAIRE', compteauxi, FALSE, fHTiers);
  if (T1 = nil) then
  begin
          if GrandNbAux then // Fiche 10589 Pb lenteur TRT  A PORTER sur 7XX
          begin
                sSQL := 'SELECT ##TOP 200## T_AUXILIAIRE,T_TIERS,T_LETTRABLE,T_NATUREAUXI,T_MULTIDEVISE,T_LIBELLE,T_COLLECTIF, '+
                'T_EAN,T_TABLE0,T_TABLE1,T_TABLE2,T_TABLE3,T_TABLE4,T_TABLE5,T_TABLE6,T_TABLE7,T_TABLE8,T_TABLE9,'+
                'T_ADRESSE1,T_ADRESSE2,T_ADRESSE3,T_CODEPOSTAL,T_VILLE,T_PAYS,T_ABREGE,T_LANGUE,T_TELEPHONE,T_FAX,'+
                'T_REGIMETVA,T_MODEREGLE,T_COMMENTAIRE,T_NIF,T_SIRET,T_APE,T_JURIDIQUE,T_FORMEJURIDIQUE,T_TVAENCAISSEMENT,'+
                'T_DEVISE,T_CONFIDENTIEL,T_DATECREATION, T_PAYEUR, T_ISPAYEUR, T_AVOIRRBT, T_RELANCEREGLEMENT, T_RELANCETRAITE '+
                ',T_FERME, T_CORRESP1, T_CORRESP2, T_ESCOMPTE, T_REMISE, T_FACTURE, T_CREDITDEMANDE,T_CREDITACCORDE,T_CREDITPLAFOND,T_CODEIMPORT '+
                ',R_PAYS, R_ETABBQ, R_GUICHET,R_NUMEROCOMPTE,R_CLERIB,R_DOMICILIATION,R_CODEIBAN ' +
                'FROM TIERS left outer join rib on ( t_auxiliaire=r_auxiliaire AND R_PRINCIPAL="X") ' +
                ' Where  T_AUXILIAIRE >= "'+compteauxi+ '" ORDER BY T_AUXILIAIRE';
                ChargeTobTiersGene (sSQL, 'T');
                T1 := RechercheOptimise(TTiers, 'T_AUXILIAIRE', compteauxi, FALSE, fHTiers);
          end
          else
          begin
                 NatAux := EnleveBlanc(Ligne, 59, 3);
                 if (NatAux <> 'FOU') and (NatAux <> 'CLI') and (NatAux <> 'SAL') and (NatAux <> 'DIV') and
                 (NatAux <> 'AUC') and (NatAux <> 'AUD') and (not TiersAutre) then
                 begin
                      sSQL := 'SELECT T_AUXILIAIRE,T_TIERS,T_LETTRABLE,T_NATUREAUXI,T_MULTIDEVISE,T_LIBELLE,T_COLLECTIF, '+
                      'T_EAN,T_TABLE0,T_TABLE1,T_TABLE2,T_TABLE3,T_TABLE4,T_TABLE5,T_TABLE6,T_TABLE7,T_TABLE8,T_TABLE9,'+
                      'T_ADRESSE1,T_ADRESSE2,T_ADRESSE3,T_CODEPOSTAL,T_VILLE,T_PAYS,T_ABREGE,T_LANGUE,T_TELEPHONE,T_FAX,'+
                      'T_REGIMETVA,T_MODEREGLE,T_COMMENTAIRE,T_NIF,T_SIRET,T_APE,T_JURIDIQUE,T_FORMEJURIDIQUE,T_TVAENCAISSEMENT,'+
                      'T_DEVISE,T_CONFIDENTIEL,T_DATECREATION, T_PAYEUR, T_ISPAYEUR, T_AVOIRRBT, T_RELANCEREGLEMENT, T_RELANCETRAITE '+
                      ',T_FERME, T_CORRESP1, T_CORRESP2, T_ESCOMPTE, T_REMISE, T_FACTURE, T_CREDITDEMANDE,T_CREDITACCORDE,T_CREDITPLAFOND,T_CODEIMPORT '+
                      ',R_PAYS, R_ETABBQ, R_GUICHET,R_NUMEROCOMPTE,R_CLERIB,R_DOMICILIATION,R_CODEIBAN ' +
                      'FROM TIERS left outer join rib on ( t_auxiliaire=r_auxiliaire AND R_PRINCIPAL="X") '
                      + ' Where  t_natureauxi<>"FOU" and t_natureauxi<>"CLI" and t_natureauxi<>"SAL" and t_natureauxi<>"DIV"'
                      + ' and t_natureauxi<>"AUC" and t_natureauxi<>"AUD"'
                      + ' ORDER BY T_AUXILIAIRE';

                      ChargeTobTiersGene (sSQL, 'T');
                      TiersAutre := TRUE;
                      T1 := RechercheOptimise(TTiers, 'T_AUXILIAIRE', compteauxi, FALSE, fHTiers);
                 end;
          end;
          //T1 := TTiers.FindFirst(['T_AUXILIAIRE'], [compteauxi], FALSE);
  end;
  // Pour PB affectation rib en automatique
  sSQL := (* PB SelectDB pour tiers update met tous les champs et écrase en update les donnée
  'SELECT T_AUXILIAIRE,T_TIERS,T_LETTRABLE,T_NATUREAUXI,T_MULTIDEVISE,T_LIBELLE,T_COLLECTIF, '+
  'T_EAN,T_TABLE0,T_TABLE1,T_TABLE2,T_TABLE3,T_TABLE4,T_TABLE5,T_TABLE6,T_TABLE7,T_TABLE8,T_TABLE9,'+
  'T_ADRESSE1,T_ADRESSE2,T_ADRESSE3,T_CODEPOSTAL,T_VILLE,T_PAYS,T_ABREGE,T_LANGUE,T_TELEPHONE,T_FAX,'+
  'T_REGIMETVA,T_MODEREGLE,T_COMMENTAIRE,T_NIF,T_SIRET,T_APE,T_JURIDIQUE,T_FORMEJURIDIQUE,T_TVAENCAISSEMENT,'+
  'T_DEVISE,T_CONFIDENTIEL,T_DATECREATION, T_PAYEUR, T_ISPAYEUR, T_AVOIRRBT, T_RELANCEREGLEMENT, T_RELANCETRAITE '+
  ',R_PAYS, R_ETABBQ, R_GUICHET,R_NUMEROCOMPTE,R_CLERIB,R_DOMICILIATION,R_CODEIBAN ' +
  'FROM TIERS left outer join rib on ( t_auxiliaire=r_auxiliaire AND R_PRINCIPAL="X") ' + *)

  'SELECT  * FROM TIERS left outer join rib on ( t_auxiliaire=r_auxiliaire AND R_PRINCIPAL="X") ' +
  ' Where  T_AUXILIAIRE="'+compteauxi+ '"';

  if T1 = nil then
  begin
    //T := TTiersInsert.FindFirst(['T_AUXILIAIRE'], [compteauxi], FALSE);
    T := RechercheOptimise(TTiersInsert, 'T_AUXILIAIRE', compteauxi, FALSE, fHTiersInsert);
    if T = nil then
    begin
         T := TOB.Create('TIERS', TTiersInsert, -1);
         RemplissageTiers (T, compteauxi, Ligne, TRUE, FALSE);

         if (OrigineFichier = 'S1') then
         begin
             T.PutValue('T_MULTIDEVISE', EnleveBlanc(Ligne, 486, 1));
             T.PutValue('T_DEVISE', '');
             if EnleveBlanc(Ligne, 486, 1) <> 'X' then T.PutValue('T_MULTIDEVISE', 'X');
         end;
         fHTiersInsert.Put(compteauxi, TObject(T));
    end
    else
    begin
         if OKUpdateTiers (T, compteauxi, Ligne) then
         begin
              T2 := TOB.Create('TIERS', TTiersupdate, -1);
// Pb affectation rib en automatique         if T2.SelectDB('"' + compteauxi + '"', nil, false) then
                QDos := OpenSQL(sSQL, true);
                if T2.SelectDB('', QDos) then
                begin
                     Ferme(QDos);
                     RemplissageTiers (T2, compteauxi, Ligne, FALSE, TRUE);
                end
                else
                Ferme(QDos);
              if (OrigineFichier = 'S1') then
              begin
                   T2.PutValue('T_MULTIDEVISE', T.getValue('T_MULTIDEVISE'));
                   T2.PutValue('T_DEVISE', T.getValue('T_DEVISE'));
              end;
         end;
    end;
  end
  else
  begin
       // s'il y a une modification d'un champ on met à jour tiers
       if OKUpdateTiers (T1, compteauxi, Ligne) then
       begin
         T2 := TOB.Create('TIERS', TTiersupdate, -1);
// Pb affectation rib en automatique        if T2.SelectDB('"' + compteauxi + '"', nil, false) then
         QDos := OpenSQL(sSQL, true);
         if T2.SelectDB('', QDos) then
         begin
                     Ferme(QDos);
                     RemplissageTiers (T2, compteauxi, Ligne, TRUE, TRUE);
         end
         else Ferme(QDos);

         if (OrigineFichier = 'S1') then
         begin
             T2.PutValue('T_MULTIDEVISE', T1.getValue('T_MULTIDEVISE'));
             T2.PutValue('T_DEVISE', T1.getValue('T_DEVISE'));
         end;

       end;
  end;
  // AJOUT ME 14-01-2005
  if (TTiers.detail.count > MaxNombre) or (TTiersInsert.detail.count > MaxNombre)
  or (TTiersupdate.detail.count > MaxNombre) then
  begin
       GrandNbAux := TRUE;
       InsertTiers;
  end;
end;

// Import Tiers CCIMPEX fiche 10134
procedure TImportCom.ImporteTiersAncienFormat(Ligne: string);
var
  T,Tr,Tc, Ti            : TOB;
  compteauxi,sSQL        : string;
  Cpte,Collectif         : string;
  ExisteAuxi             : Boolean;
    Procedure AjoutUpdateRib;
    begin
             if not T.FieldExists ('R_PAYS') then
             begin
                  T.AddChampSupValeur ('R_PAYS',  Tr.GetValue ('R_PAYS'));
                  T.AddChampSupValeur ('R_ETABBQ', Tr.GetValue ('R_ETABBQ'));
                  T.AddChampSupValeur ('R_GUICHET', Tr.GetValue ('R_GUICHET'));
                  T.AddChampSupValeur ('R_NUMEROCOMPTE', Tr.GetValue ('R_NUMEROCOMPTE'));
                  T.AddChampSupValeur ('R_CLERIB', Tr.GetValue ('R_CLERIB'));
                  T.AddChampSupValeur ('R_DOMICILIATION', Tr.GetValue ('R_DOMICILIATION'));
                  T.AddChampSupValeur ('R_CODEIBAN', Tr.GetValue ('R_CODEIBAN'));
              end
              else
              begin
                  T.putValue ('R_PAYS',  Tr.GetValue ('R_PAYS'));
                  T.putValue ('R_ETABBQ', Tr.GetValue ('R_ETABBQ'));
                  T.putValue ('R_GUICHET', Tr.GetValue ('R_GUICHET'));
                  T.putValue ('R_NUMEROCOMPTE', Tr.GetValue ('R_NUMEROCOMPTE'));
                  T.putValue ('R_CLERIB', Tr.GetValue ('R_CLERIB'));
                  T.putValue ('R_DOMICILIATION', Tr.GetValue ('R_DOMICILIATION'));
                  T.putValue ('R_CODEIBAN', Tr.GetValue ('R_CODEIBAN'));
              end;
    end;
    procedure MajRIB;
    begin
            if EnleveBlanc(Ligne, 310, 11) <> '' then
            begin
                 if TRibCAU = nil then
                    TRibCAU := TOB.Create('', nil, -1);

                 if TRibCAU.FindFirst(['R_AUXILIAIRE'], [compteauxi], FALSE) = nil then
                 begin
                       Tr := TOB.Create('RIB', TRibCAU, -1);
                       Tr.PutValue('R_AUXILIAIRE', Compteauxi);
                       Tr.PutValue('R_DOMICILIATION', EnleveBlanc(Ligne, 276, 24));
                       Tr.PutValue('R_ETABBQ', EnleveBlanc(Ligne, 300, 5));
                       Tr.PutValue('R_GUICHET', EnleveBlanc(Ligne, 305, 5));
                       Tr.PutValue('R_NUMEROCOMPTE', EnleveBlanc(Ligne, 310, 11));
                       Tr.PutValue('R_CLERIB', EnleveBlanc(Ligne, 321, 2));
                       // modif pour produflex
                       Tr.PutValue('R_PAYS', EnleveBlanc(Ligne, 323, 3));
                       if EnleveBlanc(Ligne, 323, 3) = 'FR' then
                        Tr.PutValue('R_PAYS', 'FRA');

                       Tr.PutValue('R_PRINCIPAL', EnleveBlanc(Ligne, 717, 1));
                       Tr.PutValue ('R_NUMERORIB', 1);
                        Tr.InsertDB(nil, TRUE);
                 end;
            end;

    end;
begin
// modif pour produflex
  InitTiers;

  if TTiersCAU = nil then
  begin
    TTiersCAU := TOB.Create('', nil, -1);
    OnAfficheListeCom('Import des tiers' ,LISTEIMPORT);
  end;

  compteauxi := BourreOuTronque(EnleveBlanc(Ligne, 7, 17), fbAux);
  compteauxi :=FindEtReplace(compteauxi,'''','0',TRUE) ;
    // si compte avec minuscule dedans
  Compteauxi := UpperCase (compteauxi);

  // application de la Table de correspondance
  Cpte := Compteauxi;
  TraiteCorrespCpt(1, Cpte, Compteauxi);

  T := TOB.Create('TIERS', TTiersCAU, -1);
  ExisteAuxi := T.SelectDB('"' + compteauxi + '"', nil, false);


  if (not BMAJCPTE) and  ExisteAuxi then
  begin
        sSQL := 'SELECT T_AUXILIAIRE,T_TIERS,T_LETTRABLE,T_NATUREAUXI,T_MULTIDEVISE,T_LIBELLE,T_COLLECTIF, '+
        'T_EAN,T_TABLE0,T_TABLE1,T_TABLE2,T_TABLE3,T_TABLE4,T_TABLE5,T_TABLE6,T_TABLE7,T_TABLE8,T_TABLE9,'+
        'T_ADRESSE1,T_ADRESSE2,T_ADRESSE3,T_CODEPOSTAL,T_VILLE,T_PAYS,T_ABREGE,T_LANGUE,T_TELEPHONE,T_FAX,'+
        'T_REGIMETVA,T_MODEREGLE,T_COMMENTAIRE,T_NIF,T_SIRET,T_APE,T_JURIDIQUE,T_FORMEJURIDIQUE,T_TVAENCAISSEMENT,'+
        'T_DEVISE,T_CONFIDENTIEL,T_DATECREATION, T_PAYEUR, T_ISPAYEUR, T_AVOIRRBT, T_RELANCEREGLEMENT, T_RELANCETRAITE '+
        ',T_FERME, R_PAYS, R_ETABBQ, R_GUICHET,R_NUMEROCOMPTE,R_CLERIB,R_DOMICILIATION,R_CODEIBAN ' +
        'FROM TIERS left outer join rib on ( t_auxiliaire=r_auxiliaire) '
        + ' Where  (t_natureauxi="FOU" or t_natureauxi="CLI" OR t_natureauxi="SAL" or t_natureauxi="DIV"'
        + ' or t_natureauxi="AUC" or t_natureauxi="AUD")'
        + 'AND T_AUXILIAIRE="'+compteauxi+'"';
        ChargeTobTiersGene (sSQL, 'T');
        exit;
  end;


  if not ExisteAuxi then
  begin
            T.PutValue('T_AUXILIAIRE', compteauxi);
            T.PutValue('T_TIERS', EnleveBlanc(Ligne, 7, 17));
            if EnleveBlanc(Ligne, 24, 35) = '' then
               T.PutValue('T_LIBELLE', '.')
            else
            begin
               if EnleveBlanc(Ligne, 24, 35) <> '' then
                  T.PutValue('T_LIBELLE', VireTouteSub(EnleveBlanc(Ligne, 24, 35),'"'));
            end;
            T.PutValue('T_NATUREAUXI', EnleveBlanc(Ligne, 59, 3));
            T.PutValue('T_LETTRABLE', EnleveBlanc(Ligne, 62, 1));

            Collectif := BourreOuTronque(EnleveBlanc(Ligne, 63, 17), fbGene);
            T.PutValue('T_COLLECTIF', Collectif);
            T.PutValue('T_EAN', EnleveBlanc(Ligne, 80, 17));
            T.PutValue('T_TABLE0', Veriftablelibre('T00', EnleveBlanc(Ligne, 97, 3)));
            T.PutValue('T_TABLE1', Veriftablelibre('T01', EnleveBlanc(Ligne, 100, 3)));
            T.PutValue('T_TABLE2', Veriftablelibre('T02', EnleveBlanc(Ligne, 103, 3)));
            T.PutValue('T_TABLE3', Veriftablelibre('T03', EnleveBlanc(Ligne, 106, 3)));
            T.PutValue('T_TABLE4', Veriftablelibre('T04', EnleveBlanc(Ligne, 109, 3)));
            T.PutValue('T_TABLE5', Veriftablelibre('T05', EnleveBlanc(Ligne, 112, 3)));
            T.PutValue('T_TABLE6', Veriftablelibre('T06', EnleveBlanc(Ligne, 115, 3)));
            T.PutValue('T_TABLE7', Veriftablelibre('T07', EnleveBlanc(Ligne, 118, 3)));
            T.PutValue('T_TABLE8', Veriftablelibre('T08', EnleveBlanc(Ligne, 121, 3)));
            T.PutValue('T_TABLE9', Veriftablelibre('T09', EnleveBlanc(Ligne, 124, 3)));
            T.PutValue('T_ADRESSE1', VireTouteSub(EnleveBlanc(Ligne, 127, 35),'"'));
            T.PutValue('T_ADRESSE2', VireTouteSub(EnleveBlanc(Ligne, 162, 35),'"'));
            T.PutValue('T_ADRESSE3', VireTouteSub(EnleveBlanc(Ligne, 197, 35),'"'));
            T.PutValue('T_CODEPOSTAL', EnleveBlanc(Ligne, 232, 9));
            T.PutValue('T_VILLE', EnleveBlanc(Ligne, 241, 35));
            // mettre les ribs
            MajRIB;
            
            T.PutValue('T_PAYS', EnleveBlanc(Ligne, 323, 3));
            if EnleveBlanc(Ligne, 326, 17) <> '' then
               T.PutValue('T_ABREGE', VireTouteSub(EnleveBlanc(Ligne, 326, 17),'"'))
            else
               T.PutValue('T_ABREGE', VireTouteSub(EnleveBlanc(Ligne, 24, 17),'"'));
            T.PutValue('T_LANGUE', EnleveBlanc(Ligne, 343, 3));
            T.PutValue('T_MULTIDEVISE', EnleveBlanc(Ligne, 346, 1));
            T.PutValue('T_DEVISE', EnleveBlanc(Ligne, 347, 3));

            T.PutValue('T_TELEPHONE', EnleveBlanc(Ligne, 350, 25));
            T.PutValue('T_FAX', EnleveBlanc(Ligne, 375, 25));
            if (Regdefaut <> '') and (EnleveBlanc(Ligne, 400, 3) <> '') then
                 T.PutValue('T_REGIMETVA', EnleveBlanc(Ligne, 400, 3))
            else
                 T.PutValue('T_REGIMETVA', RendRegTvadefaut);
            if (Mrdefaut <> '') and (EnleveBlanc(Ligne, 403, 3) <> '') then
                 T.PutValue('T_MODEREGLE', EnleveBlanc(Ligne, 403, 3))
            else
                 T.PutValue('T_MODEREGLE', RendMRDefaut);

            T.PutValue('T_COMMENTAIRE', EnleveBlanc(Ligne, 406, 35));
            T.PutValue('T_NIF', EnleveBlanc(Ligne, 441, 17));
            T.PutValue('T_SIRET', EnleveBlanc(Ligne, 458, 17));
            T.PutValue('T_APE', EnleveBlanc(Ligne, 475, 5));
            // mettre les contacts
            if EnleveBlanc(Ligne, 480, 35) <> '' then
            begin
                 if TContactCAU  = nil then
                    TContactCAU := TOB.Create('', nil, -1);
                 Tc := TOB.Create('CONTACT', TContactCAU, -1);
                 Tc.PutValue('C_AUXILIAIRE', Compteauxi);
                 Tc.PutValue('C_TYPECONTACT', 'T');
                 Tc.PutValue('C_NATUREAUXI', EnleveBlanc(Ligne, 59, 3));
                 Tc.PutValue('C_NOM', EnleveBlanc(Ligne, 480, 35));
                 Tc.PutValue('C_SERVICE', EnleveBlanc(Ligne, 515, 35));
                 Tc.PutValue('C_FONCTION', EnleveBlanc(Ligne, 550, 35));
                 Tc.PutValue('C_TELEPHONE', EnleveBlanc(Ligne, 585, 25));
                 Tc.PutValue('C_FAX', EnleveBlanc(Ligne, 610, 25));
                 Tc.PutValue('C_TELEX', EnleveBlanc(Ligne, 635, 25));
                 Tc.PutValue('C_RVA', EnleveBlanc(Ligne, 660, 50));
                 Tc.PutValue('C_CIVILITE', EnleveBlanc(Ligne, 710, 3));
                 Tc.PutValue('C_PRINCIPAL', EnleveBlanc(Ligne, 713, 1));
            end;

            T.PutValue('T_FORMEJURIDIQUE', EnleveBlanc(Ligne, 714, 3));
            T.PutValue('T_CREERPAR', 'COM');
            T.PutValue('T_DATECREATION', Date-1);

            T.PutValue('T_PAYEUR', EnleveBlanc(Ligne, 718, 17));
            T.PutValue('T_ISPAYEUR', EnleveBlanc(Ligne, 735, 1));
            T.PutValue('T_AVOIRRBT', EnleveBlanc(Ligne, 736, 1));

            T.PutValue('T_CONFIDENTIEL', '0');
            Ti := RechercheOptimise(TTiersInsert, 'T_AUXILIAIRE', compteauxi, FALSE, fHTiersInsert);
            if Ti = nil then
            begin
                  Ti := TOB.Create('TIERS', TTiersInsert, -1);
                  Ti.Dupliquer(T, TRUE, TRUE, TRUE);
            end;
  end
  else
  begin
            if EnleveBlanc(Ligne, 24, 35) <> '' then T.PutValue('T_LIBELLE', VireTouteSub(EnleveBlanc(Ligne, 24, 35),'"'));
            T.PutValue('T_EAN', EnleveBlanc(Ligne, 80, 17));
            if EnleveBlanc(Ligne, 97, 3) <> '' then T.PutValue('T_TABLE0', Veriftablelibre('T00', EnleveBlanc(Ligne, 97, 3)));
            if EnleveBlanc(Ligne, 100, 3) <> '' then T.PutValue('T_TABLE1', Veriftablelibre('T01', EnleveBlanc(Ligne, 100, 3)));
            if EnleveBlanc(Ligne, 103, 3) <> '' then T.PutValue('T_TABLE2', Veriftablelibre('T02', EnleveBlanc(Ligne, 103, 3)));
            if EnleveBlanc(Ligne, 106, 3) <> '' then T.PutValue('T_TABLE3', Veriftablelibre('T03', EnleveBlanc(Ligne, 106, 3)));
            if EnleveBlanc(Ligne, 109, 3) <> '' then T.PutValue('T_TABLE4', Veriftablelibre('T04', EnleveBlanc(Ligne, 109, 3)));
            if EnleveBlanc(Ligne, 112, 3) <> '' then T.PutValue('T_TABLE5', Veriftablelibre('T05', EnleveBlanc(Ligne, 112, 3)));
            if EnleveBlanc(Ligne, 115, 3) <> '' then T.PutValue('T_TABLE6', Veriftablelibre('T06', EnleveBlanc(Ligne, 115, 3)));
            if EnleveBlanc(Ligne, 118, 3) <> '' then T.PutValue('T_TABLE7', Veriftablelibre('T07', EnleveBlanc(Ligne, 118, 3)));
            if EnleveBlanc(Ligne, 121, 3) <> '' then T.PutValue('T_TABLE8', Veriftablelibre('T08', EnleveBlanc(Ligne, 121, 3)));
            if EnleveBlanc(Ligne, 124, 3) <> '' then T.PutValue('T_TABLE9', Veriftablelibre('T09', EnleveBlanc(Ligne, 124, 3)));
            T.PutValue('T_ADRESSE1', VireTouteSub(EnleveBlanc(Ligne, 127, 35),'"'));
            T.PutValue('T_ADRESSE2', VireTouteSub(EnleveBlanc(Ligne, 162, 35),'"'));
            T.PutValue('T_ADRESSE3', VireTouteSub(EnleveBlanc(Ligne, 197, 35),'"'));
            T.PutValue('T_CODEPOSTAL', EnleveBlanc(Ligne, 232, 9));
            T.PutValue('T_VILLE', EnleveBlanc(Ligne, 241, 35));
            T.PutValue('T_PAYS', EnleveBlanc(Ligne, 323, 3));
            if EnleveBlanc(Ligne, 326, 17) <> '' then
              T.PutValue('T_ABREGE', VireTouteSub(EnleveBlanc(Ligne, 326, 17),'"'));
            T.PutValue('T_LANGUE', EnleveBlanc(Ligne, 343, 3));
            T.PutValue('T_TELEPHONE', EnleveBlanc(Ligne, 350, 25));
            T.PutValue('T_FAX', EnleveBlanc(Ligne, 375, 25));
            if (Regdefaut = '') and (EnleveBlanc(Ligne, 400, 3) <> '') then
                  T.PutValue('T_REGIMETVA', EnleveBlanc(Ligne, 400, 3))
            else
                  T.PutValue('T_REGIMETVA', RendRegTvadefaut);

            if (Mrdefaut = '') and (EnleveBlanc(Ligne, 403, 3) <> '') then
                  T.PutValue('T_MODEREGLE', EnleveBlanc(Ligne, 403, 3))
            else
               T.PutValue('T_MODEREGLE', RendMRDefaut);

            T.PutValue('T_COMMENTAIRE', EnleveBlanc(Ligne, 406, 35));
            T.PutValue('T_NIF', EnleveBlanc(Ligne, 441, 17));
            T.PutValue('T_SIRET', EnleveBlanc(Ligne, 458, 17));
            T.PutValue('T_APE', EnleveBlanc(Ligne, 475, 5));
            T.PutValue('T_FORMEJURIDIQUE', EnleveBlanc(Ligne, 714, 3));
            if (EnleveBlanc(Ligne, 718, 17) <> '') then
               T.PutValue('T_PAYEUR', EnleveBlanc(Ligne, 718, 17));
            if (EnleveBlanc(Ligne, 735, 1) <> '') then
               T.PutValue('T_ISPAYEUR', EnleveBlanc(Ligne, 735, 1));
            if (EnleveBlanc(Ligne, 736, 1) <> '') then
               T.PutValue('T_AVOIRRBT', EnleveBlanc(Ligne, 736, 1));

            // mettre les ribs
            if EnleveBlanc(Ligne, 276, 24) <> '' then
            begin
                 if TRibCAU = nil then
                    TRibCAU := TOB.Create('', nil, -1);
                 Tr := TOB.Create('RIB', TRibCAU, -1);
                 if Tr.SelectDB('"' + compteauxi + '"', nil, false) then
                 begin
                       if EnleveBlanc(Ligne, 276, 24) <> '' then Tr.PutValue('R_DOMICILIATION', EnleveBlanc(Ligne, 276, 24));
                       if EnleveBlanc(Ligne, 300, 5) <> '' then Tr.PutValue('R_ETABBQ', EnleveBlanc(Ligne, 300, 5));
                       if EnleveBlanc(Ligne, 305, 5) <> '' then Tr.PutValue('R_GUICHET', EnleveBlanc(Ligne, 305, 5));
                       if EnleveBlanc(Ligne, 310, 11) <> '' then Tr.PutValue('R_NUMEROCOMPTE', EnleveBlanc(Ligne, 310, 11));
                       if EnleveBlanc(Ligne, 321, 2) <> '' then Tr.PutValue('R_CLERIB', EnleveBlanc(Ligne, 321, 2));
                       // modif pour produflex
                       if EnleveBlanc(Ligne, 323, 3) <> '' then
                       begin
                        Tr.PutValue('R_PAYS', EnleveBlanc(Ligne, 323, 3));
                        if EnleveBlanc(Ligne, 323, 3) = 'FR' then
                          Tr.PutValue('R_PAYS', 'FRA');
                       end;
                       if EnleveBlanc(Ligne, 717, 1) <> '' then Tr.PutValue('R_PRINCIPAL', EnleveBlanc(Ligne, 717, 1));
                 end
                 else MajRIB;
            end;

  end;
  // modif pour produflex
   if AttribRIBAuto and (fHTiersInsert <> nil) and (Tr <> nil) then
   begin
         T := RechercheOptimise(TTiersInsert, 'T_AUXILIAIRE', compteauxi, FALSE, fHTiersInsert);
         if T <> nil then AjoutUpdateRib
         else
         begin
                 T := TTiersUpdate.FindFirst(['T_AUXILIAIRE'], [compteauxi], FALSE);
                 if (T <> nil)  then AjoutUpdateRib
                 else
                 begin
                      T := RechercheOptimise(TTiers, 'T_AUXILIAIRE', compteauxi, FALSE, fHTiers);
                      if (T <> nil)  then AjoutUpdateRib
                 end;
         end;
   end;


end;

procedure TImportCom.ImportTiersComp(Ligne : string);
var
  TC          : TOB;
  compteauxi  : string;
  ExisteAuxi  : Boolean;
  Okmj        : Boolean;
begin
        Okmj := TRUE;
        compteauxi := BourreOuTronque(EnleveBlanc(Ligne, 7, 17), fbAux);
        // fiche 10563
        TC := TOB.Create('TIERSCOMPL', nil, -1);
        TC.PutValue('YTC_AUXILIAIRE', compteauxi);
        ExisteAuxi :=   TC.SelectDB('"' + compteauxi + '"', nil, false);
        // ExisteAuxi := ExisteSQL ('SELECT YTC_AUXILIAIRE FROM TIERSCOMPL WHERE YTC_AUXILIAIRE="' + compteauxi + '"');

        if (NatureFichier = 'BAL') and BInterdireCreationTiers then
        begin
             // fiche 10385
             Okmj := ExisteSQL ('SELECT T_AUXILIAIRE FROM TIERS WHERE T_AUXILIAIRE="' + compteauxi + '"');
        end;
        if (ExisteAuxi and BMAJCPTE) or ((not ExisteAuxi) and Okmj) then
        begin
                (*  TC := TOB.Create('TIERSCOMPL', nil, -1);
                  TC.PutValue('YTC_AUXILIAIRE', compteauxi);
                *)
                TC.PutValue('YTC_TIERS', EnleveBlanc(Ligne, 24, 17));
                // fiche 10563
                MajChamp(TC, 'YTC_TABLELIBRETIERS1', EnleveBlanc(Ligne, 44, 6));
                MajChamp(TC, 'YTC_TABLELIBRETIERS2', EnleveBlanc(Ligne, 50, 6));
                MajChamp(TC, 'YTC_TABLELIBRETIERS3', EnleveBlanc(Ligne, 56, 6));
                MajChamp(TC, 'YTC_TABLELIBRETIERS4', EnleveBlanc(Ligne, 62, 6));
                MajChamp(TC, 'YTC_TABLELIBRETIERS5', EnleveBlanc(Ligne, 68, 6));
                MajChamp(TC, 'YTC_TABLELIBRETIERS6', EnleveBlanc(Ligne, 74, 6));
                MajChamp(TC, 'YTC_TABLELIBRETIERS7', EnleveBlanc(Ligne, 80, 6));
                MajChamp(TC, 'YTC_TABLELIBRETIERS8', EnleveBlanc(Ligne, 86, 6));
                MajChamp(TC, 'YTC_TABLELIBRETIERS9', EnleveBlanc(Ligne, 92, 6));
                MajChamp(TC, 'YTC_TABLELIBRETIERSA', EnleveBlanc(Ligne, 98, 6));

                TC.PutValue('YTC_VALLIBRE1', Valeur(StPoint(EnleveBlanc(Ligne, 104, 15, '0'))));
                TC.PutValue('YTC_VALLIBRE2', Valeur(StPoint(EnleveBlanc(Ligne, 119, 15, '0'))));
                TC.PutValue('YTC_VALLIBRE3', Valeur(StPoint(EnleveBlanc(Ligne, 134, 15, '0'))));

                TC.PutValue('YTC_DATELIBRE1', Format_Date_HAL(EnleveBlanc(Ligne, 149, 8, '01011900')));
                TC.PutValue('YTC_DATELIBRE2', Format_Date_HAL(EnleveBlanc(Ligne, 157, 8, '01011900')));
                TC.PutValue('YTC_DATELIBRE3', Format_Date_HAL(EnleveBlanc(Ligne, 165, 8, '01011900')));
                TC.PutValue('YTC_TABLELIBREFOU1', EnleveBlanc(Ligne, 173, 6));
                TC.PutValue('YTC_TABLELIBREFOU2', EnleveBlanc(Ligne, 179, 6));
                TC.PutValue('YTC_TABLELIBREFOU3', EnleveBlanc(Ligne, 185, 6));
                TC.PutValue('YTC_VALLIBREFOU1', Valeur(StPoint(EnleveBlanc(Ligne, 191, 15, '0'))));
                TC.PutValue('YTC_VALLIBREFOU2', Valeur(StPoint(EnleveBlanc(Ligne, 206, 15, '0'))));
                TC.PutValue('YTC_VALLIBREFOU3', Valeur(StPoint(EnleveBlanc(Ligne, 221, 15, '0'))));
                TC.PutValue('YTC_DATELIBREFOU1', Format_Date_HAL(EnleveBlanc(Ligne, 236, 8, '01011900')));
                TC.PutValue('YTC_DATELIBREFOU2', Format_Date_HAL(EnleveBlanc(Ligne, 244, 8, '01011900')));
                TC.PutValue('YTC_DATELIBREFOU3', Format_Date_HAL(EnleveBlanc(Ligne, 252, 8, '01011900')));
                TC.PutValue('YTC_TEXTELIBRE1', EnleveBlanc(Ligne, 260, 15));
                TC.PutValue('YTC_TEXTELIBRE2', EnleveBlanc(Ligne, 295, 15));
                TC.PutValue('YTC_TEXTELIBRE3', EnleveBlanc(Ligne, 330, 15));
                TC.PutValue('YTC_BOOLLIBRE1', EnleveBlanc(Ligne, 365, 1));
                TC.PutValue('YTC_BOOLLIBRE2', EnleveBlanc(Ligne, 366, 1));
                TC.PutValue('YTC_BOOLLIBRE3', EnleveBlanc(Ligne, 367, 1));
                MajChamp(TC, 'YTC_RESSOURCE1', EnleveBlanc(Ligne, 368, 17));
                MajChamp(TC, 'YTC_RESSOURCE2', EnleveBlanc(Ligne, 385, 17));
                MajChamp(TC, 'YTC_RESSOURCE3', EnleveBlanc(Ligne, 402, 17));
                TC.PutValue('YTC_DOCIDENTITE', EnleveBlanc(Ligne, 419, 3));
                TC.PutValue('YTC_DOCDATEDELIV', Format_Date_HAL(EnleveBlanc(Ligne, 422, 8,'01011900')));
                TC.PutValue('YTC_DOCDATEEXPIR', Format_Date_HAL(EnleveBlanc(Ligne, 430, 8,'01011900')));
                TC.PutValue('YTC_DOCOBSERV', EnleveBlanc(Ligne, 438, 35));
                TC.PutValue('YTC_DOCORIGINE', EnleveBlanc(Ligne, 473, 35));
                TC.PutValue('YTC_FAMREG', EnleveBlanc(Ligne, 508, 3));
                TC.PutValue('YTC_COMMSPECIAL', EnleveBlanc(Ligne, 511, 17));
                TC.PutValue('YTC_REPRESENTANT2', EnleveBlanc(Ligne, 528, 17));
                TC.PutValue('YTC_REPRESENTANT3', EnleveBlanc(Ligne, 545, 17));
                TC.PutValue('YTC_TAUXREPR1', Valeur(StPoint(EnleveBlanc(Ligne, 562, 15, '0'))));
                TC.PutValue('YTC_TAUXREPR2', Valeur(StPoint(EnleveBlanc(Ligne, 577, 15, '0'))));
                TC.PutValue('YTC_TAUXREPR3', Valeur(StPoint(EnleveBlanc(Ligne, 592, 15, '0'))));
                TC.PutValue('YTC_NADRESSELIV', StrToInt(EnleveBlanc(Ligne, 607, 6,'0')));
                TC.PutValue('YTC_NADRESSEFAC',StrToInt(EnleveBlanc(Ligne, 613, 6,'0')));

                TC.PutValue('YTC_STATIONEDI', EnleveBlanc(Ligne, 619, 17));
                TC.PutValue('YTC_NOTRECODETIERS', EnleveBlanc(Ligne, 636, 17));
                TC.PutValue('YTC_NOTRECODCOMPTA', EnleveBlanc(Ligne, 653, 17));
                TC.PutValue('YTC_SCHEMAGEN', EnleveBlanc(Ligne, 670, 17));
                TC.PutValue('YTC_ACCELERATEUR', EnleveBlanc(Ligne, 687, 1));
                TC.PutValue('YTC_DAS2', EnleveBlanc(Ligne, 688, 1));
                TC.PutValue('YTC_PROFESSION', EnleveBlanc(Ligne, 689, 35));
                TC.PutValue('YTC_REMUNERATION', EnleveBlanc(Ligne, 724, 3));
                TC.PutValue('YTC_INDEMNITE', EnleveBlanc(Ligne, 727, 3));
                TC.PutValue('YTC_AVANTAGE', EnleveBlanc(Ligne, 730, 3));
                TC.InsertOrUpdateDB(TRUE);
        end;
        TC.Free;
end;

procedure TImportCom.ImportJournaux(Ligne: string);
var
Mds, Axe                     : string;
typepiece                    : string;
nature                       : string;
Typesouche,ana               : string;
dmvt                         : integer;
Contrepartie                 : string;
typecontrpartie              : string;
TJ,TS                        : TOB;
Abrege                       : string;
Q1                           : TQuery;
TA                           : TOB;
LibelleJournal,Journal       : string;
TypSim                       : string;
Existejournal                : Boolean;
begin

Axe     := EnleveBlanc(Ligne, 71, 3);
Journal := EnleveBlanc(Ligne, 7, 3);

// application de la Table de correspondance
TraiteCorrespCpt(8, EnleveBlanc(Ligne, 7, 3), Journal);

Existejournal := ExisteSQL ('SELECT J_JOURNAL FROM JOURNAL Where J_JOURNAL="'+ Journal+'"');

if (EnleveBlanc(Ligne, 74, 3) = '') and ((EnleveBlanc(Ligne, 71, 3) = 'LIB') or (EnleveBlanc(Ligne, 71, 3) = 'BOR') or
   (EnleveBlanc(Ligne, 71, 3) = '-')) then
   begin
                   Mds := EnleveBlanc(Ligne, 71, 3); Axe := '';
   end
   else
   begin
                   Mds := EnleveBlanc(Ligne, 74, 3);
   end;
  if (Mds = '') or (((Mds = 'BOR')) and (OrigineFichier = 'S1')) then Mds := 'LIB';
   typepiece := Mds;
   nature := EnleveBlanc(Ligne, 45, 3);

   if (GetParamSocSecur ('SO_JALOUVRE', 'ANO') = Journal) then begin nature := 'ANO'; typepiece := '-'; end;
   if (GetParamSocSecur ('SO_JALFERME', 'CLO') = Journal) then begin nature := 'CLO'; typepiece := '-'; end;

   if (nature = 'ANO') and (typepiece <> '-') then typepiece := '-';
   if (nature = 'CLO') and (typepiece <> '-') then typepiece := '-';
   OKexercice := OnAfficheListeCom('Journal : '+ Journal ,LISTEIMPORT);
     // si dans le fichier le type du journal est incorrect pour journal ano, clo et reg
   if nature <> EnleveBlanc(Ligne, 45, 3) then
   begin
         OnAfficheListeCom('Nature '+ Journal + ' incohérente : '+
          'correction automatique',LISTEIMPORT);
   end;

   if typepiece <> Mds then
   begin
         Mds := typepiece;
         OnAfficheListeCom('Mode de saisie '+ Journal + ' incohérent : '+
          'correction automatique en pièce ',LISTEIMPORT);
   end;
   if EnleveBlanc(Ligne, 48, 3)= '' then  Typesouche := Journal
   else Typesouche := EnleveBlanc(Ligne, 48, 3);

   if (not ExisteSQl ('SELECT SH_SOUCHE FROM SOUCHE Where SH_TYPE="CPT" AND SH_SOUCHE="'+ Typesouche + '"')) then
   begin
        if (Paramgene) or (OrigineFichier = 'SI') then
        begin
            dmvt := 0;
            if nature = 'ODA' then ana := 'X'
            else ana := '-';
            TS := TOB.Create('SOUCHE', nil, -1);
            TS.putvalue ('SH_TYPE', 'CPT');
            TS.putvalue ('SH_SOUCHE', Typesouche);
            TS.putvalue ('SH_LIBELLE', 'SOUCHE : '+Typesouche);
            TS.putvalue ('SH_NUMDEPART', dmvt);
            TS.putvalue ('SH_SIMULATION', '-');
            TS.putvalue ('SH_ANALYTIQUE', ana);
            TS.InsertOrUpdateDB(TRUE);
            TS.Free;
         end;
   end;
   TypSim := EnleveBlanc(Ligne, 51, 3);
   if (TypSim = 'SIS') and (not ExisteSQl ('SELECT SH_SOUCHE FROM SOUCHE Where SH_TYPE="CPT" AND SH_SOUCHE="'+ TypSim + '"')) then
   begin
            dmvt := 0;
            if nature = 'ODA' then ana := 'X'
            else ana := '-';
            TS := TOB.Create('SOUCHE', nil, -1);
            TS.putvalue ('SH_TYPE', 'CPT');
            TS.putvalue ('SH_SOUCHE', TypSim);
            TS.putvalue ('SH_LIBELLE', 'SOUCHE : '+TypSim);
            TS.putvalue ('SH_NUMDEPART', dmvt);
            TS.putvalue ('SH_SIMULATION', 'X');
            TS.putvalue ('SH_ANALYTIQUE', ana);
            TS.InsertOrUpdateDB(TRUE);
            TS.Free;
   end;

  Contrepartie := BourreOuTronque(EnleveBlanc(Ligne, 54, 17), fbGene);
  typecontrpartie := 'MAN';
  if EnleveBlanc(Ligne, 477, 17)= '' then
     Abrege :=  VireTouteSub(EnleveBlanc(Ligne, 10, 17),'"')
  else
     Abrege := VireTouteSub(EnleveBlanc(Ligne, 477, 17),'"');
  if (nature <> 'BQE') and (nature <> 'CAI') then begin Contrepartie := '';  typecontrpartie := '' end;

  if (OrigineFichier = 'S1') and (nature <> 'ANO') and (nature <> 'CLO') then  
  begin
        Q1 := OpenSql(' Select J_MODESAISIE from JOURNAL Where J_JOURNAL="' + Journal
        + '" AND J_MODESAISIE <> "'+Mds+'"', True);
        if not Q1.Eof then
        begin
          if (Mds = 'LIB') and (Q1.FindField('J_MODESAISIE').asstring  <> 'LIB') and (Mds <> Q1.FindField('J_MODESAISIE').asstring) then
          begin
                   if TOBJrl = nil then
                      TOBJrl := TOB.Create('', nil, -1);
                   TA := TOB.Create ('',TOBJrl,-1);
                   TA.AddChampSupValeur('JOURNAL', Journal);
                   TA.AddChampSupValeur('MODESAISIE', Q1.FindField('J_MODESAISIE').asstring);
                   TA.AddChampSupValeur('OK', FALSE);
          end;
        end;
        Ferme(Q1);
        if not Existejournal then
        begin
              dmvt := 0;
              if nature = 'ODA' then ana := 'X'
              else ana := '-';
              TS := TOB.Create('SOUCHE', nil, -1);
              TS.putvalue ('SH_TYPE', 'CPT');
              TS.putvalue ('SH_SOUCHE', Typesouche);
              TS.putvalue ('SH_LIBELLE', 'SOUCHE : '+Typesouche);
              TS.putvalue ('SH_NUMDEPART', dmvt);
              TS.putvalue ('SH_SIMULATION', '-');
              TS.putvalue ('SH_ANALYTIQUE', ana);
              TS.InsertOrUpdateDB(TRUE);
              TS.Free;
        end;
  end;

  if Existejournal and (not Paramgene) then  exit;
  // Fiche 10340
  if (OrigineFichier = 'SI') and (NatureFichier = 'DOS') and Existejournal then  exit;

  TJ := TOB.Create('JOURNAL', nil, -1);
  TJ.putvalue ('J_JOURNAL', Journal);
  Libellejournal := EnleveBlanc(Ligne, 10, 35);
  if LibelleJournal = '' then LibelleJournal := '.';
  TJ.putvalue ('J_LIBELLE', VireTouteSub(LibelleJournal,'"'));
  TJ.putvalue ('J_NATUREJAL', nature);
  TJ.putvalue ('J_COMPTEURNORMAL', Typesouche);
  if TypSim <> '' then
     TJ.putvalue ('J_COMPTEURSIMUL',  TypSim)
  else
     TJ.putvalue ('J_COMPTEURSIMUL',  'SIM');

  TJ.putvalue ('J_CONTREPARTIE', Contrepartie);
  TJ.PutValue ('J_TYPECONTREPARTIE', typecontrpartie);
  TJ.putvalue ('J_AXE', Axe);
  TJ.putvalue ('J_MODESAISIE', Mds);
  TJ.putvalue ('J_COMPTEAUTOMAT', EnleveBlanc(Ligne, 77, 200));
  TJ.putvalue ('J_COMPTEINTERDIT',  EnleveBlanc(Ligne, 277, 200));
  TJ.putvalue ('J_CENTRALISABLE',  '-');
  TJ.putvalue ('J_MULTIDEVISE', 'X');
  TJ.putvalue ('J_ABREGE', Abrege);
  TJ.SetAllModifie (TRUE);
  TJ.InsertOrUpdateDB(TRUE);
  TJ.Free;

  if Mds = 'BOR' then
    SetParamSoc('SO_ZACTIVEPFU', 'X');
end;

procedure TImportCom.ImportRib(Ligne: string);
var
  TL              : TOB;
  compteaux       : string;
  Okmj,ExisteAuxi : Boolean;
  szIban,Iban     : string;
  Q               : TQuery;
  T               : TOB;
  numrib          : integer;
    Procedure AjoutUpdateRib;
    begin
             if not T.FieldExists ('R_PAYS') then
             begin
                  T.AddChampSupValeur ('R_PAYS',  TL.GetValue ('R_PAYS'));
                  T.AddChampSupValeur ('R_ETABBQ', TL.GetValue ('R_ETABBQ'));
                  T.AddChampSupValeur ('R_GUICHET', TL.GetValue ('R_GUICHET'));
                  T.AddChampSupValeur ('R_NUMEROCOMPTE', TL.GetValue ('R_NUMEROCOMPTE'));
                  T.AddChampSupValeur ('R_CLERIB', TL.GetValue ('R_CLERIB'));
                  T.AddChampSupValeur ('R_DOMICILIATION', TL.GetValue ('R_DOMICILIATION'));
                  T.AddChampSupValeur ('R_CODEIBAN', TL.GetValue ('R_CODEIBAN'));
              end
              else
              begin
                  T.putValue ('R_PAYS',  TL.GetValue ('R_PAYS'));
                  T.putValue ('R_ETABBQ', TL.GetValue ('R_ETABBQ'));
                  T.putValue ('R_GUICHET', TL.GetValue ('R_GUICHET'));
                  T.putValue ('R_NUMEROCOMPTE', TL.GetValue ('R_NUMEROCOMPTE'));
                  T.putValue ('R_CLERIB', TL.GetValue ('R_CLERIB'));
                  T.putValue ('R_DOMICILIATION', TL.GetValue ('R_DOMICILIATION'));
                  T.putValue ('R_CODEIBAN', TL.GetValue ('R_CODEIBAN'));
              end;
    end;
begin

  Okmj := TRUE;
  compteaux := BourreOuTronque( EnleveBlanc(Ligne, 7, 17,''), fbAux);
  numrib := StrToInt(EnleveBlanc(Ligne, 24, 6,'0'));
  Q := OpenSQL ('SELECT R_AUXILIAIRE, R_CODEIBAN FROM RIB WHERE R_AUXILIAIRE="' + compteaux + '"'+
  'AND R_NUMERORIB='+IntToStr(numrib), TRUE);
  if Q.EOF then  ExisteAuxi := FALSE
  else
  begin
   IBan := Q.FindField ('R_CODEIBAN').asstring;
   ExisteAuxi := TRUE;
  end;
  Ferme (Q);

//  ExisteAuxi := ExisteSQL ('SELECT R_AUXILIAIRE, R_CODEIBAN FROM RIB WHERE R_AUXILIAIRE="' + compteaux + '"');
  if (NatureFichier = 'BAL') and BInterdireCreationTiers then
  begin
   //Okmj := FALSE;
     Okmj := ExisteSQL ('SELECT T_AUXILIAIRE FROM TIERS WHERE T_AUXILIAIRE="' + compteaux + '"');
  end;

  if (ExisteAuxi and BMAJCPTE) or ((not ExisteAuxi) and Okmj) then
  begin
      TL := TOB.Create('RIB', nil, -1);
      TL.PutValue ('R_AUXILIAIRE', compteaux);
      TL.PutValue ('R_NUMERORIB', StrToInt(EnleveBlanc(Ligne, 24, 6,'0')));
      TL.PutValue ('R_PRINCIPAL', EnleveBlanc(Ligne, 30, 1,''));
      TL.PutValue ('R_ETABBQ', EnleveBlanc(Ligne, 31, 5,''));
      TL.PutValue ('R_GUICHET', EnleveBlanc(Ligne, 36, 5,''));
      TL.PutValue ('R_NUMEROCOMPTE', EnleveBlanc(Ligne, 41, 11,''));
      TL.PutValue ('R_CLERIB', EnleveBlanc(Ligne, 52, 2,''));
      TL.PutValue ('R_DOMICILIATION', EnleveBlanc(Ligne, 54, 24,''));
      TL.PutValue ('R_VILLE', EnleveBlanc(Ligne, 78, 35,''));
      if EnleveBlanc(Ligne, 113, 3,'') = '' then
         TL.PutValue ('R_PAYS', 'FRA')
      else
         TL.PutValue ('R_PAYS', EnleveBlanc(Ligne, 113, 3,''));
      TL.PutValue ('R_DEVISE', EnleveBlanc(Ligne, 116, 3,''));
      TL.PutValue ('R_CODEBIC', EnleveBlanc(Ligne, 119, 35,''));
      TL.PutValue ('R_SOCIETE', EnleveBlanc(Ligne, 154, 3,''));
      TL.PutValue ('R_SALAIRE', EnleveBlanc(Ligne, 157, 1,''));
      TL.PutValue ('R_ACOMPTE', EnleveBlanc(Ligne, 158, 1,''));
      TL.PutValue ('R_FRAISPROF', EnleveBlanc(Ligne, 159, 1,''));
      // fiche 10429  // fiche 10586
      szIban := calcIBAN(TrouveCodeIsoPays (TL.GetValue ('R_PAYS')), calcRIB(TL.GetValue ('R_PAYS'),TL.GetValue ('R_ETABBQ'),TL.GetValue ('R_GUICHET'),TL.GetValue ('R_NUMEROCOMPTE'),TL.GetValue ('R_CLERIB')));
      TL.PutValue ('R_CODEIBAN', EnleveBlanc(Ligne, 160, 70, szIban));
        // ajout me nature éco et nouveau format des comptes
      if ((VersionFichier <> '') and (StrToint(VersionFichier) >= 6)) then
      begin
        TL.PutValue ('R_NATECO', EnleveBlanc(Ligne, 230, 3,''));
        TL.PutValue ('R_TYPEPAYS', EnleveBlanc(Ligne, 233, 1,''));
        TL.PutValue ('R_TYPEPAYS', EnleveBlanc(Ligne, 233, 1,''));
        if  EnleveBlanc(Ligne, 234, 8,'') <> '' then
          TL.PutValue ('R_ETABBQ', EnleveBlanc(Ligne, 234, 8,''));
        if  EnleveBlanc(Ligne, 242, 20,'') <> '' then
          TL.PutValue ('R_NUMEROCOMPTE', EnleveBlanc(Ligne, 242, 20,''));
      end;

      if not ExisteAuxi then
      begin
         TL.InsertDB(nil, TRUE);
         if AttribRIBAuto and (fHTiersInsert <> nil) then
         begin
               T := RechercheOptimise(TTiersInsert, 'T_AUXILIAIRE', compteaux, FALSE, fHTiersInsert);
               if T <> nil then AjoutUpdateRib
               else
               begin
                       T := TTiersUpdate.FindFirst(['T_AUXILIAIRE'], [compteaux], FALSE);
                       if (T <> nil)  then AjoutUpdateRib
                       else
                       begin
                            T := RechercheOptimise(TTiers, 'T_AUXILIAIRE', compteaux, FALSE, fHTiers);
                            if (T <> nil)  then AjoutUpdateRib
                       end;
               end;
         end;
      end
      else
      begin
         if (szIban = '') and (IBan <> '') then   // fiche 10429
            TL.PutValue ('R_CODEIBAN', IBan);
         TL.SetAllModifie(True);
         TL.UpdateDB(TRUE);
      end;
      TL.Free;
      OKexercice := OnAfficheListeCom('Import Rib ' + EnleveBlanc(Ligne, 7, 17,'') ,LISTEIMPORT);
  end;
end;


procedure TImportCom.ImportBanqueCp(Ligne: string);
var
  TL          : TOB;
  szIban      : string;
  bqg         : string;
begin
  TL := TOB.Create('BANQUECP', nil, -1);

  TL.PutValue ('BQ_CODE', EnleveBlanc(Ligne, 7, 17,''));

// fiche 10511 TL.PutValue ('BQ_GENERAL', EnleveBlanc(Ligne, 24, 17,''));
  bqg := BourreOuTronque(EnleveBlanc(Ligne, 24, 17, ''), fbGene);
  TL.PutValue ('BQ_GENERAL', bqg);
  TL.PutValue ('BQ_LIBELLE', EnleveBlanc(Ligne, 41, 35,''));
  TL.PutValue ('BQ_DOMICILIATION', EnleveBlanc(Ligne, 76, 24,''));
  TL.PutValue ('BQ_ADRESSE1', EnleveBlanc(Ligne, 100, 35,''));
  TL.PutValue ('BQ_ADRESSE2', EnleveBlanc(Ligne, 135, 35,''));
  TL.PutValue ('BQ_ADRESSE3', EnleveBlanc(Ligne, 170, 35,''));
  TL.PutValue ('BQ_CODEPOSTAL', EnleveBlanc(Ligne, 205, 9,''));
  TL.PutValue ('BQ_VILLE', EnleveBlanc(Ligne, 214, 35,''));
  TL.PutValue ('BQ_DIVTERRIT', EnleveBlanc(Ligne, 249, 9,''));
  TL.PutValue ('BQ_PAYS', EnleveBlanc(Ligne, 258, 3,''));
  TL.PutValue ('BQ_LANGUE', EnleveBlanc(Ligne, 261, 3,''));
  TL.PutValue ('BQ_DEVISE', EnleveBlanc(Ligne, 264, 3,''));
  TL.PutValue ('BQ_TELEPHONE', EnleveBlanc(Ligne, 267, 25,''));
  TL.PutValue ('BQ_FAX', EnleveBlanc(Ligne, 292, 25,''));
  TL.PutValue ('BQ_CONTACT', EnleveBlanc(Ligne, 317, 35,''));
  TL.PutValue ('BQ_ETABBQ', EnleveBlanc(Ligne, 352, 5,''));
  TL.PutValue ('BQ_NUMEROCOMPTE', EnleveBlanc(Ligne, 357, 11,''));
  TL.PutValue ('BQ_CLERIB', EnleveBlanc(Ligne, 368, 2,''));
  TL.PutValue ('BQ_GUICHET', EnleveBlanc(Ligne, 370, 5,''));
  TL.PutValue ('BQ_CODEBIC', EnleveBlanc(Ligne, 375, 35,''));
  TL.PutValue ('BQ_NUMEMETLCR', EnleveBlanc(Ligne, 410, 6,''));
  TL.PutValue ('BQ_CONVENTIONLCR', EnleveBlanc(Ligne, 416, 6,''));
  TL.PutValue ('BQ_NUMEMETVIR', EnleveBlanc(Ligne, 422, 6,''));
  TL.PutValue ('BQ_JOURFERMETUE', EnleveBlanc(Ligne, 428, 7,''));
  TL.PutValue ('BQ_REPRELEVE', EnleveBlanc(Ligne, 435, 40,''));
  TL.PutValue ('BQ_REPLCR', EnleveBlanc(Ligne, 475, 40,''));
  TL.PutValue ('BQ_REPLCRFOURN', EnleveBlanc(Ligne, 515, 40,''));
  TL.PutValue ('BQ_REPVIR', EnleveBlanc(Ligne, 555, 40,''));
  TL.PutValue ('BQ_REPPRELEV', EnleveBlanc(Ligne, 595, 40,''));
  TL.PutValue ('BQ_REPBONAPAYER', EnleveBlanc(Ligne, 635, 40,''));
  TL.PutValue ('BQ_REPIMPAYELCR', EnleveBlanc(Ligne, 675, 40,''));
  TL.PutValue ('BQ_DELAIVIRORD', EnleveBlanc(Ligne, 715, 5,'0'));
  TL.PutValue ('BQ_DELAIVIRCHAUD', EnleveBlanc(Ligne, 720, 5,'0'));
  TL.PutValue ('BQ_DELAIVIRBRULANT', EnleveBlanc(Ligne, 725, 5,'0'));
  TL.PutValue ('BQ_DELAIPRELVORD', EnleveBlanc(Ligne, 730, 5,'0'));
  TL.PutValue ('BQ_DELAIPRELVACC', EnleveBlanc(Ligne, 735, 5,'0'));
  TL.PutValue ('BQ_DELAILCR', EnleveBlanc(Ligne, 740, 5,'0'));
  TL.PutValue ('BQ_GUIDECOMPATBLE', EnleveBlanc(Ligne, 745, 17,''));
  TL.PutValue ('BQ_DERNSOLDEFRS', EnleveBlanc(Ligne, 762, 20,'0'));
  TL.PutValue ('BQ_DERNSOLDEDEV', EnleveBlanc(Ligne, 782, 20,'0'));
  TL.PutValue ('BQ_DATEDERNSOLDE', StrToDate(EnleveBlanc(Ligne, 802, 10,stDate1900)));
  TL.PutValue ('BQ_RELEVEETRANGER', EnleveBlanc(Ligne, 812, 1,''));
  TL.PutValue ('BQ_CALENDRIER', EnleveBlanc(Ligne, 813, 3,''));
  TL.PutValue ('BQ_RAPPAUTOREL', EnleveBlanc(Ligne, 816, 1,''));
  TL.PutValue ('BQ_RAPPROAUTOLCR', EnleveBlanc(Ligne, 817, 1,''));
  TL.PutValue ('BQ_LETTREVIR', EnleveBlanc(Ligne, 818, 3,''));
  TL.PutValue ('BQ_LETTREPRELV', EnleveBlanc(Ligne, 821, 3,''));
  TL.PutValue ('BQ_LETTRELCR', EnleveBlanc(Ligne, 824, 3,''));
  TL.PutValue ('BQ_ENCOURSLCR', EnleveBlanc(Ligne, 827, 20,'0'));
  TL.PutValue ('BQ_PLAFONDLCR', EnleveBlanc(Ligne, 847, 20,'0'));
  TL.PutValue ('BQ_REPIMPAYEPRELV', EnleveBlanc(Ligne, 867, 40,''));
  TL.PutValue ('BQ_ECHEREPPRELEV', EnleveBlanc(Ligne, 907, 1,''));
  TL.PutValue ('BQ_ECHEREPLCR', EnleveBlanc(Ligne, 908, 1,''));
  TL.PutValue ('BQ_SOCIETE', EnleveBlanc(Ligne, 909, 3,''));
  TL.PutValue ('BQ_BANQUE', EnleveBlanc(Ligne, 912, 3,''));
  TL.PutValue ('BQ_DELAIBAPLCR', EnleveBlanc(Ligne, 915, 5,'0'));
  TL.PutValue ('BQ_DELAITRANSINT', EnleveBlanc(Ligne, 920, 5,'0'));
  TL.PutValue ('BQ_COMPTEFRAIS', EnleveBlanc(Ligne, 925, 17,''));
  TL.PutValue ('BQ_TYPEREMTRANS', EnleveBlanc(Ligne, 942, 3,''));
  TL.PutValue ('BQ_INDREMTRANS', EnleveBlanc(Ligne, 945, 3,''));
  TL.PutValue ('BQ_COMMENTAIRE', EnleveBlanc(Ligne, 948, 35,''));
  TL.PutValue ('BQ_LETTRECHQ', EnleveBlanc(Ligne, 983, 3,''));
  TL.PutValue ('BQ_NUMEMETPRE', EnleveBlanc(Ligne, 986, 6,''));
  TL.PutValue ('BQ_DESTINATAIRE', EnleveBlanc(Ligne, 992, 5,''));
  TL.PutValue ('BQ_MULTIDEVISE', EnleveBlanc(Ligne, 997, 1,''));

  // fiche 10429     // fiche 10586
  szIban := calcIBAN(TrouveCodeIsoPays(TL.GetValue ('BQ_PAYS')), calcRIB(TL.GetValue ('BQ_PAYS'), TL.GetValue ('BQ_ETABBQ'),TL.GetValue ('BQ_GUICHET'),TL.GetValue ('BQ_NUMEROCOMPTE'), TL.GetValue ('BQ_CLERIB')));

  TL.PutValue ('BQ_CODEIBAN', EnleveBlanc(Ligne, 998, 70, szIban));
  TL.PutValue ('BQ_NODOSSIER', EnleveBlanc(Ligne, 1028, 8, V_PGI.NoDossier));
  TL.PutValue ('BQ_CODECIB', EnleveBlanc(Ligne, 1036, 3, ''));
  TL.PutValue ('BQ_NATURECPTE', EnleveBlanc(Ligne, 1039, 3, 'BQE'));
  if not ExisteSQL ('SELECT * FROM BANQUECP Where BQ_CODE="'+EnleveBlanc(Ligne, 7, 17,'')+'"')then
// 10512 +' AND BQ_GENERAL ="'+ bqg +'"' )then
  begin
       TL.PutValue ('BQ_NODOSSIER', V_PGI.NoDossier);
       TL.InsertDB(nil, TRUE);
  end
  else
  begin
       TL.SetAllModifie(True);
       TL.UpdateDB(TRUE);
  end;
  TL.Free;
  OKexercice := OnAfficheListeCom('Import Rib ' + EnleveBlanc(Ligne, 7, 17,'') ,LISTEIMPORT);
end;

procedure TImportCom.ImportBanque(Ligne: string);
var
  TL          : TOB;
begin
  TL := TOB.Create('BANQUES', nil, -1);

  TL.PutValue ('PQ_BANQUE', EnleveBlanc(Ligne, 7, 3,''));
  TL.PutValue ('PQ_LIBELLE', EnleveBlanc(Ligne, 10, 35,''));
  TL.PutValue ('PQ_RESTCB', EnleveBlanc(Ligne, 45, 3,''));
  TL.PutValue ('PQ_REMCB', EnleveBlanc(Ligne, 48, 3,''));
  TL.PutValue ('PQ_RESTCHQ', EnleveBlanc(Ligne, 51, 3,''));
  TL.PutValue ('PQ_REMCHQ', EnleveBlanc(Ligne, 54, 3,''));
  TL.PutValue ('PQ_RESTESP', EnleveBlanc(Ligne, 57, 3,''));
  TL.PutValue ('PQ_REMESP', EnleveBlanc(Ligne, 60, 3,''));
  TL.PutValue ('PQ_RESTLCR', EnleveBlanc(Ligne, 63, 3,''));
  TL.PutValue ('PQ_REMLCR', EnleveBlanc(Ligne, 66, 3,''));
  TL.PutValue ('PQ_RESTPRE', EnleveBlanc(Ligne, 69, 3,''));
  TL.PutValue ('PQ_REMPRE', EnleveBlanc(Ligne, 72, 3,''));
  TL.PutValue ('PQ_RESTTRI', EnleveBlanc(Ligne, 75, 3,''));
  TL.PutValue ('PQ_REMTRI', EnleveBlanc(Ligne, 78, 3,''));
  TL.PutValue ('PQ_RESTVIR', EnleveBlanc(Ligne, 81, 3,''));
  TL.PutValue ('PQ_REMVIR', EnleveBlanc(Ligne, 84, 3,''));
  TL.PutValue ('PQ_RESTVIT', EnleveBlanc(Ligne, 87, 3,''));
  TL.PutValue ('PQ_REMVIT', EnleveBlanc(Ligne, 90, 3,''));
  TL.PutValue ('PQ_RESTTIP', EnleveBlanc(Ligne, 93, 3,''));
  TL.PutValue ('PQ_REMTIP', EnleveBlanc(Ligne, 96, 3,''));
  TL.PutValue ('PQ_RESTTEP', EnleveBlanc(Ligne, 99, 3,''));
  TL.PutValue ('PQ_REMTEP', EnleveBlanc(Ligne, 102, 3,''));
  TL.PutValue ('PQ_RESTVIC', EnleveBlanc(Ligne, 105, 3,''));
  TL.PutValue ('PQ_REMVIC', EnleveBlanc(Ligne, 108, 3,''));
  TL.PutValue ('PQ_DE_MODE', EnleveBlanc(Ligne, 111, 3,''));
  TL.PutValue ('PQ_DE_TXREF', EnleveBlanc(Ligne, 114, 3,''));

  TL.PutValue ('PQ_DE_CORRECTION', EnleveBlanc(Ligne, 117, 20,'0'));
  TL.PutValue ('PQ_DE_SAISIE', EnleveBlanc(Ligne, 137, 20,'0'));
  TL.PutValue ('PQ_DE_NUMERATEUR', EnleveBlanc(Ligne, 157, 5,'0'));
  TL.PutValue ('PQ_DE_DENOMINATEUR', EnleveBlanc(Ligne, 162, 5,'0'));
  TL.PutValue ('PQ_DE_TAUX', EnleveBlanc(Ligne, 167, 10,'0'));
  TL.PutValue ('PQ_DE_DATETAUX', StrToDate(EnleveBlanc(Ligne, 177, 10,stDate1900)));
  TL.PutValue ('PQ_DE_PLAFOND', EnleveBlanc(Ligne, 187, 20,'0'));
  TL.PutValue ('PQ_PD_MODE', EnleveBlanc(Ligne, 207, 3,'0'));
  TL.PutValue ('PQ_PD_TXREF', EnleveBlanc(Ligne, 210, 3,'0'));

  TL.PutValue ('PQ_PD_CORRECTION', EnleveBlanc(Ligne, 213, 20,'0'));
  TL.PutValue ('PQ_PD_SAISIE', EnleveBlanc(Ligne, 233, 20,'0'));
  TL.PutValue ('PQ_PD_NUMERATEUR', EnleveBlanc(Ligne, 253, 5,'0'));
  TL.PutValue ('PQ_PD_DENOMINATEUR', EnleveBlanc(Ligne, 258, 5,'0'));
  TL.PutValue ('PQ_PD_TAUX', EnleveBlanc(Ligne, 263, 10,'0'));
  TL.PutValue ('PQ_PD_DATETAUX', StrToDate(EnleveBlanc(Ligne, 273, 10,stDate1900)));
  TL.PutValue ('PQ_PD_PLAFOND', EnleveBlanc(Ligne, 283, 20,'0'));
  TL.PutValue ('PQ_CR_MODE', EnleveBlanc(Ligne, 303, 3,''));
  TL.PutValue ('PQ_CR_TXREF', EnleveBlanc(Ligne, 306, 3,''));

  TL.PutValue ('PQ_CR_CORRECTION', EnleveBlanc(Ligne, 309, 20,'0'));
  TL.PutValue ('PQ_CR_SAISIE', EnleveBlanc(Ligne, 329, 20,'0'));
  TL.PutValue ('PQ_CR_NUMERATEUR', EnleveBlanc(Ligne, 349, 5,'0'));
  TL.PutValue ('PQ_CR_DENOMINATEUR', EnleveBlanc(Ligne, 354, 5,'0'));
  TL.PutValue ('PQ_CR_TAUX', EnleveBlanc(Ligne, 359, 10,'0'));
  TL.PutValue ('PQ_CR_DATETAUX', StrToDate(EnleveBlanc(Ligne, 369, 10,stDate1900)));
  TL.PutValue ('PQ_CR_PLAFOND', EnleveBlanc(Ligne, 379, 20,'0'));
  TL.PutValue ('PQ_CR_TYPEPLAFOND', EnleveBlanc(Ligne, 399, 3,''));
  TL.PutValue ('PQ_CO_MODE', EnleveBlanc(Ligne, 402, 3,''));
  TL.PutValue ('PQ_CO_TXREF', EnleveBlanc(Ligne, 405, 3,''));

  TL.PutValue ('PQ_CO_CORRECTION', EnleveBlanc(Ligne, 408, 20,'0'));
  TL.PutValue ('PQ_CO_SAISIE', EnleveBlanc(Ligne, 428, 20,'0'));
  TL.PutValue ('PQ_CO_NUMERATEUR', EnleveBlanc(Ligne, 448, 5,'0'));
  TL.PutValue ('PQ_CO_DENOMINATEUR', EnleveBlanc(Ligne, 453, 5,'0'));
  TL.PutValue ('PQ_CO_TAUX', EnleveBlanc(Ligne, 458, 10,'0'));
  TL.PutValue ('PQ_CO_DATETAUX', StrToDate(EnleveBlanc(Ligne, 468, 10,stDate1900)));
  TL.PutValue ('PQ_CO_FRAIS', EnleveBlanc(Ligne, 478, 20,'0'));
  TL.PutValue ('PQ_CO_TVA', EnleveBlanc(Ligne, 498, 20,'0'));
  TL.PutValue ('PQ_BB_TAUX', EnleveBlanc(Ligne, 518, 20,'0'));
  TL.PutValue ('PQ_BB_DATETAUX', StrToDate(EnleveBlanc(Ligne, 538, 10,stDate1900)));
  // PB d'existance physique de ces champs voir OL
  // pour le pb des bases anciennes
 if ChampphysiqueExiste ( 'BANQUES', 'PQ_ETABBQ') then
 begin
  TL.PutValue ('PQ_ETABBQ', EnleveBlanc(Ligne, 548, 5,''));
  TL.PutValue ('PQ_ABREGE', EnleveBlanc(Ligne, 553, 17,''));
  TL.InsertOrUpdateDB(TRUE);
 end;
  TL.Free;
  OKexercice := OnAfficheListeCom('Import Banque ' + EnleveBlanc(Ligne, 7, 3,'') ,LISTEIMPORT);
end;

procedure TImportCom.ImportLigneReleveBanque(Ligne: string);
var
  TL          : TOB;
begin
  if IgnorePointage then exit;
  TL := TOB.Create('EEXBQLIG', nil, -1);
  try
    TL.PutValue ('CEL_GENERAL', EnleveBlanc(Ligne, 7, 17,''));
    TL.PutValue ('CEL_NUMRELEVE', EnleveBlanc(Ligne, 24, 10,'0'));
    TL.PutValue ('CEL_NUMLIGNE', EnleveBlanc(Ligne, 34, 10,'0'));
    TL.PutValue ('CEL_LIBELLE', EnleveBlanc(Ligne, 44, 35,''));
    TL.PutValue ('CEL_RIB', EnleveBlanc(Ligne, 79, 35,''));
    TL.PutValue ('CEL_VALIDE', EnleveBlanc(Ligne, 114, 1,''));
    TL.PutValue ('CEL_CODEAFB', EnleveBlanc(Ligne, 115, 3,''));
    TL.PutValue ('CEL_DATEOPERATION', Str8ToDate(EnleveBlanc(Ligne, 118, 8,'01011900'),True));
    TL.PutValue ('CEL_DATEVALEUR', Str8ToDate(EnleveBlanc(Ligne, 126, 8,'01011900'),True));
    TL.PutValue ('CEL_DEBITDEV', Valeur(StPoint(EnleveBlanc(Ligne, 134, 20, '0'))));
    TL.PutValue ('CEL_CREDITDEV', Valeur(StPoint(EnleveBlanc(Ligne, 154, 20, '0'))));
    TL.PutValue ('CEL_DEBITEURO', Valeur(StPoint(EnleveBlanc(Ligne, 174, 20, '0'))));
    TL.PutValue ('CEL_CREDITEURO', Valeur(StPoint(EnleveBlanc(Ligne, 194, 20, '0'))));
    TL.PutValue ('CEL_REFPIECE', EnleveBlanc(Ligne, 214, 35,'0'));
    TL.PutValue ('CEL_REFORIGINE', EnleveBlanc(Ligne, 249, 17,''));
    TL.PutValue ('CEL_REFPOINTAGE', EnleveBlanc(Ligne, 266, 17,''));
    TL.PutValue ('CEL_DATEPOINTAGE', Str8ToDate(EnleveBlanc(Ligne, 283, 8,'01011900'),True));
    TL.PutValue ('CEL_IMO', EnleveBlanc(Ligne, 291, 2,''));
    TL.PutValue ('CEL_CODERAPPRO', EnleveBlanc(Ligne, 293, 10,'0'));
    TL.PutValue ('CEL_EXONERE', EnleveBlanc(Ligne, 303, 10,'0'));
    TL.PutValue ('CEL_DISPONIBLE', EnleveBlanc(Ligne, 313, 10,'0'));
    TL.PutValue ('CEL_LIBELLE1', EnleveBlanc(Ligne, 323, 50,''));
    TL.PutValue ('CEL_LIBELLE2', EnleveBlanc(Ligne, 373, 50,''));
    TL.PutValue ('CEL_LIBELLE3', EnleveBlanc(Ligne, 423, 50,''));
    TL.PutValue ('CEL_NATUREINTERNE', EnleveBlanc(Ligne, 473, 3,''));
    TL.InsertOrUpdateDB(TRUE);
  finally
    TL.Free;
  end;
  if SaveLigne <> EnleveBlanc(Ligne, 7, 17,'') then
     OKexercice := OnAfficheListeCom('Import Ligne Relevé Bancaire ' + EnleveBlanc(Ligne, 7, 17,''), LISTEIMPORT);
  SaveLigne := EnleveBlanc(Ligne, 7, 17,'');
end;


{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 28/01/2004
Modifié le ... :   /  /
Description .. : Import d'un relevé bancaire : table EEXBQ
Mots clefs ... :
*****************************************************************}
procedure TImportCom.ImportReleveBanque(Ligne: string);
var
  TL          : TOB;
begin
  if IgnorePointage then exit;
  TL := TOB.Create('EEXBQ', nil, -1);
  try
    TL.PutValue ('EE_GENERAL', EnleveBlanc(Ligne, 7, 17,''));
    TL.PutValue ('EE_REFPOINTAGE', EnleveBlanc(Ligne, 24, 17,''));
    TL.PutValue ('EE_DATEOLDSOLDE', Str8ToDate((EnleveBlanc(Ligne, 41, 8,'01011900')),True));
    TL.PutValue ('EE_DATEPOINTAGE', Str8ToDate(EnleveBlanc(Ligne, 49, 8,'01011900'),True));
    TL.PutValue ('EE_OLDSOLDECRE', Valeur(StPoint(EnleveBlanc(Ligne, 57, 20, '0'))));
    TL.PutValue ('EE_NEWSOLDECRE', Valeur(StPoint(EnleveBlanc(Ligne, 77, 20, '0'))));
    TL.PutValue ('EE_OLDSOLDEDEB', Valeur(StPoint(EnleveBlanc(Ligne, 97, 20, '0'))));
    TL.PutValue ('EE_NEWSOLDEDEB', Valeur(StPoint(EnleveBlanc(Ligne, 117, 20, '0'))));
    TL.PutValue ('EE_OLDSOLDEDEBEURO', Valeur(StPoint(EnleveBlanc(Ligne, 137, 20, '0'))));
    TL.PutValue ('EE_OLDSOLDECREEURO', Valeur(StPoint(EnleveBlanc(Ligne, 157, 20, '0'))));
    TL.PutValue ('EE_NEWSOLDEDEBEURO', Valeur(StPoint(EnleveBlanc(Ligne, 177, 20, '0'))));
    TL.PutValue ('EE_NEWSOLDECREEURO', Valeur(StPoint(EnleveBlanc(Ligne, 197, 20, '0'))));
    TL.PutValue ('EE_SOCIETE', EnleveBlanc(Ligne, 217, 3,''));
    TL.PutValue ('EE_NUMRELEVE', EnleveBlanc(Ligne, 220, 10,'0'));
    TL.PutValue ('EE_RIB', EnleveBlanc(Ligne, 230, 35,''));
    TL.PutValue ('EE_DEVISE', EnleveBlanc(Ligne, 265, 3,''));
    TL.PutValue ('EE_DATESOLDE', Str8ToDate(EnleveBlanc(Ligne, 268, 8,'01011900'),True));
    TL.PutValue ('EE_OLDSOLDEDEBDEV', Valeur(StPoint(EnleveBlanc(Ligne, 276, 20, '0'))));
    TL.PutValue ('EE_OLDSOLDECREDEV', Valeur(StPoint(EnleveBlanc(Ligne, 296, 20, '0'))));
    TL.PutValue ('EE_NEWSOLDEDEBDEV', Valeur(StPoint(EnleveBlanc(Ligne, 316, 20, '0'))));
    TL.PutValue ('EE_NEWSOLDECREDEV', Valeur(StPoint(EnleveBlanc(Ligne, 336, 20, '0'))));
    TL.PutValue ('EE_NBMVT', EnleveBlanc(Ligne, 356, 10,'0'));
    TL.PutValue ('EE_ORIGINERELEVE', EnleveBlanc(Ligne, 366, 3,''));
    TL.PutValue ('EE_VALIDE', EnleveBlanc(Ligne, 369, 1,''));
    TL.PutValue ('EE_DATEINTEGRE', Str8ToDate(EnleveBlanc(Ligne, 370, 8,'01011900'), True));
    TL.PutValue ('EE_STATUTRELEVE', EnleveBlanc(Ligne, 378, 3,''));
    TL.PutValue ('EE_NUMERO', EnleveBlanc(Ligne, 381, 10,'1'));
    TL.PutValue ('EE_AVANCEMENT', EnleveBlanc(Ligne, 391, 1,''));
    TL.InsertOrUpdateDB(TRUE);
  finally
    TL.Free;
  end;
  if SaveLigne <> EnleveBlanc(Ligne, 7, 17,'') then
       OKexercice := OnAfficheListeCom('Import Relevé Bancaire ' + EnleveBlanc(Ligne, 7, 17,''), LISTEIMPORT);
  SaveLigne := EnleveBlanc(Ligne, 7, 17,'');
end;

procedure TImportCom.ImportContact(Ligne: string);
var
  TC              : TOB;
  Okmj,ExisteAuxi : Boolean;
  compteauxi      : string;
begin
  Okmj := TRUE;
  compteauxi := BourreOuTronque(EnleveBlanc(Ligne, 10, 35,''), fbAux);
// fihce 10557
  ExisteAuxi := ExisteSQL ('SELECT C_AUXILIAIRE FROM contact WHERE C_AUXILIAIRE="' + compteauxi + '" AND ' +
  'C_NUMEROCONTACT='+ EnleveBlanc(Ligne, 45, 6,'1'));
  if (NatureFichier = 'BAL') and BInterdireCreationTiers then
  begin
  // Okmj := FALSE;
     Okmj := ExisteSQL ('SELECT T_AUXILIAIRE FROM TIERS WHERE T_AUXILIAIRE="' + compteauxi + '"');
  end;
  if (ExisteAuxi and BMAJCPTE) or ((not ExisteAuxi) and Okmj) then
  begin
          TC := TOB.Create('CONTACT', nil, -1);
          try

              TC.PutValue('C_TYPECONTACT', EnleveBlanc(Ligne, 7, 3,'T'));
              TC.PutValue('C_AUXILIAIRE', compteauxi);
              TC.PutValue('C_NUMEROCONTACT', EnleveBlanc(Ligne, 45, 6,'1'));
              TC.PutValue('C_NATUREAUXI', EnleveBlanc(Ligne, 51, 3,''));
              TC.PutValue('C_PRINCIPAL', EnleveBlanc(Ligne, 54, 1,'X'));
              TC.PutValue('C_NOM', EnleveBlanc(Ligne, 55, 35,''));
              TC.PutValue('C_PRENOM', EnleveBlanc(Ligne, 90, 35,''));
              TC.PutValue('C_SERVICE', EnleveBlanc(Ligne, 125, 35,''));
              TC.PutValue('C_FONCTION', EnleveBlanc(Ligne, 160, 35,''));
              TC.PutValue('C_TELEPHONE', EnleveBlanc(Ligne, 195, 35,''));
              TC.PutValue('C_FAX', EnleveBlanc(Ligne, 220, 25,''));
              TC.PutValue('C_TELEX', EnleveBlanc(Ligne, 245, 25,''));
              TC.PutValue('C_RVA', EnleveBlanc(Ligne, 270, 250,''));
              TC.PutValue('C_SOCIETE', EnleveBlanc(Ligne, 520, 3,''));
              TC.PutValue('C_CIVILITE', EnleveBlanc(Ligne, 523, 3,''));
              TC.PutValue('C_FONCTIONCODEE', EnleveBlanc(Ligne, 526, 3,''));
              TC.PutValue('C_LIPARENT', EnleveBlanc(Ligne, 529, 3,''));
              TC.PutValue('C_JOURNAIS', EnleveBlanc(Ligne, 532, 6,'1'));
              TC.PutValue('C_MOISNAIS', EnleveBlanc(Ligne, 538, 6,'1'));
              TC.PutValue('C_ANNEENAIS', EnleveBlanc(Ligne, 544, 6,'0'));
              TC.PutValue('C_SEXE', EnleveBlanc(Ligne, 550, 3,''));
              TC.PutValue('C_PUBLIPOSTAGE', EnleveBlanc(Ligne, 553, 1,''));
              TC.PutValue('C_TEXTELIBRE1', EnleveBlanc(Ligne, 554, 35,''));
              TC.PutValue('C_TEXTELIBRE2', EnleveBlanc(Ligne, 589, 35,''));
              TC.PutValue('C_TEXTELIBRE3', EnleveBlanc(Ligne, 624, 35,''));
              TC.PutValue('C_BOOLLIBRE1', EnleveBlanc(Ligne, 659, 1,''));
              TC.PutValue('C_BOOLLIBRE2', EnleveBlanc(Ligne, 660, 1,''));
              TC.PutValue('C_BOOLLIBRE3', EnleveBlanc(Ligne, 661, 1,''));
              TC.PutValue('C_LIBRECONTACT1', EnleveBlanc(Ligne, 662, 3,''));
              TC.PutValue('C_LIBRECONTACT2', EnleveBlanc(Ligne, 665, 3,''));
              TC.PutValue('C_LIBRECONTACT3', EnleveBlanc(Ligne, 668, 3,''));
              TC.PutValue('C_DATELIBRE1', Str8ToDate(EnleveBlanc(Ligne, 671, 8,'01011900'),True));
              TC.PutValue('C_DATELIBRE2', Str8ToDate(EnleveBlanc(Ligne, 679, 8,'01011900'),True));
              TC.PutValue('C_DATELIBRE3', Str8ToDate(EnleveBlanc(Ligne, 687, 8,'01011900'),True));
              TC.PutValue('C_VALLIBRE1', EnleveBlanc(Ligne, 695, 10,'0'));
              TC.PutValue('C_VALLIBRE2', EnleveBlanc(Ligne, 710, 10,'0'));
              TC.PutValue('C_VALLIBRE3', EnleveBlanc(Ligne, 725, 10,'0'));
              TC.PutValue('C_TIERS', EnleveBlanc(Ligne, 740, 17,''));
              TC.PutValue('C_CLETELEPHONE', EnleveBlanc(Ligne, 757, 25,''));
              TC.InsertOrUpdateDB(TRUE);
        finally
          TC.Free;
        end;
        OKexercice := OnAfficheListeCom('Import Contacts ' + EnleveBlanc(Ligne, 10, 35,''), LISTEIMPORT);
  END;
end;

function TImportCom.EnleveBlanc(Chaine: string; x, y: integer ; ValNumeric : String = ''): string;
var
  st: string;
  IsNumeric : Boolean ;
begin
  IsNumeric:=ValNumeric<>'' ;
  st := Copy(Chaine, x, y);
  If IsNumeric Then St:=StStrFPoint(St) ;
  Result := Trim(St);
  If (Result='') And (ValNumeric<>'') Then Result:=ValNumeric ;
end;

procedure TImportCom.InitEcriture;
var Q3        : TQuery;
etablis       : string;
begin
{$IFDEF TTW}
 cWA.MessagesAuClient('COMSX.IMPORT','','TImportCom.InitEcriture') ;
{$ENDIF}

  if GrandNbAux and (TTiers <> nil) and ((TTiers.detail.count > 200) or (TTiersInsert.detail.count > 200)
  or (TTiersupdate.detail.count > 200)) then
  begin
       GrandNbAux := TRUE;
       InsertTiers;
       InitTiers (FALSE);
  end;

  if TEcrParJournal = nil then
  begin
    // pour création des journaux de paramsoc s'ils n'existent pas
      Insert_journal;
      if TGen = nil then
         InitGeneraux;
      if TTiers = nil then
         InitTiers;

    TEcrParJournal := TOB.Create('', nil, -1);
    if ListeEtab = nil then
    begin
       ListeEtab := TStringList.Create;
      // pour maj des section d'attente
       VerifAxeSection;
    end;
    // fiche 10396
    if (NatureFichier = 'DOS') or ((OrigineFichier = 'S1') and (NatureFichier = 'SYN')) then
    begin
         if TOBLettre = nil then
         begin
            TOBLettre := TOB.Create('', nil, -1);
            fHLettre  := THashTable.Create();
            fHLettre.useElfHash   := true;
            fHLettre.singleThread := true;
         end;
    end;
    Modesaisie := '-';
    if GetParamSocSecur ('SO_ETABLISDEFAUT', '') = '' then
    begin
    // on crée un établissement par défault
      InitSocietePCL ;
      InitEtablissement;
      Q3 := OpenSQl ('SELECT ET_ETABLISSEMENT from ETABLISS', True);
      if not Q3.EOF then etablis :=  Q3.Findfield ('ET_ETABLISSEMENT').Asstring;
      ferme (Q3);
      if etablis <> '' then
         SetParamsoc('SO_ETABLISDEFAUT', etablis);
    end;
    if GetParamSocSecur('SO_TAUXEURO', '6,55957') <= 0 then SetParamSoc ('SO_TAUXEURO', 6.55957);

    if (NatureFichier = 'JRL') and (STva = nil) then  // Fiche 10363
    begin
      Q3 := OpenSQl ('SELECT * FROM CHOIXCOD WHERE CC_TYPE="TX1"', True);
      While not Q3.EOF do
      begin
            AJouterListeSoldes (Q3.Findfield ('CC_CODE').Asstring, STva);
            Q3.Next;
      end;
      Ferme (Q3);
    end;
// obligatoire pour les exercices les chargés
{$IFNDEF EAGLSERVER}
    CHARGEMAGEXO(False);
{$ELSE}
   TCPContexte.GetCurrent.Exercice.ChargeMagExo ;
{$ENDIF}
  end;
end;

function TImportCom.AlimPiece(Ligne: string; var PieceS1 : string) : Boolean;
var
ClientSynchro : Boolean;
ExpertSynchro : Boolean;
PieceSx       : Boolean;
nump          : integer;
begin
   Result := FALSE;
   PieceSx := FALSE;
   if (OrigineEnvoie = 'CLI') and (EnleveBlanc(Ligne, 803, 6) <> '') then PieceSx := (not EstAlpha(EnleveBlanc(Ligne, 803, 6)))
   else
   if (OrigineEnvoie = 'EXP') and (EnleveBlanc(Ligne,  819, 6) <> '') then PieceSx := (not EstAlpha(EnleveBlanc(Ligne, 819, 6)));

   ClientSynchro := ((OrigineEnvoie = 'CLI') and ((NatureFichier = 'SYN') or (OrigineFichier = 'S1') or
  ((NatureFichier = 'DOS') and PieceSx)));                                     
   ExpertSynchro := ((OrigineEnvoie = 'EXP') and ((NatureFichier = 'SYN') or ((RepriseComplement = '-SYN') and (NatureFichier = 'DOS') and PieceSx)));
  if ClientSynchro then // origine client
  begin
     if (NatureFichier = 'DOS')  then
     begin
          if (OrigineFichier = 'S1') then
              PieceS1 := EnleveBlanc(Ligne, 803, 6)
          else
              numeropiece := EnleveBlanc(Ligne, 803, 6);
     end
     else
     begin
           if EnleveBlanc(Ligne, 819, 6) <> '' then
              numeropiece := IntTostr(StrToint(EnleveBlanc(Ligne, 819, 6)))   // pour enlever les 0 dans le cas 000001
           else
           if EnleveBlanc(Ligne, 803, 6) <> '' then
              PieceS1 := EnleveBlanc(Ligne, 803, 6);
     end;
  end
  else
  if ExpertSynchro then // origine expert
  begin
     // numéro de piece S5
     if (NatureFichier = 'DOS')  then
              numeropiece := EnleveBlanc(Ligne,  819, 6)
     else
     begin
           if EnleveBlanc(Ligne, 803, 6) <> '' then
              numeropiece := IntTostr(StrToint(EnleveBlanc(Ligne, 803, 6)))
           else
           if EnleveBlanc(Ligne, 819, 6) <> '' then
              PieceS1 := EnleveBlanc(Ligne, 819, 6);
     end;
  end
  else
  begin
        if (EnleveBlanc(Ligne, 152, 8) <> '')  and (EnleveBlanc(Ligne, 152, 8) <> '0') then
          numeropiece := EnleveBlanc(Ligne, 152, 8)
        else
          Result := TRUE;
  end;                                              // fiche 10340
  if (numeropiece <> '') and ((DecoupLon <> 0) or (OrigineFichier = 'SI')) then // pour enlever 0000001
  begin
       if Isnumeric(numeropiece) then
       begin
          nump := StrToInt(numeropiece);
          numeropiece := IntToStr(nump);
       end
       else Result := TRUE;
  end;
  if ((NatureFichier = 'JRL') and (Calculpiece <> 'N'))
  or ((NatureFichier = 'BAL') and (OrigineFichier = 'SI'))
  or ((Modesaisie = '-') and (OrigineFichier = 'SI')) // fiche 10340
  or ((NatureFichier = 'BAL') and (Modesaisie = '-'))
  or (ClientSynchro and (PieceS1 = '') and ((numeropiece = '') or (Valeur(numeropiece) = 0)))
  or (ExpertSynchro and (PieceS1 = '') and ((numeropiece = '') or (Valeur(numeropiece) = 0)))
  or ((not ExpertSynchro) and (not ClientSynchro)and (PieceS1 = '') and ((numeropiece = '') or (Valeur(numeropiece) = 0)))
  then
  Result := TRUE;
end;

function TImportCom.RechercheOptimise (const TobSrc: tob ; const Field: string ; const ValeurR: string; RechOpti : Boolean=TRUE; THbl : THashTable=nil ) : TOB ;
var
iStart,iMax,iNumChamp,iCount,iPrevCount  : integer ;
ValB                                     : string;
lt                                       : TOB;
begin
     result := nil ; lT := nil;
     if THbl <> nil then lT := TOB(THbl.Get(ValeurR));
     if lT <> nil then Result := lT
     else
     begin
               if (pos('_', ValeurR) = 0)  and (pos('-', ValeurR) = 0) and (pos('/', ValeurR) = 0)
               and (pos('\', ValeurR) = 0) and (pos('''', ValeurR) = 0) and RechOpti then
               begin
               if Assigned(TobSrc) and (TobSrc.Detail.Count>0) then
                  begin
                  // récupération du n° de champ
                  iNumChamp := TobSrc.Detail[0].GetNumChamp(Field) ;
                  if iNumChamp>0 then
                     begin
                     iStart     := 0 ;
                     iPrevCount := -1 ;
                     iCOunt     := 0 ;
                     iMax       := TobSrc.Detail.Count-1 ;
                     if TobSrc.Detail[iMax].GetValeur(iNumChamp) = ValeurR then result := TobSrc.Detail[iMax]
                     else if TobSrc.Detail[0].GetValeur(iNumChamp) = ValeurR then result := TobSrc.Detail[0]
                     else while not assigned(result) do
                             begin
                             if (iPrevCount=iCount) then break ;
                             iPrevCount := iCount ;
                             iCount := ((iMax - iStart) shr 1 + iStart) ;
                             ValB := TobSrc.Detail[iCount].GetValeur(iNumChamp);
                             if ValB = ValeurR then result :=  TobSrc.Detail[iCount]
                             else if ValB < ValeurR then
                                  iStart := iCount
                             else
                                  iMax := iCount ;
                             end ;
                     end ;
                  end ;
               end
               else
                         Result := TobSrc.FindFirst([Field], [ValeurR], FALSE);
     end;
end ;

function TImportCom.Okdoublon (Ligne : string; var TT : TOB; CJ : string; FindSurTob : Boolean=FALSE): boolean;
var
debit,credit       : double;
StD,STC,St,devise  : string;
QT                 : TQuery;
Exo,Auxil,Lib      : string;
TL                 : TOB;
ChampDoublon,Champ : string;
begin
     Result := FALSE;
     AffectationMontantdoublon (Ligne, debit, credit);
     Exo := QUELEXODT(Format_Date_HAL(EnleveBlanc(Ligne, 4, 8)));
     Auxil := '';
     if (EnleveBlanc(Ligne, 32, 17) <> '') and (Copy(Ligne, 31, 1) <> 'H')
     and (Copy(Ligne, 31, 1) <> 'A') and (Copy(Ligne, 31, 1) <> 'O')then
     Auxil := BourreOuTronque(EnleveBlanc(Ligne, 32, 17), fbAux);
     Lib   := VireTouteSub(EnleveBlanc(Ligne, 84, 35),'"');
     StD   := Formatfloat('###0.00', debit) ;
     StC   := Formatfloat('###0.00', credit) ;
     if not FindSurTob then
     begin
          StD := FindEtReplace(StD,',','.',TRUE) ;
          StC := FindEtReplace(StC,',','.',TRUE) ;
     end;
     devise := 'EUR';
     if (EnleveBlanc(Ligne, 160, 3) <> '') then devise := EnleveBlanc(Ligne, 160, 3);
     if (EnleveBlanc(Ligne, 160, 3) = '') and  (GetParamSocSecur('SO_DEVISEPRINC', 'EUR') <> '') then devise := GetParamSocSecur('SO_DEVISEPRINC', 'EUR');
     if FindSurTob then
     begin
             if TPieceCon = nil then exit;
             TL := TPieceCon.FindFirst(
             ['E_JOURNAL', 'E_EXERCICE', 'E_DATECOMPTABLE', 'E_NUMEROPIECE','E_GENERAL', 'E_AUXILIAIRE', 'E_DEBITDEV', 'E_CREDITDEV', 'E_DEVISE'],
             [CJ, Exo, Format_Date_HAL(EnleveBlanc(Ligne, 4, 8)),
             numeropiece, BourreOuTronque(EnleveBlanc(Ligne, 14, 17), fbGene), Auxil, Valeur(STD), Valeur(STC), devise], FALSE);
             if (TL <> nil) then
             begin
                      // fiche 10531
                       TT.PutValue('E_BLOCNOTE'          , TL.getvalue ('E_BLOCNOTE'));
                       // fiche 10510
                       TT.PutValue('E_TABLE0', TL.getvalue ('E_TABLE0'));
                       TT.PutValue('E_TABLE1', TL.getvalue ('E_TABLE1'));
                       TT.PutValue('E_TABLE2', TL.getvalue ('E_TABLE2'));
                       TT.PutValue('E_TABLE3', TL.getvalue ('E_TABLE3'));

                       if (EnleveBlanc(Ligne, 830, 1) = '1') then
                       begin
                               if TL.GetValue ( 'E_PAQUETREVISION') = 0 then
                                   TL.free;
                               Result := TRUE ;
                       end
                       else
                       begin
                               TT.PutValue('E_LETTRAGE'          , TL.getvalue ('E_LETTRAGE'));
                               TT.PutValue('E_LETTRAGEDEV'       , TL.getvalue ('E_LETTRAGEDEV'));
                               TT.PutValue('E_ECHE'              , TL.getvalue ('E_ECHE'));
                               TT.PutValue('E_ETATLETTRAGE'      , TL.getvalue ('E_ETATLETTRAGE'));
                               TT.PutValue('E_COUVERTURE'        , TL.getvalue ('E_COUVERTURE'));
                               TT.PutValue('E_COUVERTUREDEV'     , TL.getvalue ('E_COUVERTUREDEV'));
                               TT.PutValue('E_DATEPAQUETMAX'     , TL.getvalue ('E_DATEPAQUETMAX'));
                               TT.PutValue('E_DATEPAQUETMIN'     , TL.getvalue ('E_DATEPAQUETMIN'));
                               TT.PutValue('E_DATEECHEANCE'      , TL.getvalue ('E_DATEECHEANCE'));
                               TT.PutValue('E_NUMECHE'           , TL.getvalue ('E_NUMECHE'));
                               if (GetParamSocSecur ('SO_CPPOINTAGESX', 'EXP') = 'EXP')
                               and (EnleveBlanc(Ligne, 430, 17) = '') and (TL.getvalue ('E_REFPOINTAGE') <> '') then
                               begin
                                    TT.PutValue('E_REFPOINTAGE'  , TL.getvalue ('E_REFPOINTAGE'));
                                    TT.PutValue('E_DATEPOINTAGE' , TL.getvalue ('E_DATEPOINTAGE'));
                               end;
                               // cas E_REFREVISION reste à zéro sur les lignes non modifiées de la pièce
                               //exemple initialisation vers expert + modification compte 1ere ligne + synchro
                               //TT.PutValue( 'E_REFREVISION'      , TL.getvalue ('E_REFREVISION')) ;
                               TT.PutValue ( 'E_PAQUETREVISION'  , TL.getvalue ('E_PAQUETREVISION') ) ;
                               TT.PutValue ( 'E_IO'              , TL.getvalue ('E_IO')) ;
                               TL.free;
                               Result := TRUE ;
                       end;
             end;
     end
     else
     begin
               if (Copy(Ligne, 31, 1) <> 'A')  and (Copy(Ligne, 31, 1) <> 'O') then
               begin
                 //Fiche 10217
                  if GetParamSocSecur('SO_CPCONTROLEDOUBLON', FALSE) and (GetParamSocSecur('SO_CPCHAMPDOUBLON', '') <> '') then
                  begin
                        ChampDoublon := GetParamSocSecur('SO_CPCHAMPDOUBLON', False);
                        if  ChampDoublon = 'LIB' then Champ := ' AND E_LIBELLE="'+Lib+'"'
                        else
                        if  ChampDoublon = 'RFI' then Champ := 'AND E_REFINTERNE="'+ EnleveBlanc(Ligne, 49, 35)+'" '
                        else
                        if ChampDoublon = 'RFX' then Champ:=' AND E_REFEXTERNE="' + EnleveBlanc(Ligne, 223, 35) +'" '
                        else
                        if ChampDoublon = 'RFL' then Champ:=' AND E_REFLIBRE="' + EnleveBlanc(Ligne, 351, 35) +'" ' ;
                        St:='SELECT E_JOURNAL, E_EXERCICE, E_DATECOMPTABLE, E_NUMEROPIECE, E_NUMLIGNE, E_NUMECHE, E_QUALIFPIECE, E_REFINTERNE '+
                            'FROM ECRITURE WHERE E_JOURNAL="'+CJ+'" AND E_EXERCICE="'+Exo+'" AND E_DATECOMPTABLE="'+USDATETIME(Format_Date_HAL(EnleveBlanc(Ligne, 4, 8)))
                            +'" AND E_GENERAL="'+BourreOuTronque(EnleveBlanc(Ligne, 14, 17), fbGene)+'" AND '
                            +'E_AUXILIAIRE="'+Auxil+'" ' + 'AND E_QUALIFPIECE="'+EnleveBlanc(Ligne, 151, 1)+'" '+Champ

                  end
                  else
                      St:='SELECT E_JOURNAL, E_EXERCICE, E_DATECOMPTABLE, E_NUMEROPIECE, E_NUMLIGNE, E_NUMECHE, E_QUALIFPIECE, E_REFINTERNE '+
                      'FROM ECRITURE WHERE E_JOURNAL="'+CJ+'" AND E_EXERCICE="'+Exo+'" AND E_DATECOMPTABLE="'+USDATETIME(Format_Date_HAL(EnleveBlanc(Ligne, 4, 8)))
                      +'" AND E_GENERAL="'+BourreOuTronque(EnleveBlanc(Ligne, 14, 17), fbGene)+'" AND '
                      +'E_AUXILIAIRE="'+Auxil+'" '
                      +' AND E_REFINTERNE="'+ EnleveBlanc(Ligne, 49, 35)+'" '
                      +' AND E_LIBELLE="'+Lib+'" AND E_DEVISE="'+devise
                      +'" AND E_DEBITDEV="'+STD+'" AND '+
                      'E_CREDITDEV="'+STC+'"';
               end
               else
                   St:='SELECT Y_JOURNAL, Y_EXERCICE, Y_DATECOMPTABLE, Y_NUMEROPIECE, Y_NUMLIGNE, Y_NUMVENTIL, Y_QUALIFPIECE, Y_REFINTERNE '+
                      'FROM ANALYTIQ WHERE Y_JOURNAL="'+CJ+'" AND Y_EXERCICE="'+Exo+'" AND Y_DATECOMPTABLE="'+USDATETIME(Format_Date_HAL(EnleveBlanc(Ligne, 4, 8)))+
                      '" AND Y_GENERAL="'+BourreOuTronque(EnleveBlanc(Ligne, 14, 17), fbGene)+'" AND '+
                      'Y_SECTION="'+Auxil+'"'
                      +' AND Y_REFINTERNE="'+ EnleveBlanc(Ligne, 49, 35)+'" '
                      +' AND Y_LIBELLE="'+Lib+'" AND Y_DEVISE="'+
                      devise+'" AND Y_DEBITDEV="'+STD+'" AND '+
                      'Y_CREDITDEV="'+STC+'" AND Y_AXE="'+EnleveBlanc(Ligne, 219, 2, 'A1')+'" ' ;

               QT := OpenSQL(St, TRUE, -1, '', TRUE) ;
               If Not QT.Eof Then
               begin
                    if (not BIntegre) then
                       OnAfficheListeCom('Doublon : '+ EnleveBlanc(Ligne, 0, 55),LISTEIMPORT);
                    Result := TRUE ;
               end;
               Ferme(QT);
     end;
end;

Function TImportCom.remplir_ecriture(Ext, Ligne, Modesaisie: string; periode: integer; T: TOB; Codejournal : string; DateEcriture : TDateTime; Exerc : TExoDate; CpteRemplace :string='') : string;
var
  Compte,Comptegene,Compteauxi  : string;
  mt                            : double;
  Refpiecesx,j,NumAxe           : integer;
  CodeLettre,TypeMvt            : string;
  qualifpiece                   : string;
  Etab,LibelleEcr               : string;
  Lettrage,Compteattente        : string;
  Existecpte,SExisteCpte        : Boolean;
  SavNatureCpte                 : string;
  Ajoutlettre                   : string;
  CAttentGen,CAttentAux         : string;
  TA                            : TOB;
  //lDev                          : RDevise;
  IndEtab                       : integer;
  OkAlp                         : Boolean;
  TL                            : TOB;
  Dateref,DateEche              : TDateTime;
  ChaineTmp,Cpte,Axe            : string;
  Q1                            : TQuery;
  ExisteSection                 : Boolean;
  St1, MDP, CleLettre           : string;
  CLettrage,CCotation,Couverture: string;
  SaveAuto                      : Boolean;
  SaveCreat,TypeMvmt            : string;
  Natpiece, Etatlettr           : string;
  T_Mode                        : T_ModeRegl;
  T_Info                        : TFTransfertInfo;
begin
  Result := ''; SaveAuto := TRUE; SExisteCpte:= TRUE; // à porter sur V7XX
  Compte := ''; T_Info.OKLettrage := FALSE; T_Info.Pointable := '';
  CAttentGen := ''; CAttentAux := '';
  CodeLettre := '';  DateEche := iDate1900; mt:=0;

  if OptionEchange then // cas échange
  begin
     SaveAuto := AutoriseEcheck;   SaveCreat := InterdireLaCreation;
  end;

  if ext = 'Y_' then inc(NbEcrAna);

  if (OrigineFichier = 'SI') and (EnleveBlanc(Ligne, 1, 2)='AA') then
       T.PutValue(Ext + 'DATECOMPTABLE', GetEncours.Deb)
  else
       T.PutValue(Ext + 'DATECOMPTABLE', DateEcriture);
// voir pour date paquet
  if ext = 'E_' then
  begin
    T.PutValue(Ext + 'DATEPAQUETMAX', DateEcriture);
    T.PutValue(Ext + 'DATEPAQUETMIN', DateEcriture);
    CLettrage := EnleveBlanc(Ligne, 1014, 5);
    Etatlettr := EnleveBlanc(Ligne, 1021, 3);
    if CLettrage <> '' then
    begin
            if (OrigineFichier = 'S1') or (OrigineFichier = 'QU') or (OrigineFichier = 'WT') or (OrigineFichier = 'WS') or (NatureFichier = 'SYN') or (stArg <> '') then
            begin
                 if Length(CLettrage) = 1 then
                 begin
                     if Etatlettr = 'PL' then
                          CodeLettre := 'aaaa'
                     else
                          CodeLettre := 'AAAA';
                 end
                 else
                 begin
                   if length (CLettrage) <= 3 then
                   begin
                    CodeLettre := CLettrage;
                    OkAlp := TRUE;
                    for j:=1 To  length(CodeLettre) do
                    begin
                            if not (CodeLettre[j] in Alpha) then
                            begin OkAlp := FALSE; break; end;
                    end;
                    if not OKalp then
                    begin
                         if Etatlettr = 'PL' then
                            CodeLettre := 'aaaa'
                         else
                            CodeLettre := 'AAAA';
                    end
                    else
                    begin
                         for j:=1 To  4-(length(CodeLettre)) do
                         begin
                              if Etatlettr = 'PL' then
                                 Ajoutlettre := Ajoutlettre+'a'
                              else
                                 Ajoutlettre := Ajoutlettre+'A';
                         end;
                         if Etatlettr = 'PL' then
                            CodeLettre := Ajoutlettre + CLettrage
                         else
                            CodeLettre := Ajoutlettre + CLettrage;
                    end;
                   end
                   else
                   begin
                         if Etatlettr = 'PL' then
                            CodeLettre := LowerCase(CLettrage)
                         else
                            CodeLettre := UpperCase(CLettrage);
                   end;
                 end;
                 T.PutValue(Ext + 'LETTRAGE', CodeLettre);
                 CodeLettre := CLettrage;

            end
            else
            T.PutValue(Ext + 'LETTRAGE', CLettrage);
    end;
  end;

  T.PutValue(Ext + 'EXERCICE', Exerc.Code);
  if ((stArg = '') and (((not PresenceExercice)) or ((NatureFichier = 'BAL') or (NatureFichier = 'JRL'))))
   or OptionEchange  or ((stArg <> '') and ControlExercClos) then // fiche 10558
  begin
     if (Exerc.EtatCpta <> 'OUV') AND (Exerc.EtatCpta <> 'CPR') and (Exerc.EtatCpta <> 'NON') then
     begin
                if BRejet then  // On rejete l'ecriture mais on continue le traitement   // fiche 10551
                begin
                  OKRejet := True;
                  if not BIntegre then OkCreatCompte := FALSE;
                  OnAfficheListeCom('L''exercice est clos. L''importation des écritures est impossible : ' + EnleveBlanc(Ligne, 0, 20) ,LISTEIMPORT);
                end
                else
                begin
                  // On sort de l'import.
                  AffMessage ('L''exercice est clos. L''importation des écritures est impossible.');
                  OKexercice := FALSE;
                  exit;
                end;
     end;
     if OptionEchange and (Exerc.EtatCpta = 'NON') then  // fiche 18675
     begin
                AffMessage ('L''exercice est non ouvert. L''importation des écritures est impossible.');
                OKexercice := FALSE; exit;
     end;
     // fiche 10201
     if DateEcriture < GetParamSocSecur('SO_DATECLOTUREPER', iDate1900) then
     begin
                if BRejet then  // On rejete l'ecriture mais on continue le traitement  // fiche 10551
                begin
                  OKRejet := True;
                  if not BIntegre then OkCreatCompte := FALSE;
                  OnAfficheListeCom('Certaines écritures correspondent à une période clôturée. L''importation des écritures est impossible : ' + EnleveBlanc(Ligne, 0, 20) ,LISTEIMPORT);
                end
                else
                begin
                    AffMessage ('Import impossible. Certaines écritures correspondent à une période clôturée.');
                    OKexercice := FALSE; exit;
                end;
     end;
  end;
  // fiche 10587
  if (stArg = '') and (ExerciceSuperVise <> '') and (Exerc.Code = ExerciceSuperVise) and ((NatureFichier = 'BAL') or (NatureFichier = 'JRL')) then
  begin
                if BRejet then
                begin
                  OKRejet := True;
                  if not BIntegre then OkCreatCompte := FALSE;
                  OnAfficheListeCom('Certaines écritures correspondent à une période supervisée. L''importation des écritures est impossible : ' + EnleveBlanc(Ligne, 0, 20) ,LISTEIMPORT);
                end
                else
                begin
                    AffMessage ('Import impossible. Certaines écritures correspondent à une période supervisé.');
                    OKexercice := FALSE; exit;
                end;
  end;

  T.PutValue(Ext + 'JOURNAL', Codejournal);
  Natpiece := EnleveBlanc(Ligne, 12, 2);

  // fiche 10472
  if (Natpiece <> 'AF') and (Natpiece <> 'AC') and (Natpiece <> 'ECC') and
  (Natpiece <> 'FC') and (Natpiece <> 'FF') and (Natpiece <> 'OC') and (Natpiece <> 'OD')
  and (Natpiece <> 'OF') and (Natpiece <> 'RC') and (Natpiece <> 'RF') then
  begin
      OnAfficheListeCom('Nature du mouvement '+ Natpiece + ' incorrecte, changement par OD' ,LISTEIMPORT);
      Natpiece := 'OD';
  end;
  
  T.PutValue(Ext + 'NATUREPIECE', Natpiece);

  Comptegene := EnleveBlanc(Ligne, 14, 17);
  CompteAuxi := EnleveBlanc(Ligne, 32, 17);
  LibelleEcr := VireTouteSub(EnleveBlanc(Ligne, 84, 35),'"');

  Compte := BourreOuTronque(Comptegene, fbGene);
  // si compte avec minuscule dedans
  Compte := UpperCase (compte);

  // application de la Table de correspondance
  Cpte := Compte;
  TraiteCorrespCpt(0, Cpte, Compte);

  T.PutValue(Ext + 'GENERAL', compte);
  AJouterListeSoldes (Compte, SGeneraux);

  if Ext = 'E_' then
  begin
                // On controle si le compte général ou Auxiliaire
                // existent, si non et à blanc on met un compte d'attente
                // selon la nature du compte
      T_Info.iCompte := Compte;
      SExisteCpte := ExisteCompte ('G_', T_Info);
      Existecpte := SExisteCpte;
      if Copy(Ligne, 31, 1) <> 'H' then //fiche 10547
      begin
                if (CompteAuxi <> '') and ((Comptegene = '') or (not Existecpte)
                or (SubCollCli <> '') or (SubCollFou <> '')) then
                begin  // si auxiliaire rempli
                       T_Info.iCompte := BourreOuTronque(CompteAuxi, fbAux);
                       Existecpte := ExisteCompte ('T_', T_Info, FALSE);

                       if OptionEchange and (not SExisteCpte) and (not Existecpte) then   // cas échange
                       begin
                           AutoriseEcheck := FALSE;   InterdireLaCreation := 'IGE;IAU';
                       end;
                       With T_Info do
                       begin
                           if (NatureCpte = 'CLI') then NatureCpte := 'COC';
                           if (NatureCpte = 'FOU') then NatureCpte := 'COF';
                           if (NatureCpte = 'SAL') then NatureCpte := 'COS';
                           if (NatureCpte = 'DIV') then NatureCpte := 'COD';
                       end;
                       // pas de compte de substitution
                       if ((SubCollCli = '') and (SubCollFou = '')) or (T_Info.Collectif = '') then
                       begin
                            Compteattente := RetournCompteattente (T_Info.NatureCpte);
                            // Si collectif et compte de remplacement inexistant renseigné
                            if (not SExistecpte) and (CpteGene <> '') and (Compteattente = CpteGene)
                            and (not AutoriseLaCreation ('IGE', 'Le compte tiers devrait être vide')) then
                            begin
                                          OnAfficheListeCom(' dans le fichier : ' + EnleveBlanc(Ligne, 1, 70),LISTEIMPORT);
                                          OKexercice := FALSE; exit;
                             end
                            else
                                          Existecpte := FALSE;
                       end
                       else
                       begin
                            Comptegene := T_Info.Collectif;
                            ExisteCpte := SExisteCpte;
                            // pour forcer le remplaçement par collectif du compte
                            if (Compte <> T_Info.Collectif)
                            and ((Compte = BourreOuTronque(SubCollCli, fbGene)) or
                                  (Compte = BourreOuTronque(SubCollFou, fbGene))) then
                            begin
                                          Compteattente := T_Info.Collectif; Existecpte := FALSE;
                                          SExistecpte := TRUE;
                            end;
                       end;
                end
                else
                begin
                       Compteattente := RetournCompteattente (T_Info.NatureCpte);
                       // Si compte de remplacement inexistant renseigné
                       if (not SExistecpte) and (CpteGene <> '') and (Compteattente = CpteGene) then
                       begin
                            SExistecpte    := TRUE;
                            Result := CpteGene; // pour signifier qu'il y a un remplacement de compte
                       end;
                end;

                if (Comptegene = '') and (CompteAuxi = '') then
                begin
                     if CpteGene <> '' then  Compteattente := CpteGene
                     else Compteattente := GetParamSocSecur ('SO_GENATTEND', '');
                end;

                if ((T_Info.NatureCpte = 'COC') or (T_Info.NatureCpte = 'COF') or (T_Info.NatureCpte = 'COS') or (T_Info.NatureCpte = 'COD'))
                 and (CompteAuxi = '') then
                      Ligne[31] := 'X';

                 if (Comptegene = '') or (not Existecpte) then
                 begin
                    if (not SExistecpte) and (Compte <> '') then
                    begin
                            if AutoriseLaCreation ('IGE', 'Le compte général '+ compte + ' est manquant.') then
                               InsertCompte ('G_', Compte, LibelleEcr, T_Info.NatureCpte, '', Etatlettr)
                            else  if BIntegre then
                            T.PutValue(Ext + 'AUXILIAIRE', 'INCOHERENT');

                            if OptionEchange and (not SExisteCpte) and (not Existecpte) then // cas échange
                                 AutoriseLaCreation ('IAU', 'Le compte auxiliaire '+ CompteAuxi + ' est manquant.');

                         if (T_Info.NatureCpte ='DIV') and (Etatlettr = 'TL') then
                            T_Info.OKLettrage := TRUE;
                         Result := Compte; // pour signifier qu'il y a une insertion de compte
                    end
                    else
                    begin
                            if OptionEchange and (not Existecpte) then // cas échange
                                 AutoriseLaCreation ('IAU', 'Le compte auxiliaire '+ CompteAuxi + ' est manquant.')
                            else
                            begin
                                    if Compteattente = '' then
                                    begin
                                     OnAfficheListeCom(ERR_COMPTEGENEMANQ ,LISTEIMPORT);
                                     OKexercice := FALSE; exit;
                                    end;
                                    T.PutValue(Ext + 'GENERAL', compteattente);
                                    if (T_Info.Collectif <> '') and ((SubCollCli <> '') or (SubCollFou <> '')) then
                                    begin
                                          OnAfficheListeCom('Le collectif du Compte '+ CompteAuxi+ ' est remplacé par '+compteattente ,LISTEIMPORT);
                                          OnAfficheListeCom(' dans le fichier : ' + EnleveBlanc(Ligne, 1, 70),LISTEIMPORT);
                                          OnAfficheListeCom('',LISTEIMPORT);
                                          CAttentGen := compteattente;
                                    end
                                    else
                                    begin
                                          OnAfficheListeCom('Compte '+ EnleveBlanc(Ligne, 14, 17)+ ' inexistant dans l''écriture ' + 'remplacé par '+compteattente ,LISTEIMPORT);
                                          OnAfficheListeCom(' dans le fichier : ' + EnleveBlanc(Ligne, 1, 70),LISTEIMPORT);
                                          OnAfficheListeCom('',LISTEIMPORT);
                                          CAttentGen := compteattente;
                                    end;
                            end;
                    end;

                 end;
      end;
      MajDernLettrage ('G_', Compte, T.GetValue(Ext + 'LETTRAGE'));
  end;

  T.PutValue(Ext + 'PERIODE', periode);
  if Ext = 'E_' then
    T.PutValue('E_MODESAISIE', Modesaisie);
  TypeMvt := Copy(Ligne, 31, 1);

  case TypeMvt[1] of
    'X', 'E', 'L':
      begin
        // CA - 15/04/2002 - Si le compte général n'est pas collectif,
        // on ne s'occupe pas de la zone auxiliaire. On génère l'écriture
        // sur le compte général.

        // CA - 26/10/2005 - Prise en compte de l'auxiliarisation des comptes de charges et produits avec SIC
        if (((T_Info.NatureCpte = 'COC') or (T_Info.NatureCpte='COF')
            or (T_Info.NatureCpte='COS') or (T_Info.NatureCpte='COD'))
            or (GetParamSocSecur('SO_IFDEFCEGID',False) and (CompteAuxi<>'') and ((T_Info.NatureCpte='CHA') or (T_Info.NatureCpte='PRO')))) then
        begin
          Compte := BourreOuTronque(CompteAuxi, fbAux);
          Compte :=FindEtReplace(Compte,'''','0',TRUE) ;
          Compte := UpperCase (Compte);

          // application de la Table de correspondance
          Cpte := Compte;
          TraiteCorrespCpt(1, Cpte, Compte);

          AJouterListeSoldes (Compte, STiers);

          T.PutValue(Ext + 'AUXILIAIRE', Compte);
          // maj dernierlettrage  auxiliaire
          if Ext = 'E_' then
          begin
                  // On controle si le compte Auxiliaire
                  // existe, si non et à blanc on met un compte d'attente
                  // selon la nature du compte
                   SavNatureCpte := T_Info.NatureCpte;
                   T_Info.iCompte := Compte;
                   SExisteCpte   := ExisteCompte ('T_', T_Info);
                   ExisteCpte    := SExisteCpte;

                   Compteattente := RetournCompteattente (T_Info.NatureCpte);
                   if (SavNatureCpte = 'COC') and (T_Info.NatureCpte <> 'CLI') and (T_Info.NatureCpte <> 'DIV') then
                      Existecpte := FALSE;
                   if (SavNatureCpte = 'COF') and (T_Info.NatureCpte <> 'FOU') and (T_Info.NatureCpte <> 'DIV') then
                      Existecpte := FALSE;
                   if (SavNatureCpte = 'COS') and (T_Info.NatureCpte <> 'SAL') and (T_Info.NatureCpte <> 'DIV') then
                      Existecpte := FALSE;
                   if (not Existecpte) and (Comptegene <> '') then
                   begin
                        // si auxiliaire rempli
                         T_Info.iCompte := BourreOuTronque(Comptegene, fbGene);
                         ExisteCompte ('G_', T_Info, FALSE);
                         if (T_Info.NatureCpte = 'COC') then
                         begin
                              T_Info.NatureCpte := 'CLI';
                              if (CpteClient <> '') then CompteAuxi := '';
                         end
                         else
                         if (T_Info.NatureCpte = 'COF') then
                         begin
                              T_Info.NatureCpte := 'FOU';
                              if CpteFour <> '' then CompteAuxi := '';
                         end
                         else
                         if (T_Info.NatureCpte = 'COS') then
                         begin
                              T_Info.NatureCpte := 'SAL';
                              if CpteSalaire <> '' then CompteAuxi := '';
                         end
                         else
                         if (T_Info.NatureCpte = 'COD') then
                         begin
                              T_Info.NatureCpte := 'DIV';
                              if CpteDivers <> '' then  CompteAuxi := '';
                         end
                         else
                         begin
                              if CpteGene <> '' then  CompteAuxi := '';
                         end;
                         Compteattente := RetournCompteattente (T_Info.NatureCpte);
                         Existecpte := FALSE;
                   end;

                   if (CompteAuxi = '') or (not Existecpte) then
                   begin
                         if (not SExisteCpte)  and (CompteAuxi <> '') then
                         begin
                            if AutoriseLaCreation ('IAU', 'Le compte auxiliaire '+ compte + ' est manquant.') then
                               InsertCompte ('T_', Compte, LibelleEcr, T_Info.NatureCpte, CAttentGen)
                            else  if BIntegre then
                            T.PutValue(Ext + 'AUXILIAIRE', 'INCOHERENT');

                            T_Info.OKLettrage := TRUE;
                         end
                         else
                         begin
                             if Compteattente = '' then
                             begin
                               OnAfficheListeCom(ERR_COMPTEAUXIMANQ + ' ' + EnleveBlanc(Ligne, 0, 20),LISTEIMPORT);
                               OKexercice := FALSE; exit;
                             end;
                              T.PutValue(Ext + 'AUXILIAIRE', Compteattente);
                              CAttentAux := compteattente;
                              Compte := CAttentAux;
                              T_Info.iCompte := Compte;
                              ExisteCompte ('T_', T_Info);
                              OnAfficheListeCom('Compte '+ EnleveBlanc(Ligne, 32, 17) +' inexistant dans l''écriture ' + 'remplacé par '+compteattente ,LISTEIMPORT);
                              OnAfficheListeCom(' dans le fichier : ' + EnleveBlanc(Ligne, 1, 70),LISTEIMPORT);
                              OnAfficheListeCom('',LISTEIMPORT);
                         end;
                   end;
             MajDernLettrage ('T_', Compte, T.GetValue(Ext + 'LETTRAGE'));
          end;
        end
        else
        begin
                 if CompteAuxi <> '' then
                 begin
                   if OptionEchange  then   // cas échange
                   begin
                           AutoriseEcheck := FALSE;   InterdireLaCreation := 'IGE;IAU';
                   end;
                   if not AutoriseLaCreation ('IGE;IAU', Comptegene +ERR_COMPTECOLL1+ 'compte auxiliaire :'+ CompteAuxi) then
                       T.PutValue(Ext + 'AUXILIAIRE', 'INCOHERENT')
                    else
                    begin
                         if not AutoriseLaCreation ('IAU', Comptegene +ERR_COMPTECOLL1+ 'compte auxiliaire :'+ CompteAuxi) then
                              T.PutValue(Ext + 'AUXILIAIRE', 'INCOHERENT')
                         else
                         begin
                              if (SubCollCli = '') or (SubCollFou = '') then
                                 OnAfficheListeCom (Comptegene +ERR_COMPTECOLL1+CompteAuxi+ERR_COMPTECOLL1suite,LISTEIMPORT);
                         end;
                    end;
                 end;
        end;
      end;
    'A', 'O','H':
      begin
        if (TypeMvt[1] = 'H') and (not SExistecpte) and (Compte <> '') then
        begin
              if AutoriseLaCreation ('IGE', 'Le compte général '+ compte + ' est manquant.') then
               begin
                 InsertCompte ('G_', Compte, LibelleEcr, T_Info.NatureCpte, '', Etatlettr);
                 Result := Compte;
               end
               else  if BIntegre then
                 T.PutValue(Ext + 'AUXILIAIRE', 'INCOHERENT');
        end;

        if Ext = 'E_' then
          T.PutValue(Ext + 'ANA', 'X');
        if Ext = 'Y_' then
        begin
          if TypeMvt[1] = 'O' then  // OD analytique
          begin
             T.PutValue(Ext + 'TYPEANALYTIQUE', 'X');
             T.PutValue (Ext+ 'NUMLIGNE', 0);
          end;

          Axe := EnleveBlanc(Ligne, 219, 2, 'A1');
          if Axe = 'A' then Axe := 'A1';

          if Axe = 'A1' then Compte := BourreOuTronque(CompteAuxi, fbAxe1) else
          if Axe = 'A2' then Compte := BourreOuTronque(CompteAuxi, fbAxe2) else
          if Axe = 'A3' then Compte := BourreOuTronque(CompteAuxi, fbAxe3) else
          if Axe = 'A4' then Compte := BourreOuTronque(CompteAuxi, fbAxe4) else
          if Axe = 'A5' then Compte := BourreOuTronque(CompteAuxi, fbAxe5);
          //Si compte de remplaçement géné + Compte est remplacé
          if (CpteGene <> '') and (CpteRemplace <> '') then
               T.PutValue(Ext + 'GENERAL', CpteRemplace);

          // application de la Table de correspondance
          Cpte := Compte;
          TraiteCorrespCpt(StrToInt (Axe[2])+1, Cpte, Compte);

          // chargement section de la base
          if SSectionInbase = nil then
           ChargeSection;
          IndEtab := SSectionInbase.IndexOf (compte+':'+Axe);
          ExisteSection := (IndEtab >= 0);
          if (SSectionFerme <> nil) and (SSectionFerme.Indexof(compte+':'+Axe) >=0) then // fiche 10413
          begin
              OnAfficheListeCom('Attention, la section ' + Compte + ' est fermée' ,LISTEIMPORT);
              OKexercice := FALSE;
          end;

          // fiche 10568 pour ne pas créer la section sur l'axe de déclaration tva
          if (not ExisteSection) and (AxeTva <> '') and (Axe = AxeTva) then
          begin                                                                                   
              OnAfficheListeCom('Attention, la section ' + Compte + ' n''a pas été créée car elle concerne l''axe de TVA' ,LISTEIMPORT);
              OKexercice := FALSE;
          end;

          if not ExisteSection then   // création de la section
          begin
               ChaineTmp := '';
               if (Axe = 'A1') and (RAxe1 <> '') then ChaineTmp := BourreOuTronque(RAxe1, fbAxe1) else
               if (Axe = 'A2') and (RAxe2 <> '') then ChaineTmp := BourreOuTronque(RAxe2, fbAxe2) else
               if (Axe = 'A3') and (RAxe3 <> '') then ChaineTmp := BourreOuTronque(RAxe3, fbAxe3) else
               if (Axe = 'A4') and (RAxe4 <> '') then ChaineTmp := BourreOuTronque(RAxe4, fbAxe4) else
               if (Axe = 'A5') and (RAxe5 <> '') then ChaineTmp := BourreOuTronque(RAxe5, fbAxe5);

               if (ChaineTmp = '') and (not AutoriseLaCreation ('ISE', 'La section '+ compte +' est manquante.'+'Axe :'+Axe)) then
               begin
                    if BIntegre then  T.PutValue(Ext + 'AUXILIAIRE', 'INCOHERENT');
               end
               else
               begin
                    if ChaineTmp <> '' then Compte := ChaineTmp
                    else
                    begin
                         St1:='***SAT'+Format_String(Compte,17)+Format_String(Compte,35)+Axe ;
                         SSectionInbase.Add (compte + ':'+ Axe);
                         ImportSectionAnalytique(St1);
                    end
               end;
          end;

          T.PutValue(Ext + 'SECTION', Compte);
          AJouterListeSoldes (Compte, SSection);
              // si je le dossier est en mode non croisé et le fichier en croisé il faut qd même renseigner les sousplans
          //if GetParamSocSecur('SO_CROISAXE', False) then
          begin
               if EnleveBlanc(Ligne, 431, 17) = '' then
               begin
                   if Axe = 'A1' then T.PutValue(Ext + 'SOUSPLAN1', Compte) else
                   if Axe = 'A2' then T.PutValue(Ext + 'SOUSPLAN2', Compte) else
                   if Axe = 'A3' then T.PutValue(Ext + 'SOUSPLAN3', Compte) else
                   if Axe = 'A4' then T.PutValue(Ext + 'SOUSPLAN4', Compte) else
                   if Axe = 'A5' then T.PutValue(Ext + 'SOUSPLAN5', Compte);
                   for NumAxe:=1 to 5 do
                   begin
                       if GetParamSocSecur ('SO_VENTILA'+ IntTostr(NumAxe), False)
                       and (T.GetValue(Ext + 'SOUSPLAN'+ IntTostr(NumAxe)) = '') then
                       begin
                         Q1 := OpenSql ('SELECT X_SECTIONATTENTE from axe where X_AXE="'+ Axe+'"', True);
                         if  not Q1.EOF then
                             T.PutValue(Ext + 'SOUSPLAN'+ IntTostr(NumAxe), Q1.findField ('X_SECTIONATTENTE').asstring);
                         ferme (Q1);
                       end;
                   end;
               end
               else
               begin
                   Compte:= EnleveBlanc(Ligne, 431, 17);
                   if Compte <> '' then
                   begin
                        Cpte := BourreOuTronque(Compte, fbAxe1);
                        T.PutValue(Ext + 'SOUSPLAN1', Cpte);
                   end;

                   Compte:= EnleveBlanc(Ligne, 448, 17);
                   if Compte <> '' then
                   begin
                        Cpte := BourreOuTronque(Compte, fbAxe2);
                        T.PutValue(Ext + 'SOUSPLAN2', Cpte);
                        if BSoldeLigne then
                           Ajoute (lpS, Axe, Cpte, T.GetValue (Ext + 'DEBIT'), T.GetValue (Ext + 'CREDIT'), -1, -1, T.GetValue(Ext + 'DATECOMPTABLE')) ;
                   end;
                   Compte:= EnleveBlanc(Ligne, 465, 17);
                   if Compte <> '' then
                   begin
                        Cpte := BourreOuTronque(Compte, fbAxe3);
                        T.PutValue(Ext + 'SOUSPLAN3', Cpte);
                        if BSoldeLigne then
                           Ajoute (lpS, Axe, Cpte, T.GetValue (Ext + 'DEBIT'), T.GetValue (Ext + 'CREDIT'), -1, -1, T.GetValue(Ext + 'DATECOMPTABLE')) ;
                   end;
                   Compte:= EnleveBlanc(Ligne, 482, 17);
                   if Compte <> '' then
                   begin
                        Cpte := BourreOuTronque(Compte, fbAxe4);
                        T.PutValue(Ext + 'SOUSPLAN4', Cpte);
                        if BSoldeLigne then
                           Ajoute (lpS, Axe, Cpte, T.GetValue (Ext + 'DEBIT'), T.GetValue (Ext + 'CREDIT'), -1, -1, T.GetValue(Ext + 'DATECOMPTABLE')) ;
                   end;
                   Compte:= EnleveBlanc(Ligne, 499, 17);
                   if Compte <> '' then
                   begin
                        Cpte := BourreOuTronque(Compte, fbAxe5);
                        T.PutValue(Ext + 'SOUSPLAN5', Cpte);
                        if BSoldeLigne then
                           Ajoute (lpS, Axe, Cpte, T.GetValue (Ext + 'DEBIT'), T.GetValue (Ext + 'CREDIT'), -1, -1, T.GetValue(Ext + 'DATECOMPTABLE')) ;
                   end;
               end;
          end;
        end;
      end;
  end;

  T.PutValue(Ext + 'REFINTERNE', EnleveBlanc(Ligne, 49, 35));
  T.PutValue(Ext + 'LIBELLE', LibelleEcr);

  if Ext = 'E_' then
  begin
    MDP := EnleveBlanc(Ligne, 119, 3);
    TraiteCorrespCpt(7, EnleveBlanc(Ligne, 119, 3), MDP);

    if MDP <> '' then AJouterListeSoldes (MDP, SModepaie);
    T.PutValue(Ext + 'MODEPAIE', MDP);
    // ajout me pour la date d'échéance  122 au lieu de 4
    T.PutValue(Ext + 'DATEECHEANCE', Format_Date_HAL(EnleveBlanc(Ligne, 122, 8)));
    DateEche := Format_Date_HAL(EnleveBlanc(Ligne, 122, 8));
    if not TransIsValidDate(DateToStr(Format_Date_HAL(EnleveBlanc(Ligne, 122, 8)))) then
       T.PutValue(Ext + 'DATEECHEANCE', iDate1900);
  end;

  if Ext = 'E_' then  T.PutValue(Ext + 'COTATION', 1);
    T.PutValue(Ext + 'TAUXDEV', 1.0);

  T.PutValue(Ext + 'DEVISE', EnleveBlanc(Ligne, 160, 3, V_PGI.DevisePivot));
  { BVE 04.07.07 : Test sur l'existence de la devise }
  if TDevise.FindFirst(['D_DEVISE'],[T.GetValue(Ext + 'DEVISE')],true) = nil then
  begin
     // Cas devise dans le fichier et qui n'est pas chargé // fiche 10566
     if ExisteSQl ('SELECT * from DEVISE Where D_DEVISE="'+T.GetValue(Ext + 'DEVISE')+'"') then
           RemplirTOBDevise
     else
     begin
           OnAfficheListeCom('Erreur : la devise ' + T.GetValue(Ext + 'DEVISE') + ' est inconnue.', LISTEIMPORT );
           if BIntegre and (not AutoriseEcheck) then
              // On force la devise avec la devise pivot
              T.PutValue(Ext + 'DEVISE', V_PGI.DevisePivot);
           if BRejet then
              // On rejete l'ecriture mais on continue le traitement
              OKRejet := True
           else
           begin
              // On sort de l'import.
              OKexercice := FALSE;
              exit;
           end;
     end;
  end;
  { END BVE 04.07.07 }

  CCotation := EnleveBlanc(Ligne, 163, 10);
  if CCotation <> '' then
  begin
   mt := Valeur(StPoint(CCotation));
   if Ext = 'E_' then T.PutValue(Ext + 'COTATION', mt)
   else
   begin
   // fiche  19917
    if not T.FieldExists(Ext + 'COTATION') then
      T.AddChampSupValeur(Ext + 'COTATION', mt);
   end;
   if mt <> 1.0 then mt := mt*V_PGI.TauxEuro
   else mt := 1.0;
   T.PutValue(Ext + 'TAUXDEV', mt);
  end;
  // E_cotation pas renseigné
  if (CCotation = '') or (mt = 0) then
  begin
(*       if T.GetValue( Ext + 'DEVISE' ) = '' then
        T.PutValue(Ext + 'DEVISE', 'EUR');
*)
        if (CCotation = '') then
        begin
               if (Ext = 'E_') then
               begin
                   if ZDevise = nil then ZDevise := TZDevise.create;
                   ZDevise.Load([T.GetValue( Ext + 'DEVISE' )]) ;
                   ZDevise.AffecteTaux(DateEcriture) ;
                   T.PutValue(Ext + 'TAUXDEV', ZDevise.Dev.Taux);
                   if Ext = 'E_' then CSetCotation (T) ;
                   mt := T.GetValue(Ext + 'COTATION');
                   YCotation := mt;
               end
               else // cas analytique
               begin
                  if YCotation <> 0 then
                  begin
                    mt := YCotation;
                    T.PutValue(Ext + 'TAUXDEV', ZDevise.Dev.Taux);
                  end;
               end;
         end;

         if Ext = 'E_' then T.PutValue(Ext + 'COTATION', mt)
         else
         begin
         // fiche  19917
          if not T.FieldExists(Ext + 'COTATION') then
            T.AddChampSupValeur(Ext + 'COTATION', mt);
         end;
  end;

// les couvertures
   if (Ext = 'E_') then
   begin
       Couverture := EnleveBlanc(Ligne, 938, 20);
       if Couverture <> '' then
        T.PutValue(Ext + 'COUVERTURE', Valeur(StPoint(Couverture)))
       else
        T.PutValue(Ext + 'COUVERTURE', 0);

       Couverture := EnleveBlanc(Ligne, 958, 20);
       if Couverture <> ''  then
        T.PutValue(Ext + 'COUVERTUREDEV', Valeur(StPoint(Couverture)))
       else
        T.PutValue(Ext + 'COUVERTUREDEV', 0);

  end;

  AffectationMontant(Ligne, T, Ext);
// calcul des couvetures
  AffectationMontantcouverture(Ligne, T, Ext);

  qualifpiece := EnleveBlanc(Ligne, 151, 1);

  if qualifpiece <> '' then
  begin
      if (stArg <> '') and (qualifpiece = 'S') and(TypeEcrR = 'S') and (TypeEcrI = 'N')then
        T.PutValue(Ext + 'QUALIFPIECE', TypeEcrI)
      else
      begin
          { CA - 08/07/2005 - Si on lance ComSx avec un argument (depuis le bureau par exemple ) et sans fichier INI,
          TypeEcrI est à ''. On a alors TypeEcrI <> TypeEcrR et on affectait E_QUALIFPIECE à "" ==> problème
          J'ajoute donc un contrôle supplémentaire pour tester TypeEcrI<>'' }
          // if  TypeEcrI <> TypeEcrR then
          if  ((TypeEcrI <> '') and (TypeEcrI <> TypeEcrR)) then
                  T.PutValue(Ext + 'QUALIFPIECE', TypeEcrI)
          else
          begin
              if (qualifpiece <> 'N') and (qualifpiece <> 'S') and (qualifpiece <> 'U') and (qualifpiece <> 'P')
              and (qualifpiece <> 'R') and  (qualifpiece <> 'C')  then
                  T.PutValue(Ext + 'QUALIFPIECE', 'N')
              else
                  T.PutValue(Ext + 'QUALIFPIECE', qualifpiece);
          end;
      end;
  end
  else
    T.PutValue(Ext + 'QUALIFPIECE', 'N');

  if EnleveBlanc(Ligne, 216, 3) <> '' then
  begin
    Etab := EnleveBlanc(Ligne, 216, 3);
    TraiteCorrespCpt(9, EnleveBlanc(Ligne, 216, 3), Etab);

    if BControlEtab or ((not AutoriseEcheck) and (pos('IET', InterdireLaCreation) <> 0)) then
    begin
        if BIntegre or BRejet then
        begin
               if not ExisteSql ('SELECT ET_ETABLISSEMENT from ETABLISS WHERE ET_ETABLISSEMENT="'+Etab+'"') then
               begin
                  if BIntegre and (not AutoriseEcheck) then  T.PutValue(Ext + 'AUXILIAIRE', 'INCOHERENT');
                  if BRejet then OKRejet := TRUE;
               end;
        end;
    end;

    T.PutValue(Ext + 'ETABLISSEMENT', Etab);
    if ListeEtab.Count = 0 then ListeEtab.Add(Etab);
    IndEtab := ListeEtab.Indexof (Etab);
    if IndEtab < 0 then ListeEtab.Add(Etab);
  end
  else
  begin   //fiche 21967
      if OptionEchange and (GetParamSocSecur ('SO_CPEDIETABLISSEMENTDEF', '') <> '') then
              T.PutValue(Ext + 'ETABLISSEMENT', GetParamSocSecur ('SO_CPEDIETABLISSEMENTDEF', GetParamSocSecur ('SO_ETABLISDEFAUT', '001')))
      else
              T.PutValue(Ext + 'ETABLISSEMENT', GetParamSocSecur ('SO_ETABLISDEFAUT', '001'));
  end;
  T.PutValue(Ext + 'SEMAINE', NumSemaine(DateEcriture)) ;

  if Ext = 'Y_' then
    T.PutValue(Ext + 'AXE', EnleveBlanc(Ligne, 219, 2, 'A1'));

  if Ext = 'E_' then
  begin
   if EnleveBlanc(Ligne, 221, 2) <> '' then
    T.PutValue(Ext + 'NUMECHE', EnleveBlanc(Ligne, 221, 2))
   else
    T.PutValue(Ext + 'NUMECHE', 0);
    // fiche 10322
   if BValider then T.PutValue(Ext + 'VALIDE', 'X');
  end;
  if Copy((T.GetValue (Ext+'LIBELLE')),1,12) = 'DETRUITE LE ' then
  begin
    T.PutValue(Ext + 'CREERPAR', 'DET');
    T.PutValue(Ext + 'VALIDE', 'X');
  end
  else
   // en mode étendu
   if EnleveBlanc(Ligne, 223, 35) <> '' then
    T.PutValue(Ext + 'REFEXTERNE', EnleveBlanc(Ligne, 223, 35));
   if EnleveBlanc(Ligne, 258, 8) <> '' then
   begin
    T.PutValue(Ext + 'DATEREFEXTERNE', Format_Date_HAL(EnleveBlanc(Ligne, 258, 8)));
    if not TransIsValidDate(DateToStr(Format_Date_HAL(EnleveBlanc(Ligne, 258, 8)))) then
       T.PutValue(Ext + 'DATEREFEXTERNE', iDate1900);
   end;
   if EnleveBlanc(Ligne, 266, 8) <> '' then
   begin
    T.PutValue(Ext + 'DATECREATION', Format_Date_HAL(EnleveBlanc(Ligne, 266, 8)));
    if not TransIsValidDate(DateToStr(Format_Date_HAL(EnleveBlanc(Ligne, 266, 8)))) then
       T.PutValue(Ext + 'DATECREATION', iDate1900);
    if NatureFichier = 'JRL' then
         T.PutValue(Ext + 'DATECREATION',Date);
   end;

   if EnleveBlanc(Ligne, 274, 3) <> '' then
   begin
          if BCtrS then  // fiche 10362
          begin
               if EnleveBlanc(Ligne, 274, 3) <> V_PGI.CodeSociete then
               begin
                    OnAfficheListeCom('Code société '+ EnleveBlanc(Ligne, 274, 3)+ ' incohérent' ,LISTEIMPORT); OKexercice := FALSE;
               end;
          end;
          T.PutValue(Ext + 'SOCIETE', EnleveBlanc(Ligne, 274, 3));
   end;
   if EnleveBlanc(Ligne, 277, 17) <> '' then
    T.PutValue(Ext + 'AFFAIRE', EnleveBlanc(Ligne, 277, 17));
   if EnleveBlanc(Ligne, 294, 8) <> '' then
   begin
    T.PutValue(Ext + 'DATETAUXDEV', Format_Date_HAL(EnleveBlanc(Ligne, 294, 8))); // fiche 10522
    if not TransIsValidDate(DateToStr(Format_Date_HAL(EnleveBlanc(Ligne, 294, 8)))) then
       T.PutValue(Ext + 'DATETAUXDEV', iDate1900);
   end;

   if TOBAno <> nil then
   begin
       TA := TOBAno.FindFirst(['EXERCICE','JOURNAL'], [T.GetValue(Ext + 'EXERCICE'),T.GetValue(Ext + 'JOURNAL')], FALSE);
       if (TA <> nil) and (TA.GetValue('JOURNAL') <> '') then
       begin
                   if TA.GetValue ('ANO') <> '' then
                      T.PutValue(Ext + 'ECRANOUVEAU', TA.GetValue ('ANO'))
                   else
                      T.PutValue(Ext + 'ECRANOUVEAU', 'OAN');
       end
       else
                    T.PutValue(Ext + 'ECRANOUVEAU', EnleveBlanc(Ligne, 302, 3, 'N'));
       if (not OkMajExoV8) and (T.GetValue(Ext + 'ECRANOUVEAU') = 'H') then OkMajExoV8 := TRUE;
   end
   else
          T.PutValue(Ext + 'ECRANOUVEAU', EnleveBlanc(Ligne, 302, 3, 'N'));
   if EnleveBlanc(Ligne, 305, 20) <> '' then
    T.PutValue(Ext + 'QTE1', Valeur(StPoint(EnleveBlanc(Ligne, 305, 20))));
   if EnleveBlanc(Ligne, 325, 20) <> '' then
    T.PutValue(Ext + 'QTE2', Valeur(StPoint(EnleveBlanc(Ligne, 325, 20))));
   if EnleveBlanc(Ligne, 345, 3) <> '' then
    T.PutValue(Ext + 'QUALIFQTE1', EnleveBlanc(Ligne, 345, 3));
   if EnleveBlanc(Ligne, 348, 3) <> '' then
    T.PutValue(Ext + 'QUALIFQTE2', EnleveBlanc(Ligne, 348, 3));
   if EnleveBlanc(Ligne, 351, 35) <> '' then
    T.PutValue(Ext + 'REFLIBRE', EnleveBlanc(Ligne, 351, 35));
   if (EnleveBlanc(Ligne, 386, 1) <> '') and (Ext <> 'Y_') then
     T.PutValue(Ext + 'TVAENCAISSEMENT', EnleveBlanc(Ligne, 386, 1));
     // ajout me pour les encaissements
   if (Ext <> 'Y_') then
   begin
            if (T_Info.NatureCpte = 'COC') or (T_Info.NatureCpte='COF')
            or (T_Info.NatureCpte = 'FOU') or (T_Info.NatureCpte='CLI')
            or (T_Info.NatureCpte = 'TIC') or (T_Info.NatureCpte='TID') // fiche 10537
            or (T_Info.NatureCpte = 'SAL') // fiche 10324
            or (T_Info.NatureCpte='COS') or (T_Info.NatureCpte='COD') then
            begin
                 if T.GetValue (Ext + 'DEBIT') <> 0 then  T.PutValue(Ext + 'ENCAISSEMENT', 'ENC')
                 else
                 if T.GetValue (Ext + 'CREDIT') <> 0 then  T.PutValue(Ext + 'ENCAISSEMENT', 'DEC');
            end
            else
                 T.PutValue(Ext + 'ENCAISSEMENT', 'RIE');
            if (EnleveBlanc(Ligne, 387, 3) <> '') then
               T.PutValue(Ext + 'REGIMETVA', EnleveBlanc(Ligne, 387, 3))
            else
               T.PutValue(Ext + 'REGIMETVA', T_Info.Regimetva);
   end;
   if (Ext <> 'Y_') then // fiche 10363
   begin
       if EnleveBlanc(Ligne, 390, 3) <> '' then
            T_Info.ETva := EnleveBlanc(Ligne, 390, 3);
       if (NatureFichier = 'JRL') and (STva <> nil) and (T_Info.Etva <> '') then  // Test de l'existance de tva dans la base
       begin
            if (STva.Count <> 0) and (STva.Indexof (T_Info.ETva) < 0) then
                T_Info.ETva := GetParamsocSecur('SO_CODETVAGENEDEFAULT', 'TN');
       end
       else  if T_Info.Etva = '' then T_Info.ETva := GetParamsocSecur('SO_CODETVAGENEDEFAULT', 'TN');
       T.PutValue(Ext + 'TVA', T_Info.ETva);
   end;
   if EnleveBlanc(Ligne, 393, 3) <> '' then
    T.PutValue(Ext + 'TPF', EnleveBlanc(Ligne, 393, 3));
   if EnleveBlanc(Ligne, 396, 17) <> '' then
    T.PutValue(Ext + 'CONTREPARTIEGEN', BourreOuTronque(EnleveBlanc(Ligne, 396, 17), FbGene));
   if EnleveBlanc(Ligne, 413, 17) <> '' then
    T.PutValue(Ext + 'CONTREPARTIEAUX', BourreOuTronque(EnleveBlanc(Ligne, 413, 17), FbAux));
   if (EnleveBlanc(Ligne, 430, 17) <> '') and (Ext = 'E_') then
    T.PutValue(Ext + 'REFPOINTAGE', EnleveBlanc(Ligne, 430, 17));
   if (EnleveBlanc(Ligne, 447, 8) <> '') and (Ext = 'E_') then
   begin
    T.PutValue(Ext + 'DATEPOINTAGE', Format_Date_HAL(EnleveBlanc(Ligne, 447, 8)));
    if not TransIsValidDate(DateToStr(Format_Date_HAL(EnleveBlanc(Ligne, 447, 8)))) then
       T.PutValue(Ext + 'DATEPOINTAGE', iDate1900);
   end;
   if (EnleveBlanc(Ligne, 455, 8) <> '') and (Ext = 'E_') then
   begin
    T.PutValue(Ext + 'DATERELANCE', Format_Date_HAL(EnleveBlanc(Ligne, 455, 8)));
    if not TransIsValidDate(DateToStr(Format_Date_HAL(EnleveBlanc(Ligne, 455, 8)))) then
       T.PutValue(Ext + 'DATERELANCE', iDate1900);
   end;
   if (EnleveBlanc(Ligne, 463, 8) <> '') and (Ext = 'E_') then
   begin
    T.PutValue(Ext + 'DATEVALEUR', Format_Date_HAL(EnleveBlanc(Ligne, 463, 8)));
    if not TransIsValidDate(DateToStr(Format_Date_HAL(EnleveBlanc(Ligne, 463, 8)))) then
       T.PutValue(Ext + 'DATEVALEUR', iDate1900);

   end;
   if  (Ext = 'E_') then
   begin
         if (EnleveBlanc(Ligne, 471, 35) <> '') then
         begin
              { modif pour produflex ancien format }
              if TTiersCAU <> nil then
              begin
                    if pos ('/', EnleveBlanc(Ligne, 471, 35)) <> 0 then
                          T.PutValue(Ext + 'RIB', EnleveBlanc(Ligne, 471, 35))
                    else
                    if (CompteAuxi <> '') then
                          T.PutValue(Ext + 'RIB', T_Info.RibAuto);
              end
              else
                T.PutValue(Ext + 'RIB', EnleveBlanc(Ligne, 471, 35));
         end
         else
         begin
              if (CompteAuxi <> '') then
              T.PutValue(Ext + 'RIB', T_Info.RibAuto);
              //T.PutValue(Ext + 'RIB', ForceRib (NatureCpte, Compte));
         end;
   end;

   if (EnleveBlanc(Ligne, 506, 10) <> '') and (Ext = 'E_') then
    T.PutValue(Ext + 'REFRELEVE', EnleveBlanc(Ligne, 506, 10));
   if (EnleveBlanc(Ligne, 516, 17) <> '') and (Ext = 'E_') then
    T.PutValue(Ext + 'IMMO', EnleveBlanc(Ligne, 516, 17));
   if EnleveBlanc(Ligne, 533, 30) <> '' then
    T.PutValue(Ext + 'LIBRETEXTE0', EnleveBlanc(Ligne, 533, 30));
   if EnleveBlanc(Ligne, 563, 30) <> '' then
    T.PutValue(Ext + 'LIBRETEXTE1', EnleveBlanc(Ligne, 563, 30));
   if EnleveBlanc(Ligne, 593, 30) <> '' then
    T.PutValue(Ext + 'LIBRETEXTE2', EnleveBlanc(Ligne, 593, 30));
   if EnleveBlanc(Ligne, 623, 30) <> '' then
    T.PutValue(Ext + 'LIBRETEXTE3', EnleveBlanc(Ligne, 623, 30));
   if EnleveBlanc(Ligne, 653, 30) <> '' then
    T.PutValue(Ext + 'LIBRETEXTE4', EnleveBlanc(Ligne, 653, 30));
   if EnleveBlanc(Ligne, 683, 30) <> '' then
    T.PutValue(Ext + 'LIBRETEXTE5', EnleveBlanc(Ligne, 683, 30));
   if EnleveBlanc(Ligne, 713, 30) <> '' then
    T.PutValue(Ext + 'LIBRETEXTE6', EnleveBlanc(Ligne, 713, 30));
   if EnleveBlanc(Ligne, 743, 30) <> '' then
    T.PutValue(Ext + 'LIBRETEXTE7', EnleveBlanc(Ligne, 743, 30));
   if EnleveBlanc(Ligne, 773, 30) <> '' then
    T.PutValue(Ext + 'LIBRETEXTE8', EnleveBlanc(Ligne, 773, 30));

   Refpiecesx := 0;
   if (ext = 'E_')  then
   begin
         if (OrigineFichier = 'SI') and  (EnleveBlanc(Ligne, 277, 17) <> '') then
         begin
              ChaineTmp := EnleveBlanc(Ligne, 277, 17);
              if (ChaineTmp <> '') and (not EstAlpha(ChaineTmp))then
              T.PutValue (Ext + 'REFREVISION', StrToint (ChaineTmp));
         end;
         if (NatureFichier <> 'JRL') and ((OrigineFichier = 'S1') or (OrigineEnvoie = 'CLI')) and (EnleveBlanc(Ligne, 803, 30) <> '') then
         begin
                        ChaineTmp := EnleveBlanc(Ligne, 803, 6);
                        if (ChaineTmp <> '') and (not EstAlpha(ChaineTmp))then
                        begin
                           Refpiecesx := StrToint (EnleveBlanc(Ligne, 803, 6));
                        end;
                           T.PutValue (Ext + 'REFREVISION', Refpiecesx);
                        if (EnleveBlanc(Ligne, 809, 1) = 'X') then
                           T.PutValue (Ext + 'ETATREVISION', 'X')
                        else
                            T.PutValue (Ext + 'ETATREVISION', '-');
                        if (OrigineFichier = 'S1') then
                           T.PutValue(Ext + 'LIBRETEXTE9', Format ('%-23.23s%-4.4s',[Copy(Ligne, 803, 30),CodeLettre]))
                        else
                           T.PutValue(Ext + 'LIBRETEXTE9', Format ('%-30.30s',[Copy(Ligne, 803, 30)]));
         end
         else
         if (NatureFichier <> 'JRL') and ((OrigineFichier <> 'S1') or (OrigineEnvoie = 'EXP')) and (EnleveBlanc(Ligne, 819, 6) <> '') then
         begin
                        ChaineTmp := EnleveBlanc(Ligne, 819, 6);
                        if (ChaineTmp <> '') and (not EstAlpha(ChaineTmp)) then
                           Refpiecesx := StrToint (EnleveBlanc(Ligne, 819, 6));
                        T.PutValue (Ext + 'REFREVISION', Refpiecesx);
                        if (OrigineFichier = 'S1') then
                           T.PutValue(Ext + 'LIBRETEXTE9', Format ('%-23.23s%-4.4s',[Copy(Ligne, 803, 30),CodeLettre]))
                        else
                           T.PutValue(Ext + 'LIBRETEXTE9', Format ('%-30.30s',[Copy(Ligne, 803, 30)]))
         end;
   end;

   if (VersionFichier = '') or (StrToint(VersionFichier) <= 3) then
   begin
       if EnleveBlanc(Ligne, 833, 3) <> '' then
        T.PutValue(Ext + 'TABLE0',  Veriftablelibre('E00', EnleveBlanc(Ligne, 833, 3)));
       if EnleveBlanc(Ligne, 836, 3) <> '' then
        T.PutValue(Ext + 'TABLE1',  Veriftablelibre('E01', EnleveBlanc(Ligne, 836, 3)));
       if EnleveBlanc(Ligne, 839, 3) <> '' then
        T.PutValue(Ext + 'TABLE2', Veriftablelibre ('E02', EnleveBlanc(Ligne, 839, 3)));
       if EnleveBlanc(Ligne, 842, 3) <> '' then
        T.PutValue(Ext + 'TABLE3', Veriftablelibre ('E03', EnleveBlanc(Ligne, 842, 3)));
   end
   else
   begin
       if EnleveBlanc(Ligne, 1024, 17) <> '' then
        T.PutValue(Ext + 'TABLE0',  Veriftablelibre('E00', EnleveBlanc(Ligne, 1024, 17)));
       if EnleveBlanc(Ligne, 1041, 17) <> '' then
        T.PutValue(Ext + 'TABLE1',  Veriftablelibre('E01', EnleveBlanc(Ligne, 1041, 17)));
       if EnleveBlanc(Ligne, 1058, 17) <> '' then  // mettre pour plustard
        T.PutValue(Ext + 'TABLE2', EnleveBlanc(Ligne, 1058, 17));
//        T.PutValue(Ext + 'TABLE2', Veriftablelibre ('E02', EnleveBlanc(Ligne, 1058, 17)));
       if EnleveBlanc(Ligne, 1075, 17) <> '' then
        T.PutValue(Ext + 'TABLE3', EnleveBlanc(Ligne, 1075, 17));
//        T.PutValue(Ext + 'TABLE3', Veriftablelibre ('E03', EnleveBlanc(Ligne, 1075, 17)));

        if (StrToint(VersionFichier) >= 5)  and (Ext = 'E_') then
        begin
             T.PutValue(Ext + 'REFGESCOM', EnleveBlanc(Ligne, 1092, 35));
             if EnleveBlanc(Ligne, 1140, 3) <> '' then
                T.PutValue(Ext + 'TRESOSYNCHRO', EnleveBlanc(Ligne, 1140, 3))
             else
             begin  // fiche 10590
                if (TypeEcrI= 'N') and ((T_Info.OKLettrage) or (T_Info.Pointable = 'X')) then
                begin
                   if (not EstMultiSoc and (T_Info.NatureCpte <> 'DIV')) or EstMultiSoc  then
                      T.PutValue(Ext + 'TRESOSYNCHRO', ets_Nouveau)
                   else
                      T.PutValue(Ext + 'TRESOSYNCHRO', ets_Rien);
                end
                else
                   T.PutValue(Ext + 'TRESOSYNCHRO', ets_Rien)
             end;
        end;
        if (StrToint(VersionFichier) >= 6)  and (Ext = 'E_') then
        begin
             if EnleveBlanc(Ligne, 1143, 17) <> '' then
                T.PutValue(Ext + 'NUMTRAITECHQ', EnleveBlanc(Ligne, 1143, 17));
             if EnleveBlanc(Ligne, 1160, 17) <> '' then
                T.PutValue(Ext + 'NUMENCADECA', EnleveBlanc(Ligne, 1160, 17));
        end;
        if (StrToint(VersionFichier) >= 7)  and ((EnleveBlanc(Ligne, 1177, 1) = 'X'))then
                T.PutValue(Ext + 'VALIDE', 'X');
   end;

   TypeMvmt := EnleveBlanc(Ligne, 1127, 3);
   if TypeMvmt <> '' then     // fiche 10363
      T.PutValue(Ext + 'TYPEMVT', TypeMvmt)
   else
      T.PutValue(Ext + 'TYPEMVT', QuelTypeMvt (T.GetValue(Ext + 'AUXILIAIRE'), T_Info.NatureCpte, Natpiece));

   if EnleveBlanc(Ligne, 845, 20) <> '' then
    T.PutValue(Ext + 'LIBREMONTANT0', EnleveBlanc(Ligne, 845, 20));
   if EnleveBlanc(Ligne, 865, 20) <> '' then
    T.PutValue(Ext + 'LIBREMONTANT1', EnleveBlanc(Ligne, 865, 20));
   if EnleveBlanc(Ligne, 885, 20) <> '' then
    T.PutValue(Ext + 'LIBREMONTANT2', EnleveBlanc(Ligne, 885, 20));
   if EnleveBlanc(Ligne, 905, 20) <> '' then
    T.PutValue(Ext + 'LIBREMONTANT3', EnleveBlanc(Ligne, 905, 20));
   if EnleveBlanc(Ligne, 925, 8) <> '' then
   begin
    T.PutValue(Ext + 'LIBREDATE', Format_Date_HAL(EnleveBlanc(Ligne, 925, 8)));
    if not TransIsValidDate(DateToStr(Format_Date_HAL(EnleveBlanc(Ligne, 925, 8)))) then
       T.PutValue(Ext + 'LIBREDATE', iDate1900);
   end;
   if EnleveBlanc(Ligne, 933, 1) <> '' then
    T.PutValue(Ext + 'LIBREBOOL0', EnleveBlanc(Ligne, 933, 1));
   if EnleveBlanc(Ligne, 934, 1) <> '' then
    T.PutValue(Ext + 'LIBREBOOL1', EnleveBlanc(Ligne, 934, 1));
   if EnleveBlanc(Ligne, 935, 3) <> '' then
    T.PutValue(Ext + 'CONSO', EnleveBlanc(Ligne, 935, 3));

   if (EnleveBlanc(Ligne, 998, 8) <> '') and (Ext = 'E_') then
   begin
    T.PutValue(Ext + 'DATEPAQUETMAX', Format_Date_HAL(EnleveBlanc(Ligne, 998, 8)));
    if not TransIsValidDate(DateToStr(Format_Date_HAL(EnleveBlanc(Ligne, 998, 8)))) then
       T.PutValue(Ext + 'DATEPAQUETMAX', iDate1900);
   end;
   if (EnleveBlanc(Ligne, 1006, 8) <> '') and (Ext = 'E_')then
   begin
    T.PutValue(Ext + 'DATEPAQUETMIN', Format_Date_HAL(EnleveBlanc(Ligne, 1006, 8)));
    if not TransIsValidDate(DateToStr(Format_Date_HAL(EnleveBlanc(Ligne, 1006, 8)))) then
       T.PutValue(Ext + 'DATEPAQUETMIN', iDate1900);
   end;
   // fiche 10396
    if (Ext = 'E_') and  ((NatureFichier = 'DOS') or ((OrigineFichier = 'S1') and (NatureFichier = 'SYN') and (EnleveBlanc(Ligne, 830, 1) = '1')))then
    begin
        if (T.GetValue(Ext + 'LETTRAGE') <> '') and ((EnleveBlanc(Ligne, 998, 8) = '') or (EnleveBlanc(Ligne, 998, 8) = '01011900')) then
        begin
             Dateref := T.GetValue(Ext + 'DATECOMPTABLE');
//             TL := TOBLettre.FindFirst(['COMPTEGENE','COMPTEAUX' , 'LETTRAGE', 'ETATLETTRAGE'], [Comptegene, CompteAuxi, T.GetValue(Ext + 'LETTRAGE'), EnleveBlanc(Ligne, 1021, 3)], FALSE);
             CleLettre := Comptegene+';'+CompteAuxi+';'+T.GetValue(Ext + 'LETTRAGE')+';'+Etatlettr;
             TL := TOB(fHLettre.Get(CleLettre));
             if TL = nil then
             begin
                  TL := TOB.Create ('',TOBLettre,-1);
                  TL.AddChampSupValeur('COMPTEGENE', Comptegene);
                  TL.AddChampSupValeur('COMPTEAUX', CompteAuxi);
                  TL.AddChampSupValeur('LETTRAGE', T.GetValue(Ext + 'LETTRAGE'));
                  TL.AddChampSupValeur('DATEPAQUETMAX', Dateref);
                  TL.AddChampSupValeur('DATEPAQUETMIN',  Dateref);
                  TL.AddChampSupValeur('ETATLETTRAGE',  Etatlettr);
               //   TL.AddChampSupValeur('EXERCICE', T.getValue(Ext + 'EXERCICE'));
                  fHLettre.Put(CleLettre, TObject(TL));
             end
             else
             begin
                  if  (Dateref < TL.GetValue('DATEPAQUETMIN')) and (Dateref <> iDate1900) then
                       TL.putvalue('DATEPAQUETMIN',  Dateref);
                  if  T.GetValue(Ext + 'DATECOMPTABLE') > TL.GetValue('DATEPAQUETMAX') and (Dateref <> iDate1900) then
                       TL.putvalue('DATEPAQUETMAX',  T.GetValue(Ext + 'DATECOMPTABLE'));
             end;
        end;
    end;

   if Ext = 'E_' then
   begin
       if EnleveBlanc(Ligne, 1019, 1) <> '' then
        T.PutValue(Ext + 'LETTRAGEDEV', EnleveBlanc(Ligne, 1019, 1));
       if Etatlettr <> '' then
        T.PutValue(Ext + 'ETATLETTRAGE', Etatlettr)
       else
        T.PutValue(Ext + 'ETATLETTRAGE', 'RI');

       if  (Etatlettr = 'TL') then
        T.PutValue(Ext + 'ETAT', '---0GM000')
       else
        T.PutValue(Ext + 'ETAT', '0000000000');
       if T_Info.OKLettrage then
       begin
            if T.GetValue(Ext + 'MODEPAIE') = '' then
               T.PutValue(Ext + 'MODEPAIE', RendModePaiementParDefaut);
               // ajout me 122 au lieu 4
               if BCalculModeregle then    
               begin
                  RemplirDateEchance (T_Mode, DateEcriture , T );
                  T.PutValue(Ext + 'DATEECHEANCE', T_Mode.TabEche[1].DateEche);
                  DateEche := T_Mode.TabEche[1].DateEche;
                  Mdp := RendModePaiementParDefaut(T_Mode.ModeInitial);
                  if Mdp <> '' then
                    T.PutValue(Ext + 'MODEPAIE', Mdp);
               end
               else
               begin
                  if (EnleveBlanc(Ligne, 122, 8) = '') or  (EnleveBlanc(Ligne, 122, 8) = '01011900') then
                  begin
                     T.PutValue(Ext + 'DATEECHEANCE', DateEcriture);
                     DateEche := DateEcriture;
                  end;
               end;
            if (EnleveBlanc(Ligne, 221, 2) = '') or  (Valeur(EnleveBlanc(Ligne, 221, 2)) = 0) then
               T.PutValue(Ext + 'NUMECHE', 1);
            if (Etatlettr = '') or (Etatlettr = 'RI') then
               T.PutValue(Ext + 'ETATLETTRAGE', 'AL');
            T.PutValue(Ext + 'ECHE', 'X');
       end
       else
       begin
             // Fiche 10115 if T.GetValue(Ext + 'MODEPAIE') = '' then T.PutValue(Ext + 'MODEPAIE', 'DIV');
               T.PutValue(Ext + 'DATEECHEANCE', iDate1900);
               if EnleveBlanc(Ligne, 221, 2) <> '' then
                  T.PutValue(Ext + 'NUMECHE', 0);
               if Etatlettr = '' then
                  T.PutValue(Ext + 'ETATLETTRAGE', 'RI')
               else
               begin
                    if (OrigineFichier = 'S1') then  // à vérifier
                    begin
                         if Etatlettr <> 'RI' then
                            T.PutValue(Ext + 'ETATLETTRAGE', 'RI');
                    end;
               end;
               T.PutValue(Ext + 'ECHE', '-');
       end;
       // vérification du lettrage
       if (NatureFichier = 'DOS') or (NatureFichier = 'SYN') or (RepriseComplement = 'REJ') then
       begin
           if T.GetValue(Ext + 'LETTRAGE') <> '' then
           begin
               if (T.GetValue(Ext + 'ETATLETTRAGE') = 'RI') or (T.GetValue(Ext + 'ETATLETTRAGE') = 'AL')
               or (Length(T.GetValue(Ext + 'LETTRAGE')) < 4) then
               begin
                   if (T.GetValue(Ext + 'ETATLETTRAGE') <> 'RI') and (T.GetValue(Ext + 'ETATLETTRAGE') <> 'AL')
                   and (Length(T.GetValue(Ext + 'LETTRAGE')) < 4) then
                          T.PutValue(Ext + 'ETATLETTRAGE', 'AL');
                   CRemplirInfoLettrage(T);
               end
               else
               begin
                   if (T.GetValue(Ext + 'ETATLETTRAGE') = 'PL') then
                   begin
                       Lettrage := LowerCase(T.GetValue(Ext + 'LETTRAGE'));
                       T.PutValue(Ext + 'LETTRAGE', Lettrage);
                   end;
                   if (T.GetValue(Ext + 'ETATLETTRAGE') = 'TL') then
                   begin
                        Lettrage :=T.GetValue(Ext + 'LETTRAGE');
                        UpperCase(Lettrage);
                        T.PutValue(Ext + 'LETTRAGE', Lettrage);
                   end;
               end;
               if (CAttentGen <> '') then
               begin
                    T_Info.iCompte := CAttentGen;
                    ExisteCompte ('G_', T_Info);
               end;
               if (CAttentAux <> '') then
               begin
                    T_Info.iCompte := CAttentAux;
                    ExisteCompte ('T_', T_Info);
               end;
           end;
           if (T.GetValue(Ext + 'ETATLETTRAGE') = 'AL') then CRemplirInfoLettrage(T);
       end
       else
       begin
           if (NatureFichier = 'JRL') and (T.GetValue(Ext + 'ETATLETTRAGE') <> 'RI') then
                   CRemplirInfoLettrage(T);
       end;
       // CA - 16/04/2002 - si compte non lettrable, on positionne correctement
       // les infos de lettrage.
       if (not T_Info.OkLettrage) or (T.GetValue(Ext + 'ETATLETTRAGE') = 'RI') then
         CSupprimerInfoLettrage (T);

       // ajout me pour les échéances e_eche et e_numeche renseignement uniquement pour les nature banque te caisse
       if (T_Info.Pointable = 'X') and ((T_Info.NatureCpte='BQE') or (T_Info.NatureCpte='CAI'))  then
       begin
        CRemplirInfoPointage (T);
        if EnleveBlanc(Ligne, 119, 3) <> '' then
          T.PutValue(Ext + 'MODEPAIE', EnleveBlanc(Ligne, 119, 3));
       end;
       // ajout me 17/04/2003 car CRemplirInfoLettrage met la date comptable dans dateecheance
       if (T.GetValue(Ext + 'ETATLETTRAGE') <> 'RI')
       (* ajout me 02-02-2005and (DateEche >  StrToDate(T.Getvalue (Ext + 'DATEECHEANCE'))) *)then
           T.PutValue(Ext + 'DATEECHEANCE', DateEche);

   end;
   T.PutValue(Ext +'CONFIDENTIEL', '0');
   // fiche 14216
   if (VersionFichier <> '') and (StrToint(VersionFichier) >= 7) then
   begin
        if Ext = 'E_' then
        begin
           T.PutValue(Ext +'CONFIDENTIEL', EnleveBlanc(Ligne, 1237, 1, '0'));
           // ajout me 07/12/2005
           T.PutValue(Ext +'CFONBOK', EnleveBlanc(Ligne, 1238, 1, '-'));
        end
        else
            T.PutValue(Ext +'CONFIDENTIEL', EnleveBlanc(Ligne, 1178, 1, '0'));
   end;
   // fiche 10487
   if (VersionFichier <> '') and (StrToint(VersionFichier) >= 8) and (Ext = 'E_') then
   begin
        if EnleveBlanc(Ligne, 1239, 3, '') <> '' then
          T.PutValue(Ext +'CODEACCEPT', EnleveBlanc(Ligne, 1239, 3, 'NON'))
        else                                                        // fiche 10549
          T.PutValue(Ext +'CODEACCEPT', RendCodeAcceptation (T.GetValue(Ext + 'MODEPAIE')));
        // BVE 28.08.07 : Suivi des validations
{$IFDEF CERTIFNF}
        if EnleveBlanc(Ligne, 1242, 8, '') <> '' then
          T.PutValue(Ext +'DOCID', EnleveBlanc(Ligne, 1242, 8, '0'));
{$ENDIF}
        if EnleveBlanc(Ligne, 1278, 3, '') <> '' then
          T.PutValue(Ext +'QUALIFORIGINE', EnleveBlanc(Ligne, 1278, 3, 'COM'));
   end
   else  // fiche 10549
   begin
        if Ext = 'E_' then
        begin
            if T.GetValue(Ext + 'MODEPAIE') <> '' then
              T.PutValue(Ext +'CODEACCEPT', RendCodeAcceptation (T.GetValue(Ext + 'MODEPAIE')))
            else
              T.PutValue(Ext +'CODEACCEPT', EnleveBlanc(Ligne, 1239, 3, 'NON'))
        end;
   end;
(*
    - Dans le champs E_CREERPAR :
            - "EXP" si provient d'expert : Enr 6 sur 3 positions de la première ligne d'un TRA
            - "CLI" si provient d'un client : Enr 6 sur 3 positions de la première ligne d'un TRA
            - "COM" si pas d'info particulière : EX fichier trt de siscoII ou d'un import TRA sans entête ( ex : la mode )
    - Dans le champs E_QUALIFORIGINE, avoir le type de l'export ( 2 premiers caractères) et la série concernée (3 ième caractère :
    Les différents type d'export sont présents sur l' Enr 9 sur 3 positions de la première ligne d'un TRA et sur l'enr 20 sur 1 position de la deuxième ligne d'un trt.
            - "BA" : Balance, "DO" : Dossier, "EX" : Exercice, "JR" : Journal, "SY" : Synchronisation, "ND" : si non défini : Fichier sans entête
    Les différentes séries sont présentes sur l'enr 5 sur une position de la première ligne d'un TRA et sera SISCOII si le fichier est un trt :
            - "1" si provient de S1, "3" si provien de S3, "5" si provient de S5, "7" si provient de S7, "S" si provient de SISCOII
            - "0" si non défini
*)

  if (OrigineFichier <> 'SI') and (QualifOrigine <> 'ND0') then
    T.PutValue(Ext + 'CREERPAR', OrigineEnvoie)
  else
    T.PutValue(Ext + 'CREERPAR', 'COM');
  {$IFNDEF CERTIFNF}
  if Ext = 'E_' then
          T.PutValue(Ext + 'QUALIFORIGINE', QualifOrigine);
  {$ELSE}
    // on met dans qualiforigine rien et dans  e_creerpar le code CPJ_CODE de CPJALIMPORT
    T.PutValue(Ext + 'CREERPAR', QualifOrigine) ;
  {$ENDIF}
  // ANouveau dynamique
  if Ext = 'E_' then
  begin
         if GetParamSocSecur('SO_CPANODYNA',false) then
         begin
               if TSAnv = nil then  TSAnv :=TList.Create ;
               AjouteAno(TSAnv, T, T_Info.NatureCpte, False) ;
         end;
  end;
 if OptionEchange then // cas échange
 begin
     AutoriseEcheck := SaveAuto;   InterdireLaCreation := SaveCreat;
 end;

 if OptionTP and (Ext = 'E_') then // Tiers payeurs
 begin
      if (T_Info.TPayeur <> '') and (T.GetValue(Ext + 'QUALIFPIECE') = 'N') and ((Natpiece = 'FC') or (Natpiece='FF') or (Natpiece='AC') or (Natpiece='AF'))then
      begin
            if ((GetParamSocSecur('SO_JALVTP','') <> '') and ((Natpiece='FC') or (Natpiece='AC'))) or
            ((GetParamSocSecur('SO_JALATP','') <> '') and ((Natpiece='FF') or (Natpiece='AF'))) then
            begin
                 ExisteTP := TRUE;
                 T.PutValue(Ext + 'FLAGECR', 'TP');
            end;
      end;
 end;

end;

function TImportCom.OKVentilable (Compte : string) : Boolean;
var
TA1,TA2,TA3   : TOB;
OKVentil      : Boolean;
NumAxe        : integer;
QDos          : TQuery;
begin
     FillChar(VentAttent,Sizeof(VentAttent),#0);
     OKVentil := FALSE;

     if TGen <> nil then
     begin
        TA1 := TGenUpdate.FindFirst(['G_GENERAL'], [Compte], FALSE);
        if TA1 <> nil then
        begin
             if TA1.GetValue ('G_VENTILABLE') = 'X' then
             begin
                  OKVentil := TRUE;
                  for NumAxe:=1 to 5 do VentAttent[NumAxe]:=(TA1.GetValue ('G_VENTILABLE'+inttostr(NumAxe))='X') ;
             end;
        end
        else
        begin
             //TA2 := TGenInsert.FindFirst(['G_GENERAL'], [Compte], FALSE);
             TA2 := RechercheOptimise(TGenInsert, 'G_GENERAL', Compte, TRUE, fHGeneInsert);
             if TA2 <> nil then
             begin
                  if TA2.GetValue ('G_VENTILABLE') = 'X' then
                  begin
                       OKVentil := TRUE;
                       for NumAxe:=1 to 5 do VentAttent[NumAxe]:=(TA2.GetValue ('G_VENTILABLE'+inttostr(NumAxe))='X') ;
                  end;
             end
             else
             begin
                  //TA3 := TGen.FindFirst(['G_GENERAL'], [Compte], FALSE);
                  TA3 := RechercheOptimise(TGen, 'G_GENERAL', Compte, TRUE, fHGene);
                  if TA3 <> nil then
                  begin
                       if TA3.GetValue ('G_VENTILABLE') = 'X' then
                       begin
                            OKVentil := TRUE;
                            for NumAxe:=1 to 5 do VentAttent[NumAxe]:=(TA3.GetValue ('G_VENTILABLE'+inttostr(NumAxe))='X') ;
                       end;
                  end
                  else
                  begin
                          QDos := OpenSQL('SELECT * FROM GENERAUX WHERE G_GENERAL="'+Compte+'"', true);
                          if not QDos.EOF then
                          begin
                               if QDos.FindField ('G_VENTILABLE').asstring = 'X' then
                               begin
                                    OKVentil := TRUE;
                                    for NumAxe:=1 to 5 do VentAttent[NumAxe]:=(QDos.FindField ('G_VENTILABLE'+inttostr(NumAxe)).asstring='X') ;
                               end;
                               TGen.LoadDetailDB('GENERAUX', '', '', QDos, TRUE, FALSE);

                          end;
                          Ferme(QDos);
                  end;
             end;
        end;
     end;
     Result := OKVentil;
end;

procedure TImportCom.MetVentilable (Compte,Axe : string);
var
TA2           : TOB;
Numaxe        : string;
begin
     if Axe = '' then exit;
     Numaxe := Copy(Axe,2,1);
     if (Numaxe <> '1') and (Numaxe <> '2') and (Numaxe <> '3') and (Numaxe <> '4') and (Numaxe <> '5') then
     exit;
     if TGen <> nil then
     begin
             TA2 := RechercheOptimise(TGenInsert, 'G_GENERAL', Compte, TRUE, fHGeneInsert);
//             TA2 := TGenInsert.FindFirst(['G_GENERAL'], [Compte], FALSE);
             if TA2 <> nil then
             begin
                  TA2.PutValue ('G_VENTILABLE', 'X');
                  TA2.PutValue ('G_VENTILABLE'+Numaxe, 'X') ;
             end;
     end;
end;

procedure TImportCom.DecoupageEcriture;
var
i,ii,iii,ie                         : integer;
TLec,TLecPeriode,TLecpiece,TLec1    : TOB;
Numpiece                            : integer;
nbpiece                             : integer;
Soldedeb, SoldeCre                  : double;
etablissement, nature,ancnature     : string;
Jour,M,Y,Jouranc                    : Word;
Oketab                              : Boolean;
okgrp                               : Boolean;
devise,ancdevise                    : string;
begin
     if (OrigineFichier = 'S5') and (Modesaisie = 'LIB')  then exit;

     Oketab := FALSE;

     for i := 0 to TEcrParJournal.detail.Count - 1 do
     begin
          TLecPeriode := TEcrParJournal.detail[i];
          for ii := 0 to TLecPeriode.detail.Count - 1 do
          begin
               TLecpiece := TLecPeriode.detail[ii];
               nbpiece := TLecpiece.detail.Count;

               for iii := 0 to nbpiece - 1 do
               begin
                    TLec := TLecpiece.detail[iii];
                    // ajout pour renumérotation des lignes pour origine sisco
                    // ajout me 13-05-02 ajout cas journal
                    if (OrigineFichier = 'SI') and
                    ((NatureFichier = 'BAL') or (NatureFichier = 'JRL')) then
                    begin
                         numpiece :=  TLec.GetValue ('E_NUMEROPIECE');

                         DecoupageJournal(TLec,TLecpiece,TLecPeriode,numpiece);
                    end
                    else
                    begin
                         ie := 0;
                         okgrp := FALSE;
                         if Tlec.detail[ie].NomTable = 'ODA' then
                               MiseNumVentil (Tlec)
                         else
                         begin
                               if (Tlec.detail[ie].GetValue('E_MODESAISIE') = 'BOR') or
                               (Tlec.detail[ie].GetValue('E_MODESAISIE') = '-') then
                               okgrp := TRUE;

                               if okgrp  or (Tlec.detail[ie].GetValue('E_MODESAISIE') = 'LIB') then
                               begin
                                   DecodeDate(Tlec.detail[ie].GetValue ('E_DATECOMPTABLE'),Y,M,jouranc);
                                   etablissement := Tlec.detail[ie].GetValue ('E_ETABLISSEMENT');
                                   nature := Tlec.detail[ie].GetValue ('E_NATUREPIECE');
                                   ancnature :=  Tlec.detail[ie].GetValue ('E_NATUREPIECE');
                                   devise := Tlec.detail[ie].GetValue ('E_DEVISE');
                                   ancdevise := Tlec.detail[ie].GetValue ('E_DEVISE');
                                   Soldedeb := 0.0; SoldeCre := 0.0;
                                   Oketab := FALSE;
                                   ie := 0;
                                   TLec1 := TLec.Detail[0];
                                   while TLec1 <> nil  do
                                   begin
                                    //    if Tlec.detail[ie].GetValue ('E_DATECOMPTABLE')=idate1900 then continue;
                                       DecodeDate(Tlec.detail[ie].GetValue ('E_DATECOMPTABLE'),Y,M,jour);
                                       nature := Tlec.detail[ie].GetValue ('E_NATUREPIECE'); 
                                       devise := Tlec.detail[ie].GetValue ('E_DEVISE');
                                       if (jouranc = jour)
                                       and (devise = ancdevise) // Rajout du test sur la devise en plus de la nature et de la date comptable.
                                       and (nature = ancnature) then
                                       begin
                                             Soldedeb := Soldedeb + TLec.detail[ie].GetValue ('E_DEBIT');
                                             Soldecre := Soldecre + TLec.detail[ie].GetValue ('E_CREDIT');
                                       end
                                       else
                                       begin
                                            DecodeDate(Tlec.detail[ie].GetValue ('E_DATECOMPTABLE'),Y,M,jouranc);
                                            ancnature := Tlec.detail[ie].GetValue ('E_NATUREPIECE');
                                            ancdevise := Tlec.detail[ie].GetValue ('E_DEVISE');
                                            if (Arrondi(Soldedeb,V_PGI.OkDecV) <> Arrondi(Soldecre,V_PGI.OkDecV)) then
                                            begin
                                                if okgrp then
                                                     GenereCompteDattente (TLec, ie, Soldedeb, SoldeCre);
                                            end;
                                            Soldedeb := Soldedeb + TLec.detail[ie].GetValue ('E_DEBIT');
                                            Soldecre := Soldecre + TLec.detail[ie].GetValue ('E_CREDIT');
                                       end;
                                       if etablissement <> Tlec.detail[ie].GetValue ('E_ETABLISSEMENT') then
                                       begin
                                             etablissement := Tlec.detail[ie].GetValue ('E_ETABLISSEMENT');
                                             Oketab := TRUE;
                                       end;
                                       if ie < TLec.Detail.Count-1 then  TLec1 := Tlec.detail[ie+1]
                                       else TLec1 := nil;
                                       inc (ie);
                                   end;
                                   //AJOUT ME 16-03-2005
                                   if (Tlec.detail[TLec.Detail.Count-1].GetValue( 'E_DEVISE' )= V_PGI.DevisePivot) and
                                    (Arrondi(Soldedeb,V_PGI.OkDecV) <> Arrondi(Soldecre,V_PGI.OkDecV))  then
                                            GenereCompteDattente (TLec, ie, Soldedeb, SoldeCre);
                                   if Oketab then // si modification d'étabissement
                                   begin
                                        OnAfficheListeCom( FormatDateTime (Traduitdateformat('mmmm yyyy'),Tlec.detail[0].GetValue ('E_DATECOMPTABLE'))
                                        + ' Jrl :' + Tlec.detail[0].GetValue ('E_JOURNAL')
                                        + ' Pièce :' + IntToStr( Tlec.detail[0].GetValue ('E_NUMEROPIECE') )
                                        + ' établissement modifié' ,LISTEIMPORT);
                                   end;
                               end;
                               MiseajourDetail (Tlec,FALSE,Oketab,okgrp );
                         end; // non ODA
                    end;
               end;
          end;
       end;
end;

// fiche 10367 refonte de la fonction   reprendre 6.5
procedure TImportCom.AttenteAnaTOB ( TOBEcr : TOB ; NbDec : integer );
Var Pourc,XP,XD,TotP,TotD,TotPourc,TotalEcriture,TotalDevise : double ;
    i             : integer ;
    TOBAxe,TOBAna : TOB ;
    da,dv         : double;
    nn            : integer;
    fb           : TFichierBase ;
    NumAxe       : integer;
BEGIN
          for NumAxe:=1 to 5 do // fin fiche 10095
          begin
            if not VentAttent[NumAxe] then
            begin
                 TOBAxe := TOB.Create('A'+inttostr(NumAxe), TOBEcr, -1);
                 TOBAxe.AddChampSupValeur('Y_AXE', 'A'+inttostr(NumAxe));
            end;
          end;
          fb:=AxeToFb('A1');
          TotalEcriture:=TOBEcr.GetValue('E_DEBIT')+TOBEcr.GetValue('E_CREDIT') ;
          TotalDevise:=TOBEcr.GetValue('E_DEBITDEV')+TOBEcr.GetValue('E_CREDITDEV') ;
          // Prendre les axes ventilés
          for nn:=0 to TOBEcr.Detail.Count-1 do
          begin
            TotP:=0 ; TotD:=0 ; TotPourc:=0 ;
            TOBAxe:=TOBEcr.detail[nn];
            // ajout me car ECRCOMPL est au même niveau que axe
            if TOBAXE.FieldExists ('EC_CLEECR') then continue;
            for i:=0 to TOBAxe.Detail.Count-1 do
            BEGIN
              // Sommations
              TOBAna:=TOBAxe.Detail[i] ;
              Pourc:=TOBAna.GetValue('Y_POURCENTAGE') ;
              fb:=AxeToFb(TOBAna.GetValue('Y_AXE')) ;
              TotPourc:=Arrondi(TotPourc+Pourc,ADecimP) ;
              if ((i<TOBAxe.Detail.Count-1) or (Arrondi(TotPourc-100.0,ADecimP)<>0)) then
                 BEGIN
                 // Si pas dernière ligne ou tot % <> 100 % alors cumuler
                 XP:=TOBAna.GetValue('Y_DEBIT')+TOBAna.GetValue('Y_CREDIT') ;
                 XD:=TOBAna.GetValue('Y_DEBITDEV')+TOBAna.GetValue('Y_CREDITDEV') ;
                 TotP:=Arrondi(TotP+XP,V_PGI.OkDecV) ; TotD:=Arrondi(TotD+XD,NbDec) ;
                 END else
                 BEGIN
                 // Si Dernière ligne et tot % = 100 % alors répercuter systématiquement le reste sur la dernière ligne
                 XP:=Arrondi(TotalEcriture-TotP,V_PGI.OkDecV) ;
                 XD:=Arrondi(TotalDevise-TotD,NbDec) ;
                 if TOBEcr.GetValue('E_DEBIT')<>0 then
                    BEGIN
                    TOBAna.PutValue('Y_DEBIT',XP) ; TOBAna.PutValue('Y_DEBITDEV',XD) ;
                    END else
                    BEGIN
                    TOBAna.PutValue('Y_CREDIT',XP) ; TOBAna.PutValue('Y_CREDITDEV',XD) ;
                    END ;
                 END ;
            END ;

           if (TOBAxe.Detail.Count <> 0) and (TotalEcriture <> TotP) then  // ajout me car ECRCOMPL est au même niveau que axe
           begin
               // Si ventilation incomplète, ventiler le reste sur une nouvelle ligne avec section d'attente
                // fiche 10594 si plusieurs ventilations analytiques avec écart
                // équilibre de l'écart euro sur la dernière ligne en devise
               if (TOBAxe.Detail[TOBAxe.Detail.Count-1].GetValue('Y_DEVISE') = V_PGI.DevisePivot) and (Arrondi(TotPourc-100,ADecimP)<>0) and ((Arrondi(TotPourc-100,ADecimP) > 1) or
                (Arrondi(TotPourc-100,ADecimP) < -1)) then
               begin
                    TOBAna:=TOB.Create('ANALYTIQ',TOBAxe,-1) ;
                    TobAna.InitValeurs ;
                    EcrVersAna(TOBEcr,TOBAna) ;
                    TOBAna.PutValue('Y_SECTION', GetInfoCpta(fb).Attente) ;
                    TOBAna.PutValue('Y_NUMVENTIL',TOBAxe.Detail.Count) ;
                    if TotPourc <> 0 then
                       TOBAna.PutValue('Y_POURCENTAGE',Arrondi(100.0-TotPourc,ADecimP))
                    else
                       TOBAna.PutValue('Y_POURCENTAGE', 0);
                    XP:=Arrondi(TotalEcriture-TotP,V_PGI.OkDecV) ;
                    XD:=Arrondi(TotalDevise-TotD,NbDec) ;
                    if TOBEcr.GetValue('E_DEBIT')<>0 then
                       BEGIN
                        TOBAna.PutValue('Y_DEBIT',XP) ; TOBAna.PutValue('Y_DEBITDEV',XD) ;
                        if BSoldeLigne then
                         Ajoute (lpS, Tobana.GetValue ('Y_AXE'), Tobana.GetValue ('Y_SECTION'), XP, 0, -1, -1,Tobana.GetValue ('Y_DATECOMPTABLE')) ;
                       END else
                       BEGIN
                        TOBAna.PutValue('Y_CREDIT',XP) ; TOBAna.PutValue('Y_CREDITDEV',XD) ;
                        if BSoldeLigne then
                         Ajoute (lpS, Tobana.GetValue ('Y_AXE'), Tobana.GetValue ('Y_SECTION'), 0, XP, -1, -1,Tobana.GetValue ('Y_DATECOMPTABLE')) ;
                       END ;
               end
               else
               if (Arrondi(TotPourc-100,ADecimP)<>0) and (Arrondi(TotPourc-100,ADecimP) <= 1) then
                  BEGIN
                        TOBAna:=TOBAxe.Detail[TOBAxe.Detail.Count-1] ;
                        if TOBAna.GetValue('Y_DEBIT') <> 0.0 then
                           da := TOBAna.GetValue('Y_DEBIT')
                        else
                           da := TOBAna.GetValue('Y_CREDIT');
                        if TOBAna.GetValue('Y_DEBITDEV') <> 0.0 then
                           dv := TOBAna.GetValue('Y_DEBITDEV')
                        else
                           dv := TOBAna.GetValue('Y_CREDITDEV');

                        XP:=Arrondi(TotalEcriture-TotP,V_PGI.OkDecV);
                        XD:=Arrondi(TotalDevise-TotD,NbDec) ;
                        if TOBEcr.GetValue('E_DEBIT')<>0 then
                           BEGIN
                            TOBAna.PutValue('Y_DEBIT', XP+da) ; TOBAna.PutValue('Y_DEBITDEV',XD+dv) ;
                            if BSoldeLigne then
                             Ajoute (lpS, Tobana.GetValue ('Y_AXE'), Tobana.GetValue ('Y_SECTION'), XP, 0, -1, -1,Tobana.GetValue ('Y_DATECOMPTABLE')) ;
                           END else
                           BEGIN
                            TOBAna.PutValue('Y_CREDIT', XP+da) ; TOBAna.PutValue('Y_CREDITDEV',XD+dv) ;
                            if BSoldeLigne then
                            Ajoute (lpS, Tobana.GetValue ('Y_AXE'), Tobana.GetValue ('Y_SECTION'), 0, XP, -1, -1,Tobana.GetValue ('Y_DATECOMPTABLE')) ;
                           END ;
                  END ;
           end;
          end; //for
END ;

///////////////////////// MAJ des soldes
procedure TImportCom.CalcSoldeList (var LJB : TList; fb : TFichierBase);
var
T         : T_SC ;
FRM       : TFRM ;
i         : integer;
lTypeExo  : TTypeExo;
Exerc     : TExoDate;
begin
    if not BSoldeLigne then exit;
    if LJB = nil then  exit;
    for i := 0 to LJB.Count - 1 do
    begin
       T:=T_SC(LJB[i]) ;
       Fillchar(FRM,SizeOf(FRM),#0) ;
       FRM.cpt:=T.Cpte ;
       FRM.axe:=T.Identi ;
       FRM.NumD:=T.NumP ;
       FRM.DateD:=T.DateP ;
       FRM.LigD:=T.NumL ;
       CQuelExercice (FRM.DateD, Exerc);
       If Exerc.code = GetPrecedent.Code Then lTypeExo := tePrecedent Else
          If Exerc.code = GetEnCours.Code Then lTypeExo := teEncours Else
             If Exerc.code = GetSuivant.Code Then lTypeExo := teSuivant Else
                lTypeExo := teEncours;

       AttribParamsNew ( FRM , T.Debit, T.Credit , lTypeExo) ;
       if ExecReqMAJ (fb, FALSE, FALSE, FRM) <>1 then
        begin
         V_PGI.IoError:=oeSaisie ;  break;
        end;
    end;
end;


procedure  TImportCom.MiseajourDetail (var Tlec : TOB; okc,oke,okgrp : boolean);
var
ie,Ligne,numgrp,ancnumgrp                          : integer;
docid                                              : string;
DOCGUID                                            : Boolean;
Jour,M,Y,ancjour                                   : Word;
ia,iana                                            : integer;
Tobaxe,Tobana                                      : TOB;
pourcent                                           : double;
deb,cre,totana                                     : double;
TotalEcriture, TotalEcritureDev                    : double;
Totdeb,Totcre,Totpourcent                          : double;
per,nump,nume,Maxligne,Newligne                    : integer;
TA,T                                               : TOB;
Q1                                                 : TQuery;
Debiteur, OkChangement                             : Boolean;
begin
   Ligne := 1;
   numgrp := 1; ancnumgrp := 0;
   Totdeb := 0.0; Totcre := 0.0;
   MaxLigne := 0; TA := nil;
   if okc and ((OrigineFichier = 'S1') or (NatureFichier = 'SYN'))
   and   (StrToint(VersionFichier) < 3) then   // ajout me pour S1
   begin
   // création d'une liste pour les pièces dans le désordre
   // en sauvegardant le dernier numéro de ligne
                 if TOBPIECE = nil then
                    TOBPIECE := TOB.Create('', nil, -1);
                 per := Tlec.detail[TLec.Detail.Count-1].Getvalue('E_PERIODE');
                 nump := Tlec.detail[TLec.Detail.Count-1].Getvalue('E_NUMEROPIECE');
                 nume := Tlec.detail[TLec.Detail.Count-1].Getvalue('E_NUMECHE');
                 TA := TOBPIECE.FindFirst(['JOURNAL','PERIODE','NUMEROPIECE'],
                        [Tlec.detail[TLec.Detail.Count-1].Getvalue('E_JOURNAL'),
                        inttostr(per),IntToStr(nump)], FALSE);
                 if TA = nil then
                 begin
                      TA := TOB.Create ('',TOBPIECE,-1);
                      TA.AddChampSupValeur('JOURNAL', Tlec.detail[TLec.Detail.Count-1].Getvalue('E_JOURNAL'));
                      TA.AddChampSupValeur('PERIODE', IntToStr(per));
                      TA.AddChampSupValeur('NUMEROPIECE', nump);
                       TA.AddChampSupValeur('NUMECHE', nume);
                      TA.AddChampSupValeur('MAXLIGNE', Tlec.detail.Count);
                      Maxligne := 0;
                 end
                 else
                 begin
                      Maxligne := TA.Getvalue ('MAXLIGNE');
                      TA.Putvalue ('MAXLIGNE',Maxligne+Tlec.detail.Count);
                 end;
   end;
   docid := ''; OkChangement := FALSE;
   for ie:=0 to TLec.Detail.Count-1  do
   begin
                 if okc then
                 begin
                 TotalMouvDebit  := TotalMouvDebit + Tlec.detail[ie].GetValue ('E_DEBIT');
                 TotalMouvCredit := TotalMouvCredit + Tlec.detail[ie].GetValue ('E_CREDIT');
                 // ajout me pour multi echeance
                      if Tlec.detail[ie].getValue ('E_NUMECHE') > 1 then
                       begin
                           T := Tlec.FindFirst(['E_JOURNAL','E_EXERCICE', 'E_DATECOMPTABLE','E_NUMEROPIECE','E_GENERAL','E_NUMLIGNE'], [Tlec.detail[ie].getValue ('E_JOURNAL'),
                             Tlec.detail[ie].getValue ('E_EXERCICE'), Tlec.detail[ie].getValue ('E_DATECOMPTABLE'),  Tlec.detail[ie].getValue ('E_NUMEROPIECE'), Tlec.detail[ie].getValue ('E_GENERAL'), ligne-1], FALSE);
                           if T <> nil then
                           begin
                              Tlec.detail[ie].PutValue ('E_NUMLIGNE', T.getValue ('E_NUMLIGNE'));
                              if ligne > 1 then dec (ligne);
                           end
                           else
                              Tlec.detail[ie].PutValue ('E_NUMLIGNE', Maxligne+ligne);
                       end
                       else
                       Tlec.detail[ie].PutValue ('E_NUMLIGNE', Maxligne+ligne);

                       if Tlec.detail[ie].detail.count > 0 then
                       if Tlec.detail[ie].detail[0].FieldExists ('EC_CLEECR') then
                          Tlec.detail[ie].detail[0].PutValue ('EC_NUMLIGNE', Tlec.detail[ie].Getvalue ('E_NUMLIGNE'));

                       deb := Tlec.detail[ie].GetValue ('E_DEBIT');
                       cre := Tlec.detail[ie].GetValue ('E_CREDIT');
                       if deb > cre then TotalEcriture := deb - cre
                       else TotalEcriture := cre-deb;

                       // ajout me 1-02-2005 si sens des montants différents %  prend devient < 0
                       Debiteur := (deb > cre);

                       deb := Tlec.detail[ie].GetValue ('E_DEBITDEV');
                       cre := Tlec.detail[ie].GetValue ('E_CREDITDEV');
                       if deb > cre then TotalEcritureDev := deb - cre
                       else TotalEcritureDev := cre-deb;

///////////////////////// MAJ des soldes
                       if BSoldeLigne then
                       begin
                             if Tlec.detail[ie].GetValue ('E_AUXILIAIRE') <> '' then
                                Ajoute ( lpTi,'T', Tlec.detail[ie].GetValue ('E_AUXILIAIRE'), Tlec.detail[ie].GetValue ('E_DEBIT'), Tlec.detail[ie].GetValue ('E_CREDIT'),
                                Tlec.detail[ie].GetValue ('E_NUMEROPIECE'), Tlec.detail[ie].GetValue ('E_NUMLIGNE'), Tlec.detail[ie].GetValue ('E_DATECOMPTABLE')) ;

                                Ajoute ( lpG,'G', Tlec.detail[ie].GetValue ('E_GENERAL'), Tlec.detail[ie].GetValue ('E_DEBIT'), Tlec.detail[ie].GetValue ('E_CREDIT'),
                                Tlec.detail[ie].GetValue ('E_NUMEROPIECE'), Tlec.detail[ie].GetValue ('E_NUMLIGNE'), Tlec.detail[ie].GetValue ('E_DATECOMPTABLE')) ;

                                Ajoute (lpJ, 'J', Tlec.detail[ie].GetValue ('E_JOURNAL'), Tlec.detail[ie].GetValue ('E_DEBIT'), Tlec.detail[ie].GetValue ('E_CREDIT'),
                                Tlec.detail[ie].GetValue ('E_NUMEROPIECE'), Tlec.detail[ie].GetValue ('E_NUMLIGNE'), Tlec.detail[ie].GetValue ('E_DATECOMPTABLE')) ;
                       end;

                       if MajDate then // fiche 10160
                          AjoutListeDatedernMvt (lpGM, Tlec.detail[ie].GetValue ('E_GENERAL'),  Tlec.detail[ie].GetValue ('E_DATECOMPTABLE'));

                       Tobaxe := Tlec.detail[ie];
                       // ajout me car ECRCOMPL est au même niveau que axe
                       if TOBAXE.FieldExists ('EC_CLEECR') then continue;

                       for ia:=0 to Tobaxe.Detail.Count-1  do
                       begin
                            if GetParamSocSecur('SO_CROISAXE', False) then break;
                            Tobana := Tobaxe.detail[ia];

                            Totpourcent := 0;
                            for iana:=0 to Tobana.Detail.Count-1  do
                            begin
                                 if okc then
                                 begin
                                 Tobana.detail[iana].PutValue ('Y_NUMLIGNE', Maxligne+ligne);
                                 if Tobana.detail[iana].GetValue('Y_TYPEANALYTIQUE')= 'X' then
                                    Tobana.detail[iana].PutValue ('Y_NUMLIGNE', 0);
                                 end;

                                 if BEcartChange then
                                 begin
                                    Tobana.detail[iana].PutValue ('Y_JOURNAL', Tlec.detail[ie].GetValue ('E_JOURNAL'));
                                    Tobana.detail[iana].PutValue ('Y_NUMEROPIECE', Tlec.detail[ie].GetValue ('E_NUMEROPIECE'));
                                 end;
                                 // mise à jour des pourcentage, le total est fait avant dans EcrVersAna
                                                          // ajout me le test à zéro si le % est déjà calculé
                                                          // par les ventilations types
                                   if (TotalEcriture > 0) and (Tobana.detail[iana].GetValue ('Y_POURCENTAGE') = 0) then
                                   begin
                                        deb := Tobana.detail[iana].GetValue ('Y_DEBIT');
                                        cre := Tobana.detail[iana].GetValue ('Y_CREDIT');

///////////////////////// MAJ des soldes
                                        if BSoldeLigne then
                                        begin
                                          Ajoute (lpS, Tobana.detail[iana].GetValue ('Y_AXE'), Tobana.detail[iana].GetValue ('Y_SECTION'), deb, cre,
                                          -1, -1,Tobana.detail[iana].GetValue ('Y_DATECOMPTABLE')) ;
                                        end;

                                        if deb > cre then totana := deb-cre
                                        else totana := cre-deb;
                                        pourcent := (totana/TotalEcriture)*100;
                                        pourcent:=Arrondi(pourcent,ADecimP) ;
                                        Totpourcent := Totpourcent + pourcent;
                                        // dans le cas de 99.99999
                                        if (iana = Tobana.Detail.Count-1) and (Totpourcent <> 100) then
                                        begin
                                             if pourcent = 0 then pourcent := 100-Totpourcent;
                                        end;

                                        // ajout me 1-02-2005 Fiche 10138 si sens des montants différents %  prend devient < 0
                                        if (Debiteur and (cre > deb)) or ((not Debiteur) and (deb > cre)) then
                                            Tobana.detail[iana].PutValue ('Y_POURCENTAGE', pourcent * (-1))
                                        else
                                            Tobana.detail[iana].PutValue ('Y_POURCENTAGE', pourcent);
                                        Tobana.detail[iana].PutValue ('Y_TOTALECRITURE', TotalEcriture);
                                   end;
                                   if TotalEcritureDev > 0 then
                                        Tobana.detail[iana].PutValue ('Y_TOTALDEVISE', TotalEcritureDev);
                            end;
                       end;
                 end;
            if (GetParamSocSecur('SO_ETABLISDEFAUT', '001') <> '') and oke then
               Tlec.detail[ie].PutValue ('E_ETABLISSEMENT', GetParamSocSecur('SO_ETABLISDEFAUT', '001'));
            if okgrp then
            begin
                 DOCGUID := FALSE;
                 if Tlec.detail[ie].detail.count > 0 then
                 begin
                      TA :=  Tlec.detail[ie].detail[0];
                       if (TA.FieldExists ('EC_DOCGUID')) then
                            DOCGUID := TRUE;
                 end;
                 DecodeDate(Tlec.detail[ie].GetValue ('E_DATECOMPTABLE'),Y,M,jour);
                 Totdeb := Totdeb + Tlec.detail[ie].GetValue ('E_DEBIT');
                 Totcre := Totcre + Tlec.detail[ie].GetValue ('E_CREDIT');
                 if (jour <> ancjour)  then
                 begin
                     DecodeDate(Tlec.detail[ie].GetValue ('E_DATECOMPTABLE'),Y,M,ancjour);
                     ancnumgrp := Numgrp;
                     if DOCGUID then
                        docid := TA.Getvalue ('EC_DOCGUID');
                 end
                 else
                 begin
                      if Arrondi(Totdeb,V_PGI.OkDecV) = Arrondi(Totcre,V_PGI.OkDecV) then
                      begin
                           ancnumgrp := Numgrp;
                           inc (Numgrp);
                           Totdeb := 0.0; Totcre := 0.0;
                      end;
                 end;
                 Tlec.detail[ie].PutValue ('E_NUMGROUPEECR', ancnumgrp);
                 if DOCGUID then
                 begin
                      if not OkChangement then
                      begin
                         TA.PutValue ('EC_DOCGUID', docid);
                          // fiche 32913  en mode bordereau il y a un seul ecrcompl par numgroupecr
                         if ie <> 0 then
                            Tlec.detail[ie].cleardetail
                         else
                         TA.PutValue ('EC_CLEECR', IntTostr(ancnumgrp));
                      end
                      else
                      begin
                         docid := TA.Getvalue ('EC_DOCGUID');
                         OkChangement := FALSE;
                         TA.PutValue ('EC_CLEECR', IntTostr(ancnumgrp));
                      end;
                 end;
                 if ancnumgrp <> Numgrp then OkChangement := TRUE;
                 ancnumgrp := Numgrp;
            end;
            inc(Ligne);
   end;
end;

Function TImportCom.NombreExercice(WW : string; Max : Boolean=FALSE): integer;
var
Q1       : TQuery;
Nb       : integer;
begin
           if Max then
              Q1 := OpenSQl ('SELECT Max(EX_EXERCICE) FROM EXERCICE '+ WW,TRUE)
           else
              Q1 := OpenSQl ('SELECT Count(EX_EXERCICE) FROM EXERCICE '+ WW,TRUE);
           Nb := Q1.Fields[0].asinteger;
           ferme (Q1);
           Result := Nb;
end;

Function TImportCom.RenExoClo (Where : string) : Boolean;
begin
  Result := FALSE;
  if stArg = '' then
  begin
      if (OrigineFichier = 'SI') and ((NatureFichier = 'BAL')or (NatureFichier='EXE')) then
      begin
          if ExisteSQL ('SELECT EX_DATEDEBUT,EX_DATEFIN,EX_ETATCPTA from EXERCICE '+Where) then
          begin
                  if (stArg = '') then
                      AffMessage ('L''exercice est clos. L''importation des écritures est impossible.')
                  else
                      OnAfficheListeCom(ERR_EXERCICECLOS, LISTEIMPORT);
                  OKexercice := FALSE;
                  exit;
          end;
      end;
  end;
  Result := TRUE;
end;

Procedure TImportCom.ChargeTobTiersGene (sSQl : string; Quoi : Char);
var
QDos      : TQuery;
T         : Tob;
begin
     Case  Quoi of
           'G' :
           begin
                   QDos := OpenSQL(sSQL, true, -1, '', TRUE); // Fiche 10589 Pb lenteur TRT  A PORTER sur 7XX
                   try
                      while not QDos.Eof do
                      begin
                        T := TOB.Create ('GENERAUX',TGen,-1);
                        T.SelectDB('', QDos);
                        fHGene.Put( T.GetString('G_GENERAL'), TObject(T));
                        QDos.Next;
                      end;
                    finally
                      Ferme (QDos);
                    end;
           end;
           'T' :
           begin
                     QDos := OpenSQL(sSQL, true, -1, '', TRUE);  // Pb lenteur TRT  A PORTER sur 7XX
                     try
                        while not QDos.Eof do
                        begin
                          T := TOB.Create ('TIERS',TTiers,-1);
                          T.SelectDB('', QDos);
                          fHTiers.Put( T.GetString('T_AUXILIAIRE'), TObject(T));
                          QDos.Next;
                        end;
                      finally
                        Ferme (QDos);
                      end;
           end;
           'U' :  // Tiers CCIMPEX
           begin
                      QDos := OpenSQL(sSQL, true);
                      TTiersCAU.LoadDetailDB('TIERS', '', '', QDos, TRUE, FALSE);
                      Ferme(QDos);
           end;
     end;
end;

procedure TImportCom.Insert_journal;
var ii    : integer;
Q         : TQuery;
W1,W2     : string;
typepiece : string;
TSou      : TOB;
begin
     for ii:=0 to SizeRecControl do
     begin
       if (RecControl[ii].champ  = 'JAL') then
       begin
         if RecControl[ii].valeur <> '' then
         begin
            if not ExisteSQL('SELECT * FROM JOURNAL Where J_JOURNAL="'+RecControl[ii].valeur+ '"') then
            begin
                 if GetParamSocSecur (RecControl[ii].NomChamp, '') = '' then
                 begin
                      Q := OpenSQL ('SELECT PRR_NUMPLAN,PRR_SOCNOM,PRR_SOCDATA from PARSOCREF where ' +
                        ' PRR_NUMPLAN =7 and PRR_SOCNOM="'+RecControl[ii].NomChamp+'"', True);
                        if not Q.EOF then
                        begin
                                      SetParamsoc (RecControl[ii].NomChamp, Q.Findfield ('PRR_SOCDATA').asstring);
                                      W1 := ' AND JR_JOURNAL="'+ Q.Findfield ('PRR_SOCDATA').asstring+'"';
                                      W2 := ' Where J_JOURNAL="'+Q.Findfield ('PRR_SOCDATA').asstring+'"';
                                      ferme (Q);
                                      LoadStandardCompta (7, 'JOURNAL','JALREF', W2, W1);
                        end
                        else
                        Ferme(Q);
                 end
                 else
                 begin
                      if (RecControl[ii].Nature = 'CLO') or (RecControl[ii].Nature = 'ANO')
                      or (RecControl[ii].Nature = 'REG') then
                         typepiece := '-'
                      else
                         typepiece := 'LIB';

                    ExecuteSql('INSERT INTO JOURNAL (J_JOURNAL,J_LIBELLE,J_NATUREJAL,J_COMPTEURNORMAL,'+
                    'J_COMPTEURSIMUL,J_CONTREPARTIE,J_AXE,J_MODESAISIE,J_COMPTEAUTOMAT,J_COMPTEINTERDIT,J_ABREGE,J_FERME,' +
                    'J_VALIDEEN,J_VALIDEEN1,J_DATEDERNMVT,J_NUMDERNMVT, J_CENTRALISABLE,J_MULTIDEVISE,J_CREERPAR) ' +
                    'VALUES ("' + RecControl[ii].valeur + '","Journal '+RecControl[ii].Nature +
                    '","' + RecControl[ii].Nature + '","'+RecControl[ii].valeur+'","SIM","","","'+
                    typepiece+'","","","Journal '+RecControl[ii].Nature+'","-' +
                    '","------------------------","------------------------' +
                    '","'+stDate1900+'",0,"-","X","COM"' +
                    ')');
                   if (not ExisteSQl ('SELECT SH_SOUCHE FROM SOUCHE Where SH_TYPE="CPT" AND SH_SOUCHE="'+ RecControl[ii].valeur + '"')) then
                   begin
                         TSou := TOB.Create('SOUCHE', nil, -1);
                         TSou.Putvalue ('SH_TYPE', 'CPT');
                         TSou.Putvalue ('SH_SOUCHE', RecControl[ii].valeur);
                         TSou.Putvalue ('SH_LIBELLE', 'SOUCHE :'+ RecControl[ii].valeur);
                         TSou.Putvalue ('SH_NUMDEPART', 1);
                         TSou.Putvalue ('SH_SIMULATION', '-');
                         TSou.Putvalue ('SH_ANALYTIQUE', '-');
                         TSou.InsertDB(nil, TRUE);
                         TSou.Free;

                   end;
                 end;
            end
         end;
       end;
     end;
end;

procedure TImportCom.VerifAxeSection;
var
i              : integer;
Q1             : Tquery;
axe            : string;
LeFb           : TFichierBase;
Sectionatt     : string;
TSec           : TOB;
Satt           : string;
begin
 if not OkVerAxe then
 begin
       LeFb := FbAxe1;
       for i:=1 to MAXAXE do
       begin
            if i = 1 then LeFb :=fbAxe1;
            if i = 2 then LeFb :=fbAxe2;
            if i = 3 then LeFb :=fbAxe3;
            if i = 4 then LeFb :=fbAxe4;
            if i = 5 then LeFb :=fbAxe5;
            axe := 'A'+IntToStr(i);
           Q1 := OpenSql ('SELECT * from axe where X_AXE="'+axe+'"', True);
           if  not Q1.EOF then
           begin
                if Q1.FindField ('X_SECTIONATTENTE').asstring <> '' then
                begin
                     Sectionatt := Q1.FindField ('X_SECTIONATTENTE').asstring;
                     Satt := BourreOuTronque(Q1.FindField ('X_SECTIONATTENTE').asstring, LeFb)
                end
                else
                begin
                     Satt :='ATTENTE'+Q1.findField ('X_AXE').asstring;
                     Satt := BourreOuTronque(Satt, LeFb);
                     Sectionatt := '';
                end;
           end;
           ferme (Q1);
           if Satt <> '' then
           begin
                TSec := TOB.Create('SECTION', nil, -1);
                TSec.Putvalue ('S_AXE', axe);
                TSec.Putvalue ('S_SECTION', Satt);
                TSec.Putvalue ('S_LIBELLE', 'Section d''attente ');
                TSec.PutValue ('S_CONFIDENTIEL', '0');
                TSec.InsertOrUpdateDB(TRUE);
                TSec.Free;
                if length(Satt) <> length (Sectionatt)then
                begin
                         ExecuteSQL('UPDATE AXE SET X_AXE="' + axe +'",X_SECTIONATTENTE="' + Satt +
                         '" Where X_AXE="'+axe+'"');
                end;
           end;
       end;
       OkVerAxe := TRUE;
 end;
end;

procedure TImportCom.AffectationMontantdoublon(St: string; var debit,credit : double);
var
  CodeMontant, Sens: string;
  IndD, IndE       : Integer;
  Montant          : array[1..3] of Double;
  SMontant         : array[1..3] of string;
begin
  debit := 0.0; credit := 0.0;
  CodeMontant := EnleveBlanc(St, 173, 3);
  IndD := Pos('D', CodeMontant);
  IndE := Pos('E', CodeMontant);
  Fillchar(Montant, SizeOf(Montant), #0);
  SMontant[1] := StPoint(EnleveBlanc(St, 131, 20));
  SMontant[2] := StPoint(EnleveBlanc(St, 176, 20));
  SMontant[3] := StPoint(EnleveBlanc(St, 196, 20));
  if SMontant[1] <> '' then
    Montant[1] := Valeur(SMontant[1]);
  if SMontant[2] <> '' then
    Montant[2] := Valeur(SMontant[2]);
  if SMontant[3] <> '' then
    Montant[3] := Valeur(SMontant[3]);

  Sens := EnleveBlanc(St, 130, 1);
  if IndE > 0 then // si montant en euro
  begin
      if Sens = 'D' then debit := Montant[IndE]
      else if Sens = 'C' then  credit := Montant[IndE];
  end;
  if IndD > 0 then   // si montant en devise
  begin
    if Sens = 'D' then
         debit := Montant[IndD]
    else if Sens = 'C' then
         credit := Montant[IndD];
  end;
end;

Function TImportCom.ReChRibPrincipal(TA : TOB; Nature : string) : String ;
BEGIN
    Result := '';
    if not TA.FieldExists ('R_PAYS') then exit;
    if AttribRIBAuto then
    begin
          If (Nature='CLI') or  (Nature='FOU') or  (Nature='SAL')  then
          begin
              if (UpperCase(Copy(TA.GetValue('R_PAYS'), 1, 2)) = 'FR') or (TA.GetValue('R_PAYS')='') then
                Result := EncodeLeRib(TA.GetValue('R_ETABBQ'),TA.GetValue('R_GUICHET'),
                                      TA.GetValue('R_NUMEROCOMPTE'),TA.GetValue('R_CLERIB'),
                                      TA.GetValue('R_DOMICILIATION'))
              else  // fiche 10559
                if (Copy(TA.GetValue('R_PAYS'), 1, 2) <> '') and (TA.GetValue('R_CODEIBAN') <> '')  then
                    Result := '*' + TA.GetValue('R_CODEIBAN');
          end;
    end;
END ;


function TImportCom.ExisteCompte (Ext : string; var Tftr : TFTransfertInfo; ForceLettrage : Boolean=TRUE) : Boolean;
var
TA1,TA2,TA3    : TOB;
OKCompte       : Boolean;
lettrable      : string;
v1,v2,v3,v4,v5 : string;
sSQL           : string;
QD             : TQuery;
T              : TOB;
Fermeture      : string;
begin

OKCompte       := FALSE;
Tftr.OKLettrage     := FALSE;
Fermeture      := '-';
With Tftr do
begin
// ajout V590
iCompte := UpperCase(iCompte);
OKLettrage     := FALSE;

// on regarde d'abord au niveau de la liste cache
// si le compte existe

if Ext = 'T_' then
begin
     if TTiers <> nil then
     begin
        TA1 := TTiersUpdate.FindFirst([Ext+'AUXILIAIRE'], [iCompte], FALSE);
        if TA1 <> nil then
        begin
             OKCompte  := TRUE;
             NatureCpte    := TA1.Getvaleur (FiTOBT_NATUREAUXI);
             lettrable := TA1.GetValeur (FiTOBT_LETTRABLE);
             if ForceLettrage and (lettrable = 'X') then OKLettrage := TRUE;
             Regimetva := TA1.GetValeur (FiTOBT_REGIMETVA);
             Collectif :=  TA1.GetValue ('T_COLLECTIF');
             TPayeur   := TA1.GetValue ('T_PAYEUR');
             RibAuto       := ReChRibPrincipal(TA1, NatureCpte);
             Fermeture := TA1.GetValue ('T_FERME');
        end
        else
        begin
//             TA2 := TTiersInsert.FindFirst([Ext+'AUXILIAIRE'], [Compte], FALSE);
             TA2 := RechercheOptimise(TTiersInsert, Ext+'AUXILIAIRE', iCompte, FALSE, fHTiersInsert);
             if TA2 <> nil then begin
                OKCompte := TRUE;
                NatureCpte := TA2.Getvaleur (FiTOBT_NATUREAUXI);
                lettrable := TA2.GetValeur (FiTOBT_LETTRABLE);
                if ForceLettrage and (lettrable = 'X') then OKLettrage := TRUE;
                Regimetva := TA2.GetValeur (FiTOBT_REGIMETVA);
                Collectif :=  TA2.GetValue ('T_COLLECTIF');
                TPayeur   := TA2.GetValue ('T_PAYEUR');
                RibAuto       := ReChRibPrincipal(TA2, NatureCpte);
                Fermeture := TA2.GetValue ('T_FERME');
             end
             else
             begin
                  TA3 := RechercheOptimise(TTiers, Ext+'AUXILIAIRE', iCompte, FALSE, fHTiers);
                  if TA3 <> nil then
                  begin
                     OKCompte := TRUE;
                     NatureCpte := TA3.Getvaleur (FiTOBT_NATUREAUXI);
                     lettrable := TA3.GetValeur (FiTOBT_LETTRABLE);
                     if ForceLettrage and (lettrable = 'X') then OKLettrage := TRUE;
                     Regimetva := TA3.GetValeur (FiTOBT_REGIMETVA);
                     Collectif :=  TA3.GetValue ('T_COLLECTIF');
                     TPayeur   := TA3.GetValue ('T_PAYEUR');
                     RibAuto       := ReChRibPrincipal(TA3, NatureCpte);
                     Fermeture := TA3.GetValue ('T_FERME');
                  end
                  else   // ajout me 14-01-2005
                  begin
(* à porter sur V7XX  Pb dossier nb d'auxilaire > 2000 on ne trouver pas le mode de reglement
                       sSQL := 'SELECT T_AUXILIAIRE,T_NATUREAUXI,T_LETTRABLE,T_REGIMETVA,T_COLLECTIF, T_PAYEUR, T_FERME  '+
                                    ',T_CORRESP1, T_CORRESP2, T_ESCOMPTE, T_REMISE, T_FACTURE' +
                                    ',R_PAYS, R_ETABBQ, R_GUICHET,R_NUMEROCOMPTE,R_CLERIB,R_DOMICILIATION,R_CODEIBAN FROM TIERS '
                                 + ' left outer join rib on ( t_auxiliaire=r_auxiliaire AND R_PRINCIPAL="X") '
                                 + ' Where  T_AUXILIAIRE="'+Compte+ '"';
*)
                            sSQL := 'SELECT T_AUXILIAIRE,T_TIERS,T_LETTRABLE,T_NATUREAUXI,T_MULTIDEVISE,T_LIBELLE,T_COLLECTIF, '+
                            'T_EAN,T_TABLE0,T_TABLE1,T_TABLE2,T_TABLE3,T_TABLE4,T_TABLE5,T_TABLE6,T_TABLE7,T_TABLE8,T_TABLE9,'+
                            'T_ADRESSE1,T_ADRESSE2,T_ADRESSE3,T_CODEPOSTAL,T_VILLE,T_PAYS,T_ABREGE,T_LANGUE,T_TELEPHONE,T_FAX,'+
                            'T_REGIMETVA,T_MODEREGLE,T_COMMENTAIRE,T_NIF,T_SIRET,T_APE,T_JURIDIQUE,T_FORMEJURIDIQUE,T_TVAENCAISSEMENT,'+
                            'T_DEVISE,T_CONFIDENTIEL,T_DATECREATION, T_PAYEUR, T_ISPAYEUR, T_AVOIRRBT, T_RELANCEREGLEMENT, T_RELANCETRAITE '+
                            ',T_FERME, T_CORRESP1, T_CORRESP2, T_ESCOMPTE, T_REMISE, T_FACTURE, T_CREDITDEMANDE,T_CREDITACCORDE,T_CREDITPLAFOND'+
                            ',R_PAYS, R_ETABBQ, R_GUICHET,R_NUMEROCOMPTE,R_CLERIB,R_DOMICILIATION,R_CODEIBAN ' +
                            'FROM TIERS left outer join rib on ( t_auxiliaire=r_auxiliaire AND R_PRINCIPAL="X") ' +
                            ' Where  T_AUXILIAIRE="'+iCompte+ '"';

                            QD := OpenSQL(sSQL, true);
                            if not QD.EOF then
                            begin
                              T := TOB.Create ('TIERS',TTiers,-1);
                              T.SelectDB('', QD);
                              OKCompte := TRUE;
                              NatureCpte := T.GetValue ('T_NATUREAUXI');
                              lettrable := T.GetValue ('T_LETTRABLE');
                              OKLettrage := (ForceLettrage and (lettrable = 'X'));
                              Regimetva := T.GetValue('T_REGIMETVA');
                              Collectif := T.GetValue ('T_COLLECTIF');
                              TPayeur   := T.GetValue ('T_PAYEUR');
                              RibAuto       := ReChRibPrincipal(T, NatureCpte);
                              Fermeture := T.GetValue ('T_FERME');
                              fHTiers.Put( T.GetString('T_AUXILIAIRE'), TObject(T));
                            end;
                            Ferme(QD);

                  end;
             end;
        end;
     end;
end
else
begin
     if TGen <> nil then
     begin
        TA1 := TGenUpdate.FindFirst([Ext+'GENERAL'], [iCompte], FALSE);
        if TA1 <> nil then
        begin
             OKCompte  := TRUE;
             NatureCpte    := TA1.Getvaleur (FiTOBG_NATUREGENE);
             pointable := TA1.Getvaleur (FiTOBG_POINTABLE);
             lettrable := TA1.Getvaleur (FiTOBG_LETTRABLE);
             if lettrable = 'X' then OKLettrage := TRUE;
             Regimetva := TA1.GetValeur (FiTOBG_REGIMETVA);
             v1        := TA1.GetValeur (FiTOBG_VENTILABLE1);
             v2        := TA1.GetValeur (FiTOBG_VENTILABLE2);
             v3        := TA1.GetValeur (FiTOBG_VENTILABLE3);
             v4        := TA1.GetValeur (FiTOBG_VENTILABLE4);
             v5        := TA1.GetValeur (FiTOBG_VENTILABLE5);
             Fermeture := TA1.GetValue ('G_FERME');
             Etva      := TA1.GetValue ('G_TVA');
        end
        else
        begin
             //TA2 := TGenInsert.FindFirst([Ext+'GENERAL'], [Compte], FALSE);
             TA2 := RechercheOptimise(TGenInsert, Ext+'GENERAL', iCompte, TRUE, fHGeneInsert);
             if TA2 <> nil then
             begin
                  OKCompte  := TRUE;
                  NatureCpte    := TA2.Getvaleur (FiTOBG_NATUREGENE);
                  pointable := TA2.Getvaleur (FiTOBG_POINTABLE);
                  lettrable := TA2.Getvaleur (FiTOBG_LETTRABLE);
                  if lettrable = 'X' then OKLettrage := TRUE;
                  Regimetva := TA2.GetValeur (FiTOBG_REGIMETVA);
                  v1        := TA2.GetValeur (FiTOBG_VENTILABLE1);
                  v2        := TA2.GetValeur (FiTOBG_VENTILABLE2);
                  v3        := TA2.GetValeur (FiTOBG_VENTILABLE3);
                  v4        := TA2.GetValeur (FiTOBG_VENTILABLE4);
                  v5        := TA2.GetValeur (FiTOBG_VENTILABLE5);
                  Fermeture := TA2.GetValue ('G_FERME');
                  Etva      := TA2.GetValue ('G_TVA');
             end
             else
             begin
                  //TA3 := TGen.FindFirst([Ext+'GENERAL'], [Compte], FALSE);
                  TA3 := RechercheOptimise(TGen, Ext+'GENERAL', iCompte, TRUE, fHGene);
                  if TA3 <> nil then
                  begin
                       OKCompte  := TRUE;
                       NatureCpte    := TA3.Getvaleur (FiTOBG_NATUREGENE);
                       pointable := TA3.Getvaleur (FiTOBG_POINTABLE);
                       lettrable := TA3.Getvaleur (FiTOBG_LETTRABLE);
                       Regimetva := TA3.GetValeur (FiTOBG_REGIMETVA);
                       if lettrable = 'X' then OKLettrage := TRUE;
                       v1        := TA3.GetValeur (FiTOBG_VENTILABLE1);
                       v2        := TA3.GetValeur (FiTOBG_VENTILABLE2);
                       v3        := TA3.GetValeur (FiTOBG_VENTILABLE3);
                       v4        := TA3.GetValeur (FiTOBG_VENTILABLE4);
                       v5        := TA3.GetValeur (FiTOBG_VENTILABLE5);
                       Fermeture := TA3.GetValue ('G_FERME');
                       Etva      := TA3.GetValue ('G_TVA');
                  end
                  else   // ajout me 14-01-2005
                  begin
                            QD := OpenSQL('SELECT * FROM GENERAUX WHERE G_GENERAL="'+iCompte+'"', true);
                            if not QD.EOF then
                            begin
                              T := TOB.Create ('GENERAUX',TGen,-1);
                              T.SelectDB('', QD);
                              OKCompte := TRUE;
                              NatureCpte    := T.Getvaleur (FiTOBG_NATUREGENE);
                              pointable := T.Getvaleur (FiTOBG_POINTABLE);
                              lettrable := T.Getvaleur (FiTOBG_LETTRABLE);
                              Regimetva := T.GetValeur (FiTOBG_REGIMETVA);
                              OKLettrage := (lettrable = 'X');
                              v1        := T.GetValeur (FiTOBG_VENTILABLE1);
                              v2        := T.GetValeur (FiTOBG_VENTILABLE2);
                              v3        := T.GetValeur (FiTOBG_VENTILABLE3);
                              v4        := T.GetValeur (FiTOBG_VENTILABLE4);
                              v5        := T.GetValeur (FiTOBG_VENTILABLE5);
                              Fermeture := T.GetValue ('G_FERME');
                              Etva      := T.GetValue ('G_TVA');
                              fHGene.Put( T.GetString('G_GENERAL'), TObject(T));
                            end;
                            Ferme(QD);
                  end;
             end;
        end;
     end;
end;
if Fermeture = 'X' then  // fiche 10413
begin
              OnAfficheListeCom('Attention, le compte ' + iCompte + ' est fermé' ,LISTEIMPORT);
              OKexercice := FALSE;
end;
end;

Result := OKCompte;
end;

function TImportCom.RetournCompteattente (nature : string) :string;
var
Compteattente : string;
begin
      if nature = 'COC' then Compteattente := GetParamSocSecur ('SO_DEFCOLCLI', '')
      else
      if nature = 'COF' then Compteattente := GetParamSocSecur ('SO_DEFCOLFOU', '')
      else
      if nature = 'COS' then Compteattente := GetParamSocSecur ('SO_DEFCOLSAL', '')
      else
      if nature = 'COD' then Compteattente := GetParamSocSecur ('SO_DEFCOLDIV', '')
      else
      if nature = 'CLI' then
      begin
           if CpteClient <> '' then Compteattente := CpteClient
           else Compteattente := GetParamSocSecur ('SO_CLIATTEND', False);
      end
      else
      if nature = 'FOU' then
      begin
           if CpteFour <> '' then  Compteattente := CpteFour
           else Compteattente := GetParamSocSecur ('SO_FOUATTEND', False);
      end
      else
      if nature = 'SAL' then
      begin
           if CpteSalaire <> '' then Compteattente := CpteSalaire
           else Compteattente := GetParamSocSecur ('SO_SALATTEND', False);
      end
      else
      if nature = 'DIV' then
      begin
           if CpteDivers <> '' then Compteattente := CpteDivers
           else Compteattente := GetParamSocSecur ('SO_DIVATTEND', False);
      end
      else
      begin
           if CpteGene <> '' then  Compteattente := CpteGene
           else Compteattente := GetParamSocSecur ('SO_GENATTEND', False);
      end;
      // ajout V590
      Compteattente := UpperCase(Compteattente);
            // fiche 21484
      if (nature = 'CLI') or  (nature = 'FOU') or (nature = 'SAL')  or (nature = 'DIV') then
          AJouterListeSoldes (Compteattente, STiers)
      else
          AJouterListeSoldes (Compteattente, SGeneraux);

      Result := Compteattente;
end;

function TImportCom.MajDernLettrage (Ext, Compte, lettrage : string) : Boolean;
var
TA1,TA2,TA3   : TOB;
begin

if lettrage = '' then begin Result := FALSE; exit; end;

if Ext = 'T_' then
begin
     if TTiers <> nil then
     begin
        TA1 := TTiersUpdate.FindFirst([Ext+'AUXILIAIRE'], [Compte], FALSE);
        if TA1 <> nil then
        begin
             if TA1.GetValue (Ext+'LETTRABLE') = 'X' then
             begin
                  if TA1.GetValue (Ext+'DERNLETTRAGE') < lettrage then
                     TA1.PutValue (Ext+'DERNLETTRAGE', lettrage);
             end;
        end
        else
        begin
//             TA2 := TTiersInsert.FindFirst([Ext+'AUXILIAIRE'], [Compte], FALSE);
             TA2 := RechercheOptimise(TTiersInsert, Ext+'AUXILIAIRE', Compte, FALSE, fHTiersInsert);
             if TA2 <> nil then
             begin
                 if TA2.GetValue (Ext+'LETTRABLE') = 'X' then
                 begin
                      if TA2.GetValue (Ext+'DERNLETTRAGE') < lettrage then
                         TA2.PutValue (Ext+'DERNLETTRAGE', lettrage);
                 end;
             end
             else
             begin
//                  TA3 := TTiers.FindFirst([Ext+'AUXILIAIRE'], [Compte], FALSE);
                  TA3 := RechercheOptimise(TTiers, Ext+'AUXILIAIRE', Compte, FALSE, fHTiers);
                  if TA3 <> nil then
                  begin
                     if TA3.GetValue (Ext+'LETTRABLE') = 'X' then
                     begin
                          if TA3.GetValue (Ext+'DERNLETTRAGE') < lettrage then
                             TA3.PutValue (Ext+'DERNLETTRAGE', lettrage);
                     end;
                  end;
             end;
        end;
     end;
end
else
begin
     if TGen <> nil then    // ajout me 16-12-2002 'GENERAL AU lieu AUXILIAIRE'
     begin
        TA1 := TGenUpdate.FindFirst([Ext+'GENERAL'], [Compte], FALSE);
        if TA1 <> nil then
        begin
             if TA1.GetValue (Ext+'LETTRABLE') = 'X' then
             begin
                  if TA1.GetValue (Ext+'DERNLETTRAGE') < lettrage then
                     TA1.PutValue (Ext+'DERNLETTRAGE', lettrage);
             end;
        end
        else
        begin
             //TA2 := TGenInsert.FindFirst([Ext+'GENERAL'], [Compte], FALSE);
             TA2 := RechercheOptimise(TGenInsert, Ext+'GENERAL', Compte, TRUE, fHGeneInsert);
             if TA2 <> nil then
             begin
                 if TA2.GetValue (Ext+'LETTRABLE') = 'X' then
                 begin
                      if TA2.GetValue (Ext+'DERNLETTRAGE') < lettrage then
                         TA2.PutValue (Ext+'DERNLETTRAGE', lettrage);
                 end;
             end
             else
             begin
                  //TA3 := TGen.FindFirst([Ext+'GENERAL'], [Compte], FALSE);
                  TA3 := RechercheOptimise(TGen, Ext+'GENERAL', Compte, TRUE, fHGene);
                  if TA3 <> nil then
                  begin
                     if TA3.GetValue (Ext+'LETTRABLE') = 'X' then
                     begin
                          if TA3.GetValue (Ext+'DERNLETTRAGE') < lettrage then
                             TA3.PutValue (Ext+'DERNLETTRAGE', lettrage);
                     end;
                  end;
             end;
        end;
     end;
end;

Result := TRUE;
end;

procedure TImportCom.AffectationMontant(St: string; var T: TOB; ext: string);
var
  CodeMontant, Sens, lStDev : string;
  IndP, IndD, IndE: Integer;
  Montant: array[1..3] of Double;
  SMontant: array[1..3] of string;
  EnEuro: Boolean;
  debit, credit,cotation,mt       : double;
  debiteur, crediteur             : double;
begin
  EnEuro := {EuroOK and }TenueEuro;

  cotation := 1.0;
  debit := 0.0; credit := 0.0;
  debiteur := 0.0; crediteur := 0.0;
  CodeMontant := EnleveBlanc(St, 173, 3);
  if (CodeMontant = '') or (CodeMontant = '---') then
    CodeMontant := 'E--';
  IndP := Pos('F', CodeMontant);
  IndD := Pos('D', CodeMontant);
  IndE := Pos('E', CodeMontant);
  Fillchar(Montant, SizeOf(Montant), #0);
  SMontant[1] := StPoint(EnleveBlanc(St, 131, 20));
  SMontant[2] := StPoint(EnleveBlanc(St, 176, 20));
  SMontant[3] := StPoint(EnleveBlanc(St, 196, 20));
  if SMontant[1] <> '' then
    Montant[1] := Valeur(SMontant[1]);
  if SMontant[2] <> '' then
    Montant[2] := Valeur(SMontant[2]);
  if SMontant[3] <> '' then
    Montant[3] := Valeur(SMontant[3]);

  Sens := EnleveBlanc(St, 130, 1);
  if IndP > 0 then // si montant en Franc
  begin
    if EnEuro then
    begin
      if Sens = 'D' then   debiteur := Montant[IndP]
      else if Sens = 'C' then crediteur := Montant[IndP];
    end
    else
    begin
         if Sens = 'D' then  debit := Montant[IndP]
         else if Sens = 'C' then  credit := Montant[IndP];
    end;

    // on renseigne les devises dans tous les cas
    if Sens = 'D' then
      T.PutValue(ext + 'DEBITDEV', Montant[IndP])
    else if Sens = 'C' then
      T.PutValue(ext + 'CREDITDEV', Montant[IndP]);
  end;
  if IndE > 0 then // si montant en euro
  begin
    if Sens = 'D' then debit := Montant[IndE]
    else if Sens = 'C' then  credit := Montant[IndE];
    // on renseigne les devises dans tous les cas
    if Sens = 'D' then
      T.PutValue(ext + 'DEBITDEV', Montant[IndE])
    else if Sens = 'C' then
      T.PutValue(ext + 'CREDITDEV', Montant[IndE]);
  end;
  if (IndD > 0) and (EnleveBlanc(St, 163, 10)<> '') then   // si montant en devise
  begin
        if ((IndD>0) and (Montant[IndD] <> 0)) and ((IndE>0) and (Montant[IndE] <> 0)) then  //fiche 14015
        begin
                if Sens = 'D' then
                    T.PutValue(ext + 'DEBITDEV', Montant[IndD])
                else if Sens = 'C' then
                    T.PutValue(ext + 'CREDITDEV', Montant[IndD]);
                if Ext = 'E_' then
                begin
                      lStDev := T.GetValue(ext + 'DEVISE');
                      if ( ( lStDev = V_PGI.DevisePivot ) or ( lStDev = V_PGI.DeviseFongible)) then
                         T.PutValue (ext+'COTATION', 1.0)
                      else
                      begin // dans la table chancellerie h_cotation=montant devise/montant euro
                         mt :=  Arrondi(Montant[IndE] / Montant[IndD], 9);
                         T.PutValue(Ext+'COTATION', mt);
                         T.PutValue(Ext + 'TAUXDEV', mt*V_PGI.TauxEuro);
                      end;
                end;
        end
        else
        begin
               // if ext = 'E_' then fiche  19917
                if T.FieldExists(Ext + 'COTATION') then
                         cotation := T.GetValue (ext+'COTATION');
                if Montant[IndD] <> 0 then  // si montant devise renseigné
                begin
                        if Sens = 'D' then
                        begin
                            T.PutValue(ext + 'DEBITDEV', Montant[IndD]);
                            //if ext = 'E_' then   // fiche  19917
                            if cotation <> 1 then
                               debit := Montant[IndD]* cotation
                            else
                               debit := Montant[IndD];
                        end
                        else if Sens = 'C' then
                        begin
                            T.PutValue(ext + 'CREDITDEV', Montant[IndD]);
                            //if ext = 'E_' then  // fiche  19917
                            if cotation <> 1 then
                               credit := Montant[IndD]* cotation
                            else credit := Montant[IndD];
                        end;
                end
                else
                begin
  //////////////    Cas d'écart d'échange
                // DE- ou ED- montant devise à zéro et montant en euro renseigné
                if BEcartChange or (cotation = 0) then
                begin
                      if Sens = 'D' then
                      begin
                          T.PutValue(ext + 'DEBITDEV', Montant[IndD]);
                          debit := Montant[IndE];
                      end
                      else if Sens = 'C' then
                      begin
                          T.PutValue(ext + 'CREDITDEV', Montant[IndD]);
                          credit := Montant[IndE];
                      end;
                end
                else
                begin
                     if Montant[IndE] <> 0 then  
                     begin
                        if Sens = 'D' then
                        begin
                            T.PutValue(ext + 'DEBITDEV', Montant[IndE]/ cotation);
                            debit := Montant[IndE];
                        end
                        else if Sens = 'C' then
                        begin
                            T.PutValue(ext + 'CREDITDEV', Montant[IndE]/cotation);
                            credit := Montant[IndE];
                        end;
                     end;
                end;
  ///////////////
                end;
        end;
  end
  else
  begin
        if (IndD > 0) and (EnleveBlanc(St, 163, 10)= '') then
        begin
            if Sens = 'D' then
            begin
              T.PutValue(ext + 'DEBITDEV', Montant[IndD]);
// ME to GP 10620
//              debit := Montant[IndE];
              if IndE > 0 then debit := Montant[IndE] else Debit := 0;  // Fiche 10620
              // cas de cotation non renseignée D-- et uniquement montant den devise
              // alors on calcule le montant devise pivot à partir de la table chancellerie
              if (Debit = 0) then
              begin
                  if(T.GetValue (ext+'COTATION') <> 1) and (T.GetValue(Ext + 'TAUXDEV') <> 1) then
                      debit := Montant[IndD]* T.GetValue (ext+'COTATION')
                  else
                  begin
                          OnAfficheListeCom('Attention la cotation n''est pas renseignée dans la table chancellerie' ,LISTEIMPORT);
                          OKexercice := FALSE;
                  end;
              end;
            end
            else if Sens = 'C' then
            begin
              T.PutValue(ext + 'CREDITDEV', Montant[IndD]);
// ME to GP 10620
//              credit := Montant[IndE];
              if IndE > 0 then credit := Montant[IndE] else credit := 0;   // Fiche 10620
              // cas de cotation non renseignée D-- et uniquement montant den devise
              // alors on calcule le montant devise pivot à partir de la table chancellerie
              if (credit = 0) then
              begin
                  if (T.GetValue (ext+'COTATION') <> 1) and (T.GetValue(Ext + 'TAUXDEV') <> 1) then
                      credit := Montant[IndD]* T.GetValue (ext+'COTATION')
                  else
                  begin
                          OnAfficheListeCom('Attention la cotation n''est pas renseignée dans la table chancellerie' ,LISTEIMPORT);
                          OKexercice := FALSE;
                  end;
              end;
            end;
        end;
  end;
 If (Arrondi(debit+credit,V_PGI.OkDecV)=0) And (Arrondi(debiteur+crediteur,V_PGI.OkDecE)<>0)Then
  BEGIN
  If EnEuro Then
    BEGIN
    debit:=PivotToEuro(debiteur) ;
    credit:=PivotToEuro(crediteur) ;
    END Else
    BEGIN
    debit:=EuroToPivot(debiteur) ;
    credit:=EuroToPivot(crediteur) ;
    END ;
  END;

                          //AJOUT ME 16-03-2005
  T.Putvalue (ext + 'DEBIT', Arrondi(debit,V_PGI.OkDecV) );
  T.Putvalue (ext + 'CREDIT', Arrondi(credit, V_PGI.OkDecV) );
end;

procedure TImportCom.AffectationMontantcouverture(St: string; var T: TOB; ext: string);
var
  mt                   : double;
  etatlettrage         :  string;
begin
  if Ext = 'Y_' then exit;
  etatlettrage := EnleveBlanc(St, 1021, 3);

  if  (etatlettrage <> 'TL')
  and (T.GetValue(Ext + 'COUVERTURE') = 0) and  (T.GetValue(Ext + 'COUVERTUREDEV') = 0) then exit;
// à voir pour recalcul
  if  (etatlettrage = 'PL') and  (T.GetValue(Ext + 'COUVERTURE') <> 0) and
  (T.GetValue(Ext + 'COUVERTUREDEV') = 0) then
  T.PutValue(Ext + 'COUVERTUREDEV', T.GetValue(Ext + 'COUVERTURE'));

  if  (etatlettrage = 'TL') then
  begin
     if (T.GetValue(Ext + 'DEBIT') > 0.0) or (T.GetValue(Ext + 'DEBIT') < 0.0) then
        mt := T.GetValue(Ext + 'DEBIT')
     else
        mt := T.GetValue(Ext + 'CREDIT');
     if mt <> T.GetValue(Ext + 'COUVERTURE') then
        T.PutValue(Ext + 'COUVERTURE', mt);

     if (T.GetValue(Ext + 'DEBITDEV') > 0.0) or  (T.GetValue(Ext + 'DEBITDEV') < 0.0) then
        mt := T.GetValue(Ext + 'DEBITDEV')
     else
        mt := T.GetValue(Ext + 'CREDITDEV');
     if mt <> T.GetValue(Ext + 'COUVERTUREDEV') then
     T.PutValue(Ext + 'COUVERTUREDEV', mt);
  end;

end;

procedure TImportCom.InsertCompte (Ext, Compte, libelle :string; Var nature : string; Collectif : string=''; EtatLettrage : string='');
var
TA2        : TOB;
Coll,Sens  : string;
LibRef     : string;
Q          : TQuery;
CC         : string;
Cycle      : string;
T_Info     : TFTransfertInfo;
        procedure insertlmdr;
        var
        TMdr      : TOB;
        begin
            TMdr := TOB.Create('MODEREGL', nil, -1);

            TMdr.PutValue ('MR_MODEREGLE', '001');
            TMdr.PutValue ('MR_LIBELLE', 'ESPECES');
            TMdr.PutValue ('MR_APARTIRDE', 'ECR');
            TMdr.PutValue ('MR_PLUSJOUR', 0);
            TMdr.PutValue ('MR_ARRONDIJOUR', 'PAS');
            TMdr.PutValue ('MR_NOMBREECHEANCE', 1);
            TMdr.InsertOrUpdateDB(TRUE);
            TMdr.Free;

        end;
begin

if (not AutoriseEcheck) then exit;

if compte = '' then exit;
if libelle <>  '' then LibRef := VireTouteSub(libelle,'"')
else  LibRef := 'Créé par Comsx';

// ajout V590
compte := UpperCase(compte);

if Ext = 'T_' then
begin
     if TTiersInsert <> nil then
     begin
             CC := BourreOuTronque(Compte, fbaux);
//             TA2 := TTiersInsert.FindFirst([Ext+'AUXILIAIRE'], [CC], FALSE);
             TA2 := RechercheOptimise(TTiersInsert, Ext+'AUXILIAIRE', CC, FALSE, fHTiersInsert);

             if TA2 = nil then // creation
             begin
                      TA2 := TOB.Create('TIERS', TTiersInsert, -1);
                      TA2.PutValue('T_AUXILIAIRE', CC);
                      TA2.PutValue('T_TIERS', CC);
                      OnAfficheListeCom('Tiers ' + CC + ' est inexistant (création par ComSx)',LISTEIMPORT);

                      TA2.PutValue('T_LIBELLE', VireTouteSub(libelle,'"'));
                      TA2.PutValue('T_NATUREAUXI', nature);
                      if Collectif = '' then
                      begin
                            if nature ='CLI' then
                            begin
                                 Coll := GetParamSocSecur ('SO_DEFCOLCLI', '');
                                 T_Info.iCompte := Coll;
                                 if not ExisteCompte ('G_', T_Info) then Coll := '';
                                 if (Coll = '') or (T_Info.NatureCpte <> 'COC') then    // ajout me 22/05/2003
                                 begin
                                      Coll := BourreOuTronque('4119999999999', fbGene); nature := 'COC';
                                      T_Info.iCompte := Coll;
                                      if not ExisteCompte ('G_', T_Info) then
                                         InsertCompte ('G_', BourreOuTronque(Coll, fbgene), 'COMPTE GENERAL', nature);
                                 end;
                            end
                            else
                            if nature ='FOU' then
                            begin
                                 Coll := GetParamSocSecur ('SO_DEFCOLFOU', False);
                                 T_Info.iCompte := Coll;
                                 if not ExisteCompte ('G_', T_Info) then Coll := '';
                                 if (Coll = '') or (T_Info.NatureCpte <> 'COF') then
                                 begin
                                      Coll := BourreOuTronque('4019999999999', fbGene);  nature := 'COF';
                                      T_Info.iCompte := Coll;
                                      if not ExisteCompte ('G_', T_Info) then
                                      InsertCompte ('G_', Coll, 'COMPTE GENERAL', nature);
                                 end;
                            end
                            else
                            if nature ='SAL' then
                            begin
                                 Coll := GetParamSocSecur ('SO_DEFCOLSAL', False);
                                 T_Info.iCompte := Coll;
                                 if not ExisteCompte ('G_', T_Info) then Coll := '';
                                 if (Coll = '') or (T_Info.NatureCpte <> 'COS') then
                                 begin
                                      Coll := BourreOuTronque('4219999999999', fbGene); nature := 'COS';
                                      T_Info.iCompte := Coll;
                                      // fiche 10524
                                      if not ExisteCompte ('G_', T_Info) then
                                      InsertCompte ('G_', Coll, 'COMPTE GENERAL', nature);
                                 end;
                            end
                            else
                            begin
                                 Coll := GetParamSocSecur ('SO_DEFCOLDIV', False);
                                 T_Info.iCompte := Coll;
                                 if not ExisteCompte ('G_', T_Info) then Coll := '';
                                 if (Coll = '') or (T_Info.NatureCpte <> 'COD') then
                                 begin
                                      Coll := BourreOuTronque('4679999999999', fbGene); nature := 'COD';
                                      T_Info.iCompte := Coll;
                                      if not ExisteCompte ('G_', T_Info) then
                                      InsertCompte ('G_', Coll, 'COMPTE GENERAL', nature);
                                 end;
                            end;
                      end
                      else
                      Coll := Collectif;
                      TA2.PutValue('T_COLLECTIF', BourreOuTronque(Coll, fbgene));
                      if GetParamSocSecur('SO_GCMODEREGLEDEFAUT', False) = '' then
                      begin
                           Q := OpenSql ('Select MR_MODEREGLE from MODEREGL', TRUE);
                           if (Mrdefaut = '') and (not Q.EOF) then
                              TA2.PutValue('T_MODEREGLE', Q.FindField ('MR_MODEREGLE').asstring)
                           else
                           begin
                              TA2.PutValue('T_MODEREGLE', RendMRDefaut);
                              insertlmdr;
                           end;
                           Ferme (Q);
                      end
                      else
                      TA2.PutValue('T_MODEREGLE', RendMRDefaut);
                      TA2.PutValue('T_LETTRABLE', 'X');
                      TA2.PutValue('T_DATECREATION', Date-1);
                      TA2.PutValue('T_ABREGE',Copy(LibRef,1,17)) ;
                      TA2.PutValue('T_REGIMETVA', RendRegTvadefaut);
                      TA2.PutValue('T_MULTIDEVISE', 'X');
                      TA2.PutValue('T_DEVISE', '');
                      TA2.PutValue('T_CONFIDENTIEL', '0');
                      if (Nature = 'CLI') or (Nature = 'FOU') then
                      begin
                           TA2.PutValue('T_MODEREGLE',RendMRDefaut);
                           TA2.PutValue('T_TVAENCAISSEMENT', GetParamSocSecur('SO_CODETVADEFAUT', False));
                      end;
                      fHTiersInsert.Put(CC, TObject(TA2));
             end;
     end;
end
else
begin
     if TGenInsert <> nil then
     begin
             CC := BourreOuTronque(Compte, fbGene);
//             TA2 := TGenInsert.FindFirst([Ext+'GENERAL'], [CC], FALSE);
             TA2 := RechercheOptimise(TGenInsert, Ext+'GENERAL', CC, TRUE, fHGeneInsert);
             if TA2 = nil then // creation
             begin
                      TA2 := TOB.Create('GENERAUX', TGenInsert, -1);
                      TA2.PutValue('G_GENERAL', CC);
                      OnAfficheListeCom('Compte ' + CC + ' est inexistant (création par ComSx)',LISTEIMPORT);

                      RemplirNatureSens(compte,'', Nature, Sens);
                      // fiche 10107
                      if (Nature = 'DIV') and (EtatLettrage = 'TL') then
                         TA2.PutValue('G_LETTRABLE', 'X') ;


                      TA2.PutValue('G_LIBELLE',LibRef) ;
                      TA2.PutValue('G_ABREGE',Copy(LibRef,1,17)) ;
                      TA2.PutValue('G_NATUREGENE', Nature);
                      TA2.PutValue('G_SENS', Sens) ;
                      if (nature = 'COC') or (nature = 'COF') or (nature = 'COS') or (nature = 'COD') then
                       TA2.PutValue('G_COLLECTIF', 'X');

                      TA2.PutValue('G_VENTILABLE1', '-');
                      TA2.PutValue('G_VENTILABLE2', '-');
                      TA2.PutValue('G_VENTILABLE3', '-');
                      TA2.PutValue('G_VENTILABLE4', '-');
                      TA2.PutValue('G_VENTILABLE5', '-');
                      TA2.PutValue('G_VENTILABLE', '-');
                      TA2.PutValue('G_CONFIDENTIEL', '0');
                      TA2.PutValue('G_VISAREVISION', '-'); // fiche 21483
                      // fiche 10447
                      if (ctxPCL in V_PGI.PGIContexte) then
                      begin
                              if GetParamSocSecur('SO_CPPLANREVISION', '') <> '' then
                              begin
                                       Cycle := TrouveRubriqueDuCpteGene(CC);
                                          // Aucun cycle de révision ne contient le général saisi
                                       if (Cycle = '') and GetParamSocSecur('SO_CPREVISBLOQUEGENE', False) then
                                       begin
                                                 OnAfficheListeCom('Création impossible. Le compte général '+ CC +' n''appartient à aucun cycle de révision.',LISTEIMPORT);
                                                 OkCreatCompte := FALSE;
                                                 OKexercice    := TRUE;
                                       end
                                       else
                                       TA2.PutValue('G_CYCLEREVISION', Cycle);
                              end;
                              if (not GetParamSocSecur('SO_CPREVISLIASSEMODIF', False)) and
                                 (GetParamSocSecur('SO_CPCONTROLELIASSE', False) <> '') and
                                 (GetParamSocSecur('SO_CPCONTROLELIASSE', False) <> 'AUC') then
                              begin
                                      // Recherche du général dans les liasses ( sur Table RUBRIQUE )
                                      if TrouveRubriqueDuCpteGene(CC, TRUE) = '' then
                                      begin
                                        if GetParamSocSecur('SO_CPREVISBLOQUELIASSE', False) then
                                        begin
                                                 OnAfficheListeCom('Création impossible. Le compte général '+ CC +' n''appartient pas à la liasse spécifiée dans vos paramètres société.',LISTEIMPORT);
                                                 OkCreatCompte := FALSE;
                                                 OKexercice    := TRUE;
                                        end;
                                      end;
                              end;
                      end;
                      TA2.PutValue('G_DATECREATION', Date-1);
                      fHGeneInsert.Put( CC, TObject(TA2));
                      ModeEVentilType := TRUE;
             end;
     end;
end;

end;


function TImportCom.RendModePaiementParDefaut (Modereg : string= ''): string ;
var
T     : TOB;
Mdr   : string;
Q     : Tquery;
begin
  Result := '';
  if Modereg = '' then
  begin
        if not OkModePaieDefaut then
        begin
             Mdr        := RendMRDefaut;
             MdpDefault := ModePaiementParDefaut (Mdr);
             if not ExisteSQL ('SELECT * FROM MODEPAIE WHERE MP_MODEPAIE="'+MdpDefault+'"') then
             begin
                T := TOB.Create ('MODEPAIE',nil,-1);
                try
                  T.PutValue('MP_MODEPAIE', MdpDefault);
                  T.PutValue('MP_LIBELLE', MdpDefault);
                  T.PutValue('MP_ABREGE', MdpDefault);
                  T.PutValue('MP_POINTABLE','-');
                  T.InsertDB(nil);
                finally
                  T.Free;
                end;
             end;
             OkModePaieDefaut := True;
        end;
        Result := MdpDefault;
  end
  else
  begin
      Q := OPenSQL ('select MR_MP1 from moderegl WHERE MR_MODEREGLE="'+Modereg+'"', TRUE);
      if not Q.EOF then
        Result := Q.FindField ('MR_MP1').asstring;
      Ferme (Q);
  end;
end;

procedure TImportCom.DecoupageJournal(TLec,TLecpiece,TLecPeriode : TOB; var numpiece : integer);
var
ie,numligne, ii    : integer;
Soldedeb,SoldeCre  : double;
Tobnewp            : TOB;
Nb                 : integer;
begin
   if (NatureFichier = 'JRL') then Nb := 80
   else  Nb := 101;
   Soldedeb := 0.0; SoldeCre := 0.0;
   if TLec.detail.Count-1 >= Nb then
   begin
        Tobnewp := TLecpiece.detail[0];
        ie := 0;
        if (NatureFichier = 'JRL') then
        begin
            // recherche de la première piéce équilibrée
            While (TLec.detail.Count > 0) do
            begin
                 if (Arrondi(Soldedeb,V_PGI.OkDecV) = Arrondi(Soldecre,V_PGI.OkDecV)) and
                    (Soldedeb <> 0.0) and (SoldeCre <> 0.0) and (ie > 50) then break;
                 Soldedeb := Soldedeb + Tlec.detail[ie].GetValue ('E_DEBIT');
                 Soldecre := Soldecre + TLec.detail[ie].GetValue ('E_CREDIT');
                 inc (ie);
                 if TLec.detail.Count = ie then break;
            end;
        end
        else ie := 100;

        Soldedeb := 0.0; SoldeCre := 0.0;
        numligne := ie;
        While ((TLec.detail.Count - ie) > 0) do
        begin
             if ((NatureFichier = 'JRL') and (Arrondi(Soldedeb,V_PGI.OkDecV) = Arrondi(Soldecre,V_PGI.OkDecV))
             and (numligne > 50))
             or ((NatureFichier = 'BAL') and (numligne = 100))then
             begin
                inc (numpiece);
                Tobnewp     := TLecPeriode.FindFirst(['E_NUMEROPIECE'], [numpiece], FALSE);
                if Tobnewp = nil then
                begin
                     Tobnewp:=TOB.Create('NUMEROPIECE',TLecpiece,-1) ;
                     Tobnewp.AddChampSupValeur('E_NUMEROPIECE', numpiece);
                     TLecPeriode.putValue('J_NUMDERNMVT', numpiece);
                     // Correction bug découpage à partir de Sisco micro de folio
                     // les champs n'étaient pas ajoutés  03/08/2006
                     Tobnewp.AddChampSupValeur('SOLDE', 0);
                     Tobnewp.AddChampSupValeur('SOLDEDEV', 0);
                     Tobnewp.AddChampSupValeur('E_NUMEROPIECES1', '');
                     Tobnewp.AddChampSupValeur('E_DOCIDS1', '');
                     Tobnewp.AddChampSupValeur('E_DOCIDGED', '');
                     Tobnewp.AddChampSupValeur('NATUREPIECE', '');
                     Tobnewp.AddChampSupValeur('COMPTEINSERT', '');
                     Tobnewp.AddChampSupValeur ('NUMFOLIO',1);
                     Tobnewp.AddChampSupValeur('OKREJET', '-');
                     Tobnewp.AddChampSupValeur('OKDOUBLON', '-');
                end;
                numligne := 1;
             end;
             Soldedeb := Soldedeb + Tlec.detail[ie].GetValue ('E_DEBIT');
             Soldecre := Soldecre + TLec.detail[ie].GetValue ('E_CREDIT');
             TLec.detail[ie].PutValue ('E_NUMLIGNE', numligne);
             TLec.detail[ie].PutValue ('E_NUMEROPIECE', numpiece);
             inc (numligne);
             if TLec.detail.Count  > 0 then
                TLec.detail[ie].ChangeParent(Tobnewp,-1) ;
             if TLec.detail.Count = 0 then break;
        end;
         //Fiche 10518
        Soldedeb := 0.0; SoldeCre := 0.0;
        for ie := 0 to TLecpiece.detail.count-1 do
        begin
             Tobnewp := TLecpiece.detail[ie];
             for ii := 0 To Tobnewp.detail.count-1 do
             begin
             Soldedeb := Soldedeb + Tobnewp.detail[ii].GetValue ('E_DEBIT');
             Soldecre := Soldecre + Tobnewp.detail[ii].GetValue ('E_CREDIT');
             end;
             if(Arrondi(Soldedeb,V_PGI.OkDecV) <> Arrondi(Soldecre,V_PGI.OkDecV))  then
             begin
               GenereCompteDattente (Tobnewp, Tobnewp.detail.Count, Soldedeb, SoldeCre);
               Soldedeb := 0.0; SoldeCre := 0.0;
             end;
        end;
   end;
end;

procedure  TImportCom.MiseNumVentil (var Tlec : TOB);
var
ie,Ligne                                           : integer;
ia,iana                                            : integer;
Tobaxe,Tobana,Tobattente                           : TOB;
EquilODA                                           : Boolean;
Soldedeb, Soldecre                                 : double;
Axe                                                : string;
LeFb                                               : TFichierBase;
begin
   Ligne := 1;
   EquilODA := GetParamSocSecur('SO_EQUILANAODA', False);
   for ie:=0 to TLec.Detail.Count-1  do
   begin
           Soldedeb := 0.0; Soldecre := 0.0;
           Tobaxe := Tlec.detail[ie];
           for ia:=0 to Tobaxe.Detail.Count-1  do
           begin
                Tobana := Tobaxe.detail[ia];
                for iana:=0 to Tobana.Detail.Count-1  do
                begin
                     if Tobana.detail[iana].GetValue('Y_TYPEANALYTIQUE')= 'X' then
                     begin
                        Tobana.detail[iana].PutValue ('Y_NUMVENTIL', ligne);
                        Tobana.detail[iana].PutValue ('Y_NUMLIGNE', 0);
                     end;
                     if EquilODA then
                     begin
                        Soldedeb  := Soldedeb + Tobana.detail[iana].Getvalue ('Y_DEBIT');
                        Soldecre  := Soldecre + Tobana.detail[iana].GetValue ('Y_CREDIT');
                     end;
                     inc (ligne);
                end;
                // fiche 10319
                Soldedeb := Arrondi(Soldedeb, V_PGI.OkDecV);
                Soldecre := Arrondi(Soldecre, V_PGI.OkDecV);
                if (Tobana.Detail.Count <> 0) and (Soldedeb <> Soldecre) then
                begin
                      Tobattente   := TOB.Create ('ANALYTIQ', Tobana, -1);
                      Tobattente.PutValue ( 'Y_JOURNAL', Tobana.detail[0].Getvalue('Y_JOURNAL') );
                      Tobattente.PutValue ( 'Y_EXERCICE', Tobana.detail[0].Getvalue('Y_EXERCICE') );
                      Tobattente.PutValue ( 'Y_DATECOMPTABLE', Tobana.detail[0].Getvalue('Y_DATECOMPTABLE') );
                      Tobattente.PutValue ( 'Y_NUMEROPIECE', Tobana.detail[0].Getvalue('Y_NUMEROPIECE') );
                      Axe := Tobana.detail[0].Getvalue('Y_AXE');
                      Tobattente.PutValue ( 'Y_AXE', Axe);
                      Tobattente.PutValue ( 'Y_QUALIFPIECE', Tobana.detail[0].Getvalue('Y_QUALIFPIECE') );
                      LeFb :=fbAxe1;
                      if Axe = 'A1' then LeFb :=fbAxe1;
                      if Axe = 'A2' then LeFb :=fbAxe2;
                      if Axe = 'A3' then LeFb :=fbAxe3;
                      if Axe = 'A4' then LeFb :=fbAxe4;
                      if Axe = 'A5' then LeFb :=fbAxe5;

                      Tobattente.PutValue ( 'Y_SECTION', GetInfoCpta(LeFb).Attente);  
                      Tobattente.PutValue ( 'Y_GENERAL', Tobana.detail[0].Getvalue('Y_GENERAL') );
                      Tobattente.PutValue ( 'Y_REFINTERNE', Tobana.detail[0].Getvalue('Y_REFINTERNE') );
                      Tobattente.PutValue ( 'Y_NATUREPIECE', Tobana.detail[0].Getvalue('Y_NATUREPIECE') );
                      Tobattente.PutValue ( 'Y_TYPEANALYTIQUE', Tobana.detail[0].Getvalue('Y_TYPEANALYTIQUE') );
                      Tobattente.PutValue ( 'Y_DATEREFEXTERNE', Tobana.detail[0].Getvalue('Y_DATEREFEXTERNE') );
                      Tobattente.PutValue ( 'Y_ETABLISSEMENT', Tobana.detail[0].Getvalue('Y_ETABLISSEMENT') );
                      Tobattente.PutValue ( 'Y_DEVISE', Tobana.detail[0].Getvalue('Y_DEVISE') );
                      Tobattente.PutValue ( 'Y_PERIODE', Tobana.detail[0].Getvalue('Y_PERIODE') );
                      Tobattente.PutValue ( 'Y_SEMAINE', Tobana.detail[0].Getvalue('Y_SEMAINE') );
                      Tobattente.PutValue ( 'Y_CREERPAR', Tobana.detail[0].Getvalue('Y_CREERPAR') );
                      Tobattente.PutValue ( 'Y_LIBELLE', 'GENEREE PAR COMSX' );
                      Tobattente.PutValue ( 'Y_NUMVENTIL', ligne);
                      Tobattente.PutValue ( 'Y_NUMLIGNE', 0);
                      if Soldedeb > Soldecre then
                      begin
                         Tobattente.PutValue ('Y_CREDIT', (Soldedeb - Soldecre));
                         Tobattente.PutValue ('Y_DEBIT', 0.0);
                         Tobattente.PutValue ('Y_DEBITDEV', 0.0);
                         Tobattente.PutValue ('Y_CREDITDEV', (Soldedeb - Soldecre));
                      end
                      else
                      if Soldedeb < Soldecre then
                      begin
                         Tobattente.PutValue ('Y_CREDIT', 0.0);
                         Tobattente.PutValue ('Y_DEBIT', (Soldecre - Soldedeb));
                         Tobattente.PutValue ('Y_DEBITDEV', (Soldecre - Soldedeb));
                         Tobattente.PutValue ('Y_CREDITDEV', 0.0);
                      end;
                      OnAfficheListeCom('Journal déséquilibré : '+ Tobana.detail[0].Getvalue('Y_JOURNAL')+' Période : ' +FormatDateTime(Traduitdateformat('mm/dd/yyyy'),Tobana.detail[0].Getvalue('Y_DATECOMPTABLE')) + ', Pièce : ' + IntToStr(Tobana.detail[0].Getvalue('Y_NUMEROPIECE')),LISTEIMPORT);
                      if BRejet then
                         OkCreatCompte := FALSE;
                      inc(NbEcrana);
                end;
           end;
     end;
end;

procedure TImportCom.GenereCompteDattente (TLec : TOB; ie : integer; Soldedeb, SoldeCre : double);
var
Tobattente : TOB;
begin
      Tobattente   := TOB.Create ('ECRITURE', TLec, ie);
      CPutDefautEcr (Tobattente);
      CSupprimerInfoLettrage (Tobattente); // à voir pour compte général lettrable
      CDupliquerTOBEcr (TLec.detail[ie-1], Tobattente);
      Tobattente.PutValue ( 'E_GENERAL', GetParamSocSecur('SO_GENATTEND',''));
      Tobattente.PutValue ( 'E_LIBELLE', 'GENEREE PAR COMSX' );
      if Soldedeb > Soldecre then
      begin
         Tobattente.PutValue ('E_CREDIT', (Soldedeb - Soldecre));
         Tobattente.PutValue ('E_DEBIT', 0.0);
         Tobattente.PutValue ('E_DEBITDEV', 0.0);
         Tobattente.PutValue ('E_CREDITDEV', (Soldedeb - Soldecre));
         TotalMouvCredit := TotalMouvCredit + (Soldedeb - Soldecre);
      end;
      if Soldedeb < Soldecre then
      begin
         Tobattente.PutValue ('E_CREDIT', 0.0);
         Tobattente.PutValue ('E_DEBIT', (Soldecre - Soldedeb));
         Tobattente.PutValue ('E_DEBITDEV', (Soldecre - Soldedeb));
         Tobattente.PutValue ('E_CREDITDEV', 0.0);
         TotalMouvDebit  := TotalMouvDebit + (Soldecre - Soldedeb);
      end;
end;


procedure TImportCom.RemplirNatureSens ( Gene,Auxi : String ; Var Nature,Sens : String ) ;
Var CC,C2,C3 : Char ;
BEGIN
if nature = 'COC' then begin Sens := 'D'; exit; end
else
if nature = 'COF' then begin Sens := 'C'; exit; end
else
if Nature <> '' then exit;

CC:=Gene[1] ; C2:=Gene[2] ; C3:=Gene[3] ;
Nature:='' ; Sens:='' ;
Case CC of
   '1' : BEGIN Nature:='DIV' ; Sens:='C' ; END ;
   '2' : BEGIN
         Nature:='IMO' ; Sens:='D' ;
         if C2 in ['8','9'] then BEGIN Nature:='DIV' ; Sens:='C' ; END ;
         END ;
   '3' : BEGIN
         Nature:='DIV' ; Sens:='D' ;
         if C2='9' then BEGIN Nature:='DIV' ; Sens:='C' ; END ;
         END ;
   '4' : BEGIN
         Nature:='DIV' ; if Auxi<>'' then Nature:='COD' ;
         Case C2 of
            '0' : BEGIN
                  Sens:='C' ; if C3='9' then Sens:='D' ;
                  if Auxi<>'' then
                       Nature:='COF'
                  else
                       if gene = GetParamSocSecur ('SO_DEFCOLFOU', False) then Nature:='COF';
                  END ;
            '1' : BEGIN
                  Sens:='D' ; if C3='9' then Sens:='C' ;
                  if Auxi<>'' then Nature:='COC'
                  else
                       if gene = GetParamSocSecur ('SO_DEFCOLCLI', '') then Nature:='COC';
                  END ;
            '8' : BEGIN Sens:='D' ; END ;
             else BEGIN Sens:='C' ; END ;
            END ;
         END ;
   '5' : BEGIN
         Nature:='DIV' ; Sens:='D' ;
         if C2='1' then BEGIN Nature:='BQE' ; Sens:='D' ; END else
         if C2='3' then BEGIN Nature:='CAI' ; Sens:='D' ; END else
         if C2='9' then BEGIN Nature:='DIV' ; Sens:='C' ; END ;
         END ;
   '6' : BEGIN Nature:='CHA' ; Sens:='D' ; END ;
   '7' : BEGIN Nature:='PRO' ; Sens:='C' ; END ;
   END ;
END ;


// suppression des écritures

procedure TEvtSup.DegageEcriture;
var
ETab      : string;
Q         : Tquery;
          procedure suppEcriture;
          begin
            ExecuteSQL('DELETE FROM Ecriture WHERE E_EXERCICE="' + Codeexo + '"'
              + ' AND E_JOURNAL="' + Journal + '" AND E_PERIODE=' +
                IntToStr(vPeriode) + ' AND E_NUMEROPIECE=' + npiece);
            ExecuteSQL('DELETE FROM Analytiq WHERE Y_EXERCICE="' + Codeexo + '"'
              + ' AND Y_JOURNAL="' + Journal + '" AND Y_PERIODE=' +
                IntToStr(vPeriode) + ' AND Y_NUMEROPIECE=' + npiece);
              // ecriture complémentaire
             ExecuteSQL('DELETE FROM ECRCOMPL WHERE EC_EXERCICE="' + Codeexo + '"'
              + ' AND EC_JOURNAL="' + Journal + '" AND EC_DATECOMPTABLE="' + USDATETIME(DateComptable) +
              '" AND EC_NUMEROPIECE=' + npiece);
          end;
begin
  ETab := '';
{$IFNDEF EAGLSERVER}
  if Codeexo = '' then
       Codeexo := QUELEXODTBUD(Datecomptable);
{$ENDIF}

  if parpiece then  
  begin
       if (refrevision <> '') then
       begin
           Q := Opensql ('select e_refrevision,e_numeropiece,e_datecomptable from ecriture where e_refrevision='+refrevision + ' Group by e_refrevision,e_numeropiece,e_datecomptable' , TRUE);
           if not Q.EOF then
           begin
               while not Q.EOF do
               begin
                    npiece :=  IntToStr(Q.Findfield('e_numeropiece').asinteger);
                    // pour S1 si modification de la période
                    vPeriode := GetPeriode(Q.Findfield('e_datecomptable').asdatetime);
                    Codeexo := QUELEXODTBUD(Q.Findfield('e_datecomptable').asdatetime);
                    suppEcriture;
                    Q.next;
               end;
           end
           else
                    suppEcriture;
           Ferme (Q);
       end
       else
                suppEcriture;
  end
  else
  begin
        if (Etablissement <> '') and (TWhereEtablissement <> '') then ETab := 'AND ('+TWhereEtablissement+')';
        if Datec <> '' then
           ExecuteSQL('DELETE FROM Ecriture WHERE E_EXERCICE="' + Codeexo + '"'
                              + ' AND E_JOURNAL="' + Journal + '" AND E_DATECOMPTABLE="' +
                              Datec + '"'+ ETab)
        else
           ExecuteSQL('DELETE FROM Ecriture WHERE E_EXERCICE="' + Codeexo + '"'
                              + ' AND E_JOURNAL="' + Journal + '" AND E_PERIODE=' +
                              IntToStr(vPeriode) + ETab);

        // suppression des écritures complémentaires
        ExecuteSQL('DELETE FROM ECRCOMPL WHERE EC_EXERCICE="' + Codeexo + '"'
                              + ' AND EC_JOURNAL="' + Journal + '" AND EC_DATECOMPTABLE="' + Datec + '"');

        if (Etablissement <> '') and (TWhereEtabY <> '') then ETab := 'AND ('+TWhereEtabY+')';
        if Datec <> '' then
           ExecuteSQL('DELETE FROM Analytiq WHERE Y_EXERCICE="' + Codeexo + '"'
                              + ' AND Y_JOURNAL="' + Journal + '" AND Y_DATECOMPTABLE="' +
                              Datec + '"'+ ETab)
        else
           ExecuteSQL('DELETE FROM Analytiq WHERE Y_EXERCICE="' + Codeexo + '"'
                              + ' AND Y_JOURNAL="' + Journal + '" AND Y_PERIODE=' +
                              IntToStr(vPeriode) + ETab);
  end;
end;

procedure TEvtSup.DegageExercice;
begin
     ExecuteSQL('DELETE FROM EXERCICE');
end;

// fiche 10300
Function TImportCom.TraiteJockerCorrespCpt(CRTYPE, CptInput : string; Var CptOut : string) : Boolean ;
var
Q1             : TQuery;
CptEntree      : string;
CptSortie      : string;
Pi             : PChar;
ipos, ii       : integer;
S2             : string;
begin
  CptOut := CptInput; Result := FALSE;

  Q1 := OpenSql('Select CR_CORRESP,CR_LIBELLE,CR_ABREGE from CORRESP where CR_TYPE="'+CRTYPE+'" and '+
  'CR_CORRESP like "%\%"ESCAPE"\" or CR_CORRESP like "%$%" ', TRUE);
  While not Q1.EOF do
  begin
       CptEntree := Q1.Findfield('CR_CORRESP').asstring;
       S2 := '';
      if (CptEntree[1] = '%') or (CptEntree[1] = '$') or (CptEntree[1] = CptInput[1]) then
       begin
               CptSortie := Q1.Findfield('CR_LIBELLE').asstring;
               ipos := pos('%', CptEntree);
               if ipos > 0 then
               begin
                      if (ipos <> 1) and (Copy (CptInput, 0, ipos-1) = Copy (CptEntree, 0, ipos-1)) then
                      begin
                           S2 := CptInput;
                           break;
                      end
                      else
                      begin
                                 Pi := PChar(CptEntree);
                                 while Pi^ <> #0 do
                                 begin
                                      if Pi^ <> '%' then
                                         S2 := S2 + Pi^;
                                    Inc(Pi);
                                 end;
                                 if pos(S2, CptInput) <> 0 then
                                 begin
                                      S2 := CptInput; break;
                                 end;
                      end;
               end
               else
               begin
                    if pos('$', CptEntree) <> 0 then
                    begin
                                 Pi := PChar(CptEntree);
                                 ii := 1;
                                 while Pi^ <> #0 do
                                 begin
                                      if Pi^ <> '$' then
                                         S2 := S2 + Pi^
                                      else
                                         S2 := S2 + CptInput[ii];
                                    Inc(Pi); inc (ii);
                                 end;
                                 if S2 = CptInput then break;
                    end;
               end;
       end;
       Q1.next;

  end;
  if S2 = CptInput then
  begin
         CptOut := CptSortie;
         Result := TRUE;
  end;
end;


Function TImportCom.TraiteCorrespCpt(Quoi : Integer ; CptInput : string; Var CptOut : string) : Boolean ;
Var
    CRType  : String ;
    SQL,FbCb: String ;
    LeFb    : TFichierBase ;
    Q1      : TQuery;
BEGIN
Result := FALSE; CRTYPE := '';
if FCorresp = '' then exit;
LeFb := FbGene;
Case Quoi Of
  0 : BEGIN
        if (pos('IGE', FCorresp) <> 0) then
        begin
         CRTYPE:='IGE' ; LeFb:=fbGene ;
         FbCb := fbGeneCb;
        end
      END ;
  1 : BEGIN
        if (pos('IAU', FCorresp) <> 0) then
        begin
            CRTYPE:='IAU' ; LeFb:=fbAux ;
            FbCb := fbAuxCb;
        end;
      END ;
  2 : BEGIN
        if (pos('IA1', FCorresp) <> 0) then
        begin
          CRTYPE:='IA1' ; LeFb:=fbAxe1 ;
          FbCb := fbAxe1Cb;
        end;
      END ;
  3 : BEGIN
        if (pos('IA2', FCorresp) <> 0) then
        begin
          CRTYPE:='IA2' ; LeFb:=fbAxe2 ;
          FbCb := fbAxe2Cb;
        end;
      END ;
  4 : BEGIN
        if (pos('IA3', FCorresp) <> 0) then
        begin
          CRTYPE:='IA3' ; LeFb:=fbAxe3 ;
          FbCb := fbAxe3Cb;
        end;
      END ;
  5 : BEGIN
        if (pos('IA4', FCorresp) <> 0) then
        begin
          CRTYPE:='IA4' ; LeFb:=fbAxe4 ;
          FbCb := fbAxe4Cb;
        end;
      END ;
  6 : BEGIN
        if (pos('IA5', FCorresp) <> 0) then
        begin
          CRTYPE:='IA5' ; LeFb:=fbAxe5 ;
          FbCb := fbAxe5Cb;
        end;
      END ;
  7 : BEGIN
        if (pos('IPM', FCorresp) <> 0) then
           CRTYPE:='IPM' ;
      END ;
  8 : BEGIN
        if (pos('IJA', FCorresp) <> 0) then
           CRTYPE:='IJA' ;
      END ;
  9 : BEGIN       // Etablissement
        if (pos('IET', FCorresp) <> 0) then
           CRTYPE:='IET'
      END ;
  END ;
  if CRTYPE = '' then exit;

  SQL := 'Select CR_CORRESP,CR_LIBELLE,CR_ABREGE from CORRESP where CR_TYPE="'+CRTYPE+'" and CR_CORRESP="'
              + CptInput +'"';
  Q1 := OpenSql(SQL, TRUE);
  if not Q1.EOF then
  begin
       CptOut := Q1.Findfield('CR_LIBELLE').asstring;
       Result := TRUE;
  end;
  Ferme (Q1);
  If Not Result Then
  begin
       Result := TraiteJockerCorrespCpt (CRTYPE, CptInput, CptOut);
       if not Result then Exit ;
  end;
  Case Quoi Of
    0..6 : CptOut :=BourreOuTronque(CptOut,LeFb) ;
    END ;
  Result:=TRUE ;
END ;



{***********A.G.L.Privé.*****************************************
Auteur  ...... : Module e traitement d'export
Créé le ...... : 12/05/2005
Modifié le ... :   /  /
Description .. : Création de modepaiement s'il n'existe pas
Mots clefs ... :
*****************************************************************}

procedure TImportCom.CreationModePaieInexistant;
var
TMdp      : TOB;
Mdp       : string;
i         : integer;
MPGene    : string;
MPCate    : string;
MPAccept  : string;
Q         : TQuery;
begin
    if SModepaie = nil then exit;
    if SModepaie.Count > 0  then
    begin
        Q := OpenSql ('SELECT ##TOP 1## * FROM MODEPAIE ', TRUE);
        if not Q.EOF then
        begin
          MPGene := Q.FindField ('MP_GENERAL').asstring;
          MPCate := Q.FindField ('MP_CATEGORIE').asstring;
          MPAccept :=  Q.FindField ('MP_CODEACCEPT').asstring;
        end;
        Ferme(Q);
    end;

    // Création de modepaiement s'il n'existe pas
    for i := 0 to SModepaie.Count-1 do
    begin
          Mdp := SModepaie[i];
          TMdp := TOB.Create('MODEPAIE', nil, -1);
          TMdp.PutValue ('MP_MODEPAIE', Mdp);
          TMdp.PutValue ('MP_LIBELLE', Mdp);
          TMdp.PutValue('MP_ABREGE', Mdp);
          TMdp.PutValue('MP_POINTABLE','-');
          TMdp.PutValue('MP_GENERAL', MPGene);
          TMdp.PutValue('MP_CATEGORIE', MPCate);
          TMdp.PutValue('MP_CODEACCEPT', MPAccept);
          if Not TMdp.SelectDB('"'+Mdp+'"', Nil) then
               TMdp.InsertDB(nil, TRUE);
          TMdp.Free;
    end;
    for i:=0 to SModepaie.Count-1 do If SModepaie.Objects[i]<>NIL Then SModepaie.Objects[i].Free ;
    SModepaie.free;  SModepaie := nil;
end;


procedure TImportCom.ImportCpteBudget(Ligne : string);
var
  TC          : TOB;
  comptebud   : string;
  Existegene  : Boolean;
begin
        comptebud := BourreOuTronque(EnleveBlanc(Ligne, 7, 17), fbBudGen);
        Existegene := ExisteSQL ('SELECT BG_BUDGENE FROM BUDGENE WHERE BG_BUDGENE="' + comptebud + '"');
        if (ExisteGene and BMAJCPTE) or not ExisteGene then
        begin
                TC := TOB.Create('BUDGENE', nil, -1);
                TC.PutValue('BG_BUDGENE', comptebud);
                TC.PutValue('BG_LIBELLE', EnleveBlanc(Ligne, 24, 35));
                TC.PutValue('BG_ABGREGE', EnleveBlanc(Ligne, 24, 17));
                TC.PutValue('BG_SENS', EnleveBlanc(Ligne, 59, 3));
                TC.PutValue('BG_SIGNE', EnleveBlanc(Ligne, 62, 3));
                TC.PutValue('BG_ATTENTE', EnleveBlanc(Ligne, 65, 1));
                TC.PutValue('BG_BLOQUANT', EnleveBlanc(Ligne, 66, 3));

                TC.PutValue('BG_DATECREATION', StrToDate(EnleveBlanc(Ligne, 70, 10, stDate1900)));
                TC.PutValue('BG_DATEOUVERTURE', StrToDate(EnleveBlanc(Ligne, 80, 10, stDate1900)));
                TC.PutValue('BG_FERME', EnleveBlanc(Ligne, 90, 1));
                TC.PutValue('BG_DATEFERMETURE', StrToDate(EnleveBlanc(Ligne, 91, 10, stDate1900)));
                TC.PutValue('BG_REPORTDISPO', EnleveBlanc(Ligne, 101, 3));

                TC.PutValue('BG_TABLE0', Veriftablelibre ('B00', EnleveBlanc(Ligne, 104, 17)));
                TC.PutValue('BG_TABLE1', Veriftablelibre ('B01', EnleveBlanc(Ligne, 121, 17)));
                TC.PutValue('BG_TABLE2', Veriftablelibre ('B02', EnleveBlanc(Ligne, 138, 17)));
                TC.PutValue('BG_TABLE3', Veriftablelibre ('B03', EnleveBlanc(Ligne, 155, 17)));
                TC.PutValue('BG_TABLE4', Veriftablelibre ('B04', EnleveBlanc(Ligne, 172, 17)));
                TC.PutValue('BG_TABLE5', Veriftablelibre ('B05', EnleveBlanc(Ligne, 189, 17)));
                TC.PutValue('BG_TABLE6', Veriftablelibre ('B06', EnleveBlanc(Ligne, 206, 17)));
                TC.PutValue('BG_TABLE7', Veriftablelibre ('B07', EnleveBlanc(Ligne, 223, 17)));
                TC.PutValue('BG_TABLE8', Veriftablelibre ('B08', EnleveBlanc(Ligne, 240, 17)));
                TC.PutValue('BG_TABLE9', Veriftablelibre ('B09', EnleveBlanc(Ligne, 257, 17)));

                TC.PutValue('BG_COMPTERUB', EnleveBlanc(Ligne, 274, 250));
                TC.PutValue('BG_EXCLURUB', EnleveBlanc(Ligne, 524, 250));
                TC.InsertOrUpdateDB(TRUE);
                TC.Free;
        end;
end;

procedure TImportCom.ImportJalBudget(Ligne : string);
var
  TC          : TOB;
  journalbud   : string;
  ExisteJal  : Boolean;
begin
        journalbud := EnleveBlanc(Ligne, 7, 3);
        ExisteJal := ExisteSQL ('SELECT BJ_BUDJAL FROM BUDJAL WHERE BJ_BUDJAL="' + journalbud + '"');
        if (ExisteJal and BMAJCPTE) or not ExisteJAl then
        begin
                TC := TOB.Create('BUDJAL', nil, -1);
                TC.PutValue('BJ_BUDJAL', journalbud);
                TC.PutValue('BJ_LIBELLE', EnleveBlanc(Ligne, 10, 35));
                TC.PutValue('BJ_ABGREGE', EnleveBlanc(Ligne, 10, 17));
                TC.PutValue('BJ_AXE', EnleveBlanc(Ligne, 45, 3));
                TC.PutValue('BJ_DATECREATION', Format_Date_HAL(EnleveBlanc(Ligne, 48, 8)));
                TC.PutValue('BJ_DATEOUVERTURE', Format_Date_HAL(EnleveBlanc(Ligne, 58, 8)));
                TC.PutValue('BJ_DATEFERMETURE', Format_Date_HAL(EnleveBlanc(Ligne, 68, 8)));

                TC.PutValue('BJ_EXODEB', EnleveBlanc(Ligne, 78, 3));
                TC.PutValue('BJ_EXOFIN', EnleveBlanc(Ligne, 81, 3));
                TC.PutValue('BJ_COMPTEURNORMAL', EnleveBlanc(Ligne, 84, 3));
                TC.PutValue('BJ_COMPTEURSIMUL', EnleveBlanc(Ligne, 87, 3));
                TC.PutValue('BJ_FERME', EnleveBlanc(Ligne, 90, 1));
                TC.PutValue('BJ_NATJAL', EnleveBlanc(Ligne, 104, 17));

                TC.PutValue('BJ_PERDEB', Format_Date_HAL(EnleveBlanc(Ligne, 94, 8)));
                TC.PutValue('BJ_PERFIN', Format_Date_HAL(EnleveBlanc(Ligne, 104, 8)));
                TC.PutValue('BJ_SENS', EnleveBlanc(Ligne, 114, 3));

                TC.PutValue('BJ_CATEGORIE', EnleveBlanc(Ligne, 117, 3));
                TC.PutValue('BJ_SOUSPLAN', EnleveBlanc(Ligne, 120, 35));
                TC.PutValue('BJ_BUDGENES', EnleveBlanc(Ligne, 155, 1024));
                TC.PutValue('BJ_BUDGENES2', EnleveBlanc(Ligne, 1179, 1024));
                TC.PutValue('BJ_BUDSECTS', EnleveBlanc(Ligne, 2203, 1024));
                TC.PutValue('BJ_BUDSECTS2', EnleveBlanc(Ligne, 3227, 1024));
                TC.InsertOrUpdateDB(TRUE);
                TC.Free;
        end;
end;

procedure TImportCom.ImportSectionBudget(Ligne : string);
var
  TC          : TOB;
begin
        TC := TOB.Create('BUDSECT', nil, -1);
        TC.PutValue('BS_BUDSECT', EnleveBlanc(Ligne, 7, 17));
        TC.PutValue('BS_LIBELLE', EnleveBlanc(Ligne, 27, 35));
        TC.PutValue('BS_ABREGE', EnleveBlanc(Ligne, 62, 17));
        TC.PutValue('BS_AXE', EnleveBlanc(Ligne, 24, 3));
        TC.PutValue('BS_DATECREATION', StrToDate(EnleveBlanc(Ligne, 85, 10, stDate1900)));
        TC.PutValue('BS_DATEOUVERTURE', StrToDate(EnleveBlanc(Ligne, 95, 10, stDate1900)));
        TC.PutValue('BS_DATEFERMETURE', StrToDate(EnleveBlanc(Ligne, 106, 10, stDate1900)));

        TC.PutValue('BS_FERME', EnleveBlanc(Ligne, 105, 1));
        TC.PutValue('BS_CREERPAR', EnleveBlanc(Ligne, 116, 3));
        TC.PutValue('BS_REPORTDISPO', EnleveBlanc(Ligne, 119, 3));
        TC.PutValue('BS_ATTENTE', EnleveBlanc(Ligne, 122, 1));
        TC.PutValue('BS_SENS', EnleveBlanc(Ligne, 123, 1));
        TC.PutValue('BS_SIGNE', EnleveBlanc(Ligne, 124, 1));

        TC.PutValue('BS_TABLE0', Veriftablelibre ('D00', EnleveBlanc(Ligne, 125, 17)));
        TC.PutValue('BS_TABLE1', Veriftablelibre ('D01', EnleveBlanc(Ligne, 142, 17)));
        TC.PutValue('BS_TABLE2', Veriftablelibre ('D02', EnleveBlanc(Ligne, 159, 17)));
        TC.PutValue('BS_TABLE3', Veriftablelibre ('D03', EnleveBlanc(Ligne, 176, 17)));

        TC.PutValue('BS_TABLE4', Veriftablelibre ('D04', EnleveBlanc(Ligne, 193, 17)));
        TC.PutValue('BS_TABLE5', Veriftablelibre ('D05', EnleveBlanc(Ligne, 210, 17)));
        TC.PutValue('BS_TABLE6', Veriftablelibre ('D06', EnleveBlanc(Ligne, 227, 17)));

        TC.PutValue('BS_TABLE7', Veriftablelibre ('D07', EnleveBlanc(Ligne, 244, 17)));
        TC.PutValue('BS_TABLE8', Veriftablelibre ('D08', EnleveBlanc(Ligne, 261, 17)));
        TC.PutValue('BS_TABLE9', Veriftablelibre ('D09', EnleveBlanc(Ligne, 278, 17)));

        TC.PutValue('BS_SECTIONRUB', EnleveBlanc(Ligne, 295, 250));
        TC.PutValue('BS_EXCURUB', EnleveBlanc(Ligne, 545, 250));
        TC.InsertOrUpdateDB(TRUE);
        TC.Free;
end;

procedure TImportCom.ImportChoixCode(Ligne : string);
var
  TC          : TOB;
  Code        : string;
begin
        If not ChargeChoixCod then
        begin
              if not ExisteSQL ('SELECT * FROM CHOIXCOD WHERE CC_TYPE="NAT"') then InitParamTablesLibres (''); 
              ChargeChoixCod := TRUE;
        end;
        Code := EnleveBlanc(Ligne, 7, 3);
        TC := TOB.Create('CHOIXCOD', nil, -1);
        TC.PutValue('CC_TYPE', Code);
        TC.PutValue('CC_CODE', EnleveBlanc(Ligne, 10, 3));
        TC.PutValue('CC_LIBELLE', EnleveBlanc(Ligne, 27, 35));
        TC.PutValue('CC_ABREGE', EnleveBlanc(Ligne, 97, 17));
        TC.PutValue('CC_LIBRE', EnleveBlanc(Ligne, 114, 70));
        TC.SetAllModifie(True);
        TC.InsertOrUpdateDB(TRUE);
        TC.Free;
end;

procedure TImportCom.ImportChoixExt(Ligne : string);
var
  TC          : TOB;
  Code        : string;
begin
        Code := EnleveBlanc(Ligne, 7, 3);
        TC := TOB.Create('CHOIXEXT', nil, -1);
        TC.PutValue('YX_TYPE', Code);
        TC.PutValue('YX_CODE', EnleveBlanc(Ligne, 10, 17));
        TC.PutValue('YX_LIBELLE', EnleveBlanc(Ligne, 27, 70));
        TC.PutValue('YX_ABREGE', EnleveBlanc(Ligne, 97, 17));
        TC.PutValue('YX_LIBRE', EnleveBlanc(Ligne, 114, 35));
        TC.SetAllModifie(True);
        TC.InsertOrUpdateDB(TRUE);
        TC.Free;
end;

procedure TImportCom.ImportRessource(Ligne : string);
var
  TC                    : TOB;
  Code,Ressource,Salarie: string;
  Q                     : Tquery;
begin
        Salarie := '';
        Code := EnleveBlanc(Ligne, 7, 17);
        Ressource := EnleveBlanc(Ligne, 24, 17);
        Q := OpenSql ('SELECT * FROM RESSOURCE '+
        ' Where ARS_TYPERESSOURCE="'+Code+'" AND  ARS_RESSOURCE="'+ Ressource +'"', TRUE);
        if not Q.EOF then
        begin
              TC := TOB.Create('',nil, -1);
              Salarie := Q.FindField ('ARS_SALARIE').asstring;
              TC.LoadDetailDB('RESSOURCE', '', '', Q, TRUE, FALSE);
              Ferme(Q);
              TC.detail[0].PutValue('ARS_TYPERESSOURCE', Code);
              TC.detail[0].PutValue('ARS_RESSOURCE', Ressource);
              TC.detail[0].PutValue('ARS_LIBELLE', EnleveBlanc(Ligne, 41, 35));
              TC.detail[0].PutValue('ARS_LIBELLE2', EnleveBlanc(Ligne, 76, 35));
              TC.detail[0].PutValue('ARS_CREERPAR', 'COM');
        end
        else
        begin
              TC := TOB.Create('RESSOURCE',nil, -1);
              TC.PutValue('ARS_TYPERESSOURCE', Code);
              TC.PutValue('ARS_RESSOURCE', Ressource);
              TC.PutValue('ARS_LIBELLE', EnleveBlanc(Ligne, 41, 35));
              TC.PutValue('ARS_LIBELLE2', EnleveBlanc(Ligne, 76, 35));
              TC.PutValue('ARS_CREERPAR', 'COM');
        end;
        if (Salarie = '') then
        begin
             TC.SetAllModifie(True);
             TC.InsertOrUpdateDB;
        end;
        TC.Free;
end;


procedure TImportCom.ImportParamlib(Ligne : string);
var
Code : string;
    procedure InsertEnreg (Ext : string; deb, fin : integer);
    var
    ie : integer;
    TC : TOB;
    begin
       for ie := deb to fin do
       begin
               if EnleveBlanc(Ligne, (RecPramlib[ie].deb), 35,'') <> '' then
               begin
                              TC := TOB.Create('PARAMLIB', nil, -1);
                              TC.PutValue('PL_TABLE', Code);
                              TC.PutValue('PL_CHAMP', Ext + RecPramlib[ie].NomChamp);
                              TC.PutValue('PL_LIBELLE', EnleveBlanc(Ligne, RecPramlib[ie].deb, 35));
                              TC.PutValue('PL_VISIBLE', EnleveBlanc(Ligne, RecPramlib[ie].deb+35, 1, '-'));
                              TC.InsertOrUpdateDB(TRUE);
                              TC.free;
               end;
       end;
    end;
begin
     Code := EnleveBlanc(Ligne, 7, 3);
     if (Code <> 'Y') and (Code <> 'E') and (Code <> 'BE') then
          InsertEnreg ('NT_', 0, 21)
     else
     begin
          if (Code = 'Y') or (Code = 'E') then
          begin
               InsertEnreg (Code+'_LIBRE', 0, 9);
               InsertEnreg(Code+'_', 22, 32);
          end
          else
          if Code = 'BE' then
             InsertEnreg(Code+'_', 29, 37);
     end;
     OKexercice := OnAfficheListeCom('Table Libre personnalisée : '+ Code,LISTEIMPORT);
end;

procedure TImportCom.CreatJournal (CodeJournal : string);
var
TJ : TOB;
begin
        TJ := TOB.Create ('JOURNAL',nil,-1);
        TJ.PutValue('J_JOURNAL', CodeJournal);
        TJ.PutValue('J_LIBELLE', CodeJournal+' CREATION PAR COM');
        TJ.PutValue('J_ABREGE', CodeJournal);
        TJ.PutValue('J_NATUREJAL', 'OD');
        TJ.PutValue('J_COMPTEURNORMAL', CodeJournal);
        TJ.PutValue('J_CONTREPARTIE', '');
        TJ.PutValue('J_MODESAISIE', 'LIB');
        TJ.putvalue ('J_CENTRALISABLE',  '-');
        TJ.putvalue ('J_MULTIDEVISE', 'X');
        TJ.AddChampSupValeur('J_COMPTEURSIMUL', 'SIM');
        TJ.putvalue ('J_CREERPAR', 'COM');
        TJ.InsertDB(nil);
        TJ.Free;
        TJ := TOB.Create ('SOUCHE', nil, -1);
        TJ.PutValue('SH_TYPE', 'CPT');
        TJ.PutValue('SH_SOUCHE', CodeJournal);
        TJ.PutValue('SH_LIBELLE', 'SOUCHE :'+ CodeJournal);
        TJ.PutValue('SH_NUMDEPART', 1);
        TJ.InsertOrUpdateDB(TRUE);
        TJ.Free;
end;

procedure TImportCom.SuppressionDesdoublons(Var TLecpiece, TlecNew :TOB; Journal : string; ii : integer);
var
ie                                 : integer;
TLec1                              : TOB;
begin
               if TLecpiece = nil then exit;
               ie := 0;
               while ie < TLecpiece.Detail.Count  do
               begin
                    TLec1 := TLecpiece.detail[ie];
                     // si pas de doublon mais intégration d'autres pièces
                     if (BIntegre and  (TLecpiece.Detail[ie].GetValue('OKDOUBLON')= 'X')) or
                     ((not AutoriseEcheck) (*and (pos('IGE;IAU', InterdireLaCreation) <> 0)*) and  (TLecpiece.Detail[ie].GetValue('OKDOUBLON')= 'I')) then
                     begin
                          if TlecNew = nil then
                             TlecNew := TOB.create ('', nil, -1);

                          if (BIntegre and  (TLecpiece.Detail[ie].GetValue('OKDOUBLON')= 'X')) then
                             OnAfficheListeCom('Doublon, Journal : '+ Journal+' Période : ' +TLecpiece.GetValue('FPERIODE') + ', Pièce d''origine : ' + TLecpiece.Detail[ie].GetValue('NUMFOLIO'),LISTEIMPORT);
                          if (TLecpiece.Detail[ie].GetValue('OKDOUBLON')= 'I') then
                             OnAfficheListeCom('Incohérence, Journal : '+ Journal+' Période : ' +TLecpiece.GetValue('FPERIODE') + ', Pièce d''origine : ' + TLecpiece.Detail[ie].GetValue('NUMFOLIO') + ' n''est pas importée',LISTEIMPORT);

                          TLec1.ChangeParent(TlecNew, -1);
                          if (NbEcriture > TLec1.detail.count) then
                             NbEcriture := NbEcriture- TLec1.detail.count;
                     end
                     else inc(ie);
               end;
end;

procedure TImportCom.ImporteRelance (Ligne : string);
var
TR : TOB;
begin
      TR := TOB.Create ('RELANCE',nil,-1);
      TR.PutValue ('RR_TYPERELANCE', EnleveBlanc(Ligne, 7, 3));
      TR.PutValue ('RR_FAMILLERELANCE', EnleveBlanc(Ligne, 10, 3));
      TR.PutValue ('RR_LIBELLE', EnleveBlanc(Ligne, 13, 35));
      TR.PutValue ('RR_GROUPELETTRE', EnleveBlanc(Ligne, 48, 1));
      TR.PutValue ('RR_NONECHU', EnleveBlanc(Ligne, 49, 1));
      TR.PutValue ('RR_SCOORING', EnleveBlanc(Ligne, 50, 1));
      TR.PutValue ('RR_DELAI1', EnleveBlanc(Ligne, 51, 5, '0'));
      TR.PutValue ('RR_MODELE1', EnleveBlanc(Ligne, 56, 3));
      TR.PutValue ('RR_DELAI2', EnleveBlanc(Ligne, 59, 5, '0'));
      TR.PutValue ('RR_MODELE2', EnleveBlanc(Ligne, 64, 3));
      TR.PutValue ('RR_DELAI3', EnleveBlanc(Ligne, 67, 5, '0'));
      TR.PutValue ('RR_MODELE3', EnleveBlanc(Ligne, 72, 3));
      TR.PutValue ('RR_DELAI4', EnleveBlanc(Ligne, 75, 5, '0'));
      TR.PutValue ('RR_MODELE4', EnleveBlanc(Ligne, 80, 3));
      TR.PutValue ('RR_DELAI5', EnleveBlanc(Ligne, 83, 5, '0'));
      TR.PutValue ('RR_MODELE5', EnleveBlanc(Ligne, 88, 3));
      TR.PutValue ('RR_DELAI6', EnleveBlanc(Ligne, 91, 5, '0'));
      TR.PutValue ('RR_MODELE6', EnleveBlanc(Ligne, 96, 3));
      TR.PutValue ('RR_DELAI7', EnleveBlanc(Ligne, 99, 5, '0'));
      TR.PutValue ('RR_MODELE7', EnleveBlanc(Ligne, 104, 3));
      TR.PutValue ('RR_ENJOURS', EnleveBlanc(Ligne, 107, 1, '-'));
      if Not TR.SelectDB('"'+EnleveBlanc(Ligne, 7, 3)+'";"'+EnleveBlanc(Ligne, 10, 3)+'"', Nil) then
       TR.InsertDB(nil, TRUE);
      TR.free;
end;

{$IFDEF EAGLSERVER}
(*
function TImportCom.ImportBobEagl (Fich : string; var TB : TOB) : Boolean;
var
Ind        : integer;
TSql : TOB;
begin
        TobLoadFromBinFile(Fich, nil, TB);
        for Ind := 0 to Tb.detail.count-1 do
        begin
                 TSql := Tb.detail[Ind];
                 TSql.InsertOrUpdateDB(TRUE) ;
        end;
        Result := TRUE;
end;
*)
{$ENDIF}


Function TImportCom.ImportBobIMMO : Boolean;
var
    TIMO          : TOB;
    NumVersion : integer;
    NewFile    : string;
begin
   Result := FALSE;
   TIMO := TOB.Create('IMMOPGI',Nil,-1);
    try
      NewFile := ChangeFileExt (FichierEntree, '.BOB');
      RenameFile(FichierEntree, NewFile);
      if AglReadContentBob ( NewFile , TIMO, True ) then
      begin
                NumVersion := TIMO.GetValue('BOBVERSION');
                if NumVersion <> 7 then
                begin
                    AffMessage('Version Incompatible, import impossible');
                    TIMO.Free;
                    exit;
                end;
                SourisSablier;
                if AglImportBob ( NewFile ) then
                begin
                    SourisNormale;
                    OKexercice := TRUE;
                    OkCreatCompte := TRUE;
                    Result := TRUE;
                    OnAfficheListeCom ('Import du dossier des immoblisations terminé.', LISTEIMPORT);
                end
                else
                begin
                    SourisNormale;
                    AffMessage ( 'Erreur lors de l''importation du dossier IMMO');
                end;
      end;
(*
{$IFDEF EAGLSERVER}
              if ImportBobEagl (NewFile, TIMO) then
              begin
                  SourisNormale;
                  OKexercice := TRUE;
                  OkCreatCompte := TRUE;
                  Result := TRUE;
                  OnAfficheListeCom ('Import du dossier des immoblisations terminé.', LISTEIMPORT);
              end;
{$ENDIF}
*)
    finally
      TIMO.Free
    end;
end;

function TImportCom.AjouterFichierDansGed (Fichier, CodeTiers, CodeGed, Description, Auteur, Annee, Mois : String; var DocId: string): Variant;
var
   FileId           : string;
   TobDoc,TobDocGed : Tob;
   Erreur           : string;
begin
     // Par défaut, fichier non intégré en GED Cegid
     Result := False; FileId := '';
     if copy(Fichier, 1, 5)= '.\GED' then
        Fichier := RootPath + '\GED\'+ TRIM(ExtractFileName(Fichier))
     else
     if copy(Fichier, 1, 2)= '.\' then
        Fichier := RootPath + '\'+ TRIM(ExtractFileName(Fichier));
     // Insertion du document
     FileId := V_GedFiles.Import (Fichier);
     if FileId <> '' then
     begin
          DocId     := '';
          TobDoc    := Tob.Create('YDOCUMENTS', nil, -1);
          TobDocGed := Tob.Create('DPDOCUMENT', nil, -1);
          try
              TobDoc.LoadDb;
              TobDoc.PutValue('YDO_LIBELLEDOC',Copy(Description,1,70) );
              TobDoc.PutValue('YDO_AUTEUR',    Copy(Auteur,1,35)      );
              TobDoc.PutValue('YDO_ANNEE',     Copy(Annee,1,4)        );
              TobDoc.PutValue('YDO_MOIS',      Copy(Mois,1,2)         );

              TobDocGed.LoadDb;
              TobDocGed.PutValue('DPD_NODOSSIER', NoDossier);
              TobDocGed.PutValue('DPD_CODEGED', CodeGed);
              DocId := InsertDocumentGed (TobDoc, TobDocGed, FileId, Erreur);
          finally
              TobDoc.Free;
              TobDocGed.Free;
          end;

          // Si fichier non référencé, on l'enlève de la Ged
          if DocId= '' then
              V_GedFiles.Erase (FileId)
          else
              Result := True;
     end;
end;

procedure TImportCom.AJouterListeSoldes (Compte : string; var ST : TStringList);
begin
       if Compte = '' then exit;
       if not BRecalculS then exit;
       if BSoldeLigne then exit;
       // A PORTER sur 7XX
       if (OrigineFichier = 'SI') or (NatureFichier = 'DOS') then exit; // pour des fuites mémoires
       if ST = nil then
       begin
              ST :=  TStringList.Create;
              ST.Sorted:=TRUE ;
              ST.Duplicates := dupIgnore;
       end;
       ST.Add (Compte);
end;

Function TImportCom.RendMRDefaut : string;
begin
    if (Mrdefaut = '') then
          Result :=  GetParamSocSecur('SO_GCMODEREGLEDEFAUT', '', False)
    else
          Result := Mrdefaut;
end;

Function TImportCom.RendRegTvadefaut : string;
begin
     Result := 'CEE';
     if (Regdefaut = '') then
     begin
           if GetParamSocSecur('SO_REGIMEDEFAUT', False) <> '' then
              Result :=  GetParamSocSecur('SO_REGIMEDEFAUT', False);
     end
     else Result := Regdefaut;
end;

Function TImportCom.AutoriseLaCreation (Code, MMessage : string) : Boolean;
begin
    Result := TRUE;
    if (not AutoriseEcheck) and  (pos(Code, InterdireLaCreation) <> 0) then
    begin
             OnAfficheListeCom(MMessage,LISTEIMPORT);
             if not BIntegre then
              begin
                   OkCreatCompte := FALSE;
                   OKexercice    := TRUE;
              end;
              OKRejet := TRUE;
              Result := FALSE;
    end;
end;

// fiche 10330
Function TImportCom.Veriftablelibre (TypeNat,Valeur : string) : string;
var
Q   : TQuery;
Ind : integer;
begin
Result := '';
if Valeur = '' then exit;
if OrigineFichier <> 'SI' then
begin
     if STLibre = nil then
     begin
       STLibre :=  TStringList.Create;
       STLibre.Sorted:=TRUE ;
       STLibre.Duplicates := dupIgnore;
       Q := OpenSQl ('SELECT NT_TYPECPTE,NT_NATURE FROM NATCPTE where NT_TYPECPTE like "T%" or NT_TYPECPTE like "E%" or NT_TYPECPTE like "G%" '+
       'or NT_TYPECPTE like "S%" or NT_TYPECPTE like "B%" or NT_TYPECPTE like "D%"', TRUE);
       While not Q.EOF do
       begin
             STLibre.Add (Q.FindField ('NT_TYPECPTE').asstring+':'+Q.FindField ('NT_NATURE').asstring);
             Q.Next;
       end;
       Ferme (Q);
     end;
     if STLibre.count > 0 then
     begin
       Ind := STLibre.IndexOf (TypeNat+':'+Valeur);
       if (Ind >= 0) then
            Result := Valeur;
     end;
     //if ExisteSQL('SELECT NT_NATURE FROM NATCPTE Where NT_TYPECPTE="'+TypeNat+'"' + ' and NT_NATURE="'+Valeur+'"') then
     //   Result := Valeur
end
else
     Result := Valeur;

end;

procedure TImportCom.LibereHTable;
begin
         // Objet de stockage en mémoire
         if assigned(fHTiers) then
         begin
              fHTiers.Clear(True);
              FreeAndNil (fHTiers);
         end;
         if assigned(fHTiersInsert) then
         begin
              fHTiersInsert.Clear(True);
              FreeAndNil (fHTiersInsert);
         end;
         if assigned(fHGene) then
         begin
              fHGene.Clear(True);
              FreeAndNil (fHGene);
         end;
         if assigned(fHGeneInsert) then
         begin
              fHGeneInsert.Clear(True);
              FreeAndNil (fHGeneInsert);
         end;

end;

procedure TImportCom.ChargeSection;
var
Q : TQuery;
begin
       SSectionInbase :=  TStringList.Create;
       SSectionInbase.Sorted:=TRUE ;
       SSectionInbase.Duplicates := dupIgnore;
       Q := OpenSQl ('SELECT S_SECTION,S_AXE FROM SECTION ', TRUE);
       While not Q.EOF do
       begin
             SSectionInbase.Add (Q.FindField ('S_SECTION').asstring + ':'+ Q.FindField ('S_AXE').asstring);
             Q.Next;
       end;
       Ferme (Q);
       SSectionFerme :=  TStringList.Create; // fiche 10413
       SSectionFerme.Sorted:=TRUE ;
       SSectionFerme.Duplicates := dupIgnore;
       Q := OpenSQl ('SELECT S_SECTION,S_AXE FROM SECTION WHERE S_FERME="X"', TRUE);
       While not Q.EOF do
       begin
             SSectionFerme.Add (Q.FindField ('S_SECTION').asstring + ':'+ Q.FindField ('S_AXE').asstring);
             Q.Next;
       end;
       Ferme (Q);
end;

procedure TImportCom.AffectDocGuid ( TOBEcr : TOB; DocGuid : string='');
Var
    nn               : integer ;
    TOBCompl         : TOB ;
BEGIN
          // Prendre les axes ventilés
          for nn:=0 to TOBEcr.Detail.Count-1 do
          begin
            TOBCompl := TOBEcr.detail[nn];
            // ajout me car ECRCOMPL est au même niveau que axe
            if TOBCompl.FieldExists ('EC_CLEECR') then
                     TOBCompl.PutValue ('EC_DOCGUID', DocGuid);
          end;
END ;

procedure TImportCom.GenereEcartDechange(Var TLecpiece, TlecNew :TOB);
var
ie,iii,Numpiece,Numligne           : integer;
TLec1,TLec,TLecInsert,TSou      : TOB;
Mtsolde                            : double;
JournalEcart,JSouche               : string;
Q1                                 : TQuery;
TJ                                 : TOB;
begin
               if (not BEcartChange) then exit;
               if (NatureFichier <> 'JRL') and (not BEcartChange) then exit;

               JournalEcart := 'ECC';
               JSouche      := JournalEcart;
               Numpiece := 1; Numligne := 1;

               Q1 := OpenSQL ('SELECT J_JOURNAL,J_COMPTEURNORMAL,SH_NUMDEPART FROM JOURNAL LEFT JOIN SOUCHE ON J_COMPTEURNORMAL=SH_SOUCHE WHERE J_NATUREJAL="ECC"', TRUE);
               if not Q1.EOF then
               begin
                  JournalEcart := Q1.FindField ('J_JOURNAL').Asstring;
                  JSouche      := Q1.FindField('J_COMPTEURNORMAL').asstring;
                  Numpiece     := Q1.FindField ('SH_NUMDEPART').asinteger+1;
               end
               else
               begin
                      TJ := TOB.Create ('JOURNAL',nil,-1);
                      TJ.PutValue('J_JOURNAL', 'ECC');
                      TJ.PutValue('J_LIBELLE', 'Ecart de change CREATION PAR COM');
                      TJ.PutValue('J_ABREGE', 'Ecart de change');
                      TJ.PutValue('J_NATUREJAL', 'ECC');
                      TJ.PutValue('J_COMPTEURNORMAL', 'ECC');
                      TJ.PutValue('J_CONTREPARTIE', '');
                      TJ.PutValue('J_MODESAISIE', '-');
                      TJ.putvalue ('J_CENTRALISABLE',  '-');
                      TJ.putvalue ('J_MULTIDEVISE', '-');
                      TJ.AddChampSupValeur('J_COMPTEURSIMUL', 'SIM');
                      TJ.putvalue ('J_CREERPAR', 'COM');
                      TJ.InsertDB(nil);
                      TJ.Free;
                      TJ := TOB.Create ('SOUCHE', nil, -1);
                      TJ.PutValue('SH_TYPE', 'CPT');
                      TJ.PutValue('SH_SOUCHE', 'ECC');
                      TJ.PutValue('SH_LIBELLE', 'SOUCHE : ECC');
                      TJ.PutValue('SH_NUMDEPART', 1);
                      TJ.InsertOrUpdateDB(TRUE);
                      TJ.Free;
               end;
               Ferme (Q1);

               if TLecpiece = nil then exit;
               ie := 0;
               for iii := 0 to TLecpiece.detail.Count - 1 do
               begin
                    TLec := TLecpiece.detail[iii];
                     while ie < TLec.Detail.Count  do
                     begin
                          TLec1 := TLec.detail[ie];
                           if (TLec1.GetValue('E_DEBITDEV')= 0)
                           and  (TLec1.GetValue('E_CREDITDEV')= 0) then
                           begin
                                Mtsolde := TLec1.GetDouble ('E_DEBIT')- TLec1.GetDouble ('E_CREDIT');
                                TLec.PutValue('SOLDE', (TLec.GetDouble('SOLDE') - Mtsolde));
                                TLec1.PutValue ('E_JOURNAL', JournalEcart);
                                TLec1.PutValue ('E_NUMEROPIECE', Numpiece);
                                TLec1.PutValue ('E_NUMLIGNE', Numligne);
                                TLec1.PutValue ('E_TAUXDEV', 1);
                                TLec1.PutValue ('E_COTATION', 1);
                                TLec1.PutValue('E_MODESAISIE', '-');
                                TLec1.PutValue ('E_DEBITDEV', 0);
                                TLec1.PutValue ('E_CREDITDEV', 0);
                                TLec1.PutValue ('E_NATUREPIECE', 'ECC'); // Fiche 10437
                                 if TlecNew = nil then
                                   TlecNew := TOB.create ('', nil, -1);
                                TLec1.ChangeParent(TlecNew, -1);
                                TLecInsert := TOB.create('ECRITURE', TlecNew, -1);
                                TLecInsert.Dupliquer(TLec.detail[ie-1], TRUE, TRUE, TRUE);
                                TLecInsert.PutValue ('E_JOURNAL', JournalEcart);
                                TLecInsert.PutValue ('E_DATECOMPTABLE', TLec1.Getvalue ('E_DATECOMPTABLE'));
                                TLecInsert.PutValue ('E_NUMEROPIECE', Numpiece);
                                TLecInsert.PutValue ('E_DEBIT', TLec1.Getvalue ('E_CREDIT'));
                                TLecInsert.PutValue ('E_CREDIT', TLec1.Getvalue ('E_DEBIT'));
                                TLecInsert.PutValue ('E_DEBITDEV', 0);
                                TLecInsert.PutValue ('E_CREDITDEV', 0);
                                TLecInsert.PutValue ('E_NUMLIGNE', Numligne+1);
                                TLecInsert.PutValue ('E_TAUXDEV', 1);
                                TLecInsert.PutValue ('E_COTATION', 1);
                                TLecInsert.PutValue ('E_NATUREPIECE', 'ECC'); // Fiche 10437
                                TLecInsert.PutValue ('E_MODESAISIE', '-');
                                TLecInsert.PutValue ('E_DEVISE', TLec1.GetValue ('E_DEVISE'));
                                inc(Numligne);
                                if (NbEcriture > TLec1.detail.count) then
                                   NbEcriture := NbEcriture- TLec1.detail.count;
                           end
                           else inc(ie);
                     end;
               end;
               // mise à jour souche
              TSou := TOB.Create('SOUCHE', nil, -1);
              TSou.Putvalue ('SH_TYPE', 'CPT');
              TSou.Putvalue ('SH_SOUCHE', JSouche);
              if Not TSou.SelectDB('"CPT";"'+JSouche+'"', Nil) then
              begin
                    TSou.Putvalue ('SH_LIBELLE', JSouche);
                    TSou.Putvalue ('SH_NUMDEPART', 1);
                    TSou.Putvalue ('SH_SIMULATION', '-');
                    TSou.Putvalue ('SH_ANALYTIQUE', '-');
                    TSou.InsertDB(nil, TRUE);
              end
              else
              begin
                   TSou.Putvalue ('SH_NUMDEPART', Numpiece);
                   TSou.SetAllModifie (TRUE); TSou.UpdateDB(TRUE);
              end;
              if TlecNew <> nil then
                 MiseajourDetail (TlecNew,TRUE,FALSE,FALSE );
end;

procedure TImportCom.GestiondesBAP (Tlec : TOB);
var
  TPiece : TOB;
begin
  TPiece := Tlec.detail[0];
  if TPiece.Getstring ('E_QUALIFPIECE') <> 'N' then exit;

  if not Assigned(EnregBap) then
  begin
    OnAfficheListeCom('Bons à payer ', LISTEIMPORT);
    EnregBap := TBonAPayer.Create(TLec, TypeVisa)
  end
  else
    EnregBap.CreerBap(TLec);

end;

procedure TImportCom.ImporteBapVisa(Ligne: string);
var
TBap : TOB;
begin
  TBap := TOB.Create('CPTYPEVISA', nil, -1);
  TBap.Putvalue ('CTI_CODEVISA', EnleveBlanc(Ligne, 7, 3));
  TBap.Putvalue ('CTI_LIBELLE', VireTouteSub(EnleveBlanc(Ligne, 10, 35),'"'));
  TBap.Putvalue ('CTI_ABREGE', VireTouteSub(EnleveBlanc(Ligne, 45, 17),'"'));
  TBap.Putvalue ('CTI_NBVISA', StrToInt(EnleveBlanc(Ligne, 62, 5)));
  TBap.Putvalue ('CTI_NATUREPIECE', EnleveBlanc(Ligne, 67, 3));
  TBap.Putvalue ('CTI_MONTANTMIN', Valeur(StPoint(EnleveBlanc(Ligne, 70, 20))));
  TBap.Putvalue ('CTI_MONTANTMAX', Valeur(StPoint(EnleveBlanc(Ligne, 90, 20))));

  TBap.Putvalue ('CTI_ETABLISSEMENT', EnleveBlanc(Ligne, 110, 3));
  TBap.Putvalue ('CTI_TYPELIBRE1', EnleveBlanc(Ligne, 113, 3));
  TBap.Putvalue ('CTI_TYPELIBRE2', EnleveBlanc(Ligne, 116, 3));
  TBap.Putvalue ('CTI_TYPELIBRE3', EnleveBlanc(Ligne, 119, 3));
  TBap.Putvalue ('CTI_CODELIBRE1', EnleveBlanc(Ligne, 122, 3));
  TBap.Putvalue ('CTI_CODELIBRE2', EnleveBlanc(Ligne, 125, 3));
  TBap.Putvalue ('CTI_CODELIBRE3', EnleveBlanc(Ligne, 128, 3));
  TBap.Putvalue ('CTI_TEXTELIBRE1', EnleveBlanc(Ligne, 131, 35));
  TBap.Putvalue ('CTI_TEXTELIBRE2', EnleveBlanc(Ligne, 166, 35));
  TBap.Putvalue ('CTI_TEXTELIBRE3', EnleveBlanc(Ligne, 201, 35));
  TBap.Putvalue ('CTI_AXE1', EnleveBlanc(Ligne, 236, 3));
  TBap.Putvalue ('CTI_AXE2', EnleveBlanc(Ligne, 239, 3));
  TBap.Putvalue ('CTI_AXE3', EnleveBlanc(Ligne, 242, 3));
  TBap.Putvalue ('CTI_CIRCUITBAP', EnleveBlanc(Ligne, 245, 3));
  TBap.Putvalue ('CTI_COMPTE', EnleveBlanc(Ligne, 248, 250));
  TBap.Putvalue ('CTI_EXCLUSION', EnleveBlanc(Ligne, 498, 250));
  TBap.InsertOrUpdateDB(TRUE);
  TBap.free;
  OKexercice := OnAfficheListeCom('Types de visas des bons à payer : ' + EnleveBlanc(Ligne, 7, 3),LISTEIMPORT );

end;

procedure TImportCom.ImporteBapCircuit(Ligne: string);
var
TBap : TOB;
begin
  TBap := TOB.Create('CPCIRCUIT', nil, -1);
  TBap.Putvalue ('CCI_CIRCUITBAP', EnleveBlanc(Ligne, 7, 3));
  TBap.Putvalue ('CCI_NUMEROORDRE', StrToInt(EnleveBlanc(Ligne, 10, 5, '1')));
  TBap.Putvalue ('CCI_VISEUR1', EnleveBlanc(Ligne, 15, 3));
  TBap.Putvalue ('CCI_VISEUR2', EnleveBlanc(Ligne, 18, 3));
  TBap.Putvalue ('CCI_NBJOUR', StrToInt(EnleveBlanc(Ligne, 21, 5)));
  TBap.Putvalue ('CCI_LIBELLE', EnleveBlanc(Ligne, 26, 35));
  TBap.InsertOrUpdateDB(TRUE);
  TBap.free;
  OKexercice := OnAfficheListeCom('Circuit de validation des BAP : ' + EnleveBlanc(Ligne, 7, 3),LISTEIMPORT );
end;

procedure TImportCom.ImporteBonapayer(Ligne: string; exercice, Docid : string; numpiece : integer);
var
TBap : TOB;
begin
  TBap := TOB.Create('CPBONSAPAYER', nil, -1);
  TBap.Putvalue ('BAP_JOURNAL', EnleveBlanc(Ligne, 1, 3));
  TBap.Putvalue ('BAP_EXERCICE', exercice);
  TBap.Putvalue ('BAP_DATECOMPTABLE', Format_Date_HAL(EnleveBlanc(Ligne, 4, 8)));
  TBap.Putvalue ('BAP_NUMEROPIECE', numpiece);
  TBap.Putvalue ('BAP_VISEUR', EnleveBlanc(Ligne, 57, 3));
  TBap.Putvalue ('BAP_STATUTBAP', EnleveBlanc(Ligne, 60, 3));
  TBap.Putvalue ('BAP_VISEUR1', EnleveBlanc(Ligne, 63, 3));
  TBap.Putvalue ('BAP_VISEUR2', EnleveBlanc(Ligne, 66, 3));
  TBap.Putvalue ('BAP_NUMEROORDRE', StrToInt(EnleveBlanc(Ligne, 69, 8, '1')));
  TBap.Putvalue ('BAP_NBJOUR', StrToInt(EnleveBlanc(Ligne, 77, 5, '1')));
  TBap.Putvalue ('BAP_ECHEANCEBAP', Format_Date_HAL(EnleveBlanc(Ligne, 82, 8)));
  TBap.Putvalue ('BAP_TIERSPAYEUR', EnleveBlanc(Ligne, 90, 1));
  TBap.Putvalue ('BAP_CLEFFACTURE', EnleveBlanc(Ligne, 91, 35));
  TBap.Putvalue ('BAP_CODEVISA', EnleveBlanc(Ligne, 126, 3));
  TBap.Putvalue ('BAP_CIRCUITBAP', EnleveBlanc(Ligne, 129, 3));
  TBap.Putvalue ('BAP_DATEMAIL', Format_Date_HAL(EnleveBlanc(Ligne, 132, 8)));
  TBap.Putvalue ('BAP_DATERELANCE1', Format_Date_HAL(EnleveBlanc(Ligne, 140, 8)));
  TBap.Putvalue ('BAP_DATERELANCE2', Format_Date_HAL(EnleveBlanc(Ligne, 148, 8)));
  TBap.Putvalue ('BAP_DATEECHEANCE', Format_Date_HAL(EnleveBlanc(Ligne, 156, 8)));
  TBap.Putvalue ('BAP_SOCIETE', EnleveBlanc(Ligne, 164, 3));
  TBap.Putvalue ('BAP_ETABLISSEMENT', EnleveBlanc(Ligne, 167, 3));
  TBap.Putvalue ('BAP_ALERTE1', EnleveBlanc(Ligne, 170, 1));
  TBap.Putvalue ('BAP_ALERTE2', EnleveBlanc(Ligne, 171, 1));
  TBap.Putvalue ('BAP_RELANCE1', EnleveBlanc(Ligne, 172, 1));
  TBap.Putvalue ('BAP_RELANCE2', EnleveBlanc(Ligne, 173, 1));
  TBap.Putvalue ('BAP_IDGED', DocId);

  TBap.InsertOrUpdateDB(TRUE);
  TBap.free;
  OKexercice := OnAfficheListeCom('Bons à payer du journal : ' + EnleveBlanc(Ligne, 1, 3)+
   ' Période : ' + FormatDateTime (Traduitdateformat('mmmm yyyy'), Format_Date_HAL(EnleveBlanc(Ligne, 4, 8)))+
   ' Pièce : ' + IntTostr(numpiece) ,LISTEIMPORT );

end;

// fiche 10447
function TImportCom.TrouveRubriqueDuCpteGene(vStCpte: string; Liasse : Boolean=FALSE): string;
var
 lStSqlRub, St, C1, C2                : string;
 lLesComptesInclus, lLesComptesExclus : TStringList;
 Q                                    : TQuery;
 i                                    : integer;
    Function IlfautExlure (CC : string) : Boolean;
    var
    ii : integer;
    begin
          Result := False;
          if lLesComptesInclus.Count = 0 then exit;

          for ii := 0 to lLesComptesInclus.Count - 1 do
          begin
               St := lLesComptesInclus[ii];
               if (Pos ( ':', St ) > 0) then
               begin
                    C1 := ReadTokenPipe(St,':');
                    C2 := St;
                    if (Copy (CC, 1 , Length(C1))) = C1 then begin Result := TRUE; break; end
                    else
                    if (Copy (CC, 1 , Length(C2))) = C2 then begin Result := TRUE; break; end;
               end
               else
               begin
                     if (Copy (CC, 1 , Length(St))) = St then begin Result := TRUE; break; end;
               end;
          end;
    end;

begin

  if Liasse then
  lStSqlRub := 'SELECT RB_RUBRIQUE, RB_COMPTE1, RB_EXCLUSION1 FROM RUBRIQUE WHERE ' +
            'RB_FAMILLES LIKE "%@' + GetParamSocSecur('SO_CPCONTROLELIASSE', False) + ';%" ' +
            'ORDER BY RB_RUBRIQUE'
  else
  lStSqlRub := 'SELECT RB_RUBRIQUE, RB_COMPTE1, RB_EXCLUSION1 FROM RUBRIQUE WHERE ' +
               'RB_CLASSERUB = "CDR" AND RB_NATRUB = "CPT" AND ' +
               'RB_FAMILLES LIKE "%' + GetParamSocSecur('SO_CPPLANREVISION', '') + ';%" ' +
               'ORDER BY RB_RUBRIQUE';
  lLesComptesInclus := nil;
  lLesComptesExclus := nil;
  lLesComptesInclus := TStringList.Create;
  lLesComptesInclus.Duplicates := dupIgnore;
  lLesComptesInclus.Sorted := True;
  lLesComptesExclus := TStringList.Create;
  lLesComptesExclus.Duplicates := dupIgnore;
  lLesComptesExclus.Sorted := True;
  Result := '';
  Q := OpenSql( lStSqlRub, True);
  while not Q.Eof do
  begin
      if lLesComptesInclus <> nil then lLesComptesInclus.Clear;
      if lLesComptesExclus <> nil then lLesComptesExclus.Clear;
      AnalyseCompte(Q.FindField('RB_COMPTE1').AsString, fbGene, False, False, True, True, lLesComptesInclus);
      for i := 0 to lLesComptesInclus.Count - 1 do
      begin
          St := lLesComptesInclus[i];
          if Pos (':', St) > 0 then
          begin
               // Tranche de compte
               C1 := ReadTokenPipe(St,':');
               C2 := St;
          end
          else if Pos ('%', St) > 0 then
          begin
               // Like <=> tranche de compte XXX000000 à XXXZZZZZZ
               C1 := Copy (St , 1, Pos ('%', St) - 1);
               C2 := C1 + Copy ('99999999999999999', 1, GetInfoCpta (fbGene).Lg - Length (C1));
               C1 := C1 + Copy ('00000000000000000', 1, GetInfoCpta (fbGene).Lg - Length (C1));
          end
          else
          begin
               // Valeur <=> tranche de compte XXX à XXX (C1 = C2)
               C1 := St;
               C2 := St;
          end;
          if (vStCpte >= C1) and (vStCpte <=C2) and (not IlfautExlure (vStCpte)) then
          begin
                Result := Q.FindField('RB_RUBRIQUE').AsString;
                Break;
          end;
      end;
      if Result <> '' then break;
      Q.Next;
  end;
  if (lLesComptesInclus <> nil) then FreeAndNil (lLesComptesInclus);
  if (lLesComptesExclus <> nil) then FreeAndNil (lLesComptesExclus);

end;

// fiche 10477
procedure TImportCom.ImporteCorresp(Ligne: string);
var
TCorresp : TOB;
begin
  TCorresp := TOB.Create('CORRESP', nil, -1);
  TCorresp.Putvalue ('CR_CORRESP', EnleveBlanc(Ligne, 7, 17));
  TCorresp.Putvalue ('CR_LIBELLE', EnleveBlanc(Ligne, 24, 35));
  TCorresp.Putvalue ('CR_TYPE', EnleveBlanc(Ligne, 59, 3));
  TCorresp.InsertOrUpdateDB(TRUE);
  TCorresp.free;
  OKexercice := OnAfficheListeCom('Table de Correspondance : ' + EnleveBlanc(Ligne, 7, 17),LISTEIMPORT );
end;

procedure TImportCom.ImporteCorrespImp(Ligne: string);
var
TCorresp : TOB;
begin
  TCorresp := TOB.Create('CORRESP', nil, -1);
  TCorresp.Putvalue ('CR_CORRESP', EnleveBlanc(Ligne, 7, 17));
  TCorresp.Putvalue ('CR_LIBELLE', EnleveBlanc(Ligne, 24, 17));
  TCorresp.Putvalue ('CR_TYPE', EnleveBlanc(Ligne, 41, 3));
  TCorresp.InsertOrUpdateDB(TRUE);
  TCorresp.free;
  OKexercice := OnAfficheListeCom('Table de Correspondance : ' + EnleveBlanc(Ligne, 7, 17),LISTEIMPORT );
end;

procedure TImportCom.ImporteVentiltype (Ligne : string);
var
TVentil        :TOB;
Nature,Section : string;
begin
  TVentil := TOB.Create('VENTIL', nil, -1);
  nature := EnleveBlanc(Ligne, 7, 3);
  Section := EnleveBlanc(Ligne, 27, 17);
  TVentil.Putvalue ('V_NATURE', Nature);
  if nature[1] = 'G' then
     TVentil.Putvalue ('V_COMPTE', BourreOuTronque(EnleveBlanc(Ligne, 10, 17), FbGene))
  else
     TVentil.Putvalue ('V_COMPTE', EnleveBlanc(Ligne, 10, 17));
  if nature = 'TY1' then Section := BourreOuTronque(Section, Fbaxe1)
  else
  if nature = 'TY2' then Section := BourreOuTronque(Section, Fbaxe2)
  else
  if nature = 'TY3' then Section := BourreOuTronque(Section, Fbaxe3)
  else
  if nature = 'TY4' then Section := BourreOuTronque(Section, Fbaxe4)
  else
  if nature = 'TY5' then Section := BourreOuTronque(Section, Fbaxe5);

  // fiche 10605
  if nature = 'GE1' then Section := BourreOuTronque(Section, Fbaxe1)
  else
  if nature = 'GE2' then Section := BourreOuTronque(Section, Fbaxe2)
  else
  if nature = 'GE3' then Section := BourreOuTronque(Section, Fbaxe3)
  else
  if nature = 'GE4' then Section := BourreOuTronque(Section, Fbaxe4)
  else
  if nature = 'GE5' then Section := BourreOuTronque(Section, Fbaxe5);


  TVentil.Putvalue ('V_SECTION', Section);
  TVentil.Putvalue ('V_TAUXMONTANT', Valeur(StPoint(EnleveBlanc(Ligne, 44, 20))));
  TVentil.Putvalue ('V_TAUXQTE1', Valeur(StPoint(EnleveBlanc(Ligne, 64, 20))));
  TVentil.Putvalue ('V_TAUXQTE2', Valeur(StPoint(EnleveBlanc(Ligne, 84, 20))));
  TVentil.Putvalue ('V_NUMEROVENTIL', StrToInt(EnleveBlanc(Ligne, 104, 8)));
  TVentil.Putvalue ('V_SOCIETE', EnleveBlanc(Ligne, 112, 3));
  TVentil.Putvalue ('V_MONTANT', Valeur(StPoint(EnleveBlanc(Ligne, 115, 20))));
  TVentil.Putvalue ('V_SOUSPLAN1', EnleveBlanc(Ligne, 135, 17));
  TVentil.Putvalue ('V_SOUSPLAN2', EnleveBlanc(Ligne, 152, 17));
  TVentil.Putvalue ('V_SOUSPLAN3', EnleveBlanc(Ligne, 169, 17));
  TVentil.Putvalue ('V_SOUSPLAN4', EnleveBlanc(Ligne, 186, 17));
  TVentil.Putvalue ('V_SOUSPLAN5', EnleveBlanc(Ligne, 203, 17));
  TVentil.Putvalue ('V_SOUSPLAN6', EnleveBlanc(Ligne, 220, 17));
  TVentil.InsertOrUpdateDB(TRUE);
  TVentil.free;
  if EnleveBlanc(Ligne, 10, 17) <> VentilType then
  begin
       OKexercice := OnAfficheListeCom('Ventilation type : ' + EnleveBlanc(Ligne, 10, 17),LISTEIMPORT );
       VentilType := EnleveBlanc(Ligne, 10, 17)
  end;
end;

// BVE 28.08.07 : Suivi des validations
{$IFDEF CERTIFNF}
procedure TImportCom.ImporteSuiviValidation (Ligne : string);
var
TSuiviVal      :TOB;
begin
  TSuiviVal := TOB.Create('CPJALVALIDATION', nil, -1);
  TSuiviVal.Putvalue ('CPV_SESSION',      StrToInt(EnleveBlanc(Ligne, 7, 8)));
  TSuiviVal.Putvalue ('CPV_DATE',         StrToDate(EnleveBlanc(Ligne, 15, 8)));
  TSuiviVal.Putvalue ('CPV_UTILISATEUR',  EnleveBlanc(Ligne, 23, 3));
  TSuiviVal.Putvalue ('CPV_IDDEBUTVAL',   StrToInt(EnleveBlanc(Ligne, 26, 8)));
  TSuiviVal.Putvalue ('CPV_IDFINVAL',     StrToInt(EnleveBlanc(Ligne, 34, 8)));
  TSuiviVal.Putvalue ('CPV_NUMEVENT',     StrToInt(EnleveBlanc(Ligne, 42, 8)));
  TSuiviVal.InsertOrUpdateDB(TRUE);
  TSuiviVal.free;
end;
{$ENDIF}

// Fiche 10160
procedure TImportCom.AjoutListeDatedernMvt (TS : TList; Cpte : string; DateP : tDateTime);
var
i : integer;
T : T_DM ;
begin
 for i := 0 to TS.Count-1 do
  begin
    T := T_DM(TS[i]) ;
    if (T.Cpte=Cpte) then
    begin
     if T.DateP < DateP then
      T.DateP := DateP;
      exit;
    end;
  end;
  T := T_DM.Create ;
  T.Cpte  := Cpte;
  T.DateP := DateP;
  TS.Add(T) ;
end;

procedure TImportCom.MAJDateDernMvt;
var
i : integer;
T : T_DM ;
begin
  for i := 0 to lpGM.Count-1 do
  begin
    T := T_DM(lpGM[i]) ;
    ExecuteSQL('UPDATE GENERAUX SET G_DATEDERNMVT="' + UsDateTime(T.DateP) + '"' +
                ' WHERE G_GENERAL="'+BourreOuTronque(T.Cpte,FbGene)+'"');

  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 04/07/2007
Modifié le ... :   /  /    
Description .. : Permet de remplir la TOB Devise 
Mots clefs ... : 
*****************************************************************}
procedure TImportCom.RemplirTOBDevise;
begin
  if not(assigned(TDevise)) then
     TDevise := TOB.Create('DEVISE',nil,-1)
  else
     TDevise.ClearDetail;
  TDevise.LoadDetailDB('DEVISE','','',nil,false);
end;

function TImportCom.RendCodeAcceptation (MP : string) : string;
var
i   : integer;
St,MPLu, Acc : string;

begin
          for i:=0 to GetInfoAcceptation.Count-1 do
          BEGIN
          St:= GetInfoAcceptation[i] ; MPLu:=ReadtokenSt(St) ;
          if MPLu=MP then BEGIN Acc:=ReadtokenSt(St) ; Result:=Acc ; Break ; END ;
          END ;
end;

procedure  TImportCom.RemplirDateEchance (var T_Mode : T_ModeRegl; DateEcriture : TDateTime; T : TOB);
var
TA1  : TOB;
QD   : TQuery;
sSQL : string;
CompteA,CompteG : string;
TobM : TOB;
Q    : TQuery;
begin
        FillChar(T_Mode, Sizeof(T_Mode),#0) ;
        if (T.GetValue ('E_AUXILIAIRE') <> '') then CompteA := T.GetValue ('E_AUXILIAIRE')
        else CompteG := T.GetValue ('E_GENERAL');
        if CompteA <> '' then
        begin
                  TA1 := TTiersUpdate.FindFirst(['T_AUXILIAIRE'], [CompteA], FALSE);
                  if TA1 <> nil then
                      T_Mode.ModeInitial:= TA1.GetValue ('T_MODEREGLE')
                  else
                  begin
                    TA1 := RechercheOptimise(TTiersInsert, 'T_AUXILIAIRE', CompteA, FALSE, fHTiersInsert);
                    if TA1 <> nil then
                            T_Mode.ModeInitial:= TA1.GetValue ('T_MODEREGLE')
                    else
                    begin
                              TA1 := RechercheOptimise(TTiers, 'T_AUXILIAIRE', CompteA, FALSE, fHTiers);
                              if TA1 <> nil then
                                  T_Mode.ModeInitial:= TA1.GetValue ('T_MODEREGLE')
                              else
                              begin
                                        sSQL := 'SELECT T_MODEREGLE FROM TIERS '
                                             + ' Where  T_AUXILIAIRE="'+CompteA+ '"';
                                        QD := OpenSQL(sSQL, true);
                                        if not QD.EOF then
                                            T_Mode.ModeInitial:= QD.FindField ('T_MODEREGLE').asstring;
                                        Ferme (QD);
                              end;
                    end;
                  end;
        end
        else
        begin
              TA1 := TGenUpdate.FindFirst(['G_GENERAL'], [CompteG], FALSE);
              if TA1 <> nil then
                                        T_Mode.ModeInitial:= TA1.GetValue ('G_MODEREGLE')
              else
              begin
                   TA1 := RechercheOptimise(TGenInsert, 'G_GENERAL', CompteG, TRUE, fHGeneInsert);
                   if TA1 <> nil then
                                        T_Mode.ModeInitial:= TA1.GetValue ('G_MODEREGLE')

                   else
                   begin
                        TA1 := RechercheOptimise(TGen, 'G_GENERAL', CompteG, TRUE, fHGene);
                        if TA1 <> nil then
                                        T_Mode.ModeInitial:= TA1.GetValue ('G_MODEREGLE')
                        else   // ajout me 14-01-2005
                        begin
                                  QD := OpenSQL('SELECT G_MODEREGLE FROM GENERAUX WHERE G_GENERAL="'+CompteG+'"', true);
                                  if not QD.EOF then
                                      T_Mode.ModeInitial:=  QD.FindField ('G_MODEREGLE').asstring;
                                  Ferme(QD);
                        end;
                  end;
              end;
        end;
        T_Mode.ModeFinal  := T_Mode.ModeInitial ;
        T_Mode.DateFact := DateEcriture;
        T_Mode.DateFactExt := T.GetValue ('E_DATEREFEXTERNE');
        T_Mode.Action:=taCreat ;

        T_Mode.CodeDevise:=T.GetValue('E_DEVISE') ;
        if TDevise.FindFirst(['D_DEVISE'],[T.GetValue('E_DEVISE')],true) <> nil then
        begin
            T_Mode.Symbole := TDevise.getvalue ('D_SYMBOLE'); T_Mode.TauxDevise:= T.GetDouble('E_TAUXDEV') ;
            T_Mode.Quotite:=TDevise.getvalue ('D_QUOTITE'); T_Mode.Decimale:=TDevise.getvalue ('D_DECIMALE');
        end;
        T_Mode.DateFact := DateEcriture ; T_Mode.DateBL:=DateEcriture ;
        T_Mode.Aux      := CompteA ;
        TobM := nil;
        if TModeRegle = nil then
        begin
                 Q := OpenSQL('SELECT * FROM MODEREGL ORDER BY MR_MODEREGLE', TRUE);
                 TobM :=TOB.Create('TOBM',Nil,-1) ;
                 TobM.LoadDetailDB('MODEREGL', '', '', Q, TRUE, FALSE);
                 Ferme (Q);
        end;
        CalculModeRegle (T_Mode, FALSE, TobM);
        if not TransIsValidDate(DateToStr(T_Mode.TabEche[1].DateEche)) then
          T_Mode.TabEche[1].DateEche := DateEcriture;

end;

procedure TImportCom.SuppressionDesSoldesAZero(Var TLecpiece, TlecNew :TOB; Journal : string; ii : integer);
var
ie                                 : integer;
TLec1                              : TOB;
MtDev, MtDevRegul                  : double;
OkMove                             : Boolean;
begin
               if TLecpiece = nil then exit;
               ie := 0;
               while ie < TLecpiece.Detail.Count  do
               begin
                    TLec1 := TLecpiece.detail[ie];
                    if (TLec1.detail[0].GetValue('E_ECRANOUVEAU') <> 'OAN') and
                    (TLec1.detail[0].GetValue('E_ECRANOUVEAU') <> 'H')
                    and ((TLec1.getdouble('SOLDE') <> 0) or (TLec1.Getdouble ('SOLDEDEV') <> 0))then
                    begin
                          OkMove := TRUE;
                          // Fiche 10620
                          if TLec1.detail[0].GetValue('E_DEVISE') <> V_PGI.DevisePivot  then
                          begin
                               MtDev :=  (TLec1.detail[TLec1.Detail.Count-1].GetValue('E_DEBIT') - TLec1.detail[TLec1.Detail.Count-1].GetValue('E_CREDIT'));
                               TLec1.putValue('SOLDE', TLec1.getdouble('SOLDE')-MtDev);
                               MtDev :=  (TLec1.detail[TLec1.Detail.Count-1].GetValue('E_DEBITDEV') - TLec1.detail[TLec1.Detail.Count-1].GetValue('E_CREDITDEV'));
                               TLec1.putValue('SOLDEDEV', TLec1.getdouble('SOLDEDEV')-MtDev);
                               CEquilibrePiece(TLec1);
                               MtDevRegul := TLec1.detail[TLec1.Detail.Count-1].GetValue('E_DEBIT') - TLec1.detail[TLec1.Detail.Count-1].GetValue('E_CREDIT');
                               TLec1.putValue('SOLDE', TLec1.getdouble('SOLDE')+MtDevRegul);
                               MtDevRegul := TLec1.detail[TLec1.Detail.Count-1].GetValue('E_DEBITDEV') - TLec1.detail[TLec1.Detail.Count-1].GetValue('E_CREDITDEV');
                               TLec1.putValue('SOLDEDEV', TLec1.getdouble('SOLDEDEV')+MtDevRegul);
                               if (TLec1.getdouble('SOLDE') <> 0) or (TLec1.Getdouble ('SOLDEDEV') <> 0) then  OkMove := TRUE
                               else OkMove := FALSE;
                          end;
                          if OkMove then
                          begin
                                if TlecNew = nil then
                                   TlecNew := TOB.create ('', nil, -1);

                                OnAfficheListeCom('Fichier déséquilibré en montant :' ,LISTEIMPORT);
                                OnAfficheListeCom( FormatDateTime (Traduitdateformat('dd/mm/yyyy'),TLec1.detail[0].GetValue ('E_DATECOMPTABLE'))
                                              + ' Journal : ' + TLec1.detail[0].GetValue ('E_JOURNAL')
                                              + ' Référence : ' + TLec1.detail[0].GetValue ('E_REFINTERNE')
                                              + ' Libellé :' + TLec1.detail[0].GetValue ('E_LIBELLE') ,LISTEIMPORT);

                                TLec1.ChangeParent(TlecNew, -1);
                                if (NbEcriture > TLec1.detail.count) then
                                   NbEcriture := NbEcriture- TLec1.detail.count;
                          end
                          else inc (ie);
                     end
                     else inc(ie);
               end;
end;

procedure TImportCom.EcrireFicRejet(Tl : TOB);
var
Typecpte, Ligne, Sens, Axe   : string;
Montant1, Montant2, Montant3 : string;
CodeMontant                  : string;
DecDev                       : Integer;
EnDevise                     : Boolean;
tv, tt                       : integer;
TOBAxe,Tla                   : TOB;

        Function QuelMt(p : string; b : Byte ; Decim : Integer) : String ;
        Var MD,MC,Montant : Double ;
            Sup : String ;
        BEGIN
        Result:='' ; Sup:='' ; Montant:=0 ;
        Case b Of 1 : Sup:='DEV' ; END ;
        MD:=TL.GetDouble(p+'_DEBIT'+Sup); MC:=TL.GetDouble(p+'_CREDIT'+Sup) ;   // Fiche 10620
        MD:=Arrondi(MD,Decim) ; MC:=Arrondi(MC,Decim) ;
        if ((MD<0) or (MD>0)) and (MC=0) then BEGIN Montant:=MD ; Sens:='D' ; END else
         if ((MC<0) or (MC>0)) and (MD=0) then BEGIN Montant:=MC ; Sens:='C' ; END ;
        Result:=StrfMontant(Montant,20,Decim,'',False) ;
        END ;

begin
        EnDevise := FALSE;
        if TL.getvalue ('E_AUXILIAIRE') <>'' then Typecpte:= 'X' else Typecpte:=' ';

        CodeMontant := 'E--';
        DecDev := V_PGI.OkDecV;
        if (TL.GetValue('E_DEVISE')=V_PGI.DevisePivot) then
        BEGIN
          DecDev:=V_PGI.OkDecV ;  EnDevise:= FALSE;
        END else
        BEGIN
            if TDevise.FindFirst(['D_DEVISE'],[TL.GetValue('E_DEVISE')],true) <> nil then
            BEGIN
              DecDev:=TDEVISE.GetValue('D_DECIMALE');
              EnDevise:=TRUE ;
            END ;
        END ;
        If EnDevise Then
        BEGIN
                              CodeMontant:='DE-' ;
                              Montant1:=QuelMt('E', 1, DecDev) ;
                              Montant2:=QuelMt('E', 0, V_PGI.OkDecV) ;
                              Montant3:=Format_String(' ',20) ;
        END Else
        BEGIN
                              CodeMontant:='E--';
                              Montant1:=QuelMt('E', 0, V_PGI.OkDecV) ;
                              Montant2:=Format_String(' ',20) ;
                              Montant3:=Format_String(' ',20) ;
        END ;


        Ligne := Format(Formatmvtetendu, [TL.GetValue ('E_JOURNAL'), FormatDateTime(Traduitdateformat('ddmmyyyy'),TL.GetValue ('E_DATECOMPTABLE')),
        TL.Getvalue ('E_NATUREPIECE'),  TL.Getvalue ('E_GENERAL'),Typecpte,  TL.getvalue ('E_AUXILIAIRE'),
        TL.Getvalue('E_REFINTERNE'), TL.Getvalue ('E_LIBELLE'), TL.Getvalue('E_MODEPAIE'),FormatDateTime(Traduitdateformat('ddmmyyyy'),TL.GetValue ('E_DATEECHEANCE')),
        Sens, Montant1, TL.GetValue ('E_QUALIFPIECE'), IntToStr(TL.GetValue ('E_NUMEROPIECE')), TL.GetValue('E_DEVISE'), AlignDroite(StrfMontant(TL.Getdouble('E_COTATION'),20,9,'',False),10), CodeMontant,
        Montant2, Montant3, TL.Getvalue('E_ETABLISSEMENT'), Axe,
        IntToStr(TL.GetValue('E_NUMECHE')), TL.Getvalue('E_REFEXTERNE'), FormatDateTime(Traduitdateformat('ddmmyyyy'),Tl.getValue('E_DATEREFEXTERNE')),
        FormatDateTime(Traduitdateformat('ddmmyyyy'),TL.GetValue ('E_DATECREATION')), TL.GetValue('E_SOCIETE'), TL.GetValue('E_AFFAIRE'), FormatDateTime(Traduitdateformat('ddmmyyyy'),TL.GetValue ('E_DATETAUXDEV')),
        TL.GetValue('E_ECRANOUVEAU'), AlignDroite(StrfMontant(TL.GetDouble('E_QTE1'),20,4,'',False),20), AlignDroite(StrfMontant(TL.GetDouble('E_QTE2'),20,4,'',False),20),  TL.GetValue('E_QUALIFQTE1'),
        TL.GetValue('E_QUALIFQTE2'), TL.GetValue('E_REFLIBRE'), TL.GetValue('E_TVAENCAISSEMENT'), TL.GetValue('E_REGIMETVA'),
        TL.GetValue('E_TVA'), TL.GetValue('E_TPF'), TL.GetValue('E_CONTREPARTIEGEN'), TL.GetValue('E_CONTREPARTIEAUX'),
        TL.GetValue('E_REFPOINTAGE'), FormatDateTime(Traduitdateformat('ddmmyyyy'), TL.GetValue ('E_DATEPOINTAGE')), FormatDateTime(Traduitdateformat('ddmmyyyy'), TL.GetValue ('E_DATERELANCE')),
        FormatDateTime(Traduitdateformat('ddmmyyyy'), TL.GetValue ('E_DATEVALEUR')), TL.GetValue ('E_RIB'), TL.GetValue('E_REFRELEVE'), TL.GetValue('E_IMMO'),
        TL.GetValue('E_LIBRETEXTE0'), TL.GetValue('E_LIBRETEXTE1'), TL.GetValue('E_LIBRETEXTE2'), TL.GetValue('E_LIBRETEXTE3'), TL.GetValue('E_LIBRETEXTE4'),
        TL.GetValue('E_LIBRETEXTE5'), TL.GetValue('E_LIBRETEXTE6'), TL.GetValue('E_LIBRETEXTE7'), TL.GetValue('E_LIBRETEXTE8'),
        TL.GetValue('E_LIBRETEXTE9'), TL.GetValue('E_TABLE0'),TL.GetValue('E_TABLE1'), TL.GetValue('E_TABLE2'), TL.GetValue('E_TABLE3'),
        TL.GetValue('E_LIBREMONTANT0'), TL.GetValue('E_LIBREMONTANT1'),
        TL.GetValue('E_LIBREMONTANT2'), TL.GetValue('E_LIBREMONTANT3'), FormatDateTime(Traduitdateformat('ddmmyyyy'),TL.GetValue ('E_LIBREDATE')),
        TL.GetValue('E_LIBREBOOL0'), TL.GetValue('E_LIBREBOOL1'), TL.GetValue('E_CONSO'), AlignDroite(StrfMontant(TL.GetValue('E_COUVERTURE'),20,V_PGI.OkDecV,'',False),20),
        AlignDroite(StrfMontant(TL.GetValue('E_COUVERTUREDEV'),20,V_PGI.OkDecV,'',False),20), AlignDroite(StrfMontant(0,20,V_PGI.OkDecE,'',False),20),
        FormatDateTime(Traduitdateformat('ddmmyyyy'),TL.GetValue ('E_DATEPAQUETMAX')), FormatDateTime(Traduitdateformat('ddmmyyyy'),TL.GetValue('E_DATEPAQUETMIN')), TL.GetValue('E_LETTRAGE'),
        TL.GetValue('E_LETTRAGEDEV'),'-', TL.GetValue('E_ETATLETTRAGE'),TL.GetVAlue('E_TABLE0'), TL.GetVAlue('E_TABLE1'), TL.GetVAlue('E_TABLE2'),
        TL.GetValue('E_TABLE3'), TL.GetValue('E_REFGESCOM'), TL.GetValue('E_TYPEMVT'), TL.GetValue('EC_DOCGUID'), TL.GetValue('E_TRESOSYNCHRO'),
        TL.GetValue('E_NUMTRAITECHQ'), TL.GetValue('E_NUMENCADECA'), TL.GetValue('E_VALIDE'),
        '', '', '', '',  TL.GetValue('E_CONFIDENTIEL'), TL.GetValue('E_CFONBOK'), TL.GetValue('E_CODEACCEPT')
        ,TL.GetValue('E_DOCID'), '', '', '', TL.GetValue('E_QUALIFORIGINE')]);
      writeln(FRejet, Ligne);
      for tt :=0 to Tl.Detail.Count-1 do
      begin
              TOBAxe := TL.Detail[tt];
              for tv:=0 to TOBAxe.Detail.Count-1 do
              begin
                Tla := TOBAxe.Detail[tv];
                if TLa.GetValue ('Y_TYPEANALYTIQUE') = 'X' then Typecpte:='O'
                else Typecpte:='A' ;
                Ligne := Format(Formatmvtetendu, [TLa.GetValue ('Y_JOURNAL'), FormatDateTime(Traduitdateformat('ddmmyyyy'),TLa.GetValue ('Y_DATECOMPTABLE')),
                TLa.Getvalue ('Y_NATUREPIECE'),  TLa.Getvalue ('Y_GENERAL'), Typecpte,  TLa.getvalue ('Y_SECTION'),
                TLa.Getvalue('Y_REFINTERNE'), TLa.Getvalue ('Y_LIBELLE'), '','',
                Sens, Montant1, TLa.GetValue ('Y_QUALIFPIECE'), IntToStr(TLa.GetValue ('Y_NUMEROPIECE')), TLa.GetValue('Y_DEVISE'), AlignDroite(StrfMontant(TLa.Getdouble('Y_TAUXDEV'),20,9,'',False),10), CodeMontant,
                Montant2, Montant3, TLa.Getvalue('Y_ETABLISSEMENT'), TLa.GetValue('Y_AXE'), '', TLa.Getvalue('Y_REFEXTERNE'), FormatDateTime(Traduitdateformat('ddmmyyyy'),TLa.getValue('Y_DATEREFEXTERNE')),
                FormatDateTime(Traduitdateformat('ddmmyyyy'),TLa.GetValue ('Y_DATECREATION')), TLa.GetValue('Y_SOCIETE'), TLa.GetValue('Y_AFFAIRE'), FormatDateTime(Traduitdateformat('ddmmyyyy'),TLa.GetValue ('Y_DATETAUXDEV')),
                TLa.GetValue('Y_ECRANOUVEAU'), AlignDroite(StrfMontant(TLa.GetDouble('Y_QTE1'),20,4,'',False),20), AlignDroite(StrfMontant(TLa.GetDouble('Y_QTE2'),20,4,'',False),20),  TLa.GetValue('Y_QUALIFQTE1'),
                TLa.GetValue('Y_QUALIFQTE2'), TLa.GetValue('Y_REFLIBRE'),
                '', '',
                '', '', TLa.GetValue('Y_CONTREPARTIEGEN'), TLa.GetValue('Y_CONTREPARTIEAUX'),
                '', TLa.GetValue('Y_SOUSPLAN1'), TLa.GetValue('Y_SOUSPLAN2'),
                TLa.GetValue('Y_SOUSPLAN3'), TLa.GetValue('Y_SOUSPLAN4'), TLa.GetValue('Y_SOUSPLAN5'), '',
                TLa.GetValue('Y_LIBRETEXTE0'), TLa.GetValue('Y_LIBRETEXTE1'), TLa.GetValue('Y_LIBRETEXTE2'), TLa.GetValue('Y_LIBRETEXTE3'), TLa.GetValue('Y_LIBRETEXTE4'),
                TLa.GetValue('Y_LIBRETEXTE5'), TLa.GetValue('Y_LIBRETEXTE6'), TLa.GetValue('Y_LIBRETEXTE7'), TLa.GetValue('Y_LIBRETEXTE8'),
                TLa.GetValue('Y_LIBRETEXTE9'), TLa.GetValue('Y_TABLE0'),TLa.GetValue('Y_TABLE1'), TLa.GetValue('Y_TABLE2'), TLa.GetValue('Y_TABLE3'),
                TLa.GetValue('Y_LIBREMONTANT0'), TLa.GetValue('Y_LIBREMONTANT1'),
                TLa.GetValue('Y_LIBREMONTANT2'), TLa.GetValue('Y_LIBREMONTANT3'), FormatDateTime(Traduitdateformat('ddmmyyyy'),TLa.GetValue ('Y_LIBREDATE')),
                TLa.GetValue('Y_LIBREBOOL0'), TLa.GetValue('Y_LIBREBOOL1'), TLa.GetValue('Y_CONSO'), AlignDroite(StrfMontant(0,20,V_PGI.OkDecE,'',False),20),
                AlignDroite(StrfMontant(0,20,V_PGI.OkDecE,'',False),20), AlignDroite(StrfMontant(0,20,V_PGI.OkDecE,'',False),20),
                '', '', '',
                '','-', '',TLa.GetVAlue('Y_TABLE0'), TLa.GetVAlue('Y_TABLE1'), TLa.GetVAlue('Y_TABLE2'),
                TLa.GetValue('Y_TABLE3'), '', TLa.GetValue('Y_TYPEMVT'), '', '',
                '', '', TLa.GetValue('Y_VALIDE'),
                '', '', '', '',  TLa.GetValue('Y_CONFIDENTIEL'), '', ''
                ,'','', '', '', TL.GetValue('E_QUALIFORIGINE')]);
              writeln(FRejet, Ligne);
            end;
      end;

end;

// fiche 10586
function TImportCom.TrouveCodeIsoPays (Pays : string) : string;
var
CodeIso, tmp    : string;
Indpays         : integer;
begin
       Result := Pays;  CodeIso := '';
       if SPays = nil then
       begin
              SPays :=  TStringList.Create;
              SPays.Sorted:=TRUE ;
              SPays.Duplicates := dupIgnore;
       end;
      for Indpays := 0 to SPays.Count-1 do
      begin
          If pos(Pays, SPays[Indpays]) <> 0 then
          begin
              tmp := SPays[Indpays];
              CodeIso := ReadTokenPipe (tmp, ':');
              break;
          end;
      end;
      if CodeIso = '' then
      begin
           CodeIso := CodeIsoDuPays (Pays);
           if CodeIso <> '' then
           begin
               SPays.Add (CodeIso + ':'+Pays);
               Result := CodeIso;
           end;
      end
      else Result := CodeIso;
end;

end.


