{$IFDEF WIN32}
{$A-,H-}
{$ELSE}
{$A-}
{$ENDIF}
unit GDECLADF;


interface

Type TSaaProg = (spCPTA,spNEG,spINT) ;

CONST
      MaxRamDispo = 15000 ;
      MaxDataRecSize = 1000 ;
      MaxLgCpt = 13 ;
      MaxKeyLen = 40 ;
      PageSize = 32 ;
      Order = 16 ;
      PageStackSize = 16 ;
      MaxHeight = 7 ;
      Nb_FichierCPTA = 17 ;
      MaxligneBloc = 10 ;
      TailleBloc = 49 ;
      maxbanqueF = 3 ;
      MaxNbPeriode = 24 ;
      SocPath : String[85] = 'SOC0\' ;
      SocUtil : String[85] = 'UTIL\' ;
      ListeMaxZone = 20 ;
      ListeMaxLigne = 7 ;

Const EtatInit=1 ;
      Lettrabe=2 ;
      LettrePartiel=4 ;
      LettreTotal=8 ;
      Pointee=16 ;
      Validee=32 ;
      ANouveau=64 ;
      EtapeRegl=128 ;
      Exportee=16 ;
      Editee=32 ;

Const EcrReel = 0 ;
      EcrLot  = 1 ;
      EcrAbo  = 2 ;
      EcrSim  = 3 ;
      EcrTyp  = 4 ;
      EcrClo  = 5 ;
      EcrAnal = 6 ;
      EcrAnou = 7 ;

TYPE
String1 = String[1] ; String4 = String[4] ; String34 = String[34] ;
String2  = String[2]  ; String3  = String[3]  ; String5  = string[5] ;
String6  = String[6]  ; String8  = string[8]  ; String9  = String[9] ; String10 = string[10] ;
String12 = string[12] ; String14 = string[14] ; String15 = string[15] ;
String20 = string[20] ; String25 = string[25] ; String26 = string[26] ; String21 = String[21] ;
String30 = string[30] ; String35 = string[35] ; String40 = string[40] ;
String50 = string[50] ; String60 = string[60] ; String65 = string[65] ;
String69 = String[69] ; String80 = String[80] ; String70 = String[70] ;
String11 = String[11] ; String22 = String[22] ;
String100 = String[100] ; String160 = String[160] ;  String38 = String[38] ;
String13 = String[13] ; String120 = String[120] ; String81=String[81] ;
String23 = String[23] ;String24 = String[24] ; String49 = String[49] ; String19 = String[19] ;
String17 = String[17] ;String55 = String[55] ;String47 = String[47] ;
String18 = String[18] ;String7 = String[7] ;String37 = String[37] ;
String27 = String[27] ; String200 = String[200] ;String16 = String[16] ;
String29 = String[29] ; String132 = String[123] ; String85 = String[85] ;
String137 = String[137] ; String32 = String[32] ;
Tab4Real = Array[1..4] of Real48 ;
Tab2Real = Array[1..2] of Real48 ;
TabCa = array [1..24] of Real48 ;
TCloTva = RECORD
          Cpte1,Cpte2 : String20 ;
          Periode     : String10 ;
          Taux        : Real48     ;
          END ;


TOpteLigne = RECORD
             Actif : Byte ;
             Option : Array[1..5] of byte ;
             RefAuto : String7 ;
             NumSouche : String7 ;
             SaisieSurListe : Byte ;
             END ;

TOpteLigTreso = RECORD
                EcrContrep : Byte ;
                AvecBor : Byte ;
                END ;

TOptePiece = Array[1..19] of TOpteLigne ;
CumPeriode = Array[0..MaxNbPeriode+1] of Double ;
TOpteTreso = Array[1..18] Of TOpteLigTreso ;
Type TTabPer = Array[1..2*MaxNbPeriode] Of Word ;


TYPE EnregCPTA = RECORD
                 Status : Longint ;
                 LeReseau : Byte ;
                 Case Sorte : byte of

                 255: ( buffer        : array[7..MaxDataRecSize] of byte ) ;
                  {=========PLAN COMPTABLE=========}
                  0:  ( CompteG       : String17   ;  {1}
                        IntituleG     : String35   ;  {2}
                        AbregeG       : String17   ;  {3}
                        Correspond1G  : String17   ;  {4}
                        Correspond2G  : String17   ;  {5}
                        SensG         : Byte       ;  {6}
                        CollectifG    : Byte       ;  {7}
                        CentraliseG   : Byte       ;  {8}
                        PointableG    : Byte       ;  {9}
                        VentilableG   : Byte       ; {10}
                        LettrableG    : Byte       ; {11}
                        NiveauG       : Byte       ; {12}
                        SoldeProgrG   : Byte       ; {13}
                        TotMensG      : Byte       ; {14}
                        SautPageG     : Byte       ; {15}
                        UniteG        : Byte       ; {16}
                        TvaEncaisseG  : Byte       ; {17}
                        DateCreatG    : Word       ; {18}
                        DateModifG    : Word       ; {19}
                        DateFermeG    : Word       ; {20}
                        DateMvtG      : Word       ; {21}
                        TotMvtG       : Real48       ; {22}
                        DerNumPieceG  : Longint    ; {23}
                        DerNumLigneG  : SmallInt    ; {24}
                        DerNumLettreG : Longint    ; {25}
                        NatureG       : Byte       ; {26}
                        CumulFauxG    : Byte       ; {27}
                        PosBlocG      : Longint    ; {28}
                        ModePaiementG : Byte       ; {29}
                        PosAdrG       : LongInt    ; {30}
                        PaiementG     : Byte       ; {31}
                        TotPointeNG   : Tab2Real   ; {32}
                        SectionG      : String17   ; {34}
                        PosSectionG   : LongInt    ; {35}
                        SensMvtG      : Byte       ; {36}
                        TotPointeN1G  : Tab2Real   ; {37}
                        EngageONG     : Byte       ; {38}
                        DernLienG     : LongInt    ; {39}
                        JunkG         : String[41]);

                        {======== FICHIER TIERS =======}
                  1 : ( AuxiliaireT   : String17   ;  {1x}
                        CollectifT    : String17   ;  {2x}
                        IntituleT     : String35   ;  {3x}
                        AbregeT       : String17   ;  {4x}
                        NatureT       : Byte       ;  {5x}
                        Correspond1T  : String17   ;  {6x}
                        Correspond2T  : String17   ;  {7x}
                        SecteurT      : Byte       ;  {8}
                        PointableT    : Byte       ;  {9}
                        LettrableT    : Byte       ; {10}
                        EcheancableT  : Byte       ; {11}
                        NiveauT       : Byte       ; {12x}
                        RegimeTvaT    : Byte       ; {13}
                        PaiementT     : Byte       ; {14}
                        DeviseT       : Byte       ; {15}
                        LimiteT       : Real48       ; {16}
                        RelanceT      : Byte       ; {17}
                        UniteT        : Byte       ; {18}
                        TvaEncaisseT  : Byte       ; {19}
                        DerNumLettreT : Longint    ; {20x}
                        SoldeProgrT   : Byte       ; {21x}
                        TotMensT      : Byte       ; {22x}
                        SautPageT     : Byte       ; {23x}
                        DateCreatT    : Word       ; {24x}
                        DateModifT    : Word       ; {25x}
                        DateFermeT    : Word       ; {26x}
                        DateMvtT      : Word       ; {27x}
                        TotMvtT       : Real48       ; {28x}
                        DerNumPieceT  : Longint    ; {29x}
                        DerNumLigneT  : SmallInt    ; {30x}
                        CumulFauxT    : Byte       ; {31}
                        PosAdresseT   : Longint    ; {32x}
                        PosBlocT      : Longint    ; {33}
                        TotPointeNT   : Tab2Real   ; {34}
                        CompteTrieT   : String17   ; {35}
                        SensMvtT      : Byte       ; {36x}
                        TotPointeN1T  : Tab2Real   ; {37}
                        PosAuxNegT    : LongInt    ; {38}
                        MarcheONT     : Byte       ; {39}
                        DernLienT     : LongInt    ; {40}
                        JunkT         : String[17]  );

                         {======== COMPTES ANALYTIQUES =======}
                  2:  ( CompteA       : String17   ;  {1x}
                        CompteTrieA   : String17   ;  {2}
                        IntituleA     : String35   ;  {3x}
                        AbregeA       : String17   ;  {4x}
                        NiveauA       : Byte       ;  {5x}
                        SoldeProgrA   : Byte       ;  {6x}
                        TotMensA      : Byte       ;  {7x}
                        SautPageA     : Byte       ;  {8x}
                        DateCreatA    : Word       ;  {9x}
                        DateModifA    : Word       ; {10x}
                        DateFermeA    : Word       ; {11x}
                        DateMvtA      : Word       ; {12x}
                        TotMvtA       : Real48       ; {13x}
                        DerNumPieceA  : Longint    ; {14x}
                        DerNumLigneA  : SmallInt    ; {15x}
                        CumulFauxA    : Byte       ; {16}
                        PosBlocA      : Longint    ; {17}
                        SensMvtA      : Byte       ; {18x}
                        JunkA         : String1  );

                         {======== FICHIER JOURNAUX =======}
                  3:  ( CodeJ         : String[3]  ;  {1x}
                        IntituleJ     : String35   ;  {2x}
                        AbregeJ       : String17   ;  {3x}
                        NatureJ       : Byte       ;  {4x}
                        ContrepartieJ : Byte       ;  {5}
                        PieceTypeJ    : Byte       ;  {6}
                        CentraliseJ   : Byte       ;  {7}
                        AlerteCaisseJ : Byte       ;  {8}
                        DernierFolioJ : Longint    ;  {9}
                        OptionsPieceJ : TOptePiece ; {10}
                        CumulFauxJ    : Byte       ; {11}
                        CpteContrepGJ : String17   ; {12}
                        CpteContrepAJ : String17   ; {13}
                        CtrlTvaJ      : Byte       ; {14}
                        ChampCtrJ     : Byte       ; {15}
                        PosCtrPtrGJ   : Longint    ; {16}
                        PosCtrPtrAJ   : Longint    ; {17}
                        PosAdresseJ   : Longint    ; {18}
                        PosBanqueJ    : Longint    ; {19}
                        PosBlocJ      : Longint    ; {20}
                        OptionsTresoJ : TOpteTreso ; {21}
                        CompteurJ     : LongInt    ; {22}
                        LettrableJ    : Byte       ; {23}
                        DeviseJ       : Byte       ; {24x}
                        JunkJ         : String[8] );

                  {======== ECRITURES COMPTABLES =======}
                  4:  ( TypeE         : Byte       ;  {1}
                        NoOrdreE      : SmallInt    ;  {2}
                        NoPieceE      : Longint    ;  {3}
                        CodeJournalE  : String3    ;  {4}
                        DateComptaE   : Word       ;  {5}
                        PosContrepGE  : Longint    ;  {6}
                        PosContrepAE  : LongInt    ;  {7}
                        SensE         : Byte       ;  {8}
                        MontantFrcsE  : Real48       ;  {9}
                        ReferenceE    : String13   ; {10}
                        LibelleE      : String35   ; {11}
                        UniteE        : Byte       ; {12} { Secteur M9 }
                        QuantiteE     : Real48       ; {13}
                        MtantDeviseE  : Real48       ; {14}
                        DeviseE       : Byte       ; {15}
                        DateOrigineE  : Word       ; {16}
                        RefOrigineE   : String13   ; {17}
                        DateDernEchE  : Word       ; {18}
                        EtatE         : Byte       ; {19}
                        RefPointageE  : String13   ; {20}
                        EcheanceOnE   : Byte       ; {21}
                        AnalytiqueOnE : Byte       ; {22}
                        LettrageOnE   : Byte       ; {23}  {?????}
                        QuantiteOnE   : Byte       ; {24}  {?????}
                        NbModifE      : Byte       ; {25}  {Nombre de fois que cette ‚criture a ‚t‚ modifi‚e}
                        CodeUtilE     : Byte       ; {26}
                        DatePieceE    : Word       ; {27}
                        PosCpteGeneE  : Longint    ; {28}
                        PosCpteAuxE   : Longint    ; {29}
                        PosEnteteE    : Longint    ; {30}
                        PosBanqueGE   : Longint    ; {31}
                        PosJalE       : LongInt    ; {32}
                        PosFactE      : Longint    ; {33}
                        RuptE         : Byte       ; {34}
                        PremZoneE     : String1    ; {35}
                        SimulE        : Byte       ; {36}
                        RegimeE       : Byte       ; {37}
                        EncaisseE     : Byte       ; {38}
                        TypePieceE    : Byte       ; {39}{c'est 42}
                        JunkE         : String[5]  );
                        { Attention … UTILSAV4 Si agrandissement du fichier }
                        { Attention : pour la recup, le junk doit avoir au
                          moins deux caractŠres }

                     {======== FICHE GRAPH ========}
                  5 :(Tableg      : Array [1..3] of string30 ;
                      TailleLeg   : Array[1..3] of byte ;
                      FontLeg     : Array[1..3] of byte ;
                      GrHisto     : Array[1..2] of byte ;
                      GrMotif     : Array[1..2] of byte ;
                      GrCouleur   : Array[1..2] of byte ;
                      GrPoints    : Array[1..2] of byte ;
                      GrValeur    : Array[1..2] of byte ;
                      GrUnite     : Array[1..2] of byte ;
                      TailleUnite : Array[1..2] of byte ;
                      FontUnite   : Array[1..2] of byte ;
                      GrGrille    : Byte ;
                      GrPosition  : Byte ;
                      GrEfface    : Byte ;
                      GrEcran     : Byte ;
                      TitreG      : String30 );

                       {======== EN-TETES =========}
                  6:  ( TypeET        : Byte       ;  {1} { R‚el, Simul, ...}
                        NoPieceET     : Longint    ;  {2}
                        CodeJournalET : String3    ;  {3}
                        DateComptaET  : Word       ;  {4}
                        ReferenceET   : String13   ;  {5}
                        LibelleET     : String35   ;  {6}
                        TypePieceET   : Byte       ;  {7} { Facture, Avoir, ...}
                        TvaEncaisseET : Byte       ;  {8}
                        RefBordereauET: String5    ;  {9}
                        FolioET       : Longint    ; {10}
                        ContrepartieET: Byte       ; {11}
                        NoLotET       : String5    ; {12}
                        TypBordET     : Byte       ; {13}
                        TypCtrET      : Byte       ; {14}
                        DomBanqET     : Byte       ; {15}
                        RibET         : Byte       ; {16}
                        ModReglET     : Byte       ; {17}
                        DateEcheET    : Byte       ; {18}
                        FlagEditeET   : Byte       ; {19}
                        PosJournalET  : LongInt    ; {20}
                        EtatET        : Byte       ; {21}
                        EcheanceONET  : Byte       ; {22}
                        PeriodiciteET : Byte       ; {23}
                        NbRepeteET    : Byte       ; {24}
                        ReconductionET: Byte       ; {25}
                        DateDepartET  : Word       ; {26}
                        NbGenereET    : Byte       ; {27}
                        LastDateET    : Word       ; {28}
                        DateFinET     : Word       ; {29}
                        FinDeMoisET   : Byte       ; {30}
                        CodeUtilET    : Byte       ; {31}
                        PosCtrPtrGET  : LongInt    ; {32}
                        PosCtrPtrAET  : LongInt    ; {33}
                        SimulET       : Byte       ; {34}
                        EncaisseET    : Byte       ; {35}
                        RegimeET      : Byte       ; {36}
                        CompteurET    : LongInt    ; {37}
                        PieceVideET   : Byte       ; {38}
                        PosFactET     : LongInt    ; {39}
                        TvaOkET       : Byte       ; {40}
                        FactNegET     : Byte       ; {41}
                        DeviseET      : Byte       ; {42}
                        TauxDeviseET  : Real48       ; {43}
                        DateCompET    : Word       ; {44}
                        UniteET       : Byte       ; {45} { Secteur }
                        JunkET        : String[4]  );
                        { Attention : pour la recup, le junk doit avoir au
                          moins deux caractŠres }

                       {===PLAN REFERENCE======}
                  7:  ( CompteRF      : String17 ;
                        IntituleRF    : String35 ;
                        SensRF        : Byte     ;
                        CollectifRF   : Byte     ;
                        PosBlocRF     : Longint  ;
                        JunkRF        : String[61] ) ;

                  {======== ECRITURES ANALYTIQUES =======}
                  8:  ( TypeEA        : Byte       ;  {1}
                        NoLigneEA     : SmallInt    ;  {2}
                        CodeJournalEA : String3    ;  {3}
                        DateComptaEA  : Word       ;  {4}
                        SectionEA     : String17   ;  {5}
                        SensEA        : Byte       ;  {6}
                        MontantEA     : Real48       ;  {7}
                        TauxRepartEA  : Real48       ;  {8}
                        ReferenceEA   : String13   ;  {9}
                        LibelleEA     : String35   ; {10}
                        UniteEA       : Byte       ; {11} { Secteur M9 }
                        QuantiteEA    : Real48       ; {12}
                        TauxQteEA     : Real48       ; {13}
                        NoEcrGenEA    : SmallInt    ; {14}
                        NoPieceEA     : longint    ; {15}
                        PosCpteGeneEA : Longint    ; {16}
                        PosCpteAnaEA  : Longint    ; {17}
                        PosEcrGenEA   : LongInt    ; {18}
                        EtatEA        : Byte       ; {19}
                        SimulEA       : Byte       ; {20}
                        OdONEA        : Byte       ; {21}
                        JunkEA        : String[8] );
                        { Attention on s'en sert sur 2 caract. dans IMPSOC }


                  {===============ECHEANCES==============}
                  9:  ( TypeH         : Byte       ;  {1}
                        NoPieceH      : Longint    ;  {2}
                        NoEcritureH   : SmallInt    ;  {3}
                        NoEcheanceH   : Byte       ;  {4}
                        DateEcheanceH : Word       ;  {5}
                        PaiementH     : Byte       ;  {6}
                        SensH         : Byte       ;  {7}
                        MontantFrcsH  : Real48       ;  {8}
                        MtantDeviseH  : Real48       ;  {9}
                        DeviseH       : Byte       ; {10}
                        EtatH         : Byte       ; {11}
                        CompteGeneH   : String17   ; {12}
                        CompteAuxilH  : String17   ; {13}
                        NatureH       : Byte       ; {14}
                        DateComptaH   : Word       ; {15}
                        DerniereRelanceH : Word    ; {16}
                        NiveauDernRelanceH : Byte  ; {17}
                        CodeLettreH   : Longint    ; {18}
                        DateLettrageH : Word       ; {19}
                        UtilisateurH  : Byte       ; {20}
                        RefOrigineH   : String13   ; {21}
                        MtantLettreH  : Real48       ; {22}
                        MtantRestantH : Real48       ; {23}
                        DeviseLettreH : Real48       ; {24}
                        DeviseRestantH: Real48       ; {25}
                        EtatGlobalH   : Byte       ; {26}
                        PosEcritureH  : Longint    ; {27}
                        PosGeneH      : Longint    ; {28}
                        PosAuxH       : Longint    ; {29}
                        DateGlobalH   : Word       ; {30}
                        NoEcheNegH    : Byte       ; {31}
                        OldTypePieceH : Byte       ; {32}
                        PosBlocH      : LongInt    ; {33}
                        JunkReseauH   : Byte       ; {34}
                        SimulH        : Byte       ; {35}
                        RegimeH       : Byte       ; {36}
                        EncaisseH     : Byte       ; {37}
                        TypePieceH    : Byte       ; {38}{37}
                        EditeTvaH     : Byte       ; {39}
                        PosTvaH       : LongInt    ; {40}
                        PourcentTvaH  : Real48       ; {41}
                        EtatTvaH      : Byte       ; {42}
                        JunkH         : String5)   ; {43}
                        { Attention : on s'en sert sur 3 caract. dans IMPSOC }
                        { Attention : on s'en sert sur 5 caract. dans LETTAUTO }
                        { Attention : pour la recup, le junk doit avoir au
                          moins deux caractŠres }

{$IFDEF MAJOR}
                     {======== ANCIEN LETTRAGE / REGLEMENT ========}
                 10:  ( TypeEtapeER      : Byte     ; {1}
                        NoEtapeER        : SmallInt  ; {2}
                        IntituleER       : String35 ; {3}
                        JalER            : String3  ; {4}
                        Compte1ER        : String13 ; {5}
                        Compte2ER        : String13 ; {6}
                        ContreER         : String13 ; {7}
                        JunkER1          : String13 ; {8}
                        CodeAFBER        : String5  ; {9}
                        RefPieceER       : String13 ; {10}
                        LibPieceER       : String35 ; {11}
                        TypeEcritureER   : Byte     ; {12}
                        TypeContreER     : Byte     ; {13}
                        SensEcritureER   : Byte     ; {14}
                        TypePieceER      : Byte     ; {15}
                        PaiementER       : Byte     ; {16}
                        CollectifER      : Byte     ; {17}
                        PosJalER         : LongInt  ; {18}
                        PosContreER      : LongInt  ; {19}
                        PosCpteCreditER  : LongInt  ; {20}
                        BudgetER         : Byte     ; {21}
                        JunkER           : String[5] )    ;

{$ELSE}
                     {======== ANCIEN LETTRAGE ========}
                 10:  ( JunkL : String80 ) ;

{$ENDIF}

                     {======== CUMULS ==========}
{ En ce qui concerne la sorte des fiches cumuls, on adopte la convention suivante:
      0..3 : ‚quivalente … la sorte du fichier de base
      4    : Compte G‚n‚ral    --> Compte Auxiliaire
      5    : Compte Auxiliaire --> Compte G‚n‚ral
      6    : Compte G‚n‚ral    --> Compte Analytique
      7    : Compte Analytique --> Compte G‚n‚ral
      8    : Somme de Solde pour Comptes G‚n‚raux Collectifs
      + 10 pour les ‚critures de simulation }

                 11:  ( SorteC        : Byte       ; {1}
                        TypeEcritureC : Byte       ;  {2}
                        Compte1C      : String17   ;  {3}
                        Compte2C      : String17   ;  {4}
                        TotDebitC     : CumPeriode ;  {5}
                        TotCreditC    : CumPeriode ;  {6}
                        ExerciceC     : SmallInt    ;  {7}
                        NbPeriodeC    : Byte       ;  {8}
                        PremierePC    : Byte       ;  {9}
                        JunkC         : String[3]  );

                  {======== PAGE GESCOM / TIERS =========}
                 12:  ( FormeJurGT    : Byte       ;  {1}
                        NoCommunGT    : String20   ;  {2}
                        ApeGT         : String20   ;  {3}
                        ContactGT     : String35   ;  {4}
                        Adresse1GT    : String35   ;  {5}
                        Adresse2GT    : String35   ;  {6}
                        Adresse3GT    : String35   ;  {7}
                        CodePostalGT  : String9    ;  {8} { ???? sur la longueur}
                        VilleGT       : String35   ;  {9}
                        PaysGT        : Byte       ; {10}
                        TvaGT         : String20   ; {11}
                        TelephoneGT   : String20   ; {12}
                        FaxGT         : String20   ; {13}
                        LangueGT      : Byte       ; {14}
                        SorteGT       : SmallInt    ; {15}
                        PosTiersGT    : Longint    ; {16}
                        JunkGT        : String[75]  );

                      {============BANQUES===========}
                 13:  ( BanqueB       : String35 ;    {1}
                        CodeBanqueB   : String9  ;    {2}
                        CodeGuichetB  : String9  ;    {3}
                        CompteBanqueB : String20 ;    {4}
                        CleRibB       : String5  ;    {5}
                        SorteB        : Byte     ;    {6}
                        PosFicheB     : Longint  ;    {7}
                        NoOrdreB      : SmallInt  ;    {8}
                        PosJalB       : LongInt  ;    {9}
                        JunkB         : String[27]  );

                        {===VENTILATION ANALYTIQUE==}
                 14:  ( CompteGeneVA  : String17 ;    {1}
                        PosGenVA      : longint ;     {2}
                        CompteAnalVA  : String17 ;    {3}
                        PourcentVA    : Real48     ;    {5}
                        NoOrdreVA     : SmallInt ;     {6}
                        JunkVA        : String[82]  );

                        {======== VENTILATION JOURNAL =======}
                 15:  ( TypeEnrVJ     : Byte ;       {1}
                        PosCodeJalVJ  : longint ;    {2}
                        Compte1VJ     : String17 ;   {3}
                        Compte2VJ     : String17 ;   {4}
                        NoOrdreVJ     : SmallInt  ;   {5}
                        ColonneVJ     : Byte     ;   {6}
                        SensVJ        : Byte     ;   {7}
                        SigneVJ       : Byte     ;   {8}
                        TauxVJ        : Real48     ;   {9}
                        PeriodeVJ     : String10 ;   {10}
                        JunkVJ        : String[58]  );

                        {======TRANCHES ANALYTIQUES========}
                        { INDEX : NoTable + CodeTranche }
                 16:  ( NoTableTA     : Byte     ;    {1} {de 1 … 19 : Analytique / de 20 … 39 : Ruptures}
                        LongCodeTA    : Byte     ;    {2}
                        CodeTrancheTA : String17 ;    {3}
                        IntituleTA    : String35 ;    {4}
                        TypBordTA     : Byte     ;    {5}
                        TypCtrTA      : Byte     ;    {6}
                        DomBanqTA     : Byte     ;    {7}
                        RibTA         : Byte     ;    {8}
                        ModReglTA     : Byte     ;    {9}
                        DateEcheTA    : Byte     ;    {10} { Editable O/N }
                        DernBorReelTA : LongInt  ;    {11}
                        DernBorLotTA  : LongInt  ;    {12}
                        DernBorSimTA  : LongInt  ;    {13}
                        DernBorAboTA  : LongInt  ;    {14}
                        Compte1TA     : String17 ;    {15}
                        Compte2TA     : String17 ;    {16}
                        TauxTvaTA     : Real48     ;    {17}
                        JunkTA        : String[1]  );

                         {======== FICHE POSTAL ========}
                     17 :(CodeI       : String8 ;
                          VilleI      : String22 );

                        END ;

Descri_FichierCPTA = RECORD
                 nom : string60 ;
                 rec_size : SmallInt ;
                 key_size : SmallInt ;
                 duplicate: SmallInt ;
                 END ;


CONST  fichiersCPTA : array[0..Nb_fichierCPTA] of descri_fichierCPTA =
              ((nom : 'PLAN' ;
                    rec_size : 256 ;
                    key_size : 14  ;             {0}
                    duplicate: 1 ),

               (nom : 'TIERS' ;
                    rec_size : 256 ;
                    key_size : 14  ;             {1}
                    duplicate: 1 ),

               (nom : 'ANALYSE' ;
                    rec_size : 128 ;
                    key_size : 14  ;             {2}
                    duplicate: 1 ),

               (nom : 'JOURNAUX' ;
                    rec_size : 620 ;
                    key_size : 11  ;             {3}
                    duplicate: 1 ),

               (nom : 'ECRITURE' ;
                    rec_size : 176 ;
                    key_size : 31  ;             {4}
                    duplicate: 1  ),

               (nom : 'GRAPH' ;
                    rec_size : 156 ;
                    key_size : 0  ;             {5}
                    duplicate: -1  ),

               (nom : 'ENTETE' ;
                    rec_size : 148 ;
                    key_size : 11  ;             {6}
                    duplicate: 1  ),

               (nom : 'PLANREF' ;
                    rec_size : 128 ;
                    key_size : 18  ;             {7}
                    duplicate: 0  ),

               (nom : 'ANAJOUR' ;
                    rec_size : 140 ;
                    key_size : 31  ;             {8}
                    duplicate: 1  ),

               (nom : 'ECHEANCE' ;
                    rec_size : 163 ;
                    key_size : 38  ;             {9}
                    duplicate: 1  ),

{$IFDEF MAJOR}
               (nom : 'ETAPE' ;
                    rec_size : 185 ;
                    key_size : 4  ;            {10}
                    duplicate: 1  ),
{$ELSE}
               (nom : 'TOTO' ;
                    rec_size : 91 ;
                    key_size : 0  ;            {10}
                    duplicate: -1  ),
{$ENDIF}

               (nom : 'CUMULS' ;
                    rec_size : 468 ;   { Taille pour V8 et ant‚rieure : 364 }
                    key_size : 30  ;            {11}
                    duplicate: 1  ),

               (nom : 'GESCOM' ;
                    rec_size : 384 ;
                    key_size : 0  ;            {12}
                    duplicate: -1  ),

               (nom : 'BANQUES' ;
                    rec_size : 128 ;
                    key_size : 8   ;            {13}
                    duplicate: 0  ),

               { INDEX1 : PosCptGen + NoOrdre }
               (nom : 'VENTLANA' ;
                    rec_size : 128 ;
                    key_size : 7  ;            {14}
                    duplicate: 0  ),

               (nom : 'VENTJOUR' ;
                    rec_size : 128 ;
                    key_size : 8  ;            {15}
                    duplicate: 1  ),

               { INDEX1 : NoTable + CodeTranche (sur 10 caractŠres)}
               (nom : 'TRANCHES' ;
                    rec_size : 128 ;
                    key_size : 19  ;            {16}
                    duplicate: 1  ),

               (nom : 'PosTal' ;
                    rec_size : 38 ;              {7}
                    key_size : 0 ;
                    duplicate: -1)) ;


{======================================================================}
{============================= NEGOCE =================================}
{======================================================================}
CONST
      Nb_FichierNEG = 21 ;
      MaxNiveauPass = 16 ;
      NbMaxRayon = 20 ;
      NbMaxGrille = 40  ;
      NbMaxRepres = 96 ;
      NbMaxComm = 21 ;
      MaxStock = 10 ;
      MaxStat  = 24 ;
      NbMaxPrix = 32 ;
      TotalPrixP = 48 ;
      NbMaxQte  = 16 ;
      NbMaxAssemble = 18 ;
      MaxLastFourn = 5 ;
      MaxTvaTpf=6 ;
      MaxTypeTvaTpf=6 ;
      MaxModePaie=12 ;
      MaxRegle = 15 ;
      MaxFactRL = 30 ;
      NbMaxJou=14 ;
TYPE
TableauMode = array [1..MaxModePaie]  of Real48        ;
TabTyp  = Array [1..MaxModePaie]  of Byte        ;
TabTaxeTvaTpf  = Array [1..MaxTvaTpf] of Real48    ;
Fiche_Comm   = Array[1..NbMaxComm,1..2] of Single ;

ReAppro  = RECORD
           Status     : longint ;
           Sorte      : byte    ;
           Reseau     : boolean ;
           CodeOO     : longint ;
           RefFournOO : longint ;
           QteOO      : Single  ;
           PrixaOO    : Single  ;
           RemiseOO   : Single  ;
           DepotOO    : Byte    ;
           END ;

AdrDep = RECORD
         Respons  : String20 ;
         Design   : String20 ;
         Adresse  : String20 ;
         CodePost : String9  ;
         Ville    : String20 ;
         Tel      : String15 ;
         Fax      : String15 ;
         Ape      : String15 ;
         RCS      : String40 ;
         Siret    : String15 ;
         Comment  : String30 ;
         END ;

LigneRemise = RECORD
             Prix      : Single ;
             Date      : Word ;
             Eche      : Word ;
             ModePaie  : Byte ;
             NumCli    : Longint ;
             NumFact   : longint ;
             NumBanque : Byte ;
             CliBanque : Byte ;
             END ;


LigneASS = RECORD
           PosFiche : Longint ;
           Qte      : Single  ;
           Prix     : Single  ;
           END ;

ModeMode = RECORD
           TypMode   : Byte    ;
           EcheMode  : Word    ;
           MontMode  : Real48    ;
           Etat      : Byte    ;
           END ;

CaisseMode  = array[1..MaxModePaie] of ModeMode ;
Type CompteEnBanque = RECORD
                      Agence     : String25 ;
                      Rue        : String26 ;
                      Localite   : String22 ;
                      CodePost   : String8  ;
                      Ville      : String22 ;
                      RIB        : String24 ;
                      CodeJour   : String3  ;
                      Compte     : String13 ;
                      CompteVir  : String13 ;

                      Fax        : String15 ;
                      Ape        : String15 ;
                      RCS        : String40 ;
                      Siret      : String15 ;
                      Comment    : String30 ;
                      END ;

Type EnregCEE = RECORD
                ObligAch   : String1  ;
                ObligVen   : String1  ;
                NumIdent   : String14 ;
                TiersDecla : String5  ;
                Contact    : String40 ;
                Mention : Array[1..3] of String50 ;
                DateDeb : Array[1..6] of Word ;
                DateFin : Array[1..6] of Word ;
                END ;

Type Fiche_Tva = RECORD
                 CompteTva  : Array[1..2,1..MaxTypeTvaTpf] of String13 ;
                 END ;

Type LigBN = Array[1..MaxligneBloc] of String[TailleBloc] ;
Type BNNEG  = RECORD
                 Status : LongInt ;
                 PosClient : LongInt ;
                 Lignes : LigBN ;
                 END ;

      DevArt   = RECORD
                 PxDev  : Single ;
                 NumDev : byte ;
                 END ;

TabRealSav = Array [1..5] of Real48 ;
Fiche_SAV_P  = RECORD
               Code : Longint ;
               Travail : Array [1..5] of String40 ;
               Prix : Array [1..5] of Single ;
               END ;

FournArt   = RECORD
             Fourn   : longint ;
             PA      : Single  ;
             QteMin  : Single  ;
             QteArr  : Single  ;
             Delai   : Byte    ;
             Red     : Single  ;
             Date    : Word    ;
             END ;

Regle    = RECORD
           EtatV        : Byte       ;
           EcheanceV    : Word       ;
           DateReglV    : Word       ;
           ModeV        : Byte       ;
           MontantV     : Real48       ;
           RegleV       : Real48       ;
           SaisiV       : Byte       ;
           PassV        : Byte       ;
           RemisV       : Byte       ;
           RapprochV    : Longint    ;
           DateRemBaV   : Word       ;
           END ;

TabRegle = Array [1..MaxRegle] of Regle ;
TabDispo   = Array [1..MaxStock,1..8] of single ;
TabBornesQte  = Array [1..NbMaxQte] of Single ;
TabTarif1     = Array [1..NbMaxPrix] of Single  ;
TabTarif     = RECORD
               Prix  : TabTarif1 ;
               Qte   : TabBornesQte ;
               NbQte : byte ;
               END ;

type    CleCaisse  = RECORD
          Date    : Word    ;
          Caisse  : byte    ;
          END ;


TYPE EnregNEG  = RECORD
                 status : LongInt ;
                 reseau : Boolean ;
                 case sorte : byte of
                         {======== FICHE CLIENT =======}
                 255: ( buffer        : array[7..MaxDataRecSize] of byte ) ;
                  0 : ( NOM           : STRING25   ; {*}
                        CODEMAG       : STRING18   ; {*}
                        CONTACTC      : STRING20   ; {*}
                        NUMAUXILI     : STRING13   ; {*}
                        RUE           : STRING26   ; {*}
                        LOCALITE      : STRING22   ; {*}
                        CODEPOSTAL    : STRING8    ; {*}
                        VILLE         : STRING22   ; {*}
                        NUMCOMPTA     : STRING13   ; {*}
                        CommentaireC  : String20   ; {*}
                        Date          : Word       ; {*}
                        HistOuiC      : Byte       ;
                        Vendeur       : byte       ;
                        CoefCommRC    : Single     ;
                        Secteur       : byte       ;
                        TELEPHONE     : STRING15   ; {*}
                        Telex         : string10   ; {*}
                        Telecopie     : string15   ; {*}
                        BANQUE        : STRING20   ; {*}
                        NUMBANQUE     : STRING5    ; {*}
                        NUMGUICHET    : STRING5    ; {*}
                        NUMCOMPTE     : STRING12   ; {*}
                        NUMRIB        : STRING2    ; {*}
                        DateCC        : Word       ; {*}
                        DateMC        : Word       ; {*}

                        Reglement     : Byte       ;
                        TARIF         : BYTE       ;
                        TypTVAC       : BYTE       ;
                        TTC           : Byte       ;
                        TypeTpfC      : Byte       ;
                        Remise        : Single     ;
                        Escompte      : Single     ;
                        Groupe        : Byte       ;
                        Releve        : Byte       ;
                        Exemplaire    : Array [1..4] of Byte ;
                        Expedition    : Byte       ;
                        Transport     : Byte       ;

                        Acompte       : Real48       ;
                        Avoir         : Real48       ;
                        BlNonFactC    : Real48       ;
                        TotEffetC     : Real48       ;
                        SoldeTTCC     : Real48       ;
                        EncoursMaxC   : Real48       ;
                        DateFact      : Word       ;
                        MontFact      : Real48       ;
                        DateRegle     : Word       ;
                        MontRegle     : Real48       ;
                        DateReser     : word       ;
                        MontReser     : Real48       ;
                        DeviseC       : Byte       ;
                        Apporteur     : byte       ;
                        PosTarifCCA   : longint    ;
                        BlocC         : Longint    ;
                        CoefCommAC    : Single     ;
                        TotRisqueC    : Single     ;
                        ANC           : Single     ;
                        SoldeHTC      : Real48       ;
                        NifCEEC       : String14   ;
                        NiveauC       : Byte       ;
                        JunkC         : String16  );

                         {======== FICHE ARTICLE ========}
                      1:(Categorie    : byte      ;
                         Famille      : byte      ;
                         Marque       : byte      ;
                         Modele       : string30  ;
                         Generic      : Byte      ;
                         CommentaireA : String20  ;
                         RefMag       : String18  ;
                         RefFourn     : String18  ;
                         EAN13        : string18  ;
                         Place        : Byte      ;
                         Grille       : byte      ;
                         Taille       : byte      ;
                         GrilleCoul   : byte      ;
                         Couleur      : byte      ;
                         Tva          : byte      ;
                         TPF          : Byte      ;
                         TPFFRS       : Single    ;
                         EncaisseA    : Byte      ;
                         Prix_Achat   : Single    ;
                         PMPA         : Single    ;
                         Prix_Base    : Single    ;
                         Prix_TTC     : Single    ;
                         Px_Rev       : Single    ;
                         QteMin       : Single    ;
                         QteMax       : Single    ;
                         CondAchat    : Single    ;
                         CondVente    : Single    ;
                         AffichCondA  : Byte      ;
                         UniteA       : String10  ;
                         HistOuiA     : Byte      ;
                         SommeilA     : Byte      ;
                         Junk111      : Byte      ;
                         Junk112      : Byte      ;
                         Contremar    : Byte      ;
                         Liaison      : Longint   ;
                         Substitution : Longint   ;
                         CoefCommRA   : single    ;
                         Volume       : Single    ;
                         Poids        : Single    ;
                         Freinte      : Single    ;
                         Suivi        : byte      ;

                         Prix_V       : TabTarif1 ;
                         Prix_Q       : TabBornesQte ;
                         Promotion    : Single    ;
                         Promo1       : Word      ;
                         Promo2       : Word      ;

                         PMoyenA      : Single    ;
                         TotQteA      : Single    ;
                         DateCA       : Word      ;
                         DateMA       : Word      ;
                         PinventA     : Single    ;
                         DateInvent   : Word      ;
                         Last_Vente   : Word      ;
                         TypePrix     : Byte      ;
                         NbQte        : Byte      ;
                         BlocA        : Longint   ;
                         CoefCommAA   : Single    ;
                         Disp         : TabDispo  ;
                         Assemble     : Longint   ;
                         Multi_Fourn  : Array [1..MaxLastFourn] of FournArt ;
                         FlagInv      : SmallInt ;
                         TabDevA      : Array[1..3] of DevArt ;
                         JunkA        : String4 ) ;

                         {======= FICHE FOURNISSEUR =====}
                      2:
                        (NomF         : String25   ;
                         CodeFourn    : String18   ;
                         NumAuxiliF   : String13   ;
                         ContactF     : string20   ;
                         RueF         : string26   ;
                         LocaliteF    : String22   ;
                         CodePostalF  : String8    ;
                         VilleF       : string22   ;
                         CommentaireF : String20   ;
                         telephoneF   : string15   ;
                         telexF       : string10   ;
                         telefaxF     : string15   ;
                         BANQUEF      : STRING20   ;
                         NUMBANQUEF   : STRING5    ;
                         NUMGUICHETF  : STRING5    ;
                         NUMCOMPTEF   : STRING12   ;
                         NUMRIBF      : STRING2    ;
                         NumComptaF   : String13   ;
                         SecteurF     : byte       ;
                         DateF        : Word       ;
                         HistOuiF     : Byte       ;
                         DelaiF       : SmallInt    ;
                         QteMinF      : Single     ;
                         QteArrF      : Single     ;
                         DateCF       : Word       ;
                         DateMF       : Word       ;

                         ReglementF   : Byte       ;
                         TypTVAF      : BYTE       ;
                         TTCF         : Byte       ;
                         TypeTpfF     : Byte       ;
                         RemiseF      : Single     ;
                         EscompteF    : Single     ;
                         GroupeF      : Byte       ;
                         ReleveF      : Byte       ;
                         ExemplaireF  : Array [1..3] of Byte ;
                         ExpeditionF  : Byte       ;
                         TransportF   : Byte       ;

                         AcompteF     : Real48       ;
                         AvoirF       : Real48       ;
                         BlNonFactF   : Real48       ;
                         TotEffetF    : Real48       ;
                         SoldeTTCF    : Real48       ;
                         EncoursMaxF  : Real48       ;
                         DateFactF    : Word       ;
                         MontFactF    : Real48       ;
                         DateRegleF   : Word       ;
                         MontRegleF   : Real48       ;
                         DateReserF   : word       ;
                         MontReserF   : Real48       ;
                         BlocF        : Longint    ;
                         ANF          : Single     ;
                         SoldeHTF     : Real48       ;
                         TotRisqueF   : Real48       ;
                         NifCEEF      : String14   ;
                         JunkCEEF     : String14   ;
                         DeviseF      : Byte       ;
                         NiveauF      : Byte       ;
                         HorsMarcheF  : Single     ;
                         JunkF        : String1   );

                         {====== PIECE =====}
                       4:(SoucheP     : byte     ;
                         NumeroP      : Longint  ;
                         Nom2P        : String26 ;
                         Rue2P        : string26 ;
                         Localite2P   : String22 ;
                         CodePostal2P : String8  ;
                         Ville2P      : string22 ;
                         TarifP       : Byte     ;
                         EscompteP    : Single   ;
                         ModeRegleP   : byte     ;
                         VendeurP     : byte     ;
                         ExpeditionP  : Byte     ;
                         TransportP   : Byte     ;
                         TypTVAP      : Byte     ;
                         TypeTpfP     : Byte     ;
                         NbExP        : Byte     ;
                         EffetP       : Word     ;
                         AffaireP     : String13 ;
                         Piece1P      : String20 ;
                         Piece2P      : String20 ;
                         Piece3P      : String20 ;
                         DateP        : Word     ;
                         NatureP      : Byte     ;
                         ClientP      : Longint  ;
                         PortP        : Single   ;
                         TvaPort      : Byte     ;
                         TOTHTPP      : Real48     ;
                         TvaP         : Real48     ;
                         TpfP         : Real48     ;
                         TOTTTCPP     : Real48     ;
                         AcompteP     : Single   ;
                         QteP         : Single   ;
                         DeviseP      : Byte     ;
                         TauxP        : Single   ;
                         EnCaisseP    : Byte     ;
                         AbonnementP  : Byte     ;
                         PeriodiqueP  : Byte     ;
                         MaxRepeteP   : Byte     ;
                         ReconduireP  : Byte     ;
                         DateAboP     : Word     ;
                         RepeteP      : Byte     ;
                         LastAboP     : Word     ;
                         TPFFRSP      : Single   ;
                         ModeTTCP     : Byte     ;
                         ApporteurP   : byte     ;
                         FraisP       : Array[1..10] of single ;
                         FreinteP     : Byte     ;
                         ReleveP      : Byte     ;
                         RefFP        : String10 ;
                         PosP         : Longint  ;
                         PosContre    : Longint  ;
                         ContreMP     : Byte     ;
                         DateBLP      : word     ;
                         PosPieResP   : Longint  ;
                         FlagExportP  : Byte     ;
                         PortExoP     : Single   ;
                         PosDuplicP   : Longint  ;
                         TotPoidsP    : Real48     ;
                         TotVolumeP   : Real48     ;
                         DepartP      : Byte     ;
                         ArriveeP     : Byte     ;
                         RegroupeP    : byte     ;
                         QteTPFP      : Single   ;
                         NomCliP      : String10 ;
                         MessFiscP    : Byte    );

                         {===== LIGNE PIECE =======}
                       5 :
                        (RefCR        : Longint    ;
                         RefAR        : Longint    ;
                         RefPR        : Longint    ;
                         DesignR      : String40   ;
                         PrixDevR     : Single     ;
                         QteR         : Single     ;
                         RemiseR      : Single     ;
                         CommentR     : String20   ;
                         CommR        : Single     ;
                         AffaireR     : String13   ;
                         TvaR         : Byte       ;
                         TpfR         : Byte       ;
                         DepotR       : Byte       ;
                         VendeurR     : Byte       ;
                         TarifR       : Byte       ;
                         DateR        : Word       ;
                         NatureR      : Byte       ;
                         NumeroR      : Longint    ;
                         DeviseR      : Byte       ;
                         DispoRR      : Single     ;
                         TypeRedR     : Byte       ;
                         NextR        : LongInt    ;
                         ContreR      : SmallInt    ;
                         TypeCommR    : byte       ;
                         TypTVAR      : Byte       ;
                         TPFFRSR      : Single     ;
                         TypeTpfR     : Byte       ;
                         ApporteurR   : Byte       ;
                         TypeCalcR    : Byte       ;
                         CoefCommRR   : Single     ;
                         CoefCommAR   : Single     ;
                         ConditionR   : single     ;
                         NoOrdreR     : Byte       ;
                         ReliqQteR    : Single     ;
                         PrixUR       : Real48      );

                         {======== FICHE SAV ========}
                      6 :(RefSAV      : Array [0..2] of Longint  ;
                          NumSAV      : Longint  ;
                          DateSAV     : Word     ;
                          CommentSAV  : String20 ;
                          SerieSAV    : String20 ;
                          PrixSAV     : Single   ;
                          NatSAV      : Byte     ;
                          SoucheSAV   : Byte     ;
                          JunkSAV     : String12 );

                         {======== FICHE POSTAL ========}
                      7 :(CodeI       : String8 ;
                          VilleI      : String22 );

                         {======== LITIGE COMMISSION ========}
                       9 :(EcheLC      : Word ;
                           DateLC      : Word ;
                           MontantLC   : Real48 ;
                           NumEcheLC   : Byte ;
                           Pos14LC     : Longint ) ;

                         {======== FICHE JOURNEE ========}
                      10:(NumJ        : longint    ;
                          dateJ       : Word       ;
                          EcheanceJ   : Word       ;
                          natureJ     : byte       ;
                          qteJ        : Real48       ;
                          prixUJ      : Real48       ;
                          RemiseJ     : Real48       ;
                          prixPAJ     : Real48       ;
                          dispoJJ     : Real48       ;
                          TVAJ        : byte       ;
                          TPFJ        : Byte       ;
                          DepotJ      : byte       ;
                          VendeurJ    : Byte       ;
                          TarifJ      : Byte       ;
                          EscompteJ   : Real48       ;

                          RefAJ       : Longint    ;
                          RefCJ       : Longint    ;
                          HeureJ      : SmallInt    ;
                          CaisseJ     : Byte       ;
                          PassageJ    : byte       ;
                          TPFFRSJ     : Real48       ;
                          UserJ       : Byte       ;
                          PrixTJ      : Real48       ;
                          PrixBaseJ   : Real48       ;
                          CommJ       : Real48       ;
                          TypeCommJ   : Byte       ;
                          SoucheJ     : Byte       ;
                          TypTVAJ     : Byte       ;
                          TypeTpfJ    : Byte       ;
                          ApporteurJ  : Byte       ;
                          AffaireJ    : String13   ;
                          TypeCalcJ   : Byte       ;
                          CoefCommRJ  : Real48       ;
                          CoefCommAJ  : Real48       ;
                          TypeAnouvJ  : Byte       ;
                          DateBlJ     : Word       ;
                          PrixRAJ     : Real48       ;
                          FraisJ      : Real48       ;
                          DeviseJ     : Byte       ;
                          PrixDevUJ   : Real48       ;
                          JunkJ       : String2   );

                         {======== FICHE STAT ========}
                      12:(SorteS      : Byte       ;
                          DepotS      : Byte       ;
                          RefS        : Longint    ;            { pointeur sur la fiche }
                          AnneeS      : Byte       ;
                          PeriodeS    : Byte       ;
                          CaS         : Array[1..4] of Real48) ;
                          { 1: CA Vente/CA Vente/CA Achat  ; 3: Qt‚ Vente/Qt‚ Vente/Qt‚ Achat     }
                          { 2: Marge/Marge/ 0              ; 4: Nb passages/Qt‚ Achat/Nb passages }
                         {======== FICHE SAUVE PIED  ========}
                      14:(DateSD       : word       ;
                          CaisseSD     : byte       ;
                          VendeurSD    : byte       ;
                          DepotSD      : Byte       ;
                          HeureSD      : SmallInt    ;
                          TaxeTVASD    : TabTaxeTvaTpf  ;
                          TaxeTPFSD    : TabTaxeTvaTpf  ;
                          BaseTVASD    : TabTaxeTvaTpf  ;
                          BaseTPFSD    : TabTaxeTvaTpf  ;
                          TTCSD        : Real48       ;
                          HTSD         : Real48       ;
                          PortHTSD     : Single     ;
                          TvaPortSD    : Byte       ;
                          EscompteSD   : Single     ;
                          PEscSD       : Single     ;
                          RefCliSD     : longint    ;
                          NomCliSD     : String19   ;
                          RefPSD       : String10   ;
                          NatureSD     : Byte       ;
                          NumeroSD     : Longint    ;
                          AcompteSD    : Single     ;
                          SoucheSD     : Byte       ;
                          TypTVASD     : Byte       ;
                          TypeTpfSD    : Byte       ;
                          ApporteurSD  : Byte       ;
                          QTESD        : Real48       ;
                          AffaireSD    : String13   ;
                          PassageSD    : Byte       ;
                          ModeRegleSD  : Byte       ;
                          TauxSD       : Single     ;
                          TabRegleSD   : TabRegle   ;
                          LettrageSD   : String13   ;
                          ReleveSD     : Byte       ;
                          PortExoSD    : Single     ;
                          XLitigeSD    : Single     ;
                          DeviseSD     : Byte      );


                         {======== FICHE VOL  ========}
                       16:(SorteL     : Byte       ;
                          CodeL       : Longint    ;
                          QteL        : Single     ;
                          PrixL       : Single     ;
                          DateL       : Word       ;
                          DepotL      : Byte       ;
                          CommentL    : String50   ;
                          FournL      : Longint    ;
                          TypeL       : Byte       ;
                          JunkL       : String11  );

                         {======== FICHE ADRESSE  ========}
                       17:(RaisonLiR    : String30   ;
                          RueLiR        : string26   ;
                          LocaliteLiR   : string22   ;
                          CodePostalLiR : String8    ;
                          VilleLiR      : string22   ;
                          PosClientLiR  : Longint    ;
                          NumAdrLiR     : SmallInt   );

                         {======== FICHE NOMENCLATURE  ========}
                       18:(TabASS     : Array [1..NbMaxAssemble] of LigneASS ;
                          PosArticle  : Longint    ;
                          Typegene    : Byte       ;
                          Junk18      : string12  );

                         {======== FICHE REGLEMENT ========}
                     19:(SorteO         : Byte     ;
                         ModeO          : Byte     ;
                         DateO          : Word     ;
                         EcheO          : Word     ;
                         TotalO         : Real48     ;
                         RefClientO     : longint  ;
                         PassageO       : byte     ;
                         RemisO         : byte     ;
                         BanqueO        : String10 ;
                         EtatO          : Byte     ;
                         LitigeO        : Boolean  ;
                         NumeroO        : String13 ;
                         TabFactO       : Array[1..MaxFactRL] of Longint ;
                         DateReglO      : Word ;
                         DateRemBaO     : Word ;
                         XLitigeO       : Single) ;

                         {======== FICHE REPRESENTANT  ========}
                     20:(NOMW           : STRING25   ;
                         RUEW           : STRING26   ;
                         LOCALITEW      : STRING22   ;
                         CODEPOSTALW    : STRING8    ;
                         VILLEW         : STRING22   ;
                         Comment1W      : String30   ;
                         Comment2W      : String30   ;
                         TELEPHONEW     : STRING15   ;
                         BANQUEW        : STRING20   ;
                         NUMBANQUEW     : STRING5    ;
                         NUMGUICHETW    : STRING5    ;
                         NUMCOMPTEW     : STRING12   ;
                         NUMRIBW        : STRING2    ;
                         SURCAW         : Byte       ;
                         SURFactW       : Byte       ;
                         TypeCalcW      : Byte       ;
                         EnPourcW       : Byte       ;
                         ReducW         : Fiche_Comm ;
                         TrancheW       : Fiche_Comm ;
                         JunkW          : String10  );
                      END ;


descri_fichierNEG = RECORD
                 nom : string30 ;
                 rec_size : SmallInt ;
                 key_size : SmallInt ;
                 duplicate: SmallInt ;
                 END ;

Fiche_CA  = RECORD
            Nom : string50 ;
            CaPast : Real48 ;
            CaMonth : Real48 ;
            Cacumul : Real48 ;
            Code : longint ;
            Date :  Word   ;
            rep : byte ;
            Nb : Real48 ;
            END ;



CONST  fichiersNEG : array[0..Nb_fichierNEG] of descri_fichierNEG =
              ((nom : 'client' ;
                    rec_size : 450 ;
                    key_size : 10  ;             {0}
                    duplicate: 1   ) ,

                    (nom : 'mat' ;
                    rec_size : 900 ;             {1}
                    key_size : 37  ;
                    duplicate: 1   ) ,

                    (nom : 'fourn' ;
                    rec_size : 450 ;             {2}
                    key_size : 10  ;
                    duplicate: 1   ),

                    (nom : 'histoire' ;
                    rec_size : 120 ;             {3}
                    key_size : 15  ;
                    duplicate: 1  ),

                    (nom : 'Piece' ;
                    rec_size : 384 ;             {4}
                    key_size : 17  ;
                    duplicate: 1  ),

                    (nom : 'LIGNER' ;
                    rec_size : 170 ;             {5}
                    key_size : 0  ;
                    duplicate: -1  ),

                    (nom : 'SAV' ;
                    rec_size : 85 ;              {6}
                    key_size : 0  ;
                    duplicate: -1  ),

                    (nom : 'PosTal' ;
                    rec_size : 38 ;              {7}
                    key_size : 0  ;
                    duplicate: -1  ),

                    (nom : 'Graph' ;
                    rec_size : 156 ;              {8}
                    key_size : 0  ;
                    duplicate: -1  ),

                    (nom : 'LITCOMM' ;
                    rec_size : 25 ;              {9}
                    key_size : 0  ;
                    duplicate: -1  ),

                    (nom : 'SAUVJOUR' ;
                    rec_size : 150 ;              {10}
                    key_size : 9  ;
                    duplicate: 1 ),

                    (nom : 'JOURNEE' ;
                    rec_size : 50 ;              {11}
                    key_size : 0  ;
                    duplicate: -1 ),

                    (nom : 'STAT' ;
                    rec_size : 38  ;              {12}
                    key_size : 9   ;
                    duplicate: 0   ),

                    (nom : 'HELP' ;
                    rec_size : 61  ;              {13}
                    key_size : 15  ;
                    duplicate: 1   ),

                    (nom : 'SAUVPIED' ;
                    rec_size : 685  ;             {14}
                    key_size : 9   ;
                    duplicate: 1  ),

                    (nom : 'PIED' ;
                    rec_size : 50 ;              {15}
                    key_size : 0   ;
                    duplicate: -1  ),

                    (nom : 'VOL' ;
                    rec_size : 90  ;              {16}
                    key_size : 0   ;
                    duplicate: -1   ),

                    (nom : 'ADDLIV' ;
                    rec_size : 130 ;              {17}
                    key_size : 8   ;
                    duplicate: 1  ),

                    (nom : 'ASSEMBLE' ;
                    rec_size : 240 ;              {18}
                    key_size : 0  ;
                    duplicate: -1  ),

                    (nom : 'RAPPROCH' ;
                    rec_size : 179  ;              {19}
                    key_size : 0  ;
                    duplicate: -1  ),

                    (nom : 'REPRESEN' ;
                    rec_size : 700  ;             {20}
                    key_size : 0  ;
                    duplicate: -1  ),

                    (nom : 'Amacro' ;
                    rec_size : 170 ;              {21}
                    key_size : 0  ;
                    duplicate: -1  ));


(*=========================================================================================
//=========================== INTEGRATEUR =================================================
//=========================================================================================*)
Const Nb_FichierINT = 14 ;

TYPE  EnregINT = RECORD
                 status : LongInt ;
                 reseau : Boolean ;
                 case sorte : byte of
                         {======== FICHE CLIENT =======}
                 255: ( buffer        : array[1..MaxDataRecSize] of Char ) ;
                  0 : ( NOMC          : STRING[50]   ;
                        CODEC         : STRING[20]   ;
                        CONTACTC      : STRING[50]   ;
                        NATUREC       : Byte       ;
                        RUEC          : STRING[50]   ;
                        LOCALITEC     : STRING[50]   ;
                        CODEPOSTALC   : STRING[8]    ;
                        VILLEC        : STRING[50]   ;
                        SECTEURC      : Byte       ;
                        COMMERCIALC   : Byte       ;
                        TEL1C         : STRING[15]   ;
                        TEL2C         : STRING[15]   ;
                        FAXC          : string[15]   ;
                        MEMOC         : String     ;
                        ARAPPELERC    : TDateTime  ;
                        MOTIFC        : String[100]  ;
                        ORIGINEC      : String[30]   ;
                        EncoursC      : Real48       ;
                        CAHTC         : Real48       ;
                        LIBREC        : Integer    ;
                        COMMENT1C     : String[50]   ;
                        Comment2C     : String[50]   ;
                        DernModifC    : TDateTime ;
                        JunkC         : String[86]  );

                         {======== FICHE ARTICLE ========}
                      1:(NomA         : string[50]  ;
                         REFERENCEA   : String[20]  ;
                         CodeMAgA     : String[20]  ;
                         CodeBarreA   : String[20]  ;
                         RAYONA       : String[20]  ;
                         FAMILLEA     : string[20]  ;
                         SSFAMILLEA   : string[20]  ;
                         PUHTA        : Real48      ;
                         PUTTCA       : Real48      ;
                         TVAA         : Real48      ;
                         JunkA        : String[50]) ;

                         {======= FICHE AFFAIRE =========}
                      2:(RefCliF      : Longint    ;
                         NomF         : String[100]  ;
                         DateDebutF   : TDateTime       ;
                         DateFinF     : TDateTime       ;
                         CodeF        : string[20]   ;
                         CAPotentielF : Real48       ;
                         ObjectifF    : Real48       ;
                         CommercialF  : Byte       ;
                         Arappeler1F  : TDateTime       ;
                         Motif1F      : string[50]   ;
                         Arappeler2F  : TDateTime       ;
                         Motif2F      : string[50]   ;
                         Arappeler3F  : TDateTime       ;
                         Motif3F      : string[50]   ;
                         MemoF        : string     ;
                         NiveauF      : Integer    ;
                         JunkF        : String[70] );

                         {==== Note Interne ===}
                      3:(NomN         : String[50] ;
                         MemoN        : String   ;
                         CategorieN   : Integer  ;
                         NiveauN      : Integer  ;
                         JunkN        : String[50]);

                         {===== ACTIO  ============}
                      4:(REFCLIT      : Longint    ;
                         DateT        : TDateTime  ;
                         NumeroT      : Integer    ;
                         MemoT        : String     ;
                         CommercialT  : Byte       ;
                         TypeActionT  : Integer    ;
                         NomFichierT  : String[12] ;
                         RefAFFT      : Longint    ;
                         NiveauT      : Integer    ;
                         NaTureT      : Byte        ;
                         ArappelerT   : TdateTime  ;
                         MotifT       : String[20] ;
                         JunkT        : String[4] );

                         {======== Type Action=======}
                      5 :(IntituleY   : String[50] ;
                          CheminY     : String[50] ;
                          ClasseFenY  : String[30] ;
                          ModeleDocY  : String[8]  ;
                          ExtensionY  : String[3]  ;
                          ResumeY     : Boolean  ;
                          AdresseY    : Boolean  ;
                          JunkY       : String[50] );

                         {======== CATEGORIE NOTE ======}
                      6 :(NomG        : String[50];
                          NiveauG     : Integer ;
                          JunkG       : String[20] );

                         {======== CONTACT ============}
                      7 :(REFCLIO     : Longint ;
                          ContactO    : String[50] ;
                          FonctionO   : String[50] ;
                          TelephoneO  : String[15] ;
                          FaxO        : String[15] ;
                          JunkO       : String[200]);

                         {======== FACTURE ========}
                       8 :(REFCLIU     : Longint ;
                           DateU       : TDateTime ;
                           REFARTU     : Longint ;
                           REFERENCEU  : String[20] ;
                           COMMERCIALU : Byte  ;
                           PrixU       : Real48 ;
                           QteU        : Real48 ;
                           RemiseU     : Real48 ;
                           TotalU      : Real48 ;
                           PayeONU     : Boolean ;
                           CodeAffaireU: String[13] ;
                           NumPieceU   : Longint ;
                           TVAU        : Real48 ;
                           JunkU       : String[44]  ) ;

                         {======== SOCIETE ========}
                       9:(VersionS    : String[10]   ;
                          NomS        : String[35]   ;
                          Adresse1S   : String[35]   ;
                          Adresse2S   : String[35]   ;
                          CodePostalS : String[8]    ;
                          VilleS      : String[35]   ;
                          CommentaireS: String[200]  ;
                          APES        : String[5]    ;
                          SIRETS      : String[14]   ;
                          NIFS        : String[13]   ;
                          CONTACTS    : String[35]   ;
                          TELEPHONES  : String[15]   ;
                          FAXS        : String[15]   ;
                          TELEXS      : String[15]   ;
                          RepClientS  : String[50]   ;
                          RepModeleS  : String[50]   ;
                          RepGescomS  : String[50]   ;
                          CodeSynapS  : String[50]   ;
                          CreerClientS: Integer ;
                          CreerNoteS  : Integer ;
                          CreerFactS  : Integer ;
                          CreerPassS  : Integer ;
                          CreerTableS : Integer ;
                          CreerCodeTS : Integer ;
                          CreerCodeAS : Integer ;
                          EncoursCalcS: Boolean ;
                          NbLicS      : Integer ;
                          WordAPIS    : Boolean ;
                          NumVersionS : Real48 ;
                          JunkS       : String[193] );

                         {======== UTILISATEUR =======}
                      10:(NomR        : String[50]   ;
                          DateR       : TDateTime    ;
                          PresentONR  : Boolean    ;
                          NiveauR     : Integer    ;
                          PassWordR   : String[50]   ;
                          MemoR       : String[50]  ;
                          FonctionR   : String[50]  ;
                          JunkR       : String[94] );
                         {======== Secteur ======}
                     11 :(NomZ        : String[50];
                          JunkZ       : String[20] );
                         {======== Nature ======}
                     12 :(NomE        : String[50];
                          JunkE       : String[20] );
                         {======== Libre ======}
                     13 :(NomL        : String[50];
                          JunkL       : String[20] );
                         {===Nature Action======}
                     14 :(NomNA        : String[50];
                          JunkNA       : String[20] );

                      END ;


descri_fichierINT = RECORD
                      nom : string[50] ;
                      rec_size : integer ;
                      key_size : integer ;
                      duplicate: integer ;
                      END ;

CONST  fichiersINT : array[0..Nb_fichierINT] of descri_fichierINT =
              ((nom : 'SOC0\client' ;
                    rec_size : 949 ;
                    key_size : 60  ;             {0}
                    duplicate: 1   ) ,

                    (nom : 'SOC0\Article' ;
                    rec_size : 252 ;             {1}
                    key_size : 30  ;
                    duplicate: 1   ) ,

                    (nom : 'SOC0\Affaire' ;
                    rec_size : 667 ;             {2}
                    key_size : 30  ;
                    duplicate: 1   ),

                    (nom : 'SOC0\Note' ;
                    rec_size : 368 ;             {3}
                    key_size : 60  ;
                    duplicate: 1  ),

                    (nom : 'SOC0\Action' ;
                    rec_size : 343 ;             {4}
                    key_size : 15  ;
                    duplicate: 1  ),

                    (nom : 'SOC0\TypeAct' ;
                    rec_size : 205 ;             {5}
                    key_size : 0  ;
                    duplicate: -1  ),

                    (nom : 'SOC0\Categori' ;
                    rec_size : 80 ;              {6}
                    key_size : 0  ;
                    duplicate: -1  ),

                    (nom : 'SOC0\CONTACT' ;
                    rec_size : 345 ;              {7}
                    key_size : 10  ;
                    duplicate: 1  ),

                    (nom : 'SOC0\Facture' ;
                    rec_size : 138 ;              {8}
                    key_size : 16  ;
                    duplicate: 1  ),

                    (nom : 'SOC0\SOCIETE' ;
                    rec_size : 912 ;              {9}
                    key_size : 0  ;
                    duplicate: -1  ),

                    (nom : 'SOC0\UTILISAT' ;
                    rec_size : 316 ;              {10}
                    key_size : 0  ;
                    duplicate: -1),

                    (nom : 'SOC0\SECTEUR' ;
                    rec_size : 80 ;              {11}
                    key_size : 0  ;
                    duplicate: -1),

                    (nom : 'SOC0\NATURE' ;
                    rec_size : 80 ;              {12}
                    key_size : 0  ;
                    duplicate: -1),

                    (nom : 'SOC0\LIBRE' ;
                    rec_size : 80 ;              {13}
                    key_size : 0  ;
                    duplicate: -1),

                    (nom : 'SOC0NATACT' ;
                    rec_size : 80 ;              {14}
                    key_size : 0  ;
                    duplicate: -1));

(*//=========================================================================================
//=========================================================================================
//=========================================================================================*)


  NbMaxTypePiece = 19 ;
  MaxCptInt = 11 ;
  MaxVentilCpt=14 ;
  CpteInt=1 ;
  Ventil=2 ;
  BilanCpte=3 ;
  BilanRubr=4 ;
  BudCpte=5 ;
  BudRubr=6 ;
  TresoCpte=7 ;
  TvaCpte=8 ;
  TvaRub=9 ;

Type TNomTable = ARRAY[1..19] of string10 ;
Type TCarac = ARRAY[1..5] of byte ;
Type TPlage = RECORD
              Deb : SmallInt ;
              Fin : SmallInt ;
              END ;
Type TFourchette = RECORD
                   Deb : String17 ;
                   Fin : String17 ;
                   END ;
TFourchCli = ARRAY[1..4] of Tfourchette ;
TFourchFou = ARRAY[1..4] of Tfourchette ;

Type
TFormatCpt = RECORD
             Lg : byte ;
             Desc : string13 ;
             Cadrage : byte ;
             BourEntier : string1 ;
             BourAlpha : string1 ;
             BourLettre : string1 ;
             END ;

TCaracDefaut = ARRAY[1..2,1..3] of byte ;

TFourDate = RECORD
            Deb : Word ;
            Fin : Word ;
            END ;

TCaracLettrage = RECORD
                 JalReg       : String3 ;
                 CptRegDeb    : String17 ;
                 CptRegCred   : String17 ;
                 MtantRegDeb  : Real48 ;
                 MtantRegCred : Real48 ;
                 DiffChangeL  : String17 ;
                 DiffChangeP  : String17 ;
                 END ;

TParamDossier = RECORD
                NomCollabo  : String17 ;
                CodeDossier : String17 ;
                RaisonExpert: String40 ;
                END ;

{====PARAMETRES SOCIETE =======}
TParamSoc = RECORD
            CodeSociete  : String17 ;
            PassWord     : String6  ;
            RaisonSoc    : String40 ;
            Adresse1     : String35 ;
            Adresse2     : String35 ;
            Adresse3     : String35 ;
            Localite     : String35 ;
            CodePostal   : String15 ;
            Ville        : String35 ;
            Pays         : Byte     ;
            FormeJur     : Byte     ;
            NoCommun     : String20 ;
            Ape          : String20 ;
            LgCollectif  : Byte     ;
            MultiColl    : Byte     ;
            Devise       : Byte     ;
            NbDecDevise  : Byte     ;
            NbExoConserv : Byte     ;
            DernPiece    : Longint  ;
            {
            DernLettrage : Longint  ;
            }
            CtrlDoublonImport : Byte ;
            SoldeSolde   : Byte ; { Pour M9 pour l'instant }
            Libre        : Array[1..2] Of Byte ;
            MultiEcheOn  : Byte     ;
            AnalytiqueOn : Byte     ;
            QuantiteOn   : Byte     ;
            DevisesOn    : Byte     ;
            FormatGen    : TFormatCpt ;
            FormatAux    : TFormatCpt ;
            FormatAna    : TFormatCpt ;
            Cli          : TFourchCli ;
            Fou          : TFourchFou ;
            Imm          : TFourchette ;
            Sal          : TFourchette ;
            Ban          : TFourchette ;
            Cai          : TFourchette ;
            Cha          : TFourchette ;
            Pro          : TFourchette ;
            CaracDefaut  : TCaracDefaut ;
            GeneAttend   : String17 ;
            ClientAttend : String17 ;
            FournAttend  : String17 ;
            AutreAttend  : String17 ;
            SectionAttend: String17 ;
            OuvreBilan   : String17 ;
            Resultat     : String17 ;
            Benefice     : String17 ;
            Perte        : String17 ;
            EnCours      : TFourDate ;
            Suivant      : TFourDate ;
            DateExo      : Array[1..20] Of TFourDate ;
            DevisePivot  : Byte ;
            DernOdAnal   : LongInt ;
            An2000       : Byte ; { POUR DE VRAI !!! }
            DernANouv    : Longint ;
            GlAuxSurDate : Byte ; { Avant AvecTrait, passe en seeknum }
            DetailANouv  : Byte ;
            CtrlBal      : Byte ;
            DateUtil     : Word ;
            ImportAuto   : Byte ;
            ChangeJour   : Byte ;
            DernClo      : LongInt ;
            NegPath      : String[48] ;
            NegSoc       : SmallInt ;
            TypSaisieUnique : Byte ;
            CptesBilan   : TFourchette ;
            CptesGestion : TFourchette ;
            VerifRIB     : Byte ;
            NomSNEXE     : String8 ;
            BudGetReglement : Real48 ;
            EmetNat      : String6 ;
            ModifEcheNeg : Byte ;
            ExportLetNeg : Byte ;
            HistoDevise  : TFourDate ;
            OffSetDev    : Byte ;
            Canabis      : String[5] ; {identifiant devise pivot}
            LetCli       : TCaracLettrage ;
            LetFou       : TCaracLettrage ;
            LetAutre     : TCaracLettrage ;
            DernLot      : LongInt ;
            DernAbo      : LongInt ;
            DernSim      : LongInt ;
            PlageCpta    : TPlage  ;
            PlageEche    : TPlage  ;
            DernTyp      : LongInt ;
            PlanRef      : Byte ;
            Mouchard     : Byte ;
            ChangeColl   : Byte ;
            MontantNegatif: Byte ;
            Virgule      : String1 ;
            JalFerme     : String3 ;
            JalOuvre     : String3 ;
            SaisieUnique : Byte ;
            Glandu       : Byte ;
            EchUn        : Byte ;
            AnaUn        : Byte ;
            CreatUn      : Byte ;
            AvecBandeau  : Byte ;
            SaisieRafale : Byte ;
            PreImprime   : Byte ;
            CrMulti      : Byte ;
            EcrRapide    : Byte ;
            AnaRapide    : Byte ;
            EchRapide    : Byte ;
            MajCumLot    : Byte ;
            AnalOk       : Byte ; {Utilis‚ qu'en ACCESS, sinon ModulAnalyse}
            LettEnSaisie : Byte ;
            NbMaxImprime : SmallInt ;
            GLBilan      : Byte ;
            DeczUnique   : Byte ;
            ReelEgalLot  : Byte ;
            WithSim      : Byte ;
            CritGuide    : Byte ;
            TvaEncVente  : Byte ;
            TvaEncAchat  : Byte ;
            EncSurLett   : Byte ;
            Methode      : Byte ;
            ForceEDLEG   : Byte ;
            SecteurDefaut: Byte ;
            ValSecteur   : Byte ;
            JunkGG       : Byte ;
            END ;

Const
  FileHeaderSize = 14 ;
  MinDataRecSize = FileHeaderSize;

type
  DataFile  =  record
                 F          : file;
                 FirstFree,
                 NumberFree,
                 Int1       : LongInt;
                 ItemSize   : word;
                 NumRec     : LongInt;
               end;
  TaKeyStr  =  string[MaxKeyLen];
  TaItem    =  record
                 DataRef, PageRef : LongInt;
                 Key : TaKeyStr;
               end;
  TaPage    =  record
                 ItemsOnPage : 0..PageSize;
                 BckwPageRef : LongInt;
                 ItemArray   : array[1..PageSize] of TaItem;
               end;
  TaPagePtr =  ^TaPage;
  TaSearchStep =
               record
                 PageRef,ItemArrIndex : LongInt;
               end;
  TaPath    =  array[1..MaxHeight] of TaSearchStep;
  IndexFile =  record
                 DataF         : DataFile;
                 AllowDuplKeys : Boolean;
                 KeyL          : byte;
                 RR,PP         : LongInt;
                 Path          : TaPath;
                 DelAdd        : Boolean ;
               end;
  IndexFilePtr = ^IndexFile;
  TaStackRec = record
                 Page      : TaPage;
                 IndexFPtr : IndexFilePtr;
                 PageRef   : LongInt;
                 Updated   : Boolean;
               end;
  TaStackRecPtr = ^TaStackRec;
  TaPageStack = array[1..PageStackSize] of TaStackRec;
  TaPageStackPtr = ^TaPageStack;

  TaPageMap  =  array[1..PageStackSize] of word;
  TaPageMapPtr = ^TaPageMap;

  TaRecordBuffer  =
               record
                 case SmallInt of
                   0 : (Page : TaStackRec);
                   1 : (R : array[1..MaxDataRecSize] of Byte);
                   2 : (I : LongInt);
               end;
  TaRecordBufPtr = ^TaRecordBuffer;

  FileAsciiZ = array[1..64] of char;
  FileName = string[66];

const
  ItemOverhead = 9; { SizeOf(TAItem) - MaxKeyLen }
  PageOverhead = 5; { SizeOf(TaPage.ItemsOnPage) +
                      SizeOf(TaPage.BckwPageRef) }

  NoDuplicates = 0;
  Duplicates = 1;

  RecTooLarge          = 1000;
  RecTooSmall          = 1001;
  KeyTooLarge          = 1002;
  RecSizeMismatch      = 1003;
  KeySizeMismatch      = 1004;
  MemOverflow          = 1005;

Type ProcPtr = ^byte;

Type BNCPTA = RECORD
                Status : Longint ;
                LeReseau : Byte ;
                Sorte  : SmallInt ;
                Lignes : LigBN ;
                END ;

type NomSoc = RECORD
              Nom  : String13 ;
              Path : String85 ;
              END ;

Type TPerDev = Array[1..31] of Real48 ;

Type CalcModePaie = Record
                    Origine  : Byte ;
                    NbJours  : Byte ;
                    Arrivee  : Byte ;
                    NbEche   : Byte ;
                    Decal    : Byte ;
                    Remp1    : Byte ;
                    Remp2    : Byte ;
                    Paiement : Array[1..12] of Byte ;
                    Montant2 : Single ;
                    Montant1 : Real48 ;
                    OkEsc    : Boolean ;
                    Titre    : String60 ;
                    Pourcent : Array[1..12] of Single ;
                    END ;

ListeInt  = Array [1..ListeMaxZone] of SmallInt ;
ListeStr  = Array [1..ListeMaxZone] of String40 ;
ListeTab  = Array [1..ListeMaxLigne] of String200 ;

Fiche_Liste = RECORD
              sorte : SmallInt ;
              titre : String30 ;
              TitreUnun : String200 ;
              CZ,PZ,LZ : ListeInt ;
              FZ,TZ,JZ : ListeStr ;
              END ;

Const MaxNbColonne = 4 ;

Type TabExo  = Array[1..MaxNbColonne] of Byte ;

Type TFormat = RECORD
               TypeFormat,ValSelectZone,Decim,DecalB : SmallInt ;
               Vide,DesiListe,FormatLDate : Boolean ;
               SelectZone,DesiDebit,ZZZone : Byte ;
               END ;

Type TColonne = RECORD
                Tipe   : Byte ;
                Valeur : String30 ;
                END ;

Type TBilan   = RECORD
                Status   : Longint  ;
                Sorte    : Byte     ;
                Code     : String4  ;
                Intitule : String40 ;
                Colonnes : Array[1..MaxNbColonne] of TColonne ;
                JZ       : Array[1..MaxNbColonne+2] of String40 ;
                END ;

Type LaVariableSAARI =
     RECORD
     ParaSoc : TParamSoc ;
     OkDecV,OkDecQ : Byte ;
     datff : array [0..nb_fichierNEG] of datafile ;
     idxff : array [0..nb_fichierNEG] of indexfile ;
     DATFTVA,DatFGuide : DataFile  ;
     IDXFTVA,IdxFGuide : IndexFile ;
     FichierChCh : File of String13 ;
     DatfBN : Datafile ;
     ok : Boolean ;
     WarnIndex,InHandler : boolean;
     TurboResult,NumSoc : SmallInt ;
     TAStatus  : word;
     FMessage : file ;
     FichierNum : file of Real48 ;
     FichierDev : File Of TPerDev ;
     FichierRegle : File of CalcModePaie ;
     CurExercice : SmallInt ;
     NomSociete : String ;
     LastDate,CloturePer,CloturePro,ClotureTVA,LastPurge : Word ;
     PZ,CZ,LZ : ListeInt ;
     TitreListe,TitreUnUn : String40 ;
     JZ,FZ,TZ : ListeStr ;
     SorteGeneral,NumListe : LongInt ;
     TotZone : SmallInt ;
     QuelExport : TSaaProg ; 
     END ;


Var  TAErrorProc : ProcPtr;
     VSAA : ^LaVariableSaari ;

{Variables negoce}
Var   AnneeCur,CurMonth,DebutExercice : Smallint ;
      OkdecQ,OkdecV,OkdecP : smallint ;

implementation


INITIALIZATION
New(VSAA) ;

{$IFDEF WIN32}
FINALIZATION
Dispose(VSAA) ;
{$A+,H+}
{$ENDIF}

end.
{$A+}

