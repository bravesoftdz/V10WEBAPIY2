unit CRITEDT ;
{
PT1 17/03/2008 GGU V_9 Suppression des paramètres Nature ducs et Etablissements
}

interface

Uses
  CPTypeCons,
  HENT1, STDCTRLS, Ent1 ;

Const MaxColEdt = 20 ;
Const CMaxPortrait : Integer =745 ;
Const MaxNbRep = 13 ; //7 ; Rony --> pour Budget
Const MaxNbRepBud = 13 ;
Const GridPCL : Boolean = FALSE ;

Type TNatureEtat = (neBal,neGL,neJU,neBro,neEch,neGlV,neJal,neCum,nePoi,neRap,neJBq,neBrouAna,neJalR,
                    neBalBud,neGlBud,neJalBud,neBalBudRea, neBroBud, neBroBudAna,neBalLib,nePlanEdt) ;
Type TNatureEtatJal = (nejalRien,neJalPer,neJalCentr,neJaG) ;
Type TModeSelect = (msRien,msGenEcr,msAuxEcr,msGenAna,msSecAna,msGenBEcr,msGenBAna,msSecBAna) ;
Type TNatureBase = (nbGen,nbAux,nbSec,nbBudGen,nbBudSec, nbBudJal, nbGenT,nbJal,nbJalAna,nbBudAna,nbLibEcr,nbLibAna) ;
Type TNatureEtatTiers = (Age,Ventile) ;

Type TRuptEtat = (Sans,Avec,Sur) ;

Type TTotRupt = Array[0..12] of Double ;

Type TQuelAN = (AvecAN,SansAN,QueAN) ;

Type TQuelEtatBud = (Normal,Rupture,Nature) ;

Type SetOfByte = Set Of Byte ;

Type TReport = RECORD
               TotDebit  : Double ;
               TotCredit : Double ;
               Active    : Boolean ;
               END ;

Type TabTRep = Array[1..MaxNbRep] Of TReport ;
Type TabTRep13 = Array[1..MaxNbRepBud] Of TReport ;


Type TabReport = RECORD
                 OkAff : Boolean ;
                 R : Array[1..10] Of TabTRep ;
                 END ;

Type TCollAff =   RECORD
                  OkAff : Boolean ;
                  Tag   : Array[1..2] Of Integer ; { Débit / crédit }
                  Left  : Array[1..2] Of Integer ; { pour décalage des composants }
                  Width : Array[1..2] Of Integer ;
                  END ;

Type TTabCollAff = Array[1..6] Of TCollAff ; { Les 2+4=6 colonnes de la balances }

Type TFormatPrint = RECORD
                    PrColl,PrSepMontant,PrSepCompte,PrSepRupt,PrSepCompte2 : Boolean ;
                    TabColl : TTabCollAff ;
                    Report  : TabReport ;
                    END ;



Type TCollAffEdt =   RECORD
                     OkAff : Boolean ;
                     Tag   : Integer ;
                     Left  : Integer ;
                     Width : Integer ;
                   END ;

Type TTabCollAffEdt = Array[1..maxColEdt] Of TCollAffEdt ;

Type TModePrint = (PAutomatique,PPortrait,PPaysage) ;

Type TFormatPrintEdt = RECORD
                      PrColl,PrSepMontant : Boolean ;
                      prSepCompte : Array[1..10] Of Boolean ;
                      TabColl : TTabCollAffEdt ;
                      Report  : TabReport ;
                      ModePrint : TModePrint ;
                      ColMin,ColMax : Integer ;
                      END ;

Const MaxTranslate=20 ;

Type TRappelCrit = (SansRappel,AvecRappel,QueRappel) ;

Type TQuelleRupt = (rRien,rRuptures,rLibres,rCorresp) ;

Type TEtat = (etBalGen,etBalAux,etBalAna,etBalGenAux,etBalGenAna,etBalAuxGen,etBalAnaGen,etBalGeneComp, // GC
              etGlGen,etGlAux,etGlAna,etGlGenAux,etGlGenAna,etGlAuxGen,etGlAnaGen,
              etBalVen,etBalAge,etGlVen,etGlAge,etEche,etJuSold,etJustBq,
              etBalAno,etBalClo,etRappro,etPointage,etSynPointage,
              etCumulGen,etCumulAux,etCumulAna,etJalPer,etJalCentr,etJalGen,
              etJalDiv,etJalDivGen, etJalAno, etJalClo, etJalDivPCL, // GC - 20/12/2001
              etBlBudgeteGen,etBlBudgeteAna,etBlBudgeteGenAna,etBlBudgeteAnaGen,
              etBlBudEcartGen,etBlBudEcartAna,etBlBudEcartGenAna,etBlBudEcartAnaGen,
              etBrou,etBrouAna,etBrouBud,etBrouBudAna, etBalGenCerg,etBalLibEcr,etBalLibAna,
              etBalHTAG,etBalHTGA,etTvaAcc,etTvaAut,etTvaFac,etTvaHT,etBalAgeDev,etBudMulti) ;

Type TEdition = RECORD
                 Etat : TEtat ;
                 END ;

type TCritEdtChaine = RECORD
     Utiliser         : Boolean;
     NatureFiche      : string;  // Nature de la fiche CP, AM, PG, etc...
     NomFiche         : string;  // Nom de la fiche
     NomFiltre        : string;  // Clé dans la table FILTRES
     NatureEtat       : string;  // Nature de l'état
     CodeEtat         : string;  // Nom de l'état
     NombreExemplaire : Integer; //
     FiltreUtilise    : string;  // Filtre à charger au lancement
     AuFormatPDF      : Boolean; // Imprimer dans un fichier PDF
     NomPDF           : string;  // Nom du PDF à générer
     MultiPdf         : Boolean; // Créer un PDF par dossier
     ///////////////////////////////
     // Critères standard d'éditions
     UtiliseCritStd   : Boolean;
     Exercice         : TExoDate;
     ModeSelection    : string;
     TypeEcriture     : string;
     NatureCompte     : string;
     ///////////////////////////////
     //Paramètres de la paie
     PGDateDeb,PGDatefin : TDateTime;
//PT1     PGNatDeb,PGNatFin,PGEtabDeb,PGEtabFin  : string; // MF 28/03/2007 pour impression des DUCS
     ParametresEtat : String; //GGU 13/03/2008 Pour la saisie des parameters lors du lancement des etats chaînés
   end;

  ClassCritEdtChaine = Class
    CritEdtChaine   : TCritEdtChaine;
    CtxEtatsChaines : Boolean;
                       end;

Type TCritEdt = RECORD
                Cpt1,Cpt2,SCpt1,sCpt2 : String17 ;
                LCpt1,LCpt2,LSCpt1,LsCpt2 : String ;
                DateDeb,DateFin : TDateTime ;
                Date1,Date2 : TDateTime ;
                Exo : TExoDate ;
                // GC
                AvecComparatif : Boolean ; // Balance Comparative
                ExoComparatif : TExoDate ; // Balance Comparative
                CompareBalSit : boolean;   // Balance de situation
                BalSit : string;           // Balance de situation
                // FIN - GC

                // GCO - 29/09/2008
                Journal       : HString;   //
                NatureJournal : HString;   //
                // FIN GCO

                DeviseSelect,DeviseAffichee,Etab,Valide  : String3 ;
                NatureCpt : String ; // FQ 15844 SBO 27/09/2005
                Qualifpiece : String ;
                Joker, SJoker, AfficheSymbole : Boolean ;
                MonoExo  : Boolean ;
                Monnaie  : Byte ; { 0 = Francs, 1 = Devise, 2 = Euro }
                SautPage, SoldeProg, TotMens : Byte ; { 0 = Selon la Nature du Cpt ; 1 = Oui ; 2 = Non }
                Decimale : Integer ;
                DecimalePivot : Integer ; { pour cumul périodique uniquement }
                Symbole : String3 ;
                SymbolePivot : String3 ; { pour cumul périodique uniquement }
                FormatMontant, RefInterne    : String ;
                FormatMontantPivot : String ; { pour cumul périodique uniquement }
                SQLPlus,SQLPlusBase : String ;
                DEVPourExist : String3 ;
                GA      : Boolean ;
                SQLGA   : String ;
                OnCumul : Boolean ;
                Composite : Boolean ;
                ModeRevision : TCheckBoxState ;
                RappelCrit : TRappelCrit ;
                RuptTabLibre : Boolean ;
                LibreTrie : String ;
                SqlTabLibreTrie : String ;
                LibreCodes1  : String ;
                LibreCodes2  : String ;
                Rupture      : TQuelleRupt ;
                Complement   : Boolean ;
                Cptlibre1, Cptlibre2, Cptlibre3, CptLibre4,
                Cptlibre11, Cptlibre21, Cptlibre31, CptLibre41   : string17 ;
                NbLibreTrie : Integer ;
                SautPageRupt : Boolean ;
                FiltreSup : String ;
                SautPageTri : Boolean ;
                ReferenceInterne : string; // GCO
                Case NatureEtat : tNatureEtat Of
                  neBal : ( Bal : Record
                                  FormatPrint  : TFormatPrint ;
                                  CodeRupt1,CodeRupt2 : string17 ;
                                  Date11,Date21 : TDateTime ;
                                  TypCpt : Byte ; { 0 : Exo, 1 = Tous, 2 = Cpt soldés, 3 = Période }
                                  AnalPur, Axe : String3 ;
                                  Sauf,SqlSauf,SSauf,SSqlSauf,PlanRupt : String[255] ;
                                  RuptOnly : TRuptEtat ;
                                  ZoomBalRub : Boolean ;
                                  TotalBilan, BilGes : Boolean ;
                                  QuelAN : TQuelAN ;
                                  PlansCorresp : Byte ;
//                                  OnlyCptCorresp : Boolean ;
                                  OnlyCptAssocie : Boolean ;
                                  PourCloture : Boolean ;
                                  AvecCptSecond  : Boolean ;
                                  SansAnoANA : Boolean ;
                                  SpeedOk : Boolean ;
                                  VoirEcart : Byte ; { 0 : Tous, 1 : Sans écart, 2 : Que les écarts }
                                  STTS : String[255] ; // Pour tri champ sup
                                  OkRegroup : Boolean ;
                                  StRegroup : String[255] ;
                                  CycleDeRevision : string[255];
                                  End ;
                          ) ;
                  neGL  : ( GL  : Record
                                  FormatPrint  : TFormatPrintEDT ;
                                  CodeRupt1,CodeRupt2 : string17 ;
                                  Date11,Date21 : TDateTime ;
                                  TypCpt : Byte ; { 0 : Exo, 1 = Tous, 2 = Cpt soldés, 3 = Période }
                                  AnalPur, Axe : String3 ;
                                  NumPiece1, NumPiece2 : Integer ;
                                  Sauf,SqlSauf,SSauf,SSqlSauf,PlanRupt : String[255] ;
                                  RuptOnly : TRuptEtat ;
                                  ForceNonCentralisable, TotEche : Boolean ;
                                  // GC - 17/01/2002
                                  AvecMemoEcr : Boolean;     // Modif pour Grand livre Auxiliaire
                                  AvecDetailCentralise : Boolean; // Modif pour Grand livre Auxiliaire
                                  // FIN GC
                                  TotalBilan, BilGes : Boolean ;
                                  QuelAN : TQuelAN ;
                                  SansDetailCentralisation : Boolean ;
                                  Lettrable, Pointable : Byte ; { 0 = Les deux, 1 = Oui, 2 = Non }
                                  OnlyCptAssocie : Boolean ;
                                  PlansCorresp : Byte ;
                                  Encaissement : Byte ; { Pour état de contrôle TVA }
                                  Deductible   : Boolean ; { Pour état de contrôle TVA }
                                  NumTVA : Integer ; { Pour état de contrôle TVA }
                                  CodeTVA : String[3] ; { Pour état de contrôle TVA }
                                  RegimeTVA : String[3] ; { Pour état de contrôle TVA }
                                  LibRegimeTva : String[35] ; { Pour état de contrôle TVA }
                                  PrevalidTVA,ValidTVA : Boolean ; { Pour état de contrôle TVA }
                                  SansAnoANA : Boolean ;
                                  EnDateSituation : Boolean; // GCO - 17/02/2004
                                  DetailAno : Boolean;       // GCO - 15/09/2005
                                  CycleDeRevision : string[255];
                                  End ;
                          ) ;
                  neJU  : ( JU :  Record
                                  FormatPrint : TFormatPrintEDT ;
                                  EnSituation : Boolean ;
                                  TriePar     : Byte ; { 0 = Code, 1 = Libellé}
                                  sens        : Byte ; { 0 = Debit, 1 = Crédit, 2 = Mixte }
                                  Lettrage    : Integer ; {0 = Tous, 1 = Lettré, 2 = Non lettré }
                                  OnTiers     : Boolean ;
                                  RuptOnly    : TRuptEtat ;
                                  PlansCorresp : Byte ;
//                                  OnlyCptCorresp : Boolean ;
                                  OnlyCptAssocie : Boolean ;
                                  PlanRupt       : String[255] ;
                                  PasDeCollectifAAfficher : Boolean ;
                                  SaufCptSolde   : Boolean ;
                                  SansContrepartie : Boolean ;
                                  CodeRupt1,CodeRupt2 : string17 ;
                                  end ;
                          ) ;
                  neBro : (   { Jal1->Cpt1 ; Jal2->Cpt2 ; Util1->SCpt1 ; Util2->Scpt2}
                            Bro : Record
                                  FormatPrint : TFormatPrintEdt ;
                                  DateCreat1, DateCreat2  : TDateTime ;
                                  AfficheAnal, Tri,TriJal,AfficheTauxEuro : Boolean ;
                                  NumPiece1, NumPiece2      : Integer ;
                                  RefInterne     : String[255] ;
                                  Axe : String3 ;
                                  AvecJalAN   : Boolean ;
                                  User : String[3] ;
                                  TP1,TP2 : String[17] ; // Pour tiers payeur
                                  TPAchat : Boolean ;
                                  end ;
                          ) ;
                  neEch : ( Ech : Record
                                  FormatPrint  : TFormatPrintEDT ;
                                  Cpt3,Cpt4    : String17 ;
                                  SSjoker      : Boolean  ;
                                  ModePaie     : String3 ;
                                  RuptPaie     : Boolean ;
                                  Edition      : Byte ; { 0 = Mvt   , 1 = Mvt+Cumul Eche , 3 = Cumul / Aux, 4 = Cumul / Eche,  5= Mvt+Cumul Aux}
                                  sens         : Byte ; { 0 = Debit, 1 = Crédit, 2 = Mixte }
                                  Echeancier   : Byte ; { 0 = Tous, 1 = Encaissement, 2 = Décaissement }
                                  Lettrage     : Integer ;
                                  NatureSCpt   : String3 ;
                                  RecapMP      : Boolean ;
                                  NewPageRecapMP   : Boolean ;
                                  LCpt3,LCpt4  : String17 ;
                                  Jal1,Jal2    : String[3] ;
                                  AvecLibEcr   : Boolean ; 
                                  end ;
                          ) ;
                  neGlV : ( GlV : Record
                                  FormatPrint : TFormatPrintEDT ;
                                  TriePar     : Byte ; { 0 = Code, 1 = Libellé}
                                  sens        : Byte ; { 0 = Debit, 1 = Crédit, 2 = Mixte }
                                  TypePar     : Byte ; { 0 = Auxiliaire, 1 = Paiement }
                                  TypeGL      : Byte ; { 0 = Prévision, 1 = Retard }
                                  Ecart       : Integer ;
                                  ChoixEcart  : Integer ; { 0 : Jour, 1 : Période }
                                  SQLModePaie,PlanRupt : String[255] ;
                                  RuptOnly    : TRuptEtat ;
                                  CodeRupt1,CodeRupt2 : string17 ;
                                  PlansCorresp : Byte ;
//                                  OnlyCptCorresp : Boolean ;
                                  OnlyCptAssocie : Boolean ;
                                  SoldeFormate   : String3 ;
                                  Resol          : hChar ;
                                  DeviseEtPivot  : Boolean ;
                                  MultiDevise    : Boolean ;
                                  Credit         : Boolean ;
                                  SaufCptSolde   : Boolean ;
                                  AfficheLigneAZero : Boolean ;
                                  TP1,TP2 : String[17] ; // Pour tiers payeur
                                  end ;
                          ) ;
                  neJal : ( Jal : Record
                                  FormatPrint : TFormatPrintEdt ;
                                  AvecDevise  : Boolean ;
                                  NumPiece1, NumPiece2 : Integer ;
                                  Tri         : Boolean ; { True : par Date comptable ; False : par N° pièce }
                                  AvecJalAN   : Boolean ;
                                  AfficheTauxEuro : Boolean ;
                                  // GC - 20/12/2001
                                  VoirLibelleCompte : Boolean ;
                                  AvecNouvFct : Boolean;
                                  // GC - 20/12/2001 - FIN
                                  end ;
                          ) ;
                  neJalR : ( JalR : Record
                                    FormatPrint : TFormatPrintEdt ;
                                    AvecDevise  : Boolean ;
                                    NumPiece1, NumPiece2 : Integer ;
                                    AvecJalAN   : Boolean ;
                                    end ;
                          ) ;
                          (*
                  neJaG : ( { DeviseFrancOuEuro->DeviseSelect ; DeviseSup->DeviseAffiche }
                            JaG : Record
                                  FormatPrint : TFormatPrint ;
                                  AvecDevise  : Boolean ;
                                  NumPiece1, NumPiece2, DecimaleDev : Integer ;
                                  FormatMontantDevise : String[255] ;
                                  Tri         : Boolean ; { True : par Date comptable ; False : par N° pièce }
                                  end ;
                          ) ;
                          *)
                  neCum : ( Cum : Record
                                  FormatPrint  : TFormatPrint ;
                                  TypCpt : Byte ; { 0 : Exo, 1 = Tous, 2 = Cpt soldés, 3 = Période }
                                  AnalPur, Axe : String3 ;
                                  Date11,Date21 : TDateTime ;
                                  FormatMontantDevise : String[255] ;
                                  AvecDevise   : Boolean ;
                                  Sauf,SqlSauf,PlanRupt : String[255] ;
                                  CalculAN : Boolean ;
                                  DeuxMontant : Boolean ;
                                  AvecJalAN   : Boolean ;
                                  RecapMens   : Boolean ;
                                  end ;
                            ) ;
                  nePoi : ( Poi : Record
                                  FormatPrint : TFormatPrintEdt ;
                                  Banque      : String3 ;
                                  RefP1,RefP2 : String17 ;
                                  QueLesBQue  : Boolean ;
                                  SurUnCompte : Boolean ;
                                  TypSelectPointe : Byte ; {0 : Tous 1 : Pointés 2 : Non pointés}
                                  end ;
                            ) ;
                  neRap : ( Rap : Record
                                  FormatPrint  : TFormatPrintEdt ;
                                  RefP         : String17 ;
                                  Sens         : Byte ; { 0 : Comptable, 1 : Bancaire }
                                  Tout          : Boolean ; { Force l'édition dans le cas outout est pointé }
                                  end ;
                            ) ;
                  neJBq : ( JBq : Record
                                  FormatPrint  : TFormatPrintEdt ;
                                  Banque      : String3 ;
                                  RefP1,RefP2 : String17 ;
                                  end ;
                            ) ;
                  neBrouAna : (   { Jal1->Cpt1 ; Jal2->Cpt2 ; Util1->SCpt1 ; Util2->Scpt2}
                                BrouAna : Record
                                          FormatPrint : TFormatPrintEdt ;
                                          DateCreat1, DateCreat2  : TDateTime ;
                                          AfficheAnal, Tri : Boolean ;
                                          NumPiece1, NumPiece2      : Integer ;
                                          RefInterne     : String[255] ;
                                          Axe : String3 ;
                                          end ;
                              ) ;
                  neBalBud : ( BalBud : Record
                                  FormatPrint  : TFormatPrintEDT ;
                                  CodeRupt1,CodeRupt2 : string17 ;
                                  Date11,Date21 : TDateTime ;
                                  TypCpt : Byte ; { 0 : Exo, 1 = Tous, 2 = Cpt soldés, 3 = Période }
                                  AnalPur, Axe, SoldeFormate : String3 ;
                                  Sauf,SqlSauf,SSauf,SSqlSauf,PlanRupt : String[255] ;
                                  RuptOnly     : TRuptEtat ;
                                  ZoomBalRub   : Boolean ;
                                  ColonnageJal : boolean ;
                                  QuelAN       : TQuelAN ;
                                  Taux         : Double ;
                                  Exo1, Exo2,MvtSur : String3 ;
                                  Journal, Categorie : String3 ;
                                  // FQ 16536 SBO 18/10/2005 Mise en place des options de révision dans les balances budgétaires
                                  NatureBud : String[255] ;
                                  NatureLib : String[255] ;
                                  // Fin FQ 16536
                                  Resol        : hChar ;
                                  Realise, AnaGene, Qte, TotIniRev,QueTotalBud,QueTotalRea : Boolean ;
                                  OnlySecRupt, AvecCptSecond  : Boolean ;
                                  OnlyLibre    : Boolean ;
                                  OnlyCptAssocie : Boolean ;
                                  SANBud : String[10] ;
                                  PasDeRubrique : Boolean ;
                                  End ;
                          ) ;
                  neBroBud : (
                                BroBud : Record
                                FormatPrint : TFormatPrintEdt ;
                                DateCreat1, DateCreat2  : TDateTime ;
                                AfficheAnal, Tri : Boolean ;
                                NumPiece1, NumPiece2      : Integer ;
                                RefInterne     : String[255] ;
                                Axe : String3 ;
                                Exo1, Exo2 : String3 ;
                                Journal, NatureBud : String3 ;
                                end ;
                              ) ;
                  neBroBudAna : (
                                BroBudAna : Record
                                FormatPrint : TFormatPrintEdt ;
                                DateCreat1, DateCreat2  : TDateTime ;
                                AfficheAnal, Tri : Boolean ;
                                NumPiece1, NumPiece2      : Integer ;
                                RefInterne     : String[255] ;
                                Axe : String3 ;
                                Exo1, Exo2 : String3 ;
                                Journal, NatureBud : String3 ;
                                end ;
                              ) ;
                  neBalLib : ( BalLib : Record
                                  FormatPrint  : TFormatPrint ;
                                  Date11,Date21 : TDateTime ;
                                  AnalPur, Axe : String3 ;
                                  Sauf,SqlSauf,SSauf,SSqlSauf : String[255] ;
                                  NbCol : Array[1..4] of boolean ;
                                  End ;
                             ) ;
                  nePlanEdt : ( PlanEdt : Record
                                  FormatPrint  : TFormatPrintEdt ;
                                  Date11,Date21 : TDateTime ;
                                  PeriodeDates : Array[1..24] of TDateTime ;
                                  SoldeFormate : String3 ;
                                  Sauf : String[255] ;
                                  Taux         : Double ;
                                  Exo1, Exo2,MvtSur : String3 ;
                                  NatureBud : String3 ;
                                  SANBud : String[10] ;
                                  Famille : String17 ;
                                  Nature,Journal : String3 ;
                                  Resol        : hChar ;
                                  Detail1,Detail2,Detail3 : String1 ; // 0 : non , 1 : oui , 2 : suivant le paramétrage
                                  Realise, Qte, QueTotalRea : Boolean ;
                                  SousPlan : Integer ;
                                  PasDeRubrique : Boolean ;
                                  End ;
                              ) ;
                END ;

ClassCritEdt = class
                 CritEdt : TCritEdt;
               end;

Type tColPcl = Record
               Actif : Boolean ;
               Exo : TExoDate ;
               Date1,Date2 : tDateTime ;
               AvecHisto : Boolean ;
               IsFormule : Boolean ;
               IdentHisto : String ;
               StFormule : String ;
               StFormat : String ;
               StFormatDetail : string;
               StTitre : String ;
               BalSit : string;
               End ;

Type tTabColPCL = Array[0..9] Of tColPCL ;

Type ttypEcr = (Reel,Simu,Situ,Previ,Revi,Ifrs) ;         // Modif IFRS 05/05/2004
Type tEcrPCL = Array[Reel..Ifrs] Of Boolean ;             // Modif IFRS 05/05/2004
Type tAffichePCL = (OnPivot,OnDevise,OnContrevaleur) ;

{ Ajout CA le 05/02/2001 }
Type TTypeEtatSynthese = (esSIG,esBIL,esCR,esCRA,esSIGA,esBILA);
{ Fin Ajout CA }

Type tCritEdtPCL = Record
                   NbColActif : Integer ;
                   Col : TTabColPCL ;
                   AvecPourcent : Boolean ;
                   AvecDetail : Boolean ;
                   TypEcr : tEcrPCL ;
                   Etab,Devise : String ;
                   EnMonnaieOpposee,DeviseEnPivot : String ;
                   Modele : String ;
                   Affichage : tAffichePCL ;
                   StTypEcr : String ;
                   FormatMontant : String ;
                   Resolution : String ;
                   ResolutionDetail : string;
                   AvecMontant0 : Boolean ;
                   bLibelleCompte : boolean;
                   bLibelleSection : boolean;
                   TE : TTypeEtatSynthese ;
                   InfoLibre : string;
                   MentionFiligrane : string;
                   stAxe : String;    		// Pour compte de résultat et SIG analytique
                   stSousPlan: String;    // FP ANAL2
                   stSectionExclues: String; // FP ANAL2
                   stSectionDe : String;	// Pour compte de résultat et SIG analytique
                   stSectionA : String;		// Pour compte de résultat et SIG analytique
                   bTLI : boolean;				// Pour compte de résultat et SIG analytique
                   bJoker : Boolean;      // Pour joker
                   bUnEtatParSection : boolean;
                   bLibelleDansLesTitres : boolean;
                   bImpressionDate : boolean;
                   bEtat : boolean;
                   bExport : boolean;
                   Corresp : string;
                   GeneTrad  : string; // contient le champs de traduction des libellés // FQ 22154
                   End ;
implementation


end.
