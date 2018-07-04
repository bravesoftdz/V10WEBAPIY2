{***********UNITE*************************************************
Auteur  ...... : CB
Créé le ...... : 14/03/2002
Modifié le ... :   /  /
Description .. : Types et constantes utilisés dans le planning
*****************************************************************}
unit AFPlanningCst;

interface

uses UTob;

type

  AffaireCourante = record
    StAffaire  : string;
    StLibAff   : string;
    StAff0     : string;
    StAff1     : string;
    StAff2     : string;
    StAff3     : string;
    StAvenant  : string;
    StTiers    : string;
    StLibTiers : string;
    StDevise   : string;
    StChantier : String;
    StAffInit  : String;
    DtDebutAff     : TDateTime;
    DtFinAff       : TDateTime;
    DtDateSigne    : TDateTime;
    DtDateDebGener : TDateTime;
    DtResil        : TDateTime;
    DtDateFinGener : TDateTime;
    DtDateLimite   : TDateTime;
    DtDateCloture  : TDateTime;
    DtDateGarantie : TDateTime;
  end;
 
  RecordTache = record
    StNumeroTache: string;
    StLibTache: string;
    StFctTache: string;
    StArticle: string;
    StTiers: string;
    StUnite: string;
    StLastDateGene: TdateTime;
    StActiviteRepris: string;
    BoCompteur: Boolean;

    //mcd 26/07/04 ajout des champs libres, afin que ceux-ci soient copié dans planning
    StLib1: string;
    StLib2: string;
    StLib3: string;
    StLib4: string;
    StLib5: string;
    StLib6: string;
    StLib7: string;
    StLib8: string;
    StLib9: string;
    StLibA: string;
    RdVal1: Double;
    RdVal2: Double;
    RdVal3: Double;
    StBool1: string;
    StBool2: string;
    StBool3: string;
    DtDate1: TdateTime;
    DtDate2: TdateTime;
    DtDate3: TdateTime;
    StChar1: string;
    StChar2: string;
    StChar3: string;
  end;

  RecordArticle = record
    StArticle: string;
    StCodeArticle: string;
    StTypeArticle: string;
    StActReprise: string;
  end;

  {RecordPlanning = Record
      TobItems    : TOB;
      TobEtats    : TOB;
      TobRes      : TOB;
      TobCols     : TOB;
      TobRows     : TOB;
      TobEvents   : TOB;
  End;}

  RecordRessource = record
    RdAffaire: AffaireCourante;
    StOldRes: string;
    StNewRes: string;
    StNumTache: string;
    InStatut: Integer;
  end;

  RecordColonne = record
    StField: string;
    StSize: string;
    StAlign: string;
    StEntete: string;
  end;

  TTabRes = array of RecordRessource;
  TTJour = (Lun, Mar, Mer, Jeu, Ven, Sam, Dim, WE);
  TJour = set of TTJour;

const

  MAX_LEVEL_ETAT = 5;
  CadenceMois = '001';
  CadenceSemaine = '002';
  CadenceJour = '003';
  CadenceDemiJour = '004';
  CadenceHeure = '005';
  cInMenuPlanning = 153200;
  cInPopUp = 2;

  cInfsBold = 0;
  cInfsItalic = 1;
  cInfsUnderline = 2;
  cInfsStrikeOut = 3;

  cInMois = 38;
  cInColMnt = 13;
  cInTitre1 = 100;
  cInTitre2 = 42;

  cInColNomTache = 0;
  cInColPrevu = 1;
  cInColAffecte = 2;
  cInColRealise = 3;
  cInColRafTh = 16;
  cInColRaf = 17;
  cInColTotal = 18;
  cInColEcart = 19;
  cInColTache = 20;
  cInColLigFct = 21;

  cStSep = '/';

  cInColDebMois = 4;
  cInColFinMois = 15;

  cInUpdate = 0;
  cInDelete = 1;

  cInValArticle = 0;
  cInValFonction = 1;
  cInValRessource = 2;
  cInValTache = 3;
  cInValTarif = 4;

  cInColRes = 1;
  cInColResLib = 2;
  cInColQte = 3; //qte initale planning
  cInColQteRAP = 4; //rap planning
  cInColQtePC = 5; //qte initiale PC
  cInColQteRAF = 6; //reste à faire PC
  cInColStatut = 7;
  {  cInColComp1     = 8;
    cInColComp2     = 9;
    cInColComp3     = 10;
  }
  cInPlanning1 = 1;
  cInPlanning2 = 2;
  cInPlanning3 = 3;
  cInPlanning4 = 4;
  cInPlanning5 = 5;
  cInPlanning6 = 6;
  cInPlanning7 = 7;

  cInCreationDirecte = 0;
  cInCreationSansSelection = 1;
  cInCreationAvecSelection = 2;
  cInCreationNonProductive = 3;
  cInCreationSelectionEtRes = 4;

  cInQuotidienne = 1;
  cInHebdomadaire = 2;
  cInMensuelle = 3;
  cInAnnuelle = 4;
  cInNbInterv = 5;

  cInMoisMethode1 = 1; // numéro du jour
  cInMoisMethode2 = 2; // libellé du jour
  cInMoisMethode3 = 3; // semaine

  cInPremier = 1; // premiere semaine
  cInSecond = 2;
  cInTroisieme = 3;
  cInQuatrieme = 4;
  cInDernier = 5;

  cInResActive = 1;
  cInResInactive = 2;
  cInResTempo = 3;

  cInRemplace = 0;
  cInRemplaceTout = 1;
  cInSupprimer = 2;
  cInSupprimerTout = 3;
  cInAjout = 4;

  cInIconeECOLE = 78;
  cInIconeINTERNE = 99;
  cInIconeNORMALE = -1;

  cInNoCompteur = -1;

  cInDimanche = 1;
  cInLundi = 2;
  cInMardi = 3;
  cInMercredi = 4;
  cInJeudi = 5;
  cInVendredi = 6;
  cInSamedi = 7;

  cStAnnule = 'A';
  cStJour = 'J';
  cStJourPrec = 'M';
  cStJourSuiv = 'P';

  cInConserver = 0;
  cInAnnuler = 1;
  cInDecaler = 2;

implementation

end.

