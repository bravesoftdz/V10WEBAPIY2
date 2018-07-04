unit ControlParam;

interface

uses
  {$IFDEF MODENT1}
  CPTypeCons,
  {$ELSE}
  Ent1,
  {$ENDIF MODENT1}
  MesgErrCom;


type
   ControlCom = record
             NomChamp : string;
             deb,fin  : integer;
             LeFb     : TFichierBase;
             Nature   : string;
             Valeur   : string;
             champ    : string;
             end;

   ControlComEnreg = record
             deb,fin  : integer;
             mess     : string;
             end;

   ModeReglS1 = record
             MDR     : string;
             end;

const SizeRecControl = 21;
const SizeRecControlPS3 = 17;
//const SizeEnregControl = 28;
const SizeEnregControl = 26;

var
   EnregControl : array [0..SizeEnregControl] of ControlComEnreg =
   (
// exercice
   (deb : 7  ; fin : 3 ; mess : ERR_CODE_EXERCICE),    //'Code exercice manquant (pos 7)'),
   (deb : 10 ; fin : 8 ; mess : ERR_DATEDEB_EXERCICE), //'Date début exercice inexistant (pos 10)'),
   (deb : 18 ; fin : 8 ; mess : ERR_DATEFIN_EXERCICE), //'Date fin exercice inexistant (pos 18)'),
   (deb : 26 ; fin : 3 ; mess : ERR_ETAT_INEXISTANT),  //'Etat comptable inexistant, la valeur par défaut : OUV,CLO,CPR pos(26)'),
// table libre
   (deb : 7  ; fin : 17 ; mess : ERR_TLIBRE_CODE),     //'Code inexistant dans le fichier (pos 7)'),
   (deb : 24 ; fin : 35 ; mess : ERR_TLIBRE_LIBELLE),  //'Libellé manquant dans le fichier (pos 24)'),
   (deb : 59 ; fin : 3  ; mess : ERR_TLIBRE_TYPE),     // 'type manquant par défaut: GEN,TIE,SEC,ECR,ANA (pos 59)'),
// sous section
   (deb : 7  ; fin : 17 ; mess : ERR_SSECTION_CODE),   //'Code inexistant dans le fichier (pos 7)'),
   (deb : 24 ; fin : 35 ; mess : ERR_SSECTION_LIB),    //'Libellé manquant dans le fichier (pos 24)'),
   (deb : 59 ; fin : 3  ; mess : ERR_SSECTION_AXE),    //'Axe manquant dans le fichier (pos 59)'),
//  établissement et mode de paiement journal
   (deb : 7  ; fin : 3 ; mess : ERR_JRL_CODE),         //'Code inexistant dans le fichier (pos 7)'),
//   (deb : 10 ; fin : 35 ; mess : ERR_JRL_LIBELLE),     //'Libellé manquant dans le fichier (pos 10)'),
   (deb : 45 ; fin : 3 ; mess : ERR_JRL_NATURE),       //'nature manquant dans le fichier (pos 45)'),

// Devise
   (deb : 7  ; fin : 3  ; mess : ERR_DEV_CODE),        //'Code inexistant dans le fichier (pos 7)'),
   (deb : 10 ; fin : 35 ; mess : ERR_DEV_LIBELLE),     //'Libellé manquant dans le fichier (pos 10)'),
   (deb : 49 ; fin : 1  ; mess : ERR_DEV_DECIMALE),    //'Nombre décimale manquant dans le fichier (pos 49)'),
   (deb : 56 ; fin : 1  ; mess : ERR_DEV_IN),          //'Monnaie IN manquante dans le fichier (pos 56)'),
   (deb : 57 ; fin : 20 ; mess : ERR_DEV_TAUX),        //'Taux manquant dans le fichier (pos 57)'),
   (deb : 80 ; fin : 1  ; mess : ERR_DEV_EURO),        //'Subdivision de l''euro inexistante dans le fichier (pos 80)'),
// comptes généraux
   (deb : 7  ; fin : 17 ; mess : ERR_GENERAUX_CODE),    //'Code inexistant dans le fichier (pos 7)'),
//   (deb : 24 ; fin : 35 ; mess : 'Libellé manquant dans le fichier (pos 24)'),
   (deb : 59 ; fin : 3  ; mess : ERR_GENERAUX_NATURE),  //'Nature manquante dans le fichier (pos 59)'),
   (deb : 62 ; fin : 1  ; mess : ERR_GENERAUX_CDLETTE), //'Code lettrable manquant ou différent de X ou - dans le fichier (pos 62)'),
   (deb : 63 ; fin : 1  ; mess : ERR_GENERAUX_CDPOINT), //'Code pointable manquant ou différent de X ou - dans le fichier (pos 63)'),
   (deb : 64 ; fin : 1  ; mess : ERR_GENERAUX_VENTIL1), //'Ventilation axe1 manquante ou différente de X ou - dans le fichier (pos 63)'),
   (deb : 65 ; fin : 1  ; mess : ERR_GENERAUX_VENTIL2), //'Ventilation axe2 manquante ou différente de X ou - dans le fichier (pos 65)'),
   (deb : 66 ; fin : 1  ; mess : ERR_GENERAUX_VENTIL3), //'Ventilation axe3 manquante ou différente de X ou - dans le fichier (pos 66)'),
   (deb : 67 ; fin : 1  ; mess : ERR_GENERAUX_VENTIL4), //'Ventilation axe4 manquante ou différente de X ou - dans le fichier (pos 67)'),
   (deb : 68 ; fin : 1  ; mess : ERR_GENERAUX_VENTIL5)  //'Ventilation axe5 manquante ou différente de X ou - dans le fichier (pos 68)')
   );
   RecControl : array [0..SizeRecControl] of ControlCom =
  (
// PS2
   (Nomchamp : 'SO_GENATTEND' ; deb : 28  ; fin : 17 ; LeFb : fbGene ; Nature : 'DIV' ; Valeur : ''; champ : 'CPT'),
   (Nomchamp : 'SO_CLIATTEND' ; deb : 45  ; fin : 17 ; LeFb : fbAux  ; Nature : 'CLI' ; Valeur : ''; champ : 'CPT'),
   (Nomchamp : 'SO_FOUATTEND' ; deb : 62  ; fin : 17 ; LeFb : fbAux  ; Nature : 'FOU' ; Valeur : ''; champ : 'CPT'),
   (Nomchamp : 'SO_SALATTEND' ; deb : 79  ; fin : 17 ; LeFb : fbAux  ; Nature : 'SAL' ; Valeur : ''; champ : 'CPT'),
   (Nomchamp : 'SO_DIVATTEND' ; deb : 96  ; fin : 17 ; LeFb : fbAux ; Nature : 'DIV' ; Valeur : ''; champ : 'CPT'),
// PS3
   (Nomchamp : 'SO_OUVREBIL'  ; deb : 7  ; fin : 17 ; LeFb : fbGene ; Nature : 'DIV' ;  Valeur : ''; champ : 'CPT'),
   (Nomchamp : 'SO_RESULTAT'  ; deb : 24  ; fin : 17 ; LeFb : fbGene ; Nature : 'DIV' ; Valeur : ''; champ : 'CPT'),
   (Nomchamp : 'SO_OUVREBEN'  ; deb : 41  ; fin : 17 ; LeFb : fbGene ; Nature : 'DIV' ; Valeur : ''; champ : 'CPT'),
   (Nomchamp :'SO_OUVREPERTE' ; deb : 58  ; fin : 17 ; LeFb : fbGene ; Nature : 'DIV' ; Valeur : ''; champ : 'CPT'),
   (Nomchamp : 'SO_DEFCOLCLI' ; deb : 84  ; fin : 17 ; LeFb : fbGene ; Nature : 'COC' ; Valeur : ''; champ : 'CPT'),
   (Nomchamp : 'SO_DEFCOLFOU' ; deb : 101 ; fin : 17 ; LeFb : fbGene ; Nature : 'COF' ; Valeur : ''; champ : 'CPT'),
   (Nomchamp : 'SO_DEFCOLSAL' ; deb : 118 ; fin : 17 ; LeFb : fbGene ; Nature : 'COS' ; Valeur : ''; champ : 'CPT'),
   (Nomchamp : 'SO_DEFCOLDDIV'; deb : 135 ; fin : 17 ; LeFb : fbGene ; Nature : 'COD' ; Valeur : ''; champ : 'CPT'),
   (Nomchamp : 'SO_DEFCOLCDIV'; deb : 152 ; fin : 17 ; LeFb : fbGene ; Nature : 'COD' ; Valeur : ''; champ : 'CPT'),
   (Nomchamp : 'SO_DEFCOLDIV' ; deb : 169 ; fin : 17 ; LeFb : fbGene ; Nature : 'COD' ; Valeur : ''; champ : 'CPT'),
   (Nomchamp : 'SO_JALOUVRE'  ; deb : 75  ; fin : 3  ; LeFb : fbGene ; Nature : 'ANO' ; Valeur : ''; champ : 'JAL'),
   (Nomchamp : 'SO_JALFERME'  ; deb : 78  ; fin : 3  ; LeFb : fbGene ; Nature : 'CLO' ; Valeur : ''; champ : 'JAL'),
   (Nomchamp:'SO_JALREPBALAN' ; deb : 81  ; fin : 3  ; LeFb : fbGene ; Nature : 'OD'  ; Valeur : ''; champ : 'JAL'),
// PS5
//   (Nomchamp:'SO_JALECARTEURO' ; deb : 12  ; fin : 3  ; LeFb : fbGene ; Nature : 'REG'  ; Valeur : ''; champ : 'JAL'),
//   (Nomchamp: 'SO_ECCEURODEBIT'; deb : 15 ; fin : 17 ; LeFb : fbGene ; Nature : 'CHA' ; Valeur : ''; champ : 'CPT'),
//   (Nomchamp:'SO_ECCEUROCREDIT'; deb : 32 ; fin : 17 ; LeFb : fbGene ; Nature : 'PRO' ; Valeur : ''; champ : 'CPT'),
   (Nomchamp: 'SO_REGIMEDEFAUT'; deb : 59 ; fin : 3 ; LeFb : fbGene ; Nature : '' ; Valeur : ''; champ : 'DIV'),
   (Nomchamp: 'SO_GCMODEREGLEDEFAUT'; deb : 62 ; fin : 3 ; LeFb : fbGene ; Nature : '' ; Valeur : ''; champ : 'DIV'),
   (Nomchamp: 'SO_CODETVADEFAUT'; deb : 65 ; fin : 3 ; LeFb : fbGene ; Nature : '' ; Valeur : ''; champ : 'DIV'),
   (Nomchamp: 'SO_CODETVAGENEDEFAULT'; deb : 68 ; fin : 3 ; LeFb : fbGene ; Nature : '' ; Valeur : ''; champ : 'DIV')
  );
  ControlMDR : array [0..18] of ModeReglS1 =
  ((MDR : 'ARH'), (MDR : 'AVR'), (MDR : 'BOE'),
   (MDR : 'BOR'), (MDR : 'CB'),  (MDR : 'CCD'),
   (MDR : 'CHE'), (MDR : 'CHQ'), (MDR : 'DIV'),
   (MDR : 'ESP'), (MDR : 'LCR'), (MDR : 'PRE'),
   (MDR : 'REL'), (MDR : 'RTD'), (MDR : 'TEP'),
   (MDR : 'TIP'), (MDR : 'TRD'), (MDR : 'TRE'),
   (MDR : 'VIR'));

type
   ParamLib = record
             indice   : integer;
             NomChamp : string;
             deb  : integer;
             Ext  : string;
   end;


const SizeRecPramlib = 37; // structure table libre
   RecPramlib : array [0..SizeRecPramlib] of ParamLib =
  (
   (indice : 0; Nomchamp : 'TEXTE0' ; deb : 46),
   (indice : 1; Nomchamp : 'TEXTE1' ; deb : 82),
   (indice : 2; Nomchamp : 'TEXTE2' ; deb : 118),
   (indice : 3; Nomchamp : 'TEXTE3' ; deb : 154),
   (indice : 4; Nomchamp : 'TEXTE4' ; deb : 190),
   (indice : 5; Nomchamp : 'TEXTE5' ; deb : 226),
   (indice : 6; Nomchamp : 'TEXTE6' ; deb : 262),
   (indice : 7; Nomchamp : 'TEXTE7' ; deb : 298),
   (indice : 8; Nomchamp : 'TEXTE8' ; deb : 334),
   (indice : 9; Nomchamp : 'TEXTE9' ; deb : 370),
   // montant
   (indice : 10; Nomchamp : 'MONTANT1' ; deb : 406),
   (indice : 11; Nomchamp : 'MONTANT2' ; deb : 442),
   (indice : 12; Nomchamp : 'MONTANT3' ; deb : 478),
   (indice : 13; Nomchamp : 'MONTANT4' ; deb : 514),
   // montant
   (indice : 14; Nomchamp : 'BOOLEAN1' ; deb : 550),
   (indice : 15; Nomchamp : 'BOOLEAN2' ; deb : 586),
   (indice : 16; Nomchamp : 'BOOLEAN3' ; deb : 622),
   (indice : 17; Nomchamp : 'BOOLEAN4' ; deb : 658),
   // Date
   (indice : 18; Nomchamp : 'DATE1' ; deb : 694),
   (indice : 19; Nomchamp : 'DATE2' ; deb : 730),
   (indice : 20; Nomchamp : 'DATE3' ; deb : 766),
   (indice : 21; Nomchamp : 'DATE4' ; deb : 802),

   // montant  champ table Y et E
   (indice : 22; Nomchamp : 'LIBREMONTANT0' ; deb : 406),
   (indice : 23; Nomchamp : 'LIBREMONTANT1' ; deb : 442),
   (indice : 24; Nomchamp : 'LIBREMONTANT2' ; deb : 478),
   (indice : 25; Nomchamp : 'LIBREMONTANT3' ; deb : 514),

   (indice : 26; Nomchamp : 'LIBREBOOL0' ; deb : 550),
   (indice : 27; Nomchamp : 'LIBREBOOL1' ; deb : 586),
   (indice : 28; Nomchamp : 'LIBREDATE' ; deb : 658),

   (indice : 29; Nomchamp : 'TABLE0' ; deb : 694),
   (indice : 30; Nomchamp : 'TABLE1' ; deb : 730),
   (indice : 31; Nomchamp : 'TABLE2' ; deb : 766),
   (indice : 32; Nomchamp : 'TABLE3' ; deb : 802),

   (indice : 33; Nomchamp : 'LIBRETEXTE1' ; deb : 46),
   (indice : 34; Nomchamp : 'LIBRETEXTE2' ; deb : 82),
   (indice : 35; Nomchamp : 'LIBRETEXTE3' ; deb : 118),
   (indice : 36; Nomchamp : 'LIBRETEXTE4' ; deb : 154),
   (indice : 37; Nomchamp : 'LIBRETEXTE5' ; deb : 190)
  );

implementation


end.
