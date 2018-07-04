unit MesgErrCom;

interface

//POUR TOUS //AJOUT MARION 12-05-04
const ERR_ID               = 'Identifiant obligatoire.';
//const ERR_FIXE             = 'Champ obligatoire.'; 

// exercice
{----------------------------ERREUR TESTEE marion 24-05-04---------------------}
const ERR_CODE_EXERCICE    = 'Code exercice manquant (pos 7)'; //INCLUS
//AJOUT MARION 26-05-2004
const ERR_CODE_EXERCICE2   = 'Code exercice incorrecte (pos 7)'; //INCLUS
const ERR_CODE_DOUBLON     = 'Plusieurs lignes pour le même exercice.'; //INCLUS
const ERR_ETAT_EXERCICE    = 'Plus de deux exercices sont ouverts. (pos 26)'; //INCLUS
//AJOUT MARION 02-06-04
const ERR_ETATANO          = 'Typologie des A-Nouveaux manquante. (pos 67)';//INCLUS
const ERR_ETATANO_INEXIST  = 'Typologie des A-Nouveaux incorrecte. (pos 67)';//INCLUS

const ERR_DATEDEB_EXERCICE = 'Date début exercice inexistant (pos 10)';//INCLUS
const ERR_DATEFIN_EXERCICE = 'Date fin exercice inexistant (pos 18)';//INCLUS
const ERR_ETAT_INEXISTANT  = 'Etat comptable inexistant, la valeur par défaut : OUV,CLO,CPR (pos 26)';//INCLUS
{----------------------------FIN DE TEST ERREUR--------------------------------}

// table libre
const ERR_TLIBRE_CODE      = 'Code inexistant dans le fichier (pos 7)';
const ERR_TLIBRE_LIBELLE   = 'Libellé manquant dans le fichier (pos 24)';
const ERR_TLIBRE_TYPE      = 'type manquant par défaut: GEN,TIE,SEC,ECR,ANA (pos 59)';
// sous section
const ERR_SSECTION_CODE    = 'Code inexistant dans le fichier (pos 7)';
const ERR_SSECTION_LIB     = 'Libellé manquant dans le fichier (pos 24)';
const ERR_SSECTION_AXE     = 'Axe manquant dans le fichier (pos 59)';

//  établissement et mode de paiement journal
const ERR_JRL_CODE         = 'Code inexistant dans le fichier (pos 7)';//INCLUS
const ERR_JRL_LIBELLE      = 'Libellé manquant dans le fichier (pos 10)';//INCLUS
const ERR_JRL_NATURE       = 'nature manquant dans le fichier (pos 45)';

// devise
const ERR_DEV_CODE         = 'Code inexistant dans le fichier (pos 7)';
const ERR_DEV_LIBELLE      = 'Libellé manquant dans le fichier (pos 10)';
const ERR_DEV_DECIMALE     = 'Nombre décimale manquant dans le fichier (pos 49)';
const ERR_DEV_IN           = 'Monnaie IN manquante dans le fichier (pos 56)';
const ERR_DEV_TAUX         = 'Taux manquant dans le fichier (pos 57)';
const ERR_DEV_EURO         = 'Subdivision de l''euro inexistante dans le fichier (pos 80)';

// comptes généraux
const ERR_GENERAUX_CODE    = 'Code inexistant dans le fichier (pos 7)';  //INCLUS
{----------------------------ERREUR TESTEE marion 26-05-04---------------------}
const ERR_GENERAUX_NATURE  = 'Nature manquante ou incorrecte dans le fichier (pos 59)';//INCLUS
const ERR_GENERAUX_CDLETTE = 'Code lettrable manquant ou différent de X ou - dans le fichier (pos 62)';//INCLUS
const ERR_GENERAUX_CDPOINT = 'Code pointable manquant ou différent de X ou - dans le fichier (pos 63)';//INCLUS
const ERR_GENERAUX_VENTIL1 = 'Ventilation axe1 manquante ou différente de X ou - dans le fichier (pos 63)'; //INCLUS
const ERR_GENERAUX_VENTIL2 = 'Ventilation axe2 manquante ou différente de X ou - dans le fichier (pos 65)';//INCLUS
const ERR_GENERAUX_VENTIL3 = 'Ventilation axe3 manquante ou différente de X ou - dans le fichier (pos 66)';//INCLUS
const ERR_GENERAUX_VENTIL4 = 'Ventilation axe4 manquante ou différente de X ou - dans le fichier (pos 67)';//INCLUS
const ERR_GENERAUX_VENTIL5 = 'Ventilation axe5 manquante ou différente de X ou - dans le fichier (pos 68)';//INCLUS
//AJOUT MARION 12-05-2004
const ERR_GENERAUX_LETTR   = 'Attention ce compte n est pas lettrable.';//INCLUS
//ajout marion 02-06-04
const ERR_GENERAUX_SENS    = 'Sens du compte incorrect.';//INCLUS
{----------------------------FIN DE TEST ERREUR--------------------------------}
const ERR_GENERAUX_LIB    = 'Libelle du compte manquant (pos 24).';//INCLUS
const ERR_COLLECTLET      = 'Attention compte collectif  : '; //+affichage du compte
const ERR_COLLETLETSUITE  = ' est lettrable';
const ERR_COMPTEALPHA     = 'Attention Comptes Généraux alpha numériques : ';  // + compte

// entête
const ENTETE_INEXSISTANT   = 'Entête du fichier inexistante ou incorrecte.';//INCLUS
CONST ENTETE_DTEBASCUL     = 'Format de la date de bascule incorrect. (pos 18)'; //INCLUS
const ENTETE_DTEARRPER     = 'Format de la date d''arrete periodique incorrect.(pos 26)';//INCLUS
{---------------------ERREUR TESTEE-------------------marion 25-05-04}
const ENTETE_TYPEINCORRECT = 'Type de fichier incorrect.'; //INCLUS
const ENTETE_TYPEINCSUITE  = ': Entête du fichier Pos (9)';
const ENTETE_TYPEMANQUANT = 'Type de fichier manquant.';  //INCLUS

const ENTETE_ORIGINE       = 'Origine du fichier incorrect : ' ; //INCLUS
const ENTETE_ORIGINESUITE  = ' Entête du fichier Pos (6)';
const ENTETE_ORIGINE2       = 'Origine du fichier manquant.' ; //INCLUS
const ENTETE_NUMDOCAB      = ' Numero de dossier au cabinet inexistant';

const ENTETE_FORMAT        = 'Format du fichier incorrect : ';  //INCLUS
const ENTETE_FORMATSUITE   = ' Entête du fichier Pos (12)';
//AJOUT MARION 25-05-2004
const ENTETE_FORMAT2        = 'Format du fichier manquant.'; //INCLUS
const ENTETE_FORMAT_TYPE    = 'Incohérence entre le type et le format du fichier.'; //INCLUS
const ENTETE_VERSION        = 'Version de fichier manquant.(pos 34)';  //INCLUS
{-------------------------FIN ERREUR TEST-------------}

//Exercice
{---------------------ERREUR TESTEE-------------------marion 25-05-04}
const ERR_EXERCICEDATEDEB  = 'Date de début d''exercice incorrect.'; //INCLUS
const ERR_EXERCICEDATEFIN  = 'Date de fin d''exercice incorrect.'; //INCLUS
const ERR_EXERCICEETAT     = 'Etat comptable incorrect : enregistrement EXO (pos 26)'; //OUV,NON,CDE,CPR,CLO  //TESTER
//AJOUT MARION 25-05-2004
const ERR_EXERCICETYPEEXO  = 'Type d''exercice incompatible avec le mode de saisie du journal.';
{-------------------------FIN ERREUR TEST-------------}

const ERR_EXERCICEOUV      = 'Le fichier contient plus de deux exercices ouverts.';
const ERR_EXOMANQUANT      = 'Fichier incorrect : l''enregistrement exercice ***EXO est manquant.'; //INCLUS
const ERR_EXOPLUSOUV       = 'Vous avez déjà deux exercices ouverts, veuillez effectuer une clôture d''exercice';
const ERR_DATEEXEXER       = 'Les dates d''exercice sont incohérentes avec le dossier.';
const ERR_EXERCDECALE      = 'Exercice est décalé. L''importation des écritures est impossible.';
const ERR_EXERCICECLOS     = 'L''exercice est clos. L''importation des écritures est impossible.'; // cas journal, balance

//PS1,PS2,PS3,PS4,PS5
const ERR_PARAMETRE        = 'Fichier incorrect : le paramètre PS'; // + indice de psx //INCLUS
const ERR_PARAMETRESUITE   = ' est manquant.';  //INCLUS

{---------------------ERREUR TESTEE-------------------marion 26-05-04}
const ERR_MONNAIETENUE     = 'Paramètrage de la monnaie de tenue incorrect : (Enreg. PS5 : pos. 11)'; //INCLUS
const ERR_PARAMTENUE       = 'Paramètrage de la monnaie de tenue incorrect : (Enreg. PS5 : pos. 7)'; //INCLUS
//AJOUT MARION 26-05-2004
const ERR_DEVISE           = 'Devise manquante : (Enreg. PS5 : pos. 7)';  //INCLUS
const ERR_DEVISE2          = 'Devise incorrecte : (Enreg. PS5 : pos. 7)'; //INCLUS
CONST ERR_TENUE            = 'Monnaie de Tenue manquante:(Enreg. PS5 : pos. 11)'; //INCLUS
{-------------------------FIN ERREUR TEST-------------}

const ERR_MONNAIEPS5       = 'Enregistrement PS5 : la monnaie de tenue est incohérente avec le paramétrage du dossier.'; //INCLUS
const ERR_PARAMPS5         = 'Enregistrement PS5 : vous devez renseigner tous les paramètres obligatoires dans le fichier.' ;//INCLUS
const ERR_JALPS5           = 'Enregistrement PS5 : vous devez renseigner le journal d''ecrat de conversion (pos 12).'; //INCLUS
CONST ERR_PS5_GEREANA      = 'Enregistrement PS5 : Code gestion analytique différent de X ou - dans le fichier (pos 76).';//INCLUS
{---------------------ERREUR TESTEE-------------------marion 26-05-04}
const ERR_LGCPTEPS2        = 'Enregistrement PS2 : longueur des comptes inexistante.';//INCLUS
const ERR_LGSECTION        = 'Enregistrement PS2 : longueur des sections inexistantes'; //INCLUS
{-------------------------FIN ERREUR TEST-------------}

const ERR_SECTIONANA       = 'Enregistrement PS2 : la longueur des sections analytiques est incohérente avec le paramétrage du dossier.';
const VERIF_PARAMSOC       = 'Enregistrement PS2 :Vérifier les comptes d''attente dans les paramètres société ';
const VERIF_GNRLATTEN      = 'Enregistrement PS2 :Vérifier le compte general d''attente dans les paramètres société.';
const VERIF_CLIATTEN       = 'Enregistrement PS2 :Vérifier le compte client d''attente dans les paramètres société.';
const VERIF_FOUATTEN       = 'Enregistrement PS2 :Vérifier le compte fournisseur d''attente dans es paramètres société.';
const VERIF_SALATTEN       = 'Enregistrement PS2 :Vérifier le compte salarie d''attente dans les paramètres société.';
const VERIF_CDIVATTEN      = 'Enregistrement PS2 :Vérifier le compte tiers d''attente dans les paramètres société.';


// cas journal et synchro
const ERR_LGCPTE           = 'Enregistrement PS2 : la longueur des comptes est incohérente avec le paramétrage du dossier.';//INCLUS

//JOURNAUX //AJOUT MARION 12-05-2004
{---------------------ERREUR TESTEE-------------------marion 25-05-04}
const ERR_JOURN_CODE        = 'Code journal manquant(pos 7).'; //INCLUS
const ERR_JOURN_NATURE1     = 'Nature du journal manquante.';  //INCLUS
const ERR_JOURN_NATURE2     = 'Nature du journal incorrecte.'; //INCLUS
const ERR_JOURN_MODSAISI    = 'Mode de saisie incorrecte.';  //INCLUS
const ERR_JOURN_CPT         = 'Numero de compte inexistant.';//INCLUS
const ERR_JOURN_AXE         = 'Axe analytique inexistant.';//INCLUS
{-------------------------FIN ERREUR TEST-------------}
const ERR_JALMANQUANT      = 'Fichier incorrect : l''enregistrement journal ***JAL est manquant.';//INCLUS 

// cas sisco
const ERR_EXODOSINEXIST    = 'Le dossier comptable n''existe pas. Veuillez le créer.'; // cas dossier
const ERR_TRANCHEAUXIL     = 'Tranches de comptes auxiliaires incohérentes avec le paramétrage du dossier.';


const DOSSIER_INEXSISTANT  = 'Le dossier comptable n''existe pas. Veuillez le créer.';

//fichier
const FICHIER_INEXISTANT   = 'Choisir un fichier d''import ';

// ecriture
const ERR_INEXISTANT       = 'Fichier incorrect : aucunes lignes d''écriture.';//INCLUS
const ERR_DATEECRITURE     = 'Date d''écriture incompatible par rapport aux exercices du dossier ';
const ERR_MOUVMNT          = 'Statut du mouvement incorrect ou manquant. pos(1021)';//INCLUS
const ERR_MOUVMNT_CODE     = 'Code de lettrage incompatible avec l''etat lettrage.(pos 1014)';//INCLUS
{---------------------ERREUR TESTEE-------------------marion 25-05-04}
//ECRITURE  -AJOUT MARION 24-05-2004
const ERR_DTECMPABL        = 'Date comptable manquante ou incorrecte (pos 4).'; //INCLUS
const ERR_DTEECH           = 'Date d''echeance incorrecte (pos 122).'; //INCLUS
const ERR_TYPPIEC          = 'Type pièce manquante (pos 12).';  //INCLUS
const ERR_TYPPIEC2         = 'Type pièce incorrecte (pos 12).'; //INCLUS
const ERR_MONTANT1         = 'Montant1 manquant ou incorrect'; //INCLUS
const ERR_CODEMONTANT      = 'Code montant incorrect'; //INCLUS
//const ERR_AXE              = 'Axe analytique manquant'; //INCLUS
const ERR_SECTION          = 'Section analytique manquante';
const ERR_DEVISEMANQ       = 'Zone devise manquante';
//ajout marion 02-06-04
const ERR_CODE_JOURNAL      = 'Code journal incorrect.(pos 1)';
const ERR_CMPT_GNRL         = 'Compte général incorrect ou inexistant.';//INCLUS
const ERR_CMPTGNR_NAT       = 'Nature du compte général non collectif.'; //INCLUS
const ERR_TYPE_CPTE1        = 'Nature de la ligne de mouvement manquante.';
const ERR_TYPE_CPTE2        = 'Nature de la ligne de mouvement incorrecte.';
const ERR_SENS1             = 'Sens du mouvement manquant.'; //INCLUS
const ERR_SENS2             = 'Sens du mouvement incorrect';//INCLUS
const ERR_TYPE_ECR          = 'Type de mouvement manquant ou incorrect.'; //INCLUS
const ERR_AXE_ECR           = 'Axe analytique manquant.'; //INCLUS
const ERR_MODSAIS_JOURN     = 'Mode de saisie du journal incorrect';//INCLUS
const ERR_ANOUVEAU          = 'Code du journal incompatible avec la nature des mouvements d''à nouveaux (pos 302).'; //INCLUS
{-------------------------FIN ERREUR TEST-------------}

// compte remplacé par compte général d'attente.Si compte général d'attente non renseigné dans paramsoc alors :
const ERR_COMPTEGENEMANQ   = 'Compte général est manquant. Erreur Ecriture.';  //INCLUS
const ERR_COMPTEAUXIMANQ   = 'Compte auxiliaire est manquant. Erreur Ecriture.'; //INCLUS

//comptegene +ERR_COMPTECOLL1+ compteauxi+ERR_COMPTECOLL1suite
const ERR_COMPTECOLL1      = ' n''est pas un collectif,' ;
const ERR_COMPTECOLL1suite = ' est remplacé par blanc';
const VERIF_AXE            = 'Vérifier les informations correspondantes aux axes analytiques';

//AJOUT MARI0N 28-05-2004
//compte Tiers
const ERR_CMPTTIER_CODE   = 'Code du compte auxiliare manquant. (pos 7)'; //INCLUS
const ERR_CMPTTIER_LIBEL  = 'Libelle du compte manquant (pos 24).'; //INCLUS
CONST ERR_CMPTTIER_CGN    = 'Compte collectif associé manquant. (pos 63)';  //INCLUS
CONST ERR_CMPTTIER_CGN2   = 'Compte collectif associé inconnu (pos 63).';  //INCLUS
const ERR_CMPTTIER_NATURE = 'Nature du compte inexistante ou incorrecte.(pos 59)';//INCLUS
const ERR_CMPTTIER_LETTRA = 'Code lettrable manquant ou inexistant.(pos 62)';//INCLUS
const ERR_CMPTTIER_ADR1   = 'Coordonnées de l''Adresse 1 manquantes.(pos 267)';//INCLUS
const ERR_CMPTTIER_ADR2   = 'Coordonnées de l''Adresse 2 manquantes.(pos 302)'; //INCLUS
const ERR_CMPTTIER_DOMICIL= 'Domiciliation bancaire inexistante ou incorrecte.(pos 416)';//INCLUS
const ERR_CMPTTIER_ETAB   = 'Etablissement bancaire inexistant ou incorrect.(pos 440)'; //INCLUS
const ERR_CMPTTIER_GUICH  = 'Code guichet inexistant ou incorrect.(pos 445)'; //INCLUS
const ERR_CMPTTIER_CMPT   = 'Numero de compte inexistant ou incorrect.(pos 450)'; //INCLUS
const ERR_CMPTTIER_CLE    = 'Cle inexistant ou incorrect.(pos 461)';//INCLUS
const ERR_CMPTTIER_PAY    = 'Pays inexistant.(pos 463)'; //INCLUS
const ERR_CMPTTIER_MDEV   = 'Multidevise incorrecte.(pos 486)';//INCLUS
const ERR_CMPTTIER_RTVA   = 'Regime TVA inexistant ou incorrect.(pos 540)'; //INCLUS
const ERR_CMPTTIER_MRGLMT = 'Mode de réglement inexistant ou incorrect.(pos 543)';//INCLUS
const ERR_CMPTTIER_CPPAL  = 'Contact principal incorrect.(pos 853)'; //INCLUS
const ERR_CMPTTIER_RIBPAL = 'RIB principal incorrect.(pos 857)'; //INCLUS
const ERR_CREATE_CPTE = 'ERREUR CREATION ';

implementation

end.

