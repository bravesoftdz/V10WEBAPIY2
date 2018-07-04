unit MessagesErreur;

interface

uses HMsgBox, HCtrls;

type
  TypeMsgErreur = (msNone, msError, msWarning, msAsk);

procedure ChargeTHMsgBoxFromArray(var MsgBox : THMsgBox; Tableau : array of string);
function GetTypeMsg(const Mess: String): TypeMsgErreur;

{***********A.G.L.***********************************************
Auteur  ...... : JTR
Créé le ...... : 18/05/2005
Modifié le ... :
Description .. : Messages liés à la comptabilisation
Mots clefs ... : FactCpta
*****************************************************************}
const MSGERR_TitreCpta : array[1..14] of string 	= (
      {1}   'Comptabilisation des opérations de caisse'
      {2}  ,'Comptabilisation de la pièce'
      {3}  ,'Comptabilisation du tiers'
      {4}  ,'Comptabilisation des articles'
      {5}  ,'Comptabilisation des frais de port'
      {6}  ,'Comptabilisation des taxes'
      {7}  ,'Comptabilisation des remises'
      {8}  ,'Comptabilisation des escomptes'
      {9}  ,'Comptabilisation de la tva intracommunautaire'
      {10} ,'Comptabilisation'
      {11} ,'Comptabilisation des remises en banque'
      {12} ,'Comptabilisation des stocks'
      {13} ,'Comptabilisation de l''extourne'
      {14} ,'Comptabilisation des règlements'
      );

const MSGERR_TexteCpta : array[1..34] of string 	= (
      {1}  'Compte général d''escompte absent ou incorrect'
      {2} ,'Compte général de remise absent ou incorrect'
      {3} ,'Compte général de taxe absent ou incorrect'
      {4} ,'Compte collectif tiers absent ou incorrect'
      {5} ,'Compte général de HT absent ou incorrect'
      {6} ,'Compte général d''écart de conversion absent ou incorrect'
      {7} ,'Journal comptable non renseigné pour cette nature de pièce'
      {8} ,'Nature comptable non renseignée'
      {9} ,'Erreur sur la numérotation du journal comptable'
      {10} ,'Certains comptes généraux, auxiliaires ou analytiques sont incorrects'
      {11} ,'Compte général de stock ou variation absent ou incorrect'
      {12} ,'Compte de retenue de garantie absent ou incorrect'
      {13} ,'Compte général de taxe RG absent ou incorrect'
      {14} ,'Compte général de banque (associé au journal ou au mode de paiement) absent ou incorrect'
      {15} ,'Erreur, les journaux des modes de paiement sont différent'
      {16} ,'Journal non renseigné'
      {17} ,'Erreur sur le client'
      {18} ,'Certaines écritures ne sont pas équilibrées'
      {19} ,'Ecart sur contrôle des pièces d''achat' {DBR CPA}
      {20} ,' : Le compte général, le code du journal ou le compte de contrepartie du journal est incorrect'
      {21} ,' : Le compte de virement interne ou le journal de la caisse général est incorrect'
      {22} ,'Caisse générale : Le code ou le compte de contrepartie est incorrect'
      {23} ,' : Le compte d''écart Crédit est incorrect'
      {24} ,' : Le compte d''écart Débit est incorrect'
      {25} ,'Le Régime fiscal de la TVA Intracommunautaire est incorrect'
      {26} ,'Erreur lors de l''enregistrement des lignes d''écriture (le n° de pièce comptable attribué %s est peut-être déjà utilisé)'
      {27} ,'Erreur lors de l''enregistrement de la comptabilité différée'
      {28} ,'Traitement interrompu.#13Le paramètrage du mode de paiement est incorrect'
      {29} ,'Erreur lors de la mise à jour des lignes de règlements'
      {30} ,'Erreur lors de la suppression des écritures comptables de la %s n° %s'
      {31} ,'Erreur lors de la mise à jour du solde du tiers'
      {32} ,'Erreur lors de la mise à jour du lien de la %s n° %s avec les écritures'
      {33} ,'Ligne de Tiers absente ou incorrecte'
      {34} ,'Ligne de Charges ou Produits absente ou incorrecte'
      );

{***********A.G.L.***********************************************
Auteur  ...... : JTR
Créé le ...... : 18/05/2005
Modifié le ... :
Description .. : Messages liés à la fiche article
Mots clefs ... : UTomArticle
*****************************************************************}
const MSGERR_Article : array[1..78] of string 	= (
      {1}  'Vous devez renseigner un fournisseur' //Catalogue
      {2}, 'Vous devez renseigner la référence' //Catalogue
      {3}, 'La nature de la pièce suivante ne peut être égale à la nature de pièce courante'
      {4}, 'Le prix de vente est inférieur au prix d''achat'
      {5}, 'Le code barre et le qualifiant code barre doivent être renseignés  '
      {6}, 'Le nombre de caractères du code barre est incorrect'
      {7}, 'Le code barre est incorrect'
      {8}, 'Le code barre contient des caractères non autorisés'
      {9}, 'Le code barre doit être numérique'
      {10}, 'Le libellé doit être renseigné'
      {11}, 'L''article de remplacement n''existe pas'
      {12}, 'L''article de substitution n''existe pas'
      {13}, 'Vous ne pouvez pas sélectionner l''article courant comme article de remplacement'
      {14}, 'Vous ne pouvez pas sélectionner l''article courant comme article de substitution'
      {15}, 'Erreur de création de l''article dimensionné'
      {16}, 'La première famille de taxe est obligatoire '
      {17}, 'Suppression impossible, cet article est utilisé dans une pièce'
      {18}, 'Le type prestation/frais/fourniture est obligatoire' // Affaire
      {19}, 'Le code prestation/frais/fourniture est obligatoire' // Affaire
      {20}, 'Le code article est obligatoire'
      {21}, 'L''article n''existe pas' //Catalogue
      {22}, 'Impossible de supprimer un fournisseur principal' //Catalogue
      {23}, 'Le fournisseur n''existe pas' //Catalogue
      {24}, 'Votre base tarif HT fait référence au fournisseur principal'
      {25}, 'Fournisseur principal inexistant'
      {26}, 'L''article lié doit être différent de l''article courant'
      {27}, 'Cet article lié existe déjà'
      {28}, 'Le calcul de la base tarif se fait uniquement à partir des bases de l''article'
      {29}, 'L''article ayant du stock, la case tenu en stock ne peut être décochée'
      {30}, 'Vous devez renseigner le libellé' //Catalogue
      {31}, 'Suppression impossible, cet article est utilisé sur une ligne d''activité' // Affaire
      {32}, 'Le fournisseur principal n''existe pas'
      {33}, 'Vous n''avez pas calculé la base tarif'
      {34}, 'Le type d''article financier est obligatoire'
      {35}, 'Article en mouvement, impossible de modifier sa gestion par lot'
      {36}, 'Impossible de dé référencer un fournisseur principal' //Catalogue
      {37}, 'Calcul du coefficient impossible, le prix de référence est nul'
      {38}, 'Le type de flux est obligatoire'
      {39}, 'Ce code barres est déjà affecté à un autre article'
      {40}, 'Cet article est déjà mouvementé, vous ne pouvez pas modifier le fournisseur'
      {41}, 'Version de démonstration : le nombre d''articles excède le nombre autorisé.'
      {42}, 'Opération non disponible, il existe du stock sur cet article'
      {43}, 'Opération non disponible, il existe un tarif détail sur cet article'
      {44}, 'Les prix des articles dimensionnés vont être mis à jour !'
      {45}, 'Confirmation'
      {46}, 'Cette TVA n''existe pas'
      {47}, 'Cette TPF n''existe pas'
      {48}, 'Choix d''un type d''ouvrage'
      {49}, 'Choix d''un type de nomenclature'
      {50}, 'Confirmez-vous la suppression du catalogue sur l''ancien fournisseur?'
      {51}, 'Le code article est obligatoire sur le catalogue'
      {52}, 'La saisie du champs suivant est obligatoire : '
      {53}, 'Article indisponible, vous devez auparavant valider votre saisie'
      {54}, 'Cette collection n''existe pas'
      {55}, 'Cette valeur n''existe pas'
      {56}, 'La nature de prestation n''existe pas'
      {57}, 'Plusieurs utilisateurs essayent de créer le même article, veuillez valider de nouveau'
      {58}, 'Renseignez la prestation associée !'
      {59}, 'Cet article existe déjà'
      {60}, 'Le conditionnement ne peut pas être celui par défaut'
      {61}, 'Vous avez saisi un profil mais n''avez pas appliqué les valeurs de ce profil.' + #13 + ' Souhaitez-vous valider?'
      {62}, 'Caractère interdit dans le code article'
      {63}, 'Suppression interdite, cet article est utilisé dans les mouvements de stock'
      {64}, 'Cet article est déjà utilisé dans un autre profil avancé'
      {65}, 'La limite du nombre maximum de champ est atteinte'
      {66}, 'Ce profil avancé est utilisé dans des articles'
      {67}, 'L''article de référence n''existe pas'
      {68}, 'Aucun champ à inclure - Vérifier le mode de copie ou renseigner les champs à inclure.'
      {69}, 'La mise à jour des articles utilisant un profil avancé qui préserve les différences '
            + 'est impossible depuis la fiche profil.' + #13 + ' ' + #13 + ' Pour se faire il faut autoriser la mise à jour du profil depuis la fiche article et,'
            + #13 + ' après modification d''un article profil, valider la modification des articles associés.'
      {70}, 'L''article de référence n''est pas renseigné'
      {71}, 'L''unité est obligatoire'
      {72}, 'Domaine différent de celui qui vous est assigné'
      {73}, 'Cet article parc est inexistant'
      {74}, 'Du stock existe pour le catalogue#13Le référencement a un article n''est pas possible'
      {75}, 'La marque est obligatoire'
      {76}, 'Il manque des formules de conversion entre les unités de la fiche article'
      {77}, 'Le domaine de l''article n''est pas valide.'
      {78}, 'Un conditionnement par défaut existe pour ce flux.'
      );

{***********A.G.L.***********************************************
Auteur  ...... : JTR - Inventaire Contremarque
Créé le ...... : 18/05/2005
Modifié le ... :
Description .. : Messages liés à l'inventaire de Contremarque
Mots clefs ... : SaisieInv
*****************************************************************}
const MSGERR_InvContreM : array[1..14] of string 	= (
      {1}   'Confirmation'
      {2}  ,'Le stock inventorié va être forcé à la valeur du stock ordinateur, sur toutes les lignes non saisies !'
      {3}  ,'Le stock inventorié va être remis à zéro, sur toutes les lignes saisies !'
      {4}  ,'Si le stock ordinateur est négatif, le stock inventorié sera forcé à zéro.'
      {5}  ,'L''emplacement n''existe pas pour ce dépôt. %s'
      {6}  ,'Le lot est obligatoire. %s'
      {7}  ,'Le n° de série est obligatoire. %s'
      {8}  ,'La quantité doit être unitaire en gestion par n° de série. %s'
      {9}  ,'L''article existe déjà dans la liste avec les mêmes caractéristiques. %s'
      {10} ,'Le n° de série existe déjà dans la liste pour ce même article. %s'
      {11} ,'Inventaire de type initialisation des stocks. Il existe des mouvements sur cet article/dépôt. Création refusée. %s'
      {12} ,'Le n° de série existe déjà pour ce même article dans la liste : %s'
      {13} ,'Le ' + '%s' + ' contient des caractères non autorisés'
      {14} ,'Impossible de saisir une quantité <> 0 car le n° de série et/ou le n° de lot contiennent des caractères non-autorisés.'
      );

{***********A.G.L.***********************************************
Auteur  ...... : JTR - Validation d'inventaire
Créé le ...... : 18/05/2005
Modifié le ... :
Description .. : Messages liés à la validation des inventaires
Mots clefs ... : UtofValidInv
*****************************************************************}
const MSGERR_ValidInv: array[1..2] of string 	= (
      {1}  'Vous devez renseigner le Tiers Ecarts d''inventaire dans les Paramètes Société'
      {2} ,'Veuillez sélectionner au moins une liste à valider'
      );

{***********A.G.L.***********************************************
Auteur  ...... : JTR - Traitement de validation d'inventaire
Créé le ...... : 18/05/2005
Modifié le ... :
Description .. : Messages liés à la validation des inventaires
Mots clefs ... : ValidateInv
*****************************************************************}
const MSGERR_ValidInvTrt : array[1..25] of string 	= (
      {1}   'Validation en cours...'
      {2}  ,'Validation effectuée'
      {3}  ,'&Annuler'
      {4}  ,'&Fermer'
      {5}  ,'Ordre des prix retenu :'
      {6}  ,'Abandon'
      {7}  ,'Validation des listes interrompue par l''utilisateur'
      {8}  ,'-- Traitement interrompu par l''utilisateur --'
      {9}  ,'En cours...'
      {10} ,'Non validée'
      {11} ,'Validée'
      {12} ,'Saisie incomplète'
      {13} ,'(une ou plusieurs lignes non inventoriées)'
      {14} ,'Saisie en cours'
      {15} ,'(en cours de traitement par un autre utilisateur)'
      {16} ,'Echec'
      {17} ,'Validation interrompue'
      {18} ,'Tiers Ecarts inventaire inconnu'
      {19} ,''
      {20} ,'Impossible de déterminer le client par défaut.'
      {21} ,'Erreur à la création de la pièce'
      {22} ,'Erreur à la création d''une ligne'
      {23} ,'Impossible de trouver cet article :'
      {24} ,'Erreur à la validation de l''inventaire'
      {25} ,'Impossible de créer une pièce sans dépôt'
      );

{***********A.G.L.***********************************************
Auteur  ...... : JTR
Créé le ...... : 25/07/2005
Modifié le ... :
Description .. : Messages liés à la remise en banque
Mots clefs ... : GCREMBANQUEESP_TOF
*****************************************************************}
const MSGERR_RemBqe : array[1..4] of string 	= (
      {1}  'Aucun élément sélectionné'
      {2}  ,'Traitement terminé'
      {3}  ,'La banque est incorrecte'
      {4}  ,'Ce mode de paiement n''est pas paramétré pour le module Ventes comptoir' 
      );

{***********A.G.L.***********************************************
Auteur  ...... : JTR
Créé le ...... : 15/09/2005
Modifié le ... :
Description .. : Messages liés à la comptabilisation des stocks
Mots clefs ... : GCCPTASTOCK_TOF
*****************************************************************}
const MSGERR_CptaStock : array[1..10] of string 	= (
      {1}   'Le journal comptable n''est pas renseigné.'
      {2},  'Veuillez confirmer la comptabilisation des stocks déjà comptabilisés.'
      {3},  'Veuillez confirmer la comptabilisation des stocks.'
      {4},  'Traitement terminé'
      {5},  'Aucune ligne à traiter'
      {6},  'Erreur lors de la numération des écritures'
      {7},  'Erreur lors de la création des écritures'
      {8},  'Veuillez sélectionner des établissements'
      {9},  'Erreur lors de la mise à jour des mouvements de stock'
      {10}, 'Il n''y a pas de nature de mouvement de stock paramétrée pour la comptabilisation'
      );

{***********A.G.L.***********************************************
Auteur  ...... : JTR
Créé le ...... : 20/09/2005
Modifié le ... :
Description .. : Messages liés aux entrées exceptionnelles
Mots clefs ... : GCSTKMOTIF_TOF
*****************************************************************}
const MSGERR_EntExcept : array[1..20] of string 	= (
	    {1}  'Dépôt obligatoire.'
      {2}  ,'Article obligatoire.'
			{3}  ,'Motif obligatoire.'
      {4}  ,'Article non géré en n° de série.'
      {5}  ,'Veuillez renseigner la dimension.'
      {6}  ,'Veuillez renseigner la dimension.'
      {7}  ,'Veuillez renseigner la dimension.'
      {8}  ,'Veuillez renseigner la dimension.'
      {9}  ,'Veuillez renseigner la dimension.'
      {10} ,'L''article est non géré en stock.'
      {11} ,'Stock disponible insuffisant.'
      {12} ,'Article inexistant.'
      {13} ,'Article inexistant pour votre domaine.'
      {14} ,'L''affaire n''existe pas.'
      {15} ,'Le fournisseur est obligatoire.'
      {16} ,'Erreur sur la référence.'
      {17} ,'Cet article n''est pas géré en contremarque.'
      {18} ,'Article ou référence de contremarque inexistant.'
      {19} ,'Choix abandonné.'
      {20} ,''
		  );

{***********A.G.L.***********************************************
Auteur  ...... : JTR
Créé le ...... : 26/09/2005
Modifié le ... :
Description .. : Messages liés aux sorties exceptionnelles
Mots clefs ... : STKDEPLACEMENT_TOF
*****************************************************************}
Const	MSGERR_SortExcept: array[1..12] of string = (
      {1}   'Le dépôt est obligatoire.'
      {2}  ,'L''emplacement est obligatoire.'
      {3}  ,'L''emplacement n''existe pas pour le dépôt.'
      {4}  ,'L''emplacement de destination doit être différent de l''emplacement de sélection.'
      {5}  ,'Le statut de disponibilité de l''emplacement est imposé.'
      {6}  ,'Le motif mouvement est obligatoire.'
      {7}  ,'L''affaire est obligatoire.'
      {8}  ,'L''affaire n''existe pas.'
      {9}  ,'Le client est obligatoire.'
      {10}  ,'Le client n''existe pas.'
		  {11}  ,'Le statut de disponibilité est obligatoire.'
      {12}  ,'Le statut de flux est obligatoire.'
    );

{***********A.G.L.***********************************************
Auteur  ...... : JTR
Créé le ...... : 16/11/2005
Modifié le ... :
Description .. : Messages liés à la gestion des arrondis
Mots clefs ... : GCARRONDI_TOM
*****************************************************************}
Const MSG_Arrondi : array[1..8] of string = (
	    {1}   'Libellé d''arrondi est obligatoire'
      {2}  ,'Code d''arrondi est obligatoire'
      {3}  ,'L''emplacement d''arrivée n''est pas renseigné.'
      {4}  ,'La ligne est incorrecte'
      {5}  ,'Voulez-vous enregistrer les modifications ?'
      {6}  ,'Valeur est incohérente,sélectionnez les valeurs en cliquant sur la cellule constante méthode'
      {7}  ,'Erreur!,Il faut au moins une ligne dans un tableau'
      {8}  ,'Valeur constante est incohérente'
		  );

{***********A.G.L.***********************************************
Auteur  ...... : BBI
Créé le ...... : 16/11/2005
Modifié le ... :
Description .. : Messages liés au paramétrage ABC
Mots clefs ... : ParametrageABC_TOM
*****************************************************************}
Const MSG_ParamABC_TOM: array[1..5] of string = (
    {1}  'La formule est incorrecte'
    {2} ,'La nature de pièce doit être renseignée'
    {3} ,'La fonction est obligatoire pour tout autre flux que l''article'
    {4} ,'Saisie incorrecte. La somme des pourcentages doit être égale à 100'
    {5} ,'Saisie incorrecte. Vous ne pouvez renseigner un seuil sans définir son libellé'
        );

{***********A.G.L.***********************************************
Auteur  ...... : BBI
Créé le ...... : 16/11/2005
Modifié le ... :
Description .. : Messages liés au paramétrage ABC
Mots clefs ... : ParamOblig_TOF
*****************************************************************}
Const MSG_ParamOblig_TOF: array[1..2] of string 	= (
          {1}  'Voulez-vous enregistrer les modifications ?'
          {2}  ,'Vos modifications seront actives à la prochaine connexion.'
          );

{***********A.G.L.***********************************************
Auteur  ...... : BBI
Créé le ...... : 16/11/2005
Modifié le ... :
Description .. : Messages liés au paramétrage ABC
Mots clefs ... : PieceEpure_TOF
*****************************************************************}
const MSG_PieceEpure_TOF : array[1..4] of string 	= (
          {1}   'Le traitement s''est correctement effectué.'
          {2}  ,'Le traitement ne s''est pas terminé correctement.'
          {3}  ,'Le traitement a été interrompu'
          {4}  ,'Confirmez vous la suppression des documents sélectionnés ?'
          );

{***********A.G.L.***********************************************
Auteur  ...... : JTR
Créé le ...... : 16/11/2005
Modifié le ... :
Description .. : Messages liés Bascule de la gestion des unités de mesure en mode avancé
Mots clefs ... : gcBasculeUnite_Tof
*****************************************************************}
const	MSG_BasculeUnite : array[1..5] of string = (
      {1}   'La bascule des unités a déjà été effectuée.'
      {2}  ,'La table des unités contient des doublons, le lancement de la bascule est impossible.'
      {3}  ,'La table des unités contient des quotités nulles, le lancement de la bascule est impossible.'
      {4}  ,'Le qualifiant par défaut n''est pas renseigné.'
      {5}  ,'L''unité par défaut n''est pas renseignée.'
      );

{***********A.G.L.***********************************************
Auteur  ...... : JTR
Créé le ...... : 16/11/2005
Modifié le ... :
Description .. : Messages liés à l'archivage des mouvements de stock
Mots clefs ... : ArchivageMouvement_TOF
*****************************************************************}
const MSG_ArchiveMvt : array[1..3] of string = (
      {1}  'Les mouvements seront archivés jusqu''au '
      {2} ,'Les mouvements de stock ne peuvent être archivés au delà du '
      {3} ,''
      );

{***********A.G.L.***********************************************
Auteur  ...... : JTR
Créé le ...... : 16/11/2005
Modifié le ... :
Description .. : Messages liés à l'affectation des mvt de contremarque
Mots clefs ... : GCAFFECTCONTREM_TOF
*****************************************************************}
const MSG_AffecteContreM: array[1..1] of string 	= (
      {1} 'Sans fournisseur, la ligne ne sera plus de contremarque. Voulez vous continuer ?'
      );

{***********A.G.L.***********************************************
Auteur  ...... : JTR
Créé le ...... : 18/11/2005
Modifié le ... :
Description .. : Messages liés à la gestion des Articles/Qtés
Mots clefs ... : ArticleQte_Tom
*****************************************************************}
const MSG_ArticleQte: array[1..2] of string 	= (
      {1} 'Le code article doit être renseigné'
      {2} ,''
      );

{***********A.G.L.***********************************************
Auteur  ...... : JTR
Créé le ...... : 18/11/2005
Modifié le ... :
Description .. : Messages liés à la duplication de nomenclature
Mots clefs ... : ChoixDupNomen
*****************************************************************}
const MSG_DupNomen: array[1..3] of string 	= (
      {1}  'La référence de la nomenclature existe déjà'
      {2} ,'La référence de l''ouvrage existe déjà'
      {3} ,'Veuillez renseigner le code'
      );

{***********A.G.L.***********************************************
Auteur  ...... : JTR
Créé le ...... : 18/11/2005
Modifié le ... :
Description .. : Messages liés aux formules de calcul ABC
Mots clefs ... : FormuleCalculABC_TOF
*****************************************************************}
const MSG_FrmlCalcABC : array[1..3] of string = (
      {1}  'La formule est incorrecte'
      {2} ,'Les champs doivent être de type numérique'
      {3} ,''
      );

{***********A.G.L.***********************************************
Auteur  ...... : JTR
Créé le ...... : 18/11/2005
Modifié le ... :
Description .. : Messages liés aux Articles/Parc
Mots clefs ... : GCArticleParc_Tom
*****************************************************************}
const MSG_ArticleParc : array[1..4] of string = (
      {1}  'Suppression impossible : article parc référencé dans les articles'
      {2} ,'L''activation de la gestion de versions#13 vous oblige à créer au moins une version'
      {3} ,'Suppression impossible : article parc référencé dans les versions',
      {4} 'La saisie du champ suivant est obligatoire : '
      );

{***********A.G.L.***********************************************
Auteur  ...... : JTR
Créé le ...... : 18/11/2005
Modifié le ... :
Description .. : Messages liés aux Commandes ouvertes
Mots clefs ... : GcCdeOuv_TOF
*****************************************************************}
const MSG_CdeOuv : array[1..6] of string = (
      {1}  'Vous ne pouvez pas saisir une quantité négative.'
      {2} ,'Vous n''avez pas renseigné l''article.'
      {3} ,'L''article : %s n''existe pas.'
      {4} ,'Vous ne pouvez pas saisir une périodicité négative.'
      {5} ,'Vous ne pouvez pas saisir un nombre de cadence négatif ou nul.'
      {6} ,'Vous n''avez pas renseigné le type de cadence.'
		  );

{***********A.G.L.***********************************************
Auteur  ...... : JTR
Créé le ...... : 18/11/2005
Modifié le ... :
Description .. : Messages liés à la gestion du colisage
Mots clefs ... : GcColisage_Tom
*****************************************************************}
const MSG_Colisage : array[1..3] of string = (
      {1} 'Le poids brut d''une unité logistique ne peut être inférieur à son poids net.'
      {2} ,'Vous devez choisir un préparateur existant dans la liste.'
      {3} ,'Vous devez choisir un emballeur existant dans la liste.'
      );

{***********A.G.L.***********************************************
Auteur  ...... : JTR
Créé le ...... : 18/11/2005
Modifié le ... :
Description .. : Messages liés à la gestion du colisage
Mots clefs ... : AssistAffecteDepot
*****************************************************************}
const MSG_AssAffecteDepot : array[1..10] of string = (
      {1}  'INFORMATION'
      {2}  ,'Tous les mouvements de la base contiennent un dépôt.'+#13+#10+'L''éxécution de cet utilitaire n''est pas nécessaire.'
      {3}  ,'CONFIRMATION'
      {4}  ,'Veuillez confirmer la mise à jour des mouvements sans dépôt.'
      {5}  ,'ERREUR'
      {6}  ,'Le code du dépôt ne peut être supérieur à 3 alphanumériques.'
      {7}  ,'Le libellé du dépôt ne peut être supérieur à 35 alphanumériques.'
      {8}  ,'Ce code de dépôt existe déjà.'
      {9}  ,''
      {10} ,''
      );

{***********A.G.L.***********************************************
Auteur  ...... : JTR
Créé le ...... : 18/11/2005
Modifié le ... :
Description .. : Messages liés au RAZ de l'activité
Mots clefs ... : AssistRazActivite
*****************************************************************}
const MSG_AssRazActivite : array[1..8] of string = (
      {1}  'ERREUR'
      {2}  ,'ATTENTION'
      {3}  ,'CONFIRMATION'
      {4}  ,'Le mot de passe est incorrect.'
      {5}  ,'Il faut impérativement une sauvegarde récente de votre base avant de continuer.#13#10Si vous n''en avez pas, cliquez sur "Abandonner" pour annuler l''opération sinon, cliquez sur "Oui" pour continuer.'
      {6}  ,'Veuillez confirmer l''éxécution de la remise à zéro de l''activité'
      {7}  ,'Cet utilitaire supprime définitivement toutes les informations liées à la société.#13#10Veuillez confirmer...'
      {8}  ,'L''éxécution de cet utilitaire est impossible tant qu''il y a d''autres utilisateurs connectés sur la base.'
      );

{***********A.G.L.***********************************************
Auteur  ...... : JTR
Créé le ...... : 18/11/2005
Modifié le ... :
Description .. :
Mots clefs ... : EditAcompteDiff_TOF
*****************************************************************}
const MSG_EditAcpteDiff : array[1..3] of string = (
     {1}  'Aucun élément sélectionné'
     {2}, 'Aucun état modèle paramétré'
     {3}, 'Trop de documents sélectionnés. Veuiller réduire la sélection'
     );

{***********A.G.L.***********************************************
Auteur  ...... : JTR
Créé le ...... : 15/11/2005
Modifié le ... :
Description .. : Messages liés aux Paramétrage des natures de pièces
Mots clefs ... : UtomPiece.pas
*****************************************************************}
Const	MSG_PARPIECE: array[1..23] of string = (
      {1}    'Le code "Nature de Pièce" doit être renseigné'
      {2}   ,'L''intitulé doit être renseigné'
      {3}   ,'Le type de flux doit être renseigné'
      {4}   ,'Le mode de création doit être renseigné'
      {5}   ,'La nature suivante XXX n''est pas une nature de pièce existante'
      {6}   ,'Les pièces suivantes sont circulaires : XXX'
      {7}   ,'La nature équivalente XXX n''est pas une nature de pièce existante'
      {8}   ,'Cette nature de pièce ne peut être supprimée, des pièces y font références'
      {9}   ,'Suppression impossible, cette nature est référencée comme nature suivante pour le type de pièce XXX'
      {10}  ,'Le montant de visa doit être supérieur ou égal à zéro'
      {11}  ,'Vos modifications seront actives à la prochaine connexion'
      {12}  ,'Le domaine d''activité doit être renseigné'
      {13}  ,'Vous ne pouvez pas utiliser les tables libres si aucune d''elles n''est renseignée'
      {14}  ,'ATTENTION !!! La Gestion des reliquats et la Gestion des lots sont activées.'+#13+' Il sera impossible de générer des reliquats sur les articles gérés en lots.'
      {15}  ,'La nature suivante XXX ne peut être sélectionnée.#13 Se reporter à l''aide sur le "Solde avant transformation".'
      {16}  ,'Le modèle WORD n''existe pas'
      {17}  ,'Liste affaire'
      {18}  ,'Liste de saisie'
      {19}  ,'ERREUR !!! Une des pièces précédentes est paramétrée avec la comptabilisation "Normale"'
      {20}  ,'Certains comptes de TVA ou TPF n''ont pas de comptes FAR FAE'
      {21}  ,'Le compte FAR FAE est incorrect'
      {22}  ,'La gestion de FAR FAE ne peut se faire que sur un passage en comptabilité de type "Ecriture normale"'
      {23}  ,'ERREUR !!! Les nomenclatures ne sont pas acceptées dans les pièce d''achat '
    );

{***********A.G.L.***********************************************
Auteur  ...... : JTR
Créé le ...... : 15/11/2005
Modifié le ... :
Description .. : Messages liés à la saisie de pièces
Mots clefs ... : Facture.pas
*****************************************************************

ATTENTION !!! Les array sont en décalage de -1 par rapport aux THMsgBox
exemple : MSG_FactErr[1] = Herr.execute[0]
}

Const MSG_FactTitre: array[1..30] of string = ( // = HTitres
      {1}   'Liste des articles'
      {2}  ,'Articles'
      {3}  ,'Tiers'
      {4}  ,'Liste des affaires'
      {5}  ,'Affaires'
      {6}  ,'ATTENTION : Pièce non enregistrée !'
      {7}  ,'ATTENTION : Cette pièce en cours de traitement par un autre utilisateur n''a pas été enregistrée !'
      {8}  ,'Liste des commerciaux'
      {9}  ,'Conditionnements'
      {10}  ,'Conditionnement en cours :'
      {11}  ,'Non affecté'
      {12}  ,'Liste des commerciaux'
      {13}  ,'Liste des dépôts '
      {14}  ,'ATTENTION : Cette pièce ne peut pas passer en comptabilité et n''a pas été enregistrée !'
      {15}  ,'Euro '
      {16}  ,'Autres'
      {17}  ,'Libellés automatiques'
      {18}  ,'(référence de substitution possible :'
      {19}  ,'(remplacement par référence :'
      {20}  ,'ATTENTION : l''impression ne s''est pas correctement effectuée !'
      {21}  ,'ATTENTION : la suppression ne s''est pas correctement effectuée !'
      {22}  ,'Crédit accordé :'
      {23}  ,'Encours actuel :'
      {24}  ,'ATTENTION : La pièce présente un problème de numérotation et n''a pas été enregistrée !'
      {25}  ,'Client'
      {26}  ,'Fournisseur'
      {27}  ,''
      {28}  ,'ATTENTION. Le stock disponible est insuffisant pour certains articles.'
      {29}  ,'Changement de code TVA'
      {30}  ,'Changement du régime fiscal de la pièce'
      );

      MSG_FactPiece: array[1..84] of string = ( // = HPiece
      {1}   '1;?caption?;Il n''existe pas de tarif pour cet article correspondant à la devise de la pièce !;E;O;O;O;'
      {2}  ,'2;?caption?;Saisie impossible. Il n''existe pas de tarif pour cet article correspondant à la devise de la pièce !;E;O;O;O;'
      {3}  ,'3;?caption?;Saisie impossible. Cet article fermé n''est pas autorisé pour cette nature de pièce;E;O;O;O;'
      {4}  ,'4;?caption?;Choix impossible. Ce tiers fermé ne peut pas être appelé pour cette nature de pièce;E;O;O;O;'
      {5}  ,'5;?caption?;Choix impossible. La nature du tiers associé à l''affaire n''est pas compatible avec la nature de pièce;E;O;O;O;'
      {6}  ,'6;?caption?;Vous devez renseigner un tiers valide;E;O;O;O;'
      {7}  ,'7;?caption?;Confirmez-vous l''abandon de la saisie ?;Q;YN;Y;N;'
      {8}  ,'8;?caption?;Voulez-vous affecter ce dépôt sur toutes les lignes concernées ?;Q;YN;Y;N;'
      {9}  ,'9;?caption?;Choix impossible. Ce code tiers est incorrect;E;O;O;O;'
      {10}  ,'10;?caption?;Cette pièce ne peut pas être modifiée ni transformée, seule la consultation est autorisée;E;O;O;O;'
      {11}  ,'11;?caption?;Cet article est en rupture, confirmez-vous malgré tout la quantité ?;Q;YN;Y;N;'
      {12}  ,'12;?caption?;Cet article est en rupture ;E;O;O;O;'
      {13}  ,'13;?caption?;Cet article n''est plus géré ;E;O;O;O;'
      {14}  ,'14;?caption?;Le plafond de crédit accordé au tiers est dépassé ! ;E;O;O;O;'
      {15}  ,'15;?caption?;Le crédit accordé au tiers est dépassé ! ;E;O;O;O;'
      {16}  ,'16;?caption?;Voulez-vous saisir ce taux dans la table de chancellerie ?;Q;YN;Y;N;O;'
      {17}  ,'17;?caption?;ATTENTION : Le taux en cours est de 1. Voulez-vous saisir ce taux dans la table de chancellerie;Q;YN;Y;N;O;'
      {18}  ,'18;?caption?;ATTENTION : la marge de la ligne est inférieure au minimum requis;E;O;O;O;'
      {19}  ,'19;?caption?;La date que vous avez renseignée n''est pas valide;E;O;O;O;'
      {20}  ,'20;?caption?;La date que vous avez renseignée n''est pas dans un exercice ouvert;E;O;O;O;'
      {21}  ,'21;?caption?;La date que vous avez renseignée est antérieure à une clôture;E;O;O;O;'
      {22}  ,'22;?caption?;La date que vous avez renseignée est antérieure à une clôture;E;O;O;O;'
      {23}  ,'23;?caption?;La date que vous avez renseignée est en dehors des limites autorisées;E;O;O;O;'
      {24}  ,'24;?caption?;ATTENTION : la marge de la ligne est inférieure au minimum requis. Voulez-vous changer de code utilisateur ?;Q;YN;N;N;'
      {25}  ,'25;?caption?;Le plafond de crédit accordé au tiers est dépassé. Voulez-vous changer de code utilisateur ? ;Q;YN;N;N;'
      {26}  ,'26;?caption?;Vous ne pouvez pas saisir avant le ;E;O;O;O;'
      {27}  ,'27;?caption?;Cette pièce contient déjà des lignes. Reprise des lignes de l''affaire impossible ;E;O;O;O;'
      {28}  ,'28;?caption?;Attention : Vous n''avez pas affecté d''affaire sur cette pièce ;E;O;O;O;'
      {29}  ,'29;?caption?;La date est antérieure à celle de dernière clôture de stock;E;O;O;O;'
      {30}  ,'30;?caption?;Confirmez-vous la suppression de la pièce ?;Q;YN;N;N;'
      {31}  ,'31;?caption?;Voulez-vous répercuter la date de livraison de l''entête sur toutes les lignes?;Q;YN;Y;N;O;'
      {32}  ,'32;?caption?;Voulez-vous répercuter la date de livraison de l''entête sur les lignes sélectionnées ?;Q;YN;Y;N;O;'
      {33}  ,'33;?caption?;Vous devez renseigner une devise;E;O;O;O;'
      {34}  ,'34;?caption?;Vous avez enregistré des acomptes ou des règlements. Voulez-vous les supprimer ?;Q;YN;Y;N;O;'
      {35}  ,'35;?caption?;Ce client est en rouge ! Vous ne pouvez pas lui créer de pièce à  ce niveau de risque;E;O;O;O;'
      {36}  ,'36;?caption?;Confirmez-vous la suppression de tout le paragraphe ?;Q;YN;N;N;'
      {37}  ,'37;?caption?;Cet article est tenu en stock, vous ne pouvez le contremarquer;E;O;O;O;'
      {38}  ,'38;?caption?;Cette pièce contient déjà des lignes, vous ne pouvez pas changer le fournisseur;E;O;O;O;'
      {39}  ,'39;?caption?;Confirmez vous le changement du code TVA sur toutes les lignes saisies ?;Q;YN;Y;N;'
      {40}  ,'40;?caption?;Confirmez vous le changement du régime fiscal de la pièce ?;Q;YN;Y;N;'
      {41}  ,'41;?caption?;Erreur sur le code du commercial;E;O;O;O;'
      {42}  ,'42;?caption?;Vous devez renseigner un motif par défaut;E;O;O;O;'
      {43}  ,'43;?caption?;Vous ne pouvez pas saisir une quantité négative;E;O;O;O;'
      {44}  ,'44;?caption?;Cette pièce ne peut pas être modifiée, seule la dernière situation est modifiable;E;O;O;O;'
      {45}  ,'45;?caption?;Certains articles n''ont pu être recupérés;E;O;O;O;'
      {46}  ,'46;?caption?;Génération impossible : le paramétrage est incomplet;E;O;O;O;'
      {47}  ,'47;?caption?;Date inférieure à arrêté de période;E;O;O;O;'
      {48}  ,'48;?caption?;Saisie impossible. Cet article n''est gérable qu''en contremarque.;E;O;O;O;'
      {49}  ,'49;?caption?;La saisie par code barre n''a pas été validée, voulez-vous la valider ?;Q;YN;Y;N;'
      {50}  ,'50;?caption?;Le taux d''escompte ne peut excèder 100%. ;E;O;O;O;'
      {51}  ,'51;?caption?;Le taux de remise ne peut excèder 100%. ;E;O;O;O;'
      {52}  ,'52;?caption?;Attention ! La quantité saisie est inférieure au total des quantités des pièces suivantes.;E;O;O;O;'
      {53}  ,'53;?caption?;Vous ne pouvez pas supprimer cette ligne, d''autres lignes y sont associées.;E;O;O;O;'
      {54}  ,'54;?caption?;Vous ne pouvez pas changer l''article de cette ligne.;E;O;O;O;'
      {55}  ,'55;?caption?;Toutes les lignes de la pièce d''origine ont été transformées.;E;O;O;O;'
      {56}  ,'56;?caption?;ATTENTION: La suppression est impossible car de l''activité a été saisie ;E;O;O;O;'
      {57}  ,'57;?caption?;La quantité de la ligne est supérieure au maximum autorisé;E;O;O;O;'
      {58}  ,'58;?caption?;Le dépôt est obligatoire.;E;O;O;O;'
      {59}  ,'59;?caption?;Confirmez-vous le solde de cette ligne ?;Q;YN;N;N;'
      {60}  ,'60;?caption?;Voulez-vous dé-solder de cette ligne ?;Q;YN;N;N;'
      {61}  ,'61;?caption?;Vous ne pouvez pas supprimer cette ligne, la ligne de la pièce précédente est soldée.;E;O;O;O;'
      {62}  ,'62;?caption?;Confirmez-vous le solde de toutes les lignes de la piece ?;Q;YN;N;N;'
      {63}  ,'63;?caption?;Cet article n''a pas de catalogue, vous ne pouvez le contremarquer.;E;O;O;O;'
      {64}  ,'64;?caption?;Cet article est géré en lot ou en numéro de série, vous ne pouvez le contremarquer;E;O;O;O;'
      {65}  ,'65;?caption?;La quantité saisie est supérieure à la quantité restant à lancer;E;O;O;O;'
      {66}  ,'66;?caption?;Le n° d''ordre est obligatoire.;E;O;O;O;'
      {67}  ,'67;?caption?;La phase est obligatoire.;E;O;O;O;'
      {68}  ,'68;?caption?;Le traitement est obligatoire.;E;O;O;O;'
      {69}  ,'69;?caption?;Voulez-vous affecter la modification aux lignes commissionnables de la pièce ?;Q;YN;Y;N;'
      {70}  ,'70;?caption?;Confirmez vous le changement du code TVA sur toutes les lignes sélectionnées ?;Q;YN;Y;N;'
      {71}  ,'71;?caption?;Un compteur différent est paramétré pour cet établissement. Le changement est impossible;E;O;O;O;'
      {72}  ,'72;?caption?;Confirmez-vous le dé-solde de toutes les lignes de la piece ?;Q;YN;N;N;'
      {73}  ,'73;?caption?;La facture d''acompte ne pourra être produite qu''une fois comptabilisée;E;O;O;O;'
      {74}  ,'74;?caption?;Une prestation école existe déjà.;E;O;O;O;'
      {75}  ,'75;?caption?;Vous devez saisir une prestation de type planification article école.;E;O;O;O;'
      {76}  ,'76;?caption?;Voulez-vous modifier le code affaire sur toutes les lignes ?;Q;YN;Y;N;'
      {77}  ,'77;?caption?;Ce tiers n''a pas de commercial;W;O;O;O;'
      {78}  ,'78;?caption?;Choix impossible. Ce type de tiers n''est pas autorisé pour la nature de pièce;E;O;O;O;'
      {79}  ,'79;?caption?;Choix impossible. Cet article n''a pas de nomenclature d''assemblage;E;O;O;O;'
      {80}  ,'80;?caption?;Cet article ne peut plus être utilisé, sa date de suppression est dépassée;E;O;O;O;'
      {81}  ,'81;?caption?;L''article %s n''existe pas;E;O;O;O;'
      {82}  ,'82;?caption?;Choix impossible. La nomenclature d''assemblage n''a pas de composants;E;O;O;O;'
      {83}  ,'83;?caption?;La date de pièce n''est pas comprise entre les dates limites des paramètres société !;E;O;O;O;'
      {84}  ,'84;?caption?;L''article n''est pas tenu en Stock !;E;O;O;O;'
         );

      MSG_FactAveugle: array[1..5] of String = ( // = HAveugle
      {1}   'ATTENTION : code barre inconnu.'
      {2}  ,'2;?Caption?;La quantité de cet article est supérieure à celle attendue dans le document. #13 Voulez-vous l''ajouter au document ?;Q;YN;Y;N;'
      {3}  ,'ATTENTION : vous ne pouvez pas saisir une quantité négative.'
      {4}  ,'3;?Caption?;Voulez-vous supprimer tous les articles saisis ?;Q;YN;Y;N;'
      {5}  ,'ATTENTION : cet article ne correspond pas à ce fournisseur.'
      );

      MSG_FactErr: array[1..17] of String = ( // = HErr
      {1}   'Erreurs en validation'
      {2}  ,'2;?caption?;Vous ne pouvez pas enregistrer une pièce vide;E;O;O;O;'
      {3}  ,'3;?caption?;Vous ne pouvez pas enregistrer une pièce sans articles;E;O;O;O;'
      {4}  ,'4;?caption?;Vous ne pouvez pas enregistrer une pièce sans mouvements;E;O;O;O;'
      {5}  ,'5;?caption?;Vous ne pouvez pas enregistrer une pièce à cette date;E;O;O;O;'
      {6}  ,'6;?caption?;Enregistrement impossible : l''acompte est supérieur au total de la pièce;E;O;O;O;'
      {7}  ,'7;?caption?;Enregistrement impossible : la devise est incorrecte;E;O;O;O;'
      {8}  ,'8;?caption?;La pièce d''origine est déjà traitée;E;O;O;O;'
      {9}  ,'9;?caption?;Le contrôle facture n''est pas valide;E;O;O;O;'
      {10}  ,'10;?caption?;Le circuit est obligatoire;E;O;O;O;'
      {11}  ,'11;?caption?;L''ordre, la phase et le traitement de production sont obligatoires;E;O;O;O;'
      {12}  ,'12;?caption?;Le code marché est obligatoire;E;O;O;O;'
      {13}  ,'13;?caption?;La date de début est supérieure à la date de fin;E;O;O;O;'
      {14}  ,'14;?caption?;Ce code marché existe déjà dans une autre pièce, pour le même couple article, tiers;E;O;O;O;'
      {15}  ,'15;?caption?;Marché : Doublon sur code marché et/ou chevauchement des dates de marché dans cette pièce;E;O;O;O;'
      {16}  ,'16;?caption?;Vous ne pouvez pas enregistrer une pièce avec une ou des lignes dont le montant est égal à zéro;E;O;O;O;'
      {17}  ,'17;?caption?;Les dates de validité ne couvrent pas les appels passés sur ce marché;E;O;O;O;'
      );

      MSG_FactTexteSortie: array[1..6] of string 	= ( // = TexteSortie
      {1}   'Vous ne pouvez pas modifier la quantité de cette pièce'
      {2}  ,'Vous ne pouvez pas facturer plus que ce que vous avez livré.'#13'Veuillez saisir une ligne de complément pour cet article'
      {3}  ,'Vous ne pouvez pas facturer plus que ce que vous avez réceptionné.'#13'Veuillez saisir une ligne de complément pour cet article'
      {4}  ,'Vous ne pouvez pas supprimer la ligne'
      {5}  ,'Vous ne pouvez pas changer le signe de la ligne'
      {6}  ,'Saisir la prestation à planifier et le nombre de participants'
      );
      
{***********A.G.L.***********************************************
Auteur  ...... : JSI
Créé le ...... : 08/02/2006
Modifié le ... :
Description .. : Messages liés aux contrôles de suppression
...............  et fermeture des tiers
Mots clefs ... :
*****************************************************************}
Const	MSGERR_SupprTiers : array[1..38] of string = (
        {1}        'Ce compte auxiliaire est mouvementé',
        {2}        'Ce compte auxiliaire est référencé dans un guide d''écriture comptable',
        {3}        'Ce compte auxiliaire sert de modèle pour la création',
        {4}        'Ce compte auxiliaire est référencé dans un catalogue',
        {5}        'Ce compte auxiliaire est référencé dans une ligne d''écriture',
        {6}        'Ce compte auxiliaire est référencé dans une ligne de pièce interne',
        {7}        'Ce compte auxiliaire est référencé dans une pièce',
        {8}        'Ce compte auxiliaire est référencé dans une pièce interne',
        {9}       'Ce compte auxiliaire est un compte de correspondance',
        {10}       'Ce compte auxiliaire est un client ou un fournisseur d''attente',
        {11}       'Ce compte auxiliaire est référencé dans une section analytique',
        {12}       'Ce compte auxiliaire est référencé pour un autre compte auxiliaire',
        {13}       'Ce compte auxiliaire est un utilisateur',
        {14}       '14;Suppression des comptes auxiliaires;Vous n''avez rien sélectionné;E;O;O;O;',
        {15}       '15;Suppression des comptes auxiliaires;Désirez-vous un compte rendu des comptes détruits?;Q;YNC;N;C;',
        {16}       '16;Suppression des comptes auxiliaires;Désirez-vous un compte rendu des comptes non détruits?;Q;YNC;N;C;',
        {17}       'Compte supprimé',
        {18}       'Compte en cours d''utilisation !',
        {19}       'Aucun',
        {20}       'élément sélectionné',
        {21}       'éléments sélectionnés',
        {22}       '22;',
        {23}       '23;',
        {24}       '24;',
        {25}       '25;',
        {26}       'Ce compte est mouvementé en pièces de Gestion Commerciale',
        {27}       'Ce compte est mouvementé en règlements de Gestion Commerciale',
        {28}       'Ce compte est référencé dans les activités',
        {29}       'Ce compte est référencé dans les ressources',
        {30}       'Ce compte est référencé en affaire',
        {31}       'Ce compte est référencé dans les actions',
        {32}       'Ce compte est référencé dans les propositions',
        {33}       'Ce compte est référencé comme salarié',
        {34}       'Ce compte est référencé dans le catalogue',
        {35}       'Ce compte est mouvementé dans la paie comme salarié',
        {36}       'Ce compte est référencé dans les projets',
        {37}       'Ce compte est référencé dans les documents GED',
        {38}       'Ce tiers a du parc en service'
          );

{***********A.G.L.***********************************************
Auteur  ...... : JTR
Créé le ...... : 02/02/2006
Modifié le ... :
Description .. : Messages liés à la fiche tiers
Mots clefs ... : UtomTiers.pas
*****************************************************************}
Const	MSG_Tiers : array[1..58] of string = (
      {1}  'Le code doit être renseigné'
      {2}  ,'La raison sociale doit être renseignée'
      {3}  ,'Le client payeur spécifié n''existe pas ou n''est pas un client payeur'
      {4}  ,'Le code du représentant n''existe pas'
      {5}  ,'Le code du tiers facturé n''existe pas ou n''est pas un client'
      {6}  ,'Le code de l''apporteur n''existe pas'
      {7}  ,'Le tiers n''est pas un client'
      {8}  ,'Le destinataire de l''adresse est obligatoire'
      {9}  ,'Le code postal de l''adresse est obligatoire'
      {10} ,'La ville de l''adresse est obligatoire'
      {11} ,'Les natures de pièces suivantes et courantes doivent être différentes'
      {12} ,'Vous ne pouvez choisir ce type de pièce suivante du fait de la circularité : XXX'
      {13} ,'Le plafond autorisé ne doit pas être inférieur au crédit accordé'
      {14} ,'Le montant de visa doit etre supérieur à zéro'
      {15} ,'Le compte collectif doit être renseigné#13Veuillez le renseigner dans les paramètres comptables'
      {16} ,'La nature d''auxiliaire n''est pas définie'
      {17} ,'Une fiche existe déjà avec ce code tiers'
      {18} ,'Le mode de règlement doit être renseigné'
      {19} ,'Le régime fiscal doit être renseigné'
      {20} ,'L''adresse doit être renseignée'
      {21} ,'Une fiche existe déjà avec ce code auxiliaire'
      {22} ,'Le mois de clôture à une valeur incorrecte'   // AFfaire
      {23} ,'Le compte Auxiliaire doit être renseigné '
      {24} ,'Le compte Collectif n''existe pas'
      {25} ,'Le code groupe n''existe pas'
      {26} ,'Le code prescripteur n''existe pas'
      {27} ,'Problème de création de la fiche Suspect'  // GRC
      {28} ,'Version de démonstration, vous n''êtes pas autorisé à créer plus de 10 tiers'
      {29} ,'Le numéro de Siret ou Siren n''est pas correct '
      {30} ,'Ce numéro de Siret existe déjà sur la fiche tiers : '// GRC
      {31} ,'Vous ne pouvez créer qu''une seule adresse de facturation' // Affaire
      {32} ,'Version de démonstration : le nombre de tiers excède le nombre autorisé.'
      {33} ,'Champ obligatoire : '
      {34} ,'Ce numéro de Siren existe déjà sur la fiche tiers : ' // mng : 34
      {35} ,'Date d''expiration inférieure à la date de délivrance' //CHR
      {36} ,'La saisie du champ suivant est obligatoire : '
      {37} ,'Le code du tiers livré n''existe pas ou n''est pas un client'
      {38} ,'Cette région n''appartient pas à ce pays'
      {39} ,'Le fournisseur payé spécifié n''existe pas'
      {40} ,'L''emetteur de factures spécifié n''existe pas'
      {41} ,'Le code NIF est obligatoire'
      {42} ,'Un code ressource libre est inexistant'
      {43} ,'Un code assistant libre est inexistant'
      {44} ,'Ce tiers n''est pas un tiers payeur'
      {45} ,'Le jour de naissance n''est pas correct'
      {46} ,'L''année de naissance doit être saisie sur 4 caractères'
      {47} ,'L''adresse de livraison n''existe pas'
      {48} ,'L''adresse de facturation n''existe pas'
      {49} ,'Le commercial est obligatoire'
      {50} ,'Cette enseigne n''existe pas'
      {51} ,'Ce transporteur n''existe pas'
      {52} ,'Ce fournisseur transporteur est utilisé dans les tarifs.' + #13 + ' Vous ne pouvez supprimer le type de fournisseur "transporteur".'
      {53} ,'Ce fournisseur transporteur est affecté à un tiers.' + #13 + ' Vous ne pouvez supprimer le type de fournisseur "transporteur".'
      {54} ,'Ce fournisseur transporteur est affecté à une adresse.' + #13 + ' Vous ne pouvez supprimer le type de fournisseur "transporteur".'
      {55} ,'Le compte n''est pas un collectif fournisseur.'
      {56} ,'Le compte n''est pas un collectif client.'
      {57} ,'La prestation de transport est incorrecte .'
      {58} ,'Le code frais renseigné est inconnu.'
      );

{***********A.G.L.***********************************************
Auteur  ...... : MNG
Créé le ...... : 23/03/2006
Modifié le ... :
Description .. : Messages liés à la gestion des alertes
Mots clefs ... : Alerte
*****************************************************************}
Const	MSG_Alertes : array[1..3] of string  = (
          {1}         'La zone "Alerte sur" doit être renseignée',
          {2}         'Vous devez sélectionner un ou plusieurs utilisateur(s)',
          {3}         'Vous devez sélectionner un ou plusieurs groupe(s)'
        );          
{***********A.G.L.***********************************************
Auteur  ...... : Thibaut Sublet
Créé le ...... : 25/08/2006
Modifié le ... :
Description .. : Messages d'erreur en gestion des pièces/lignes
Mots clefs ... : PREPAEXPE
*****************************************************************}
const MSGERR_PrepaExpe: Array[1..4] of String = (
          {1}  'Il existe des lignes extraites en préparation d''expédition.'
          {2} ,'Il existe du stock affecté par la préparation d''expédition.'
          {3} ,'Cette ligne est extraite en préparation d''expédition.'
          {4} ,'Il existe du stock affecté à cette ligne par la préparation d''expédition.'
  );

implementation

{***********A.G.L.***********************************************
Auteur  ...... : JTR
Créé le ...... : 18/11/2005
Modifié le ... :
Description .. : Creer un THMsgBox depuis un tableau de message
Mots clefs ... :
*****************************************************************}
procedure ChargeTHMsgBoxFromArray(var MsgBox : THMsgBox; Tableau : array of string);
var Cpt : integer;
begin
  for Cpt := 0 to High(Tableau) do
    MsgBox.Mess.add(Tableau[Cpt]);
end;

{***********A.G.L.***********************************************
Auteur  ...... : BBI
Créé le ...... : 19/04/2006
Modifié le ... :
Description .. : Indique la nature du message en prenant en compte
                 le caractere situé aprés le texte (E, W, Q)
Mots clefs ... :
*****************************************************************}
function GetTypeMsg(const Mess: String): TypeMsgErreur;
var stType : string;
begin
  stType := Mess;
  ReadTokenSt(stType);
  ReadTokenSt(stType);
  ReadTokenSt(stType);
  stType := ReadTokenSt(stType);
  Result := msNone;
  if (stType = 'E') or (stType = 'X') then Result := msError
  else if stType = 'W' then Result := msWarning
  else if stType = 'Q' then Result := msAsk
  ;
end;

end.
