{$IFNDEF DUCS41}
// DUCS 4.2
{***********UNITE*************************************************
Auteur  ...... : PAIE : MF
Créé le ...... : 12/04/2002

Modifié le ... : 22/05/2002
Description .. : Ensemble des procédures de la confection des
Suite ........ : fichiers Ducs Edi dossier
Mots clefs ... : PAIE, PGDUCSEDI
*****************************************************************}
{
 PT1 : 07/08/2002 : V585  MF   PGDUCSEDI
                             1- Traitement des champs code application et serveur
                             unique suite ajout dans table ORGANISMEPAIE.
                             2- traitement du cas UNEDIC : déclaration avec
                             paiement groupé (VLU) Autant de GP2 que d'établis-
                             sements secondaires déclarés --> Création de nouveaux
                             segments NADDCE, QTYTE, RFFDCE et RFFACE
                             3- La date de versement des salaires est renseignée
                             pour les IRC avec décalage de paie et pour l'UNEDIC.
                             Pour l'UNEDIC, s'il sagit d'une déclaration Néant
                             cette date est renseignée à 000000
                             4- Traitement du cas IRC multi centre payeur.
                             On a autant de segments de paiement que
                             d'établissements--> Création de nouveaux segments
                             (par type de paiement)
                             5- Traitement du paiement par virement et
                             modification du paiement par prélèvement ou
                             télérèglement (on récupère le RIB prévu pour le
                             paiement des charges sociales)
                             6- rectification du formatage de la référence sociale
                             utile au contrôle de la clé IBAN (chèque ou virement)
                             7- Correction : qd multi établissement l'édition de
                             contrôle ne faisait pas apparaître le bon
                             établissement.
                             8- Traitement des champs PES_MTPAYE, PES_DUCSNEANT
                             et PES_PAIEMODE ajouter à la table ENVOISOCIAL

 PT2 : 04/09/2002 : V582  MF  (pour version tests de validation)
                             1- Alimentation de la clé IBAN sur segment RFFPQ
 PT3 : 19/09/2002 : V585  MF  (pour version tests de validation)
                             1- correction alimentation segment FII. Récupé-
                             -ration du RIB des SALAIRES à partir de BQ_GENERAL
                             (et non pas BQ_CODE)
 PT4 : 24/09/2002 : V585  MF  (pour version tests de validation)
                             1- Le segment DTM du groupe 6 était renseigné avec
                             un mauvais format de date (CCYYDDMM au lieu de
                             CCYYMMDD)
                             2- Le segment MAOTOT était mal alimenté pour une
                             déclaration UNEDIC mono etab. Le Segement MOAPAI ne faisait
                             pas apparaître la monnaie de tenue.
 PT5 : 22/10/2002 : V585  MF Modification apportées suite au complément du cahier
                             charges (COPAYM V4.1 - mars 1999). Concerne le
                             Tableau Récapitulatif pour l'Acoss.
 PT6 : 23/10/2002 : V585  MF fiche qualité 10252 : le fichier .log est
                             maintenant nommé DUCSEDI+Siret.log
 PT7 : 06/01/2003 : V591  MF
                             1- Correction des avertissements de compile
                             2- Vérification que la nature Ducs est égale à A ou I ou U
                                Sinon anomalie signalée et fichier non généré
                             3- contrôles lors de la génération du fichier (dossier)
                                absence de RIB
                             4- Initialisation de l'enregistrement EnrDucsEdiGroupee
 PT8 : 08/01/2003 : V591  MF
                             1- Correction initialisation et tests sur doubles
                             2- Modification du calcul du montant à payer : =
                                Montant déclaré - acomptes + régularisations
                                (Idem consultation)
                                Modification du traitement des régularisations
                                et des acomptes (IRC uniquement) 1 occcurence
                                pour les régularisations et 1 occurence pour
                                les acomptes
                                Edition sur le fichier de log du montant déclaré,
                                de la régularisation et de l'acompte.
 PT9 : 12/01/2003 : V591  MF
                             1- Correction d'un access Vio quand les champs
                                ET_ADRESS1 et ET_ADRESSE3 étaient rensignés
                                et ET_ADRESSE2 non renseigné (table ETABLIS)
                                Erreur sur le nom du champ utilisé pour
                                l'adresse 3.
                             2- Suite tests validation UNEDIC
                                Correction formatage des périodes mensuelles
                                et de le clé IBAN quan unedic (tjs 99)
 PT10 : 15/01/2003 : V591 MF
                             1- Le segment MOATOT n'est présent pour l'ACOOS
                                que s'il s'agit d'une ducs dossier avec paiement
                                groupé
 PT11 : 31/01/2003 : V591 MF
                             1- Correction anomalie suite portage CWAS
                                Le qualifiant de cotisation n'était pas repris
                                (IRC)
 PT12  : 04/02/2003 : V591 MF
                             1- cas où pour un établissement on a deux
                                déclarations IRC. Un seul fichier
                                Duxyyyyyyyyyyyy.DUC était créé.
                                Pour les caisses IRC le nom du fichier
                                est maintenant bâti comme suit :
                                DUI+SIRET+ABREGE+CODE INSTITUTION
 PT13  : 10/02/2003 : V591 MF
                             1- correction alimentation de la date de règlement
                             des cotisations (prenait à tort la date de paiement
                             des salaires)
 PT14  : 28/02/2003 : V42  MF
                             1- Correction recherche des qualifiants de
                             cotisation
                             2- Rectification traitement des TOB (suite MEMCHECK)
 PT15  : 05/03/2003 : V42  MF (corrections suit tests GBdF pour IRC)
                             1- segment GIS
                             2- segment RFFMR groupe 1
                             3- segment RFFNUM (adhérent) groupe 1
                             4- segment RFFACT (adhérent) groupe 1
                             5- segments Tiers rédacteur groupe 1
                             6- segment MOADUE groupe 2
                             7- segment RFFDC groupe 2 (centre payeur)
                             8- segment MOATOT fin de déclaration
                             9- segment MOAPAI (montant réglé)
 PT16  : 07/03/2003 : V42  MF (corrections suit tests GBdF )
                             1- URSSAF : cotisation transport : une ligne de
                             codification 900 était créée à tort quand pls taux
                             transport.
                             2- IRC : segment DLI gp3 (n° de ligne): Il ne doit
                             pas y avoi de rupture de numérotation entre les
                             centres payeurs.
                             3- Taux de cotisation exprimé avec 4 décimales
                             sans point décimal.
                             4- IRC : Segment CNT
                             5- IRC : Segment QTY+913 (QTEF) présent uniquement
                             pour les forfaits ( M ou G) . Avec 4 décimales.
 PT17 : 11/03/2003 : V42  MF
                             1- UNEDIC : Calcul effectifs Hommmes et Femmes par
                             établissement.
                             2 segments QTY (920 = Hommes, 9210 = Femmes)
PT18  : 17/04/2003 VG V_42  Critère de sélection d'une clause WHERE sans libellé
PT19  : 06/05/2003 : V_42 MF Suite test de validation IRC
                             Les Montant du groupe 3 sont arrondis, exprimés en
                             centimes (forfait non arrondis).
PT20  : 21/07/2003 : V_421 MF Traitement du Centre Payeur (caisses IRC)
PT21  : 23/09/2003 MF V_421
                             1-FQ 10829 : pour les envois DUCS sélection possible
                             de l'émetteur.
                             2- Mise au point CWAS
PT22  : 24/02/2004  : MF V_5.0
                             UNEDIC : prise en compte de la date de changement
                             de taux sur segment DTM du groupe 3. On doit
                             fournir début et fin de période
PT23  : 08/04/2004  : MF V_5.0
                             1- Ajout contrôle n° interne obligatoire pour Acoss
                             et Unedic
                             2- FQ 11250 : Ajout contrôle pour Acoss pas de régul
                             s'il ne s'agit pas d'un TR Annuel
PT24  : 20/04/2004  : MF V_5.0
                             Si le répertoire de stockage n'existe pas le
                             traitement est abandonné  (CheminEagl)
PT25  : 19/01/2005  : MF V_6.0
                             FQ 11752 : Contrôle de la présence de la date
                             de règlement dans le cas d'une ducs IRC.
                             (règle le problème de ducs en partie doublée)
PT26  : 17/02/2005  : MF V_60
                             FQ 12013 : Correction traitement de récupération du
                             RIB de la banque
PT27  : 27/05/2005  : MF v_602
                              FQ 12329 : correction DTM+273 (date de référence)
                              et RFF+CR (Référence paiement)
PT28  : 25/07/2005  : MF V_604
                              FQ 12220  : correction du contrôle du qualifiant
                              de cotisation pour les ducs edi IRC.
PT29  : 15/12/2005  : MF V_650
                              FQ  12761 : Correction Référence du paiement
                              quand ASSEDIC paiement par chèque.
PT30  : 09/02/2006 : MF V_650 DUCS EDI V4.2
PT31  : 28/03/2006 : MF V_65  DUCS EDI V4.2 : mise au point suite tests de labellisation UNEDIC
                              Segment : RFF de type " ZE1/…/ZE8 " Ces segments
                              ne doivent pas être transmis pour les bordereaux
                              " 920 " et " 921 ".
PT32  : 07/04/2006 : MF V_65  DUCS  EDI V4.2 : mise au point suite tests de labellisation ACOSS
                              les segments MOA+134 et MOA+277 doivent être présents, même pour
                              une ducs néant.
PT33  : 16/06/2006 : MF V_70  DUCS EDI V4.1 : memcheck

PT34  : 07/07/2006 : MF V_70   DUCS EDI V4.2 : le calcul des effectifs UNEDIC est
                              revu pour satisfaire aux évolutions DRA 2005
                              + correction memcheck
                              + modification nom du logiciel (CEGID PAIE remplace CEGID PGI)
PT35  :  27/07/2006 : MF V_70  DUCS EDI V4.2 : suite test de labellisation UNEDIC
                              - correction de la référence sociale du paiement RFF+PQ
                              période de référence en cas paie décalée.
                              - Ajout segment RFF+AFN dans groupe 2 pour établissement
                              secondaire. Il faut enseigner le n° de compte de recouvrement
PT36  :  10/08/2006 : MF_V70   DUCS EDI V4.2 :
                              - Contrôle de l'identifiant de l'OPS obligatoire pour IRC
                              - Info sur le correspondant detinataire (IRC & ACOSS) facultative
                              - Mode de communication avec le contact destinataire (Tel) (IRC & ACOSS) facultatif
                              - Mode de communication avec le contact destinataire (Fax) (IRC & ACOSS) facultatif
                              - pas d'Email destinataire
PT37  :  29/01/2007 MF V_70  Modifs DUCS EDI V4.2
PT38  :  14/03/2007 VG V_72 BQ_GENERAL n'est pas forcément unique
PT39  :  11/04/2007 MF V_72 Modifs Mise en base des fichiers Ducs EDI
PT40  :  18/04/2007 MF V_72 FQ 13969 : modif traitement des effectifs.
PT41  :  19/04/2007 MF V_72 DUCS V4.2 : Correction alimentation date de paiement
PT42  :  18/07/2007 MF V_72 FQ 14588 : on encadre le nom du fichier (Chemin + Fichier) par des guillemets
PT43  :  01/08/2007 MF V_72 FQ 14637
PT44  :  02/08/2007 MF V_42 FQ 14638
PT45  :  07/08/2007 MF V_42 On ne génère plus le segemnt 'AUT' (cf cahier des charges : segment non utilisé)
PT46  :  14/09/2007 MF V_80 FQ 14777 : suite pilotage COGEPARC . On passe le
                            n° de dossier en paramètre à la fonction AGL__YFILESTD_IMPORT
PT47  :  27/09/2007 MF  Suite pilotage V8 : Pour les IRC le segment COM EM n'est pas attendu
                                            + contrôle pésence d'un titulaire du cpte (paiement)
PT48  :  10/10/2007 MF V_80 suite 1er envois V7 et V8 ducs V4.2 : l'identifiant SIREN n'est plus utilisé FQ 14844
PT49  :  10/10/2007 MF V_80 suite 1er envois V7 et V8 ducs V4.2 : correction de la référence de paiement chèque ou virement
                                                pour les déclaration trimetrielle et annuelles FQ 14839
PT50  :  10/10/2007 MF V_80 mise en place de la loi TEPA
PT51  :  25/10/2007 MF V_80 FQ 14897 : correction DTM+173 : date de déclaration
PT52  :  25/10/2007 MF V_80 Ne pas rendre obligatoire la coche paiement groupé pour les IRC
}
unit PGDucsEdiOutils;

interface
uses
  {$IFDEF VER150}
  Variants,
  {$ENDIF}
{$IFDEF EAGLCLIENT}
//unused  UtileAGL,
{$ELSE}
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}

  controls,HEnt1,hmsgbox,ShellAPI,Windows,SysUtils,HCtrls, PgOutils,PgOutils2,EntPaie,
  uYFILESTD,  // PT39
  UTOB,  ParamSoc, PgOutilsTreso;

Type TDucsEdi = record
     {Infos Liste du multi-critère}
     Etab : string;
     Organisme : string;
     DateDebut : TDateTime;
     DateFin : TDateTime;
     NoDucs : Integer;
     Siret : string;
     NatureOrg : string;
     TypDucs : string;
     Periode : string;
     {Infos DUCSENTETE}
     Abrege : string;
     IdentQual : string;
     IdentEmet : string;
     IdentDest : string;
     DucsDossier : Boolean;
     Declarant : string;
     EmettSoc : string;
     TelephoneDecl : string;
     FaxDeclarant : string;
     EmailDeclarant : string;      //  PT30 DUCS EDI V4.2
     DeclarantSuite : string;
     ApePdu : string;
     NbSalFpe : double;
     TotHommes : double;
     TotFemmes : double;
     TotApprenti : double;
//d PT30 DUCS EDI V4.2
// d PT34
{
     HorsAppH    : double;
     HorsAppF    : double;}
     CDIH        : double;
     CDIF        : double;
// f PT34
     ApprentiH   : double;
     ApprentiF   : double;
     CadreH      : double;
     CadreF      : double;
     E65H        : double;
     E65F        : double;
     ProfH       : double; //PT37
     ProfF       : double; //PT37
// d PT34
{
     PermanH     : double;
     PermanF     : double;}
     InterimH    : double;
     InterimF    : double;
{
     TEEPermH    : double;
     TEEPermF    : double;}
     SimplifH    : double;
     SimplifF    : double;
{
     TEEOccasH   : double;
     TEEOccasF   : double;}
     CNEH        : double;
     CNEF        : double;
{
     CEAH        : double;
     CEAF        : double;
     TTSH        : double;
     TTSF        : double;}
     CDDH        : double;
     CDDF        : double;
// f PT34
     EcartZe1    : boolean;
     EcartZe2    : boolean;
     EcartZe3    : boolean;
     EcartZe4    : boolean;
     EcartZe5    : boolean;
     EcartZe6    : boolean;
     EcartZe7    : boolean;
     EcartZe8    : boolean;
// d PT34
     EcartZe9    : boolean;
// f PT34
     MasseAnnuel : double;
     RemunDads   : double;
//f PT30 DUCS EDI V4.2
     Paiement : TdateTime;
     Reglement : TdateTime;
     Regularisation : double;
     Acompte : double;
     MtTotRegul : double;
     MtTotAcpte : double;
     NumeroPdu : string;
     MonnaieTenue : string;
     NumeroInterne : string;
     CentrePayeur : string;
     TypBordereau : string; //PT30 DUCS EDI V4.2
     { Infos DUCSDETAIL}
     MtDeclare  : double;
     MtTransport : double;
     MtTotal : double;
     MtAPayer : double;
     NbTransport : integer;
     NbCot : integer;
     CleIBAN : string;
     ErrTransport : boolean;
     ErrCotisqual : boolean;
     ErrInstitution : boolean;
     ErrCondition : boolean;
     Neant : boolean;
// d PT30 DUCS EDI V4.2
     NbTAT : integer;
     MtTAT : double;
     ErrTAT: boolean;
// f PT30 DUCS EDI V4.2
     { Infos ORGANISMEPAIE}
     SiretPog : string;
     LibellePog : string;
     Adresse1Pog : string;
     Adresse2Pog : string;
     VillePog : string;
     CpPog : string;
     ContactPog : string;
     TelPog : string;
     FaxPog : string;
     EmailPog : string;  // PT30 DUCS EDI V4.2
     InstitutionPog : string;
     AdherContact: string;
     NoContEmet : string;
     PaiementGroupe : boolean;
     ModePaiement : string;
     IdentOPS : string;
//PT30 DUCS EDI V4.2     CodAppliPog : string;
     ServUniqPog : boolean;
     { Infos EMETTEURSOCIAL}
     SiretPet : string;
     AdressePet : string;
     VillePet : string;
     CpPet : string;
     Adresse2Pet : string;
     { Infos ETABLISS}
     LibelleET : string;
     Adresse1ET : string;
     Adresse2ET : string;
     VilleET : string;
     CpET : string;
     TelET : string;
     FaxET : string;
     EmailET : string;   // PT30 DUCS EDI V4.2
     Juridique : string;
     NoCentrePayeur : integer;
     TitulaireCpt : string;  // PT30 DUCS EDI V4.2
end;

Type TDucsEdiGroupee = record
     Etab : string;
     Siret : string;
     LibelleET : string;
     Adresse1ET : string;
     Adresse2ET : string;
     VilleET : string;
     CpET : string;
     MtDeclare  : double;
     MtTransport : double;
     NbTransport : integer;
     Regularisation : double;
     Acompte : double;
     MtAPayer : double;
     ErrTransport : boolean;
     ErrCotisqual : boolean;
     ErrInstitution : boolean;
     ErrCondition : boolean;
     ApePdu : string;
     NbSalFpe : double;
     TotHommes : double;
     TotFemmes : double;
     TotApprenti : double;
//d PT30 DUCS EDI V4.2
     ErrTAT : boolean;
     MtTAT   : double;
     NbTAT : integer;
// d PT34
{
     HorsAppH    : double;
     HorsAppF    : double;}
     CDIH        : double;
     CDIF        : double;
// f PT34
     ApprentiH   : double;
     ApprentiF   : double;
     CadreH      : double;
     CadreF      : double;
     E65H        : double;
     E65F        : double;
     ProfH        : double;  //PT37
     ProfF        : double;  //PT37
// d PT34
{
     PermanH     : double;
     PermanF     : double;}
     InterimH    : double;
     InterimF    : double;
{
     TEEPermH    : double;
     TEEPermF    : double;}
     SimplifH    : double;
     SimplifF    : double;
{
     TEEOccasH   : double;
     TEEOccasF   : double;}
     CNEH        : double;
     CNEF        : double;
{
     CEAH        : double;
     CEAF        : double;
     TTSH        : double;
     TTSF        : double;}
     CDDH        : double;
     CDDF        : double;
// f PT34
     EcartZe1    : boolean;
     EcartZe2    : boolean;
     EcartZe3    : boolean;
     EcartZe4    : boolean;
     EcartZe5    : boolean;
     EcartZe6    : boolean;
     EcartZe7    : boolean;
     EcartZe8    : boolean;
// d PT34
     EcartZe9    : boolean;
// f PT34
     MasseAnnuel : double;
     RemunDads   : double;
//f PT30 DUCS EDI V4.2
     Abrege : string;
     MonnaieTenue : string;
     NumeroInterne : string;
     Paiement : TdateTime;
     Reglement : TdateTime;
     NoDucs : Integer;
     ModePaiement : string;
     CentrePayeur : string;
// d PT30 DUCS EDI V4.2
     TypBordereau : string;
     TitulaireCpt : string;
// f PT30 DUCS EDI V4.2
end;
         procedure OuvFicDucsEdiLOG(var Siret : string);
         procedure FermFicDucsEdiLOG();
         procedure EdiFicDucsEdiLOG(Tsiret : TOB);
         procedure OuvreFichierDucsEdi(var EnregDucsEdi : TDucsEdi);
         procedure InitVarDucsPDU(var EnregDucsEdi : TDucsEdi);
         procedure InitVarDucsPDD(var EnregDucsEdi : TDucsEdi);
         procedure InitVarDucsPOG(var EnregDucsEdi : TDucsEdi);
         procedure InitVarDucsPET(var EnregDucsEdi : TDucsEdi);
         procedure InitVarDucsET(var EnregDucsEdi : TDucsEdi);
         procedure InitEffectifsUNEDIC (var EnregDucsEdi : TDucsEdi);
         procedure InitCotisation(var EnregDucsEdi : TDucsEdi;var EnrDucsEdiGroupee : TDucsEdiGroupee;Trait : string);
         procedure EnTeteDeclaration(var EnregDucsEdi : TDucsEdi);
         procedure Groupe1Declaration(var EnregDucsEdi : TDucsEdi);
         procedure Groupe2Declaration (EnregDucsEdi : TDucsEdi; EnrDucsEdiGroupee : TDucsEdiGroupee; Trait : string);
         procedure Groupe3ACOSS(EnregDucsEdi : TDucsEdi; EnrDucsEdiGroupee : TDucsEdiGroupee; Trait : string);
         procedure Groupe4_5ACOSS(EnregDucsEdi : TDucsEdi;EnrDucsEdiGroupee : TDucsEdiGroupee;  Trait : string);
         procedure Groupe3IRC(var EnregDucsEdi : TDucsEdi; EnrDucsEdiGroupee : TDucsEdiGroupee; Trait : string; var NoLg : integer);
         procedure Groupe3UNEDIC(var EnregDucsEdi : TDucsEdi);
         procedure FinDeclaration(var EnregDucsEdi : TDucsEdi;EnrDucsEdiGroupee : TDucsEdiGroupee);
         procedure PaiementEdi(var EnregDucsEdi : TDucsEdi;EnrDucsEdiGroupee : TDucsEdiGroupee; Trait :string; var DossOK : boolean; TMtsEtabs : TOB);
         procedure MultiGp2Gp3(EtabPrinc : string;var EnregDucsEdi : TDucsEdi;var EnrDucsEdiGroupee : TDucsEdiGroupee; var DossOK : boolean; var NoLg : integer; var TMtsEtabs :TOB);//PT37
         procedure InitVarGp2(var EnregDucsEdi : TDucsEdi; var EnrDucsEdiGroupee : TDucsEdiGroupee;LesDucsGroupee : TOB; var TMtsEtabs : TOB);   //PT37
         procedure InitVarGp2UNEDIC(var EnregDucsEdi : TDucsEdi; var EnrDucsEdiGroupee : TDucsEdiGroupee;LesDucsGroupee : TOB);
         procedure FinFin (var EnregDucsEdi : TDucsEdi; var TMtsEtabs : TOB; sansFichierBase : boolean);  // PT39
         procedure MajEnvoiSocialDUCS (EnregDucsEdi : TDucsEdi; sansFichierBase : boolean); // PT39
// d PT30 DUCS EDI V4.2
         procedure MultiGp2Gp3TOB(EtabPrinc : string;var EnregDucsEdi : TDucsEdi;var EnrDucsEdiGroupee : TDucsEdiGroupee; var DossOK : boolean; var NoLg : integer);
         procedure InitEffGlobal(var EnregDucsEdi : TDucsEdi; var EnrDucsEdiGroupee : TDucsEdiGroupee;LesDucsGroupee : TOB);
         procedure DetUnedicGroupee(var EnregDucsEdi : TDucsEdi; var EnrDucsEdiGroupee : TDucsEdiGroupee);
// f PT30 DUCS EDI V4.2
         Function  ControlVarDucsEnTete (EnregDucsEdi : TDucsEdi) : Boolean;
         Function  ControlVarDucsGrp1 (EnregDucsEdi : TDucsEdi) : Boolean;
         Function  ControlCotisation(EnregDucsEdi : TDucsEdi;EnrDucsEdiGroupee : TDucsEdiGroupee;Trait : string) : Boolean;
         Function  ControlPaiement (EnregDucsEdi : TDucsEdi) : Boolean;
         Function  CleIBAN (creancier, periode,reference: string;var cleIBAN : string ) : boolean;
         procedure Abandonfichier(Trait : string);

var
         FRapport : TextFile;
         NomFic, NomFicDucs : string;
         FDUCSEDI : TextFile;
         TOB_Lignes, TOB_Transport, TOB_DucsGroupee : TOB;
         TOB_TAT     : TOB; // PT30 DUCS EDI V4.2
         PeriodRef : string;
implementation
uses Pg_OutilsEnvoi;
{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 12/04/2002
Modifié le ... : 22/05/2002
Description .. : Ouverture du fichier de log (en append ou cration)
Mots clefs ... : PAIE, PGDUCSEDI
*****************************************************************}
procedure OuvFicDucsEdiLOG(var Siret : String);       
var
  reponse                       : integer;
  Noerreur                      : integer;
begin
// d PT30 DUCS EDI V4.2
//{$IFDEF EAGLCLIENT}
  NomFic := VH_Paie.PGCheminEagl+'\DUCSEDI'+Siret+'.log';
//{$ELSE}
//  NomFic := V_PGI.DatPath+'\DUCSEDI'+Siret+'.log';
//{$ENDIF}
// PT30 DUCS EDI V4.2
  if FileExists(NomFic) then
  begin
    reponse := PGIAsk ('Voulez-vous réinitialiser le fichier de contrôle '+NomFic+' ?#13#10'+
                       'Si vous répondez ''NON'', le fichier existant sera complété.',
                       'Confection de fichiers Ducs-EDI');
    if reponse = mrYes then
      DeleteFile(PChar(NomFic));
  end;

  AssignFile(FRapport, NomFic);
  if FileExists(NomFic) then
    Append(FRapport)
  else
  begin
  {$i-}ReWrite(FRapport); {$i+}
    NoErreur := IoResult;
    if NoErreur<>0 Then
      Begin
       PGIBox ('Le répertoire de stockage n''est pas défini', 'Traitement impossible');
       Siret := '';
       Exit ;
      End;
    Writeln(FRapport, '');
    Writeln(FRapport, '*** CONFECTION DE FICHIERS DUCS EDI ***');
    Writeln(FRapport, '---------------------------------------');
  end;

  Writeln(FRapport, '');
  Writeln(FRapport, '********** Début de traitement : '+DateTimeToStr(Now)+
                    '  ******************************************');
end;
{ fin OuvFicDucsEdiLOG}


{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 12/04/2002
Modifié le ... : 22/05/2002
Description .. : fermeture du fichier de log
Mots clefs ... : PAIE, PGDUCSEDI
*****************************************************************}
procedure FermFicDucsEdiLOG();
begin
  Writeln(FRapport, '');
  Writeln(FRapport, '********** Traitement terminé : '+DateTimeToStr(Now)+
                    '  ******************************************');

  CloseFile(FRapport);
end;
{ fin FermFicDucsEdiLOG}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 12/04/2002
Modifié le ... : 22/05/2002
Description .. : lancement de l'édition du fichier de log
Mots clefs ... : PAIE, PGDUCSEDI
*****************************************************************}
procedure EdiFicDucsEdiLOG(TSiret : TOB);
var
  i, II                         : integer;
  Siret, Anomalie               : string;
  TSiretFille                   : TOB;
  NomFic                        : string;

begin
  for II := 0 to TSiret.Detail.Count-1 do
  begin
    TSiretFille := TSiret.Detail[II];
    Siret := TSiretFille.GetValue('SIRET');
    Anomalie := TSiretFille.GetValue('ANOMALIE') ;
// d PT30 DUCS EDI V4.2
//{$IFDEF EAGLCLIENT}
    NomFic := '"'+VH_Paie.PGCheminEagl+'\DUCSEDI'+Siret+'.log"';  // PT42
//{$ELSE}
//    NomFic := V_PGI.DatPath+'\DUCSEDI'+Siret+'.log';
//{$ENDIF}
// f PT30 DUCS EDI V4.2
    i:=PGIAsk ('Siret : '+Siret+'#13#10 Voulez-vous visualiser le fichier de contrôle ?', 'DUCS EDI') ;
    if i=mrYes then
       ShellExecute( 0, PCHAR('open'),PChar('WordPad'), PChar(NomFic),Nil,SW_RESTORE);
  end;
end;
{ fin EdiFicDucsEdiLOG}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 12/04/2002
Modifié le ... :   /  /
Description .. : Initialisation des variables nécessaires à la confection du
Suite ........ : fichier Ducs Edi, issues de la table DUCSENTETE
Mots clefs ... : PAIE, PGDUCSEDI
*****************************************************************}
procedure InitVarDucsPDU(var EnregDucsEdi : TDucsEdi);
var
  QQ                            : TQuery;
begin
  QQ := OpenSQL('SELECT PDU_ABREGEPERIODE,PDU_DUCSDOSSIER,PDU_DECLARANT,'+
                'PDU_EMETTSOC,PDU_TELEPHONEDECL,PDU_FAXDECLARANT,PDU_NUMERO,'+
                'PDU_DECLARANTSUITE,PDU_APE,PDU_NBSALFPE,PDU_TOTHOMMES,'+
                'PDU_TOTFEMMES,PDU_TOTAPPRENTI,PDU_PAIEMENT,PDU_DATEREGLEMENT,PDU_ACOMPTES,'+
                'PDU_REGULARISATION,PDU_NUMERO,PDU_MONNAIETENUE, '+
                'PDU_CENTREPAYEUR '+
// d PT30 DUCS EDI V4.2
                ',PDU_TYPBORDEREAU '+
                ',PDU_NBSALQ922,PDU_NBSALQ923,PDU_NBSALQ924,PDU_NBSALQ925'+
                ',PDU_NBSALQ926,PDU_NBSALQ927,PDU_NBSALQ928,PDU_NBSALQ929 '+
                ',PDU_NBSALQ934,PDU_NBSALQ935,PDU_NBSALQ944,PDU_NBSALQ945 '+
                ',PDU_NBSALQ936,PDU_NBSALQ937'+
                ',PDU_NBSALQ960,PDU_NBSALQ961,PDU_NBSALQ962,PDU_NBSALQ963 '+
                ',PDU_NBSALQ964,PDU_NBSALQ965,PDU_NBSALQ966,PDU_NBSALQ967 '+
                ',PDU_ECARTZE1,PDU_ECARTZE2,PDU_ECARTZE3,PDU_ECARTZE4 '+
                ',PDU_ECARTZE5,PDU_ECARTZE6,PDU_ECARTZE7,PDU_ECARTZE8 '+
                ',PDU_MASSEANNUEL,PDU_REMUNDADS '+
// f PT30 DUCS EDI V4.2
                'FROM DUCSENTETE WHERE '+
                'PDU_ETABLISSEMENT = "'+EnregDucsEdi.Etab+'" AND '+
                'PDU_ORGANISME = "'+EnregDucsEdi.Organisme+'" AND '+
                'PDU_DATEDEBUT = "'+UsDateTime(EnregDucsEdi.DateDebut)+'" AND '+
                'PDU_DATEFIN = "'+UsDateTime(EnregDucsEdi.DateFin)+'" AND '+
                'PDU_NUM = '+IntToStr(EnregDucsEdi.NoDucs)+'',True) ;        // DB2
  if Not QQ.EOF then
  begin
    EnregDucsEdi.Abrege:=QQ.FindField('PDU_ABREGEPERIODE').AsString;
    EnregDucsEdi.DucsDossier:=QQ.FindField('PDU_DUCSDOSSIER').AsString='X';
    EnregDucsEdi.Declarant:=QQ.FindField('PDU_DECLARANT').AsString;
    EnregDucsEdi.EmettSoc:=QQ.FindField('PDU_EMETTSOC').AsString;
    EnregDucsEdi.TelephoneDecl:=QQ.FindField('PDU_TELEPHONEDECL').AsString;
    EnregDucsEdi.FaxDeclarant:=QQ.FindField('PDU_FAXDECLARANT').AsString;
    EnregDucsEdi.NumeroInterne:=QQ.FindField('PDU_NUMERO').AsString;
    EnregDucsEdi.DeclarantSuite:=QQ.FindField('PDU_DECLARANTSUITE').AsString;
    EnregDucsEdi.ApePdu:=QQ.FindField('PDU_APE').AsString;
    EnregDucsEdi.NbSalFpe:=QQ.FindField('PDU_NBSALFPE').AsFloat;
    EnregDucsEdi.TotHommes:=QQ.FindField('PDU_TOTHOMMES').AsFloat;
    EnregDucsEdi.TotFemmes:=QQ.FindField('PDU_TOTFEMMES').AsFloat;
    EnregDucsEdi.TotApprenti:=QQ.FindField('PDU_TOTAPPRENTI').AsFloat;
// d PT30 DUCS EDI V4.2
// d PT34
{
    EnregDucsEdi.HorsAppH:=QQ.FindField('PDU_NBSALQ922').AsFloat;
    EnregDucsEdi.HorsAppF:=QQ.FindField('PDU_NBSALQ923').AsFloat;}
    EnregDucsEdi.CDIH:=QQ.FindField('PDU_NBSALQ922').AsFloat;
    EnregDucsEdi.CDIF:=QQ.FindField('PDU_NBSALQ923').AsFloat;
// d PT34
    EnregDucsEdi.ApprentiH:=QQ.FindField('PDU_NBSALQ924').AsFloat;
    EnregDucsEdi.ApprentiF:=QQ.FindField('PDU_NBSALQ925').AsFloat;
    EnregDucsEdi.CadreH:=QQ.FindField('PDU_NBSALQ928').AsFloat;
    EnregDucsEdi.CadreF:=QQ.FindField('PDU_NBSALQ929').AsFloat;
    EnregDucsEdi.E65H:=QQ.FindField('PDU_NBSALQ926').AsFloat;
    EnregDucsEdi.E65F:=QQ.FindField('PDU_NBSALQ927').AsFloat;
    EnregDucsEdi.ProfH:=QQ.FindField('PDU_NBSALQ934').AsFloat;    // PT37
    EnregDucsEdi.ProfF:=QQ.FindField('PDU_NBSALQ935').AsFloat;    // PT37
// d PT34
{
    EnregDucsEdi.PermanH:=QQ.FindField('PDU_NBSALQ944').AsFloat;
    EnregDucsEdi.PermanF:=QQ.FindField('PDU_NBSALQ945').AsFloat;}
    EnregDucsEdi.InterimH:=QQ.FindField('PDU_NBSALQ944').AsFloat;
    EnregDucsEdi.InterimF:=QQ.FindField('PDU_NBSALQ945').AsFloat;
{
    EnregDucsEdi.TEEPermH:=QQ.FindField('PDU_NBSALQ960').AsFloat;
    EnregDucsEdi.TEEPermF:=QQ.FindField('PDU_NBSALQ961').AsFloat;}
    EnregDucsEdi.SimplifH:=QQ.FindField('PDU_NBSALQ960').AsFloat;
    EnregDucsEdi.SimplifF:=QQ.FindField('PDU_NBSALQ961').AsFloat;
{    EnregDucsEdi.TEEOccasH:=QQ.FindField('PDU_NBSALQ962').AsFloat;
    EnregDucsEdi.TEEOccasF:=QQ.FindField('PDU_NBSALQ963').AsFloat;}
    EnregDucsEdi.CNEH:=QQ.FindField('PDU_NBSALQ962').AsFloat;
    EnregDucsEdi.CNEF:=QQ.FindField('PDU_NBSALQ963').AsFloat;
{
    EnregDucsEdi.CEAH:=QQ.FindField('PDU_NBSALQ964').AsFloat;
    EnregDucsEdi.CEAF:=QQ.FindField('PDU_NBSALQ965').AsFloat;
    EnregDucsEdi.TTSH:=QQ.FindField('PDU_NBSALQ966').AsFloat;
    EnregDucsEdi.TTSF:=QQ.FindField('PDU_NBSALQ967').AsFloat;}
    EnregDucsEdi.CDDH:=QQ.FindField('PDU_NBSALQ936').AsFloat;
    EnregDucsEdi.CDDF:=QQ.FindField('PDU_NBSALQ937').AsFloat;
// f PT34
    EnregDucsEdi.EcartZe1:=QQ.FindField('PDU_ECARTZE1').AsString='X';
    EnregDucsEdi.EcartZe2:=QQ.FindField('PDU_ECARTZE2').AsString='X';
    EnregDucsEdi.EcartZe3:=QQ.FindField('PDU_ECARTZE3').AsString='X';
    EnregDucsEdi.EcartZe4:=QQ.FindField('PDU_ECARTZE4').AsString='X';
    EnregDucsEdi.EcartZe5:=QQ.FindField('PDU_ECARTZE5').AsString='X';
    EnregDucsEdi.EcartZe6:=QQ.FindField('PDU_ECARTZE6').AsString='X';
    EnregDucsEdi.EcartZe7:=QQ.FindField('PDU_ECARTZE7').AsString='X';
    EnregDucsEdi.EcartZe8:=QQ.FindField('PDU_ECARTZE8').AsString='X';
// d PT34
    if  (QQ.FindField('PDU_NBSALQ966').AsFloat = 1) then
        EnregDucsEdi.EcartZe9 := TRUE
    else
        EnregDucsEdi.EcartZe9 := FALSE;
// f PT34
    EnregDucsEdi.MasseAnnuel:=QQ.FindField('PDU_MASSEANNUEL').AsFloat;
    EnregDucsEdi.RemunDads:=QQ.FindField('PDU_REMUNDADS').AsFloat;
// f PT30 DUCS EDI V4.2
    EnregDucsEdi.Paiement:=QQ.FindField('PDU_PAIEMENT').AsDateTime;
    EnregDucsEdi.Reglement:=QQ.FindField('PDU_DATEREGLEMENT').AsDateTime;
    EnregDucsEdi.Regularisation:=QQ.FindField('PDU_REGULARISATION').AsFloat;
    EnregDucsEdi.Acompte:=QQ.FindField('PDU_ACOMPTES').AsFloat;
    EnregDucsEdi.NumeroPdu:=QQ.FindField('PDU_NUMERO').AsString;
    EnregDucsEdi.MonnaieTenue:=QQ.FindField('PDU_MONNAIETENUE').AsString;
    if (QQ.FindField('PDU_CENTREPAYEUR').AsString <> '') then
      EnregDucsEdi.CentrePayeur:=QQ.FindField('PDU_CENTREPAYEUR').AsString;
// d PT30 DUCS EDI V4.2
    EnregDucsEdi.TypBordereau:=QQ.FindField('PDU_TYPBORDEREAU').AsString;
// f PT30 DUCS EDI V4.2
  end ;
  Ferme(QQ) ;
end;
{ fin InitVarDucsPDU}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 13/05/2002
Modifié le ... :   /  /
Description .. : Initialisation des variables nécessaires à la confection
Suite ........ : du fichier Ducs Edi issues de la table DUCSDETAIL
Mots clefs ... : PAIE, PGDUCSEDI
*****************************************************************}
procedure InitVarDucsPDD(var EnregDucsEdi : TDucsEdi);
var
  QQ : TQuery;
  StQQ : string;

begin
  StQQ := 'PDD_ETABLISSEMENT = "'+EnregDucsEdi.Etab+'" AND '+
          'PDD_ORGANISME = "'+EnregDucsEdi.Organisme+'" AND '+
          'PDD_DATEDEBUT = "'+UsDateTime(EnregDucsEdi.DateDebut)+'" AND '+
          'PDD_DATEFIN = "'+UsDateTime(EnregDucsEdi.DateFin)+'" AND '+
          'PDD_NUM = '+IntToStr(EnregDucsEdi.NoDucs)+' AND '+          // DB2
          'PDD_TYPECOTISATION <> "S"';
  QQ := OpenSQL('SELECT SUM(PDD_MTCOTISAT) AS DECLAR FROM DUCSDETAIL WHERE '+StQQ,True) ;
  EnregDucsEdi.MtDeclare := 0.0;
  EnregDucsEdi.MtTotal := 0.0;
  EnregDucsEdi.MtAPayer := 0.0;

  if not QQ.EOF then // PortageCWAS
    EnregDucsEdi.MtDeclare := QQ.FindField('DECLAR').asFloat;
  EnregDucsEdi.MtTotal := EnregDucsEdi.MtTotal + EnregDucsEdi.MtDeclare;
{ PT8-2 }
  EnregDucsEdi.MtAPayer := EnregDucsEdi.MtAPayer+
                           EnregDucsEdi.MtDeclare+
                           EnregDucsEdi.Regularisation-
                           EnregDucsEdi.Acompte;
  EnregDucsEdi.MtTotRegul :=0.0;
  EnregDucsEdi.MtTotAcpte :=0.0;

  EnregDucsEdi.MtTotRegul :=  EnregDucsEdi.MtTotRegul + EnregDucsEdi.Regularisation;
  EnregDucsEdi.MtTotAcpte :=  EnregDucsEdi.MtTotAcpte + EnregDucsEdi.Acompte;

  Ferme (QQ);
// d PT30 DUCS EDI V4.2
  if (Copy(EnregDucsEdi.TypDucs,1,1) = 'A') then
  begin
// f PT30 DUCS EDI V4.2
    // Versement Transport
    StQQ := '';
    StQQ := 'PDD_ETABLISSEMENT = "'+EnregDucsEdi.Etab+'" AND '+
            'PDD_ORGANISME = "'+EnregDucsEdi.Organisme+'" AND '+
            'PDD_DATEDEBUT = "'+UsDateTime(EnregDucsEdi.DateDebut)+'" AND '+
            'PDD_DATEFIN = "'+UsDateTime(EnregDucsEdi.DateFin)+'" AND '+
            'PDD_NUM = '+IntToStr(EnregDucsEdi.NoDucs)+' AND '+  // DB2
            'PDD_TYPECOTISATION <> "S" AND PDD_TYPECOTISATION <> "I" '+
            'AND PDD_CODIFICATION LIKE "1%900P"';
    QQ := OpenSQL('SELECT COUNT(DISTINCT(PDD_TAUXCOTISATION)) AS NBTRANSPORT, '+
                  'SUM(PDD_MTCOTISAT) AS MTTRANSPORT FROM DUCSDETAIL WHERE '+StQQ,True) ;
    EnregDucsEdi.MtTransport:=0.0;
    if not QQ.EOF then // PortageCWAS
    begin
      EnregDucsEdi.MtTransport:= QQ.FindField('MTTRANSPORT').asFloat;
      EnregDucsEdi.NbTransport:= QQ.FindField('NBTRANSPORT').AsInteger;
    end;
    Ferme(QQ);
    EnregDucsEdi.ErrTransport := False;
// d PT30 DUCS EDI V4.2
//  if (copy(EnregDucsEdi.TypDucs,1,1) = 'A') and
//    (EnregDucsEdi.NbTransport > 1) then
    if (EnregDucsEdi.NbTransport > 1) then
// f PT30 DUCS EDI V4.2
    {Création TOB_Transport}
    begin
// d PT30 DUCS EDI V4.2
//    QQ := OpenSQL('SELECT Count (*) FROM DUCSDETAIL WHERE (PDD_COMURBAINE="" '+
//                    'OR PDD_COMURBAINE IS NULL OR PDD_COMURBAINE="    ") and '+
//                    StQQ,TRUE);
      QQ := OpenSQL('SELECT Count (*) FROM DUCSDETAIL WHERE (PDD_CODECOMMUNE="" '+
                    'OR PDD_CODECOMMUNE IS NULL OR PDD_CODECOMMUNE="     ") and '+
                    StQQ,TRUE);
// f PT30 DUCS EDI V4.2
      if not QQ.EOF then
      begin
        if (QQ.Fields[0].AsInteger <> 0) then
          EnregDucsEdi.ErrTransport :=  True;
      end;
      ferme (QQ);

      QQ := OpenSQL('SELECT * FROM DUCSDETAIL WHERE '+StQQ,True);

      TOB_Transport := TOB.Create ('Les lignes de tansport', NIL, -1);
      TOB_Transport.LoadDetailDB ('DUCSDETAIL', '','',QQ,FALSE,FALSE);
      Ferme(QQ);
      TOB_Transport.Detail.Sort('PDD_ETABLISSEMENT;PDD_ORGANISME;PDD_DATEDEBUT;'+
                                 'PDD_DATEFIN;PDD_INSTITUTION;PDD_CODIFICATION;PDD_DATECHGTTAUX');
    end;

    // Taxes additionnelles de transpaort
    StQQ := '';
    StQQ := 'PDD_ETABLISSEMENT = "'+EnregDucsEdi.Etab+'" AND '+
             'PDD_ORGANISME = "'+EnregDucsEdi.Organisme+'" AND '+
             'PDD_DATEDEBUT = "'+UsDateTime(EnregDucsEdi.DateDebut)+'" AND '+
             'PDD_DATEFIN = "'+UsDateTime(EnregDucsEdi.DateFin)+'" AND '+
             'PDD_NUM = '+IntToStr(EnregDucsEdi.NoDucs)+' AND '+  // DB2
             'PDD_TYPECOTISATION <> "S" AND PDD_TYPECOTISATION <> "I" '+
             'AND PDD_CODIFICATION LIKE "1%901P"';
     QQ := OpenSQL('SELECT COUNT(DISTINCT(PDD_TAUXCOTISATION)) AS NBTAT, '+
                   'SUM(PDD_MTCOTISAT) AS MTTAT FROM DUCSDETAIL WHERE '+StQQ,True) ;
     EnregDucsEdi.MtTAT:=0.0;
     if not QQ.EOF then
     begin
       EnregDucsEdi.MtTAT:= QQ.FindField('MTTAT').asFloat;
       EnregDucsEdi.NbTAT:= QQ.FindField('NBTAT').AsInteger;
     end;
     Ferme(QQ);
     EnregDucsEdi.ErrTAT := False;
     if (EnregDucsEdi.NbTAT > 1) then
    {Création TOB_TAT}
     begin
// d PT30 DUCS EDI V4.2
//     QQ := OpenSQL('SELECT Count (*) FROM DUCSDETAIL WHERE (PDD_COMURBAINE="" '+
//                     'OR PDD_COMURBAINE IS NULL OR PDD_COMURBAINE="    ") and '+
//                     StQQ,TRUE);
       QQ := OpenSQL('SELECT Count (*) FROM DUCSDETAIL WHERE (PDD_CODECOMMUNE="" '+
                     'OR PDD_CODECOMMUNE IS NULL OR PDD_CODECOMMUNE="     ") and '+
                     StQQ,TRUE);
// f PT30 DUCS EDI V4.2
       if not QQ.EOF then
       begin
         if (QQ.Fields[0].AsInteger <> 0) then
           EnregDucsEdi.ErrTAT :=  True;
       end;
       ferme (QQ);

       QQ := OpenSQL('SELECT * FROM DUCSDETAIL WHERE '+StQQ,True);

       TOB_TAT := TOB.Create ('Les lignes de taxes additionnelles de transport', NIL, -1);

       TOB_TAT.LoadDetailDB ('DUCSDETAIL', '','',QQ,FALSE,FALSE);
       Ferme(QQ);
       TOB_TAT.Detail.Sort('PDD_ETABLISSEMENT;PDD_ORGANISME;PDD_DATEDEBUT;'+
                           'PDD_DATEFIN;PDD_INSTITUTION;PDD_CODIFICATION;PDD_DATECHGTTAUX');

    end;
// f PT30 DUCS EDI V4.2
  end;
end;
{ fin InitVarDucsPDD}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 12/04/2002
Modifié le ... :   /  /
Description .. : Initialisation des variables nécessaires à la confection du
Suite ........ : fichier Ducs Edi, issues de la table ORGANISMEPAIE
Mots clefs ... : PAIE, PGDUCSEDI
*****************************************************************}
procedure InitVarDucsPOG(var EnregDucsEdi : TDucsEdi);
var
  QQ : TQuery;
begin

  QQ := OpenSQL('SELECT POG_NATUREORG,POG_IDENTQUAL,POG_IDENTEMET,'+
                'POG_IDENTDEST,POG_NOCONTEMET,POG_SIRET,POG_LIBELLE,'+
                'POG_ADRESSE1,POG_ADRESSE2, POG_ADRESSE3,POG_VILLE,'+
                'POG_CODEPOSTAL,POG_CONTACT,POG_TELEPHONE,POG_FAX,'+
                'POG_INSTITUTION,POG_ADHERCONTACT,POG_PAIEGROUPE, '+
//PT30 DUCS EDI V4.2
//               'POG_PAIEMODE,POG_IDENTOPS,POG_CODAPPLI,POG_SERVUNIQ '+
                'POG_PAIEMODE,POG_IDENTOPS,POG_SERVUNIQ,POG_EMAIL,POG_TITULAIRECPT '+
                'FROM ORGANISMEPAIE WHERE '+
                'POG_ETABLISSEMENT = "'+EnregDucsEdi.Etab+'" AND '+
                'POG_ORGANISME = "'+EnregDucsEdi.Organisme+'"',True) ;
  if not QQ.EOF then
  begin
    EnregDucsEdi.NatureOrg:=QQ.FindField('POG_NATUREORG').AsString;
    EnregDucsEdi.IdentQual:=QQ.FindField('POG_IDENTQUAL').AsString;
    EnregDucsEdi.IdentEmet:=QQ.FindField('POG_IDENTEMET').AsString;
    EnregDucsEdi.IdentDest:=QQ.FindField('POG_IDENTDEST').AsString;
    EnregDucsEdi.SiretPog:=QQ.FindField('POG_SIRET').AsString;
    EnregDucsEdi.LibellePog:=QQ.FindField('POG_LIBELLE').AsString;

    if (QQ.FindField('POG_ADRESSE1').AsString <> '') then
      EnregDucsEdi.Adresse1Pog:=QQ.FindField('POG_ADRESSE1').AsString;
    if (QQ.FindField('POG_ADRESSE2').AsString <> '') then
      if (EnregDucsEdi.Adresse1Pog  <> '') then
        EnregDucsEdi.Adresse2Pog:=QQ.FindField('POG_ADRESSE2').AsString
      else
        EnregDucsEdi.Adresse1Pog:=QQ.FindField('POG_ADRESSE2').AsString;
    if (QQ.FindField('POG_ADRESSE3').AsString <> '') then
      if (EnregDucsEdi.Adresse1Pog  = '') then
        EnregDucsEdi.Adresse1Pog:=QQ.FindField('POG_ADRESSE3').AsString
      else
        if (EnregDucsEdi.Adresse2Pog  = '') then
          EnregDucsEdi.Adresse2Pog:=QQ.FindField('POG_ADRESSE3').AsString;

    EnregDucsEdi.VillePog:=QQ.FindField('POG_VILLE').AsString;
    EnregDucsEdi.CpPog:=QQ.FindField('POG_CODEPOSTAL').AsString;
    EnregDucsEdi.ContactPog:=QQ.FindField('POG_CONTACT').AsString;
    EnregDucsEdi.TelPog:=QQ.FindField('POG_TELEPHONE').AsString;
    EnregDucsEdi.FaxPog:=QQ.FindField('POG_FAX').AsString;
    EnregDucsEdi.EmailPog:=QQ.FindField('POG_EMAIL').AsString;   // PT30 DUCS EDI V4.2
    EnregDucsEdi.InstitutionPog:=QQ.FindField('POG_INSTITUTION').AsString;
    EnregDucsEdi.AdherContact:=QQ.FindField('POG_ADHERCONTACT').AsString;
    EnregDucsEdi.NoContEmet:=QQ.FindField('POG_NOCONTEMET').AsString;
    EnregDucsEdi.PaiementGroupe:=QQ.FindField('POG_PAIEGROUPE').AsString='X';
    EnregDucsEdi.ModePaiement:=QQ.FindField('POG_PAIEMODE').AsString;
    EnregDucsEdi.IdentOPS:=QQ.FindField('POG_IDENTOPS').AsString;
    EnregDucsEdi.TitulaireCpt:=QQ.FindField('POG_TITULAIRECPT').AsString;

{PT30 DUCS EDI V4.2
pour ducs edi v4.1
   EnregDucsEdi.CodAppliPog:=QQ.FindField('POG_CODAPPLI').AsString;
    if (EnregDucsEdi.CodAppliPog = '') then
      EnregDucsEdi.CodAppliPog := '013';
    EnregDucsEdi.ServUniqPog:=QQ.FindField('POG_SERVUNIQ').AsString='X';*}

  end ;
  Ferme(QQ) ;
  if (EnregDucsEdi.NatureOrg = '100') then
    EnregDucsEdi.TypDucs := 'A ACOSS ';         { Acoss}
  if (EnregDucsEdi.NatureOrg = '200') then
    EnregDucsEdi.TypDucs := 'U UNEDIC';         {Unedic}
  if (EnregDucsEdi.NatureOrg = '300') then
    EnregDucsEdi.TypDucs := 'I AGIRC ';         { Irc}
end;
{ fin InitVarDucsPOG}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 15/04/2002
Modifié le ... :   /  /
Description .. : Initialisation des variables nécessaires à la confection du
Suite ........ : fichier Ducs Edi, issues de la table  EMETTEURSOCIAL
Mots clefs ... : PAIE, PGDUCSEDI
*****************************************************************}
procedure InitVarDucsPET(var EnregDucsEdi : TDucsEdi);
var
  QQ : TQuery;
  Adresse     : string;      //PT37
begin
  Adresse := '';

  QQ := OpenSQL('SELECT PET_SIRET,PET_COMPLADR,PET_VILLE,PET_CODEPOSTAL,'+
                'PET_COMPLADR2, '+
                'PET_NUMEROVOIE, PET_BISTER, PET_NOMVOIE, PET_BURDISTRIB '+     //PT37
                ',PET_EMAILDUCS '+        // PT30 DUCS EDI V4.2
                'FROM EMETTEURSOCIAL WHERE '+
                'PET_EMETTSOC = "'+EnregDucsEdi.EmettSoc+'" ',True) ;
  if not QQ.EOF then
  begin
    EnregDucsEdi.SiretPet:=QQ.FindField('PET_SIRET').AsString;

// d PT37
    if (QQ.FindField('PET_NUMEROVOIE').AsString <> '') then
      Adresse := QQ.FindField('PET_NUMEROVOIE').AsString + ' ';
    if (QQ.FindField('PET_BISTER').AsString <> '') then
      Adresse := ' '+ QQ.FindField('PET_BISTER').AsString + ' ';
    if (QQ.FindField('PET_NOMVOIE').AsString <> '') then
      Adresse := ' ' + QQ.FindField('PET_NOMVOIE').AsString + ' ';
    if (Adresse = '') then
    begin
      Adresse :=  QQ.FindField('PET_COMPLADR').AsString;
      EnregDucsEdi.Adresse2Pet:=QQ.FindField('PET_COMPLADR2').AsString;
    end
    else
    if (QQ.FindField('PET_COMPLADR').AsString <> '') then
      EnregDucsEdi.Adresse2Pet:=QQ.FindField('PET_COMPLADR').AsString
    else
      EnregDucsEdi.Adresse2Pet:=QQ.FindField('PET_COMPLADR2').AsString;


//    EnregDucsEdi.AdressePet:=QQ.FindField('PET_COMPLADR').AsString;
    EnregDucsEdi.AdressePet := Adresse;

    if (QQ.FindField('PET_VILLE').AsString <> '') then
       EnregDucsEdi.VillePet:=QQ.FindField('PET_VILLE').AsString
    else
       EnregDucsEdi.VillePet:=QQ.FindField('PET_BURDISTRIB').AsString;
// f PT37

    EnregDucsEdi.CpPet:=QQ.FindField('PET_CODEPOSTAL').AsString;
//PT37    EnregDucsEdi.Adresse2Pet:=QQ.FindField('PET_COMPLADR2').AsString;
    EnregDucsEdi.EmailDeclarant:=QQ.FindField('PET_EMAILDUCS').AsString;
  end ;
  Ferme(QQ) ;
end;
{ fin InitVarDucsPET}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 15/04/2002
Modifié le ... :   /  /
Description .. : Initialisation des variables nécessaires à la confection du
Suite ........ : fichier Ducs Edi, issues de la table  ETABLISS
Mots clefs ... : PAIE, PGDUCSEDI
*****************************************************************}
procedure InitVarDucsET(var EnregDucsEdi : TDucsEdi);
var
  QQ : TQuery;
begin

  QQ := OpenSQL('SELECT ET_LIBELLE,ET_ADRESSE1,ET_ADRESSE2,ET_ADRESSE3,'+
                'ET_VILLE,ET_CODEPOSTAL,ET_TELEPHONE,ET_FAX,ET_JURIDIQUE '+
                ',ET_EMAIL '+                  // PT30 DUCS EDI V4.2
                'FROM ETABLISS WHERE '+
                'ET_ETABLISSEMENT = "'+EnregDucsEdi.Etab+'" ',True) ;
  if not QQ.EOF then
  begin
    EnregDucsEdi.LibelleET:=QQ.FindField('ET_LIBELLE').AsString;

    if (QQ.FindField('ET_ADRESSE1').AsString <> '') then
      EnregDucsEdi.Adresse1ET:=QQ.FindField('ET_ADRESSE1').AsString;
    if (QQ.FindField('ET_ADRESSE2').AsString <> '') then
      if (EnregDucsEdi.Adresse1ET  <> '') then
        EnregDucsEdi.Adresse2ET:=QQ.FindField('ET_ADRESSE2').AsString
      else
        EnregDucsEdi.Adresse1ET:=QQ.FindField('ET_ADRESSE2').AsString;
    if (QQ.FindField('ET_ADRESSE3').AsString <> '') then
      if (EnregDucsEdi.Adresse1ET  = '') then
        EnregDucsEdi.Adresse1ET:=QQ.FindField('ET_ADRESSE3').AsString
      else
        if (EnregDucsEdi.Adresse2ET  = '') then
          EnregDucsEdi.Adresse2ET:=QQ.FindField('ET_ADRESSE3').AsString;

    EnregDucsEdi.VilleET:=QQ.FindField('ET_VILLE').AsString;
    EnregDucsEdi.CpET:=QQ.FindField('ET_CODEPOSTAL').AsString;
    EnregDucsEdi.TelET:=QQ.FindField('ET_TELEPHONE').AsString;
    EnregDucsEdi.FaxET:=QQ.FindField('ET_FAX').AsString;
    EnregDucsEdi.EmailET:=QQ.FindField('ET_EMAIL').AsString;     // PT30 DUCS EDI V4.2
    EnregDucsEdi.Juridique:=QQ.FindField('ET_JURIDIQUE').AsString;
  end ;
  Ferme(QQ) ;
end;
{ fin InitVarDucsET}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 28/05/2002
Modifié le ... :   /  /
Description .. : Calcul des effectifs pour l'établissement principal seulement .
Suite ........ : Il s'agit d'une déclaration UNEDIC avec VLU, lors de
Suite ........ : l'initialisation de la ducs les effectifs calculés correspondent
Suite ........ : aux effectifs tous établissements confondus.
Mots clefs ... : PAIE, PGDUCSEDI
*****************************************************************}
procedure InitEffectifsUNEDIC (var EnregDucsEdi : TDucsEdi);
var
  StQL                  : string;
  QL                    : Tquery;
// d PT30 DUCS EDI V4.2
  TEffectifs, LeSal     : TOB;
  DateNaiss             : TDateTime;
  I                     : integer;
// f PT30 DUCS EDI V4.2
begin
  {Calcul des Effectifs en fin de période}
  StQL := '';
  StQL :='PSA_ETABLISSEMENT="'+EnregDucsEdi.Etab+'" AND ';

  StQL := StQL+'(PSA_DATESORTIE="'+UsDateTime(IDate1900)+'"'+
               ' OR  PSA_DATESORTIE IS NULL '+
               ' OR PSA_DATESORTIE>="'+UsDateTime(EnregDucsEdi.DateFin)+'")'+
               ' AND PSA_DATEENTREE<="'+UsDateTime(EnregDucsEdi.DateFin)+'"';
// d PT34
   StQL := StQL +' and ((pci_debutcontrat <= "'+UsDateTime(EnregDucsEdi.DateFin)+'" and '+
                 'pci_fincontrat >= "'+UsDateTime(EnregDucsEdi.DateDebut)+'" or '+
                 'pci_fincontrat = "'+UsDateTime(IDate1900)+'") or '+
                 '(pci_typecontrat is  NULL) or (psa_salarie not in '+
                 '(select pci_Salarie from contrattravail where '+
                 '(pci_debutcontrat <= "'+UsDateTime(EnregDucsEdi.DateFin)+'" and '+
                 'pci_fincontrat >= "'+UsDateTime(EnregDucsEdi.DateDebut)+'" or '+
                 'pci_fincontrat = "'+UsDateTime(IDate1900)+'"))))';

// f PT34
//d PT30 DUCS EDI V4.2
  QL := OpenSQL('SELECT PSA_SALARIE,PSA_SEXE,PSA_CATDADS,PSA_DATENAISSANCE,'+
                'PSA_DADSPROF,PSA_PROFIL FROM SALARIES '+
// d PT34
                ' left join contrattravail on psa_salarie = pci_salarie'+
// f PT34
                ' WHERE '+StQL,True) ;
  TEffectifs := TOB.Create('Les effectifs', NIL, -1);
  TEffectifs.LoadDetailDB('SAL','','',QL,False);
  Ferme(QL) ;
  if (TEffectifs.Detail.Count <> 0) then
  begin
    DateNaiss := PlusMois(EnregDucsEdi.DateFin,-780);

    for I := 0 to TEffectifs.Detail.Count-1 do
    begin
      LeSal := TEffectifs.Detail[I];
      if (LeSal.GetValue('PSA_DATENAISSANCE') <=  DateNaiss) then
        LeSal.AddChampSupValeur('+65ANS','X')
      else
         LeSal.AddChampSupValeur('+65ANS','-');
    end;
    {nbre d'hommes}
    EnregDucsEdi.TotHommes := TEffectifs.Somme('PSA_SEXE',['PSA_SEXE'],['M'],TRUE, TRUE);
    {nbre de femmes}
    EnregDucsEdi.TotFemmes := TEffectifs.Somme('PSA_SEXE',['PSA_SEXE'],['F'],TRUE, TRUE);
    {Nbre d'apprentis}
    EnregDucsEdi.TotApprenti:= TEffectifs.Somme('PSA_CATDADS',['PSA_CATDADS'],['003'],TRUE, TRUE);
    {Nbre d'apprentis hommes}
    EnregDucsEdi.ApprentiH:= TEffectifs.Somme('PSA_SEXE',['PSA_SEXE','PSA_CATDADS'],['M','003'],TRUE, TRUE);
    {Nbre d'apprentis femmes}
    EnregDucsEdi.ApprentiF:= TEffectifs.Somme('PSA_SEXE',['PSA_SEXE','PSA_CATDADS'],['F','003'],TRUE, TRUE);
    {Nbre de cadres hommes (cadre ou dirigeant)}
    EnregDucsEdi.CadreH := TEffectifs.Somme('PSA_SEXE',['PSA_SEXE','PSA_DADSPROF'],['M','22'],TRUE, TRUE)+
                           TEffectifs.Somme('PSA_SEXE',['PSA_SEXE','PSA_DADSPROF'],['M','23'],TRUE, TRUE)+
                           TEffectifs.Somme('PSA_SEXE',['PSA_SEXE','PSA_DADSPROF'],['M','24'],TRUE, TRUE)+
                           TEffectifs.Somme('PSA_SEXE',['PSA_SEXE','PSA_DADSPROF'],['M','25'],TRUE, TRUE);
    {Nbre de cadres femmes (cadre ou dirigeant)}
    EnregDucsEdi.CadreF := TEffectifs.Somme('PSA_SEXE',['PSA_SEXE','PSA_DADSPROF'],['F','22'],TRUE, TRUE)+
                           TEffectifs.Somme('PSA_SEXE',['PSA_SEXE','PSA_DADSPROF'],['F','23'],TRUE, TRUE)+
                           TEffectifs.Somme('PSA_SEXE',['PSA_SEXE','PSA_DADSPROF'],['F','24'],TRUE, TRUE)+
                           TEffectifs.Somme('PSA_SEXE',['PSA_SEXE','PSA_DADSPROF'],['F','25'],TRUE, TRUE);
    {nbre salariés hommes de + de 65 ans}
    EnregDucsEdi.E65H := TEffectifs.Somme('PSA_SEXE',['PSA_SEXE','+65ANS'],['M','X'],TRUE, TRUE);

    {nbre salariés femmes de + de 65 ans}
    EnregDucsEdi.E65F := TEffectifs.Somme('PSA_SEXE',['PSA_SEXE','+65ANS'],['F','X'],TRUE, TRUE);

    {nbre salariés hommes titulaire d'un contrat de professionnalisation}
//PT37    EnregDucsEdi.CesH := TEffectifs.Somme('PSA_SEXE',['PSA_SEXE','PSA_PROFIL'],['M','023'],TRUE, TRUE);
    EnregDucsEdi.ProfH := TEffectifs.Somme('PSA_SEXE',['PSA_SEXE','PSA_PROFIL'],['M','296'],TRUE, TRUE);

    {nbre salariés femmes titulaire d'un contrat de professionnalisation}
//PT37    EnregDucsEdi.CesF := TEffectifs.Somme('PSA_SEXE',['PSA_SEXE','+PSA_PROFIL'],['F','023'],TRUE, TRUE);
    EnregDucsEdi.pROFF := TEffectifs.Somme('PSA_SEXE',['PSA_SEXE','+PSA_PROFIL'],['F','296'],TRUE, TRUE);

// d PT34
    {Nbre d'hommes en CDI}
    EnregDucsEdi.CDIH:= TEffectifs.Somme('PSA_SEXE',['PSA_SEXE','PCI_TYPECONTRAT'],['M','CDI'],TRUE, TRUE);

    {Nbre de femmes en CDI}
    EnregDucsEdi.CDIF:= TEffectifs.Somme('PSA_SEXE',['PSA_SEXE','PCI_TYPECONTRAT'],['F','CDI'],TRUE, TRUE);

    {Nbre d'hommes en CDD}
    EnregDucsEdi.CDDH:= TEffectifs.Somme('PSA_SEXE',['PSA_SEXE','PCI_TYPECONTRAT'],['M','CCD'],TRUE, TRUE);

    {Nbre de femmes en CDD}
    EnregDucsEdi.CDDF:= TEffectifs.Somme('PSA_SEXE',['PSA_SEXE','PCI_TYPECONTRAT'],['F','CCD'],TRUE, TRUE);

    {Nbre d'hommes en CNE}
    EnregDucsEdi.CNEH:= TEffectifs.Somme('PSA_SEXE',['PSA_SEXE','PCI_TYPECONTRAT'],['M','CNE'],TRUE, TRUE);

    {Nbre de femmes en CNE}
    EnregDucsEdi.CNEF:= TEffectifs.Somme('PSA_SEXE',['PSA_SEXE','PCI_TYPECONTRAT'],['F','CNE'],TRUE, TRUE);

// f PT34
  end;
  FreeAndNIl(Teffectifs);
// f PT30 DUCS EDI V4.2

  {Nbre de salariés rémunérés pour la période}
   StQL := '';
   StQL :='PHB_ETABLISSEMENT="'+EnregDucsEdi.Etab+'" AND ';
   StQL := StQL+ 'PHB_ORGANISME="'+EnregDucsEdi.Organisme+'" ';

   StQL := StQL+'AND '+
                'PHB_DATEDEBUT >="'+UsDateTime(EnregDucsEdi.DateDebut)+'" AND '+
                'PHB_DATEFIN <="'+UsDateTime(EnregDucsEdi.DateFin)+'" AND '+
                '(PHB_BASECOT <> 0 OR PHB_MTSALARIAL <> 0 OR PHB_MTPATRONAL <> 0)';

   QL := OpenSQL('SELECT Count(DISTINCT PHB_SALARIE) AS NOMBRE FROM HISTOBULLETIN WHERE '+
                 StQL,True) ;
   if not QL.EOF then // PortageCWAS
     EnregDucsEdi.NbSalFpe := QL.FindField('NOMBRE').AsInteger ;
   Ferme(QL) ;
end;  
{ fin InitEffectifsUNEDIC}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 13/05/2002
Modifié le ... :   /  /
Description .. : Chargement de la TOB des lignes détails à traiter +
Suite ........ : contrôles sur  champs obligatoires
Mots clefs ... : PAIE, PGDUCSEDI
*****************************************************************}
procedure  InitCotisation(var EnregDucsEdi : TDucsEdi;var EnrDucsEdiGroupee : TDucsEdiGroupee;Trait : string);
var
   QL                           : TQuery;
   StQL                         : string;
   Erreur                       : boolean;
   LesLignes                    : TOB;
   Wetab,WNoDucs                : string;
   Pred, QualifiantCotis        : string;
   DebCodif, SuitCodif          : string;
begin
  Erreur := false;
  if (Trait <> 'G') then //PT37
    EnregDucsEdi.Neant := False;

  if (Trait <> 'G') then
  begin
    EnregDucsEdi.ErrCotisqual := False;
    EnregDucsEdi.ErrInstitution := False;
    EnregDucsEdi.ErrCondition := False;
    Wetab := EnregDucsEdi.Etab;
    WNoducs := IntToStr(EnregDucsEdi.NoDucs);
  end
  else
  begin
    EnrDucsEdiGroupee.ErrCotisqual := False;
    EnrDucsEdiGroupee.ErrInstitution := False;
    EnrDucsEdiGroupee.ErrCondition := False;
    Wetab := EnrDucsEdiGroupee.Etab;
    WNoducs := IntToStr(EnrDucsEdiGroupee.NoDucs);
  end;

  if (copy(EnregDucsEdi.TypDucs,1,1) = 'I') then
  { IRC : Contrôle alimentation des champs PDD_INSTITUTION et PDD_CONDITION}
  {       non renseignés}
  begin
    StQL := 'PDD_ETABLISSEMENT = "'+WEtab+'" AND '+
            'PDD_ORGANISME = "'+EnregDucsEdi.Organisme+'" AND '+
            'PDD_DATEDEBUT = "'+UsDateTime(EnregDucsEdi.DateDebut)+'" AND '+
            'PDD_DATEFIN = "'+UsDateTime(EnregDucsEdi.DateFin)+'" AND '+
            'PDD_NUM = '+ WNoDucs+' AND '+
            'PDD_TYPECOTISATION <> "S" AND PDD_TYPECOTISATION <> "I"';
    QL := OpenSQL('SELECT COUNT(*)FROM DUCSDETAIL WHERE '+StQL+
                  ' AND (PDD_INSTITUTION = "" OR PDD_INSTITUTION IS NULL '+
                  'OR PDD_INSTITUTION = "    ")',True) ;
    if not QL.EOF then
    begin
      if (QL.Fields[0].AsInteger <> 0) then
      begin
        if (Trait <> 'G') then
          EnregDucsEdi.ErrInstitution :=  True
        else
          EnrDucsEdiGroupee.ErrInstitution :=  True;
        Erreur := true;
      end;
    end;
    ferme (QL);
    QL := OpenSQL('SELECT COUNT(*)FROM DUCSDETAIL WHERE '+StQL+
                  ' AND (PDD_CONDITION = "" OR PDD_CONDITION IS NULL '+
                  'OR PDD_CONDITION = "    ")',True) ;
    if not QL.EOF then
    begin
      if (QL.Fields[0].AsInteger <> 0) then
      begin
        if (Trait <>'G') then
          EnregDucsEdi.ErrCondition :=  True
        else
          EnrDucsEdiGroupee.ErrCondition :=  True;

        Erreur := true;
      end;
    end;
    ferme (QL);
  end;

  if (Erreur = false) then
  { Aucune erreur détectée}
  begin
// PT30 DUCS EDI V4.2
//  if (TOB_Lignes <> NIL) then
    if (TOB_Lignes <> NIL) and
       ((copy(EnregDucsEdi.TypDucs,1,1) <> 'U')  or (Trait <> 'G')) then
    begin
      TOB_Lignes.Free;
      TOB_Lignes := NIL;
    end;
    { Chargement de la TOB des Lignes Détail}
    if (TOB_Lignes = NIL) then        // PT30 DUCS EDI V4.2
      TOB_Lignes := TOB.Create ('Les lignes de cotisation', NIL, -1);

    QL :=OpenSql ('SELECT * FROM DUCSDETAIL '+
                  'WHERE '+
                  '(PDD_ETABLISSEMENT ="'+ WEtab +'") AND '+
                  '(PDD_ORGANISME ="'+ EnregDucsEdi.Organisme+'") AND '+
                  '(PDD_DATEDEBUT ="'+UsDateTime(EnregDucsEdi.DateDebut)+'") AND '+
                  '(PDD_DATEFIN = "'+UsDateTime(EnregDucsEdi.DateFin)+'") AND '+
                  '(PDD_NUM = '+WNoDucs+') AND'+
// d PT30 DUCS EDI V4.2
                  '(PDD_TYPECOTISATION <> "S" AND PDD_TYPECOTISATION <> "I")', TRUE);
// f PT30 DUCS EDI V4.2
    TOB_Lignes.LoadDetailDB ('DUCSDETAIL', '','',QL,FALSE,FALSE);
    Ferme(QL);

// d PT30 DUCS EDI V4.2
    // pour une ducs edi UNEDIC avec paiement groupé le GP2 de l'établissement
    // principal est suivi du GP3 correspondant à l'ensemble des cotisations (tous
    // établissements). Il faut donc compléter la TOB_Lignes (cumul des codifs
    // et des taux identiques).
{PT34    if (copy(EnregDucsEdi.TypDucs,1,1) = 'U') and
       ((EnregDucsEdi.TypBordereau = '923') or (EnregDucsEdi.TypBordereau = '926')) then }
// d PT43
   if (copy(EnregDucsEdi.TypDucs,1,1) = 'U') and
       ((EnregDucsEdi.TypBordereau = '923') or (EnregDucsEdi.TypBordereau = '926')or (EnregDucsEdi.TypBordereau = '921')) then
// if (copy(EnregDucsEdi.TypDucs,1,1) = 'U') then
// f PT43
      DetUnedicGroupee(EnregDucsEdi,EnrDucsEdiGroupee)
    else
// f PT30 DUCS EDI V4.2

      TOB_Lignes.Detail.Sort('PDD_ETABLISSEMENT;PDD_ORGANISME;PDD_DATEDEBUT;'+
                           'PDD_DATEFIN;PDD_INSTITUTION;PDD_CODIFICATION;PDD_DATECHGTTAUX');

// d PT30 DUCS EDI V4.2
//    if (TOB_Lignes.FillesCount(1) = 0) then EnregDucsEdi.Neant := True;
    if (TOB_Lignes.FillesCount(1) = 0) and (Trait <> 'G') then
      EnregDucsEdi.Neant := True;
    if (TOB_Lignes.FillesCount(1) <> 0) and (Trait = 'G') and
       (EnregDucsEdi.Neant = True) then
      EnregDucsEdi.Neant := False;
// f PT30 DUCS EDI V4.2


    if (copy (EnregDucsEdi.TypDucs,1,1)= 'I') and
       (TOB_Lignes <> NIL) then
    {IRC : Le champ PDP_COTISQUAL de la table DUCSPARAM doit être
           renseigné                                             }
    begin
      LesLignes := TOB_Lignes.FindFirst ([''],[''],TRUE);
      While LesLignes <> NIL  do
      begin
        if (Leslignes.getValue('PDD_TYPECOTISATION') <> 'I') and
           (Leslignes.getValue('PDD_TYPECOTISATION') <> 'S') then
        { les lignes Intitulé et Sous-Total ne sont pas traitées.}
        begin
          DebCodif := copy(Leslignes.getValue('PDD_CODIFICATION'),1,2);
          SuitCodif := copy(Leslignes.getValue('PDD_CODIFICATION'),3,5);
// PT28   while (((QualifiantCotis = '') or
          if (((QualifiantCotis = '') or
                  (QualifiantCotis = NULL) or
                  (QualifiantCotis = '   ')) and
// PT28          (StrToInt(copy(DebCodif,2,1)) >= 0)) do
                 (StrToInt(copy(DebCodif,2,1)) >= 0)) then
          begin
            QL := OpenSql('SELECT PDP_PREDEFINI,PDP_COTISQUAL'+
                          ' FROM DUCSPARAM '+
                          ' WHERE ##PDP_PREDEFINI## PDP_CODIFICATION = "'+
                          DebCodif +  SuitCodif +'"'+
                          'ORDER BY PDP_PREDEFINI',True);
            while  not QL.eof do
            begin
              Pred :=  QL.FindField('PDP_PREDEFINI').AsString;
              QualifiantCotis := QL.FindField('PDP_COTISQUAL').AsString;
              if (QL.FindField('PDP_PREDEFINI').AsString = 'DOS') then break;
                QL.Next;
            end;
            ferme (QL);
//PT28      DebCodif := IntToStr(StrToInt(DebCodif)-1);
          end;
          if (Pred  <> '') then
            if (QualifiantCotis = '') or
               (QualifiantCotis = NULL) or
               (QualifiantCotis = '   ') then
              if (Trait <> 'G') then
                EnregDucsEdi.ErrCotisqual := True
              else
                EnrDucsEdiGroupee.ErrCotisqual := True;
        end;
        LesLignes := TOB_Lignes.FindNext ([''],[''],TRUE);
      end;
    end;
  end;
end;
{ fin InitCotisation}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 12/04/2002
Modifié le ... :   /  /
Description .. : Contrôle de la validité des champs nécessaire à l'en tête de
Suite ........ : la déclaration
Mots clefs ... : PAIE, PGDUCSEDI
*****************************************************************}
Function  ControlVarDucsEnTete (EnregDucsEdi : TDucsEdi) : Boolean;
var
  i                            : integer;
  st, tiret                    : string;
  An, Mois,Trimestre,Libperiod : string;
  NumMois                       : integer;
begin
  result := True;

  tiret := '';
  Writeln(FRapport, '');
  if (EnregDucsEdi.DucsDossier = true) then
    st :=  EnregDucsEdi.Etab+' '+EnregDucsEdi.LibelleET+' (MULTI)'
  else
    st := EnregDucsEdi.Etab+' '+EnregDucsEdi.LibelleET+' (MONO)';

  libperiod := '';
  An := copy(EnregDucsEdi.Abrege,1,2);
  Mois := copy (EnregDucsEdi.Abrege,4,1);
  Trimestre := copy (EnregDucsEdi.Abrege,3,1);
  if (Mois <> '0') and (trimestre<>'0') then
  {Il s'agit d'une déclaration mensuelle}
  begin
    NumMois:=((StrToInt(Trimestre)-1)*3)+StrToInt(Mois);
    Mois:=IntToStr(NumMois);
    if Length(Mois)=1 then Mois:='0'+Mois;
    libperiod := RechDom ('PGMOIS', Mois,FALSE) +' 20'+An ;
  end
  else
    if (Mois='0') and (Trimestre<>'0') then
    { Il s'agit d'une déclaration trimestrielle}
    begin
      LibPeriod := Trimestre+' TRIMESTRE 20'+An;
    end
    else
      if ((Mois = '0') and (Trimestre = '0')) then
      begin
      {Il s'agit d'une déclaration annuelle}
// d PT30 DUCS EDI V4.2
{        if (VH_Paie.PGDecalage = true) or (VH_Paie.PGDecalagePetit = true) then
          libperiod := 'ANNEE 20'+ColleZeroDevant((StrToInt(An)-1),2)
        else}
// f PT30 DUCS EDI V4.2
          libperiod := 'ANNEE 20'+An;
      end;
  libperiod := Uppercase(libperiod);

  if (Copy(EnregDucsEdi.TypDucs,1,1) <> 'I') then
  st := st+'  '+Libperiod+'  Déclaration '+copy(EnregDucsEdi.TypDucs,3,6)+
        ' pour '+EnregDucsEdi.LibellePog
  else
    st := st+'  '+Libperiod+'  Déclaration IRC pour '+
          EnregDucsEdi.LibellePog ;


  Writeln(FRapport,st);

  for i:=0 to (Length(EnregDucsEdi.Etab)+length(EnregDucsEdi.LibelleET)+9) do
    tiret := tiret + '-';
  Writeln(FRapport, tiret);

// d PT30 DUCS EDI V4.2
  Writeln(FRapport,RechDom('PGTYPBORDEREAU',EnregDucsEdi.TypBordereau, False));
// f PT30 DUCS EDI V4.2

   { PT12 formatage du nom de fichier }
  if (Copy(EnregDucsEdi.TypDucs,1,1) = 'I') then
  { Caisse IRC : on ajoute le code institution }
    NomFicDucs := 'DU'+
                  Copy(EnregDucsEdi.TypDucs,1,1)+
                  EnregDucsEdi.Siret+
                  EnregDucsEdi.Abrege+
                  EnregDucsEdi.InstitutionPog+
                  '.DUC'
  else
  { Caisse non IRC}
    NomFicDucs := 'DU'+
                  Copy(EnregDucsEdi.TypDucs,1,1)+
                  EnregDucsEdi.Siret+
                  EnregDucsEdi.Abrege+
                  '.DUC';

  Writeln(FRapport, '');
  Writeln(FRapport,'  Fichier : '+NomFicDucs);
  Writeln(FRapport, '');
  if (Copy(EnregDucsEdi.TypDucs,1,1) <> 'A') and
     (Copy(EnregDucsEdi.TypDucs,1,1) <> 'I') and
     (Copy(EnregDucsEdi.TypDucs,1,1) <> 'U') then
  begin
    Writeln(FRapport, '  Anomalie   : L''organisme traité n''est pas de nature DUCS-EDI');
    Writeln(FRapport, '  --> Action : Vérification de la fiche organisme.');
    result := False;
  end;


  { Qualifiant (Segment UNB)}
  if (EnregDucsEdi.IdentQual <> '5') and
//PT48 Correction     (EnregDucsEdi.IdentQual <> '22') and
     (EnregDucsEdi.IdentQual <> 'ZZZ') then
  begin
    if (EnregDucsEdi.IdentQual = '22') then
        Writeln(FRapport, '  Anomalie   : Le qualifiant de type SIREN n''est plus accepté (DUCS V4.2)')
    else
      Writeln(FRapport, '  Anomalie   : Le qualifiant n''est pas renseigné');
    Writeln(FRapport, '  --> Action : Modification de la fiche organisme.');
    result := False;
  end;
  {Identifiant émetteur (Segment UNB)}
  if (EnregDucsEdi.IdentEmet = '') then
  begin
    Writeln(FRapport, '  Anomalie   : L''identifiant émetteur n''est pas renseigné');
    Writeln(FRapport, '  --> Action : Modification de la fiche organisme.');
    result := False;
  end;
  {Identifiant destinataire (Segment UNB)}
  if (EnregDucsEdi.IdentDest = '') then
  begin
    Writeln(FRapport, '  Anomalie   : L''identifiant destinataire n''est pas renseigné');
    Writeln(FRapport, '  --> Action : Modification de la fiche organisme.');
   result := False;
  end;
// d PT30 DUCS EDI V4.2
  {Indicateur de traitement (type de bordereau) (Segment GIS)}
  if (EnregDucsEdi.TypBordereau = '') then
  begin
    Writeln(FRapport, '  Anomalie   : Le type de bordereau n''est pas renseigné');
    Writeln(FRapport, '  --> Action : Modification de la fiche organisme (onglet affiliation). ');
   result := False;
  end;
// f PT30 DUCS EDI V4.2
end;
{fin ControlVarDucsEnTete}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 15/04/2002
Modifié le ... :   /  /
Description .. : message d'abandon dans le fichier de log
Mots clefs ... : PAIE, PGDUCSEDI
*****************************************************************}
procedure Abandonfichier(Trait : string);
var
  Nomcomplet                    : string;
begin
  Writeln(FRapport, '');
  Writeln(FRapport, '!! La confection du fichier est abandonnée.');
  if (Trait = 'G') then
  begin
    CloseFile(FDUCSEDI);
// d PT30 DUCS EDI V4.2
//{$IFDEF EAGLCLIENT}
    Nomcomplet := VH_Paie.PGCheminEagl+'\'+NomFicDucs;
//{$ELSE}
//    Nomcomplet := V_PGI.DatPath+'\'+NomFicDucs;
//{$ENDIF}
// f PT30 DUCS EDI V4.2
    DeleteFile(PChar(Nomcomplet));
  end;
// d PT30 DUCS EDI V4.2 (suite memcheck on vide les tob)
    TOB_Lignes.Free; TOB_Lignes := NIL;

    if (TOB_Transport <> Nil) then
    begin
      TOB_Transport.Free;
      TOB_Transport := NIL;
    end;
    if (TOB_TAT <> Nil) then
    begin
      TOB_TAT.Free;
      TOB_TAT := NIL;
    end;
//PT34    TOB_DucsGroupee.Free; TOB_DucsGroupee := NIL;
      FreeAndNIl(TOB_DucsGroupee);

// f PT30 DUCS EDI V4.2
end;
{Abandonfichier}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 15/04/2002
Modifié le ... :   /  /
Description .. : Contrôle de validité des champs nécessaires au
Suite ........ : groupe 1 de la déclaration
Mots clefs ... : PAIE, PGDUCSEDI
*****************************************************************}
Function  ControlVarDucsGrp1 (EnregDucsEdi : TDucsEdi) : Boolean;
begin
  result := True;

  { Infos EMETTEUR
    --------------
  Nom du déclarant(émetteur) (segment NADFR)}
  if (EnregDucsEdi.Declarant ='') then
  begin
    Writeln(FRapport, '  Anomalie   : Le nom de l''émetteur n''est pas '+
                      'renseigné');
    Writeln(FRapport, '  --> Action : En consultation des Ducs renseigner '+
                      'les coordonnées du déclarant.');
    result := False;
  end;
  { Nom du contact chez l'émetteur (segment CTAFR)}
  if (EnregDucsEdi.DeclarantSuite ='') then
  begin
    Writeln(FRapport, '  Anomalie   : Le nom du contact chez l''émetteur '+
                      'n''est pas renseigné');
    Writeln(FRapport, '  --> Action : En consultation des Ducs renseigner '+
                      'les coordonnées du déclarant.');
    result := False;
  end;
  { N° Tel ou n° Fax chez l'émetteur (segments COMFRT et COMFRF)}
  if (EnregDucsEdi.TelephoneDecl = '') and
     (EnregDucsEdi.FaxDeclarant = '')  and
     (EnregDucsEdi.EmailDeclarant = '') then // PT30 DUCS EDI V4.2
  begin
    Writeln(FRapport, '  Anomalie   : Le n° de téléphone ou de fax '+
                      'du contact chez l''émetteur n''est pas renseigné');
    Writeln(FRapport, '  --> Action : En consultation des Ducs renseigner '+
                      'les coordonnées du déclarant.');
    result := False;
  end;
  { Identification émetteur chez le destinataire (segment RFF IRC uniquement)}
  if (Copy(EnregDucsEdi.TypDucs,1,1)= 'I') then
    if (EnregDucsEdi.NoContEmet = '') then
    begin
      Writeln(FRapport, '  Anomalie   : La référence de l''émetteur chez le '+
                        'destinataire n''est pas renseignée');
      Writeln(FRapport, '  --> Action : Modification de la fiche organisme '+
                        'renseigner le n° contrat émetteur');
      result := False;
    end;

  {Infos DESTINATAIRE
   ------------------
   Adresse Oragnisme (destinataire) (segment NADMR)}
  if ((EnregDucsEdi.Adresse1Pog = '') and
      (EnregDucsEdi.Adresse2Pog = '')) or
     (EnregDucsEdi.VillePog = '') or (EnregDucsEdi.CpPog = '') then
  begin
    Writeln(FRapport, '  Anomalie   : L''adresse du destinataire est'+
                      ' incomplète');
    Writeln(FRapport, '  --> Action : Modification de la fiche organisme,'+
                      ' renseigner Ville + Code postal + au moins une ligne '+
                      'd''adresse.' );
    result := False;
  end;
  {Référence destinataire (n° institution) (segment RFFMR - IRC uniquement)}
  if (copy(EnregDucsEdi.TypDucs,1,1) = 'I') then
    if (EnregDucsEdi.InstitutionPog = ' ') then
    begin
      Writeln(FRapport, '  Anomalie   : La référence destinataire (code '+
                        'institution)  n''est pas renseignée');
      Writeln(FRapport, '  --> Action : Modification de la fiche organisme,'+
                        ' renseigner le code institution.' );
      result := False;
    end;
// d PT30 DUCS EDI V4.2
  {Siret du destinataire}
  if (EnregDucsEdi.SiretPog = '') then
  begin
    Writeln(FRapport, '  Anomalie   : Le siret du destinataire '+
                        'n''est pas renseigné');
    Writeln(FRapport, '  --> Action : Modification de la fiche organisme,'+
                        ' renseigner le siret (onglet Adresse).' );
    result := False;
  end;
// f PT30 DUCS EDI V4.2
  {Infos ADHERENT
   --------------

  Siret adhérent (segment NADDT)}
  if (EnregDucsEdi.Siret = '') or
     ((Length(EnregDucsEdi.Siret) <> 14) and (Length(EnregDucsEdi.Siret) <> 9)) or
     (ControlSiret(EnregDucsEdi.Siret) <> True) then
  begin
    Writeln(FRapport, '  Anomalie   : Le siret de l''adhérent est incomplet '+
                      'ou erroné ');
    Writeln(FRapport, '  --> Action : En consultation des Ducs renseigner '+
                      'le siret. Penser à rectifier '+
                      'la fiche établissement');
    result := False;
  end;
  {Nom de la personne à contacter chez l'adhérent.
   ----------------------------------------------}
  if (EnregDucsEdi.AdherContact = '') then
  begin
    Writeln(FRapport, '  Anomalie   : Le nom du contact chez l''adhérent n'''+
                      'est pas renseigné');
    Writeln(FRapport, '  --> Action : Modification de la fiche organisme '+
                      'renseigner le Contact adhérent');
    result := False;
  end;

  {Adresse adhérent  (segment NADDT)}
  if ((EnregDucsEdi.Adresse1ET = '') and
      (EnregDucsEdi.Adresse2ET = '')) or
     (EnregDucsEdi.VilleET = '') or (EnregDucsEdi.CpET = '') then
  begin
    Writeln(FRapport, '  Anomalie   : L''adresse de l''adhérent est'+
                      ' incomplète');
    Writeln(FRapport, '  --> Action : Modification de la fiche établissement,'+
                      ' renseigner Ville + Code postal + au moins une ligne '+
                      'd''adresse.' );
    result := False;
  end;
  {N° de téléphone ou de fax du contact chez l'adhérent (segments COMDTT et COMDTF)}
  if (EnregDucsEdi.TelET = '') and (EnregDucsEdi.FaxET = '')and
     (EnregDucsEdi.EmailET = '') then // PT30 DUCS EDI V4.2
  begin
    Writeln(FRapport, '  Anomalie   : Le n° de téléphone ou de fax ou l''adresse'+
                      ' Email de l''adhérent n''est pas renseigné');
    Writeln(FRapport, '  --> Action : Modification de la fiche établissement,'+
                      ' renseigner le n° de téléphone ou de fax ou l''adresse Email.');
    result := False;
  end;

  { Code Ape adhérent (segment RFFACT)}
  if (EnregDucsEdi.ApePdu = '') then
  begin
    Writeln(FRapport, '  Anomalie   : Le code Ape de l''adhérent est incomplet '+
                      'ou erroné ');
    Writeln(FRapport, '  --> Action : En consultation des Ducs renseigner '+
                      'le code Ape. Penser à '+
                      'rectifier la fiche établissement');
    result := False;
  end;
//  d PT36
  { Identifiant de l'entreprise chez l'OPS - n° interne (segment RFFNUM)}
  if (copy(EnregDucsEdi.TypDucs,1,1) = 'I') then
    if (EnregDucsEdi.NumeroInterne = '') then
    begin
      Writeln(FRapport, '  Anomalie   : Le n° interne de l''adhérent est incomplet '+
                      'ou erroné ');
      Writeln(FRapport, '  --> Action : En consultation des Ducs renseigner '+
                      'le n° interne. Penser à '+
                      'rectifier la fiche de l''oragnisme destinataire');
      result := False;
  end;
// f  PT36  
end;
{ fin ControlVarDucsGrp1}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 13/05/2002
Modifié le ... : 13/05/2002
Description .. : Contrôle des champs obligatoires au paiement
Suite ........ : Inscription des anomalies dans le fichier LOG
Mots clefs ... : PAIE, PGDUCSEDI
*****************************************************************}
Function ControlPaiement (EnregDucsEdi : TDucsEdi) : Boolean;
begin
  result := true;
  if (copy(EnregDucsEdi.TypDucs,1,1) = 'I') and
     ((EnregDucsEdi.ModePaiement  = 'Z10') or
      (EnregDucsEdi.ModePaiement = '30') or
      (EnregDucsEdi.ModePaiement = '31')) and
     (EnregDucsEdi.Reglement = IDate1900) then
    begin
      Writeln(FRapport, '  Anomalie   : Date de règlement, '+
                        'la date n''est pas renseignée');

      Writeln(FRapport, '  --> Action : En consultation des Ducs renseigner '+
                        'la date de règlement');
      result := False;
    end;

  if (EnregDucsEdi.ModePaiement = 'Z10') and
     ((EnregDucsEdi.IdentOPS = '' ) or
      (EnregDucsEdi.IdentOPS = Null) or
      (EnregDucsEdi.IdentOPS = '        ' )) then
    begin
      Writeln(FRapport, '  Anomalie   : Télérèglement EDI, l'''+
                        'identification OPS n''est pas renseigné');

      Writeln(FRapport, '  --> Action : Modification organisme : renseigner'+
                        ' l''identification OPS ou bien modifier le mode de'+
                        ' règlement ');
      result := False;
    end;
// d PT47
// contrôle du champ Titulaire du compte dans le cas du télérèglement pour
// toutes les natures d'organisme et pour les IRC du virement, du prélèvement
// et du télérèglement
// (champ obligatoire)
  if ((EnregDucsEdi.ModePaiement = 'Z10') or
       ((copy(EnregDucsEdi.TypDucs,1,1) = 'I') and
        ((EnregDucsEdi.ModePaiement = '30') or
         (EnregDucsEdi.ModePaiement = '31')))) then
    if (EnregDucsEdi.TitulaireCpt ='') then
    begin
      Writeln(FRapport, '  Anomalie   : Télérèglement EDI, '+
                        'le titulaire du compte n''est pas renseigné');

      Writeln(FRapport, '  --> Action : Modification organisme : renseigner'+
                        ' le titulaire du compte');
      result := False;
    end;
// f PT47
end;
// fin ControlPaiement

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 13/05/2002
Modifié le ... :   /  /
Description .. : Inscription dans le fichier LOG des erreurs liées
Suite ........ : au contrôle des champs obligatoires des lignes
Suite ........ : de cotisation
Mots clefs ... : PAIE, PGDUCSEDI
*****************************************************************}
Function ControlCotisation(EnregDucsEdi : TDucsEdi;EnrDucsEdiGroupee : TDucsEdiGroupee;  Trait : string) : Boolean;
var
   st : string;
begin
result := True;
st := '  Etablissement concerné : ';
     if (trait <> 'G') then
       st := st + EnregDucsEdi.Etab+' '+EnregDucsEdi.LibelleET
     else
       st := st + EnrDucsEdiGroupee.Etab+' '+EnrDucsEdiGroupee .LibelleET ;
// Contrôle : n° interne obligatoire pour ACOSS et UNEDIC
    if ((copy(EnregDucsEdi.TypDucs,1,1)='A') and
        (copy(EnregDucsEdi.Abrege,3,2) <> '00')) or
       (copy(EnregDucsEdi.TypDucs,1,1)<>'A')then
      begin
       if (copy(EnregDucsEdi.TypDucs,1,1)<>'I') then
        begin
         // ACOSS UNEDIC    (n° interne)
         if ((Trait <> 'G') and (EnregDucsEdi.NumeroInterne = '')) or
         ((Trait = 'G') and (EnrDucsEdiGroupee.NumeroInterne = '')) then
           begin
             Writeln(FRapport, '  Anomalie   : Le Numéro interne n''est pas renseigné ');
             Writeln(FRapport,st);
             Writeln(FRapport, '  --> Action : En consultation des Ducs renseigner '+
                               'le numéro interne. Numéro à renseigner sur la fiche organisme');
             result := False;
           end;
        end ;
      end;
// contrôle : Régularisations uniquement sur TR annuel (ACOSS)
    if ((copy(EnregDucsEdi.TypDucs,1,1)='A') and
        (copy(EnregDucsEdi.Abrege,3,2) <> '00')) then
      if ((Trait <> 'G') and (EnregDucsEdi.Regularisation <> 0.0)) or
         ((Trait = 'G') and  (EnrDucsEdiGroupee.Regularisation <> 0.0)) then
        begin
          Writeln(FRapport, '  Anomalie   : Il ne s''agit pas d''un TR Annuel, '+
                            'pas de régularisation possible.');
          Writeln(FRapport,st);
          Writeln(FRapport, '  --> Action : En consultation des Ducs renseigner '+
                            'le montant de la régularisation à 0');
          result := False;
        end;
     if ((EnregDucsEdi.ErrTransport = True) and (Trait <> 'G')) or
        ((EnrDucsEdiGroupee.ErrTransport = True) and (Trait = 'G')) then
       begin
// PT30 DUCS EDI V4.2
//       Writeln(FRapport, '  Anomalie   : Le code Communauté Urbaine n''est pas '+
         Writeln(FRapport, '  Anomalie   : Le code Commune n''est pas '+
                           'renseigné sur une ligne détail concernant le transport');
         Writeln(FRapport,st);
         Writeln(FRapport, '  --> Action : En consultation des Ducs renseigner '+
//                           'le code Communauté Urbaine des lignes de transport');
                           'le code Commune des lignes de transport');
         result := False;
       end;
// PT30 DUCS EDI V4.2
     if ((EnregDucsEdi.ErrTAT = True) and (Trait <> 'G')) or
        ((EnrDucsEdiGroupee.ErrTAT = True) and (Trait = 'G')) then
       begin
         Writeln(FRapport, '  Anomalie   : Le code Commune n''est pas '+
                           'renseigné sur une ligne détail concernant la taxe'+
                           ' additionnelle de transport');
         Writeln(FRapport,st);
         Writeln(FRapport, '  --> Action : En consultation des Ducs renseigner '+
                           'le code Commune des lignes de taxe '+
                           'additionnelle de transport');
         result := False;
       end;
// PT30 DUCS EDI V4.2
     if ((EnregDucsEdi.ErrCotisqual = True) and (Trait <> 'G')) or
        ((EnrDucsEdiGroupee.ErrCotisqual = True) and (Trait = 'G'))then
       begin
         Writeln(FRapport, '  Anomalie   : Il manque un qualifiant de cotisation.');
         Writeln(FRapport,st);
         Writeln(FRapport, '  --> Action : Revenir sur la fiche codifications'+
                           ' des Ducs au menu paramètres et renseigner le champ.');
         result := False;
       end;
     if ((EnregDucsEdi.ErrInstitution = True) and (Trait <> 'G')) or
        ((EnrDucsEdiGroupee.ErrInstitution = True) and (Trait = 'G'))then
       begin
         Writeln(FRapport, '  Anomalie   : Il manque un code institution sur '+
                           'une ligne détail.');
         Writeln(FRapport,st);
         Writeln(FRapport, '  --> Action : En consultation des Ducs renseigner '+
                           'le code intitution sur les lignes détail.');
         result := False;
       end;
     if ((EnregDucsEdi.ErrCondition = True) and (Trait <> 'G')) or
        ((EnrDucsEdiGroupee.ErrCondition = True) and (Trait = 'G')) then
       begin
         Writeln(FRapport, '  Anomalie   : Il manque une condition spéciale'+
                           ' de cotisations sur une ligne détail.');
         Writeln(FRapport,st);
         Writeln(FRapport, '  --> Action : En consultation des Ducs renseigner '+
                           'la condition spéciale de cotisation sur les lignes '+
                           'détail');
         result := False;
       end;
end;
// fin  ControlCotisation

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 12/04/2002
Modifié le ... :   /  /
Description .. : Ouverture (création) du fichier Ducs Edi dossier
Mots clefs ... : PAIE, PGDUCSEDI
*****************************************************************}
procedure OuvreFichierDucsEdi(var EnregDucsEdi : TDucsEdi);
var
   Nomcomplet : string;
begin
// d PT30 DUCS EDI V4.2
//{$IFDEF EAGLCLIENT}
     Nomcomplet := Vh_Paie.PgCheminEagl+'\'+NomFicDucs;
//{$ELSE}
//     Nomcomplet := V_PGI.DatPath+'\'+NomFicDucs;
//{$ENDIF}
// f PT30 DUCS EDI V4.2
     if FileExists(Nomcomplet) then
      DeleteFile(PChar(Nomcomplet));

     AssignFile(FDUCSEDI, Nomcomplet);
{$i-}ReWrite(FDUCSEDI); {$i+}
     If IoResult<>0 Then
      Begin
       Writeln(FRapport,  'Fichier inaccessible : '+Nomcomplet);
       Exit ;
      End;
end;
// fin OuvreFichierDucsEdi

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 12/04/2002
Modifié le ... :   /  /    
Description .. : Construction des segments d'en-tête de déclaration
Mots clefs ... : PAIE, PGDUCSEDI
*****************************************************************}
procedure EnTeteDeclaration(var EnregDucsEdi : TDucsEdi);
var
   tampon, blanc : string;
begin
     blanc := StringOfChar(' ',35);
     // TYP - type de destinataire
     //---------------------------
     Writeln(FDUCSEDI,'TYP   '+
                      Copy(EnregDucsEdi.TypDucs,3,6));

     // UNB - En-tête d'interchange : Identification émetteur et destinataire
     //----------------------------------------------------------------------
     tampon := '';
     if (EnregDucsEdi.IdentQual = '5') then
          // identifiant Siret
          tampon := EnregDucsEdi.IdentQual+'   ';
{*PT48     if (EnregDucsEdi.IdentQual = '22') then
          // identifiant Siren
          tampon := EnregDucsEdi.IdentQual+'  '; }
     if (EnregDucsEdi.IdentQual = 'ZZZ') then
          // identifiant selon accord
          tampon := EnregDucsEdi.IdentQual+' ';
     Writeln(FDUCSEDI,'UNB   '+
                      tampon +
                      EnregDucsEdi.IdentEmet+
                      copy(blanc,1,(35-length(EnregDucsEdi.IdentEmet)))+
                      EnregDucsEdi.identDest+
                      copy(blanc,1,(35-length(EnregDucsEdi.identDest)))+
                      Copy(DateTimeToStr(Now),9,2)+
                      Copy(DateTimeToStr(Now),4,2)+
                      Copy(DateTimeToStr(Now),1,2)+
                      Copy(DateTimeToStr(Now),12,2)+
                      Copy(DateTimeToStr(Now),15,2));

     // BGM - Début de message
     //-----------------------
     //       Référence interne de la déclaration chez l'émetteur (nom fichier)
     if (copy(EnregDucsEdi.TypDucs,1,1) = 'I') then 
        { Caisse IRC }
        Writeln(FDUCSEDI,'BGM   '+
                         Copy(NomFicDucs ,1,25)+
                         copy(blanc,1,(35-25)))
     else
        {Caisse non IRC }
        Writeln(FDUCSEDI,'BGM   '+
                         Copy(NomFicDucs ,1,21)+
                         copy(blanc,1,(35-21)));

     // GIS - Type de déclaration(périodicité)
     //---------------------------------------
//     tampon := '';
//     if (Copy(EnregDucsEdi.TypDucs,1,1) = 'A') then
        // ACOSS
//        begin
//          if  (Copy(EnregDucsEdi.Abrege,3,1) <> '0') then
           // Mensuel ou Trimestriel
//           begin
//            if (EnregDucsEdi.DucsDossier <> True ) then
              // Un seul établissement
//              tampon := '913'
//            else
              // Plusieurs établissements
//              tampon := '914';
//           end
//         else
           // Annuel
//           begin
//            if (EnregDucsEdi.DucsDossier <> True) then
              // Un seul établissement
//              tampon := '915'
//            else
              // Plusieurs établissements
//              tampon := '916';
//           end;
//        end; // fin ACOSS

//     if (Copy(EnregDucsEdi.TypDucs,1,1) = 'U') then
        // UNEDIC
//        begin
//         if (Copy(EnregDucsEdi.Abrege,4,1) <> '0') or
//            (Copy(EnregDucsEdi.Abrege,3,1) <> '0') then
           // Mensuel ou Trimestriel
//           begin
//            if (EnregDucsEdi.DucsDossier <> True) then
              // Un seul établissement
//              tampon := '920'
//            else
              // Plusieurs établissements
//              tampon := '921';
//           end
//         else
           // Annuel
//           begin
//            if (EnregDucsEdi.DucsDossier <> True) then
              // Un seul établissement
//              tampon := '922'
//            else
              // Plusieurs établissements
//              tampon := '923';
//           end;
//        end; // fin UNEDIC

//     if (Copy(EnregDucsEdi.TypDucs,1,1) = 'I') then
        // IRC
{*        begin
         if (Copy(EnregDucsEdi.Abrege,4,1) <> '0') then*}
           // Mensuel
{*           tampon := '931';
         if (Copy(EnregDucsEdi.Abrege,4,1) = '0') and
            (Copy(EnregDucsEdi.Abrege,3,1) <> '0')then*}
           // Trimestriel
{*           tampon := '930';
         if (Copy(EnregDucsEdi.Abrege,4,1) = '0') and
            (Copy(EnregDucsEdi.Abrege,3,1) = '0')then
           // Annuel
           tampon := '932';
        end;*}
        // fin IRC
{*     Writeln(FDUCSEDI,'GIS   '+
                      tampon);
*}
     Writeln(FDUCSEDI,'GIS   '+
                      EnregDucsEdi.TypBordereau);
end;
// fin  EnTeteDeclaration

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 15/04/2002
Modifié le ... :   /  /    
Description .. : Construction des segments du groupe 1 de la déclaration
Mots clefs ... : PAIE, PGDUCSEDI
*****************************************************************}
procedure Groupe1Declaration(var EnregDucsEdi : TDucsEdi);
var
   tampon, blanc, periode : string;
begin
     blanc := StringOfChar(' ',70);          // PT30 DUCS EDI V4.2
     // EMETTEUR DU MESSAGE
     //--------------------
     //
     // NADFR -  Nom & adresse de l'émetteur
     // ------------------------------------
     Writeln(FDUCSEDI,'NADFR '+
                      EnregDucsEdi.SiretPet+
                      copy(blanc,1,(35-length(EnregDucsEdi.SiretPet)))+
                      EnregDucsEdi.Declarant+
                      copy(blanc,1,(35-length(EnregDucsEdi.Declarant)))+
                      EnregDucsEdi.AdressePet+
                      copy(blanc,1,(35-length(EnregDucsEdi.AdressePet)))+
                      EnregDucsEdi.Adresse2Pet+
                      copy(blanc,1,(35-length(EnregDucsEdi.Adresse2Pet)))+
                      EnregDucsEdi.VillePet+
                      copy(blanc,1,(35-length(EnregDucsEdi.VillePet)))+
                      EnregDucsEdi.CpPet);

      // CTAFR - Info sur le correspondant émetteur, personne à contacter
      // ----------------------------------------------------------------
     Writeln(FDUCSEDI,'CTAFR '+
                      EnregDucsEdi.DeclarantSuite+
                      copy(blanc,1,(35-length(EnregDucsEdi.DeclarantSuite))));

      // COMFRT - Mode de communication avec le contact (téléphone)
      // ----------------------------------------------------------
     if (EnregDucsEdi.TelephoneDecl <> '') then
        Writeln(FDUCSEDI,'COMFRT'+
                         EnregDucsEdi.TelephoneDecl+
                         copy(blanc,1,(70-length(EnregDucsEdi.TelephoneDecl)))+ // PT30 DUCS EDI V4.2
                         'TE');
      // COMFRF - Mode de communication avec le contact (fax)
      // ----------------------------------------------------------
      if (EnregDucsEdi.FaxDeclarant <> '') then
         Writeln(FDUCSEDI,'COMFRF'+
                          EnregDucsEdi.FaxDeclarant+
                          copy(blanc,1,(70-length(EnregDucsEdi.FaxDeclarant)))+ // PT30 DUCS EDI V4.2
                          'FX');
// d PT30 DUCS EDI V4.2
      // COMFRM - Mode de communication avec le contact (Email)
      // ----------------------------------------------------------
      if (EnregDucsEdi.EmailDeclarant <> '') then
         Writeln(FDUCSEDI,'COMFRM'+
                          EnregDucsEdi.EmailDeclarant+
                          copy(blanc,1,(70-length(EnregDucsEdi.EmailDeclarant)))+
                          'EM');
// f PT30 DUCS EDI V4.2
      // RFFFR - Référence de l'émetteur chez le destinataire (IRC uniquement)
      // ---------------------------------------------------------------------
      if (copy(EnregDucsEdi.TypDucs,1,1) = 'I') then
         Writeln(FDUCSEDI,'RFFFR '+
                          EnregDucsEdi.NoContEmet+
                          copy(blanc,1,(35-length(EnregDucsEdi.NoContEmet)))+
                          'FX');

     // DESTINATAIRE DU MESSAGE
     //--------------------
     //
     // NADMR -  Nom & adresse du destinataire
     // ------------------------------------
     Writeln(FDUCSEDI,'NADMR '+
                      EnregDucsEdi.SiretPog+
                      copy(blanc,1,(35-length(EnregDucsEdi.SiretPog)))+
                      EnregDucsEdi.LibellePog+
                      copy(blanc,1,(35-length(EnregDucsEdi.LibellePog)))+
                      EnregDucsEdi.Adresse1Pog+
                      copy(blanc,1,(35-length(EnregDucsEdi.Adresse1Pog)))+
                      EnregDucsEdi.Adresse2Pog+
                      copy(blanc,1,(35-length(EnregDucsEdi.Adresse2Pog)))+
                      EnregDucsEdi.VillePog+
                      copy(blanc,1,(35-length(EnregDucsEdi.VillePog)))+
                      EnregDucsEdi.CpPog);
     // CTAMR - Info sur le correspondant detinataire (IRC & ACOSS)
     // -------------------------------------------------------------
// d PT36
     if (((copy(EnregDucsEdi.TypDucs,1,1) = 'I') or
          (copy(EnregDucsEdi.TypDucs,1,1) = 'A')) and
         (EnregDucsEdi.ContactPog <> '')) then
// f PT36
       Writeln(FDUCSEDI,'CTAMR '+
                        EnregDucsEdi.ContactPog+
                        copy(blanc,1,(35-length(EnregDucsEdi.ContactPog))));
     // COMMRT - Mode de communication avec le contact destinataire (Tel)
     //          (IRC & ACOSS)
     // -----------------------------------------------------------------
// d PT36
     if (((copy(EnregDucsEdi.TypDucs,1,1) = 'I') or
          (copy(EnregDucsEdi.TypDucs,1,1) = 'A')) and
          (EnregDucsEdi.TelPog <> '')) then
// f PT36
       Writeln(FDUCSEDI,'COMMRT'+
                        EnregDucsEdi.TelPog+
                        copy(blanc,1,(70-length(EnregDucsEdi.TelPog)))+ // PT30 DUCS EDI V4.2
                        'TE');
     // COMMRF - Mode de communication avec le contact destinataire (Fax)
     //          (IRC & ACOSS)
     // -----------------------------------------------------------------
// d PT36
     if (((copy(EnregDucsEdi.TypDucs,1,1) = 'I') or
          (copy(EnregDucsEdi.TypDucs,1,1) = 'A')) and
          (EnregDucsEdi.FaxPog <> '')) then
// f PT36
       Writeln(FDUCSEDI,'COMMRF'+
                        EnregDucsEdi.FaxPog+
                        copy(blanc,1,(70-length(EnregDucsEdi.FaxPog)))+ // PT30 DUCS EDI V4.2
                        'FX');
// d PT30 DUCS EDI V4.2
     // COMMRM - Mode de communication avec le contact destinataire (Email)
     //          (IRC & ACOSS)
     // -----------------------------------------------------------------
{PT36     if (copy(EnregDucsEdi.TypDucs,1,1) = 'I') and
        (EnregDucsEdi.EmailPog <> '') and
        (EnregDucsEdi.EmailPog <> '') then
       Writeln(FDUCSEDI,'COMMRM'+
                        EnregDucsEdi.EmailPog+
                        copy(blanc,1,(70-length(EnregDucsEdi.EmailPog)))+
                        'EM');}
// f PT30 DUCS EDI V4.2

     // RFFMR - n° gagrément ministériel du groupe d'institution ou
     //         n° d'institution (IRC uniquement)
     // -----------------------------------------------------------------
     if (copy(EnregDucsEdi.TypDucs,1,1) = 'I') then
       Writeln(FDUCSEDI,'RFFMR '+
                        EnregDucsEdi.InstitutionPog+
                        copy(blanc,1,(35-length(EnregDucsEdi.InstitutionPog))));


    // ADHERENT (ENTREPRISE DECLARANTE)
    // --------------------------------
    //
    // NADDT - Nom & adresse de l'adhérent
    // -----------------------------------
    Writeln(FDUCSEDI,'NADDT '+
                     EnregDucsEdi.Siret+
                     Copy(blanc,1,(35-length(EnregDucsEdi.Siret)))+
                     EnregDucsEdi.LibelleET+
                     Copy(blanc,1,(35-length(EnregDucsEdi.LibelleET)))+
                     EnregDucsEdi.Adresse1ET+
                     Copy(blanc,1,(35-length(EnregDucsEdi.Adresse1ET)))+
                     EnregDucsEdi.Adresse2ET+
                     Copy(blanc,1,(35-length(EnregDucsEdi.Adresse2ET)))+
                     EnregDucsEdi.VilleET+
                     Copy(blanc,1,(35-length(EnregDucsEdi.VilleET)))+
                     EnregDucsEdi.CpET);
    // CTADT - Nom du contact chez l'adhérent
    // --------------------------------------
    Writeln(FDUCSEDI,'CTADT '+
                     EnregDucsEdi.AdherContact+
                     Copy(blanc,1,(35-length(EnregDucsEdi.AdherContact))));
    // COMDTT - Mode de communication avec le contact chez l'adhérent (Tel)
    // --------------------------------------------------------------------
    if (EnregDucsEdi.TelET <> '') then
      Writeln(FDUCSEDI,'COMDTT'+
                       EnregDucsEdi.TelET+
                       Copy(blanc,1,(70-length(EnregDucsEdi.TelET)))+ // PT30 DUCS EDI V4.2
                       'TE');
    // COMDTF - Mode de communication avec le contact chez l'adhérent (Fax)
    // --------------------------------------------------------------------
    if (EnregDucsEdi.FaxET <> '') then
      Writeln(FDUCSEDI,'COMDTF'+
                       EnregDucsEdi.FaxET+
                       Copy(blanc,1,(70-length(EnregDucsEdi.FaxET)))+ // PT30 DUCS EDI V4.2
                     'FX');
// PT30 DUCS EDI V4.2
    // COMDTM - Mode de communication avec le contact chez l'adhérent (Email)
    // --------------------------------------------------------------------
    if (EnregDucsEdi.EmailET <> '') then
      Writeln(FDUCSEDI,'COMDTM'+
                       EnregDucsEdi.EmailET+
                       Copy(blanc,1,(70-length(EnregDucsEdi.EmailET)))+
                     'EM');
// d PT30 DUCS EDI V4.2

   // RFFNUM - Identification adhérent chez l'institution (n° interne)
   // ----------------------------------------------------------------
   if (copy(EnregDucsEdi.TypDucs,1,1) = 'I') then
   Writeln(FDUCSEDI,'RFFNUM'+
                    EnregDucsEdi.NumeroInterne+
                    Copy(blanc,1,(35-length(EnregDucsEdi.NumeroInterne))));
   // RFFACT - Référence de l'activité pricipale de l'adhérent
   // --------------------------------------------------------
   if (copy(EnregDucsEdi.TypDucs,1,1) = 'I') then
   Writeln(FDUCSEDI,'RFFACT'+
                    EnregDucsEdi.ApePdu+
                    Copy(blanc,1,(35-length(EnregDucsEdi.ApePdu))));

   // RFFAMG - Acoss, Entreprise adhérent TDS (O/N)
   // --------------------------------------------------------
   if (copy(EnregDucsEdi.TypDucs,1,1) = 'A') and
      (copy(EnregDucsEdi.Abrege,3,2) = '00') then
     // TR Acoss
     begin
      if (GetParamSocSecur('SO_PGADHTDS',TRUE) = True)  then     
        // Adhérent TDS
        Writeln(FDUCSEDI,'RFFAMG1'+Copy(blanc,1,34))
      else
        // Non adhérent TDS
        Writeln(FDUCSEDI,'RFFAMG0'+Copy(blanc,1,34));
     end;

   // TIERS REDACTEUR
   // ---------------
   //  Ces segments sont alimentés par les  mêmes informations que les segments
   //  concernant l'émetteur du message. il ne seront générés que si le siret
   //  émetteur est différent du siret adhérent. (cela suppose que le rédacteur
   //  sera toujours l'émetteur)
   if (EnregDucsEdi.SiretPet <> EnregDucsEdi.Siret) then
   begin
     // NADAE  Nom & adresse du tiers rédacteur
     // ---------------------------------------
     Writeln(FDUCSEDI,'NADAE '+
                      EnregDucsEdi.SiretPet+
                      copy(blanc,1,(35-length(EnregDucsEdi.SiretPet)))+
                      EnregDucsEdi.Declarant+
                      copy(blanc,1,(35-length(EnregDucsEdi.Declarant)))+
                      EnregDucsEdi.AdressePet+
                      copy(blanc,1,(35-length(EnregDucsEdi.AdressePet)))+
                      EnregDucsEdi.Adresse2Pet+
                      copy(blanc,1,(35-length(EnregDucsEdi.Adresse2Pet)))+
                      EnregDucsEdi.VillePet+
                      copy(blanc,1,(35-length(EnregDucsEdi.VillePet)))+
                      EnregDucsEdi.CpPet);

      // CTAAE - Info sur le correspondant tiers rédacteur, personne à contacter
      // -----------------------------------------------------------------------
     Writeln(FDUCSEDI,'CTAAE '+
                      EnregDucsEdi.DeclarantSuite+
                      copy(blanc,1,(35-length(EnregDucsEdi.DeclarantSuite))));

      // COMAET - Mode de communication avec le contact (téléphone)
      // ----------------------------------------------------------
     if (EnregDucsEdi.TelephoneDecl <> '') then
        Writeln(FDUCSEDI,'COMAET'+
                         EnregDucsEdi.TelephoneDecl+
                         copy(blanc,1,(70-length(EnregDucsEdi.TelephoneDecl)))+ // PT30 DUCS EDI V4.2
                         'TE');
      // COMAEF - Mode de communication avec le contact (fax)
      // ----------------------------------------------------------
      if (EnregDucsEdi.FaxDeclarant <> '') then
         Writeln(FDUCSEDI,'COMAEF'+
                          EnregDucsEdi.FaxDeclarant+
                          copy(blanc,1,(70-length(EnregDucsEdi.FaxDeclarant)))+ // PT30 DUCS EDI V4.2
                          'FX');
// d PT30 DUCS EDI V4.2
      // COMAEM- Mode de communication avec le contact (Email)
      // ----------------------------------------------------------
      if (EnregDucsEdi.EmailDeclarant <> '') then
         Writeln(FDUCSEDI,'COMAEM'+
                          EnregDucsEdi.EmailDeclarant+
                          copy(blanc,1,(70-length(EnregDucsEdi.EmailDeclarant)))+
                          'EM');
// f PT30 DUCS EDI V4.2

      // RFFAE - Référence de l'émetteur chez le destinataire (IRC uniquement)
      // ---------------------------------------------------------------------
      if (copy(EnregDucsEdi.TypDucs,1,1) = 'I') then
         Writeln(FDUCSEDI,'RFFAE '+
                          EnregDucsEdi.NoContEmet+
                          copy(blanc,1,(35-length(EnregDucsEdi.NoContEmet)))+
                          'FX');
   end;

   // PAIEMENT GROUPE  (ACOSS et UNEDIC uniquement) PT30 DUCS EDI V4.2
   // ----------------------------------------------------------------
   // La présence des segments "paiement groupé" dépend de la valeur du champ
   // POG_PAIEGROUPE (=TRUE).
   // Les segments "paiement groupé" concernent l'établissement principal
   // responsable du paiement groupé. Correspond à une ducs dossier
   if ((copy(EnregDucsEdi.TypDucs,1,1) = 'U') or
       (copy(EnregDucsEdi.TypDucs,1,1) = 'A')) and
      (EnregDucsEdi.PaiementGroupe = true) then
   begin
     // NADFZ - Nom & adresses établissemnt principal du paiement groupé
     // ----------------------------------------------------------------
     Writeln(FDUCSEDI,'NADFZ '+
                          EnregDucsEdi.Siret+
                          copy(blanc,1,(35-length(EnregDucsEdi.Siret)))+
                          EnregDucsEdi.LibelleET+
                          copy (blanc,1,(35-length(EnregDucsEdi.LibelleET)))+
                          EnregDucsEdi.Adresse1ET+
                          copy (blanc,1,(35-length(EnregDucsEdi.Adresse1ET)))+
                          EnregDucsEdi.Adresse2ET+
                          copy (blanc,1,(35-length(EnregDucsEdi.Adresse2ET)))+
                          EnregDucsEdi.VilleET+
                          copy (blanc,1,(35-length(EnregDucsEdi.VilleET)))+
                          EnregDucsEdi.CpET);
     // CTAFZ - Info sur le correspondant de paiement groupé, personne à contacter
     // --------------------------------------------------------------------------
     Writeln(FDUCSEDI,'CTAFZ '+
                      EnregDucsEdi.AdherContact+
                      copy(blanc,1,(35-length(EnregDucsEdi.AdherContact))));
     // COMFZT - Mode de communication avec la personne du paiement groupé à contacter (Tel)
     // ------------------------------------------------------------------------------------
     if (EnregDucsEdi.TelET <> '') then
       Writeln(FDUCSEDI,'COMFZT'+
                        EnregDucsEDi.TelET+
                        copy(blanc,1,(70-length(EnregDucsEdi.TelET)))+ // PT30 DUCS EDI V4.2
                        'TE');
     // COMFZF - Mode de communication avec la personne du paiement groupé à contacter (Fax)
     // ------------------------------------------------------------------------------------
      if (EnregDucsEdi.FaxET <> '') then
        Writeln(FDUCSEDI,'COMFZF'+
                        EnregDucsEDi.FaxET+
                        copy(blanc,1,(70-length(EnregDucsEdi.FaxET)))+ // PT30 DUCS EDI V4.2
                        'FX');
// d PT30 DUCS EDI V4.2
     // COMFZM - Mode de communication avec la personne du paiement groupé à contacter (Email)
     // ------------------------------------------------------------------------------------
      if (EnregDucsEdi.EmailET <> '') then
        Writeln(FDUCSEDI,'COMFZM'+
                        EnregDucsEDi.EmailET+
                        copy(blanc,1,(70-length(EnregDucsEdi.EmailET)))+
                        'EM');
// f PT30 DUCS EDI V4.2

   end;
   // PEDp00 - Période de la déclaration
   // ----------------------------------
   tampon :='';
   if (copy(EnregDucsEdi.Abrege,4,1) <> '0') then
     // mensuelle
     begin
     tampon := 'M';
     EnregDucsEdi.Periode := 'M';
     end
   else
     if (copy(EnregDucsEdi.Abrege,3,1) <> '0') then
       // Trimestrielle
       begin
       tampon := 'T';
       EnregDucsEdi.Periode := 'T';
       end
     else
       begin
       tampon :='A';
       EnregDucsEdi.Periode := 'A';
       end;
   if (Copy(EnregDucsEdi.Typducs,1,1) = 'A') then
     // ACOSS
     begin
      if (tampon <> 'A') then
        // Mensuelle ou trimestrielle   (PEDp00+abrégé période)
        Writeln(FDUCSEDI,'PED'+
                         tampon+
                         '00'+
                         EnregDucsEdi.Abrege+
                         copy(blanc,1,(17-(length(EnregDucsEdi.Abrege)))))
      else
        // Annuelle   (PEDp00+année+62)
// d PT30 DUCS EDI V4.2
{PT37  il n' y pas lieu de modifier l'annee de EnregDucsEdi.Abrege

        if (VH_Paie.PGDecalage = true) or (VH_Paie.PGDecalagePetit = true) then
          Writeln(FDUCSEDI,'PED'+
                           'A'+
                           '00'+
                           ColleZeroDevant((StrToInt(copy(EnregDucsEdi.Abrege,1,2))-1),2)+
                           '62'+
                            copy(blanc,1,(17-(length(EnregDucsEdi.Abrege)))))
         else
}
// f PT30 DUCS EDI V4.2
           Writeln(FDUCSEDI,'PED'+
                          'A'+
                          '00'+
                          copy(EnregDucsEdi.Abrege,1,2)+
                          '62'+
                          copy(blanc,1,(17-(length(EnregDucsEdi.Abrege)))));
     end
   else
   // IRC ou UNEDIC
   begin
     periode := '';
     periode :=  copy(DateToStr(EnregDucsEdi.DateFin),7,2);

     PeriodRef := EnregDucsEdi.Abrege;     //  PT35

     if (tampon =  'A') then
     begin
         // Annuelle  (PEDp00+ssaa)
// d PT30 DUCS EDI V4.2
{        if (VH_Paie.PGDecalage = true) or (VH_Paie.PGDecalagePetit = true) then
          periode := periode+
                     ColleZeroDevant((StrToInt(copy(EnregDucsEdi.Abrege,1,2))-1),2)
        else}
// f PT30 DUCS EDI V4.2
{* PT49 correction }
         periode := periode+copy(EnregDucsEdi.Abrege,1,2);
{        PeriodRef := periode;  // PT51 }
     end;
     if (tampon = 'T') then
     begin
         // Trimestrielle  (PEDp00+ssaat)
{* PT49 correction  }
         periode := periode+copy(EnregDucsEdi.Abrege,1,3); 
{         PeriodRef := periode;    //PT51  }

     end;
     if (tampon = 'M') then
     begin
         // Mensuelle   (PEDp00+ssaamm)
       if (((VH_Paie.PGDecalage = true) or (VH_Paie.PGDecalagePetit = true)) and
         (Copy(EnregDucsEdi.Typducs,1,1) = 'U')) then
       begin
         if (copy(EnregDucsEdi.Abrege,3,2) = '11') then
         begin
           periode := periode+
                      format('%2.2d',[StrToInt(copy(EnregDucsEdi.Abrege,1,2))-1])+'12';
           PeriodRef := format('%2.2d',[StrToInt(copy(EnregDucsEdi.Abrege,1,2))-1])+'43';
         end
         else
         begin
           periode := periode+
                      copy(EnregDucsEdi.Abrege,1,2)+
                      format('%2.2d',[((StrToint(copy(EnregDucsEdi.Abrege,3,1))-1)*3)+
                      StrToInt(copy(EnregDucsEdi.Abrege,4,1))-1]);

           PeriodRef :=ColleZeroDevant(StrToInt(EnregDucsEdi.Abrege)-1,4);
           if (copy(PeriodRef,4,1) = '0') then
             PeriodRef :=ColleZeroDevant(StrToInt(EnregDucsEdi.Abrege)-8,4);

         end;
       end
       else
         periode := periode+
                    copy(EnregDucsEdi.Abrege,1,2)+
                    format('%2.2d',[((StrToint(copy(EnregDucsEdi.Abrege,3,1))-1)*3)+
                    StrToInt(copy(EnregDucsEdi.Abrege,4,1))]);
        end;
//PT35        PeriodRef := EnregDucsEdi.Abrege;     //  PT29
     Writeln(FDUCSEDI,'PED'+
                       tampon+
                       '00'+
                       periode+
                       copy(blanc,1,(17-(length(periode)))));
   end;
   // PED 01 - Période de versement des salaires pour paie décalée (IRC uniquement)
   // ----------------------------------------------------------------------------
{* PT30 DUCS EDI V4.2 plus utile
   if ((VH_Paie.PGDecalage = true) or (VH_Paie.PGDecalagePetit = true)) and
      (copy(EnregDucsEdi.TypDucs,1,1) = 'I') then
        Writeln(FDUCSEDI,'PED '+
                         '01'+
                         copy(DateToStr(EnregDucsEdi.DateDebut),7,4)+
                         copy(DateToStr(EnregDucsEDi.DateDebut),1,2)+
                         copy(DateToStr(EnregDucsEDi.DateDebut),4,2)+
                         '-'+
                         copy(DateToStr(EnregDucsEdi.DateFin),7,4)+
                         copy(DateToStr(EnregDucsEDi.DateFin),1,2)+
                         copy(DateToStr(EnregDucsEDi.DateFin),4,2));*}
                         
   // QTY - Effectif au 31/12, déclarations annuelles (Acoss et Unedic uniquement)
   // ----------------------------------------------------------------------------
   tampon := '';
   if (copy(EnregDucsEdi.Abrege,3,2) ='00') and
      (copy(EnregDucsEdi.TypDucs,1,1) <> 'I') then
     begin
      if (copy(EnregDucsEdi.TypDucs,1,1) = 'A') then
        // Annuel ACOSS (TR)
//PT40         tampon := FloatToStr((EnregDucsEdi.TotHommes+EnregDucsEdi.TotFemmes));
         tampon := FloatToStr((EnregDucsEdi.TotHommes+
                               EnregDucsEdi.TotFemmes-
                               EnregDucsEdi.TotApprenti));
       if (copy(EnregDucsEdi.TypDucs,1,1) = 'U') then
        // Annuel UNEDIC
        tampon :=  FloatToStr(EnregDucsEdi.TotHommes+EnregDucsEdi.TotFemmes);

      Writeln(FDUCSEDI,'QTY   '+
                       tampon+
                       copy(blanc,1,(18-(length(tampon)))));
     end;
// d PT30 DUCS EDI V4.2
// RFFZST ne fait plus partie du fichier
   // RFFZST - Référence du statut juridique de l'entreprise
   //-------------------------------------------------------
   //          Déclarations annuelles (UNEDIC seulement)
{*   if (copy(EnregDucsEdi.Abrege,3,2) ='00') and
      (copy(EnregDucsEdi.TypDucs,1,1) = 'U') then
      Writeln(FDUCSEDI,'RFFZST'+
                       EnregDucsEdi.Juridique+
                       copy(blanc,1,(35-(length(EnregDucsEdi.Juridique)))));*}

   // RFFAUM - Référence nom de l’éditeur du logiciel de paie
   //-------------------------------------------------------

      Writeln(FDUCSEDI,'REFAUM'+
                       Apalatys+
                       copy(blanc,1,(20-(length(Apalatys)))));
   // RFFAUN - Référence nom du logiciel de paie et version
   //-------------------------------------------------------

{PT34      Writeln(FDUCSEDI,'REFAUN'+
                       'CEGID PGI'+
                       copy(blanc,1,(20-9))+
                       copy(V_PGI.NumVersion,1,10));}
      Writeln(FDUCSEDI,'REFAUN'+
                       'CEGID PAIE'+
                       copy(blanc,1,(20-10))+
                       copy(V_PGI.NumVersion,1,10));


end;
// fin Groupe1Declaration

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 17/04/2002
Modifié le ... :   /  /
Description .. : Construction des segments du groupe 2 de la déclaration
Mots clefs ... : PAIE, PGDUCSEDI
*****************************************************************}
procedure Groupe2Declaration (EnregDucsEdi : TDucsEdi; EnrDucsEdiGroupee : TDucsEdiGroupee; Trait : string);

var
   blanc, tampon              : string;
   typenr                     : string;
   WMontant                   : double; // PT30 DUCS EDI V4.2
begin
     blanc := StringOfChar(' ',70);    // PT30 DUCS EDI V4.2
     // NADDC - Etablissement déclaré Nom & adresse
     // -------------------------------------------
     if (Trait <> 'G') then
      Writeln(FDUCSEDI,'NADDC '+
                       EnregDucsEdi.Siret+
                       copy(blanc,1,(35-(length(EnregDucsEdi.Siret))))+
                       EnregDucsEdi.LibelleET+
                       copy(blanc,1,(35-(length(EnregDucsEdi.LibelleET))))+
                       EnregDucsEdi.Adresse1ET+
                       copy(blanc,1,(35-(length(EnregDucsEdi.Adresse1ET))))+
                       EnregDucsEdi.Adresse2ET+
                       copy(blanc,1,(35-(length(EnregDucsEdi.Adresse2ET))))+
                       EnregDucsEdi.VilleET+
                       copy(blanc,1,(35-(length(EnregDucsEdi.VilleET))))+
                       EnregDucsEdi.CpET)
      else
       begin
        typenr := 'NADDC ';
        if (copy(EnregDucsEdi.TypDucs,1,1) = 'U') then typenr := 'NADDCE';
        Writeln(FDUCSEDI,typenr+
                        EnrDucsEdiGroupee.Siret+
                        copy(blanc,1,(35-(length(EnrDucsEdiGroupee.Siret))))+
                        EnrDucsEdiGroupee.LibelleET+
                        copy(blanc,1,(35-(length(EnrDucsEdiGroupee.LibelleET))))+
                        EnrDucsEdiGroupee.Adresse1ET+
                        copy(blanc,1,(35-(length(EnrDucsEdiGroupee.Adresse1ET))))+
                        EnrDucsEdiGroupee.Adresse2ET+
                        copy(blanc,1,(35-(length(EnrDucsEdiGroupee.Adresse2ET))))+
                        EnrDucsEdiGroupee.VilleET+
                        copy(blanc,1,(35-(length(EnrDucsEdiGroupee.VilleET))))+
                        EnrDucsEdiGroupee.CpET);

       end;

    // CTADC - Nom du contact celui de  l'adhérent - IRC seulement
    // -----------------------------------------------------------
//    if (Trait <> 'G') then (info. idem ts etabs.)
//    begin
      if (copy(EnregDucsEdi.TypDucs,1,1) = 'I') then
      begin
        Writeln(FDUCSEDI,'CTADC '+
                         EnregDucsEdi.AdherContact+
                         Copy(blanc,1,(35-length(EnregDucsEdi.AdherContact))));
      end;
//    end;

//    if (Trait <> 'G') then  (info. idem ts etabs.)
//      begin
        if (copy(EnregDucsEdi.TypDucs,1,1) = 'I') then
          begin
          // COMDC - Mode de communication avec le contact (Tel) - IRC seulement
          // -------------------------------------------------------------------
          if (EnregDucsEdi.TelET <> '') then
            Writeln (FDUCSEDI,'COMDCT'+
                              EnregDucsEdi.TelET+
                              copy(blanc,1,(70-length(EnregDucsEdi.TelET)))+ // PT30 DUCS EDI V4.2
                              'TE');
          // COMDCF - Mode de communication avec le contact (Fax) - IRC seulement
          // --------------------------------------------------------------------
          if (EnregDucsEdi.FaxET <> '') then
            Writeln (FDUCSEDI,'COMDCF'+
                              EnregDucsEdi.FaxET+
                              copy(blanc,1,(70-length(EnregDucsEdi.FaxET)))+ // PT30 DUCS EDI V4.2
                              'FX');
// PT30 DUCS EDI V4.2
          // COMDCM - Mode de communication avec le contact (Email) - IRC seulement
          // --------------------------------------------------------------------
{* PT42
         if (EnregDucsEdi.EmailET <> '') then
            Writeln (FDUCSEDI,'COMDCM'+
                              EnregDucsEdi.EmailET+
                              copy(blanc,1,(70-length(EnregDucsEdi.EmailET)))+
                              'EM');*}
// f PT30 DUCS EDI V4.2
          end;
//      end;

// d dPT30 DUCS EDI V4.2
    // PEV   - Date de versement des salaires
    // --------------------------------------
    // A renseigner pour
    // - ACOSS si BRC (913 ou 914)
    // - UNEDIC si DRA (922 ou 923) et etab principal uniquement
    if (Trait <> 'G') then
    begin
      if ((copy(EnregDucsEdi.TypDucs,1,1) = 'A') or
          (copy(EnregDucsEdi.TypDucs,1,1) = 'U')) and
         (EnregDucsEdi.Paiement <> IDate1900) then
      begin
        if (EnregDucsEdi.TypBordereau = '913') or
           (EnregDucsEdi.TypBordereau = '914') or
           (EnregDucsEdi.TypBordereau = '922') or
           (EnregDucsEdi.TypBordereau = '923')then
        begin
            Writeln(FDUCSEDI,'PEV   '+
                             Copy(DateToStr(EnregDucsEdi.Paiement),7,4)+
                             Copy(DateToStr(EnregDucsEdi.Paiement),4,2)+
                             Copy(DateToStr(EnregDucsEdi.Paiement),1,2)+
                             Copy(blanc,1,9));
        end;
      end;
    end
    else
    begin
      if ((copy(EnregDucsEdi.TypDucs,1,1) = 'A') and
         (EnrDucsEdiGroupee.Paiement <> IDate1900))   then
      begin
        if (EnregDucsEdi.TypBordereau = '913') or
           (EnregDucsEdi.TypBordereau = '914') then
        begin
            Writeln(FDUCSEDI,'PEV   '+
                             Copy(DateToStr(EnrDucsEdiGroupee.Paiement),7,4)+
                             Copy(DateToStr(EnrDucsEdiGroupee.Paiement),4,2)+
                             Copy(DateToStr(EnrDucsEdiGroupee.Paiement),1,2)+
                             Copy(blanc,1,9));

        end;
      end;
    end;


    // PEM   - Quantième habituel de versement des salaires
    //         Jour fin de mois
    // ----------------------------------------------------
    // A renseigner pour
    // - UNEDIC si DRA (922 ou 923) ou AVA (924 ou 926
    //   et etab principal uniquement
    if ((copy(EnregDucsEdi.TypDucs,1,1) = 'U') and (Trait <> 'G') and
        ((VH_Paie.PGDecalage = false) and (VH_Paie.PGDecalagePetit = false)))then
    begin
      if (EnregDucsEdi.TypBordereau = '922') or
         (EnregDucsEdi.TypBordereau = '923') or
         (EnregDucsEdi.TypBordereau = '924') or
         (EnregDucsEdi.TypBordereau = '926')then
      begin
          Writeln(FDUCSEDI,'PEM   '+
                           Copy(DateToStr(EnregDucsEdi.Paiement),1,2)+
                           Copy(blanc,1,9));
      end;
    end;
    // PES   - Quantième habituel de versement des salaires
    //         Jour mois suivant
    // ----------------------------------------------------
    // A renseigner pour
    // - UNEDIC si DRA (922 ou 923) ou AVA (924 ou 926
    //   et etab principal uniquement
    if ((copy(EnregDucsEdi.TypDucs,1,1) = 'U') and (Trait <> 'G') and
        ((VH_Paie.PGDecalage = true) or (VH_Paie.PGDecalagePetit = true)))then
    begin
      if (EnregDucsEdi.TypBordereau = '922') or
         (EnregDucsEdi.TypBordereau = '923') or
         (EnregDucsEdi.TypBordereau = '924') or
         (EnregDucsEdi.TypBordereau = '926')then
      begin
          Writeln(FDUCSEDI,'PES   '+
                           Copy(DateToStr(EnregDucsEdi.Paiement),1,2)+
                           Copy(blanc,1,9));
      end;
    end;

// f PT30 DUCS EDI V4.2
    // MOADUE - Montant total déclaré (URSSAF) ou *
    //          mt total des cotisations (UNEDIC, IRC)
    //  pour TR URSSAF (915 ou 916) tient compte de la régul
    // -------------------------------------------------------------------
    if (Trait <> 'G') then
    begin
// d PT30 DUCS EDI V4.2
      WMontant := EnregDucsEdi.MtDeclare;
      if (EnregDucsEdi.TypBordereau = '915') or
         (EnregDucsEdi.TypBordereau = '916') then
        WMontant := WMontant + EnregDucsEdi.Regularisation;

//        if (EnregDucsEdi.MtDeclare <> 0.0) or
//           ((EnregDucsEdi.MtDeclare = 0.0) and
{PT32      if (WMontant <> 0.0) or
         ((WMontant = 0.0) and (copy(EnregDucsEdi.TypDucs,1,1) <> 'A')) then
      begin}
        if (copy(EnregDucsEdi.TypDucs,1,1)<>'I') then
        // Acoss et Unedic mt à l'euro
        begin
//               tampon :=FormatFloat('#0',EnregDucsEdi.MtDeclare);
          tampon :=FormatFloat('#0',WMontant);
        end
        else
        // IRC mt au centime d'euro
        begin
//               tampon := FormatFloat('#0',EnregDucsEdi.MtDeclare*100);
          tampon := FormatFloat('#0',WMontant*100);
// f PT30 DUCS EDI V4.2
        end;
           Writeln(FDUCSEDI,'MOADUE'+
                            tampon+
                            copy(blanc,1,(18-length(tampon)))+
                            EnregDucsEdi.MonnaieTenue);

// d PT30 DUCS EDI V4.2
      Writeln(FRapport, '  Etablissement : '+ EnregDucsEdi.Etab+' '+EnregDucsEdi.LibelleET);

          if (copy(EnregDucsEdi.TypDucs,1,1) = 'A') then
            Writeln(FRapport,'    Montant total déclaré : '+
                              tampon+' '+
                              EnregDucsEdi.MonnaieTenue)
          else
          begin
            if (copy(EnregDucsEdi.TypDucs,1,1) = 'I') then
          tampon :=FormatFloat('#0.00',WMontant); // édition à l'euro
            Writeln(FRapport,'    Montant total des cotisations : '+
                              tampon+' '+
                              EnregDucsEdi.MonnaieTenue);
        Writeln(FRapport,'');
          end ;
// f PT30 DUCS EDI V4.2
{PT32      end;}
// d PT30 DUCS EDI V4.2
{      if (copy(EnregDucsEdi.TypDucs,1,1) = 'A') then
      begin
        if (WMontant = 0.0) then
        begin
          tampon := '0,00';
          Writeln(FRapport, EnregDucsEdi.Etab+' '+EnregDucsEdi.LibelleET);
          Writeln(FRapport,'    Montant total déclaré : '+
                              tampon+' '+
                              EnregDucsEdi.MonnaieTenue);
        end;
      end;}
// f PT30 DUCS EDI V4.2
    end
    else
    begin
// d PT30 DUCS EDI V4.2
      WMontant := EnrDucsEdiGroupee.MtDeclare;
      if (EnregDucsEdi.TypBordereau = '915') or
         (EnregDucsEdi.TypBordereau = '916') then
        WMontant := WMontant + EnrDucsEdiGroupee.Regularisation;

//        if (EnrDucsEdiGroupee.MtDeclare <> 0.0) then
//      if (WMontant <> 0.0) then
//      begin
        if (copy(EnregDucsEdi.TypDucs,1,1)<>'I') then
        // Acoss et Unedic mt à l'euro
        begin
//             tampon :=FormatFloat('#0',EnrDucsEdiGroupee.MtDeclare);
          tampon :=FormatFloat('#0',WMontant);
        end
        else
        // IRC mt au centime d'euro
        begin
//          tampon := FormatFloat('#0',EnrDucsEdiGroupee.MtDeclare*100);
          tampon := FormatFloat('#0',WMontant*100);
// f PT30 DUCS EDI V4.2
        end;
        if (Copy(EnregDucsEdi.TypDucs,1,1) <> 'U') then  // ducs edi v4.2 pour unedic uniqt etab principal
          Writeln(FDUCSEDI,'MOADUE'+
                         tampon+
                         copy(blanc,1,(18-length(tampon)))+
                         EnrDucsEdiGroupee.MonnaieTenue);
// d PT30 DUCS EDI V4.2
        Writeln(FRapport,'  Etablissement : '+ EnrDucsEdiGroupee.Etab+' '+EnrDucsEdiGroupee.LibelleET);

        if (copy(EnregDucsEdi.TypDucs,1,1) = 'A') then
          Writeln(FRapport,'    Montant total déclaré : '+
                           tampon+' '+
                           EnrDucsEdiGroupee.MonnaieTenue)
        else
        begin
          if (copy(EnregDucsEdi.TypDucs,1,1) = 'I') then
              tampon :=FormatFloat('#0.00',WMontant); // édition à l'euro

          Writeln(FRapport,'    Montant total des cotisations : '+
                           tampon+' '+
                           EnrDucsEdiGroupee.MonnaieTenue);
          Writeln(FRapport,'');
        end;
// f PT30 DUCS EDI V4.2
//      end;
{      if (WMontant = 0.0) then
      begin
        tampon := '0';
        Writeln(FRapport,'  Etablissement : '+  EnrDucsEdiGroupee.Etab+' '+EnrDucsEdiGroupee.LibelleET);
        Writeln(FRapport,'    Montant total déclaré : '+
                         tampon+' '+
                         EnrDucsEdiGroupee.MonnaieTenue);
      end;}
    end;

// d PT30 DUCS EDI V4.2
    // MOAMSA - Masses salariales annuelles
    //  pour DRA UNEDIC (922 ou 923)
    // ------------------------------------
    if ((copy(EnregDucsEdi.TypDucs,1,1) = 'U') and
        ((EnregDucsEdi.TypBordereau = '922') or
         (EnregDucsEdi.TypBordereau = '923'))) then
    //  Unedic mt à l'euro
    begin
      tampon :=FormatFloat('#0',EnregDucsEdi.MasseAnnuel);
      if (Trait <> 'G') or ((Trait = 'G') and (Copy(EnregDucsEdi.TypDucs,1,1) <> 'U')) then  // PT30 DUCS EDI V4.2
//    pour unedic uniqt etab principal
        Writeln(FDUCSEDI,'MOAMSA'+
                         tampon+
                         copy(blanc,1,(18-length(tampon)))+
                         EnregDucsEdi.MonnaieTenue);
    end;

    // MOAREM - Rémunération déclarées sur la Dads
    //  pour DRA et AVA UNEDIC (922 ou 923 ou 924 ou 926)
    // --------------------------------------------------
    if ((copy(EnregDucsEdi.TypDucs,1,1) = 'U') and
        ((EnregDucsEdi.TypBordereau = '922') or
         (EnregDucsEdi.TypBordereau = '923') or
         (EnregDucsEdi.TypBordereau = '924') or
         (EnregDucsEdi.TypBordereau = '926'))) then
    //  Unedic mt à l'euro
    begin
      tampon :=FormatFloat('#0',EnregDucsEdi.RemunDads);
      if (Trait <> 'G') or ((Trait = 'G') and (Copy(EnregDucsEdi.TypDucs,1,1) <> 'U')) then  // PT30 DUCS EDI V4.2
//    pour unedic uniqt etab principal
        Writeln(FDUCSEDI,'MOAREM'+
                         tampon+
                         copy(blanc,1,(18-length(tampon)))+
                         EnregDucsEdi.MonnaieTenue);
    end;
// f PT30 DUCS EDI V4.2

    // MOATRA - Montant total des cotisations transport (ACOSS)
    // --------------------------------------------------------
    if (Trait <> 'G') then
     begin
      if (copy(EnregDucsEdi.TypDucs,1,1)='A') and
         (EnregDucsEdi.NbTransport > 1) then
         // Acoss  et plusieurs taux de cotisation transport
         begin
           tampon :=FormatFloat('#0',EnregDucsEdi.MtTransport);
           tampon := ColleZeroDevant(StrToInt(tampon),8);      // PT30 DUCS EDI V4.2
           Writeln(FDUCSEDI,'MOATRA'+
                            tampon+
                            copy(blanc,1,(18-length(tampon)))+
                            EnregDucsEdi.MonnaieTenue);
         end;
      end
    else
     begin
      if (copy(EnregDucsEdi.TypDucs,1,1)='A') and
         (EnrDucsEdiGroupee.NbTransport > 1) then
         // Acoss  et plusieurs taux de cotisation transport
         begin
           tampon :=FormatFloat('#0',EnrDucsEdiGroupee.MtTransport);
           tampon := ColleZeroDevant(StrToInt(tampon),8);      // PT30 DUCS EDI V4.2
           Writeln(FDUCSEDI,'MOATRA'+
                            tampon+
                            copy(blanc,1,(18-length(tampon)))+
                            EnrDucsEdiGroupee.MonnaieTenue);
         end;
     end;
// d PT30 DUCS EDI V4.2

    // MOATAT - Montant total des Taxes additionnelles de transport (ACOSS)
    // ---------------------------------------------------------------------
    if (Trait <> 'G') then
     begin
      if (copy(EnregDucsEdi.TypDucs,1,1)='A') and
         (EnregDucsEdi.NbTAT > 1) then
         // Acoss  et plusieurs taux de cotisation transport
         begin
           tampon :=FormatFloat('#0',EnregDucsEdi.MtTAT);
           tampon := ColleZeroDevant(StrToInt(tampon),8);  // PT30 DUCS EDI V4.2
           Writeln(FDUCSEDI,'MOATAT'+
                            tampon+
                            copy(blanc,1,(18-length(tampon)))+
                            EnregDucsEdi.MonnaieTenue);
         end;
      end
    else
     begin
      if (copy(EnregDucsEdi.TypDucs,1,1)='A') and
         (EnrDucsEdiGroupee.NbTAT > 1) then
         // Acoss  et plusieurs taux de cotisation transport
         begin
           tampon :=FormatFloat('#0',EnrDucsEdiGroupee.MtTAT);
           tampon := ColleZeroDevant(StrToInt(tampon),8);  // PT30 DUCS EDI V4.2
           Writeln(FDUCSEDI,'MOATAT'+
                            tampon+
                            copy(blanc,1,(18-length(tampon)))+
                            EnrDucsEdiGroupee.MonnaieTenue);
         end;
     end;
// f ducs edi V4.2

    // MOAREG - Montant de la régularisation (ACOSS)
    // POur TR URSSAF (type boredreau 915 ou 916)
    // ------------------------------------------------------
    if (copy(EnregDucsEdi.TypDucs,1,1)='A') and
       ((EnregDucsEdi.TypBordereau = '915') or
        (EnregDucsEdi.TypBordereau = '916'))  then         // PT30 DUCS EDI V4.2
      // Acoss
      if (Trait <> 'G') then
      begin
// d PT30 DUCS EDI V4.2
        if (EnregDucsEdi.Regularisation<> 0.0) then
        begin
// f PT30 DUCS EDI V4.2
          tampon :=FormatFloat('#0',EnregDucsEdi.Regularisation);
          Writeln(FDUCSEDI,'MOAREG'+
                           tampon+
                           copy(blanc,1,(18-length(tampon)))+
                           EnregDucsEdi.MonnaieTenue);
// PT30 DUCS EDI V4.2
// if (EnregDucsEdi.Regularisation <> 0.0) then
{PT37          Writeln(FRapport,'    Montant de la régularisation : '+
                           Formatfloat('#0',EnregDucsEdi.Regularisation)+' '+
                           EnregDucsEdi.MonnaieTenue);}
        end;  // PT30 DUCS EDI V4.2
      end
      else
      begin
// d PT30 DUCS EDI V4.2
        if (EnrDucsEdiGroupee.Regularisation<> 0.0) then
        begin
// f PT30 DUCS EDI V4.2

          tampon :=FormatFloat('#0',EnrDucsEdiGroupee.Regularisation);
          Writeln(FDUCSEDI,'MOAREG'+
                           tampon+
                           copy(blanc,1,(18-length(tampon)))+
                           EnrDucsEdiGroupee.MonnaieTenue);
// PT30 DUCS EDI V4.2
//        if (EnrDucsEdiGroupee.Regularisation <> 0.0) then
{PT37          Writeln(FRapport,'    Montant de la régularisation : '+
                             Formatfloat('#0',EnrDucsEdiGroupee.Regularisation)+' '+
                             EnrDucsEdiGroupee.MonnaieTenue);}
        end;  // PT30 DUCS EDI V4.2
      end;

    // QTYD  - Effectif des salariés de l'établissement rémunérés pour la période (ACOSS)
    // ----------------------------------------------------------------------------------
    if (copy(EnregDucsEdi.TypDucs,1,1)='A') then
      // Acoss
      begin
        if (EnregDucsEdi.Neant = False ) then
         begin
           if (Trait <>'G') then
             tampon :=FormatFloat('#0',EnregDucsEdi.NbSalFpe)
           else
             tampon :=FormatFloat('#0',EnrDucsEdiGroupee.NbSalFpe);
         end
        else
         // Ducs Néant et Acoss - Effectif à zéro
         tampon :=FormatFloat('#0',0);

        Writeln(FDUCSEDI,'QTYD  '+
                          tampon+
                          copy(blanc,1,(18-length(tampon))));
     end;

    // QTYT  - Effectif des salariés inscrits dans l'établissement au dernier
    // jour de la période (ACOSS)
    // ----------------------------------------------------------------------
    if (copy(EnregDucsEdi.TypDucs,1,1) = 'A') then
    begin
      typenr:='QTYT  ';
      if (EnregDucsEdi.Neant = False ) then
      begin
        if (Trait <> 'G') then
          tampon :=FormatFloat('#0',
//PT40                  (EnregDucsEdi.TotHommes+EnregDucsEdi.TotFemmes))
                  (EnregDucsEdi.TotHommes+
                   EnregDucsEdi.TotFemmes-
                   EnregDucsEdi.TotApprenti))
        else
           tampon :=FormatFloat('#0',
//PT40                   (EnrDucsEdiGroupee.TotHommes+EnrDucsEdiGroupee.TotFemmes));
                   (EnrDucsEdiGroupee.TotHommes+
                    EnrDucsEdiGroupee.TotFemmes-
                    EnrDucsEdiGroupee.TotApprenti));
      end
      else
        // Ducs Néant - Effectif à zéro
        tampon :=FormatFloat('#0',0);

      Writeln(FDUCSEDI,typenr+
                       tampon+
                       copy(blanc,1,(18-length(tampon))));
    end;
    // QTYTH  - Effectif des salariés Hommes inscrits dans l'établissement au
    //          dernier jour de la période (UNEDIC)
    // ADV : type bordereau 920 ou 921
    // -------------------------------------------------------------------------
// d PT30 DUCS EDI V4.2
    if (copy(EnregDucsEdi.TypDucs,1,1)='U') and
       ((EnregDucsEdi.TypBordereau = '920') or
        (EnregDucsEdi.TypBordereau = '921'))then
// f PT30 DUCS EDI V4.2
    begin
      typenr:='QTYTH ';
      if (Trait <> 'G') then
        tampon :=FormatFloat('#0',EnregDucsEdi.TotHommes)
      else
        tampon :=FormatFloat('#0',EnrDucsEdiGroupee.TotHommes);
      if (Trait = 'G') then typenr := 'QTYTHE';
       Writeln(FDUCSEDI,typenr+
                        tampon+
                        copy(blanc,1,(18-length(tampon))));
    end;

    // QTYTF  - Effectif des salariés Femmes inscrits dans l'établissement au
    //          dernier jour de la période (UNEDIC)
    // ADV : type bordereau 920 ou 921
    // --------------------------------------------------------------------------------------------------------
// d PT30 DUCS EDI V4.2
    if (copy(EnregDucsEdi.TypDucs,1,1)='U') and
       ((EnregDucsEdi.TypBordereau = '920') or
        (EnregDucsEdi.TypBordereau = '921'))then
// f PT30 DUCS EDI V4.2
    begin
      typenr:='QTYTF ';
      if (Trait <> 'G') then
        tampon :=FormatFloat('#0',EnregDucsEdi.TotFemmes)
      else
        tampon :=FormatFloat('#0',EnrDucsEdiGroupee.TotFemmes);
      if (Trait = 'G') then typenr := 'QTYTFE';
        Writeln(FDUCSEDI,typenr+
                         tampon+
                         copy(blanc,1,(18-length(tampon))));
    end;

// d PT30 DUCS EDI V4.2
    // QTYT2  - Effectif des salariés hommes en CDI inscrits dans
    //          l'établissement au dernier jour de la période (UNEDIC)
    // DRA et AVA : type bordereau 922 ou 923 ou 924
    // -------------------------------------------------------------------------

    if (copy(EnregDucsEdi.TypDucs,1,1)='U') and
       ((EnregDucsEdi.TypBordereau = '922') or
        (EnregDucsEdi.TypBordereau = '923') or
        (EnregDucsEdi.TypBordereau = '924' )) then
    begin
      typenr:='QTYT2 ';
      if (Trait <> 'G') then
        tampon :=FormatFloat('#0',EnregDucsEdi.CDIH)
      else
        tampon :=FormatFloat('#0',EnrDucsEdiGroupee.CDIH);
      if (Trait = 'G') then typenr := 'QTYT2E';
        Writeln(FDUCSEDI,typenr+
                         tampon+
                         copy(blanc,1,(18-length(tampon))));
    end;
    // QTYT3  - Effectif des salariés femmes en CDI inscrits dans
    //          l'établissement au dernier jour de la période (UNEDIC)
    // DRA et AVA : type bordereau 922 ou 923 ou 924
    // -------------------------------------------------------------------------

    if (copy(EnregDucsEdi.TypDucs,1,1)='U') and
       ((EnregDucsEdi.TypBordereau = '922') or
        (EnregDucsEdi.TypBordereau = '923') or
        (EnregDucsEdi.TypBordereau = '924' )) then
    begin
      typenr:='QTYT3 ';
      if (Trait <> 'G') then
        tampon :=FormatFloat('#0',EnregDucsEdi.CDIF)
      else
        tampon :=FormatFloat('#0',EnrDucsEdiGroupee.CDIF);
      if (Trait = 'G') then typenr := 'QTYT3E';
        Writeln(FDUCSEDI,typenr+
                         tampon+
                         copy(blanc,1,(18-length(tampon))));
    end;
    // QTYT4  - Effectif des salariés hommes  apprentis inscrits dans
    //          l'établissement au dernier jour de la période (UNEDIC)
    // DRA et AVA : type bordereau 922 ou 923 ou 924 ou 926
    // -------------------------------------------------------------------------

    if (copy(EnregDucsEdi.TypDucs,1,1)='U') and
       ((EnregDucsEdi.TypBordereau = '922') or
        (EnregDucsEdi.TypBordereau = '923') or
        (EnregDucsEdi.TypBordereau = '924' ) or
        (EnregDucsEdi.TypBordereau = '926' )) then
    begin
      typenr:='QTYT4 ';
      if (Trait <> 'G') then
        tampon :=FormatFloat('#0',EnregDucsEdi.ApprentiH)
      else
        tampon :=FormatFloat('#0',EnrDucsEdiGroupee.ApprentiH);
      if (Trait = 'G') then typenr := 'QTYT4E';
        Writeln(FDUCSEDI,typenr+
                         tampon+
                         copy(blanc,1,(18-length(tampon))));
    end;
    // QTYT5  - Effectif des salariés femmes apprentis inscrits dans
    //          l'établissement au dernier jour de la période (UNEDIC)
    // DRA et AVA : type bordereau 922 ou 923 ou 924 ou 926
    // -------------------------------------------------------------------------

    if (copy(EnregDucsEdi.TypDucs,1,1)='U') and
       ((EnregDucsEdi.TypBordereau = '922') or
        (EnregDucsEdi.TypBordereau = '923') or
        (EnregDucsEdi.TypBordereau = '924' ) or
        (EnregDucsEdi.TypBordereau = '926' )) then
    begin
      typenr:='QTYT5 ';
      if (Trait <> 'G') then
        tampon :=FormatFloat('#0',EnregDucsEdi.ApprentiF)
      else
        tampon :=FormatFloat('#0',EnrDucsEdiGroupee.ApprentiF);
      if (Trait = 'G') then typenr := 'QTYT5E';
        Writeln(FDUCSEDI,typenr+
                         tampon+
                         copy(blanc,1,(18-length(tampon))));
    end;
    // QTYT6  - Effectif des salariés hommes de + de 65 ans inscrits dans
    //          l'établissement au dernier jour de la période (UNEDIC)
    // DRA et AVA : type bordereau 922 ou 923 ou 924
    // -------------------------------------------------------------------------

    if (copy(EnregDucsEdi.TypDucs,1,1)='U') and
       ((EnregDucsEdi.TypBordereau = '922') or
        (EnregDucsEdi.TypBordereau = '923') or
        (EnregDucsEdi.TypBordereau = '924' )) then
    begin
      typenr:='QTYT6 ';
      if (Trait <> 'G') then
        tampon :=FormatFloat('#0',EnregDucsEdi.E65H)
      else
        tampon :=FormatFloat('#0',EnrDucsEdiGroupee.E65H);
      if (Trait = 'G') then typenr := 'QTYT6E';
        Writeln(FDUCSEDI,typenr+
                         tampon+
                         copy(blanc,1,(18-length(tampon))));
    end;

    // QTYT7  - Effectif des salariés femmes de + de 65 ans inscrits dans
    //          l'établissement au dernier jour de la période (UNEDIC)
    // DRA et AVA : type bordereau 922 ou 923 ou 924
    // -------------------------------------------------------------------------

    if (copy(EnregDucsEdi.TypDucs,1,1)='U') and
       ((EnregDucsEdi.TypBordereau = '922') or
        (EnregDucsEdi.TypBordereau = '923') or
        (EnregDucsEdi.TypBordereau = '924' )) then
    begin
      typenr:='QTYT7 ';
      if (Trait <> 'G') then
        tampon :=FormatFloat('#0',EnregDucsEdi.E65F)
      else
        tampon :=FormatFloat('#0',EnrDucsEdiGroupee.E65F);
      if (Trait = 'G') then typenr := 'QTYT7E';
        Writeln(FDUCSEDI,typenr+
                         tampon+
                         copy(blanc,1,(18-length(tampon))));
    end;
    // QTYT8  - Effectif des salariés hommes dirigeants  inscrits dans
    //          l'établissement au dernier jour de la période (UNEDIC)
    // DRA et AVA : type bordereau 922 ou 923 ou 924
    // -------------------------------------------------------------------------

    if (copy(EnregDucsEdi.TypDucs,1,1)='U') and
       ((EnregDucsEdi.TypBordereau = '922') or
        (EnregDucsEdi.TypBordereau = '923') or
        (EnregDucsEdi.TypBordereau = '924' )) then
    begin
      typenr:='QTYT8 ';
      if (Trait <> 'G') then
        tampon :=FormatFloat('#0',EnregDucsEdi.CadreH)
      else
        tampon :=FormatFloat('#0',EnrDucsEdiGroupee.CadreH);
      if (Trait = 'G') then typenr := 'QTYT8E';
        Writeln(FDUCSEDI,typenr+
                         tampon+
                         copy(blanc,1,(18-length(tampon))));
    end;
    // QTYT9  - Effectif des salariés femmes dirigeants inscrits dans
    //          l'établissement au dernier jour de la période (UNEDIC)
    // DRA et AVA : type bordereau 922 ou 923 ou 924
    // -------------------------------------------------------------------------

    if (copy(EnregDucsEdi.TypDucs,1,1)='U') and
       ((EnregDucsEdi.TypBordereau = '922') or
        (EnregDucsEdi.TypBordereau = '923') or
        (EnregDucsEdi.TypBordereau = '924' )) then
    begin
      typenr:='QTYT9 ';
      if (Trait <> 'G') then
        tampon :=FormatFloat('#0',EnregDucsEdi.CadreF)
      else
        tampon :=FormatFloat('#0',EnrDucsEdiGroupee.CadreF);
      if (Trait = 'G') then typenr := 'QTYT9E';
        Writeln(FDUCSEDI,typenr+
                         tampon+
                         copy(blanc,1,(18-length(tampon))));
    end;
    // QTYTA  - Effectif des salariés hommes en contrat de professionnalisation
    //          inscrits dans l'établissement au dernier jour de la période
    //          (UNEDIC)
    // DRA et AVA : type bordereau 922 ou 923 ou 924 ou 926
    // -------------------------------------------------------------------------

    if (copy(EnregDucsEdi.TypDucs,1,1)='U') and
       ((EnregDucsEdi.TypBordereau = '922') or
        (EnregDucsEdi.TypBordereau = '923') or
        (EnregDucsEdi.TypBordereau = '924' ) or
        (EnregDucsEdi.TypBordereau = '926' )) then
    begin
      typenr:='QTYTA ';
      if (Trait <> 'G') then
        tampon :=FormatFloat('#0',EnregDucsEdi.ProfH)   //PT37
      else
        tampon :=FormatFloat('#0',EnrDucsEdiGroupee.ProfH);  //PT37
      if (Trait = 'G') then typenr := 'QTYTAE';
        Writeln(FDUCSEDI,typenr+
                         tampon+
                         copy(blanc,1,(18-length(tampon))));
    end;
    // QTYTB  - Effectif des salariés femmes titulaires d'un CES inscrits dans
    //          l'établissement au dernier jour de la période (UNEDIC)
    // DRA et AVA : type bordereau 922 ou 923 ou 924 ou 926
    // -------------------------------------------------------------------------

    if (copy(EnregDucsEdi.TypDucs,1,1)='U') and
       ((EnregDucsEdi.TypBordereau = '922') or
        (EnregDucsEdi.TypBordereau = '923') or
        (EnregDucsEdi.TypBordereau = '924' ) or
        (EnregDucsEdi.TypBordereau = '926' )) then
    begin
      typenr:='QTYTB ';
      if (Trait <> 'G') then
        tampon :=FormatFloat('#0',EnregDucsEdi.ProfF)   //PT37
      else
        tampon :=FormatFloat('#0',EnrDucsEdiGroupee.ProfF);   //PT37
      if (Trait = 'G') then typenr := 'QTYTBE';
        Writeln(FDUCSEDI,typenr+
                         tampon+
                         copy(blanc,1,(18-length(tampon))));
    end;
    // QTYTQ  - Effectif des salariés hommes intérimaires inscrits dans
    //          l'établissement au dernier jour de la période (UNEDIC)
    // DRA et AVA : type bordereau 922 ou 923 ou 924 ou 926
    // -------------------------------------------------------------------------

    if (copy(EnregDucsEdi.TypDucs,1,1)='U') and
       ((EnregDucsEdi.TypBordereau = '922') or
        (EnregDucsEdi.TypBordereau = '923') or
        (EnregDucsEdi.TypBordereau = '924' ) or
        (EnregDucsEdi.TypBordereau = '926' )) then
    begin
      if ((EnregDucsEdi.ApePdu = '745B')and (Trait <>'G')) or
         ((EnrDucsEdiGroupee.ApePdu = '745B')and (Trait = 'G')) then
      begin
        typenr:='QTYTQ ';
        if (Trait <> 'G') then
// d PT34
{          tampon :=FormatFloat('#0',EnregDucsEdi.PermanH)}
          tampon :=FormatFloat('#0',EnregDucsEdi.InterimH)
        else
{          tampon :=FormatFloat('#0',EnrDucsEdiGroupee.PermanH);}
          tampon :=FormatFloat('#0',EnrDucsEdiGroupee.InterimH);
// f PT34
        if (Trait = 'G') then typenr := 'QTYTQE';
          Writeln(FDUCSEDI,typenr+
                           tampon+
                           copy(blanc,1,(18-length(tampon))));
      end;
    end;
    // QTYTR - Effectif des salariés femmes intérimaires inscrits dans
    //          l'établissement au dernier jour de la période (UNEDIC)
    // DRA et AVA : type bordereau 922 ou 923 ou 924 ou 926
    // -------------------------------------------------------------------------

    if (copy(EnregDucsEdi.TypDucs,1,1)='U') and
       ((EnregDucsEdi.TypBordereau = '922') or
        (EnregDucsEdi.TypBordereau = '923') or
        (EnregDucsEdi.TypBordereau = '924' ) or
        (EnregDucsEdi.TypBordereau = '926' )) then
    begin
      if ((EnregDucsEdi.ApePdu = '745B') and (Trait <>'G')) or
         ((EnrDucsEdiGroupee.ApePdu = '745B') and (Trait = 'G')) then
      begin
        typenr:='QTYTR ';
        if (Trait <> 'G') then
// d PT34
 {         tampon :=FormatFloat('#0',EnregDucsEdi.PermanF)}
          tampon :=FormatFloat('#0',EnregDucsEdi.InterimF)
        else
{          tampon :=FormatFloat('#0',EnrDucsEdiGroupee.PermanF);}
          tampon :=FormatFloat('#0',EnrDucsEdiGroupee.InterimF);
// f PT34
        if (Trait = 'G') then typenr := 'QTYTRE';
          Writeln(FDUCSEDI,typenr+
                           tampon+
                           copy(blanc,1,(18-length(tampon))));
      end;
    end;

    // QTYTI - Effectif des salariés hommes CEA + CETPE + TEE occ +TTS
    //         (Dispositif de simplification)
    //         inscrits dans l'établissement au dernier jour de la période (UNEDIC)
    // DRA et AVA : type bordereau 922 ou 923 ou 924
    // -------------------------------------------------------------------------

    if (copy(EnregDucsEdi.TypDucs,1,1)='U') and
       ((EnregDucsEdi.TypBordereau = '922') or
        (EnregDucsEdi.TypBordereau = '923') or
        (EnregDucsEdi.TypBordereau = '924' )) then
    begin
      typenr:='QTYTI ';
      if (Trait <> 'G') then
// d PT34
{        tampon :=FormatFloat('#0',EnregDucsEdi.TEEPermH)}
        tampon :=FormatFloat('#0',EnregDucsEdi.SimplifH)
      else
{        tampon :=FormatFloat('#0',EnrDucsEdiGroupee.TEEPermH);}
        tampon :=FormatFloat('#0',EnrDucsEdiGroupee.SimplifH);
// f PT34
      if (Trait = 'G') then typenr := 'QTYTIE';
        Writeln(FDUCSEDI,typenr+
                         tampon+
                         copy(blanc,1,(18-length(tampon))));
    end;

    // QTYTJ  - Effectif des salariés femmesCEA + CETPE + TEE occ +TTS
    //          (Dispositif de simplification)
    //          inscrits dans l'établissement au dernier jour de la période (UNEDIC)
    // DRA et AVA : type bordereau 922 ou 923 ou 924
    // -------------------------------------------------------------------------

    if (copy(EnregDucsEdi.TypDucs,1,1)='U') and
       ((EnregDucsEdi.TypBordereau = '922') or
        (EnregDucsEdi.TypBordereau = '923') or
        (EnregDucsEdi.TypBordereau = '924' )) then
    begin
      typenr:='QTYTJ ';
      if (Trait <> 'G') then
// d PT34
{        tampon :=FormatFloat('#0',EnregDucsEdi.TEEPermF)}
        tampon :=FormatFloat('#0',EnregDucsEdi.SimplifF)
      else
{        tampon :=FormatFloat('#0',EnrDucsEdiGroupee.TEEPermF);}
        tampon :=FormatFloat('#0',EnrDucsEdiGroupee.SimplifF);
// f PT34
      if (Trait = 'G') then typenr := 'QTYTJE';
        Writeln(FDUCSEDI,typenr+
                         tampon+
                         copy(blanc,1,(18-length(tampon))));
    end;

    // QTYTK  - Effectif des salariés homme en CNE inscrits dans
    //          l'établissement au dernier jour de la période (UNEDIC)
    // DRA et AVA : type bordereau 922 ou 923 ou 924
    // -------------------------------------------------------------------------

    if (copy(EnregDucsEdi.TypDucs,1,1)='U') and
       ((EnregDucsEdi.TypBordereau = '922') or
        (EnregDucsEdi.TypBordereau = '923') or
        (EnregDucsEdi.TypBordereau = '924' )) then
    begin
      typenr:='QTYTK ';
      if (Trait <> 'G') then
// d PT34
{        tampon :=FormatFloat('#0',EnregDucsEdi.TEEOccasH)}
        tampon :=FormatFloat('#0',EnregDucsEdi.CNEH)
      else
{        tampon :=FormatFloat('#0',EnrDucsEdiGroupee.TEEOccasH);}
        tampon :=FormatFloat('#0',EnrDucsEdiGroupee.CNEH);
// f PT34
      if (Trait = 'G') then typenr := 'QTYTKE';
        Writeln(FDUCSEDI,typenr+
                         tampon+
                         copy(blanc,1,(18-length(tampon))));
    end;

    // QTYTL  - Effectif des salariés femmes en CNE  inscrits dans
    //          l'établissement au dernier jour de la période (UNEDIC)
    // DRA et AVA : type bordereau 922 ou 923 ou 924
    // -------------------------------------------------------------------------

    if (copy(EnregDucsEdi.TypDucs,1,1)='U') and
       ((EnregDucsEdi.TypBordereau = '922') or
        (EnregDucsEdi.TypBordereau = '923') or
        (EnregDucsEdi.TypBordereau = '924' ) ) then
    begin
      typenr:='QTYTL ';
      if (Trait <> 'G') then
// d PT34
{        tampon :=FormatFloat('#0',EnregDucsEdi.TEEOccasF)}
        tampon :=FormatFloat('#0',EnregDucsEdi.CNEF)
      else
{        tampon :=FormatFloat('#0',EnrDucsEdiGroupee.TEEOccasF);}
        tampon :=FormatFloat('#0',EnrDucsEdiGroupee.CNEF);
// f PT34
      if (Trait = 'G') then typenr := 'QTYTLE';
        Writeln(FDUCSEDI,typenr+
                         tampon+
                         copy(blanc,1,(18-length(tampon))));
    end;

// d PT34
    // QTYTM  - Effectif des salariés Hommes en CDD inscrits dans
    //          l'établissement au dernier jour de la période (UNEDIC)
    // DRA et AVA : type bordereau 922 ou 923 ou 924
    // -------------------------------------------------------------------------

    if (copy(EnregDucsEdi.TypDucs,1,1)='U') and
       ((EnregDucsEdi.TypBordereau = '922') or
        (EnregDucsEdi.TypBordereau = '923') or
        (EnregDucsEdi.TypBordereau = '924' )) then
    begin
      typenr:='QTYTM ';
      if (Trait <> 'G') then
        tampon :=FormatFloat('#0',EnregDucsEdi.CDDH)
      else
        tampon :=FormatFloat('#0',EnrDucsEdiGroupee.CDDH);
      if (Trait = 'G') then typenr := 'QTYTME';
        Writeln(FDUCSEDI,typenr+
                         tampon+
                         copy(blanc,1,(18-length(tampon))));
    end;

    // QTYTN - Effectif des salariés femmes CEA inscrits dans
    //          l'établissement au dernier jour de la période (UNEDIC)
    // DRA et AVA : type bordereau 922 ou 923 ou 924 ou 926
    // -------------------------------------------------------------------------

    if (copy(EnregDucsEdi.TypDucs,1,1)='U') and
       ((EnregDucsEdi.TypBordereau = '922') or
        (EnregDucsEdi.TypBordereau = '923') or
        (EnregDucsEdi.TypBordereau = '924' )) then
    begin
      typenr:='QTYTN ';
      if (Trait <> 'G') then
        tampon :=FormatFloat('#0',EnregDucsEdi.CDDF)
      else
        tampon :=FormatFloat('#0',EnrDucsEdiGroupee.CDDF);
      if (Trait = 'G') then typenr := 'QTYTNE';
        Writeln(FDUCSEDI,typenr+
                         tampon+
                         copy(blanc,1,(18-length(tampon))));
    end;

    // QTYTM  - Effectif des salariés Hommes CEA inscrits dans
    //          l'établissement au dernier jour de la période (UNEDIC)
    // DRA et AVA : type bordereau 922 ou 923 ou 924
    // -------------------------------------------------------------------------
{
    if (copy(EnregDucsEdi.TypDucs,1,1)='U') and
       ((EnregDucsEdi.TypBordereau = '922') or
        (EnregDucsEdi.TypBordereau = '923') or
        (EnregDucsEdi.TypBordereau = '924' )) then
    begin
      typenr:='QTYTM ';
      if (Trait <> 'G') then
        tampon :=FormatFloat('#0',EnregDucsEdi.CEAH)
      else
        tampon :=FormatFloat('#0',EnrDucsEdiGroupee.CEAH);
      if (Trait = 'G') then typenr := 'QTYTME';
        Writeln(FDUCSEDI,typenr+
                         tampon+
                         copy(blanc,1,(18-length(tampon))));
    end;}

    // QTYTN - Effectif des salariés femmes CEA inscrits dans
    //          l'établissement au dernier jour de la période (UNEDIC)
    // DRA et AVA : type bordereau 922 ou 923 ou 924 ou 926
    // -------------------------------------------------------------------------

{
    if (copy(EnregDucsEdi.TypDucs,1,1)='U') and
       ((EnregDucsEdi.TypBordereau = '922') or
        (EnregDucsEdi.TypBordereau = '923') or
        (EnregDucsEdi.TypBordereau = '924' )) then
    begin
      typenr:='QTYTN ';
      if (Trait <> 'G') then
        tampon :=FormatFloat('#0',EnregDucsEdi.CEAF)
      else
        tampon :=FormatFloat('#0',EnrDucsEdiGroupee.CEAF);
      if (Trait = 'G') then typenr := 'QTYTNE';
        Writeln(FDUCSEDI,typenr+
                         tampon+
                         copy(blanc,1,(18-length(tampon))));
    end;}

    // QTYTO  - Effectif des salariés hommes TTS inscrits dans
    //          l'établissement au dernier jour de la période (UNEDIC)
    // DRA et AVA : type bordereau 922 ou 923 ou 924 ou 926
    // -------------------------------------------------------------------------

{
    if (copy(EnregDucsEdi.TypDucs,1,1)='U') and
       ((EnregDucsEdi.TypBordereau = '922') or
        (EnregDucsEdi.TypBordereau = '923') or
        (EnregDucsEdi.TypBordereau = '924' )) then
    begin
      typenr:='QTYTO ';
      if (Trait <> 'G') then
        tampon :=FormatFloat('#0',EnregDucsEdi.TTSH)
      else
        tampon :=FormatFloat('#0',EnrDucsEdiGroupee.TTSH);
      if (Trait = 'G') then typenr := 'QTYTOE';
        Writeln(FDUCSEDI,typenr+
                         tampon+
                         copy(blanc,1,(18-length(tampon))));
    end;}

    // QTYTP  - Effectif des salariés femmes TTS inscrits dans
    //          l'établissement au dernier jour de la période (UNEDIC)
    // DRA et AVA : type bordereau 922 ou 923 ou 924
    // -------------------------------------------------------------------------

{
    if (copy(EnregDucsEdi.TypDucs,1,1)='U') and
       ((EnregDucsEdi.TypBordereau = '922') or
        (EnregDucsEdi.TypBordereau = '923') or
        (EnregDucsEdi.TypBordereau = '924' )) then
    begin
      typenr:='QTYTP ';
      if (Trait <> 'G') then
        tampon :=FormatFloat('#0',EnregDucsEdi.TTSF)
      else
        tampon :=FormatFloat('#0',EnrDucsEdiGroupee.TTSF);
      if (Trait = 'G') then typenr := 'QTYTPE';
        Writeln(FDUCSEDI,typenr+
                         tampon+
                         copy(blanc,1,(18-length(tampon))));
    end;}
// f PT34

    // RFFDC - Référence d'identification complémentaire de l'établissement
    // --------------------------------------------------------------------
{    if ((copy(EnregDucsEdi.TypDucs,1,1)='A') and
        (copy(EnregDucsEdi.Abrege,3,2) <> '00')) or
       (copy(EnregDucsEdi.TypDucs,1,1)<>'A')then
      begin}
       if (copy(EnregDucsEdi.TypDucs,1,1)<>'I') then
        begin
         // ACOSS UNEDIC    (n° interne)
         typenr :='RFFDC ';
         if (Trait <> 'G') then
          Writeln(FDUCSEDI,'RFFDC '+
                           EnregDucsEdi.NumeroInterne+
                           copy(blanc,1,(35-length(EnregDucsEdi.NumeroInterne))))
         else
         begin
          if (copy(EnregDucsEdi.TypDucs,1,1) = 'U') then
// d PT35
          begin
            Writeln(FDUCSEDI,'RFFDC '+
                              EnregDucsEdi.NumeroInterne+
                              copy(blanc,1,(35-length(EnregDucsEdi.NumeroInterne))));

            Typenr :=  'RFFDCE';
            Writeln(FDUCSEDI,typenr+
                             EnrDucsEdiGroupee.NumeroInterne+
                             copy(blanc,1,(35-length(EnrDucsEdiGroupee.NumeroInterne))));
          end
          else
            Writeln(FDUCSEDI,'RFFDC '+
                              EnrDucsEdiGroupee.NumeroInterne+
                              copy(blanc,1,(35-length(EnrDucsEdiGroupee.NumeroInterne))));
// f  PT35
         end;
        end
       else
        // IRC   (Centre payeur)
        begin
         if (Trait <> 'G') then
         begin
           if (EnregDucsEdi.CentrePayeur <> '') then
             Tampon  := EnregDucsEdi.CentrePayeur
           else
           begin
             Tampon  := Format('%4.4d',[EnregDucsEdi.NoCentrePayeur]);
             PGZeroAGauche(tampon,5);
           end;
         end
         else
         begin
           if (EnrDucsEdiGroupee.CentrePayeur <> '') then
             Tampon  := EnrDucsEdiGroupee.CentrePayeur
           else
           begin
             Tampon  := Format('%4.4d',[EnregDucsEdi.NoCentrePayeur]);
             PGZeroAGauche(tampon,5);
           end;
         end;

         Writeln(FDUCSEDI,'RFFDC '+
                          tampon+
                          copy(blanc,1,(35-length(tampon))));
        end;
      //end;
    // RFFACT - Référence complémentaire de l'établissement, Code APE
    // --------------------------------------------------------------
    if ((copy(EnregDucsEdi.TypDucs,1,1)='A') and
        (copy(EnregDucsEdi.Abrege,3,2) = '00')) or
       (copy(EnregDucsEdi.TypDucs,1,1)='U')then
      begin
        // ACOSS (TR) -  UNEDIC    (Code APE)
        typenr := 'RFFACT';
        if (Trait <> 'G') then
         Writeln(FDUCSEDI,'RFFACT'+
                          EnregDucsEdi.ApePdu+
                          copy(blanc,1,(35-length(EnregDucsEdi.ApePdu))))
        else
        begin
         if (copy(EnregDucsEdi.TypDucs,1,1) = 'U') then  typenr := 'RFFACE';
         Writeln(FDUCSEDI,typenr+
                          EnrDucsEdiGroupee.ApePdu+
                          copy(blanc,1,(35-length(EnrDucsEdiGroupee.ApePdu))))
        end;
      end;
// d PT30 DUCS EDI V4.2
    // RFFZDA - UNEDIC DRA - Salaires de décembre dans DADS (petit décalage)
    // ---------------------------------------------------------------------

    if (copy(EnregDucsEdi.TypDucs,1,1)='U') and
       ((EnregDucsEdi.TypBordereau = '922') or
        (EnregDucsEdi.TypBordereau = '923'))then
    begin
       if (Trait <> 'G') then
         typenr := 'RFFZDA'
       else
         typenr := 'RFFZAE';
      if (VH_Paie.PGDecalagePetit = true) then
        Writeln(FDUCSEDI,typenr+
                          '1'+
                          copy(blanc,1,34))
      else
        Writeln(FDUCSEDI,typenr+
                          '0'+
                          copy(blanc,1,34));
    end;

    // RFFZE1 - UNEDIC  - Ecart effectif pour Licenciements économiques
    // ----------------------------------------------------------------

// d PT31
//    if (copy(EnregDucsEdi.TypDucs,1,1)='U') then
    if (copy(EnregDucsEdi.TypDucs,1,1)='U') and
       (EnregDucsEdi.TypBordereau <> '920') and
       (EnregDucsEdi.TypBordereau <> '921') then
// f PT31
    begin
      if (Trait <> 'G') then
         typenr := 'RFFZE1'
       else
         typenr := 'RFFZ1E';

      if EnregDucsEdi.EcartZe1 then
        Writeln(FDUCSEDI,typenr+
                         '1'+
                         copy(blanc,1,34))
      else
        Writeln(FDUCSEDI,typenr+
                          '0'+
                          copy(blanc,1,34));
    end;
    // RFFZE2 - UNEDIC  - Ecart effectif pour Fins de CDD
    // --------------------------------------------------
// d PT31
//    if (copy(EnregDucsEdi.TypDucs,1,1)='U') then
    if (copy(EnregDucsEdi.TypDucs,1,1)='U') and
       (EnregDucsEdi.TypBordereau <> '920') and
       (EnregDucsEdi.TypBordereau <> '921') then
// f PT31
    begin
      if (Trait <> 'G') then
        typenr := 'RFFZE2'
      else
        typenr := 'RFFZ2E';

      if EnregDucsEdi.EcartZe2 then
        Writeln(FDUCSEDI,typenr+
                         '1'+
                         copy(blanc,1,34))
      else
        Writeln(FDUCSEDI,typenr+
                          '0'+
                          copy(blanc,1,34));
    end;
    // RFFZE3 - UNEDIC  - Ecart effectif pour Embauches CDI
    // ----------------------------------------------------
// d PT31
//    if (copy(EnregDucsEdi.TypDucs,1,1)='U') then
    if (copy(EnregDucsEdi.TypDucs,1,1)='U') and
       (EnregDucsEdi.TypBordereau <> '920') and
       (EnregDucsEdi.TypBordereau <> '921') then
// f PT31
    begin
      if (Trait <> 'G') then
        typenr := 'RFFZE3'
      else
        typenr := 'RFFZ3E';

      if EnregDucsEdi.EcartZe3 then
        Writeln(FDUCSEDI,typenr+
                         '1'+
                         copy(blanc,1,34))
      else
        Writeln(FDUCSEDI,typenr+
                          '0'+
                          copy(blanc,1,34));
    end;
    // RFFZE4 - UNEDIC  - Ecart effectif pour Embauches CDD
    // ----------------------------------------------------
// d PT31
//    if (copy(EnregDucsEdi.TypDucs,1,1)='U') then
    if (copy(EnregDucsEdi.TypDucs,1,1)='U') and
       (EnregDucsEdi.TypBordereau <> '920') and
       (EnregDucsEdi.TypBordereau <> '921')then
// f PT31
    begin
      if (Trait <> 'G') then
        typenr := 'RFFZE4'
      else
        typenr := 'RFFZ4E';

      if EnregDucsEdi.EcartZe4 then
        Writeln(FDUCSEDI,typenr+
                         '1'+
                         copy(blanc,1,34))
      else
        Writeln(FDUCSEDI,typenr+
                          '0'+
                          copy(blanc,1,34));
    end;
    // RFFZE5 - UNEDIC  - Ecart effectif pour Transfert des salariés vers une autre entreprise
    // ---------------------------------------------------------------------------------------

// d PT31
//    if (copy(EnregDucsEdi.TypDucs,1,1)='U') then
    if (copy(EnregDucsEdi.TypDucs,1,1)='U') and
       (EnregDucsEdi.TypBordereau <> '920') and
       (EnregDucsEdi.TypBordereau <> '921') then
// f PT31
    begin
      if (Trait <> 'G') then
        typenr := 'RFFZE5'
      else
        typenr := 'RFFZ5E';

      if EnregDucsEdi.EcartZe5 then
        Writeln(FDUCSEDI,typenr+
                         '1'+
                         copy(blanc,1,34))
      else
        Writeln(FDUCSEDI,typenr+
                          '0'+
                          copy(blanc,1,34));

    end;
    // RFFZE6 - UNEDIC  - Ecart effectif pour Reprise des salariés d’une autre entreprise
    // -----------------------------------------------------------------------------------

// d PT31
//    if (copy(EnregDucsEdi.TypDucs,1,1)='U') then
    if (copy(EnregDucsEdi.TypDucs,1,1)='U') and
       (EnregDucsEdi.TypBordereau <> '920') and
       (EnregDucsEdi.TypBordereau <> '921') then
// f PT31
    begin
      if (Trait <> 'G') then
        typenr := 'RFFZE6'
      else
        typenr := 'RFFZ6E';

      if EnregDucsEdi.EcartZe6 then
        Writeln(FDUCSEDI,typenr+
                         '1'+
                         copy(blanc,1,34))
      else
        Writeln(FDUCSEDI,typenr+
                          '0'+
                          copy(blanc,1,34));
    end;

    // RFFZE7 - UNEDIC  - Ecart effectif pour Redressement/liquidation judiciaire
    // --------------------------------------------------------------------------
// d PT31
//    if (copy(EnregDucsEdi.TypDucs,1,1)='U') then
    if (copy(EnregDucsEdi.TypDucs,1,1)='U') and
       (EnregDucsEdi.TypBordereau <> '920') and
       (EnregDucsEdi.TypBordereau <> '921') then
// f PT31
    begin
      if (Trait <> 'G') then
        typenr := 'RFFZE7'
      else
        typenr := 'RFFZ7E';
      if EnregDucsEdi.EcartZe7 then
        Writeln(FDUCSEDI,typenr+
                         '1'+
                         copy(blanc,1,34))
      else
        Writeln(FDUCSEDI,typenr+
                          '0'+
                          copy(blanc,1,34));
    end;
    // RFFZE8 - UNEDIC  - Ecart effectif pour motif Autre
    // ---------------------------------------------------------------------
// d PT31
//    if (copy(EnregDucsEdi.TypDucs,1,1)='U') then
    if (copy(EnregDucsEdi.TypDucs,1,1)='U') and
       (EnregDucsEdi.TypBordereau <> '920') and
       (EnregDucsEdi.TypBordereau <> '921') then
// f PT31
    begin
      if (Trait <> 'G') then
        typenr := 'RFFZE8'
      else
        typenr := 'RFFZ8E';
      if EnregDucsEdi.EcartZe8 then
        Writeln(FDUCSEDI,typenr+
                         '1'+
                         copy(blanc,1,34))
      else
        Writeln(FDUCSEDI,typenr+
                          '0'+
                          copy(blanc,1,34));
    end;
// f PT30 DUCS EDI V4.2

// d PT34
    // RFFZE9 - UNEDIC  - Ecart effectif pour embauche CNE
    // ---------------------------------------------------------------------
    if (copy(EnregDucsEdi.TypDucs,1,1)='U') and
       (EnregDucsEdi.TypBordereau <> '920') and
       (EnregDucsEdi.TypBordereau <> '921') then
// f PT31
    begin
      if (Trait <> 'G') then
        typenr := 'RFFZE9'
      else
        typenr := 'RFFZ9E';
      if EnregDucsEdi.EcartZe9 then
        Writeln(FDUCSEDI,typenr+
                         '1'+
                         copy(blanc,1,34))
      else
        Writeln(FDUCSEDI,typenr+
                          '0'+
                          copy(blanc,1,34));
    end;
// f PT34

end;
// fin Groupe2Declaration

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 13/05/2002
Modifié le ... :   /  /
Description .. : Traitement des segments du groupe 3 de la déclaration
Suite ........ : Spécificité URSSAF
Mots clefs ... : PAIE, PGDUCSEDI
*****************************************************************}
procedure Groupe3ACOSS(EnregDucsEdi : TDucsEdi;EnrDucsEdiGroupee : TDucsEdiGroupee;Trait : string);
var
   CodeTypeEnCours, TauxEnCours                        : string;
   TypePersonnel,Taux,MtDeplafonne,MtPlafonne,Effectif : string;
   DateChgtTx                                          : string;
   MtNonZero                                           : boolean;
   blanc                                               : string;
   tampon                                              : string;
   LesLignes                                           : TOB;
   WMonnaieTenue                                       : string;
   WNbTransport                                        : integer;
   WnbTAT                                              : integer; // PT37
   DateChgtTxEnCours                                   : string;  // PT30 DUCS EDI V4.2
   IndicAT                                             : boolean; // PT37
   DateChgtTxAT                                        : string;  // PT37
   DateChgtTxATEnCours                                 : string;  // PT37
   QteEnBase                                           : string;  // PT50
begin
  if (Trait <> 'G') then
   begin
    WMonnaieTenue := EnregDucsEdi.MonnaieTenue;
    WNbtransport :=  EnregDucsEdi.NbTransport;
    WnbTAT := EnregDucsEdi.NbTAT;                    // PT37
   end
  else
   begin
    WMonnaieTenue := EnrDucsEdiGroupee.MonnaieTenue;
    WNbtransport :=  EnrDucsEdiGroupee.NbTransport;
    WnbTAT := EnrDucsEdiGroupee.NbTAT;                    // PT37
   end;
  MtNonZero := False;
  blanc := StringOfChar(' ',35);
  CodeTypeEnCours := '';
  TauxEnCours :=  '';
  DateChgtTxEnCours := ''; // PT30 DUCS EDI V4.2
  DateChgtTx :='';
  IndicAT := false;        // PT37
  DateChgtTxAT := '';      // PT37
 if TOB_Lignes <> NIL then
    begin
     LesLignes := TOB_Lignes.FindFirst ([''],[''],TRUE);
     While LesLignes <> NIL  do
      begin
        if (Leslignes.getValue('PDD_TYPECOTISATION') <> 'I') and
           (Leslignes.getValue('PDD_TYPECOTISATION') <> 'S') then
        // les lignes Intitulé et Sous-Total ne sont pas traitées.
        begin
          tampon := copy(Leslignes.getValue('PDD_CODIFICATION'),4,3) ;
          tampon := FloatToStr(LesLignes.getValue('PDD_TAUXCOTISATION'));
          tampon := Leslignes.getValue('PDD_TYPECOTISATION');
// d PT37
          if (not IndicAT) then
          begin
            if (Leslignes.getValue('PDD_DATECHGTTAUX') = IDate1900) then
              DateChgtTx := '0'
            else
              DateChgtTx := Copy(DateToStr(Leslignes.getValue('PDD_DATECHGTTAUX')),7,10)+
                            Copy(DateToStr(Leslignes.getValue('PDD_DATECHGTTAUX')),4,2)+
                            Copy(DateToStr(Leslignes.getValue('PDD_DATECHGTTAUX')),1,2);
          end
          else
//          if (IndicAT) then
          begin
            if (Leslignes.getValue('PDD_DATECHGTTAUX') <> IDate1900) then
              DateChgtTxAT := Copy(DateToStr(Leslignes.getValue('PDD_DATECHGTTAUX')),7,10)+
                              Copy(DateToStr(Leslignes.getValue('PDD_DATECHGTTAUX')),4,2)+
                              Copy(DateToStr(Leslignes.getValue('PDD_DATECHGTTAUX')),1,2)
            else
              DateChgtTxAT := '0';
          end;
// f PT37
          if (copy(Leslignes.getValue('PDD_CODIFICATION'),4,3) <> CodeTypeEnCours) or
             ((FloatToStr(LesLignes.getValue('PDD_TAUXCOTISATION')) <> TauxEnCours) and
              (Leslignes.getValue('PDD_TYPECOTISATION') = 'A') or
              (DateChgtTx <> DateChgtTxEnCours)) then   // PT30 DUCS EDI V4.2
           // chgt de codif ou de taux s'il s'agit d'un taux AT
           begin
            if (CodeTypeEnCours <> '') and (MtNonZero = True) then
            begin
               if (taux = '') then taux :='0';
               if (MtDeplafonne = '') then MtDeplafonne := '0';
               if (MtPlafonne = '') then MtPlafonne := '0';
               if (Effectif = '') then Effectif := '0';
               if (((TypePersonnel <> '900') or (WNbTransport <= 1)) and
                  ((TypePersonnel <> '901') or (WNbTAT <= 1))) then  // PT37
               begin
// d PT37
                 if (IndicAT) then
                   Writeln (FDUCSEDI,'ACO   '+
                                 TypePersonnel+
                                 Taux+
                                 copy(blanc,1,(8-length(Taux)))+
                                 MtDeplafonne+
                                 copy(blanc,1,(18-length(MtDeplafonne)))+
                                 MtPlafonne+
                                 copy(blanc,1,(18-length(MtPlafonne)))+
                                 Effectif+
                                 copy(blanc,1,(15-length(Effectif)))+
                                 DateChgtTxATEnCours+      // PT37
                                 copy(blanc,1,(8-length(DateChgtTxATEnCours)))+ // PT37
                                 copy(blanc,1,15)+//PT50
                                 WMonnaieTenue)
                 else
// f PT37

                   Writeln (FDUCSEDI,'ACO   '+
                                 TypePersonnel+
                                 Taux+
                                 copy(blanc,1,(8-length(Taux)))+
                                 MtDeplafonne+
                                 copy(blanc,1,(18-length(MtDeplafonne)))+
                                 MtPlafonne+
                                 copy(blanc,1,(18-length(MtPlafonne)))+
                                 Effectif+
                                 copy(blanc,1,(15-length(Effectif)))+
// PT30 DUCS EDI V4.2            DateChgtTx+
//                               copy(blanc,1,(8-length(DateChgtTx)))+
                                 DateChgtTxEnCours+
                                 copy(blanc,1,(8-length(DateChgtTxEnCours)))+
// d PT50
                                 QteEnBase+
                                 copy(blanc,1,(15-length(QteEnBase)))+
// f PT50
                                 WMonnaieTenue);
               end;

               EnregDucsEdi.NbCot := EnregDucsEdi.NbCot+1;
// PT37
               if (IndicAT) and ((copy(Leslignes.getValue('PDD_CODIFICATION'),4,3) <> CodeTypeEnCours) or
                                 ((FloatToStr(LesLignes.getValue('PDD_TAUXCOTISATION')) <> TauxEnCours) and
                                  (Leslignes.getValue('PDD_TYPECOTISATION') = 'A'))) then
               begin
                 IndicAT := false;
                 DateChgtTxAT := '';
                 if (Leslignes.getValue('PDD_DATECHGTTAUX') = IDate1900) then
                    DateChgtTx := '0'
                 else
                     DateChgtTx := Copy(DateToStr(Leslignes.getValue('PDD_DATECHGTTAUX')),7,10)+
                                   Copy(DateToStr(Leslignes.getValue('PDD_DATECHGTTAUX')),4,2)+
                                   Copy(DateToStr(Leslignes.getValue('PDD_DATECHGTTAUX')),1,2);
                end;
// PT37
            end;
             TypePersonnel := '';
             Taux := '';
             MtDeplafonne := '';
             MtPlafonne := '';
             Effectif := '';
             MtNonZero := False;
             QteEnBase := ''; //PT50
             CodeTypeEnCours :=copy(Leslignes.getValue('PDD_CODIFICATION'),4,3);
// d PT30 DUCS EDI V4.2
             DateChgtTxEnCours := DateChgtTx;
             DateChgtTxATEnCours := DateChgtTxAT; //PT37
// f PT30 DUCS EDI V4.2
             TauxEnCours := '';
          end;
          if (DateChgtTxATEnCours = '') then DateChgtTxATEnCours := DateChgtTxAT; // PT37

          TypePersonnel := copy(Leslignes.getValue('PDD_CODIFICATION'),4,3);
          if ((copy(Leslignes.getValue('PDD_CODIFICATION'),4,4) = '900P') and
              (WNbTransport <=1)) or
             ((copy(Leslignes.getValue('PDD_CODIFICATION'),4,4) = '901P') and       // PT37
              (WNbTAT <=1)) or
             (Leslignes.getValue('PDD_TYPECOTISATION') = 'A') then
           // Transport ou taux AT
           begin
            taux := FormatFloat('#0',(Leslignes.getValue('PDD_TAUXCOTISATION')*10000));
            taux := ColleZeroDevant(StrToInt(taux),8);         // PT30 DUCS EDI V4.2
            TauxEnCours := Taux;
            MtNonZero := True;
// d PT37
            if (Leslignes.getValue('PDD_TYPECOTISATION') = 'A') then
              IndicAT := true;
// f PT37
           end;
// d PT37
           if ((TypePersonnel <> '900') or (WNbTransport <= 1)) and
              ((TypePersonnel <> '901') or (WNbTAT <= 1)) then
           // les lignes 900 et 901 ne sont traitées que si elles sont uniques
           begin
// f PT37
             if (copy(Leslignes.getValue('PDD_CODIFICATION'),7,1) = 'D') or
                ((copy(Leslignes.getValue('PDD_CODIFICATION'),7,1) = ' ') and
                ((Leslignes.getValue('PDD_TYPECOTISATION') <> 'M') and
                 (Leslignes.getValue('PDD_TYPECOTISATION') <> 'Q') and
                 (Leslignes.getValue('PDD_TYPECOTISATION') <> 'D'))) then   //PT50
               // Déplafonné (codif en D, ou blanc mais pas forfait)
               begin
                 MtDeplafonne :=  FormatFloat('#0',(Leslignes.getValue('PDD_BASECOTISATION')));
                 MtNonZero := True;
               end;

             if (copy(Leslignes.getValue('PDD_CODIFICATION'),7,1) = 'P') then
               // Plafonne ( codif en P, ou blanc et forfait)
               begin
                 MtPlafonne := FormatFloat('#0',(Leslignes.getValue('PDD_BASECOTISATION')));
                 MtNonZero := True;
               end;

             if ((Leslignes.getValue('PDD_TYPECOTISATION') = 'M') or
                 (Leslignes.getValue('PDD_TYPECOTISATION') = 'Q') or
                 (Leslignes.getValue('PDD_TYPECOTISATION') = 'C') or        //PT50
                 (Leslignes.getValue('PDD_TYPECOTISATION') = 'D')) then     //PT50
             begin
                MtNonZero := True;

             if (Leslignes.getValue('PDD_MTCOTISAT') < 0) then
               // si le forfait est négatif il est renseigné en valeur absolue
               MtPlafonne := FormatFloat('#0',(Leslignes.getValue('PDD_MTCOTISAT')*-1))
             else
               MtPlafonne := FormatFloat('#0',Leslignes.getValue('PDD_MTCOTISAT'));
             end;
// d PT30 DUCS EDI V4.2
// effectif non utilisé pour un TR
             if (EnregDucsEdi.TypBordereau <> '915') and
                (EnregDucsEdi.TypBordereau <> '916') then
// f PT30 DUCS EDI V4.2
               Effectif :=  Leslignes.getValue('PDD_EFFECTIF');
// d PT50
             if (Leslignes.getValue('PDD_TYPECOTISATION') = 'D') then
               QteEnBase := FormatFloat('#0',(Leslignes.getValue('PDD_BASECOTISATION')*10000));
// f PT50
           end;
        end;
        LesLignes := TOB_Lignes.FindNext ([''],[''],TRUE);
      end;
     if ((TypePersonnel <> '900') or (WNbTransport <= 1)) and
        ((TypePersonnel <> '901') or (WNbTAT <= 1)) then // PT37
     begin
      if (CodeTypeEnCours <> '') and (MtNonZero = True) then
       begin
         if (taux = '') then taux :='0';
         if (MtDeplafonne = '') then MtDeplafonne := '0';
         if (MtPlafonne = '') then MtPlafonne := '0';
         if (Effectif = '') then Effectif := '0';

// d PT37
         if (IndicAT) then
           Writeln (FDUCSEDI,'ACO   '+
                             TypePersonnel+
                             Taux+
                             copy(blanc,1,(8-length(Taux)))+
                             MtDeplafonne+
                             copy(blanc,1,(18-length(MtDeplafonne)))+
                             MtPlafonne+
                             copy(blanc,1,(18-length(MtPlafonne)))+
                             Effectif+
                             copy(blanc,1,(15-length(Effectif)))+
                             DateChgtTxATEnCours+        // PT37
                             copy(blanc,1,(8-length(DateChgtTxATEnCours)))+ // PT37
                             copy(blanc,1,15)+ // PT50
                             WMonnaieTenue)
         else
// f PT37


         Writeln (FDUCSEDI,'ACO   '+
                           TypePersonnel+
                           Taux+
                           copy(blanc,1,(8-length(Taux)))+
                           MtDeplafonne+
                           copy(blanc,1,(18-length(MtDeplafonne)))+
                           MtPlafonne+
                           copy(blanc,1,(18-length(MtPlafonne)))+
                           Effectif+
                           copy(blanc,1,(15-length(Effectif)))+
//                         DateChgtTx+                            // PT30 DUCS EDI V4.2
//                         copy(blanc,1,(8-length(DateChgtTx)))+  // PT30 DUCS EDI V4.2
                           DateChgtTxEnCours+
                           copy(blanc,1,(8-length(DateChgtTxEnCours)))+
// d PT50
                           QteEnBase+
                           copy(blanc,1,(15-length(QteEnBase)))+
// f PT50
                           WMonnaieTenue);

         EnregDucsEdi.NbCot := EnregDucsEdi.NbCot+1;
      end;
     end;
    end;
end;
// fin Groupe3ACOSS
{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 13/05/2002
Modifié le ... : 29/06/2005
Description .. : Traitement des segments des groupes 4 et 5 de la
Suite ........ : déclaration. Il s'agit des segments liés au "transport"
Suite ........ : versement transport et taxe additionnelle de transport
Suite ........ : Spécifictié URSSAF
Mots clefs ... : PAIE, PGDUCSEDI
*****************************************************************}
procedure Groupe4_5ACOSS(EnregDucsEdi : TDucsEdi;EnrDucsEdiGroupee : TDucsEdiGroupee;Trait : string);
var
 blanc,localisation,taux1,montant1,taux2,montant2,taux3,montant3 : string;
 CuEnCours : string;
 LTransport : TOB;
 LTAT       : TOB;  // PT30 DUCS EDI V4.2
 WMonnaieTenue : string;
begin
 if (Trait <> 'G') then
  WMonnaieTenue := EnregDucsEdi.MonnaieTenue
 else
  WMonnaieTenue := EnrDucsEdiGroupee.MonnaieTenue;

 blanc := StringOfChar(' ',35);
 localisation := '';
 taux1 := '0';
 montant1 := '0';
 taux2 := '0';
 montant2 := '0';
 taux3 := '0';
 montant3 := '0';
 CuEnCours := '';
 // Versement transport
 if TOB_Transport <> NIL then
 begin
   LTransport := TOB_Transport.FindFirst ([''],[''],TRUE);
   While LTransport <> NIL  do
   begin
     if (CuEnCours <> '') and
// PT30 DUCS EDI V4.2
//      (CuEnCours <> LTransport.getValue('PDD_COMURBAINE')) then
        (CuEnCours <> LTransport.getValue('PDD_CODECOMMUNE')) then
     begin
       Writeln (FDUCSEDI,'ACOLOC'+
                         Localisation+
//PT30 DUCS EDI V4.2
//                       copy(blanc,1,(4-length(Localisation)))+
                         copy(blanc,1,(5-length(Localisation)))+
                         Taux1+
                         copy(blanc,1,(8-length(Taux1)))+
                         Montant1+
                         copy(blanc,1,(18-length(Montant1)))+
                         Taux2+
                         copy(blanc,1,(8-length(Taux2)))+
                         Montant2+
                         copy(blanc,1,(18-length(Montant2)))+
                         Taux3+
                         copy(blanc,1,(8-length(Taux3)))+
                         Montant3+
                         copy(blanc,1,(18-length(Montant3)))+
                         WMonnaieTenue);

       localisation := '';
       taux1 := '0';
       montant1 := '0';
       taux2 := '0';
       montant2 := '0';
       taux3 := '0';
       montant3 := '0';
       CuEnCours := '';
     end;

// PT30 DUCS EDI V4.2
//   if (CuEnCours = LTransport.getValue('PDD_COMURBAINE')) or
     if (CuEnCours = LTransport.getValue('PDD_CODECOMMUNE')) or
        (CuEnCours = '') then
     begin
       if (Taux1 = '0') then
       begin
         Taux1 := FormatFloat('#0',(LTransport.getValue('PDD_TAUXCOTISATION')*10000));
         Taux1 := ColleZeroDevant(StrToInt(Taux1),8);      // PT30 DUCS EDI V4.2
         Montant1 := FormatFloat('#0',(LTransport.getValue('PDD_BASECOTISATION')));
       end
       else
         if (Taux2 = '0') then
         begin
           Taux2 := FormatFloat('#0',(LTransport.getValue('PDD_TAUXCOTISATION')*10000));
           Taux2 := ColleZeroDevant(StrToInt(Taux2),8);      // PT30 DUCS EDI V4.2
           Montant2 := FormatFloat('#0',(LTransport.getValue('PDD_BASECOTISATION')));
         end
         else
           if (Taux3 = '0') then
           begin
             Taux3 := FormatFloat('#0',(LTransport.getValue('PDD_TAUXCOTISATION')*10000));
             Taux3 := ColleZeroDevant(StrToInt(Taux3),8);      // PT30 DUCS EDI V4.2
             Montant3 := FormatFloat('#0',(LTransport.getValue('PDD_BASECOTISATION')));
           end;
     end;
// PT30 DUCS EDI V4.2        CuEnCours := LTransport.getValue('PDD_COMURBAINE');
// PT30 DUCS EDI V4.2        Localisation := LTransport.getValue('PDD_COMURBAINE');
       CuEnCours := LTransport.getValue('PDD_CODECOMMUNE');
       Localisation := LTransport.getValue('PDD_CODECOMMUNE');

       LTransport := TOB_Transport.FindNext ([''],[''],TRUE);
   end;
   if (CuEnCours <> '') then
   // dernière CU
   begin
     Writeln (FDUCSEDI,'ACOLOC'+
                        Localisation+
// PT30 DUCS EDI V4.2
//   copy(blanc,1,(4-length(Localisation)))+
     copy(blanc,1,(5-length(Localisation)))+
                  Taux1+
                  copy(blanc,1,(8-length(Taux1)))+
                  Montant1+
                  copy(blanc,1,(18-length(Montant1)))+
                  Taux2+
                  copy(blanc,1,(8-length(Taux2)))+
                  Montant2+
                  copy(blanc,1,(18-length(Montant2)))+
                  Taux3+
                  copy(blanc,1,(8-length(Taux3)))+
                  Montant3+
                  copy(blanc,1,(18-length(Montant3)))+
                  WMonnaieTenue);

     Localisation := '';
     taux1 := '0';
     montant1 := '0';
     taux2 := '0';
     montant2 := '0';
     taux3 := '0';
     montant3 := '0';
     CuEnCours := '';
   end;
 end;

 if TOB_TAT <> NIL then
 begin
   LTAT := TOB_TAT.FindFirst ([''],[''],TRUE);
   While LTAT <> NIL  do
   begin
     if (CuEnCours <> '') and
        (CuEnCours <> LTAT.getValue('PDD_CODECOMMUNE')) then
     begin
       Writeln (FDUCSEDI,'ACOTAT'+
                          Localisation+
                          copy(blanc,1,(5-length(Localisation)))+
                          Taux1+
                          copy(blanc,1,(8-length(Taux1)))+
                          Montant1+
                          copy(blanc,1,(18-length(Montant1)))+
                          Taux2+
                          copy(blanc,1,(8-length(Taux2)))+
                          Montant2+
                          copy(blanc,1,(18-length(Montant2)))+
                          Taux3+
                          copy(blanc,1,(8-length(Taux3)))+
                          Montant3+
                          copy(blanc,1,(18-length(Montant3)))+
                          WMonnaieTenue);

       localisation := '';
       taux1 := '0';
       montant1 := '0';
       taux2 := '0';
       montant2 := '0';
       taux3 := '0';
       montant3 := '0';
       CuEnCours := '';
     end;

     if (CuEnCours = LTAT.getValue('PDD_CODECOMMUNE')) or
        (CuEnCours = '') then
     begin
       if (Taux1 = '0') then
       begin
         Taux1 := FormatFloat('#0',(LTAT.getValue('PDD_TAUXCOTISATION')*10000));
         Taux1 := ColleZeroDevant(StrToInt(Taux1),8);      // PT30 DUCS EDI V4.2
         Montant1 := FormatFloat('#0',(LTAT.getValue('PDD_BASECOTISATION')));
       end
       else
         if (Taux2 = '0') then
         begin
           Taux2 := FormatFloat('#0',(LTAT.getValue('PDD_TAUXCOTISATION')*10000));
           Taux2 := ColleZeroDevant(StrToInt(Taux2),8);      //PT30 DUCS EDI V4.2
           Montant2 := FormatFloat('#0',(LTAT.getValue('PDD_BASECOTISATION')));
         end
         else
           if (Taux3 = '0') then
           begin
             Taux3 := FormatFloat('#0',(LTAT.getValue('PDD_TAUXCOTISATION')*10000));
             Taux3 := ColleZeroDevant(StrToInt(Taux3),8);      // PT30 DUCS EDI V4.2
             Montant3 := FormatFloat('#0',(LTAT.getValue('PDD_BASECOTISATION')));
           end;
     end;
     CuEnCours := LTAT.getValue('PDD_CODECOMMUNE');
     Localisation := LTAT.getValue('PDD_CODECOMMUNE');

     LTAT := TOB_TAT.FindNext ([''],[''],TRUE);
   end;
   if (CuEnCours <> '') then
   // dernière CU
   begin
     Writeln (FDUCSEDI,'ACOTAT'+
                        Localisation+
                        copy(blanc,1,(5-length(Localisation)))+
                        Taux1+
                        copy(blanc,1,(8-length(Taux1)))+
                        Montant1+
                        copy(blanc,1,(18-length(Montant1)))+
                        Taux2+
                        copy(blanc,1,(8-length(Taux2)))+
                        Montant2+
                        copy(blanc,1,(18-length(Montant2)))+
                        Taux3+
                        copy(blanc,1,(8-length(Taux3)))+
                        Montant3+
                        copy(blanc,1,(18-length(Montant3)))+
                        WMonnaieTenue);

    Localisation := '';
    taux1 := '0';
    montant1 := '0';
    taux2 := '0';
    montant2 := '0';
    taux3 := '0';
    montant3 := '0';
    CuEnCours := '';
  end;
 end;
end;
// fin Groupe4_5ACOSS

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 13/05/2002
Modifié le ... :   /  /
Description .. : Traitement des segment du groupe 3 de la déclaration.
Suite ........ : Spécificité IRC
Mots clefs ... : PAIE, PGDUCSEDI
*****************************************************************}
procedure Groupe3IRC(var EnregDucsEdi : TDucsEdi; EnrDucsEdiGroupee : TDucsEdiGroupee; Trait : string; var Nolg : integer);

var
   NatureAssiette, QualifiantCotis,ConditionSpe,Institution, GrpCotisant : string;
   CategSocio, NoLigne, DateChgtTx, Taux, Cotisation, Base, forfait :  string;
   EffCotisation, EffForfait, blanc : string;
   LesLignes : TOB;
   QQ : Tquery;
   WMonnaieTenue : string;
   DebCodif, SuitCodif : string;        

begin
  if (Trait <> 'G') then
    WMonnaieTenue := EnregDucsEdi.MonnaieTenue
  else
    WMonnaieTenue := EnrDucsEdiGroupee.MonnaieTenue;

   blanc := StringOfChar(' ',35);
   NatureAssiette :='';
   QualifiantCotis :='';
   ConditionSpe :='';
   Institution :='';
   GrpCotisant :='';
   CategSocio :='';
   NoLigne :='';
   DateChgtTx :='';
   Taux :='0';
   Cotisation:='0';
   Base :='0';
   forfait :='0';
   EffCotisation :='0';
   EffForfait :='0';

  if TOB_Lignes <> NIL then
    begin
     LesLignes := TOB_Lignes.FindFirst ([''],[''],TRUE);
     While LesLignes <> NIL  do
      begin
        if (Leslignes.getValue('PDD_TYPECOTISATION') <> 'I') and
           (Leslignes.getValue('PDD_TYPECOTISATION') <> 'S') then
        // les lignes Intitulé et Sous-Total ne sont pas traitées.
        begin
          NoLg := NoLg+1;

          NatureAssiette := copy(Leslignes.getValue('PDD_CODIFICATION'),6,2);
          DebCodif := copy(Leslignes.getValue('PDD_CODIFICATION'),1,2);
          SuitCodif := copy(Leslignes.getValue('PDD_CODIFICATION'),3,5);
          while (((QualifiantCotis = '') or
                  (QualifiantCotis = NULL) or
                  (QualifiantCotis = '   ')) and (StrToInt(copy(DebCodif,2,1)) >= 0)) do
          begin
            QQ := OpenSql('SELECT PDP_PREDEFINI,PDP_COTISQUAL'+
                          ' FROM DUCSPARAM '+
                          ' WHERE ##PDP_PREDEFINI## PDP_CODIFICATION = "'+
                          DebCodif +  SuitCodif +'"' +
                          'ORDER BY PDP_PREDEFINI',True);
            While  not QQ.eof do
            begin
              QualifiantCotis :=QQ.FindField('PDP_COTISQUAL').AsString;  
              if (QQ.FindField('PDP_PREDEFINI').AsString = 'DOS') then break;
              QQ.Next;
            end;
            ferme (QQ);
            DebCodif := IntToStr(StrToInt(DebCodif)-1);
          end;

          ConditionSpe :=Leslignes.getValue('PDD_CONDITION');

          Institution :=Leslignes.getValue('PDD_INSTITUTION');

          GrpCotisant := copy(Leslignes.getValue('PDD_CODIFICATION'),3,3);

          CategSocio :='';

          NoLigne := IntToStr(NoLg);

          if (Leslignes.getValue('PDD_DATECHGTTAUX') = IDate1900) then
           DateChgtTx := '0'
          else
           DateChgtTx := Copy(DateToStr(Leslignes.getValue('PDD_DATECHGTTAUX')),7,10)+
                         Copy(DateToStr(Leslignes.getValue('PDD_DATECHGTTAUX')),4,2)+
                         Copy(DateToStr(Leslignes.getValue('PDD_DATECHGTTAUX')),1,2);

          if (Leslignes.getValue('PDD_TYPECOTISATION')<> 'M') and
             (Leslignes.getValue('PDD_TYPECOTISATION')<> 'Q') then
           // ce n'est pas un forfait
           Taux :=formatFloat('#0',Leslignes.getValue('PDD_TAUXCOTISATION')*10000);
           if (taux <> '0') then // PT30 DUCS EDI V4.2
           Taux := ColleZeroDevant(StrToInt(Taux),8);      // PT30 DUCS EDI V4.2

           Cotisation:=FormatFloat('#0',Leslignes.getValue('PDD_MTCOTISAT')*100);

          if (Leslignes.getValue('PDD_TYPECOTISATION')<> 'M') and
             (Leslignes.getValue('PDD_TYPECOTISATION')<> 'Q') then
           // ce n'est pas un forfait
          Base :=FormatFloat('#0',Leslignes.getValue('PDD_BASECOTISATION')*100);

          if (Leslignes.getValue('PDD_TYPECOTISATION') = 'M') or
             (Leslignes.getValue('PDD_TYPECOTISATION') = 'Q') then
          // forfait
          begin
             // Montant cotisation arrondi/effectif
              forfait := FormatFloat('#0',
                        ((StrToFloat(FormatFloat('#0',Leslignes.getValue('PDD_MTCOTISAT')*100)))/
                        Leslignes.getValue('PDD_EFFECTIF')));
          end;

// d PT30 DUCS EDI V4.2
//        if (Leslignes.getValue('PDD_TYPECOTISATION') <> 'M') and
          if (Leslignes.getValue('PDD_TYPECOTISATION') <> 'Q') then
            // ce n'est pas un forfait
            EffCotisation := Leslignes.getValue('PDD_EFFECTIF');
// f PT30 DUCS EDI V4.2
          if (Leslignes.getValue('PDD_TYPECOTISATION') = 'M') or
             (Leslignes.getValue('PDD_TYPECOTISATION') = 'Q') then
            // forfait
            EffForfait := IntToStr(StrToint(Leslignes.getValue('PDD_EFFECTIF')) * 10000);

          Writeln(FDUCSEDI,'AGI   0'+
                           NatureAssiette+
                           QualifiantCotis+
                           ConditionSpe+
                           Institution+
                           GrpCotisant+
                           CategSocio+
                           copy(blanc,1,(4-length(CategSocio)))+
                           NoLigne+
                           copy(blanc,1,(6-length(NoLigne)))+
                           DateChgtTx+
                           copy(blanc,1,(8-length(DateChgtTx)))+
                           Taux+
                           copy(blanc,1,(8-length(Taux)))+
                           Cotisation+
                           copy(blanc,1,(18-length(Cotisation)))+
                           Base+
                           copy(blanc,1,(18-length(Base)))+
                           forfait+
                           copy(blanc,1,(18-length(forfait)))+
                           EffCotisation+
                           copy(blanc,1,(18-length(EffCotisation)))+
                           EffForfait+
                           copy(blanc,1,(18-length(EffForfait)))+
                           EnregDucsEdi.MonnaieTenue);

          EnregDucsEdi.NbCot := EnregDucsEdi.NbCot+1;

          NatureAssiette :='';
          QualifiantCotis :='';
          ConditionSpe :='';
          Institution :='';
          GrpCotisant :='';
          CategSocio :='';
          NoLigne :='';
          DateChgtTx :='';
          Taux :='0';
          Cotisation:='0';
          Base :='0';
          forfait :='0';
          EffCotisation :='0';
          EffForfait :='0';
        end;
        LesLignes := TOB_Lignes.FindNext ([''],[''],TRUE);
      end;
    end;
end;
// fin Groupe3IRC

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 13/05/2002
Modifié le ... :   /  /
Description .. : Traitement des segments du groupe 3 de la déclaration
Suite ........ : Spécificité ASSEDIC
Mots clefs ... : PAIE, PGDUCSEDI
*****************************************************************}
procedure Groupe3UNEDIC(var EnregDucsEdi : TDucsEdi);
var
   blanc, NatureAssiette, DateD, DateF, Taux, Base, Forfait : string;
   LesLignes, LaLigneSuivante                               : TOB;
   i                                                        : integer; 
begin
   blanc := StringOfChar(' ',35);
   NatureAssiette := '';
   Taux := '';
   Base := '';
   Forfait := '0';

{  PT22 }
    if TOB_Lignes <> NIL then
    begin
      i := 0;

      While (i < TOB_Lignes.Detail.Count)  do
      begin
        LesLignes := TOB_Lignes.Detail[i];
        DateD := '0';
        DateF := '0';
        if (Leslignes.getValue('PDD_TYPECOTISATION') <> 'I') and
           (Leslignes.getValue('PDD_TYPECOTISATION') <> 'S') then
        // les lignes Intitulé et Sous-Total ne sont pas traitées.
        begin
        NatureAssiette := copy(Leslignes.getValue('PDD_CODIFICATION'),5,3);

        if (Leslignes.getValue('PDD_DATECHGTTAUX') <> IDate1900) then
        begin
          DateD := Copy(DateToStr(Leslignes.getValue('PDD_DATECHGTTAUX')),7,10)+
                   Copy(DateToStr(Leslignes.getValue('PDD_DATECHGTTAUX')),4,2)+
                   Copy(DateToStr(Leslignes.getValue('PDD_DATECHGTTAUX')),1,2);
          DateF := Copy(DateToStr(EnregDucsEdi.DateFin),7,10)+
                   Copy(DateToStr(EnregDucsEdi.DateFin),4,2)+
                   Copy(DateToStr(EnregDucsEdi.DateFin),1,2);

          if (i+1 < TOB_Lignes.Detail.Count) then
          begin
            LaLigneSuivante := TOB_Lignes.Detail[i+1];
            if(LaLigneSuivante <> NIL) then
            begin
              if (copy(LaLigneSuivante.getValue('PDD_CODIFICATION'),5,3) =
                  NatureAssiette) then
                if (LaLigneSuivante.getValue('PDD_DATECHGTTAUX') <> IDate1900) then
                  DateF := copy(DateToStr(PlusDate(LaLigneSuivante.getValue('PDD_DATECHGTTAUX'),-1,'J')),7,4)+
                           copy(DateToStr(PlusDate(LaLigneSuivante.getValue('PDD_DATECHGTTAUX'),-1,'J')),4,2)+
                           copy(DateToStr(PlusDate(LaLigneSuivante.getValue('PDD_DATECHGTTAUX'),-1,'J')),1,2)
                else
                  DateF := copy(DateToStr(EnregDucsEdi.DateFin),7,4)+
                           copy(DateToStr(EnregDucsEDi.DateFin),1,2)+
                           copy(DateToStr(EnregDucsEDi.DateFin),4,2);
            end;
          end;
        end;
        Taux := FormatFloat ('#0',(Leslignes.getValue('PDD_TAUXCOTISATION')*10000));
        Taux := ColleZeroDevant(StrToInt(Taux),8);      //PT30 DUCS EDI V4.2
        Base := FormatFloat ('#0',Leslignes.getValue('PDD_BASECOTISATION'));
        Forfait := '0'; // facultatif --> !! Non Traité

        Writeln(FDUCSEDI,'CUN   '+
                         NatureAssiette+
                         DateD+
                         copy(blanc,1,(8-length(DateD)))+
                         DateF+
                         copy(blanc,1,(8-length(DateF)))+
                         Taux+
                         copy(blanc,1,(8-length(Taux)))+
                         Base+
                         copy(blanc,1,(18-length(Base)))+
                         Forfait+
                         copy(blanc,1,(18-length(Forfait)))+
                         EnregDucsEdi.MonnaieTenue);

        EnregDucsEdi.NbCot := EnregDucsEdi.NbCot+1;

        NatureAssiette := '';
        Taux := '';
        Base := '';
        Forfait := '0';
        end;
        i := i+1;
      end;
    end;
end;
// fin Groupe3UNEDIC
{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 13/05/2002
Modifié le ... :   /  /
Description .. : Traitement des segments de fin de déclaration
Mots clefs ... : PAIE, PGDUCSEDI
*****************************************************************}
procedure FinDeclaration(var EnregDucsEdi : TDucsEdi;EnrDucsEdiGroupee : TDucsEdiGroupee);
var
   blanc,Montant             : string;
   tampon                    : string;    // PT30 DUCS EDI V4.2
begin
     blanc := StringOfChar(' ',35);

     TOB_Lignes.Free; TOB_Lignes := NIL;

     if (TOB_Transport <> Nil) then
     begin
       TOB_Transport.Free;
       TOB_Transport := NIL;
     end;
//d memcheck
   if (TOB_TAT <> Nil) then
    begin
      TOB_TAT.Free;
      TOB_TAT := NIL;
    end;
// f memcheck
//PT34     TOB_DucsGroupee.Free; TOB_DucsGroupee := NIL;
         FreeAndNil(TOB_DucsGroupee);
     // UNS   S - Délimiteur de section de message
     // ------------------------------------------
     Writeln(FDUCSEDI,'UNS   S');

     // MOATOT - Montant Total
     // ----------------------
     if (copy(EnregDucsEdi.TypDucs,1,1) = 'I')  then
      // IRC  : monatnt déclaré  en cts
      Montant := FormatFloat ('#0',(EnregDucsEdi.MtTotal*100))
     else
       // ACOSS ou UNEDIC  : montant payé à l'euro ( ou au franc)
       // Pour l'ACOSS : ne concerne que les déclarations multi établissements.
       // Pour L'UNEDIC ne concerne que les ADV et AVA
{PT32       if ((EnregDucsEdi.DucsDossier = TRUE) and
          (EnregDucsEdi.PaiementGroupe = TRUE)) or
          (copy(EnregDucsEdi.TypDucs,1,1) = 'U')  then}
       if ((copy(EnregDucsEdi.TypDucs,1,1) = 'U') or (copy(EnregDucsEdi.TypDucs,1,1) = 'A'))   then
         Montant := FormatFloat ('#0',EnregDucsEdi.MtTotal);
// PT30 DUCS EDI V4.2
     if   (copy(EnregDucsEdi.TypDucs,1,1) = 'A') and
           ((EnregDucsEdi.TypBordereau = '915') or (EnregDucsEdi.TypBordereau = '916')) then
         Montant := FormatFloat ('#0',EnregDucsEdi.MtTotal+EnregDucsEdi.MtTotRegul);
// PT30 DUCS EDI V4.2
{PT32     if ((EnregDucsEdi.DucsDossier = TRUE) and
          (EnregDucsEdi.PaiementGroupe = TRUE) and
          (copy(EnregDucsEdi.TypDucs,1,1) = 'A')) or    }
     if ((copy(EnregDucsEdi.TypDucs,1,1) = 'A')) or
// PT30 DUCS EDI V4.2
//        (copy(EnregDucsEdi.TypDucs,1,1) <> 'A') then
          (copy(EnregDucsEdi.TypDucs,1,1) = 'I') or
          ((copy(EnregDucsEdi.TypDucs,1,1) = 'U') and
           (EnregDucsEdi.TypBordereau = '920') or
           (EnregDucsEdi.TypBordereau = '921') or
           (EnregDucsEdi.TypBordereau = '924') or
           (EnregDucsEdi.TypBordereau = '926')) then
        Writeln(FDUCSEDI,'MOATOT'+
                         Montant+
                         copy(blanc,1,(18-length(Montant)))+
                         EnregDucsEdi.MonnaieTenue);
      Writeln(FRapport,''); // PT30 DUCS EDI V4.2
      Writeln(FRapport,'    Montant total des cotisations : '+FormatFloat ('#0',EnregDucsEdi.MtTotal)+' '+EnregDucsEdi.MonnaieTenue);
      Writeln(FRapport,'');
// d PT30 DUCS EDI V4.2
    // MOAMSA - Masses salariales annuelles
    //  pour DRA UNEDIC (922 ou 923)
    // ------------------------------------
    if ((copy(EnregDucsEdi.TypDucs,1,1) = 'U') and
        ((EnregDucsEdi.TypBordereau = '922') or
         (EnregDucsEdi.TypBordereau = '923'))) then
    //  Unedic mt à l'euro
    begin
      tampon :=FormatFloat('#0',EnregDucsEdi.MasseAnnuel);
      Writeln(FDUCSEDI,'MOAMSA'+
                       tampon+
                       copy(blanc,1,(18-length(tampon)))+
                       EnregDucsEdi.MonnaieTenue);
    end;
// f PT30 DUCS EDI V4.2

      // MOAREG - Montant de la régularisation (IRC seulement)
      // -----------------------------------------------------
     if (copy(EnregDucsEdi.TypDucs,1,1) = 'I') then
        begin
          if (EnregDucsEdi.MtTotRegul <> 0.0) then
            begin
              Montant := Formatfloat('#0',(EnregDucsEdi.MtTotRegul*100));
              Writeln (FDUCSEDI,'MOAREG'+
                                Montant+
                                copy(blanc,1,(18-length(Montant)))+
                                EnregDucsEdi.MonnaieTenue);
              Writeln(FRapport,'    Montant de la régularisation : '+Formatfloat('#0',EnregDucsEdi.MtTotRegul)+' '+EnregDucsEdi.MonnaieTenue);
            end;
        end;
      // MOAACP - Montant de l'acompte (IRC seulement)
      // -----------------------------------------------------
     if (copy(EnregDucsEdi.TypDucs,1,1) = 'I') then
        begin
          if (EnregDucsEdi.MtTotAcpte <> 0.0) then
            begin
              Montant := Formatfloat('#0',(EnregDucsEdi.MtTotAcpte*100));
              Writeln (FDUCSEDI,'MOAACP'+
                                Montant+
                                copy(blanc,1,(18-length(Montant)))+
                                EnregDucsEdi.MonnaieTenue);
              Writeln(FRapport,'    Montant de l''acompte : '+Formatfloat('#0',EnregDucsEdi.MtTotAcpte)+' '+EnregDucsEdi.MonnaieTenue);
            end;
        end;
end;
// fin FinDeclaration
{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 13/05/2002
Modifié le ... :   /  /
Description .. : Traitement des segments liés au paiement
Mots clefs ... : PAIE, PGDUCSEDI
*****************************************************************}
procedure PaiementEdi(var EnregDucsEdi : TDucsEdi;EnrDucsEdiGroupee : TDucsEdiGroupee; Trait :string; var DossOK : boolean; TMtsEtabs : TOB); // PT37
var
   intervenant,Montant, blanc                     : string;
   tampon,creancier, ModePaiement, NumeroInterne  : string;
   paiement                                       : TdateTime ;
   CodeBanque ,NoCpte, CleRib, NoBanque,NoGuichet : string;
   QQ                                             : TQuery;
   StQQ, LibPaiement, LibMt,st, StPlus            : string;
   typenr                                         : string;
   LBanque, LDom                                  : String;
   TitulaireCpt                                   : string; // PT30 DUCS EDI V4.2
   II                                             : integer; // PT37
   TMtsEtabsFille                                 : TOB ; // PT37
begin
     CodeBanque:= '';

     if (EnregDucsEdi.Neant = True) or
        ((Trait <> 'G') and (EnregDucsEdi.MtApayer = 0)) or
        ((Trait = 'G') and (EnrDucsEdiGroupee.MtApayer = 0)) then
       // Ducs Néant ou Montant à payer = 0 - pas de paiement
       begin
        st := '  Etablissement : ';
        if (trait <> 'G') then
          st := st + EnregDucsEdi.Etab+' '+EnregDucsEdi.LibelleET
        else
          st := st + EnrDucsEdiGroupee.Etab+' '+EnrDucsEdiGroupee .LibelleET ;

        Writeln(FRapport, st);
        if (EnregDucsEdi.Neant = True) then
          Writeln(FRapport, '    Ducs Néant')
        else
         begin
           // Ducs non néant mais montant à payer = 0
           if (Trait <> 'G') then
             begin
               tampon :=FormatFloat('#0',EnregDucsEdi.Acompte);
             end
           else
             if (Trait = 'G') then
               begin
                 tampon :=FormatFloat('#0',EnrDucsEdiGroupee.Acompte);
               end;
            Writeln(FRapport, '    Montant acompte        : '+tampon + ' ' + EnregDucsEdi.MonnaieTenue);

           if (Trait <> 'G') then
             begin
               tampon :=FormatFloat('#0',EnregDucsEdi.Regularisation);
             end
           else
             if (Trait = 'G') then
               begin
                 tampon :=FormatFloat('#0',EnrDucsEdiGroupee.Regularisation);
               end;
            Writeln(FRapport, '    Montant régularisation : '+tampon + ' ' + EnregDucsEdi.MonnaieTenue);
            Writeln(FRapport, '    Montant à payer        : 0'+ ' ' + EnregDucsEdi.MonnaieTenue);

             if ((copy(EnregDucsEdi.TypDucs,1,1) = 'A') and
            (EnregDucsEdi.DucsDossier = TRUE) and
            (EnregDucsEdi.PaiementGroupe = TRUE)) or
            ((copy(EnregDucsEdi.TypDucs,1,1) = 'I') and
            (EnregDucsEdi.DucsDossier = TRUE) {and
            (EnregDucsEdi.PaiementGroupe = TRUE)}) or // PT52
            ((copy(EnregDucsEdi.TypDucs,1,1) = 'U') and
            (EnregDucsEdi.DucsDossier = TRUE) and
            (EnregDucsEdi.PaiementGroupe = TRUE)) then
            begin
// d PT44
              if (TMtsEtabs <> nil) then
              begin
// f PT44
              for II := 0 to TMtsEtabs.Detail.Count-1 do
              begin
                TMtsEtabsFille := TMtsEtabs.Detail[II];

                st := '  Etablissement : ';
                st := st + TMtsEtabsFille.GetValue('ETAB')+' '+TMtsEtabsFille.GetValue('LibelleET');

                Writeln(FRapport, st);

                tampon := FormatFloat('#0',TMtsEtabsFille.GetValue('ACOMPTE'));
                Writeln(FRapport, '    Montant acompte        : '+tampon + ' ' + EnregDucsEdi.MonnaieTenue);

                tampon := FormatFloat('#0',TMtsEtabsFille.GetValue('REGUL'));
                Writeln(FRapport, '    Montant régularisation : '+tampon + ' ' + EnregDucsEdi.MonnaieTenue);

                Writeln(FRapport, '    Montant à payer        : 0'+ ' ' + EnregDucsEdi.MonnaieTenue);
              end;
              end; //PT44
            end;


        end;
        exit;
       end;

     tampon := '';
     creancier := '';
     TitulaireCpt := '';  // PT30 DUCS EDI V4.2
     blanc := StringOfChar(' ',35);

     if (Trait <> 'G') then
      begin
// d PT37
        st := '  Etablissement : ';
        st := st + EnregDucsEdi.Etab+' '+EnregDucsEdi.LibelleET  ;
        Writeln(FRapport, st);

        tampon :=FormatFloat('#0.00',EnregDucsEdi.Acompte);
        Writeln(FRapport, '    Montant acompte        : '+tampon + ' ' + EnregDucsEdi.MonnaieTenue);

        tampon :=FormatFloat('#0.00',EnregDucsEdi.Regularisation);
        Writeln(FRapport, '    Montant régularisation : '+tampon + ' ' + EnregDucsEdi.MonnaieTenue);

// PT37        tampon :=FormatFloat('#0.00',EnregDucsEdi.MtApayer);
            tampon := FormatFloat('#0.00',EnregDucsEdi.MtDeclare-EnregDucsEdi.Acompte+EnregDucsEdi.Regularisation);
        Writeln(FRapport, '    Montant à payer        : '+tampon + ' ' + EnregDucsEdi.MonnaieTenue);

        if ((copy(EnregDucsEdi.TypDucs,1,1) = 'A') and
           (EnregDucsEdi.DucsDossier = TRUE) and
           (EnregDucsEdi.PaiementGroupe = TRUE)) or
           ((copy(EnregDucsEdi.TypDucs,1,1) = 'I') and
           (EnregDucsEdi.DucsDossier = TRUE) {and
           (EnregDucsEdi.PaiementGroupe = TRUE)}) or            // PT52
           ((copy(EnregDucsEdi.TypDucs,1,1) = 'U') and
           (EnregDucsEdi.DucsDossier = TRUE) and
           (EnregDucsEdi.PaiementGroupe = TRUE)) then
        begin
// d PT44
          if (TMtsEtabs <> nil) then
          begin
// f PT44
          for II := 0 to TMtsEtabs.Detail.Count-1 do
          begin
            TMtsEtabsFille := TMtsEtabs.Detail[II];

            st := '  Etablissement : ';
            st := st + TMtsEtabsFille.GetValue('ETAB')+' '+TMtsEtabsFille.GetValue('LibelleET');

            Writeln(FRapport, st);

            tampon := FormatFloat('#0.00',TMtsEtabsFille.GetValue('ACOMPTE'));
            Writeln(FRapport, '    Montant acompte        : '+tampon + ' ' + EnregDucsEdi.MonnaieTenue);

            tampon := FormatFloat('#0.00',TMtsEtabsFille.GetValue('REGUL'));
            Writeln(FRapport, '    Montant régularisation : '+tampon + ' ' + EnregDucsEdi.MonnaieTenue);

            tampon := FormatFloat('#0.00',TMtsEtabsFille.GetValue('APAYER'));
            Writeln(FRapport, '    Montant à payer        : '+tampon + ' ' + EnregDucsEdi.MonnaieTenue);
          end;
          end; //PT44
        end;

// f PT37
        ModePaiement := EnregDucsEdi.ModePaiement;
        Montant := FormatFloat ('#0',EnregDucsEdi.MtApayer);
        if (copy(EnregDucsEdi.TypDucs,1,1) = 'I')  then
          Montant := FormatFloat ('#0',EnregDucsEdi.MtApayer * 100);
        LibMt := FormatFloat ('#0.00',EnregDucsEdi.MtApayer)+' '+
                              EnregDucsEdi.MonnaieTenue;
        NumeroInterne := EnregDucsEdi.NumeroInterne;
        paiement := EnregDucsEdi.reglement ;
// d PT30 DUCS EDI V4.2
// d PT37
{        if (ModePaiement = 'Z10') then
          TitulaireCpt := EnregDucsEdi.TitulaireCpt;}
        if (((ModePaiement = 'Z10') and (copy(EnregDucsEdi.TypDucs,1,1) <> 'I')) or
            (((ModePaiement = 'Z10') or
              (ModePaiement = '30') or
              (ModePaiement = '31')) and
             (copy(EnregDucsEdi.TypDucs,1,1) = 'I'))) then
// f PT37             
          TitulaireCpt := EnregDucsEdi.TitulaireCpt;
// f PT30 DUCS EDI V4.2
      end
     else
      begin
// d PT37
        if ((copy(EnregDucsEdi.TypDucs,1,1) = 'A') and
           (EnregDucsEdi.DucsDossier = TRUE) and
           (EnregDucsEdi.PaiementGroupe = TRUE)) or
           ((copy(EnregDucsEdi.TypDucs,1,1) = 'I') and
           (EnregDucsEdi.DucsDossier = TRUE) {and
           (EnregDucsEdi.PaiementGroupe = TRUE)}) or // PT52
           ((copy(EnregDucsEdi.TypDucs,1,1) = 'U') and
           (EnregDucsEdi.DucsDossier = TRUE) and
           (EnregDucsEdi.PaiementGroupe = TRUE)) then
        begin
// d PT44
{          if (TMtsEtabs <> nil) then
          begin
// f PT44
          for II := 0 to TMtsEtabs.Detail.Count-1 do
          begin
            TMtsEtabsFille := TMtsEtabs.Detail[II];

            st := '  Etablissement : ';
            st := st + TMtsEtabsFille.GetValue('ETAB')+' '+TMtsEtabsFille.GetValue('LibelleET');

            Writeln(FRapport, st);

            tampon := FormatFloat('#0',TMtsEtabsFille.GetValue('ACOMPTE'));
            Writeln(FRapport, '    Montant acompte        : '+tampon + ' ' + EnregDucsEdi.MonnaieTenue);

            tampon := FormatFloat('#0',TMtsEtabsFille.GetValue('REGUL'));
            Writeln(FRapport, '    Montant régularisation : '+tampon + ' ' + EnregDucsEdi.MonnaieTenue);

            tampon := FormatFloat('#0.00',TMtsEtabsFille.GetValue('APAYER'));
            Writeln(FRapport, '    Montant à payer        : 0'+ ' ' + EnregDucsEdi.MonnaieTenue);
          end;
          end; // PT44  }
        end;
// f PT37
        ModePaiement := EnrDucsEdiGroupee.ModePaiement;
        Montant := FormatFloat ('#0',EnrDucsEdiGroupee.MtApayer);
        if (copy(EnregDucsEdi.TypDucs,1,1) = 'I')  then
          Montant := FormatFloat ('#0',EnrDucsEdiGroupee.MtApayer * 100);
        LibMt := FormatFloat ('#0.00',EnrDucsEdiGroupee.MtApayer)+' '+
                              EnrDucsEdiGroupee.MonnaieTenue;
        NumeroInterne := EnrDucsEdiGroupee.NumeroInterne;
        paiement := EnregDucsEdi.reglement;
// d PT30 DUCS EDI V4.2
        if (ModePaiement = 'Z10') then
          TitulaireCpt := EnrDucsEdiGroupee.TitulaireCpt;
// f PT30 DUCS EDI V4.2
      end;

    if (ModePaiement = '20') then libPaiement := 'Chèque';
    if (ModePaiement = '30') then libPaiement := 'Virement';
    if (ModePaiement = '31') then libPaiement := 'Prélèvement';
    if (ModePaiement = 'Z10') then libPaiement := 'Télé-Règlement EDI';
    st := 'PAIEMENT : Etablissement : ';       // PT30 DUCS EDI V4.2
     if (trait <> 'G') then
       st := st + EnregDucsEdi.Etab+' '+EnregDucsEdi.LibelleET
     else
       st := st + EnrDucsEdiGroupee.Etab+' '+EnrDucsEdiGroupee .LibelleET ;

    Writeln(FRapport, '');       // PT30 DUCS EDI V4.2
    Writeln(FRapport, st);
    Writeln(FRapport, '    Paiement par '+Libpaiement+'  '+LibMT);

     // PAI   - Mode de paiement
     //-------------------------
    typenr := 'PAI   ';
    if (copy(EnregDucsEdi.TypDucs,1,1) = 'I') and (ModePaiement = '20') then
      typenr := 'PAIC  ';
    if (copy(EnregDucsEdi.TypDucs,1,1) = 'I') and (ModePaiement = '30') then
      typenr := 'PAIV  ';
    if (copy(EnregDucsEdi.TypDucs,1,1) = 'I') and (ModePaiement = '31') then
      typenr := 'PAIP  ';
    if (copy(EnregDucsEdi.TypDucs,1,1) = 'I') and (ModePaiement = 'Z10') then
      typenr := 'PAIZ  ';
     Writeln (FDUCSEDI,typenr+
                         ModePaiement+
                         copy(blanc,1,(3-length(ModePaiement))));

     // FII   - Identification Financière
     // ---------------------------------
    typenr := 'FII   ';
    if (copy(EnregDucsEdi.TypDucs,1,1) = 'I') and (ModePaiement = '30') then
      typenr := 'FIIV  ';
    if (copy(EnregDucsEdi.TypDucs,1,1) = 'I') and (ModePaiement = '31') then
      typenr := 'FIIP  ';
    if (copy(EnregDucsEdi.TypDucs,1,1) = 'I') and (ModePaiement = 'Z10') then
      typenr := 'FIIZ  ';

     if (ModePaiement = '30') or
        (ModePaiement = '31') or
        (ModePaiement = 'Z10') then
       // Virement, Prélèvement, Télérèglement EDI
       begin
         if (ModePaiement = '30') then
           // Virement - bénéficiaire
           begin
             intervenant := 'BF ';
             StQQ:= '';
             if (Trait <> 'G') then
               StQQ := '"'+EnregDucsEdi.Etab+'"'
             else
               StQQ := '"'+EnrDucsEdiGroupee.Etab+'"';

             QQ := OpenSQL('SELECT POG_RIBDUCSEDI FROM ORGANISMEPAIE WHERE '+
                           'POG_ETABLISSEMENT='+StQQ+
                           ' AND POG_ORGANISME="'+EnregDucsEdi.Organisme+'"',True) ;
             if Not QQ.EOF then
               begin
                 CodeBanque :=QQ.FindField('POG_RIBDUCSEDI').AsString;
               end;
             Ferme(QQ) ;
             if (CodeBanque = '') then
               begin
                 Writeln(FRapport, '  Anomalie   : Virement, la'+
                             ' banque bénéficiaire n''est pas renseignée');

                 Writeln(FRapport, '  --> Action : Modification organisme : renseigner'+
                             ' la banque bénéficiaire ou bien modifier le mode de'+
                             ' règlement ');
                 DossOK := False;
                 exit;
               end ;

             QQ := OpenSQL('SELECT PBO_NUMCPTEBQORG,PBO_CLEBQORG,PBO_ETABBQORG,'+
                           'PBO_GUICHBQORG,PBO_LIBELLEBQORG,PBO_DOMBQORG '+
                           'FROM PAIEBQORG WHERE '+
                           'PBO_CODEBQORG="'+CodeBanque+'"' ,True) ;
             if Not QQ.EOF then
               begin
                 NoCpte :=QQ.FindField('PBO_NUMCPTEBQORG').AsString;
                 CleRib :=QQ.FindField('PBO_CLEBQORG').AsString;
                 NoBanque :=QQ.FindField('PBO_ETABBQORG').AsString;
                 NoGuichet :=QQ.FindField('PBO_GUICHBQORG').AsString;
                 LBanque := QQ.FindField('PBO_LIBELLEBQORG').AsString;
                 LDom   := QQ.FindField('PBO_DOMBQORG').AsString;
               end ;
             Ferme(QQ) ;
           end
         else
           // donneur d'ordre - prélèvement ou télérèglement EDI
           begin
           intervenant := 'OR ';
           StQQ:= '';
           if (Trait <> 'G') then
             StQQ := '"'+EnregDucsEdi.Etab+'"'
           else
             StQQ := '"'+EnrDucsEdiGroupee.Etab+'"';
             QQ := OpenSQL('SELECT POG_RIBDUCSEDI FROM ORGANISMEPAIE WHERE '+
                           'POG_ETABLISSEMENT='+StQQ+
                           ' AND POG_ORGANISME="'+EnregDucsEdi.Organisme+'"',True) ;
             if Not QQ.EOF then
               begin
                 CodeBanque :=QQ.FindField('POG_RIBDUCSEDI').AsString;
               end;
             Ferme(QQ) ;

           if (CodeBanque = '') then
              begin
                Writeln(FRapport, '  Anomalie   : pélèvement ou télérèglement, la'+
                             ' banque du donneur d''ordre n''est pas renseignée');

                Writeln(FRapport, '  --> Action : Modification établissement social : renseigner'+
                             ' l''identifiant bancaire pour le paiement des charges sociale ou bien modifier le mode de'+
                             ' règlement ');
                DossOK := False;
                exit;
              end;

           StPlus:= PGBanqueCP (True);   //PT38

           QQ := OpenSQL('SELECT BQ_NUMEROCOMPTE,BQ_CLERIB,BQ_ETABBQ,'+
                         'BQ_GUICHET,BQ_LIBELLE,BQ_DOMICILIATION '+
                         'FROM BANQUECP WHERE '+
                         'BQ_GENERAL="'+CodeBanque+'"'+StPlus ,True) ;
           if Not QQ.EOF then
            begin
             NoCpte :=QQ.FindField('BQ_NUMEROCOMPTE').AsString;
             CleRib :=QQ.FindField('BQ_CLERIB').AsString;
             NoBanque :=QQ.FindField('BQ_ETABBQ').AsString;
             NoGuichet :=QQ.FindField('BQ_GUICHET').AsString;
             LBanque := QQ.FindField('BQ_LIBELLE').AsString;
             LDom   := QQ.FindField('BQ_DOMICILIATION').AsString;

            end ;
           Ferme(QQ) ;
           end;

           Writeln (FDUCSEDI,typenr+
                             intervenant+
                             NoCpte+
                             CleRib+
// PT30 DUCS EDI V4.2
//                           copy(blanc,1,(35-(length(NoCpte)+length(CleRib))))+
                             copy(blanc,1,(17-(length(NoCpte)+length(CleRib))))+
                             NoBanque+
                             copy(blanc,1,(11-(length(NoBanque))))+
                             NoGuichet+
                             copy(blanc,1,(17-(length(NoGuichet))))+
// PT30 DUCS EDI V4.2
                             TitulaireCpt+
                             copy(blanc,1,(35-(length(TitulaireCpt)))));
// PT30 DUCS EDI V4.2
          if (ModePaiement = '30') then
            Writeln(FRapport, '    RIB du bénéficiaire : '+NoBanque+
                            '  '+NoGuichet+'  '+NoCpte+'  '+CleRib)
          else
            Writeln(FRapport, '    RIB du donneur d''ordre : '+NoBanque+
                            '  '+NoGuichet+'  '+NoCpte+'  '+CleRib);
          Writeln(FRapport,'                              '+LBanque+'  '+LDom);

       end;

       // MOAPAI - Montant réglé
       // ----------------------
    typenr := 'MOAPAI';
    if (copy(EnregDucsEdi.TypDucs,1,1) = 'I') and (ModePaiement = '20') then
      typenr := 'MOAPAC';
    if (copy(EnregDucsEdi.TypDucs,1,1) = 'I') and (ModePaiement = '30') then
      typenr := 'MOAPAV';
    if (copy(EnregDucsEdi.TypDucs,1,1) = 'I') and (ModePaiement = '31') then
      typenr := 'MOAPAP';
    if (copy(EnregDucsEdi.TypDucs,1,1) = 'I') and (ModePaiement = 'Z10') then
      typenr := 'MOAPAZ';
       Writeln(FDUCSEDI,typenr+
                        Montant+
                        copy(blanc,1,(18-length(Montant)))+
                        EnregDucsEdi.MonnaieTenue);

       // RFFqual - Référence sociale (chèque, virement, télérèglement EDI)
       // -----------------------------------------------------------------
       if (ModePaiement <> '31') and
          (ModePaiement <> 'Z11') then
         begin
           if (ModePaiement = '20') or
              (ModePaiement = '30') then
              // Chèque ou virement
              begin
               typenr := 'RFFPQ ';
               if (copy(EnregDucsEdi.TypDucs,1,1) = 'I') and (ModePaiement = '20') then
                 typenr := 'RFFPQC';
               if (copy(EnregDucsEdi.TypDucs,1,1) = 'I') and (ModePaiement = '30') then
                 typenr := 'RFFPQV';

               EnregDucsEDi.CleIBAN :='00';
// pour test       "281199923999999999999990000000" --> clé 07
// pour test CleIBAN ('S1','199923','99999999999999',EnregDucsEDi.CleIBAN);
               if (copy(EnregDucsEdi.TypDucs,1,1) = 'A') then creancier := 'S1';
               if (copy(EnregDucsEdi.TypDucs,1,1) = 'U') then creancier := 'S2';
               if (copy(EnregDucsEdi.TypDucs,1,1) = 'I') then creancier := 'S3';

               if (copy(EnregDucsEdi.TypDucs,1,1) = 'U') then
                EnregDucsEDi.CleIBAN := '99'    // toujours 99 pour UNEDIC
               else
                CleIBAN (creancier,
                        copy (DateToStr(EnregDucsEDi.DateFin),7,2)+EnregDucsEdi.Abrege,
                        copy(NumeroInterne,1,19),
                        EnregDucsEDi.CleIBAN);
               if (copy(EnregDucsEdi.TypDucs,1,1) = 'U') then
                 Writeln (FDUCSEDI, typenr+
                                  EnregDucsEDi.CleIBAN+
                                  creancier+
                                  copy(DateToStr(EnregDucsEdi.DateFin),7,2)+
                                  PeriodRef+
                                  NumeroInterne+
                                  copy(blanc,1,(35-(8+length(copy(NumeroInterne,1,19))))))
               else
                 Writeln (FDUCSEDI, typenr+
                                  EnregDucsEDi.CleIBAN+
                                  creancier+
                                  copy(DateToStr(EnregDucsEdi.DateFin),7,2)+
                                  EnregDucsEdi.Abrege+
                                  NumeroInterne+
                                  copy(blanc,1,(35-(8+length(copy(NumeroInterne,1,19))))));
               if (ModePaiement = '30') then
                 // Virement
                 st :='  clé et référence sociale à renseigner dans la partie'+
                      ' référence commerciale du virement :'
               else
                 // chèque
                 st := '   clé et référence sociale à inscrire au dos du chèque:';

               Writeln (FRapport,st);
               if (copy(EnregDucsEdi.TypDucs,1,1) = 'U') then
                 Writeln (FRapport,'    Clé : '+EnregDucsEDi.CleIBAN+
                                ' Référence : '+creancier+' '+
                                copy(DateToStr(EnregDucsEdi.DateFin),7,2)+
                                PeriodRef+ ' '+
                                NumeroInterne)
               else
                 Writeln (FRapport,'    Clé : '+EnregDucsEDi.CleIBAN+
                                ' Référence : '+creancier+' '+
                                copy(DateToStr(EnregDucsEdi.DateFin),7,2)+
                                EnregDucsEdi.Abrege+ ' '+
                                NumeroInterne);

              end;
              if (ModePaiement = 'Z10') then
              // Télérèglement
              begin
                typenr := 'RFFCR ';
                if (copy(EnregDucsEdi.TypDucs,1,1) = 'I') and (ModePaiement = 'Z10') then
                  typenr :='RFFCRZ';
                tampon:=EnregDucsEdi.IdentOPS+copy(blanc,1,(8-length(EnregDucsEdi.IdentOPS)));
                if (copy(EnregDucsEdi.TypDucs,1,1) = 'U') then
                  tampon:= tampon + PeriodRef
                else
// d PT30 DUCS EDI V4.2
                  if ((VH_Paie.PGDecalage = true) or
                      (VH_Paie.PGDecalagePetit = true)) and
                     (copy(EnregDucsEdi.Abrege,3,2)='00') then
                    tampon:= tampon +
                             ColleZeroDevant((StrToInt(Copy(EnregDucsEdi.Abrege,1,2))-1),2)+
                             Copy(EnregDucsEdi.Abrege,3,2)
                  else
// f PT30 DUCS EDI V4.2
                    tampon:= tampon + EnregDucsEdi.Abrege;
                Writeln (FDUCSEDI, typenr+
                                   tampon+
                                   copy(blanc,1,(35-length(tampon))));
              end;
         end;

       // DTMPAI - Date de paiement (date de paiement renseignée quand
       //                            prélèvement (31) ou Télérèglement EDI (Z10)
       //                            pour les IRC ou Chèque (20) ou virement (30)
       //                            pour ACOSS et UNEDIC)
       // ---------------------------------------------------------------
       typenr := 'DTMPAI';
       if (copy(EnregDucsEdi.TypDucs,1,1) = 'I') and (ModePaiement = '20') then
         typenr := 'DTMPAC';
       if (copy(EnregDucsEdi.TypDucs,1,1) = 'I') and (ModePaiement = '30') then
         typenr := 'DTMPAV';
       if (copy(EnregDucsEdi.TypDucs,1,1) = 'I') and (ModePaiement = '31') then
         typenr := 'DTMPAP';
       if (copy(EnregDucsEdi.TypDucs,1,1) = 'I') and (ModePaiement = 'Z10') then
         typenr := 'DTMPAZ';
       if (paiement <> IDate1900) then
// d PT37
{          if (ModePaiement = '31') or
             (ModePaiement = 'Z10') or
             ((copy(EnregDucsEdi.TypDucs,1,1) = 'I') and
              (ModePaiement = '30')) then}
              
          if (((copy(EnregDucsEdi.TypDucs,1,1) = 'I') and
               ((ModePaiement = '31') or (ModePaiement = 'Z10'))) or
              (((copy(EnregDucsEdi.TypDucs,1,1) = 'A') or
                (copy(EnregDucsEdi.TypDucs,1,1) = 'U'))and
//d PT41
//  mauvaise lecture CdC
//  (chèque(20) ou virement(30) : inutilisé;  renseigné pour Télérèglement(Z10))
//               ((ModePaiement = '20') or (ModePaiement = '30'))))  then
                (ModePaiement = 'Z10')))  then
// f PT41
// f PT37
               Writeln (FDUCSEDI,typenr+
                               copy(DateToStr(Paiement),7,4)+
                               copy(DateToStr(Paiement),4,2)+
                               copy(DateToStr(Paiement),1,2)+
                               copy(blanc,1,(35-8)));
       Writeln (FRapport,'');
end;
// fin PaiementEdi
{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 28/05/2002
Modifié le ... :   /  /
Description .. : Traitement des derniers segments du ficheir Ducs Edi
Mots clefs ... : PAIE, PGDUCSEDI
*****************************************************************}
procedure FinFin (var EnregDucsEdi : TDucsEdi; var TMtsEtabs : TOB; SansFichierBase : boolean); // PT37 PT39
var
 blanc : string;
// d PT39
 Libelle  : string;
 CodeRetour : integer;
 Nature   : string;
// f PT39
 NoDossier : string; // PT46
begin
 if Assigned(TMtsEtabs) then FreeAndNil(TMtsEtabs); 

 blanc := StringOfChar(' ',35);
 // AUT   - Authentification déclarant
 // ----------------------------------
{* PT44
   Writeln (FDUCSEDI,'AUT   '+
                   EnregDucsEdi.LibelleET+
                   Copy(blanc,1,(35-length(EnregDucsEdi.LibelleET))));
*}

 // CNT   - Nombre de lignes de cotisation (IRC uniquemet)
 // ------------------------------------------------------
 if (copy(EnregDucsEdi.TypDucs,1,1) = 'I') then
     Writeln (FDUCSEDI, 'CNT   '+
                       copy(blanc,1,(18-length(IntToStr(EnregDucsEdi.NbCot))))+
                       IntToStr(EnregDucsEdi.NbCot));

 // UNT   1 - Fin de la déclaration
 // -------------------------------
 Writeln (FDUCSEDI,'UNT   1');

 Writeln(FRapport, '');
 Writeln(FRapport, '  fichier OK');
 CloseFile(FDUCSEDI);
// d PT39
 Libelle := EnregDucsEdi.Etab +' '+
            EnregDucsEdi.Organisme +' '+
            DateToStr(EnregDucsEdi.DateDebut)+' '+
            DateToStr(EnregDucsEdi.DateFin);

 Nature := 'DU';
 Nature :=  Nature + copy(EnregDucsEdi.TypDucs,1,1);

 NoDossier := PgRendNoDossier(); //PT46

 if (not SansFichierBase) then
 // mise en base du fichier Ducs EDI confectionné
 begin
   CodeRetour :=  AGL_YFILESTD_IMPORT (VH_Paie.PGCheminEagl+'\'+NomFicDucs,
                                       'PAIE',
                                       NomFicDucs,
                                       'DUC',
                                       'DUC',Nature, '','','',
                                       '-','-','-','-','-',
                                       'FRA','DOS', Libelle,NoDossier); // PT46
   if (CodeRetour <> -1 ) then
     PGIInfo(AGL_YFILESTD_GET_ERR(CodeRetour) + #13#10 + VH_Paie.PGCheminEagl+'\'+NomFicDucs);
 end;
// f PT39

end;
// fin FinFin
procedure MajEnvoiSocialDUCS (EnregDucsEdi : TDucsEdi; sansFichierBase : boolean); // PT39
var
   FileAttrs, Ordre : integer;
   sr : TSearchRec;
   EnregEnvoi : TEnvoiSocial;
   Size : double;
   Millesime, StQQ : string;
   QQ : Tquery;
   Guid1  : string; // PT39
   StPcl  : string; // PT39
begin
   Guid1 := '';// PT39
   Size := 0.0;
   // calcul de la taille du fichier confectionné
   FileAttrs := 0;
   FileAttrs := FileAttrs + faAnyFile;
// d PT30 DUCS EDI V4.2
//{$IFDEF EAGLCLIENT}
   if FindFirst(VH_Paie.PgCheminEagl+'\'+NomFicDucs, FileAttrs, sr) = 0 then
//{$ELSE}
//   if FindFirst(V_PGI.DatPath+'\'+NomFicDucs, FileAttrs, sr) = 0 then
//{$ENDIF}
// f PT30 DUCS EDI V4.2
      begin
      if (sr.Attr and FileAttrs) = sr.Attr then
         Size := Arrondi(sr.Size/1024, 2);
      SysUtils.FindClose(sr);
      end;

   // formatage du millésime
   Millesime := Copy(DateToStr(EnregDucsEdi.DateFin),7,4);
   StQQ:= 'SELECT CO_CODE'+
            ' FROM COMMUN WHERE'+
            ' CO_TYPE="PGA" AND'+
            ' CO_LIBELLE="'+Millesime+'"';
   QQ:=OpenSQL(StQQ,TRUE) ;
   if Not QQ.EOF then
      Millesime := QQ.FindField ('CO_CODE').AsString;
   Ferme (QQ);

   // suppression des éléments déjà existants dans ENVOISOCIAL
   StQQ := 'DELETE FROM ENVOISOCIAL WHERE'+
           ' PES_TYPEMESSAGE= "DUC" AND '+
           ' PES_MILLESSOC = "'+Millesime+'" AND'+
           ' PES_DATEDEBUT = "'+UsDateTime(EnregDucsEdi.DateDebut)+'" AND'+
           ' PES_DATEFIN = "'+UsDateTime(EnregDucsEdi.DateFin)+'" AND'+
           ' PES_SIRETDO = "'+EnregDucsEdi.Siret+'" AND'+
           ' PES_FRACTIONDADS = "Z" AND'+
           ' PES_FICHIERRECU = "'+ExtractFileName(NomFicDucs)+'"';
   ExecuteSQL(StQQ) ;

   // détermination du Chrono
   StQQ := 'SELECT MAX(PES_CHRONOMESS) AS MAXI FROM ENVOISOCIAL';
   QQ:=OpenSQL(StQQ,TRUE) ;
   if Not QQ.EOF then
      Ordre := QQ.FindField ('MAXI').AsInteger+1
   else
      Ordre := 1;
   Ferme(QQ);
// d PT39
// Récupération du GUID pour maj ENVOISOCIAL
   if (not SansFichierBase) then
   begin
     if (V_PGI.ModePcl = '1') then
       StPcl:= ' AND (YFS_PREDEFINI="DOS" AND YFS_NODOSSIER="'+V_PGI.NoDossier+'")'
     else
       StPcl:= '';
//Recherche enregistrement dans YFILESTD
     StQQ :=  'SELECT YFS_FILEGUID'+
               ' FROM YFILESTD WHERE'+
               ' YFS_NOM ="'+ NomFicDucs +'" AND'+
               ' YFS_CODEPRODUIT="PAIE" AND'+
               ' YFS_CRIT1="DUC" AND'+
               ' YFS_PREDEFINI<>"CEG"'+ StPcl; 
     QQ:=OpenSQL(StQQ,TRUE) ;
     if not QQ.EOF then
      Guid1 := QQ.FindField ('YFS_FILEGUID').AsString;
     Ferme(QQ);
   end;
// f PT39

   // création de
   ChargeTOBENVOI ;

   EnregEnvoi.Ordre := Ordre;
   EnregEnvoi.TypeE := 'DUC';
   EnregEnvoi.Millesime := Millesime;
   EnregEnvoi.Periodicite := EnregDucsEdi.Periode;
   EnregEnvoi.DateD := EnregDucsEdi.DateDebut;
   EnregEnvoi.DateF := EnregDucsEdi.DateFin;
   EnregEnvoi.Siret := EnregDucsEdi.Siret;
   EnregEnvoi.Fraction := 'Z';
   EnregEnvoi.Libelle := EnregDucsEdi.LibelleET;
   EnregEnvoi.Size := size;
   EnregEnvoi.NomFic := ExtractFileName(NomFicDucs);
   EnregEnvoi.Statut := '';
   EnregEnvoi.Monnaie := EnregDucsEdi.MonnaieTenue;
//d PT30 DUCS EDI V4.2   EnregEnvoi.CodAppli := EnregDucsEdi.CodAppliPog;
   EnregEnvoi.CodAppli := 'V42';
// PT30 DUCS EDI V4.2
   EnregEnvoi.ServUniq := EnregDucsEdi.ServUniqPog;
   EnregEnvoi.MtPaye :=  EnregDucsEdi.MtAPayer;
   EnregEnvoi.DucsNeant := EnregDucsEdi.Neant;
   EnregEnvoi.ModReglt := EnregDucsEdi.ModePaiement;
   if (Copy(EnregDucsEDi.Typducs,1,1) = 'A') then
    EnregEnvoi.Inst := 'ZDAC';
   if (Copy(EnregDucsEDi.Typducs,1,1) = 'U') then
    EnregEnvoi.Inst := 'ZDUN';
   if (Copy(EnregDucsEDi.Typducs,1,1) = 'I') then
    EnregEnvoi.Inst := 'ZDIR';

   EnregEnvoi.EmettSoc := EnregDucsEdi.EmettSoc;
   EnregEnvoi.Guid1 := Guid1; // PT39

   CreeEnvoi (EnregEnvoi);
   LibereTOBENVOI;

// d PT39
  if (not sansFichierBase) then
  // suppression du fichier sur disque
  begin
     if FileExists(VH_Paie.PgCheminEagl+'\'+NomFicDucs) then
      DeleteFile(PChar(VH_Paie.PgCheminEagl+'\'+NomFicDucs));
  end;
// f PT39
end; { fin MajEnvoiSocialDUCS}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 28/05/2002
Modifié le ... :   /  /
Description .. : Traitement des cas :
Suite ........ : ACOSS : Multi établissements avec paiement groupé
Suite ........ :               autant de  groupes 2 et groupes 3 que de
Suite ........ :               déclarations (Ducs papier)
Suite ........ : IRC : Multi établissements, multi centres payeurs
Suite ........ :          autant de  groupes 2, groupes 3
Suite ........ :          que de  déclarations (Ducs papier)
Suite ........ : UNEDIC : Multi établissements avec paiement groupé (VLU)
Suite ........ :                 Segments FZ sur groupe 1,  1 groupe 2, groupe 3
Suite ........ :                 et autant de groupes 2 supplémentaires que
Suite ........ :                d'établissements secondaires.
Mots clefs ... : PAIE, PGDUCSEDI
*****************************************************************}
// PT16-2 procedure MultiGp2Gp3(EtabPrinc : string;var EnregDucsEdi : TDucsEdi;var EnrDucsEdiGroupee : TDucsEdiGroupee; var DossOK : boolean);
procedure MultiGp2Gp3(EtabPrinc : string;var EnregDucsEdi : TDucsEdi;var EnrDucsEdiGroupee : TDucsEdiGroupee; var DossOK : boolean; var NoLg : integer; var TMtsEtabs : TOB);
var
   QL : TQuery;
   LesDucsGroupee,TOBIrcFille,TOBIrc : TOB;
begin
// d dPT30 DUCS EDI V4.2
  if (Copy(EnregDucsEdi.TypDucs,1,1) <> 'A') then
  // IRC - UNEDC
  begin
// d PT34
    if (TOB_DucsGroupee <> nil) then
       FreeAndNil(TOB_DucsGroupee);
// f PT34
    TOB_DucsGroupee := TOB.Create ('Les ducs paiement groupé', NIL, -1);

    QL :=OpenSql ('SELECT PDU_ETABLISSEMENT,PDU_ORGANISME,'+
                  'PDU_DATEDEBUT,PDU_DATEFIN,PDU_DUCSDOSSIER,PDU_NUM'+
                  ' FROM DUCSENTETE '+
                  'WHERE '+
                  '(PDU_ETABLISSEMENT <>"'+ EtabPrinc +'") AND '+
                  '(PDU_ORGANISME ="'+ EnregDucsEdi.Organisme+'") AND '+
                  '(PDU_DATEDEBUT ="'+UsDateTime(EnregDucsEdi.DateDebut)+'") AND '+
                  '(PDU_DATEFIN = "'+UsDateTime(EnregDucsEdi.DateFin)+'") AND '+
                  '(PDU_DUCSDOSSIER = "X")', TRUE);

    TOB_DucsGroupee.LoadDetailDB ('TABLEDUCSENTETE', '','',QL,FALSE,FALSE);

    Ferme(QL);

    TOB_DucsGroupee.Detail.Sort('PDU_NUM')
  end; // f PT30 DUCS EDI V4.2

  if (TOB_DucsGroupee <> NIL) then
  begin
    if (copy(EnregDucsEdi.TypDucs,1,1) = 'I') then
    // IRC
    begin
      TOBIrc := Tob.Create('Les infos paiement IRC', NIL,-1);
    end;
    LesDucsGroupee := TOB_DucsGroupee.FindFirst([''],[''],TRUE);
    EnregDucsEdi.NoCentrePayeur := 1;
    While LesDucsGroupee <> NIL  do
    begin
      InitVarGp2(EnregDucsEdi, EnrDucsEdiGroupee, LesDucsGroupee, TMtsEtabs); // PT37

      if (copy(EnregDucsEdi.TypDucs,1,1) = 'I') then
      // IRC
      begin
        TOBIrcFille := Tob.Create('', TOBIrc,-1);
        TOBIrcFille.AddChampSup('ETAB',FALSE);
        TOBIrcFille.AddChampSup('LIBELLEET',FALSE);
        TOBIrcFille.AddChampSup('MODEPAIEMENT',FALSE);
        TOBIrcFille.AddChampSup('MTAPAYER',FALSE);
        TOBIrcFille.AddChampSup('DATEFIN',FALSE);
        TOBIrcFille.AddChampSup('ABREGE',FALSE);
        TOBIrcFille.AddChampSup('NOINTERNE',FALSE);
        TOBIrcFille.AddChampSup('PAIEMENT',FALSE);
        TOBIrcFille.AddChampSup('TITULAIRECPT',FALSE);   // PT30 DUCS EDI V4.2
        TOBIrcFille.PutValue('ETAB',EnrDucsEdiGroupee.Etab);
        TOBIrcFille.PutValue('LIBELLEET',EnrDucsEdiGroupee.LibelleET);
        TOBIrcFille.PutValue('MODEPAIEMENT',EnrDucsEdiGroupee.ModePaiement);
        TOBIrcFille.PutValue('MTAPAYER',EnrDucsEdiGroupee.MtAPayer);
        TOBIrcFille.PutValue('DATEFIN',EnregDucsEdi.DateFin);
        TOBIrcFille.PutValue('ABREGE',EnrDucsEdiGroupee.Abrege);
        TOBIrcFille.PutValue('NOINTERNE',EnrDucsEdiGroupee.NumeroInterne);
        TOBIrcFille.PutValue('PAIEMENT',EnrDucsEdiGroupee.Paiement);
        TOBIrcFille.PutValue('TITULAIRECPT',EnrDucsEdiGroupee.TitulaireCpt);  // PT30 DUCS EDI V4.2
        if (copy(EnregDucsEdi.TypDucs,1,1) = 'I') and
           (EnregDucsEdi.PaiementGroupe = False) then
                  // On recalcule le Mt à payer total car dans le cas IRC Multi Centre
                  // Payeur le mt de chaque Groupe 2 correspond au montant à payer
                  // par chaque centre payeur (établissement)
          EnregDucsEdi.MtAPayer := EnregDucsEdi.MtAPayer - EnrDucsEdiGroupee.MtAPayer;
      end;

      if (copy(EnregDucsEdi.TypDucs,1,1) <> 'U') then
      // Acoss et Irc
      begin
        InitCotisation (EnregDucsEdi, EnrDucsEdiGroupee,'G');
        if (ControlCotisation (EnregDucsEdi, EnrDucsEdiGroupee,'G') = false) then
        begin
          DossOK :=  false;
          exit;
        end;
      end;
      EnregDucsEdi.NoCentrePayeur := EnregDucsEdi.NoCentrePayeur + 1;
      Groupe2Declaration (EnregDucsEdi, EnrDucsEdiGroupee,'G');
      if (copy(EnregDucsEdi.TypDucs,1,1) = 'A') then
      // Acoss
      begin
        Groupe3ACOSS (EnregDucsEdi, EnrDucsEdiGroupee,'G');
        if (EnrDucsEdiGroupee.NbTransport > 1) or
           (EnrDucsEdiGroupee.NbTAT > 1) then  // PT37
          Groupe4_5ACOSS(EnregDucsEdi, EnrDucsEdiGroupee,'G');
      end;
      if (copy(EnregDucsEdi.TypDucs,1,1) = 'I') then
        // IRC
        Groupe3IRC (EnregDucsEdi, EnrDucsEdiGroupee,'G', NoLg);

        LesDucsGroupee := TOB_DucsGroupee.FindNext ([''],[''],TRUE);
    end;

// pour IRC Multi centre payeur  PT52
// une findéclaration + boucle paiement edi
// faire une TOB avec les éléments nécessaire à fin déclaration et paiement edi
    if (copy(EnregDucsEdi.TypDucs,1,1) = 'I') and
       (EnregDucsEdi.PaiementGroupe = False) then
    begin
      FinDeclaration(EnregDucsEdi,EnrDucsEdiGroupee);
      PaiementEdi(EnregDucsEdi,EnrDucsEdiGroupee,'', DossOK, TMtsEtabs);
      if (DossOK = False) then
        exit;
      TOBIrcFille := TOBIrc.FindFirst([''],[''],TRUE);
      While   TOBIrcFille <> NIL do
      begin
        EnrDucsEdiGroupee.Etab := TOBIrcFille.GetValue('ETAB');
        EnrDucsEdiGroupee.LibelleET := TOBIrcFille.GetValue('LIBELLEET');
        EnrDucsEdiGroupee.ModePaiement := TOBIrcFille.GetValue('MODEPAIEMENT');
        EnrDucsEdiGroupee.MtAPayer := TOBIrcFille.GetValue('MTAPAYER');
        EnrDucsEdiGroupee.Abrege := TOBIrcFille.GetValue('ABREGE');
        EnrDucsEdiGroupee.NumeroInterne := TOBIrcFille.GetValue('NOINTERNE');
        EnrDucsEdiGroupee.Paiement := TOBIrcFille.GetValue('PAIEMENT');
// d PT30 DUCS EDI V4.2
        if (EnrDucsEdiGroupee.ModePaiement = 'Z10') then
          EnrDucsEdiGroupee.TitulaireCpt := TOBIrcFille.GetValue('TITULAIRECPT');
// f PT30 DUCS EDI V4.2


        if (EnrDucsEdiGroupee.ModePaiement <> '') then
          PaiementEdi(EnregDucsEdi,EnrDucsEdiGroupee,'G', DossOK, TMtsEtabs);
        if (DossOK = False) then
          exit;
        TOBIrcFille := TOBIrc.FindNext ([''],[''],TRUE);
      end;
      TOBIrc.Free; TOBIrc := NIL;

    end;
  end;
end;
// fin  MultiGp2Gp3

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 28/05/2002
Modifié le ... :   /  /
Description .. : Initialisation des variables concernant les établissements
Suite ........ : secondaires nécessaires au traitements Multi établissemens
Suite ........ : tels que décrits dans MultiGp2GP3
Mots clefs ... : PAIE, PGDUCSEDI
*****************************************************************}
procedure InitVarGp2(var EnregDucsEdi : TDucsEdi; var EnrDucsEdiGroupee : TDucsEdiGroupee;LesDucsGroupee : TOB; var TMtsEtabs : TOB);   //PT37
var
   QL : TQuery;
   StQL : string;
   TMtsEtabsFille :  TOB;             // PT37
begin

  EnrDucsEdiGroupee.Etab := LesDucsGroupee.getValue('PDU_ETABLISSEMENT');
  EnrDucsEdiGroupee.NoDucs := LesDucsGroupee.getValue('PDU_NUM');

  // ETABLISS
  QL := OpenSQL('SELECT ET_LIBELLE,ET_ADRESSE1,ET_ADRESSE2,ET_ADRESSE3,'+
                'ET_VILLE,ET_CODEPOSTAL,ET_TELEPHONE,ET_FAX,ET_JURIDIQUE '+
                'FROM ETABLISS WHERE '+
                'ET_ETABLISSEMENT = "'+EnrDucsEdiGroupee.Etab+'" ',True) ;
  if Not QL.EOF then
    begin
      EnrDucsEdiGroupee.LibelleET:=QL.FindField('ET_LIBELLE').AsString;

      if (QL.FindField('ET_ADRESSE1').AsString <> '') then
         EnrDucsEdiGroupee.Adresse1ET:=QL.FindField('ET_ADRESSE1').AsString;
      if (QL.FindField('ET_ADRESSE2').AsString <> '') then
         if (EnrDucsEdiGroupee.Adresse1ET  <> '') then
           EnrDucsEdiGroupee.Adresse2ET:=QL.FindField('ET_ADRESSE2').AsString
         else
           EnrDucsEdiGroupee.Adresse1ET:=QL.FindField('ET_ADRESSE2').AsString;
      if (QL.FindField('ET_ADRESSE3').AsString <> '') then
         if (EnregDucsEdi.Adresse1ET  = '') then
           EnrDucsEdiGroupee.Adresse1ET:=QL.FindField('ET_ADRESSE3').AsString
         else
      if (EnrDucsEdiGroupee.Adresse2ET  = '') then
         EnrDucsEdiGroupee.Adresse2ET:=QL.FindField('ET_ADRESSE3').AsString;

      EnrDucsEdiGroupee.VilleET:=QL.FindField('ET_VILLE').AsString;
      EnrDucsEdiGroupee.CpET:=QL.FindField('ET_CODEPOSTAL').AsString;
    end ;
  Ferme(QL) ;

  // DUCSENTETE
  QL := OpenSql('SELECT PDU_SIRET,PDU_MONNAIETENUE,PDU_NBSALFPE,'+
                'PDU_TOTHOMMES,PDU_TOTFEMMES,PDU_ACOMPTES,'+
// d PT34
                'PDU_NBSALQ922,PDU_NBSALQ923,PDU_NBSALQ924,PDU_NBSALQ925,'+
                'PDU_NBSALQ926,PDU_NBSALQ927,PDU_NBSALQ928,PDU_NBSALQ929,'+
                'PDU_NBSALQ934,PDU_NBSALQ935,PDU_NBSALQ936,PDU_NBSALQ937,'+
                'PDU_NBSALQ944,PDU_NBSALQ945,PDU_NBSALQ960,PDU_NBSALQ961,'+
                'PDU_NBSALQ962,PDU_NBSALQ963,PDU_NBSALQ966,'+
                'PDU_ECARTZE1,PDU_ECARTZE2,PDU_ECARTZE3,PDU_ECARTZE4,'+
                'PDU_ECARTZE5,PDU_ECARTZE6,PDU_ECARTZE7,PDU_ECARTZE8,'+
// f PT34
                'PDU_REGULARISATION,PDU_TOTAPPRENTI, '+
                'PDU_ABREGEPERIODE,PDU_NUMERO,PDU_APE,PDU_PAIEMENT,PDU_DATEREGLEMENT, '+
                'PDU_CENTREPAYEUR '+
                ' FROM DUCSENTETE '+
                ' WHERE '+
                'PDU_ETABLISSEMENT = "'+EnrDucsEdiGroupee.Etab+'" AND '+
                'PDU_ORGANISME = "'+EnregDucsEdi.Organisme+'" AND '+
                'PDU_DATEDEBUT = "'+UsDateTime(EnregDucsEdi.DateDebut)+'" AND '+
                'PDU_DATEFIN = "'+UsDateTime(EnregDucsEdi.DateFin)+'" AND '+
                'PDU_NUM = '+IntToStr(EnrDucsEdiGroupee.NoDucs)+'',True) ;        // DB2
  if Not QL.EOF then
    begin
      EnrDucsEdiGroupee.Siret := QL.FindField('PDU_SIRET').AsString;
      EnrDucsEdiGroupee.Abrege := QL.FindField('PDU_ABREGEPERIODE').AsString;
      EnrDucsEdiGroupee.MonnaieTenue := QL.FindField('PDU_MONNAIETENUE').AsString;
      EnrDucsEdiGroupee.NbSalFpe := QL.FindField('PDU_NBSALFPE').AsFloat;
      EnrDucsEdiGroupee.TotHommes := QL.FindField('PDU_TOTHOMMES').AsFloat;
      EnrDucsEdiGroupee.TotFemmes := QL.FindField('PDU_TOTFEMMES').AsFloat;
      EnrDucsEdiGroupee.Regularisation:=QL.FindField('PDU_REGULARISATION').AsFloat;
      EnrDucsEdiGroupee.Acompte:=QL.FindField('PDU_ACOMPTES').AsFloat;
      EnrDucsEdiGroupee.TotApprenti := QL.FindField('PDU_TOTAPPRENTI').AsFloat;
      EnrDucsEdiGroupee.NumeroInterne := QL.FindField('PDU_NUMERO').AsString;
      EnrDucsEdiGroupee.ApePdu := QL.FindField('PDU_APE').AsString;
      EnrDucsEdiGroupee.Paiement:=QL.FindField('PDU_PAIEMENT').AsDateTime;
      EnrDucsEdiGroupee.Reglement:=QL.FindField('PDU_DATEREGLEMENT').AsDateTime;
      if (QL.FindField('PDU_CENTREPAYEUR').AsString <> '') then
        EnrDucsEdiGroupee.CentrePayeur:=QL.FindField('PDU_CENTREPAYEUR').AsString;
// d PT34
     EnrDucsEdiGroupee.CDIH:=QL.FindField('PDU_NBSALQ922').AsFloat;
     EnrDucsEdiGroupee.CDIF:=QL.FindField('PDU_NBSALQ923').AsFloat;
     EnrDucsEdiGroupee.ApprentiH:=QL.FindField('PDU_NBSALQ924').AsFloat;
     EnrDucsEdiGroupee.ApprentiF:=QL.FindField('PDU_NBSALQ925').AsFloat;
     EnrDucsEdiGroupee.CadreH:=QL.FindField('PDU_NBSALQ928').AsFloat;
     EnrDucsEdiGroupee.CadreF:=QL.FindField('PDU_NBSALQ929').AsFloat;
     EnrDucsEdiGroupee.E65H:=QL.FindField('PDU_NBSALQ926').AsFloat;
     EnrDucsEdiGroupee.E65F:=QL.FindField('PDU_NBSALQ927').AsFloat;
     EnrDucsEdiGroupee.ProfH:=QL.FindField('PDU_NBSALQ934').AsFloat;    // PT37
     EnrDucsEdiGroupee.ProfF:=QL.FindField('PDU_NBSALQ935').AsFloat;    // PT37
     EnrDucsEdiGroupee.InterimH:=QL.FindField('PDU_NBSALQ944').AsFloat;
     EnrDucsEdiGroupee.InterimF:=QL.FindField('PDU_NBSALQ945').AsFloat;
     EnrDucsEdiGroupee.SimplifH:=QL.FindField('PDU_NBSALQ960').AsFloat;
     EnrDucsEdiGroupee.SimplifF:=QL.FindField('PDU_NBSALQ961').AsFloat;
     EnrDucsEdiGroupee.CNEH:=QL.FindField('PDU_NBSALQ962').AsFloat;
     EnrDucsEdiGroupee.CNEF:=QL.FindField('PDU_NBSALQ963').AsFloat;
     EnrDucsEdiGroupee.CDDH:=QL.FindField('PDU_NBSALQ936').AsFloat;
     EnrDucsEdiGroupee.CDDF:=QL.FindField('PDU_NBSALQ937').AsFloat;
     EnrDucsEdiGroupee.EcartZe1:=QL.FindField('PDU_ECARTZE1').AsString='X';
     EnrDucsEdiGroupee.EcartZe2:=QL.FindField('PDU_ECARTZE2').AsString='X';
     EnrDucsEdiGroupee.EcartZe3:=QL.FindField('PDU_ECARTZE3').AsString='X';
     EnrDucsEdiGroupee.EcartZe4:=QL.FindField('PDU_ECARTZE4').AsString='X';
     EnrDucsEdiGroupee.EcartZe5:=QL.FindField('PDU_ECARTZE5').AsString='X';
     EnrDucsEdiGroupee.EcartZe6:=QL.FindField('PDU_ECARTZE6').AsString='X';
     EnrDucsEdiGroupee.EcartZe7:=QL.FindField('PDU_ECARTZE7').AsString='X';
     EnrDucsEdiGroupee.EcartZe8:=QL.FindField('PDU_ECARTZE8').AsString='X';
     if  (QL.FindField('PDU_NBSALQ966').AsFloat = 1) then
        EnrDucsEdiGroupee.EcartZe9 := TRUE
     else
        EnrDucsEdiGroupee.EcartZe9 := FALSE;
// f PT34
    end;
  ferme (QL);

// PT30 DUCS EDI V4.2
//QL := OpenSql('SELECT POG_PAIEMODE FROM ORGANISMEPAIE WHERE '+
  QL := OpenSql('SELECT POG_PAIEMODE,POG_TITULAIRECPT FROM ORGANISMEPAIE WHERE '+
                'POG_ETABLISSEMENT = "'+EnrDucsEdiGroupee.Etab+'" AND '+
                'POG_ORGANISME = "'+EnregDucsEdi.Organisme+'"',True) ;
   if Not QL.EOF then
    begin
      EnrDucsEdiGroupee.ModePaiement := QL.FindField('POG_PAIEMODE').AsString;
      EnrDucsEdiGroupee.TitulaireCpt := QL.FindField('POG_TITULAIRECPT').AsString; // PT30 DUCS EDI V4.2
    end;
  ferme (QL);


  // DUCSDETAIL
  StQL := 'PDD_ETABLISSEMENT = "'+EnrDucsEdiGroupee.Etab+'" AND '+
          'PDD_ORGANISME = "'+EnregDucsEdi.Organisme+'" AND '+
          'PDD_DATEDEBUT = "'+UsDateTime(EnregDucsEdi.DateDebut)+'" AND '+
          'PDD_DATEFIN = "'+UsDateTime(EnregDucsEdi.DateFin)+'" AND '+
          'PDD_NUM = '+IntToStr(EnrDucsEdiGroupee.NoDucs)+' AND '+        // DB2
          'PDD_TYPECOTISATION <> "S"';

  QL := OpenSQL('SELECT SUM(PDD_MTCOTISAT) AS DECLAR FROM DUCSDETAIL WHERE '+StQL,True) ;

  EnrDucsEdiGroupee.MtDeclare := 0.0;
  IF NOT QL.EOF then // PortageCWAS
   EnrDucsEdiGroupee.MtDeclare := QL.FindField('DECLAR').asFloat;
  EnregDucsEdi.MtTotal :=  EnregDucsEdi.MtTotal +  EnrDucsEdiGroupee.MtDeclare;

  EnrDucsEdiGroupee.MtAPayer := 0.0;
  EnrDucsEdiGroupee.MtAPayer := EnrDucsEdiGroupee.MtAPayer+
                              EnrDucsEdiGroupee.MtDeclare+
                              EnrDucsEdiGroupee.Regularisation-
                              EnrDucsEdiGroupee.Acompte;
// d PT37
  if (TMtsEtabs = NIL) then
    TMtsEtabs := TOB.Create ('les montants par etab', NIL,-1);
  TMtsEtabsFille := TOB.Create ('les montants par etab', TMtsEtabs,-1);
  TMtsEtabsFille.AddChampSupValeur ('ETAB',EnrDucsEdiGroupee.Etab, FALSE);
  TMtsEtabsFille.AddChampSupValeur ('LibelleET',EnrDucsEdiGroupee.LibelleET, FALSE);
  TMtsEtabsFille.AddChampSupValeur ('ACOMPTE',EnrDucsEdiGroupee.Acompte, FALSE);
  TMtsEtabsFille.AddChampSupValeur ('REGUL',EnrDucsEdiGroupee.Regularisation, FALSE);
  TMtsEtabsFille.AddChampSupValeur ('APAYER',EnrDucsEdiGroupee.MtAPayer, FALSE);
// f PT37

  Ferme (QL);
// d PT30 DUCS EDI V4.2
  if (copy(EnregDucsEdi.TypDucs,1,1) = 'A') then
  begin
    // Versement transport
    StQL := '';
    StQL := 'PDD_ETABLISSEMENT = "'+EnrDucsEdiGroupee.Etab+'" AND '+
            'PDD_ORGANISME = "'+EnregDucsEdi.Organisme+'" AND '+
            'PDD_DATEDEBUT = "'+UsDateTime(EnregDucsEdi.DateDebut)+'" AND '+
            'PDD_DATEFIN = "'+UsDateTime(EnregDucsEdi.DateFin)+'" AND '+
            'PDD_NUM = '+IntToStr(EnrDucsEdiGroupee.NoDucs)+' AND '+        // DB2
            'PDD_TYPECOTISATION <> "S" AND PDD_TYPECOTISATION <> "I" '+
            'AND PDD_CODIFICATION LIKE "1%900P"';
    QL := OpenSQL('SELECT COUNT(DISTINCT(PDD_TAUXCOTISATION)) AS NBTRANSPORT, '+
                  'SUM(PDD_MTCOTISAT) AS MTTRANSPORT FROM DUCSDETAIL WHERE '+StQL,True) ;
    EnrDucsEdiGroupee.MtTransport:=0.0;
    IF NOT QL.EOF then // PortageCWAS
    begin
      EnrDucsEdiGroupee.MtTransport:= QL.FindField('MTTRANSPORT').asFloat;
      EnrDucsEdiGroupee.NbTransport:= QL.FindField('NBTRANSPORT').AsInteger;
    end;
    Ferme(QL);

    EnrDucsEdiGroupee.ErrTransport := False;
// PT30 DUCS EDI V4.2
//  if (copy(EnregDucsEdi.TypDucs,1,1) = 'A') and (EnrDucsEdiGroupee.NbTransport > 1) then
    if (EnrDucsEdiGroupee.NbTransport > 1) then
    // Acoss - Création TOB_Transport
    begin
      QL := OpenSQL('SELECT Count (*) FROM DUCSDETAIL WHERE (PDD_CODECOMMUNE="" '+
                     'OR PDD_CODECOMMUNE IS NULL OR PDD_CODECOMMUNE="    ") and '+
                     StQL,TRUE);
       if NOT QL.EOF then
       begin
         if (QL.Fields[0].AsInteger <> 0) then
            EnrDucsEdiGroupee.ErrTransport :=  True;
       end;
     ferme (QL);
    end;

    QL := OpenSQL('SELECT * FROM DUCSDETAIL WHERE '+StQL,True);

    TOB_Transport := TOB.Create ('Les lignes de tansport', NIL, -1);

    TOB_Transport.LoadDetailDB ('DUCSDETAIL', '','',QL,FALSE,FALSE);
    Ferme(QL);
    TOB_Transport.Detail.Sort('PDD_ETABLISSEMENT;PDD_ORGANISME;PDD_DATEDEBUT;'+
                              'PDD_DATEFIN;PDD_INSTITUTION;PDD_CODIFICATION;PDD_DATECHGTTAUX');
    // Taxes additionnelles de transport
    StQL := '';
    StQL := 'PDD_ETABLISSEMENT = "'+EnrDucsEdiGroupee.Etab+'" AND '+
            'PDD_ORGANISME = "'+EnregDucsEdi.Organisme+'" AND '+
            'PDD_DATEDEBUT = "'+UsDateTime(EnregDucsEdi.DateDebut)+'" AND '+
            'PDD_DATEFIN = "'+UsDateTime(EnregDucsEdi.DateFin)+'" AND '+
            'PDD_NUM = '+IntToStr(EnrDucsEdiGroupee.NoDucs)+' AND '+        // DB2
            'PDD_TYPECOTISATION <> "S" AND PDD_TYPECOTISATION <> "I" '+
            'AND PDD_CODIFICATION LIKE "1%901P"';
    QL := OpenSQL('SELECT COUNT(DISTINCT(PDD_TAUXCOTISATION)) AS NBTAT, '+
                  'SUM(PDD_MTCOTISAT) AS MTTAT FROM DUCSDETAIL WHERE '+StQL,True) ;
    EnrDucsEdiGroupee.MtTAT:=0.0;
    IF NOT QL.EOF then
    begin
      EnrDucsEdiGroupee.MtTAT:= QL.FindField('MTTAT').asFloat;
//PT37      EnrDucsEdiGroupee.NbTransport:= QL.FindField('NBTAT').AsInteger;
      EnrDucsEdiGroupee.NbTAT:= QL.FindField('NBTAT').AsInteger;
    end;
    Ferme(QL);

    EnrDucsEdiGroupee.ErrTAT := False;
    if (EnrDucsEdiGroupee.NbTAT > 1) then
    //Création TOB_TAT
    begin
      QL := OpenSQL('SELECT Count (*) FROM DUCSDETAIL WHERE (PDD_CODECOMMUNE="" '+
                     'OR PDD_CODECOMMUNE IS NULL OR PDD_CODECOMMUNE="    ") and '+
                     StQL,TRUE);
       if NOT QL.EOF then
       begin
         if (QL.Fields[0].AsInteger <> 0) then
            EnrDucsEdiGroupee.ErrTAT :=  True;
       end;
     ferme (QL);
    end;

    QL := OpenSQL('SELECT * FROM DUCSDETAIL WHERE '+StQL,True);

    TOB_TAT := TOB.Create ('Les lignes de taxes additionnelles de transport', NIL, -1);

    TOB_TAT.LoadDetailDB ('DUCSDETAIL', '','',QL,FALSE,FALSE);
    Ferme(QL);
    TOB_TAT.Detail.Sort('PDD_ETABLISSEMENT;PDD_ORGANISME;PDD_DATEDEBUT;'+
                              'PDD_DATEFIN;PDD_INSTITUTION;PDD_CODIFICATION;PDD_DATECHGTTAUX');
  end; // f PT30 DUCS EDI V4.2
  EnregDucsEdi.MtAPayer := EnregDucsEdi.MtAPayer+
                           EnrDucsEdiGroupee.MtDeclare+
                           EnrDucsEdiGroupee.Regularisation-
                           EnrDucsEdiGroupee.Acompte;

  EnregDucsEdi.MtTotRegul :=  EnregDucsEdi.MtTotRegul + EnrDucsEdiGroupee.Regularisation;
  EnregDucsEdi.MtTotAcpte :=  EnregDucsEdi.MtTotAcpte + EnrDucsEdiGroupee.Acompte;

end;
// fin InitVarGp2
{***********A.G.L.***********************************************
Auteur  ...... : PAIE : MF
Créé le ...... : 22/05/2002
Modifié le ... :   /  /
Description .. : Initialisation des variables concernant les
Suite ........ : établissements secondaires pour l'UNEDIC
Mots clefs ... :
*****************************************************************}
procedure InitVarGp2UNEDIC(var EnregDucsEdi : TDucsEdi; var EnrDucsEdiGroupee : TDucsEdiGroupee;LesDucsGroupee : TOB);
var
   QL : TQuery;
   StQL : string;
   TOB_LignesCot : TOB;
begin

  EnrDucsEdiGroupee.Etab := LesDucsGroupee.getValue('POG_ETABLISSEMENT');
  EnrDucsEdiGroupee.NumeroInterne := LesDucsGroupee.getValue('POG_NUMINTERNE');
  EnrDucsEdiGroupee.ModePaiement := '';
  EnrDucsEdiGroupee.Abrege := EnregDucsEdi.Abrege;
  EnrDucsEdiGroupee.MtDeclare  :=0.0;
  EnrDucsEdiGroupee.MtTransport :=0.0;
  EnrDucsEdiGroupee.NbTransport :=0;
  EnrDucsEdiGroupee.Regularisation :=0.0;
  EnrDucsEdiGroupee.Acompte :=0.0;
  EnrDucsEdiGroupee.MtAPayer :=0.0;
  EnrDucsEdiGroupee.ErrTransport := False;
  EnrDucsEdiGroupee.ErrTAT := False;      // PT30 DUCS EDI V4.2
  EnrDucsEdiGroupee.ErrCotisqual := False;
  EnrDucsEdiGroupee.ErrInstitution := False;
  EnrDucsEdiGroupee.ErrCondition := False;
  if VH_Paie.PGTenueEuro=FALSE then
    EnrDucsEdiGroupee.MonnaieTenue := 'FRF'          // Franc
  else
    EnrDucsEdiGroupee.MonnaieTenue := 'EUR';         // Euro

  EnrDucsEdiGroupee.Paiement := Idate1900;

  // ETABLISS
  QL := OpenSQL('SELECT ET_LIBELLE,ET_ADRESSE1,ET_ADRESSE2,ET_ADRESSE3,'+
                'ET_VILLE,ET_CODEPOSTAL,ET_TELEPHONE,ET_FAX,ET_JURIDIQUE, '+
                'ET_SIRET, ET_APE '+
                'FROM ETABLISS WHERE '+
                'ET_ETABLISSEMENT = "'+EnrDucsEdiGroupee.Etab+'" ',True) ;
  if Not QL.EOF then
    begin
      EnrDucsEdiGroupee.LibelleET:=QL.FindField('ET_LIBELLE').AsString;

      if (QL.FindField('ET_ADRESSE1').AsString <> '') then
         EnrDucsEdiGroupee.Adresse1ET:=QL.FindField('ET_ADRESSE1').AsString;
      if (QL.FindField('ET_ADRESSE2').AsString <> '') then
         if (EnrDucsEdiGroupee.Adresse1ET  <> '') then
           EnrDucsEdiGroupee.Adresse2ET:=QL.FindField('ET_ADRESSE2').AsString
         else
           EnrDucsEdiGroupee.Adresse1ET:=QL.FindField('ET_ADRESSE2').AsString;
      if (QL.FindField('ET_ADRESSE3').AsString <> '') then
         if (EnregDucsEdi.Adresse1ET  = '') then
           EnrDucsEdiGroupee.Adresse1ET:=QL.FindField('ET_ADRESSE3').AsString
         else
      if (EnrDucsEdiGroupee.Adresse2ET  = '') then
         EnrDucsEdiGroupee.Adresse2ET:=QL.FindField('ET_ADRESSE3').AsString;

      EnrDucsEdiGroupee.VilleET:=QL.FindField('ET_VILLE').AsString;
      EnrDucsEdiGroupee.CpET:=QL.FindField('ET_CODEPOSTAL').AsString;
      EnrDucsEdiGroupee.Siret:=QL.FindField('ET_SIRET').AsString;
      EnrDucsEdiGroupee.ApePdu:=QL.FindField('ET_APE').AsString;
    end ;
  Ferme(QL) ;

  // Calcul des Effectifs
  StQL := '';
  StQL :='PSA_ETABLISSEMENT="'+EnrDucsEdiGroupee.Etab+'" AND ';

  StQL := StQL+'(PSA_DATESORTIE="'+UsDateTime(IDate1900)+'"'+
               ' OR  PSA_DATESORTIE IS NULL '+
               ' OR PSA_DATESORTIE>="'+UsDateTime(EnregDucsEdi.DateDebut)+'")'+
               ' AND PSA_DATEENTREE<="'+UsDateTime(EnregDucsEdi.DateFin)+'"';

 // nbre d'hommes
 QL := OpenSQL('SELECT Count(*) FROM SALARIES WHERE '+StQL+
               ' AND PSA_SEXE="M"',True) ;
 IF NOT QL.EOF then // PortageCWAS
  EnrDucsEdiGroupee.TotHommes := QL.Fields[0].AsInteger ;
 Ferme(QL) ;

 // nbre de femmes
 QL := OpenSQL('SELECT Count(*) FROM SALARIES WHERE '+StQL+
               ' AND PSA_SEXE="F"',True) ;
 IF NOT QL.EOF then // PortageCWAS
  EnrDucsEdiGroupee.TotFemmes := QL.Fields[0].AsInteger ;
 Ferme(QL) ;

 // Nbre d'apprentis
 QL := OpenSQL('SELECT Count(*) FROM SALARIES WHERE '+StQL+
               ' AND PSA_CATDADS="003"',True) ;
 IF NOT QL.EOF then //  PortageCWAS
  EnrDucsEdiGroupee.TotApprenti := QL.Fields[0].AsInteger ;
 Ferme(QL) ;

 // Nbre de salariés rémunérés pour la période
 StQL := '';
 StQL :='PHB_ETABLISSEMENT="'+EnrDucsEdiGroupee.Etab+'" AND ';
 StQL := StQL+ 'PHB_ORGANISME="'+EnregDucsEdi.Organisme+'" ';

 StQL := StQL+'AND '+
              'PHB_DATEDEBUT >="'+UsDateTime(EnregDucsEdi.DateDebut)+'" AND '+
              'PHB_DATEFIN <="'+UsDateTime(EnregDucsEdi.DateFin)+'" AND '+
              '(PHB_BASECOT <> 0 OR PHB_MTSALARIAL <> 0 OR PHB_MTPATRONAL <> 0)';

 QL := OpenSQL('SELECT Count(DISTINCT PHB_SALARIE) AS NOMBRE FROM HISTOBULLETIN WHERE '+
                   StQL,True) ;
 IF NOT QL.EOF then // PortageCWAS
  EnrDucsEdiGroupee.NbSalFpe := QL.FindField('NOMBRE').AsInteger ;
 Ferme (QL);
// d PT30 DUCS EDI V4.2
  // DUCSENTETE
  QL := OpenSql('SELECT PDU_ACOMPTES,'+
                'PDU_REGULARISATION '+
                ' FROM DUCSENTETE '+
                ' WHERE '+
                'PDU_ETABLISSEMENT = "'+EnrDucsEdiGroupee.Etab+'" AND '+
                'PDU_ORGANISME = "'+EnregDucsEdi.Organisme+'" AND '+
                'PDU_DATEDEBUT = "'+UsDateTime(EnregDucsEdi.DateDebut)+'" AND '+
                'PDU_DATEFIN = "'+UsDateTime(EnregDucsEdi.DateFin)+'" AND '+
                'PDU_NUM = '+IntToStr(EnrDucsEdiGroupee.NoDucs)+'',True) ;        // DB2
  if Not QL.EOF then
  begin
      EnrDucsEdiGroupee.Regularisation:=QL.FindField('PDU_REGULARISATION').AsFloat;
      EnrDucsEdiGroupee.Acompte:=QL.FindField('PDU_ACOMPTES').AsFloat;
  end;
  ferme (QL);

  if (TOB_LignesCot = NIL) then
    TOB_LignesCot := TOB.Create ('Les lignes de cotisation', NIL, -1);

  QL :=OpenSql ('SELECT * FROM DUCSDETAIL '+
                 'WHERE '+
                  '(PDD_ETABLISSEMENT ="'+ EnrDucsEdiGroupee.Etab +'") AND '+
                  '(PDD_ORGANISME ="'+ EnregDucsEdi.Organisme+'") AND '+
                  '(PDD_DATEDEBUT ="'+UsDateTime(EnregDucsEdi.DateDebut)+'") AND '+
                  '(PDD_DATEFIN = "'+UsDateTime(EnregDucsEdi.DateFin)+'") AND '+
                  '(PDD_NUM = '+IntToStr(EnrDucsEdiGroupee.NoDucs)+') AND '+
                  '(PDD_TYPECOTISATION <> "S" AND PDD_TYPECOTISATION <> "I")', TRUE);

  TOB_LignesCot.LoadDetailDB ('ASSEDICDETAIL', '','',QL,FALSE,FALSE);

  EnrDucsEdiGroupee.MtDeclare := TOB_LignesCot.Somme('PDD_MTCOTISAT',[''],[''],TRUE, FALSE);

  ferme (QL);

  EnrDucsEdiGroupee.MtAPayer := 0.0;
  EnrDucsEdiGroupee.MtAPayer := EnrDucsEdiGroupee.MtAPayer+
                              EnrDucsEdiGroupee.MtDeclare+
                              EnrDucsEdiGroupee.Regularisation-
                              EnrDucsEdiGroupee.Acompte;
  EnregDucsEdi.MtAPayer := EnregDucsEdi.MtAPayer +EnrDucsEdiGroupee.MtAPayer

// f PT30 DUCS EDI V4.2
end;
// fin InitVarGp2UNEDIC
{***********A.G.L.***********************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 13/05/2002
Modifié le ... : 13/05/2002
Description .. : Fonction de calcul et de contrôle de la clé IBAN
Suite ........ : Paramètres : Type de créancier  (2 car)
Suite ........ :              Période            (6 car)
Suite ........ :              référence interne  (19 car)
Suite ........ :              Clé Iban           (2 car) = 00 pour calcul
Suite ........ :                                        <> 00 por contrôle
Mots clefs ... : PAIE, PGDUCSEDI
*****************************************************************}
Function CleIBAN (creancier, periode,reference : string;var cleIBAN : string) : boolean;
var
   car : char;
   RefSociale : string;
   RefSocialeSize,icar : integer;
   Divi : string;
   i,j,IRest,IDivi: integer;
   NbNonCar : integer;
begin
 IDivi := 0;
 // Constitution de la référence sociale
 RefSociale := '';
 result := True;

   for i := 1 to 2 do
    begin
     car := creancier[i];
     if (IsNumeric(car) = false) then
       begin
         icar := Ord(creancier[i])-55;
         RefSociale := RefSociale+IntToStr(icar);
       end
     else
       RefSociale := RefSociale+car;
    end;

   for i := 1 to 6 do
    begin
     car := periode[i];
     if (IsNumeric(car) = false) then
       begin
         icar := Ord(car)-55;
         RefSociale := RefSociale+IntToStr(icar);
       end
     else
      RefSociale := RefSociale+car;
    end;

     reference := UpperCase(reference);  
     NbNonCar := 0;
     for i := 1 to length(reference) do
       begin
         car := reference[i];
         if ((Ord(car)>=65) and (Ord(car)<=90)) or
            ((IsNumeric(car) = true) and
            (car <> '') and (car <> '.') and(car <> ' ')) then
           begin
             if (IsNumeric(car) = false) then
               begin
                 icar := Ord(car)-55;
                 RefSociale := RefSociale+IntToStr(icar);
               end
             else
               RefSociale := RefSociale+car;
           end
         else
          NbNonCar := NbNonCar+1;
       end;
     if ((length(reference)-NbNonCar) < 19) then
      for i:= (length(reference)-NbNonCar)+1 to 19 do
        begin
         car:= '0';
         RefSociale := RefSociale+car;
        end;

   RefSociale :=  RefSociale+ CleIBAN;
   RefSocialeSize :=  length (RefSociale);

  // Division par 97 :
  // On procède par divisions successives sur la chaîne RefSociale
  j := 1;
  divi := '';
  IRest := 0;
  while (j <= RefSocialeSize) do
    begin
     for i:=j to RefSocialeSize do
       begin
         j := j+1;
         divi := divi+copy(RefSociale,i,1);
         IDivi := StrToInt(divi);
         if (Idivi > 97) then
          begin
            divi := '';
            break;
          end;
          IDivi := 0;
       end;
       Irest := Idivi mod 97;
       if (j <= RefSocialeSize) then
       // Irest est repris pour déterminer la nouvelle chaîne à diviser
       divi := IntToStr(Irest);
    end;
  // Calcul ou contrôle de la clé
  if (StrToInt(CleIBAN) = 0) then
    begin
     if ((98-IRest) > 9) then
        CleIBAN := IntToStr(98-IRest)
     else
        CleIBAN := '0'+ IntToStr(98-IRest);
    end
  else
      if (Irest <> 1) then result := False;

end;
//  fin CleIBAN

procedure MultiGp2Gp3TOB(EtabPrinc : string;var EnregDucsEdi : TDucsEdi;var EnrDucsEdiGroupee : TDucsEdiGroupee; var DossOK : boolean; var NoLg : integer);
var
   QL : TQuery;
   //LesDucsGroupee,
   //TOBIrcFille,
   //TOBIrc :  TOB;
begin
// d PT34
       if (TOB_DucsGroupee <> nil) then
          FreeAndNil(TOB_DucsGroupee);
// PT34
       TOB_DucsGroupee := TOB.Create ('Les ducs paiement groupé', NIL, -1);

       QL :=OpenSql ('SELECT PDU_ETABLISSEMENT,PDU_ORGANISME,'+
                     'PDU_DATEDEBUT,PDU_DATEFIN,PDU_DUCSDOSSIER,PDU_NUM'+
                     ' FROM DUCSENTETE '+
                     'WHERE '+
                     '(PDU_ETABLISSEMENT <>"'+ EtabPrinc +'") AND '+
                     '(PDU_ORGANISME ="'+ EnregDucsEdi.Organisme+'") AND '+
                     '(PDU_DATEDEBUT ="'+UsDateTime(EnregDucsEdi.DateDebut)+'") AND '+
                     '(PDU_DATEFIN = "'+UsDateTime(EnregDucsEdi.DateFin)+'") AND '+
                     '(PDU_DUCSDOSSIER = "X")', TRUE);

       TOB_DucsGroupee.LoadDetailDB ('TABLEDUCSENTETE', '','',QL,FALSE,FALSE);

       Ferme(QL);

       TOB_DucsGroupee.Detail.Sort('PDU_NUM')
end;
procedure InitEffGlobal(var EnregDucsEdi : TDucsEdi; var EnrDucsEdiGroupee : TDucsEdiGroupee;LesDucsGroupee : TOB);
var
   QL : TQuery;
//   StQL : string;
begin

  EnrDucsEdiGroupee.Etab := LesDucsGroupee.getValue('PDU_ETABLISSEMENT');
  EnrDucsEdiGroupee.NoDucs := LesDucsGroupee.getValue('PDU_NUM');

    // DUCSENTETE
  QL := OpenSql('SELECT PDU_TOTHOMMES,PDU_TOTFEMMES '+
                ' FROM DUCSENTETE '+
                ' WHERE '+
                'PDU_ETABLISSEMENT = "'+EnrDucsEdiGroupee.Etab+'" AND '+
                'PDU_ORGANISME = "'+EnregDucsEdi.Organisme+'" AND '+
                'PDU_DATEDEBUT = "'+UsDateTime(EnregDucsEdi.DateDebut)+'" AND '+
                'PDU_DATEFIN = "'+UsDateTime(EnregDucsEdi.DateFin)+'" AND '+
                'PDU_NUM = '+IntToStr(EnrDucsEdiGroupee.NoDucs)+'',True) ;        // DB2
  if Not QL.EOF then
    begin
      EnrDucsEdiGroupee.TotHommes := QL.FindField('PDU_TOTHOMMES').AsFloat;
      EnrDucsEdiGroupee.TotFemmes := QL.FindField('PDU_TOTFEMMES').AsFloat;
    end;
  ferme (QL);

end;
procedure DetUnedicGroupee(var EnregDucsEdi : TDucsEdi; var EnrDucsEdiGroupee : TDucsEdiGroupee);
var
   QL                          : TQuery;
   TOB_AssedicGroupee          : TOB;
   TAssedic                    : TOB;
   TOB_LigneSCot               : TOB;
   TligneCot                   : TOB;
   LesLignes                   : TOB;
   i                           : integer;
   CodifLigne                  : string;
   TauxLigne                   : double;
begin
  // On établit la TOB des différents établissements secondaires
  TOB_AssedicGroupee := TOB.Create ('Les ducs assedic paiement groupé', NIL, -1);

  QL :=OpenSql ('SELECT PDU_ETABLISSEMENT,PDU_NUM,PDU_MASSEANNUEL,PDU_REMUNDADS'+
                ' FROM DUCSENTETE '+
                'WHERE '+
               '(PDU_ETABLISSEMENT <>"'+ EnregDucsEdi.Etab +'") AND '+
               '(PDU_ORGANISME ="'+ EnregDucsEdi.Organisme+'") AND '+
               '(PDU_DATEDEBUT ="'+UsDateTime(EnregDucsEdi.DateDebut)+'") AND '+
               '(PDU_DATEFIN = "'+UsDateTime(EnregDucsEdi.DateFin)+'") AND '+
               '(PDU_DUCSDOSSIER = "X")', TRUE);

  TOB_AssedicGroupee.LoadDetailDB ('DUCSASSEDIC', '','',QL,FALSE,FALSE);
  Ferme(QL);

  if (TOB_AssedicGroupee <> NIL) and  (TOB_AssedicGroupee.detail.count <> 0) then
  begin
    TAssedic := TOB_AssedicGroupee.FindFirst ([''],[''],TRUE);
    While TAssedic <> NIL  do
      begin
        // On établit la TOB des lignes de cotisation de l'établissement
        if (TOB_LignesCot = NIL) then
          TOB_LignesCot := TOB.Create ('Les lignes de cotisation', NIL, -1);

        QL :=OpenSql ('SELECT * FROM DUCSDETAIL '+
                      'WHERE '+
                  '(PDD_ETABLISSEMENT ="'+ TAssedic.GetValue('PDU_ETABLISSEMENT') +'") AND '+
                  '(PDD_ORGANISME ="'+ EnregDucsEdi.Organisme+'") AND '+
                  '(PDD_DATEDEBUT ="'+UsDateTime(EnregDucsEdi.DateDebut)+'") AND '+
                  '(PDD_DATEFIN = "'+UsDateTime(EnregDucsEdi.DateFin)+'") AND '+
                  '(PDD_NUM = '+IntToStr(TAssedic.GetValue('PDU_NUM'))+') AND '+
                  '(PDD_TYPECOTISATION <> "S" AND PDD_TYPECOTISATION <> "I")', TRUE);

        TOB_LignesCot.LoadDetailDB ('ASSEDICDETAIL', '','',QL,FALSE,FALSE);
      EnregDucsEdi.MasseAnnuel := EnregDucsEdi.MasseAnnuel+
                                  TAssedic.GetValue('PDU_MASSEANNUEL');
      EnregDucsEdi.RemunDads := EnregDucsEdi.RemunDads+
                                TAssedic.GetValue('PDU_REMUNDADS');

        TAssedic := TOB_AssedicGroupee.FindNext ([''],[''],TRUE);
      end;
      if (TOB_Lignes <> NIL) then
      begin
        i := 0;
        if (TOB_LignesCOT <> NIL) then
        begin
          While (i < TOB_Lignes.Detail.Count)  do
          // pour chaque ligne de la TOB_Lignes on tient le cumul avec les lignes
          // de cotisation de l'établissement secondaire
          begin
            LesLignes := TOB_Lignes.Detail[i];
            CodifLigne := Leslignes.getValue('PDD_CODIFICATION');
            TauxLigne := Leslignes.getValue('PDD_TAUXCOTISATION');
            TligneCot := TOB_LignesCot.FindFirst(['PDD_CODIFICATION', 'PDD_TAUXCOTISATION'],
                                                 [CodifLigne, TauxLigne], FALSE);
            while (TligneCot <> nil) do
            begin
              if (TligneCot.GetValue('PDD_ETABLISSEMENT') <> EnregDucsEdi.Etab) then
              begin
                LesLignes.PutValue('PDD_BASECOTISATION',
                                   LesLignes.GetValue('PDD_BASECOTISATION')+
                                   TligneCot.GetValue('PDD_BASECOTISATION'));
                LesLignes.PutValue('PDD_EFFECTIF',
                                   LesLignes.GetValue('PDD_EFFECTIF')+
                                   TligneCot.GetValue('PDD_EFFECTIF'));
                LesLignes.PutValue('PDD_MTCOTISAT',
                                   LesLignes.GetValue('PDD_MTCOTISAT')+
                                   TligneCot.GetValue('PDD_MTCOTISAT'));
                TligneCot.free; // on épure TOB_lignesCot au fur et à mesure

                TligneCot := TOB_LignesCot.FindNext(['PDD_CODIFICATION', 'PDD_TAUXCOTISATION'],
                                                    [CodifLigne, TauxLigne], FALSE);
              end;
            end;
            i := i+1;
          end;
          if (TOB_LignesCot <> NIL) and (TOB_LignesCot.Detail.Count <> 0) then
          // le nbre de lignes de TOB_LignesCot est supérieur à celui de TOB_Lignes
          // il faut créer des filles dans TOB_Lignes
          begin
            TligneCot := TOB_LignesCot.FindFirst ([''],[''],TRUE);
            While (TligneCot <> NIL) do
            begin
              COdifLigne := TligneCot.getValue('PDD_CODIFICATION');
              TauxLigne := TligneCot.getValue('PDD_TAUXCOTISATION');

              // La ligne TOB_Lignes existe-t-elle déjà?
              LesLignes := TOB_Lignes.FindFirst(['PDD_CODIFICATION', 'PDD_TAUXCOTISATION'],
                                                 [CodifLigne, TauxLigne], FALSE);
              if (LesLignes <> NIL) then
              // existe déjà
              begin
                LesLignes.PutValue('PDD_BASECOTISATION',
                                   LesLignes.GetValue('PDD_BASECOTISATION')+
                                   TligneCot.GetValue('PDD_BASECOTISATION'));
                LesLignes.PutValue('PDD_EFFECTIF',
                                   LesLignes.GetValue('PDD_EFFECTIF')+
                                   TligneCot.GetValue('PDD_EFFECTIF'));
                LesLignes.PutValue('PDD_MTCOTISAT',
                                   LesLignes.GetValue('PDD_MTCOTISAT')+
                                   TligneCot.GetValue('PDD_MTCOTISAT'));
              end
              else
              // Nouvelle ligne
              begin
                LesLignes := TOB.Create('DUCSDETAIL',TOB_Lignes,-1);

                LesLignes.PutValue('PDD_ETABLISSEMENT',EnregDucsEdi.Etab);
                LesLignes.PutValue('PDD_ORGANISME',TligneCot.GetValue('PDD_ORGANISME'));
                LesLignes.PutValue('PDD_DATEDEBUT',TligneCot.GetValue('PDD_DATEDEBUT'));
                LesLignes.PutValue('PDD_DATEFIN',TligneCot.GetValue('PDD_DATEFIN'));
                LesLignes.PutValue('PDD_CODIFICATION',TligneCot.GetValue('PDD_CODIFICATION'));
                LesLignes.PutValue('PDD_BASECOTISATION',TligneCot.GetValue('PDD_BASECOTISATION'));
                LesLignes.PutValue('PDD_TAUXCOTISATION',TligneCot.GetValue('PDD_TAUXCOTISATION'));
                LesLignes.PutValue('PDD_EFFECTIF',TligneCot.GetValue('PDD_EFFECTIF'));
                LesLignes.PutValue('PDD_LIBELLE',TligneCot.GetValue('PDD_LIBELLE'));
                LesLignes.PutValue('PDD_LIBELLESUITE',TligneCot.GetValue('PDD_LIBELLESUITE'));
                LesLignes.PutValue('PDD_TYPECOTISATION',TligneCot.GetValue('PDD_TYPECOTISATION'));
                LesLignes.PutValue('PDD_COMURBAINE',TligneCot.GetValue('PDD_COMURBAINE'));
                LesLignes.PutValue('PDD_CONDITION',TligneCot.GetValue('PDD_CONDITION'));
                LesLignes.PutValue('PDD_SIRET',TligneCot.GetValue('PDD_SIRET'));
                LesLignes.PutValue('PDD_APE',TligneCot.GetValue('PDD_APE'));
                LesLignes.PutValue('PDD_GROUPE',TligneCot.GetValue('PDD_GROUPE'));
                LesLignes.PutValue('PDD_NUMERO',TligneCot.GetValue('PDD_NUMERO'));
                LesLignes.PutValue('PDD_DATECHGTTAUX',TligneCot.GetValue('PDD_DATECHGTTAUX'));
                LesLignes.PutValue('PDD_NUMORDRE',TligneCot.GetValue('PDD_NUMORDRE'));
                LesLignes.PutValue('PDD_CODIFEDITEE',TligneCot.GetValue('PDD_CODIFEDITEE'));
                LesLignes.PutValue('PDD_MTCOTISAT',TligneCot.GetValue('PDD_MTCOTISAT'));
                LesLignes.PutValue('PDD_INSTITUTION',TligneCot.GetValue('PDD_INSTITUTION'));
                LesLignes.PutValue('PDD_NUM',TligneCot.GetValue('PDD_NUM'));
                LesLignes.PutValue('PDD_CODECOMMUNE',TligneCot.GetValue('PDD_CODECOMMUNE'));
              end;
              TligneCot := TOB_LignesCot.FindNext ([''],[''],TRUE);
            end;
          end;
//          EnregDucsEdi.MtTotal := TOB_Lignes.Somme('PDD_MTCOTISAT',[''],[''],TRUE, FALSE);
        end;

      end;
  end;
  if (TOB_AssedicGroupee <> NIL) then
  begin
    TOB_AssedicGroupee.Free;
    TOB_AssedicGroupee := NIL;
  end;
  if (TOB_LignesCOT <> nil) then
  begin
    TOB_LignesCOT.Free;
    TOB_LignesCOT := NIL;
  end;

end;
// f PT30 DUCS EDI V4.2
//end.
{$ELSE}
// DUCS 4.1
{***********UNITE*************************************************
Auteur  ...... : PAIE : MF
Créé le ...... : 12/04/2002

Modifié le ... : 22/05/2002
Description .. : Ensemble des procédures de la confection des
Suite ........ : fichiers Ducs Edi dossier
Mots clefs ... : PAIE, PGDUCSEDI
*****************************************************************}
{
 PT1 : 07/08/2002 : V585  MF   PGDUCSEDI
                             1- Traitement des champs code application et serveur
                             unique suite ajout dans table ORGANISMEPAIE.
                             2- traitement du cas UNEDIC : déclaration avec
                             paiement groupé (VLU) Autant de GP2 que d'établis-
                             sements secondaires déclarés --> Création de nouveaux
                             segments NADDCE, QTYTE, RFFDCE et RFFACE
                             3- La date de versement des salaires est renseignée
                             pour les IRC avec décalage de paie et pour l'UNEDIC.
                             Pour l'UNEDIC, s'il sagit d'une déclaration Néant
                             cette date est renseignée à 000000
                             4- Traitement du cas IRC multi centre payeur.
                             On a autant de segments de paiement que
                             d'établissements--> Création de nouveaux segments
                             (par type de paiement)
                             5- Traitement du paiement par virement et
                             modification du paiement par prélèvement ou
                             télérèglement (on récupère le RIB prévu pour le
                             paiement des charges sociales)
                             6- rectification du formatage de la référence sociale
                             utile au contrôle de la clé IBAN (chèque ou virement)
                             7- Correction : qd multi établissement l'édition de
                             contrôle ne faisait pas apparaître le bon
                             établissement.
                             8- Traitement des champs PES_MTPAYE, PES_DUCSNEANT
                             et PES_PAIEMODE ajouter à la table ENVOISOCIAL

 PT2 : 04/09/2002 : V582  MF  (pour version tests de validation)
                             1- Alimentation de la clé IBAN sur segment RFFPQ
 PT3 : 19/09/2002 : V585  MF  (pour version tests de validation)
                             1- correction alimentation segment FII. Récupé-
                             -ration du RIB des SALAIRES à partir de BQ_GENERAL
                             (et non pas BQ_CODE)
 PT4 : 24/09/2002 : V585  MF  (pour version tests de validation)
                             1- Le segment DTM du groupe 6 était renseigné avec
                             un mauvais format de date (CCYYDDMM au lieu de
                             CCYYMMDD)
                             2- Le segment MAOTOT était mal alimenté pour une
                             déclaration UNEDIC mono etab. Le Segement MOAPAI ne faisait
                             pas apparaître la monnaie de tenue.
 PT5 : 22/10/2002 : V585  MF Modification apportées suite au complément du cahier
                             charges (COPAYM V4.1 - mars 1999). Concerne le
                             Tableau Récapitulatif pour l'Acoss.
 PT6 : 23/10/2002 : V585  MF fiche qualité 10252 : le fichier .log est
                             maintenant nommé DUCSEDI+Siret.log
 PT7 : 06/01/2003 : V591  MF
                             1- Correction des avertissements de compile
                             2- Vérification que la nature Ducs est égale à A ou I ou U
                                Sinon anomalie signalée et fichier non généré
                             3- contrôles lors de la génération du fichier (dossier)
                                absence de RIB
                             4- Initialisation de l'enregistrement EnrDucsEdiGroupee
 PT8 : 08/01/2003 : V591  MF
                             1- Correction initialisation et tests sur doubles
                             2- Modification du calcul du montant à payer : =
                                Montant déclaré - acomptes + régularisations
                                (Idem consultation)
                                Modification du traitement des régularisations
                                et des acomptes (IRC uniquement) 1 occcurence
                                pour les régularisations et 1 occurence pour
                                les acomptes
                                Edition sur le fichier de log du montant déclaré,
                                de la régularisation et de l'acompte.
 PT9 : 12/01/2003 : V591  MF
                             1- Correction d'un access Vio quand les champs
                                ET_ADRESS1 et ET_ADRESSE3 étaient rensignés
                                et ET_ADRESSE2 non renseigné (table ETABLIS)
                                Erreur sur le nom du champ utilisé pour
                                l'adresse 3.
                             2- Suite tests validation UNEDIC
                                Correction formatage des périodes mensuelles
                                et de le clé IBAN quan unedic (tjs 99)
 PT10 : 15/01/2003 : V591 MF
                             1- Le segment MOATOT n'est présent pour l'ACOOS
                                que s'il s'agit d'une ducs dossier avec paiement
                                groupé
 PT11 : 31/01/2003 : V591 MF
                             1- Correction anomalie suite portage CWAS
                                Le qualifiant de cotisation n'était pas repris
                                (IRC)
 PT12  : 04/02/2003 : V591 MF
                             1- cas où pour un établissement on a deux
                                déclarations IRC. Un seul fichier
                                Duxyyyyyyyyyyyy.DUC était créé.
                                Pour les caisses IRC le nom du fichier
                                est maintenant bâti comme suit :
                                DUI+SIRET+ABREGE+CODE INSTITUTION
 PT13  : 10/02/2003 : V591 MF
                             1- correction alimentation de la date de règlement
                             des cotisations (prenait à tort la date de paiement
                             des salaires)
 PT14  : 28/02/2003 : V42  MF
                             1- Correction recherche des qualifiants de
                             cotisation
                             2- Rectification traitement des TOB (suite MEMCHECK)
 PT15  : 05/03/2003 : V42  MF (corrections suit tests GBdF pour IRC)
                             1- segment GIS
                             2- segment RFFMR groupe 1
                             3- segment RFFNUM (adhérent) groupe 1
                             4- segment RFFACT (adhérent) groupe 1
                             5- segments Tiers rédacteur groupe 1
                             6- segment MOADUE groupe 2
                             7- segment RFFDC groupe 2 (centre payeur)
                             8- segment MOATOT fin de déclaration
                             9- segment MOAPAI (montant réglé)
 PT16  : 07/03/2003 : V42  MF (corrections suit tests GBdF )
                             1- URSSAF : cotisation transport : une ligne de
                             codification 900 était créée à tort quand pls taux
                             transport.
                             2- IRC : segment DLI gp3 (n° de ligne): Il ne doit
                             pas y avoi de rupture de numérotation entre les
                             centres payeurs.
                             3- Taux de cotisation exprimé avec 4 décimales
                             sans point décimal.
                             4- IRC : Segment CNT
                             5- IRC : Segment QTY+913 (QTEF) présent uniquement
                             pour les forfaits ( M ou G) . Avec 4 décimales.
 PT17 : 11/03/2003 : V42  MF
                             1- UNEDIC : Calcul effectifs Hommmes et Femmes par
                             établissement.
                             2 segments QTY (920 = Hommes, 9210 = Femmes)
PT18  : 17/04/2003 VG V_42  Critère de sélection d'une clause WHERE sans libellé
PT19  : 06/05/2003 : V_42 MF Suite test de validation IRC
                             Les Montant du groupe 3 sont arrondis, exprimés en
                             centimes (forfait non arrondis).
PT20  : 21/07/2003 : V_421 MF Traitement du Centre Payeur (caisses IRC)
PT21  : 23/09/2003 MF V_421
                             1-FQ 10829 : pour les envois DUCS sélection possible
                             de l'émetteur.
                             2- Mise au point CWAS
PT22  : 24/02/2004  : MF V_5.0
                             UNEDIC : prise en compte de la date de changement
                             de taux sur segment DTM du groupe 3. On doit
                             fournir début et fin de période
PT23  : 08/04/2004  : MF V_5.0
                             1- Ajout contrôle n° interne obligatoire pour Acoss
                             et Unedic
                             2- FQ 11250 : Ajout contrôle pour Acoss pas de régul
                             s'il ne s'agit pas d'un TR Annuel
PT24  : 20/04/2004  : MF V_5.0
                             Si le répertoire de stockage n'existe pas le
                             traitement est abandonné  (CheminEagl)
PT25  : 19/01/2005  : MF V_6.0
                             FQ 11752 : Contrôle de la présence de la date
                             de règlement dans le cas d'une ducs IRC.
                             (règle le problème de ducs en partie doublée)
PT26  : 17/02/2005  : MF V_60
                             FQ 12013 : Correction traitement de récupération du
                             RIB de la banque
PT27  : 27/05/2005  : MF v_602
                              FQ 12329 : correction DTM+273 (date de référence)
                              et RFF+CR (Référence paiement)
PT28  : 25/07/2005  : MF V_604
                              FQ 12220  : correction du contrôle du qualifiant
                              de cotisation pour les ducs edi IRC.
PT29  : 15/12/2005  : MF V_650
                              FQ  12761 : Correction Référence du paiement
                              quand ASSEDIC paiement par chèque.
PTA   : 11/05/2006  : MF V_650 : pour ducs 4.1 uniquement : résoud pb URSSAF VLU
                                 livrer également coplib.a  (ducs 4.1) du 10/05/06
PT33  : 16/06/2006 : MF V_70  DUCS EDI V4.1 : memcheck

PT34  : 07/07/2006 : MF V_70  memcheck

PT35 :  27/07/2006 : MF V_70  idem DUCS EDI V4.2 : suite test de labellisation UNEDIC
                              - correction de la référence sociale du paiement RFF+PQ
                              période de référence en cas paie décalée.

}
unit PGDucsEdiOutils;

interface
uses
  {$IFDEF VER150}
  Variants,
  {$ENDIF}
{$IFDEF EAGLCLIENT}
//unused  UtileAGL,
{$ELSE}
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}

  controls,HEnt1,hmsgbox,ShellAPI,Windows,SysUtils,HCtrls, PgOutils,PgOutils2,EntPaie,
  UTOB,  ParamSoc;

Type TDucsEdi = record
     {Infos Liste du multi-critère}
     Etab : string;
     Organisme : string;
     DateDebut : TDateTime;
     DateFin : TDateTime;
     NoDucs : Integer;
     Siret : string;
     NatureOrg : string;
     TypDucs : string;
     Periode : string;
     {Infos DUCSENTETE}
     Abrege : string;
     IdentQual : string;
     IdentEmet : string;
     IdentDest : string;
     DucsDossier : Boolean;
     Declarant : string;
     EmettSoc : string;
     TelephoneDecl : string;
     FaxDeclarant : string;
     DeclarantSuite : string;
     ApePdu : string;
     NbSalFpe : double;
     TotHommes : double;
     TotFemmes : double;
     TotApprenti : double;
     Paiement : TdateTime;
     Reglement : TdateTime; //PT13
     Regularisation : double;
     Acompte : double;
     MtTotRegul : double;
     MtTotAcpte : double;
     NumeroPdu : string;
     MonnaieTenue : string;
     NumeroInterne : string;
     CentrePayeur : string; //PT20
     { Infos DUCSDETAIL}
     MtDeclare  : double;
     MtTransport : double;
     MtTotal : double;
     MtAPayer : double;
     NbTransport : integer;
     NbCot : integer;
     CleIBAN : string;
     ErrTransport : boolean;
     ErrCotisqual : boolean;
     ErrInstitution : boolean;
     ErrCondition : boolean;
     Neant : boolean;
     { Infos ORGANISMEPAIE}
     SiretPog : string;
     LibellePog : string;
     Adresse1Pog : string;
     Adresse2Pog : string;
     VillePog : string;
     CpPog : string;
     ContactPog : string;
     TelPog : string;
     FaxPog : string;
     InstitutionPog : string;
     AdherContact: string;
     NoContEmet : string;
     PaiementGroupe : boolean;
     ModePaiement : string;
     IdentOPS : string;
     CodAppliPog : string;  // PT1-1
     ServUniqPog : boolean; // PT1-1
     { Infos EMETTEURSOCIAL}
     SiretPet : string;
     AdressePet : string;
     VillePet : string;
     CpPet : string;
     Adresse2Pet : string;
     { Infos ETABLISS}
     LibelleET : string;
     Adresse1ET : string;
     Adresse2ET : string;
     VilleET : string;
     CpET : string;
     TelET : string;
     FaxET : string;
     Juridique : string;
     NoCentrePayeur : integer;     // PT20
end;

Type TDucsEdiGroupee = record
     Etab : string;
     Siret : string;
     LibelleET : string;
     Adresse1ET : string;
     Adresse2ET : string;
     VilleET : string;
     CpET : string;
     MtDeclare  : double;
     MtTransport : double;
     NbTransport : integer;
     Regularisation : double;
     Acompte : double;
     MtAPayer : double;
     ErrTransport : boolean;
     ErrCotisqual : boolean;
     ErrInstitution : boolean;
     ErrCondition : boolean;
     ApePdu : string;
     NbSalFpe : double;
     TotHommes : double;
     TotFemmes : double;
     TotApprenti : double;
     Abrege : string;
     MonnaieTenue : string;
     NumeroInterne : string;
     Paiement : TdateTime;
     Reglement : TdateTime; //PT13
     NoDucs : Integer;
     ModePaiement : string;
     CentrePayeur : string; //PT20
end;
// d PT6         procedure OuvFicDucsEdiLOG();
         procedure OuvFicDucsEdiLOG(var Siret : string); 
// f PT6
         procedure FermFicDucsEdiLOG();
         procedure EdiFicDucsEdiLOG(Tsiret : TOB);   // PT6
         procedure OuvreFichierDucsEdi(var EnregDucsEdi : TDucsEdi);
         procedure InitVarDucsPDU(var EnregDucsEdi : TDucsEdi);
         procedure InitVarDucsPDD(var EnregDucsEdi : TDucsEdi);
         procedure InitVarDucsPOG(var EnregDucsEdi : TDucsEdi);
         procedure InitVarDucsPET(var EnregDucsEdi : TDucsEdi);
         procedure InitVarDucsET(var EnregDucsEdi : TDucsEdi);
         procedure InitEffectifsUNEDIC (var EnregDucsEdi : TDucsEdi);
         procedure InitCotisation(var EnregDucsEdi : TDucsEdi;var EnrDucsEdiGroupee : TDucsEdiGroupee;Trait : string);
         procedure EnTeteDeclaration(var EnregDucsEdi : TDucsEdi);
         procedure Groupe1Declaration(var EnregDucsEdi : TDucsEdi);
         procedure Groupe2Declaration (EnregDucsEdi : TDucsEdi; EnrDucsEdiGroupee : TDucsEdiGroupee; Trait : string);
         procedure Groupe3ACOSS(EnregDucsEdi : TDucsEdi; EnrDucsEdiGroupee : TDucsEdiGroupee; Trait : string);
         procedure Groupe4_5ACOSS(EnregDucsEdi : TDucsEdi;EnrDucsEdiGroupee : TDucsEdiGroupee;  Trait : string);
         procedure Groupe3IRC(var EnregDucsEdi : TDucsEdi; EnrDucsEdiGroupee : TDucsEdiGroupee; Trait : string; var NoLg : integer);// PT16-2
         procedure Groupe3UNEDIC(var EnregDucsEdi : TDucsEdi);
         procedure FinDeclaration(var EnregDucsEdi : TDucsEdi;EnrDucsEdiGroupee : TDucsEdiGroupee);
         procedure PaiementEdi(var EnregDucsEdi : TDucsEdi;EnrDucsEdiGroupee : TDucsEdiGroupee; Trait :string; var DossOK : boolean);
// PT7-3         procedure PaiementEdi(var EnregDucsEdi : TDucsEdi;EnrDucsEdiGroupee : TDucsEdiGroupee; Trait :string);
// PT16-2         procedure MultiGp2Gp3(EtabPrinc : string;var EnregDucsEdi : TDucsEdi;var EnrDucsEdiGroupee : TDucsEdiGroupee; var DossOK : boolean);
         procedure MultiGp2Gp3(EtabPrinc : string;var EnregDucsEdi : TDucsEdi;var EnrDucsEdiGroupee : TDucsEdiGroupee; var DossOK : boolean; var NoLg : integer);
         procedure InitVarGp2(var EnregDucsEdi : TDucsEdi; var EnrDucsEdiGroupee : TDucsEdiGroupee;LesDucsGroupee : TOB);
         procedure InitVarGp2UNEDIC(var EnregDucsEdi : TDucsEdi; var EnrDucsEdiGroupee : TDucsEdiGroupee;LesDucsGroupee : TOB);
         procedure FinFin (var EnregDucsEdi : TDucsEdi);
         procedure MajEnvoiSocialDUCS (EnregDucsEdi : TDucsEdi);

         Function  ControlVarDucsEnTete (EnregDucsEdi : TDucsEdi) : Boolean;
         Function  ControlVarDucsGrp1 (EnregDucsEdi : TDucsEdi) : Boolean;
         Function  ControlCotisation(EnregDucsEdi : TDucsEdi;EnrDucsEdiGroupee : TDucsEdiGroupee;Trait : string) : Boolean;
         Function  ControlPaiement (EnregDucsEdi : TDucsEdi) : Boolean;
         Function  CleIBAN (creancier, periode,reference: string;var cleIBAN : string ) : boolean;
         procedure Abandonfichier(Trait : string);

var
         FRapport : TextFile;
         NomFic, NomFicDucs : string;
         FDUCSEDI : TextFile;
         TOB_Lignes, TOB_Transport, TOB_DucsGroupee : TOB;
         PeriodRef : string;
implementation
uses Pg_OutilsEnvoi;
{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 12/04/2002
Modifié le ... : 22/05/2002
Description .. : Ouverture du fichier de log (en append ou cration)
Mots clefs ... : PAIE, PGDUCSEDI
*****************************************************************}
procedure OuvFicDucsEdiLOG(var Siret : String);       // PT6
var
  reponse                       : integer;
  Noerreur                      : integer; // PT24
begin
// d PT21-2
{$IFDEF EAGLCLIENT}
  NomFic := VH_Paie.PGCheminEagl+'\DUCSEDI'+Siret+'.log';
{$ELSE}
  NomFic := V_PGI.DatPath+'\DUCSEDI'+Siret+'.log';    // PT6
{$ENDIF}
// f PT21-2
  if FileExists(NomFic) then
  begin
    reponse := PGIAsk ('Voulez-vous réinitialiser le fichier de contrôle '+NomFic+' ?#13#10'+
                       'Si vous répondez ''NON'', le fichier existant sera complété.',
                       'Confection de fichiers Ducs-EDI');
    if reponse = mrYes then
      DeleteFile(PChar(NomFic));
  end;

  AssignFile(FRapport, NomFic);
  if FileExists(NomFic) then
    Append(FRapport)
  else
  begin
// d PT24
  {$i-}ReWrite(FRapport); {$i+}
    NoErreur := IoResult;
    if NoErreur<>0 Then
      Begin
       PGIBox ('Le répertoire de stockage n''est pas défini', 'Traitement impossible');
       Siret := '';
       Exit ;
      End;
// f PT24
    Writeln(FRapport, '');
    Writeln(FRapport, '*** CONFECTION DE FICHIERS DUCS EDI ***');
    Writeln(FRapport, '---------------------------------------');
  end;

  Writeln(FRapport, '');
  Writeln(FRapport, '********** Début de traitement : '+DateTimeToStr(Now)+
                    '  ******************************************');
end;
{ fin OuvFicDucsEdiLOG}


{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 12/04/2002
Modifié le ... : 22/05/2002
Description .. : fermeture du fichier de log
Mots clefs ... : PAIE, PGDUCSEDI
*****************************************************************}
procedure FermFicDucsEdiLOG();
begin
  Writeln(FRapport, '');
  Writeln(FRapport, '********** Traitement terminé : '+DateTimeToStr(Now)+
                    '  ******************************************');

  CloseFile(FRapport);
end;
{ fin FermFicDucsEdiLOG}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 12/04/2002
Modifié le ... : 22/05/2002
Description .. : lancement de l'édition du fichier de log
Mots clefs ... : PAIE, PGDUCSEDI
*****************************************************************}
procedure EdiFicDucsEdiLOG(TSiret : TOB); // PT6
var
// d PT6
  i, II                         : integer;
  Siret, Anomalie               : string;
  TSiretFille                   : TOB;
  NomFic                        : string;

begin
  for II := 0 to TSiret.Detail.Count-1 do
  begin
    TSiretFille := TSiret.Detail[II];
    Siret := TSiretFille.GetValue('SIRET');
    Anomalie := TSiretFille.GetValue('ANOMALIE') ;
// d PT21-2
{$IFDEF EAGLCLIENT}
    NomFic := '"'+VH_Paie.PGCheminEagl+'\DUCSEDI'+Siret+'.log"';   // PT42
{$ELSE}
    NomFic := V_PGI.DatPath+'\DUCSEDI'+Siret+'.log';
{$ENDIF}
// f PT21-2
// f PT6
    i:=PGIAsk ('Siret : '+Siret+'#13#10 Voulez-vous visualiser le fichier de contrôle ?', 'DUCS EDI') ;
    if i=mrYes then
      i := ShellExecute( 0, PCHAR('open'),PChar('WordPad'), PChar(NomFic),Nil,SW_RESTORE);
  end;
end;
{ fin EdiFicDucsEdiLOG}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 12/04/2002
Modifié le ... :   /  /
Description .. : Initialisation des variables nécessaires à la confection du
Suite ........ : fichier Ducs Edi, issues de la table DUCSENTETE
Mots clefs ... : PAIE, PGDUCSEDI
*****************************************************************}
procedure InitVarDucsPDU(var EnregDucsEdi : TDucsEdi);
var
  QQ                            : TQuery;
begin
  QQ := OpenSQL('SELECT PDU_ABREGEPERIODE,PDU_DUCSDOSSIER,PDU_DECLARANT,'+
                'PDU_EMETTSOC,PDU_TELEPHONEDECL,PDU_FAXDECLARANT,PDU_NUMERO,'+
                'PDU_DECLARANTSUITE,PDU_APE,PDU_NBSALFPE,PDU_TOTHOMMES,'+
                'PDU_TOTFEMMES,PDU_TOTAPPRENTI,PDU_PAIEMENT,PDU_DATEREGLEMENT,PDU_ACOMPTES,'+//PT13
                'PDU_REGULARISATION,PDU_NUMERO,PDU_MONNAIETENUE, '+
                'PDU_CENTREPAYEUR '+ // PT20
                'FROM DUCSENTETE WHERE '+
                'PDU_ETABLISSEMENT = "'+EnregDucsEdi.Etab+'" AND '+
                'PDU_ORGANISME = "'+EnregDucsEdi.Organisme+'" AND '+
                'PDU_DATEDEBUT = "'+UsDateTime(EnregDucsEdi.DateDebut)+'" AND '+
                'PDU_DATEFIN = "'+UsDateTime(EnregDucsEdi.DateFin)+'" AND '+
                'PDU_NUM = '+IntToStr(EnregDucsEdi.NoDucs)+'',True) ;        // DB2
  if Not QQ.EOF then
  begin
    EnregDucsEdi.Abrege:=QQ.FindField('PDU_ABREGEPERIODE').AsString;
    EnregDucsEdi.DucsDossier:=QQ.FindField('PDU_DUCSDOSSIER').AsString='X';
    EnregDucsEdi.Declarant:=QQ.FindField('PDU_DECLARANT').AsString;
    EnregDucsEdi.EmettSoc:=QQ.FindField('PDU_EMETTSOC').AsString;
    EnregDucsEdi.TelephoneDecl:=QQ.FindField('PDU_TELEPHONEDECL').AsString;
    EnregDucsEdi.FaxDeclarant:=QQ.FindField('PDU_FAXDECLARANT').AsString;
    EnregDucsEdi.NumeroInterne:=QQ.FindField('PDU_NUMERO').AsString;
    EnregDucsEdi.DeclarantSuite:=QQ.FindField('PDU_DECLARANTSUITE').AsString;
    EnregDucsEdi.ApePdu:=QQ.FindField('PDU_APE').AsString;
    EnregDucsEdi.NbSalFpe:=QQ.FindField('PDU_NBSALFPE').AsFloat;
    EnregDucsEdi.TotHommes:=QQ.FindField('PDU_TOTHOMMES').AsFloat;
    EnregDucsEdi.TotFemmes:=QQ.FindField('PDU_TOTFEMMES').AsFloat;
    EnregDucsEdi.TotApprenti:=QQ.FindField('PDU_TOTAPPRENTI').AsFloat;
    EnregDucsEdi.Paiement:=QQ.FindField('PDU_PAIEMENT').AsDateTime;
    EnregDucsEdi.Reglement:=QQ.FindField('PDU_DATEREGLEMENT').AsDateTime; // PT13
    EnregDucsEdi.Regularisation:=QQ.FindField('PDU_REGULARISATION').AsFloat;
    EnregDucsEdi.Acompte:=QQ.FindField('PDU_ACOMPTES').AsFloat;
    EnregDucsEdi.NumeroPdu:=QQ.FindField('PDU_NUMERO').AsString;
    EnregDucsEdi.MonnaieTenue:=QQ.FindField('PDU_MONNAIETENUE').AsString;
    // d PT20
    if (QQ.FindField('PDU_CENTREPAYEUR').AsString <> '') then
      EnregDucsEdi.CentrePayeur:=QQ.FindField('PDU_CENTREPAYEUR').AsString;
    // f PT20

  end ;
  Ferme(QQ) ;
end;
{ fin InitVarDucsPDU}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 13/05/2002
Modifié le ... :   /  /
Description .. : Initialisation des variables nécessaires à la confection
Suite ........ : du fichier Ducs Edi issues de la table DUCSDETAIL
Mots clefs ... : PAIE, PGDUCSEDI
*****************************************************************}
procedure InitVarDucsPDD(var EnregDucsEdi : TDucsEdi);
var
  QQ : TQuery;
  StQQ : string;

begin
  StQQ := 'PDD_ETABLISSEMENT = "'+EnregDucsEdi.Etab+'" AND '+
          'PDD_ORGANISME = "'+EnregDucsEdi.Organisme+'" AND '+
          'PDD_DATEDEBUT = "'+UsDateTime(EnregDucsEdi.DateDebut)+'" AND '+
          'PDD_DATEFIN = "'+UsDateTime(EnregDucsEdi.DateFin)+'" AND '+
          'PDD_NUM = '+IntToStr(EnregDucsEdi.NoDucs)+' AND '+          // DB2
          'PDD_TYPECOTISATION <> "S"';
  QQ := OpenSQL('SELECT SUM(PDD_MTCOTISAT) AS DECLAR FROM DUCSDETAIL WHERE '+StQQ,True) ;
  EnregDucsEdi.MtDeclare := 0.0; // PT8-1
  EnregDucsEdi.MtTotal := 0.0;   // PT8-1
  EnregDucsEdi.MtAPayer := 0.0;  // PT8-1

  if not QQ.EOF then // PortageCWAS
    EnregDucsEdi.MtDeclare := QQ.FindField('DECLAR').asFloat;
  EnregDucsEdi.MtTotal := EnregDucsEdi.MtTotal + EnregDucsEdi.MtDeclare;
{ PT8-2     EnregDucsEdi.MtAPayer := EnregDucsEdi.MtAPayer+
                              EnregDucsEdi.MtDeclare-
                              EnregDucsEdi.Regularisation-
                              EnregDucsEdi.Acompte;  }
  EnregDucsEdi.MtAPayer := EnregDucsEdi.MtAPayer+
                           EnregDucsEdi.MtDeclare+
                           EnregDucsEdi.Regularisation-
                           EnregDucsEdi.Acompte;
  EnregDucsEdi.MtTotRegul :=0.0; // PT8-1
  EnregDucsEdi.MtTotAcpte :=0.0; // PT8-1

  EnregDucsEdi.MtTotRegul :=  EnregDucsEdi.MtTotRegul + EnregDucsEdi.Regularisation;
  EnregDucsEdi.MtTotAcpte :=  EnregDucsEdi.MtTotAcpte + EnregDucsEdi.Acompte;

  Ferme (QQ);
  StQQ := '';
  StQQ := 'PDD_ETABLISSEMENT = "'+EnregDucsEdi.Etab+'" AND '+
          'PDD_ORGANISME = "'+EnregDucsEdi.Organisme+'" AND '+
          'PDD_DATEDEBUT = "'+UsDateTime(EnregDucsEdi.DateDebut)+'" AND '+
          'PDD_DATEFIN = "'+UsDateTime(EnregDucsEdi.DateFin)+'" AND '+
          'PDD_NUM = '+IntToStr(EnregDucsEdi.NoDucs)+' AND '+  // DB2
          'PDD_TYPECOTISATION <> "S" AND PDD_TYPECOTISATION <> "I" '+
          'AND PDD_CODIFICATION LIKE "1%900P"';
  QQ := OpenSQL('SELECT COUNT(DISTINCT(PDD_TAUXCOTISATION)) AS NBTRANSPORT, '+
                'SUM(PDD_MTCOTISAT) AS MTTRANSPORT FROM DUCSDETAIL WHERE '+StQQ,True) ;
  EnregDucsEdi.MtTransport:=0.0; // PT8-1
  if not QQ.EOF then // PortageCWAS
  begin
    EnregDucsEdi.MtTransport:= QQ.FindField('MTTRANSPORT').asFloat;
    EnregDucsEdi.NbTransport:= QQ.FindField('NBTRANSPORT').AsInteger;
  end;
  Ferme(QQ);
  EnregDucsEdi.ErrTransport := False;
  if (copy(EnregDucsEdi.TypDucs,1,1) = 'A') and
     (EnregDucsEdi.NbTransport > 1) then
  { Acoss - Création TOB_Transport}
  begin
    QQ := OpenSQL('SELECT Count (*) FROM DUCSDETAIL WHERE (PDD_COMURBAINE="" '+
                  'OR PDD_COMURBAINE IS NULL OR PDD_COMURBAINE="    ") and '+
                  StQQ,TRUE);
    if not QQ.EOF then
    begin
      if (QQ.Fields[0].AsInteger <> 0) then
        EnregDucsEdi.ErrTransport :=  True;
    end;
    ferme (QQ);

    QQ := OpenSQL('SELECT * FROM DUCSDETAIL WHERE '+StQQ,True);

    TOB_Transport := TOB.Create ('Les lignes de tansport', NIL, -1);

    TOB_Transport.LoadDetailDB ('DUCSDETAIL', '','',QQ,FALSE,FALSE);
    Ferme(QQ);
    TOB_Transport.Detail.Sort('PDD_ETABLISSEMENT;PDD_ORGANISME;PDD_DATEDEBUT;'+
                              'PDD_DATEFIN;PDD_INSTITUTION;PDD_CODIFICATION;PDD_DATECHGTTAUX');
  end;
end;
{ fin InitVarDucsPDD}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 12/04/2002
Modifié le ... :   /  /
Description .. : Initialisation des variables nécessaires à la confection du
Suite ........ : fichier Ducs Edi, issues de la table ORGANISMEPAIE
Mots clefs ... : PAIE, PGDUCSEDI
*****************************************************************}
procedure InitVarDucsPOG(var EnregDucsEdi : TDucsEdi);
var
  QQ : TQuery;
begin

  QQ := OpenSQL('SELECT POG_NATUREORG,POG_IDENTQUAL,POG_IDENTEMET,'+
                'POG_IDENTDEST,POG_NOCONTEMET,POG_SIRET,POG_LIBELLE,'+
                'POG_ADRESSE1,POG_ADRESSE2, POG_ADRESSE3,POG_VILLE,'+
                'POG_CODEPOSTAL,POG_CONTACT,POG_TELEPHONE,POG_FAX,'+
                'POG_INSTITUTION,POG_ADHERCONTACT,POG_PAIEGROUPE, '+
                'POG_PAIEMODE,POG_IDENTOPS,POG_CODAPPLI,POG_SERVUNIQ '+ // PT1-1
                'FROM ORGANISMEPAIE WHERE '+
                'POG_ETABLISSEMENT = "'+EnregDucsEdi.Etab+'" AND '+
                'POG_ORGANISME = "'+EnregDucsEdi.Organisme+'"',True) ;
  if not QQ.EOF then
  begin
    EnregDucsEdi.NatureOrg:=QQ.FindField('POG_NATUREORG').AsString;
    EnregDucsEdi.IdentQual:=QQ.FindField('POG_IDENTQUAL').AsString;
    EnregDucsEdi.IdentEmet:=QQ.FindField('POG_IDENTEMET').AsString;
    EnregDucsEdi.IdentDest:=QQ.FindField('POG_IDENTDEST').AsString;
    EnregDucsEdi.SiretPog:=QQ.FindField('POG_SIRET').AsString;
    EnregDucsEdi.LibellePog:=QQ.FindField('POG_LIBELLE').AsString;

    if (QQ.FindField('POG_ADRESSE1').AsString <> '') then
      EnregDucsEdi.Adresse1Pog:=QQ.FindField('POG_ADRESSE1').AsString;
    if (QQ.FindField('POG_ADRESSE2').AsString <> '') then
      if (EnregDucsEdi.Adresse1Pog  <> '') then
        EnregDucsEdi.Adresse2Pog:=QQ.FindField('POG_ADRESSE2').AsString
      else
        EnregDucsEdi.Adresse1Pog:=QQ.FindField('POG_ADRESSE2').AsString;
    if (QQ.FindField('POG_ADRESSE3').AsString <> '') then
      if (EnregDucsEdi.Adresse1Pog  = '') then
        EnregDucsEdi.Adresse1Pog:=QQ.FindField('POG_ADRESSE3').AsString
      else
        if (EnregDucsEdi.Adresse2Pog  = '') then
          EnregDucsEdi.Adresse2Pog:=QQ.FindField('POG_ADRESSE3').AsString;

    EnregDucsEdi.VillePog:=QQ.FindField('POG_VILLE').AsString;
    EnregDucsEdi.CpPog:=QQ.FindField('POG_CODEPOSTAL').AsString;
    EnregDucsEdi.ContactPog:=QQ.FindField('POG_CONTACT').AsString;
    EnregDucsEdi.TelPog:=QQ.FindField('POG_TELEPHONE').AsString;
    EnregDucsEdi.FaxPog:=QQ.FindField('POG_FAX').AsString;
    EnregDucsEdi.InstitutionPog:=QQ.FindField('POG_INSTITUTION').AsString;
    EnregDucsEdi.AdherContact:=QQ.FindField('POG_ADHERCONTACT').AsString;
    EnregDucsEdi.NoContEmet:=QQ.FindField('POG_NOCONTEMET').AsString;
    EnregDucsEdi.PaiementGroupe:=QQ.FindField('POG_PAIEGROUPE').AsString='X';
    EnregDucsEdi.ModePaiement:=QQ.FindField('POG_PAIEMODE').AsString;
    EnregDucsEdi.IdentOPS:=QQ.FindField('POG_IDENTOPS').AsString;

// d PT1-1
    EnregDucsEdi.CodAppliPog:=QQ.FindField('POG_CODAPPLI').AsString;
    if (EnregDucsEdi.CodAppliPog = '') then
      EnregDucsEdi.CodAppliPog := '013';
    EnregDucsEdi.ServUniqPog:=QQ.FindField('POG_SERVUNIQ').AsString='X';
// f PT1-1

  end ;
  Ferme(QQ) ;
  if (EnregDucsEdi.NatureOrg = '100') then
    EnregDucsEdi.TypDucs := 'A ACOSS ';         { Acoss}
  if (EnregDucsEdi.NatureOrg = '200') then
    EnregDucsEdi.TypDucs := 'U UNEDIC';         {Unedic}
  if (EnregDucsEdi.NatureOrg = '300') then
    EnregDucsEdi.TypDucs := 'I AGIRC ';         { Irc}
end;
{ fin InitVarDucsPOG}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 15/04/2002
Modifié le ... :   /  /
Description .. : Initialisation des variables nécessaires à la confection du
Suite ........ : fichier Ducs Edi, issues de la table  EMETTEURSOCIAL
Mots clefs ... : PAIE, PGDUCSEDI
*****************************************************************}
procedure InitVarDucsPET(var EnregDucsEdi : TDucsEdi);
var
  QQ : TQuery;
begin

  QQ := OpenSQL('SELECT PET_SIRET,PET_COMPLADR,PET_VILLE,PET_CODEPOSTAL,'+
                'PET_COMPLADR2 '+
                'FROM EMETTEURSOCIAL WHERE '+
                'PET_EMETTSOC = "'+EnregDucsEdi.EmettSoc+'" ',True) ;
  if not QQ.EOF then
  begin
    EnregDucsEdi.SiretPet:=QQ.FindField('PET_SIRET').AsString;
    EnregDucsEdi.AdressePet:=QQ.FindField('PET_COMPLADR').AsString;
    EnregDucsEdi.VillePet:=QQ.FindField('PET_VILLE').AsString;
    EnregDucsEdi.CpPet:=QQ.FindField('PET_CODEPOSTAL').AsString;
    EnregDucsEdi.Adresse2Pet:=QQ.FindField('PET_COMPLADR2').AsString;
  end ;
  Ferme(QQ) ;
end;
{ fin InitVarDucsPET}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 15/04/2002
Modifié le ... :   /  /
Description .. : Initialisation des variables nécessaires à la confection du
Suite ........ : fichier Ducs Edi, issues de la table  ETABLISS
Mots clefs ... : PAIE, PGDUCSEDI
*****************************************************************}
procedure InitVarDucsET(var EnregDucsEdi : TDucsEdi);
var
  QQ : TQuery;
begin

  QQ := OpenSQL('SELECT ET_LIBELLE,ET_ADRESSE1,ET_ADRESSE2,ET_ADRESSE3,'+
                'ET_VILLE,ET_CODEPOSTAL,ET_TELEPHONE,ET_FAX,ET_JURIDIQUE '+
                'FROM ETABLISS WHERE '+
                'ET_ETABLISSEMENT = "'+EnregDucsEdi.Etab+'" ',True) ;
  if not QQ.EOF then
  begin
    EnregDucsEdi.LibelleET:=QQ.FindField('ET_LIBELLE').AsString;

    if (QQ.FindField('ET_ADRESSE1').AsString <> '') then
      EnregDucsEdi.Adresse1ET:=QQ.FindField('ET_ADRESSE1').AsString;
    if (QQ.FindField('ET_ADRESSE2').AsString <> '') then
      if (EnregDucsEdi.Adresse1ET  <> '') then
        EnregDucsEdi.Adresse2ET:=QQ.FindField('ET_ADRESSE2').AsString
      else
        EnregDucsEdi.Adresse1ET:=QQ.FindField('ET_ADRESSE2').AsString;
    if (QQ.FindField('ET_ADRESSE3').AsString <> '') then
      if (EnregDucsEdi.Adresse1ET  = '') then
        EnregDucsEdi.Adresse1ET:=QQ.FindField('ET_ADRESSE3').AsString
      else
        if (EnregDucsEdi.Adresse2ET  = '') then
          EnregDucsEdi.Adresse2ET:=QQ.FindField('ET_ADRESSE3').AsString;// PT9-1

    EnregDucsEdi.VilleET:=QQ.FindField('ET_VILLE').AsString;
    EnregDucsEdi.CpET:=QQ.FindField('ET_CODEPOSTAL').AsString;
    EnregDucsEdi.TelET:=QQ.FindField('ET_TELEPHONE').AsString;
    EnregDucsEdi.FaxET:=QQ.FindField('ET_FAX').AsString;
    EnregDucsEdi.Juridique:=QQ.FindField('ET_JURIDIQUE').AsString;
  end ;
  Ferme(QQ) ;
end;
{ fin InitVarDucsET}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 28/05/2002
Modifié le ... :   /  /
Description .. : Calcul des effectifs pour l'établissement principal seulement .
Suite ........ : Il s'agit d'une déclaration UNEDIC avec VLU, lors de
Suite ........ : l'initialisation de la ducs les effectifs calculés correspondent
Suite ........ : aux effectifs tous établissements confondus.
Mots clefs ... : PAIE, PGDUCSEDI
*****************************************************************}
procedure InitEffectifsUNEDIC (var EnregDucsEdi : TDucsEdi);
var
  StQL                  : string;
  QL                    : Tquery;
begin
  {Calcul des Effectifs}
  StQL := '';
  StQL :='PSA_ETABLISSEMENT="'+EnregDucsEdi.Etab+'" AND ';

  StQL := StQL+'(PSA_DATESORTIE="'+UsDateTime(IDate1900)+'"'+
               ' OR  PSA_DATESORTIE IS NULL '+
               ' OR PSA_DATESORTIE>="'+UsDateTime(EnregDucsEdi.DateDebut)+'")'+
               ' AND PSA_DATEENTREE<="'+UsDateTime(EnregDucsEdi.DateFin)+'"';

  {nbre d'hommes}
  QL := OpenSQL('SELECT Count(*) FROM SALARIES WHERE '+StQL+
                ' AND PSA_SEXE="M"',True) ;
  if not QL.EOF then // PortageCWAS
    EnregDucsEdi.TotHommes := QL.Fields[0].AsInteger ;
  Ferme(QL) ;

  { nbre de femmes}
  QL := OpenSQL('SELECT Count(*) FROM SALARIES WHERE '+StQL+
                ' AND PSA_SEXE="F"',True) ;
  if not QL.EOF then // PortageCWAS
    EnregDucsEdi.TotFemmes := QL.Fields[0].AsInteger ;
  Ferme(QL) ;

  {Nbre d'apprentis}
  QL := OpenSQL('SELECT Count(*) FROM SALARIES WHERE '+StQL+
                ' AND PSA_CATDADS="003"',True) ;
  if not QL.EOF then // PortageCWAS
    EnregDucsEdi.TotApprenti := QL.Fields[0].AsInteger ;
  Ferme(QL) ;

  {Nbre de salariés rémunérés pour la période}
   StQL := '';
   StQL :='PHB_ETABLISSEMENT="'+EnregDucsEdi.Etab+'" AND ';
   StQL := StQL+ 'PHB_ORGANISME="'+EnregDucsEdi.Organisme+'" ';

   StQL := StQL+'AND '+
                'PHB_DATEDEBUT >="'+UsDateTime(EnregDucsEdi.DateDebut)+'" AND '+
                'PHB_DATEFIN <="'+UsDateTime(EnregDucsEdi.DateFin)+'" AND '+
                '(PHB_BASECOT <> 0 OR PHB_MTSALARIAL <> 0 OR PHB_MTPATRONAL <> 0)';

   QL := OpenSQL('SELECT Count(DISTINCT PHB_SALARIE) AS NOMBRE FROM HISTOBULLETIN WHERE '+
                 StQL,True) ;
   if not QL.EOF then // PortageCWAS
     EnregDucsEdi.NbSalFpe := QL.FindField('NOMBRE').AsInteger ;
   Ferme(QL) ;
end;
{ fin InitEffectifsUNEDIC}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 13/05/2002
Modifié le ... :   /  /
Description .. : Chargement de la TOB des lignes détails à traiter +
Suite ........ : contrôles sur  champs obligatoires
Mots clefs ... : PAIE, PGDUCSEDI
*****************************************************************}
procedure  InitCotisation(var EnregDucsEdi : TDucsEdi;var EnrDucsEdiGroupee : TDucsEdiGroupee;Trait : string);
var
   QL                           : TQuery;
   StQL                         : string;
   Erreur                       : boolean;
   LesLignes                    : TOB;
   Wetab,WNoDucs                : string;
   Pred, QualifiantCotis        : string; // PT11-1
   DebCodif, SuitCodif  : string;        // PT14-1
begin
  Erreur := false;
  EnregDucsEdi.Neant := False;
// d PT7-4
{with EnrDucsEdiGroupee do
  begin
     Etab := '';
     Siret := '';
     LibelleET := '';
     Adresse1ET := '';
     Adresse2ET := '';
     VilleET := '';
     CpET := '';
     MtDeclare  := 0.0;
     MtTransport := 0.0;
     NbTransport := 0;
     Regularisation := 0.0;
     Acompte := 0.0;
     MtAPayer := 0.0;
     ErrTransport := False;
     ErrCotisqual := False;
     ErrInstitution := False;
     ErrCondition := False;
     ApePdu := '';
     NbSalFpe := 0.0;
     TotHommes := 0.0;
     TotFemmes := 0.0;
     TotApprenti := 0.0;
     Abrege := '';
     MonnaieTenue := '';
     NumeroInterne := '';
     Paiement := IDate1900;
     NoDucs := 0;
     ModePaiement := '';
  end;  }
// f PT7-4

  if (Trait <> 'G') then
  begin
    EnregDucsEdi.ErrCotisqual := False;
    EnregDucsEdi.ErrInstitution := False;
    EnregDucsEdi.ErrCondition := False;
    Wetab := EnregDucsEdi.Etab;
    WNoducs := IntToStr(EnregDucsEdi.NoDucs);
  end
  else
  begin
    EnrDucsEdiGroupee.ErrCotisqual := False;
    EnrDucsEdiGroupee.ErrInstitution := False;
    EnrDucsEdiGroupee.ErrCondition := False;
    Wetab := EnrDucsEdiGroupee.Etab;
    WNoducs := IntToStr(EnrDucsEdiGroupee.NoDucs);
  end;

  if (copy(EnregDucsEdi.TypDucs,1,1) = 'I') then
  { IRC : Contrôle alimentation des champs PDD_INSTITUTION et PDD_CONDITION}
  {       non renseignés}
  begin
    StQL := 'PDD_ETABLISSEMENT = "'+WEtab+'" AND '+
            'PDD_ORGANISME = "'+EnregDucsEdi.Organisme+'" AND '+
            'PDD_DATEDEBUT = "'+UsDateTime(EnregDucsEdi.DateDebut)+'" AND '+
            'PDD_DATEFIN = "'+UsDateTime(EnregDucsEdi.DateFin)+'" AND '+
            'PDD_NUM = '+ WNoDucs+' AND '+
            'PDD_TYPECOTISATION <> "S" AND PDD_TYPECOTISATION <> "I"';
    QL := OpenSQL('SELECT COUNT(*)FROM DUCSDETAIL WHERE '+StQL+
                  ' AND (PDD_INSTITUTION = "" OR PDD_INSTITUTION IS NULL '+
                  'OR PDD_INSTITUTION = "    ")',True) ;
    if not QL.EOF then
    begin
      if (QL.Fields[0].AsInteger <> 0) then
      begin
        if (Trait <> 'G') then
          EnregDucsEdi.ErrInstitution :=  True
        else
          EnrDucsEdiGroupee.ErrInstitution :=  True;
        Erreur := true;
      end;
    end;
    ferme (QL);
    QL := OpenSQL('SELECT COUNT(*)FROM DUCSDETAIL WHERE '+StQL+
                  ' AND (PDD_CONDITION = "" OR PDD_CONDITION IS NULL '+
                  'OR PDD_CONDITION = "    ")',True) ;
    if not QL.EOF then
    begin
      if (QL.Fields[0].AsInteger <> 0) then
      begin
        if (Trait <>'G') then
          EnregDucsEdi.ErrCondition :=  True
        else
          EnrDucsEdiGroupee.ErrCondition :=  True;

        Erreur := true;
      end;
    end;
    ferme (QL);
  end;

  if (Erreur = false) then
  { Aucune erreur détectée}
  begin
// d PT14-2
    if (TOB_Lignes <> NIL) then
    begin
      TOB_Lignes.Free;
      TOB_Lignes := NIL;
    end;
// f PT14-2
    { Chargement de la TOB des Lignes Détail}
    TOB_Lignes := TOB.Create ('Les lignes de cotisation', NIL, -1);

    QL :=OpenSql ('SELECT * FROM DUCSDETAIL '+
                  'WHERE '+
                  '(PDD_ETABLISSEMENT ="'+ WEtab +'") AND '+
                  '(PDD_ORGANISME ="'+ EnregDucsEdi.Organisme+'") AND '+
                  '(PDD_DATEDEBUT ="'+UsDateTime(EnregDucsEdi.DateDebut)+'") AND '+
                  '(PDD_DATEFIN = "'+UsDateTime(EnregDucsEdi.DateFin)+'") AND '+
                  '(PDD_NUM = '+WNoDucs+')', TRUE);

    TOB_Lignes.LoadDetailDB ('DUCSDETAIL', '','',QL,FALSE,FALSE);
    Ferme(QL);

    TOB_Lignes.Detail.Sort('PDD_ETABLISSEMENT;PDD_ORGANISME;PDD_DATEDEBUT;'+
                           'PDD_DATEFIN;PDD_INSTITUTION;PDD_CODIFICATION;PDD_DATECHGTTAUX');

    if (TOB_Lignes.FillesCount(1) = 0) then EnregDucsEdi.Neant := True;

    if (copy (EnregDucsEdi.TypDucs,1,1)= 'I') and
       (TOB_Lignes <> NIL) then
    {IRC : Le champ PDP_COTISQUAL de la table DUCSPARAM doit être
           renseigné                                             }
    begin
      LesLignes := TOB_Lignes.FindFirst ([''],[''],TRUE);
      While LesLignes <> NIL  do
      begin
        if (Leslignes.getValue('PDD_TYPECOTISATION') <> 'I') and
           (Leslignes.getValue('PDD_TYPECOTISATION') <> 'S') then
        { les lignes Intitulé et Sous-Total ne sont pas traitées.}
        begin
// // PT14-1
          DebCodif := copy(Leslignes.getValue('PDD_CODIFICATION'),1,2);
          SuitCodif := copy(Leslignes.getValue('PDD_CODIFICATION'),3,5);
// PT28   while (((QualifiantCotis = '') or
          if (((QualifiantCotis = '') or
                  (QualifiantCotis = NULL) or
                  (QualifiantCotis = '   ')) and
// PT28          (StrToInt(copy(DebCodif,2,1)) >= 0)) do
                 (StrToInt(copy(DebCodif,2,1)) >= 0)) then
          begin
            QL := OpenSql('SELECT PDP_PREDEFINI,PDP_COTISQUAL'+
                          ' FROM DUCSPARAM '+
                          ' WHERE ##PDP_PREDEFINI## PDP_CODIFICATION = "'+
// PT14- 1                         Leslignes.getValue('PDD_CODIFICATION')+'"'+
                          DebCodif +  SuitCodif +'"'+
                          'ORDER BY PDP_PREDEFINI',True);
            while  not QL.eof do
            begin
              Pred :=  QL.FindField('PDP_PREDEFINI').AsString; // PT11-1
              QualifiantCotis := QL.FindField('PDP_COTISQUAL').AsString;  // PT11-1
              if (QL.FindField('PDP_PREDEFINI').AsString = 'DOS') then break;
                QL.Next;
            end;
            ferme (QL);
//PT28      DebCodif := IntToStr(StrToInt(DebCodif)-1);
          end;
// PT14-1
          if (Pred  <> '') then
            if (QualifiantCotis = '') or
               (QualifiantCotis = NULL) or
               (QualifiantCotis = '   ') then
              if (Trait <> 'G') then
                EnregDucsEdi.ErrCotisqual := True
              else
                EnrDucsEdiGroupee.ErrCotisqual := True;
// PT14-1               ferme (QL);
        end;
        LesLignes := TOB_Lignes.FindNext ([''],[''],TRUE);
      end;
    end;
  end;
end;
{ fin InitCotisation}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 12/04/2002
Modifié le ... :   /  /
Description .. : Contrôle de la validité des champs nécessaire à l'en tête de
Suite ........ : la déclaration
Mots clefs ... : PAIE, PGDUCSEDI
*****************************************************************}
Function  ControlVarDucsEnTete (EnregDucsEdi : TDucsEdi) : Boolean;
var
  i                            : integer;
  st, tiret                    : string;
  An, Mois,Trimestre,Libperiod : string;
// PT7-1   IntMois, IntTrim, NumMois : integer;
  NumMois                       : integer;
begin
  result := True;

  tiret := '';
  Writeln(FRapport, '');
  if (EnregDucsEdi.DucsDossier = true) then
    st :=  EnregDucsEdi.Etab+' '+EnregDucsEdi.LibelleET+' (MULTI)'
  else
    st := EnregDucsEdi.Etab+' '+EnregDucsEdi.LibelleET+' (MONO)';

  libperiod := '';
  An := copy(EnregDucsEdi.Abrege,1,2);
  Mois := copy (EnregDucsEdi.Abrege,4,1);
// PT7-1   IntMois:=StrToInt(Mois);
  Trimestre := copy (EnregDucsEdi.Abrege,3,1);
// PT7-1   IntTrim:=StrToInt(Trimestre);
  if (Mois <> '0') and (trimestre<>'0') then
  {Il s'agit d'une déclaration mensuelle}
  begin
    NumMois:=((StrToInt(Trimestre)-1)*3)+StrToInt(Mois);
    Mois:=IntToStr(NumMois);
    if Length(Mois)=1 then Mois:='0'+Mois;
    libperiod := RechDom ('PGMOIS', Mois,FALSE) +' 20'+An ;
  end
  else
    if (Mois='0') and (Trimestre<>'0') then
    { Il s'agit d'une déclaration trimestrielle}
    begin
      LibPeriod := Trimestre+' TRIMESTRE 20'+An;
    end
    else
      if ((Mois = '0') and (Trimestre = '0')) then
      begin
      {Il s'agit d'une déclaration annuelle}
        libperiod := 'ANNEE 20'+An;
      end;
  libperiod := Uppercase(libperiod);

  if (Copy(EnregDucsEdi.TypDucs,1,1) <> 'I') then
  st := st+'  '+Libperiod+'  Déclaration '+copy(EnregDucsEdi.TypDucs,3,6)+
        ' pour '+EnregDucsEdi.LibellePog
  else
    st := st+'  '+Libperiod+'  Déclaration IRC pour '+
          EnregDucsEdi.LibellePog ;


  Writeln(FRapport,st);

  for i:=0 to (Length(EnregDucsEdi.Etab)+length(EnregDucsEdi.LibelleET)+9) do
    tiret := tiret + '-';
  Writeln(FRapport, tiret);

   { PT12 formatage du nom de fichier }
  if (Copy(EnregDucsEdi.TypDucs,1,1) = 'I') then
  { Caisse IRC : on ajoute le code institution }
    NomFicDucs := 'DU'+
                  Copy(EnregDucsEdi.TypDucs,1,1)+
                  EnregDucsEdi.Siret+
                  EnregDucsEdi.Abrege+
                  EnregDucsEdi.InstitutionPog+
                  '.DUC'
  else
  { Caisse non IRC}
    NomFicDucs := 'DU'+
                  Copy(EnregDucsEdi.TypDucs,1,1)+
                  EnregDucsEdi.Siret+
                  EnregDucsEdi.Abrege+
                  '.DUC';

  Writeln(FRapport, '');
  Writeln(FRapport,'  Fichier : '+NomFicDucs);
  Writeln(FRapport, '');
// d PT7-2
  if (Copy(EnregDucsEdi.TypDucs,1,1) <> 'A') and
     (Copy(EnregDucsEdi.TypDucs,1,1) <> 'I') and
     (Copy(EnregDucsEdi.TypDucs,1,1) <> 'U') then
  begin
    Writeln(FRapport, '  Anomalie   : L''organisme traité n''est pas de nature DUCS-EDI');
    Writeln(FRapport, '  --> Action : Vérification de la fiche organisme.');
    result := False;
  end;
// f PT7-2


  { Qualifiant (Segment UNB)}
  if (EnregDucsEdi.IdentQual <> '5') and
     (EnregDucsEdi.IdentQual <> '22') and
     (EnregDucsEdi.IdentQual <> 'ZZZ') then
  begin
    Writeln(FRapport, '  Anomalie   : Le qualifiant n''est pas renseigné');
    Writeln(FRapport, '  --> Action : Modification de la fiche organisme.');
    result := False;
  end;
  {Identifiant émetteur (Segment UNB)}
  if (EnregDucsEdi.IdentEmet = '') then
  begin
    Writeln(FRapport, '  Anomalie   : L''identifiant émetteur n''est pas renseigné');
    Writeln(FRapport, '  --> Action : Modification de la fiche organisme.');
    result := False;
  end;
  {Identifiant destinataire (Segment UNB)}
  if (EnregDucsEdi.IdentDest = '') then
  begin
    Writeln(FRapport, '  Anomalie   : L''identifiant destinataire n''est pas renseigné');
    Writeln(FRapport, '  --> Action : Modification de la fiche organisme.');
   result := False;
  end;
end;
{fin ControlVarDucsEnTete}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 15/04/2002
Modifié le ... :   /  /
Description .. : message d'abandon dans le fichier de log
Mots clefs ... : PAIE, PGDUCSEDI
*****************************************************************}
procedure Abandonfichier(Trait : string);
var
  Nomcomplet                    : string;
begin
  Writeln(FRapport, '');
  Writeln(FRapport, '!! La confection du fichier est abandonnée.');
  if (Trait = 'G') then
  begin
    CloseFile(FDUCSEDI);
// d PT21-2
{$IFDEF EAGLCLIENT}
    Nomcomplet := VH_Paie.PGCheminEagl+'\'+NomFicDucs;
{$ELSE}
    Nomcomplet := V_PGI.DatPath+'\'+NomFicDucs;
{$ENDIF}
// f PT21-2
    DeleteFile(PChar(Nomcomplet));
  end;
// d PT33  (suite memcheck on vide les tob)
    TOB_Lignes.Free; TOB_Lignes := NIL;

    if (TOB_Transport <> Nil) then
    begin
      TOB_Transport.Free;
      TOB_Transport := NIL;
    end;

//PT34    TOB_DucsGroupee.Free; TOB_DucsGroupee := NIL;
      FreeAndNIl(TOB_DucsGroupee);

// f PT33

end;
{Abandonfichier}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 15/04/2002
Modifié le ... :   /  /
Description .. : Contrôle de validité des champs nécessaires au
Suite ........ : groupe 1 de la déclaration
Mots clefs ... : PAIE, PGDUCSEDI
*****************************************************************}
Function  ControlVarDucsGrp1 (EnregDucsEdi : TDucsEdi) : Boolean;
begin
  result := True;

  { Infos EMETTEUR
    --------------
  Nom du déclarant(émetteur) (segment NADFR)}
  if (EnregDucsEdi.Declarant ='') then
  begin
    Writeln(FRapport, '  Anomalie   : Le nom de l''émetteur n''est pas '+
                      'renseigné');
    Writeln(FRapport, '  --> Action : En consultation des Ducs renseigner '+
                      'les coordonnées du déclarant.');
    result := False;
  end;
  { Nom du contact chez l'émetteur (segment CTAFR)}
  if (EnregDucsEdi.DeclarantSuite ='') then
  begin
    Writeln(FRapport, '  Anomalie   : Le nom du contact chez l''émetteur '+
                      'n''est pas renseigné');
    Writeln(FRapport, '  --> Action : En consultation des Ducs renseigner '+
                      'les coordonnées du déclarant.');
    result := False;
  end;
  { N° Tel ou n° Fax chez l'émetteur (segments COMFRT et COMFRF)}
  if (EnregDucsEdi.TelephoneDecl = '') and
     (EnregDucsEdi.FaxDeclarant = '') then
  begin
    Writeln(FRapport, '  Anomalie   : Le n° de téléphone ou de fax du contact '+
                      'chez l''émetteur n''est pas renseigné');
    Writeln(FRapport, '  --> Action : En consultation des Ducs renseigner '+
                      'les coordonnées du déclarant.');
    result := False;
  end;
  { Identification émetteur chez le destinataire (segment RFF IRC uniquement)}
  if (Copy(EnregDucsEdi.TypDucs,1,1)= 'I') then
    if (EnregDucsEdi.NoContEmet = '') then
    begin
      Writeln(FRapport, '  Anomalie   : La référence de l''émetteur chez le '+
                        'destinataire n''est pas renseignée');
      Writeln(FRapport, '  --> Action : Modification de la fiche organisme '+
                        'renseigner le n° contrat émetteur');
      result := False;
    end;

  {Infos DESTINATAIRE
   ------------------
   Adresse Oragnisme (destinataire) (segment NADMR)}
  if ((EnregDucsEdi.Adresse1Pog = '') and
      (EnregDucsEdi.Adresse2Pog = '')) or
     (EnregDucsEdi.VillePog = '') or (EnregDucsEdi.CpPog = '') then
  begin
    Writeln(FRapport, '  Anomalie   : L''adresse du destinataire est'+
                      ' incomplète');
    Writeln(FRapport, '  --> Action : Modification de la fiche organisme,'+
                      ' renseigner Ville + Code postal + au moins une ligne '+
                      'd''adresse.' );
    result := False;
  end;
  {Référence destinataire (n° institution) (segment RFFMR - IRC uniquement)}
  if (copy(EnregDucsEdi.TypDucs,1,1) = 'I') then
    if (EnregDucsEdi.InstitutionPog = ' ') then
    begin
      Writeln(FRapport, '  Anomalie   : La référence destinataire (code '+
                        'institution)  n''est pas renseignée');
      Writeln(FRapport, '  --> Action : Modification de la fiche organisme,'+
                        ' renseigner le code institution.' );
      result := False;
    end;

  {Infos ADHERENT
   --------------

  Siret adhérent (segment NADDT)}
  if (EnregDucsEdi.Siret = '') or
     ((Length(EnregDucsEdi.Siret) <> 14) and (Length(EnregDucsEdi.Siret) <> 9)) or
     (ControlSiret(EnregDucsEdi.Siret) <> True) then
  begin
    Writeln(FRapport, '  Anomalie   : Le siret de l''adhérent est incomplet '+
                      'ou erroné ');
    Writeln(FRapport, '  --> Action : En consultation des Ducs renseigner '+
                      'le siret. Penser à rectifier '+
                      'la fiche établissement');
    result := False;
  end;
  {Nom de la personne à contacter chez l'adhérent.
   ----------------------------------------------}
  if (EnregDucsEdi.AdherContact = '') then
  begin
    Writeln(FRapport, '  Anomalie   : Le nom du contact chez l''adhérent n'''+
                      'est pas renseigné');
    Writeln(FRapport, '  --> Action : Modification de la fiche organisme '+
                      'renseigner le Contact adhérent');
    result := False;
  end;

  {Adresse adhérent  (segment NADDT)}
  if ((EnregDucsEdi.Adresse1ET = '') and
      (EnregDucsEdi.Adresse2ET = '')) or
     (EnregDucsEdi.VilleET = '') or (EnregDucsEdi.CpET = '') then
  begin
    Writeln(FRapport, '  Anomalie   : L''adresse de l''adhérent est'+
                      ' incomplète');
    Writeln(FRapport, '  --> Action : Modification de la fiche établissement,'+
                      ' renseigner Ville + Code postal + au moins une ligne '+
                      'd''adresse.' );
    result := False;
  end;
  {N° de téléphone ou de fax du contact chez l'adhérent (segments COMDTT et COMDTF)}
  if (EnregDucsEdi.TelET = '') and (EnregDucsEdi.FaxET = '') then
  begin
    Writeln(FRapport, '  Anomalie   : Le n° de téléphone ou de fax de l'''+
                      'adhérent n''est pas renseigné');
    Writeln(FRapport, '  --> Action : Modification de la fiche établissement,'+
                      ' renseigner le n° de téléphone ou de fax.');
    result := False;
  end;

  { Code Ape adhérent (segment RFFACT)}
  if (EnregDucsEdi.ApePdu = '') then
  begin
    Writeln(FRapport, '  Anomalie   : Le code Ape de l''adhérent est incomplet '+
                      'ou erroné ');
    Writeln(FRapport, '  --> Action : En consultation des Ducs renseigner '+
                      'le code Ape. Penser à '+
                      'rectifier la fiche établissement');
    result := False;
  end;
end;
{ fin ControlVarDucsGrp1}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 13/05/2002
Modifié le ... : 13/05/2002
Description .. : Contrôle des champs obligatoires au paiement
Suite ........ : Inscription des anomalies dans le fichier LOG
Mots clefs ... : PAIE, PGDUCSEDI
*****************************************************************}
Function ControlPaiement (EnregDucsEdi : TDucsEdi) : Boolean;
begin
  result := true;
// d PT25
  if (copy(EnregDucsEdi.TypDucs,1,1) = 'I') and
     ((EnregDucsEdi.ModePaiement  = 'Z10') or
      (EnregDucsEdi.ModePaiement = '30') or
      (EnregDucsEdi.ModePaiement = '31')) and
     (EnregDucsEdi.Reglement = IDate1900) then
    begin
      Writeln(FRapport, '  Anomalie   : Date de règlement, '''+
                        'la date n''est pas renseignée');

      Writeln(FRapport, '  --> Action : En consultation des Ducs renseigner '+
                        'la date de règlement');
      result := False;
    end;

// f PT25
  if (EnregDucsEdi.ModePaiement = 'Z10') and
     ((EnregDucsEdi.IdentOPS = '' ) or
      (EnregDucsEdi.IdentOPS = Null) or
      (EnregDucsEdi.IdentOPS = '        ' )) then
    begin
      Writeln(FRapport, '  Anomalie   : Télérèglement EDI, l'''+
                        'identification OPS n''est pas renseigné');

      Writeln(FRapport, '  --> Action : Modification organisme : renseigner'+
                        ' l''identification OPS ou bien modifier le mode de'+
                        ' règlement ');
      result := False;
    end;


end;
// fin ControlPaiement

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 13/05/2002
Modifié le ... :   /  /    
Description .. : Inscription dans le fichier LOG des erreurs liées
Suite ........ : au contrôle des champs obligatoires des lignes
Suite ........ : de cotisation
Mots clefs ... : PAIE, PGDUCSEDI
*****************************************************************}
Function ControlCotisation(EnregDucsEdi : TDucsEdi;EnrDucsEdiGroupee : TDucsEdiGroupee;  Trait : string) : Boolean;
var
   st : string;
begin
result := True;
st := '  Etablissement concerné : ';
     if (trait <> 'G') then
       st := st + EnregDucsEdi.Etab+' '+EnregDucsEdi.LibelleET
     else
       st := st + EnrDucsEdiGroupee.Etab+' '+EnrDucsEdiGroupee .LibelleET ;
// d PT23-1
// Contrôle : n° interne obligatoire pour ACOSS et UNEDIC
    if ((copy(EnregDucsEdi.TypDucs,1,1)='A') and
        (copy(EnregDucsEdi.Abrege,3,2) <> '00')) or
       (copy(EnregDucsEdi.TypDucs,1,1)<>'A')then
      begin
       if (copy(EnregDucsEdi.TypDucs,1,1)<>'I') then
        begin
         // ACOSS UNEDIC    (n° interne)
         if ((Trait <> 'G') and (EnregDucsEdi.NumeroInterne = '')) or
         ((Trait = 'G') and (EnrDucsEdiGroupee.NumeroInterne = '')) then
           begin
             Writeln(FRapport, '  Anomalie   : Le Numéro interne n''est pas renseigné ');
             Writeln(FRapport,st);
             Writeln(FRapport, '  --> Action : En consultation des Ducs renseigner '+
                               'le numéro interne. Numéro à renseigner sur la fiche organisme');
             result := False;
           end;
        end ;
      end;
// f PT23-1
// d PT23-2
// contrôle : Régularisations uniquement sur TR annuel (ACOSS)
    if ((copy(EnregDucsEdi.TypDucs,1,1)='A') and
        (copy(EnregDucsEdi.Abrege,3,2) <> '00')) then
      if ((Trait <> 'G') and (EnregDucsEdi.Regularisation <> 0.0)) or
         ((Trait = 'G') and  (EnrDucsEdiGroupee.Regularisation <> 0.0)) then
        begin
          Writeln(FRapport, '  Anomalie   : Il ne s''agit pas d''un TR Annuel, '+
                            'pas de régularisation possible.');
          Writeln(FRapport,st);
          Writeln(FRapport, '  --> Action : En consultation des Ducs renseigner '+
                            'le montant de la régularisation à 0');
          result := False;
        end;
// f PT23-2
     if ((EnregDucsEdi.ErrTransport = True) and (Trait <> 'G')) or
        ((EnrDucsEdiGroupee.ErrTransport = True) and (Trait = 'G')) then
       begin
         Writeln(FRapport, '  Anomalie   : Le code Communauté Urbaine n''est pas '+
                           'renseigné sur une ligne détail concernant le transport');
         Writeln(FRapport,st);
         Writeln(FRapport, '  --> Action : En consultation des Ducs renseigner '+
                           'le code Communauté Urbaine des lignes de transport');
         result := False;
       end;
     if ((EnregDucsEdi.ErrCotisqual = True) and (Trait <> 'G')) or
        ((EnrDucsEdiGroupee.ErrCotisqual = True) and (Trait = 'G'))then
       begin
         Writeln(FRapport, '  Anomalie   : Il manque un qualifiant de cotisation.');
         Writeln(FRapport,st);
         Writeln(FRapport, '  --> Action : Revenir sur la fiche codifications'+
                           ' des Ducs au menu paramètres et renseigner le champ.');
         result := False;
       end;
     if ((EnregDucsEdi.ErrInstitution = True) and (Trait <> 'G')) or
        ((EnrDucsEdiGroupee.ErrInstitution = True) and (Trait = 'G'))then
       begin
         Writeln(FRapport, '  Anomalie   : Il manque un code institution sur '+
                           'une ligne détail.');
         Writeln(FRapport,st);
         Writeln(FRapport, '  --> Action : En consultation des Ducs renseigner '+
                           'le code intitution sur les lignes détail.');
         result := False;
       end;
     if ((EnregDucsEdi.ErrCondition = True) and (Trait <> 'G')) or
        ((EnrDucsEdiGroupee.ErrCondition = True) and (Trait = 'G')) then
       begin
         Writeln(FRapport, '  Anomalie   : Il manque une condition spéciale'+
                           ' de cotisations sur une ligne détail.');
         Writeln(FRapport,st);
         Writeln(FRapport, '  --> Action : En consultation des Ducs renseigner '+
                           'la condition spéciale de cotisation sur les lignes '+
                           'détail');
         result := False;
       end;
end;
// fin  ControlCotisation

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 12/04/2002
Modifié le ... :   /  /
Description .. : Ouverture (création) du fichier Ducs Edi dossier
Mots clefs ... : PAIE, PGDUCSEDI
*****************************************************************}
procedure OuvreFichierDucsEdi(var EnregDucsEdi : TDucsEdi);
var
   Nomcomplet : string;
begin
// d PT21-2
{$IFDEF EAGLCLIENT}
     Nomcomplet := Vh_Paie.PgCheminEagl+'\'+NomFicDucs;
{$ELSE}
     Nomcomplet := V_PGI.DatPath+'\'+NomFicDucs;
{$ENDIF}
// f PT21-2     
     if FileExists(Nomcomplet) then
      DeleteFile(PChar(Nomcomplet));

     AssignFile(FDUCSEDI, Nomcomplet);
{$i-}ReWrite(FDUCSEDI); {$i+}
     If IoResult<>0 Then
      Begin
       Writeln(FRapport,  'Fichier inaccessible : '+Nomcomplet);
       Exit ;
      End;
end;
// fin OuvreFichierDucsEdi

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 12/04/2002
Modifié le ... :   /  /    
Description .. : Construction des segments d'en-tête de déclaration
Mots clefs ... : PAIE, PGDUCSEDI
*****************************************************************}
procedure EnTeteDeclaration(var EnregDucsEdi : TDucsEdi);
var
   tampon, blanc : string;
begin
     blanc := StringOfChar(' ',35);
     // TYP - type de destinataire
     //---------------------------
     Writeln(FDUCSEDI,'TYP   '+
                      Copy(EnregDucsEdi.TypDucs,3,6));

     // UNB - En-tête d'interchange : Identification émetteur et destinataire
     //----------------------------------------------------------------------
     tampon := '';
     if (EnregDucsEdi.IdentQual = '5') then
          // identifiant Siret
          tampon := EnregDucsEdi.IdentQual+'   ';
     if (EnregDucsEdi.IdentQual = '22') then
          // identifiant Siren
          tampon := EnregDucsEdi.IdentQual+'  ';
     if (EnregDucsEdi.IdentQual = 'ZZZ') then
          // identifiant selon accord
          tampon := EnregDucsEdi.IdentQual+' ';
     Writeln(FDUCSEDI,'UNB   '+
                      tampon +
                      EnregDucsEdi.IdentEmet+
                      copy(blanc,1,(35-length(EnregDucsEdi.IdentEmet)))+
                      EnregDucsEdi.identDest+
                      copy(blanc,1,(35-length(EnregDucsEdi.identDest)))+
                      Copy(DateTimeToStr(Now),9,2)+
                      Copy(DateTimeToStr(Now),4,2)+
                      Copy(DateTimeToStr(Now),1,2)+
                      Copy(DateTimeToStr(Now),12,2)+
                      Copy(DateTimeToStr(Now),15,2));

     // BGM - Début de message
     //-----------------------
     //       Référence interne de la déclaration chez l'émetteur (nom fichier)
     if (copy(EnregDucsEdi.TypDucs,1,1) = 'I') then // PT12
        { Caisse IRC }
        Writeln(FDUCSEDI,'BGM   '+
                         Copy(NomFicDucs ,1,25)+
                         copy(blanc,1,(35-25)))
     else
        {Caisse non IRC }
        Writeln(FDUCSEDI,'BGM   '+
                         Copy(NomFicDucs ,1,21)+
                         copy(blanc,1,(35-21)));

     // GIS - Type de déclaration(périodicité)
     //---------------------------------------
     tampon := '';
     if (Copy(EnregDucsEdi.TypDucs,1,1) = 'A') then
        // ACOSS
        begin
          if  (Copy(EnregDucsEdi.Abrege,3,1) <> '0') then
           // Mensuel ou Trimestriel
           begin
            if (EnregDucsEdi.DucsDossier <> True ) then
              // Un seul établissement
              tampon := '913'
            else
              // Plusieurs établissements
              tampon := '914';
           end
         else
           // Annuel
           begin
            if (EnregDucsEdi.DucsDossier <> True) then
              // Un seul établissement
              tampon := '915'
            else
              // Plusieurs établissements
              tampon := '916';
           end;
        end; // fin ACOSS

     if (Copy(EnregDucsEdi.TypDucs,1,1) = 'U') then
        // UNEDIC
        begin
         if (Copy(EnregDucsEdi.Abrege,4,1) <> '0') or
            (Copy(EnregDucsEdi.Abrege,3,1) <> '0') then
           // Mensuel ou Trimestriel
           begin
            if (EnregDucsEdi.DucsDossier <> True) then
              // Un seul établissement
              tampon := '920'
            else
              // Plusieurs établissements
              tampon := '921';
           end
         else
           // Annuel
           begin
            if (EnregDucsEdi.DucsDossier <> True) then
              // Un seul établissement
              tampon := '922'
            else
              // Plusieurs établissements
              tampon := '923';
           end;
        end; // fin UNEDIC

     if (Copy(EnregDucsEdi.TypDucs,1,1) = 'I') then
        // IRC
        begin
         if (Copy(EnregDucsEdi.Abrege,4,1) <> '0') then
           // Mensuel
// PT15-1           tampon := '930';
           tampon := '931';
         if (Copy(EnregDucsEdi.Abrege,4,1) = '0') and
            (Copy(EnregDucsEdi.Abrege,3,1) <> '0')then
           // Trimestriel
// PT15-1           tampon := '931';
           tampon := '930';
         if (Copy(EnregDucsEdi.Abrege,4,1) = '0') and
            (Copy(EnregDucsEdi.Abrege,3,1) = '0')then
           // Annuel
           tampon := '932';
        end; // fin IRC
     Writeln(FDUCSEDI,'GIS   '+
                      tampon);
end;
// fin  EnTeteDeclaration

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 15/04/2002
Modifié le ... :   /  /    
Description .. : Construction des segments du groupe 1 de la déclaration
Mots clefs ... : PAIE, PGDUCSEDI
*****************************************************************}
procedure Groupe1Declaration(var EnregDucsEdi : TDucsEdi);
var
   tampon, blanc, periode : string;
begin
     blanc := StringOfChar(' ',35);
     // EMETTEUR DU MESSAGE
     //--------------------
     //
     // NADFR -  Nom & adresse de l'émetteur
     // ------------------------------------
     Writeln(FDUCSEDI,'NADFR '+
                      EnregDucsEdi.SiretPet+
                      copy(blanc,1,(35-length(EnregDucsEdi.SiretPet)))+
                      EnregDucsEdi.Declarant+
                      copy(blanc,1,(35-length(EnregDucsEdi.Declarant)))+
                      EnregDucsEdi.AdressePet+
                      copy(blanc,1,(35-length(EnregDucsEdi.AdressePet)))+
                      EnregDucsEdi.Adresse2Pet+
                      copy(blanc,1,(35-length(EnregDucsEdi.Adresse2Pet)))+
                      EnregDucsEdi.VillePet+
                      copy(blanc,1,(35-length(EnregDucsEdi.VillePet)))+
                      EnregDucsEdi.CpPet);

      // CTAFR - Info sur le correspondant émetteur, personne à contacter
      // ----------------------------------------------------------------
     Writeln(FDUCSEDI,'CTAFR '+
                      EnregDucsEdi.DeclarantSuite+
                      copy(blanc,1,(35-length(EnregDucsEdi.DeclarantSuite))));

      // COMFRT - Mode de communication avec le contact (téléphone)
      // ----------------------------------------------------------
     if (EnregDucsEdi.TelephoneDecl <> '') then
        Writeln(FDUCSEDI,'COMFRT'+
                         EnregDucsEdi.TelephoneDecl+
                         copy(blanc,1,(25-length(EnregDucsEdi.TelephoneDecl)))+
                         'TE');
      // COMFRF - Mode de communication avec le contact (fax)
      // ----------------------------------------------------------
      if (EnregDucsEdi.FaxDeclarant <> '') then
         Writeln(FDUCSEDI,'COMFRF'+
                          EnregDucsEdi.FaxDeclarant+
                          copy(blanc,1,(25-length(EnregDucsEdi.FaxDeclarant)))+
                          'FX');
      // RFFFR - Référence de l'émetteur chez le destinataire (IRC uniquement)
      // ---------------------------------------------------------------------
      if (copy(EnregDucsEdi.TypDucs,1,1) = 'I') then
         Writeln(FDUCSEDI,'RFFFR '+
                          EnregDucsEdi.NoContEmet+
                          copy(blanc,1,(35-length(EnregDucsEdi.NoContEmet)))+
                          'FX');

     // DESTINATAIRE DU MESSAGE
     //--------------------
     //
     // NADMR -  Nom & adresse du destinataire
     // ------------------------------------
     Writeln(FDUCSEDI,'NADMR '+
                      EnregDucsEdi.SiretPog+
                      copy(blanc,1,(35-length(EnregDucsEdi.SiretPog)))+
                      EnregDucsEdi.LibellePog+
                      copy(blanc,1,(35-length(EnregDucsEdi.LibellePog)))+
                      EnregDucsEdi.Adresse1Pog+
                      copy(blanc,1,(35-length(EnregDucsEdi.Adresse1Pog)))+
                      EnregDucsEdi.Adresse2Pog+
                      copy(blanc,1,(35-length(EnregDucsEdi.Adresse2Pog)))+
                      EnregDucsEdi.VillePog+
                      copy(blanc,1,(35-length(EnregDucsEdi.VillePog)))+
                      EnregDucsEdi.CpPog);
     // CTAMR - Info sur le correspondant detinatair (IRC uniquement)
     // -------------------------------------------------------------
     if (copy(EnregDucsEdi.TypDucs,1,1) = 'I') then
       Writeln(FDUCSEDI,'CTAMR '+
                        EnregDucsEdi.ContactPog+
                        copy(blanc,1,(35-length(EnregDucsEdi.ContactPog))));
     // COMMRT - Mode de communication avec le contact destinataire (Tel)
     //          (IRC uniquement)
     // -----------------------------------------------------------------
     if (copy(EnregDucsEdi.TypDucs,1,1) = 'I') and
        (EnregDucsEdi.TelPog <> '') and
        (EnregDucsEdi.TelPog <> '') then
       Writeln(FDUCSEDI,'COMMRT'+
                        EnregDucsEdi.TelPog+
                        copy(blanc,1,(25-length(EnregDucsEdi.TelPog)))+
                        'TE');
     // COMMRF - Mode de communication avec le contact destinataire (Fax)
     //          (IRC uniquement)
     // -----------------------------------------------------------------
     if (copy(EnregDucsEdi.TypDucs,1,1) = 'I') and
        (EnregDucsEdi.FaxPog <> '') and
        (EnregDucsEdi.FaxPog <> '') then
       Writeln(FDUCSEDI,'COMMRF'+
                        EnregDucsEdi.FaxPog+
                        copy(blanc,1,(25-length(EnregDucsEdi.FaxPog)))+
                        'FX');
     // RFFMR - n° gagrément ministériel du groupe d'institution ou
     //         n° d'institution (IRC uniquement)
     // -----------------------------------------------------------------
     if (copy(EnregDucsEdi.TypDucs,1,1) = 'I') then
// PT15-2       Writeln(FDUCSEDI,'CRFFMR '+
       Writeln(FDUCSEDI,'RFFMR '+
                        EnregDucsEdi.InstitutionPog+
                        copy(blanc,1,(35-length(EnregDucsEdi.InstitutionPog))));


    // ADHERENT (ENTREPRISE DECLARANTE)
    // --------------------------------
    //
    // NADDT - Nom & adresse de l'adhérent
    // -----------------------------------
    Writeln(FDUCSEDI,'NADDT '+
                     EnregDucsEdi.Siret+
                     Copy(blanc,1,(35-length(EnregDucsEdi.Siret)))+
                     EnregDucsEdi.LibelleET+
                     Copy(blanc,1,(35-length(EnregDucsEdi.LibelleET)))+
                     EnregDucsEdi.Adresse1ET+
                     Copy(blanc,1,(35-length(EnregDucsEdi.Adresse1ET)))+
                     EnregDucsEdi.Adresse2ET+
                     Copy(blanc,1,(35-length(EnregDucsEdi.Adresse2ET)))+
                     EnregDucsEdi.VilleET+
                     Copy(blanc,1,(35-length(EnregDucsEdi.VilleET)))+
                     EnregDucsEdi.CpET);
    // CTADT - Nom du contact chez l'adhérent
    // --------------------------------------
    Writeln(FDUCSEDI,'CTADT '+
                     EnregDucsEdi.AdherContact+
                     Copy(blanc,1,(35-length(EnregDucsEdi.AdherContact))));
    // COMDTT - Mode de communication avec le contact chez l'adhérent (Tel)
    // --------------------------------------------------------------------
    if (EnregDucsEdi.TelET <> '') then
      Writeln(FDUCSEDI,'COMDTT'+
                       EnregDucsEdi.TelET+
                       Copy(blanc,1,(25-length(EnregDucsEdi.TelET)))+
                       'TE');
    // COMDTF - Mode de communication avec le contact chez l'adhérent (Fax)
    // --------------------------------------------------------------------
    if (EnregDucsEdi.FaxET <> '') then
      Writeln(FDUCSEDI,'COMDTF'+
                       EnregDucsEdi.FaxET+
                       Copy(blanc,1,(25-length(EnregDucsEdi.FaxET)))+
                     'FX');
// d PT15-3
   // RFFNUM - Identification adhérent chez l'institution (n° interne)
   // ----------------------------------------------------------------
   if (copy(EnregDucsEdi.TypDucs,1,1) = 'I') then
   Writeln(FDUCSEDI,'RFFNUM'+
                    EnregDucsEdi.NumeroInterne+
                    Copy(blanc,1,(35-length(EnregDucsEdi.NumeroInterne))));
// f PT15-3
   // RFFACT - Référence de l'activité pricipale de l'adhérent
   // --------------------------------------------------------
//PT15-4   if (copy(EnregDucsEdi.TypDucs,1,1) <> 'A') then  // PT5
   if (copy(EnregDucsEdi.TypDucs,1,1) = 'I') then
   Writeln(FDUCSEDI,'RFFACT'+
                    EnregDucsEdi.ApePdu+
//PT15-4                    Copy(blanc,1,(35-length(EnregDucsEdi.FaxET))));
                    Copy(blanc,1,(35-length(EnregDucsEdi.ApePdu))));

// d PT5
   // RFFAMG - Acoss, Entreprise adhérent TDS (O/N)
   // --------------------------------------------------------
   if (copy(EnregDucsEdi.TypDucs,1,1) = 'A') and
      (copy(EnregDucsEdi.Abrege,3,2) = '00') then
     // TR Acoss
     begin
      if (GetParamSoc('SO_PGADHTDS') = True)  then
        // Adhérent TDS
        Writeln(FDUCSEDI,'RFFAMG1'+Copy(blanc,1,34))
      else
        // Non adhérent TDS
        Writeln(FDUCSEDI,'RFFAMG0'+Copy(blanc,1,34));
     end;
// f PT5

   // TIERS REDACTEUR
   // ---------------
   //  Ces segments sont alimentés par les  mêmes informations que les segments
   //  concernant l'émetteur du message. il ne seront générés que si le siret
   //  émetteur est différent du siret adhérent. (cela suppose que le rédacteur
   //  sera toujours l'émetteur)
//PT15-5   if (EnregDucsEdi.SiretPet = EnregDucsEdi.Siret) then
   if (EnregDucsEdi.SiretPet <> EnregDucsEdi.Siret) then
   begin
     // NADAE  Nom & adresse du tiers rédacteur
     // ---------------------------------------
     Writeln(FDUCSEDI,'NADAE '+
                      EnregDucsEdi.SiretPet+
                      copy(blanc,1,(35-length(EnregDucsEdi.SiretPet)))+
                      EnregDucsEdi.Declarant+
                      copy(blanc,1,(35-length(EnregDucsEdi.Declarant)))+
                      EnregDucsEdi.AdressePet+
                      copy(blanc,1,(35-length(EnregDucsEdi.AdressePet)))+
                      EnregDucsEdi.Adresse2Pet+
                      copy(blanc,1,(35-length(EnregDucsEdi.Adresse2Pet)))+
                      EnregDucsEdi.VillePet+
                      copy(blanc,1,(35-length(EnregDucsEdi.VillePet)))+
                      EnregDucsEdi.CpPet);

      // CTAAE - Info sur le correspondant tiers rédacteur, personne à contacter
      // -----------------------------------------------------------------------
     Writeln(FDUCSEDI,'CTAAE '+
                      EnregDucsEdi.DeclarantSuite+
                      copy(blanc,1,(35-length(EnregDucsEdi.DeclarantSuite))));

      // COMAET - Mode de communication avec le contact (téléphone)
      // ----------------------------------------------------------
     if (EnregDucsEdi.TelephoneDecl <> '') then
        Writeln(FDUCSEDI,'COMAE '+
                         EnregDucsEdi.TelephoneDecl+
                         copy(blanc,1,(25-length(EnregDucsEdi.TelephoneDecl)))+
                         'TE');
      // COMAEF - Mode de communication avec le contact (fax)
      // ----------------------------------------------------------
      if (EnregDucsEdi.FaxDeclarant <> '') then
         Writeln(FDUCSEDI,'COMAE '+
                          EnregDucsEdi.FaxDeclarant+
                          copy(blanc,1,(25-length(EnregDucsEdi.FaxDeclarant)))+
                          'FX');
      // RFFAE - Référence de l'émetteur chez le destinataire (IRC uniquement)
      // ---------------------------------------------------------------------
      if (copy(EnregDucsEdi.TypDucs,1,1) = 'I') then
         Writeln(FDUCSEDI,'RFFAE '+
                          EnregDucsEdi.NoContEmet+
                          copy(blanc,1,(35-length(EnregDucsEdi.NoContEmet)))+
                          'FX');
   end;

   // PAIEMENT GROUPE  (UNEDIC uniquement)
   // ------------------------------------
   // La présence des segments "paiement groupé" dépend de la valeur du champ
   // POG_PAIEGROUPE (=TRUE).
   // Les segments "paiement groupé" concernent l'établissement principal
   // responsable du paiement groupé. Correspond à une ducs dossier
   if (copy(EnregDucsEdi.TypDucs,1,1) = 'U') and
      (EnregDucsEdi.PaiementGroupe = true) then
   begin
     // NADFZ - Nom & adresses établissemnt principal du paiement groupé
     // ----------------------------------------------------------------
     Writeln(FDUCSEDI,'NADFZ '+
                          EnregDucsEdi.Siret+
                          copy(blanc,1,(35-length(EnregDucsEdi.Siret)))+
                          EnregDucsEdi.LibelleET+
                          copy (blanc,1,(35-length(EnregDucsEdi.LibelleET)))+
                          EnregDucsEdi.Adresse1ET+
                          copy (blanc,1,(35-length(EnregDucsEdi.Adresse1ET)))+
                          EnregDucsEdi.Adresse2ET+
                          copy (blanc,1,(35-length(EnregDucsEdi.Adresse2ET)))+
                          EnregDucsEdi.VilleET+
                          copy (blanc,1,(35-length(EnregDucsEdi.VilleET)))+
                          EnregDucsEdi.CpET);
     // CTAFZ - Info sur le correspondant de paiement groupé, personne à contacter
     // --------------------------------------------------------------------------
     Writeln(FDUCSEDI,'CTAFZ '+
                      EnregDucsEdi.AdherContact+
                      copy(blanc,1,(35-length(EnregDucsEdi.AdherContact))));
     // COMFZT - Mode de communication avec la personne du paiement groupé à contacter (Tel)
     // ------------------------------------------------------------------------------------
     if (EnregDucsEdi.TelET <> '') then
       Writeln(FDUCSEDI,'COMFZT'+
                        EnregDucsEDi.TelET+
                        copy(blanc,1,(25-length(EnregDucsEdi.TelET)))+
                        'TE');
     // COMFZF - Mode de communication avec la personne du paiement groupé à contacter (Fax)
     // ------------------------------------------------------------------------------------
      if (EnregDucsEdi.FaxET <> '') then
        Writeln(FDUCSEDI,'COMFZF'+
                        EnregDucsEDi.FaxET+
                        copy(blanc,1,(25-length(EnregDucsEdi.FaxET)))+
                        'FX');
   end;
   // PEDp00 - Période de la déclaration
   // ----------------------------------
   tampon :='';
   if (copy(EnregDucsEdi.Abrege,4,1) <> '0') then
     // mensuelle
     begin
     tampon := 'M';
     EnregDucsEdi.Periode := 'M';
     end
   else
     if (copy(EnregDucsEdi.Abrege,3,1) <> '0') then
       // Trimestrielle
       begin
       tampon := 'T';
       EnregDucsEdi.Periode := 'T';
       end
     else
       begin
       tampon :='A';
       EnregDucsEdi.Periode := 'A';
       end;
   if (Copy(EnregDucsEdi.Typducs,1,1) = 'A') then
     // ACOSS
     begin
      if (tampon <> 'A') then
        // Mensuelle ou trimestrielle   (PEDp00+abrégé période)
        Writeln(FDUCSEDI,'PED'+
                         tampon+
                         '00'+
                         EnregDucsEdi.Abrege+
                         copy(blanc,1,(17-(length(EnregDucsEdi.Abrege)))))
      else
        // Annuelle   (PEDp00+année+62)
        Writeln(FDUCSEDI,'PED'+
                         'A'+
                         '00'+
                         copy(EnregDucsEdi.Abrege,1,2)+
                         '62'+
                         copy(blanc,1,(17-(length(EnregDucsEdi.Abrege)))));
     end
   else
   // IRC ou UNEDIC
   begin
    periode := '';
    periode :=  copy(DateToStr(EnregDucsEdi.DateFin),7,2);
    PeriodRef := EnregDucsEdi.Abrege;     //  PT35
    if (tampon =  'A') then
        // Annuelle  (PEDp00+ssaa)
        periode := periode+
                   copy(EnregDucsEdi.Abrege,1,2);
    if (tampon = 'T') then
        // Trimestrielle  (PEDp00+ssaat)
        periode := periode+
// correction suite test validation    copy(EnregDucsEdi.Abrege,3,1);
                   copy(EnregDucsEdi.Abrege,1,3);
    if (tampon = 'M') then
        // Mensuelle   (PEDp00+ssaamm)
// d PT27
    if (((VH_Paie.PGDecalage = true) or (VH_Paie.PGDecalagePetit = true)) and
        (Copy(EnregDucsEdi.Typducs,1,1) = 'U')) then
    begin
       if (copy(EnregDucsEdi.Abrege,3,2) = '11') then
       begin
         periode := periode+
                   format('%2.2d',[StrToInt(copy(EnregDucsEdi.Abrege,1,2))-1])+'12';
         PeriodRef := format('%2.2d',[StrToInt(copy(EnregDucsEdi.Abrege,1,2))-1])+'43';
       end
       else
       begin
         periode := periode+
                   copy(EnregDucsEdi.Abrege,1,2)+
                   format('%2.2d',[((StrToint(copy(EnregDucsEdi.Abrege,3,1))-1)*3)+
                                  StrToInt(copy(EnregDucsEdi.Abrege,4,1))-1]);

       PeriodRef :=ColleZeroDevant(StrToInt(EnregDucsEdi.Abrege)-1,4);
       if (copy(PeriodRef,4,1) = '0') then
          PeriodRef :=ColleZeroDevant(StrToInt(EnregDucsEdi.Abrege)-8,4);

       end;
    end
    else
// f PT27
        periode := periode+
                   copy(EnregDucsEdi.Abrege,1,2)+ // correction suite test validation
// PT9-2                  IntToStr((StrToint(copy(EnregDucsEdi.Abrege,3,1))-1)*3+
// PT9-2                           StrToInt(copy(EnregDucsEdi.Abrege,4,1)));
                   format('%2.2d',[((StrToint(copy(EnregDucsEdi.Abrege,3,1))-1)*3)+ //PT9-2
                                  StrToInt(copy(EnregDucsEdi.Abrege,4,1))]);        // PT9-2
//PT35       PeriodRef := EnregDucsEdi.Abrege;     //  PT29
        Writeln(FDUCSEDI,'PED'+
                         tampon+
                         '00'+
                         periode+
                         copy(blanc,1,(17-(length(periode)))));
   end;
   // PED 01 - Période de versement des salares pour paie décalée (IRC uniquement)
   // ----------------------------------------------------------------------------
   if ((VH_Paie.PGDecalage = true) or (VH_Paie.PGDecalagePetit = true)) and
      (copy(EnregDucsEdi.TypDucs,1,1) = 'I') then
        Writeln(FDUCSEDI,'PED '+
                         '01'+
                         copy(DateToStr(EnregDucsEdi.DateDebut),7,4)+
                         copy(DateToStr(EnregDucsEDi.DateDebut),1,2)+
                         copy(DateToStr(EnregDucsEDi.DateDebut),4,2)+
                         '-'+
                         copy(DateToStr(EnregDucsEdi.DateFin),7,4)+
                         copy(DateToStr(EnregDucsEDi.DateFin),1,2)+
                         copy(DateToStr(EnregDucsEDi.DateFin),4,2));
   // QTY - Effectif au 31/12, déclarations annuelles (Acoss et Unedic uniquement)
   // ----------------------------------------------------------------------------
   tampon := '';
   if (copy(EnregDucsEdi.Abrege,3,2) ='00') and
      (copy(EnregDucsEdi.TypDucs,1,1) <> 'I') then
     begin
      if (copy(EnregDucsEdi.TypDucs,1,1) = 'A') then
        // ACOSS
//PT5        tampon := FloatToStr((EnregDucsEdi.TotHommes+EnregDucsEdi.TotFemmes)-EnregDucsEdi.TotApprenti);
         tampon := FloatToStr((EnregDucsEdi.TotHommes+EnregDucsEdi.TotFemmes));
      if (copy(EnregDucsEdi.TypDucs,1,1) = 'U') then
        // UNEDIC
        tampon :=  FloatToStr(EnregDucsEdi.TotHommes+EnregDucsEdi.TotFemmes);

      Writeln(FDUCSEDI,'QTY   '+
                       tampon+
                       copy(blanc,1,(18-(length(tampon)))));
     end;
   // RFFZST - Référence du statut juridique de l'entreprise
   //-------------------------------------------------------
   //          Déclarations annuelles (UNEDIC seulement)
   if (copy(EnregDucsEdi.Abrege,3,2) ='00') and
      (copy(EnregDucsEdi.TypDucs,1,1) = 'U') then
      Writeln(FDUCSEDI,'RFFZST'+
                       EnregDucsEdi.Juridique+
                       copy(blanc,1,(35-(length(EnregDucsEdi.Juridique)))));
end;
// fin Groupe1Declaration

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 17/04/2002
Modifié le ... :   /  /    
Description .. : Construction des segments du groupe 2 de la déclaration
Mots clefs ... : PAIE, PGDUCSEDI
*****************************************************************}
procedure Groupe2Declaration (EnregDucsEdi : TDucsEdi; EnrDucsEdiGroupee : TDucsEdiGroupee; Trait : string);

var
   blanc, tampon : string;
   typenr : string;       // PT1-2
begin
     blanc := StringOfChar(' ',35);
     // NADDC - Etablissement déclaré Nom & adresse
     // -------------------------------------------
     if (Trait <> 'G') then
      Writeln(FDUCSEDI,'NADDC '+
                       EnregDucsEdi.Siret+
                       copy(blanc,1,(35-(length(EnregDucsEdi.Siret))))+
                       EnregDucsEdi.LibelleET+
                       copy(blanc,1,(35-(length(EnregDucsEdi.LibelleET))))+
                       EnregDucsEdi.Adresse1ET+
                       copy(blanc,1,(35-(length(EnregDucsEdi.Adresse1ET))))+
                       EnregDucsEdi.Adresse2ET+
                       copy(blanc,1,(35-(length(EnregDucsEdi.Adresse2ET))))+
                       EnregDucsEdi.VilleET+
                       copy(blanc,1,(35-(length(EnregDucsEdi.VilleET))))+
                       EnregDucsEdi.CpET)
      else
// d PT1-2
       begin
        typenr := 'NADDC ';
        if (copy(EnregDucsEdi.TypDucs,1,1) = 'U') then typenr := 'NADDCE';
//        Writeln(FDUCSEDI,'NADDC '+
        Writeln(FDUCSEDI,typenr+
// f PT1-2
                        EnrDucsEdiGroupee.Siret+
                        copy(blanc,1,(35-(length(EnrDucsEdiGroupee.Siret))))+
                        EnrDucsEdiGroupee.LibelleET+
                        copy(blanc,1,(35-(length(EnrDucsEdiGroupee.LibelleET))))+
                        EnrDucsEdiGroupee.Adresse1ET+
                        copy(blanc,1,(35-(length(EnrDucsEdiGroupee.Adresse1ET))))+
                        EnrDucsEdiGroupee.Adresse2ET+
                        copy(blanc,1,(35-(length(EnrDucsEdiGroupee.Adresse2ET))))+
                        EnrDucsEdiGroupee.VilleET+
                        copy(blanc,1,(35-(length(EnrDucsEdiGroupee.VilleET))))+
                        EnrDucsEdiGroupee.CpET);

       end; // PT1-2

    // CTADC - Nom du contact celui de  l'adhérent
    // -------------------------------------------
    if (Trait <> 'G') then
      Writeln(FDUCSEDI,'CTADC '+
                       EnregDucsEdi.AdherContact+
                       Copy(blanc,1,(35-length(EnregDucsEdi.AdherContact))));

    if (Trait <> 'G') then
      begin
        if (copy(EnregDucsEdi.TypDucs,1,1) <> 'I') then
          begin
          // COMDC - Mode de communication avec le contact (Tel) - IRC seulement
          // -------------------------------------------------------------------
          if (EnregDucsEdi.TelET <> '') then
            Writeln (FDUCSEDI,'COMDCT'+
                              EnregDucsEdi.TelET+
                              copy(blanc,1,(25-length(EnregDucsEdi.TelET)))+
                              'TE');
          // COMDCF - Mode de communication avec le contact (Fax) - IRC seulement
          // --------------------------------------------------------------------
          if (EnregDucsEdi.FaxET <> '') then
            Writeln (FDUCSEDI,'COMDCF'+
                              EnregDucsEdi.FaxET+
                              copy(blanc,1,(25-length(EnregDucsEdi.FaxET)))+
                              'FX');
          end;
      end;

    // PEV   - Date de versement des salaires
    // --------------------------------------
    // Pour l'Acoss cette date, sous la forme YYQM, représente la période
    // de référence si elle est différente de la période de déclaration
    // du groupe 1. (cas de cotisations à effet rétroactif)
    // !! Ce cas n'est pas traité
    if (Trait <> 'G') then
    begin
      if (copy(EnregDucsEdi.TypDucs,1,1) = 'I') and // PT1-3
         ((VH_Paie.PGDecalage = true) or (VH_Paie.PGDecalagePetit = true)) then
        // Pour IRC et s'il y a décalage de paie
        begin
          Writeln(FDUCSEDI,'PEV   '+
                           Copy(DateToStr(EnregDucsEdi.DateDebut),7,10)+
                           Copy(DateToStr(EnregDucsEdi.DateDebut),4,2)+
                           Copy(DateToStr(EnregDucsEdi.DateDebut),1,2)+
                           '-'+
                           Copy(DateToStr(EnregDucsEdi.DateFin),7,10)+
                           Copy(DateToStr(EnregDucsEdi.DateFin),4,2)+
                           Copy(DateToStr(EnregDucsEdi.DateFin),1,2));
        end;
//d PT1-3
//      if (copy(EnregDucsEdi.TypDucs,1,1) <> 'U') then
      if (copy(EnregDucsEdi.TypDucs,1,1) = 'U') then
        if (EnregDucsEdi.Neant = false) then
         Writeln(FDUCSEDI,'PEV   '+
                          Copy(DateToStr(EnregDucsEdi.Paiement),7,10)+
                          Copy(DateToStr(EnregDucsEdi.Paiement),4,2)+
                          Copy(DateToStr(EnregDucsEdi.Paiement),1,2)+
                          Copy(blanc,1,9))
         else
          Writeln(FDUCSEDI,'PEV   '+
                           '00000000',
                           Copy(blanc,1,9)) ;
// f PT1-3
    end;

    // MOADUE - Montant total déclaré
    // ------------------------------
    if (Trait <> 'G') then
    begin
        if (EnregDucsEdi.MtDeclare <> 0.0) or    // PT8-1
           ((EnregDucsEdi.MtDeclare = 0.0) and   // PT8-1
//PT15-6           (copy(EnregDucsEdi.TypDucs,1,1)='U')) then
           (copy(EnregDucsEdi.TypDucs,1,1) <> 'A')) then
         begin
           if (copy(EnregDucsEdi.TypDucs,1,1)<>'I') then
             // Acoss et Unedic mt à l'euro
             begin
               tampon :=FormatFloat('#0',EnregDucsEdi.MtDeclare);
             end
           else
             // IRC mt au centime d'euro
             begin
               tampon := FormatFloat('#0',EnregDucsEdi.MtDeclare*100);
             end;
           Writeln(FDUCSEDI,'MOADUE'+
                            tampon+
                            copy(blanc,1,(18-length(tampon)))+
                            EnregDucsEdi.MonnaieTenue);
         end;
    end
    else
      begin
        if (EnrDucsEdiGroupee.MtDeclare <> 0.0) then // PT8-1
         begin
           if (copy(EnregDucsEdi.TypDucs,1,1)<>'I') then
             // Acoss et Unedic mt à l'euro
             begin
               tampon :=FormatFloat('#0',EnrDucsEdiGroupee.MtDeclare);
             end
           else
             // IRC mt au centime d'euro
             begin
               tampon := FormatFloat('#0',EnrDucsEdiGroupee.MtDeclare*100);
             end;
           Writeln(FDUCSEDI,'MOADUE'+
                            tampon+
                            copy(blanc,1,(18-length(tampon)))+
                            EnrDucsEdiGroupee.MonnaieTenue);
         end;
      end;

    // MOATRA - Montant total des cotisations transport (ACOSS)
    // --------------------------------------------------------
    if (Trait <> 'G') then
     begin
      if (copy(EnregDucsEdi.TypDucs,1,1)='A') and
         (EnregDucsEdi.NbTransport > 1) then
         // Acoss  et plusieurs taux de cotisation transport
         begin
           tampon :=FormatFloat('#0',EnregDucsEdi.MtTransport);
           Writeln(FDUCSEDI,'MOATRA'+
                            tampon+
                            copy(blanc,1,(18-length(tampon)))+
                            EnregDucsEdi.MonnaieTenue);
         end;
      end
    else
     begin
      if (copy(EnregDucsEdi.TypDucs,1,1)='A') and
         (EnrDucsEdiGroupee.NbTransport > 1) then
         // Acoss  et plusieurs taux de cotisation transport
         begin
           tampon :=FormatFloat('#0',EnrDucsEdiGroupee.MtTransport);
           Writeln(FDUCSEDI,'MOATRA'+
                            tampon+
                            copy(blanc,1,(18-length(tampon)))+
                            EnrDucsEdiGroupee.MonnaieTenue);
         end;
     end;

    // MOAREG - Montant de la régularisation (ACOSS)
    // ---------------------------------------------
    if (copy(EnregDucsEdi.TypDucs,1,1)='A') then
      // Acoss
      if (Trait <> 'G') then
       begin
         tampon :=FormatFloat('#0',EnregDucsEdi.Regularisation);
         Writeln(FDUCSEDI,'MOAREG'+
                          tampon+
                          copy(blanc,1,(18-length(tampon)))+
                          EnregDucsEdi.MonnaieTenue);
// d PT23-2
         if (EnregDucsEdi.Regularisation <> 0.0) then
         Writeln(FRapport,'    Montant de la régularisation : '+
                          Formatfloat('#0',EnregDucsEdi.Regularisation)+' '+
                          EnregDucsEdi.MonnaieTenue);
// f PT23-2
       end
      else
       begin
         tampon :=FormatFloat('#0',EnrDucsEdiGroupee.Regularisation);
         Writeln(FDUCSEDI,'MOAREG'+
                          tampon+
                          copy(blanc,1,(18-length(tampon)))+
                          EnrDucsEdiGroupee.MonnaieTenue);
// d PT23-2
         if (EnrDucsEdiGroupee.Regularisation <> 0.0) then
           Writeln(FRapport,'    Montant de la régularisation : '+
                           Formatfloat('#0',EnrDucsEdiGroupee.Regularisation)+' '+
                           EnrDucsEdiGroupee.MonnaieTenue);
// f PT23-2
       end;

    // QTYD  - Effectif des salariés de l'établissement rémunérés pour la période (ACOSS)
    // ----------------------------------------------------------------------------------
    if (copy(EnregDucsEdi.TypDucs,1,1)='A') then
      // Acoss
      begin
        if (EnregDucsEdi.Neant = False ) then
         begin
           if (Trait <>'G') then
             tampon :=FormatFloat('#0',EnregDucsEdi.NbSalFpe)
           else
             tampon :=FormatFloat('#0',EnrDucsEdiGroupee.NbSalFpe);
         end
        else
         // Ducs Néant et Acoss - Effectif à zéro
         tampon :=FormatFloat('#0',0);

        Writeln(FDUCSEDI,'QTYD  '+
                          tampon+
                          copy(blanc,1,(18-length(tampon))));
     end;
// d PT17
{    // QTYT  - Effectif des salariés inscrits dans l'établissement au dernier jour de la période (ACOSS UNEDIC)
    // --------------------------------------------------------------------------------------------------------
// PT9-2    if ((copy(EnregDucsEdi.TypDucs,1,1)='A') or
// PT9-2       (copy(EnregDucsEdi.TypDucs,1,1)='U')) and
// PT9-2       (copy(EnregDucsEdi.Abrege,3,2) = '00') then
     if ((copy(EnregDucsEdi.TypDucs,1,1)='A') or
         (copy(EnregDucsEdi.TypDucs,1,1)='U')) then
       // Acoss ou UNEDIC
      begin
        typenr:='QTYT  '; // PT1-2
        if (EnregDucsEdi.Neant = False ) or (copy(EnregDucsEdi.TypDucs,1,1) <>'A') then
        begin

         if (Trait <> 'G') or (copy(EnregDucsEdi.TypDucs,1,1)='A') then //PT5
           tampon :=FormatFloat('#0',
//PT5                                (EnregDucsEdi.TotHommes+EnregDucsEdi.TotFemmes-EnregDucsEdi.TotApprenti))
                                 (EnregDucsEdi.TotHommes+EnregDucsEdi.TotFemmes))
         else
           tampon :=FormatFloat('#0',
//PT5                                 (EnrDucsEdiGroupee.TotHommes+EnrDucsEdiGroupee.TotFemmes-EnrDucsEdiGroupee.TotApprenti));
                                 (EnrDucsEdiGroupee.TotHommes+EnrDucsEdiGroupee.TotFemmes));
        end
        else
           // Ducs Néant et Acoss - Effectif à zéro
         tampon :=FormatFloat('#0',0);
// d PT1-2
        if (Trait = 'G') and (copy(EnregDucsEdi.TypDucs,1,1) = 'U') then typenr := 'QTYTE ';
//        Writeln(FDUCSEDI,'QTYT  '+
        Writeln(FDUCSEDI,typenr+
// f PT1-2
                        tampon+
                        copy(blanc,1,(18-length(tampon))));

      end;   }
    // QTYT  - Effectif des salariés inscrits dans l'établissement au dernier
    // jour de la période (ACOSS)
    // ----------------------------------------------------------------------
    if (copy(EnregDucsEdi.TypDucs,1,1) = 'A') then
    begin
      typenr:='QTYT  ';
      if (EnregDucsEdi.Neant = False ) then
      begin
        if (Trait <> 'G') then
          tampon :=FormatFloat('#0',
                  (EnregDucsEdi.TotHommes+EnregDucsEdi.TotFemmes))
        else
           tampon :=FormatFloat('#0',
                   (EnrDucsEdiGroupee.TotHommes+EnrDucsEdiGroupee.TotFemmes));
      end
      else
        // Ducs Néant - Effectif à zéro
        tampon :=FormatFloat('#0',0);

      Writeln(FDUCSEDI,typenr+
                       tampon+
                       copy(blanc,1,(18-length(tampon))));
    end;
    // QTYTH  - Effectif des salariés Hommes inscrits dans l'établissement au
    //          dernier jour de la période (UNEDIC)
    // --------------------------------------------------------------------------------------------------------
    if (copy(EnregDucsEdi.TypDucs,1,1)='U') then
    // UNEDIC
    begin
      typenr:='QTYTH ';
      if (Trait <> 'G') then
        tampon :=FormatFloat('#0',EnregDucsEdi.TotHommes)
      else
        tampon :=FormatFloat('#0',EnrDucsEdiGroupee.TotHommes);
      if (Trait = 'G') then typenr := 'QTYTHE';
       Writeln(FDUCSEDI,typenr+
                        tampon+
                        copy(blanc,1,(18-length(tampon))));
    end;

    // QTYTF  - Effectif des salariés Femmes inscrits dans l'établissement au
    //          dernier jour de la période (UNEDIC)
    // --------------------------------------------------------------------------------------------------------
    if (copy(EnregDucsEdi.TypDucs,1,1)='U') then
    // UNEDIC
    begin
      typenr:='QTYTF ';
      if (Trait <> 'G') then
        tampon :=FormatFloat('#0',EnregDucsEdi.TotFemmes)
      else
        tampon :=FormatFloat('#0',EnrDucsEdiGroupee.TotFemmes);
      if (Trait = 'G') then typenr := 'QTYTFE';
        Writeln(FDUCSEDI,typenr+
                         tampon+
                         copy(blanc,1,(18-length(tampon))));
    end;

// f PT17

    // RFFDC - Référence d'identification complémentaire de l'établissement
    // --------------------------------------------------------------------
    if ((copy(EnregDucsEdi.TypDucs,1,1)='A') and
        (copy(EnregDucsEdi.Abrege,3,2) <> '00')) or
       (copy(EnregDucsEdi.TypDucs,1,1)<>'A')then
    begin
      if (copy(EnregDucsEdi.TypDucs,1,1)<>'I') then
      begin
        // ACOSS UNEDIC    (n° interne)
        typenr :='RFFDC '; // PT1-2
        if (Trait <> 'G') then
          Writeln(FDUCSEDI,'RFFDC '+
                           EnregDucsEdi.NumeroInterne+
                           copy(blanc,1,(35-length(EnregDucsEdi.NumeroInterne))))
        else
// d PT1-2
        begin  // PTA
            if (copy(EnregDucsEdi.TypDucs,1,1) = 'U') then Typenr :=  'RFFDCE';
//          Writeln(FDUCSEDI,'RFFDC '+
            Writeln(FDUCSEDI,typenr+
// f PT1-2
                           EnrDucsEdiGroupee.NumeroInterne+
                           copy(blanc,1,(35-length(EnrDucsEdiGroupee.NumeroInterne))));
        end;   // PTA
      end
      else
      // IRC   (Centre payeur)
      begin
//PT15-7         Tampon  := IntToStr (EnregDucsEdi.CentrePayeur);
// d PT20
        if (Trait <> 'G') then
        begin
          if (EnregDucsEdi.CentrePayeur <> '') then
            Tampon  := EnregDucsEdi.CentrePayeur
          else
          begin
            Tampon  := Format('%4.4d',[EnregDucsEdi.NoCentrePayeur]);
            PGZeroAGauche(tampon,5);
          end;
        end
        else
        begin
          if (EnrDucsEdiGroupee.CentrePayeur <> '') then
            Tampon  := EnrDucsEdiGroupee.CentrePayeur
          else
          begin
            Tampon  := Format('%4.4d',[EnregDucsEdi.NoCentrePayeur]);
            PGZeroAGauche(tampon,5);
          end;
        end;
// f PT20

        Writeln(FDUCSEDI,'RFFDC '+
                          tampon+
                          copy(blanc,1,(35-length(tampon))));
      end;
    end;
    // RFFACT - Référence complémentaire de l'établissement, Code APE
    // --------------------------------------------------------------
    if ((copy(EnregDucsEdi.TypDucs,1,1)='A') and
        (copy(EnregDucsEdi.Abrege,3,2) = '00')) or
       (copy(EnregDucsEdi.TypDucs,1,1)='U')then
      begin
        // ACOSS (TR) -  UNEDIC    (Code APE)
        typenr := 'RFFACT';    // PT1-2
        if (Trait <> 'G') then
         Writeln(FDUCSEDI,'RFFACT'+
                          EnregDucsEdi.ApePdu+
                          copy(blanc,1,(35-length(EnregDucsEdi.ApePdu))))
        else
        begin //PT5
// d PT1-2
         if (copy(EnregDucsEdi.TypDucs,1,1) = 'U') then  typenr := 'RFFACE';
//         Writeln(FDUCSEDI,'RFFACT'+
         Writeln(FDUCSEDI,typenr+
// f PT1-2
                          EnrDucsEdiGroupee.ApePdu+
                          copy(blanc,1,(35-length(EnrDucsEdiGroupee.ApePdu))))
        end; //PT5
      end;
end;
// fin Groupe2Declaration

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 13/05/2002
Modifié le ... :   /  /    
Description .. : Traitement des segments du groupe 3 de la déclaration
Suite ........ : Spécificité URSSAF
Mots clefs ... : PAIE, PGDUCSEDI
*****************************************************************}
procedure Groupe3ACOSS(EnregDucsEdi : TDucsEdi;EnrDucsEdiGroupee : TDucsEdiGroupee;Trait : string);
var
   CodeTypeEnCours, TauxEnCours : string;
   TypePersonnel,Taux,MtDeplafonne,MtPlafonne,Effectif : string;
   DateChgtTx : string; //PT5
   MtNonZero : boolean;
   blanc : string;
   tampon : string;
   LesLignes : TOB;
   WMonnaieTenue : string;
   WNbTransport : integer;

begin
  if (Trait <> 'G') then
   begin
    WMonnaieTenue := EnregDucsEdi.MonnaieTenue;
    WNbtransport :=  EnregDucsEdi.NbTransport;
   end
  else
   begin
    WMonnaieTenue := EnrDucsEdiGroupee.MonnaieTenue;
    WNbtransport :=  EnrDucsEdiGroupee.NbTransport;
   end;
  MtNonZero := False;
  blanc := StringOfChar(' ',35);
  CodeTypeEnCours := '';
  TauxEnCours :=  '';
  DateChgtTx :='';   //PT5
 if TOB_Lignes <> NIL then
    begin
     LesLignes := TOB_Lignes.FindFirst ([''],[''],TRUE);
     While LesLignes <> NIL  do
      begin
        if (Leslignes.getValue('PDD_TYPECOTISATION') <> 'I') and
           (Leslignes.getValue('PDD_TYPECOTISATION') <> 'S') then
        // les lignes Intitulé et Sous-Total ne sont pas traitées.
        begin
          tampon := copy(Leslignes.getValue('PDD_CODIFICATION'),4,3) ;
          tampon := FloatToStr(LesLignes.getValue('PDD_TAUXCOTISATION'));
          tampon := Leslignes.getValue('PDD_TYPECOTISATION');
// d PT5
          if (Leslignes.getValue('PDD_DATECHGTTAUX') = IDate1900) then
           DateChgtTx := '0'
          else
           DateChgtTx := Copy(DateToStr(Leslignes.getValue('PDD_DATECHGTTAUX')),7,10)+
                         Copy(DateToStr(Leslignes.getValue('PDD_DATECHGTTAUX')),4,2)+
                         Copy(DateToStr(Leslignes.getValue('PDD_DATECHGTTAUX')),1,2);
// f PT5                         


          if (copy(Leslignes.getValue('PDD_CODIFICATION'),4,3) <> CodeTypeEnCours) or
             ((FloatToStr(LesLignes.getValue('PDD_TAUXCOTISATION')) <> TauxEnCours) and
              (Leslignes.getValue('PDD_TYPECOTISATION') = 'A')) then
           // chgt de codif ou de taux s'il s'agit d'un taux AT
           begin
            if (CodeTypeEnCours <> '') and (MtNonZero = True) then
              begin
               if (taux = '') then taux :='0';
               if (MtDeplafonne = '') then MtDeplafonne := '0';
               if (MtPlafonne = '') then MtPlafonne := '0';
               if (Effectif = '') then Effectif := '0';
// d PT16-1
               if (TypePersonnel <> '900') or (WNbTransport <= 1) then
               begin
// f PT16-1
                 Writeln (FDUCSEDI,'ACO   '+
                                 TypePersonnel+
                                 Taux+
                                 copy(blanc,1,(8-length(Taux)))+
                                 MtDeplafonne+
                                 copy(blanc,1,(18-length(MtDeplafonne)))+
                                 MtPlafonne+
                                 copy(blanc,1,(18-length(MtPlafonne)))+
                                 Effectif+
                                 copy(blanc,1,(15-length(Effectif)))+
// d PT5
                                 DateChgtTx+
                                 copy(blanc,1,(8-length(DateChgtTx)))+
// f PT5
                                 WMonnaieTenue);
               end; // PT16-1

               EnregDucsEdi.NbCot := EnregDucsEdi.NbCot+1;
              end;
             TypePersonnel := '';
             Taux := '';
             MtDeplafonne := '';
             MtPlafonne := '';
             Effectif := '';
             MtNonZero := False;
             CodeTypeEnCours :=copy(Leslignes.getValue('PDD_CODIFICATION'),4,3);
             TauxEnCours := '';
           end;
          TypePersonnel := copy(Leslignes.getValue('PDD_CODIFICATION'),4,3);
          if ((copy(Leslignes.getValue('PDD_CODIFICATION'),4,4) = '900P') and
              (WNbTransport <=1)) or
             (Leslignes.getValue('PDD_TYPECOTISATION') = 'A') then
           // Transport ou taux AT
           begin
            taux := FormatFloat('#0',(Leslignes.getValue('PDD_TAUXCOTISATION')*10000));
            TauxEnCours := Taux;
            MtNonZero := True;
           end;

          if (copy(Leslignes.getValue('PDD_CODIFICATION'),7,1) = 'D') or
             ((copy(Leslignes.getValue('PDD_CODIFICATION'),7,1) = ' ') and
             ((Leslignes.getValue('PDD_TYPECOTISATION') <> 'M') and
             (Leslignes.getValue('PDD_TYPECOTISATION') <> 'Q'))) then
           // Déplafonné (codif en D, ou blanc mais pas forfait)
           begin
            MtDeplafonne :=  FormatFloat('#0',(Leslignes.getValue('PDD_BASECOTISATION')));
            MtNonZero := True;
           end;

          if (copy(Leslignes.getValue('PDD_CODIFICATION'),7,1) = 'P') then
           // Plafonne ( codif en P, ou blanc et forfait)
           begin
            MtPlafonne := FormatFloat('#0',(Leslignes.getValue('PDD_BASECOTISATION')));
            MtNonZero := True;
           end;

          if ((Leslignes.getValue('PDD_TYPECOTISATION') = 'M') or
              (Leslignes.getValue('PDD_TYPECOTISATION') = 'Q')) then
           begin
            MtNonZero := True;

            if (Leslignes.getValue('PDD_MTCOTISAT') < 0) then
                // si le forfait est négatif il est renseigné en valeur absolue
                MtPlafonne := FormatFloat('#0',(Leslignes.getValue('PDD_MTCOTISAT')*-1))
            else
                MtPlafonne := FormatFloat('#0',Leslignes.getValue('PDD_MTCOTISAT'));
           end;

          Effectif :=  Leslignes.getValue('PDD_EFFECTIF');
        end;
        LesLignes := TOB_Lignes.FindNext ([''],[''],TRUE);
      end;
     if (TypePersonnel <> '900') or (WNbTransport <= 1) then
     begin
      if (CodeTypeEnCours <> '') and (MtNonZero = True) then
       begin
         if (taux = '') then taux :='0';
         if (MtDeplafonne = '') then MtDeplafonne := '0';
         if (MtPlafonne = '') then MtPlafonne := '0';
         if (Effectif = '') then Effectif := '0';

         Writeln (FDUCSEDI,'ACO   '+
                           TypePersonnel+
                           Taux+
                           copy(blanc,1,(8-length(Taux)))+
                           MtDeplafonne+
                           copy(blanc,1,(18-length(MtDeplafonne)))+
                           MtPlafonne+
                           copy(blanc,1,(18-length(MtPlafonne)))+
                           Effectif+
                           copy(blanc,1,(15-length(Effectif)))+
                           WMonnaieTenue);

         EnregDucsEdi.NbCot := EnregDucsEdi.NbCot+1;
      end;
     end;
//     TOB_Lignes.Free; TOB_Lignes := NIL;
    end;
end;
// fin Groupe3ACOSS
{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 13/05/2002
Modifié le ... :   /  /    
Description .. : Traitement des segments des groupes 4 et 5 de la
Suite ........ : déclaration. Il s'agit des segments liés au "transport"
Suite ........ : Spécifictié URSSAF
Mots clefs ... : PAIE, PGDUCSEDI
*****************************************************************}
procedure Groupe4_5ACOSS(EnregDucsEdi : TDucsEdi;EnrDucsEdiGroupee : TDucsEdiGroupee;Trait : string);
var
 blanc,localisation,taux1,montant1,taux2,montant2,taux3,montant3 : string;
 CuEnCours : string;
 LTransport : TOB;
 WMonnaieTenue : string;
begin
 if (Trait <> 'G') then
  WMonnaieTenue := EnregDucsEdi.MonnaieTenue
 else
  WMonnaieTenue := EnrDucsEdiGroupee.MonnaieTenue;

 blanc := StringOfChar(' ',35);
 localisation := '';
 taux1 := '';
 montant1 := '';
 taux2 := '';
 montant2 := '';
 taux3 := '';
 montant3 := '';
 CuEnCours := '';
 if TOB_Transport <> NIL then
    begin
     LTransport := TOB_Transport.FindFirst ([''],[''],TRUE);
     While LTransport <> NIL  do
      begin
         if (CuEnCours <> '') and
            (CuEnCours <> LTransport.getValue('PDD_COMURBAINE')) then
          begin
           Writeln (FDUCSEDI,'ACOLOC'+
                             Localisation+
                             copy(blanc,1,(4-length(Localisation)))+
                             Taux1+
                             copy(blanc,1,(8-length(Taux1)))+
                             Montant1+
                             copy(blanc,1,(18-length(Montant1)))+
                             Taux2+
                             copy(blanc,1,(8-length(Taux2)))+
                             Montant2+
                             copy(blanc,1,(18-length(Montant2)))+
                             Taux3+
                             copy(blanc,1,(8-length(Taux3)))+
                             Montant3+
                             copy(blanc,1,(18-length(Montant3)))+
                             WMonnaieTenue);

            localisation := '';
            taux1 := '';
            montant1 := '';
            taux2 := '';
            montant2 := '';
            taux3 := '';
            montant3 := '';
            CuEnCours := '';
           end;

          if (CuEnCours = LTransport.getValue('PDD_COMURBAINE')) or
             (CuEnCours = '') then
            begin
             if (Taux1 = '') then
               begin
                 Taux1 := FormatFloat('#0',(LTransport.getValue('PDD_TAUXCOTISATION')*10000));
                 Montant1 := FormatFloat('#0',(LTransport.getValue('PDD_BASECOTISATION')));
               end
             else
              if (Taux2 = '') then
                begin
                 Taux2 := FormatFloat('#0',(LTransport.getValue('PDD_TAUXCOTISATION')*10000));
                 Montant3 := FormatFloat('#0',(LTransport.getValue('PDD_BASECOTISATION')));
                end
              else
               if (Taux3 = '') then
                begin
                 Taux3 := FormatFloat('#0',(LTransport.getValue('PDD_TAUXCOTISATION')*10000));
                 Montant3 := FormatFloat('#0',(LTransport.getValue('PDD_BASECOTISATION')));
                end;
             end;
           CuEnCours := LTransport.getValue('PDD_COMURBAINE');
           Localisation := LTransport.getValue('PDD_COMURBAINE');

         LTransport := TOB_Transport.FindNext ([''],[''],TRUE);
      end; //3
     if (CuEnCours <> '') then
     // dernière CU
       begin
         Writeln (FDUCSEDI,'ACOLOC'+
                           Localisation+
                           copy(blanc,1,(4-length(Localisation)))+
                           Taux1+
                           copy(blanc,1,(8-length(Taux1)))+
                           Montant1+
                           copy(blanc,1,(18-length(Montant1)))+
                           Taux2+
                           copy(blanc,1,(8-length(Taux2)))+
                           Montant2+
                           copy(blanc,1,(18-length(Montant2)))+
                           Taux3+
                           copy(blanc,1,(8-length(Taux3)))+
                           Montant3+
                           copy(blanc,1,(18-length(Montant3)))+
                           WMonnaieTenue);

         Localisation := '';
         taux1 := '';
         montant1 := '';
         taux2 := '';
         montant2 := '';
         taux3 := '';
         montant3 := '';
         CuEnCours := '';
       end;
    end;
end;
// fin Groupe4_5ACOSS

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 13/05/2002
Modifié le ... :   /  /    
Description .. : Traitement des segment du groupe 3 de la déclaration.
Suite ........ : Spécificité IRC
Mots clefs ... : PAIE, PGDUCSEDI
*****************************************************************}
procedure Groupe3IRC(var EnregDucsEdi : TDucsEdi; EnrDucsEdiGroupee : TDucsEdiGroupee; Trait : string; var Nolg : integer);// PT16-2

var
   NatureAssiette, QualifiantCotis,ConditionSpe,Institution, GrpCotisant : string;
   CategSocio, NoLigne, DateChgtTx, Taux, Cotisation, Base, forfait :  string;
   EffCotisation, EffForfait, blanc : string;
   LesLignes : TOB;
// PT16-2   NoLg : integer;
   QQ : Tquery;
   WMonnaieTenue : string;
   DebCodif, SuitCodif : string;        // PT14-1

begin
  if (Trait <> 'G') then
    WMonnaieTenue := EnregDucsEdi.MonnaieTenue
  else
    WMonnaieTenue := EnrDucsEdiGroupee.MonnaieTenue;

   blanc := StringOfChar(' ',35);
   NatureAssiette :='';
   QualifiantCotis :='';
   ConditionSpe :='';
   Institution :='';
   GrpCotisant :='';
   CategSocio :='';
   NoLigne :='';
   DateChgtTx :='';
   Taux :='0';
   Cotisation:='0';
   Base :='0';
   forfait :='0';
   EffCotisation :='0';
   EffForfait :='0';
// PT16-2   NoLg := 0;

  if TOB_Lignes <> NIL then
    begin
     LesLignes := TOB_Lignes.FindFirst ([''],[''],TRUE);
     While LesLignes <> NIL  do
      begin
        if (Leslignes.getValue('PDD_TYPECOTISATION') <> 'I') and
           (Leslignes.getValue('PDD_TYPECOTISATION') <> 'S') then
        // les lignes Intitulé et Sous-Total ne sont pas traitées.
        begin
          NoLg := NoLg+1;

          NatureAssiette := copy(Leslignes.getValue('PDD_CODIFICATION'),6,2);
// PT14-1
          DebCodif := copy(Leslignes.getValue('PDD_CODIFICATION'),1,2);
          SuitCodif := copy(Leslignes.getValue('PDD_CODIFICATION'),3,5);
          while (((QualifiantCotis = '') or
                  (QualifiantCotis = NULL) or
                  (QualifiantCotis = '   ')) and (StrToInt(copy(DebCodif,2,1)) >= 0)) do
          begin
            QQ := OpenSql('SELECT PDP_PREDEFINI,PDP_COTISQUAL'+
                          ' FROM DUCSPARAM '+
                          ' WHERE ##PDP_PREDEFINI## PDP_CODIFICATION = "'+
// PT14-1                      Leslignes.getValue('PDD_CODIFICATION')+'"'+
                          DebCodif +  SuitCodif +'"' +
                          'ORDER BY PDP_PREDEFINI',True);
            While  not QQ.eof do
            begin
              QualifiantCotis :=QQ.FindField('PDP_COTISQUAL').AsString;  // PT11-1
              if (QQ.FindField('PDP_PREDEFINI').AsString = 'DOS') then break;
              QQ.Next;
            end;
            ferme (QQ);
            DebCodif := IntToStr(StrToInt(DebCodif)-1);
          end;
// PT14-1


          ConditionSpe :=Leslignes.getValue('PDD_CONDITION');

          Institution :=Leslignes.getValue('PDD_INSTITUTION');

          GrpCotisant := copy(Leslignes.getValue('PDD_CODIFICATION'),3,3);

          CategSocio :='';

          NoLigne := IntToStr(NoLg);

          if (Leslignes.getValue('PDD_DATECHGTTAUX') = IDate1900) then
           DateChgtTx := '0'
          else
           DateChgtTx := Copy(DateToStr(Leslignes.getValue('PDD_DATECHGTTAUX')),7,10)+
                         Copy(DateToStr(Leslignes.getValue('PDD_DATECHGTTAUX')),4,2)+
                         Copy(DateToStr(Leslignes.getValue('PDD_DATECHGTTAUX')),1,2);

          if (Leslignes.getValue('PDD_TYPECOTISATION')<> 'M') and
             (Leslignes.getValue('PDD_TYPECOTISATION')<> 'Q') then
           // ce n'est pas un forfait
// PT16-3  Taux :=formatFloat('#0',Leslignes.getValue('PDD_TAUXCOTISATION')*100);
           Taux :=formatFloat('#0',Leslignes.getValue('PDD_TAUXCOTISATION')*10000);

// PT19    Cotisation:=FormatFloat('#0',Leslignes.getValue('PDD_MTCOTISAT'));
           Cotisation:=FormatFloat('#0',Leslignes.getValue('PDD_MTCOTISAT')*100);

          if (Leslignes.getValue('PDD_TYPECOTISATION')<> 'M') and
             (Leslignes.getValue('PDD_TYPECOTISATION')<> 'Q') then
           // ce n'est pas un forfait
// PT19   Base :=FormatFloat('#0',Leslignes.getValue('PDD_BASECOTISATION'));
          Base :=FormatFloat('#0',Leslignes.getValue('PDD_BASECOTISATION')*100);

          if (Leslignes.getValue('PDD_TYPECOTISATION') = 'M') or
             (Leslignes.getValue('PDD_TYPECOTISATION') = 'Q') then
          // forfait
          begin
{ PT19           if  (Leslignes.getValue('PDD_BASECOTISATION') <> 0) then
              //base cotisation
              forfait :=FormatFloat('#0',Leslignes.getValue('PDD_BASECOTISATION'))
            else }
             // Montant cotisation arrondi/effectif
              forfait := FormatFloat('#0',
                        ((StrToFloat(FormatFloat('#0',Leslignes.getValue('PDD_MTCOTISAT')*100)))/
                        Leslignes.getValue('PDD_EFFECTIF')));
          end;

          EffCotisation := Leslignes.getValue('PDD_EFFECTIF');

{ PT16-5          if (Leslignes.getValue('PDD_TYPECOTISATION')<> 'M') and
             (Leslignes.getValue('PDD_TYPECOTISATION')<> 'Q') then
           // ce n'est pas un forfait }
          if (Leslignes.getValue('PDD_TYPECOTISATION') = 'M') or
             (Leslignes.getValue('PDD_TYPECOTISATION') = 'Q') then
           // forfait
//PT16-5    EffForfait :=Leslignes.getValue('PDD_EFFECTIF');
            EffForfait := IntToStr(StrToint(Leslignes.getValue('PDD_EFFECTIF')) * 10000);

          Writeln(FDUCSEDI,'AGI   0'+
                           NatureAssiette+
                           QualifiantCotis+
                           ConditionSpe+
                           Institution+
                           GrpCotisant+
                           CategSocio+
                           copy(blanc,1,(4-length(CategSocio)))+
                           NoLigne+
                           copy(blanc,1,(6-length(NoLigne)))+
                           DateChgtTx+
                           copy(blanc,1,(8-length(DateChgtTx)))+
                           Taux+
                           copy(blanc,1,(8-length(Taux)))+
                           Cotisation+
                           copy(blanc,1,(18-length(Cotisation)))+
                           Base+
                           copy(blanc,1,(18-length(Base)))+
                           forfait+
                           copy(blanc,1,(18-length(forfait)))+
                           EffCotisation+
                           copy(blanc,1,(18-length(EffCotisation)))+
                           EffForfait+
                           copy(blanc,1,(18-length(EffForfait)))+
                           EnregDucsEdi.MonnaieTenue);

          EnregDucsEdi.NbCot := EnregDucsEdi.NbCot+1;

          NatureAssiette :='';
          QualifiantCotis :='';
          ConditionSpe :='';
          Institution :='';
          GrpCotisant :='';
          CategSocio :='';
          NoLigne :='';
          DateChgtTx :='';
          Taux :='0';
          Cotisation:='0';
          Base :='0';
          forfait :='0';
          EffCotisation :='0';
          EffForfait :='0';
        end;
        LesLignes := TOB_Lignes.FindNext ([''],[''],TRUE);
      end;
    end;
end;
// fin Groupe3IRC

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 13/05/2002
Modifié le ... :   /  /
Description .. : Traitement des segments du groupe 3 de la déclaration
Suite ........ : Spécificité ASSEDIC
Mots clefs ... : PAIE, PGDUCSEDI
*****************************************************************}
procedure Groupe3UNEDIC(var EnregDucsEdi : TDucsEdi);
var
   blanc, NatureAssiette, DateD, DateF, Taux, Base, Forfait : string;
   LesLignes, LaLigneSuivante                               : TOB;
   i                                                        : integer; // PT22
begin
   blanc := StringOfChar(' ',35);
   NatureAssiette := '';
// PT22   DateRef := '0' ;
   Taux := '';
   Base := '';
   Forfait := '0';

{  PT22 if TOB_Lignes <> NIL then
    begin
     LesLignes := TOB_Lignes.FindFirst ([''],[''],TRUE);
     While LesLignes <> NIL  do
      begin
        if (Leslignes.getValue('PDD_TYPECOTISATION') <> 'I') and
           (Leslignes.getValue('PDD_TYPECOTISATION') <> 'S') then
        // les lignes Intitulé et Sous-Total ne sont pas traitées.
        begin
        NatureAssiette := copy(Leslignes.getValue('PDD_CODIFICATION'),5,3);
        DateRef := '0' ; // période de référence (déclaration rétroactive) !! Non Traité
//PT16-3Taux := FormatFloat ('#0',(Leslignes.getValue('PDD_TAUXCOTISATION')*100));
        Taux := FormatFloat ('#0',(Leslignes.getValue('PDD_TAUXCOTISATION')*10000));
        Base := FormatFloat ('#0',Leslignes.getValue('PDD_BASECOTISATION'));
        Forfait := '0'; // facultatif --> !! Non Traité

        Writeln(FDUCSEDI,'CUN   '+
                         NatureAssiette+
                         DateRef+
                         copy(blanc,1,(8-length(DateRef)))+
                         Taux+
                         copy(blanc,1,(8-length(Taux)))+
                         Base+
                         copy(blanc,1,(18-length(Base)))+
                         Forfait+
                         copy(blanc,1,(18-length(Forfait)))+
                         EnregDucsEdi.MonnaieTenue);

        EnregDucsEdi.NbCot := EnregDucsEdi.NbCot+1;

        NatureAssiette := '';
        DateRef := '0' ;
        Taux := '';
        Base := '';
        Forfait := '0';
        end;
        LesLignes := TOB_Lignes.FindNext ([''],[''],TRUE);
      end;
    end;}
    if TOB_Lignes <> NIL then
    begin
      i := 0;

      While (i < TOB_Lignes.Detail.Count)  do
      begin
        LesLignes := TOB_Lignes.Detail[i];
        DateD := '0';
        DateF := '0';
        if (Leslignes.getValue('PDD_TYPECOTISATION') <> 'I') and
           (Leslignes.getValue('PDD_TYPECOTISATION') <> 'S') then
        // les lignes Intitulé et Sous-Total ne sont pas traitées.
        begin
        NatureAssiette := copy(Leslignes.getValue('PDD_CODIFICATION'),5,3);

        if (Leslignes.getValue('PDD_DATECHGTTAUX') <> IDate1900) then
        begin
          DateD := Copy(DateToStr(Leslignes.getValue('PDD_DATECHGTTAUX')),7,10)+
                   Copy(DateToStr(Leslignes.getValue('PDD_DATECHGTTAUX')),4,2)+
                   Copy(DateToStr(Leslignes.getValue('PDD_DATECHGTTAUX')),1,2);
          DateF := Copy(DateToStr(EnregDucsEdi.DateFin),7,10)+
                   Copy(DateToStr(EnregDucsEdi.DateFin),4,2)+
                   Copy(DateToStr(EnregDucsEdi.DateFin),1,2);

          if (i+1 < TOB_Lignes.Detail.Count) then
          begin
            LaLigneSuivante := TOB_Lignes.Detail[i+1];
            if(LaLigneSuivante <> NIL) then
            begin
              if (copy(LaLigneSuivante.getValue('PDD_CODIFICATION'),5,3) =
                  NatureAssiette) then
                if (LaLigneSuivante.getValue('PDD_DATECHGTTAUX') <> IDate1900) then
                  DateF := copy(DateToStr(PlusDate(LaLigneSuivante.getValue('PDD_DATECHGTTAUX'),-1,'J')),7,4)+
                           copy(DateToStr(PlusDate(LaLigneSuivante.getValue('PDD_DATECHGTTAUX'),-1,'J')),4,2)+
                           copy(DateToStr(PlusDate(LaLigneSuivante.getValue('PDD_DATECHGTTAUX'),-1,'J')),1,2)
                else
                  DateF := copy(DateToStr(EnregDucsEdi.DateFin),7,4)+
                           copy(DateToStr(EnregDucsEDi.DateFin),1,2)+
                           copy(DateToStr(EnregDucsEDi.DateFin),4,2);
            end;
          end;
        end;
        Taux := FormatFloat ('#0',(Leslignes.getValue('PDD_TAUXCOTISATION')*10000));
        Base := FormatFloat ('#0',Leslignes.getValue('PDD_BASECOTISATION'));
        Forfait := '0'; // facultatif --> !! Non Traité

        Writeln(FDUCSEDI,'CUN   '+
                         NatureAssiette+
                         DateD+
                         copy(blanc,1,(8-length(DateD)))+
                         DateF+
                         copy(blanc,1,(8-length(DateF)))+
                         Taux+
                         copy(blanc,1,(8-length(Taux)))+
                         Base+
                         copy(blanc,1,(18-length(Base)))+
                         Forfait+
                         copy(blanc,1,(18-length(Forfait)))+
                         EnregDucsEdi.MonnaieTenue);

        EnregDucsEdi.NbCot := EnregDucsEdi.NbCot+1;

        NatureAssiette := '';
//        DateRef := '0' ;
        Taux := '';
        Base := '';
        Forfait := '0';
        end;
        i := i+1;
      end;
    end;
end;
// fin Groupe3UNEDIC
{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 13/05/2002
Modifié le ... :   /  /
Description .. : Traitement des segments de fin de déclaration
Mots clefs ... : PAIE, PGDUCSEDI
*****************************************************************}
procedure FinDeclaration(var EnregDucsEdi : TDucsEdi;EnrDucsEdiGroupee : TDucsEdiGroupee);
var
   blanc,Montant : string;
// PT8-2   valR, valA : double;
begin
     blanc := StringOfChar(' ',35);

     TOB_Lignes.Free; TOB_Lignes := NIL;

{PT14-2     if ((EnregDucsEdi.NbTransport > 1) or
         (EnrDucsEdiGroupee.NbTransport > 1)) and
        (copy(EnregDucsEdi.Typducs,1,1) = 'A') then}
     if (TOB_Transport <> Nil) then
     begin
       TOB_Transport.Free;
       TOB_Transport := NIL;
     end; // PT14-2

//PT34    TOB_DucsGroupee.Free; TOB_DucsGroupee := NIL;
      FreeAndNIl(TOB_DucsGroupee);

     // UNS   S - Délimiteur de section de message
     // ------------------------------------------
     Writeln(FDUCSEDI,'UNS   S');

     // MOATOT - Montant Total
     // ----------------------
     if (copy(EnregDucsEdi.TypDucs,1,1) = 'I')  then
      // IRC  : monatnt déclaré  en cts
      Montant := FormatFloat ('#0',(EnregDucsEdi.MtTotal*100))
     else
      // ACOSS ou UNEDIC  : montant payé à l'euro ( ou au franc)
      // Pour l'ACOSS : ne concerne que les déclarations multi établissements.
      if ((EnregDucsEdi.DucsDossier = TRUE) and
         (EnregDucsEdi.PaiementGroupe = TRUE)) or
         (copy(EnregDucsEdi.TypDucs,1,1) = 'U')  then           // PT4-2
        Montant := FormatFloat ('#0',EnregDucsEdi.MtTotal);

      if ((EnregDucsEdi.DucsDossier = TRUE) and
          (EnregDucsEdi.PaiementGroupe = TRUE) and
          (copy(EnregDucsEdi.TypDucs,1,1) = 'A')) or
//PT15-8          (copy(EnregDucsEdi.TypDucs,1,1) = 'U')  then   // PT10-1
          (copy(EnregDucsEdi.TypDucs,1,1) <> 'A')  then   
        Writeln(FDUCSEDI,'MOATOT'+
                         Montant+
                         copy(blanc,1,(18-length(Montant)))+
                         EnregDucsEdi.MonnaieTenue);
// d PT8-2
      Writeln(FRapport,'    Montant total des cotisations : '+FormatFloat ('#0',EnregDucsEdi.MtTotal)+' '+EnregDucsEdi.MonnaieTenue);
      Writeln(FRapport,'');
// f PT8-2
      // MOAREG - Montant de la régularisation (IRC seulement)
      // -----------------------------------------------------
{ dPT8-2
      if (copy(EnregDucsEdi.TypDucs,1,1) = 'I') then
        begin
          valR := EnregDucsEdi.MtTotRegul;
          valA := EnregDucsEdi.MtTotAcpte ;
          if (valR > 0) or (valA > 0) then
            begin
              if (valR < 0) then valR := 0;
              if (valA < 0) then valA := 0;
              Montant := Formatfloat('#0',((valR+valA)*100));
              Writeln (FDUCSEDI,'MOAREG'+
                                Montant+
                                copy(blanc,1,(18-length(Montant)))+
                                EnregDucsEdi.MonnaieTenue);
              Writeln(FRapport,'    Montant de la régularisation : '+Formatfloat('#0',(valR+valA))+' '+EnregDucsEdi.MonnaieTenue);
            end;
        end;  }
     if (copy(EnregDucsEdi.TypDucs,1,1) = 'I') then
        begin
          if (EnregDucsEdi.MtTotRegul <> 0.0) then
            begin
              Montant := Formatfloat('#0',(EnregDucsEdi.MtTotRegul*100));
              Writeln (FDUCSEDI,'MOAREG'+
                                Montant+
                                copy(blanc,1,(18-length(Montant)))+
                                EnregDucsEdi.MonnaieTenue);
              Writeln(FRapport,'    Montant de la régularisation : '+Formatfloat('#0',EnregDucsEdi.MtTotRegul)+' '+EnregDucsEdi.MonnaieTenue);
            end;
        end;
// f PT8-2        
      // MOAACP - Montant de l'acompte (IRC seulement)
      // -----------------------------------------------------
{ d PT8-2
      if (copy(EnregDucsEdi.TypDucs,1,1) = 'I') then
        begin
          valR := EnregDucsEdi.MtTotRegul;
          valA := EnregDucsEdi.MtTotAcpte;
          if (valR < 0) or (valA < 0) then
            begin
              if (valR > 0) then valR := 0;
              if (valA > 0) then valA := 0;
              Montant := Formatfloat('#0',(((valR*-1)+(valA*-1))*100));
              Writeln (FDUCSEDI,'MOAACP'+
                                Montant+
                                copy(blanc,1,(18-length(Montant)))+
                                EnregDucsEdi.MonnaieTenue);
              Writeln(FRapport,'    Montant de l''acompte : '+Formatfloat('#0',((valR*-1)+(valA*-1)))+' '+EnregDucsEdi.MonnaieTenue);
            end;
        end; }
     if (copy(EnregDucsEdi.TypDucs,1,1) = 'I') then
        begin
          if (EnregDucsEdi.MtTotAcpte <> 0.0) then
            begin
              Montant := Formatfloat('#0',(EnregDucsEdi.MtTotAcpte*100));
              Writeln (FDUCSEDI,'MOAACP'+
                                Montant+
                                copy(blanc,1,(18-length(Montant)))+
                                EnregDucsEdi.MonnaieTenue);
              Writeln(FRapport,'    Montant de l''acompte : '+Formatfloat('#0',EnregDucsEdi.MtTotAcpte)+' '+EnregDucsEdi.MonnaieTenue);
            end;
        end;
// f PT8-2        
end;
// fin FinDeclaration
{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 13/05/2002
Modifié le ... :   /  /
Description .. : Traitement des segments liés au paiement
Mots clefs ... : PAIE, PGDUCSEDI
*****************************************************************}
// PT7-3 procedure PaiementEdi(var EnregDucsEdi : TDucsEdi;EnrDucsEdiGroupee : TDucsEdiGroupee; Trait :string);
procedure PaiementEdi(var EnregDucsEdi : TDucsEdi;EnrDucsEdiGroupee : TDucsEdiGroupee; Trait :string; var DossOK : boolean);
var
   intervenant,Montant, blanc : string;
   tampon,creancier, ModePaiement, NumeroInterne : string;
   paiement : TdateTime ;
   CodeBanque ,NoCpte, CleRib, NoBanque,NoGuichet : string;
   QQ : TQuery;
   StQQ, LibPaiement, LibMt,st : string;
   typenr : string;   // PT1-4
   LBanque, LDom : String; // PT6

begin
     CodeBanque:= ''; // PT7-3

     if (EnregDucsEdi.Neant = True) then
// d PT6
       // Ducs Néant - pas de paiement
       begin
        st := '  Etablissement : ';
        if (trait <> 'G') then
          st := st + EnregDucsEdi.Etab+' '+EnregDucsEdi.LibelleET
        else
          st := st + EnrDucsEdiGroupee.Etab+' '+EnrDucsEdiGroupee .LibelleET ;

        Writeln(FRapport, st);
        Writeln(FRapport, '    Ducs Néant');

        exit;
       end;
// f PT6

     tampon := '';
     creancier := '';
     blanc := StringOfChar(' ',35);

     if (Trait <> 'G') then
      begin
        ModePaiement := EnregDucsEdi.ModePaiement;
        Montant := FormatFloat ('#0',EnregDucsEdi.MtApayer);
// d PT15-9
        if (copy(EnregDucsEdi.TypDucs,1,1) = 'I')  then
          Montant := FormatFloat ('#0',EnregDucsEdi.MtApayer * 100);
// f PT15-9
        LibMt := FormatFloat ('#0.00',EnregDucsEdi.MtApayer)+' '+
                              EnregDucsEdi.MonnaieTenue;
        NumeroInterne := EnregDucsEdi.NumeroInterne;
        paiement := EnregDucsEdi.reglement
//PT13        paiement := EnregDucsEdi.paiement;
      end
     else
      begin
        ModePaiement := EnrDucsEdiGroupee.ModePaiement;
        Montant := FormatFloat ('#0',EnrDucsEdiGroupee.MtApayer);
 // d PT15-9
        if (copy(EnregDucsEdi.TypDucs,1,1) = 'I')  then
          Montant := FormatFloat ('#0',EnrDucsEdiGroupee.MtApayer * 100);
// f PT15-9
       LibMt := FormatFloat ('#0.00',EnrDucsEdiGroupee.MtApayer)+' '+
                              EnrDucsEdiGroupee.MonnaieTenue;
        NumeroInterne := EnrDucsEdiGroupee.NumeroInterne;
//      paiement := EnregDucsEdi.paiement;   //PT13
        paiement := EnregDucsEdi.reglement;
      end;

    if (ModePaiement = '20') then libPaiement := 'Chèque';
    if (ModePaiement = '30') then libPaiement := 'Virement';
    if (ModePaiement = '31') then libPaiement := 'Prélèvement';
    if (ModePaiement = 'Z10') then libPaiement := 'Télé-Règlement EDI';
    st := '  Etablissement : ';
     if (trait <> 'G') then
       st := st + EnregDucsEdi.Etab+' '+EnregDucsEdi.LibelleET
     else
       st := st + EnrDucsEdiGroupee.Etab+' '+EnrDucsEdiGroupee .LibelleET ;

    Writeln(FRapport, st);
    Writeln(FRapport, '    Paiement par '+Libpaiement+'  '+LibMT);

     // PAI   - Mode de paiement
     //-------------------------
// d PT1-4
    typenr := 'PAI   ';
    if (copy(EnregDucsEdi.TypDucs,1,1) = 'I') and (ModePaiement = '20') then
      typenr := 'PAIC  ';
    if (copy(EnregDucsEdi.TypDucs,1,1) = 'I') and (ModePaiement = '30') then
      typenr := 'PAIV  ';
    if (copy(EnregDucsEdi.TypDucs,1,1) = 'I') and (ModePaiement = '31') then
      typenr := 'PAIP  ';
    if (copy(EnregDucsEdi.TypDucs,1,1) = 'I') and (ModePaiement = 'Z10') then
      typenr := 'PAIZ  ';
//     Writeln (FDUCSEDI,'PAI   '+
     Writeln (FDUCSEDI,typenr+
// f PT1-4
                         ModePaiement+
                         copy(blanc,1,(3-length(ModePaiement))));

     // FII   - Identification Financière
     // ---------------------------------
// d PT1-4
    typenr := 'FII   ';
    if (copy(EnregDucsEdi.TypDucs,1,1) = 'I') and (ModePaiement = '30') then
      typenr := 'FIIV  ';
    if (copy(EnregDucsEdi.TypDucs,1,1) = 'I') and (ModePaiement = '31') then
      typenr := 'FIIP  ';
    if (copy(EnregDucsEdi.TypDucs,1,1) = 'I') and (ModePaiement = 'Z10') then
      typenr := 'FIIZ  ';
// f PT1-4

     if (ModePaiement = '30') or
        (ModePaiement = '31') or
        (ModePaiement = 'Z10') then
       // Virement, Prélèvement, Télérèglement EDI
       begin
         if (ModePaiement = '30') then
// d PT1-5
           // Virement - bénéficiaire
           begin
             intervenant := 'BF ';
             StQQ:= '';
             if (Trait <> 'G') then
               StQQ := '"'+EnregDucsEdi.Etab+'"'
             else
               StQQ := '"'+EnrDucsEdiGroupee.Etab+'"';

             QQ := OpenSQL('SELECT POG_RIBDUCSEDI FROM ORGANISMEPAIE WHERE '+
                           'POG_ETABLISSEMENT='+StQQ+
                           ' AND POG_ORGANISME="'+EnregDucsEdi.Organisme+'"',True) ;
             if Not QQ.EOF then
               begin
                 CodeBanque :=QQ.FindField('POG_RIBDUCSEDI').AsString;
               end;
             Ferme(QQ) ;
//d PT7-3
             if (CodeBanque = '') then
               begin
                 Writeln(FRapport, '  Anomalie   : Virement, la'+
                             ' banque bénéficiaire n''est pas renseignée');

                 Writeln(FRapport, '  --> Action : Modification organisme : renseigner'+
                             ' la banque bénéficiaire ou bien modifier le mode de'+
                             ' règlement ');
                 DossOK := False;
                 exit;
// f PT7-3
               end ;

             QQ := OpenSQL('SELECT PBO_NUMCPTEBQORG,PBO_CLEBQORG,PBO_ETABBQORG,'+
                           'PBO_GUICHBQORG,PBO_LIBELLEBQORG,PBO_DOMBQORG '+
                           'FROM PAIEBQORG WHERE '+
                           'PBO_CODEBQORG="'+CodeBanque+'"' ,True) ;
             if Not QQ.EOF then
               begin
                 NoCpte :=QQ.FindField('PBO_NUMCPTEBQORG').AsString;
                 CleRib :=QQ.FindField('PBO_CLEBQORG').AsString;
                 NoBanque :=QQ.FindField('PBO_ETABBQORG').AsString;
                 NoGuichet :=QQ.FindField('PBO_GUICHBQORG').AsString;
                 LBanque := QQ.FindField('PBO_LIBELLEBQORG').AsString;
                 LDom   := QQ.FindField('PBO_DOMBQORG').AsString;
               end ;
             Ferme(QQ) ;
           end
         else
           // donneur d'ordre - pélèvement ou télérèglement EDI
           begin
           intervenant := 'OR ';
           StQQ:= '';
           if (Trait <> 'G') then
             StQQ := '"'+EnregDucsEdi.Etab+'"'
           else
             StQQ := '"'+EnrDucsEdiGroupee.Etab+'"';
// d PT26
             QQ := OpenSQL('SELECT POG_RIBDUCSEDI FROM ORGANISMEPAIE WHERE '+
                           'POG_ETABLISSEMENT='+StQQ+
                           ' AND POG_ORGANISME="'+EnregDucsEdi.Organisme+'"',True) ;
             if Not QQ.EOF then
               begin
                 CodeBanque :=QQ.FindField('POG_RIBDUCSEDI').AsString;
               end;
             Ferme(QQ) ;
// f PT26

{PT26           QQ := OpenSQL('SELECT ETB_RIBDUCSEDI FROM ETABCOMPL WHERE '+
                         'ETB_ETABLISSEMENT='+StQQ ,True) ;
           if Not QQ.EOF then
            begin
             CodeBanque :=QQ.FindField('ETB_RIBDUCSEDI').AsString;
            end ;
           Ferme(QQ) ;}
//d PT7-3
           if (CodeBanque = '') then
              begin
                Writeln(FRapport, '  Anomalie   : pélèvement ou télérèglement, la'+
                             ' banque du donneur d''ordre n''est pas renseignée');

                Writeln(FRapport, '  --> Action : Modification établissement social : renseigner'+
                             ' l''identifiant bancaire pour le paiement des charges sociale ou bien modifier le mode de'+
                             ' règlement ');
                DossOK := False;
                exit;
              end;
// f PT7-3

           QQ := OpenSQL('SELECT BQ_NUMEROCOMPTE,BQ_CLERIB,BQ_ETABBQ,'+
                         'BQ_GUICHET,BQ_LIBELLE,BQ_DOMICILIATION '+
                         'FROM BANQUECP WHERE '+
                         'BQ_GENERAL="'+CodeBanque+'"' ,True) ;    //PT3
           if Not QQ.EOF then
            begin
             NoCpte :=QQ.FindField('BQ_NUMEROCOMPTE').AsString;
             CleRib :=QQ.FindField('BQ_CLERIB').AsString;
             NoBanque :=QQ.FindField('BQ_ETABBQ').AsString;
             NoGuichet :=QQ.FindField('BQ_GUICHET').AsString;
             LBanque := QQ.FindField('BQ_LIBELLE').AsString;
             LDom   := QQ.FindField('BQ_DOMICILIATION').AsString;

            end ;
           Ferme(QQ) ;
           end;
// f  PT1-5

// PT1-4           Writeln (FDUCSEDI,'FII   '+
           Writeln (FDUCSEDI,typenr+    // PT1-4
                             intervenant+
                             NoCpte+
                             CleRib+
                             copy(blanc,1,(35-(length(NoCpte)+length(CleRib))))+
                             NoBanque+
                             copy(blanc,1,(11-(length(NoBanque))))+
                             NoGuichet+
                             copy(blanc,1,(17-(length(NoGuichet)))));
// d PT6
          if (ModePaiement = '30') then
            Writeln(FRapport, '    RIB du bénéficiaire : '+NoBanque+
                            '  '+NoGuichet+'  '+NoCpte+'  '+CleRib)
          else
            Writeln(FRapport, '    RIB du donneur d''ordre : '+NoBanque+
                            '  '+NoGuichet+'  '+NoCpte+'  '+CleRib);
          Writeln(FRapport,'                              '+LBanque+'  '+LDom);

// f PT6
       end;

       // MOAPAI - Montant réglé
       // ----------------------
// d PT1-4
    typenr := 'MOAPAI';
    if (copy(EnregDucsEdi.TypDucs,1,1) = 'I') and (ModePaiement = '20') then
      typenr := 'MOAPAC';
    if (copy(EnregDucsEdi.TypDucs,1,1) = 'I') and (ModePaiement = '30') then
      typenr := 'MOAPAV';
    if (copy(EnregDucsEdi.TypDucs,1,1) = 'I') and (ModePaiement = '31') then
      typenr := 'MOAPAP';
    if (copy(EnregDucsEdi.TypDucs,1,1) = 'I') and (ModePaiement = 'Z10') then
      typenr := 'MOAPAZ';
//       Writeln(FDUCSEDI,'MOAPAI'+
       Writeln(FDUCSEDI,typenr+
// f PT1-4
                        Montant+
                        copy(blanc,1,(18-length(Montant)))+
                        EnregDucsEdi.MonnaieTenue);  // PT4-2

       // RFFqual - Référence sociale (chèque, virement, télérèglement EDI)
       // -----------------------------------------------------------------
       if (ModePaiement <> '31') and
          (ModePaiement <> 'Z11') then
         begin
           if (ModePaiement = '20') or
              (ModePaiement = '30') then
              // Chèque ou virement
              begin
// d PT1-4
               typenr := 'RFFPQ ';
               if (copy(EnregDucsEdi.TypDucs,1,1) = 'I') and (ModePaiement = '20') then
                 typenr := 'RFFPQC';
               if (copy(EnregDucsEdi.TypDucs,1,1) = 'I') and (ModePaiement = '30') then
                 typenr := 'RFFPQV';

               EnregDucsEDi.CleIBAN :='00';
// pour test       "281199923999999999999990000000" --> clé 07
// pour test CleIBAN ('S1','199923','99999999999999',EnregDucsEDi.CleIBAN);
               if (copy(EnregDucsEdi.TypDucs,1,1) = 'A') then creancier := 'S1';
               if (copy(EnregDucsEdi.TypDucs,1,1) = 'U') then creancier := 'S2';
               if (copy(EnregDucsEdi.TypDucs,1,1) = 'I') then creancier := 'S3';

               if (copy(EnregDucsEdi.TypDucs,1,1) = 'U') then    // PT9-2
                EnregDucsEDi.CleIBAN := '99'    // toujours 99 pour UNEDIC
               else
                CleIBAN (creancier,
                        copy (DateToStr(EnregDucsEDi.DateFin),7,2)+EnregDucsEdi.Abrege,
                        copy(NumeroInterne,1,19),
                        EnregDucsEDi.CleIBAN);
// PT1-4               Writeln (FDUCSEDI, 'RFFPQ '+
// d PT27
               if (copy(EnregDucsEdi.TypDucs,1,1) = 'U') then
                 Writeln (FDUCSEDI, typenr+
                                  EnregDucsEDi.CleIBAN+
                                  creancier+
                                  copy(DateToStr(EnregDucsEdi.DateFin),7,2)+
                                  PeriodRef+
                                  NumeroInterne+
                                  copy(blanc,1,(35-(8+length(copy(NumeroInterne,1,19))))))
               else
// f PT27
                 Writeln (FDUCSEDI, typenr+ // PT1-4
                                  EnregDucsEDi.CleIBAN+        // PT2-1
                                  creancier+
                                  copy(DateToStr(EnregDucsEdi.DateFin),7,2)+
                                  EnregDucsEdi.Abrege+
                                  NumeroInterne+
                                  copy(blanc,1,(35-(8+length(copy(NumeroInterne,1,19))))));
               if (ModePaiement = '30') then
                 // Virement
                 st :='  clé et référence sociale à renseigner dans la partie'+
                      ' référence commerciale du virement :'
               else
                 // chèque
                 st := '   clé et référence sociale à inscrire au dos du chèque:';

               Writeln (FRapport,st);
// d PT27
               if (copy(EnregDucsEdi.TypDucs,1,1) = 'U') then
                 Writeln (FRapport,'    Clé : '+EnregDucsEDi.CleIBAN+
                                ' Référence : '+creancier+' '+
                                copy(DateToStr(EnregDucsEdi.DateFin),7,2)+
                                PeriodRef+ ' '+
                                NumeroInterne)
               else
// f PT27
                 Writeln (FRapport,'    Clé : '+EnregDucsEDi.CleIBAN+
                                ' Référence : '+creancier+' '+
                                copy(DateToStr(EnregDucsEdi.DateFin),7,2)+
                                EnregDucsEdi.Abrege+ ' '+
                                NumeroInterne);

              end;
              if (ModePaiement = 'Z10') then
              // Télérèglement
              begin
// d PT1-4
                typenr := 'RFFCR ';
                if (copy(EnregDucsEdi.TypDucs,1,1) = 'I') and (ModePaiement = 'Z10') then
                  typenr :='RFFCRZ';
// f PT1-4
                tampon:=EnregDucsEdi.IdentOPS+copy(blanc,1,(8-length(EnregDucsEdi.IdentOPS)));
// d PT27
                if (copy(EnregDucsEdi.TypDucs,1,1) = 'U') then
                  tampon:= tampon + PeriodRef
                else
// f PT27
                  tampon:= tampon + EnregDucsEdi.Abrege;
// PT1-4                Writeln (FDUCSEDI, 'RFFCR '+
                Writeln (FDUCSEDI, typenr+   // PT1-4
                                   tampon+
                                   copy(blanc,1,(35-length(tampon))));
              end;
         end;

       // DTMPAI - Date de paiement (si date de paiement renseignée quand
       //                            prélèvement ou Télérèglement EDI ou
       //                            IRC et Chèque ou IRC et virement)
       // ---------------------------------------------------------------
// d PT1-4
       typenr := 'DTMPAI';
       if (copy(EnregDucsEdi.TypDucs,1,1) = 'I') and (ModePaiement = '20') then
         typenr := 'DTMPAC';
       if (copy(EnregDucsEdi.TypDucs,1,1) = 'I') and (ModePaiement = '30') then
         typenr := 'DTMPAV';
       if (copy(EnregDucsEdi.TypDucs,1,1) = 'I') and (ModePaiement = '31') then
         typenr := 'DTMPAP';
       if (copy(EnregDucsEdi.TypDucs,1,1) = 'I') and (ModePaiement = 'Z10') then
         typenr := 'DTMPAZ';
// f  PT1-4
       if (paiement <> IDate1900) then
          if (ModePaiement = '31') or
             (ModePaiement = 'Z10') or
             ((copy(EnregDucsEdi.TypDucs,1,1) = 'I') and        // PT25
// PT25            (ModePaiement = '20') or
              (ModePaiement = '30')) then
//             Writeln (FDUCSEDI,'DTMPAI'+
               Writeln (FDUCSEDI,typenr+   // PT1-4
                               copy(DateToStr(Paiement),7,4)+
// d PT4-1
                               copy(DateToStr(Paiement),4,2)+
                               copy(DateToStr(Paiement),1,2)+
//                               copy(DateToStr(Paiement),1,2)+
//                               copy(DateToStr(Paiement),4,2)+
// f PT4-1
                              copy(blanc,1,(35-8)));
       Writeln (FRapport,'');
end;
// fin PaiementEdi
{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 28/05/2002
Modifié le ... :   /  /
Description .. : Traitement des derniers segments du ficheir Ducs Edi
Mots clefs ... : PAIE, PGDUCSEDI
*****************************************************************}
procedure FinFin (var EnregDucsEdi : TDucsEdi);
var
 blanc : string;

begin
 blanc := StringOfChar(' ',35);
 // AUT   - Authentification déclarant
 // ----------------------------------
 Writeln (FDUCSEDI,'AUT   '+
                   EnregDucsEdi.LibelleET+
                   Copy(blanc,1,(35-length(EnregDucsEdi.LibelleET))));

 // CNT   - Nombre de lignes de cotisation (IRC uniquemet)
 // ------------------------------------------------------
 if (copy(EnregDucsEdi.TypDucs,1,1) = 'I') then   // PT16-4
     Writeln (FDUCSEDI, 'CNT   '+
                       copy(blanc,1,(18-length(IntToStr(EnregDucsEdi.NbCot))))+
                       IntToStr(EnregDucsEdi.NbCot));

 // UNT   1 - Fin de la déclaration
 // -------------------------------
 Writeln (FDUCSEDI,'UNT   1');

 Writeln(FRapport, '');
 Writeln(FRapport, '  fichier OK');
 CloseFile(FDUCSEDI);
end;
// fin FinFin
procedure MajEnvoiSocialDUCS (EnregDucsEdi : TDucsEdi);
var
   FileAttrs, Ordre : integer;
   sr : TSearchRec;
   EnregEnvoi : TEnvoiSocial;
   Size : double;
   Millesime, StQQ : string;
   QQ : Tquery;
begin
   Size := 0.0; // PT7-1
   // calcul de la taille du fichier confectionné
   FileAttrs := 0;
   FileAttrs := FileAttrs + faAnyFile;
// d PT21-2
{$IFDEF EAGLCLIENT}
   if FindFirst(VH_Paie.PgCheminEagl+'\'+NomFicDucs, FileAttrs, sr) = 0 then
{$ELSE}
   if FindFirst(V_PGI.DatPath+'\'+NomFicDucs, FileAttrs, sr) = 0 then
{$ENDIF}
// f PT21-2
      begin
      if (sr.Attr and FileAttrs) = sr.Attr then
         Size := Arrondi(sr.Size/1024, 2);
      SysUtils.FindClose(sr);
      end;

   // formatage du millésime
   Millesime := Copy(DateToStr(EnregDucsEdi.DateFin),7,4);
   StQQ:= 'SELECT CO_CODE'+
            ' FROM COMMUN WHERE'+
            ' CO_TYPE="PGA" AND'+
            ' CO_LIBELLE="'+Millesime+'"';
   QQ:=OpenSQL(StQQ,TRUE) ;
   if Not QQ.EOF then
      Millesime := QQ.FindField ('CO_CODE').AsString;
   Ferme (QQ);

   // suppression des éléments déjà existants dans ENVOISOCIAL
   StQQ := 'DELETE FROM ENVOISOCIAL WHERE'+
           ' PES_TYPEMESSAGE= "DUC" AND '+
           ' PES_MILLESSOC = "'+Millesime+'" AND'+
           ' PES_DATEDEBUT = "'+UsDateTime(EnregDucsEdi.DateDebut)+'" AND'+
           ' PES_DATEFIN = "'+UsDateTime(EnregDucsEdi.DateFin)+'" AND'+
           ' PES_SIRETDO = "'+EnregDucsEdi.Siret+'" AND'+
           ' PES_FRACTIONDADS = "Z" AND'+
//PT18           ' PES_LIBELLE = "'+EnregDucsEdi.LibelleET+'" AND'+
           ' PES_FICHIERRECU = "'+ExtractFileName(NomFicDucs)+'"';
   ExecuteSQL(StQQ) ;

   // détermination du Chrono
   StQQ := 'SELECT MAX(PES_CHRONOMESS) AS MAXI FROM ENVOISOCIAL';
   QQ:=OpenSQL(StQQ,TRUE) ;
   if Not QQ.EOF then
      Ordre := QQ.FindField ('MAXI').AsInteger+1
   else
      Ordre := 1;
   Ferme(QQ);

   // création de
   ChargeTOBENVOI ;

   EnregEnvoi.Ordre := Ordre;
   EnregEnvoi.TypeE := 'DUC';
   EnregEnvoi.Millesime := Millesime;
   EnregEnvoi.Periodicite := EnregDucsEdi.Periode;
   EnregEnvoi.DateD := EnregDucsEdi.DateDebut;
   EnregEnvoi.DateF := EnregDucsEdi.DateFin;
   EnregEnvoi.Siret := EnregDucsEdi.Siret;
   EnregEnvoi.Fraction := 'Z';
   EnregEnvoi.Libelle := EnregDucsEdi.LibelleET;
   EnregEnvoi.Size := size;
   EnregEnvoi.NomFic := ExtractFileName(NomFicDucs);
   EnregEnvoi.Statut := '';
   EnregEnvoi.Monnaie := EnregDucsEdi.MonnaieTenue;
   EnregEnvoi.CodAppli := EnregDucsEdi.CodAppliPog; // PT1-1
   EnregEnvoi.ServUniq := EnregDucsEdi.ServUniqPog; // PT1-1
   // d PT1-8
   EnregEnvoi.MtPaye :=  EnregDucsEdi.MtAPayer;
   EnregEnvoi.DucsNeant := EnregDucsEdi.Neant;
   EnregEnvoi.ModReglt := EnregDucsEdi.ModePaiement;
   // f PT1-8
   if (Copy(EnregDucsEDi.Typducs,1,1) = 'A') then
    EnregEnvoi.Inst := 'ZDAC';
   if (Copy(EnregDucsEDi.Typducs,1,1) = 'U') then
    EnregEnvoi.Inst := 'ZDUN';
   if (Copy(EnregDucsEDi.Typducs,1,1) = 'I') then
    EnregEnvoi.Inst := 'ZDIR';

   EnregEnvoi.EmettSoc := EnregDucsEdi.EmettSoc;        // PT21-1

   CreeEnvoi (EnregEnvoi);
   LibereTOBENVOI;
end; { fin MajEnvoiSocialDUCS}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 28/05/2002
Modifié le ... :   /  /
Description .. : Traitement des cas :
Suite ........ : ACOSS : Multi établissements avec paiement groupé
Suite ........ :               autant de  groupes 2 et groupes 3 que de
Suite ........ :               déclarations (Ducs papier)
Suite ........ : IRC : Multi établissements, multi centres payeurs
Suite ........ :          autant de  groupes 2, groupes 3 et groupe 6
Suite ........ :         (paiement) que de  déclarations (Ducs papier)
Suite ........ : UNEDIC : Multi établissements avec paiement groupé (VLU)
Suite ........ :                 Segments FZ sur groupe 1,  1 groupe 2, groupe 3
Suite ........ :                 et autant de groupes 2 supplémentaires que
Suite ........ :                d'établissements secondaires.
Mots clefs ... : PAIE, PGDUCSEDI
*****************************************************************}
// PT16-2 procedure MultiGp2Gp3(EtabPrinc : string;var EnregDucsEdi : TDucsEdi;var EnrDucsEdiGroupee : TDucsEdiGroupee; var DossOK : boolean);
procedure MultiGp2Gp3(EtabPrinc : string;var EnregDucsEdi : TDucsEdi;var EnrDucsEdiGroupee : TDucsEdiGroupee; var DossOK : boolean; var NoLg : integer);
var
   QL : TQuery;
   LesDucsGroupee,TOBIrcFille,TOBIrc : TOB;
begin
// d PT34
    if (TOB_DucsGroupee <> nil) then
       FreeAndNil(TOB_DucsGroupee);
// f PT34
       TOB_DucsGroupee := TOB.Create ('Les ducs paiement groupé', NIL, -1);

       if (copy(EnregDucsEdi.TypDucs,1,1) <> 'U') then
        // ACOSS - IRC
        begin
         QL :=OpenSql ('SELECT PDU_ETABLISSEMENT,PDU_ORGANISME,'+
                       'PDU_DATEDEBUT,PDU_DATEFIN,PDU_DUCSDOSSIER,PDU_NUM'+
                       ' FROM DUCSENTETE '+
                       'WHERE '+
                       '(PDU_ETABLISSEMENT <>"'+ EtabPrinc +'") AND '+
                       '(PDU_ORGANISME ="'+ EnregDucsEdi.Organisme+'") AND '+
                       '(PDU_DATEDEBUT ="'+UsDateTime(EnregDucsEdi.DateDebut)+'") AND '+
                       '(PDU_DATEFIN = "'+UsDateTime(EnregDucsEdi.DateFin)+'") AND '+
                       '(PDU_DUCSDOSSIER = "X")', TRUE);

         TOB_DucsGroupee.LoadDetailDB ('TABLEDUCSENTETE', '','',QL,FALSE,FALSE);

         Ferme(QL);
        end
       else
        // UNEDIC
        begin
           QL :=OpenSql ('SELECT POG_ETABLISSEMENT,POG_ORGANISME,'+
                         'POG_DUCSDOSSIER,POG_PAIEMODE,POG_NUMINTERNE'+
                         ' FROM ORGANISMEPAIE '+
                         'WHERE '+
                         '(POG_ETABLISSEMENT <>"'+ EtabPrinc +'") AND '+
                         '(POG_ORGANISME ="'+ EnregDucsEdi.Organisme+'") AND '+
                         '(POG_DUCSDOSSIER = "X")', TRUE);

           TOB_DucsGroupee.LoadDetailDB ('TABLEORGANISME', '','',QL,FALSE,FALSE);

           Ferme(QL);
        end;
       if (copy(EnregDucsEdi.TypDucs,1,1) <> 'U') then
         TOB_DucsGroupee.Detail.Sort('PDU_NUM')
       else
         TOB_DucsGroupee.Detail.Sort('POG_ETABLISSEMENT');

//if V_PGI.Debug=True then
//       PGVisuUnObjet(TOB_DucsGroupee, '', '');

       if (TOB_DucsGroupee <> NIL) then
          begin
           if (copy(EnregDucsEdi.TypDucs,1,1) = 'I') then
             // IRC
             begin
              TOBIrc := Tob.Create('Les infos paiement IRC', NIL,-1);
             end;
           LesDucsGroupee := TOB_DucsGroupee.FindFirst([''],[''],TRUE);
// PT15-7           EnregDucsEdi.CentrePayeur := 0;
           EnregDucsEdi.NoCentrePayeur := 1;   // PT20
            While LesDucsGroupee <> NIL  do
             begin
              if (copy(EnregDucsEdi.TypDucs,1,1) <> 'U') then
                InitVarGp2(EnregDucsEdi, EnrDucsEdiGroupee, LesDucsGroupee)
              else
                InitVarGp2UNEDIC(EnregDucsEdi, EnrDucsEdiGroupee, LesDucsGroupee);

              if (copy(EnregDucsEdi.TypDucs,1,1) = 'I') then
               // IRC
               begin
                TOBIrcFille := Tob.Create('', TOBIrc,-1);
// d PT1-7
                TOBIrcFille.AddChampSup('ETAB',FALSE);
                TOBIrcFille.AddChampSup('LIBELLEET',FALSE);
// f PT1-7
                TOBIrcFille.AddChampSup('MODEPAIEMENT',FALSE);
                TOBIrcFille.AddChampSup('MTAPAYER',FALSE);
                TOBIrcFille.AddChampSup('DATEFIN',FALSE);
                TOBIrcFille.AddChampSup('ABREGE',FALSE);
                TOBIrcFille.AddChampSup('NOINTERNE',FALSE);
                TOBIrcFille.AddChampSup('PAIEMENT',FALSE);

// d PT1-7
                TOBIrcFille.PutValue('ETAB',EnrDucsEdiGroupee.Etab);
                TOBIrcFille.PutValue('LIBELLEET',EnrDucsEdiGroupee.LibelleET);
// f PT1-7
                TOBIrcFille.PutValue('MODEPAIEMENT',EnrDucsEdiGroupee.ModePaiement);
                TOBIrcFille.PutValue('MTAPAYER',EnrDucsEdiGroupee.MtAPayer);
                TOBIrcFille.PutValue('DATEFIN',EnregDucsEdi.DateFin);
                TOBIrcFille.PutValue('ABREGE',EnrDucsEdiGroupee.Abrege);
                TOBIrcFille.PutValue('NOINTERNE',EnrDucsEdiGroupee.NumeroInterne);
                TOBIrcFille.PutValue('PAIEMENT',EnrDucsEdiGroupee.Paiement);
               if (copy(EnregDucsEdi.TypDucs,1,1) = 'I') and
                  (EnregDucsEdi.PaiementGroupe = False) then
                  // On recalcule le Mt à payer total car dans le cas IRC Multi Centre
                  // Payeur le mt de chaque Groupe 2 correspond au montant à payer
                  // par chaque centre payeur (établissement)
                  EnregDucsEdi.MtAPayer := EnregDucsEdi.MtAPayer - EnrDucsEdiGroupee.MtAPayer;
               end;

              if (copy(EnregDucsEdi.TypDucs,1,1) <> 'U') then
               // Acoss et Irc
               begin
                InitCotisation (EnregDucsEdi, EnrDucsEdiGroupee,'G');
                if (ControlCotisation (EnregDucsEdi, EnrDucsEdiGroupee,'G') = false) then
                  begin
                   DossOK :=  false;
                   exit;
                  end;
               end;
              EnregDucsEdi.NoCentrePayeur := EnregDucsEdi.NoCentrePayeur + 1; // PT20
              Groupe2Declaration (EnregDucsEdi, EnrDucsEdiGroupee,'G');
              if (copy(EnregDucsEdi.TypDucs,1,1) = 'A') then
               // Acoss
               begin
                Groupe3ACOSS (EnregDucsEdi, EnrDucsEdiGroupee,'G');
                if (EnrDucsEdiGroupee.NbTransport > 1) then
                 Groupe4_5ACOSS(EnregDucsEdi, EnrDucsEdiGroupee,'G');
               end;
              if (copy(EnregDucsEdi.TypDucs,1,1) = 'I') then
               // IRC
// PT16-2       Groupe3IRC (EnregDucsEdi, EnrDucsEdiGroupee,'G');
                Groupe3IRC (EnregDucsEdi, EnrDucsEdiGroupee,'G', NoLg);

              LesDucsGroupee := TOB_DucsGroupee.FindNext ([''],[''],TRUE);
             end;

// pour IRC Multi centre payeur
// une findéclaration + boucle paiement edi
// faire une TOB avec les éléments nécessaire à fin déclaration et paiement edi
           if (copy(EnregDucsEdi.TypDucs,1,1) = 'I') and
               (EnregDucsEdi.PaiementGroupe = False) then
            begin
              FinDeclaration(EnregDucsEdi,EnrDucsEdiGroupee);
// d PT7-3               PaiementEdi(EnregDucsEdi,EnrDucsEdiGroupee,'');
              PaiementEdi(EnregDucsEdi,EnrDucsEdiGroupee,'', DossOK);
              if (DossOK = False) then exit;
// f PT7-3
              TOBIrcFille := TOBIrc.FindFirst([''],[''],TRUE);
              While   TOBIrcFille <> NIL do
                begin
// d PT1-7
                  EnrDucsEdiGroupee.Etab := TOBIrcFille.GetValue('ETAB');
                  EnrDucsEdiGroupee.LibelleET := TOBIrcFille.GetValue('LIBELLEET');
                  EnrDucsEdiGroupee.ModePaiement := TOBIrcFille.GetValue('MODEPAIEMENT');

// f PT1-7
                  EnrDucsEdiGroupee.MtAPayer := TOBIrcFille.GetValue('MTAPAYER');
                  EnrDucsEdiGroupee.Abrege := TOBIrcFille.GetValue('ABREGE');
                  EnrDucsEdiGroupee.NumeroInterne := TOBIrcFille.GetValue('NOINTERNE');
                  EnrDucsEdiGroupee.Paiement := TOBIrcFille.GetValue('PAIEMENT');

                  if (EnrDucsEdiGroupee.ModePaiement <> '') then
// d PT7-3                       PaiementEdi(EnregDucsEdi,EnrDucsEdiGroupee,'G');
                       PaiementEdi(EnregDucsEdi,EnrDucsEdiGroupee,'G', DossOK);
                       if (DossOK = False) then exit;
// f PT7-3
                  TOBIrcFille := TOBIrc.FindNext ([''],[''],TRUE);
                end;
              TOBIrc.Free; TOBIrc := NIL;

            end;
          end;
end;
// fin  MultiGp2Gp3

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 28/05/2002
Modifié le ... :   /  /    
Description .. : Initialisation des variables concernant les établissements 
Suite ........ : secondaires nécessaires au traitements Multi établissemens 
Suite ........ : tels que décrits dans MultiGp2GP3
Mots clefs ... : PAIE, PGDUCSEDI
*****************************************************************}
procedure InitVarGp2(var EnregDucsEdi : TDucsEdi; var EnrDucsEdiGroupee : TDucsEdiGroupee;LesDucsGroupee : TOB);
var
   QL : TQuery;
   StQL : string;
begin

  EnrDucsEdiGroupee.Etab := LesDucsGroupee.getValue('PDU_ETABLISSEMENT');
  EnrDucsEdiGroupee.NoDucs := LesDucsGroupee.getValue('PDU_NUM');

  // ETABLISS
  QL := OpenSQL('SELECT ET_LIBELLE,ET_ADRESSE1,ET_ADRESSE2,ET_ADRESSE3,'+
                'ET_VILLE,ET_CODEPOSTAL,ET_TELEPHONE,ET_FAX,ET_JURIDIQUE '+
                'FROM ETABLISS WHERE '+
                'ET_ETABLISSEMENT = "'+EnrDucsEdiGroupee.Etab+'" ',True) ;
  if Not QL.EOF then
    begin
      EnrDucsEdiGroupee.LibelleET:=QL.FindField('ET_LIBELLE').AsString;

      if (QL.FindField('ET_ADRESSE1').AsString <> '') then
         EnrDucsEdiGroupee.Adresse1ET:=QL.FindField('ET_ADRESSE1').AsString;
      if (QL.FindField('ET_ADRESSE2').AsString <> '') then
         if (EnrDucsEdiGroupee.Adresse1ET  <> '') then
           EnrDucsEdiGroupee.Adresse2ET:=QL.FindField('ET_ADRESSE2').AsString
         else
           EnrDucsEdiGroupee.Adresse1ET:=QL.FindField('ET_ADRESSE2').AsString;
      if (QL.FindField('ET_ADRESSE3').AsString <> '') then
         if (EnregDucsEdi.Adresse1ET  = '') then
           EnrDucsEdiGroupee.Adresse1ET:=QL.FindField('ET_ADRESSE3').AsString
         else
      if (EnrDucsEdiGroupee.Adresse2ET  = '') then
         EnrDucsEdiGroupee.Adresse2ET:=QL.FindField('ET_ADRESSE3').AsString;  // PT9-1

      EnrDucsEdiGroupee.VilleET:=QL.FindField('ET_VILLE').AsString;
      EnrDucsEdiGroupee.CpET:=QL.FindField('ET_CODEPOSTAL').AsString;
    end ;
  Ferme(QL) ;

  // DUCSENTETE
  QL := OpenSql('SELECT PDU_SIRET,PDU_MONNAIETENUE,PDU_NBSALFPE,'+
                'PDU_TOTHOMMES,PDU_TOTFEMMES,PDU_ACOMPTES,'+
                'PDU_REGULARISATION,PDU_TOTAPPRENTI, '+
                'PDU_ABREGEPERIODE,PDU_NUMERO,PDU_APE,PDU_PAIEMENT,PDU_DATEREGLEMENT, '+ // PT13
                'PDU_CENTREPAYEUR '+ // PT20
                ' FROM DUCSENTETE '+
                ' WHERE '+
                'PDU_ETABLISSEMENT = "'+EnrDucsEdiGroupee.Etab+'" AND '+
                'PDU_ORGANISME = "'+EnregDucsEdi.Organisme+'" AND '+
                'PDU_DATEDEBUT = "'+UsDateTime(EnregDucsEdi.DateDebut)+'" AND '+
                'PDU_DATEFIN = "'+UsDateTime(EnregDucsEdi.DateFin)+'" AND '+
                'PDU_NUM = '+IntToStr(EnrDucsEdiGroupee.NoDucs)+'',True) ;        // DB2
  if Not QL.EOF then
    begin
      EnrDucsEdiGroupee.Siret := QL.FindField('PDU_SIRET').AsString;
      EnrDucsEdiGroupee.Abrege := QL.FindField('PDU_ABREGEPERIODE').AsString;
      EnrDucsEdiGroupee.MonnaieTenue := QL.FindField('PDU_MONNAIETENUE').AsString;
      EnrDucsEdiGroupee.NbSalFpe := QL.FindField('PDU_NBSALFPE').AsFloat;
      EnrDucsEdiGroupee.TotHommes := QL.FindField('PDU_TOTHOMMES').AsFloat;
      EnrDucsEdiGroupee.TotFemmes := QL.FindField('PDU_TOTFEMMES').AsFloat;
      EnrDucsEdiGroupee.Regularisation:=QL.FindField('PDU_REGULARISATION').AsFloat;
      EnrDucsEdiGroupee.Acompte:=QL.FindField('PDU_ACOMPTES').AsFloat;
      EnrDucsEdiGroupee.TotApprenti := QL.FindField('PDU_TOTAPPRENTI').AsFloat;
      EnrDucsEdiGroupee.NumeroInterne := QL.FindField('PDU_NUMERO').AsString;
      EnrDucsEdiGroupee.ApePdu := QL.FindField('PDU_APE').AsString;
      EnrDucsEdiGroupee.Paiement:=QL.FindField('PDU_PAIEMENT').AsDateTime;
      EnrDucsEdiGroupee.Reglement:=QL.FindField('PDU_DATEREGLEMENT').AsDateTime; //PT13
      // d PT20
      if (QL.FindField('PDU_CENTREPAYEUR').AsString <> '') then
        EnrDucsEdiGroupee.CentrePayeur:=QL.FindField('PDU_CENTREPAYEUR').AsString;
      //f PT20
    end;
  ferme (QL);


  QL := OpenSql('SELECT POG_PAIEMODE FROM ORGANISMEPAIE WHERE '+
                'POG_ETABLISSEMENT = "'+EnrDucsEdiGroupee.Etab+'" AND '+
                'POG_ORGANISME = "'+EnregDucsEdi.Organisme+'"',True) ;
   if Not QL.EOF then
    begin
      EnrDucsEdiGroupee.ModePaiement := QL.FindField('POG_PAIEMODE').AsString;
    end;
  ferme (QL);


  // DUCSDETAIL
  StQL := 'PDD_ETABLISSEMENT = "'+EnrDucsEdiGroupee.Etab+'" AND '+
          'PDD_ORGANISME = "'+EnregDucsEdi.Organisme+'" AND '+
          'PDD_DATEDEBUT = "'+UsDateTime(EnregDucsEdi.DateDebut)+'" AND '+
          'PDD_DATEFIN = "'+UsDateTime(EnregDucsEdi.DateFin)+'" AND '+
          'PDD_NUM = '+IntToStr(EnrDucsEdiGroupee.NoDucs)+' AND '+        // DB2
          'PDD_TYPECOTISATION <> "S"';

  QL := OpenSQL('SELECT SUM(PDD_MTCOTISAT) AS DECLAR FROM DUCSDETAIL WHERE '+StQL,True) ;

  EnrDucsEdiGroupee.MtDeclare := 0.0; // PT8-1
  IF NOT QL.EOF then // PortageCWAS
   EnrDucsEdiGroupee.MtDeclare := QL.FindField('DECLAR').asFloat; 
  EnregDucsEdi.MtTotal :=  EnregDucsEdi.MtTotal +  EnrDucsEdiGroupee.MtDeclare;

  EnrDucsEdiGroupee.MtAPayer := 0.0; // PT8-1
{ PT8-2 EnrDucsEdiGroupee.MtAPayer := EnrDucsEdiGroupee.MtAPayer+
                              EnrDucsEdiGroupee.MtDeclare-
                              EnrDucsEdiGroupee.Regularisation-
                              EnrDucsEdiGroupee.Acompte; }
  EnrDucsEdiGroupee.MtAPayer := EnrDucsEdiGroupee.MtAPayer+
                              EnrDucsEdiGroupee.MtDeclare+
                              EnrDucsEdiGroupee.Regularisation-
                              EnrDucsEdiGroupee.Acompte;

  Ferme (QL);

  StQL := '';
  StQL := 'PDD_ETABLISSEMENT = "'+EnrDucsEdiGroupee.Etab+'" AND '+
          'PDD_ORGANISME = "'+EnregDucsEdi.Organisme+'" AND '+
          'PDD_DATEDEBUT = "'+UsDateTime(EnregDucsEdi.DateDebut)+'" AND '+
          'PDD_DATEFIN = "'+UsDateTime(EnregDucsEdi.DateFin)+'" AND '+
          'PDD_NUM = '+IntToStr(EnrDucsEdiGroupee.NoDucs)+' AND '+        // DB2
          'PDD_TYPECOTISATION <> "S" AND PDD_TYPECOTISATION <> "I" '+
          'AND PDD_CODIFICATION LIKE "1%900P"';
  QL := OpenSQL('SELECT COUNT(DISTINCT(PDD_TAUXCOTISATION)) AS NBTRANSPORT, '+
                'SUM(PDD_MTCOTISAT) AS MTTRANSPORT FROM DUCSDETAIL WHERE '+StQL,True) ;
  EnrDucsEdiGroupee.MtTransport:=0.0; // PT8-1
  IF NOT QL.EOF then // PortageCWAS
  begin
    EnrDucsEdiGroupee.MtTransport:= QL.FindField('MTTRANSPORT').asFloat;
    EnrDucsEdiGroupee.NbTransport:= QL.FindField('NBTRANSPORT').AsInteger;
  end;
  Ferme(QL);

  EnrDucsEdiGroupee.ErrTransport := False;
  if (copy(EnregDucsEdi.TypDucs,1,1) = 'A') and (EnrDucsEdiGroupee.NbTransport > 1) then
     // Acoss - Création TOB_Transport
     begin
      QL := OpenSQL('SELECT Count (*) FROM DUCSDETAIL WHERE (PDD_COMURBAINE="" '+
                    'OR PDD_COMURBAINE IS NULL OR PDD_COMURBAINE="    ") and '+
                    StQL,TRUE);
      if NOT QL.EOF then
        begin
          if (QL.Fields[0].AsInteger <> 0) then
             EnrDucsEdiGroupee.ErrTransport :=  True;
        end;
      ferme (QL);
     end;

     QL := OpenSQL('SELECT * FROM DUCSDETAIL WHERE '+StQL,True);

      TOB_Transport := TOB.Create ('Les lignes de tansport', NIL, -1);

      TOB_Transport.LoadDetailDB ('DUCSDETAIL', '','',QL,FALSE,FALSE);
      Ferme(QL);
      TOB_Transport.Detail.Sort('PDD_ETABLISSEMENT;PDD_ORGANISME;PDD_DATEDEBUT;'+
                                'PDD_DATEFIN;PDD_INSTITUTION;PDD_CODIFICATION;PDD_DATECHGTTAUX');
{  EnregDucsEdi.MtAPayer := EnregDucsEdi.MtAPayer+
                           EnrDucsEdiGroupee.MtDeclare-
                           EnrDucsEdiGroupee.Regularisation-
                           EnrDucsEdiGroupee.Acompte; }
  EnregDucsEdi.MtAPayer := EnregDucsEdi.MtAPayer+
                           EnrDucsEdiGroupee.MtDeclare+
                           EnrDucsEdiGroupee.Regularisation-
                           EnrDucsEdiGroupee.Acompte;

  EnregDucsEdi.MtTotRegul :=  EnregDucsEdi.MtTotRegul + EnrDucsEdiGroupee.Regularisation;
  EnregDucsEdi.MtTotAcpte :=  EnregDucsEdi.MtTotAcpte + EnrDucsEdiGroupee.Acompte;


//if V_PGI.Debug=True then
//      PGVisuUnObjet(TOB_Transport, '', '');
end;
// fin InitVarGp2
{***********A.G.L.***********************************************
Auteur  ...... : PAIE : MF
Créé le ...... : 22/05/2002
Modifié le ... :   /  /    
Description .. : Initialisation des variables concernant les
Suite ........ : établissements secondaires pour l'UNEDIC
Mots clefs ... : 
*****************************************************************}
procedure InitVarGp2UNEDIC(var EnregDucsEdi : TDucsEdi; var EnrDucsEdiGroupee : TDucsEdiGroupee;LesDucsGroupee : TOB);
var
   QL : TQuery;
   StQL : string;
begin

  EnrDucsEdiGroupee.Etab := LesDucsGroupee.getValue('POG_ETABLISSEMENT');
  EnrDucsEdiGroupee.NumeroInterne := LesDucsGroupee.getValue('POG_NUMINTERNE');
  EnrDucsEdiGroupee.ModePaiement := '';
  EnrDucsEdiGroupee.Abrege := EnregDucsEdi.Abrege;
  EnrDucsEdiGroupee.MtDeclare  :=0.0; // PT8-1
  EnrDucsEdiGroupee.MtTransport :=0.0; // PT8-1
  EnrDucsEdiGroupee.NbTransport :=0;
  EnrDucsEdiGroupee.Regularisation :=0.0; // PT8-1
  EnrDucsEdiGroupee.Acompte :=0.0; // PT8-1
  EnrDucsEdiGroupee.MtAPayer :=0.0; // PT8-1
  EnrDucsEdiGroupee.ErrTransport := False;
  EnrDucsEdiGroupee.ErrCotisqual := False;
  EnrDucsEdiGroupee.ErrInstitution := False;
  EnrDucsEdiGroupee.ErrCondition := False;
  if VH_Paie.PGTenueEuro=FALSE then
    EnrDucsEdiGroupee.MonnaieTenue := 'FRF'          // Franc
  else
    EnrDucsEdiGroupee.MonnaieTenue := 'EUR';         // Euro

  EnrDucsEdiGroupee.Paiement := Idate1900;

  // ETABLISS
  QL := OpenSQL('SELECT ET_LIBELLE,ET_ADRESSE1,ET_ADRESSE2,ET_ADRESSE3,'+
                'ET_VILLE,ET_CODEPOSTAL,ET_TELEPHONE,ET_FAX,ET_JURIDIQUE, '+
                'ET_SIRET, ET_APE '+
                'FROM ETABLISS WHERE '+
                'ET_ETABLISSEMENT = "'+EnrDucsEdiGroupee.Etab+'" ',True) ;
  if Not QL.EOF then
    begin
      EnrDucsEdiGroupee.LibelleET:=QL.FindField('ET_LIBELLE').AsString;

      if (QL.FindField('ET_ADRESSE1').AsString <> '') then
         EnrDucsEdiGroupee.Adresse1ET:=QL.FindField('ET_ADRESSE1').AsString;
      if (QL.FindField('ET_ADRESSE2').AsString <> '') then
         if (EnrDucsEdiGroupee.Adresse1ET  <> '') then
           EnrDucsEdiGroupee.Adresse2ET:=QL.FindField('ET_ADRESSE2').AsString
         else
           EnrDucsEdiGroupee.Adresse1ET:=QL.FindField('ET_ADRESSE2').AsString;
      if (QL.FindField('ET_ADRESSE3').AsString <> '') then
         if (EnregDucsEdi.Adresse1ET  = '') then
           EnrDucsEdiGroupee.Adresse1ET:=QL.FindField('ET_ADRESSE3').AsString
         else
      if (EnrDucsEdiGroupee.Adresse2ET  = '') then
         EnrDucsEdiGroupee.Adresse2ET:=QL.FindField('ET_ADRESSE3').AsString;

      EnrDucsEdiGroupee.VilleET:=QL.FindField('ET_VILLE').AsString;
      EnrDucsEdiGroupee.CpET:=QL.FindField('ET_CODEPOSTAL').AsString;
      EnrDucsEdiGroupee.Siret:=QL.FindField('ET_SIRET').AsString;
      EnrDucsEdiGroupee.ApePdu:=QL.FindField('ET_APE').AsString;
    end ;
  Ferme(QL) ;

  // Calcul des Effectifs
  StQL := '';
  StQL :='PSA_ETABLISSEMENT="'+EnrDucsEdiGroupee.Etab+'" AND ';

  StQL := StQL+'(PSA_DATESORTIE="'+UsDateTime(IDate1900)+'"'+
               ' OR  PSA_DATESORTIE IS NULL '+
               ' OR PSA_DATESORTIE>="'+UsDateTime(EnregDucsEdi.DateDebut)+'")'+
               ' AND PSA_DATEENTREE<="'+UsDateTime(EnregDucsEdi.DateFin)+'"';

 // nbre d'hommes
 QL := OpenSQL('SELECT Count(*) FROM SALARIES WHERE '+StQL+
               ' AND PSA_SEXE="M"',True) ;
 IF NOT QL.EOF then // PortageCWAS
  EnrDucsEdiGroupee.TotHommes := QL.Fields[0].AsInteger ;
 Ferme(QL) ;

 // nbre de femmes
 QL := OpenSQL('SELECT Count(*) FROM SALARIES WHERE '+StQL+
               ' AND PSA_SEXE="F"',True) ;
 IF NOT QL.EOF then // PortageCWAS
  EnrDucsEdiGroupee.TotFemmes := QL.Fields[0].AsInteger ;
 Ferme(QL) ;

 // Nbre d'apprentis
 QL := OpenSQL('SELECT Count(*) FROM SALARIES WHERE '+StQL+
               ' AND PSA_CATDADS="003"',True) ;
 IF NOT QL.EOF then //  PortageCWAS
  EnrDucsEdiGroupee.TotApprenti := QL.Fields[0].AsInteger ;
 Ferme(QL) ;

 // Nbre de salariés rémunérés pour la période
 StQL := '';
 StQL :='PHB_ETABLISSEMENT="'+EnrDucsEdiGroupee.Etab+'" AND ';
 StQL := StQL+ 'PHB_ORGANISME="'+EnregDucsEdi.Organisme+'" ';

 StQL := StQL+'AND '+
              'PHB_DATEDEBUT >="'+UsDateTime(EnregDucsEdi.DateDebut)+'" AND '+
              'PHB_DATEFIN <="'+UsDateTime(EnregDucsEdi.DateFin)+'" AND '+
              '(PHB_BASECOT <> 0 OR PHB_MTSALARIAL <> 0 OR PHB_MTPATRONAL <> 0)';

 QL := OpenSQL('SELECT Count(DISTINCT PHB_SALARIE) AS NOMBRE FROM HISTOBULLETIN WHERE '+
                   StQL,True) ;
 IF NOT QL.EOF then // PortageCWAS
  EnrDucsEdiGroupee.NbSalFpe := QL.FindField('NOMBRE').AsInteger ;
 Ferme (QL);
end;
// fin InitVarGp2UNEDIC
{***********A.G.L.***********************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 13/05/2002
Modifié le ... : 13/05/2002
Description .. : Fonction de calcul et de contrôle de la clé IBAN
Suite ........ : Paramètres : Type de créancier  (2 car)
Suite ........ :              Période            (6 car)
Suite ........ :              référence interne  (19 car)
Suite ........ :              Clé Iban           (2 car) = 00 pour calcul
Suite ........ :                                        <> 00 por contrôle
Mots clefs ... : PAIE, PGDUCSEDI
*****************************************************************}
Function CleIBAN (creancier, periode,reference : string;var cleIBAN : string) : boolean;
var
   car : char;
   RefSociale : string;
   RefSocialeSize,icar : integer;
   Divi : string;
   i,j,IRest,IDivi: integer;
// PT7-1    tampon : string;
   NbNonCar : integer;
begin
 IDivi := 0; // PT7-1 
 // Constitution de la référence sociale
 RefSociale := '';
 result := True;

   for i := 1 to 2 do
    begin
     car := creancier[i];
     if (IsNumeric(car) = false) then
       begin
         icar := Ord(creancier[i])-55;
         RefSociale := RefSociale+IntToStr(icar);
       end
     else
       RefSociale := RefSociale+car;
    end;

   for i := 1 to 6 do
    begin
     car := periode[i];
     if (IsNumeric(car) = false) then
       begin
         icar := Ord(car)-55;
         RefSociale := RefSociale+IntToStr(icar);
       end
     else
      RefSociale := RefSociale+car;
    end;

     reference := UpperCase(reference);  // PT1-6
     NbNonCar := 0;
     for i := 1 to length(reference) do
       begin
         car := reference[i];
         if ((Ord(car)>=65) and (Ord(car)<=90)) or
            ((IsNumeric(car) = true) and
            (car <> '') and (car <> '.') and(car <> ' ')) then // PT1-6
           begin                                       //PT1-6
             if (IsNumeric(car) = false) then
               begin
                 icar := Ord(car)-55;
                 RefSociale := RefSociale+IntToStr(icar);
               end
             else
               RefSociale := RefSociale+car;
           end
         else
          NbNonCar := NbNonCar+1;
       end;
     if ((length(reference)-NbNonCar) < 19) then
      for i:= (length(reference)-NbNonCar)+1 to 19 do
        begin
         car:= '0';
         RefSociale := RefSociale+car;
        end;

   RefSociale :=  RefSociale+ CleIBAN;
   RefSocialeSize :=  length (RefSociale);

  // Division par 97 :
  // On procède par divisions successives sur la chaîne RefSociale
// PT7-1   i := 0;
  j := 1;
  divi := '';
  IRest := 0;
  while (j <= RefSocialeSize) do
    begin
     for i:=j to RefSocialeSize do
       begin
         j := j+1;
         divi := divi+copy(RefSociale,i,1);
         IDivi := StrToInt(divi);
         if (Idivi > 97) then
          begin
            divi := '';
            break;
          end;
          IDivi := 0;
       end;
       Irest := Idivi mod 97;
       if (j <= RefSocialeSize) then
       // Irest est repris pour déterminer la nouvelle chaîne à diviser
       divi := IntToStr(Irest);
    end;
  // Calcul ou contrôle de la clé
  if (StrToInt(CleIBAN) = 0) then
    begin
     if ((98-IRest) > 9) then
        CleIBAN := IntToStr(98-IRest)
     else
        CleIBAN := '0'+ IntToStr(98-IRest);
    end
  else
      if (Irest <> 1) then result := False;

end;
//  fin CleIBAN

//end.
{$ENDIF}
end.
