{***********UNITE*************************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 13/07/2000
Modifié le ... : 04/03/2002
Description .. : Gestion de l'attestation ASSEDIC
Mots clefs ... : PAIE;PGATTESTATION;PGASSEDIC
*****************************************************************}
{
PT1   : 04/09/2001 VG V547 Formatage des éléments de la grille
PT2   : 10/09/2001 VG V547 Ajout d'un bouton pour calcul des données de paie par
                           rapport au dernier jour travaillé . On ne gère plus
                           l'évenement onchange de la date de dernier jour
                           travaillé
PT3   : 18/09/2001 VG V547 Affichage  du titre de la fenêtre
PT4   : 29/10/2001 VG V563 La caisse ASSEDIC n'était pas reprise automatiquement
PT5-1 : 09/11/2001 VG V563 Modification de la tablette PGMONNAIE : 001 => FRF et
                           002 => EUR
PT6   : 14/12/2001 VG V563 Quand un montant d'indemnité est vide alors que la
                           case est cochée, on avait "conversion de type variant
                           incorrecte" . Dans le cas ou le montant est vidé, on
                           décoche la case correspondante .
PT7   : 28/02/2002 VG V571 On ne complète plus à zéro et on n'édite plus les
                           lignes de détail de salaire dont la date est
                           antérieure à la date d'entrée du salarié
                           Fiche de bug N° 412
PT8-1 : 04/03/2002 VG V571 Récupération du nombre de jours ouvrables CP -
                           Fiche qualité N° 10045
PT8-2 : 04/03/2002 VG V571 Pour le solde de tout compte, date de fin devient
                           date du dernier jour travaillé au lieu du dernier
                           jour du mois - Fiche qualité N° 10045
PT9   : 28/05/2002 VG V582 Les montants des indemnités doivent être récupérés
                           uniquement sur la dernière paye
PT10  : 17/06/2002 VG V582 Ajout traitement des indemnités conventionnelles
                           Fiche qualité N° 10090
PT11  : 15/07/2002 VG V585 Avant l'impression de l'attestation, on force la
                           validation de l'enregistrement                            
PT12  : 04/09/2002 VG V585 Modification de la zone qualification du salarié
PT13  : 03/10/2002 VG V585 Traitement dans le cas ou le salarié n'est pas sorti
PT14-1: 13/11/2002 VG V585 La qualification de la fiche salarié ne correspond
                           pas à la qualification ASSEDIC
PT14-2: 13/11/2002 VG V591 On remplace EDIT_QUALIF par PAS_QUALIF
PT15  : 18/02/2003 VG V_42 Ajout de champs pour sauvegarder les données saisies
                           FQ N° 10336
PT16  : 27/02/2003 VG V_42 Dans les tableaux, on réinitialise à '01/01/1900' les
                           dates qui ne sont pas valides au moment de la
                           validation - FQ N° 10529
PT17  : 24/03/2003 VG V_42 Le libellé emploi est tronqué à 18 caractères pour
                           éviter qu'il ne déborde sur la zone Lieu de Travail -
                           FQ N°10587
PT18  : 22/08/2003 VG V_42 On ne reprenait que le code de la forme juridique de
                           l'établissement au lieu de reprendre le libellé
PT19  : 17/09/2003 VG V_42 Reprise du motif de fin de contrat en automatique au
                           niveau de l'authentification
PT20  : 23/09/2003 VG V_42 En Eagl, certaines zones ne sont pas conservées à la
                           validation de l'attestation - FQ N°10807
PT21-1: 01/10/2003 VG V_42 Complément du PT19
PT21-2: 01/10/2003 VG V_42 Pour les salariés affectés au régime MSA, On coche
                           désormais la case Régime Spécial de Sécurité Sociale
PT22  : 23/10/2003 VG V_42 Attestation ASSEDIC E-AGL - FQ N°10940                           
PT23  : 09/02/2004 VG V_50 Cas de l'existance de différentes périodes de préavis
                           FQ N°10867
PT24-1: 03/05/2004 VG V_50 Le cadre 7.3 ne doit être alimenté que lorsque la
                           période du dernier bulletin n'est pas une fin de mois
                           FQ N°11172
PT24-2: 02/06/2004 VG V_50 On exclut du brut l'indemnité CP et préavis
                           FQ N°11172
PT25  : 28/04/2004 SB V_50 FQ 10812 Intégration de la gestion des déclarants
PT26-1: 17/06/2004 VG V_50 La suppression de l'attestation ne supprimait pas les
                           données salariales - FQ N°11227
PT26-2: 17/06/2004 VG V_50 Optimisation
PT27  : 12/08/2004 VG V_50 Reprise des montants contenus dans les cumuls
                           salariés pour la période ou il n'existe pas de paie
                           en parallèle. Ceci afin d'avoir le montant du brut
                           suite à une reprise ISIS II - FQ N°11359
PT28  : 13/08/2004 VG V_50 Attestation ASSEDIC pour un apprenti : alimenter la
                           zone Salaire brut par le cumul 01 Brut - FQ N°11387
PT29  : 29/09/2004 VG V_50 reprendre automatiquement le département - FQ N°11630
PT30  : 03/01/2005 VG V_60 Ne pas permettre de cocher en même temps case
                           "Régime Alsace Moselle" et case "Régime spécial SS"
                           FQ N°11723                           
PT31  : 18/10/2005 VG V_60 Mise en place d'un format de saisie pour les dates
                           dans les grilles - FQ N°12051
PT32  : 27/10/2005 VG V_65 Nouveau format de l'attestation ASSEDIC - FQ N°12632
PT33  : 10/11/2005 VG V_65 Mise à jour des contenus des tableaux sur validation
                           en CWAS - FQ N°12365
PT34  : 09/02/2006 VG V_65 Traitement des préavis non effectué-payé - FQ N°11805
PT35  : 13/02/2006 VG V_65 Modification de l'alimentation de l'ICCP - FQ N°11860
PT36  : 25/04/2006 VG V_65 Alimentation en provenance des contrats de travail
                           FQ N°12302
PT37  : 04/05/2006 VG V_65 Contrôle des éléments saisis dans les grilles de
                           salaire, prime et solde - FQ N°12740
PT38  : 09/05/2006 VG V_65 Ajout du motif de sortie Décès dans PGMOTIFSORTIE
                           entraîne la modification du traitement de
                           l'attestation pour l'enlever de la tablette
                           FQ N°12819
PT39  : 28/06/2006 VG V_70 Complément PT37 pour executer le contrôle sans sortie
                           de cellule
PT40-1: 12/07/2006 VG V_70 Correction PT34 - FQ N°11805
PT40-2: 12/07/2006 VG V_70 Complément PT37 - Le contrôle s'exécutait aussi pour
                           la première ligne
PT41  : 18/08/2006 VG V_70 Complément PT37 - Pas de message pour les montants
                           FQ N°13445
PT42-1: 31/10/2006 VG V_70 Enregistrement en base des nouvelles indemnités
PT42-2: 31/10/2006 VG V_70 Alimentation automatique de PAS_INDTRANS2
                           FQ N°13570                           
PT43  : 24/11/2006 VG V_70 Correction suite à la multi-édition des attestations
                           FQ N°13700
PT44  : 27/04/2007 GGS V_72 FQ 14098 Motif fin de stage equivalent Autre Motif
PT45  : 30/05/2007 GGU V_72 FQ 13631 Modifications pour rendre compatible
                           l'attestation avec l'edition en lot et pour unifier
                           la version CWAS et la version normale
PT46  : 02/07/2007 VG V_72 Ajout "Statut particulier" dans la fiche salarié
                           FQ N°12301
PT47-1: 04/07/2007 VG V_72 Alimentation automatique des primes - FQ N°13068
PT47-2: 04/07/2007 VG V_72 Optimisation
PT48  : 06/09/2007 VG V_80 Suppression du message "Voulez-vous enregistrer les
                           modifications" en cas de suppression de
                           l'enregistrement - FQ N°14374
PT49  : 07/09/2007 VG V_80 Corrections suite à passage par LanceEtatTob
                           FQ N°14746+FQ N°14754
PT50  : 02/10/2007 VG V_80 Rechercher d'abord les infos dans le contrat de
                           travail puis dans la fiche du salarié - FQ N°12301
PT51  : 09/10/2008 JS FQ N°15741 Ajout indemnité spécifique de rupture conventionnelle
}
unit UTOMAttestations;

interface
uses
         {$IFDEF VER150}
         Variants,
         {$ENDIF}
         UTOM,Classes,Hctrls,StdCtrls,SysUtils,UTOB,
{$IFNDEF EAGLCLIENT}
         DBCtrls,HDB,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
         DB,Fiche,EdtREtat,
{$ELSE}
         eFiche,UtileAgl,
{$ENDIF}
         UTofPG_MulAssedic,HTB97,ComCtrls,
         HEnt1,EntPaie,HMsgBox,P5Util,buttons,PgOutils2 ;
                              
Type
      TOM_Attestations = Class (TOM)
      public
        procedure OnArgument (stArgument : String ) ; override ;
        procedure OnNewRecord ; override ;
        procedure OnLoadRecord ; override ;
        procedure OnUpdateRecord  ; override ;
        procedure OnAfterUpdateRecord ; override;
        procedure OnDeleteRecord  ; override ;
        procedure OnAfterDeleteRecord; override ;
        procedure OnChangeField (F : TField) ; override ;
        procedure OnClose ; override ;
        procedure Imprimer(Sender: TObject);

      private
{$IFNDEF EAGLCLIENT}
        AgeSalarieFNE, Alsace, ChomTot, DDTEFP, EmpPub, FNGSRED : TDBCheckBox;
        FNGS, FNGS2, IndApprenti, IndAutre, IndAvion, IndCDD : TDBCheckBox;
        IndClient, IndCNE, IndFinMiss, IndJournal, IndLeg : TDBCheckBox;
        IndRetraite, IndSinistre, IndSpcfiq, IndSpec, LienParente : TDBCheckBox;
        PlanSoc, Preavis1, Preavis2, Preavis3, RegimeSS : TDBCheckBox;

        Agirc, Arrco, Autres, HorEtab, HorEtabann, HorSal, HorSalann : THDBEdit;
        MotRupCont : THDBEdit;
{$ELSE}
        AgeSalarieFNE, Alsace, ChomTot, DDTEFP, EmpPub, FNGSRED : TCheckBox;
        FNGS, FNGS2, IndApprenti, IndAutre, IndAvion, IndCDD : TCheckBox;
        IndClient, IndCNE, IndFinMiss, IndJournal, IndLeg : TCheckBox;
        IndRetraite, IndSinistre, IndSpcfiq, IndSpec, LienParente : TCheckBox;
        PlanSoc, Preavis1, Preavis2, Preavis3, RegimeSS : TCheckBox;

        Agirc, Arrco, Autres, HorEtab, HorEtabann, HorSal, HorSalann : THEdit;
        MotRupCont : THEdit;
{$ENDIF}
        DerJour : TDateTime;

        CatDADS, UsDerJour : String;

        Salaires, Primes, Solde : THGRID;

        Leg, CDD, CNE, FinMiss, Retraite, Sinistre, Spec, Spcfiq : Double;
        Journal, Client, Avion, Apprenti, Autre : Double;
        Conv, Transac, CP, Preavis, Total : Double;
        Indrupcon : Double; //PT51
        Indrupconv : TCheckBox; //PT51
        Indrupconp: THEdit; //PT51

        Imprim, Valid : TToolBarButton97;

        OnFerme, Suppr : Boolean;

        DerniereCreate, ColEnCours, LigEnCours : integer;

        NbreMois : Word;

        procedure ChargeZones ();
        procedure ChargeIndemnites ();
        procedure ChargeSalaires (ModifJour :Boolean);
        procedure ChargePrimes (ModifJour :Boolean);
        procedure ChargeSolde (ModifJour :Boolean);
        procedure ChargeCaisses ();
        procedure EmpPubClick (Sender : TObject);
        procedure PreavisClick (Sender : TObject);
        procedure RegimeSSClick (Sender : TObject);
        procedure AlsaceClick (Sender : TObject);
        procedure PubTypeChange ();
        procedure LienParenteClick (Sender: TObject);
        procedure PlanSocClick (Sender: TObject);
        procedure AgeSalarieFNEClick (Sender: TObject);
        procedure ChomTotClick (Sender: TObject);
        procedure FNGSClick (Sender: TObject);
        procedure IndemniteClick (Sender: TObject);
        procedure DernierJourChange (Sender: TObject);
        procedure QualiteEmpChange ();
        procedure ContratPartChange ();
        procedure FonctionChange ();
        procedure MotifChange();
        procedure DifHorChange ();
        procedure AgircChange (Sender: TObject);
        procedure ArrcoChange (Sender: TObject);
        procedure AutresChange (Sender: TObject);
        procedure HorExit (Sender: TObject);
        procedure UpdateTable();
        procedure CellExit (Sender: TObject; var ACol,ARow: Integer;
                            var Cancel: Boolean);
        procedure GridExit (Sender: TObject);
        procedure CellEnter (Sender: TObject; var ACol, ARow: Longint;
                             var Cancel: Boolean);
        procedure Validation (Sender: TObject);
        procedure AffichDeclarant (Sender: TObject);
        procedure IndemniteExit (Sender: TObject); // PT51
(*Mise en commentaire PT45
{$IFDEF EAGLCLIENT}
        procedure ESalaires();
        procedure EPrimes();
        procedure ESolde();
{$ENDIF}
*)
     END ;

Procedure GetInfoContratTravail(Salarie : String; DateFinContrat : TDateTime; var MotifSortie, TypeContrat, TypeContratPart, FonctionSal : String);


var PGAttesEtab : string;
implementation

Uses
  Grids;  // PT45

Procedure GetInfoContratTravail(Salarie : String; DateFinContrat : TDateTime; var MotifSortie, TypeContrat, TypeContratPart, FonctionSal : String);
var
StContrat : String;
QRechContrat : TQuery;
begin
StContrat:= 'SELECT PCI_MOTIFSORTIE, PCI_TYPECONTRAT, PCI_FONCTIONSAL'+
            ' FROM CONTRATTRAVAIL WHERE'+
            ' PCI_SALARIE="'+Salarie+'" AND'+
            ' PCI_FINCONTRAT<="'+UsDateTime (DateFinContrat)+'"'+
            ' ORDER BY PCI_FINCONTRAT DESC';
QRechContrat:= OpenSql (StContrat,TRUE);
if not QRechContrat.EOF then
   begin
   MotifSortie:= QRechContrat.FindField ('PCI_MOTIFSORTIE').Asstring;
   TypeContrat:= QRechContrat.FindField ('PCI_TYPECONTRAT').Asstring;
   FonctionSal:= QRechContrat.FindField ('PCI_FONCTIONSAL').Asstring;  //PT50
   if (TypeContrat='CAA') then
      begin
      TypeContrat:= '002';
      TypeContratPart:= '1';
      end
   else
   if (TypeContrat='CAC') then
      begin
      TypeContrat:= '002';
      TypeContratPart:= '1';
      end
   else
   if (TypeContrat='CAE') then
      begin
      TypeContrat:= '002';
      TypeContratPart:= '6';
      end
   else
   if (TypeContrat='CAP') then
      begin
      TypeContrat:= '002';
      TypeContratPart:= '1';
      end
   else
   if (TypeContrat='CCD') then
      TypeContrat:= '002'
   else
   if (TypeContrat='CDI') then
      TypeContrat:= '001'
   else
   if (TypeContrat='CEJ') then
      begin
      TypeContrat:= '002';
      TypeContratPart:= '6';
      end
   else
   if (TypeContrat='CES') then
      begin
      TypeContrat:= '002';
      TypeContratPart:= '5';
      end
   else
   if (TypeContrat='CTT') then
      begin
      TypeContrat:= '002';
      TypeContratPart:= '6';
      end
   else
   if (TypeContrat='SC') then
      TypeContrat:= '001';
   end;
Ferme (QRechContrat);
end;



{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 13/07/2000
Modifié le ... :   /  /
Description .. : OnArgument
Mots clefs ... : PAIE;PGATTESTATION;PGASSEDIC
*****************************************************************}
procedure TOM_Attestations.OnArgument(stArgument: String);
var
Pages : TPageControl;
ValidJour : TBitBtn;
St,Etab   : String;
Edit      : THEdit;
begin
Inherited ;

if PGAttesSalarie = '' then
begin
  PGAttesSalarie:= Trim(ReadTokenPipe(stArgument, ';'));
end;


// Positionnement sur le premier onglet
Pages:= TPageControl(GetControl('PAGES'));
if Pages<>nil then
   Pages.ActivePageIndex:=0;

// Gestion des grilles
Salaires:= THGRID (GetControl ('GRID_SALAIRES'));
if Salaires <> NIL then
   begin
// valeurs des colonnes numériques cadrées à droite de la cellule
   Salaires.ColAligns[0]:= taCenter;
   Salaires.ColAligns[1]:= taCenter;
   Salaires.ColAligns[2]:= taCenter;
   Salaires.ColAligns[3]:= taRightJustify;
   Salaires.ColAligns[4]:= taRightJustify;
   Salaires.ColAligns[5]:= taRightJustify;
   Salaires.ColAligns[6]:= taRightJustify;
   Salaires.HideSelectedWhenInactive:= true;
   Salaires.OnCellExit:=CellExit;
   Salaires.OnExit:= GridExit;
   Salaires.OnCellEnter:= CellEnter;
   Salaires.ColTypes[0]:= 'D';
   Salaires.colformats[0]:= ShortDateFormat;
   Salaires.ColTypes[1]:= 'D';
   Salaires.colformats[1]:= ShortDateFormat;
   Salaires.ColTypes[2]:= 'D';
   Salaires.colformats[2]:= ShortDateFormat;
   end;
Primes:= THGRID (GetControl ('GRID_PRIMES'));
if Primes <> NIL then
   begin
   Primes.ColAligns[0]:= taCenter;
   Primes.ColAligns[1]:= taCenter;
   Primes.ColAligns[3]:= taCenter;
   Primes.ColAligns[4]:= taRightJustify;
   Primes.HideSelectedWhenInactive:= true;
   Primes.OnCellExit:= CellExit;
   Primes.OnExit:= GridExit;
   Primes.OnCellEnter:= CellEnter;
   Primes.ColTypes[0]:= 'D';
   Primes.colformats[0]:= ShortDateFormat;
   Primes.ColTypes[1]:= 'D';
   Primes.colformats[1]:= ShortDateFormat;
   Primes.ColTypes[3]:= 'D';
   Primes.colformats[3]:= ShortDateFormat;
   end;
Solde:= THGRID (GetControl ('GRID_SOLDE'));
if Solde <> NIL then
   begin
   Solde.ColAligns[0]:= taCenter;
   Solde.ColAligns[1]:= taCenter;
   Solde.ColAligns[2]:= taCenter;
   Solde.ColAligns[3]:= taRightJustify;
   Solde.ColAligns[4]:= taRightJustify;
   Solde.ColAligns[5]:= taRightJustify;
   Solde.HideSelectedWhenInactive:= true;
   Solde.OnCellExit:= CellExit;
   Solde.OnExit:= GridExit;
   Solde.OnCellEnter:= CellEnter;
   Solde.Enabled:= True;
   Solde.ColTypes[0]:= 'D';
   Solde.colformats[0]:= ShortDateFormat;
   Solde.ColTypes[1]:= 'D';
   Solde.colformats[1]:= ShortDateFormat;
   Solde.ColTypes[2]:= 'D';
   Solde.colformats[2]:= ShortDateFormat;
   end;

// Gestion du bouton
ValidJour:= TBitBtn(GetControl('BVALIDJOUR'));
if ValidJour<>nil then
   ValidJour.OnClick:= DernierJourChange;

// Gestion du navigateur
Valid:= TToolBarButton97(GetControl('BValider'));
if Valid<>nil then
   Valid.OnClick:= Validation;

// Gestion de l'impression
Imprim:= TToolBarButton97(GetControl('BImprimer'));
if Imprim<>nil then
   begin
   Imprim.OnClick:= Imprimer;
   Imprim.Visible:= True;
   end;

// Gestion des clicks sur Checkbox
{$IFNDEF EAGLCLIENT}
EmpPub:= TDBCheckBox (GetControl ('PAS_PUBLIC'));
LienParente:= TDBCheckBox (GetControl ('PAS_LIENPARENTE'));
Preavis1:= TDBCheckBox (GetControl ('PAS_MOTIFPREAVIS1'));
Preavis2:= TDBCheckBox (GetControl ('PAS_MOTIFPREAVIS2'));
Preavis3:= TDBCheckBox (GetControl ('PAS_MOTIFPREAVIS3'));
Alsace:= TDBCheckBox (GetControl ('PAS_ALSACEMOSEL'));
RegimeSS:= TDBCheckBox (GetControl ('PAS_REGIMESPECSS'));
PlanSoc:= TDBCheckBox (GetControl ('PAS_PLANSOC'));
AgeSalarieFNE:= TDBCheckBox (GetControl ('PAS_AGESALARIE'));
ChomTot:= TDBCheckBox (GetControl ('PAS_CHOMTOTAL'));
DDTEFP:= TDBCheckBox (GetControl ('PAS_DDTEFP'));
FNGSRED:= TDBCheckBox (GetControl ('PAS_FNGSRED'));
FNGS:= TDBCheckBox (GetControl ('PAS_FNGS'));
FNGS2:= TDBCheckBox (GetControl ('PAS_FNGS2'));
IndLeg:= TDBCheckBox (GetControl ('PAS_INDLEGB'));
IndCDD:= TDBCheckBox (GetControl ('PAS_INDCDDB'));
IndCNE:= TDBCheckBox (GetControl ('PAS_INDCNEB'));  //PT42-1
IndFinMiss:= TDBCheckBox (GetControl ('PAS_INDFINMISSB'));
IndRetraite:= TDBCheckBox (GetControl ('PAS_INDRETRAITEB'));
IndSinistre:= TDBCheckBox (GetControl ('PAS_INDSINISTREB'));     //PT42-1
IndSpec:= TDBCheckBox (GetControl ('PAS_INDSPECB'));
IndSpcfiq:= TDBCheckBox (GetControl ('PAS_INDSPCFIQB'));         //PT42-1
IndJournal:= TDBCheckBox (GetControl ('PAS_INDJOURNALB'));
IndClient:= TDBCheckBox (GetControl ('PAS_INDCLIENTB'));
IndAvion:= TDBCheckBox (GetControl ('PAS_INDAVIONB'));
IndApprenti:= TDBCheckBox (GetControl ('PAS_INDAPPRENTIB'));     //PT42-1
IndAutre:= TDBCheckBox (GetControl ('PAS_INDAUTREB'));

Arrco:= THDBEdit (GetControl ('PAS_ORGARRCO'));
Agirc:= THDBEdit (GetControl ('PAS_ORGAGIRC'));
Autres:= THDBEdit (GetControl ('PAS_ORGAUTRE'));
HorEtab:= THDBEdit (GetControl ('PAS_HORHEBENT'));
HorSal:= THDBEdit (GetControl ('PAS_HORHEBSAL'));
HorEtabann:= THDBEdit (GetControl ('PAS_HORANNENT'));
HorSalann:= THDBEdit (GetControl ('PAS_HORANNSAL'));
MotRupCont:= THDBEdit (GetControl ('PAS_MOTRUPCONT'));
{$ELSE}
EmpPub:= TCheckBox (GetControl ('PAS_PUBLIC'));
LienParente:= TCheckBox (GetControl ('PAS_LIENPARENTE'));
Preavis1:= TCheckBox (GetControl ('PAS_MOTIFPREAVIS1'));
Preavis2:= TCheckBox (GetControl ('PAS_MOTIFPREAVIS2'));
Preavis3:= TCheckBox (GetControl ('PAS_MOTIFPREAVIS3'));
Alsace:= TCheckBox (GetControl ('PAS_ALSACEMOSEL'));
RegimeSS:= TCheckBox (GetControl ('PAS_REGIMESPECSS'));
PlanSoc:= TCheckBox (GetControl ('PAS_PLANSOC'));
AgeSalarieFNE:= TCheckBox (GetControl ('PAS_AGESALARIE'));
ChomTot:= TCheckBox (GetControl ('PAS_CHOMTOTAL'));
DDTEFP:= TCheckBox (GetControl ('PAS_DDTEFP'));
FNGSRED:= TCheckBox (GetControl ('PAS_FNGSRED'));
FNGS:= TCheckBox (GetControl ('PAS_FNGS'));
FNGS2:= TCheckBox (GetControl ('PAS_FNGS2'));
IndLeg:= TCheckBox (GetControl ('PAS_INDLEGB'));
IndCDD:= TCheckBox (GetControl ('PAS_INDCDDB'));
IndCNE:= TCheckBox (GetControl ('PAS_INDCNEB'));  //PT42-1
IndFinMiss:= TCheckBox (GetControl ('PAS_INDFINMISSB'));
IndRetraite:= TCheckBox (GetControl ('PAS_INDRETRAITEB'));
IndSinistre:= TCheckBox (GetControl ('PAS_INDSINISTREB'));    //PT42-1
IndSpec:= TCheckBox (GetControl ('PAS_INDSPECB'));
IndSpcfiq:= TCheckBox (GetControl ('PAS_INDSPCFIQB'));        //PT42-1
IndJournal:= TCheckBox (GetControl ('PAS_INDJOURNALB'));
IndClient:= TCheckBox (GetControl ('PAS_INDCLIENTB'));
IndAvion:= TCheckBox (GetControl ('PAS_INDAVIONB'));
IndApprenti:= TCheckBox (GetControl ('PAS_INDAPPRENTIB'));    //PT42-1
IndAutre:= TCheckBox (GetControl ('PAS_INDAUTREB'));

Arrco:= THEdit (GetControl ('PAS_ORGARRCO'));
Agirc:= THEdit (GetControl ('PAS_ORGAGIRC'));
Autres:= THEdit (GetControl ('PAS_ORGAUTRE'));
HorEtab:= THEdit (GetControl ('PAS_HORHEBENT'));
HorSal:= THEdit (GetControl ('PAS_HORHEBSAL'));
HorEtabann:= THEdit (GetControl ('PAS_HORANNENT'));
HorSalann:= THEdit (GetControl ('PAS_HORANNSAL'));
MotRupCont:= THEdit (GetControl ('PAS_MOTRUPCONT'));
{$ENDIF}
{PT42-1
IndCNE:= TCheckBox (GetControl ('INDCNEB'));
IndCNEM:= THEdit (GetControl ('INDCNEM'));
IndSinistre:= TCheckBox (GetControl ('INDSINISTREB'));
IndSinistreM:= THEdit (GetControl ('INDSINISTREM'));
IndSpcfiq:= TCheckBox (GetControl ('INDSPCFIQB'));
IndSpcfiqM:= THEdit (GetControl ('INDSPCFIQM'));
IndApprenti:= TCheckBox (GetControl ('INDAPPRENTIB'));
IndApprentiM:= THEdit (GetControl ('INDAPPRENTIM'));
}
//deb PT51
Indrupconv:= TCheckBox (GetControl ('INDSPECONVENTIONB'));
Indrupconp:= THEdit (GetControl ('INDSPECONVENTIONM'));
//fin PT51

If EmpPub <> NIL then
   EmpPub.OnClick:= EmpPubClick;
If LienParente <> NIL then
   LienParente.OnClick:= LienParenteClick;
If Preavis1 <> NIL then
   Preavis1.OnClick:= PreavisClick;
If Preavis2 <> NIL then
   Preavis2.OnClick:= PreavisClick;
If Preavis3 <> NIL then
   Preavis3.OnClick:= PreavisClick;
If Alsace <> NIL then
   Alsace.OnClick:= AlsaceClick;
If RegimeSS <> NIL then
   RegimeSS.OnClick:= RegimeSSClick;
If PlanSoc <> NIL then
   PlanSoc.OnClick:= PlanSocClick;
If AgeSalarieFNE <> NIL then
   AgeSalarieFNE.OnClick:= AgeSalarieFNEClick;

If ChomTot <> NIL then
   ChomTot.OnClick:= ChomTotClick;
If DDTEFP <> NIL then
   DDTEFP.OnClick:= ChomTotClick;

If FNGSRED <> NIL then
   FNGSRED.OnClick:= FNGSClick;
If FNGS <> NIL then
   FNGS.OnClick:= FNGSClick;
If FNGS2 <> NIL then
   FNGS2.OnClick:= FNGSClick;

If IndLeg <> NIL then
   IndLeg.OnClick:= IndemniteClick;
If IndCDD <> NIL then
   IndCDD.OnClick:= IndemniteClick;
If IndCNE <> NIL then
   IndCNE.OnClick:= IndemniteClick;
{PT42-1
if IndCNEM <> NIL then
   IndCNEM.OnExit:= IndemniteExit;
}
If IndFinMiss <> NIL then
   IndFinMiss.OnClick:= IndemniteClick;
If IndRetraite <> NIL then
   IndRetraite.OnClick:= IndemniteClick;
If IndSinistre <> NIL then
   IndSinistre.OnClick:= IndemniteClick;
{PT42-1
if IndSinistreM <> NIL then
   IndSinistreM.OnExit:= IndemniteExit;
}
If IndSpec <> NIL then
   IndSpec.OnClick:= IndemniteClick;
If IndSpcfiq <> NIL then
   IndSpcfiq.OnClick:= IndemniteClick;
{PT42-1
if IndSpcfiqM <> NIL then
   IndSpcfiqM.OnExit:= IndemniteExit;
}
If IndJournal <> NIL then
   IndJournal.OnClick:= IndemniteClick;
If IndClient <> NIL then
   IndClient.OnClick:= IndemniteClick;
If IndAvion <> NIL then
   IndAvion.OnClick:= IndemniteClick;
If IndApprenti <> NIL then
   IndApprenti.OnClick:= IndemniteClick;
{PT42-1
if IndApprentiM <> NIL then
   IndApprentiM.OnExit:= IndemniteExit;
}
//deb PT51
if Indrupconp <> NIL then
   Indrupconp.OnExit:= IndemniteExit;
if Indrupconv <> nil then
   Indrupconv.OnClick := IndemniteClick;
//fin PT51

If IndAutre <> NIL then
   IndAutre.OnClick:= IndemniteClick;

// Gestion des changements dans les combo et dans les edit
If HorEtab <> NIL then
   HorEtab.OnExit:= HorExit;
If HorSal <> NIL then
   HorSal.OnExit:= HorExit;
If HorEtabann <> NIL then
   HorEtabann.OnExit:= HorExit;
If HorSalann <> NIL then
   HorSalann.OnExit:= HorExit;

If Arrco <> NIL then
   Arrco.OnChange:= ArrcoChange;
If Agirc <> NIL then
   Agirc.OnChange:= AgircChange;
If Autres <> NIL then
   Autres.OnChange:= AutresChange;

Etab:= RechDom('PGSALARIEETAB',PGAttesSalarie,False);
St:= '(PDA_ETABLISSEMENT = "" OR PDA_ETABLISSEMENT LIKE "%'+Etab+'%")'+
     ' AND (PDA_TYPEATTEST = "" OR  PDA_TYPEATTEST LIKE "%ASS%")';
SetControlProperty ('DECLARANT', 'Plus' ,St );

Edit:= THEdit (GetControl('DECLARANT'));
if Edit<>nil then
   Edit.OnExit:= AffichDeclarant;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 13/07/2000
Modifié le ... :   /  /
Description .. : OnNewRecord
Mots clefs ... : PAIE;PGATTESTATION;PGASSEDIC
*****************************************************************}
procedure TOM_Attestations.OnNewRecord;
var
IMax, j : Integer;
QRechNumAttest :TQuery;
begin
Inherited ;

Salaires := THGRID (GetControl ('GRID_SALAIRES'));
Primes   := THGRID (GetControl ('GRID_PRIMES'));
Solde   := THGRID (GetControl ('GRID_SOLDE'));

IMax:=1;
QRechNumAttest:=OpenSQL('SELECT MAX(PAS_ORDRE)'+
                        ' FROM ATTESTATIONS',TRUE);
if Not QRechNumAttest.EOF then
   try
   IMax:= QRechNumAttest.Fields[0].AsInteger+1;
   except
         on E: EConvertError do
            IMax:= 1;
   end;
Ferme(QRechNumAttest);
SetField('PAS_ORDRE',IMax);
SetField('PAS_TYPEATTEST','ASS');
SetField('PAS_DECLARDATE',Date);
SetField('PAS_DATEATTEST',IDate1900);
SetField('PAS_DATEADHESION',IDate1900);
SetField('PAS_DATEACCIDENT',IDate1900);
SetField('PAS_DATEARRET',IDate1900);
SetField('PAS_REPRISEARRET',IDate1900);
SetField('PAS_PERIODEDEBUT',IDate1900);
SetField('PAS_PERIODEFIN',IDate1900);
SetField('PAS_DATEPLANSOC',IDate1900);
SetField('PAS_DEBCHOMAGE',IDate1900);
SetField('PAS_FINCHOMAGE',IDate1900);
SetField('PAS_REPRISECHOM',IDate1900);
SetField('PAS_SUBDEBUT',IDate1900);
SetField('PAS_SUBFIN',IDate1900);

For j:=1 to 12 do
    begin
    Salaires.CellValues[0,j]:= '01/01/1900';
    Salaires.CellValues[1,j]:= '01/01/1900';
    Salaires.CellValues[2,j]:= '01/01/1900';
    end;


For j:=1 to 3 do
    begin
    Primes.CellValues[0,j]:= '01/01/1900';
    Primes.CellValues[1,j]:= '01/01/1900';
    Primes.CellValues[3,j]:= '01/01/1900';
    end;

For j:=1 to 2 do
    begin
    Solde.CellValues[0,j]:= '01/01/1900';
    Solde.CellValues[1,j]:= '01/01/1900';
    Solde.CellValues[2,j]:= '01/01/1900';
    end;
ChargeZones;
QRechNumAttest:= OpenSql ('SELECT PDA_DECLARANTATTES'+
                         ' FROM DECLARANTATTEST WHERE'+
                         ' (PDA_ETABLISSEMENT = "" OR'+
                         ' PDA_ETABLISSEMENT LIKE "%'+PGAttesEtab+'%") AND'+
                         ' (PDA_TYPEATTEST = "" OR'+
                         ' PDA_TYPEATTEST LIKE "%ASS%" )'+
                         ' ORDER BY PDA_ETABLISSEMENT DESC',True);
If Not QRechNumAttest.eof then
   Begin
   SetControlText('DECLARANT' ,QRechNumAttest.FindField('PDA_DECLARANTATTES').AsString);
   AffichDeclarant(nil);
   End;
Ferme(QRechNumAttest);
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 13/07/2000
Modifié le ... :   /  /
Description .. : OnLoadRecord
Mots clefs ... : PAIE;PGATTESTATION;PGASSEDIC
*****************************************************************}
procedure TOM_Attestations.OnLoadRecord;
begin
Inherited ;
If Not(DS.State in [dsInsert]) Then
   DerniereCreate:= 0;
ChargeZones;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 13/07/2000
Modifié le ... :   /  /
Description .. : OnUpdateRecord
Mots clefs ... : PAIE;PGATTESTATION;PGASSEDIC
*****************************************************************}
procedure TOM_Attestations.OnUpdateRecord;
begin
inherited;
OnFerme:= False;
If (DS.State in [dsInsert]) Then
   DerniereCreate:= GetField('PAS_ORDRE')
else
   if (DerniereCreate=GetField('PAS_ORDRE')) then
      OnFerme:= True; // le bug arrive on se casse !!!
try
   begintrans;
   UpdateTable;
   ExecuteSQL('UPDATE SALARIES SET PSA_ASSEDIC="X" WHERE'+
              ' PSA_SALARIE="'+PGAttesSalarie+'"');
   CommitTrans;
Except
   Rollback;
   PGIBox ('Une erreur est survenue lors de la mise à jour de la base',
           TFFiche(Ecran).caption);
END;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 13/07/2000
Modifié le ... :   /  /
Description .. : OnAfterUpdateRecord
Mots clefs ... : PAIE;PGATTESTATION;PGASSEDIC
*****************************************************************}
procedure TOM_Attestations.OnAfterUpdateRecord;
begin
inherited;
If OnFerme then
   Ecran.Close;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 17/06/2004
Modifié le ... :   /  /
Description .. : Suppression des éléments de la table secondaire lors de la
Suite ........ : suppression de l'attestation
Mots clefs ... : PAIE;PGATTESTATION;PGASSEDIC
*****************************************************************}
procedure TOM_Attestations.OnDeleteRecord;
begin
inherited;
try
   begintrans;
   ExecuteSQL('DELETE FROM ATTSALAIRES WHERE'+
              ' PAL_ORDRE='+IntToStr(GetField('PAS_ORDRE'))+' AND'+
              ' PAL_SALARIE="'+PGAttesSalarie+'" AND'+
              ' PAL_TYPEATTEST="ASS"');
   CommitTrans;
Except
   Rollback;
   PGIBox ('Une erreur est survenue lors de la mise à jour de la base',
           TFFiche(Ecran).caption);
   end;

Suppr:= True;
end;


//PT48
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 06/09/2007
Modifié le ... :   /  /
Description .. : Mise à jour du champ "Attestation faite"
Mots clefs ... : PAIE;PGATTESTATION;PGASSEDIC
*****************************************************************}
procedure TOM_Attestations.OnAfterDeleteRecord;
begin
inherited;
if (ExisteSQL ('SELECT PAS_SALARIE'+
               ' FROM ATTESTATIONS WHERE'+
               ' PAS_SALARIE="'+PGAttesSalarie+'" AND'+
               ' PAS_TYPEATTEST="ASS"')=False) then
   ExecuteSQL ('UPDATE SALARIES SET PSA_ASSEDIC="-" WHERE'+
               ' PSA_SALARIE="'+PGAttesSalarie+'"');
end;
//FIN PT48


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 13/07/2000
Modifié le ... :   /  /
Description .. : OnChangeField
Mots clefs ... : PAIE;PGATTESTATION;PGASSEDIC
*****************************************************************}
procedure TOM_Attestations.OnChangeField;
begin
inherited;
//PT48
if (Suppr=True) then
   exit;
//FIN PT48

if (F.FieldName=('PAS_PUBLICTYPE')) then
   PubTypeChange;

if (F.FieldName=('PAS_QUALITEEMPLOI')) then
   QualiteEmpChange;

if (F.FieldName=('PAS_CONTRATPARTIC')) then
   ContratPartChange;

if (F.FieldName=('PAS_FONCTIONSAL')) then
   FonctionChange;

if (F.FieldName=('PAS_MOTIFDIFF')) then
   DifHorChange;

if (F.FieldName=('PAS_QUALIF')) then
   SetControlText ('LABEL_QUALIF',
                  RechDom ('PGQUALIF',GetField('PAS_QUALIF'),FALSE,''));

if (F.FieldName=('PAS_MOTRUPCONT')) then
   begin
   if (MotRupCont<>nil) then
      MotRupCont.Plus:= ' AND PMS_ASSEDICCOR <> "403"';
   MotifChange;
   end;

if (F.FieldName=('PAS_AUTRERUPT')) then
   MotifChange;

if (F.FieldName=('PAS_INDLEGB')) or (F.FieldName=('PAS_INDLEGM')) then
   begin
   if (GetField('PAS_INDLEGB') = 'X') and (GetField('PAS_INDLEGM') <> 0) and
      (GetField('PAS_INDLEGM') <> null) then
      Leg := StrToFloat(GetField('PAS_INDLEGM'))
   else
      begin
      Leg := 0;
      if (GetField('PAS_INDLEGB') = 'X') and
         (GetField('PAS_INDLEGM') = null) then
         SetField('PAS_INDLEGB', '-');
      end;
   end;

if (F.FieldName=('PAS_INDCDDB')) or (F.FieldName=('PAS_INDCDDM')) then
   begin
   if (GetField('PAS_INDCDDB') = 'X') and (GetField('PAS_INDCDDM') <> 0) and
      (GetField('PAS_INDCDDM') <> null) then
      CDD := StrToFloat(GetField('PAS_INDCDDM'))
   else
      begin
      CDD := 0;
      if (GetField('PAS_INDCDDB') = 'X') and
         (GetField('PAS_INDCDDM') = null) then
         SetField('PAS_INDCDDB', '-');
      end;
   end;

//PT42-1
if (F.FieldName=('PAS_INDCNEB')) or (F.FieldName=('PAS_INDCNEM')) then
   begin
   if (GetField('PAS_INDCNEB') = 'X') and (GetField('PAS_INDCNEM') <> 0) and
      (GetField('PAS_INDCNEM') <> null) then
      CNE:= StrToFloat(GetField('PAS_INDCNEM'))
   else
      begin
      CNE:= 0;
      if (GetField('PAS_INDCNEB') = 'X') and
         (GetField('PAS_INDCNEM') = null) then
         SetField('PAS_INDCNEB', '-');
      end;
   end;
//FIN PT42-1

if (F.FieldName=('PAS_INDFINMISSB')) or (F.FieldName=('PAS_INDFINMISSM')) then
   begin
   if (GetField('PAS_INDFINMISSB') = 'X') and
      (GetField('PAS_INDFINMISSM') <> 0) and
      (GetField('PAS_INDFINMISSM') <> null) then
      FinMiss := StrToFloat(GetField('PAS_INDFINMISSM'))
   else
      begin
      FinMiss := 0;
      if (GetField('PAS_INDFINMISSB') = 'X') and
         (GetField('PAS_INDFINMISSM') = null) then
         SetField('PAS_INDFINMISSB', '-');
      end;
   end;

if (F.FieldName=('PAS_INDRETRAITEB')) or (F.FieldName=('PAS_INDRETRAITEM')) then
   begin
   if (GetField('PAS_INDRETRAITEB') = 'X') and
      (GetField('PAS_INDRETRAITEM') <> 0) and
      (GetField('PAS_INDRETRAITEM') <> null) then
      Retraite := StrToFloat(GetField('PAS_INDRETRAITEM'))
   else
      begin
      Retraite := 0;
      if (GetField('PAS_INDRETRAITEB') = 'X') and
         (GetField('PAS_INDRETRAITEM') = null) then
         SetField('PAS_INDRETRAITEB', '-');
      end;
   end;

//PT42-1
if (F.FieldName=('PAS_INDSINISTREB')) or (F.FieldName=('PAS_INDSINISTREM')) then
   begin
   if (GetField('PAS_INDSINISTREB') = 'X') and
      (GetField('PAS_INDSINISTREM') <> 0) and
      (GetField('PAS_INDSINISTREM') <> null) then
      Sinistre:= StrToFloat(GetField('PAS_INDSINISTREM'))
   else
      begin
      Sinistre:= 0;
      if (GetField('PAS_INDSINISTREB') = 'X') and
         (GetField('PAS_INDSINISTREM') = null) then
         SetField('PAS_INDSINISTREB', '-');
      end;
   end;
//FIN PT42-1

if (F.FieldName=('PAS_INDSPECB')) or (F.FieldName=('PAS_INDSPECM')) then
   begin
   if (GetField('PAS_INDSPECB') = 'X') and (GetField('PAS_INDSPECM') <> 0) and
      (GetField('PAS_INDSPECM') <> null) then
      Spec := StrToFloat(GetField('PAS_INDSPECM'))
   else
      begin
      Spec := 0;
      if (GetField('PAS_INDSPECB') = 'X') and
         (GetField('PAS_INDSPECM') = null) then
         SetField('PAS_INDSPECB', '-');
      end;
   end;

//PT42-1
if (F.FieldName=('PAS_INDSPCFIQB')) or (F.FieldName=('PAS_INDSPCFIQM')) then
   begin
   if (GetField('PAS_INDSPCFIQB') = 'X') and
      (GetField('PAS_INDSPCFIQM') <> 0) and
      (GetField('PAS_INDSPCFIQM') <> null) then
      Spcfiq:= StrToFloat(GetField('PAS_INDSPCFIQM'))
   else
      begin
      Spcfiq:= 0;
      if (GetField('PAS_INDSPCFIQB') = 'X') and
         (GetField('PAS_INDSPCFIQM') = null) then
         SetField('PAS_INDSPCFIQB', '-');
      end;
   end;
//FIN PT42-1

if (F.FieldName=('PAS_INDJOURNALB')) or (F.FieldName=('PAS_INDJOURNALM')) then
   begin
   if (GetField('PAS_INDJOURNALB') = 'X') and
      (GetField('PAS_INDJOURNALM') <> 0) and
      (GetField('PAS_INDJOURNALM') <> null) then
      Journal := StrToFloat(GetField('PAS_INDJOURNALM'))
   else
      begin
      Journal := 0;
      if (GetField('PAS_INDJOURNALB') = 'X') and
         (GetField('PAS_INDJOURNALM') = null) then
         SetField('PAS_INDJOURNALB', '-');
      end;
   end;

if (F.FieldName=('PAS_INDCLIENTB')) or (F.FieldName=('PAS_INDCLIENTM')) then
   begin
   if (GetField('PAS_INDCLIENTB') = 'X') and
      (GetField('PAS_INDCLIENTM') <> 0) and
      (GetField('PAS_INDCLIENTM') <> null) then
      Client := StrToFloat(GetField('PAS_INDCLIENTM'))
   else
      begin
      Client := 0;
      if (GetField('PAS_INDCLIENTB') = 'X') and
         (GetField('PAS_INDCLIENTM') = null) then
         SetField('PAS_INDCLIENTB', '-');
      end;
   end;

if (F.FieldName=('PAS_INDAVIONB')) or (F.FieldName=('PAS_INDAVIONM')) then
   begin
   if (GetField('PAS_INDAVIONB') = 'X') and (GetField('PAS_INDAVIONM') <> 0) and
      (GetField('PAS_INDAVIONM') <> null) then
      Avion := StrToFloat(GetField('PAS_INDAVIONM'))
   else
      begin
      Avion := 0;
      if (GetField('PAS_INDAVIONB') = 'X') and
         (GetField('PAS_INDAVIONM') = null) then
         SetField('PAS_INDAVIONB', '-');
      end;
   end;

//PT42-1
if (F.FieldName=('PAS_INDAPPRENTIB')) or (F.FieldName=('PAS_INDAPPRENTIM')) then
   begin
   if (GetField('PAS_INDAPPRENTIB') = 'X') and
      (GetField('PAS_INDAPPRENTIM') <> 0) and
      (GetField('PAS_INDAPPRENTIM') <> null) then
      Apprenti:= StrToFloat(GetField('PAS_INDAPPRENTIM'))
   else
      begin
      Apprenti:= 0;
      if (GetField('PAS_INDAPPRENTIB') = 'X') and
         (GetField('PAS_INDAPPRENTIM') = null) then
         SetField('PAS_INDAPPRENTIB', '-');
      end;
   end;
//FIN PT42-1   

if (F.FieldName=('PAS_INDAUTREB')) or (F.FieldName=('PAS_INDAUTREM')) then
   begin
   if (GetField('PAS_INDAUTREB') = 'X') and (GetField('PAS_INDAUTREM') <> 0) and
      (GetField('PAS_INDAUTREM') <> null) then
      Autre := StrToFloat(GetField('PAS_INDAUTREM'))
   else
      begin
      Autre := 0;
      if (GetField('PAS_INDAUTREB') = 'X') and
         (GetField('PAS_INDAUTREM') = null) then
         SetField('PAS_INDAUTREB', '-');
      end;
   end;

if (F.FieldName=('PAS_INDLEGB')) or (F.FieldName=('PAS_INDLEGM')) or
   (F.FieldName=('PAS_INDCDDB')) or (F.FieldName=('PAS_INDCDDM')) or
   (F.FieldName=('PAS_INDCNEB')) or (F.FieldName=('PAS_INDCNEM')) or  //PT42-1
   (F.FieldName=('PAS_INDFINMISSB')) or (F.FieldName=('PAS_INDFINMISSM')) or
   (F.FieldName=('PAS_INDRETRAITEB')) or (F.FieldName=('PAS_INDRETRAITEM')) or
   (F.FieldName=('PAS_INDSINISTREB')) or (F.FieldName=('PAS_INDSINISTREM')) or  //PT42-1
   (F.FieldName=('PAS_INDSPECB')) or (F.FieldName=('PAS_INDSPECM')) or
   (F.FieldName=('PAS_INDSPCFIQB')) or (F.FieldName=('PAS_INDSPCFIQM')) or  //PT42-1
   (F.FieldName=('PAS_INDJOURNALB')) or (F.FieldName=('PAS_INDJOURNALM')) or
   (F.FieldName=('PAS_INDCLIENTB')) or (F.FieldName=('PAS_INDCLIENTM')) or
   (F.FieldName=('PAS_INDAVIONB')) or (F.FieldName=('PAS_INDAVIONM')) or
   (F.FieldName=('PAS_INDAPPRENTIB')) or (F.FieldName=('PAS_INDAPPRENTIM')) or  //PT42-1
   (F.FieldName=('PAS_INDAUTREB')) or (F.FieldName=('PAS_INDAUTREM')) then
   begin
{PT42-1
   CNE:= StrToFloat(GetControlText('INDCNEM'));
   Sinistre:= StrToFloat(GetControlText('INDSINISTREM'));
   SpcFiq:= StrToFloat(GetControlText('INDSPCFIQM'));
   Apprenti:= StrToFloat(GetControlText('INDAPPRENTIM'));
}

     Indrupcon:= StrToFloat(GetControlText('INDSPECONVENTIONM'));    //PT51

   Total:= Leg+CDD+CNE+FinMiss+Retraite+Sinistre+Spec+Spcfiq+Journal+Client+
           Avion+Apprenti+Autre+Indrupcon;  //PT51
   SetControlText('EDIT_INDTOT', FloatToStr(Total));
   end;


(*Mise en commentaire PT45
{$IFDEF EAGLCLIENT}
if F.FieldName=('GRID_SALAIRES') then
   begin
   Salaires := THGrid (GetControl ('GRID_SALAIRES'));
   If Salaires <> NIL then
      BEGIN
      ESalaires;
      END;
   end;

if F.FieldName=('GRID_PRIMES') then
   begin
   Primes := THGrid (GetControl ('GRID_PRIMES'));
   If Primes <> NIL then
      BEGIN
      EPrimes;
      END;
   end;

if F.FieldName=('GRID_SOLDE') then
   begin
   Solde := THGrid (GetControl ('GRID_SOLDE'));
   If Solde <> NIL then
      BEGIN
      ESolde;
      END;
   end;
{$ENDIF}
*)

Imprim.Enabled:=TFFiche(Ecran).BDelete.Enabled;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 13/07/2000
Modifié le ... :   /  /
Description .. : OnClose
Mots clefs ... : PAIE;PGATTESTATION;PGASSEDIC
*****************************************************************}
procedure TOM_Attestations.OnClose;
begin
inherited;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 13/07/2000
Modifié le ... :   /  /
Description .. : ChargeZones
Mots clefs ... : PAIE;PGATTESTATION;PGASSEDIC
*****************************************************************}
procedure TOM_Attestations.ChargeZones();
var
QEntree, QRechEtab, QRechEtabComp, QRechOrg, QRechSal, QRechSalMSA : TQuery;
QSortie : TQuery;
NbSal : Integer;
AnneePrec, FonctionSal, MotifSortie, StDate, StEntree, StEtab : String;
StEtabComp, StOrg, StSal, StSortie, TypeContrat, TypeContratPart : String;
UsDateDeb, UsDateFin : String;
JourJ : TDateTime;
AnneeA, Jour, MoisM : Word;
begin

SetField ('PAS_SALARIE', PGAttesSalarie);

if (GetField ('PAS_DERNIERJOUR')<>IDate1900) then
   DerJour:= StrToDate (GetField ('PAS_DERNIERJOUR'));
UsDerJour:= UsDateTime (DerJour);
UsDateDeb:= UsDateTime (DebutDeMois (PlusMois (DerJour, -12)));
UsDateFin:= UsDateTime (DebutDeMois (DerJour));

StSal:= 'SELECT PSA_LIBELLE, PSA_PRENOM, PSA_ETABLISSEMENT, PSA_NUMEROSS,'+
        ' PSA_LIBELLEEMPLOI, PSA_STATUT, PSA_DATEENTREE, PSA_DATESORTIE,'+
        ' PSA_HORHEBDO, PSA_HORANNUEL, PSA_MOTIFSORTIE, PSA_CATDADS,'+
        ' PSA_FONCTIONSAL'+
        ' FROM SALARIES WHERE'+
        ' PSA_SALARIE="'+PGAttesSalarie+'"';
QRechSal:= OpenSql (StSal,TRUE);
SetControlText ('EDIT_NOM', QRechSal.FindField ('PSA_LIBELLE').Asstring);
SetControlText ('EDIT_PRENOM', QRechSal.FindField ('PSA_PRENOM').Asstring);
SetControlText ('EDIT_NUMSS', QRechSal.FindField ('PSA_NUMEROSS').Asstring);
CatDADS:= QRechSal.FindField ('PSA_CATDADS').Asstring;

PGAttesEtab:= QRechSal.FindField ('PSA_ETABLISSEMENT').Asstring;
StEtab:= 'SELECT ET_LIBELLE, ET_ADRESSE1, ET_ADRESSE2, ET_ADRESSE3,'+
         ' ET_CODEPOSTAL, ET_VILLE, ET_TELEPHONE, ET_JURIDIQUE, ET_SIRET,'+
         ' ET_APE'+
         ' FROM ETABLISS WHERE'+
         ' ET_ETABLISSEMENT="'+PGAttesEtab+'"';
QRechEtab:= OpenSql (StEtab,TRUE);
if (NOT QRechEtab.EOF) then
   begin
   SetControlText ('EDIT_ETABNOM', QRechEtab.FindField ('ET_LIBELLE').Asstring);
   SetControlText ('EDIT_ETABADR1', QRechEtab.FindField ('ET_ADRESSE1').Asstring);
   SetControlText ('EDIT_ETABADR2', QRechEtab.FindField ('ET_ADRESSE2').Asstring);
   SetControlText ('EDIT_ETABADR3', QRechEtab.FindField ('ET_ADRESSE3').Asstring);
   SetControlText ('EDIT_ETABCP', QRechEtab.FindField ('ET_CODEPOSTAL').Asstring);
   SetControlText ('EDIT_ETABVILLE', QRechEtab.FindField ('ET_VILLE').Asstring);
   SetControlText ('EDIT_ETABSTAT', RechDom ('TTFORMEJURIDIQUE',
                                             QRechEtab.FindField ('ET_JURIDIQUE').Asstring, False));
   SetControlText ('EDIT_ETABTEL', QRechEtab.FindField ('ET_TELEPHONE').Asstring);
   SetControlText ('EDIT_ETABSIRET', QRechEtab.FindField ('ET_SIRET').Asstring);
   if ControlSiret (GetControlText ('EDIT_ETABSIRET'))=False then
      PGIBox ('Le SIRET de l''établissement '+PGAttesEtab+' n''est pas valide',
              TFFiche(Ecran).caption);

   SetControlText ('EDIT_ETABAPE', QRechEtab.FindField ('ET_APE').Asstring);
   if (GetControlText ('DECLARANT')='') then
      SetField ('PAS_DECLARLIEU', QRechEtab.FindField ('ET_VILLE').Asstring);
   end;
Ferme (QRechEtab);

if (DS.State=dsInsert) then
   begin
   SetField ('PAS_PERIODEDEBUT', QRechSal.FindField ('PSA_DATEENTREE').AsDateTime);
   SetField ('PAS_PERIODEFIN', QRechSal.FindField ('PSA_DATESORTIE').AsDateTime);
   if (GetField ('PAS_PERIODEFIN') > IDate1900) then
      SetField ('PAS_DERNIERJOUR',GetField ('PAS_PERIODEFIN'))
   else
      SetField ('PAS_DERNIERJOUR', Date);
   SetField ('PAS_DERNIEREMPLOI', QRechSal.FindField ('PSA_LIBELLEEMPLOI').Asstring);
   StSal:= 'SELECT PSE_SALARIE, PSE_MSA'+
           ' FROM DEPORTSAL WHERE'+
           ' PSE_SALARIE="'+PGAttesSalarie+'" AND'+
           ' PSE_MSA="X"';
   QRechSalMSA:= OpenSql (StSal,TRUE);
   if (NOT QRechSalMSA.EOF) then
      SetField ('PAS_REGIMESPECSS', 'X');
   Ferme (QRechSalMSA);
   SetField ('PAS_HORHEBSAL', QRechSal.FindField ('PSA_HORHEBDO').AsFloat);
   SetField ('PAS_HORANNSAL', QRechSal.FindField ('PSA_HORANNUEL').AsFloat);

   MotifSortie:= '';
   TypeContrat:= '';
   TypeContratPart:= '';
   FonctionSal:= ''; //PT50
   if (GetField ('PAS_DERNIERJOUR')<>IDate1900) then
      begin
{PT50
      GetInfoContratTravail (PGAttesSalarie,
                             StrToDate(GetField('PAS_DERNIERJOUR')),
                             MotifSortie, TypeContrat, TypeContratPart);
}
      if (QRechSal.FindField ('PSA_MOTIFSORTIE').Asstring<>'') then
         MotifSortie:= QRechSal.FindField ('PSA_MOTIFSORTIE').Asstring;

      if (QRechSal.FindField ('PSA_FONCTIONSAL').Asstring<>'') then
         FonctionSal:= QRechSal.FindField ('PSA_FONCTIONSAL').Asstring;

      GetInfoContratTravail (PGAttesSalarie,
                             StrToDate (GetField ('PAS_DERNIERJOUR')),
                             MotifSortie, TypeContrat, TypeContratPart,
                             FonctionSal);
//FIN PT50
{PT45
      StContrat:= 'SELECT PCI_MOTIFSORTIE, PCI_TYPECONTRAT'+
                  ' FROM CONTRATTRAVAIL WHERE'+
                  ' PCI_SALARIE="'+PGAttesSalarie+'" AND'+
                  ' PCI_FINCONTRAT<="'+UsDateTime(StrToDate(GetField('PAS_DERNIERJOUR')))+'"'+
                  ' ORDER BY PCI_FINCONTRAT DESC';
      QRechContrat:= OpenSql(StContrat,TRUE);
      if NOT QRechContrat.EOF then
         begin
         MotifSortie:= QRechContrat.FindField('PCI_MOTIFSORTIE').Asstring;
         TypeContrat:= QRechContrat.FindField('PCI_TYPECONTRAT').Asstring;
         if (TypeContrat='CAA') then
            begin
            TypeContrat:= '002';
            TypeContratPart:= '1';
            end
         else
         if (TypeContrat='CAC') then
            begin
            TypeContrat:= '002';
            TypeContratPart:= '1';
            end
         else
         if (TypeContrat='CAE') then
            begin
            TypeContrat:= '002';
            TypeContratPart:= '6';
            end
         else
         if (TypeContrat='CAP') then
            begin
            TypeContrat:= '002';
            TypeContratPart:= '1';
            end
         else
         if (TypeContrat='CCD') then
            TypeContrat:= '002'
         else
         if (TypeContrat='CDI') then
            TypeContrat:= '001'
         else
         if (TypeContrat='CEJ') then
            begin
            TypeContrat:= '002';
            TypeContratPart:= '6';
            end
         else
         if (TypeContrat='CES') then
            begin
            TypeContrat:= '002';
            TypeContratPart:= '5';
            end
         else
         if (TypeContrat='CTT') then
            begin
            TypeContrat:= '002';
            TypeContratPart:= '6';
            end
         else
         if (TypeContrat='SC') then
            TypeContrat:= '001';
         end;
}

{PT50
      if (QRechSal.FindField ('PSA_MOTIFSORTIE').Asstring<>'') then
         MotifSortie:= QRechSal.FindField ('PSA_MOTIFSORTIE').Asstring;
}

      if (MotifSortie='403') then
         MotifSortie:= '';

      SetField ('PAS_MOTRUPCONT', MotifSortie);

{PT45
      Ferme(QRechContrat);
}
      end;
   SetField ('PAS_CONTRATNAT', TypeContrat);
   SetField ('PAS_CONTRATPARTIC', TypeContratPart);

{PT50
   SetField ('PAS_FONCTIONSAL', QRechSal.FindField('PSA_FONCTIONSAL').Asstring);//PT46
}
   SetField ('PAS_FONCTIONSAL', FonctionSal);
//FIN PT50

   StEtabComp:= 'SELECT PHE_HORAIREHEURE, PHE_HORAIREAN'+
                ' FROM HORAIREETAB WHERE'+
                ' PHE_ETABLISSEMENT = "'+PGAttesEtab+'"';
   QRechEtabComp:= OpenSql (StEtabComp, TRUE);
   if (NOT QRechEtabComp.EOF) then
      begin
      SetField ('PAS_HORHEBENT',QRechEtabComp.FindField ('PHE_HORAIREHEURE').AsFloat);
      SetField ('PAS_HORANNENT',QRechEtabComp.FindField ('PHE_HORAIREAN').AsFloat);
      end;
   Ferme (QRechEtabComp);

   StOrg:= 'SELECT POG_LIBELLE, POG_NUMAFFILIATION'+
           ' FROM ORGANISMEPAIE WHERE'+
           ' POG_ETABLISSEMENT="'+PGAttesEtab+'" AND'+
           ' POG_NATUREORG="200"';
   QRechOrg:= OpenSql (StOrg, TRUE);
   if (NOT QRechOrg.EOF) then
      begin
      SetField ('PAS_ASSEDICCAISSE', QRechOrg.FindField ('POG_LIBELLE').Asstring);
      SetField ('PAS_ASSEDICNUM', QRechOrg.FindField ('POG_NUMAFFILIATION').Asstring);
      end;
   Ferme (QRechOrg);
   end;
Ferme (QRechSal);

ArrcoChange (nil);
AgircChange (nil);
AutresChange (nil);

if (GetField ('PAS_PUBLICTYPE')='') then
   begin
   SetControlVisible ('LABEL_CONVNUM', FALSE);
   SetControlVisible ('PAS_CONVNUMERO', FALSE);
   SetControlVisible ('LABEL_CONVANA', FALSE);
   SetControlVisible ('PAS_CONVCODEANA', FALSE);
   SetControlVisible ('PAS_ADHESION', FALSE);
   SetControlVisible ('LABEL_COLLTERRDATE', FALSE);
   SetControlVisible ('PAS_DATEADHESION', FALSE);
   SetField ('PAS_CONVNUMERO', '');
   SetField ('PAS_CONVCODEANA', '');
   SetField ('PAS_ADHESION', '-');
   SetField ('PAS_DATEADHESION', IDate1900);
   end
else
if (GetField('PAS_PUBLICTYPE')='001') then
   begin
   SetControlVisible ('LABEL_CONVNUM', FALSE);
   SetControlVisible ('PAS_CONVNUMERO', FALSE);
   SetControlVisible ('LABEL_CONVANA', FALSE);
   SetControlVisible ('PAS_CONVCODEANA', FALSE);
   SetControlVisible ('PAS_ADHESION', TRUE);
   SetControlVisible ('LABEL_COLLTERRDATE', FALSE);
   SetControlVisible ('PAS_DATEADHESION', FALSE);
   SetField ('PAS_CONVNUMERO', '');
   SetField ('PAS_CONVCODEANA', '');
   SetField ('PAS_DATEADHESION', IDate1900);
   end
else
if (GetField ('PAS_PUBLICTYPE')='002') then
   begin
   SetControlVisible ('LABEL_CONVNUM', TRUE);
   SetControlVisible ('PAS_CONVNUMERO', TRUE);
   SetControlVisible ('LABEL_CONVANA', TRUE);
   SetControlVisible ('PAS_CONVCODEANA', TRUE);
   SetControlVisible ('PAS_ADHESION', TRUE);
   SetControlVisible ('LABEL_COLLTERRDATE', FALSE);
   SetControlVisible ('PAS_DATEADHESION', FALSE);
   SetField ('PAS_DATEADHESION', IDate1900);
   end
else
if (GetField ('PAS_PUBLICTYPE')='003') then
   begin
   SetControlVisible ('LABEL_CONVNUM', FALSE);
   SetControlVisible ('PAS_CONVNUMERO', FALSE);
   SetControlVisible ('LABEL_CONVANA', FALSE);
   SetControlVisible ('PAS_CONVCODEANA', FALSE);
   SetControlVisible ('PAS_ADHESION', FALSE);
   SetControlVisible ('LABEL_COLLTERRDATE', TRUE);
   SetControlVisible ('PAS_DATEADHESION', TRUE);
   SetField ('PAS_CONVNUMERO', '');
   SetField ('PAS_CONVCODEANA', '');
   SetField ('PAS_ADHESION', '-');
   end;

if (GetField ('PAS_QUALITEEMPLOI')='7') then
   SetControlVisible ('PAS_LQUALITE', TRUE)
else
   SetControlVisible ('PAS_LQUALITE', FALSE);

if (GetField ('PAS_CONTRATPARTIC')='6') then
   begin
   SetControlVisible ('LABELSAISIECONTRAT', TRUE);
   SetControlVisible ('PAS_LCONTRATPART', TRUE);
   end
else
   begin
   SetControlVisible ('LABELSAISIECONTRAT', FALSE);
   SetControlVisible ('PAS_LCONTRATPART', FALSE);
   end;

if (GetField ('PAS_FONCTIONSAL')='8') then
   begin
   SetControlVisible ('LABELSAISIEFONC', TRUE);
   SetControlVisible ('PAS_LIBFONCTION', TRUE);
   end
else
    begin
    SetControlVisible ('LABELSAISIEFONC', FALSE);
    SetControlVisible ('PAS_LIBFONCTION', FALSE);
    end;

if (GetField ('PAS_MOTIFDIFF')='3') then
   begin
   SetControlVisible ('LABELSAISIEDIFF', TRUE);
   SetControlVisible ('PAS_AUTREDIFF', TRUE);
   end
else
   begin
   SetControlVisible ('LABELSAISIEDIFF', FALSE);
   SetControlVisible ('PAS_AUTREDIFF', FALSE);
   end;

if (GetField ('PAS_MOTRUPCONT')='20') or (GetField ('PAS_MOTRUPCONT')='59') or
   (GetField ('PAS_MOTRUPCONT')='60') then
   begin
   SetControlVisible ('LABELSAISIEMOTIF', TRUE);
   SetControlVisible ('PAS_AUTRERUPT', TRUE);
   end
else
   begin
   SetControlVisible ('LABELSAISIEMOTIF', FALSE);
   SetControlVisible ('PAS_AUTRERUPT', FALSE);
   SetField ('PAS_AUTRERUPT', '');
   end;

if (GetField ('PAS_MOTRUPCONT')='14') then
   begin
   SetControlVisible ('GROUPBOXMOT2', TRUE);
   SetControlVisible ('GROUPBOXMOT3', TRUE);
   end
else
   begin
   SetControlVisible ('GROUPBOXMOT2', FALSE);
   SetControlVisible ('GROUPBOXMOT3', FALSE);
   if (PlanSoc<>NIL) then
      PlanSoc.Checked:= FALSE;
   if (AgeSalarieFNE<>NIL) then
      AgeSalarieFNE.Checked:= FALSE;
   end;

if (DS.State=dsInsert) then
   begin
   JourJ:= Date;
   DecodeDate (JourJ, AnneeA, MoisM, Jour);
   AnneePrec:= IntToStr (AnneeA-1);
   StDate:= UsDateTime (StrToDate ('31/12/'+AnneePrec));
   StEntree:= 'SELECT COUNT(PSA_SALARIE) AS NBENTREE'+
              ' FROM SALARIES WHERE'+
              ' PSA_DATEENTREE<"'+StDate+'" AND'+
              ' PSA_DATEENTREE<>"'+UsDateTime(IDate1900)+'" AND'+
              ' PSA_DATEENTREE IS NOT NULL AND'+
              ' PSA_ETABLISSEMENT="'+PGAttesEtab+'"';
   QEntree:= OpenSql (StEntree,True);

   StSortie:= 'SELECT COUNT(PSA_SALARIE) AS NBSORTIE'+
              ' FROM SALARIES WHERE'+
              ' PSA_DATESORTIE<"'+StDate+'" AND'+
              ' PSA_DATESORTIE<>"'+UsDateTime(IDate1900)+'" AND'+
              ' PSA_DATESORTIE IS NOT NULL AND'+
              ' PSA_ETABLISSEMENT="'+PGAttesEtab+'"';
   QSortie:= OpenSql (StSortie, True);

   NbSal:= 0;
   if ((not QEntree.eof) and (not QSortie.eof)) then
      NbSal:= (QEntree.FindField ('NBENTREE').asinteger)-
              (QSortie.FindField ('NBSORTIE').asinteger);
   SetField ('PAS_EFFECTIF', NbSal);
   Ferme (QEntree);
   Ferme (QSortie);
   end;

if (DS.State=dsInsert) then
   begin
   SetField ('PAS_LIEUTRAVAIL', GetControlText ('EDIT_ETABVILLE'));
   SetField ('PAS_DEPART', Copy (GetControlText ('EDIT_ETABCP'), 1, 2));
   end;

// Chargement contenu des Caisses
ChargeCaisses;

// Chargement contenu des grilles de salaires et de primes
if (DS.State = dsInsert) then
   ChargeIndemnites;
ChargeSalaires (FALSE);
ChargePrimes (FALSE);
ChargeSolde (FALSE);

Ecran.Caption:= 'Attestation ASSEDIC : '+PGAttesSalarie+' '+
                GetControlText ('EDIT_NOM')+' '+GetControlText ('EDIT_PRENOM');
TFFiche (Ecran).DisabledMajCaption:= True;

SetControlEnabled ('PAS_SALARIE', False);

HorExit (Nil);

Suppr:= False;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 13/07/2000
Modifié le ... :   /  /
Description .. : ChargeCaisses
Mots clefs ... : PAIE;PGATTESTATION;PGASSEDIC
*****************************************************************}
procedure TOM_Attestations.ChargeCaisses();
var
QRechCaisse : TQuery;
StCaisse : String;
begin
StCaisse:= 'SELECT POG_ADRESSE1, POG_ADRESSE2, POG_ADRESSE3, POG_CODEPOSTAL,'+
           ' POG_VILLE'+
           ' FROM ORGANISMEPAIE WHERE'+
           ' POG_ETABLISSEMENT = "'+PGAttesEtab+'" AND'+
           ' POG_ORGANISME="'+GetField('PAS_ORGARRCO')+'"';
QRechCaisse:=OpenSql(StCaisse,TRUE);
if (NOT QRechCaisse.EOF) then
   SetControlText('LABEL_COADR1',
                  Copy(QRechCaisse.FindField('POG_ADRESSE1').Asstring, 1,
                       Length(QRechCaisse.FindField('POG_ADRESSE1').Asstring))+' '+
                  Copy(QRechCaisse.FindField('POG_ADRESSE2').Asstring, 1,
                       Length(QRechCaisse.FindField('POG_ADRESSE2').Asstring))+' '+
                  Copy(QRechCaisse.FindField('POG_ADRESSE3').Asstring, 1,
                       Length(QRechCaisse.FindField('POG_ADRESSE3').Asstring))+' '+
                  Copy(QRechCaisse.FindField('POG_CODEPOSTAL').Asstring, 1,
                       Length(QRechCaisse.FindField('POG_CODEPOSTAL').Asstring))+' '+
                  Copy(QRechCaisse.FindField('POG_VILLE').Asstring, 1,
                       Length(QRechCaisse.FindField('POG_VILLE').Asstring)));
Ferme(QRechCaisse);

StCaisse:= 'SELECT POG_ADRESSE1, POG_ADRESSE2, POG_ADRESSE3, POG_CODEPOSTAL,'+
           ' POG_VILLE'+
           ' FROM ORGANISMEPAIE WHERE'+
           ' POG_ETABLISSEMENT = "'+PGAttesEtab+'" AND'+
           ' POG_ORGANISME="'+GetField('PAS_ORGAGIRC')+'"';
QRechCaisse:=OpenSql(StCaisse,TRUE);
if (NOT QRechCaisse.EOF) then
   SetControlText('LABEL_RCADR1',
                  Copy(QRechCaisse.FindField('POG_ADRESSE1').Asstring, 1,
                       Length(QRechCaisse.FindField('POG_ADRESSE1').Asstring))+' '+
                  Copy(QRechCaisse.FindField('POG_ADRESSE2').Asstring, 1,
                       Length(QRechCaisse.FindField('POG_ADRESSE2').Asstring))+' '+
                  Copy(QRechCaisse.FindField('POG_ADRESSE3').Asstring, 1,
                       Length(QRechCaisse.FindField('POG_ADRESSE3').Asstring))+' '+
                  Copy(QRechCaisse.FindField('POG_CODEPOSTAL').Asstring, 1,
                       Length(QRechCaisse.FindField('POG_CODEPOSTAL').Asstring))+' '+
                  Copy(QRechCaisse.FindField('POG_VILLE').Asstring, 1,
                       Length(QRechCaisse.FindField('POG_VILLE').Asstring)));
Ferme(QRechCaisse);

StCaisse:= 'SELECT POG_ADRESSE1, POG_ADRESSE2, POG_ADRESSE3, POG_CODEPOSTAL,'+
           ' POG_VILLE'+
           ' FROM ORGANISMEPAIE WHERE'+
           ' POG_ETABLISSEMENT = "'+PGAttesEtab+'" AND'+
           ' POG_ORGANISME="'+GetField('PAS_ORGAUTRE')+'"';
QRechCaisse:=OpenSql(StCaisse,TRUE);
if (NOT QRechCaisse.EOF) then
   SetControlText('LABEL_AUADR1',
                  Copy(QRechCaisse.FindField('POG_ADRESSE1').Asstring, 1,
                       Length(QRechCaisse.FindField('POG_ADRESSE1').Asstring))+' '+
                  Copy(QRechCaisse.FindField('POG_ADRESSE2').Asstring, 1,
                       Length(QRechCaisse.FindField('POG_ADRESSE2').Asstring))+' '+
                  Copy(QRechCaisse.FindField('POG_ADRESSE3').Asstring, 1,
                       Length(QRechCaisse.FindField('POG_ADRESSE3').Asstring))+' '+
                  Copy(QRechCaisse.FindField('POG_CODEPOSTAL').Asstring, 1,
                       Length(QRechCaisse.FindField('POG_CODEPOSTAL').Asstring))+' '+
                  Copy(QRechCaisse.FindField('POG_VILLE').Asstring, 1,
                       Length(QRechCaisse.FindField('POG_VILLE').Asstring)));
Ferme(QRechCaisse);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 13/07/2000
Modifié le ... :   /  /
Description .. : ChargeIndemnités
Mots clefs ... : PAIE;PGATTESTATION;PGASSEDIC
*****************************************************************}
procedure TOM_Attestations.ChargeIndemnites();
var
QRechConv, QRechRub, QrechRCP, QRechRPreavis : TQuery;
StConv, StRub, StRCP, StRPreavis, UsDateFin : String;
begin

CP:= 0;
Preavis:= 0;

Leg:= 0;
CDD:= 0;
FinMiss:= 0;
Retraite:= 0;
Spec:= 0;
Journal:= 0;
Client:= 0;
Avion:= 0;
Autre:= 0;
CNE:= 0;
Sinistre:= 0;
Spcfiq:= 0;
Apprenti:= 0;

if GetField('PAS_DERNIERJOUR') <> IDate1900 then
   DerJour := StrToDate(GetField('PAS_DERNIERJOUR'));
UsDerJour := UsDateTime(DerJour);
if (StrToDate(GetField('PAS_PERIODEDEBUT'))<DebutDeMois(DerJour)) then
   UsDateFin := UsDateTime (DebutDeMois(DerJour))
else
   UsDateFin := UsDateTime (StrToDate(GetField('PAS_PERIODEDEBUT')));

StRub:='SELECT PRM_INDEMLEGALES, SUM(PHB_MTREM) AS MTINDEMLEGALES'+
       ' FROM HISTOBULLETIN'+
       ' LEFT JOIN REMUNERATION ON'+
       ' PHB_NATURERUB=PRM_NATURERUB AND'+
       ' ##PRM_PREDEFINI## PHB_RUBRIQUE=PRM_RUBRIQUE WHERE'+
       ' PRM_INDEMLEGALES<>"" AND'+
       ' PHB_SALARIE="'+PGAttesSalarie+'" AND'+
       ' PHB_DATEFIN<="'+UsDerJour+'" AND'+
       ' PHB_DATEFIN>="'+UsDateFin+'" AND'+
       ' PHB_NATURERUB="AAA"'+
       ' GROUP BY PRM_INDEMLEGALES';
QRechRub:=OpenSql(StRub,TRUE);

if ((not QRechRub.eof) and
   (QRechRub.FindField('PRM_INDEMLEGALES').Asstring = '001')) then
   begin
   if QRechRub.FindField('MTINDEMLEGALES').AsFloat <> 0 then
      begin
      SetField('PAS_INDLEGB','X');
      Leg := QRechRub.FindField('MTINDEMLEGALES').AsFloat;
      SetField ('PAS_INDLEGM',Leg);
      end;
   QRechRub.Next;
   end;

if ((not QRechRub.eof) and
   (QRechRub.FindField('PRM_INDEMLEGALES').Asstring = '003')) then
   begin
   if QRechRub.FindField('MTINDEMLEGALES').AsFloat <> 0 then
      begin
      SetField('PAS_INDCDDB','X');
      CDD := QRechRub.FindField('MTINDEMLEGALES').AsFloat;
      SetField ('PAS_INDCDDM',CDD);
      end;
   QRechRub.Next;
   end;

if ((not QRechRub.eof) and
   (QRechRub.FindField('PRM_INDEMLEGALES').Asstring = '011')) then
   begin
   if QRechRub.FindField('MTINDEMLEGALES').AsFloat <> 0 then
      begin
{PT42-1
      SetField('INDCNEB','X');
}
      SetField('PAS_INDCNEB','X');
      CNE:= QRechRub.FindField('MTINDEMLEGALES').AsFloat;
{PT42-1
      SetField ('INDCNEM', CNE);
}
      SetField ('PAS_INDCNEM', CNE);
      end;
   QRechRub.Next;
   end;

if ((not QRechRub.eof) and
   (QRechRub.FindField('PRM_INDEMLEGALES').Asstring = '004')) then
   begin
   if QRechRub.FindField('MTINDEMLEGALES').AsFloat <> 0 then
      begin
      SetField('PAS_INDAUTREB','X');
      Autre := QRechRub.FindField('MTINDEMLEGALES').AsFloat;
      SetField ('PAS_INDAUTREM',Autre);
      end;
   QRechRub.Next;
   end;

if ((not QRechRub.eof) and
   (QRechRub.FindField('PRM_INDEMLEGALES').Asstring = '005')) then
   begin
   if QRechRub.FindField('MTINDEMLEGALES').AsFloat <> 0 then
      begin
      SetField('PAS_INDFINMISSB','X');
      FinMiss := QRechRub.FindField('MTINDEMLEGALES').AsFloat;
      SetField ('PAS_INDFINMISSM',FinMiss);
      end;
   QRechRub.Next;
   end;

if ((not QRechRub.eof) and
   (QRechRub.FindField('PRM_INDEMLEGALES').Asstring = '006')) then
   begin
   if QRechRub.FindField('MTINDEMLEGALES').AsFloat <> 0 then
      begin
      SetField('PAS_INDRETRAITEB','X');
      Retraite := QRechRub.FindField('MTINDEMLEGALES').AsFloat;
      SetField ('PAS_INDRETRAITEM',Retraite);
      end;
   QRechRub.Next;
   end;

if ((not QRechRub.eof) and
   (QRechRub.FindField('PRM_INDEMLEGALES').Asstring = '012')) then
   begin
   if QRechRub.FindField('MTINDEMLEGALES').AsFloat <> 0 then
      begin
{PT42-1
      SetField('INDSINISTREB','X');
}
      SetField ('PAS_INDSINISTREB','X');
      Sinistre:= QRechRub.FindField('MTINDEMLEGALES').AsFloat;
{PT42-1
      SetField ('INDSINISTREM', Sinistre);
}
      SetField ('PAS_INDSINISTREM', Sinistre);
      end;
   QRechRub.Next;
   end;

if ((not QRechRub.eof) and
   (QRechRub.FindField('PRM_INDEMLEGALES').Asstring = '007')) then
   begin
   if QRechRub.FindField('MTINDEMLEGALES').AsFloat <> 0 then
      begin
      SetField('PAS_INDSPECB','X');
      Spec := QRechRub.FindField('MTINDEMLEGALES').AsFloat;
      SetField ('PAS_INDSPECM',Spec);
      end;
   QRechRub.Next;
   end;

if ((not QRechRub.eof) and
   (QRechRub.FindField('PRM_INDEMLEGALES').Asstring = '013')) then
   begin
   if QRechRub.FindField('MTINDEMLEGALES').AsFloat <> 0 then
      begin
{PT42-1
      SetField('INDSPCFIQB','X');
}
      SetField ('PAS_INDSPCFIQB','X');
      Spcfiq:= QRechRub.FindField('MTINDEMLEGALES').AsFloat;
{PT42-1
      SetField ('INDSPCFIQM',Spcfiq);
}
      SetField ('PAS_INDSPCFIQM',Spcfiq);
      end;
   QRechRub.Next;
   end;

if ((not QRechRub.eof) and
   (QRechRub.FindField('PRM_INDEMLEGALES').Asstring = '008')) then
   begin
   if QRechRub.FindField('MTINDEMLEGALES').AsFloat <> 0 then
      begin
      SetField('PAS_INDJOURNALB','X');
      Journal := QRechRub.FindField('MTINDEMLEGALES').AsFloat;
      SetField ('PAS_INDJOURNALM',Journal);
      end;
   QRechRub.Next;
   end;

if ((not QRechRub.eof) and
   (QRechRub.FindField('PRM_INDEMLEGALES').Asstring = '009')) then
   begin
   if QRechRub.FindField('MTINDEMLEGALES').AsFloat <> 0 then
      begin
      SetField('PAS_INDCLIENTB','X');
      Client := QRechRub.FindField('MTINDEMLEGALES').AsFloat;
      SetField ('PAS_INDCLIENTM',Client);
      end;
   QRechRub.Next;
   end;

if ((not QRechRub.eof) and
   (QRechRub.FindField('PRM_INDEMLEGALES').Asstring = '010')) then
   begin
   if QRechRub.FindField('MTINDEMLEGALES').AsFloat <> 0 then
      begin
      SetField('PAS_INDAVIONB','X');
      Avion := QRechRub.FindField('MTINDEMLEGALES').AsFloat;
      SetField ('PAS_INDAVIONM',Avion);
      end;
   QRechRub.Next;
   end;

if ((not QRechRub.eof) and
   (QRechRub.FindField('PRM_INDEMLEGALES').Asstring = '014')) then
   begin
   if QRechRub.FindField('MTINDEMLEGALES').AsFloat <> 0 then
      begin
{PT42-1
      SetField('INDAPPRENTIB','X');
}
      SetField('PAS_INDAPPRENTIB','X');
      Apprenti := QRechRub.FindField('MTINDEMLEGALES').AsFloat;
{PT42-1
      SetField ('INDAPPRENTIM',Apprenti);
}
      SetField ('PAS_INDAPPRENTIM',Apprenti);
      end;
   QRechRub.Next;
   end;

//PT42-2
if ((not QRechRub.eof) and
   (QRechRub.FindField('PRM_INDEMLEGALES').Asstring = '015')) then
   begin
   if (QRechRub.FindField('MTINDEMLEGALES').AsFloat<>0) then
      begin
      Transac:= QRechRub.FindField ('MTINDEMLEGALES').AsFloat;
      SetField ('PAS_INDTRANS2', Transac);
      end;
   QRechRub.Next;
   end;
//FIN PT42-2

//deb PT51
if ((not QRechRub.eof) and
   (QRechRub.FindField('PRM_INDEMLEGALES').Asstring = '016')) then
   begin
   if QRechRub.FindField('MTINDEMLEGALES').AsFloat <> 0 then
      begin
      SetField('INDSPECONVENTIONB','X');
      Indrupcon := QRechRub.FindField('MTINDEMLEGALES').AsFloat;
      SetField ('INDSPECONVENTIONM',Indrupcon);
      end;
   QRechRub.Next;
   end;
//fin PT51
Ferme(QRechRub);

StRCP:='SELECT SUM(PHB_MTREM) AS MTCP, SUM(PHB_BASEREM) AS BASECP'+
       ' FROM HISTOBULLETIN'+
       ' LEFT JOIN REMUNERATION ON'+
       ' PHB_NATURERUB=PRM_NATURERUB AND'+
       ' ##PRM_PREDEFINI## PHB_RUBRIQUE=PRM_RUBRIQUE WHERE'+
       ' PRM_INDEMCOMPCP="X" AND'+
       ' PHB_SALARIE="'+PGAttesSalarie+'" AND'+
       ' PHB_DATEFIN<="'+UsDerJour+'" AND'+
       ' PHB_DATEFIN>="'+UsDateFin+'" AND'+
       ' PHB_NATURERUB="AAA"';

QRechRCP:=OpenSql(StRCP,TRUE);
if (not QRechRCP.eof) then
   begin
   if QRechRCP.FindField('MTCP').AsFloat <> 0 then
      begin
      CP := QRechRCP.FindField('MTCP').AsFloat;
      SetField ('PAS_INDCPMONTANT',CP);
      end;
   if QRechRCP.FindField('BASECP').AsFloat <> 0 then
      begin
      CP := QRechRCP.FindField('BASECP').AsFloat;
      SetField ('PAS_INDCPJOURS',CP);
      end;
   end;
Ferme(QRechRCP);

StRPreavis:='SELECT SUM(PHB_MTREM) AS MTPREAVIS'+
            ' FROM HISTOBULLETIN'+
            ' LEFT JOIN REMUNERATION ON'+
            ' PHB_NATURERUB=PRM_NATURERUB AND'+
            ' ##PRM_PREDEFINI## PHB_RUBRIQUE=PRM_RUBRIQUE WHERE'+
            ' PRM_INDEMPREAVIS="X" AND'+
            ' PHB_SALARIE="'+PGAttesSalarie+'" AND'+
            ' PHB_DATEFIN<="'+UsDerJour+'" AND'+
            ' PHB_DATEFIN>="'+UsDateFin+'" AND'+
            ' PHB_NATURERUB="AAA"';
QRechRPreavis:=OpenSql(StRPreavis,TRUE);

if ((not QRechRPreavis.eof) and
   (QRechRPreavis.FindField('MTPREAVIS').AsFloat <> 0)) then
   begin
   Preavis := QRechRPreavis.FindField('MTPREAVIS').AsFloat;
   SetField ('PAS_INDPREAVIS',Preavis);
   end;
Ferme(QRechRPreavis);

StConv:='SELECT SUM(PHB_MTREM) AS MTCONV'+
        ' FROM HISTOBULLETIN'+
        ' LEFT JOIN REMUNERATION ON'+
        ' PHB_NATURERUB=PRM_NATURERUB AND'+
        ' ##PRM_PREDEFINI## PHB_RUBRIQUE=PRM_RUBRIQUE WHERE'+
        ' PRM_INDEMCONVENT="X" AND'+
        ' PHB_SALARIE="'+PGAttesSalarie+'" AND'+
        ' PHB_DATEFIN<="'+UsDerJour+'" AND'+
        ' PHB_DATEFIN>="'+UsDateFin+'" AND'+
        ' PHB_NATURERUB="AAA"';
QRechConv:=OpenSql(StConv,TRUE);

if ((not QRechConv.eof) and (QRechConv.FindField('MTCONV').AsFloat <> 0)) then
   begin
   Conv := QRechConv.FindField('MTCONV').AsFloat;
   SetField('PAS_INDCONV2', Conv);
   end;
Ferme(QRechConv);
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 13/07/2000
Modifié le ... :   /  /
Description .. : ChargeSalaires
Mots clefs ... : PAIE;PGATTESTATION;PGASSEDIC
*****************************************************************}
procedure TOM_Attestations.ChargeSalaires(ModifJour :Boolean);
var
StGrid, StPaye, StPayeCumSal, StPayele, UsDateDeb, UsDateFin : String;
QRechGrid, QRechPayele : TQuery;
Date1, Date2, DateD, DateF : TDateTime;
j, Ordre : integer;
Payele : array[1..12] of TDateTime;
Heures, Precompte, SommeBrut : array[1..12] of Double;
PremAnnee, PremMois : Word;
TAssietteBrut, TAssietteBrutD, TCP, TCPD, THistoCumSal, THistoCumSalD : TOB;
begin
For j:=1 to 12 do
    begin
    Payele[j]:= IDate1900;
    Heures[j]:= 0;
    Precompte[j]:= 0;
    SommeBrut[j]:= 0;
    Salaires.CellValues[0,j]:= DateToStr (IDate1900);
    Salaires.CellValues[1,j]:= DateToStr (IDate1900);
    Salaires.CellValues[2,j]:= DateToStr (IDate1900);
    Salaires.CellValues[3,j]:= '';
    Salaires.CellValues[4,j]:= '';
    Salaires.CellValues[5,j]:= '';
    Salaires.CellValues[6,j]:= '';
    Salaires.CellValues[7,j]:= '';
    end;

if (GetField ('PAS_DERNIERJOUR')<>IDate1900) then
   DerJour:= StrToDate (GetField ('PAS_DERNIERJOUR'));
NbreMois:= 0;
PremAnnee:= 0;
PremMois:= 0;
NbreMois:= 12;
Date1:= GetField ('PAS_PERIODEDEBUT');
if (DerJour<FinDeMois (DerJour)) then
   Date2:= FinDeMois (PlusMois (DerJour, -1))
else
   Date2:= DerJour;
NOMBREMOIS (DebutDeMois (Date1), FinDeMois (Date2), PremMois, PremAnnee, NbreMois);
if (NbreMois>12) then
   NbreMois:= 12;

UsDerJour:= UsDateTime (DerJour);
UsDateDeb:= UsDateTime (DebutDeMois (PlusMois (DerJour, -NbreMois)));
if (DerJour<FinDeMois (DerJour)) then
   UsDateFin:= UsDateTime (DebutDeMois (DerJour))
else
   UsDateFin:= UsDateTime (DebutDeMois (PlusMois (DerJour, 1)));

if ((DS.State=dsInsert) OR (ModifJour=TRUE)) AND (Salaires<>NIL) then
// chargement uniquement en mode creation attestation
   begin
   StPayele:= 'SELECT PPU_PAYELE, PPU_DATEDEBUT, PPU_DATEFIN'+
              ' FROM PAIEENCOURS WHERE'+
              ' PPU_SALARIE="'+PGAttesSalarie+'" AND'+
              ' PPU_DATEDEBUT >= "'+UsDateDeb+'" AND'+
              ' PPU_DATEDEBUT <= "'+UsDateFin+'" ORDER BY PPU_DATEDEBUT';
   QRechPayele:= OpenSql (StPayele, TRUE);
   QRechPayele.First;
   While Not QRechPayele.eof do
         begin
         DateF:= QRechPayele.FindField ('PPU_DATEFIN').AsDateTime;
         For j:=1 to NbreMois do
             begin
             if (DateF>=DebutDeMois (PlusMois (Date2, j-(NbreMois)))) and
                (DateF<=FinDeMois (PlusMois (Date2, j-(NbreMois)))) and
                (QRechPayele.FindField ('PPU_PAYELE').AsDateTime>=Payele[j]) then
                begin
                Payele[j]:= QRechPayele.FindField ('PPU_PAYELE').AsDateTime;
                break;
                end;
             end;
         QRechPayele.next;
         end;
   Ferme (QRechPayele);

   StPayeCumSal:= 'SELECT PHC_MONTANT, PHC_OMONTANT, PHC_DATEDEBUT,'+
                  ' PHC_DATEFIN, PHC_CUMULPAIE'+
                  ' FROM HISTOCUMSAL WHERE'+
                  ' PHC_SALARIE="'+PGAttesSalarie+'" AND'+
                  ' (PHC_CUMULPAIE="01" OR PHC_CUMULPAIE="20") AND'+
                  ' PHC_DATEDEBUT>="'+UsDateDeb+'" AND'+
                  ' PHC_DATEFIN<="'+UsDateFin+'" ORDER BY PHC_DATEDEBUT';
   THistoCumSal:= TOB.Create ('Les HistoCumSal', NIL, -1);
{PT47-2
   QRechPayeCumSal:= OpenSql (StPayeCumSal, TRUE);
   THistoCumSal.LoadDetailDB ('Les HistoCumSal', '', '', QRechPayeCumSal, False);
   Ferme (QRechPayeCumSal);
}
   THistoCumSal.LoadDetailDBFromSQL ('Les HistoCumSal', StPayeCumSal);
//FIN PT47-2   

   THistoCumSalD:= THistoCumSal.FindFirst (['PHC_CUMULPAIE'], ['20'], FALSE);
   while (THistoCumSalD<>nil) do
         begin
         DateF:= THistoCumSalD.GetValue ('PHC_DATEFIN');
         For j:=1 to NbreMois do
             begin
             if (DateF>=DebutDeMois (PlusMois (Date2, j-(NbreMois)))) and
                (DateF<=FinDeMois (PlusMois (Date2, j-(NbreMois)))) then
                begin
                Heures[j]:= Heures[j]+THistoCumSalD.GetValue ('PHC_MONTANT');
                break;
                end;
             end;
         THistoCumSalD:= THistoCumSal.FindNext (['PHC_CUMULPAIE'], ['20'], FALSE);
         end;

   StPaye:= 'SELECT PHB_BASECOT, PHB_OBASECOT, PHB_MTSALARIAL,'+
            ' PHB_OMTSALARIAL, PHB_DATEDEBUT, PHB_DATEFIN, PCT_ASSIETTEBRUT,'+
            ' PCT_PRECOMPTEASS'+
            ' FROM HISTOBULLETIN'+
            ' LEFT JOIN COTISATION ON'+
            ' PHB_RUBRIQUE=PCT_RUBRIQUE WHERE'+
            ' ##PCT_PREDEFINI## (PCT_ASSIETTEBRUT="X" OR'+
            ' PCT_PRECOMPTEASS="X") AND'+
            ' PHB_SALARIE="'+PGAttesSalarie+'" AND'+
            ' PHB_DATEDEBUT>="'+UsDateDeb+'" AND'+
            ' PHB_DATEFIN<="'+UsDateFin+'" AND'+
            ' PHB_NATURERUB<>"AAA" ORDER BY PHB_DATEDEBUT';
   TAssietteBrut:= TOB.Create ('Les AssietteBrut', NIL, -1);
{PT47-2
   QRechPaye:= OpenSql (StPaye, TRUE);
   TAssietteBrut.LoadDetailDB ('Les AssietteBrut', '', '', QRechPaye, False);
   Ferme (QRechPaye);
}
   TAssietteBrut.LoadDetailDBFromSQL ('Les AssietteBrut', StPaye);
//FIN PT47-2

   StPaye:= 'SELECT PHB_DATEDEBUT, PHB_DATEFIN, PHB_MTREM, PRM_INDEMPREAVIS,'+
            ' PRM_INDEMCOMPCP'+
            ' FROM HISTOBULLETIN'+
            ' LEFT JOIN REMUNERATION ON'+
            ' PHB_NATURERUB=PRM_NATURERUB AND'+
            ' ##PRM_PREDEFINI## PHB_RUBRIQUE=PRM_RUBRIQUE WHERE'+
            ' (PRM_INDEMPREAVIS="X" OR PRM_INDEMCOMPCP="X" OR'+
            ' PRM_PRIMEASSEDIC="X") AND'+
            ' PHB_SALARIE="'+PGAttesSalarie+'" AND'+
            ' PHB_DATEDEBUT>="'+UsDateDeb+'" AND'+
            ' PHB_DATEFIN<="'+UsDateFin+'" AND'+
            ' PHB_NATURERUB="AAA" AND'+
            ' PHB_MTREM<>0';
   TCP:= TOB.Create ('Les CP', NIL, -1);
{PT47-2
   QRechPaye:= OpenSql (StPaye, TRUE);
   TCP.LoadDetailDB ('Les CP', '', '', QRechPaye, False);
   Ferme (QRechPaye);
}
   TCP.LoadDetailDBFromSQL ('Les CP', StPaye);
//FIN PT47-2

   THistoCumSalD:= THistoCumSal.FindFirst (['PHC_CUMULPAIE'], ['01'], FALSE);
   while (THistoCumSalD<>nil) do
         begin
         DateD:= THistoCumSalD.GetValue ('PHC_DATEDEBUT');
         DateF:= THistoCumSalD.GetValue ('PHC_DATEFIN');
         TAssietteBrutD:= TAssietteBrut.FindFirst (['PHB_DATEDEBUT', 'PHB_DATEFIN', 'PCT_ASSIETTEBRUT'],
                                                   [DateD, DateF, 'X'], TRUE);
         if (CatDADS='003') then
            TAssietteBrutD:= nil;
         if (TAssietteBrutD<>nil) then //On n'est pas dans le cas de la reprise
            begin
            For j:=1 to NbreMois do
                begin
                if (DateF>=DebutDeMois (PlusMois (Date2, j-(NbreMois)))) and
                   (DateF<=FinDeMois (PlusMois (Date2, j-(NbreMois)))) then
                   begin
                   if (VH_Paie.PGTenueEuro=FALSE) then
                      SommeBrut[j]:= SommeBrut[j]+
                                     TAssietteBrutD.GetValue ('PHB_BASECOT')
                   else
                      begin
                      if (VH_Paie.PGDateBasculEuro<=DebutDeMois (PlusMois (Date2, j-(NbreMois)))) then
                         SommeBrut[j]:= SommeBrut[j]+
                                        TAssietteBrutD.GetValue ('PHB_BASECOT')
                      else
                         SommeBrut[j]:= SommeBrut[j]+
                                        TAssietteBrutD.GetValue ('PHB_OBASECOT');
                      end;

                   TCPD:= TCP.FindFirst (['PHB_DATEDEBUT', 'PHB_DATEFIN', 'PRM_INDEMPREAVIS'],
                                         [DateD, DateF, 'X'], TRUE);
                   if (TCPD<>nil) then
                      SommeBrut[j]:= SommeBrut[j]-TCPD.GetValue ('PHB_MTREM');
                   TCPD:= TCP.FindFirst (['PHB_DATEDEBUT', 'PHB_DATEFIN', 'PRM_INDEMCOMPCP'],
                                         [DateD, DateF, 'X'], TRUE);
                   if (TCPD<>nil) then
                      SommeBrut[j]:= SommeBrut[j]-TCPD.GetValue ('PHB_MTREM');
//PT47-1
                   TCPD:= TCP.FindFirst (['PHB_DATEDEBUT', 'PHB_DATEFIN', 'PRM_PRIMEASSEDIC'],
                                         [DateD, DateF, 'X'], TRUE);
                   if (TCPD<>nil) then
                      SommeBrut[j]:= SommeBrut[j]-TCPD.GetValue ('PHB_MTREM');
//FIN PT47-1
                   end;
                end;
            end
         else  //reprise ISIS 2
            begin
            For j:=1 to NbreMois do
                begin
                if (DateF>=DebutDeMois (PlusMois (Date2, j-(NbreMois)))) and
                   (DateF<=FinDeMois (PlusMois (Date2, j-(NbreMois)))) then
                   begin
                   if (VH_Paie.PGTenueEuro=FALSE) then
                      SommeBrut[j]:= SommeBrut[j]+
                                     THistoCumSalD.GetValue ('PHC_MONTANT')
                   else
                      begin
                      if (VH_Paie.PGDateBasculEuro<=DebutDeMois (PlusMois (Date2, j-(NbreMois)))) then
                         SommeBrut[j]:= SommeBrut[j]+
                                        THistoCumSalD.GetValue ('PHC_MONTANT')
                      else
                         SommeBrut[j]:= SommeBrut[j]+
                                        THistoCumSalD.GetValue ('PHC_OMONTANT');
                      end;

                   TCPD:= TCP.FindFirst (['PHB_DATEDEBUT', 'PHB_DATEFIN', 'PRM_INDEMPREAVIS'],
                                         [DateD, DateF, 'X'], TRUE);
                   if (TCPD<>nil) then
                      SommeBrut[j]:= SommeBrut[j]-TCPD.GetValue ('PHB_MTREM');
                   TCPD:= TCP.FindFirst (['PHB_DATEDEBUT', 'PHB_DATEFIN', 'PRM_INDEMCOMPCP'],
                                         [DateD, DateF, 'X'], TRUE);
                   if (TCPD<>nil) then
                      SommeBrut[j]:= SommeBrut[j]-TCPD.GetValue ('PHB_MTREM');
//PT47-1
                   TCPD:= TCP.FindFirst (['PHB_DATEDEBUT', 'PHB_DATEFIN', 'PRM_PRIMEASSEDIC'],
                                         [DateD, DateF, 'X'], TRUE);
                   if (TCPD<>nil) then
                      SommeBrut[j]:= SommeBrut[j]-TCPD.GetValue ('PHB_MTREM');
//FIN PT47-1
                   end;
                end;
            end;
         THistoCumSalD:= THistoCumSal.FindNext (['PHC_CUMULPAIE'], ['01'], TRUE);
         end;
   FreeAndNil (THistoCumSal);
   FreeAndNil (TCP);

   TAssietteBrutD:= TAssietteBrut.FindFirst (['PCT_PRECOMPTEASS'], ['X'], FALSE);
   While (TAssietteBrutD<>nil) do
         begin
         DateF:= TAssietteBrutD.GetValue ('PHB_DATEFIN');
         For j:=1 to NbreMois do
             begin
             if (DateF >= DebutDeMois (PlusMois (Date2, j-(NbreMois)))) and
                (DateF <= FinDeMois (PlusMois (Date2, j-(NbreMois)))) then
                if (VH_Paie.PGTenueEuro=FALSE) then
                   begin
                   Precompte[j]:= Precompte[j]+TAssietteBrutD.GetValue ('PHB_MTSALARIAL');
                   break;
                   end
                else
                   BEGIN
                   if (VH_Paie.PGDateBasculEuro <= DebutDeMois (PlusMois (Date2, j-(NbreMois)))) then
                      begin
                      Precompte[j]:= Precompte[j]+TAssietteBrutD.GetValue ('PHB_MTSALARIAL');
                      break;
                      end
                   else
                      begin
                      Precompte[j]:= Precompte[j]+TAssietteBrutD.GetValue ('PHB_OMTSALARIAL');
                      break;
                      end;
                   END;
             end;
         TAssietteBrutD:= TAssietteBrut.FindNext (['PCT_PRECOMPTEASS'], ['X'], FALSE);
         end;
   FreeAndNil (TAssietteBrut);

   For j:=1 to NbreMois do
       begin
       Salaires.CellValues[0,j]:= DateToStr (DebutDeMois (PlusMois (Date2, j-(NbreMois))));
       Salaires.CellValues[1,j]:= DateToStr (FinDeMois (PlusMois (Date2, j-(NbreMois))));
       if (Payele[j]> 0) then
          Salaires.CellValues[2,j]:= DateToStr (Payele[j]);
       Salaires.CellValues[3,j]:= DoubleToCell (Heures[j], 2);
       Salaires.CellValues[5,j]:= DoubleToCell (SommeBrut[j], 2);
       Salaires.CellValues[6,j]:= DoubleToCell (Precompte[j], 2);
       Salaires.CellValues[4,j]:= '0';
       Salaires.CellValues[7,j]:= '';
       end;
   end
else
   if (DS.State=dsBrowse) AND (Salaires<>NIL) then
// chargement uniquement en mode modification attestation
      begin
      Ordre:= GetField ('PAS_ORDRE');

      StGrid:= 'SELECT PAL_DATEDEBUT, PAL_DATEFIN, PAL_PAYELE, PAL_NBHEURES,'+
               ' PAL_JNONPAYES, PAL_SALAIRE, PAL_PRECOMPTE, PAL_MONNAIE,'+
               ' PAL_OBSERVATIONS, PAL_MOIS'+
               ' FROM ATTSALAIRES WHERE'+
               ' PAL_ORDRE='+IntToStr(Ordre)+' AND'+
               ' PAL_SALARIE="'+PGAttesSalarie+'" AND'+
               ' PAL_TYPEATTEST="ASS"';
      QRechGrid:= OpenSql (StGrid, TRUE);
      For j:=1 to 12 do
          begin
          QRechGrid.First;
          While not QRechGrid.Eof do
                begin
                if (QRechGrid.FindField ('PAL_MOIS').AsInteger=j) then
                   begin
                   Salaires.CellValues[0,j]:= QRechGrid.FindField ('PAL_DATEDEBUT').Asstring;
                   Salaires.CellValues[1,j]:= QRechGrid.FindField ('PAL_DATEFIN').Asstring;
                   Salaires.CellValues[2,j]:= QRechGrid.FindField ('PAL_PAYELE').Asstring;
                   Salaires.CellValues[3,j]:= FloatToStr (QRechGrid.FindField ('PAL_NBHEURES').AsFloat);
                   Salaires.CellValues[4,j]:= FloatToStr (QRechGrid.FindField ('PAL_JNONPAYES').AsFloat);
                   Salaires.CellValues[5,j]:= FloatToStr (QRechGrid.FindField ('PAL_SALAIRE').AsFloat);
                   Salaires.CellValues[6,j]:= FloatToStr (QRechGrid.FindField ('PAL_PRECOMPTE').AsFloat);
                   Salaires.CellValues[7,j]:= QRechGrid.FindField ('PAL_OBSERVATIONS').Asstring;
                   end;
                QRechGrid.Next;
                end;
          end;
      Ferme (QRechGrid);
      end;
(*Mise en commentaire PT45
{$IFDEF EAGLCLIENT}
      ESalaires;
{$ENDIF}
*)
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 13/07/2000
Modifié le ... :   /  /
Description .. : ChargePrimes
Mots clefs ... : PAIE;PGATTESTATION;PGASSEDIC
*****************************************************************}
procedure TOM_Attestations.ChargePrimes(ModifJour :Boolean);
var
StGrid, StPaye, UsDateDeb, UsDateFin : String;
QRechGrid, QRechPaye : TQuery;
i, j, Ordre : integer;
PremAnnee, PremMois : Word;
Date1, Date2 : TDateTime;
begin
//PT47-1
For i:=1 to 3 do
    begin
    Primes.CellValues[0,i]:= '01/01/1900';
    Primes.CellValues[1,i]:= '01/01/1900';
    Primes.CellValues[2,i]:= '';
    Primes.CellValues[3,i]:= '01/01/1900';
    Primes.CellValues[4,i]:= '';
    end;
//FIN PT47-1

{PT47-1
if ((DS.State=dsBrowse) OR (ModifJour=FALSE)) AND (Primes<>NIL) then
}
if (GetField ('PAS_DERNIERJOUR')<>IDate1900) then
   DerJour:= StrToDate (GetField ('PAS_DERNIERJOUR'));
NbreMois:= 0;
PremAnnee:= 0;
PremMois:= 0;
NbreMois:= 12;
Date1:= GetField ('PAS_PERIODEDEBUT');
if (DerJour<FinDeMois (DerJour)) then
   Date2:= FinDeMois (PlusMois (DerJour, -1))
else
   Date2:= DerJour;
NOMBREMOIS (DebutDeMois (Date1), FinDeMois (Date2), PremMois, PremAnnee, NbreMois);
if (NbreMois>12) then
   NbreMois:= 12;

UsDerJour:= UsDateTime (DerJour);
UsDateDeb:= UsDateTime (DebutDeMois (PlusMois (DerJour, -NbreMois)));
if (DerJour<FinDeMois (DerJour)) then
   UsDateFin:= UsDateTime (DebutDeMois (DerJour))
else
   UsDateFin:= UsDateTime (DebutDeMois (PlusMois (DerJour, 1)));

if ((DS.State=dsInsert) OR (ModifJour=TRUE)) AND (Primes<>NIL) then
// chargement uniquement en mode creation attestation
   begin
   StPaye:= 'SELECT PHB_DATEDEBUT, PHB_DATEFIN, PHB_LIBELLE, PPU_PAYELE,'+
            ' PHB_MTREM, PRM_PRIMEASSEDIC'+
            ' FROM HISTOBULLETIN'+
            ' LEFT JOIN REMUNERATION ON'+
            ' PHB_NATURERUB=PRM_NATURERUB AND'+
            ' ##PRM_PREDEFINI## PHB_RUBRIQUE=PRM_RUBRIQUE'+
            ' LEFT JOIN PAIEENCOURS ON'+
            ' PHB_SALARIE=PPU_SALARIE AND'+
            ' PHB_DATEDEBUT=PPU_DATEDEBUT AND'+
            ' PHB_DATEFIN=PPU_DATEFIN WHERE'+
            ' PRM_PRIMEASSEDIC="X" AND'+
            ' PHB_SALARIE="'+PGAttesSalarie+'" AND'+
            ' PHB_DATEDEBUT>="'+UsDateDeb+'" AND'+
            ' PHB_DATEFIN<="'+UsDateFin+'" AND'+
            ' PHB_NATURERUB="AAA" AND'+
            ' PHB_MTREM<>0';
  QRechPaye:= OpenSql (StPaye, TRUE);
  if (NOT QRechPaye.EOF) then
     begin
     i:= 1;
     QRechPaye.First;
     While not QRechPaye.Eof do
           begin
           if (i>3) then
              break;
           Primes.CellValues[0,i]:= QRechPaye.FindField ('PHB_DATEDEBUT').Asstring;
           Primes.CellValues[1,i]:= QRechPaye.FindField ('PHB_DATEFIN').Asstring;
           Primes.CellValues[2,i]:= QRechPaye.FindField ('PHB_LIBELLE').Asstring;
           Primes.CellValues[3,i]:= QRechPaye.FindField ('PPU_PAYELE').Asstring;
           Primes.CellValues[4,i]:= DoubleToCell (QRechPaye.FindField ('PHB_MTREM').AsFloat, 2);
           QRechPaye.Next;
           i:= i+1;
           end;
     end;
   Ferme (QRechPaye);
   end
else
if (DS.State=dsBrowse) AND (Salaires<>NIL) then
// chargement uniquement en mode modification attestation
//FIN PT47-1
   begin
   Ordre:= GetField ('PAS_ORDRE');

   StGrid:= 'SELECT PAL_DATEDEBUT, PAL_DATEFIN, PAL_OBSERVATIONS, PAL_PAYELE,'+
            ' PAL_SALAIRE, PAL_MONNAIE, PAL_MOIS'+
            ' FROM ATTSALAIRES WHERE'+
            ' PAL_ORDRE='+IntToStr(Ordre)+' AND'+
            ' PAL_SALARIE="'+PGAttesSalarie+'" AND'+
            ' PAL_TYPEATTEST="ASS"';
   QRechGrid:= OpenSql (StGrid, TRUE);
   For j:=13 to 15 do
       begin
       i:= j-12;
       QRechGrid.First;
       While not QRechGrid.Eof do
             begin
             if (QRechGrid.FindField ('PAL_MOIS').AsInteger=j) then
                begin
                if (QRechGrid.FindField ('PAL_DATEDEBUT').Asstring<>'') then
                   Primes.CellValues[0,i]:= QRechGrid.FindField ('PAL_DATEDEBUT').Asstring;
{PT47-1
                else
                   Primes.CellValues[0,i]:= '01/01/1900';
}
                if (QRechGrid.FindField ('PAL_DATEFIN').Asstring<>'') then
                   Primes.CellValues[1,i]:= QRechGrid.FindField ('PAL_DATEFIN').Asstring;
{PT47-1
                else
                   Primes.CellValues[1,i]:= '01/01/1900';
}
                Primes.CellValues[2,i]:= QRechGrid.FindField ('PAL_OBSERVATIONS').Asstring;
                if (QRechGrid.FindField ('PAL_PAYELE').Asstring<>'') then
                   Primes.CellValues[3,i]:= QRechGrid.FindField ('PAL_PAYELE').Asstring;
{PT47-1
                else
                   Primes.CellValues[3,i]:= '01/01/1900';
}
                Primes.CellValues[4,i]:= DoubleToCell (QRechGrid.FindField ('PAL_SALAIRE').AsFloat, 2);
                end;
             QRechGrid.Next;
             end;
       end;
   Ferme (QRechGrid);
   end;
{PT47-1
else
   if (ModifJour = TRUE) then
      BEGIN
      For i:=1 to 3 do
          begin
          Primes.CellValues[0,i]:= '01/01/1900';
          Primes.CellValues[1,i]:= '01/01/1900';
          Primes.CellValues[2,i]:= '';
          Primes.CellValues[3,i]:= '01/01/1900';
          Primes.CellValues[4,i]:= '';
          end;
      END;
}
(*Mise en commentaire PT45
{$IFDEF EAGLCLIENT}
      EPrimes;
{$ENDIF}
*)
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 13/07/2000
Modifié le ... :   /  /
Description .. : ChargeSolde
Mots clefs ... : PAIE;PGATTESTATION;PGASSEDIC
*****************************************************************}
procedure TOM_Attestations.ChargeSolde(ModifJour :Boolean);
var
StGrid, StPaye, StPayeCumSal, StPayele : String;
UsDateDeb, UsDateFin : String;
QRechGrid, QRechPayele : TQuery;
i, j, Ordre : integer;
Date1, Date2, DateD, DateF : TDateTime;
PayeLe : array[1..2] of TDateTime;
Heures, Precompte, SommeBrut : array[1..2] of Double;
TAssietteBrut, TAssietteBrutD, TCP, TCPD, THistoCumSal, THistoCumSalD : TOB;
begin
For j:=1 to 2 do
    begin
    Payele[j]:= 0;
    Heures[j]:= 0;
    Precompte[j]:= 0;
    SommeBrut[j]:= 0;
    end;

if (GetField ('PAS_DERNIERJOUR')<>IDate1900) then
   begin
   DerJour:= StrToDate (GetField ('PAS_DERNIERJOUR'));
   if (GetField ('PAS_PERIODEFIN')>GetField ('PAS_DERNIERJOUR')) then
      begin
      Date2:= FinDeMois (StrToDate (GetField ('PAS_PERIODEFIN')));
      UsDateFin:= UsDateTime (Date2);
      end
   else
      begin
      Date2:= FinDeMois (DerJour);
      UsDateFin:= UsDateTime (Date2);
      end;
   end;
UsDerJour:= UsDateTime (DerJour);
Date1:= DebutDeMois (DerJour);
UsDateDeb:= UsDateTime (Date1);

For j:=1 to 2 do
    begin
    Solde.CellValues[0,j]:= '01/01/1900';
    Solde.CellValues[1,j]:= '01/01/1900';
    Solde.CellValues[2,j]:= '01/01/1900';
    Solde.CellValues[3,j]:= '0';
    Solde.CellValues[4,j]:= '0';
    Solde.CellValues[5,j]:= '0';
    end;

if ((DS.State=dsInsert) OR (ModifJour=TRUE)) AND (Solde<>NIL) then
// chargement uniquement en mode creation attestation
   begin
   if (DerJour<Date2) then
      begin
      StPayele:= 'SELECT PPU_PAYELE, PPU_DATEDEBUT, PPU_DATEFIN'+
                 ' FROM PAIEENCOURS WHERE'+
                 ' PPU_SALARIE="'+PGAttesSalarie+'" AND'+
                 ' PPU_DATEDEBUT>="'+UsDateDeb+'" AND'+
                 ' PPU_DATEDEBUT<="'+UsDateFin+'" ORDER BY PPU_DATEDEBUT';
      QRechPayele:= OpenSql (StPayele, TRUE);
      QRechPayele.First;
      while NOT QRechPayele.EOF do
            begin
            DateD:= QRechPayele.FindField ('PPU_DATEDEBUT').AsDateTime;
            DateF:= QRechPayele.FindField ('PPU_DATEFIN').AsDateTime;
            For j:=1 to 2 do
                begin
                if (DateF>=DebutDeMois (PlusMois (Date2, j-2))) and
                   (DateF<=FinDeMois (PlusMois (Date2, j-2))) and
                   (QRechPayele.FindField ('PPU_PAYELE').AsDateTime>=Payele[j]) then
                   begin
                   Solde.CellValues[0,j]:= DateToStr (DateD);
                   Solde.CellValues[1,j]:= DateToStr (DateF);
                   Payele[j]:= QRechPayele.FindField ('PPU_PAYELE').AsDateTime;
                   break;
                   end;
                end;
            QRechPayele.next;
            end;
      Ferme (QRechPayele);

      StPayeCumSal:= 'SELECT PHC_MONTANT, PHC_OMONTANT, PHC_DATEDEBUT,'+
                     ' PHC_DATEFIN, PHC_CUMULPAIE'+
                     ' FROM HISTOCUMSAL WHERE'+
                     ' PHC_SALARIE="'+PGAttesSalarie+'" AND'+
                     ' (PHC_CUMULPAIE="01" OR PHC_CUMULPAIE="20") AND'+
                     ' PHC_DATEDEBUT>="'+UsDateDeb+'" AND'+
                     ' PHC_DATEFIN<="'+UsDateFin+'"';
      THistoCumSal:= TOB.Create ('Les HistoCumSal', NIL, -1);
{PT47-2
      QRechPayeCumSal:= OpenSql (StPayeCumSal, TRUE);
      THistoCumSal.LoadDetailDB ('Les HistoCumSal', '', '', QRechPayeCumSal, False);
      Ferme (QRechPayeCumSal);
}
      THistoCumSal.LoadDetailDBFromSQL ('Les HistoCumSal', StPayeCumSal);
//FIN PT47-2

      THistoCumSalD:= THistoCumSal.FindFirst (['PHC_CUMULPAIE'], ['20'], FALSE);
      while (THistoCumSalD<>nil) do
            begin
            DateF:= THistoCumSalD.GetValue ('PHC_DATEFIN');
            For j:=1 to 2 do
                begin
                if (DateF>=DebutDeMois (PlusMois (Date2, j-2))) and
                   (DateF<=FinDeMois (PlusMois (Date2, j-2))) then
                   begin
                   Heures[j]:= Heures[j]+THistoCumSalD.GetValue ('PHC_MONTANT');
                   break;
                   end;
                end;
            THistoCumSalD:= THistoCumSal.FindNext (['PHC_CUMULPAIE'], ['20'], FALSE);
            end;

      StPaye:= 'SELECT PHB_BASECOT, PHB_OBASECOT, PHB_MTSALARIAL,'+
               ' PHB_OMTSALARIAL, PHB_DATEDEBUT, PHB_DATEFIN,'+
               ' PCT_ASSIETTEBRUT, PCT_PRECOMPTEASS'+
               ' FROM HISTOBULLETIN'+
               ' LEFT JOIN COTISATION ON'+
               ' PHB_RUBRIQUE=PCT_RUBRIQUE WHERE'+
               ' ##PCT_PREDEFINI## (PCT_ASSIETTEBRUT="X" OR'+
               ' PCT_PRECOMPTEASS="X") AND'+
               ' PHB_SALARIE="'+PGAttesSalarie+'" AND'+
               ' PHB_DATEDEBUT>="'+UsDateDeb+'" AND'+
               ' PHB_DATEFIN<="'+UsDateFin+'" AND'+
               ' PHB_NATURERUB <> "AAA"';
      TAssietteBrut:= TOB.Create ('Les AssietteBrut', NIL, -1);
{PT47-2
      QRechPaye:= OpenSql (StPaye, TRUE);
      TAssietteBrut.LoadDetailDB ('Les AssietteBrut', '', '', QRechPaye, False);
      Ferme (QRechPaye);
}
      TAssietteBrut.LoadDetailDBFromSQL ('Les AssietteBrut', StPaye);
//FIN PT47-2

      StPaye:= 'SELECT PHB_DATEDEBUT, PHB_DATEFIN, PHB_MTREM,'+
               ' PRM_INDEMPREAVIS, PRM_INDEMCOMPCP'+
               ' FROM HISTOBULLETIN'+
               ' LEFT JOIN REMUNERATION ON'+
               ' PHB_NATURERUB=PRM_NATURERUB AND'+
               ' ##PRM_PREDEFINI## PHB_RUBRIQUE=PRM_RUBRIQUE WHERE'+
               ' (PRM_INDEMPREAVIS="X" OR PRM_INDEMCOMPCP="X") AND'+
               ' PHB_SALARIE="'+PGAttesSalarie+'" AND'+
               ' PHB_DATEDEBUT>="'+UsDateDeb+'" AND'+
               ' PHB_DATEFIN<="'+UsDateFin+'" AND'+
               ' PHB_NATURERUB="AAA"';
      TCP:= TOB.Create ('Les CP', NIL, -1);
{PT47-2
      QRechPaye:= OpenSql (StPaye, TRUE);
      TCP.LoadDetailDB ('Les CP', '', '', QRechPaye, False);
      Ferme (QRechPaye);
}
      TCP.LoadDetailDBFromSQL ('Les CP', StPaye);
//FIN PT47-2

      THistoCumSalD:= THistoCumSal.FindFirst (['PHC_CUMULPAIE'], ['01'], FALSE);
      while (THistoCumSalD<>nil) do
            begin
            DateD:= THistoCumSalD.GetValue ('PHC_DATEDEBUT');
            DateF:= THistoCumSalD.GetValue ('PHC_DATEFIN');
            TAssietteBrutD:= TAssietteBrut.FindFirst (['PHB_DATEDEBUT', 'PHB_DATEFIN', 'PCT_ASSIETTEBRUT'],
                                                      [DateD, DateF, 'X'], TRUE);
            if (CatDADS='003') then
               TAssietteBrutD:= nil;
            if (TAssietteBrutD<>nil) then //On n'est pas dans le cas de la reprise
               begin
               For j:=1 to 2 do
                   begin
                   if (DateF>=DebutDeMois (PlusMois (Date2, j-2))) and
                      (DateF<=FinDeMois (PlusMois (Date2, j-2))) then
                      begin
                      if (VH_Paie.PGTenueEuro=FALSE) then
                         SommeBrut[j]:= SommeBrut[j]+
                                        TAssietteBrutD.GetValue ('PHB_BASECOT')
                      else
                         begin
                         if (VH_Paie.PGDateBasculEuro<=DebutDeMois (PlusMois (Date2, j-2))) then
                            SommeBrut[j]:= SommeBrut[j]+
                                           TAssietteBrutD.GetValue ('PHB_BASECOT')
                         else
                            SommeBrut[j]:= SommeBrut[j]+
                                           TAssietteBrutD.GetValue ('PHB_OBASECOT');
                         end;

                      TCPD:= TCP.FindFirst (['PHB_DATEDEBUT', 'PHB_DATEFIN', 'PRM_INDEMPREAVIS'],
                                            [DateD, DateF, 'X'], TRUE);
                      if (TCPD<>nil) then
                         SommeBrut[j]:= SommeBrut[j]-TCPD.GetValue ('PHB_MTREM');
                      TCPD:= TCP.FindFirst (['PHB_DATEDEBUT', 'PHB_DATEFIN', 'PRM_INDEMCOMPCP'],
                                            [DateD, DateF, 'X'], TRUE);
                      if (TCPD<>nil) then
                         SommeBrut[j]:= SommeBrut[j]-TCPD.GetValue ('PHB_MTREM');
                      end;
                   end;
               end
            else  //reprise ISIS 2
               begin
               For j:=1 to 2 do
                   begin
                   if (DateF>=DebutDeMois (PlusMois (Date2, j-2))) and
                      (DateF<=FinDeMois (PlusMois (Date2, j-2))) then
                      begin
                      if (VH_Paie.PGTenueEuro=FALSE) then
                         SommeBrut[j]:= SommeBrut[j]+
                                        THistoCumSalD.GetValue ('PHC_MONTANT')
                      else
                         begin
                         if (VH_Paie.PGDateBasculEuro<=DebutDeMois (PlusMois (Date2, j-2))) then
                            SommeBrut[j]:= SommeBrut[j]+
                                           THistoCumSalD.GetValue ('PHC_MONTANT')
                         else
                            SommeBrut[j]:= SommeBrut[j]+
                                           THistoCumSalD.GetValue ('PHC_OMONTANT');
                         end;

                      TCPD:= TCP.FindFirst (['PHB_DATEDEBUT', 'PHB_DATEFIN', 'PRM_INDEMPREAVIS'],
                                            [DateD, DateF, 'X'], TRUE);
                      if (TCPD<>nil) then
                         SommeBrut[j]:= SommeBrut[j]-TCPD.GetValue ('PHB_MTREM');
                      TCPD:= TCP.FindFirst (['PHB_DATEDEBUT', 'PHB_DATEFIN', 'PRM_INDEMCOMPCP'],
                                            [DateD, DateF, 'X'], TRUE);
                      if (TCPD<>nil) then
                         SommeBrut[j]:= SommeBrut[j]-TCPD.GetValue ('PHB_MTREM');
                      end;
                   end;
               end;
            THistoCumSalD:= THistoCumSal.FindNext (['PHC_CUMULPAIE'], ['01'], TRUE);
            end;

      FreeAndNil (THistoCumSal);
      FreeAndNil (TCP);

      TAssietteBrutD:= TAssietteBrut.FindFirst (['PCT_PRECOMPTEASS'], ['X'], FALSE);
      While (TAssietteBrutD<>nil) do
            begin
            DateF:= TAssietteBrutD.GetValue ('PHB_DATEFIN');
            For j:=1 to 2 do
                begin
                if (DateF>=DebutDeMois (PlusMois (Date2, j-2))) and
                   (DateF<=FinDeMois (PlusMois (Date2, j-2))) then
                   if (VH_Paie.PGTenueEuro=FALSE) then
                      begin
                      Precompte[j]:= Precompte[j]+TAssietteBrutD.GetValue ('PHB_MTSALARIAL');
                      break;
                      end
                   else
                      BEGIN
                      if (VH_Paie.PGDateBasculEuro<=DebutDeMois (PlusMois (Date2, j-2))) then
                         begin
                         Precompte[j]:= Precompte[j]+TAssietteBrutD.GetValue ('PHB_MTSALARIAL');
                         break;
                         end
                      else
                         begin
                         Precompte[j]:= Precompte[j]+TAssietteBrutD.GetValue ('PHB_OMTSALARIAL');
                         break;
                         end;
                      END;
                end;
            TAssietteBrutD:= TAssietteBrut.FindNext (['PCT_PRECOMPTEASS'], ['X'], FALSE);
            end;
      FreeAndNil (TAssietteBrut);
      For j:=1 to 2 do
          begin
          if (Payele[j]> 0) then
             Solde.CellValues[2,j]:= DateToStr (Payele[j]);
          Solde.CellValues[3,j]:= DoubleToCell (Heures[j],2);
          Solde.CellValues[4,j]:= DoubleToCell (SommeBrut[j],2);
          Solde.CellValues[5,j]:= DoubleToCell (Precompte[j],2);
          end;
      end;
   end
else
   if (DS.State=dsBrowse) AND (Solde<>NIL) then
// chargement uniquement en mode modification attestation
      begin
      Ordre:= GetField ('PAS_ORDRE');
      StGrid:= 'SELECT PAL_DATEDEBUT, PAL_DATEFIN, PAL_PAYELE, PAL_NBHEURES,'+
               ' PAL_SALAIRE, PAL_PRECOMPTE, PAL_MONNAIE, PAL_MOIS'+
               ' FROM ATTSALAIRES WHERE'+
               ' PAL_ORDRE='+IntToStr(Ordre)+' AND'+
               ' PAL_SALARIE="'+PGAttesSalarie+'" AND'+
               ' PAL_TYPEATTEST="ASS"';
      QRechGrid:= OpenSql (StGrid,TRUE);
      For j:=16 to 17 do
          begin
          i:= j-15;
          QRechGrid.First;
          While not QRechGrid.Eof do
                begin
                if (QRechGrid.FindField ('PAL_MOIS').AsInteger=j) then
                   begin
                   if (QRechGrid.FindField ('PAL_DATEDEBUT').Asstring<>'') then
                      Solde.CellValues[0,i]:= QRechGrid.FindField ('PAL_DATEDEBUT').Asstring
                   else
                      Solde.CellValues[0,i]:= '01/01/1900';
                   if (QRechGrid.FindField ('PAL_DATEFIN').Asstring<>'') then
                      Solde.CellValues[1,i]:= QRechGrid.FindField ('PAL_DATEFIN').Asstring
                   else
                      Solde.CellValues[1,i]:= '01/01/1900';
                   if (QRechGrid.FindField ('PAL_PAYELE').Asstring<>'') then
                      Solde.CellValues[2,i]:= QRechGrid.FindField ('PAL_PAYELE').Asstring
                   else
                      Solde.CellValues[2,i]:= '01/01/1900';
                   Solde.CellValues[3,i]:= FloatToStr (QRechGrid.FindField ('PAL_NBHEURES').AsFloat); //DB2
                   Solde.CellValues[4,i]:= FloatToStr (QRechGrid.FindField ('PAL_SALAIRE').AsFloat);  //DB2
                   Solde.CellValues[5,i]:= FloatToStr (QRechGrid.FindField ('PAL_PRECOMPTE').AsFloat);     //DB2
                   end;
                QRechGrid.Next;
                end;
          end;
          Ferme(QRechGrid);
      end;
(*Mise en commentaire PT45
{$IFDEF EAGLCLIENT}
      ESolde;
{$ENDIF}
*)
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 13/07/2000
Modifié le ... :   /  /
Description .. : EmpPubClick
Mots clefs ... : PAIE;PGATTESTATION;PGASSEDIC
*****************************************************************}
procedure TOM_Attestations.EmpPubClick(Sender: TObject);
begin
//PT48
if (Suppr=True) then
   exit;
//FIN PT48

if EmpPub <> NIL then
   if EmpPub.Checked = TRUE then
      begin
      SetControlVisible ('PAS_PUBLICTYPE', TRUE);
      SetControlVisible ('LABEL_STATUT', TRUE);
      SetControlVisible ('PAS_COLLTERRITOR', TRUE);
      end
   else
      begin
      SetControlVisible ('PAS_PUBLICTYPE', FALSE);
      SetControlVisible ('LABEL_STATUT', FALSE);
      SetControlVisible ('PAS_COLLTERRITOR', FALSE);
      if (DS.State <> dsBrowse) then
         begin
         SetField('PAS_PUBLICTYPE','');
         SetField('PAS_COLLTERRITOR','');
         end;
      end;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 06/02/2004
Modifié le ... :   /  /
Description .. : PreavisClick
Mots clefs ... : PAIE;PGATTESTATION;PGASSEDIC
*****************************************************************}
procedure TOM_Attestations.PreavisClick(Sender: TObject);
begin
//PT48
if (Suppr=True) then
   exit;
//FIN PT48

if (Tcheckbox(Sender).name='PAS_MOTIFPREAVIS1') then
   begin
   if Preavis1 <> NIL then
      if Preavis1.Checked = TRUE then
         begin
         SetControlVisible ('TPAS_DATEARRET', TRUE);
         SetControlVisible ('PAS_DATEARRET', TRUE);
         SetControlVisible ('TPAS_REPRISEARRET', TRUE);
         SetControlVisible ('PAS_REPRISEARRET', TRUE);
         end
      else
         begin
         SetControlVisible ('TPAS_DATEARRET', FALSE);
         SetControlVisible ('PAS_DATEARRET', FALSE);
         SetControlVisible ('TPAS_REPRISEARRET', FALSE);
         SetControlVisible ('PAS_REPRISEARRET', FALSE);
         SetField ('PAS_DATEARRET',IDate1900);
         SetField ('PAS_REPRISEARRET',IDate1900);
         end;
   end;

if (Tcheckbox(Sender).name='PAS_MOTIFPREAVIS2') then
   begin
   if Preavis2 <> NIL then
      if Preavis2.Checked = TRUE then
         begin
         SetControlVisible ('TPAS_DATEPREAVISD2', TRUE);
         SetControlVisible ('PAS_DATEPREAVISD2', TRUE);
         SetControlVisible ('TPAS_DATEPREAVISF2', TRUE);
         SetControlVisible ('PAS_DATEPREAVISF2', TRUE);
         end
      else
         begin
         SetControlVisible ('TPAS_DATEPREAVISD2', FALSE);
         SetControlVisible ('PAS_DATEPREAVISD2', FALSE);
         SetControlVisible ('TPAS_DATEPREAVISF2', FALSE);
         SetControlVisible ('PAS_DATEPREAVISF2', FALSE);
         SetField ('PAS_DATEPREAVISD2',IDate1900);
         SetField ('PAS_DATEPREAVISF2',IDate1900);
         end;
   end;

if (Tcheckbox(Sender).name='PAS_MOTIFPREAVIS3') then
   begin
   if Preavis3 <> NIL then
      if Preavis3.Checked = TRUE then
         begin
         SetControlVisible ('TPAS_DATEPREAVISD3', TRUE);
         SetControlVisible ('PAS_DATEPREAVISD3', TRUE);
         SetControlVisible ('TPAS_DATEPREAVISF3', TRUE);
         SetControlVisible ('PAS_DATEPREAVISF3', TRUE);
         end
      else
         begin
         SetControlVisible ('TPAS_DATEPREAVISD3', FALSE);
         SetControlVisible ('PAS_DATEPREAVISD3', FALSE);
         SetControlVisible ('TPAS_DATEPREAVISF3', FALSE);
         SetControlVisible ('PAS_DATEPREAVISF3', FALSE);
         SetField ('PAS_DATEPREAVISD3',IDate1900);
         SetField ('PAS_DATEPREAVISF3',IDate1900);
         end;
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 21/01/2005
Modifié le ... :   /  /
Description .. : AlsaceClick
Mots clefs ... : PAIE;PGATTESTATION;PGASSEDIC
*****************************************************************}
procedure TOM_Attestations.AlsaceClick(Sender: TObject);
begin
//PT48
if (Suppr=True) then
   exit;
//FIN PT48

if Alsace <> NIL then
   if Alsace.Checked = TRUE then
      begin
      SetField('PAS_REGIMESPECSS', '-');
      SetControlEnabled('PAS_REGIMESPECSS', False);
      end
   else
      begin
      SetControlEnabled('PAS_REGIMESPECSS', True);
      end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 13/07/2000
Modifié le ... :   /  /
Description .. : RegimeSSClick
Mots clefs ... : PAIE;PGATTESTATION;PGASSEDIC
*****************************************************************}
procedure TOM_Attestations.RegimeSSClick(Sender: TObject);
begin
//PT48
if (Suppr=True) then
   exit;
//FIN PT48

if RegimeSS <> NIL then
   if RegimeSS.Checked = TRUE then
      begin
      SetControlVisible ('LABEL_CAISSESS', TRUE);
      SetControlVisible ('PAS_CAISSESS', TRUE);
      SetControlVisible ('LABEL_NOMATRICULE', TRUE);
      SetControlVisible ('PAS_MATRICULE', TRUE);
      SetField('PAS_ALSACEMOSEL', '-');
      SetControlEnabled('PAS_ALSACEMOSEL', False);
      end
   else
      begin
      SetControlVisible ('LABEL_CAISSESS', FALSE);
      SetControlVisible ('PAS_CAISSESS', FALSE);
      SetControlVisible ('LABEL_NOMATRICULE', FALSE);
      SetControlVisible ('PAS_MATRICULE', FALSE);
      SetField ('PAS_CAISSESS','');
      SetField ('PAS_MATRICULE','');
      SetControlEnabled('PAS_ALSACEMOSEL', True);
      end;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 13/07/2000
Modifié le ... :   /  /
Description .. : PubTypeChange
Mots clefs ... : PAIE;PGATTESTATION;PGASSEDIC
*****************************************************************}
procedure TOM_Attestations.PubTypeChange();
begin
//PT48
if (Suppr=True) then
   exit;
//FIN PT48

if GetField('PAS_PUBLICTYPE') = '' then
   begin
   SetControlVisible ('LABEL_CONVNUM', FALSE);
   SetControlVisible ('PAS_CONVNUMERO', FALSE);
   SetControlVisible ('LABEL_CONVANA', FALSE);
   SetControlVisible ('PAS_CONVCODEANA', FALSE);
   SetControlVisible ('PAS_ADHESION', FALSE);
   SetControlVisible ('LABEL_COLLTERRDATE', FALSE);
   SetControlVisible ('PAS_DATEADHESION', FALSE);
   SetField ('PAS_CONVNUMERO','');
   SetField ('PAS_CONVCODEANA','');
   SetField ('PAS_ADHESION','-');
   SetField ('PAS_DATEADHESION',IDate1900);
   end;

if GetField('PAS_PUBLICTYPE') = '001' then
   begin
   SetControlVisible ('LABEL_CONVNUM', FALSE);
   SetControlVisible ('PAS_CONVNUMERO', FALSE);
   SetControlVisible ('LABEL_CONVANA', FALSE);
   SetControlVisible ('PAS_CONVCODEANA', FALSE);
   SetControlVisible ('PAS_ADHESION', TRUE);
   SetControlVisible ('LABEL_COLLTERRDATE', FALSE);
   SetControlVisible ('PAS_DATEADHESION', FALSE);
   SetField ('PAS_CONVNUMERO','');
   SetField ('PAS_CONVCODEANA','');
   SetField ('PAS_DATEADHESION',IDate1900);
   end;

if GetField('PAS_PUBLICTYPE') = '002' then
   begin
   SetControlVisible ('LABEL_CONVNUM', TRUE);
   SetControlVisible ('PAS_CONVNUMERO', TRUE);
   SetControlVisible ('LABEL_CONVANA', TRUE);
   SetControlVisible ('PAS_CONVCODEANA', TRUE);
   SetControlVisible ('PAS_ADHESION', TRUE);
   SetControlVisible ('LABEL_COLLTERRDATE', FALSE);
   SetControlVisible ('PAS_DATEADHESION', FALSE);
   SetField ('PAS_DATEADHESION',IDate1900);
   end;

if GetField('PAS_PUBLICTYPE') = '003' then
   begin
   SetControlVisible ('LABEL_CONVNUM', FALSE);
   SetControlVisible ('PAS_CONVNUMERO', FALSE);
   SetControlVisible ('LABEL_CONVANA', FALSE);
   SetControlVisible ('PAS_CONVCODEANA', FALSE);
   SetControlVisible ('PAS_ADHESION', FALSE);
   SetControlVisible ('LABEL_COLLTERRDATE', TRUE);
   SetControlVisible ('PAS_DATEADHESION', TRUE);
   SetField ('PAS_CONVNUMERO','');
   SetField ('PAS_CONVCODEANA','');
   SetField ('PAS_ADHESION','-');
   end;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 13/07/2000
Modifié le ... :   /  /
Description .. : LienParenteClick
Mots clefs ... : PAIE;PGATTESTATION;PGASSEDIC
*****************************************************************}
procedure TOM_Attestations.LienParenteClick(Sender: TObject);
begin
//PT48
if (Suppr=True) then
   exit;
//FIN PT48

if LienParente <> NIL then
   if LienParente.Checked = TRUE then
      begin
      SetControlVisible ('LABELPARENT2', TRUE);
      SetControlVisible ('PAS_LIBLIENP', TRUE);
      end
   else
      begin
      SetControlVisible ('LABELPARENT2', FALSE);
      SetControlVisible ('PAS_LIBLIENP', FALSE);
      SetField ('PAS_LIBLIENP','');
      end;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 13/07/2000
Modifié le ... :   /  /
Description .. : AgircChange
Mots clefs ... : PAIE;PGATTESTATION;PGASSEDIC
*****************************************************************}
procedure TOM_Attestations.AgircChange(Sender: TObject);
var
StOrg : String;
QRechOrg :TQuery;
begin
{$IFNDEF EAGLCLIENT}
Agirc := THDBEdit (GetControl ('PAS_ORGAGIRC'));
{$ELSE}
Agirc := THEdit (GetControl ('PAS_ORGAGIRC'));
{$ENDIF}
//PT48
if (Suppr=True) then
   exit;
//FIN PT48

if Agirc <> NIL then
   begin
   Agirc.Plus := PGAttesEtab;
   StOrg:= 'SELECT POG_ETABLISSEMENT, POG_ORGANISME, POG_LIBELLE,'+
           ' POG_ADRESSE1, POG_ADRESSE2, POG_ADRESSE3, POG_CODEPOSTAL,'+
           ' POG_VILLE'+
           ' FROM ORGANISMEPAIE WHERE'+
           ' POG_ETABLISSEMENT="'+PGAttesEtab+'" AND'+
           ' POG_ORGANISME="'+Agirc.Text+'"';
   QRechOrg:=OpenSql(StOrg, True);
   if (not QRechOrg.eof) then
      begin
      SetControlText('LABEL_AGIRC', QRechOrg.FindField ('POG_LIBELLE').AsString);

      SetControlText('LABEL_RCADR1',
                     Copy(QRechOrg.FindField('POG_ADRESSE1').Asstring, 1,
                          Length(QRechOrg.FindField('POG_ADRESSE1').Asstring))+' '+
                     Copy(QRechOrg.FindField('POG_ADRESSE2').Asstring, 1,
                          Length(QRechOrg.FindField('POG_ADRESSE2').Asstring))+' '+
                     Copy(QRechOrg.FindField('POG_ADRESSE3').Asstring, 1,
                          Length(QRechOrg.FindField('POG_ADRESSE3').Asstring))+' '+
                     Copy(QRechOrg.FindField('POG_CODEPOSTAL').Asstring, 1,
                          Length(QRechOrg.FindField('POG_CODEPOSTAL').Asstring))+' '+
                     Copy(QRechOrg.FindField('POG_VILLE').Asstring, 1,
                          Length(QRechOrg.FindField('POG_VILLE').Asstring)));
      end;
   Ferme(QRechOrg);
   end;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 13/07/2000
Modifié le ... :   /  /
Description .. : ArrcoChange
Mots clefs ... : PAIE;PGATTESTATION;PGASSEDIC
*****************************************************************}
procedure TOM_Attestations.ArrcoChange(Sender: TObject);
var
StOrg : String;
QRechOrg :TQuery;
begin
{$IFNDEF EAGLCLIENT}
Arrco := THDBEdit (GetControl ('PAS_ORGARRCO'));
{$ELSE}
Arrco := THEdit (GetControl ('PAS_ORGARRCO'));
{$ENDIF}
//PT48
if (Suppr=True) then
   exit;
//FIN PT48

if Arrco <> NIL then
   begin
   Arrco.Plus := PGAttesEtab;
   StOrg:= 'SELECT POG_ETABLISSEMENT, POG_ORGANISME, POG_LIBELLE,'+
           ' POG_ADRESSE1, POG_ADRESSE2, POG_ADRESSE3, POG_CODEPOSTAL,'+
           ' POG_VILLE'+
           ' FROM ORGANISMEPAIE WHERE'+
           ' POG_ETABLISSEMENT="'+PGAttesEtab+'" AND'+
           ' POG_ORGANISME="'+Arrco.Text+'"';
   QRechOrg:=OpenSql(StOrg, True);
   if (not QRechOrg.EOF) then
      begin
      SetControlText('LABEL_ARRCO', QRechOrg.FindField ('POG_LIBELLE').AsString);

      SetControlText('LABEL_COADR1',
                     Copy(QRechOrg.FindField('POG_ADRESSE1').Asstring, 1,
                          Length(QRechOrg.FindField('POG_ADRESSE1').Asstring))+' '+
                     Copy(QRechOrg.FindField('POG_ADRESSE2').Asstring, 1,
                          Length(QRechOrg.FindField('POG_ADRESSE2').Asstring))+' '+
                     Copy(QRechOrg.FindField('POG_ADRESSE3').Asstring, 1,
                          Length(QRechOrg.FindField('POG_ADRESSE3').Asstring))+' '+
                     Copy(QRechOrg.FindField('POG_CODEPOSTAL').Asstring, 1,
                          Length(QRechOrg.FindField('POG_CODEPOSTAL').Asstring))+' '+
                     Copy(QRechOrg.FindField('POG_VILLE').Asstring, 1,
                          Length(QRechOrg.FindField('POG_VILLE').Asstring)));
      end;
   Ferme(QRechOrg);
   end;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 13/07/2000
Modifié le ... :   /  /
Description .. : AutresChange
Mots clefs ... : PAIE;PGATTESTATION;PGASSEDIC
*****************************************************************}
procedure TOM_Attestations.AutresChange(Sender: TObject);
var
StOrg : String;
QRechOrg :TQuery;
begin
{$IFNDEF EAGLCLIENT}
Autres := THDBEdit (GetControl ('PAS_ORGAUTRE'));
{$ELSE}
Autres := THEdit (GetControl ('PAS_ORGAUTRE'));
{$ENDIF}
//PT48
if (Suppr=True) then
   exit;
//FIN PT48

if Autres <> NIL then
   begin
   Autres.Plus := PGAttesEtab;
   StOrg:= 'SELECT POG_ETABLISSEMENT, POG_ORGANISME, POG_LIBELLE,'+
           ' POG_ADRESSE1, POG_ADRESSE2, POG_ADRESSE3, POG_CODEPOSTAL,'+
           ' POG_VILLE'+
           ' FROM ORGANISMEPAIE WHERE'+
           ' POG_ETABLISSEMENT="'+PGAttesEtab+'" AND'+
           ' POG_ORGANISME="'+Autres.Text+'"';
   QRechOrg:=OpenSql(StOrg, True);
   if (not QRechOrg.EOF) then
      begin
      SetControlText('LABEL_AUTRES', QRechOrg.FindField ('POG_LIBELLE').AsString);

      SetControlText('LABEL_AUADR1',
                     Copy(QRechOrg.FindField('POG_ADRESSE1').Asstring, 1,
                          Length(QRechOrg.FindField('POG_ADRESSE1').Asstring))+' '+
                     Copy(QRechOrg.FindField('POG_ADRESSE2').Asstring, 1,
                          Length(QRechOrg.FindField('POG_ADRESSE2').Asstring))+' '+
                     Copy(QRechOrg.FindField('POG_ADRESSE3').Asstring, 1,
                          Length(QRechOrg.FindField('POG_ADRESSE3').Asstring))+' '+
                     Copy(QRechOrg.FindField('POG_CODEPOSTAL').Asstring, 1,
                          Length(QRechOrg.FindField('POG_CODEPOSTAL').Asstring))+' '+
                     Copy(QRechOrg.FindField('POG_VILLE').Asstring, 1,
                          Length(QRechOrg.FindField('POG_VILLE').Asstring)));
      end;
   Ferme(QRechOrg);
   end;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 13/07/2000
Modifié le ... :   /  /
Description .. : QualiteEmpChange
Mots clefs ... : PAIE;PGATTESTATION;PGASSEDIC
*****************************************************************}
procedure TOM_Attestations.QualiteEmpChange({Sender: TObject});
begin
//PT48
if (Suppr=True) then
   exit;
//FIN PT48

if GetField('PAS_QUALITEEMPLOI') = '7' then
   SetControlVisible ('PAS_LQUALITE', TRUE)
else
   begin
   SetControlVisible ('PAS_LQUALITE', FALSE);
   SetField ('PAS_LQUALITE','');
   end;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 13/07/2000
Modifié le ... :   /  /    
Description .. : HorExit
Mots clefs ... : PAIE;PGATTESTATION;PGASSEDIC
*****************************************************************}
procedure TOM_Attestations.HorExit(Sender: TObject);
begin
//PT48
if (Suppr=True) then
   exit;
//FIN PT48

if (HorEtab <> NIL) AND (HorEtabann <> NIL) AND (HorSal <> NIL) AND
   (HorSalann <> NIL) then
   if (HorEtab.Text = HorSal.Text) AND (HorEtabann.Text = HorSalann.Text) then
      begin
      SetControlVisible ('LABEL_DIFHOR', FALSE);
      SetControlVisible ('PAS_MOTIFDIFF', FALSE);
      SetField('PAS_MOTIFDIFF','');
      end
   else
      begin
      SetControlVisible ('LABEL_DIFHOR', TRUE);
      SetControlVisible ('PAS_MOTIFDIFF', TRUE);
      end;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 13/07/2000
Modifié le ... :   /  /    
Description .. : DifHorChange
Mots clefs ... : PAIE;PGATTESTATION;PGASSEDIC
*****************************************************************}
procedure TOM_Attestations.DifHorChange({Sender: TObject});
begin
//PT48
if (Suppr=True) then
   exit;
//FIN PT48

if GetField('PAS_MOTIFDIFF') = '3' then
   begin
   SetControlVisible ('LABELSAISIEDIFF', TRUE);
   SetControlVisible ('PAS_AUTREDIFF', TRUE);
   end
else
   begin
   SetControlVisible ('LABELSAISIEDIFF', FALSE);
   SetControlVisible ('PAS_AUTREDIFF', FALSE);
   SetField ('PAS_AUTREDIFF','');
   end;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 13/07/2000
Modifié le ... :   /  /
Description .. : ChomTotClick
Mots clefs ... : PAIE;PGATTESTATION;PGASSEDIC
*****************************************************************}
procedure TOM_Attestations.ChomTotClick(Sender: TObject);
begin
//PT48
if (Suppr=True) then
   exit;
//FIN PT48

if (Tcheckbox(Sender).name='PAS_CHOMTOTAL') then
   begin
   if ChomTot <> NIL then
      begin
      if ChomTot.Checked = TRUE then
         SetControlVisible ('PAS_DDTEFP', TRUE)
      else
         begin
         SetControlVisible ('PAS_DDTEFP', FALSE);
         if DDTEFP <> NIL then
            DDTEFP.Checked := FALSE;
         end;
      end;
   end;

if (Tcheckbox(Sender).name='PAS_DDTEFP') then
   begin
   if DDTEFP <> NIL then
      begin
      if DDTEFP.Checked = TRUE then
         begin
         SetControlVisible ('LABELCHOMPER', TRUE);
         SetControlVisible ('PAS_DEBCHOMAGE', TRUE);
         SetControlVisible ('LABELCHOMAU', TRUE);
         SetControlVisible ('PAS_FINCHOMAGE', TRUE);
         SetControlVisible ('LABELCHOMREP', TRUE);
         SetControlVisible ('PAS_REPRISECHOM', TRUE);
         end
      else
         begin
         SetControlVisible ('LABELCHOMPER', FALSE);
         SetControlVisible ('PAS_DEBCHOMAGE', FALSE);
         SetControlVisible ('LABELCHOMAU', FALSE);
         SetControlVisible ('PAS_FINCHOMAGE', FALSE);
         SetControlVisible ('LABELCHOMREP', FALSE);
         SetControlVisible ('PAS_REPRISECHOM', FALSE);
         SetField('PAS_DEBCHOMAGE',IDate1900);
         SetField('PAS_FINCHOMAGE',IDate1900);
         SetField('PAS_REPRISECHOM',IDate1900);
         end;
      end;
   end;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 13/07/2000
Modifié le ... :   /  /    
Description .. : ContratPartChange
Mots clefs ... : PAIE;PGATTESTATION;PGASSEDIC
*****************************************************************}
procedure TOM_Attestations.ContratPartChange({Sender: TObject});
begin
//PT48
if (Suppr=True) then
   exit;
//FIN PT48

if GetField('PAS_CONTRATPARTIC') = '6' then
   begin
   SetControlVisible ('LABELSAISIECONTRAT', TRUE);
   SetControlVisible ('PAS_LCONTRATPART', TRUE);
   end
else
   begin
   SetControlVisible ('LABELSAISIECONTRAT', FALSE);
   SetControlVisible ('PAS_LCONTRATPART', FALSE);
   SetField ('PAS_LCONTRATPART','');
   end;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 13/07/2000
Modifié le ... :   /  /
Description .. : FonctionChange
Mots clefs ... : PAIE;PGATTESTATION;PGASSEDIC
*****************************************************************}
procedure TOM_Attestations.FonctionChange({Sender: TObject});
begin
//PT48
if (Suppr=True) then
   exit;
//FIN PT48

if GetField('PAS_FONCTIONSAL') = '8' then
   begin
   SetControlVisible ('LABELSAISIEFONC', TRUE);
   SetControlVisible ('PAS_LIBFONCTION', TRUE);
   end
else
   begin
   SetControlVisible ('LABELSAISIEFONC', FALSE);
   SetControlVisible ('PAS_LIBFONCTION', FALSE);
   SetField ('PAS_LIBFONCTION','');
   end;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 13/07/2000
Modifié le ... :   /  /
Description .. : MotifChange
Mots clefs ... : PAIE;PGATTESTATION;PGASSEDIC
*****************************************************************}
procedure TOM_Attestations.MotifChange({Sender: TObject});
var
Motif : string;
begin
//PT48
if (Suppr=True) then
   exit;
//FIN PT48

Motif := GetField ('PAS_MOTRUPCONT');
if (Motif='403') then
   Motif:= '';
//PT44
if (Motif='90') then
   Motif:= '60';
//FIN PT44
SetControlText('LABEL_MOTIFLIB', RechDom ('PGMOTIFSORTIE',Motif,FALSE,''));
if (Motif = '20') or (Motif = '59') or (Motif = '60') then
   begin
   SetControlVisible ('LABELSAISIEMOTIF', TRUE);
   SetControlVisible ('PAS_AUTRERUPT', TRUE);
   SetControlText('PAS_DECLARRENS1', GetControlText('PAS_AUTRERUPT'));
   SetField('PAS_DECLARRENS1', GetControlText('PAS_AUTRERUPT'));
   end
else
   begin
   SetControlVisible ('LABELSAISIEMOTIF', FALSE);
   SetControlVisible ('PAS_AUTRERUPT', FALSE);
   if (GetControlText('PAS_AUTRERUPT') <> '') then
      SetField ('PAS_AUTRERUPT','');
   SetControlText('PAS_DECLARRENS1', GetControlText('LABEL_MOTIFLIB'));
   SetField('PAS_DECLARRENS1', GetControlText('LABEL_MOTIFLIB'));
   end;
if Motif = '14' then
   begin
   SetControlVisible ('GROUPBOXMOT2', TRUE);
   SetControlVisible ('GROUPBOXMOT3', TRUE);
   end
else
   begin
   SetControlVisible ('GROUPBOXMOT2', FALSE);
   SetControlVisible ('GROUPBOXMOT3', FALSE);
   if PlanSoc <> NIL then
      PlanSoc.Checked := FALSE;
   if AgeSalarieFNE <> NIL then
      AgeSalarieFNE.Checked := FALSE;
   end;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 13/07/2000
Modifié le ... :   /  /
Description .. : PlanSocClick
Mots clefs ... : PAIE;PGATTESTATION;PGASSEDIC
*****************************************************************}
procedure TOM_Attestations.PlanSocClick(Sender: TObject);
begin
//PT48
if (Suppr=True) then
   exit;
//FIN PT48

if PlanSoc <> NIL then
   if PlanSoc.Checked = TRUE then
      begin
      SetControlVisible ('LABEL_PLANSOC', TRUE);
      SetControlVisible ('PAS_DATEPLANSOC', TRUE);
      end
   else
      begin
      SetControlVisible ('LABEL_PLANSOC', FALSE);
      SetControlVisible ('PAS_DATEPLANSOC', FALSE);
      SetField ('PAS_DATEPLANSOC',IDate1900);
      end;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 13/07/2000
Modifié le ... :   /  /    
Description .. : AgeSalarieFNEClick
Mots clefs ... : PAIE;PGATTESTATION;PGASSEDIC
*****************************************************************}
procedure TOM_Attestations.AgeSalarieFNEClick(Sender: TObject);
begin
//PT48
if (Suppr=True) then
   exit;
//FIN PT48

if AgeSalarieFNE <> NIL then
   if AgeSalarieFNE.Checked = TRUE then
      SetControlVisible ('PAS_REFUSFNE', TRUE)
   else
      begin
      SetControlVisible ('PAS_REFUSFNE', FALSE);
      SetField ('PAS_REFUSFNE','-');
      end;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 13/07/2000
Modifié le ... :   /  /
Description .. : DernierJourChange
Mots clefs ... : PAIE;PGATTESTATION;PGASSEDIC
*****************************************************************}
procedure TOM_Attestations.DernierJourChange(Sender: TObject);
begin
//PT48
if (Suppr=True) then
   exit;
//FIN PT48

if not (ds.state in [dsinsert,dsedit]) then
   ds.edit;
ChargeIndemnites;
ChargeSalaires (TRUE);
ChargePrimes (TRUE);
ChargeSolde (TRUE);
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 19/10/2005
Modifié le ... :   /  /
Description .. : IndemniteClick
Mots clefs ... : PAIE;PGATTESTATION;PGASSEDIC
*****************************************************************}
procedure TOM_Attestations.IndemniteClick(Sender: TObject);
var
NomEdit : string;
Longueur : integer;
begin
//PT48
if (Suppr=True) then
   exit;
//FIN PT48

Longueur:= Length(Tcheckbox(Sender).name)-1;
NomEdit:= Copy(Tcheckbox(Sender).name, 0, Longueur)+'M';

if Tcheckbox(Sender).Checked = TRUE then
   begin
   SetControlVisible (NomEdit, TRUE);
   SetControlVisible ('LABEL_INDTOT', TRUE);
   SetControlVisible ('EDIT_INDTOT', TRUE);
   end
else
   begin
   SetControlVisible (NomEdit, FALSE);
   If NomEdit = 'INDSPECONVENTIONM' then   //pt51
   begin
   Setcontroltext('INDSPECONVENTIONM' ,'0'); //pt51
   Indrupcon := StrToFloat(GetControlText('INDSPECONVENTIONM'));
   Total:= Leg+CDD+CNE+FinMiss+Retraite+Sinistre+Spec+Spcfiq+Journal+Client+
           Avion+Apprenti+Autre+Indrupcon;  //PT51
   SetControlText('EDIT_INDTOT', FloatToStr(Total))
   end
   else                                      //pt51
   SetField (NomEdit, 0);
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 20/10/2005
Modifié le ... :   /  /
Description .. : IndemniteExit
Mots clefs ... : PAIE;PGATTESTATION;PGASSEDIC
*****************************************************************}
{PT42-1}   //PT51
procedure TOM_Attestations.IndemniteExit(Sender: TObject);
var
NomCHKBOX : string;
Longueur : integer;
begin
//deb PT51
   Longueur:= Length(THEdit(Sender).name)-1;
   NomCHKBOX:= Copy(THEdit(Sender).name, 0, Longueur)+'B';

if (THEdit(Sender).Name=('INDSPECONVENTIONM')) then
   begin
   if (GetControlText('INDSPECONVENTIONM') <> '0') and
      (GetControlText('INDSPECONVENTIONM') <> '') then
      Indrupcon:= StrToFloat(GetControlText('INDSPECONVENTIONM'))
   else
      begin
      Indrupcon:= 0;
      SetControlChecked(NomCHKBOX, False);
      end;
   end;
Total:= Leg+CDD+CNE+FinMiss+Retraite+Sinistre+Spec+Spcfiq+Journal+Client+Avion+
        Apprenti+Autre+Indrupcon;
SetControlText('EDIT_INDTOT', FloatToStr(Total));
end;
//fin PT51
{//PT42-1
Longueur:= Length(THEdit(Sender).name)-1;
NomCHKBOX:= Copy(THEdit(Sender).name, 0, Longueur)+'B';

if (THEdit(Sender).Name=('INDCNEM')) then
   begin
   if (GetControlText('INDCNEM') <> '0') and
      (GetControlText('INDCNEM') <> '') then
      CNE:= StrToFloat(GetControlText('INDCNEM'))
   else
      begin
      CNE:= 0;
      SetControlChecked(NomCHKBOX, False);
      end;
   end;

if (THEdit(Sender).Name=('INDSINISTREM')) then
   begin
   if (GetControlText('INDSINISTREM') <> '0') and
      (GetControlText('INDSINISTREM') <> '') then
      Sinistre:= StrToFloat(GetControlText('INDSINISTREM'))
   else
      begin
      Sinistre:= 0;
      SetControlChecked(NomCHKBOX, False);
      end;
   end;

if (THEdit(Sender).Name=('INDSPCFIQM')) then
   begin
   if (GetControlText('INDSPCFIQM') <> '0') and
      (GetControlText('INDSPCFIQM') <> '') then
      Spcfiq:= StrToFloat(GetControlText('INDSPCFIQM'))
   else
      begin
      Spcfiq:= 0;
      SetControlChecked(NomCHKBOX, False);
      end;
   end;

if (THEdit(Sender).Name=('INDAPPRENTIM')) then
   begin
   if (GetControlText('INDAPPRENTIM') <> '0') and
      (GetControlText('INDAPPRENTIM') <> '') then
      Apprenti:= StrToFloat(GetControlText('INDAPPRENTIM'))
   else
      begin
      Apprenti:= 0;
      SetControlChecked(NomCHKBOX, False);
      end;
   end;

Total:= Leg+CDD+CNE+FinMiss+Retraite+Sinistre+Spec+Spcfiq+Journal+Client+Avion+
        Apprenti+Autre;
SetControlText('EDIT_INDTOT', FloatToStr(Total));
end;
}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 13/07/2000
Modifié le ... :   /  /
Description .. : FNGSClick
Mots clefs ... : PAIE;PGATTESTATION;PGASSEDIC
*****************************************************************}
procedure TOM_Attestations.FNGSClick(Sender: TObject);
begin
//PT48
if (Suppr=True) then
   exit;
//FIN PT48

if (Tcheckbox(Sender).name='PAS_FNGSRED') then
   begin
   if (FNGSRED <> NIL) then
      begin
      if (FNGSRED.Checked = TRUE) then
         begin
         SetControlVisible ('PAS_FNGS', TRUE);
         SetControlVisible ('PAS_FNGS2', TRUE);
         SetControlVisible ('LABEL_FNGSMOT1', TRUE);
         SetControlVisible ('PAS_FNGSMOT1', TRUE);
         SetControlVisible ('LABEL_FNGSMOT2', TRUE);
         SetControlVisible ('PAS_FNGSMOT2', TRUE);
         end
      else
         begin
         SetControlVisible ('PAS_FNGS', FALSE);
         SetControlVisible ('LABEL_FNGSCRE1', FALSE);
         SetControlVisible ('PAS_FNGSCRE1', FALSE);
         SetControlVisible ('LABEL_FNGSMOT1', FALSE);
         SetControlVisible ('PAS_FNGSMOT1', FALSE);
         SetControlVisible ('PAS_FNGS2', FALSE);
         SetControlVisible ('LABEL_FNGSCRE2', FALSE);
         SetControlVisible ('PAS_FNGSCRE2', FALSE);
         SetControlVisible ('LABEL_FNGSMOT2', FALSE);
         SetControlVisible ('PAS_FNGSMOT2', FALSE);
         if FNGS <> NIL then
            FNGS.Checked := FALSE;
         if FNGS2 <> NIL then
            FNGS2.Checked := FALSE;
         SetField('PAS_FNGSCRE1','');
         SetField('PAS_FNGSMOT1','');
         SetField('PAS_FNGSCRE2','');
         SetField('PAS_FNGSMOT2','');
         end;
      end;
   end;

if (Tcheckbox(Sender).name='PAS_FNGS') then
   begin
   if ((FNGS <> NIL) and (FNGSRED <> NIL) and (FNGSRED.Checked = True)) then
      begin
      if FNGS.Checked = TRUE then
         begin
         SetControlVisible ('LABEL_FNGSCRE1', TRUE);
         SetControlVisible ('PAS_FNGSCRE1', TRUE);
         SetControlVisible ('LABEL_FNGSMOT1', FALSE);
         SetControlVisible ('PAS_FNGSMOT1', FALSE);
         SetField('PAS_FNGSMOT1','');
         end
      else
         begin
         SetControlVisible ('LABEL_FNGSCRE1', FALSE);
         SetControlVisible ('PAS_FNGSCRE1', FALSE);
         SetField('PAS_FNGSCRE1','');
         SetControlVisible ('LABEL_FNGSMOT1', TRUE);
         SetControlVisible ('PAS_FNGSMOT1', TRUE);
         end;
      end;
   end;

if (Tcheckbox(Sender).name='PAS_FNGS2') then
   begin
   if ((FNGS2 <> NIL) and (FNGSRED <> NIL) and (FNGSRED.Checked = True)) then
      begin
      if FNGS2.Checked = TRUE then
         begin
         SetControlVisible ('LABEL_FNGSCRE2', TRUE);
         SetControlVisible ('PAS_FNGSCRE2', TRUE);
         SetControlVisible ('LABEL_FNGSMOT2', FALSE);
         SetControlVisible ('PAS_FNGSMOT2', FALSE);
         SetField('PAS_FNGSMOT2','');
         end
      else
         begin
         SetControlVisible ('LABEL_FNGSCRE2', FALSE);
         SetControlVisible ('PAS_FNGSCRE2', FALSE);
         SetField('PAS_FNGSCRE2','');
         SetControlVisible ('LABEL_FNGSMOT2', TRUE);
         SetControlVisible ('PAS_FNGSMOT2', TRUE);
         end;
      end;
   end;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 13/07/2000
Modifié le ... :   /  /
Description .. : Imprimer
Mots clefs ... : PAIE;PGATTESTATION;PGASSEDIC
*****************************************************************}
procedure TOM_Attestations.Imprimer(Sender: TObject);
var
Pages : TPageControl;
(*Mise en commentaire PT45
{$IFDEF EAGLCLIENT}
StPages : String;
{$ENDIF}
*)
//Debut PT45
  TobDonnees, TempTob : Tob;
  i,j,k : Integer;
  Q : TQuery;
//Fin PT45
  Adresse, Libelle, Rech : String;
  NumOrdre : Integer;

  Procedure GetInfoCaisse (Etab, Organisme : String; var Libelle, Adresse : String);
  var
    StCaisse : String;
    QRechCaisse : TQuery;
  begin
    Libelle := ''; Adresse := '';
    StCaisse:= 'SELECT POG_ETABLISSEMENT, POG_ORGANISME, POG_LIBELLE, POG_ADRESSE1, '
             + ' POG_ADRESSE2, POG_ADRESSE3, POG_CODEPOSTAL, POG_VILLE'
             + ' FROM ORGANISMEPAIE WHERE'
             + ' POG_ETABLISSEMENT = "'+Etab+'" AND POG_ORGANISME="'+Organisme+'"';
    QRechCaisse:=OpenSql(StCaisse,TRUE);
    if (NOT QRechCaisse.EOF) then
    begin
      Libelle := QRechCaisse.FindField ('POG_LIBELLE').AsString;
      Adresse := Copy(QRechCaisse.FindField('POG_ADRESSE1').Asstring, 1,
                      Length(QRechCaisse.FindField('POG_ADRESSE1').Asstring))+' '+
                 Copy(QRechCaisse.FindField('POG_ADRESSE2').Asstring, 1,
                      Length(QRechCaisse.FindField('POG_ADRESSE2').Asstring))+' '+
                 Copy(QRechCaisse.FindField('POG_ADRESSE3').Asstring, 1,
                      Length(QRechCaisse.FindField('POG_ADRESSE3').Asstring))+' '+
                 Copy(QRechCaisse.FindField('POG_CODEPOSTAL').Asstring, 1,
                      Length(QRechCaisse.FindField('POG_CODEPOSTAL').Asstring))+' '+
                 Copy(QRechCaisse.FindField('POG_VILLE').Asstring, 1,
                      Length(QRechCaisse.FindField('POG_VILLE').Asstring));
    end;
    Ferme(QRechCaisse);
  end;
begin
NumOrdre:= GetField ('PAS_ORDRE');
{PT43
QRechNumAttest:=OpenSQL('SELECT MAX(PAS_ORDRE)'+
                        ' FROM ATTESTATIONS',TRUE);
if (not QRechNumAttest.Eof) then
   begin
   if (QRechNumAttest.Fields[0].AsInteger <> NumOrdre) then
      BEGIN
      try
      IMax:= QRechNumAttest.Fields[0].AsInteger+1;
      except
            on E: EConvertError do
               IMax:= 1;
      end;
      SetField('PAS_ORDRE',IMax);
      ExecuteSQL('UPDATE ATTESTATIONS SET PAS_ORDRE='+IntToStr(IMax)+' WHERE'+
                 ' PAS_ORDRE='+IntToStr(QRechNumAttest.Fields[0].AsInteger)+'');
      ExecuteSQL('UPDATE ATTSALAIRES SET PAL_ORDRE='+IntToStr(IMax)+' WHERE'+
                 ' PAL_ORDRE='+IntToStr(QRechNumAttest.Fields[0].AsInteger)+'');
      END;
   end;
Ferme(QRechNumAttest);
}

Rech := 'SELECT PAS_SALARIE, PSA_CIVILITE, PSA_NOMJF, PSA_ADRESSE1,'+
        ' PSA_ADRESSE2, PSA_ADRESSE3, PSA_CODEPOSTAL, PSA_VILLE,'+
        ' PSA_DATENAISSANCE'+
        ' FROM ATTESTATIONS'+
        ' LEFT JOIN SALARIES ON'+
        ' PAS_SALARIE=PSA_SALARIE WHERE'+
        ' PAS_SALARIE="'+PGAttesSalarie+'" AND'+
        ' PAS_TYPEATTEST="ASS" AND'+
        ' PAS_ORDRE='+IntToStr(NumOrdre);
Pages := TPageControl(GetControl('PAGES'));

if (Pages<>nil) then
(*
Debut PT45
{$IFNDEF EAGLCLIENT}
*)
   begin
   TobDonnees:= TOB.Create ('les données', nil, -1);
   TempTob:= TOB.Create ('L''enregistrement', TobDonnees, -1);
   TempTob.GetEcran (Ecran, nil, true);
   for i:= 0 to TobDonnees.FillesCount(0)-1 do
       begin
       for j:= 0 to Salaires.RowCount -1 do
           for k:= 0 to Salaires.ColCount -1 do
               TobDonnees.Detail[i].AddChampSupValeur ('GRID_SALAIRES_'+IntToStr(k)+'_'+IntToStr(j)+'_',
                                                       Salaires.CellValues[k,j]);
       for j:= 0 to Primes.RowCount -1 do
           for k:= 0 to Primes.ColCount -1 do
               TobDonnees.Detail[i].AddChampSupValeur ('GRID_PRIMES_'+IntToStr(k)+'_'+IntToStr(j)+'_',
                                                       Primes.CellValues[k,j]);
       for j:= 0 to Solde.RowCount -1 do
           for k:= 0 to Solde.ColCount -1 do
               TobDonnees.Detail[i].AddChampSupValeur ('GRID_SOLDE_'+IntToStr(k)+'_'+IntToStr(j)+'_',
                                                       Solde.CellValues[k,j]);
       Q:= OpenSQL (Rech, TRUE);
       if (Not Q.EOF) then
          For j:= 0 to Q.FieldCount -1 do
              TobDonnees.Detail[i].AddChampSupValeur (Q.Fields[j].FieldName,
                                                      Q.Fields[j].AsString);
       Ferme(Q);

//On remplace la valeur par le libellé
       TobDonnees.detail[i].SetString ('PAS_PUBLICTYPE',
                                       RechDom ('PGPUBLICTYPE',
                                                TobDonnees.detail[i].GetString ('PAS_PUBLICTYPE'),
                                                False));
       TobDonnees.detail[i].SetString ('PAS_COLLTERRITOR',
                                       RechDom ('PGCOLLTERR',
                                                TobDonnees.detail[i].GetString ('PAS_COLLTERRITOR'),
                                                False));
       TobDonnees.detail[i].SetString ('PAS_DERNIEREMPLOI',
                                       RechDom ('PGLIBEMPLOI',
                                                TobDonnees.detail[i].GetString ('PAS_DERNIEREMPLOI'),
                                                False));
//PT49
       TobDonnees.detail[i].SetString ('PAS_DEPART',
                                       RechDom ('JUDEPART',
                                                TobDonnees.detail[i].GetString ('PAS_DEPART'),
                                                False));
       TobDonnees.detail[i].SetString ('PAS_QUALITEEMPLOI',
                                       RechDom ('PGQUALITEEMPLOI',
                                                TobDonnees.detail[i].GetString ('PAS_QUALITEEMPLOI'),
                                                False));
       TobDonnees.detail[i].SetString ('PAS_MOTIFDIFF',
                                       RechDom ('PGMOTIFDIFF',
                                                TobDonnees.detail[i].GetString ('PAS_MOTIFDIFF'),
                                                False));
       TobDonnees.detail[i].SetString ('PAS_CONTRATNAT',
                                       RechDom ('PGCONTRATNATURE',
                                                TobDonnees.detail[i].GetString ('PAS_CONTRATNAT'),
                                                False));
       TobDonnees.detail[i].SetString ('PAS_CONTRATPARTIC',
                                       RechDom ('PGCONTRATPART',
                                                TobDonnees.detail[i].GetString ('PAS_CONTRATPARTIC'),
                                                False));
       TobDonnees.detail[i].SetString ('PAS_FONCTIONSAL',
                                       RechDom ('PGSTATUT',
                                                TobDonnees.detail[i].GetString ('PAS_FONCTIONSAL'),
                                                False));
//FIN PT49
       TobDonnees.detail[i].SetString ('LABEL_MOTIFLIB',
                                       TobDonnees.detail[i].GetString ('PAS_DECLARRENS1'));

//PT49
//ARRCO
       GetInfoCaisse (PGAttesEtab,
                      TobDonnees.Detail[i].GetString ('PAS_ORGARRCO'), Libelle,
                      Adresse);
    TobDonnees.Detail[i].AddChampSupValeur ('LABEL_ARRCO', Libelle);
    TobDonnees.Detail[i].AddChampSupValeur ('LABEL_COADR1', Adresse);
//AGIRC
    GetInfoCaisse (PGAttesEtab, TobDonnees.Detail[i].GetString('PAS_ORGAGIRC'),
                   Libelle, Adresse);
    TobDonnees.Detail[i].AddChampSupValeur ('LABEL_AGIRC', Libelle);
    TobDonnees.Detail[i].AddChampSupValeur ('LABEL_RCADR1', Adresse);
//Autre caisse de retraite complémentaire
    GetInfoCaisse (PGAttesEtab, TobDonnees.Detail[i].GetString('PAS_ORGAUTRE'),
                   Libelle, Adresse);
    TobDonnees.Detail[i].AddChampSupValeur ('LABEL_AUTRES', Libelle);
    TobDonnees.Detail[i].AddChampSupValeur ('LABEL_AUADR1', Adresse);
//FIN PT49
       end;

   LanceEtatTob ('E', 'PAT', 'ASS', TobDonnees, True, False, False, nil, Rech, '', False);
   FreeAndNil (TobDonnees);
   end;
//Fin PT45
(*Mise en commentaire PT45
{$ELSE}
   begin
   StPages := AglGetCriteres (Pages, FALSE);
   LanceEtat('E','PAT','ASE',True,False,False,NIL,Rech,'',False,0,StPages);
   end;
{$ENDIF}
*)
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 04/05/2001
Modifié le ... :   /  /
Description .. : Validation
Mots clefs ... : PAIE;PGATTESTATION;PGASSEDIC
*****************************************************************}
procedure TOM_Attestations.Validation(Sender: TObject);
begin
if Sender = Valid then
   begin
   TFFiche(Ecran).MonoFiche := False;
(*Mise en commentaire PT45
{$IFDEF EAGLCLIENT}
   ESalaires;
   EPrimes;
   ESolde;
{$ENDIF}
*)
   TFFiche(Ecran).BValiderClick(Nil);
   TFFiche(Ecran).MonoFiche := True;
   Imprim.Enabled:=True;
   end;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 13/07/2000
Modifié le ... :   /  /    
Description .. : UpdateTable
Mots clefs ... : PAIE;PGATTESTATION;PGASSEDIC
*****************************************************************}
procedure TOM_Attestations.UpdateTable();
var
StAtt, StCreatSal, StModifSal, StSal : String;
i, j, k, Ordre : integer;
QRechAtt:TQuery;
begin
StSal:= 'SELECT PAS_SALARIE'+
        ' FROM ATTESTATIONS WHERE'+
        ' PAS_SALARIE="'+PGAttesSalarie+'" AND'+
        ' PAS_TYPEATTEST="ASS"';
if ExisteSQL(StSal)=FALSE then
   begin                  //insertion table principale
   Ordre:= 1;
   StAtt:='SELECT MAX(PAS_ORDRE)'+
          ' FROM ATTESTATIONS';
   QRechAtt:=OpenSql(StAtt,True);
   if (not QRechAtt.Eof) then
      try
      Ordre:= QRechAtt.Fields[0].asinteger+1;
      except
            on E: EConvertError do
               Ordre:= 1;
      end;
   Ferme(QRechAtt);
   For j:=1 to 12 do
       begin
       For k:=0 to 2 do
           begin
           if (IsValidDate(Salaires.CellValues[k,j])=FALSE) then
              Salaires.CellValues[k,j]:= '01/01/1900';
           end;
       StCreatSal:= 'INSERT INTO ATTSALAIRES (PAL_ORDRE, PAL_MOIS,'+
                    ' PAL_SALARIE, PAL_TYPEATTEST, PAL_DATEDEBUT, PAL_DATEFIN,'+
                    ' PAL_PAYELE, PAL_NBHEURES, PAL_JNONPAYES, PAL_SALAIRE,'+
                    ' PAL_PRECOMPTE, PAL_OBSERVATIONS)'+
                    ' VALUES'+
                    ' ('+IntToStr(Ordre)+', '+IntToStr(j)+','+
                    ' "'+PGAttesSalarie+'", "ASS",'+
                    ' "'+UsDateTime(StrToDate(Salaires.CellValues[0,j]))+'",'+
                    ' "'+UsDateTime(StrToDate(Salaires.CellValues[1,j]))+'",'+
                    ' "'+UsDateTime(StrToDate(Salaires.CellValues[2,j]))+'",'+
                    ' '+StrFPoint(Valeur(Salaires.CellValues[3,j]))+','+
                    ' '+StrFPoint(Valeur(Salaires.CellValues[4,j]))+','+
                    ' '+StrFPoint(Valeur(Salaires.CellValues[5,j]))+','+
                    ' '+StrFPoint(Valeur(Salaires.CellValues[6,j]))+','+
                    ' "'+Salaires.CellValues[7,j]+'")';
       ExecuteSQL(StCreatSal) ;
       end;
   For j:=13 to 15 do
       begin
       i := j-12;
       if (IsValidDate(Primes.CellValues[0,i])=FALSE) then
          Primes.CellValues[0,i]:= '01/01/1900';
       if (IsValidDate(Primes.CellValues[1,i])=FALSE) then
          Primes.CellValues[1,i]:= '01/01/1900';
       if (IsValidDate(Primes.CellValues[3,i])=FALSE) then
          Primes.CellValues[3,i]:= '01/01/1900';
       StCreatSal:= 'INSERT INTO ATTSALAIRES (PAL_ORDRE, PAL_MOIS,'+
                    ' PAL_SALARIE, PAL_TYPEATTEST, PAL_DATEDEBUT, PAL_DATEFIN,'+
                    ' PAL_OBSERVATIONS, PAL_PAYELE, PAL_SALAIRE)'+
                    ' VALUES'+
                    ' ('+IntToStr(Ordre)+', '+IntToStr(j)+','+
                    ' "'+PGAttesSalarie+'", "ASS",'+
                    ' "'+UsDateTime(StrToDate(Primes.CellValues[0,i]))+'",'+
                    ' "'+UsDateTime(StrToDate(Primes.CellValues[1,i]))+'",'+
                    ' "'+Primes.CellValues[2,i]+'",'+
                    ' "'+UsDateTime(StrToDate(Primes.CellValues[3,i]))+'",'+
                    ' '+StrFPoint(Valeur(Primes.CellValues[4,i]))+')';
       ExecuteSQL(StCreatSal) ;
       end;
   For j:=16 to 17 do
       begin
       i := j-15;
       For k:=0 to 2 do
           begin
           if (IsValidDate(Solde.CellValues[k,i])=FALSE) then
              Solde.CellValues[k,i]:= '01/01/1900';
           end;
       StCreatSal:= 'INSERT INTO ATTSALAIRES (PAL_ORDRE, PAL_MOIS,'+
                    ' PAL_SALARIE, PAL_TYPEATTEST, PAL_DATEDEBUT, PAL_DATEFIN,'+
                    ' PAL_PAYELE, PAL_NBHEURES, PAL_SALAIRE, PAL_PRECOMPTE)'+
                    ' VALUES'+
                    ' ('+IntToStr(Ordre)+', '+IntToStr(j)+','+
                    ' "'+PGAttesSalarie+'", "ASS",'+
                    ' "'+UsDateTime(StrToDate(Solde.CellValues[0,i]))+'",'+
                    ' "'+UsDateTime(StrToDate(Solde.CellValues[1,i]))+'",'+
                    ' "'+UsDateTime(StrToDate(Solde.CellValues[2,i]))+'",'+
                    ' '+StrFPoint(Valeur(Solde.CellValues[3,i]))+','+
                    ' '+StrFPoint(Valeur(Solde.CellValues[4,i]))+','+
                    ' '+StrFPoint(Valeur(Solde.CellValues[5,i]))+')';
       ExecuteSQL(StCreatSal) ;
       end;
   ExecuteSQL('UPDATE SALARIES SET PSA_ASSEDIC="X" WHERE'+
              ' PSA_SALARIE="'+PGAttesSalarie+'"');
   end
else
   begin      //Modification
   Ordre:=GetField('PAS_ORDRE');

   For j:=1 to 12 do
       begin
       For k:=0 to 2 do
           begin
           if (IsValidDate(Salaires.CellValues[k,j])=FALSE) then
              Salaires.CellValues[k,j]:= '01/01/1900';
           end;
       StSal:= 'SELECT PAL_SALARIE'+
               ' FROM ATTSALAIRES WHERE'+
               ' PAL_ORDRE='+IntToStr(Ordre)+' AND'+
               ' PAL_MOIS='+IntToStr(j)+' AND'+
               ' PAL_SALARIE="'+PGAttesSalarie+'" AND'+
               ' PAL_TYPEATTEST="ASS"';
       if ExisteSQL (StSal)=FALSE then
          begin     //insertion table secondaire
          StCreatSal:= 'INSERT INTO ATTSALAIRES (PAL_ORDRE, PAL_MOIS,'+
                       ' PAL_SALARIE, PAL_TYPEATTEST, PAL_DATEDEBUT,'+
                       ' PAL_DATEFIN, PAL_PAYELE, PAL_NBHEURES, PAL_JNONPAYES,'+
                       ' PAL_SALAIRE, PAL_PRECOMPTE, PAL_OBSERVATIONS)'+
                       ' VALUES'+
                       ' ('+IntToStr(Ordre)+', '+IntToStr(j)+','+
                       ' "'+PGAttesSalarie+'", "ASS",'+
                       ' "'+UsDateTime(StrToDate(Salaires.CellValues[0,j]))+'",'+
                       ' "'+UsDateTime(StrToDate(Salaires.CellValues[1,j]))+'",'+
                       ' "'+UsDateTime(StrToDate(Salaires.CellValues[2,j]))+'",'+
                       ' '+StrFPoint(Valeur(Salaires.CellValues[3,j]))+','+
                       ' '+StrFPoint(Valeur(Salaires.CellValues[4,j]))+','+
                       ' '+StrFPoint(Valeur(Salaires.CellValues[5,j]))+','+
                       ' '+StrFPoint(Valeur(Salaires.CellValues[6,j]))+','+
                       ' "'+Salaires.CellValues[7,j]+'")';
          ExecuteSQL(StCreatSal) ;
          end
       else
          begin                  //MAJ table secondaire
          StModifSal:= 'UPDATE ATTSALAIRES  SET'+
                       ' PAL_DATEDEBUT="'+UsDateTime(StrToDate(Salaires.CellValues[0,j]))+'",'+
                       ' PAL_DATEFIN="'+UsDateTime(StrToDate(Salaires.CellValues[1,j]))+'",'+
                       ' PAL_PAYELE="'+UsDateTime(StrToDate(Salaires.CellValues[2,j]))+'",'+
                       ' PAL_NBHEURES='+StrFPoint(Valeur(Salaires.CellValues[3,j]))+','+
                       ' PAL_JNONPAYES='+StrFPoint(Valeur(Salaires.CellValues[4,j]))+','+
                       ' PAL_SALAIRE='+StrFPoint(Valeur(Salaires.CellValues[5,j]))+','+
                       ' PAL_PRECOMPTE='+StrFPoint(Valeur(Salaires.CellValues[6,j]))+','+
                       ' PAL_OBSERVATIONS="'+Salaires.CellValues[7,j]+'"'+
                       ' WHERE PAL_ORDRE='+IntToStr(Ordre)+' AND'+
                       ' PAL_MOIS='+IntToStr(j)+' AND'+
                       ' PAL_SALARIE="'+PGAttesSalarie+'" AND'+
                       ' PAL_TYPEATTEST="ASS"';
          ExecuteSQL(StModifSal);
          end;
       end;

   For j:=13 to 15 do
       begin
       i := j-12;
       if (IsValidDate(Primes.CellValues[0,i])=FALSE) then
          Primes.CellValues[0,i]:= '01/01/1900';
       if (IsValidDate(Primes.CellValues[1,i])=FALSE) then
          Primes.CellValues[1,i]:= '01/01/1900';
       if (IsValidDate(Primes.CellValues[3,i])=FALSE) then
          Primes.CellValues[3,i]:= '01/01/1900';
       StSal:= 'SELECT PAL_SALARIE'+
               ' FROM ATTSALAIRES WHERE'+
               ' PAL_ORDRE='+IntToStr(Ordre)+' AND'+
               ' PAL_MOIS='+IntToStr(j)+' AND'+
               ' PAL_SALARIE="'+PGAttesSalarie+'" AND'+
               ' PAL_TYPEATTEST="ASS"';
       if ExisteSQL (StSal)=FALSE then
          begin
          StCreatSal:= 'INSERT INTO ATTSALAIRES (PAL_ORDRE, PAL_MOIS,'+
                       ' PAL_SALARIE, PAL_TYPEATTEST, PAL_DATEDEBUT,'+
                       ' PAL_DATEFIN, PAL_OBSERVATIONS, PAL_PAYELE,'+
                       ' PAL_SALAIRE)'+
                       ' VALUES'+
                       ' ('+IntToStr(Ordre)+', '+IntToStr(j)+','+
                       ' "'+PGAttesSalarie+'", "ASS",'+
                       ' "'+UsDateTime(StrToDate(Primes.CellValues[0,i]))+'",'+
                       ' "'+UsDateTime(StrToDate(Primes.CellValues[1,i]))+'",'+
                       ' "'+Primes.CellValues[2,i]+'",'+
                       ' "'+UsDateTime(StrToDate(Primes.CellValues[3,i]))+'",'+
                       ' '+StrFPoint(Valeur(Primes.CellValues[4,i]))+')';
          ExecuteSQL(StCreatSal) ;
          end
       else
          begin
          StModifSal:= 'UPDATE ATTSALAIRES  SET'+
                       ' PAL_DATEDEBUT="'+UsDateTime(StrToDate(Primes.CellValues[0,i]))+'",'+
                       ' PAL_DATEFIN="'+UsDateTime(StrToDate(Primes.CellValues[1,i]))+'",'+
                       ' PAL_OBSERVATIONS="'+Primes.CellValues[2,i]+'",'+
                       ' PAL_PAYELE="'+UsDateTime(StrToDate(Primes.CellValues[3,i]))+'",'+
                       ' PAL_SALAIRE='+StrFPoint(Valeur(Primes.CellValues[4,i]))+
                       ' WHERE PAL_ORDRE='+IntToStr(Ordre)+' AND'+
                       ' PAL_MOIS='+IntToStr(j)+' AND'+
                       ' PAL_SALARIE="'+PGAttesSalarie+'" AND'+
                       ' PAL_TYPEATTEST="ASS"';
          ExecuteSQL(StModifSal);
          end;
       end;

   For j:=16 to 17 do
       begin
       i := j-15;
       For k:=0 to 2 do
           begin
           if (IsValidDate(Solde.CellValues[k,i])=FALSE) then
              Solde.CellValues[k,i]:= '01/01/1900';
           end;
       StSal:= 'SELECT PAL_SALARIE'+
               ' FROM ATTSALAIRES WHERE'+
               ' PAL_ORDRE='+IntToStr(Ordre)+' AND'+
               ' PAL_MOIS='+IntToStr(j)+' AND'+
               ' PAL_SALARIE="'+PGAttesSalarie+'" AND'+
               ' PAL_TYPEATTEST="ASS"';
       if ExisteSQL (StSal)=FALSE then
          begin
          StCreatSal:= 'INSERT INTO ATTSALAIRES (PAL_ORDRE, PAL_MOIS,'+
                       ' PAL_SALARIE, PAL_TYPEATTEST, PAL_DATEDEBUT,'+
                       ' PAL_DATEFIN, PAL_PAYELE, PAL_NBHEURES, PAL_SALAIRE,'+
                       ' PAL_PRECOMPTE)'+
                       ' VALUES'+
                       ' ('+IntToStr(Ordre)+', '+IntToStr(j)+','+
                       ' "'+PGAttesSalarie+'", "ASS",'+
                       ' "'+UsDateTime(StrToDate(Solde.CellValues[0,i]))+'",'+
                       ' "'+UsDateTime(StrToDate(Solde.CellValues[1,i]))+'",'+
                       ' "'+UsDateTime(StrToDate(Solde.CellValues[2,i]))+'",'+
                       ' '+StrFPoint(Valeur(Solde.CellValues[3,i]))+','+
                       ' '+StrFPoint(Valeur(Solde.CellValues[4,i]))+','+
                       ' '+StrFPoint(Valeur(Solde.CellValues[5,i]))+')';
          ExecuteSQL(StCreatSal) ;
          end
       else
          begin
          StModifSal:= 'UPDATE ATTSALAIRES  SET'+
                       ' PAL_DATEDEBUT="'+UsDateTime(StrToDate(Solde.CellValues[0,i]))+'",'+
                       ' PAL_DATEFIN="'+UsDateTime(StrToDate(Solde.CellValues[1,i]))+'",'+
                       ' PAL_PAYELE="'+UsDateTime(StrToDate(Solde.CellValues[2,i]))+'",'+
                       ' PAL_NBHEURES='+StrFPoint(Valeur(Solde.CellValues[3,i]))+','+
                       ' PAL_SALAIRE='+StrFPoint(Valeur(Solde.CellValues[4,i]))+','+
                       ' PAL_PRECOMPTE='+StrFPoint(Valeur(Solde.CellValues[5,i]))+
                       ' WHERE PAL_ORDRE='+IntToStr(Ordre)+' AND'+
                       ' PAL_MOIS='+IntToStr(j)+' AND'+
                       ' PAL_SALARIE="'+PGAttesSalarie+'" AND'+
                       ' PAL_TYPEATTEST="ASS"';
          ExecuteSQL(StModifSal);
          end;
       end;
   end;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 02/04/2001
Modifié le ... :   /  /
Description .. : Sortie d'une cellule de grille
Mots clefs ... : PAIE;PGATTESTATION;PGASSEDIC
*****************************************************************}
procedure TOM_Attestations.CellExit (Sender: TObject; var ACol,ARow: Integer;
                                     var Cancel: Boolean);
var
Valide : boolean;
Valeur : string;
begin
if ((ACol=-1) and (ARow=-1)) then
   exit;
if (Sender = Salaires) then
   Valeur:= Salaires.CellValues[ACol,ARow]
else
if (Sender = Primes) then
   Valeur:= Primes.CellValues[ACol,ARow]
else
if (Sender = Solde) then
   Valeur:= Solde.CellValues[ACol,ARow];

if ((((Sender = Salaires) and ((ACol = 0) or (ACol = 1) or (ACol = 2))) or
   ((Sender = Primes) and ((ACol = 0) or (ACol = 1) or (ACol = 3))) or
   ((Sender = Solde) and ((ACol = 0) or (ACol = 1) or (ACol = 2)))) and
   (ARow<>0)) then
   begin
   Valide:= IsValidDate (Valeur);
   if (Valide=False) then
      begin
      if (Sender = Salaires) then
         Salaires.CellValues[ACol, ARow]:= '01/01/1900'
      else
      if (Sender = Primes) then
         Primes.CellValues[ACol, ARow]:= '01/01/1900'
      else
      if (Sender = Solde) then
         Solde.CellValues[ACol, ARow]:= '01/01/1900';
      end;
   end
else
if ((((Sender = Salaires) and ((ACol = 3) or (ACol = 4) or (ACol = 5) or
   (ACol = 6))) or ((Sender = Primes) and (ACol = 4)) or ((Sender = Solde) and
   ((ACol = 3) or (ACol = 4) or (ACol = 5)))) and (ARow<>0)) then
   begin
   Valide:= IsNumeric (Valeur);
   if (Valide=False) then
      begin
      if (Sender = Salaires) then
         Salaires.CellValues[ACol, ARow]:= '0'
      else
      if (Sender = Primes) then
         Primes.CellValues[ACol, ARow]:= '0'
      else
      if (Sender = Solde) then
         Solde.CellValues[ACol, ARow]:= '0';
      end;
   Valide:= True;
   end
else
if ((((Sender = Salaires) and (ACol = 7)) or ((Sender = Primes) and
   (ACol = 2))) and (ARow<>0)) then
   begin
   Valide:= (Length (Valeur)<=35);
   if (Valide=False) then
      begin
      if (Sender = Salaires) then
         Salaires.CellValues[ACol, ARow]:= Copy (Valeur, 1, 35)
      else
      if (Sender = Primes) then
         Primes.CellValues[ACol, ARow]:= Copy (Valeur, 1, 35);
      end;
   end
else
   Valide:= True;

if (Valide=True) then
   begin
   if not (ds.state in [dsinsert,dsedit]) then
      ds.edit
   end
else
   PGIBox ('Saisie incorrecte', TFFiche(Ecran).caption);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 27/06/2006
Modifié le ... :   /  /
Description .. : Sortie d'une grille
Mots clefs ... : PAIE;PGATTESTATION;PGASSEDIC
*****************************************************************}
procedure TOM_Attestations.GridExit (Sender: TObject);
var
Bool : boolean;
begin
CellExit (Sender, ColEnCours, LigEnCours, Bool);
ColEnCours:= -1;
LigEnCours:= -1;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 27/06/2006
Modifié le ... :   /  /
Description .. : Entrée sur une cellule d'une grille
Mots clefs ... : PAIE;PGATTESTATION;PGASSEDIC
*****************************************************************}
procedure TOM_Attestations.CellEnter (Sender: TObject; var ACol, ARow: Longint;
                                      var Cancel: Boolean);
begin
if (Sender = Salaires) then
   begin
   ColEnCours:= Salaires.Col;
   LigEnCours:= Salaires.Row;
   end
else
if (Sender = Primes) then
   begin
   ColEnCours:= Primes.Col;
   LigEnCours:= Primes.Row;
   end
else
if (Sender = Solde) then
   begin
   ColEnCours:= Solde.Col;
   LigEnCours:= Solde.Row;
   end;
end;


procedure TOM_Attestations.AffichDeclarant(Sender: TObject);
Var St : String;
begin
//PT48
if (Suppr=True) then
   exit;
//FIN PT48

if GetControlText('DECLARANT')='' then exit;
St := RechDom('PGDECLARANTATTEST',GetControlText('DECLARANT'),False);
If Ds.State in ([DsBrowse]) then ds.Edit;
SetField('PAS_DECLARNOM'         ,RechDom('PGDECLARANTNOM'   ,GetControlText('DECLARANT'),False));
SetField('PAS_DECLARPRENOM'      ,RechDom('PGDECLARANTPRENOM',GetControlText('DECLARANT'),False));
SetField('PAS_DECLARPERS'        ,RechDom('PGDECLARANTATTEST',GetControlText('DECLARANT'),False));
SetField('PAS_DECLARLIEU'        ,RechDom('PGDECLARANTVILLE' ,GetControlText('DECLARANT'),False));
SetField('PAS_DECLARTEL'         ,RechDom('PGDECLARANTTEL  ' ,GetControlText('DECLARANT'),False));
St := RechDom('PGDECLARANTQUAL'  ,GetControlText('DECLARANT'),False);
if St = 'AUT' then SetField('PAS_DECLARQUAL' ,RechDom('PGDECLARANTAUTRE' ,GetControlText('DECLARANT'),False))
else               SetField('PAS_DECLARQUAL' ,RechDom('PGQUALDECLARANT2' ,St,False));
end;

(*
Mise en commentaire PT45
{$IFDEF EAGLCLIENT}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 10/11/2005
Modifié le ... :   /  /
Description .. : Mise à jour des Label à partir de la grille Salaires
Mots clefs ... : PAIE;PGATTESTATION;PGASSEDIC
*****************************************************************}
procedure TOM_Attestations.ESalaires();
begin
SetControlText('SAL_TET_DEB', Salaires.CellValues[0,0]);
SetControlText('SAL_TET_FIN', Salaires.CellValues[1,0]);
SetControlText('SAL_TET_PAI', Salaires.CellValues[2,0]);
SetControlText('SAL_TET_NBH', Salaires.CellValues[3,0]);
SetControlText('SAL_TET_J', Salaires.CellValues[4,0]);
SetControlText('SAL_TET_SAL', Salaires.CellValues[5,0]);
SetControlText('SAL_TET_PRE', Salaires.CellValues[6,0]);
SetControlText('SAL_TET_OBS', Salaires.CellValues[7,0]);
SetControlText('SAL_DET1_DEB', Salaires.CellValues[0,1]);
SetControlText('SAL_DET1_FIN', Salaires.CellValues[1,1]);
SetControlText('SAL_DET1_PAI', Salaires.CellValues[2,1]);
SetControlText('SAL_DET1_NBH', Salaires.CellValues[3,1]);
SetControlText('SAL_DET1_J', Salaires.CellValues[4,1]);
SetControlText('SAL_DET1_SAL', Salaires.CellValues[5,1]);
SetControlText('SAL_DET1_PRE', Salaires.CellValues[6,1]);
SetControlText('SAL_DET1_OBS', Salaires.CellValues[7,1]);
SetControlText('SAL_DET2_DEB', Salaires.CellValues[0,2]);
SetControlText('SAL_DET2_FIN', Salaires.CellValues[1,2]);
SetControlText('SAL_DET2_PAI', Salaires.CellValues[2,2]);
SetControlText('SAL_DET2_NBH', Salaires.CellValues[3,2]);
SetControlText('SAL_DET2_J', Salaires.CellValues[4,2]);
SetControlText('SAL_DET2_SAL', Salaires.CellValues[5,2]);
SetControlText('SAL_DET2_PRE', Salaires.CellValues[6,2]);
SetControlText('SAL_DET2_OBS', Salaires.CellValues[7,2]);
SetControlText('SAL_DET3_DEB', Salaires.CellValues[0,3]);
SetControlText('SAL_DET3_FIN', Salaires.CellValues[1,3]);
SetControlText('SAL_DET3_PAI', Salaires.CellValues[2,3]);
SetControlText('SAL_DET3_NBH', Salaires.CellValues[3,3]);
SetControlText('SAL_DET3_J', Salaires.CellValues[4,3]);
SetControlText('SAL_DET3_SAL', Salaires.CellValues[5,3]);
SetControlText('SAL_DET3_PRE', Salaires.CellValues[6,3]);
SetControlText('SAL_DET3_OBS', Salaires.CellValues[7,3]);
SetControlText('SAL_DET4_DEB', Salaires.CellValues[0,4]);
SetControlText('SAL_DET4_FIN', Salaires.CellValues[1,4]);
SetControlText('SAL_DET4_PAI', Salaires.CellValues[2,4]);
SetControlText('SAL_DET4_NBH', Salaires.CellValues[3,4]);
SetControlText('SAL_DET4_J', Salaires.CellValues[4,4]);
SetControlText('SAL_DET4_SAL', Salaires.CellValues[5,4]);
SetControlText('SAL_DET4_PRE', Salaires.CellValues[6,4]);
SetControlText('SAL_DET4_OBS', Salaires.CellValues[7,4]);
SetControlText('SAL_DET5_DEB', Salaires.CellValues[0,5]);
SetControlText('SAL_DET5_FIN', Salaires.CellValues[1,5]);
SetControlText('SAL_DET5_PAI', Salaires.CellValues[2,5]);
SetControlText('SAL_DET5_NBH', Salaires.CellValues[3,5]);
SetControlText('SAL_DET5_J', Salaires.CellValues[4,5]);
SetControlText('SAL_DET5_SAL', Salaires.CellValues[5,5]);
SetControlText('SAL_DET5_PRE', Salaires.CellValues[6,5]);
SetControlText('SAL_DET5_OBS', Salaires.CellValues[7,5]);
SetControlText('SAL_DET6_DEB', Salaires.CellValues[0,6]);
SetControlText('SAL_DET6_FIN', Salaires.CellValues[1,6]);
SetControlText('SAL_DET6_PAI', Salaires.CellValues[2,6]);
SetControlText('SAL_DET6_NBH', Salaires.CellValues[3,6]);
SetControlText('SAL_DET6_J', Salaires.CellValues[4,6]);
SetControlText('SAL_DET6_SAL', Salaires.CellValues[5,6]);
SetControlText('SAL_DET6_PRE', Salaires.CellValues[6,6]);
SetControlText('SAL_DET6_OBS', Salaires.CellValues[7,6]);
SetControlText('SAL_DET7_DEB', Salaires.CellValues[0,7]);
SetControlText('SAL_DET7_FIN', Salaires.CellValues[1,7]);
SetControlText('SAL_DET7_PAI', Salaires.CellValues[2,7]);
SetControlText('SAL_DET7_NBH', Salaires.CellValues[3,7]);
SetControlText('SAL_DET7_J', Salaires.CellValues[4,7]);
SetControlText('SAL_DET7_SAL', Salaires.CellValues[5,7]);
SetControlText('SAL_DET7_PRE', Salaires.CellValues[6,7]);
SetControlText('SAL_DET7_OBS', Salaires.CellValues[7,7]);
SetControlText('SAL_DET8_DEB', Salaires.CellValues[0,8]);
SetControlText('SAL_DET8_FIN', Salaires.CellValues[1,8]);
SetControlText('SAL_DET8_PAI', Salaires.CellValues[2,8]);
SetControlText('SAL_DET8_NBH', Salaires.CellValues[3,8]);
SetControlText('SAL_DET8_J', Salaires.CellValues[4,8]);
SetControlText('SAL_DET8_SAL', Salaires.CellValues[5,8]);
SetControlText('SAL_DET8_PRE', Salaires.CellValues[6,8]);
SetControlText('SAL_DET8_OBS', Salaires.CellValues[7,8]);
SetControlText('SAL_DET9_DEB', Salaires.CellValues[0,9]);
SetControlText('SAL_DET9_FIN', Salaires.CellValues[1,9]);
SetControlText('SAL_DET9_PAI', Salaires.CellValues[2,9]);
SetControlText('SAL_DET9_NBH', Salaires.CellValues[3,9]);
SetControlText('SAL_DET9_J', Salaires.CellValues[4,9]);
SetControlText('SAL_DET9_SAL', Salaires.CellValues[5,9]);
SetControlText('SAL_DET9_PRE', Salaires.CellValues[6,9]);
SetControlText('SAL_DET9_OBS', Salaires.CellValues[7,9]);
SetControlText('SAL_DET10_DEB', Salaires.CellValues[0,10]);
SetControlText('SAL_DET10_FIN', Salaires.CellValues[1,10]);
SetControlText('SAL_DET10_PAI', Salaires.CellValues[2,10]);
SetControlText('SAL_DET10_NBH', Salaires.CellValues[3,10]);
SetControlText('SAL_DET10_J', Salaires.CellValues[4,10]);
SetControlText('SAL_DET10_SAL', Salaires.CellValues[5,10]);
SetControlText('SAL_DET10_PRE', Salaires.CellValues[6,10]);
SetControlText('SAL_DET10_OBS', Salaires.CellValues[7,10]);
SetControlText('SAL_DET11_DEB', Salaires.CellValues[0,11]);
SetControlText('SAL_DET11_FIN', Salaires.CellValues[1,11]);
SetControlText('SAL_DET11_PAI', Salaires.CellValues[2,11]);
SetControlText('SAL_DET11_NBH', Salaires.CellValues[3,11]);
SetControlText('SAL_DET11_J', Salaires.CellValues[4,11]);
SetControlText('SAL_DET11_SAL', Salaires.CellValues[5,11]);
SetControlText('SAL_DET11_PRE', Salaires.CellValues[6,11]);
SetControlText('SAL_DET11_OBS', Salaires.CellValues[7,11]);
SetControlText('SAL_DET12_DEB', Salaires.CellValues[0,12]);
SetControlText('SAL_DET12_FIN', Salaires.CellValues[1,12]);
SetControlText('SAL_DET12_PAI', Salaires.CellValues[2,12]);
SetControlText('SAL_DET12_NBH', Salaires.CellValues[3,12]);
SetControlText('SAL_DET12_J', Salaires.CellValues[4,12]);
SetControlText('SAL_DET12_SAL', Salaires.CellValues[5,12]);
SetControlText('SAL_DET12_PRE', Salaires.CellValues[6,12]);
SetControlText('SAL_DET12_OBS', Salaires.CellValues[7,12]);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 02/04/2001
Modifié le ... :   /  /
Description .. : Sortie d'une grille
Mots clefs ... : PAIE;PGATTESTATION;PGASSEDIC
*****************************************************************}
procedure TOM_Attestations.EPrimes();
begin
SetControlText('PRI_TET_DEB', Primes.CellValues[0,0]);
SetControlText('PRI_TET_FIN', Primes.CellValues[1,0]);
SetControlText('PRI_TET_OBS', Primes.CellValues[2,0]);
SetControlText('PRI_TET_PAI', Primes.CellValues[3,0]);
SetControlText('PRI_TET_SAL', Primes.CellValues[4,0]);
SetControlText('PRI_DET1_DEB', Primes.CellValues[0,1]);
SetControlText('PRI_DET1_FIN', Primes.CellValues[1,1]);
SetControlText('PRI_DET1_OBS', Primes.CellValues[2,1]);
SetControlText('PRI_DET1_PAI', Primes.CellValues[3,1]);
SetControlText('PRI_DET1_SAL', Primes.CellValues[4,1]);
SetControlText('PRI_DET2_DEB', Primes.CellValues[0,2]);
SetControlText('PRI_DET2_FIN', Primes.CellValues[1,2]);
SetControlText('PRI_DET2_OBS', Primes.CellValues[2,2]);
SetControlText('PRI_DET2_PAI', Primes.CellValues[3,2]);
SetControlText('PRI_DET2_SAL', Primes.CellValues[4,2]);
SetControlText('PRI_DET3_DEB', Primes.CellValues[0,3]);
SetControlText('PRI_DET3_FIN', Primes.CellValues[1,3]);
SetControlText('PRI_DET3_OBS', Primes.CellValues[2,3]);
SetControlText('PRI_DET3_PAI', Primes.CellValues[3,3]);
SetControlText('PRI_DET3_SAL', Primes.CellValues[4,3]);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 02/04/2001
Modifié le ... :   /  /
Description .. : Sortie d'une grille
Mots clefs ... : PAIE;PGATTESTATION;PGASSEDIC
*****************************************************************}
procedure TOM_Attestations.ESolde();
begin
SetControlText('SOL_TET_DEBUT', Solde.CellValues[0,0]);
SetControlText('SOL_TET_FIN', Solde.CellValues[1,0]);
SetControlText('SOL_TET_NBH', Solde.CellValues[3,0]);
SetControlText('SOL_TET_SAL', Solde.CellValues[4,0]);
SetControlText('SOL_TET_DAT', Solde.CellValues[2,0]);
SetControlText('SOL_TET_PRE', Solde.CellValues[5,0]);
SetControlText('SOL_DET1_DEBUT', Solde.CellValues[0,1]);
SetControlText('SOL_DET1_FIN', Solde.CellValues[1,1]);
SetControlText('SOL_DET1_NBH', Solde.CellValues[3,1]);
SetControlText('SOL_DET1_SAL', Solde.CellValues[4,1]);
SetControlText('SOL_DET1_DAT', Solde.CellValues[2,1]);
SetControlText('SOL_DET1_PRE', Solde.CellValues[5,1]);
SetControlText('SOL_DET2_DEBUT', Solde.CellValues[0,2]);
SetControlText('SOL_DET2_FIN', Solde.CellValues[1,2]);
SetControlText('SOL_DET2_NBH', Solde.CellValues[3,2]);
SetControlText('SOL_DET2_SAL', Solde.CellValues[4,2]);
SetControlText('SOL_DET2_DAT', Solde.CellValues[2,2]);
SetControlText('SOL_DET2_PRE', Solde.CellValues[5,2]);
end;
{$ENDIF}
*)

Initialization
registerclasses([TOM_Attestations]) ;

end.

