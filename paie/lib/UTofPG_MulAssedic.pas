{***********UNITE*************************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 13/07/2000
Modifié le ... : 03/04/2002
Description .. : Multi-critère de l'attestation ASSEDIC
Mots clefs ... : PAIE;PGATTESTATION;ASSEDIC
*****************************************************************}
{
PT1   : 27/06/2001 SB V535 rafraîchissement de la liste
                           Malgré que ce soit trés sympa, on laisse libre le
                           rafraîchissement de la liste
PT2   : 22/08/2001 VG V547 Fermeture d'une requête non fermée
PT3   : 01/03/2002 VG V571 La liste des salariés ne s'affiche pas sous ORACLE
                           Fiche qualité N° 10028 + 10029
PT4   : 03/04/2002 VG V571 Affichage de tous les salariés
                           Fiche qualité N° 10069
PT5   : 17/06/2002 VG V582 Procédure de recopie du champ PAS_INDCONV vers
                           PAS_INDCONV2 et du champ PAS_INDTRANS vers
                           PAS_INDTRANS2
PT6   : 15/07/2002 VG V585 Rectification du Multi-critère : L'écran "sautait"
                           lorsqu'on modifiait un critère de sélection
PT7   : 03/10/2002 VG V585 Traitement dans le cas ou le salarié n'est pas sorti
PT8   : 20/02/2003 VG V_42 Cas du double-click sur une liste vide
PT9   : 13/06/2003 VG V_42 Correction du problème de l'erreur SQL lorsqu'on met
                           à jour le montant des indemnités conventionnelles ou
                           transactionnelles avec une valeur alphanumérique
PT10  : 09/02/2004 VG V_50 Cas de l'existance de différentes périodes de préavis
                           FQ N°10867
PT11  : 29/03/2006 EPI V_65 Ajout gestion des processus (FQ N°12791)
													 	Passage numéro salarié en paramètre
PT12  : 19/04/2006 EPI V_65 Gestion abandon en cours 
PT13  : 03/05/2006 VG V_65 Gestion multi-attestations par salarié - FQ N°12425
                           et FQ N°12704 
PT14  : 24/11/2006 VG V_70 Correction suite à la multi-édition des attestations
                           FQ N°13700
PT15  : 31/05/2007 GGU V_72 Unification de la version EAGL et de la version
                            normale de l'attestation ASSEDIC
PT16  : 01/06/2007 GGU V_72 Création d'un bouton pour appeler la fiche d'edition
                            en lot des attestations et gestion du calcul en lot
PT17  : 02/07/2007 GGU V_8 FQ14591 si le salarié n'a pas 12 périodes, les lignes
                           de salaires, primes et solde se décalent et se
                           répetent.
PT18  : 05/09/2007 VG V_80 Traitement des "attestations à faire" par lot
                           FQ N°14373
PT19  : 02/10/2007 VG V_80 Rechercher d'abord les infos dans le contrat de
                           travail puis dans la fiche du salarié - FQ N°12301
}
unit UTofPG_MulAssedic;

interface
uses Controls, Classes, sysutils, HTB97, UTob,
{$IFNDEF EAGLCLIENT}
     db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} HDB, Mul, Fe_Main,
{$ELSE}
     MaineAgl, emul,
{$ENDIF}
     HCtrls, HEnt1, EntPaie, Hqry, UTOF,
     PgOutils, PgOutils2;

type
  TOF_PGMULASSEDIC = class(TOF)
  private
    WW: THEdit; // Clause XX_WHERE
    SAL: THEdit;
    AnneeSortie, Situation: THValComboBox;
    Q_Mul: THQuery; // Query pour changer la liste associee
    BCherche: TToolbarButton97;
    AppelProc : boolean;       // PT11 appel processus
{$IFNDEF EAGLCLIENT}
    Grille: THDBGrid;
{$ELSE}
    Grille: THGrid;
{$ENDIF}
    procedure AfaireChange(Sender: TObject);
    procedure SalarieExit(Sender: TObject);
    procedure ActiveWhere(Sender: TObject);
    procedure GrilleDblClick(Sender: TObject);
    procedure MAJPreavis;
    // PT12 ajout fonction de sortie
    procedure ClickSortie (Sender: TObject);
    procedure OnVoirListeAttestClick(Sender: TObject); //PT16
    procedure CreateIndemnites(TobSalarie : TOB; Ordre : Integer);
    procedure CreatePrimes(TobSalarie : TOB; Ordre : Integer);
    procedure CreateSalaires(TobSalarie : TOB; Ordre : Integer);
    procedure CreateSolde(TobSalarie : TOB; Ordre : Integer);
    Function GetOrdre() : Integer;
  public
    procedure CalculAttest(Sender: TObject); //PT16
    procedure CalculAttestLot(Sender: TObject); //PT16
    procedure OnArgument(Arguments: string); override;
    procedure OnLoad; override;
    procedure OnUpdate; override;
  end;

var PGAttesSalarie: string;
implementation

Uses
  HmsgBox, ed_tools, HStatus, StrUtils ,UTofPG_MULEDITATTESTLOT, UTOMAttestations;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 15/07/2002
Modifié le ... :   /  /
Description .. : OnLoad
Mots clefs ... : PAIE;PGATTESTATION;PGASSEDIC
*****************************************************************}
procedure TOF_PGMULASSEDIC.OnLoad;
begin
inherited;
ActiveWhere(nil);
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 13/07/2000
Modifié le ... :   /  /
Description .. : ActiveWhere
Mots clefs ... : PAIE;PGATTESTATION;PGASSEDIC
*****************************************************************}
procedure TOF_PGMULASSEDIC.ActiveWhere(Sender: TObject);
var
Annee: string;
begin

if (GetControlText('ANNEESORTIE') <> '') then
   Annee:= AnneeSortie.Text;

if (Annee <> '') then
   begin
   if (WW <> nil) and (GetControlText('SITUATION') = '1') then
      begin // liste des salaries sortis
      if Q_Mul <> nil then
         TFMul(Ecran).SetDBListe('PGMULSALASSEDIC');
      WW.Text:= ' PSA_DATESORTIE >= "'+UsDateTime(StrToDate('01/01/'+Annee))+'" AND'+
                ' PSA_DATESORTIE <= "'+UsDateTime(StrToDate('31/12/'+Annee))+'"';
      end;

   if (WW <> nil) and (GetControlText('SITUATION') = '2') then
      begin // liste des salaries n'ayant pas d'attestations assedic
      if Q_Mul <> nil then
         TFMul(Ecran).SetDBListe('PGMULSALARIE');
      WW.Text:= 'PSA_SALARIE NOT IN (SELECT PAS_SALARIE FROM ATTESTATIONS WHERE'+
                ' PAS_TYPEATTEST = "ASS") AND'+
                ' PSA_DATESORTIE >= "'+UsDateTime(StrToDate('01/01/'+Annee))+'" AND'+
                ' PSA_DATESORTIE <= "'+UsDateTime(StrToDate('31/12/'+Annee))+'"';
      end;

   if (WW <> nil) and (GetControlText('SITUATION') = '3') then
      begin // liste des Attestations
      if Q_Mul <> nil then
         TFMul(Ecran).SetDBListe('PGASSEDICFAITES');
{PT14
      WW.Text:= ' PAS_TYPEATTEST = "ASS" AND'+
                ' PSA_DATESORTIE >= "'+UsDateTime(StrToDate('01/01/'+Annee))+'" AND'+
                ' PSA_DATESORTIE <= "'+UsDateTime(StrToDate('31/12/'+Annee))+'"';
}
      if ((Situation <> nil) and (Situation.Enabled=False)) then
         WW.Text:= ' PAS_TYPEATTEST="ASS" AND'+
                   ' PSA_DATESORTIE<="'+UsDateTime(StrToDate('31/12/'+Annee))+'"'
      else
         WW.Text:= ' PAS_TYPEATTEST="ASS" AND'+
                   ' PSA_DATESORTIE>="'+UsDateTime(StrToDate('01/01/'+Annee))+'" AND'+
                   ' PSA_DATESORTIE<="'+UsDateTime(StrToDate('31/12/'+Annee))+'"';
//FIN PT14
      end;

   if (WW <> nil) and (GetControlText('SITUATION') = '4') then
      begin // liste des salaries
      WW.Text:= '';
      if Q_Mul <> nil then
         TFMul(Ecran).SetDBListe('PGMULSALASSEDIC');
// PT11 début
// PT12 If AppelProc = True then
// PT12   WW.Text := SelectSal;
// PT11 fin
      end;
   end;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 13/07/2000
Modifié le ... :   /  /
Description .. :
Mots clefs ... : PAIE;PGATTESTATION;PGASSEDIC
*****************************************************************}
procedure TOF_PGMULASSEDIC.OnArgument(Arguments: string);
var
MoisE, AnneeE, ComboExer: string;
DebExer, FinExer: TDateTime;
// PT12 ajout bouton STOP
BTNSTOP : TToolbarButton97;
Txt : String;                  // PT12
begin
  inherited;
  // PT11 début
  AppelProc:= False;
  // PT12 SelectSal := Arguments;
  Txt:= Arguments;  // PT12
  // PT11 fin
  MAJAttestations;
  MAJPreavis;
  Q_Mul:= THQuery(Ecran.FindComponent('Q'));
  {$IFNDEF EAGLCLIENT}
  Grille:= THDBGrid(GetControl('Fliste'));
  {$ELSE}
  Grille:= THGrid(GetControl('Fliste'));
  {$ENDIF}
  if Grille <> nil then
  begin
    Grille.OnDblClick:= GrilleDblClick;
  end;
  TFMul(Ecran).BInsert.OnClick := GrilleDblClick;  //PT13


  SAL:= THEdit(GetControl('PSA_SALARIE'));
  if SAL <> nil then
     SAL.OnExit:= SalarieExit;

  AnneeSortie:= THValComboBox(GetControl('ANNEESORTIE'));
  if RendExerSocialEnCours(MoisE, AnneeE, ComboExer, DebExer, FinExer) = TRUE then
     if AnneeSortie <> nil then
        AnneeSortie.value:= ComboExer;

  Situation:= THValComboBox(GetControl('SITUATION'));
  if Situation <> nil then
     Situation.OnChange:= AFaireChange;
  SetControlText('SITUATION', '1');
  WW:= THEdit(GetControl('XX_WHERE'));
  // PT11 début
  // PT12 if (WW <> nil) and (SelectSal <> '') then
  if (WW <> nil) and (Txt = 'P') then
     begin
     AppelProc:= True;
  // PT12 WW.text := SelectSal;
     SetControlText('SITUATION', '4');
     End;
  // PT11 fin

  // PT12 début
  If AppelProc = True then
     begin
     SetControlVisible('BTNSTOP', TRUE);
     BTNSTOP:= TToolbarButton97 (GetControl ('BTNSTOP'));
     if BTNSTOP <> NIL then
        BTNSTOP.Onclick:= ClickSortie;
     end;
  // PT12 fin

  //PT13
  if (Txt='S') then
     begin
     SetControlText ('PSA_SALARIE', PGAttesSalarie);
     SetControlText ('SITUATION', '3');
     SetControlEnabled ('PSA_SALARIE', False);
     SetControlEnabled ('PSA_LIBELLE', False);
     SetControlEnabled ('PSA_ETABLISSEMENT', False);
     SetControlEnabled ('ANNEESORTIE', False);
     SetControlEnabled ('SITUATION', False);
     end;
  //FIN PT13

  BCherche:= TToolbarButton97(GetControl('BCherche'));
  ActiveWhere(nil);

  //PT16 Bouton d'appel du mul d'édition en lot des attestations ASSEDIC
  (GetControl('BEDITLOT') as TToolbarButton97).OnClick := OnVoirListeAttestClick;
  //PT16 Bouton du calcul en lot des attestations ASSEDIC
  (GetControl('BCALCULLOT') as TToolbarButton97).OnClick := CalculAttestLot;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 13/07/2000
Modifié le ... : 22/08/2001
Description .. : Double click sur la grille du multi-critère
Mots clefs ... : PAIE;PGATTESTATION;PGASSEDIC;CLICK
*****************************************************************}
procedure TOF_PGMULASSEDIC.GrilleDblClick(Sender: TObject);
begin
  if (Sender=Grille) then  //Si on viens du DoubleClick:
  begin
    //Si une seule ligne est sélectionnée, on fait le traitement normale (création - modification)
    if (Grille.nbSelected <= 1) and (not Grille.AllSelected) then
      CalculAttest(Sender)
    else //Si on a une multisélection , on fait le traitement par lot
      CalculAttestLot(Sender);
  end else begin   //Si on viens du bouton ouvrir ou insert:
    //Si une seule ligne est sélectionnée, on fait le traitement normale (création - modification)
    if (Grille.nbSelected <= 1) and (not Grille.AllSelected) then
      CalculAttest(Sender)
    else //Si on a une multisélection , on affiche un message d'erreur
      PGIInfo(TraduireMemoire('L''édition d''une attestation n''est pas possible en multi-sélection'),Ecran.caption);
  end;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 13/07/2000
Modifié le ... :   /  /
Description .. :
Mots clefs ... : PAIE;PGATTESTATION;PGASSEDIC
*****************************************************************}
procedure TOF_PGMULASSEDIC.AfaireChange(Sender: TObject);
var
TT: TFMul;
i: Integer;
begin
if (GetControlText('SITUATION') <> '') then
   i:= StrToInt(GetControlText('SITUATION'))
else
   begin
// PT11 i := 1;
// PT11 SetControlText('SITUATION', '1');
// PT11 début
   if AppelProc = False then
      begin
      i:= 1;
      SetControlText('SITUATION', '1');
      end
   else
      begin
      i:= 4;
      SetControlText('SITUATION', '4');
      end;
// PT11 fin
   end;

  case i of
     1: TFMul(Ecran).Caption:= 'Liste des Salariés sortis';
     2: TFMul(Ecran).Caption:= 'Liste des Attestations à préparer';
     3: TFMul(Ecran).Caption:= 'Liste des Attestations effectuées';
     4: TFMul(Ecran).Caption:= 'Liste des Salariés'; //PT4
  end;
  SetControlEnabled('BCALCULLOT',False);

TT:= TFMul(Ecran);
if TT <> nil then
   UpdateCaption(TT);
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 13/07/2000
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : PAIE;PGATTESTATION;PGASSEDIC
*****************************************************************}
procedure TOF_PGMULASSEDIC.SalarieExit(Sender: TObject);
begin
if (isnumeric(SAL.Text) and (VH_PAIE.PgTypeNumSal = 'NUM')) then
   SAL.Text:= ColleZeroDevant(StrToInt(SAL.Text), 10);
end;

//PT10
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 09/02/2004
Modifié le ... :
Description .. : Mise à jour des nouveaux champs avec les valeurs déjà
Suite ........ : existantes
Mots clefs ... : PAIE;PGATTESTATION;PGASSEDIC
*****************************************************************}
procedure TOF_PGMULASSEDIC.MAJPreavis;
var
StDonnees, StInsert: string;
Q: TQuery;
DateBuf, DateBuf2: TDateTime;
begin
StInsert:= 'UPDATE ATTESTATIONS SET PAS_MOTIFARRET="", PAS_MOTIFPREAVIS1="X"'+
           ' WHERE PAS_MOTIFARRET="1" AND'+
           ' PAS_TYPEATTEST="ASS"';
ExecuteSQL(StInsert);

StDonnees:= 'SELECT PAS_ORDRE, PAS_DATEARRET, PAS_REPRISEARRET'+
            ' FROM ATTESTATIONS WHERE'+
            ' PAS_MOTIFARRET = "2" AND'+
            ' PAS_TYPEATTEST="ASS"';
Q:= OpenSQL(StDonnees, True);
while not Q.EOF do
      begin
      DateBuf:= Q.FindField('PAS_DATEARRET').AsDateTime;
      DateBuf2:= Q.FindField('PAS_REPRISEARRET').AsDateTime;
      StInsert:= 'UPDATE ATTESTATIONS SET PAS_MOTIFARRET="",'+
                 ' PAS_DATEARRET="'+UsDateTime(IDate1900)+'",'+
                 ' PAS_REPRISEARRET="'+UsDateTime(IDate1900)+'",'+
                 ' PAS_MOTIFPREAVIS2="X",'+
                 ' PAS_DATEPREAVISD2="'+UsDateTime(DateBuf)+'",'+
                 ' PAS_DATEPREAVISF2="'+UsDateTime(DateBuf2)+'" WHERE'+
                 ' PAS_ORDRE='+IntToStr(Q.FindField('PAS_ORDRE').AsInteger)+' AND'+
                 ' PAS_TYPEATTEST="ASS"';
      ExecuteSQL(StInsert);
      Q.Next;
      end;
Ferme(Q);

StDonnees:= 'SELECT PAS_ORDRE, PAS_DATEARRET, PAS_REPRISEARRET'+
            ' FROM ATTESTATIONS WHERE'+
            ' PAS_MOTIFARRET = "3" AND'+
            ' PAS_TYPEATTEST="ASS"';
Q:= OpenSQL(StDonnees, True);
while not Q.EOF do
      begin
      DateBuf:= Q.FindField('PAS_DATEARRET').AsDateTime;
      DateBuf2:= Q.FindField('PAS_REPRISEARRET').AsDateTime;
      StInsert:= 'UPDATE ATTESTATIONS SET PAS_MOTIFARRET="",'+
                 ' PAS_DATEARRET="'+UsDateTime(IDate1900)+'",'+
                 ' PAS_REPRISEARRET="'+UsDateTime(IDate1900)+'",'+
                 ' PAS_MOTIFPREAVIS3="X",'+
                 ' PAS_DATEPREAVISD3="'+UsDateTime(DateBuf)+'",'+
                 ' PAS_DATEPREAVISF3="'+UsDateTime(DateBuf2)+'" WHERE'+
                 ' PAS_ORDRE='+IntToStr(Q.FindField('PAS_ORDRE').AsInteger)+' AND'+
                 ' PAS_TYPEATTEST="ASS"';
      ExecuteSQL(StInsert);
      Q.Next;
      end;
Ferme(Q);
end;
//FIN PT10

// PT12
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Eliane PION
Créé le ...... : 19/04/2006
Modifié le ... :   /  /    
Description .. : Gestion arret du processus
Mots clefs ... : PAIE;PGATTESTATION;PGASSEDIC
*****************************************************************}
procedure TOF_PGMULASSEDIC.ClickSortie(Sender: TObject);
var BTNAnn :  TToolbarButton97;
begin
BTNAnn:= TToolbarButton97 (GetControl ('BAnnuler'));
if BTNAnn <> NIL then
   begin
   TFMul(Ecran).Retour:= 'STOP';
   BTNAnn.Click;
   end;
end;

//PT16 Bouton d'appel du mul d'édition en lot des attestations ASSEDIC
procedure TOF_PGMULASSEDIC.OnVoirListeAttestClick(Sender: TObject);
begin
  AglLanceFiche('PAY', 'MULEDITATTESTLOT', '', '', GetControlText('PSA_SALARIE')+';'+GetControlText('PSA_LIBELLE')+';'+GetControlText('PSA_ETABLISSEMENT')+';'+GetControlText('ANNEESORTIE')+';');
end;

procedure TOF_PGMULASSEDIC.CalculAttestLot(Sender: TObject);
var
i, indexSalarie, NbSal, Ordre : Integer;
AnneePrec, Declarant, DeclarantQual, FonctionSal, ListeSalForEdition : String;
ListSal, MotifSortie, StDate, StSQL, TypeContrat, TypeContratPart : String;
Q, QEntree, QRechSalMSA, QSortie : TQuery;
TobSalarie, TobSalaries : Tob;
JourJ : TDateTime;
AnneeA, Jour, MoisM : Word;
HadFailed : Boolean;

  Procedure AddSal(Q : THQuery; var Liste : String);
  var
  Salarie : String;
  begin
  Salarie := Q.FindField('PSA_SALARIE').asstring;
  Liste := Liste + '"'+Salarie+'", ';
  end;
begin
HadFailed:= False;
if (Grille.nbSelected > 0) and (not Grille.AllSelected) then
   begin
   InitMoveProgressForm (nil, 'Début du traitement',
                         'Veuillez patienter SVP ...', Grille.nbSelected, FALSE,
                         TRUE);
   InitMove (Grille.nbSelected, '');
   for i:= 0 to Grille.NbSelected-1 do
       begin
       Grille.GotoLeBOOKMARK(i);
{$IFDEF EAGLCLIENT}
       TFmul (Ecran).Q.TQ.Seek (Grille.Row-1);
{$ENDIF}
       AddSal (Q_Mul, ListSal);
       MoveCurProgressForm ();
       end;
   FiniMoveProgressForm;
   end
else
if Grille.AllSelected then
   begin
{$IFDEF EAGLCLIENT}
   if (TFMul(Ecran).bSelectAll.Down) then
      TFMul(Ecran).Fetchlestous; //PT3
{$ENDIF}
   InitMoveProgressForm (nil, 'Début du traitement',
                         'Veuillez patienter SVP ...',
                         TFmul(Ecran).Q.RecordCount, FALSE, TRUE);
   InitMove (TFmul (Ecran).Q.RecordCount, '');
   Q_Mul.First;
   while not Q_Mul.EOF do
         begin
         AddSal (Q_Mul, ListSal);
         Q_Mul.Next;
         end;
   FiniMoveProgressForm;
   end;

if (ListSal='') then
   exit;
ListSal:= LeftStr (ListSal,Length (ListSal)-2);
//Boucle sur chaque salarié
TobSalaries:= Tob.Create ('TobSalaries', nil, -1);
TobSalaries.LoadDetailFromSQL ('SELECT PSA_SALARIE, PSA_LIBELLE, PSA_PRENOM,'+
                               ' PSA_ETABLISSEMENT, PSA_NUMEROSS,'+
                               ' PSA_LIBELLEEMPLOI, PSA_STATUT,'+
                               ' PSA_DATEENTREE, PSA_DATESORTIE, PSA_HORHEBDO,'+
                               ' PSA_HORANNUEL, PSA_MOTIFSORTIE, PSA_CATDADS,'+
                               ' PSA_FONCTIONSAL'+
                               ' FROM SALARIES WHERE'+
                               ' PSA_SALARIE IN ('+ListSal+')');
for indexSalarie:= 0 to TobSalaries.FillesCount (0)-1 do
    begin
// récupérer les infos nécéssaire au calcul de l'attestation
    TobSalarie:= TobSalaries.Detail[indexSalarie];
//Initialisation des champs aux valeurs par défaut
    TobSalarie.AddChampSupValeur ('PAS_SALARIE',
                                  TobSalarie.GetString ('PSA_SALARIE'));
    TobSalarie.AddChampSupValeur ('PAS_TYPEATTEST', 'ASS');
    TobSalarie.AddChampSupValeur ('PAS_TRAVAILTEMP', '-');
    TobSalarie.AddChampSupValeur ('PAS_DATEATTEST', IDate1900);
    TobSalarie.AddChampSupValeur ('PAS_PUBLIC', '-');
    TobSalarie.AddChampSupValeur ('PAS_PUBLICTYPE', '');
    TobSalarie.AddChampSupValeur ('PAS_CONVNUMERO', '');
    TobSalarie.AddChampSupValeur ('PAS_ADHESION', '-');
    TobSalarie.AddChampSupValeur ('PAS_CONVCODEANA', '');
    TobSalarie.AddChampSupValeur ('PAS_DATEADHESION', iDate1900);
    TobSalarie.AddChampSupValeur ('PAS_CAISSESS', '');
    TobSalarie.AddChampSupValeur ('PAS_MATRICULE', '');
    TobSalarie.AddChampSupValeur ('PAS_ALSACEMOSEL', '-');
    TobSalarie.AddChampSupValeur ('PAS_QUALIF', '');
    TobSalarie.AddChampSupValeur ('PAS_CADRE', '-');
    TobSalarie.AddChampSupValeur ('PAS_LIENPARENTE', '-');
    TobSalarie.AddChampSupValeur ('PAS_LIBLIENP', '');
    TobSalarie.AddChampSupValeur ('PAS_QUALITEEMPLOI', '');
    TobSalarie.AddChampSupValeur ('PAS_LQUALITE', '');
    TobSalarie.AddChampSupValeur ('PAS_LCONTRATPART', '');
    TobSalarie.AddChampSupValeur ('PAS_COLLTERRITOR', '');
    TobSalarie.AddChampSupValeur ('PAS_FONCTIONSAL', '');
    TobSalarie.AddChampSupValeur ('PAS_LIBFONCTION', '');
    TobSalarie.AddChampSupValeur ('PAS_AUTREVICTIME', '-');
    TobSalarie.AddChampSupValeur ('PAS_DATEACCIDENT', iDate1900);
    TobSalarie.AddChampSupValeur ('PAS_DATEARRET', iDate1900);
    TobSalarie.AddChampSupValeur ('PAS_REPRISEARRET', iDate1900);
    TobSalarie.AddChampSupValeur ('PAS_MOTIFARRET', '');
    TobSalarie.AddChampSupValeur ('PAS_PREAVISMOTIF', '');
    TobSalarie.AddChampSupValeur ('PAS_SITUATION', '');
    TobSalarie.AddChampSupValeur ('PAS_NONREPRIS', '-');
    TobSalarie.AddChampSupValeur ('PAS_MAINTIEN', '-');
    TobSalarie.AddChampSupValeur ('PAS_TYPMAINTIEN', '');
    TobSalarie.AddChampSupValeur ('PAS_REPRISEPARTIEL', '');
    TobSalarie.AddChampSupValeur ('PAS_CASGEN', '');
    TobSalarie.AddChampSupValeur ('PAS_MONTANT', 0);
    TobSalarie.AddChampSupValeur ('PAS_PLUSDE', '-');
    TobSalarie.AddChampSupValeur ('PAS_AUTRERUPT', '');
    TobSalarie.AddChampSupValeur ('PAS_PLANSOC', '-');
    TobSalarie.AddChampSupValeur ('PAS_DATEPLANSOC', iDate1900);
    TobSalarie.AddChampSupValeur ('PAS_RECLASS', '-');
    TobSalarie.AddChampSupValeur ('PAS_DATEPREAVISD2', iDate1900);
    TobSalarie.AddChampSupValeur ('PAS_DATEPREAVISF2', iDate1900);
    TobSalarie.AddChampSupValeur ('PAS_MOTIFPREAVIS2', '-');
    TobSalarie.AddChampSupValeur ('PAS_DATEPREAVISD3', iDate1900);
    TobSalarie.AddChampSupValeur ('PAS_DATEPREAVISF3', iDate1900);
    TobSalarie.AddChampSupValeur ('PAS_MOTIFPREAVIS3', '-');
    TobSalarie.AddChampSupValeur ('PAS_MOTIFPREAVIS1', '-');
    TobSalarie.AddChampSupValeur ('PAS_AGESALARIE', '-');
    TobSalarie.AddChampSupValeur ('PAS_REFUSFNE', '-');
    TobSalarie.AddChampSupValeur ('PAS_CHOMTOTAL', '-');
    TobSalarie.AddChampSupValeur ('PAS_DDTEFP', '-');
    TobSalarie.AddChampSupValeur ('PAS_DEBCHOMAGE', iDate1900);
    TobSalarie.AddChampSupValeur ('PAS_FINCHOMAGE', iDate1900);
    TobSalarie.AddChampSupValeur ('PAS_REPRISECHOM', iDate1900);
    TobSalarie.AddChampSupValeur ('PAS_MOTIFDIFF', '');
    TobSalarie.AddChampSupValeur ('PAS_AUTREDIFF', '');
    TobSalarie.AddChampSupValeur ('PAS_ORGARRCO', '');
    TobSalarie.AddChampSupValeur ('PAS_ORGAGIRC', '');
    TobSalarie.AddChampSupValeur ('PAS_ORGAUTRE', '');
    TobSalarie.AddChampSupValeur ('PAS_INDCPEB', '-');
    TobSalarie.AddChampSupValeur ('PAS_INDCPEM', 0);
    TobSalarie.AddChampSupValeur ('PAS_INDTRANSC', '-');
    TobSalarie.AddChampSupValeur ('PAS_FNGSRED', '-');
    TobSalarie.AddChampSupValeur ('PAS_FNGS', '-');
    TobSalarie.AddChampSupValeur ('PAS_FNGSCRE1', '');
    TobSalarie.AddChampSupValeur ('PAS_FNGSMOT1', '');
    TobSalarie.AddChampSupValeur ('PAS_FNGS2', '-');
    TobSalarie.AddChampSupValeur ('PAS_FNGSCRE2', '');
    TobSalarie.AddChampSupValeur ('PAS_FNGSMOT2', '');
    TobSalarie.AddChampSupValeur ('PAS_DECLARQAUTRE', '');
    TobSalarie.AddChampSupValeur ('PAS_DECLARRENS1', '');
    TobSalarie.AddChampSupValeur ('PAS_DECLARDATE', Date);
    TobSalarie.AddChampSupValeur ('PAS_SUBDEBUT', iDate1900);
    TobSalarie.AddChampSupValeur ('PAS_SUBFIN', iDate1900);
    TobSalarie.AddChampSupValeur ('PAS_SUBINTEGRAL', '-');
    TobSalarie.AddChampSupValeur ('PAS_SUBCOMPTE', '');
    TobSalarie.AddChampSupValeur ('PAS_SUBCPTEINT', '');
    TobSalarie.AddChampSupValeur ('PAS_SUBMONNAIE', '-');
    TobSalarie.AddChampSupValeur ('PAS_TYPEABS', '');

    StSQL:= 'SELECT ET_LIBELLE, ET_ABREGE, ET_ADRESSE1, ET_ADRESSE2,'+
            ' ET_ADRESSE3, ET_CODEPOSTAL, ET_VILLE, ET_PAYS, ET_SIRET, ET_APE,'+
            ' ET_ACTIVITE, ET_NODOSSIER'+
            ' FROM ETABLISS WHERE'+
            ' ET_ETABLISSEMENT="'+TobSalarie.GetString ('PSA_ETABLISSEMENT')+'"';
    Q:= OpenSql (StSQL, TRUE);
    if NOT Q.EOF then
       begin
       TobSalarie.AddChampSupValeur ('PAS_DECLARLIEU',
                                     Q.FindField ('ET_VILLE').AsString);
       TobSalarie.AddChampSupValeur ('PAS_LIEUTRAVAIL',
                                     Q.FindField ('ET_VILLE').AsString);
       TobSalarie.AddChampSupValeur ('PAS_DEPART',
                                     Copy (Q.FindField ('ET_CODEPOSTAL').AsString, 1, 2));
       end;
    Ferme (Q);

    StSQL:= 'SELECT PHE_HORAIREHEURE, PHE_HORAIREAN'+
            ' FROM HORAIREETAB WHERE'+
            ' PHE_ETABLISSEMENT="'+TobSalarie.GetString ('PSA_ETABLISSEMENT')+'"';
    Q:= OpenSql (StSQL,TRUE);
    if NOT Q.EOF then
       begin
       TobSalarie.AddChampSupValeur ('PAS_HORHEBENT',
                                     Q.FindField ('PHE_HORAIREHEURE').AsFloat);
       TobSalarie.AddChampSupValeur ('PAS_HORANNENT',
                                     Q.FindField ('PHE_HORAIREAN').AsFloat);
       end;
    Ferme (Q);

    StSQL:= 'SELECT POG_LIBELLE, POG_NUMAFFILIATION'+
            ' FROM ORGANISMEPAIE WHERE'+
            ' POG_ETABLISSEMENT="'+TobSalarie.GetString ('PSA_ETABLISSEMENT')+'" AND'+
            ' POG_NATUREORG="200"';
    Q:= OpenSql (StSQL,TRUE);
    if NOT Q.EOF then
       begin
       TobSalarie.AddChampSupValeur ('PAS_ASSEDICCAISSE',
                                     leftStr (Q.FindField ('POG_LIBELLE').Asstring, 17));
       TobSalarie.AddChampSupValeur ('PAS_ASSEDICNUM',
                                     Q.FindField ('POG_NUMAFFILIATION').Asstring);
       end;
    Ferme (Q);

    TobSalarie.AddChampSupValeur ('PSA_CATDADS',
                                  TobSalarie.GetString ('PSA_CATDADS'));
    TobSalarie.AddChampSupValeur ('PAS_PERIODEDEBUT',
                                  TobSalarie.GetString ('PSA_DATEENTREE'));
    TobSalarie.AddChampSupValeur ('PAS_PERIODEFIN',
                                  TobSalarie.GetString ('PSA_DATESORTIE'));
    if (TobSalarie.GetDateTime ('PAS_PERIODEFIN')>IDate1900) then
       TobSalarie.AddChampSupValeur ('PAS_DERNIERJOUR',
                                     TobSalarie.GetString ('PAS_PERIODEFIN'))
    else
       TobSalarie.AddChampSupValeur ('PAS_DERNIERJOUR', Date);
    TobSalarie.AddChampSupValeur ('PAS_DERNIEREMPLOI',
                                  TobSalarie.GetString ('PSA_LIBELLEEMPLOI'));
    TobSalarie.AddChampSupValeur ('PAS_HORHEBSAL',
                                  TobSalarie.GetDouble ('PSA_HORHEBDO'));
    TobSalarie.AddChampSupValeur ('PAS_HORANNSAL',
                                  TobSalarie.GetDouble ('PSA_HORANNUEL'));

    StSQL:= 'SELECT PSE_SALARIE, PSE_MSA'+
            ' FROM DEPORTSAL WHERE'+
            ' PSE_SALARIE="'+TobSalarie.GetString ('PSA_SALARIE')+'" AND'+
            ' PSE_MSA="X"';
    QRechSalMSA:= OpenSql(StSQL,TRUE);
    if (NOT QRechSalMSA.EOF) then
       TobSalarie.AddChampSupValeur ('PAS_REGIMESPECSS', 'X')
    else
       TobSalarie.AddChampSupValeur ('PAS_REGIMESPECSS', '-');
    Ferme (QRechSalMSA);
    MotifSortie:= '';
    TypeContrat:= '';
    TypeContratPart:= '';
    FonctionSal:= ''; //PT19
    if (TobSalarie.GetDateTime ('PAS_DERNIERJOUR')<>IDate1900) then
       begin
{PT19
      GetInfoContratTravail( TobSalarie.GetString('PSA_SALARIE')
                           , TobSalarie.GetDateTime('PAS_DERNIERJOUR')
                           , MotifSortie
                           , TypeContrat
                           , TypeContratPart);

      if (TobSalarie.GetString('PSA_MOTIFSORTIE') <> '') then
         MotifSortie:= TobSalarie.GetString('PSA_MOTIFSORTIE');
}
       if (TobSalarie.GetString ('PSA_MOTIFSORTIE')<>'') then
          MotifSortie:= TobSalarie.GetString ('PSA_MOTIFSORTIE');

       if (TobSalarie.GetString ('PSA_FONCTIONSAL')<>'') then
          FonctionSal:= TobSalarie.GetString ('PSA_FONCTIONSAL');

       GetInfoContratTravail (TobSalarie.GetString('PSA_SALARIE'),
                              TobSalarie.GetDateTime('PAS_DERNIERJOUR'),
                              MotifSortie, TypeContrat, TypeContratPart,
                              FonctionSal);
//FIN PT19
       if (MotifSortie='403') then
           MotifSortie:= '';
       TobSalarie.AddChampSupValeur ('PAS_MOTRUPCONT', MotifSortie);
       if (MotifSortie='90') then
          MotifSortie:= '60';
       if (MotifSortie='20') or (MotifSortie='59') or (MotifSortie='60') then
          TobSalarie.AddChampSupValeur ('PAS_DECLARRENS1', '')
       else
          TobSalarie.AddChampSupValeur ('PAS_DECLARRENS1',
                                        RechDom ('PGMOTIFSORTIE', MotifSortie,
                                                 FALSE, ''));
       end;
    TobSalarie.AddChampSupValeur ('PAS_CONTRATNAT', TypeContrat);
    TobSalarie.AddChampSupValeur ('PAS_CONTRATPARTIC', TypeContratPart);
    TobSalarie.AddChampSupValeur ('PAS_FONCTIONSAL', FonctionSal);      //PT19
    JourJ:= Date;
    DecodeDate (JourJ, AnneeA, MoisM, Jour);
    AnneePrec:= IntToStr (AnneeA-1);
    StDate:= UsDateTime (StrToDate ('31/12/'+AnneePrec));
    StSQL:= 'SELECT COUNT(PSA_SALARIE) AS NBENTREE'+
            ' FROM SALARIES WHERE'+
            ' PSA_DATEENTREE<"'+StDate+'" AND'+
            ' PSA_DATEENTREE<>"'+UsDateTime (IDate1900)+'" AND'+
            ' PSA_DATEENTREE IS NOT NULL AND'+
            ' PSA_ETABLISSEMENT="'+TobSalarie.GetString ('PSA_ETABLISSEMENT')+'"';
    QEntree:= OpenSql (StSQL,True);
    StSQL:= 'SELECT COUNT(PSA_SALARIE) AS NBSORTIE'+
            ' FROM SALARIES WHERE'+
            ' PSA_DATESORTIE<"'+StDate+'" AND'+
            ' PSA_DATESORTIE<>"'+UsDateTime (IDate1900)+'" AND'+
            ' PSA_DATESORTIE IS NOT NULL AND'+
            ' PSA_ETABLISSEMENT="'+TobSalarie.GetString ('PSA_ETABLISSEMENT')+'"';
    QSortie:= OpenSql (StSQL, True);
    NbSal:= 0;
    if ((not QEntree.eof) and (not QSortie.eof)) then
       NbSal:= (QEntree.FindField ('NBENTREE').asinteger)-
               (QSortie.FindField ('NBSORTIE').asinteger);
    TobSalarie.AddChampSupValeur ('PAS_EFFECTIF', NbSal);
    Ferme (QEntree);
    Ferme (QSortie);

//Chargement des données salariales
    Ordre:= GetOrdre ();
    TobSalarie.AddChampSupValeur ('PAS_ORDRE', Ordre);
    begintrans;
    try
       CreateIndemnites (TobSalarie, Ordre);
       CreateSalaires (TobSalarie, Ordre);
       CreatePrimes (TobSalarie, Ordre);
       CreateSolde (TobSalarie, Ordre);
       Q:= OpenSql ('SELECT PDA_DECLARANTATTES'+
                    ' FROM DECLARANTATTEST WHERE'+
                    ' (PDA_ETABLISSEMENT = "" OR'+
                    ' PDA_ETABLISSEMENT LIKE "%'+TobSalarie.GetString ('PSA_ETABLISSEMENT')+'%") AND'+
                    ' (PDA_TYPEATTEST="" OR'+
                    ' PDA_TYPEATTEST LIKE "%ASS%" )'+
                    ' ORDER BY PDA_ETABLISSEMENT DESC', True);
       If (Not Q.eof) then
          Begin
          Declarant:= Q.FindField ('PDA_DECLARANTATTES').AsString;
          TobSalarie.AddChampSupValeur ('PAS_DECLARNOM', '');
          TobSalarie.AddChampSupValeur ('PAS_DECLARPRENOM', '');
          TobSalarie.AddChampSupValeur ('PAS_DECLARPERS', '');
          TobSalarie.AddChampSupValeur ('PAS_DECLARLIEU', '');
          TobSalarie.AddChampSupValeur ('PAS_DECLARTEL', '');
          TobSalarie.AddChampSupValeur ('PAS_DECLARQUAL', '');
          if (Declarant<>'') then
             begin
             DeclarantQual:= RechDom ('PGDECLARANTATTEST', Declarant, False);
             TobSalarie.PutValue ('PAS_DECLARNOM',
                                  RechDom ('PGDECLARANTNOM', Declarant, False));
             TobSalarie.PutValue ('PAS_DECLARPRENOM',
                                  RechDom ('PGDECLARANTPRENOM', Declarant, False));
             TobSalarie.PutValue ('PAS_DECLARPERS',
                                  RechDom ('PGDECLARANTATTEST', Declarant, False));
             TobSalarie.PutValue ('PAS_DECLARLIEU',
                                  RechDom ('PGDECLARANTVILLE', Declarant, False));
             TobSalarie.PutValue ('PAS_DECLARTEL',
                                  RechDom ('PGDECLARANTTEL', Declarant, False));
             DeclarantQual:= RechDom ('PGDECLARANTQUAL', Declarant, False);
             if (DeclarantQual='AUT') then
                TobSalarie.PutValue ('PAS_DECLARQUAL',
                                     RechDom ('PGDECLARANTAUTRE', Declarant, False))
             else
                TobSalarie.PutValue ('PAS_DECLARQUAL',
                                     RechDom ('PGQUALDECLARANT2', DeclarantQual, False));
             end;
          End;
       Ferme (Q);
       TobSalarie.DelChampSup ('PSA_SALARIE', False);
       TobSalarie.DelChampSup ('PSA_LIBELLE', False);
       TobSalarie.DelChampSup ('PSA_PRENOM', False);
       TobSalarie.DelChampSup ('PSA_ETABLISSEMENT', False);
       TobSalarie.DelChampSup ('PSA_NUMEROSS', False);
       TobSalarie.DelChampSup ('PSA_LIBELLEEMPLOI', False);
       TobSalarie.DelChampSup ('PSA_STATUT', False);
       TobSalarie.DelChampSup ('PSA_DATEENTREE', False);
       TobSalarie.DelChampSup ('PSA_DATESORTIE', False);
       TobSalarie.DelChampSup ('PSA_HORHEBDO', False);
       TobSalarie.DelChampSup ('PSA_HORANNUEL', False);
       TobSalarie.DelChampSup ('PSA_MOTIFSORTIE', False);
       TobSalarie.DelChampSup ('PSA_FONCTIONSAL', False);  //PT19
       TobSalarie.DelChampSup ('PSA_CATDADS', False);
       TobSalarie.VirtuelleToReelle ('ATTESTATIONS');
       TobSalarie.SetAllModifie (True);
       TobSalarie.InsertDB (nil);
       ExecuteSQL ('UPDATE SALARIES SET PSA_ASSEDIC="X" WHERE'+
                   ' PSA_SALARIE="'+TobSalarie.GetString ('PAS_SALARIE')+'"');
       COMMITTRANS;
       ListeSalForEdition:= ListeSalForEdition+' OR'+
                            ' ((PAS_SALARIE="'+TobSalarie.GetString ('PAS_SALARIE')+'") AND'+
                            ' (PAS_ORDRE="'+TobSalarie.GetString ('PAS_ORDRE')+'"))';
    except
          ROLLBACK;
          HadFailed:= True;
          end;

    end; //fin de la boucle sur chaque salarié

// proposer l'edition des attestations calculées
if HadFailed then
   PGIError (TraduireMemoire ('Le calcul automatique des attestations ASSEDIC'+
                              ' a échoué.'), Ecran.Caption)
else
if (PGIAsk (TraduireMemoire ('Voulez-vous éditer les attestations qui viennent'+
                            ' d''être calculées ?'), Ecran.Caption)=mrYes) then
   begin
   if (Length(ListeSalForEdition)>0) then
      begin
      ListeSalForEdition:= RightStr (ListeSalForEdition,
                                     Length (ListeSalForEdition)-4);
      EditionAttestASSEDIC (ListeSalForEdition);
      end;
   end;
FreeAndNil (TobSalaries);
end;

procedure TOF_PGMULASSEDIC.CalculAttest(Sender: TObject);
var
  Faite, Ordre, Salarie, StSal : string;
  i: Integer;
  QRechSal: TQuery;
  Plusieurs : boolean;
begin
  if ((Q_Mul <> nil) and (Q_Mul.RecordCount = 0)) then
     exit;

  Plusieurs:= False; //PT13
  {$IFDEF EAGLCLIENT}
  TFMul(Ecran).Q.TQ.Seek(TFMul(Ecran).FListe.Row - 1);
  {$ENDIF}
  if ((GetControlText('SITUATION') = '1') or
     (GetControlText('SITUATION') = '4')) then
     begin
     Faite:= Q_Mul.FindField('PSA_ASSEDIC').AsString;
  //PT13
     if (Sender=TFMul(Ecran).BInsert) then
        Faite:= '-';
  //FIN PT13
     if Faite = 'X' then
        begin
        Salarie:= Q_Mul.FindField('PSA_SALARIE').AsString;
  //PT13
        StSal:= 'SELECT COUNT(PAS_SALARIE) AS NBRE'+
                ' FROM ATTESTATIONS WHERE'+
                ' PAS_SALARIE = "'+Salarie+'" AND'+
                ' PAS_TYPEATTEST = "ASS"';
        QRechSal:= OpenSQL (StSal, TRUE);
        Plusieurs:= QRechSal.FindField('NBRE').AsInteger>1;
        Ferme (QRechSal);
  //FIN PT13
        QRechSal:= OpenSQL ('SELECT PAS_ORDRE FROM ATTESTATIONS WHERE'+
                            ' PAS_SALARIE = "'+Salarie+'" AND'+
                            ' PAS_TYPEATTEST = "ASS"'+
                            ' ORDER BY PAS_DERNIERJOUR DESC', TRUE);
        Ordre:= IntToStr(QRechSal.FindField('PAS_ORDRE').AsInteger);
        Ferme(QRechSal);
        end
     else
        Salarie:= Q_Mul.FindField('PSA_SALARIE').AsString;
     end;

  if GetControlText('SITUATION') = '2' then
     Salarie:= Q_Mul.FindField('PSA_SALARIE').AsString;

  if GetControlText('SITUATION') = '3' then
     begin
     Salarie:= Q_Mul.FindField('PAS_SALARIE').AsString;
     Ordre:= IntToStr(Q_Mul.FindField('PAS_ORDRE').AsInteger);
     end;

  PGAttesSalarie:= Salarie;
  i:= StrToInt(GetControlText('SITUATION'));
  //PT15 {$IFNDEF EAGLCLIENT}
  case i of
       // Tous les salariés sortis
       1: if Faite = 'X' then
  {PT13
             AGLLanceFiche ('PAY', 'ASSEDIC', '', Salarie+';ASS;'+Ordre,
                            'ACTION=MODIFICATION')
  }
             begin
             if (Plusieurs=True) then
                AGLLanceFiche ('PAY', 'MUL_ATTESTASSED', '', '', 'S')
             else
                AGLLanceFiche ('PAY', 'ASSEDIC', '', Salarie+';ASS;'+Ordre,
                               'ACTION=MODIFICATION');
             end
  //FIN PT13
          else
             begin
             AGLLanceFiche ('PAY', 'ASSEDIC', '', '', Salarie+';ASS;'+Ordre+
                            ';ACTION=CREATION');
             if BCherche <> nil then
                BCherche.click;
             end;

       //Cas ou attest non faite
       2: begin
          AGLLanceFiche ('PAY', 'ASSEDIC', '', '', Salarie+';ASS;'+Ordre+
                         ';ACTION=CREATION');
          if BCherche <> nil then
             BCherche.click;
          end;

       //Cas ou attest déjà faite
  {PT13
       3: AGLLanceFiche ('PAY', 'ASSEDIC', '', Salarie+';ASS;'+Ordre,
                         'ACTION=MODIFICATION');
  }
       3: begin
          if (Plusieurs=True) then
             AGLLanceFiche ('PAY', 'MUL_ATTESTASSED', '', '', 'S')
          else
             AGLLanceFiche ('PAY', 'ASSEDIC', '', Salarie+';ASS;'+Ordre,
                            'ACTION=MODIFICATION');
          end;
  //FIN PT13

       // Tous les salariés
       4: if Faite = 'X' then
  {PT13
             AGLLanceFiche ('PAY', 'ASSEDIC', '', Salarie+';ASS;'+Ordre,
                            'ACTION=MODIFICATION')
  }
             begin
             if (Plusieurs=True) then
                AGLLanceFiche ('PAY', 'MUL_ATTESTASSED', '', '', 'S')
             else
                AGLLanceFiche ('PAY', 'ASSEDIC', '', Salarie+';ASS;'+Ordre,
                               'ACTION=MODIFICATION');
             end
  //FIN PT13
          else
             begin
             AGLLanceFiche ('PAY', 'ASSEDIC', '', '', Salarie+';ASS;'+Ordre+
                            ';ACTION=CREATION');
             if BCherche <> nil then
                BCherche.click;
             end;
       end;
end;

procedure TOF_PGMULASSEDIC.CreateIndemnites(TobSalarie : TOB; Ordre : Integer);
var
  QRechConv, QRechRub, QrechRCP, QRechRPreavis : TQuery;
  StConv, StRub, StRCP, StRPreavis, UsDateFin : String;
  DerJour : TDateTime;
  UsDerJour : String;
  CP, Preavis, Leg, CDD, FinMiss, Retraite, Spec, Journal, Conv : Extended;
  Client, Avion, Autre, CNE, Sinistre, Spcfiq, Apprenti, Transac: Extended;

begin
//  CP      := 0;                Retraite:= 0;                Spcfiq  := 0;
//  Preavis := 0;                Spec    := 0;                Apprenti:= 0;
//  Leg     := 0;                Journal := 0;                Autre   := 0;
//  CDD     := 0;                Client  := 0;                CNE     := 0;
//  FinMiss := 0;                Avion   := 0;                Sinistre:= 0;
//  Transac := 0;                Conv    := 0;
  DerJour := TobSalarie.GetDateTime('PAS_DERNIERJOUR');
  UsDerJour := UsDateTime(DerJour);
  if (TobSalarie.GetDateTime('PAS_PERIODEDEBUT')<DebutDeMois(DerJour)) then
     UsDateFin := UsDateTime(DebutDeMois(DerJour))
  else
     UsDateFin := UsDateTime(TobSalarie.GetDateTime('PAS_PERIODEDEBUT'));

  //Initialisation des champs
  TobSalarie.AddChampSupValeur('PAS_INDLEGB','-');
  TobSalarie.AddChampSupValeur ('PAS_INDLEGM',0);
  TobSalarie.AddChampSupValeur('PAS_INDCDDB','-');
  TobSalarie.AddChampSupValeur ('PAS_INDCDDM',0);
  TobSalarie.AddChampSupValeur('PAS_INDCNEB','-');
  TobSalarie.AddChampSupValeur ('PAS_INDCNEM', 0);
  TobSalarie.AddChampSupValeur('PAS_INDAUTREB','-');
  TobSalarie.AddChampSupValeur('PAS_INDAUTREM',0);
  TobSalarie.AddChampSupValeur('PAS_INDFINMISSB','-');
  TobSalarie.AddChampSupValeur ('PAS_INDFINMISSM',0);
  TobSalarie.AddChampSupValeur('PAS_INDRETRAITEB','-');
  TobSalarie.AddChampSupValeur ('PAS_INDRETRAITEM',0);
  TobSalarie.AddChampSupValeur ('PAS_INDSINISTREB','-');
  TobSalarie.AddChampSupValeur ('PAS_INDSINISTREM', 0);
  TobSalarie.AddChampSupValeur('PAS_INDSPECB','-');
  TobSalarie.AddChampSupValeur ('PAS_INDSPECM',0);
  TobSalarie.AddChampSupValeur ('PAS_INDSPCFIQB','-');
  TobSalarie.AddChampSupValeur ('PAS_INDSPCFIQM',0);
  TobSalarie.AddChampSupValeur('PAS_INDJOURNALB','-');
  TobSalarie.AddChampSupValeur ('PAS_INDJOURNALM',0);
  TobSalarie.AddChampSupValeur('PAS_INDCLIENTB','-');
  TobSalarie.AddChampSupValeur ('PAS_INDCLIENTM',0);
  TobSalarie.AddChampSupValeur('PAS_INDAVIONB','-');
  TobSalarie.AddChampSupValeur ('PAS_INDAVIONM',0);
  TobSalarie.AddChampSupValeur('PAS_INDAPPRENTIB','-');
  TobSalarie.AddChampSupValeur ('PAS_INDAPPRENTIM',0);
  TobSalarie.AddChampSupValeur ('PAS_INDTRANS2', 0);

  StRub:='SELECT PRM_INDEMLEGALES, SUM(PHB_MTREM) AS MTINDEMLEGALES'+
         ' FROM HISTOBULLETIN'+
         ' LEFT JOIN REMUNERATION ON'+
         ' PHB_NATURERUB=PRM_NATURERUB AND'+
         ' ##PRM_PREDEFINI## PHB_RUBRIQUE=PRM_RUBRIQUE WHERE'+
         ' PRM_INDEMLEGALES<>"" AND'+
         ' PHB_SALARIE="'+TobSalarie.GetString('PSA_SALARIE')+'" AND'+
         ' PHB_DATEFIN<="'+UsDerJour+'" AND'+
         ' PHB_DATEFIN>="'+UsDateFin+'" AND'+
         ' PHB_NATURERUB="AAA"'+
         ' GROUP BY PRM_INDEMLEGALES';
  QRechRub:=OpenSql(StRub,TRUE);
  if ((not QRechRub.eof) and (QRechRub.FindField('PRM_INDEMLEGALES').Asstring = '001')) then
  begin
    if QRechRub.FindField('MTINDEMLEGALES').AsFloat <> 0 then
    begin
      TobSalarie.PutValue('PAS_INDLEGB','X');
      Leg := QRechRub.FindField('MTINDEMLEGALES').AsFloat;
      TobSalarie.PutValue ('PAS_INDLEGM',Leg);
    end;
    QRechRub.Next;
  end;
  if ((not QRechRub.eof) and (QRechRub.FindField('PRM_INDEMLEGALES').Asstring = '003')) then
  begin
    if QRechRub.FindField('MTINDEMLEGALES').AsFloat <> 0 then
    begin
      TobSalarie.PutValue('PAS_INDCDDB','X');
      CDD := QRechRub.FindField('MTINDEMLEGALES').AsFloat;
      TobSalarie.PutValue ('PAS_INDCDDM',CDD);
    end;
    QRechRub.Next;
  end;
  if ((not QRechRub.eof) and (QRechRub.FindField('PRM_INDEMLEGALES').Asstring = '011')) then
  begin
    if QRechRub.FindField('MTINDEMLEGALES').AsFloat <> 0 then
    begin
      TobSalarie.PutValue('PAS_INDCNEB','X');
      CNE:= QRechRub.FindField('MTINDEMLEGALES').AsFloat;
      TobSalarie.PutValue ('PAS_INDCNEM', CNE);
    end;
    QRechRub.Next;
  end;
  if ((not QRechRub.eof) and (QRechRub.FindField('PRM_INDEMLEGALES').Asstring = '004')) then
  begin
    if QRechRub.FindField('MTINDEMLEGALES').AsFloat <> 0 then
    begin
      TobSalarie.PutValue('PAS_INDAUTREB','X');
      Autre := QRechRub.FindField('MTINDEMLEGALES').AsFloat;
      TobSalarie.PutValue('PAS_INDAUTREM',Autre);
    end;
    QRechRub.Next;
  end;
  if ((not QRechRub.eof) and (QRechRub.FindField('PRM_INDEMLEGALES').Asstring = '005')) then
  begin
    if QRechRub.FindField('MTINDEMLEGALES').AsFloat <> 0 then
    begin
      TobSalarie.PutValue('PAS_INDFINMISSB','X');
      FinMiss := QRechRub.FindField('MTINDEMLEGALES').AsFloat;
      TobSalarie.PutValue ('PAS_INDFINMISSM',FinMiss);
    end;
    QRechRub.Next;
  end;
  if ((not QRechRub.eof) and (QRechRub.FindField('PRM_INDEMLEGALES').Asstring = '006')) then
  begin
    if QRechRub.FindField('MTINDEMLEGALES').AsFloat <> 0 then
    begin
      TobSalarie.PutValue('PAS_INDRETRAITEB','X');
      Retraite := QRechRub.FindField('MTINDEMLEGALES').AsFloat;
      TobSalarie.PutValue ('PAS_INDRETRAITEM',Retraite);
    end;
    QRechRub.Next;
  end;
  if ((not QRechRub.eof) and (QRechRub.FindField('PRM_INDEMLEGALES').Asstring = '012')) then
  begin
    if QRechRub.FindField('MTINDEMLEGALES').AsFloat <> 0 then
    begin
      TobSalarie.PutValue ('PAS_INDSINISTREB','X');
      Sinistre:= QRechRub.FindField('MTINDEMLEGALES').AsFloat;
      TobSalarie.PutValue ('PAS_INDSINISTREM', Sinistre);
    end;
    QRechRub.Next;
  end;
  if ((not QRechRub.eof) and  (QRechRub.FindField('PRM_INDEMLEGALES').Asstring = '007')) then
  begin
    if QRechRub.FindField('MTINDEMLEGALES').AsFloat <> 0 then
    begin
      TobSalarie.PutValue('PAS_INDSPECB','X');
      Spec := QRechRub.FindField('MTINDEMLEGALES').AsFloat;
      TobSalarie.PutValue ('PAS_INDSPECM',Spec);
    end;
    QRechRub.Next;
  end;
  if ((not QRechRub.eof) and (QRechRub.FindField('PRM_INDEMLEGALES').Asstring = '013')) then
  begin
    if QRechRub.FindField('MTINDEMLEGALES').AsFloat <> 0 then
    begin
      TobSalarie.PutValue ('PAS_INDSPCFIQB','X');
      Spcfiq:= QRechRub.FindField('MTINDEMLEGALES').AsFloat;
      TobSalarie.PutValue ('PAS_INDSPCFIQM',Spcfiq);
    end;
    QRechRub.Next;
  end;
  if ((not QRechRub.eof) and (QRechRub.FindField('PRM_INDEMLEGALES').Asstring = '008')) then
  begin
    if QRechRub.FindField('MTINDEMLEGALES').AsFloat <> 0 then
    begin
      TobSalarie.PutValue('PAS_INDJOURNALB','X');
      Journal := QRechRub.FindField('MTINDEMLEGALES').AsFloat;
      TobSalarie.PutValue ('PAS_INDJOURNALM',Journal);
    end;
    QRechRub.Next;
  end;
  if ((not QRechRub.eof) and (QRechRub.FindField('PRM_INDEMLEGALES').Asstring = '009')) then
  begin
    if QRechRub.FindField('MTINDEMLEGALES').AsFloat <> 0 then
    begin
      TobSalarie.PutValue('PAS_INDCLIENTB','X');
      Client := QRechRub.FindField('MTINDEMLEGALES').AsFloat;
      TobSalarie.PutValue ('PAS_INDCLIENTM',Client);
    end;
    QRechRub.Next;
  end;
  if ((not QRechRub.eof) and (QRechRub.FindField('PRM_INDEMLEGALES').Asstring = '010')) then
  begin
    if QRechRub.FindField('MTINDEMLEGALES').AsFloat <> 0 then
    begin
      TobSalarie.PutValue('PAS_INDAVIONB','X');
      Avion := QRechRub.FindField('MTINDEMLEGALES').AsFloat;
      TobSalarie.PutValue ('PAS_INDAVIONM',Avion);
    end;
    QRechRub.Next;
  end;
  if ((not QRechRub.eof) and (QRechRub.FindField('PRM_INDEMLEGALES').Asstring = '014')) then
  begin
    if QRechRub.FindField('MTINDEMLEGALES').AsFloat <> 0 then
    begin
      TobSalarie.PutValue('PAS_INDAPPRENTIB','X');
      Apprenti := QRechRub.FindField('MTINDEMLEGALES').AsFloat;
      TobSalarie.PutValue ('PAS_INDAPPRENTIM',Apprenti);
    end;
    QRechRub.Next;
  end;
  if ((not QRechRub.eof) and (QRechRub.FindField('PRM_INDEMLEGALES').Asstring = '015')) then
  begin
    if (QRechRub.FindField('MTINDEMLEGALES').AsFloat <> 0) then
    begin
      Transac:= QRechRub.FindField ('MTINDEMLEGALES').AsFloat;
      TobSalarie.PutValue ('PAS_INDTRANS2', Transac);
    End;
    QRechRub.Next;
  end;
  Ferme(QRechRub);

  StRCP:='SELECT SUM(PHB_MTREM) AS MTCP, SUM(PHB_BASEREM) AS BASECP'+
         ' FROM HISTOBULLETIN'+
         ' LEFT JOIN REMUNERATION ON'+
         ' PHB_NATURERUB=PRM_NATURERUB AND'+
         ' ##PRM_PREDEFINI## PHB_RUBRIQUE=PRM_RUBRIQUE WHERE'+
         ' PRM_INDEMCOMPCP="X" AND'+
         ' PHB_SALARIE="'+TobSalarie.GetString('PSA_SALARIE')+'" AND'+
         ' PHB_DATEFIN<="'+UsDerJour+'" AND'+
         ' PHB_DATEFIN>="'+UsDateFin+'" AND'+
         ' PHB_NATURERUB="AAA"';
  QRechRCP:=OpenSql(StRCP,TRUE);
  if (not QRechRCP.eof) then
     begin
     if QRechRCP.FindField('MTCP').AsFloat <> 0 then
        begin
        CP := QRechRCP.FindField('MTCP').AsFloat;
        TobSalarie.AddChampSupValeur('PAS_INDCPMONTANT',CP);
        end;
     if QRechRCP.FindField('BASECP').AsFloat <> 0 then
        begin
        CP := QRechRCP.FindField('BASECP').AsFloat;
        TobSalarie.AddChampSupValeur('PAS_INDCPJOURS',CP);
        end;
     end;
  Ferme(QRechRCP);

  StRPreavis:='SELECT SUM(PHB_MTREM) AS MTPREAVIS'+
              ' FROM HISTOBULLETIN'+
              ' LEFT JOIN REMUNERATION ON'+
              ' PHB_NATURERUB=PRM_NATURERUB AND'+
              ' ##PRM_PREDEFINI## PHB_RUBRIQUE=PRM_RUBRIQUE WHERE'+
              ' PRM_INDEMPREAVIS="X" AND'+
              ' PHB_SALARIE="'+TobSalarie.GetString('PSA_SALARIE')+'" AND'+
              ' PHB_DATEFIN<="'+UsDerJour+'" AND'+
              ' PHB_DATEFIN>="'+UsDateFin+'" AND'+
              ' PHB_NATURERUB="AAA"';
  QRechRPreavis:=OpenSql(StRPreavis,TRUE);
  if ((not QRechRPreavis.eof) and
     (QRechRPreavis.FindField('MTPREAVIS').AsFloat <> 0)) then
     begin
     Preavis := QRechRPreavis.FindField('MTPREAVIS').AsFloat;
     TobSalarie.AddChampSupValeur ('PAS_INDPREAVIS',Preavis);
     end;
  Ferme(QRechRPreavis);

  StConv:='SELECT SUM(PHB_MTREM) AS MTCONV'+
          ' FROM HISTOBULLETIN'+
          ' LEFT JOIN REMUNERATION ON'+
          ' PHB_NATURERUB=PRM_NATURERUB AND'+
          ' ##PRM_PREDEFINI## PHB_RUBRIQUE=PRM_RUBRIQUE WHERE'+
          ' PRM_INDEMCONVENT="X" AND'+
          ' PHB_SALARIE="'+TobSalarie.GetString('PSA_SALARIE')+'" AND'+
          ' PHB_DATEFIN<="'+UsDerJour+'" AND'+
          ' PHB_DATEFIN>="'+UsDateFin+'" AND'+
          ' PHB_NATURERUB="AAA"';
  QRechConv:=OpenSql(StConv,TRUE);

  if ((not QRechConv.eof) and (QRechConv.FindField('MTCONV').AsFloat <> 0)) then
     begin
     Conv := QRechConv.FindField('MTCONV').AsFloat;
     TobSalarie.AddChampSupValeur('PAS_INDCONV2', Conv);
     end;
  Ferme(QRechConv);
end;

procedure TOF_PGMULASSEDIC.CreateSalaires(TobSalarie : TOB; Ordre : Integer);
var
  StPaye, StPayeCumSal, StPayele, UsDateDeb, UsDateFin : String;
  QRechPaye, QRechPayeCumSal, QRechPayele : TQuery;
  Date1, Date2, DateD, DateF : TDateTime;
  j : integer;
  Payele : array[1..12] of TDateTime;
  Heures, Precompte, SommeBrut : array[1..12] of Double;
  PremAnnee, PremMois : Word;
  TAssietteBrut, TAssietteBrutD, TCP, TCPD, THistoCumSal, THistoCumSalD : TOB;
  DerJour : TDateTime;
  UsDerJour : String;
  NbreMois : Word;
begin
  For j:=1 to 12 do
  begin
    Payele[j]:=IDate1900;
    Heures[j]:=0;
    Precompte[j]:=0;
    SommeBrut[j]:=0;
  end;
  DerJour := TobSalarie.GetDateTime('PAS_DERNIERJOUR');
  PremAnnee := 0;
  PremMois := 0;
  NbreMois := 12;
  Date1 := TobSalarie.GetDateTime('PAS_PERIODEDEBUT');
  if (DerJour<FinDeMois(DerJour)) then
    Date2:=FinDeMois(PlusMois(DerJour, -1))
  else
    Date2:=DerJour;
  NOMBREMOIS (DebutDeMois(Date1), FinDeMois(Date2), PremMois, PremAnnee, NbreMois);
  if NbreMois > 12 then
    NbreMois := 12;

  UsDerJour := UsDateTime(DerJour);
  UsDateDeb := UsDateTime (DebutDeMois(PlusMois (DerJour, -NbreMois)));
  if (DerJour<FinDeMois(DerJour)) then
    UsDateFin := UsDateTime (DebutDeMois(DerJour))
  else
    UsDateFin := UsDateTime (DebutDeMois(PlusMois(DerJour, 1)));
  StPayele := 'SELECT PPU_PAYELE, PPU_DATEDEBUT, PPU_DATEFIN'+
             ' FROM PAIEENCOURS WHERE'+
             ' PPU_SALARIE="'+TobSalarie.GetString('PSA_SALARIE')+'" AND'+
             ' PPU_DATEDEBUT >= "'+UsDateDeb+'" AND'+
             ' PPU_DATEDEBUT <= "'+UsDateFin+'" ORDER BY PPU_DATEDEBUT';
  QRechPayele := OpenSql(StPayele, TRUE);
  QRechPayele.First;
  While Not QRechPayele.eof do
  begin
    DateF := QRechPayele.FindField ('PPU_DATEFIN').AsDateTime;
    For j:=1 to NbreMois do
    begin
      if (DateF >= DebutDeMois(PlusMois (Date2, j-(NbreMois))))
        and (DateF <= FinDeMois(PlusMois (Date2, j-(NbreMois))))
        and (QRechPayele.FindField ('PPU_PAYELE').AsDateTime>=Payele[j]) then
      begin
        Payele[j]:= QRechPayele.FindField ('PPU_PAYELE').AsDateTime;
        break;
      end;
    end;
    QRechPayele.next;
  end;
  Ferme(QRechPayele);

  StPayeCumSal:= 'SELECT PHC_MONTANT, PHC_OMONTANT, PHC_DATEDEBUT,'+
                 ' PHC_DATEFIN, PHC_CUMULPAIE'+
                 ' FROM HISTOCUMSAL WHERE'+
                 ' PHC_SALARIE="'+TobSalarie.GetString('PSA_SALARIE')+'" AND'+
                 ' (PHC_CUMULPAIE="01" OR PHC_CUMULPAIE="20") AND'+
                 ' PHC_DATEDEBUT>="'+UsDateDeb+'" AND'+
                 ' PHC_DATEFIN<="'+UsDateFin+'" ORDER BY PHC_DATEDEBUT';
  QRechPayeCumSal:= OpenSql(StPayeCumSal, TRUE);
  THistoCumSal:= TOB.Create ('Les HistoCumSal', NIL, -1);
  THistoCumSal.LoadDetailDB ('Les HistoCumSal','','',QRechPayeCumSal,False);
  Ferme(QRechPayeCumSal);
  THistoCumSalD:= THistoCumSal.FindFirst(['PHC_CUMULPAIE'], ['20'], FALSE);
  while (THistoCumSalD<>nil) do
  begin
    DateF:= THistoCumSalD.GetValue('PHC_DATEFIN');
    For j:=1 to NbreMois do
    begin
      if (DateF >= DebutDeMois(PlusMois (Date2, j-(NbreMois)))) and
        (DateF <= FinDeMois(PlusMois (Date2, j-(NbreMois)))) then
      begin
        Heures[j]:= Heures[j]+THistoCumSalD.GetValue('PHC_MONTANT');
        break;
      end;
    end;
    THistoCumSalD:= THistoCumSal.FindNext(['PHC_CUMULPAIE'], ['20'], FALSE);
  end;
  StPaye:= 'SELECT PHB_BASECOT, PHB_OBASECOT, PHB_MTSALARIAL,'+
           ' PHB_OMTSALARIAL, PHB_DATEDEBUT, PHB_DATEFIN, PCT_ASSIETTEBRUT,'+
           ' PCT_PRECOMPTEASS'+
           ' FROM HISTOBULLETIN'+
           ' LEFT JOIN COTISATION ON'+
           ' PHB_RUBRIQUE=PCT_RUBRIQUE WHERE'+
           ' ##PCT_PREDEFINI## (PCT_ASSIETTEBRUT="X" OR'+
           ' PCT_PRECOMPTEASS="X") AND'+
           ' PHB_SALARIE="'+TobSalarie.GetString('PSA_SALARIE')+'" AND'+
           ' PHB_DATEDEBUT>="'+UsDateDeb+'" AND'+
           ' PHB_DATEFIN<="'+UsDateFin+'" AND'+
           ' PHB_NATURERUB <> "AAA" ORDER BY PHB_DATEDEBUT';
  QRechPaye:= OpenSql(StPaye, TRUE);
  TAssietteBrut:= TOB.Create('Les AssietteBrut', NIL, -1);
  TAssietteBrut.LoadDetailDB('Les AssietteBrut','','',QRechPaye,False);
  Ferme(QRechPaye);
  StPaye:= 'SELECT PHB_DATEDEBUT, PHB_DATEFIN, PHB_MTREM, PRM_INDEMPREAVIS,'+
           ' PRM_INDEMCOMPCP'+
           ' FROM HISTOBULLETIN'+
           ' LEFT JOIN REMUNERATION ON'+
           ' PHB_NATURERUB=PRM_NATURERUB AND'+
           ' ##PRM_PREDEFINI## PHB_RUBRIQUE=PRM_RUBRIQUE WHERE'+
           ' (PRM_INDEMPREAVIS="X" OR PRM_INDEMCOMPCP="X") AND'+
           ' PHB_SALARIE="'+TobSalarie.GetString('PSA_SALARIE')+'" AND'+
           ' PHB_DATEDEBUT>="'+UsDateDeb+'" AND'+
           ' PHB_DATEFIN<="'+UsDateFin+'" AND'+
           ' PHB_NATURERUB="AAA"';
  QRechPaye:= OpenSql(StPaye, TRUE);
  TCP:= TOB.Create('Les CP', NIL, -1);
  TCP.LoadDetailDB('Les CP','','',QRechPaye,False);
  Ferme(QRechPaye);
  THistoCumSalD:= THistoCumSal.FindFirst(['PHC_CUMULPAIE'], ['01'], FALSE);
  while (THistoCumSalD<>nil) do
  begin
    DateD:= THistoCumSalD.GetValue('PHC_DATEDEBUT');
    DateF:= THistoCumSalD.GetValue('PHC_DATEFIN');
    TAssietteBrutD:= TAssietteBrut.FindFirst(['PHB_DATEDEBUT', 'PHB_DATEFIN', 'PCT_ASSIETTEBRUT'],
                                              [DateD, DateF, 'X'], TRUE);
    if (TobSalarie.GetString('PSA_CATDADS') = '003') then
       TAssietteBrutD:= nil;
    if (TAssietteBrutD<>nil) then //On n'est pas dans le cas de la reprise
    begin
      For j:=1 to NbreMois do
      begin
        if (DateF>=DebutDeMois(PlusMois (Date2, j-(NbreMois)))) and
           (DateF<=FinDeMois(PlusMois (Date2, j-(NbreMois)))) then
        begin
          if (VH_Paie.PGTenueEuro=FALSE) then
            SommeBrut[j]:= SommeBrut[j]+ TAssietteBrutD.GetValue('PHB_BASECOT')
          else begin
            if (VH_Paie.PGDateBasculEuro<=DebutDeMois(PlusMois (Date2, j-(NbreMois)))) then
              SommeBrut[j]:= SommeBrut[j]+ TAssietteBrutD.GetValue('PHB_BASECOT')
            else
              SommeBrut[j]:= SommeBrut[j]+ TAssietteBrutD.GetValue('PHB_OBASECOT');
          end;
          TCPD:= TCP.FindFirst (['PHB_DATEDEBUT', 'PHB_DATEFIN', 'PRM_INDEMPREAVIS'], [DateD, DateF, 'X'], TRUE);
          if (TCPD<>nil) then
            SommeBrut[j]:= SommeBrut[j]-TCPD.GetValue('PHB_MTREM');
          TCPD:= TCP.FindFirst (['PHB_DATEDEBUT', 'PHB_DATEFIN', 'PRM_INDEMCOMPCP'], [DateD, DateF, 'X'], TRUE);
          if (TCPD<>nil) then
            SommeBrut[j]:= SommeBrut[j]-TCPD.GetValue('PHB_MTREM');
        end;
      end;
    end else begin
      For j:=1 to NbreMois do
      begin
        if (DateF>=DebutDeMois(PlusMois (Date2, j-(NbreMois)))) and
           (DateF<=FinDeMois(PlusMois (Date2, j-(NbreMois)))) then
        begin
          if (VH_Paie.PGTenueEuro=FALSE) then
            SommeBrut[j]:= SommeBrut[j]+ THistoCumSalD.GetValue('PHC_MONTANT')
          else begin
            if (VH_Paie.PGDateBasculEuro<=DebutDeMois(PlusMois (Date2, j-(NbreMois)))) then
              SommeBrut[j]:= SommeBrut[j]+ THistoCumSalD.GetValue('PHC_MONTANT')
            else
              SommeBrut[j]:= SommeBrut[j]+ THistoCumSalD.GetValue('PHC_OMONTANT');
          end;
          TCPD:= TCP.FindFirst (['PHB_DATEDEBUT', 'PHB_DATEFIN', 'PRM_INDEMPREAVIS'], [DateD, DateF, 'X'], TRUE);
          if (TCPD<>nil) then
            SommeBrut[j]:= SommeBrut[j]-TCPD.GetValue('PHB_MTREM');
          TCPD:= TCP.FindFirst (['PHB_DATEDEBUT', 'PHB_DATEFIN', 'PRM_INDEMCOMPCP'], [DateD, DateF, 'X'], TRUE);
          if (TCPD<>nil) then
            SommeBrut[j]:= SommeBrut[j]-TCPD.GetValue('PHB_MTREM');
        end;
      end;
    end;
    THistoCumSalD:= THistoCumSal.FindNext(['PHC_CUMULPAIE'], ['01'], TRUE);
  end;
  FreeAndNil(THistoCumSal);
  FreeAndNil(TCP);
  TAssietteBrutD:= TAssietteBrut.FindFirst(['PCT_PRECOMPTEASS'], ['X'], FALSE);
  While (TAssietteBrutD<>nil) do
  begin
    DateF := TAssietteBrutD.GetValue('PHB_DATEFIN');
    For j:=1 to NbreMois do
    begin
      if (DateF >= DebutDeMois(PlusMois (Date2, j-(NbreMois)))) and
         (DateF <= FinDeMois(PlusMois (Date2, j-(NbreMois)))) then
        if (VH_Paie.PGTenueEuro = FALSE) then
        begin
           Precompte[j]:= Precompte[j]+TAssietteBrutD.GetValue ('PHB_MTSALARIAL');
           break;
        end else BEGIN
          if (VH_Paie.PGDateBasculEuro <= DebutDeMois(PlusMois (Date2, j-(NbreMois)))) then
          begin
            Precompte[j]:= Precompte[j]+TAssietteBrutD.GetValue ('PHB_MTSALARIAL');
            break;
          end else begin
            Precompte[j]:= Precompte[j]+TAssietteBrutD.GetValue ('PHB_OMTSALARIAL');
            break;
          end;
        END;
    end;
    TAssietteBrutD:= TAssietteBrut.FindNext(['PCT_PRECOMPTEASS'], ['X'], FALSE);
  end;
  FreeAndNil(TAssietteBrut);
  For j:=1 to NbreMois do
  begin
    if not (Payele[j] > 0) then
      Payele[j]:= IDate1900;
    ExecuteSQL( 'INSERT INTO ATTSALAIRES( '
              + ' PAL_DATEDEBUT, PAL_DATEFIN,'
              + ' PAL_PAYELE, PAL_NBHEURES,'
              + ' PAL_JNONPAYES, PAL_SALAIRE, PAL_PRECOMPTE, '
              + ' PAL_OBSERVATIONS, PAL_MOIS, PAL_ORDRE, PAL_SALARIE, PAL_TYPEATTEST'
              + ' ) VALUES ( '
              + '"'+USDATETIME(DEBUTDEMOIS(PLUSMOIS(Date2, j-(NbreMois))))+'", "'+USDATETIME(FINDEMOIS  (PLUSMOIS(Date2, j-(NbreMois))))+'", '
              + '"'+USDATETIME(Payele[j])+'", '+StrFPoint(Heures[j])+', '
              + '0, '+StrFPoint(SommeBrut[j])+', '+StrFPoint(Precompte[j])+', '
              + '"", '+IntToStr(j)+', '+IntToStr(ordre)+', "'+TobSalarie.GetString('PSA_SALARIE')+'", "ASS"'
              + ' )');
  end;
  For j:=NbreMois+1 to 12 do  //PT17 Il faut quand même insérer les lignes vides sinon, l'edition se décale
  begin
    if not (Payele[j] > 0) then
      Payele[j]:= IDate1900;
    ExecuteSQL( 'INSERT INTO ATTSALAIRES( '
              + ' PAL_DATEDEBUT, PAL_DATEFIN,'
              + ' PAL_PAYELE, PAL_NBHEURES,'
              + ' PAL_JNONPAYES, PAL_SALAIRE, PAL_PRECOMPTE, '
              + ' PAL_OBSERVATIONS, PAL_MOIS, PAL_ORDRE, PAL_SALARIE, PAL_TYPEATTEST'
              + ' ) VALUES ( '
              + '"'+USDATETIME(iDate1900)+'", "'+USDATETIME(iDate1900)+'", '
              + '"'+USDATETIME(Payele[j])+'", '+StrFPoint(Heures[j])+', '
              + '0, '+StrFPoint(SommeBrut[j])+', '+StrFPoint(Precompte[j])+', '
              + '"", '+IntToStr(j)+', '+IntToStr(ordre)+', "'+TobSalarie.GetString('PSA_SALARIE')+'", "ASS"'
              + ' )');
  end;
end;

procedure TOF_PGMULASSEDIC.CreatePrimes(TobSalarie : TOB; Ordre : Integer);
var
  j : integer;
  DateDebut, DateFin, PayeLe : array[1..3] of TDateTime;
  Heures, Precompte, SommeBrut : array[1..3] of Double;
begin
  //PT17, on insere des lignes vides
  For j:=1 to 3 do
  begin
    DateDebut[j]:= iDate1900;
    DateFin[j]  := iDate1900;
    Payele[j]   := iDate1900;
    Heures[j]   := 0;
    Precompte[j]:= 0;
    SommeBrut[j]:= 0;
  end;
  For j:=1 to 3 do
  begin
    if not(Payele[j]> 0) then
      Payele[j] := iDate1900;
    ExecuteSQL( 'INSERT INTO ATTSALAIRES( '
              + ' PAL_DATEDEBUT, PAL_DATEFIN,'
              + ' PAL_PAYELE, PAL_NBHEURES,'
              + ' PAL_JNONPAYES, PAL_SALAIRE, PAL_PRECOMPTE,'
              + ' PAL_OBSERVATIONS, PAL_MOIS, PAL_ORDRE, PAL_SALARIE, PAL_TYPEATTEST'
              + ' ) VALUES ( '
              + '"'+USDATETIME(DateDebut[j])+'", "'+USDATETIME(DateFin[j])+'", '
              + '"'+USDATETIME(Payele[j])+'", '+StrFPoint(Heures[j])+', '
              + '0, '+StrFPoint(SommeBrut[j])+', '+StrFPoint(Precompte[j])+', '
              + '"", '+IntToStr(12+j)+', '+IntToStr(ordre)+', "'+TobSalarie.GetString('PSA_SALARIE')+'", "ASS"'
              + ' )');
  end;
end;

procedure TOF_PGMULASSEDIC.CreateSolde(TobSalarie : TOB; Ordre : Integer);
var
  StPaye, StPayeCumSal, StPayele : String;
  UsDateDeb, UsDateFin : String;
  QRechPaye, QRechPayeCumSal, QRechPayele : TQuery;
  j : integer;
  Date1, Date2, DateD, DateF : TDateTime;
  DateDebut, DateFin, PayeLe : array[1..2] of TDateTime;
  Heures, Precompte, SommeBrut : array[1..2] of Double;
  TAssietteBrut, TAssietteBrutD, TCP, TCPD, THistoCumSal, THistoCumSalD : TOB;
  DerJour : TDateTime;
  UsDerJour : String;
begin
  For j:=1 to 2 do
  begin
    DateDebut[j]:= iDate1900;
    DateFin[j]  := iDate1900;
    Payele[j]   := iDate1900;
    Heures[j]   := 0;
    Precompte[j]:= 0;
    SommeBrut[j]:= 0;
  end;
  DerJour := TobSalarie.GetDateTime('PAS_DERNIERJOUR');
  if ( TobSalarie.GetDateTime('PAS_PERIODEFIN')>TobSalarie.GetDateTime('PAS_DERNIERJOUR') ) then
  begin
    Date2:= FinDeMois (TobSalarie.GetDateTime('PAS_PERIODEFIN'));
    UsDateFin:= UsDateTime (Date2);
  end else begin
    Date2:= FinDeMois (DerJour);
    UsDateFin:= UsDateTime (Date2);
  end;
  UsDerJour:= UsDateTime (DerJour);
  Date1:= DebutDeMois (DerJour);
  UsDateDeb:= UsDateTime (Date1);
  if (DerJour<Date2) then
  begin
    StPayele := 'SELECT PPU_PAYELE, PPU_DATEDEBUT, PPU_DATEFIN'+
                ' FROM PAIEENCOURS WHERE'+
                ' PPU_SALARIE="'+TobSalarie.GetString('PSA_SALARIE')+'" AND'+
                ' PPU_DATEDEBUT>="'+UsDateDeb+'" AND'+
                ' PPU_DATEDEBUT<="'+UsDateFin+'" ORDER BY PPU_DATEDEBUT';
    QRechPayele := OpenSql(StPayele, TRUE);
    QRechPayele.First;
    while NOT QRechPayele.EOF do
    begin
      DateD:= QRechPayele.FindField ('PPU_DATEDEBUT').AsDateTime;
      DateF:= QRechPayele.FindField ('PPU_DATEFIN').AsDateTime;
      For j:=1 to 2 do
      begin
        if (DateF >= DebutDeMois(PlusMois (Date2, j-2))) and
           (DateF <= FinDeMois(PlusMois (Date2, j-2))) and
           (QRechPayele.FindField ('PPU_PAYELE').AsDateTime>=Payele[j]) then
        begin
          DateDebut[j]:= DateD;
          DateFin[j]:= DateF;
          Payele[j]:= QRechPayele.FindField ('PPU_PAYELE').AsDateTime;
          break;
        end;
      end;
      QRechPayele.next;
    end;
    Ferme(QRechPayele);
    StPayeCumSal := 'SELECT PHC_MONTANT, PHC_OMONTANT, PHC_DATEDEBUT,'+
                    ' PHC_DATEFIN, PHC_CUMULPAIE'+
                    ' FROM HISTOCUMSAL WHERE'+
                    ' PHC_SALARIE="'+TobSalarie.GetString('PSA_SALARIE')+'" AND'+
                    ' (PHC_CUMULPAIE="01" OR PHC_CUMULPAIE="20") AND'+
                    ' PHC_DATEDEBUT>="'+UsDateDeb+'" AND'+
                    ' PHC_DATEFIN<="'+UsDateFin+'"';
    QRechPayeCumSal := OpenSql(StPayeCumSal, TRUE);
    THistoCumSal := TOB.Create('Les HistoCumSal', NIL, -1);
    THistoCumSal.LoadDetailDB('Les HistoCumSal','','',QRechPayeCumSal,False);
    Ferme(QRechPayeCumSal);
    THistoCumSalD:= THistoCumSal.FindFirst(['PHC_CUMULPAIE'], ['20'], FALSE);
    while (THistoCumSalD<>nil) do
    begin
      DateF:= THistoCumSalD.GetValue('PHC_DATEFIN');
      For j:=1 to 2 do
      begin
        if (DateF >= DebutDeMois(PlusMois (Date2, j-2))) and
           (DateF <= FinDeMois(PlusMois (Date2, j-2))) then
        begin
          Heures[j]:= Heures[j]+THistoCumSalD.GetValue('PHC_MONTANT');
          break;
        end;
      end;
      THistoCumSalD:= THistoCumSal.FindNext(['PHC_CUMULPAIE'], ['20'], FALSE);
    end;
    StPaye:= 'SELECT PHB_BASECOT, PHB_OBASECOT, PHB_MTSALARIAL,'+
             ' PHB_OMTSALARIAL, PHB_DATEDEBUT, PHB_DATEFIN,'+
             ' PCT_ASSIETTEBRUT, PCT_PRECOMPTEASS'+
             ' FROM HISTOBULLETIN'+
             ' LEFT JOIN COTISATION ON'+
             ' PHB_RUBRIQUE=PCT_RUBRIQUE WHERE'+
             ' ##PCT_PREDEFINI## (PCT_ASSIETTEBRUT="X" OR'+
             ' PCT_PRECOMPTEASS="X") AND'+
             ' PHB_SALARIE="'+TobSalarie.GetString('PSA_SALARIE')+'" AND'+
             ' PHB_DATEDEBUT>="'+UsDateDeb+'" AND'+
             ' PHB_DATEFIN<="'+UsDateFin+'" AND'+
             ' PHB_NATURERUB <> "AAA"';
    QRechPaye := OpenSql(StPaye, TRUE);
    TAssietteBrut:= TOB.Create('Les AssietteBrut', NIL, -1);
    TAssietteBrut.LoadDetailDB('Les AssietteBrut','','',QRechPaye,False);
    Ferme(QRechPaye);
    StPaye:= 'SELECT PHB_DATEDEBUT, PHB_DATEFIN, PHB_MTREM,'+
             ' PRM_INDEMPREAVIS, PRM_INDEMCOMPCP'+
             ' FROM HISTOBULLETIN'+
             ' LEFT JOIN REMUNERATION ON'+
             ' PHB_NATURERUB=PRM_NATURERUB AND'+
             ' ##PRM_PREDEFINI## PHB_RUBRIQUE=PRM_RUBRIQUE WHERE'+
             ' (PRM_INDEMPREAVIS="X" OR PRM_INDEMCOMPCP="X") AND'+
             ' PHB_SALARIE="'+TobSalarie.GetString('PSA_SALARIE')+'" AND'+
             ' PHB_DATEDEBUT>="'+UsDateDeb+'" AND'+
             ' PHB_DATEFIN<="'+UsDateFin+'" AND'+
             ' PHB_NATURERUB="AAA"';
    QRechPaye:= OpenSql(StPaye, TRUE);
    TCP:= TOB.Create('Les CP', NIL, -1);
    TCP.LoadDetailDB('Les CP','','',QRechPaye,False);
    Ferme(QRechPaye);
    THistoCumSalD:= THistoCumSal.FindFirst(['PHC_CUMULPAIE'], ['01'], FALSE);
    while (THistoCumSalD<>nil) do
    begin
      DateD:= THistoCumSalD.GetValue('PHC_DATEDEBUT');
      DateF:= THistoCumSalD.GetValue('PHC_DATEFIN');
      TAssietteBrutD:= TAssietteBrut.FindFirst(['PHB_DATEDEBUT', 'PHB_DATEFIN', 'PCT_ASSIETTEBRUT'],
                                               [DateD, DateF, 'X'], TRUE);
      if (TobSalarie.GetString('PSA_CATDADS') = '003') then
        TAssietteBrutD := nil;
      if (TAssietteBrutD<>nil) then //On n'est pas dans le cas de la reprise
      begin
        For j:=1 to 2 do
        begin
          if (DateF>=DebutDeMois(PlusMois (Date2, j-2))) and
             (DateF<=FinDeMois(PlusMois (Date2, j-2))) then
          begin
            if (VH_Paie.PGTenueEuro=FALSE) then
              SommeBrut[j]:= SommeBrut[j]+ TAssietteBrutD.GetValue('PHB_BASECOT')
            else begin
              if (VH_Paie.PGDateBasculEuro<=DebutDeMois(PlusMois (Date2, j-2))) then
                SommeBrut[j]:= SommeBrut[j]+ TAssietteBrutD.GetValue('PHB_BASECOT')
              else
                SommeBrut[j]:= SommeBrut[j]+ TAssietteBrutD.GetValue('PHB_OBASECOT');
            end;
            TCPD:= TCP.FindFirst (['PHB_DATEDEBUT', 'PHB_DATEFIN', 'PRM_INDEMPREAVIS'],
                                  [DateD, DateF, 'X'], TRUE);
            if (TCPD<>nil) then
              SommeBrut[j]:= SommeBrut[j]-TCPD.GetValue('PHB_MTREM');
            TCPD:= TCP.FindFirst (['PHB_DATEDEBUT', 'PHB_DATEFIN', 'PRM_INDEMCOMPCP'],
                                  [DateD, DateF, 'X'], TRUE);
            if (TCPD<>nil) then
              SommeBrut[j]:= SommeBrut[j]-TCPD.GetValue('PHB_MTREM');
          end;
        end;
      end else begin
        For j:=1 to 2 do
        begin
          if (DateF>=DebutDeMois(PlusMois (Date2, j-2))) and
             (DateF<=FinDeMois(PlusMois (Date2, j-2))) then
          begin
            if (VH_Paie.PGTenueEuro=FALSE) then
              SommeBrut[j]:= SommeBrut[j]+ THistoCumSalD.GetValue('PHC_MONTANT')
            else begin
              if (VH_Paie.PGDateBasculEuro<=DebutDeMois(PlusMois (Date2, j-2))) then
                SommeBrut[j]:= SommeBrut[j]+ THistoCumSalD.GetValue('PHC_MONTANT')
              else
                SommeBrut[j]:= SommeBrut[j]+ THistoCumSalD.GetValue('PHC_OMONTANT');
            end;
            TCPD:= TCP.FindFirst (['PHB_DATEDEBUT', 'PHB_DATEFIN', 'PRM_INDEMPREAVIS'],
                                  [DateD, DateF, 'X'], TRUE);
            if (TCPD<>nil) then
              SommeBrut[j]:= SommeBrut[j]-TCPD.GetValue('PHB_MTREM');
            TCPD:= TCP.FindFirst (['PHB_DATEDEBUT', 'PHB_DATEFIN', 'PRM_INDEMCOMPCP'],
                                  [DateD, DateF, 'X'], TRUE);
            if (TCPD<>nil) then
              SommeBrut[j]:= SommeBrut[j]-TCPD.GetValue('PHB_MTREM');
          end;
        end;
      end;
      THistoCumSalD:= THistoCumSal.FindNext(['PHC_CUMULPAIE'], ['01'], TRUE);
    end;
    FreeAndNil(THistoCumSal);
    FreeAndNil(TCP);
    TAssietteBrutD:= TAssietteBrut.FindFirst(['PCT_PRECOMPTEASS'], ['X'], FALSE);
    While (TAssietteBrutD<>nil) do
    begin
      DateF:= TAssietteBrutD.GetValue('PHB_DATEFIN');
      For j:=1 to 2 do
      begin
        if (DateF >= DebutDeMois(PlusMois (Date2, j-2))) and
           (DateF <= FinDeMois(PlusMois (Date2, j-2))) then
          if (VH_Paie.PGTenueEuro = FALSE) then
          begin
              Precompte[j]:= Precompte[j]+TAssietteBrutD.GetValue ('PHB_MTSALARIAL');
              break;
          end else BEGIN
            if (VH_Paie.PGDateBasculEuro <= DebutDeMois(PlusMois (Date2, j-2))) then
            begin
              Precompte[j]:= Precompte[j]+TAssietteBrutD.GetValue ('PHB_MTSALARIAL');
              break;
            end else begin
              Precompte[j]:= Precompte[j]+TAssietteBrutD.GetValue ('PHB_OMTSALARIAL');
              break;
            end;
          END;
      end;
      TAssietteBrutD:= TAssietteBrut.FindNext(['PCT_PRECOMPTEASS'], ['X'], FALSE);
    end;
    FreeAndNil(TAssietteBrut);
  end;
  For j:=1 to 2 do    //PT17 dans tous les cas, on insère les lignes
  begin
    if not(Payele[j]> 0) then
      Payele[j] := iDate1900;
    ExecuteSQL( 'INSERT INTO ATTSALAIRES( '
              + ' PAL_DATEDEBUT, PAL_DATEFIN,'
              + ' PAL_PAYELE, PAL_NBHEURES,'
              + ' PAL_JNONPAYES, PAL_SALAIRE, PAL_PRECOMPTE,'
              + ' PAL_OBSERVATIONS, PAL_MOIS, PAL_ORDRE, PAL_SALARIE, PAL_TYPEATTEST'
              + ' ) VALUES ( '
              + '"'+USDATETIME(DateDebut[j])+'", "'+USDATETIME(DateFin[j])+'", '
              + '"'+USDATETIME(Payele[j])+'", '+StrFPoint(Heures[j])+', '
              + '0, '+StrFPoint(SommeBrut[j])+', '+StrFPoint(Precompte[j])+', '
              + '"", '+IntToStr(15+j)+', '+IntToStr(ordre)+', "'+TobSalarie.GetString('PSA_SALARIE')+'", "ASS"'
              + ' )');
  end;
end;

function TOF_PGMULASSEDIC.GetOrdre: Integer;
var
  Q : TQuery;
  stSQL : String;
begin
  result := 1;
  stSQL := 'SELECT MAX(PAS_ORDRE) FROM ATTESTATIONS';
  Q:= OpenSQL(stSQL, True);
  if not Q.EOF then
    result := (Q.Fields[0].AsInteger)+1;
  Ferme(Q);
end;

procedure TOF_PGMULASSEDIC.OnUpdate;
var
i : Integer;
begin
inherited;
i:= 0;
if (GetControlText ('SITUATION')<>'') then
   i:= StrToInt (GetControlText ('SITUATION'));

{PT18
SetControlEnabled ('BCALCULLOT', ((i=1) or (i=4)));
}
SetControlEnabled ('BCALCULLOT', ((i=1) or (i=2) or (i=4)));
//FIN PT18
end;

initialization
registerclasses([TOF_PGMULASSEDIC]);
end.

