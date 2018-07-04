unit UControlFicU;

interface

uses
  SysUtils, Classes, Controls, Forms, Dialogs,
  stat, hmsgbox, ExtCtrls, StdCtrls, UTobView,
  HTB97, Hctrls, ComCtrls, UTOB, UVerif,Hgauge,UDecoupTRa, HEnt1, HSysMenu,
  ImgList, Menus, Graphics, HImgList, HPanel, Mask;

type

  TFControlFicU = class(TFStat)
        BOuvrirTRA        : TToolbarButton97;
        OpenDialog1       : TOpenDialog;
        CmbTypeLigne      : THValComboBox;
        LabelTypeAff      : THLabel;
        RGrpTypeAffich    : TRadioGroup;
        ComboModDecoup    : THValComboBox;
        HLabel2           : THLabel;
        HLabel3           : THLabel;
        BValidDecoup      : TToolbarButton97;
        LabelPeriode      : THLabel;
        ComboPeriode      : THValComboBox;
        SelectAffLabel    : THLabel;
        ComboTypeAffichage: THValComboBox;
        ComboChoixLettrage: THValComboBox;
        ComboNumCompt     : THValComboBox;
        LabelNomFichier   : THLabel;
        Edit_NomFichier   : TEdit97;
        LabelLigne        : THLabel;
        ComboBloc         : THValComboBox;
        ProgressTraitement: TEnhancedGauge;

        procedure FormClose(Sender: TObject; var Action: TCloseAction);
        procedure FormShow(Sender: TObject);
        procedure BOuvrirTRAClick(Sender: TObject);
        procedure CmbTypeLigneChange(Sender: TObject);
        procedure RGrpTypeAffichClick(Sender: TObject);
        procedure BValidDecoupClick(Sender: TObject);
        procedure ComboModDecoupChange(Sender: TObject);
        procedure ComboPeriodeChange(Sender: TObject);
        procedure ComboTypeAffichageChange(Sender: TObject);
        procedure ComboChoixLettrageChange(Sender: TObject);
        procedure ComboNumComptChange(Sender: TObject);
        procedure B_DecoupAutoClick(Sender: TObject);
        procedure ComboBlocChange(Sender: TObject);

        private
               FichierTRA : string;
               procedure CreerTOBMere;
               procedure ViderTOB;
               procedure LibererTOB;
               procedure initTableau;
               procedure RemplirCombo;
               procedure RecupEntete  (ligne : string);
               Procedure RecupParamG1 (ligne : string);
               Procedure RecupParamG2 (ligne : string);
               Procedure RecupParamG3 (ligne : string);
               Procedure RecupParamG4 (ligne : string);
               Procedure RecupParamG5 (ligne : string);
               procedure RecupJournaux(ligne : string);
               procedure RecupEcriture(ligne : string);
               procedure RecupCmptGnrx(ligne : string);
               procedure RecupExercice(ligne : string);
               procedure RecupCmptTiers(ligne : string);
               procedure RecupEtablsmnt(ligne : string);
               procedure RecupModPaie(ligne :string);

               procedure appelDecoup(adresse : string;Decoup:string);
               procedure TraiterLigne;
               procedure RechercheLettrage(EtatLettrage:string;numcompt:string);
               procedure AffTypLettrable(i:integer;TD:TOB;EtatLettrage:string);
               procedure AffNumComptTypeLettrable(i:integer;TD:TOB;EtatLettrage:string);
               procedure GererLigneInconnue(ligne : string);
               procedure LanceVerif;
               procedure RemplirComboPeriode(MTOBTouExercice,MTOBExercice:TOB);
               procedure RemplirComboNumCompt(MTOBTouCmptGnrx,MTOBCmptGnrx:TOB);
               
               function  TraiterTaille(Taille:variant):integer;
               function  ConvertTaille(var Taille:variant):integer;
               function  RecupPremChampTRA( ligne:string ):string;
               function  EstType(ligne:string ; Parametre:AffComEnreg):boolean;
               function  RecherchTOB(TypeLigne:TTypeligne):integer;
               function  ExistCombo(valeur:string):boolean;
               procedure TraiterBloc(Source:TFileStream;var adresse:string);
               Function FileStreamRecup(adresse:string) : Boolean;
               procedure TraiterLigneStream(Source:TFileStream;var adresse:string);

        public

     end;
     
Procedure ControlFichierTra (Fichier : string);

var
  FControlFicU                    : TFControlFicU;
  MTOBEnTete,MTOBParamG2          : TOB ;
  MTOBParamG3,MTOBParamG4         : TOB ;
  MTOBParamG5,MTOBParamG1         : TOB ;
  MTOBExercice,MTOBJournaux       : TOB ;
  MTOBEcriture,MTOBCmptTiers      : TOB ;
  MTOBCmptGnrx,MTOBInconnue       : TOB ;
  MTOBEtab,MTOBModPaie            : TOB ;

  FTOBEnTeteErr,FTOBParamG2Err    : TOB ;
  FTOBParamG3Err,FTOBParamG4Err   : TOB ;
  FTOBParamG5Err,FTOBParamG1Err   : TOB ;
  FTOBExerciceErr,FTOBJournauxErr : TOB ;
  FTOBEtabErr,FTOBModPaieErr      : TOB ;
  FTOBEcritureErr,FTOBCmptTiersErr,
  FTOBCmptGnrxErr,TOBMERE         : TOB ;

  MTOBTouEnTete,MTOBTouParamG2    : TOB ;
  MTOBTouParamG3,MTOBTouParamG4   : TOB ;
  MTOBTouParamG5,MTOBTouParamG1   : TOB ;
  MTOBTouExercice,MTOBTouJournaux : TOB ;
  MTOBTouEtab,MTOBTouModPaie      : TOB ;
  MTOBTouEcriture,MTOBTouCmptTiers,
  MTOBTouCmptGnrx                 : TOB;
  TOBLettrage                     : TOB;
  FICHIER                         : TextFile;
  Tab1Entete                      : array[1..35] of String;
  TabNatureJournaux               : array[1..12] of String;
  TabNatureCmptGnrx               : array[1..13] of String;
  numligne,nbEnreg,erreur         : integer;
  Verif                           : TVerif;
  EnregAffCom                     : array [0..15] of AffComEnreg;
  uniteTaille                     : string;
  
implementation

{$R *.DFM}

{****************TRAITEMENT DES LIGNE AVEC UTILISATION DE FILE STREAM***********}
{***********A.G.L.***********************************************
Auteur  ...... : Marion Koszczuk
Créé le ...... : 06/07/2004
Modifié le ... :   /  /
Description .. : Récupération des lignes soit avec FileStream soit avec un
Suite ........ : Text File suivant la taille du fichier pour optimiser les
Suite ........ : traitements
Mots clefs ... : OUVERTURE FICHIER RECUP LIGNE
*****************************************************************}
function TFControlFicU.FileStreamRecup(adresse:string) : Boolean;
var Source        : TFileStream;
    nomfichier    : string;
    size          : integer;
    Decoup        : Boolean;
begin
inherited;
       Decoup := FALSE; Source := nil;
       try
          nomfichier:=ExtractFileName(adresse);
          ReadTokenPipe(nomfichier,'.');
          ViderTob;
          if UpperCase(nomfichier)='TRA' then
          begin
               ProgressTraitement.Visible := True;
               Source := TFileStream.Create(adresse,fmOpenRead or fmShareDenyNone);
               size := TraiterTaille(Source.Size);
               Source.Position := 0;
               numligne := 0;
               ProgressTraitement.MaxValue:=Source.Size;
               if ((size>100) and (unitetaille='Mo')) then
               begin
                    PGIBox('Un découpage automatique du fichier améliorerait son temps de traitement');
                    TraiterBloc(Source,adresse);
                    Combobloc.visible:=True;
                    Decoup := TRUE;
               end else
               begin
                   if (unitetaille='Ko') then
                   begin
                         TraiterLigneStream(Source,adresse);
                   end else
                   begin
                        AssignFile(FICHIER , adresse);
                        Reset(Fichier);
                        Screen.Cursor := crHourGlass;
                        TraiterLigne;
                        CloseFile(Fichier);
                        Screen.Cursor := crDefault;
                   end;
                   FControlFicU.Caption := 'Contrôle fichier TRA : '+adresse;
                   RemplirComboPeriode(MTOBTouExercice,MTOBExercice);
                   RemplirComboNumCompt(MTOBTouCmptGnrx,MTOBCmptGnrx);
                   SelectAffLabel.Visible:=true;
                   ComboTypeAffichage.Visible :=true;
                   LanceVerif;
               end;
          end else
          begin
               PGIInfo('Le fichier selectionné n''est pas un fichier au format TRA');
               FichierTra :='';
          end;
     except
         PGIBox ('Erreur lors de la  récupération des données');
         Source.Free;
     end;
     ProgressTraitement.Visible:=false;
     Result := Decoup;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Marion Koszczuk
Créé le ...... : 06/07/2004
Modifié le ... : 06/07/2004
Description .. : RECUP DES LIGNES A L AIDE DE FILE STREAM POUR
Suite ........ : SUIVANT LA TAILLE DU FICHIER
Suite ........ :
Mots clefs ... : RECUP LIGNE STREAM
*****************************************************************}
procedure TFControlFicU.TraiterLigneStream(Source:TFileStream;var adresse:string);
var j : integer;
    chaine,id,ligne   : string;
begin
     try
        while (Source.Position<Source.Size) do
        begin
             SetLength(chaine,6);
             Source.Read(chaine[1],6);
             id := chaine;
             j := 0;
             while ( (j <= nbEnreg) and (EstType(id,EnregAffCom[j])=FALSE) ) do
                   j := j+1;
             ligne := '';
             while ((Source.Position<Source.Size) and (chaine<>#$D)) do  //tant q pas retour a la ligne=#$D
                  begin
                  Ligne := Ligne+chaine;
                  setLength(chaine,1);
                  Source.Read(chaine[1],1);
                  end;
             numligne := numligne+1;
             if( j<=nbEnreg ) then
             begin
                  EnregAffCom[j].TOBCharge(ligne);
             end else
                  GererLigneInconnue(ligne);
             Source.Position := Source.Position+1;
             ProgressTraitement.Progress :=Source.Position;
        end;
        Source.Free;
     except
        PGIInfo('Problème lors du traitement du fichier selectionné');
        Source.Free;
     end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Marion Koszczuk
Créé le ...... : 06/07/2004
Modifié le ... :   /  /    
Description .. : Decoupage du fichier en bloc de 10Mo maximum à l'aide de 
Suite ........ : File Steam pour l'affichage.
Mots clefs ... : DECOUP FICHIER FILE STREAM
*****************************************************************}
procedure TFControlFicU.TraiterBloc(Source:TFileStream;var adresse:string);
var TailleBloc,i,numbloc    :integer;
    nomfichier,chemin       :string;
    Cible                   :TFileStream;
begin
     ProgressTraitement.Visible:=TRUE;
     SelectAffLabel.Visible:=False;
     ComboTypeAffichage.Visible :=False;
     TV.Clear ;
     Source.Position:=0;
     numbloc := 0; Cible := nil;
     TailleBloc := 10485760; //10Mo
     nomfichier := ExtractFileName(adresse);
     chemin := ExtractFilePath(adresse);
     try
        for i:=1 to Source.Size div TailleBloc do
        begin
             ProgressTraitement.Progress :=Source.Position;
             numbloc := numbloc+1;
             Combobloc.Items.Add ('Bloc'+IntToStr(numbloc));
             adresse:= chemin+ IntToStr(numbloc)+nomfichier;
             Combobloc.Values.Add (adresse);
             Cible := TFileStream.Create(adresse, fmCreate or fmShareDenyNone);
             Cible.CopyFrom(Source, TailleBloc);
             Cible.Free;
        end;
        if (Source.Size mod TailleBloc) > 0 then
        begin
             // Il en reste un peu
             ProgressTraitement.Progress :=Source.Position;
             numbloc := numbloc+1;
             Combobloc.Items.Add ('Bloc'+IntToStr(numbloc));
             adresse:= chemin+ IntToStr(numbloc)+nomfichier;
             Combobloc.Values.Add (adresse);
             Cible := TFileStream.Create(adresse, fmCreate or fmShareDenyNone);
             Cible.CopyFrom(Source, Source.Size mod TailleBloc);
             Cible.Free;
        end;
     except
           PGIBox('Probleme lors du découpage du fichier');
           Cible.Free;
     end;
     ProgressTraitement.Visible:=false;
end;


Procedure ControlFichierTra (Fichier : string);
BEGIN
     FControlFicU := TFControlFicU.Create(Application) ;
     with  FControlFicU do
     begin
         CritereVisible := TRUE;
         FichierTRA := Fichier; // ajout me
         ShowModal ;
         Free ;
     end;
END;


{***********A.G.L.***********************************************
Auteur  ...... : Marion Koszczuk
Créé le ...... : 24/05/2004
Modifié le ... :   /  /
Description .. : Creation des TOBMere qui vont etre utilisées.
Mots clefs ... : TOB MERE
*****************************************************************}
procedure TFControlFicU.CreerTOBMere;
begin
   TOBMERE         := TOB.Create('TOBMERE',nil,-1);
   MTOBInconnue    := TOB.Create('TOBMereLigneInconnue',TOBMERE,-1);
   MTOBEnTete      := TOB.Create('TobMereEntete',TOBMERE,-1);
   MTOBParamG1     := TOB.Create('TobMereParamG1',TOBMERE,-1);
   MTOBParamG2     := TOB.Create('TobMereParamG2',TOBMERE,-1);
   MTOBParamG3     := TOB.Create('TobMereParamG3',TOBMERE,-1);
   MTOBParamG4     := TOB.Create('TobMereParamG4',TOBMERE,-1);
   MTOBParamG5     := TOB.Create('TobMereParamG5',TOBMERE,-1);
   MTOBExercice    := TOB.Create('TobMereExercice',TOBMERE,-1);
   MTOBEcriture    := TOB.Create('TobMereEcriture',TOBMERE,-1);
   MTOBJournaux    := TOB.Create('TobMereJournaux',TOBMERE,-1);
   MTOBCmptGnrx    := TOB.Create('TobMereCmptGnrx',TOBMERE,-1);
   MTOBCmptTiers   := TOB.Create('TobMereCmptTiers',TOBMERE,-1);
   MTOBEtab        := TOB.Create('TobMereEtablissmnt',TOBMERE,-1);
   MTOBModPaie     := TOB.Create('TobMerePai',TOBMERE,-1);

   FTOBEnTeteErr   := TOB.Create('TOBFilleErreur',TOBMERE,-1);
   FTOBParamG1Err  := TOB.Create('TOBFilleErreur2',TOBMERE,-1);
   FTOBParamG2Err  := TOB.Create('TOBFilleErreur3',TOBMERE,-1);
   FTOBParamG3Err  := TOB.Create('TOBFilleErreur4',TOBMERE,-1);
   FTOBParamG4Err  := TOB.Create('TOBFilleErreur5',TOBMERE,-1);
   FTOBParamG5Err  := TOB.Create('TOBFilleErreur6',TOBMERE,-1);
   FTOBExerciceErr := TOB.Create('TOBFilleErreur7',TOBMERE,-1);
   FTOBEcritureErr := TOB.Create('TOBFilleErreur8',TOBMERE,-1);
   FTOBJournauxErr := TOB.Create('TOBFilleErreur9',TOBMERE,-1);
   FTOBCmptGnrxErr := TOB.Create('TOBFilleErreur10',TOBMERE,-1);
   FTOBCmptTiersErr:= TOB.Create('TOBFilleErreur11',TOBMERE,-1);
   FTOBEtabErr     := TOB.Create('TOBFilleErreur12',TOBMERE,-1);
   FTOBModPaieErr  := TOB.Create('TobMerePaiErreur',TOBMERE,-1);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Marion Koszczuk
Créé le ...... : 24/05/2004
Modifié le ... :   /  /
Description .. : On vide le contenu des TOB pour leur réutilisation
Mots clefs ... : VIDE TOB
*****************************************************************}
procedure TFControlFicU.ViderTOB;
begin
   if MTOBInconnue <> nil then MTOBInconnue.ClearDetail;
   if MTOBEnTete <> nil  then MTOBEnTete.ClearDetail;
   if MTOBParamG1 <> nil  then MTOBParamG1.ClearDetail;
   if MTOBParamG2 <> nil  then MTOBParamG2.ClearDetail;
   if MTOBParamG3 <> nil  then MTOBParamG3.ClearDetail;
   if MTOBParamG4 <> nil  then MTOBParamG4.ClearDetail;
   if MTOBParamG5 <> nil  then MTOBParamG5.ClearDetail;
   if MTOBExercice <> nil  then MTOBExercice.ClearDetail;
   if MTOBEcriture <> nil  then MTOBEcriture.ClearDetail;
   if MTOBJournaux <> nil  then MTOBJournaux.ClearDetail;
   if MTOBCmptGnrx <> nil  then MTOBCmptGnrx.ClearDetail;
   if MTOBCmptTiers <> nil  then MTOBCmptTiers.ClearDetail;
   if MTOBEtab <> nil  then MTOBEtab.ClearDetail;
   if MTOBModPaie <> nil  then MTOBModPaie.ClearDetail;

   if FTOBEnTeteErr <> nil  then FTOBEnTeteErr  .ClearDetail;
   if FTOBParamG1Err <> nil  then FTOBParamG1Err .ClearDetail;
   if FTOBParamG2Err <> nil  then FTOBParamG2Err .ClearDetail;
   if FTOBParamG3Err <> nil  then FTOBParamG3Err .ClearDetail;
   if FTOBParamG4Err <> nil  then FTOBParamG4Err .ClearDetail;
   if FTOBParamG5Err <> nil  then FTOBParamG5Err .ClearDetail;
   if FTOBExerciceErr <> nil  then FTOBExerciceErr.ClearDetail;
   if FTOBEcritureErr <> nil  then FTOBEcritureErr.ClearDetail;
   if FTOBJournauxErr <> nil  then FTOBJournauxErr.ClearDetail;
   if FTOBCmptGnrxErr <> nil  then FTOBCmptGnrxErr.ClearDetail;
   if FTOBCmptTiersErr <> nil  then FTOBCmptTiersErr.ClearDetail;
   if FTOBEtabErr <> nil  then FTOBEtabErr.ClearDetail ;

End;

{***********A.G.L.***********************************************
Auteur  ...... : Marion Koszczuk
Créé le ...... : 25/05/2004
Modifié le ... :   /  /
Description .. : Libere la memoire occupée par les TOBMere créees au
Suite ........ : debut du programme
Mots clefs ... : LIBERE TOB
*****************************************************************}
procedure TFControlFicU.LibererTOB;
begin
   TOBMere.Free;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Marion Koszczuk
Créé le ...... : 24/05/2004
Modifié le ... :   /  /
Description .. : Correspon a l initialisation du tableau contenant le debut de
Suite ........ : chaq typ de ligne qi permet d'identifier la ligne
Mots clefs ... : INIT TABLEAU TYPE LIGNE
*****************************************************************}
procedure TFControlFicU.initTableau;
var i : integer;
begin
  i:=0;
  EnregAffCom[i].Typeligne[1] := '***S1';  EnregAffCom[i].Typeligne[2] := '***S3';
  EnregAffCom[i].Typeligne[3] := '***S5';  EnregAffCom[i].Typeligne[4] := '***S7';
  EnregAffCom[i].Typeligne[5] := '***QU';  EnregAffCom[i].Typeligne[6] := '***S5';
  EnregAffCom[i].Typeligne[7] := '***WT';  EnregAffCom[i].Typeligne[7] := '***WS';
  EnregAffCom[i].MTOB := MTOBEnTete;
  EnregAffCom[i].FTOBErreur  := FTOBEnTeteErr;
  EnregAffCom[i].MTOBTout  :=  MTOBTouEnTete;
  EnregAffCom[i].TOBCharge := RecupEntete ;
  EnregAffCom[i].Verif := ENT;
  EnregAffCom[i].ColName := 'Position ligne;Identifiant;Origine Fichier;Type Fichier;Format;CodeExoClo;DateBascule;DteArrPeriodiq;Version;NumDoCab;DteH;Utilisateur;Raison Sociale';
  i := i+1;
  EnregAffCom[i].Typeligne[1] := '***PS1';
  EnregAffCom[i].MTOB := MTOBParamG1;
  EnregAffCom[i].FTOBErreur := FTOBParamG1Err;
  EnregAffCom[i].MTOBTout  :=  MTOBTouParamG1;
  EnregAffCom[i].TOBCharge := RecupParamG1 ;
  EnregAffCom[i].Verif := PS1;
  EnregAffCom[i].ColName := 'Position ligne;Identifiant;Nom;Adr1;Adr2;CodePostal;Ville;Pays;Tel;Fax;Telex;Email;RVA;Contact;Nif;Siret;Rcs;APE;Capital;APE';
  i := i+1;
  EnregAffCom[i].Typeligne[1] := '***PS2';
  EnregAffCom[i].MTOB := MTOBParamG2;
  EnregAffCom[i].FTOBErreur := FTOBParamG2Err;
  EnregAffCom[i].MTOBTout  := MTOBTouParamG2;
  EnregAffCom[i].TOBCharge := RecupParamG2 ;
  EnregAffCom[i].Verif := PS2;
  EnregAffCom[i].ColName := 'Position ligne;Identifiant;Lggen;Bourregen;Lgaux;Bourreaux;Lgsec1;Bourresec1;Lgsec2;Bourresec2;Lgsec3;Bourresec3;Lgsec4;Bourresec4;Lgsec5;Bourresec5;GenAttend;CliAttend;FouAttend;SalaAttend;DivAttend;Sect1Attend;Sect2Attend;Sect3Attend';
  EnregAffCom[i].ColName := EnregAffCom[i].ColName+'Sect4Attend;Sect5Attend;GestPoint';
  i := i+1;
  EnregAffCom[i].Typeligne[1] := '***PS3';
  EnregAffCom[i].MTOB := MTOBParamG3;
  EnregAffCom[i].FTOBErreur := FTOBParamG3Err;
  EnregAffCom[i].MTOBTout  := MTOBTouParamG3;
  EnregAffCom[i].TOBCharge := RecupParamG3 ;
  EnregAffCom[i].Verif := PS3;
  EnregAffCom[i].ColName := 'Position ligne;Identifiant;Cptbil;Cptres;Cptben;Cptper;Jalouvre;JalFerme;JalBal;CliDefaut;FouDefaut;SalDefaut;DebDefaut;CreDefaut';
  i := i+1;
  EnregAffCom[i].Typeligne[1] := '***PS4';
  EnregAffCom[i].MTOB := MTOBParamG4;
  EnregAffCom[i].FTOBErreur := FTOBParamG4Err;
  EnregAffCom[i].MTOBTout  := MTOBTouParamG4;
  EnregAffCom[i].TOBCharge := RecupParamG4 ;
  EnregAffCom[i].Verif := PS4;
  EnregAffCom[i].ColName := 'Position ligne;Identifiant;BilDeb1;BilFin1;BilDeb2;BilFin2;BilDeb3;BilFin3;BilDeb4;BilFin4;BilDeb5;BilFin5;ChaDeb1;ChaFin1;ChaDeb2;ChaFin2;ChaDeb3;ChaFin3;ChaDeb4;ChaFin4;ChaDeb5;ChaFin5';
  EnregAffCom[i].ColName := EnregAffCom[i].ColName + 'ProDeb1;ProFin1;ProDeb2;ProFin2;ProDeb3;ProFin3;ProDeb4;ProFin4;ProDeb5;ProFin5';
  i := i+1;
  EnregAffCom[i].Typeligne[1] := '***PS5';
  EnregAffCom[i].MTOB := MTOBParamG5;
  EnregAffCom[i].FTOBErreur := FTOBParamG5Err;
  EnregAffCom[i].MTOBTout  := MTOBTouParamG5;
  EnregAffCom[i].TOBCharge := RecupParamG5 ;
  EnregAffCom[i].Verif := PS5;
  EnregAffCom[i].ColName := 'Position ligne;Identifiant;Devise;DecDevise;TenueEuro;JalECC;CpteCDD;CpteCCC;PariteEuro;RegimDefaut;MrDefaut;ExigDefaut;TvaDefaut;EtaBDefaut;PlanRef;GereANA';
  i := i+1;
  EnregAffCom[i].Typeligne[1] := '***EXO';
  EnregAffCom[i].MTOB := MTOBExercice;
  EnregAffCom[i].FTOBErreur := FTOBExerciceErr;
  EnregAffCom[i].MTOBTout  :=  MTOBTouExercice;
  EnregAffCom[i].TOBCharge := RecupExercice ;
  EnregAffCom[i].Verif := EXO;
  EnregAffCom[i].ColName := 'Position ligne;Identifiant;Code;DateDebut;DateFin;EtatCpta;EtatBudget;Libelle;ETATANO';
  i := i+1;
  EnregAffCom[i].Typeligne[1] := '***CGN';
  EnregAffCom[i].MTOB := MTOBCmptGnrx;
  EnregAffCom[i].FTOBErreur := FTOBCmptGnrxErr;
  EnregAffCom[i].MTOBTout  := MTOBTouCmptGnrx;
  EnregAffCom[i].TOBCharge := RecupCmptGnrx ;
  EnregAffCom[i].Verif := CGN;
  EnregAffCom[i].ColName := 'Position ligne;Identifiant;Code;Libelle;Nature;Lettrabl;Pointabl;VentilAx1;VentilAx2;VentilAx3;VentilAx4;VentilAx5;Tab1;Tab2;Tab3;Tab4;Tab5;Tab6;Tab7;Tab8;Tab9;Tab10;Abrege;Sens;Co1;Co2';
  i := i+1;
  EnregAffCom[i].Typeligne[1] := '***CAE';
  EnregAffCom[i].MTOB := MTOBCmptTiers;
  EnregAffCom[i].FTOBErreur := FTOBCmptTiersErr;
  EnregAffCom[i].MTOBTout  :=  MTOBTouCmptTiers;
  EnregAffCom[i].TOBCharge := RecupCmptTiers ;
  EnregAffCom[i].Verif := CAE;
  EnregAffCom[i].ColName := 'Position ligne;Identifiant;Code;libelle;Nature;Lettrabl;Collectif;Ean;Tab1;Tab2;Tab3;Tab4;Tab5;Tab6;Tab7;Tab8;Tab9;Tab10;Adr1;Adr2;Adr3;Cp;Ville;Domiciliation;Etablssmnt;Guichet;NumCpt;Cle;Pay;LibellAbrege;Langue;MultiDevise;DevisTier;Tel';
  EnregAffCom[i].ColName := EnregAffCom[i].ColName + 'Fax;RegimTVA;ModReglmnt;Commentaire;Nif;Siret;Ape;CNom;CService;CFonction;CTel;CFax;CTelex;CRva;CCivilite;CPpal;FormJur;RibPpal;TvaEncaissmnt';
  i := i+1;
  EnregAffCom[i].Typeligne[1] := '***JAL';
  EnregAffCom[i].MTOB := MTOBJournaux;
  EnregAffCom[i].FTOBErreur  := FTOBJournauxErr;
  EnregAffCom[i].MTOBTout  :=  MTOBTouJournaux;
  EnregAffCom[i].TOBCharge := RecupJournaux ;
  EnregAffCom[i].Verif := JAL;
  EnregAffCom[i].ColName :='Position ligne;Identifiant;Code;Libelle;Nature;Souches1;Souches2;Compte;Axe;Mode Saisie;Compte Auto;Compte interdit';
  i := i+1;
  EnregAffCom[i].Typeligne[1] := 'ECR';
  EnregAffCom[i].MTOB := MTOBEcriture;
  EnregAffCom[i].FTOBErreur := FTOBEcritureErr;
  EnregAffCom[i].MTOBTout  := MTOBTouEcriture;
  EnregAffCom[i].TOBCharge := RecupEcriture ;
  EnregAffCom[i].Verif := ECR;
  EnregAffCom[i].ColName :='Position ligne;CODE;DteCptable;TYPEPIECE;GENERAL;TYPECPTE;AUXILIARE OU SECTION;REFINTERNE;LIBELLE;MODEPAIE;ECHEANCE;SENS;MONTANT1;TYPEECRITURE;NumeroPieceDEVISE;TAUXDEV;CODEMONTANT;MONTANT2;MONTANT3;ETABLISSEMENT;AXE;NUMECHE;Nature mouvement;EtaLettrag;';
  i := i+1;
  EnregAffCom[i].Typeligne[1] := '***ETB';
  EnregAffCom[i].MTOB := MTOBEtab;
  EnregAffCom[i].FTOBErreur := FTOBEtabErr;
  EnregAffCom[i].MTOBTout  :=  MTOBTouEtab;
  EnregAffCom[i].TOBCharge := RecupEtablsmnt ;
  EnregAffCom[i].Verif := ETB;
  EnregAffCom[i].ColName :='Position ligne;identifiant;Code;Libelle';
  i := i+1;
  EnregAffCom[i].Typeligne[1]:='***MDP';
  EnregAffCom[i].MTOB:=MTOBModPaie;
  EnregAffCom[i].FTOBErreur := FTOBModPaieErr;
  EnregAffCom[i].MTOBTout  :=  MTOBTouModPaie;
  EnregAffCom[i].TOBCharge:= RecupModPaie ;
  EnregAffCom[i].Verif := MDP;
  EnregAffCom[i].ColName :='Position ligne;Identifiant;Code;Libelle;Categorie;Code Acceptation;Edition lettre cheques;Edition lettre traite;Condition montant;Montant Max;Mode de remplacement';

  nbEnreg:=i;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Marion Koszczuk
Créé le ...... : 24/05/2004
Modifié le ... :   /  /
Description .. : Rempli la combo de selection du type de ligne à afficher à
Suite ........ : l'aide d'une TStringList
Mots clefs ... : REMPLIR COMBO TYPE LIGNE
*****************************************************************}
procedure TFControlFicU.RemplirCombo;
begin
   CmbTypeLigne.Items.Add ('En Tete');
   CmbTypeLigne.Items.Add ('Parametre Generaux1');
   CmbTypeLigne.Items.Add ('Parametre Generaux2');
   CmbTypeLigne.Items.Add ('Parametre Generaux3');
   CmbTypeLigne.Items.Add ('Parametre Generaux4');
   CmbTypeLigne.Items.Add ('Parametre Generaux5');
   CmbTypeLigne.Items.Add ('Exercice');
   CmbTypeLigne.Items.Add ('Compte Generaux');
   CmbTypeLigne.Items.Add ('Compte de Tiers');
   CmbTypeLigne.Items.Add ('Journaux');
   CmbTypeLigne.Items.Add ('Ecriture');
   CmbTypeLigne.Items.Add('Etablissement');
   CmbTypeLigne.Items.Add('Mode de payement');

   CmbTypeLigne.Values.add('ENT');
   CmbTypeLigne.Values.add('PS1');
   CmbTypeLigne.Values.add('PS2');
   CmbTypeLigne.Values.add('PS3');
   CmbTypeLigne.Values.add('PS4');
   CmbTypeLigne.Values.add('PS5');
   CmbTypeLigne.Values.add('EXO');
   CmbTypeLigne.Values.add('CGN');
   CmbTypeLigne.Values.add('CAE');
   CmbTypeLigne.Values.add('JAL');
   CmbTypeLigne.Values.add('ECR');
   CmbTypeLigne.Values.add('ETB');
   CmbTypeLigne.Values.add('MDP');

   CmbTypeLigne.ItemIndex := 0;
end;


procedure TFControlFicU.FormShow(Sender: TObject);
var
Decoup                  : Boolean;
SavTRA,FileDir,Filename : string;
begin
  inherited;
  {  MMaTOB.SaveToFile('journal.txt',True,True,True);   }
   erreur:=0;
   CreerTOBMere;
   initTableau;
   if FichierTra <> '' then   // ajout me
   begin
            Combobloc.Clear;
            ComboPeriode.Clear;
            Combobloc.Visible:=false;
            Edit_NomFichier.Text := FichierTra;
            ProgressTraitement.Visible:=False;
            SavTRA := FichierTra;
            Decoup := FileStreamRecup(FichierTra);
            if Decoup then
            begin
                 FichierTra := SavTra;
                 FileDir    := ExtractFileDir (FichierTra) ;
                 Filename   := ExtractFileName(FichierTra);
                 FileStreamRecup(FileDir+'\1'+Filename);
            end;
   end;

   RemplirCombo;
   if FichierTra <> '' then     // ajout me
   begin
      ComboTypeAffichage.ItemIndex := 0;
      CmbTypeLigne.ItemIndex := 10;
      CmbTypeLigneChange (Sender);
   end;
end;


{***********A.G.L.***********************************************
Auteur  ...... : Marion Koszczuk
Créé le ...... : 24/05/2004
Modifié le ... :   /  /
Description .. : Recup les premier champ d une ligne de fichier TRA pour
Suite ........ : pouvoir ensuite
Mots clefs ... :
*****************************************************************}
function TFControlFicU.RecupPremChampTRA(ligne:string):string;//MARCHE
var caractere : string;
begin
     caractere := Copy(ligne,0,4);
     if(caractere[1]<>'*') then
     begin
          Result := Copy(ligne,0,3);
          exit;
     end;
     if ((caractere = '***S') or (caractere = '***Q') or (caractere = '***W')) then
     begin
               caractere := Copy(ligne,0,5);
     end else
     begin
           caractere := Copy(ligne,0,6);
     end;
     Result := caractere;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Marion Koszczuk
Créé le ...... : 17/05/2004
Modifié le ... :   /  /
Description .. : Fonction qui verifie de quelle type est la ligne passée en
Suite ........ : parametre
Mots clefs ... : VERIF TYPE LIGNE
*****************************************************************}
function TFControlFicU.EstType(ligne:string ; Parametre:AffComEnreg) : boolean;//MARCHE//
var resultat : string;
    VRAIFAUX : boolean;
    j,max    : integer;
begin
  VraiFaux := FALSE;
  j := 1;
  if( (trim(copy(ligne,0,1))='!') and (Parametre.Verif=ENT) ) then
  begin
      VraiFaux := TRUE;
  end else
  begin
       resultat := RecupPremChampTRA(ligne); //on recup les premier caractere de la ligne
       IF LENGTH(resultat)=5 then
       begin
            max:=6
       end else
            max:=1;
       while ((VRAIFAUX=FALSE)and (J<=max)) do
       begin
            if (Parametre.Typeligne[J] =resultat) then
            begin
                 VRAIFAUX := True;
            end else
            begin
                 if( (Copy(resultat,1,3)<>'***') and  (Parametre.Typeligne[J]='ECR')) then
                     VRAIFAUX:=TRUE;
            end;
            J := J+1;
       end;
  end;
  Result := VraiFaux;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Marion Koszczuk
Créé le ...... : 26/05/2004
Modifié le ... :   /  /
Description .. : Recuperation des lignes de type EnTete dans des TOB Fille
Suite ........ :  de la TOB MTOBEntete.
Mots clefs ... : RECUP ENTETE
*****************************************************************}
procedure TFControlFicU.RecupEnTete(ligne:string);
var TOBEnTete : TOB;
begin
     TOBEnTete := TOB.create('',MTOBEnTete,-1);
     TOBEntete.AddChampSupValeur ('Position ligne', IntToStr(numligne));
     if (Trim(Copy(ligne,0,1))='!')then
     begin
          TOBEntete.AddChampSupValeur('Identifiant',Trim(Copy(ligne,0,1)));
          TOBEntete.AddChampSupValeur('Origine Fichier','Le dossier existe deja dans la base');
     end else
     begin

          TOBEntete.AddChampSupValeur ('Identifiant',    Copy(ligne,4,2));
          TOBEntete.AddChampSupValeur ('Origine Fichier',Copy(ligne,6,3));
          TOBEntete.AddChampSupValeur ('Type Fichier',   Copy(ligne,9,3));
          TOBEntete.AddChampSupValeur ('Format',         Copy(ligne,12,3));
          TOBEntete.AddChampSupValeur ('CodeExoClo',     Copy(ligne,15,3));
          TOBEntete.AddChampSupValeur ('DateBascule',    Copy(ligne,18,8));
          TOBEntete.AddChampSupValeur ('DteArrPeriodiq', Copy(ligne,26,8));
          TOBEntete.AddChampSupValeur ('Version',        Copy(ligne,34,3));
          TOBEntete.AddChampSupValeur ('NumDoCab',       Copy(ligne,37,5));
          TOBEntete.AddChampSupValeur ('DteH',           Copy(ligne,42,12));
          TOBEntete.AddChampSupValeur ('Utilisateur',    Copy(ligne,54,35));
          TOBEntete.AddChampSupValeur ('Raison Sociale', Copy(ligne,89,35));
     end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Marion Koszczuk
Créé le ...... : 26/05/2004
Modifié le ... : 26/05/2004
Description .. : Recuperation des lignes de type ParamG1 dans des TOB
Suite ........ : Fille
Suite ........ :  de la TOB MTOBParamG1
Mots clefs ... : RECUP PARAMETRE G1
*****************************************************************}
Procedure TFControlFicU.RecupParamG1(ligne:string);
var TOBParamG1 : TOB;
begin
      TOBParamG1 := Tob.Create ('',MTOBParamG1,-1);
      TOBParamG1.AddChampSupValeur ('Position ligne', IntToStr(numligne));
      TOBParamG1.AddChampSupValeur ('Identifiant',Copy(ligne,4,3));
      TOBParamG1.AddChampSupValeur ('Nom',        Copy(ligne,7,35));
      TOBParamG1.AddChampSupValeur ('Adr1',       Copy(ligne,42,35));
      TOBParamG1.AddChampSupValeur ('Adr2',       Copy(ligne,112,35));
      TOBParamG1.AddChampSupValeur ('CodePostal', Copy(ligne,147,9));
      TOBParamG1.AddChampSupValeur ('Ville',      Copy(ligne,156,35));
      TOBParamG1.AddChampSupValeur ('Pays',       Copy(ligne,191,3));
      TOBParamG1.AddChampSupValeur ('Tel',        Copy(ligne,194,25));
      TOBParamG1.AddChampSupValeur ('Fax',        Copy(ligne,219,25));
      TOBParamG1.AddChampSupValeur ('Telex',      Copy(ligne,244,25));
      TOBParamG1.AddChampSupValeur ('Email',      Copy(ligne,269,25));
      TOBParamG1.AddChampSupValeur ('RVA',        Copy(ligne,304,35));
      TOBParamG1.AddChampSupValeur ('Contact',    Copy(ligne,339,35));
      TOBParamG1.AddChampSupValeur ('Nif',        Copy(ligne,374,35));
      TOBParamG1.AddChampSupValeur ('Siret',      Copy(ligne,391,17));
      TOBParamG1.AddChampSupValeur ('Rcs',        Copy(ligne,426,35));
      TOBParamG1.AddChampSupValeur ('APE',        Copy(ligne,443,17));
      TOBParamG1.AddChampSupValeur ('Capital',    Copy(ligne,478,35));
      TOBParamG1.AddChampSupValeur ('APE',        Copy(ligne,498,1));
end;

{***********A.G.L.***********************************************
Auteur  ...... : Marion Koszczuk
Créé le ...... : 26/05/2004
Modifié le ... :   /  /
Description .. : Recuperation des lignes de type ParamG2 dans des TOB
Suite ........ : Fille
Suite ........ :  de la TOB MTOBParamG2
Mots clefs ... : RECUP PARAM G2
*****************************************************************}
Procedure TFControlFicU.RecupParamG2(ligne:string);
var TOBParamG2 : TOB;
begin
      TOBParamG2 := Tob.Create ('',MTOBParamG2,-1);
      TOBParamG2.AddChampSupValeur ('Position ligne', IntToStr(numligne));
      TOBParamG2.AddChampSupValeur ('Identifiant',Copy(ligne,4,3));
      TOBParamG2.AddChampSupValeur ('Lggen',      Copy(ligne,7,2));
      TOBParamG2.AddChampSupValeur ('Bourregen',  Copy(ligne,9,1));
      TOBParamG2.AddChampSupValeur ('Lgaux',      Copy(ligne,10,2));
      TOBParamG2.AddChampSupValeur ('Bourreaux',  Copy(ligne,12,1));
      TOBParamG2.AddChampSupValeur ('Lgsec1',     Copy(ligne,13,2));
      TOBParamG2.AddChampSupValeur ('Bourresec1', Copy(ligne,15,1));
      TOBParamG2.AddChampSupValeur ('Lgsec2',     Copy(ligne,16,2));
      TOBParamG2.AddChampSupValeur ('Bourresec2', Copy(ligne,18,1));
      TOBParamG2.AddChampSupValeur ('Lgsec3',     Copy(ligne,19,2));
      TOBParamG2.AddChampSupValeur ('Bourresec3', Copy(ligne,21,1));
      TOBParamG2.AddChampSupValeur ('Lgsec4',     Copy(ligne,22,2));
      TOBParamG2.AddChampSupValeur ('Bourresec4', Copy(ligne,24,1));
      TOBParamG2.AddChampSupValeur ('Lgsec5',     Copy(ligne,25,2));
      TOBParamG2.AddChampSupValeur ('Bourresec5', Copy(ligne,27,1));
      TOBParamG2.AddChampSupValeur ('GenAttend',  Copy(ligne,28,17));
      TOBParamG2.AddChampSupValeur ('CliAttend',  Copy(ligne,45,17));
      TOBParamG2.AddChampSupValeur ('FouAttend',  Copy(ligne,62,17));
      TOBParamG2.AddChampSupValeur ('SalaAttend', Copy(ligne,79,17));
      TOBParamG2.AddChampSupValeur ('DivAttend',  Copy(ligne,96,17));
      TOBParamG2.AddChampSupValeur ('Sect1Attend',Copy(ligne,113,17));
      TOBParamG2.AddChampSupValeur ('Sect2Attend',Copy(ligne,130,17));
      TOBParamG2.AddChampSupValeur ('Sect3Attend',Copy(ligne,147,17));
      TOBParamG2.AddChampSupValeur ('Sect4Attend',Copy(ligne,164,17));
      TOBParamG2.AddChampSupValeur ('Sect5Attend',Copy(ligne,181,17));
      TOBParamG2.AddChampSupValeur ('GestPoint',  Copy(ligne,198,3));
end;

{***********A.G.L.***********************************************
Auteur  ...... : Marion Koszczuk
Créé le ...... : 26/05/2004
Modifié le ... :   /  /
Description .. : Recuperation des lignes de type ParamG1 dans des TOB
Suite ........ : Fille
Suite ........ :  de la TOB MTOBParamG1
Mots clefs ... : RECUP PARAM G3
*****************************************************************}
procedure TFControlFicU.RecupParamG3(ligne:string);
var TOBParamG3 : TOB;
begin
      TOBParamG3 := TOB.Create('',MTOBParamG3,-1);
      TOBParamG3.AddChampSupValeur ('Position ligne', IntToStr(numligne));
      TOBParamG3.AddChampSupValeur ('Identifiant',Copy(ligne,4,3));
      TOBParamG3.AddChampSupValeur ('Cptbil',     Copy(ligne,7,17));
      TOBParamG3.AddChampSupValeur ('Cptres',     Copy(ligne,24,17));
      TOBParamG3.AddChampSupValeur ('Cptben',     Copy(ligne,41,17));
      TOBParamG3.AddChampSupValeur ('Cptper',     Copy(ligne,58,17));
      TOBParamG3.AddChampSupValeur ('Jalouvre',   Copy(ligne,75,3));
      TOBParamG3.AddChampSupValeur ('JalFerme',   Copy(ligne,78,3));
      TOBParamG3.AddChampSupValeur ('JalBal',     Copy(ligne,81,3));
      TOBParamG3.AddChampSupValeur ('CliDefaut',  Copy(ligne,84,17));
      TOBParamG3.AddChampSupValeur ('FouDefaut',  Copy(ligne,101,17));
      TOBParamG3.AddChampSupValeur ('SalDefaut',  Copy(ligne,118,17));
      TOBParamG3.AddChampSupValeur ('DebDefaut',  Copy(ligne,135,17));
      TOBParamG3.AddChampSupValeur ('CreDefaut',  Copy(ligne,152,17));
end;


{***********A.G.L.***********************************************
Auteur  ...... : Marion Koszczuk
Créé le ...... : 26/05/2004
Modifié le ... :   /  /
Description .. : Recuperation des lignes de type ParamG4 dans des TOB
Suite ........ : Fille
Suite ........ :  de la TOB MTOBParamG4
Mots clefs ... : RECUP PARAM G4
*****************************************************************}
procedure TFControlFicU.RecupParamG4(ligne:string);
var TOBParamG4 : TOB;
begin
       TOBParamG4 := TOB.Create('',MTOBParamG4,-1);
       TOBParamG4.AddChampSupValeur ('Position ligne', IntToStr(numligne));
       TOBParamG4.AddChampSupValeur ('Identifiant',Copy(ligne,4,3));
       TOBParamG4.AddChampSupValeur ('BilDeb1',    Copy(ligne,7,17));
       TOBParamG4.AddChampSupValeur ('BilFin1',    Copy(ligne,24,17));
       TOBParamG4.AddChampSupValeur ('BilDeb2',    Copy(ligne,41,17));
       TOBParamG4.AddChampSupValeur ('BilFin2',    Copy(ligne,58,17));
       TOBParamG4.AddChampSupValeur ('BilDeb3',    Copy(ligne,75,17));
       TOBParamG4.AddChampSupValeur ('BilFin3',    Copy(ligne,92,17));
       TOBParamG4.AddChampSupValeur ('BilDeb4',    Copy(ligne,109,17));
       TOBParamG4.AddChampSupValeur ('BilFin4',    Copy(ligne,126,17));
       TOBParamG4.AddChampSupValeur ('BilDeb5',    Copy(ligne,143,17));
       TOBParamG4.AddChampSupValeur ('BilFin5',    Copy(ligne,160,17));
       TOBParamG4.AddChampSupValeur ('ChaDeb1 ',   Copy(ligne,177,17));
       TOBParamG4.AddChampSupValeur ('ChaFin1',    Copy(ligne,194,17));
       TOBParamG4.AddChampSupValeur ('ChaDeb2 ',   Copy(ligne,211,17));
       TOBParamG4.AddChampSupValeur ('ChaFin2',    Copy(ligne,228,17));
       TOBParamG4.AddChampSupValeur ('ChaDeb3 ',   Copy(ligne,245,17));
       TOBParamG4.AddChampSupValeur ('ChaFin3',    Copy(ligne,262,17));
       TOBParamG4.AddChampSupValeur ('ChaDeb4 ',   Copy(ligne,279,17));
       TOBParamG4.AddChampSupValeur ('ChaFin4',    Copy(ligne,296,17));
       TOBParamG4.AddChampSupValeur ('ChaDeb5 ',   Copy(ligne,313,17));
       TOBParamG4.AddChampSupValeur ('ChaFin5',    Copy(ligne,330,17));
       TOBParamG4.AddChampSupValeur ('ProDeb1 ',   Copy(ligne,347,17));
       TOBParamG4.AddChampSupValeur ('ProFin1',    Copy(ligne,364,17));
       TOBParamG4.AddChampSupValeur ('ProDeb2 ',   Copy(ligne,381,17));
       TOBParamG4.AddChampSupValeur ('ProFin2',    Copy(ligne,398,17));
       TOBParamG4.AddChampSupValeur ('ProDeb3 ',   Copy(ligne,415,17));
       TOBParamG4.AddChampSupValeur ('ProFin3',    Copy(ligne,432,17));
       TOBParamG4.AddChampSupValeur ('ProDeb4 ',   Copy(ligne,449,17));
       TOBParamG4.AddChampSupValeur ('ProFin4',    Copy(ligne,466,17));
       TOBParamG4.AddChampSupValeur ('ProDeb5 ',   Copy(ligne,483,17));
       TOBParamG4.AddChampSupValeur ('ProFin5',    Copy(ligne,500,17));
end;

{***********A.G.L.***********************************************
Auteur  ...... : Marion Koszczuk
Créé le ...... : 26/05/2004
Modifié le ... :   /  /
Description .. : Recuperation des lignes de type ParamG5 dans des TOB
Suite ........ : Fille
Suite ........ :  de la TOB MTOBParamG5
Mots clefs ... : RECUP PARAM G5
*****************************************************************}
procedure TFControlFicU.RecupParamG5(ligne:string);
var TOBParamG5 : TOB;
begin
       TOBParamG5 := TOB.Create('',MTOBParamG5,-1);
       TOBParamG5.AddChampSupValeur ('Position ligne', IntToStr(numligne));
       TOBParamG5.AddChampSupValeur ('Identifiant',Copy(ligne,4,3));
       TOBParamG5.AddChampSupValeur ('Devise',     Trim(Copy(ligne,7,3)));
       TOBParamG5.AddChampSupValeur ('DecDevise',  Copy(ligne,10,1));
       TOBParamG5.AddChampSupValeur ('TenueEuro',  Copy(ligne,11,1));
       TOBParamG5.AddChampSupValeur ('JalECC',     Copy(ligne,12,3));
       TOBParamG5.AddChampSupValeur ('CpteCDD',    Copy(ligne,15,17));
       TOBParamG5.AddChampSupValeur ('CpteCCC',    Copy(ligne,32,17));
       TOBParamG5.AddChampSupValeur ('PariteEuro', Copy(ligne,49,10));
       TOBParamG5.AddChampSupValeur ('RegimDefaut',Copy(ligne,59,3));
       TOBParamG5.AddChampSupValeur ('MrDefaut',   Copy(ligne,62,3));
       TOBParamG5.AddChampSupValeur ('ExigDefaut', Copy(ligne,65,3));
       TOBParamG5.AddChampSupValeur ('TvaDefaut',  Copy(ligne,68,3));
       TOBParamG5.AddChampSupValeur ('EtaBDefaut', Copy(ligne,71,3));
       TOBParamG5.AddChampSupValeur ('PlanRef',    Copy(ligne,74,2));
       TOBParamG5.AddChampSupValeur ('GereANA',    Copy(ligne,76,1));
end;

{***********A.G.L.***********************************************
Auteur  ...... : Marion Koszczuk
Créé le ...... : 26/05/2004
Modifié le ... :   /  /
Description .. : Recuperation des lignes de type Exercice dans des TOB
Suite ........ : Fille
Suite ........ :  de la TOB MTOBExercice
Mots clefs ... :
*****************************************************************}
procedure TFControlFicU.RecupExercice(ligne:string);
var TOBExercice : TOB;
begin
     TOBExercice := TOB.Create('',MTOBExercice,-1);
     TOBExercice.AddChampSupValeur ('Position ligne', IntToStr(numligne));
     TOBExercice.AddChampSupValeur ('Identifiant',Copy(ligne,4,3));
     TOBExercice.AddChampSupValeur ('Code',       Copy(ligne,7,3));
     TOBExercice.AddChampSupValeur ('DateDebut',  Copy(ligne,10,8));
     TOBExercice.AddChampSupValeur ('DateFin',    Copy(ligne,18,8));
     TOBExercice.AddChampSupValeur ('EtatCpta',   Copy(ligne,26,3));
     TOBExercice.AddChampSupValeur ('EtatBudget', Copy(ligne,29,3));
     TOBExercice.AddChampSupValeur ('Libelle',    Copy(ligne,32,15));
     TOBExercice.AddChampSupValeur ('EtatAno',    Copy(ligne,67,3));
end;

{***********A.G.L.***********************************************
Auteur  ...... : Marion Koszczuk
Créé le ...... : 24/05/2004
Modifié le ... :   /  /
Description .. : Procedure qui cree la TOB FILLE concernan les lignes de
Suite ........ : type journaux.
Mots clefs ... : TOB JOURNAUX
*****************************************************************}
procedure TFControlFicU.RecupJournaux(ligne:string);
var TOBJournaux : TOB;
begin
     TOBJournaux := TOB.Create('',MTOBJournaux,-1);
     TOBJournaux.AddChampSupValeur ('Position ligne', IntToStr(numligne));
     TOBJournaux.AddChampSupValeur ('Identifiant',    Copy(ligne,4,3));
     TOBJournaux.AddChampSupValeur ('Code',           Copy(ligne,7,3));
     TOBJournaux.AddChampSupValeur ('Libelle',        Copy(ligne,10,35));
     TOBJournaux.AddChampSupValeur ('Nature',         Copy(ligne,45,3));
     TOBJournaux.AddChampSupValeur ('Souches1',       Copy(ligne,48,3));
     TOBJournaux.AddChampSupValeur ('Souches2',       Copy(ligne,51,3));
     TOBJournaux.AddChampSupValeur ('Compte',         Copy(ligne,54,17));
     TOBJournaux.AddChampSupValeur ('Axe',            Copy(ligne,71,3));
     TOBJournaux.AddChampSupValeur ('Mode Saisie',    Copy(ligne,74,3));
     TOBJournaux.AddChampSupValeur ('Compte Auto',    Copy(ligne,77,200));
     TOBJournaux.AddChampSupValeur ('Compte Interdit',Copy(ligne,277,200));
end;

{***********A.G.L.***********************************************
Auteur  ...... : Marion Koszczuk
Créé le ...... : 24/05/2004
Modifié le ... :   /  /
Description .. : Recuperation des lignes de type Ecriture dans une TOB
Suite ........ : Ecriture
Mots clefs ... : TOB ECRITURE
*****************************************************************}
procedure TFControlFicU.RecupEcriture(ligne:string);
var TOBEcriture : TOB;
begin
     TOBEcriture := TOB.Create('',MTOBEcriture,-1);
     TOBEcriture.AddChampSupValeur ('Position ligne', IntToStr(numligne));
     TOBEcriture.AddChampSupValeur ('Code',                Copy(ligne,1,3));
     TOBEcriture.AddChampSupValeur ('DteCptable',          Copy(ligne,4,8));
     TOBEcriture.AddChampSupValeur ('TypePiece',           Copy(ligne,12,2));
     TOBEcriture.AddChampSupValeur ('General',             Copy(ligne,14,17));
     TOBEcriture.AddChampSupValeur ('TypeCpte',            Copy(ligne,31,1));
     TOBEcriture.AddChampSupValeur ('Auxiliare ou Section',Copy(ligne,32,17));
     TOBEcriture.AddChampSupValeur ('RefInterne',          Copy(ligne,49,35));
     TOBEcriture.AddChampSupValeur ('Libelle',             Copy(ligne,84,35));
     TOBEcriture.AddChampSupValeur ('ModePaie',            Copy(ligne,119,3));
     TOBEcriture.AddChampSupValeur ('Echeance',            Copy(ligne,122,8));
     TOBEcriture.AddChampSupValeur ('Sens',                Copy(ligne,130,1));
     if Trim(Copy(Ligne,131,20))<>'' then
     begin
        TOBEcriture.AddChampSupValeur ('Montant1',           Valeur(StPoint(TRIM(Copy(Ligne, 131, 20)))));
     end else
     begin
        TOBEcriture.AddChampSupValeur ('Montant1',           0);
     end;
     TOBEcriture.AddChampSupValeur ('TypeEcriture',        Copy(ligne,151,1));
     TOBEcriture.AddChampSupValeur ('NumeroPiece',         Copy(ligne,152,8));
     TOBEcriture.AddChampSupValeur ('Devise',              Copy(ligne,160,3));
     TOBEcriture.AddChampSupValeur ('TauxDev',             Copy(ligne,163,10));
     TOBEcriture.AddChampSupValeur ('CodeMontant',         Copy(ligne,173,3));

     if Trim(Copy(Ligne,176,20))<>'' then
     begin
          TOBEcriture.AddChampSupValeur ('Montant2',            Copy(ligne,176,20));
     end else
     begin
          TOBEcriture.AddChampSupValeur ('Montant2',          0);
     end;
     if Trim(Copy(Ligne,196,20))<>'' then
     begin
          TOBEcriture.AddChampSupValeur ('Montant3',            Copy(ligne,196,20));
     end else
     begin
          TOBEcriture.AddChampSupValeur ('Montant3',           0);
     end;

     TOBEcriture.AddChampSupValeur ('Etablissement',       Copy(ligne,216,3));
     TOBEcriture.AddChampSupValeur ('Axe',                 Copy(ligne,219,2));
     TOBEcriture.AddChampSupValeur ('NumEche',             Copy(ligne,221,3));
     TOBEcriture.AddChampSupValeur ('DateCreation',        Copy(ligne,266,8));
     TOBEcriture.AddChampSupValeur ('Nature mouvement',    Copy(ligne,302,3));
     TOBEcriture.AddChampSupValeur ('RegimTVA',            Copy(ligne,386,1));
     TOBEcriture.AddChampSupValeur ('CodeTVA',             Copy(ligne,387,3));
     TOBEcriture.AddChampSupValeur ('CodeTPF',             Copy(ligne,390,3));
     TOBEcriture.AddChampSupValeur ('Lettrage',            Copy(ligne,1014,5));
     TOBEcriture.AddChampSupValeur ('EtaLettrag',          Trim(Copy(ligne,1021,17)));
end;

procedure TFControlFicU.RecupCmptGnrx(ligne:string);
var TOBCmptGnrx : TOB;
begin
    TOBCmptGnrx := TOB.Create('',MTOBCmptGnrx,-1);
    TOBCmptGnrx.AddChampSupValeur ('Position ligne', IntToStr(numligne));
    TOBCmptGnrx.AddChampSupValeur ('Identifiant ',Copy(ligne,4,3));
    TOBCmptGnrx.AddChampSupValeur ('Code  ',      Copy(ligne,7,17));
    TOBCmptGnrx.AddChampSupValeur ('Libelle ',    Copy(ligne,24,35));
    TOBCmptGnrx.AddChampSupValeur ('Nature ',     Copy(ligne,59,3));
    TOBCmptGnrx.AddChampSupValeur ('Lettrabl',    Copy(ligne,62,1));
    TOBCmptGnrx.AddChampSupValeur ('Pointabl',    Copy(ligne,63,1));
    TOBCmptGnrx.AddChampSupValeur ('VentilAx1',   Copy(ligne,64,1));
    TOBCmptGnrx.AddChampSupValeur ('VentilAx2',   Copy(ligne,65,1));
    TOBCmptGnrx.AddChampSupValeur ('VentilAx3',   Copy(ligne,66,1));
    TOBCmptGnrx.AddChampSupValeur ('VentilAx4',   Copy(ligne,67,1));
    TOBCmptGnrx.AddChampSupValeur ('VentilAx5',   Copy(ligne,68,1));
    TOBCmptGnrx.AddChampSupValeur ('Tab1 ',       Copy(ligne,69,17));
    TOBCmptGnrx.AddChampSupValeur ('Tab2 ',       Copy(ligne,86,17));
    TOBCmptGnrx.AddChampSupValeur ('Tab3 ',       Copy(ligne,103,17));
    TOBCmptGnrx.AddChampSupValeur ('Tab4 ',       Copy(ligne,120,17));
    TOBCmptGnrx.AddChampSupValeur ('Tab5 ',       Copy(ligne,137,17));
    TOBCmptGnrx.AddChampSupValeur ('Tab6 ',       Copy(ligne,154,17));
    TOBCmptGnrx.AddChampSupValeur ('Tab7 ',       Copy(ligne,171,17));
    TOBCmptGnrx.AddChampSupValeur ('Tab8 ',       Copy(ligne,188,17));
    TOBCmptGnrx.AddChampSupValeur ('Tab9 ',       Copy(ligne,205,17));
    TOBCmptGnrx.AddChampSupValeur ('Tab10',       Copy(ligne,222,17));
    TOBCmptGnrx.AddChampSupValeur ('Abrege',      Copy(ligne,239,17));
    TOBCmptGnrx.AddChampSupValeur ('Sens ',       Copy(ligne,256,3));
    TOBCmptGnrx.AddChampSupValeur ('Co1 ',        Copy(ligne,259,17));
    TOBCmptGnrx.AddChampSupValeur ('Co2 ',        Copy(ligne,276,17));
end;

procedure TFControlFicU.RecupCmptTiers(ligne:string);
var TOBCmptTiers : TOB;
begin
     TOBCmptTiers := TOB.Create('',MTOBCmptTiers,-1);
     TOBCmptTiers.AddChampSupValeur ('Position ligne', IntToStr(numligne));
     TOBCmptTiers.AddChampSupValeur ('Identifiant ',Copy(ligne,4,3));
     TOBCmptTiers.AddChampSupValeur ('Code ',       Copy(ligne,7,17));
     TOBCmptTiers.AddChampSupValeur ('libelle ',    Copy(ligne,24,35));
     TOBCmptTiers.AddChampSupValeur ('Nature ',     Copy(ligne,59,3));
     TOBCmptTiers.AddChampSupValeur ('Lettrabl',    Copy(ligne,62,1));
     TOBCmptTiers.AddChampSupValeur ('Collectif ',  Copy(ligne,63,17));
     TOBCmptTiers.AddChampSupValeur ('Ean ',        Copy(ligne,80,17));
     TOBCmptTiers.AddChampSupValeur ('Tab1  ',      Copy(ligne,97,17));
     TOBCmptTiers.AddChampSupValeur ('Tab2 ',       Copy(ligne,114,17));
     TOBCmptTiers.AddChampSupValeur ('Tab3 ',       Copy(ligne,131,17));
     TOBCmptTiers.AddChampSupValeur ('Tab4 ',       Copy(ligne,148,17));
     TOBCmptTiers.AddChampSupValeur ('Tab5 ',       Copy(ligne,165,17));
     TOBCmptTiers.AddChampSupValeur ('Tab6 ',       Copy(ligne,182,17));
     TOBCmptTiers.AddChampSupValeur ('Tab7 ',       Copy(ligne,199,17));
     TOBCmptTiers.AddChampSupValeur ('Tab8 ',       Copy(ligne,216,17));
     TOBCmptTiers.AddChampSupValeur ('Tab9 ',       Copy(ligne,233,17));
     TOBCmptTiers.AddChampSupValeur ('Tab10 ',      Copy(ligne,250,17));
     TOBCmptTiers.AddChampSupValeur ('Adr1  ',      Copy(ligne,267,35));
     TOBCmptTiers.AddChampSupValeur ('Adr2 ',       Copy(ligne,302,35));
     TOBCmptTiers.AddChampSupValeur ('Adr3 ',       Copy(ligne,337,35));
     TOBCmptTiers.AddChampSupValeur ('Cp  ',        Copy(ligne,372,9));
     TOBCmptTiers.AddChampSupValeur ('Ville ',      Copy(ligne,381,35));
     TOBCmptTiers.AddChampSupValeur ('Domiciliation ',Copy(ligne,416,24));
     TOBCmptTiers.AddChampSupValeur ('Etablssmnt',  Copy(ligne,440,5));
     TOBCmptTiers.AddChampSupValeur ('Guichet ',    Copy(ligne,445,5));
     TOBCmptTiers.AddChampSupValeur ('NumCpt ',     Copy(ligne,450,11));
     TOBCmptTiers.AddChampSupValeur ('Cle',         Copy(ligne,461,2));
     TOBCmptTiers.AddChampSupValeur ('Pay ',        Copy(ligne,463,3));
     TOBCmptTiers.AddChampSupValeur ('LibellAbrege',Copy(ligne,466,17));
     TOBCmptTiers.AddChampSupValeur ('Langue',      Copy(ligne,483,3));
     TOBCmptTiers.AddChampSupValeur ('MultiDevise', Copy(ligne,486,1));
     TOBCmptTiers.AddChampSupValeur ('DevisTier  ', Copy(ligne,487,3));
     TOBCmptTiers.AddChampSupValeur ('Tel  ',       Copy(ligne,490,25));
     TOBCmptTiers.AddChampSupValeur ('Fax ',        Copy(ligne,515,25));
     TOBCmptTiers.AddChampSupValeur ('RegimTVA  ',  Copy(ligne,540,3));
     TOBCmptTiers.AddChampSupValeur ('ModReglmnt',  Copy(ligne,543,3));
     TOBCmptTiers.AddChampSupValeur ('Commentaire', Copy(ligne,546,35));
     TOBCmptTiers.AddChampSupValeur ('Nif  ',       Copy(ligne,581,17));
     TOBCmptTiers.AddChampSupValeur ('Siret ',      Copy(ligne,598,17));
     TOBCmptTiers.AddChampSupValeur ('Ape ',        Copy(ligne,615,5));
     TOBCmptTiers.AddChampSupValeur ('CNom  ',      Copy(ligne,620,35));
     TOBCmptTiers.AddChampSupValeur ('CService ',   Copy(ligne,655,35));
     TOBCmptTiers.AddChampSupValeur ('CFonction  ', Copy(ligne,690,35));
     TOBCmptTiers.AddChampSupValeur ('CTel  ',      Copy(ligne,725,25));
     TOBCmptTiers.AddChampSupValeur ('CFax ',       Copy(ligne,750,25));
     TOBCmptTiers.AddChampSupValeur ('CTelex ',     Copy(ligne,775,25));
     TOBCmptTiers.AddChampSupValeur ('CRva  ',      Copy(ligne,800,50));
     TOBCmptTiers.AddChampSupValeur ('CCivilite ',  Copy(ligne,850,5));
     TOBCmptTiers.AddChampSupValeur ('CPpal  ',     Copy(ligne,853,1));
     TOBCmptTiers.AddChampSupValeur ('FormJur  ',   Copy(ligne,854,3));
     TOBCmptTiers.AddChampSupValeur ('RibPpal  ',   Copy(ligne,857,1));
     TOBCmptTiers.AddChampSupValeur ('TvaEncaissmnt',Copy(ligne,858,3));
end;

{***********A.G.L.***********************************************
Auteur  ...... : Marion Koszczuk
Créé le ...... : 28/05/2004
Modifié le ... :   /  /
Description .. : Recup de toutes les lignes du fichier dont le type n'est pas
Suite ........ : definie dans EnregAffCom
Mots clefs ... : RECUP LIGNE INCONNUES
*****************************************************************}
procedure TFControlFicU.GererLigneInconnue(ligne : string);
var TOBInconnue : TOB;
begin
     TOBInconnue := TOB.Create('',MTOBInconnue,-1);
     TOBInconnue.AddChampSupValeur ('Position ligne', IntToStr(numligne));
     TOBInconnue.AddChampSupValeur('Ligne',ligne);
end;

procedure TFControlFicU.RecupEtablsmnt(ligne :string);
var TOBEtab : TOB;
begin
     TOBEtab := TOB.Create('',MTOBEtab,-1);
     TOBEtab.AddChampSupValeur ('Position ligne',IntToStr(numligne));
     TOBEtab.AddChampSupValeur ('Identifiant',Copy(ligne,4,3));
     TOBEtab.AddChampSupValeur ('Code',Copy(ligne,7,3));
     TOBEtab.AddChampSupValeur ('Libelle',Copy(ligne,10,30));
end;

procedure TFControlFicU.RecupModPaie(ligne :string);
var TOBModPaie:TOB;
begin
    TobModPaie := TOB.Create('',MTOBModPaie,-1);
    TobModPaie.AddChampSupValeur('Position ligne',IntToStr(numligne));
    TobModPaie.AddChampSupValeur('Identifiant', Copy(ligne,4,3));
    TobModPaie.AddChampSupValeur('Code'       , Copy(ligne,7,3));
    TobModPaie.AddChampSupValeur('Libelle'    , Copy(ligne,10,35));
    TobModPaie.AddChampSupValeur('Categorie'  , Copy(ligne,45,3));
    TobModPaie.AddChampSupValeur('Code Acceptation' , Copy(ligne,48,3));
    TobModPaie.AddChampSupValeur('Edition lettre cheques' , Copy(ligne,51,1));
    TobModPaie.AddChampSupValeur('Edition lettre traite'  , Copy(ligne,52,1));
    TobModPaie.AddChampSupValeur('Condition montant'      , Copy(ligne,53,1));
    TobModPaie.AddChampSupValeur('Montant Max'            , Copy(ligne,54,20));
    TobModPaie.AddChampSupValeur('Mode de remplacement'   , Copy(ligne,74,3));

end;
{***********A.G.L.***********************************************
Auteur  ...... : Marion Koszczuk
Créé le ...... : 24/05/2004
Modifié le ... :   /  /
Description .. : parcour du fichier TRA en appelant pour chaque ligne la
Suite ........ : fonction GestLigne qui verifie l'éta du fichier (fin ,type de
Suite ........ : ligne trouvé ou non) et qui appel la recup de champs
Mots clefs ... : TRAITER LIGNE
*****************************************************************}
procedure TFControlFicU.TraiterLigne;
var i       : integer;
    ligne   : string;
begin
     inherited;
     numligne := 0;
     ProgressTraitement.Visible:=TRUE;
     ProgressTraitement.MaxValue := FileSize(Fichier) ;
     while(not EOF(Fichier)) do
     BEGIN
          Readln(Fichier,ligne);
          i := 0;
          numligne := numligne+1;
          while ( (i <= nbEnreg) and (EstType(ligne,EnregAffCom[i])=FALSE) ) do
             i := i+1;
          if( i<=nbEnreg ) then
          begin
             EnregAffCom[i].TOBCharge(ligne);
          end else
             GererLigneInconnue(ligne);
          ProgressTraitement.Progress := length(ligne);
     end;
     LanceVerif;
     ProgressTraitement.Visible:=False;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Marion Koszczuk
Créé le ...... : 24/05/2004
Modifié le ... :   /  /
Description .. : Bouton d'ouverture de fichier lorsque l'on ouvre un fichier
Suite ........ : TRA
Mots clefs ... : OUVERTURE FICHIER TRA
*****************************************************************}
procedure TFControlFicU.BOuvrirTRAClick(Sender: TObject);
begin
  inherited;
   if OpenDialog1.Execute then
   begin
        Combobloc.Clear;
        ComboPeriode.Clear;
        Combobloc.Visible:=false;
        ViderTOB;

        Edit_NomFichier.Text := OpenDialog1.FileName;
        ProgressTraitement.Visible:=False;
        FileStreamRecup(OpenDialog1.FileName);
        TV.Clear;
   end;
end;

function  TFControlFicU.TraiterTaille(Taille:variant):integer;
begin
     Taille := ConvertTaille(Taille);
     if (uniteTaille='Mo')then
        PComplement.Visible :=TRUE;
     Result := Taille;
end;

function  TFControlFicU.ConvertTaille(var Taille:variant):integer;
begin
     if Taille>1024 then //convert en ko
     begin
        Taille := (Taille/1024);
        uniteTaille := 'Ko';
        if Taille>1024 then //convert en mo
        begin
            Taille := (Taille/1024);
            uniteTaille := 'Mo';
        end;
     end else
         uniteTaille := 'octet(s)';
     Result := taille;
end;

{--------------------------------------------------------------------------
----------------AFFICHAGE LIGNE SELON SELECTION COMBO----------------------
---------------------------------------------------------------------------}
procedure TFControlFicU.FormClose(Sender: TObject; var Action: TCloseAction);
var  i : integer;
begin
  inherited;
  if ComboBloc.visible=TRUE then
  begin
       //  PGI('Conserver les blocs?',mtConfirmation,[mbYes,mbNo],0);
         if MessageDlg('Conserver les blocs du fichiers?',mtConfirmation, [mbYes, mbNo], 0) = mrNo then
         begin
              with Combobloc do
              begin
                   for i:=0 to Items.Count-1 do
                   begin
                        ItemIndex:=i;
                        DeleteFile(ComboBloc.Values[i]);
                   end;
              end;
              PGIBox('Blocs de fichiers supprimés');
         end;
  end;
  LibererTOB;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Marion Koszczuk
Créé le ...... : 24/05/2004
Modifié le ... :   /  /
Description .. : Sur changement de selection dans le combo du type de
Suite ........ : ligne à afficher
Mots clefs ... : TOB A AFFICHER, SELECTION COMBO
*****************************************************************}
procedure TFControlFicU.CmbTypeLigneChange(Sender: TObject);
begin
  inherited;
   if RGrpTypeAffich.ItemIndex=0 then//toute les lignes
   begin
      LATOB := EnregAffCom[CmbTypeLigne.ItemIndex].MTOBTout;
   end else
   begin
      if RGrpTypeAffich.ItemIndex=1 THEN  //lignes avec erreurs
      begin
         LATOB := EnregAffCom[CmbTypeLigne.ItemIndex].FTOBErreur ;
      end else
      begin
          if RGrpTypeAffich.ItemIndex=2 THEN //lignes sans erreurs
             LATOB := EnregAffCom[CmbTypeLigne.ItemIndex].MTOB;
      end;
   end;
   if ((LATOB=nil) or (LATOB.Detail.Count<=0)) then
      PGIInfo('Aucune ligne à afficher pour cette sélection');
   ColNames := EnregAffCom[CmbTypeLigne.ItemIndex].ColName ;
   ModeAlimentation := Stat.maTOB;
   BChercheClick(Sender);
end;

PROCEDURE TFControlFicU.LanceVerif;
var i       : integer;
    Control : TVerif;
begin
     Screen.Cursor:=crHourGlass;
     ProgressTraitement.Visible:=True;
     ProgressTraitement.MaxValue := nbEnreg;

     Control:=TVerif.Create;
     for i := 0 to nbEnreg do
         Control.ControlTRA(EnregAffCom,i);
     for i := 0 to nbEnreg do
       Control.RecupLigneErreur(EnregAffCom[i]);

     ProgressTraitement.Progress := nbEnreg;
     Screen.Cursor:=crDefault;
     Control.Free;
     ProgressTraitement.Visible:=False;
end;

procedure TFControlFicU.RGrpTypeAffichClick(Sender: TObject);
begin
  inherited;
  if(CmbTypeLigne.Text<>'')then
  begin
       if RGrpTypeAffich.ItemIndex=0 then  //toutes les lignes
       begin
           LATOB := EnregAffCom[CmbTypeLigne.ItemIndex].MTOBTout;
       end else
       begin
             if RGrpTypeAffich.ItemIndex=1 then  //lignes avec erreur
             begin
                 LATOB := EnregAffCom[CmbTypeLigne.ItemIndex].FTOBErreur ;
             end else
             begin
                  if RGrpTypeAffich.ItemIndex=2 then  //lignes avec erreur
                     LATOB := EnregAffCom[CmbTypeLigne.ItemIndex].MTOB ;
             end;
       end;
       ColNames := EnregAffCom[CmbTypeLigne.ItemIndex].ColName ;
       ModeAlimentation := Stat.maTOB;
       BChercheClick(Sender);
  end;
End;

procedure TFControlFicU.BValidDecoupClick(Sender: TObject);
begin
  inherited;
    if (ComboModDecoup.ItemIndex =0) then
    begin
       Screen.Cursor:=crHourGlass;
        appelDecoup(Edit_NomFichier.Text,'JAL');
        Screen.Cursor:=crDefault;
    end else
    begin
         if ((ComboModDecoup.ItemIndex =1) and (ComboPeriode.ItemIndex<>-1)) then
         begin
             Screen.Cursor:=crHourGlass;
             appelDecoup(Edit_NomFichier.Text,'EXO');
             Screen.Cursor:=crDefault;
         end else
         begin
              if ((ComboModDecoup.ItemIndex =1) and (ComboPeriode.ItemIndex<0)) then
              begin
                 PGIBox('Selectionner une periode de découpage');
                 exit;
              end;
         end;
    end;
    BChercheClick(Sender);
    PGIINFO('Fin de decoupage');
end;

procedure TFControlFicU.appelDecoup(adresse : string;Decoup:string);
var Decoupe : TDecoupTRA;
    TypeDecoup : TTypeDecoup;
begin
    Decoupe := TDecoupTra.Create;
    if Decoup = 'JAL' then
    begin
         TypeDecoup := JOURNAUX ;
         Decoupe.Decouper(adresse,TypeDecoup,'');
    end else
    begin
        if Decoup = 'EXO'then
        begin
           TypeDecoup := PERIODE ;
           Decoupe.Decouper(adresse,TypeDecoup,ComboPeriode.Text);
        end;
    end;
    Decoupe.Free;
end;

procedure TFControlFicU.ComboModDecoupChange(Sender: TObject);
begin
  inherited;
  if ComboBloc.Visible=TRUE then
  begin
       BValidDecoup.Enabled := TRUE;
       ComboPeriode.Visible := FALSE;
       LabelPeriode.Visible := FALSE;
       PGIINFO('Impossible de découper un fichier lorsqu''il est en plusieurs blocs.');
  end else
  begin
       if ((Edit_NomFichier.Text <>'') and (ComboModDecoup.ItemIndex<>-1)) then
       begin
            if ComboModDecoup.ItemIndex=1 then
            begin
                 ComboPeriode.Visible := TRUE;
                 LabelPeriode.Visible := TRUE;
                 BValidDecoup.Enabled := TRUE;
            End else
            begin
                 BValidDecoup.Enabled := TRUE;
                 ComboPeriode.Visible := FALSE;
                 LabelPeriode.Visible := FALSE;
            end;
       End;
  End;
End;

procedure TFControlFicU.RemplirComboPeriode(MTOBTouExercice,MTOBExercice:TOB);
var i         : integer;
    TOBTmp,TD : TOB;
    periode : string;
begin
    TOBTmp := nil;
    if ExistCombo('Automatique')=FALSE then
        ComboPeriode.Items.Add ('Automatique');

    i := RecherchTOB(EXO);
    if i<>-1 then
        TOBTmp := EnregAffCom[i].MTOBTout;
    if((Assigned(TOBTmp)) and (TobTmp.Detail.Count>0)) then
    begin
          i := 0;
          while i<=TOBTmp.Detail.Count-1 do
          begin
               TD := TOBTmp.Detail[i];
               if ((Trim(Copy(TD.GetValue('DateDebut'),1,1))<>'') and (Trim(Copy(TD.GetValue('DateFin'),1,1))<>''))then
               begin
                   periode := Copy(TD.GetValue('DateDebut'),5,4)+'/'+Copy(TD.GetValue('DateFin'),5,4);
                   IF  (ExistCombo(periode)=False) then
                     ComboPeriode.Items.Add(periode);
               end;
               i := i+1;
          end;
    end;
end;

procedure TFControlFicU.RemplirComboNumCompt(MTOBTouCmptGnrx,MTOBCmptGnrx:TOB);
var i         : integer;
    TOBTmp,TD : TOB;
begin
    TOBTmp := nil;
    i := RecherchTOB(CGN);
    if i<>-1 then
        TOBTmp := EnregAffCom[i].MTOBTout;

    if((Assigned(TOBTmp)) and (TobTmp.Detail.Count>0)) then
    begin
          i := 0;
          while i<=TOBTmp.Detail.Count-1 do
          begin
               TD := TOBTmp.Detail[i];
               if Trim(Copy(TD.GetValue('Code'),1,1))<>''then
               begin
                   IF  ExistCombo(TD.GetValue('Code'))=False then
                     ComboNumCompt.Items.Add(TD.GetValue('Code'));
               end;
               i := i+1;
          end;
    end;
end;

function TFControlFicU.ExistCombo(valeur:string):boolean;
var i        : integer;
    vraifaux : boolean;
begin
    VRAIFAUX := FALSE;
    with ComboPeriode do
    begin
         for i:=0 to Items.Count do
         begin
              if ComboPeriode.Items[i]=valeur then
                 vraifaux := TRUE;
         end;
    end;
    Result := VRAIfaux
end;

function  TFControlFicU.RecherchTOB(TypeLigne:TTypeligne):integer;
var i : integer;
begin
    i := 0;
    while( (i<=Length(EnregAffCom))and (EnregAffCom[i].Verif<>Typeligne)) do
    begin
         i := i+1;
    end;
    if i<=Length(EnregAffCom) then
    begin
      Result := i;
    end else
      Result := -1;
end;

procedure TFControlFicU.ComboPeriodeChange(Sender: TObject);
begin
  inherited;
  if ((ComboPeriode.ItemIndex<>-1) and (ComboPeriode.Text<>'')) then
     BValidDecoup.Enabled := TRUE;
end;

procedure TFControlFicU.ComboTypeAffichageChange(Sender: TObject);
begin
  inherited;
  if (ComboTypeAffichage.Text='Par type de lignes') or (ComboTypeAffichage.ItemIndex=0) then
  begin
       ComboChoixLettrage.Visible :=FALSE;
       ComboNumCompt.Visible :=FALSE;
       LabelTypeAff.Caption := 'Type de Ligne:';
       CmbTypeLigne.Visible := TRUE;
       RGrpTypeAffich.Visible := TRUE;
  end else
  begin
      if (ComboTypeAffichage.Text='Par type de lettrage') or (ComboTypeAffichage.ItemIndex=1) then
      begin
           CmbTypeLigne.Visible := FALSE;
           ComboNumCompt.Visible :=FALSE;
           LabelTypeAff.Caption := 'Type de Lettrage:';
           ComboChoixLettrage.Top:=32;
           ComboChoixLettrage.Visible :=TRUE;
           RGrpTypeAffich.Visible := FALSE;
      end else
      begin
           CmbTypeLigne.Visible := FALSE;
           LabelTypeAff.Caption := 'Numero de compte:';
           ComboNumCompt.Visible :=TRUE;
           RGrpTypeAffich.Visible := FALSE;
      end;
   end;
end;

procedure TFControlFicU.AffTypLettrable(i:integer;TD:TOB;EtatLettrage:string);
var TOBFilleLettrag,TOBLettrage : TOB;
    trouv,j : integer;
begin
     trouv := 0;
     if ((Assigned(TD)) and (TD.Detail.Count>0)) then
     begin
         TOBLettrage := TOB.Create('TOBMere',nil,-1);
         j := 0;
         while j<=TD.Detail.Count-1 do
         begin
            if TD.Detail[j].GetValue('EtaLettrag')=EtatLettrage then
            begin
                TOBFilleLettrag := TOB.Create('',TOBLettrage,-1);
                TOBFilleLettrag.Dupliquer(TD.Detail[j],TRUE,TRUE,FALSE);
                trouv := 1;
             end;
             j := j+1;
          end;
          if (trouv=0) then
          begin
              TOBFilleLettrag := TOB.Create('',TOBLettrage,-1);
              TOBFilleLettrag.AddChampSupValeur('Resultat','Aucunes lignes trouvées avec ce type de lettrage');
              ColNames := 'Resultat; ' ;
          end else
          begin
               if TOBLettrage<>nil then
                  ColNames := EnregAffCom[i].ColName;
          end;
          if ((LATOB=nil) or (LATOB.Detail.Count<=0)) then
             PGIInfo('Aucune ligne à afficher pour cette sélection');
          LATOB := TOBLettrage;
          ModeAlimentation := Stat.maTOB;
     end;
end;

procedure TFControlFicU.AffNumComptTypeLettrable(i:integer;TD:TOB;EtatLettrage:string);
var trouv                             : integer;
    TOBTmp,TOBNumCompte,TOBFilleNum   : TOB;
begin
     trouv := 0;
     if ((Assigned(TD)) and (TD.Detail.Count>0)) then
     begin
         TOBTmp := TD.FindFirst(['General'],[ComboNumCompt.Text],TRUE);
         TOBNumCompte := TOB.Create('TOBMere',nil,-1);
         while TOBTmp<>nil do
         begin
             if(TOBTmp.GetValue('EtaLettrag')=EtatLettrage) then
             begin
                  TOBFilleNum := TOB.Create('',TOBNumCompte,-1);
                  TOBFilleNum.Dupliquer(TOBTmp,TRUE,TRUE,FALSE);
                  trouv := 1;
             end;
             TOBTmp := TD.FindNext(['General'],[ComboNumCompt.Text],FALSE);
          end;
          if (trouv=0) then
          begin
              TOBFilleNum := TOB.Create('',TOBNumCompte,-1);
              TOBFilleNum.AddChampSupValeur('Resultat','Aucunes lignes trouvées avec ce type de lettrage');
              ColNames := 'Resultat; ' ;
          end else
          begin
               if TOBNumCompte<>nil then
                  ColNames := EnregAffCom[i].ColName;
          end;
          if ((LATOB=nil) or (LATOB.Detail.Count<=0)) then
             PGIInfo('Aucune ligne à afficher pour cette sélection');
          LATOB := TOBNumCompte;
          ModeAlimentation := Stat.maTOB;
     end;
end;

procedure TFControlFicU.RechercheLettrage(EtatLettrage:string;numcompt:string);
var TD  : TOB;
    i   : integer;
begin
     i := RecherchTOB(ECR);
     if i>-1 then
     begin
         TD := EnregAffCom[i].MTOBTout ;
         if ComboNumCompt.Visible=FALSE THEN
         begin
             Screen.Cursor:=crHourGlass;
             AffTypLettrable(i,TD,EtatLettrage);
             Screen.Cursor:=crDefault;
         end else
         begin
             Screen.Cursor:=crHourGlass;
             AffNumComptTypeLettrable(i,TD,EtatLettrage);
             Screen.Cursor:=crDefault;
         end;
     end;
end;

procedure TFControlFicU.ComboChoixLettrageChange(Sender: TObject);
var numcompt : string;
begin
  inherited;
  if ComboNumCompt.Text='' then
  begin
      numcompt := '';
  end else
      numcompt := ComboNumCompt.Text;

  if ((ComboChoixLettrage.Text='Totalement lettrés') or (ComboChoixLettrage.ItemIndex =0)) then
  begin
      RechercheLettrage('TL',numcompt);
  end else
  begin
       if ((ComboChoixLettrage.Text='Partiellement lettrés') or (ComboChoixLettrage.ItemIndex =1)) then
       begin
           RechercheLettrage('PL',numcompt);
       end else
       begin
            if ((ComboChoixLettrage.Text='Non Lettrable') or (ComboChoixLettrage.ItemIndex =2)) then
            begin
                 RechercheLettrage('RI',numcompt);
            end else
            begin
                 if((ComboChoixLettrage.Text='A lettrer') or (ComboChoixLettrage.ItemIndex =3)) then
                    RechercheLettrage('AL',numcompt);
            end;
       end;
  end;
  BChercheClick(Sender);
end;

procedure TFControlFicU.ComboNumComptChange(Sender: TObject);
begin
  inherited;
    ComboChoixLettrage.Top:=56;
    ComboChoixLettrage.Visible :=TRUE;
end;

procedure TFControlFicU.B_DecoupAutoClick(Sender: TObject);
var  Decoupe : TDecoupTRA;
begin
  inherited;
   Decoupe := TDecoupTRA.Create;
   Decoupe.DecoupAuto(Edit_NomFichier.Text);
   Decoupe.Free;
end;

procedure TFControlFicU.ComboBlocChange(Sender: TObject);
begin
  inherited;
  ShowMessage(Combobloc.Value);
  FileStreamRecup(Combobloc.Value);
end;

end.


