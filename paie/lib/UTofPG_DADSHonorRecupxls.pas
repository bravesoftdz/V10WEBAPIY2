{***********UNITE*************************************************
Auteur  ...... : VG
Créé le ...... : 22/10/2002
Modifié le ... :
Description .. : Récupération d'honoraires à partir d'une feuille EXCEL
Mots clefs ... : PAIE,PGDADSU
*****************************************************************}
{
PT1   : 29/01/2003 VG V591 Affichage d'un message si l'exercice par défaut
                           n'existe pas - FQ N°10469
PT2   : 31/01/2003 VG V591 Arrondi des montants à l'euro le plus proche                           
PT3   : 12/02/2003 VG V_42 FQ N°10469 bis
PT4   : 14/04/2003 VG V_42 Modification de la tablette PGADRESSECOMPL
PT5   : 16/09/2003 VG V_42 Adaptation cahier des charges V7R01
PT6   : 12/02/2004 VG V_50 Sur les honoraires, le cahier des charges interdit
                           d'avoir Raison sociale et Nom/Prénom renseignés
                           simultanément - FQ N°11107
PT7-1 : 12/05/2004 VG V_50 Rendre la saisie du SIRET obligatoire si Raison
                           sociale présente - FQ N°11108
PT7-2 : 12/05/2004 VG V_50 Ajout du traitement lié aux raccourcis clavier
PT8   : 05/07/2004 VG V_50 Adaptation cahier des charges V8R00
PT9   : 24/08/2004 VG V_50 Ajout d'une ligne dans le fichier des honoraires afin
                           d'afficher le millésime du fichier
PT10  : 10/01/2005 VG V_60 Erreur de format : JJMAAAA au lieu de JJMMAAAA si
                           mois de clôture antérieur à octobre - FQ N°11913
PT11  : 07/10/2005 VG V_60 Adaptation cahier des charges DADS-U V8R02
PT12  : 23/11/2005 VG V_65 Modification du traitement pour pouvoir récupérer
                           plusieurs types de "modalités de prise en charge des
                           indemnités"
PT13-1: 16/10/2006 VG V_70 Adaptation cahier des charges DADS-U V8R04
PT13-2: 16/10/2006 VG V_70 Valeur zéro non acceptée dans segment S70.G01.00.013
                           FQ N°12853
PT13-3: 16/10/2006 VG V_70 Suppression du fichier de contrôle - mise en table
                           des erreurs
PT13-4: 16/10/2006 VG V_70 Utilisation d'un type pour la cle DADS-U
PT14  : 25/10/2006 VG V_70 Mise en base des fichiers - FQ N°13441
                           17/11/2006 Mise en place cohabitation des 2 méthodes
                           (fichiers en base ou sur répertoire)
PT15  : 21/11/2006 VG V_70 Permettre de déclarer des honoraires en DADS-U
                           complémentaire - FQ N°13613
PT16  : 14/12/2006 VG V_70 Import incorrect de bis ou ter - FQ N°13725
PT17-1: 17/08/2007 VG V_80 Adaptation cahier des charges V8R05 - FQ N°14424
PT18  : 10/12/2007 VG V_80 Montant à zéro interdit - FQ N°13869
PT19  : 12/12/2007 VG V_80 Ajout de la date de fin d'exercice comptable dans le
                           fichier - FQ N°13894
PT20  : 23/01/2008 VG V_80 Pas de montant à blanc - FQ N°13869                           
}
unit UTofPG_DADSHonorRecupxls;

interface
uses  UTOF,HTB97,HMsgBox,Controls,HCtrls,HEnt1,UTOB,sysutils,UTOBXLS,StdCtrls,
      ed_tools,Classes,Pgoutils2,PgDADSCommun,PgDADSOutils,comctrls,windows,
      Vierge,UyFileSTD
{$IFNDEF EAGLCLIENT}
      ,QRGrid{$IFNDEF DBXPRESS} ,dbTables {$ELSE} ,uDbxDataSet {$ENDIF}
{$ELSE}
{$ENDIF}
      ;
Type
     TOF_PG_DADSHonorRecupxls = Class (TOF)
       private

       LaGrille  : THGrid;
       NbreCol   : Integer;
       BtnRecop, BtnValid : TToolbarButton97;
       THAnnee:THValCombobox;
       TraceErr : TListBox;
       FileN : String;

       function  OnSauve: boolean;
       procedure ValiderClick(Sender: TObject);
       procedure GrilleCopierColler (Sender: TObject);
       procedure ActiveBtn (Sender: TObject);
       procedure ImpClik (Sender: TObject);
       procedure LigneGrilleToBase(Ligne, OrdreDest : integer);
       procedure ExerciceChange(Sender: TObject);
       procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);

       public
       procedure OnArgument(Arguments : String ) ; override ;
       procedure OnClose; override ;
     END ;

implementation

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 22/10/2002
Modifié le ... :   /  /
Description .. : OnArgument
Mots clefs ... : PAIE,PGDADSU
*****************************************************************}
procedure TOF_PG_DADSHonorRecupxls.OnArgument(Arguments: String);
var
BtnImp : TToolbarButton97;
JourJ : TDateTime;
AnneeA, Jour, MoisM : Word;
AnneeCours, AnneeE, ComboExer, MoisE,StPcl,Mes : string;
TrouveCours : boolean;
begin
inherited ;
TFVierge (Ecran).OnKeyDown:= FormKeyDown;

BtnImp    := TToolbarButton97 (GetControl ('Bimprimer'));
If BtnImp <> NIL then
   BtnImp.OnClick := ImpClik;
BtnRecop:=TToolbarButton97 (GetControl ('RECOPIER'));
if BtnRecop<>NIL then
   BtnRecop.OnClick := ActiveBtn ;
BtnValid:=TToolbarButton97 (GetControl ('BVALIDER'));
if BtnValid<>NIL then
   begin
   BtnValid.OnClick := ValiderClick ;
   BtnValid.Enabled := False;
   end;
LaGrille := THGrid(Getcontrol('GRILLE'));
if LaGrille <> nil then
   begin
   LaGrille.ColWidths[0]:= 296;
   LaGrille.ColWidths[1]:= 296;
   LaGrille.ColWidths[2]:= 116;
   LaGrille.ColWidths[3]:= 100;
   LaGrille.ColWidths[4]:= 176;
   LaGrille.ColWidths[5]:= 76;
   LaGrille.ColWidths[6]:= 66;
   LaGrille.ColWidths[7]:= 145;
   LaGrille.ColWidths[8]:= 140;
   LaGrille.ColWidths[9]:= 55;
   LaGrille.ColWidths[10]:= 140;
   LaGrille.ColWidths[11]:= 111;
   LaGrille.ColWidths[12]:= 160;
   LaGrille.ColWidths[13]:= 88;
   LaGrille.ColWidths[14]:= 74;
   LaGrille.ColWidths[15]:= 59;
   LaGrille.ColWidths[16]:= 58;
   LaGrille.ColWidths[17]:= 58;
   LaGrille.ColWidths[18]:= 81;
   LaGrille.ColWidths[19]:= 123;
   LaGrille.ColWidths[20]:= 148;
   LaGrille.ColWidths[21]:= 80;
   LaGrille.ColWidths[22]:= 82;
   LaGrille.ColWidths[23]:= 66;
   LaGrille.ColWidths[24]:= 82;
   LaGrille.ColWidths[25]:= 66;
   LaGrille.ColWidths[26]:= 82;
   LaGrille.ColWidths[27]:= 66;
   LaGrille.ColWidths[28]:= 82;
   LaGrille.ColWidths[29]:= 66;
   LaGrille.ColWidths[30]:= 82;
   LaGrille.ColWidths[31]:= 66;
   LaGrille.ColWidths[32]:= 66; //PT19
   end;

THAnnee:=THValCombobox(GetControl('EXERCICE'));
JourJ := Date;
DecodeDate(JourJ, AnneeA, MoisM, Jour);
if MoisM>9 then
   AnneeCours := IntToStr(AnneeA)
else
   AnneeCours := IntToStr(AnneeA-1);

TrouveCours:= RendExerSocialPrec (MoisE, AnneeE, ComboExer, DebExer, FinExer,
                                  AnneeCours);
if (TrouveCours = FALSE) then
   begin
   PGIBox ('L''exercice '+AnneeCours+' n''existe pas', Ecran.Caption);
   if (PGAnnee = '') then
      BtnRecop.Enabled:= False;
   end;

if THAnnee <> NIL then
   begin
   THAnnee.Value:=ComboExer;
   PGAnnee := THAnnee.value;
   THAnnee.OnChange := ExerciceChange;
   PGExercice:= AnneeE;               //PT13-4
   end;

//DEBUT PT14
    If (V_PGI.ModePcl = '1') then
       StPcl:= ' AND NOT (YFS_PREDEFINI="DOS" AND YFS_NODOSSIER<>"'+V_PGI.NoDossier+'")'
    else
       StPcl:= '';
//Recherche fichier <> CEG ou modèle par défaut pour savoir si client utilise fichier en base
    If not existeSQL ('SELECT YFS_NOM'+
                      ' FROM YFILESTD WHERE'+
                      ' YFS_CODEPRODUIT="PAIE" AND'+
                      ' YFS_CRIT1="XLS" AND'+
                      ' YFS_CRIT2="HON" AND'+
                      ' YFS_PREDEFINI<>"CEG"'+StPcl) then
       begin
       Mes:= 'Cette version apporte un nouveau fonctionnement dans#13#10'+
             'la gestion du stockage des fichiers.#13#10#13#10'+
             'Par défaut l''utilisation reste identique.#13#10#13#10'+
             'Toutefois nous vous invitons à mettre en place cette nouvelle fonctionnalité';
       SetControlChecked ('SANSFICHIERBASE', true);
       HShowMessage ('1;Gestion des fichiers en base ;'+Mes+';E;HO;;;;;41910025', '', '');
       SetControlChecked ('CREPERTOIRE', True);
       SetControlEnabled ('DOCUMENT', False);
       end;
//FIN PT14
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 22/10/2002
Modifié le ... :   /  /
Description .. : Fonction de validation de la saisie
Mots clefs ... : PAIE,PGDADSU
*****************************************************************}
procedure TOF_PG_DADSHonorRecupxls.ValiderClick(Sender: TObject);
var
rep : Integer;
begin
inherited ;
Rep:=PGIAsk ('Voulez vous sauvegarder votre saisie ?', Ecran.Caption) ;
if rep=mrNo then
   exit ;
if rep=mrCancel then
   exit;
if rep=mryes then
   OnSauve;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 22/10/2002
Modifié le ... :   /  /
Description .. : fonction d'ecriture des elements recuperer de la feuille excel
Suite ........ : dans la grille de Saisie
Suite ........ : Il s'agit en fait de construire la tob qui mettra à jour la table
Mots clefs ... : PAIE,PGDADSU
*****************************************************************}
function TOF_PG_DADSHonorRecupxls.OnSauve: boolean;
Var
StHonor, StSelect : String;
QHonor:TQuery;
i, OrdreDest : integer;
Pages : TPageControl;
begin
result := TRUE;
StHonor := '--H';

OrdreDest:=1;
StSelect := 'SELECT MAX(PDE_ORDRE) AS NUMMAX'+
            ' FROM DADSPERIODES WHERE'+
            ' PDE_SALARIE LIKE "'+StHonor+'%" AND'+
            ' PDE_DATEDEBUT >= "'+UsDateTime(DebExer)+'" AND'+
            ' PDE_DATEFIN <= "'+UsDateTime(FinExer)+'"';

QHonor:=OpenSQL(StSelect,True);
if NOT QHonor.EOF then
   try
   OrdreDest:= (QHonor.FindField('NUMMAX').AsInteger)+1;
   except
         on E: EConvertError do
            OrdreDest:= 1;
   end;
Ferme(QHonor);

try
   begintrans;
   TraceErr.Items.Add('Début d''import : '+TimeToStr(Now));
   InitMoveProgressForm (NIL,'Import en cours',
                                'Veuillez patienter SVP ...',LaGrille.RowCount,
                                FALSE,TRUE);
   ChargeTOBDADS;
   DeleteErreur ('', 'SKO');	//PT13-3
   for i := 1 to LaGrille.RowCount-1 do
       begin
       LigneGrilleToBase(i, OrdreDest);
       OrdreDest := OrdreDest + 1;
       end;
   EcrireErreurKO;	//PT13-3
   LibereTOBDADS;
   TraceErr.Items.Add('Fin d''import : '+TimeToStr(Now));
   CommitTrans;
Except
   result := FALSE;
   Rollback;
   TraceErr.Items.Add('Import annulé : '+TimeToStr(Now));
   PGIBox ('Une erreur est survenue lors de la mise à jour de la base', 'DADS-U');
   END;

Pages := TPageControl(GetControl('PGCTRL1'));
if Pages<>nil then
   Pages.ActivePageIndex:=1;
FiniMoveProgressForm;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 22/10/2002
Modifié le ... :   /  /
Description .. : OnClose
Mots clefs ... : PAIE,PGDADSU
*****************************************************************}
procedure TOF_PG_DADSHonorRecupxls.OnClose;
begin
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 22/10/2002
Modifié le ... :   /  /
Description .. : Cette fonction recupère les données de la feuille EXCEL
Suite ........ : dans une TOB
Suite ........ : Elle alimente la grille pour voir les données qui seront
Suite ........ : recupérées.
Suite ........ : Les salariés inexistants ne sont pas pris et idem pour les
Suite ........ : colonnes.
Suite ........ : Les entetes de colonne sont acceptées en fonction de
Suite ........ : l'existance
Suite ........ : dans la table paramètrage de l'import.
Mots clefs ... : PAIE,PGDADSU
*****************************************************************}
procedure TOF_PG_DADSHonorRecupxls.GrilleCopierColler (Sender: TObject);
var
T,TT : TOB;
NomF, Predef, St : string;
CodeRetour, i,j, Long : integer;
begin
if LaGrille = NIL then
   begin
   PgiBox ('Attention, grille non identifiée',Ecran.Caption);
   exit;
   end;

//DEBUT PT14
If (GetCheckBoxState ('CREPERTOIRE')=CbChecked) then
   FileN:= GetControlText ('FICHIER')
else
   begin
   NomF:= GetControlText ('DOCUMENT');
   Long:= Length (NomF);
   Predef:= Copy (NomF, 0, 3);
   NomF:= Copy (NomF,4,Long);
   CodeRetour:= AGL_YFILESTD_EXTRACT (FileN, 'PAIE', NomF , 'XLS', 'HON', '', '',
                                      '', False, 'FRA', Predef);
   end;
//FIN PT14

if FileN = '' then
   begin
   PgiBox ('Attention, vous n''avez pas sélectionné de fichier EXCEL',Ecran.Caption);
   exit;
   end;

if NOT FileExists(FileN) then
   begin
   PgiBox ('Attention, votre fichier EXCEL n''existe pas',Ecran.Caption);
   exit;
   end;

T := TOB.Create('Ma tob',Nil, -1);
SourisSablier;
try
   ImportTOBFromXLS(T, FileN, True, False, 'HONORAIRES', '', nil, False);
Except
   END;

SourisNormale;

TraceErr:= TListBox (GetControl ('LSTBXERROR'));
if TraceErr = NIL then
   begin
   PgiBox ('Attention, composant trace non trouvé, abandon du traitement',Ecran.Caption);
   exit;
   end;
TraceErr.Refresh ;
if (T.Detail.Count > 2) then
   begin
// determination du nombre de ligne de la grille
   LaGrille.RowCount := T.Detail.Count-1;
// récupération du nombre de champs saisis dans la feuille EXCEL
   InitMoveProgressForm (NIL,'Traitement en cours',
                                'Veuillez patienter SVP ...',T.Detail.Count-1,
                                FALSE,TRUE);
   NbreCol :=  T.Detail[1].ChampsSup.Count;
   LaGrille.ColCount := NbreCol;
   for j := 0 to T.Detail[1].ChampsSup.Count-1 do
       begin
       st :=  TCS(T.Detail[1].Champssup[j]).Nom;
       if st<>LaGrille.Cells[j,0] then
          begin
          FiniMoveProgressForm;
          PgiBox ('Attention, Le fichier EXCEL n''est pas aux normes',Ecran.Caption);
          exit;
          end;
       end;
   for i := 2 to T.Detail.Count-1 do
       begin
       TT := T.Detail[i];
       for j := 0 to T.Detail[1].ChampsSup.Count-1 do
           LaGrille.Cells[j,i-1] := TT.GetValue(LaGrille.Cells[j,0]);
       end;
   end;

if BtnValid<>NIL then
   BtnValid.Enabled := True;
T.Free;
FiniMoveProgressForm;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 22/10/2002
Modifié le ... :   /  /
Description .. : ActiveBtn
Mots clefs ... : PAIE,PGDADSU
*****************************************************************}
procedure TOF_PG_DADSHonorRecupxls.ActiveBtn(Sender: TObject);
begin
GrilleCopierColler (Sender);
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 22/10/2002
Modifié le ... :   /  /
Description .. : Impression
Mots clefs ... : PAIE,PGDADSU
*****************************************************************}
procedure TOF_PG_DADSHonorRecupxls.ImpClik (Sender: TObject);
begin
{$IFNDEF EAGLCLIENT}
if LaGrille <> NIL then
   QRGrid.PrintGrid ( [LaGrille] , Ecran.Caption) ;
{$ENDIF}
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 22/10/2002
Modifié le ... :   /  /
Description .. : Enregistrement dans la base
Mots clefs ... : PAIE,PGDADSU
*****************************************************************}
procedure TOF_PG_DADSHonorRecupxls.LigneGrilleToBase(Ligne, OrdreDest : integer);
var
BufDest, BufOrig, CodeIso, Libelle, Nic, SDateCloture, SDateCloture4 : string;
Siren, StHonor : string;
QEtab : TQuery;
annee, jour, mois : word;
i : integer;
CleDADS : TCleDADS;
DateCloture : TDateTime;
begin
StHonor := '--H';
{PT15
StType := '001';
}
if (GetControlText ('CHCOMPL')='X') then
   TypeD:= '201'
else
   TypeD:= '001';
//FIN PT15

{PT13-4
CreeEntete (StHonor, StType, OrdreDest, DebExer, FinExer, '', '');
}
CleDADS.Salarie:= StHonor+IntToStr(OrdreDest);
{PT15
CleDADS.TypeD:= StType;
}
CleDADS.TypeD:= TypeD;
//FIN PT15
CleDADS.Num:= OrdreDest;
CleDADS.DateDeb:= DebExer;
CleDADS.DateFin:= FinExer;
CleDADS.Exercice:= PGExercice;
CreeEntete (CleDADS, '', '', '01');
//FIN PT13-4

if (LaGrille.Cells[0,Ligne] <> '') then
   begin
   CreeDetail (CleDADS, 1, 'S70.G01.00.001', LaGrille.Cells[0,Ligne],
               LaGrille.Cells[0,Ligne]);
   if (LaGrille.Cells[1,Ligne] <> '') then
      begin
      TraceErr.Items.Add('Pour l''honoraire N°'+IntToStr(OrdreDest)+', la '+
                         'raison sociale du bénéficiaire et le nom du '+
                         'bénéficiaire ne peuvent pas être renseignés '+
                         'simultanément. La raison sociale étant renseignée, '+
                         'le nom et le prénom sont effacés.');
      LaGrille.Cells[1,Ligne]:='';
      LaGrille.Cells[2,Ligne]:='';
      end;
   end
else
   begin
   if (LaGrille.Cells[1,Ligne] <> '') then
      CreeDetail (CleDADS, 2, 'S70.G01.00.002.001', LaGrille.Cells[1,Ligne],
                  LaGrille.Cells[1,Ligne]);

   if (LaGrille.Cells[2,Ligne] <> '') then
      CreeDetail (CleDADS, 3, 'S70.G01.00.002.002', LaGrille.Cells[2,Ligne],
                  LaGrille.Cells[2,Ligne]);
   if ((LaGrille.Cells[1,Ligne] <> '') and (LaGrille.Cells[2,Ligne] = '')) then
      TraceErr.Items.Add('Pour l''honoraire N°'+IntToStr(OrdreDest)+', le '+
                         'prénom est obligatoire.');
   end;

if ((LaGrille.Cells[0,Ligne] = '') and (LaGrille.Cells[1,Ligne] = '')) then
   TraceErr.Items.Add('La raison sociale ou le nom de la personne physique de'+
                      ' l''honoraire N°'+IntToStr(OrdreDest)+' n''est pas'+
                      ' renseigné');


if (LaGrille.Cells[3,Ligne] <> '') then
   begin
   Siren := Copy(LaGrille.Cells[3,Ligne], 1, 9);
   Nic := Copy(LaGrille.Cells[3,Ligne], 10, 5);
   CreeDetail (CleDADS, 4, 'S70.G01.00.003.001', Siren, LaGrille.Cells[3,Ligne]);

   CreeDetail (CleDADS, 5, 'S70.G01.00.003.002', Nic, LaGrille.Cells[3,Ligne]);
   end
else
   if (LaGrille.Cells[0,Ligne] <> '') then
      TraceErr.Items.Add('Le SIRET de l''honoraire N°'+IntToStr(OrdreDest)+
                         ' n''est pas renseigné');

if (LaGrille.Cells[4,Ligne] <> '') then
   CreeDetail (CleDADS, 6, 'S70.G01.00.004.001', LaGrille.Cells[4,Ligne],
               LaGrille.Cells[4,Ligne]);

if (LaGrille.Cells[5,Ligne] <> '') then
   CreeDetail (CleDADS, 7, 'S70.G01.00.004.003', LaGrille.Cells[5,Ligne],
               LaGrille.Cells[5,Ligne]);

if (LaGrille.Cells[6,Ligne] <> '') then
{PT16
   CreeDetail (CleDADS, 8, 'S70.G01.00.004.004',
               RechDom('PGADRESSECOMPL', LaGrille.Cells[6,Ligne], True),
               LaGrille.Cells[6,Ligne]);
}
   begin
   BufOrig:= LaGrille.Cells[6,Ligne];
   if (BufOrig = 'B') then
      BufDest:= '2'
   else
   if (BufOrig = 'T') then
      BufDest:= '3'
   else
   if (BufOrig = 'Q') then
      BufDest:= '4'
   else
   if (BufOrig = 'C') then
      BufDest:= '5';
   CreeDetail (CleDADS, 8, 'S70.G01.00.004.004', BufOrig, BufDest);
   end;
//FIN PT16

if (LaGrille.Cells[7,Ligne] <> '') then
   CreeDetail (CleDADS, 9, 'S70.G01.00.004.006', LaGrille.Cells[7,Ligne],
               LaGrille.Cells[7,Ligne]);

if (LaGrille.Cells[8,Ligne] <> '') then
   CreeDetail (CleDADS, 11, 'S70.G01.00.004.009', LaGrille.Cells[8,Ligne],
               LaGrille.Cells[8,Ligne]);

if (LaGrille.Cells[9,Ligne] = '') then
   TraceErr.Items.Add('Le code postal de l''honoraire N°'+IntToStr(OrdreDest)+
                      ' n''est pas renseigné');
CreeDetail (CleDADS, 12, 'S70.G01.00.004.010', LaGrille.Cells[9,Ligne],
            LaGrille.Cells[9,Ligne]);

if (LaGrille.Cells[10,Ligne] = '') then
   TraceErr.Items.Add('Le bureau distributeur de l''honoraire N°'+
                      IntToStr(OrdreDest)+' n''est pas renseigné');
CreeDetail (CleDADS, 13, 'S70.G01.00.004.012',
            PGUpperCase(LaGrille.Cells[10,Ligne]),
            LaGrille.Cells[10,Ligne]);

{PT17-1
if ((LaGrille.Cells[11,Ligne] <> 'FRA') and
}
if ((LaGrille.Cells[11,Ligne]<>'FRA') and (LaGrille.Cells[11,Ligne]<>'GUF') and
   (LaGrille.Cells[11,Ligne]<>'GLP') and (LaGrille.Cells[11,Ligne]<>'MCO') and
   (LaGrille.Cells[11,Ligne]<>'MTQ') and (LaGrille.Cells[11,Ligne]<>'NCL') and
   (LaGrille.Cells[11,Ligne]<>'PYF') and (LaGrille.Cells[11,Ligne]<>'SPM') and
   (LaGrille.Cells[11,Ligne]<>'REU') and (LaGrille.Cells[11,Ligne]<>'ATF') and
   (LaGrille.Cells[11,Ligne]<>'WLF') and (LaGrille.Cells[11,Ligne]<>'MYT') and
//FIN PT17-1
   (LaGrille.Cells[11,Ligne]<>'')) then
   begin
   PaysISOLib(LaGrille.Cells[11,Ligne], CodeIso, Libelle);
   if (CodeIso <> '') then
      begin
      CreeDetail (CleDADS, 14, 'S70.G01.00.004.013', CodeIso,
                  LaGrille.Cells[11,Ligne]);

      CreeDetail (CleDADS, 15, 'S70.G01.00.004.014', Libelle, Libelle);
      end;
   end;

if (LaGrille.Cells[12,Ligne] = '') then
   TraceErr.Items.Add('La profession de l''honoraire N°'+IntToStr(OrdreDest)+
                      ' n''est pas renseignée');
CreeDetail (CleDADS, 16, 'S70.G01.00.005', LaGrille.Cells[12,Ligne],
            LaGrille.Cells[12,Ligne]);

if (LaGrille.Cells[13,Ligne] <> '') then
   CreeDetail (CleDADS, 17, 'S70.G01.00.006', LaGrille.Cells[13,Ligne],
               LaGrille.Cells[13,Ligne]);

if (LaGrille.Cells[14,Ligne] <> '') then
   CreeDetail (CleDADS, 18, 'S70.G01.00.007', LaGrille.Cells[14,Ligne],
               LaGrille.Cells[14,Ligne]);

if (LaGrille.Cells[15,Ligne] <> '') then
   CreeDetail (CleDADS, 19, 'S70.G01.00.008', LaGrille.Cells[15,Ligne],
               LaGrille.Cells[15,Ligne]);

if (LaGrille.Cells[16,Ligne] <> '') then
   CreeDetail (CleDADS, 20, 'S70.G01.00.009', LaGrille.Cells[16,Ligne],
               LaGrille.Cells[16,Ligne]);

if (LaGrille.Cells[17,Ligne] <> '') then
   CreeDetail (CleDADS, 27, 'S70.G01.00.016', LaGrille.Cells[17,Ligne],
               LaGrille.Cells[17,Ligne]);

if (LaGrille.Cells[18,Ligne] <> '') then
{PT12
   CreeDetail (StHonor, StType, OrdreDest, DebExer, FinExer, 21,
              'S70.G01.00.010', LaGrille.Cells[18,Ligne],
              LaGrille.Cells[18,Ligne]);
}
   begin
   BufOrig:= LaGrille.Cells[18,Ligne];
   BufDest:= '';
   if (Length(BufOrig)>1) then
      begin
      for i := 1 to Length(BufOrig) do
          BufDest:= BufDest+Copy(BufOrig, i, 1)+';';
      end
   else
      BufDest:= BufOrig;
   CreeDetail (CleDADS, 21, 'S70.G01.00.010', BufOrig, BufDest);
   end;
//FIN PT12


if (LaGrille.Cells[19,Ligne] <> '') then
   CreeDetail (CleDADS, 22, 'S70.G01.00.011', LaGrille.Cells[19,Ligne],
               LaGrille.Cells[19,Ligne]);

{PT13-2
if (LaGrille.Cells[20,Ligne] <> '') then
}
if ((LaGrille.Cells[20,Ligne] <> '') and (LaGrille.Cells[20,Ligne] <> '0')) then
   CreeDetail (CleDADS, 23, 'S70.G01.00.013',
               FloatToStr(Arrondi(StrToFloat(LaGrille.Cells[20,Ligne]), 0)),
               FloatToStr(Arrondi(StrToFloat(LaGrille.Cells[20,Ligne]), 0)));

if (LaGrille.Cells[21,Ligne] = '') then
   TraceErr.Items.Add('L''établissement de l''honoraire N°'+IntToStr(OrdreDest)+
                      ' n''est pas renseignée')
else
   begin
   QEtab:=OpenSQL('SELECT ET_SIRET'+
                  ' FROM ETABLISS WHERE'+
                  ' ET_ETABLISSEMENT="'+LaGrille.Cells[21,Ligne]+'"',True);
   if NOT QEtab.EOF then BufOrig := QEtab.FindField('ET_SIRET').AsString
     else BufOrig := '';
   Ferme(QEtab);
   ForceNumerique(BufOrig, BufDest);
   NIC := Copy(BufDest, 10, 5);
   if (NIC = '') then
      TraceErr.Items.Add('L''établissement de l''honoraire N°'+
                         IntToStr(OrdreDest)+' est incorrect');
   CreeDetail (CleDADS, 25, 'S70.G01.00.014', NIC, LaGrille.Cells[21,Ligne]);
   end;

{PT19
DecodeDate (FinExer, annee, mois, jour);
}
if (LaGrille.Cells[32,Ligne]<>'') then
   DateCloture:= StrToDate (LaGrille.Cells[32,Ligne])
else
   DateCloture:= FinExer;

DecodeDate (DateCloture, annee, mois, jour);
//FIN PT19
SDateCloture:= ColleZeroDevant (jour, 2)+ColleZeroDevant (mois, 2)+
               IntToStr (annee);
SDateCloture4:= Copy (SDateCloture, 1, 4);
if (SDateCloture4<>'3112') then
{PT19
   CreeDetail (CleDADS, 26, 'S70.G01.00.015', SDateCloture, DateToStr (FinExer));
}
   CreeDetail (CleDADS, 26, 'S70.G01.00.015', SDateCloture,
               DateToStr (DateCloture));
//FIN PT19

if ((LaGrille.Cells[22,Ligne]<>'') and (LaGrille.Cells[23,Ligne]<>'')) then
   begin
   CreeDetail (CleDADS, 101, 'S70.G01.01.001', LaGrille.Cells[22,Ligne],
               LaGrille.Cells[22,Ligne]);
{PT20
   if (LaGrille.Cells[23,Ligne]<>'0') then     //PT18
}
   if ((LaGrille.Cells[23,Ligne]<>'0') and (LaGrille.Cells[23,Ligne]<>'')) then
//FIN PT20
      CreeDetail (CleDADS, 102, 'S70.G01.01.002',
                  FloatToStr(Arrondi(StrToFloat(LaGrille.Cells[23,Ligne]), 0)),
                  FloatToStr(Arrondi(StrToFloat(LaGrille.Cells[23,Ligne]), 0)))
//PT18
   else
      TraceErr.Items.Add('Le montant 1 de l''honoraire N°'+IntToStr(OrdreDest)+
                         ' n''est pas renseigné');
//FIN PT18
   end;

if ((LaGrille.Cells[24,Ligne] <> '') and (LaGrille.Cells[25,Ligne] <> '')) then
   begin
   CreeDetail (CleDADS, 104, 'S70.G01.01.001', LaGrille.Cells[24,Ligne],
               LaGrille.Cells[24,Ligne]);
{PT20
   if (LaGrille.Cells[25,Ligne]<>'0') then     //PT18
}
   if ((LaGrille.Cells[25,Ligne]<>'0') and (LaGrille.Cells[25,Ligne]<>'')) then
//FIN PT20
      CreeDetail (CleDADS, 105, 'S70.G01.01.002',
                  FloatToStr(Arrondi(StrToFloat(LaGrille.Cells[25,Ligne]), 0)),
                  FloatToStr(Arrondi(StrToFloat(LaGrille.Cells[25,Ligne]), 0)))
//PT18
   else
      TraceErr.Items.Add('Le montant 2 de l''honoraire N°'+IntToStr(OrdreDest)+
                         ' n''est pas renseigné');
//FIN PT18
   end;

if ((LaGrille.Cells[26,Ligne] <> '') and (LaGrille.Cells[27,Ligne] <> '')) then
   begin
   CreeDetail (CleDADS, 107, 'S70.G01.01.001', LaGrille.Cells[26,Ligne],
               LaGrille.Cells[26,Ligne]);
{PT20
   if (LaGrille.Cells[27,Ligne]<>'0') then     //PT18
}
   if ((LaGrille.Cells[27,Ligne]<>'0') and (LaGrille.Cells[27,Ligne]<>'')) then
//FIN PT20
      CreeDetail (CleDADS, 108, 'S70.G01.01.002',
                  FloatToStr(Arrondi(StrToFloat(LaGrille.Cells[27,Ligne]), 0)),
                  FloatToStr(Arrondi(StrToFloat(LaGrille.Cells[27,Ligne]), 0)))
//PT18
   else
      TraceErr.Items.Add('Le montant 3 de l''honoraire N°'+IntToStr(OrdreDest)+
                         ' n''est pas renseigné');
//FIN PT18
   end;

if ((LaGrille.Cells[28,Ligne] <> '') and (LaGrille.Cells[29,Ligne] <> '')) then
   begin
   CreeDetail (CleDADS, 110, 'S70.G01.01.001', LaGrille.Cells[28,Ligne],
               LaGrille.Cells[28,Ligne]);
{PT20
   if (LaGrille.Cells[29,Ligne]<>'0') then     //PT18
}
   if ((LaGrille.Cells[29,Ligne]<>'0') and (LaGrille.Cells[29,Ligne]<>'')) then
//FIN PT20
      CreeDetail (CleDADS, 111, 'S70.G01.01.002',
                  FloatToStr(Arrondi(StrToFloat(LaGrille.Cells[29,Ligne]), 0)),
                  FloatToStr(Arrondi(StrToFloat(LaGrille.Cells[29,Ligne]), 0)))
//PT18
   else
      TraceErr.Items.Add('Le montant 4 de l''honoraire N°'+IntToStr(OrdreDest)+
                         ' n''est pas renseigné');
//FIN PT18
   end;

if ((LaGrille.Cells[30,Ligne] <> '') and (LaGrille.Cells[31,Ligne] <> '')) then
   begin
   CreeDetail (CleDADS, 113, 'S70.G01.01.001', LaGrille.Cells[30,Ligne],
               LaGrille.Cells[30,Ligne]);
{PT20
   if (LaGrille.Cells[31,Ligne]<>'0') then     //PT18
}
   if ((LaGrille.Cells[31,Ligne]<>'0') and (LaGrille.Cells[31,Ligne]<>'')) then
//FIN PT20
      CreeDetail (CleDADS, 114, 'S70.G01.01.002',
                  FloatToStr(Arrondi(StrToFloat(LaGrille.Cells[31,Ligne]), 0)),
                  FloatToStr(Arrondi(StrToFloat(LaGrille.Cells[31,Ligne]), 0)))
//PT18
   else
      TraceErr.Items.Add('Le montant 5 de l''honoraire N°'+IntToStr(OrdreDest)+
                         ' n''est pas renseigné');
//FIN PT18
   end;
TraceErr.Items.Add('Import de l''honoraire N°'+IntToStr(OrdreDest)+' terminé');
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 22/10/2002
Modifié le ... :   /  /
Description .. : En cas de changement d'exercice
Mots clefs ... : PAIE,PGDADSU
*****************************************************************}
procedure TOF_PG_DADSHonorRecupxls.ExerciceChange(Sender: TObject);
var
StExer : string;
QExer : TQuery;
begin
PGAnnee := THAnnee.value;
StExer:= 'SELECT PEX_DATEDEBUT, PEX_DATEFIN'+
         ' FROM EXERSOCIAL WHERE'+
         ' PEX_EXERCICE="'+PGAnnee+'"';

QExer:=OpenSQL(StExer,TRUE) ;
DebExer := IDate1900;   // PORTAGECWAS
FinExer := IDate1900;
if NOT QExer.EOF then
   begin
   DebExer := QExer.FindField ('PEX_DATEDEBUT').AsDateTime;
   FinExer := QExer.FindField ('PEX_DATEFIN').AsDateTime;
   end;
Ferme(QExer);
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 04/05/2004
Modifié le ... :   /  /
Description .. : Complément des raccourcis claviers
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_PG_DADSHonorRecupxls.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
TFVierge(Ecran).FormKeyDown(Sender, Key, Shift);

case Key of
     ord('V'): if ((GetControlVisible('RECOPIER')) and
                  (GetControlEnabled('RECOPIER')) and (ssAlt in Shift)) then
                  BtnRecop.Click; //Copie des éléments
     VK_F6: if ((GetControlVisible('BVALIDER')) and
               (GetControlEnabled('BVALIDER'))) then
               BtnValid.Click; //Copie des éléments
     end;
end;


Initialization
registerclasses([TOF_PG_DADSHonorRecupxls]);
end.

