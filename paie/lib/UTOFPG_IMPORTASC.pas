{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 13/02/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PG_IMPORTASC ()
Mots clefs ... : TOF;PG_IMPORTASC
*****************************************************************}
{
PT1-1  : 25/06/2007 VG V_72 Optimisation
PT1-2  : 25/06/2007 VG V_72 On continue l'import, même en cas d'erreur
PT1-3  : 25/06/2007 VG V_72 MemCheck
PT1-4  : 25/06/2007 VG V_72 Modification du principe, on ne traite que la liste
                            des tables connues
PT1-5  : 25/06/2007 VG V_72 Ajout de la gestion des modèles
PT2    : 13/07/2007 VG V_72 "Condition d''emploi" remplacé par "Caractéristique
                           activité" - FQ N°14568
PT3    : 24/07/2007 FC V_72 FQ 14423 Gestion des différents régimes
PT4    : 23/08/2007 VG V_80 Modification suite à la réunion
PT5    : 28/08/2007 VG V_80 Correction PT3
PT6    : 30/10/2007 GGU V_80 Gestion des absences en Horaire
PT7    : 10/12/2007 FC V_81 Rajout de contrôles de base
PT12   : 29/09/2008 VG Correction access vio
PT13   : 13/10/2008 JS      FQ n°15288 Création automatique du compte tiers lorsque le matricule est alpha-numérique
}
Unit UTOFPG_IMPORTASC;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     FE_Main,
{$ELSE}
     eMul,
     MaineAgl,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF,
     HTB97,
     HPanel,
     Agent,
     ParamSoc,
     uTob,
     Ent1,
     UTobDebug,
     UTOZ,
     MajTable,
     P5Util,
     StrUtils,
     PGCalendrier,
     PGOutils2,
     LookUp,
     TiersUtil,
     EntPaie,
     ExtCtrls,
     ULibEditionPaie;

Type TMotif = record
     ControlMotif : TControl;    //TControl
     TitreMotif : string;        //Titre de la fenêtre
     TableMotif : string;        //Table
     ColonneMotif : string;      //Code
     SelectMotif : string;       //Libellé
     WhereMotif : string;        //Condition
     EditMotif : string;         //THEdit
     LabelMotif : string;        //TLabel
     end;

Type
  TOF_PG_IMPORTASC = Class (TOF)
    procedure OnNew                    ; override;
    procedure OnDelete                 ; override;
    procedure OnUpdate                 ; override;
    procedure OnLoad                   ; override;
    procedure OnArgument (S : String ) ; override;
    procedure OnDisplay                ; override;
    procedure OnClose                  ; override;
    procedure OnCancel                 ; override;
  private
    BFerme, BLancer, BParam, BPrecedent, BSuivant : TToolbarButton97;
    P: TPageControl;
    Plan: THPanel;
    cControls: THListBox;
    lEtape: THLabel;
    Buzy, Erreur, NormalClose, VerifPresenceParam : boolean;
{VerifPresenceParam pour savoir si on doit faire les contrôles pour savoir si
les valeurs sont présentes dans les paramètres}
    PCodageSal, PCumuls, PLancer, PSelect: TTabSheet;
    EdtFichier, Modele : THEdit;
    ChampOblig, Chemin, FichierATraiter, NomTable, PGTypeNumSal : string;
    StTitre : string;
    F, FRapport : TextFile;
    TozFile : TOZ;
    TExtract, TListeFile, TListeZip, TModele, Tob_MotifAbs : TOB;
    FREPORT : TListBox;
    NbChampOblig, ResultAnnee, ResultMois, ResultDepart : Integer;
    DTclot : TDateTime;
    LeMotif : TMotif;
    FCodeSal: TRadioGroup;
    ListeCodeSal : TStringList;
    Index : integer;
    Num : Longint;

    procedure RestorePage;
    function FirstPage: TTabSheet;
    procedure bSuivantClick (Sender: TObject);
    function NextPage: TTabSheet;
    procedure bPrecedentClick (Sender: TObject);
    function PreviousPage: TTabSheet;
    procedure EdtFichierChange (Sender: TObject);
    procedure PChange (Sender: TObject);
    procedure SetTabs;
    procedure DisplayStep;
    function PageNumber: Integer;
    function PageCount: Integer;
    procedure ModeleChange(Sender: TObject);
    procedure ModeleElipsisClick (Sender: TObject);
    procedure bParamClick (Sender: TObject);
    procedure BLancerClick (Sender: TObject);
    function Traitement_ASC : boolean;
    procedure TraitementCumulPaie ();
    procedure EnregCumulPaieDefaut (TC : TOB);
    function ControleValCumulPaie (Cumul,Predefini:String) : boolean;
    procedure TraitementTable (NomTable : string);
//    procedure TraitementEtabliss();
    procedure EnregEtablissDefaut (TEtablissFille : TOB);
    function ControleValEtabliss (TEtablissFille : TOB) : boolean;
//    procedure TraitementEtabCompl();
    procedure EnregEtabComplDefaut (TEtabComplFille : TOB);
    procedure MajChampOblig(LibChamp : String);
    function ControleValEtabCompl (TEtabComplFille : TOB) : boolean;
//    procedure TraitementSalaries();
    procedure EnregSalDefaut (TSalariesFille : TOB);
    function ControleValSalaries (TSalariesFille : TOB) : boolean;
    procedure VerifTablette(Champ,NomTablette,LibCleErr,LibChampErr : String;T : TOB);
    function TestHoraire(HoraireInf, HoraireSup: Double): Boolean;
//    function TestNumeroSSNaissance(NoSS: string; SSNaiss: string;T : TOB): integer;
//    procedure TraitementHistoCumul();
    function ControleValHistoCumul (THistoCumulFille : TOB) : boolean;
//    procedure TraitementAbsenceSalarie();
    procedure EnregAbsenceSalarieDefaut (TAbsenceFille : TOB);
    function ControleValAbsenceSalarie (TAbsenceFille : TOB) : boolean;
    procedure EnregRibDefaut (TRibFille : TOB);
    function ControleValRib (TRibFille : TOB) : boolean;
    procedure EnregContratDefaut (TContratFille : TOB);
    function ControleValContrat (TContratFille : TOB) : boolean;
    procedure EnregEnfantDefaut (TEnfantFille : TOB);
    function ControleValEnfant (TEnfantFille : TOB) : boolean;
    procedure EnregHistoBullDefaut (THistoBullFille : TOB);
    function ControleValHistoBull (THistoBullFille : TOB) : boolean;
    function RechercheReprise (Validite : Tdatetime; Salarie, TypeConge,
                               TypeImpute : string; TypeConge2 : string='') : Tquery;
    procedure RechercheExerciceCp(Validite: tdatetime; var DTdeb, DtFin: tdatetime);
    function ExtractConnu (Table : string; Enregistre : boolean) : boolean;
    function RecupereNomTable (NomFic : string) : string;
    function Extraction (NomFic, NomTable : string; TFille : TOB;
                         Enregistre : boolean) : boolean;
    procedure AppliqueLeParam (TExtractFille, TModele : TOB;
                               ChampsModele : string);
    procedure AppliqueParamDefaut (NomTable : string; TExtract : TOB);
    function ControleExtract (NomTable : string) : boolean;
    function ControleCumulPresent : boolean;
    function RecupChamp (Fichier, Champ, Where : String;
                         var Trouve : boolean) : Variant;
    function ChangeCodeSalarie (T : TOB; Prefixe : string ;
                                PeutCreer : boolean) : boolean;
    procedure InitCodeSal;
    procedure FinCodeSal;
    function SalarieAbsent (St : string) : boolean;
    function SalarieDansListe (St : string; PeutCreer : boolean;
                               var Code : Integer) : boolean;

  end;

Implementation

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 20/03/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_PG_IMPORTASC.OnNew;
begin
Inherited;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 20/03/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_PG_IMPORTASC.OnDelete;
begin
Inherited;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 20/03/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_PG_IMPORTASC.OnUpdate;
begin
Inherited;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 20/03/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_PG_IMPORTASC.OnLoad;
var
i, j: Integer;
T: TTabSheet;
begin
Inherited;
AgentAnimate (aaAssistant);
Buzy:= false;
NormalClose:= False;
if P.ControlCount > 0 then
   begin
   for i := 0 to P.ControlCount - 1 do
       begin
       T:= TTabSheet (P.Controls[i]);
       for j := 0 to T.ControlCount - 1 do
           begin
           if T.Controls[j] is TComboBox then
              TComboBox (T.Controls[j]).ItemIndex:= 0;
           end;
       end;
   P.ActivePage:= FirstPage;
   PChange (nil);
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 20/03/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_PG_IMPORTASC.OnArgument (S : String);
begin
Inherited;
BFerme:= TToolbarButton97 (GetControl ('BFERME'));

BLancer:= TToolbarButton97 (GetControl ('BLANCER'));
if (BLancer<>nil) then
   BLancer.OnClick:= BLancerClick;

BPrecedent:= TToolbarButton97 (GetControl ('BPRECEDENT'));
if (BPrecedent<>nil) then
   BPrecedent.OnClick:= bPrecedentClick;

BSuivant:= TToolbarButton97 (GetControl ('BSUIVANT'));
if (BSuivant<>nil) then
   BSuivant.OnClick:= bSuivantClick;

//PT1-5
BParam:= TToolbarButton97 (GetControl ('BPARAM'));
if (BParam<>nil) then
   BParam.OnClick:= bParamClick;
//FIN PT1-5

cControls:= THListBox (GetControl ('CCONTROLS'));

EdtFichier:= THEdit (GetControl ('EDTFICHIER'));
if (EdtFichier<>nil) then
   begin
   EdtFichier.OnChange:= EdtFichierChange;
   EdtFichier.Text := '';
   end;

//PT1-5
Modele:= THEdit (GetControl ('MODELE'));
if Modele <> NIL then
   Modele.OnElipsisClick:= ModeleElipsisClick;
ModeleChange (Modele);
SetControlEnabled ('LIBMODELE', False);
//FIN PT1-5
   
lEtape:= THLabel (GetControl ('LETAPE'));

P:= TPageControl (GetControl ('P'));

Plan:= THPanel (GetControl ('PLAN'));

PSelect:= TTabSheet (GetControl ('PSELECT'));
PCodageSal:= TTabSheet (GetControl ('PCODAGESAL'));
PCumuls:= TTabSheet (GetControl ('PCUMULS'));
PLancer:= TTabSheet (GetControl ('PLANCER'));

SetControlText ('CBCUMULS', 'STD');


StTitre:= 'Import Fichier ASC';

PGTypeNumSal := GetParamSocSecur('SO_PGTYPENUMSAL', '');

FCodeSal:= TRadioGroup (GetControl ('FCODESAL'));   //PT4

end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 20/03/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_PG_IMPORTASC.OnClose;
begin
Inherited;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 20/03/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_PG_IMPORTASC.OnDisplay ();
begin
Inherited;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 20/03/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_PG_IMPORTASC.OnCancel ();
begin
Inherited;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 20/03/2007
Modifié le ... :   /  /
Description .. : Restore la page (recopie les controles de Plan vers la page
Suite ........ : active)
Mots clefs ... :
*****************************************************************}
procedure TOF_PG_IMPORTASC.RestorePage;
var
i: Integer;
begin
// Sauvegarde des ItemIndex
for i := 0 to Ecran.ComponentCount - 1 do
    begin
    if Ecran.Components[i] is TComboBox then
       TComboBox(Ecran.Components[i]).Tag:= TComboBox(Ecran.Components[i]).ItemIndex;
    end;

// Changement de parent
while Plan.ControlCount > 0 do
      Plan.Controls[0].Parent:= P.ActivePage;
// Restauration des ItemIndex
for i := 0 to Ecran.ComponentCount - 1 do
    begin
    if Ecran.Components[i] is TComboBox then
       TComboBox(Ecran.Components[i]).ItemIndex:= TComboBox(Ecran.Components[i]).Tag;
    end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 20/03/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TOF_PG_IMPORTASC.FirstPage: TTabSheet;
begin
result:= P.Pages[0];
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 20/03/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_PG_IMPORTASC.bSuivantClick(Sender: TObject);
var
t: TTabSheet;
begin
RestorePage;
if P.ActivePage.PageIndex < P.PageCount - 1 then
   t:= P.Pages[P.ActivePage.PageIndex + 1]
else
   t:= nil;

if t = nil then
   P.SelectNextPage(True)
else
   begin
   P.ActivePage:= t;
   PChange(nil);
   end;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 20/03/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TOF_PG_IMPORTASC.NextPage: TTabSheet;
Var
TFille : TOB;
EstPresent : Boolean;
GCumuls : THGrid;
begin
if P.ActivePage.PageIndex < P.PageCount - 1 then
   result:= P.Pages[P.ActivePage.PageIndex + 1]
else
   result:= nil;

if (P.ActivePage=PSelect) then
   begin
   if (edtFichier.text='') then
      result:=nil  ;
   if not FileExists(edtFichier.text) then
      result:=nil ;
   end
else
if (P.ActivePage=PCumuls) then
   begin
   TozFile:= TOZ.Create;
   TozFile.OpenZipFile (GetControlText ('EdtFichier'), MoOpen);
   TListeZip:= TozFile.ConvertListInTob;
//TobDebug (TListeZip);
   TozFile.OpenSession (OsExt);
   TListeFile:= TOB.Create ('Mes fichiers', nil, -1);
   while (TListeZip.Detail.Count>0) do
         begin
         TFille:= TListeZip.Detail[0];
         FichierATraiter:= TFille.GetValue ('ZF_NAME');
         TozFile.ProcessFile (FichierATraiter);
         TFille.ChangeParent (TListeFile, -1);
         NomTable:= RecupereNomTable (FichierATraiter);
         TFille.AddChampSupValeur ('NOMTABLE', NomTable);
         end;
   FreeAndNil (TListeZip);
//TobDebug (TListeFile);

   Chemin:= ExtractFileDir (GetControlText ('EdtFichier'));
   TozFile.SetDirOut (Chemin);
   TozFile.CloseSession;
//   TozFile.Destroy;
   FreeAndNil (TozFile);

   TExtract:= TOB.Create (NomTable, nil, -1);
   ExtractConnu ('CUMULPAIE', False);
//TobDebug (TExtract);
   EstPresent:= ControleCumulPresent;
   if (EstPresent=True) then
      begin
      GCumuls:= THGrid (GetControl ('GCUMULS'));
      TExtract.Detail.Sort ('PCL_CUMULPAIE');
      TExtract.PutGridDetail (GCumuls, FALSE, FALSE,
                              'PCL_CUMULPAIE;PCL_LIBELLE;IMPORT', TRUE);
      end;
{PT4
   else
      BSuivant.Click;
}      
   FreeAndNil (TExtract);     //PT1-3 réactivation
   end
else
if (P.ActivePage=PLancer) then
   begin
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 20/03/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_PG_IMPORTASC.bPrecedentClick(Sender: TObject);
var
t: TTabSheet;
begin
RestorePage;
t:= PreviousPage;
if t = nil then
   P.SelectNextPage(False)
else
   begin
   P.ActivePage:= t;
   PChange(nil);
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 20/03/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TOF_PG_IMPORTASC.PreviousPage: TTabSheet;
begin
if P.ActivePage.PageIndex > 0 then
   result:= P.Pages[P.ActivePage.PageIndex - 1]
else
   result:= nil;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 20/03/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_PG_IMPORTASC.edtFichierChange(Sender: TObject);
begin
inherited;
PChange(Nil) ;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 20/03/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_PG_IMPORTASC.PChange(Sender: TObject);
var
i: Integer;
T: TTabSheet;
fs, fp: Boolean;
begin
T:= P.ActivePage;
if T = nil then
   exit;
while T.ControlCount > 0 do
      begin
      if T.Controls[0] is TComboBox then
// Sauvegarde de ItemIndex dans Tag
         TComboBox(T.Controls[0]).Tag:= TComboBox(T.Controls[0]).ItemIndex;
      T.Controls[0].Parent:= Plan;
      end;
//Restauration des ItemIndex
for i := 0 to Plan.ControlCount - 1 do
    begin
    if Plan.Controls[i] is TComboBox then
       TComboBox(Plan.Controls[i]).ItemIndex:= TComboBox(Plan.Controls[i]).Tag;
    end;
SetTabs;
if PreviousPage <> nil then
   fp:= (T.PageIndex > 0)
else
   fp:= FALSE;
if NextPage <> nil then
   fs:= (T.PageIndex < P.ControlCount - 1)
else
   fs:= FALSE;
bPrecedent.Enabled:= fp;
bSuivant.Enabled:= fs;
bSuivant.Default:= bSuivant.Enabled;
bFerme.Default:= not bSuivant.Default;
DisplayStep;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 20/03/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_PG_IMPORTASC.SetTabs;
var
i: Integer;
C: TWinControl;
begin
cControls.Items.Clear;
for i := 0 to Plan.ControlCount - 1 do
    begin
    if (Plan.Controls[i] is TWinControl) then
       begin
       if (Plan.Controls[i].Left < 0) or (Plan.Controls[i].Top < 0) then
          continue;
       cControls.Items.Add (FormatFloat ('0000', Plan.Controls[i].Top)+
                            FormatFloat ('0000', Plan.Controls[i].Left)+
                            Format_String(Plan.Controls[i].Name, 20));
       end;
    end;

for i := 0 to cControls.Items.Count - 1 do
    begin
    C:= TWinControl (Ecran.FindComponent (Trim (Copy (cControls.Items[i], 9, 20))));
    if C <> nil then
       C.TabOrder:= i;
    end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 20/03/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_PG_IMPORTASC.DisplayStep;
begin
lEtape.Caption:= 'Etape '+IntToStr(PageNumber)+'/'+IntToStr(PageCount);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 20/03/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TOF_PG_IMPORTASC.PageNumber: Integer;
begin
result:= P.ActivePage.PageIndex + 1;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 20/03/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TOF_PG_IMPORTASC.PageCount: Integer;
begin
result:= P.ControlCount;
end;

//PT1-5
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 09/05/2007
Modifié le ... :   /  /
Description .. : Modification de la valeur du modèle
Mots clefs ... : PAIE;PGIMPORT
*****************************************************************}
procedure TOF_PG_IMPORTASC.ModeleChange(Sender: TObject);
var
sWhere : string;
QRechModele : TQuery;
begin
LeMotif.ControlMotif:= Modele;
LeMotif.TitreMotif:= 'Modèles d''import paie';
LeMotif.TableMotif:= 'PGDEFAUTPARAM';
LeMotif.ColonneMotif:= 'PDM_CODE';
LeMotif.SelectMotif:= 'PDM_LIBELLEPARAM';
LeMotif.WhereMotif:= '##PDM_PREDEFINI##';
LeMotif.EditMotif:= 'MODELE';
LeMotif.LabelMotif:= 'LIBMODELE';

sWhere:= 'SELECT '+LeMotif.ColonneMotif+', '+LeMotif.SelectMotif+
         ' FROM '+LeMotif.TableMotif+' WHERE ';
if (GetControlText (LeMotif.EditMotif)<>'') then
   sWhere:= sWhere+
            LeMotif.ColonneMotif+'="'+GetControlText (LeMotif.EditMotif)+'"'
else
   sWhere:= sWhere+'PDM_DEFAUT="X"';
QRechModele:= OpenSql(sWhere,TRUE);
if (not QRechModele.EOF) then
   begin
   SetControlText(LeMotif.EditMotif, QRechModele.Fields[0].asstring);
   SetControlText(LeMotif.LabelMotif, QRechModele.Fields[1].asstring);
   end
else
   begin
   SetControlText(LeMotif.EditMotif, '');
   SetControlText(LeMotif.LabelMotif, '');
   end;
Ferme(QRechModele);
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 09/05/2007
Modifié le ... :   /  /
Description .. : Click sur l'ellipsis modèle d'import
Mots clefs ... : PAIE;PGIMPORT
*****************************************************************}
procedure TOF_PG_IMPORTASC.ModeleElipsisClick(Sender: TObject);
var
StSelect, sWhere : string;
begin
ModeleChange (Sender);

StSelect:= 'SELECT '+LeMotif.ColonneMotif+', '+LeMotif.SelectMotif+
           ' FROM '+LeMotif.TableMotif;
if (LeMotif.WhereMotif<> '') then
   begin
   sWhere:= StSelect+' WHERE '+LeMotif.WhereMotif;
   sWhere:= sWhere+' GROUP BY '+LeMotif.ColonneMotif+', '+LeMotif.SelectMotif;
   end;
{$IFNDEF EAGLCLIENT}
LookUpList (LeMotif.ControlMotif, LeMotif.TitreMotif, LeMotif.TableMotif,
            LeMotif.ColonneMotif, LeMotif.SelectMotif, LeMotif.WhereMotif,
            LeMotif.ColonneMotif, TRUE, -1, sWhere);
{$ELSE}
LookupList (LeMotif.ControlMotif, LeMotif.TitreMotif, LeMotif.TableMotif,
            LeMotif.ColonneMotif, LeMotif.SelectMotif, LeMotif.WhereMotif,
            LeMotif.ColonneMotif, TRUE, -1, StSelect);
{$ENDIF}
ModeleChange (Sender);
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 02/04/2007
Modifié le ... :   /  /
Description .. : Click sur le bouton "Paramétrage"
Mots clefs ... : PAIE;PG_IMPORTASC
*****************************************************************}
procedure TOF_PG_IMPORTASC.bParamClick (Sender: TObject);
var
Lequel, State : string;
begin
if (GetControlText ('MODELE')<>'') then
   begin
   Lequel:= GetControlText ('MODELE');
   State:= 'MODIFICATION;';
   end
else
   begin
   Lequel:= '001';
   State:= 'CREATION;';
   end;
AGLLanceFiche ('PAY', 'PARAMDEFAUT', '', Lequel, State+Lequel);
end;
//FIN PT1-5

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 20/03/2007
Modifié le ... :   /  /
Description .. : Click sur le bouton "Lancer l'importation" : récupération des
Suite ........ : données
Mots clefs ... : PAIE;PG_IMPORTASC
*****************************************************************}
procedure TOF_PG_IMPORTASC.BLancerClick(Sender: TObject);
var
FileName, StMessage : string;
reponse : integer;
ContinueT : boolean;
begin
InitCodeSal;        //PT4
SetControlEnabled ('BLANCER', FALSE);
SetControlEnabled ('BSUIVANT', FALSE);
SetControlEnabled ('BPRECEDENT', FALSE);

AssignFile(F, GetControlText ('EDTFICHIER'));
Reset(F);

//Création ou modification du fichier .log
FileName:= 'ascpgi.log';
AssignFile(FRapport, FileName);
if FileExists(FileName) then
   Append(FRapport)
else
   ReWrite(FRapport);
Writeln(FRapport, '');
Writeln(FRapport, 'Début de traitement : '+DateTimeToStr(Now));

//Fonction d'import
ContinueT:= Traitement_ASC;

bSuivantClick(Nil) ;

if ((GetCheckBoxState ('FCumul')=CbChecked) and (ContinueT = True)) then
   begin
   StMessage:= 'Important : vous devez vérifier les cumuls des salariés#13#10'+
               'AVANT de commencer la saisie des bulletins de paie.#13#10'+
               'Voulez-vous les éditer immédiatement ?';
   reponse := PGIAsk(StMessage, StTitre);
   if (reponse=mrYes) then
      AglLanceFiche ('PAY','FICHECUMSAL', '', '' , '' );

   StMessage:= 'Important : vous devez vérifier les bases de cotisation#13#10'+
               'AVANT de commencer la saisie des bulletins de paie.#13#10'+
               'Voulez-vous les éditer immédiatement ?';
   reponse := PGIAsk(StMessage, StTitre);
   if (reponse=mrYes) then
      AglLanceFiche ('PAY','BASECOT_ETAT', '', '' , '' );
   StMessage:='Traitement terminé';
   end
else
   if (ContinueT = True) then
      begin
      StMessage:='Traitement terminé';
      PGIBox(StMessage, StTitre);
      end;
Writeln(FRapport, StMessage+' : '+DateTimeToStr(Now));
CloseFile(FRapport);

FinCodeSal;         //PT4
BSuivant.Click;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 20/03/2007
Modifié le ... :   /  /
Description .. : Traitement concernant le fichier ASC
Mots clefs ... : PAIE;PG_IMPORTASC
*****************************************************************}
function TOF_PG_IMPORTASC.Traitement_ASC : boolean;
var
NomFic, sWhere : string;
i : integer;
ListeTable : Array[0..13] of string;
TTable, TTableD : TOB;
begin

VerifPresenceParam := True; // Passer par un PARAMSOC ou autre ???

//ListeTable[0]:= 'CUMULPAIE';
ListeTable[0]:= 'SOCIETE';
ListeTable[1]:= 'ETABLISS';
ListeTable[2]:= 'ETABCOMPL';
ListeTable[3]:= 'CONVENTIONCOLL';
ListeTable[4]:= 'TAUXAT';
ListeTable[5]:= 'CHOIXCOD';
ListeTable[6]:= 'SALARIES';
//PT4
ListeTable[7]:= 'CONTRATTRAVAIL';
ListeTable[8]:= 'RIB';
//FIN PT4
ListeTable[9]:= 'ENFANTSALARIE';
ListeTable[10]:= 'HISTOCUMSAL';
ListeTable[11]:= 'HISTOBULLETIN';     //PT4
ListeTable[12]:= 'MINIMUMCONVENT';
ListeTable[13]:= 'ABSENCESALARIE';

// initialiser la liste contenant le rapport
FREPORT := TListBox(GetControl('FREPORT'));
FREPORT.clear;

// Encadrer la mise à jour des tables par un BeginTrans/CommitTrans/RollBack
{PT1-2 Ne pas bloquer en cas d'erreur
BeginTrans;
}
Erreur := False;

// Traitement particulier pour le fichier CUMULPAIE
TraitementCumulPaie;

//Construction de la tob sur les données du modèle
//PT1-5
sWhere:= 'SELECT *'+
         ' FROM PGDEFAUTPARAM WHERE'+
         ' PDM_CODE="'+GetControlText ('MODELE')+'"';
TModele:= TOB.Create ('Le modele', NIL, -1);
{Optimisation PT1-1
QModele:=OpenSql (sWhere, TRUE);
TModele.LoadDetailDB ('PGDEFAUTPARAM', '', '', QModele, False);
Ferme (QModele);
}
TModele.LoadDetailDBFromSQL ('PGDEFAUTPARAM', SWhere);
//FIN PT1-1

//PT4
sWhere:= 'SELECT DT_NOMTABLE'+
         ' FROM DETABLES WHERE'+
         ' DT_DOMAINE="P"';
TTable:= TOB.Create ('Les tables Paie', nil, -1);
TTable.LoadDetailDBFromSQL ('DETABLES', SWhere);
//FIN PT4

for i:=0 to 13 do
    result:= ExtractConnu (ListeTable[i], True);

while (TListeFile.Detail.Count>0) do
      begin
      NomFic:= TListeFile.Detail[0].GetValue ('ZF_NAME');
      NomTable:= RecupereNomTable (NomFic);

//PT1-4 Ne pas traiter les tables inconnues si absence du fichier CEGID.par
{PT4
      if FileExists('CEGID.par') then
}
      TTableD:= TTable.FindFirst (['DT_NOMTABLE'], [NomTable], True);
      if ((FileExists('CEGID.par')) and (TTableD<>nil)) then
//FIN PT4      
         begin
         NomFic:= Chemin+'\'+NomFic;
         result:= Extraction (NomFic, NomTable, TListeFile.Detail[0], True);
         end;
      end;
FreeAndNil (TListeFile);
FreeAndNil (TModele);
//FIN PT1-5
FreeAndNil (TTable); //PT4

{PT1-2 Ne pas bloquer en cas d'erreur
if Erreur then
   RollBack
else
   CommitTrans;
}

//EnregSalDefaut (T);
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 27/03/2007
Modifié le ... :   /  /
Description .. : Procédure d'extraction des tables connues
Mots clefs ... : PAIE;PG_IMPORTASC
*****************************************************************}
function TOF_PG_IMPORTASC.ExtractConnu (Table : string; Enregistre : boolean) : boolean;
var
NomFic : string;
TFille : TOB;
begin
TFille:= TListeFile.FindFirst (['NOMTABLE'], [Table], True);
if (TFille <> nil) then
   begin
   NomFic:= TFille.GetValue ('ZF_NAME');
   NomTable:= TFille.GetValue ('NOMTABLE');
   NomFic:= Chemin+'\'+NomFic;
   result:= Extraction (NomFic, NomTable, TFille, Enregistre);
   end
else
   result:= False;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 27/03/2007
Modifié le ... :   /  /
Description .. : Procédure de lecture du nom de la table
Mots clefs ... : PAIE;PG_IMPORTASC
*****************************************************************}
function TOF_PG_IMPORTASC.RecupereNomTable (NomFic : string) : string;
begin
result:= ExtractFileName (NomFic);
result:= Copy (result, 11, Length (result)-14);
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 27/03/2007
Modifié le ... :   /  /
Description .. : Procédure d'extraction
Mots clefs ... : PAIE;PG_IMPORTASC
*****************************************************************}
function TOF_PG_IMPORTASC.Extraction (NomFic, NomTable : string; TFille : TOB;
                                      Enregistre : boolean) : boolean;
begin
if (Enregistre=True) then
   TExtract:= TOB.Create (NomTable, nil, -1);
TOBLoadFromFile (NomFic, nil, TExtract);
AppliqueParamDefaut (NomTable, TExtract);	//PT1-5
FreeAndNil (TFille);
result:= ControleExtract (NomTable);
//TobDebug (TExtract);
if (result=True) then
   begin
   if (Enregistre=True) then
      begin
      TraitementTable (NomTable);
      DeleteFile (NomFic);   
      end;
   end;

if (Enregistre=True) then
   FreeAndNil (TExtract);
end;

//PT1-5
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 11/06/2007
Modifié le ... :   /  /
Description .. : Procédure d'application d'un paramètre du modèle choisi
Mots clefs ... : PAIE;PG_IMPORTASC
*****************************************************************}
procedure TOF_PG_IMPORTASC.AppliqueLeParam (TExtractFille, TModele : TOB; ChampsModele : string);
var
TModeleD : TOB;
begin
if (TExtractFille.FieldExists (ChampsModele)=False) then
   begin
   TModeleD:= TModele.FindFirst (['PDM_NOMCHAMP'], [ChampsModele], True);
   if (TModeleD<>nil) then
      TExtractFille.AddChampSupValeur (ChampsModele,
                                       TModeleD.GetValue ('PDM_PGVALEUR'), True);
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 29/05/2007
Modifié le ... :   /  /
Description .. : Procédure d'application des paramètres du modèle choisi
Mots clefs ... : PAIE;PG_IMPORTASC
*****************************************************************}
procedure TOF_PG_IMPORTASC.AppliqueParamDefaut (NomTable : string; TExtract : TOB);
var
TExtractFille : TOB;
ChampsModeleSO : Array[0..20] of string;
ChampsModeleET : Array[0..1] of string;
ChampsModeleETB : Array[0..61] of string;
ChampsModelePSA : Array[0..8] of string;
i : integer;
begin
ChampsModeleSO[0]:= 'SO_PGRACINEAUXI';
ChampsModeleSO[1]:= 'SO_PGNUMAUXI';
ChampsModeleSO[2]:= 'SO_PGTYPENUMSAL';
ChampsModeleSO[3]:= 'SO_PGINCSALARIE';
ChampsModeleSO[4]:= 'SO_PGTIERSAUXIAUTO';
ChampsModeleSO[5]:= 'SO_PGDECALAGE';
ChampsModeleSO[6]:= 'SO_PGDECALAGEPETIT';
ChampsModeleSO[7]:= 'SO_PGANALYTIQUE';
ChampsModeleSO[8]:= 'SO_PGCONGES';
ChampsModeleSO[9]:= 'SO_PGMODIFLIGNEIMP';
ChampsModeleSO[10]:= 'SO_PGABSENCE';
ChampsModeleSO[11]:= 'SO_PGABSENCECHEVAL';
ChampsModeleSO[12]:= 'SO_PGCHR6SEMAINE';
ChampsModeleSO[13]:= 'SO_PGMSA';
ChampsModeleSO[14]:= 'SO_PGINTERMITTENTS';
ChampsModeleSO[15]:= 'SO_PGBTP';
ChampsModeleSO[16]:= 'SO_PGRESPONSABLES';
ChampsModeleSO[17]:= 'SO_PGCOEFFEVOSAL';
ChampsModeleSO[18]:= 'SO_PGSAISIEARRET';
ChampsModeleSO[19]:= 'SO_PGTICKETRESTAU';
ChampsModeleSO[20]:= 'SO_PGPCS2003';

ChampsModeleET[0]:= 'ET_ACTIVITE';
ChampsModeleET[1]:= 'ET_JURIDIQUE';

ChampsModeleETB[0]:= 'ETB_HORAIREETABL';
ChampsModeleETB[1]:= 'ETB_SMIC';
ChampsModeleETB[2]:= 'ETB_PCTFRAISPROF';
ChampsModeleETB[3]:= 'ETB_TAUXVERSTRANS';
ChampsModeleETB[4]:= 'ETB_NBJOUTRAV';
ChampsModeleETB[5]:= 'ETB_NBREACQUISCP';
ChampsModeleETB[6]:= 'ETB_NBACQUISCP';
ChampsModeleETB[7]:= 'ETB_NBRECPSUPP';
ChampsModeleETB[8]:= 'ETB_VALANCCP';
ChampsModeleETB[9]:= 'ETB_DATEACQCPANC';
ChampsModeleETB[10]:= 'ETB_VALODXMN';
ChampsModeleETB[11]:= 'ETB_TYPSMIC';
ChampsModeleETB[12]:= 'ETB_STANDCALEND';
ChampsModeleETB[13]:= 'ETB_JOURHEURE';
ChampsModeleETB[14]:= 'ETB_MEDTRAVGU';
ChampsModeleETB[15]:= 'ETB_CODEDDTEFPGU';
ChampsModeleETB[16]:= 'ETB_PROFILREM';
ChampsModeleETB[17]:= 'ETB_PERIODBUL';
ChampsModeleETB[18]:= 'ETB_REDREPAS';
ChampsModeleETB[19]:= 'ETB_REDRTT1';
ChampsModeleETB[20]:= 'ETB_PROFILAFP';
ChampsModeleETB[21]:= 'ETB_PROFILRET';
ChampsModeleETB[22]:= 'ETB_PROFILMUT';
ChampsModeleETB[23]:= 'ETB_PROFILANCIEN';
ChampsModeleETB[24]:= 'ETB_PROFIL';
ChampsModeleETB[25]:= 'ETB_PROFILRBS';
ChampsModeleETB[26]:= 'ETB_REDRTT2';
ChampsModeleETB[27]:= 'ETB_PROFILPRE';
ChampsModeleETB[28]:= 'ETB_PROFILTSS';
ChampsModeleETB[29]:= 'ETB_PROFILAPP';
ChampsModeleETB[30]:= 'ETB_PROFILTRANS';
ChampsModeleETB[31]:= 'ETB_PROFILCGE';
ChampsModeleETB[32]:= 'ETB_1ERREPOSH';
ChampsModeleETB[33]:= 'ETB_EDITBULCP';
ChampsModeleETB[34]:= 'ETB_RELIQUAT';
ChampsModeleETB[35]:= 'ETB_2EMEREPOSH';
ChampsModeleETB[36]:= 'ETB_BASANCCP';
ChampsModeleETB[37]:= 'ETB_TYPDATANC';
ChampsModeleETB[38]:= 'ETB_VALORINDEMCP';
ChampsModeleETB[39]:= 'ETB_MVALOMS';
ChampsModeleETB[40]:= 'ETB_PAIEVALOMS';
ChampsModeleETB[41]:= 'ETB_MOISPAIEMENT';
ChampsModeleETB[42]:= 'ETB_PGMODEREGLE';
ChampsModeleETB[43]:= 'ETB_BCMODEREGLE';
ChampsModeleETB[44]:= 'ETB_BCMOISPAIEMENT';
ChampsModeleETB[45]:= 'ETB_PAIACOMPTE';
ChampsModeleETB[46]:= 'ETB_PAIFRAIS';
ChampsModeleETB[47]:= 'ETB_CODESECTION';
ChampsModeleETB[48]:= 'ETB_PRUDHCOLL';
ChampsModeleETB[49]:= 'ETB_PRUDHSECT';
ChampsModeleETB[50]:= 'ETB_PRUDHVOTE';
ChampsModeleETB[51]:= 'ETB_PRUDH';
ChampsModeleETB[52]:= 'ETB_DADSSECTION';
ChampsModeleETB[53]:= 'ETB_TYPDADSSECT';
ChampsModeleETB[54]:= 'ETB_SUBROGATION';
ChampsModeleETB[55]:= 'ETB_PERIODCT';
ChampsModeleETB[56]:= 'ETB_TRANS9SAL';
ChampsModeleETB[57]:= 'ETB_CONGESPAYES';
ChampsModeleETB[58]:= 'ETB_JOURPAIEMENT';
ChampsModeleETB[59]:= 'ETB_BCJOURPAIEMENT';
ChampsModeleETB[60]:= 'ETB_JEDTDU';
ChampsModeleETB[61]:= 'ETB_JEDTAU';

ChampsModelePSA[0]:= 'PSA_UNITEPRISEFF';
ChampsModelePSA[1]:= 'PSA_SALAIRETHEO';
ChampsModelePSA[2]:= 'PSA_ETATBULLETIN';
ChampsModelePSA[3]:= 'PSA_PROFILTPS';
ChampsModelePSA[4]:= 'PSA_REGIMESS';
ChampsModelePSA[5]:= 'PSA_DADSCAT';
ChampsModelePSA[6]:= 'PSA_DADSPROF';
ChampsModelePSA[7]:= 'PSA_UNITETRAVAIL';
ChampsModelePSA[8]:= 'PSA_PRISEFFECTIF';

TExtractFille:= TExtract.FindFirst([''], [''], TRUE);
//TobDebug (TModele);
if (NomTable='SOCIETE') then
   begin
   for i:= 0 to 20 do
       AppliqueLeParam (TExtractFille, TModele, ChampsModeleSO[i]);
   end
else
if (NomTable='ETABLISS') then
   begin
   for i:= 0 to 1 do
       AppliqueLeParam (TExtractFille, TModele, ChampsModeleET[i]);
   end
else
if (NomTable='ETABCOMPL') then
   begin
   for i:= 0 to 61 do
       AppliqueLeParam (TExtractFille, TModele, ChampsModeleETB[i]);
   end
else
if (NomTable='SALARIES') then
   begin
   for i:= 0 to 8 do
       AppliqueLeParam (TExtractFille, TModele, ChampsModelePSA[i]);
   end;
end;
//FIN PT1-5


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 26/03/2007
Modifié le ... :   /  /
Description .. : Procédure de contrôle de cohérence de la TOB
Mots clefs ... : PAIE;PG_IMPORTASC
*****************************************************************}
function TOF_PG_IMPORTASC.ControleExtract (NomTable : string) : boolean;
var
TFille, TParamSoc, TParamSocFille : TOB;
i : integer;
begin
result:= True;
if (NomTable='SOCIETE') then
   begin
   TFille:= TExtract.FindFirst ([''], [''], True);
   TFille.PutValue ('SO_VERSIONBASE',
                    RecupChamp ('SOCIETE', 'SO_VERSIONBASE', '', result));
   TParamSoc:= TOB.Create ('PARAMSOC', nil, -1);
   TParamSoc.LoadDetailDB ('PARAMSOC', '', '', nil, False);
   For i:=0 to TFille.NombreChampReel do
       begin
       TParamSocFille:= TParamSoc.FindFirst (['SOC_NOM'], [TFille.GetNomChamp (i)], True);
       if ((TParamSocFille <> nil) and ((TFille.GetNomChamp (i)='SO_SOCIETE') or
          (TFille.GetNomChamp (i)='SO_LIBELLE') or
          (TFille.GetNomChamp (i)='SO_ADRESSE1') or
          (TFille.GetNomChamp (i)='SO_ADRESSE2') or
          (TFille.GetNomChamp (i)='SO_ADRESSE3') or
          (TFille.GetNomChamp (i)='SO_CODEPOSTAL') or
          (TFille.GetNomChamp (i)='SO_VILLE') or
          (TFille.GetNomChamp (i)='SO_PAYS') or
          (TFille.GetNomChamp (i)='SO_TELEPHONE') or
          (TFille.GetNomChamp (i)='SO_FAX') or
          (TFille.GetNomChamp (i)='SO_TELEX') or
          (TFille.GetNomChamp (i)='SO_APE') or
          (TFille.GetNomChamp (i)='SO_SIRET') or
          (TFille.GetNomChamp (i)='SO_DATEDEBUTEURO'))) then
          TParamSocFille.PutValue ('SOC_DATA', TFille.GetValeur (i));
       end;
   TParamSoc.InsertOrUpdateDB;
   FreeAndNil (TParamSoc);
   end;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 29/03/2007
Modifié le ... :   /  /
Description .. : Fonction de Controle d'existence des cumuls importés
Mots clefs ... : PAIE;PG_IMPORTASC
*****************************************************************}
function TOF_PG_IMPORTASC.ControleCumulPresent : boolean;
Var
StSQL : string;
TCumul, TCumulD, TFille, TInconnu, TPresent : TOB;
begin
StSQL:= 'SELECT * FROM CUMULPAIE WHERE ##PCL_PREDEFINI##';
TInconnu:= TOB.Create('CUMULPAIE', NIL, -1);
TPresent:= TOB.Create('CUMULPAIE', NIL, -1);
TCumul:= TOB.Create('CUMULPAIE', NIL, -1);
{Optimisation PT1-1
QRechCumul:= OpenSql(StSQL,TRUE);
TCumul.LoadDetailDB('CUMULPAIE','','',QRechCumul,False);
Ferme(QRechCumul);
}
TCumul.LoadDetailDBFromSQL ('CUMULPAIE', StSQL);
//FIN PT1-1
TFille:= TExtract.FindFirst ([''], [''], True);
While (TFille<>nil) do
      begin
      TCumulD:= TCumul.FindFirst (['PCL_CUMULPAIE'],
                                  [TFille.GetValue ('PCL_CUMULPAIE')], True);
      if (TCumulD<>nil) then
         TFille.ChangeParent(TPresent, -1)
      else
         TFille.ChangeParent(TInconnu, -1);
      TFille:= TExtract.FindNext ([''], [''], True);
      end;

if (TInconnu.Detail.Count<>0) then
   begin
   TFille:= TInconnu.FindFirst ([''], [''], True);
   While (TFille<>nil) do
         begin
         TFille.ChangeParent(TExtract, -1);
         TFille:= TInconnu.FindNext ([''], [''], True);
         end;
   result:= True;
   end
else
   result:= False;
//TobDebug (TExtract);
FreeAndNil (TCumul);
FreeAndNil (TInconnu);
FreeAndNil (TPresent);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 26/03/2007
Modifié le ... :   /  /
Description .. : Fonction de reprise d'un champ de la base
Mots clefs ... : PAIE;PG_IMPORTASC
*****************************************************************}
function TOF_PG_IMPORTASC.RecupChamp (Fichier, Champ, Where : String;
                                      var Trouve : boolean) : Variant;
Var
Q : TQuery;
SQL : String;
BEGIN
SQL:= 'SELECT '+Champ+' FROM '+Fichier+' '+Where;
Q:= OpenSQL(SQL,TRUE);
Trouve:= (not Q.EOF);
if (not Q.Eof) then
   result:= Q.Fields[0].Value;
Ferme (Q);
END;

{***********A.G.L.***********************************************
Auteur  ...... : FC
Créé le ...... : 29/03/2007
Modifié le ... :   /  /
Description .. : Traitement du fichier CUMULPAIE servant à alimenter la 
Suite ........ : table du même nom
Mots clefs ... : PAIE;PG_IMPORTASC
*****************************************************************}
procedure TOF_PG_IMPORTASC.TraitementCumulPaie ();
var
  TCumuls,TC : TOB;
  GCumuls : THGrid;
  i : integer;
  CBCUMULS : THValComboBox;
  StARecup : String;
begin
  FREPORT.Items.add(TraduireMemoire('Traitement du fichier CUMULPAIE'));

  CBCUMULS := THValComboBox(GetControl('CBCUMULS'));

  // Récupérer les cumuls dans la grille, ce sont ceux à insérer
  GCumuls := THGrid(GetControl('GCUMULS'));
  if (GCumuls <> nil) and (GCumuls.RowCount <> 1) then
  begin
    TCumuls := TOB.Create('CUMULPAIE', Nil, -1);
{warning PT1-3
    TC := TOB.Create('CUMULPAIE', TCumuls, -1);
}    
    for i := 1 to (GCumuls.RowCount - 1) do
    begin
      StARecup := GCumuls.CellValues[2,i];
      if (StARecup <> '') then
      begin
        TC := TExtract.FindFirst(['PCL_CUMULPAIE'], [GCumuls.CellValues[0,i]], False);
        // Faire les controles nécessaire sur le cumul en cours
        if ControleValCumulPaie(GCumuls.CellValues[0,i],CBCUMULS.Value) then
        begin
          // Renseigner les valeurs par défaut des champs
          EnregCumulPaieDefaut(TC);

          // Le prédéfini est récupéré dans la combo à l'écran
          if (CBCUMULS <> nil) then
          begin
            TC.PutValue('PCL_PREDEFINI', CBCUMULS.Value);
            if (CBCUMULS.Value = 'STD') then
              TC.PutValue('PCL_NODOSSIER', '000000')
            else
              TC.PutValue('PCL_NODOSSIER', V_PGI.NoDossier);
          end
          else
          begin
            TC.PutValue('PCL_PREDEFINI', 'DOS');
            TC.PutValue('PCL_NODOSSIER', V_PGI.NoDossier);
          end;

          TC.PutValue('PCL_CUMULPAIE', GCumuls.CellValues[0,i]);
          TC.PutValue('PCL_LIBELLE', GCumuls.CellValues[1,i]);
          // Insérer le cumul dans la base
          TC.InsertOrUpdateDB;
        end;
      end;
    end;
    FreeAndNil(TCumuls);
  end;
end;

// Controle des valeurs du fichier
function TOF_PG_IMPORTASC.ControleValCumulPaie (Cumul,Predefini:String) : boolean;
begin
  Result := False;

  if length(Cumul) < 2 then
  begin
    FREPORT.Items.add(TraduireMemoire('ERREUR : Cumul ' + Cumul + ' : Le code cumul doit comporter 2 caractères'));
    Erreur := True;
  end;

  if (Predefini = 'DOS') then
  begin
    if (StrToInt(RightStr(Cumul,1)) <> 5) and (StrToInt(RightStr(Cumul,1)) <> 7) and (StrToInt(RightStr(Cumul,1)) <> 9) then
    begin
      FREPORT.Items.add(TraduireMemoire('ERREUR : Cumul ' + Cumul + ' : Pour un prédéfini Dossier, le code cumul doit être impair terminant par 5, 7 ou 9 et supérieur à 50'));
      Erreur := True;
    end;
   end
else
  if (Predefini = 'STD') then
  begin
    if (StrToInt(RightStr(Cumul,1)) <> 1) and (StrToInt(RightStr(Cumul,1)) <> 3) then
    begin
      FREPORT.Items.add(TraduireMemoire('ERREUR : Cumul ' + Cumul + ' : Pour un prédéfini Standard, le code cumul doit être impair terminant par 1 ou 3 et supérieur à 50'));
      Erreur := True;
    end;
   end
else
if (Predefini='') then
   begin
   Erreur:= True;
   exit;
  end;

  Result := True;
end;

// Alimentation des valeurs par défaut non fournies
procedure TOF_PG_IMPORTASC.EnregCumulPaieDefaut (TC : TOB);
begin
  if not TC.IsFieldModified('PCL_THEMECUM') then TC.PutValue('PCL_THEMECUM', '001');
  if not TC.IsFieldModified('PCL_TYPECUMUL') then TC.PutValue('PCL_TYPECUMUL', 'X');
  if not TC.IsFieldModified('PCL_RAZCUMUL') then TC.PutValue('PCL_RAZCUMUL', '00');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 14/05/2007
Modifié le ... :   /  /    
Description .. : Traitement du fichier NOMTABLE servant à alimenter la
Suite ........ : table du même nom
Mots clefs ... : PAIE;PG_IMPORTASC
*****************************************************************}
procedure TOF_PG_IMPORTASC.TraitementTable (NomTable : string);
var
i:integer;
TTable, TTableFille : TOB;
ControleOK : Boolean;
StReq : string;
begin
FREPORT.Items.add (' ');
FREPORT.Items.add(TraduireMemoire ('Traitement du fichier '+NomTable));

if (TExtract.Detail.Count <> 0) then
   begin
   TTable:= TOB.Create (NomTable, nil, -1);
{warning
   TTableFille:= TOB.Create(NomTable, TTable, -1);
}

   if (NomTable='HISTOCUMSAL') then
      TExtract.Detail.Sort ('PHC_ETABLISSEMENT;PHC_SALARIE')
   else
   if (NomTable='ABSENCESALARIE') then
      begin
      TExtract.Detail.Sort ('PCN_SALARIE;PCN_ORDRE');

// Chargement des motifs concernant le type de congés
      Tob_MotifAbs:= tob.Create ('tob_motifabs', nil, -1);
{Optimisation PT1-1
      Q:= OpenSQL ('SELECT PMA_MOTIFABSENCE,PMA_JOURHEURE,PMA_JRSMAXI'+
                   ' FROM MOTIFABSENCE WHERE'+
                   ' ##PMA_PREDEFINI##', True);
      Tob_MotifAbs.LoadDetailDB ('MOTIFABSENCE', '', 'PMA_MOTIFABSENCE', Q,
                                 False);
      Ferme (Q);
}
      StReq:= 'SELECT PMA_MOTIFABSENCE,PMA_JOURHEURE,PMA_JRSMAXI'+
              ' FROM MOTIFABSENCE WHERE'+
              ' ##PMA_PREDEFINI##';
      Tob_MotifAbs.LoadDetailDBFromSQL ('MOTIFABSENCE', StReq);
//FIN PT1-1
      end
   else
   if (NomTable='SALARIES') then
      TExtract.Detail.Sort ('PSA_SALARIE');

   for i := 0 to (TExtract.Detail.Count -1) do
       begin
       TTableFille:= TExtract.detail[i];
       ControleOK:= False;
//Faire les controles nécessaire sur l'enregistrement en cours
       if (NomTable='HISTOCUMSAL') then
//PT4
          begin
          ControleOK:= ControleValHistoCumul(TTableFille);
          ChangeCodeSalarie (TTableFille, 'PHC', False);
          end
//FIN PT4
       else
       if (NomTable='ABSENCESALARIE') then
          begin
          ControleOK:= ControleValAbsenceSalarie(TTableFille);
          if (ControleOK) then
// Renseigner les valeurs par défaut des champs
             EnregAbsenceSalarieDefaut(TTableFille);
          end
       else
       if (NomTable='SALARIES') then
          begin
          ControleOK:= ControleValSalaries(TTableFille);
          if (ControleOK) then
// Renseigner les valeurs par défaut des champs
             EnregSalDefaut(TTableFille);
          end
       else
       if (NomTable='ETABLISS') then
          begin
          ControleOK:= ControleValEtabliss(TTableFille);
          if (ControleOK) then
// Renseigner les valeurs par défaut des champs
             EnregEtablissDefaut(TTableFille);
          end
       else
       if (NomTable='ETABCOMPL') then
          begin
          ControleOK:= ControleValEtabCompl(TTableFille);
          if (ControleOK) then
// Renseigner les valeurs par défaut des champs
             EnregEtabComplDefaut(TTableFille);
          end
//PT4
       else
       if (NomTable='RIB') then
          begin
//Transformation du code salarié en code tiers + alimentation enregistrement TIERS
          ControleOK:= ControleValRib(TTableFille);
          if (ControleOK) then
// Renseigner les valeurs par défaut des champs + Création enregistrement TIERS
             EnregRibDefaut(TTableFille);
          end
       else
       if (NomTable='CONTRATTRAVAIL') then
          begin
//Transformation du code salarié
          ControleOK:= ControleValContrat(TTableFille);
          if (ControleOK) then
// Renseigner les valeurs par défaut des champs
             EnregContratDefaut(TTableFille);
          end
       else
       if (NomTable='ENFANTSALARIE') then
          begin
//Transformation du code salarié
          ControleOK:= ControleValEnfant(TTableFille);
          if (ControleOK) then
// Renseigner les valeurs par défaut des champs
             EnregEnfantDefaut(TTableFille);
          end
       else
       if (NomTable='HISTOBULLETIN') then
          begin
//Transformation du code salarié
          ControleOK:= ControleValHistoBull(TTableFille);
          if (ControleOK) then
// Renseigner les valeurs par défaut des champs
             EnregHistoBullDefaut(TTableFille);
          end
       else
          ControleOK:= True;   
//FIN PT4
       if (ControleOK) then
// Insérer l'établissement dans la base
          TTableFille.InsertOrUpdateDB;
       end;
   FreeAndNil (TTable);
   if (NomTable='ABSENCESALARIE') then
      FreeAndNil (Tob_MotifAbs);
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : FC
Créé le ...... : 02/04/2007
Modifié le ... :
Description .. : Controle des valeurs du fichier ETABLISS
Mots clefs ... : PAIE;PG_IMPORTASC
*****************************************************************}
function TOF_PG_IMPORTASC.ControleValEtabliss (TEtablissFille : TOB) : boolean;
var
St : string;
begin
// A priori renvoie toujours Vrai car on ne fait que signaler les anomalies, rien n'est bloquant
Result:= False;

// Le code établissement doit être unique
{PT1-1 - Optimisation
Q:= OpenSQL ('SELECT ET_ETABLISSEMENT'+
             ' FROM ETABLISS WHERE'+
             ' ET_ETABLISSEMENT="'+TEtablissFille.getValue ('ET_ETABLISSEMENT')+'"',
             True);
if not Q.eof then
   begin
   FREPORT.Items.add (TraduireMemoire ('ERREUR : L''établissement '+
                                       TEtablissFille.getValue ('ET_ETABLISSEMENT')+
                                       ' existe déjà'));
   Erreur:= True;
   end;
Ferme (Q);
}
St:= 'SELECT ET_ETABLISSEMENT'+
     ' FROM ETABLISS WHERE'+
     ' ET_ETABLISSEMENT="'+TEtablissFille.getValue ('ET_ETABLISSEMENT')+'"';
if (ExisteSQL (St)) then
   begin
   FREPORT.Items.add (TraduireMemoire ('ERREUR : L''établissement '+
                                       TEtablissFille.getValue ('ET_ETABLISSEMENT')+
                                       ' existe déjà'));
   Erreur:= True;
   end;
//FIN PT1-1

// Le libellé de l'établissement doit être unique
{PT1-1 - Optimisation
Q:= OpenSQL ('SELECT ET_ETABLISSEMENT'+
             ' FROM ETABLISS WHERE'+
             ' ET_LIBELLE="'+TEtablissFille.getValue ('ET_LIBELLE')+'" AND'+
             ' ET_ETABLISSEMENT<>"'+TEtablissFille.getValue ('ET_ETABLISSEMENT')+'"',
             True);
if not Q.eof then
   begin
   FREPORT.Items.add (TraduireMemoire ('ERREUR : Il existe déjà un'+
                                       ' établissement avec le même libellé ('+
                                       TEtablissFille.getValue ('ET_ETABLISSEMENT')+
                                       ' '+TEtablissFille.getValue('ET_LIBELLE')+')'));
   Erreur := True;
   end;
Ferme (Q);
}
St:='SELECT ET_ETABLISSEMENT'+
    ' FROM ETABLISS WHERE'+
    ' ET_LIBELLE="'+TEtablissFille.getValue ('ET_LIBELLE')+'" AND'+
    ' ET_ETABLISSEMENT<>"'+TEtablissFille.getValue ('ET_ETABLISSEMENT')+'"';
if (ExisteSQL (St)) then
   begin
   FREPORT.Items.add (TraduireMemoire ('ERREUR : Il existe déjà un'+
                                       ' établissement avec le même libellé ('+
                                       TEtablissFille.getValue ('ET_ETABLISSEMENT')+
                                       ' '+TEtablissFille.getValue('ET_LIBELLE')+')'));
   Erreur := True;
   end;
//FIN PT1-1

Result:= True;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : FC
Créé le ...... : 02/04/2007
Modifié le ... :
Description .. : Alimentation des valeurs par défaut du fichier ETABLISS
Mots clefs ... : PAIE;PG_IMPORTASC
*****************************************************************}
procedure TOF_PG_IMPORTASC.EnregEtablissDefaut (TEtablissFille : TOB);
begin
  if not TEtablissFille.IsFieldModified('ET_DATELIBRE1') then TEtablissFille.PutValue('ET_DATELIBRE1', IDate1900);
  if not TEtablissFille.IsFieldModified('ET_DATELIBRE2') then TEtablissFille.PutValue('ET_DATELIBRE2', IDate1900);
  if not TEtablissFille.IsFieldModified('ET_DATELIBRE3') then TEtablissFille.PutValue('ET_DATELIBRE3', IDate1900);
  if not TEtablissFille.IsFieldModified('ET_SURSITE') then TEtablissFille.PutValue('ET_SURSITE', 'X');
  if not TEtablissFille.IsFieldModified('ET_NODOSSIER') then TEtablissFille.PutValue('ET_NODOSSIER', V_PGI.NoDossier);
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : FC
Créé le ...... : 02/04/2007
Modifié le ... :
Description .. : Controle des valeurs du fichier ETABCOMPL
Mots clefs ... : PAIE;PG_IMPORTASC
*****************************************************************}
function TOF_PG_IMPORTASC.ControleValEtabCompl(TEtabComplFille : TOB): boolean;
var
lib, St, StEtab : String;
begin
// A priori renvoie toujours Vrai car on ne fait que signaler les anomalies, rien n'est bloquant
Result:= False;

StEtab:= TEtabComplFille.getValue ('ETB_ETABLISSEMENT')+' '+
         TEtabComplFille.getValue ('ETB_LIBELLE');

// L'établissement doit exister dans ETABLISS
{PT1-1 - Optimisation
Q:= OpenSQL ('SELECT ET_ETABLISSEMENT'+
             ' FROM ETABLISS WHERE'+
             ' ET_ETABLISSEMENT="'+TEtabComplFille.getValue ('ETB_ETABLISSEMENT')+'"',
             True);
if Q.Eof then
   begin
   FREPORT.Items.add (TraduireMemoire ('ERREUR : Etablissement '+
                                       TEtabComplFille.getValue ('ETB_ETABLISSEMENT')+
                                       ' doit exister dans la table ETABLISS'+
                                       ' pour pouvoir saisir des informations'+
                                       ' complémentaires'));
   Erreur:= True;
   end;
Ferme (Q);
}
St:= 'SELECT ET_ETABLISSEMENT'+
     ' FROM ETABLISS WHERE'+
     ' ET_ETABLISSEMENT="'+TEtabComplFille.getValue ('ETB_ETABLISSEMENT')+'"';
if (ExisteSQL (St)) then
   begin
   FREPORT.Items.add (TraduireMemoire ('ERREUR : Etablissement '+
                                       TEtabComplFille.getValue ('ETB_ETABLISSEMENT')+
                                       ' doit exister dans la table ETABLISS'+
                                       ' pour pouvoir saisir des informations'+
                                       ' complémentaires'));
   Erreur:= True;
   end;
//FIN PT1-1

// Le code établissement doit être unique
{PT1-1 - Optimisation
Q:= OpenSQL ('SELECT ETB_ETABLISSEMENT'+
             ' FROM ETABCOMPL WHERE'+
             ' ETB_ETABLISSEMENT="'+TEtabComplFille.getValue ('ETB_ETABLISSEMENT')+'"',
             True);
if Q.Eof then
   FREPORT.Items.add (TraduireMemoire ('AVERTISSEMENT : Des informations'+
                                       ' complémentaires existent déjà pour'+
                                       ' l''établissement '+
                                       TEtabComplFille.getValue ('ETB_ETABLISSEMENT')));
Ferme (Q);
}
St:= 'SELECT ETB_ETABLISSEMENT'+
     ' FROM ETABCOMPL WHERE'+
     ' ETB_ETABLISSEMENT="'+TEtabComplFille.getValue ('ETB_ETABLISSEMENT')+'"';
if (ExisteSQL (St)) then
   FREPORT.Items.add (TraduireMemoire ('AVERTISSEMENT : Des informations'+
                                       ' complémentaires existent déjà pour'+
                                       ' l''établissement '+
                                       TEtabComplFille.getValue ('ETB_ETABLISSEMENT')));
//FIN PT1-1

// Le libellé de l'établissement doit être unique
{PT1-1 - Optimisation
Q:= OpenSQL ('SELECT ETB_ETABLISSEMENT'+
             ' FROM ETABCOMPL WHERE'+
             ' ETB_LIBELLE="'+TEtabComplFille.getValue ('ETB_LIBELLE')+'" AND'+
             ' ETB_ETABLISSEMENT<>"'+TEtabComplFille.getValue ('ETB_ETABLISSEMENT')+'"',
             True);
if Q.Eof then
   FREPORT.Items.add (TraduireMemoire ('AVERTISSEMENT : Il existe déjà des'+
                                       ' informations complémentaires pour un'+
                                       ' établissement ayant le même libellé ('+
                                       TEtabComplFille.getValue ('ETB_ETABLISSEMENT')+
                                       ' '+TEtabComplFille.getValue ('ETB_LIBELLE')+')'));
Ferme(Q);
}
St:= 'SELECT ETB_ETABLISSEMENT'+
     ' FROM ETABCOMPL WHERE'+
     ' ETB_LIBELLE="'+TEtabComplFille.getValue ('ETB_LIBELLE')+'" AND'+
     ' ETB_ETABLISSEMENT<>"'+TEtabComplFille.getValue ('ETB_ETABLISSEMENT')+'"';
if (ExisteSQL (St)) then
   FREPORT.Items.add (TraduireMemoire ('AVERTISSEMENT : Il existe déjà des'+
                                       ' informations complémentaires pour un'+
                                       ' établissement ayant le même libellé ('+
                                       StEtab+')'));
//FIN PT1-1

// ETB_SMIC Test au minimum(smic/conventionnel) si ETB_TYPSMIC saisi
if (TEtabComplFille.getValue ('ETB_TYPSMIC')<>'') and
   (TEtabComplFille.getValue ('ETB_SMIC')='') then
   FREPORT.Items.add (TraduireMemoire ('AVERTISSEMENT : Etablissement '+StEtab+
                                       ' : Il faudrait renseigner la valeur'+
                                       ' Test au minimum(smic/conventionnel)'));

// ETB_HORAIREETABL Horaire de référence
if (TEtabComplFille.getValue ('ETB_HORAIREETABL')='0') then
   begin
   FREPORT.Items.add (TraduireMemoire ('ERREUR : Etablissement '+StEtab+' :'+
                                       ' l''horaire de référence est'+
                                       ' obligatoire.'));
   Erreur:= True;
   end;

// Si ETB_CONGESPAYES coché alors faire les tests liés aux CP
if (TEtabComplFille.getValue ('ETB_CONGESPAYES')='X') then
   begin
   if (not GetParamSocSecur ('SO_PGCONGES', False)) then
      begin
      FREPORT.Items.add (TraduireMemoire ('ERREUR : Etablissement '+StEtab+' :'+
                                          ' Il faut cocher l''option gestion'+
                                          ' des congés payés au niveau dossier'+
                                          ' avant de pouvoir la positionner au'+
                                          ' niveau établissement.'));
      Erreur:= True;
      end;

// ETB_DATECLOTURECPN Date de cloture cp
   if not IsValidDate (string (TEtabComplFille.getValue ('ETB_DATECLOTURECPN'))) then
      begin
      FREPORT.Items.add (TraduireMemoire ('ERREUR : Etablissement '+StEtab+' :'+
                                          ' La date de clôture n''est pas'+
                                          ' valide.'));
      Erreur:= True;
      end;

// ETB_NBJOUTRAV Controle zone égale à 5 ou 6
   if ((Valeur (TEtabComplFille.getValue ('ETB_NBJOUTRAV'))>6) or
      (Valeur (TEtabComplFille.getValue ('ETB_NBJOUTRAV'))<5)) then
      begin
      FREPORT.Items.add (TraduireMemoire ('ERREUR : Etablissement '+StEtab+' :'+
                                          ' la zone ETB_NBJOUTRAV doit'+
                                          ' contenir la valeur 5 (ouvrés) ou'+
                                          ' (6) ouvrables.'));
      Erreur:= True;
      end;

// ETB_1ERREPOSH 1er jour de repos obligatoire
   if (TEtabComplFille.getValue ('ETB_1ERREPOSH')='') then
      begin
      FREPORT.Items.add (TraduireMemoire ('ERREUR : Etablissement '+StEtab+' :'+
                                          ' le 1er jour de repos hebdomadaire'+
                                          ' est obligatoire.'));
      Erreur:= True;
      end;

// ETB_2EMEREPOSH 2è jour de repos obligatoire si ETB_NBJOUTRAV = 5
   if (Valeur (TEtabComplFille.getValue ('ETB_NBJOUTRAV'))=5) and
      (TEtabComplFille.getValue ('ETB_2EMEREPOSH')='') then
      begin
      FREPORT.Items.add (TraduireMemoire ('ERREUR : Etablissement '+StEtab+' :'+
                                          ' le 2ème jour de repos hebdomadaire'+
                                          ' est obligatoire.'));
      Erreur:= True;
      end;

// ETB_1ERREPOSH doit être différent de ETB_2EMEREPOSH
   if (TEtabComplFille.getValue ('ETB_1ERREPOSH')=TEtabComplFille.getValue ('ETB_2EMEREPOSH')) then
      FREPORT.Items.add (TraduireMemoire ('AVERTISSEMENT : Etablissement '+
                                          StEtab+' : le 2ème jour de repos'+
                                          ' hebdomadaire doit être différent'+
                                          ' du 1er.'));

// ETB_RELIQUAT la méthode de gestion de reliquat est obligatoire
   if (TEtabComplFille.getValue ('ETB_RELIQUAT')='') then
      begin
      FREPORT.Items.add (TraduireMemoire ('ERREUR : Etablissement '+StEtab+' :'+
                                          ' la méthode de gestion de reliquat'+
                                          ' est obligatoire.'));
      Erreur:= True;
      end;

// ETB_VALORINDEMCP la méthode de valorisation des indemnité CP est obligatoire
   if (TEtabComplFille.getValue ('ETB_VALORINDEMCP')='') then
      begin
      FREPORT.Items.add (TraduireMemoire ('ERREUR : Etablissement '+StEtab+' :'+
                                          ' la méthode de valorisation des'+
                                          ' indemnité CP est obligatoire'));
      Erreur:= True;
      end;

// ETB_MVALOMS la valorisation du maintien est obligatoire
   if (TEtabComplFille.getValue ('ETB_MVALOMS')='') then
      begin
      FREPORT.Items.add (TraduireMemoire ('ERREUR : Etablissement '+StEtab+' :'+
                                          ' la méthode de calcul du maintien'+
                                          ' de salaire est obligatoire.'));
      Erreur:= True;
      end;

// ETB_NBREACQUISCP nb de jours acquis inférieur à 10
   if (Valeur (TEtabComplFille.getValue ('ETB_NBREACQUISCP'))>10) then
      begin
      FREPORT.Items.add (TraduireMemoire ('ERREUR : Etablissement '+StEtab+' :'+
                                          ' le nombre de jour acquis doit être'+
                                          ' inférieur à 10.'));
      Erreur:= True;
      end;

// ETB_DATEACQCPANC Date d'acquisition des cp ancienneté
   if (TEtabComplFille.getValue ('ETB_DATEACQCPANC')<>'') then
      if not PgOkFormatDateJJMM (Copy (TEtabComplFille.getValue ('ETB_DATEACQCPANC'), 1, 2)+
                                 Copy (TEtabComplFille.getValue ('ETB_DATEACQCPANC'), 3, 2)) then
         begin
         FREPORT.Items.add (TraduireMemoire ('ERREUR : Etablissement '+StEtab+
                                             ' : la date Date d''acquisition'+
                                             ' des CP anciennetée doit être au'+
                                             ' format jj/mm.'));
         Erreur:= True;
         end;
   end;

// Date Validité Transport ETB_DATEVALTRANS
if not IsValidDate (string (TEtabComplFille.getValue ('ETB_DATEVALTRANS'))) then
   begin
   FREPORT.Items.add (TraduireMemoire ('ERREUR : Etablissement '+StEtab+' : La'+
                                       ' date de validité transport n''est pas'+
                                       ' valide.'));
   Erreur:= True;
   end;

// ETB_JOURPAIEMENT jour de paiement compris entre 1 et 31
if (TEtabComplFille.getValue ('ETB_JOURPAIEMENT')<1) or
   (TEtabComplFille.getValue ('ETB_JOURPAIEMENT')>31) then
   begin
   FREPORT.Items.add (TraduireMemoire ('ERREUR : Etablissement '+StEtab+' : Le'+
                                       ' jour de paiement doit être compris'+
                                       ' entre 1 et 31.'));
   Erreur:= True;
   end;

// ETB_JOURHEURE Edition du calendrier en jour ou heure
if (TEtabComplFille.getValue ('ETB_JOURHEURE')<>'JOU') and
   (TEtabComplFille.getValue ('ETB_JOURHEURE')<>'HEU') then
   FREPORT.Items.add (TraduireMemoire ('AVERTISSEMENT : Etablissement '+StEtab+
                                       ' : Le mode d''édition du calendrier'+
                                       ' doit être en jour ou heure.'));

// Vérification présence valeurs dans les paramètres tablettes
if VerifPresenceParam then
   begin
// ETB_SMIC Test au minimum(smic/conventionnel) si ETB_TYPSMIC saisi
   if (TEtabComplFille.getValue ('ETB_TYPSMIC')='ELN') then
      begin
      Lib:= RechDom ('PGELEMENTNAT', TEtabComplFille.getValue ('ETB_SMIC'),
                     false);
      if (Lib='') or (Lib='Error') then
         FREPORT.Items.add (TraduireMemoire ('AVERTISSEMENT : Etablissement '+
                                             StEtab+' : La valeur Test au'+
                                             ' minimum(smic/conventionnel)'+
                                             ' n''existe pas dans le'+
                                             ' paramétrage'));
      end
   else
   if (TEtabComplFille.getValue ('ETB_TYPSMIC')='VAR') then
      begin
      Lib:= RechDom ('PGVARIABLE', TEtabComplFille.getValue ('ETB_SMIC'),
                     false);
      if (Lib='') or (Lib='Error') then
         FREPORT.Items.add (TraduireMemoire ('AVERTISSEMENT : Etablissement '+
                                             StEtab+' : La valeur Test au'+
                                             ' minimum(smic/conventionnel)'+
                                             ' n''existe pas dans le'+
                                             ' paramétrage'));
      end;

// Si ETB_CONGESPAYES coché alors faire les tests liés aux CP
   if (TEtabComplFille.getValue ('ETB_CONGESPAYES')='X') then
      begin
// ETB_NBACQUISCP Nombre de jours CP acquis par mois
      VerifTablette ('ETB_NBACQUISCP', 'PGVARIABLE', 'Etablissement '+StEtab,
                     'La variable permettant la détermination du Nombre de'+
                     ' jours CP acquis par mois', TEtabComplFille);

// ETB_1ERREPOSH 1er jour de repos obligatoire
      VerifTablette ('ETB_1ERREPOSH', 'YYJOURSSEMAINE', 'Etablissement '+StEtab,
                     'Le type du 1er jour de repos hebdomadaire',
                     TEtabComplFille);

// ETB_2EMEREPOSH 2è jour de repos obligatoire si ETB_NBJOUTRAV = 5
      VerifTablette ('ETB_2EMEREPOSH', 'YYJOURSSEMAINE',
                     'Etablissement '+StEtab,
                     'Le type du 2ème jour de repos hebdomadaire',
                     TEtabComplFille);

// ETB_RELIQUAT la méthode de gestion de reliquat est obligatoire
      VerifTablette ('ETB_RELIQUAT', 'PGRELIQUATCP', 'Etablissement '+StEtab,
                     'La méthode de gestion de reliquat', TEtabComplFille);

// ETB_VALORINDEMCP la méthode de valorisation des indemnité CP est obligatoire
      VerifTablette ('ETB_VALORINDEMCP', 'PGVALORINDEMCP',
                     'Etablissement '+StEtab,
                     'La méthode de valorisation des indemnités CP',
                     TEtabComplFille);

// ETB_MVALOMS la valorisation du maintien est obligatoire
      VerifTablette ('ETB_MVALOMS', 'PGVALOMS', 'Etablissement '+StEtab,
                     'Le type de valorisation du maintien', TEtabComplFille);

// ETB_PAIEVALOMS valeur présente dans la tablette PGPAIEVALOMS
      VerifTablette ('ETB_PAIEVALOMS', 'PGPAIEVALOMS', 'Etablissement '+StEtab,
                     'Le type de salaire retenu pour le maintien',
                     TEtabComplFille);

// ETB_EDITBULCP valeur présente dans la tablette PGEDITBULCP
      VerifTablette ('ETB_EDITBULCP', 'PGEDITBULCP', 'Etablissement '+StEtab,
                     'Le type d''édition des compteurs sur bulletin',
                     TEtabComplFille);

// ETB_PROFILCGE profil congés payés
      VerifTablette ('ETB_PROFILCGE', 'PGPROFILCGE', 'Etablissement '+StEtab,
                     'Le type de profil congés payés', TEtabComplFille);
      end;

// ETB_PERIODBUL Plafond bulletin
   VerifTablette ('ETB_PERIODBUL', 'PGPROFILPERIODE', 'Etablissement '+StEtab,
                  'Le type de plafond bulletin', TEtabComplFille);

// ETB_PROFILRBS profil Réduction loi Fillon
   VerifTablette ('ETB_PROFILRBS', 'PGPROFILRBS', 'Etablissement '+StEtab,
                  'Le type de profil Réduction loi Fillon', TEtabComplFille);

// ETB_PROFILTRANS profil transport
   VerifTablette ('ETB_PROFILTRANS', 'PGPROFILTRANS', 'Etablissement '+StEtab,
                  'Le type de profil de transport', TEtabComplFille);

// ETB_CONVENTION
   VerifTablette ('ETB_CONVENTION', 'PGCONVENTION', 'Etablissement '+StEtab,
                  'Le type de convention', TEtabComplFille);

// ETB_PGMODEREGLE mode de réglément
   VerifTablette ('ETB_PGMODEREGLE', 'PGMODEREGLE', 'Etablissement '+StEtab,
                  'Le mode de réglement', TEtabComplFille);

// ETB_MOISPAIEMENT mois de paiement
   VerifTablette ('ETB_MOISPAIEMENT', 'PGMOISPAIEMENT', 'Etablissement '+StEtab,
                  'Le mois de paiement', TEtabComplFille);

// ETB_PAIACOMPTE mode de réglement des acomptes
   VerifTablette ('ETB_PAIACOMPTE', 'PGMODEREGLE', 'Etablissement '+StEtab,
                  'Le mode de réglement des acomptes', TEtabComplFille);

// ETB_PAIFRAIS mode de réglement des frais professionnels
   VerifTablette ('ETB_PAIFRAIS', 'PGMODEREGLE', 'Etablissement '+StEtab,
                  'Le mode de réglement des frais professionnels',
                  TEtabComplFille);

// ETB_DADSSECTION Section d'établissement DADS Bilatérale
   VerifTablette ('ETB_DADSSECTION', 'PGDADSSECTION', 'Etablissement '+StEtab,
                  'La section d''établissement DADS Bilatérale',
                  TEtabComplFille);

// ETB_TYPDADSSECT type de section d'établissement DADS Bilatérale
   VerifTablette ('ETB_TYPDADSSECT', 'PGDADSTYPE', 'Etablissement '+StEtab,
                  'Le type de section d''établissement DADS Bilatérale',
                  TEtabComplFille);

// ETB_PRUDH section prud'hommale établissement
   VerifTablette ('ETB_PRUDH', 'PGPRUDH', 'Etablissement '+StEtab,
                  'La section prud''hommale établissement', TEtabComplFille);

// ETB_PRUDHCOLL Collège prud'hommal
   VerifTablette ('ETB_PRUDHCOLL', 'PGCOLLEGEPRUD', 'Etablissement '+StEtab,
                  'Le Collège prud''hommal', TEtabComplFille);

// ETB_PRUDHSECT Section prud'homamle
   VerifTablette ('ETB_PRUDHSECT', 'PGSECTIONPRUD', 'Etablissement '+StEtab,
                  'La Section prud''hommale', TEtabComplFille);

// ETB_PRUDHVOTE Lieu de vote prud'hommal
   VerifTablette ('ETB_PRUDHVOTE', 'PGLIEUVOTEPRUD', 'Etablissement '+StEtab,
                  'Le Lieu de vote pour les prud''hommes', TEtabComplFille);
   end;

Result:= True;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : FC
Créé le ...... : 02/04/2007
Modifié le ... :   /  /
Description .. : Alimentation des valeurs par défaut non fournies du fichier
Suite ........ : ETABCOMPL
Mots clefs ... : PAIE;PG_IMPORTASC
*****************************************************************}
procedure TOF_PG_IMPORTASC.EnregEtabComplDefaut(TEtabComplFille: TOB);
var
  DateCloture : Tdatetime;
  ANow, MNow, DNow : Word;
begin
  if not TEtabComplFille.IsFieldModified('ETB_DATECLOTURECPN') then
  begin
    // Déterminer la date de cloture CP
    Decodedate(date, ANow, MNow, DNow);
    DateCloture := EncodeDate(ANow, 05, 31);
    if DateCloture <= Now then DateCloture := EncodeDate(ANow + 1, 05, 31);
    TEtabComplFille.PutValue('ETB_DATECLOTURECPN', DateCloture);
  end;
  if not TEtabComplFille.IsFieldModified('ETB_DATEVALTRANS') then TEtabComplFille.PutValue('ETB_DATEVALTRANS', iDate1900);
  if not TEtabComplFille.IsFieldModified('ETB_JOURHEURE') then TEtabComplFille.PutValue('ETB_JOURHEURE', 'HEU');
  if not TEtabComplFille.IsFieldModified('ETB_CODESECTION') then TEtabComplFille.PutValue('ETB_CODESECTION', '1');
  if not TEtabComplFille.IsFieldModified('ETB_PRUDHCOLL') then TEtabComplFille.PutValue('ETB_PRUDHCOLL', '1');
  if not TEtabComplFille.IsFieldModified('ETB_PRUDHSECT') then TEtabComplFille.PutValue('ETB_PRUDHSECT', '4');
  if not TEtabComplFille.IsFieldModified('ETB_PRUDHVOTE') then TEtabComplFille.PutValue('ETB_PRUDHVOTE', '1');
  if not TEtabComplFille.IsFieldModified('ETB_MEDTRAV') then TEtabComplFille.PutValue('ETB_MEDTRAV', -1);
  if not TEtabComplFille.IsFieldModified('ETB_CODEDDTEFP') then TEtabComplFille.PutValue('ETB_CODEDDTEFP', -1);
end;

procedure TOF_PG_IMPORTASC.MajChampOblig(LibChamp : String);
begin
  NbChampOblig := NbChampOblig + 1;
  if (ChampOblig <> '') then
    ChampOblig := ChampOblig + ', ' + LibChamp
  else
    ChampOblig := LibChamp;
end;

{***********A.G.L.***********************************************
Auteur  ...... : FC
Créé le ...... : 03/04/2007
Modifié le ... :   /  /
Description .. : Controle des valeurs du fichier SALARIES
Mots clefs ... : PAIE;PG_IMPORTASC
*****************************************************************}
function TOF_PG_IMPORTASC.ControleValSalaries(TSalariesFille : TOB): boolean;
var
Civilite, CleSexe, Lib, Lib1, Lib2, Lib3, Lib4, LibAvetSal, LibErrSal : String;
LibSal, NomJFille, NoSS, Numero, Sexe, St : String;
Nombre, Num, Resultat : integer;
DEntree,DNaissance : TDateTime;
const
Voyelle = ['a', 'e', 'i', 'u', 'y', 'o', 'A', 'E', 'I', 'U', 'Y', 'O'];
begin
// A priori renvoie toujours Vrai car on ne fait que signaler les anomalies, rien n'est bloquant
Result := False;

LibSal:= 'Salarié '+TSalariesFille.getValue ('PSA_SALARIE')+' '+
         TSalariesFille.getValue ('PSA_LIBELLE')+' '+
         TSalariesFille.getValue('PSA_PRENOM');
LibAvetSal:= 'AVERTISSEMENT : '+LibSal;
LibErrSal:= 'ERREUR : '+LibSal;
// Rajouter les éventuels zéros de gauche dans le numéro du salarié
if (PGTypeNumSal = 'NUM') and
   (isnumeric (TSalariesFille.getValue ('PSA_SALARIE'))) and
   (TSalariesFille.getValue ('PSA_SALARIE') <> '') and
   (Length(TSalariesFille.getValue ('PSA_SALARIE')) < 10) then
   TSalariesFille.PutValue ('PSA_SALARIE',
                            ColleZeroDevant (StrToInt (trim (TSalariesFille.getValue ('PSA_SALARIE'))), 10));

// Le salarié ne doit pas déjà exister
{PT1-1 - Optimisation
Q:= OpenSQL ('SELECT PSA_SALARIE'+
             ' FROM SALARIES WHERE'+
             ' PSA_SALARIE="'+TSalariesFille.getValue ('PSA_SALARIE')+'"', True);
if not Q.Eof then
   begin
   FREPORT.Items.add (TraduireMemoire (LibErrSal+' déjà présent'));
   Erreur:= True;
   end;
Ferme(Q);
}
St:= 'SELECT PSA_SALARIE'+
     ' FROM SALARIES WHERE'+
     ' PSA_SALARIE="'+TSalariesFille.getValue ('PSA_SALARIE')+'"';
if (ExisteSQL (St)) then
   begin
   FREPORT.Items.add (TraduireMemoire (LibErrSal+' déjà présent'));
   Erreur:= True;
   end;
//FIN PT1-1

// L'établissement doit exister
{PT1-1 - Optimisation
Q:= OpenSQL ('SELECT ET_ETABLISSEMENT'+
             ' FROM ETABLISS WHERE'+
             ' ET_ETABLISSEMENT="'+TSalariesFille.getValue ('PSA_ETABLISSEMENT')+'"', True);
if Q.Eof then
   begin
   FREPORT.Items.add (TraduireMemoire (LibErrSal+' : L''établissement'+
                                       ' n''existe pas.'));
   Erreur:= True;
   end
else
   begin
   Ferme(Q);
// L'établissement doit comporter des informations sur le social ETABCOMPL
   Q:= OpenSQL ('SELECT ETB_ETABLISSEMENT'+
                ' FROM ETABCOMPL WHERE'+
                ' ETB_ETABLISSEMENT = "'+TSalariesFille.getValue('PSA_ETABLISSEMENT')+'"', True);
   if Q.Eof then
      begin
      FREPORT.Items.add (TraduireMemoire (LibErrSal+' : L''établissement ne'+
                                          ' comporte pas d''informations sur'+
                                          ' le social.'));
      Erreur:= True;
      end;
   end;
Ferme(Q);
}
St:= 'SELECT ET_ETABLISSEMENT'+
     ' FROM ETABLISS WHERE'+
     ' ET_ETABLISSEMENT="'+TSalariesFille.getValue ('PSA_ETABLISSEMENT')+'"';
if (ExisteSQL (St)=False) then
   begin
   FREPORT.Items.add (TraduireMemoire (LibErrSal+' : L''établissement'+
                                       ' n''existe pas.'));
   Erreur:= True;
   end
else
   begin
// L'établissement doit comporter des informations sur le social ETABCOMPL
   St:= 'SELECT ETB_ETABLISSEMENT'+
        ' FROM ETABCOMPL WHERE'+
        ' ETB_ETABLISSEMENT="'+TSalariesFille.getValue('PSA_ETABLISSEMENT')+'"';
   if (ExisteSQL (St)=False) then
      begin
      FREPORT.Items.add (TraduireMemoire (LibErrSal+' : L''établissement ne'+
                                          ' comporte pas d''informations sur'+
                                          ' le social.'));
      Erreur:= True;
      end;
   end;
//FIN PT1-1

// PSA_TYPNBACQUISCP
if (TSalariesFille.getValue ('PSA_TYPNBACQUISCP') = 'PER') then
   begin
   Lib:= RechDom ('PGVARIABLE', TSalariesFille.getValue ('PSA_NBACQUISCP'), FALSE);
   if (Lib = '') or (Lib = 'Error') then
      begin
      FREPORT.Items.add (TraduireMemoire (LibErrSal+' : Le champ jour acquis'+
                                          ' bulletin doit être renseigné avec'+
                                          ' une variable existante.'));
      Erreur:= True;
      end;
   end;

// Zones obligatoires
ChampOblig:= '';
NbChampOblig:= 0;
if (TSalariesFille.getValue ('PSA_SALARIE') <> '') and
   (length (TSalariesFille.getValue ('PSA_SALARIE'))<11) then
   begin
   if (TSalariesFille.getValue ('PSA_ETABLISSEMENT')='') then
      MajChampOblig ('l''établissement');
   if (TSalariesFille.getValue ('PSA_LIBELLE')='') then
      MajChampOblig ('le nom');
   if (TSalariesFille.getValue ('PSA_PRENOM')='') then
      MajChampOblig ('le prénom');
   if (TSalariesFille.getValue ('PSA_CIVILITE')='') then
      MajChampOblig ('la civilité');
   if (TSalariesFille.getValue ('PSA_SEXE')='') then
      MajChampOblig ('le sexe');
   if (TSalariesFille.getValue ('PSA_ADRESSE1')='') then
      MajChampOblig ('l''adresse');
   if (TSalariesFille.getValue ('PSA_CODEPOSTAL')='') then
      MajChampOblig ('le code postal');
   if (TSalariesFille.getValue ('PSA_VILLE')='') then
      MajChampOblig ('la ville');
   if (TSalariesFille.getValue ('PSA_DATEENTREE')=iDate1900) then
      MajChampOblig ('la date d''entrée');
   end;

Lib:= GetParamSocSecur ('SO_PGLIBCODESTAT', '');
if ((TSalariesFille.getValue ('PSA_CODESTAT')='') and (Lib<>'')) then
   begin
   if (Lib[1] in Voyelle) then
      MajChampOblig ('l''' + Lib)
   else
      MajChampOblig ('le ' + Lib);
   end;

Nombre:= GetParamSocSecur ('SO_PGNBRESTATORG', 0);
if (Nombre <> 0) then
   begin
   Lib1:= GetParamSocSecur ('SO_PGLIBORGSTAT1', '');
   Lib2:= GetParamSocSecur ('SO_PGLIBORGSTAT2', '');
   Lib3:= GetParamSocSecur ('SO_PGLIBORGSTAT3', '');
   Lib4:= GetParamSocSecur ('SO_PGLIBORGSTAT4', '');
   for Num := 1 to Nombre do
       begin
       Numero:= InttoStr (Num);
       if (Num > Nombre) then
          exit;
       if (TSalariesFille.getValue ('PSA_TRAVAILN'+Numero)='') then
          begin
          if ((Num = 1) and (Lib1 <> '')) then
             begin
             if (Lib1[1] in Voyelle) then
                MajChampOblig ('l'''+Lib1)
             else
                MajChampOblig ('le '+Lib1);
             end;
          if ((Num = 2) and (Lib2 <> '')) then
             begin
             if (Lib2[1] in Voyelle) then
                MajChampOblig ('l'''+Lib2)
             else
                MajChampOblig ('le '+Lib2);
             end;
          if ((Num = 3) and (Lib3 <> '')) then
             begin
             if (Lib3[1] in Voyelle) then
                MajChampOblig ('l'''+Lib3)
             else
                MajChampOblig ('le '+Lib3);
             end;
          if ((Num = 4) and (Lib4 <> '')) then
             begin
             if (Lib4[1] in Voyelle) then
                MajChampOblig ('l'''+Lib4)
             else
                MajChampOblig ('le '+Lib4);
             end;
          end;
       end;
   end;

Nombre:= GetParamSocSecur ('SO_PGNBCOMBO', 0);
if (Nombre <> 0) then
   begin
   Lib1:= GetParamSocSecur ('SO_PGLIBCOMBO1', '');
   Lib2:= GetParamSocSecur ('SO_PGLIBCOMBO2', '');
   Lib3:= GetParamSocSecur ('SO_PGLIBCOMBO3', '');
   Lib4:= GetParamSocSecur ('SO_PGLIBCOMBO4', '');

   for Num := 1 to Nombre do
       begin
       Numero:= InttoStr (Num);
       if (Num > Nombre) then
          exit;
       if (TSalariesFille.getValue ('PSA_LIBREPCMB'+Numero)='') then
          begin
          if ((Num = 1) and (Lib1 <> '')) then
             begin
             if (Lib1[1] in Voyelle) then
                MajChampOblig ('l'''+Lib1)
             else
                MajChampOblig ('le '+Lib1);
             end;
          if ((Num = 2) and (Lib2 <> '')) then
             begin
             if (Lib2[1] in Voyelle) then
                MajChampOblig ('l'''+Lib2)
             else
                MajChampOblig ('le '+Lib2);
             end;
          if ((Num = 3) and (Lib3 <> '')) then
             begin
             if (Lib3[1] in Voyelle) then
                MajChampOblig ('l'''+Lib3)
             else
                MajChampOblig ('le '+Lib3);
             end;
          if ((Num = 4) and (Lib4 <> '')) then
             begin
             if (Lib4[1] in Voyelle) then
                MajChampOblig ('l'''+Lib4)
             else
                MajChampOblig ('le '+Lib4);
             end;
          end;
       end;
   end;

if (TSalariesFille.getValue ('PSA_CONDEMPLOI')='P') and
   (TSalariesFille.getValue ('PSA_TAUXPARTIEL') = 0) then
   MajChampOblig ('le taux temps partiel');

if (NbChampOblig > 1) then
   begin
   FREPORT.Items.add (TraduireMemoire (LibErrSal+' : Les champs suivants sont'+
                                       ' obligatoires : '+ChampOblig));
   Erreur:= True;
   end
else
if (NbChampOblig = 1) then
   begin
   FREPORT.Items.add (TraduireMemoire (LibErrSal+' : Le champ suivant est'+
                                       ' obligatoire : '+ChampOblig));
   Erreur:= True;
   end;

// Erreurs liés au sexe et numéro de sécurité sociale
Sexe:= TSalariesFille.getValue ('PSA_SEXE');
NoSS:= TSalariesFille.getValue ('PSA_NUMEROSS');
if (Sexe <> '') then
   begin
   if (NoSS <> '') then
      begin
// Les infos ci dessous seront nécessaire dans TestNumeroSSNaissance
      if (TSalariesFille.getValue ('PSA_DEPTNAISSANCE')<>'') and
         (TSalariesFille.getValue ('PSA_DATENAISSANCE')<>iDate1900) then
         begin
//Les test sur l'année, le mois et le départements de naissance
//sont effectués que si le numéro SS fait 15 caractères, et avant le test de la clé
         if (Length(NoSS)=15) or (Length(NoSS)=13) then
            begin
            TestNumeroSSNaissance (NoSS, 'Annee',
                                   TSalariesFille.GetValue ('PSA_DATENAISSANCE'),
                                   TSalariesFille.getValue ('PSA_DEPTNAISSANCE'),
                                   ResultAnnee, ResultMois, ResultDepart);
            TestNumeroSSNaissance (NoSS, 'Mois',
                                   TSalariesFille.GetValue ('PSA_DATENAISSANCE'),
                                   TSalariesFille.getValue ('PSA_DEPTNAISSANCE'),
                                   ResultAnnee, ResultMois, ResultDepart);
            TestNumeroSSNaissance (NoSS, 'Depart',
                                   TSalariesFille.GetValue ('PSA_DATENAISSANCE'),
                                   TSalariesFille.getValue ('PSA_DEPTNAISSANCE'),
                                   ResultAnnee, ResultMois, ResultDepart);
            if (ResultAnnee=-8) then
               begin
               FREPORT.Items.add (TraduireMemoire (LibErrSal+' : L''année de'+
                                                   ' naissance dans le numéro'+
                                                   ' de sécurité sociale est'+
                                                   ' erronée.'));
               Erreur:= True;
               end;
            if (ResultMois=-9) then
               begin
               FREPORT.Items.add (TraduireMemoire (LibErrSal+' : Le mois de'+
                                                   ' naissance dans le numéro'+
                                                   ' de sécurité sociale est'+
                                                   ' erroné.'));
               Erreur:= True;
               end;
            if (ResultDepart=-10) then
               begin
               FREPORT.Items.add (TraduireMemoire (LibErrSal+' : Le département'+
                                                   ' de naissance dans le'+
                                                   ' numéro de sécurité'+
                                                   ' sociale est erroné.'));
               Erreur:= True;
               end;
            if (ResultDepart=-12) then
               begin
               FREPORT.Items.add (TraduireMemoire (LibErrSal+' : La date de naissance'+
                                                   ' du salarié étant inférieure'+
                                                   ' au 01/01/1976, le code du département'+
                                                   ' dans le numéro de sécurité sociale'+
                                                   ' ne peut être 2A ou 2B.'));
               Erreur:= True;
               end;
            end;
         end;

      Resultat:= TestNumeroSS (NoSS, Sexe);
      if (Length (NoSS) <> 0) then
         begin
         CleSexe:= NoSS[1];
         if (Resultat <> 0) then
            begin
            Erreur:= True;
            case Resultat of
                 1: FREPORT.Items.add (TraduireMemoire (LibErrSal+' : Vous'+
                                                        ' n''avez pas'+
                                                        ' renseigné le numéro'+
                                                        ' de sécurité sociale.'));
                 3: FREPORT.Items.add (TraduireMemoire (LibErrSal+' : Le numéro'+
                                                        ' de sécurité sociale'+
                                                        ' est provisoire. Vous'+
                                                        ' devrez le remplacer'+
                                                        ' par un numéro définitif.'));
                 -2: FREPORT.Items.add (TraduireMemoire (LibErrSal+' : Le numéro'+
                                                         ' de sécurité sociale'+
                                                         ' ne comporte pas de'+
                                                         ' clé.'));
                 -3: FREPORT.Items.add (TraduireMemoire (LibErrSal+' : Le numéro'+
                                                         ' de sécurité sociale'+
                                                         ' est incomplet, 15'+
                                                         ' positions'+
                                                         ' obligatoires.'));
                 -7: FREPORT.Items.add (TraduireMemoire (LibErrSal+' : Le numéro'+
                                                         ' de sécurité sociale'+
                                                         ' est incorrect. Ce'+
                                                         ' doit être une'+
                                                         ' valeur numérique.'));
                 -5: FREPORT.Items.add (TraduireMemoire (LibErrSal+' : La clé sexe'+
                                                         ' est erronée : "'+CleSexe+'".'+
                                                         ' La clé sexe du'+
                                                         ' numéro de sécurité'+
                                                         ' sociale vaut "2".'));
                 -6: FREPORT.Items.add (TraduireMemoire (LibErrSal+' : La clé sexe'+
                                                         ' est erronée : "'+CleSexe+'".'+
                                                         ' La clé sexe du'+
                                                         ' numéro de sécurité'+
                                                         ' sociale vaut "1".'));
                 -4: begin
                     if (Sexe = 'M') then
                        FREPORT.Items.add (TraduireMemoire (LibErrSal+' : La clé'+
                                                            ' sexe est erronée :'+
                                                            ' "'+CleSexe+'".'+
                                                            ' La clé sexe vaut'+
                                                            ' "1".'));
                     if (Sexe = 'F') then
                        FREPORT.Items.add (TraduireMemoire (LibErrSal+' : La clé'+
                                                            ' sexe est erronée :'+
                                                            ' "'+CleSexe+'".'+
                                                            ' La clé sexe vaut'+
                                                            ' "2".'));
                     end;
                 -1: FREPORT.Items.add (TraduireMemoire (LibErrSal+' : La clé'+
                                                         ' est erronée.'));
                 end;
            end;
         end;
      end;
   end;

// Controle contenu des champs Horaires
if TestHoraire (TSalariesFille.getValue ('PSA_HORHEBDO'),
                TSalariesFille.getValue ('PSA_HORAIREMOIS'))=True then
   begin
   FREPORT.Items.add (TraduireMemoire (LibErrSal+' : L''horaire hebdomadaire'+
                                       ' ne peut être supérieur à l''horaire'+
                                       ' mensuel.'));
   Erreur:= True;
   end;
if TestHoraire (TSalariesFille.getValue ('PSA_HORHEBDO'),
                TSalariesFille.getValue('PSA_HORANNUEL'))=True then
   begin
   FREPORT.Items.add (TraduireMemoire (LibErrSal+' : L''horaire hebdomadaire'+
                                       ' ne peut être supérieur à l''horaire'+
                                       ' annuel.'));
   Erreur:= True;
   end;
if TestHoraire (TSalariesFille.getValue ('PSA_HORAIREMOIS'),
                TSalariesFille.getValue('PSA_HORANNUEL'))=True then
   begin
   FREPORT.Items.add (TraduireMemoire (LibErrSal+' : L''horaire mensuel ne'+
                                       ' peut être supérieur à l''horaire'+
                                       ' annuel.'));
   Erreur:= True;
   end;

// Incohérence Civilité/Sexe/Nom de jeune fille
Civilite:= TSalariesFille.getValue ('PSA_CIVILITE');
NomJFille:= TSalariesFille.getValue ('PSA_NOMJF');
if (Sexe = 'M') and (NomJFille <> '') then
   begin
   FREPORT.Items.add (TraduireMemoire (LibErrSal+' : Incohérence entre le nom'+
                                       ' de jeune fille et le sexe.'));
   Erreur:= True;
   end;
if (((Civilite = 'MLE') or (Civilite = 'MME')) and (Sexe = 'M')) or
   ((Civilite = 'MR') and (Sexe = 'F')) then
   begin
   FREPORT.Items.add (TraduireMemoire (LibErrSal+' : Incohérence entre la'+
                                       ' civilité et le sexe.'));
   Erreur:= True;
   end;
if (Sexe = 'M') and (Civilite = 'MLE') and (NomJFille <> '') then
   begin
   FREPORT.Items.add (TraduireMemoire (LibErrSal+' : Incohérence entre le nom'+
                                       ' de jeune fille, la civilité et le'+
                                       ' sexe.'));
   Erreur:= True;
   end;
if (Sexe = 'M') and (Civilite <> 'MLE') and (NomJFille <> '') then
   begin
   FREPORT.Items.add (TraduireMemoire (LibErrSal+' : Incohérence entre le nom'+
                                       ' de jeune fille et le sexe.'));
   Erreur:= True;
   end;
if (Sexe <> 'M') and (Civilite = 'MLE') and (NomJFille <> '') then
   begin
   FREPORT.Items.add (TraduireMemoire (LibErrSal+' : Incohérence entre le nom'+
                                       ' de jeune fille et la civilité.'));
   Erreur:= True;
   end;
if ((Sexe = 'M') or  (Civilite = 'MLE')) and (NomJFille <> '') then
   begin
   FREPORT.Items.add (TraduireMemoire (LibErrSal+' : Incohérence entre le nom'+
                                       ' de jeune fille et la civilité et le'+
                                       ' sexe.'));
   Erreur:= True;
   end;

// PSA_DATENAISSANCE Date de naissance valide
if not IsValidDate (String (TSalariesFille.getValue ('PSA_DATENAISSANCE'))) then
   begin
   FREPORT.Items.add (TraduireMemoire (LibErrSal+' : La date de naissance'+
                                       ' n''est pas valide.'));
   Erreur:= True;
   end;

// La date de naissance ne doit pas être inférieure à la date d'entrée
DEntree:= TSalariesFille.getValue ('PSA_DATEENTREE');
DNaissance:= TSalariesFille.getValue ('PSA_DATENAISSANCE');
if (DEntree <> idate1900) and (DNaissance <> idate1900) and
   (DEntree < DNaissance) then
   begin
   FREPORT.Items.add (TraduireMemoire (LibErrSal+' : La date de naissance ne'+
                                       ' peut être supérieure à la date'+
                                       ' d''entrée.'));
   Erreur:= True;
   end;

// PSA_DATEENTREE Date d'entrée
if not IsValidDate (string (TSalariesFille.getValue ('PSA_DATEENTREE'))) then
   begin
   FREPORT.Items.add (TraduireMemoire (LibErrSal+' : La date d''entrée n''est'+
                                       ' pas valide.'));
   Erreur:= True;
   end;

// PSA_DATESORTIE Date de sortie
if not IsValidDate (string (TSalariesFille.getValue ('PSA_DATESORTIE'))) then
   begin
   FREPORT.Items.add (TraduireMemoire (LibErrSal+' : La date de sortie n''est'+
                                       ' pas valide.'));
   Erreur:= True;
   end;

// Pas de motif de sortie si pas de date de sortie
if (TSalariesFille.getValue ('PSA_DATESORTIE')<Idate1900) and
   (TSalariesFille.getValue ('PSA_MOTIFSORTIE')<>'') then
   begin
   FREPORT.Items.add (TraduireMemoire (LibErrSal+' : Aucun motif de sortie ne'+
                                       ' peut être saisi car la date de sortie'+
                                       ' n''est pas une date valide.'));
   Erreur:= True;
   end;

// PSA_DATEANCIENNETE Date de début d'ancienneté
if not IsValidDate (string (TSalariesFille.getValue ('PSA_DATEANCIENNETE'))) then
   begin
   FREPORT.Items.add (TraduireMemoire (LibErrSal+' : La date de début'+
                                       ' d''ancienneté n''est pas valide.'));
   Erreur:= True;
   end;

// PSA_ANCIENPOSTE Ancienneté dans le poste
if not IsValidDate (string (TSalariesFille.getValue ('PSA_ANCIENPOSTE'))) then
   begin
   FREPORT.Items.add (TraduireMemoire (LibErrSal+' : La date d''ancienneté'+
                                       ' dans le poste n''est pas valide.'));
   Erreur:= True;
   end;

// Vérification présence valeurs dans les paramètres tablettes
if VerifPresenceParam then
   begin
// PSA_CIVILITE Civilité
   VerifTablette ('PSA_CIVILITE', 'YYCIVILITE', LibSal, 'La Civilité',
                  TSalariesFille);

// PSA_SEXE Sexe
   VerifTablette ('PSA_SEXE', 'PGSEXE', LibSal, 'Le sexe', TSalariesFille);

// PSA_PAYS Pays
   VerifTablette ('PSA_PAYS', 'TTPAYS', LibSal, 'Le pays', TSalariesFille);

// PSA_NATIONALITE Nationalité
   VerifTablette ('PSA_NATIONALITE', 'YYNATIONALITE', LibSal, 'La nationalité',
                  TSalariesFille);

// PSA_PAYSNAISSANCE Pays de naissance
   VerifTablette ('PSA_PAYSNAISSANCE', 'TTPAYS', LibSal, 'Le pays de naissance',
                  TSalariesFille);

// PSA_SITUATIONFAMIL Situation de famille
   VerifTablette ('PSA_SITUATIONFAMIL', 'PGSITUATIONFAMIL', LibSal,
                  'La situation de famille', TSalariesFille);

// PSA_MOTIFENTREE Motif d'entrée
   VerifTablette ('PSA_MOTIFENTREE', 'PGMOTIFENTREELIGHT', LibSal,
                  'Le motif d''entrée', TSalariesFille);

// PSA_MOTIFSORTIE Motif de sortie
   VerifTablette ('PSA_MOTIFSORTIE', 'PGMOTIFSORTIE', LibSal,
                  'Le motif de sortie', TSalariesFille);

// PSA_CATDADS Catégorie DUCS
   VerifTablette ('PSA_CATDADS', 'PGCATBILAN', LibSal,
                  'La catégorie DUCS', TSalariesFille);

// PSA_CONVENTION Convention collective
   VerifTablette ('PSA_CONVENTION', 'PGCONVENTION', LibSal,
                  'La convention collective', TSalariesFille);

// PSA_CODEEMPLOI Nomenclature PCS
   VerifTablette ('PSA_CODEEMPLOI', 'PGCODEPCSESE', LibSal,
                  'La nomenclature PCS', TSalariesFille);

// PSA_LIBELLEEMPLOI Libellé emploi
   VerifTablette ('PSA_LIBELLEEMPLOI', 'PGLIBEMPLOI',LibSal,
                  'Le libellé emploi', TSalariesFille);

// PSA_QUALIFICATION Qualification
   VerifTablette ('PSA_QUALIFICATION', 'PGLIBQUALIFICATION', LibSal,
                  'La qualification', TSalariesFille);

// PSA_COEFFICIENT Coefficient
   VerifTablette ('PSA_COEFFICIENT', 'PGLIBCOEFFICIENT', LibSal,
                  'Le coefficient', TSalariesFille);

// PSA_INDICE Indice
   VerifTablette ('PSA_INDICE', 'PGLIBINDICE', LibSal,
                  'L''indice', TSalariesFille);

// PSA_NIVEAU Niveau
   VerifTablette ('PSA_NIVEAU', 'PGLIBNIVEAU', LibSal, 'Le niveau',
                  TSalariesFille);

// PSA_EDITORG Organisme à éditer
   VerifTablette ('PSA_EDITORG', 'PGORGANISME1', LibSal,
                  'L''organisme à éditer', TSalariesFille);

// PSA_ACTIVITE Activité
   VerifTablette ('PSA_ACTIVITE', 'PGACTIVITE', LibSal, 'L''activité',
                  TSalariesFille);

// PSA_PROFILREM Profil de rémunération
   VerifTablette ('PSA_PROFILREM', 'PGPROFILREM', LibSal,
                  'Le profil rémunération', TSalariesFille);

// PSA_PERIODBUL Périodicité plafond
   VerifTablette ('PSA_PERIODBUL', 'PGPROFILPERIODE', LibSal,
                  'La périodicité plafond', TSalariesFille);

// PSA_PROFILTPS Temps partiel
   VerifTablette ('PSA_PROFILTPS', 'PGPROFILTPS', LibSal,
                  'Le type de temps partiel', TSalariesFille);

// PSA_PROFILRBS Profil réduction loi Fillon
   VerifTablette ('PSA_PROFILRBS', 'PGPROFILRBS', LibSal,
                  'Le profil réduction loi Fillon', TSalariesFille);

// PSA_REDRTT2 Profil minoration loi Fillon
   VerifTablette ('PSA_REDRTT2', 'PGPROFILRTT2', LibSal,
                  'Le profil minoration loi Fillon', TSalariesFille);

// PSA_REDRTT1 Profil Réduction RTT loi Aubry 2
   VerifTablette ('PSA_REDRTT1', 'PGPROFILRTT1', LibSal,
                  'Le profil réduction RTT loi Aubry 2', TSalariesFille);

// PSA_REDREPAS Profil Réduction repas
   VerifTablette ('PSA_REDREPAS', 'PGPROFILREPAS', LibSal,
                  'Le profil réduction repas', TSalariesFille);

// PSA_PROFILAFP Profil abattement frais professionnels
   VerifTablette ('PSA_PROFILAFP', 'PGPROFILAFP', LibSal,
                  'Le profil abattement frais professionnels', TSalariesFille);

// PSA_PROFILAPP Profil gestion des appoints
   VerifTablette ('PSA_PROFILAPP', 'PGPROFILAPP', LibSal,
                  'Le profil gestion des appoints', TSalariesFille);

// PSA_PROFILRET Profil retraite
   VerifTablette ('PSA_PROFILRET', 'PGPROFILRET', LibSal,
                  'Le profil retraite', TSalariesFille);

// PSA_PROFILMUT Profil Cotisations mutuelle
   VerifTablette ('PSA_PROFILMUT', 'PGPROFILMUT', LibSal,
                  'Le profil cotisations mutuelle', TSalariesFille);

// PSA_PROFILPRE Profil Cotisations prévoyance
   VerifTablette ('PSA_PROFILPRE', 'PGPROFILPRE', LibSal,
                  'Le profil cotisations prévoyance', TSalariesFille);

// PSA_PROFILTSS Profil taxe sur les salaires
   VerifTablette ('PSA_PROFILTSS', 'PGPROFILTTS', LibSal,
                  'Le profil taxe sur les salaires', TSalariesFille);

// PSA_PROFILFNAL Profil FNAL + 9 salariés
   VerifTablette ('PSA_PROFILFNAL', 'PGPROFILFNAL', LibSal,
                  'Le profil FNAL + 9 salariés', TSalariesFille);

// PSA_PROFILTRANS Profil Transport + 9 salariés
   VerifTablette ('PSA_PROFILTRANS', 'PGPROFILTRANS', LibSal,
                  'Le profil Transport + 9 salariés', TSalariesFille);

// PSA_CALENDRIER calendrier
   VerifTablette ('PSA_CALENDRIER', 'AFTSTANDCALEN', LibSal, 'Le calendrier',
                  TSalariesFille);

// PSA_PROFILCDD Calcul prime de précarité
   VerifTablette ('PSA_PROFILCDD', 'PGPROFILCDD', LibSal,
                  'Le type de calcul prime de précarité', TSalariesFille);

// PSA_REGIMESS Régime sécurité sociale
   VerifTablette ('PSA_REGIMESS', 'PGREGIMESS', LibSal,
                  'Le régime de sécurité sociale', TSalariesFille);
//DEB PT3
// PSA_REGIMEMAL Régime obligatoire risque maladie
   VerifTablette ('PSA_REGIMEMAL', 'PGREGIMESS', LibSal,
                  'Le régime obligatoire risque maladie', TSalariesFille);

// PSA_REGIMEMAL Régime obligatoire risque AT
   VerifTablette ('PSA_REGIMEAT', 'PGREGIMESS', LibSal,
                  'Le régime obligatoire risque AT', TSalariesFille);

// PSA_REGIMEMAL Régime obligatoire vieillesse (PP)
   VerifTablette ('PSA_REGIMEVIP', 'PGREGIMESS', LibSal,
                  'Le régime obligatoire vieillesse (PP)', TSalariesFille);

// PSA_REGIMEMAL Régime obligatoire vieillesse (PS)
   VerifTablette ('PSA_REGIMEVIS', 'PGREGIMESS', LibSal,
                  'Le régime obligatoire vieillesse (PS)', TSalariesFille);
//FIN PT3

// PSA_DADSCAT Statut catégoriel
   VerifTablette ('PSA_DADSCAT', 'PGSCATEGORIEL', LibSal,
                  'Le statut catégoriel', TSalariesFille);

// PSA_DADSPROF Statut professionnel
   VerifTablette ('PSA_DADSPROF', 'PGSPROFESSIONNEL', LibSal,
                  'Le statut professionnel', TSalariesFille);

// PSA_DADSFRACTION Fraction DADS
   VerifTablette ('PSA_DADSFRACTION', 'PGCODESECTION', LibSal,
                  'La fraction DADS', TSalariesFille);

// PSA_TRAVETRANG Type travailleur étranger ou frontalier
   VerifTablette ('PSA_TRAVETRANG', 'PGTRAVAILETRANGER', LibSal,
                  'Le type travailleur étranger ou frontalier', TSalariesFille);

// PSA_CONDEMPLOI Condition d'emploi
   if (TSalariesFille.getValue ('PSA_CONDEMPLOI')<>'') then
      begin
{PT1-1 - Optimisation
      Q:= OpenSQL ('SELECT CO_LIBRE'+
                   ' FROM COMMUN WHERE'+
                   ' CO_TYPE="PCI" AND'+
                   ' CO_CODE="'+TSalariesFille.getValue ('PSA_CONDEMPLOI')+'"',TRUE);
      if Q.Eof then
         FREPORT.Items.add (TraduireMemoire (LibAvetSal+' : La condition'+
                                             ' d''emploi n''existe pas dans le'+
                                             ' paramétrage.'));
      Ferme(Q);
}
      St:= 'SELECT CO_LIBRE'+
           ' FROM COMMUN WHERE'+
           ' CO_TYPE="PCI" AND'+
           ' CO_CODE="'+TSalariesFille.getValue ('PSA_CONDEMPLOI')+'"';
      if (ExisteSQL (St)=False) then
{PT2
         FREPORT.Items.add (TraduireMemoire (LibAvetSal+' : La condition'+
                                             ' d''emploi n''existe pas dans le'+
                                             ' paramétrage.'));
}
         FREPORT.Items.add (TraduireMemoire (LibAvetSal+' : La caractéristique'+
                                             ' activité n''existe pas dans le'+
                                             ' paramétrage.'));
//FIN PT2
//FIN PT1-1
      end;

// PSA_PRUDHCOLL Collège prud'hommal
   VerifTablette ('PSA_PRUDHCOLL', 'PGCOLLEGEPRUD', LibSal,
                  'Le collège prud''hommal', TSalariesFille);

// PSA_PRUDHSECT Section prud'hommale
   VerifTablette ('PSA_PRUDHSECT', 'PGSECTIONPRUD', LibSal,
                  'La section prud''hommale', TSalariesFille);

// PSA_PRUDHVOTE Bureau de votre des prud'hommes
   VerifTablette ('PSA_PRUDHVOTE', 'PGLIEUVOTEPRUD', LibSal,
                  'Le bureau de vote des prud''hommes', TSalariesFille);
   end;

Result:= True;
end;

procedure TOF_PG_IMPORTASC.VerifTablette(Champ,NomTablette,LibCleErr,LibChampErr : String;T : TOB);
var
  Lib : String;
begin
  if (T.getValue(Champ) <> '') then
  begin
    if (Champ = 'PSA_CODEEMPLOI') then   // cas particulier où l'on récupère l'abrégé
      Lib := RechDom(NomTablette,T.getValue(Champ),True)
    else
      Lib := RechDom(NomTablette,T.getValue(Champ),False);
    if (Lib = '') or (Lib = 'Error') then
      FREPORT.Items.add(TraduireMemoire('AVERTISSEMENT : ' + LibCleErr + ' : ' + LibChampErr + ' n''existe pas dans le paramétrage.'));
  end;
end;

function TOF_PG_IMPORTASC.TestHoraire(HoraireInf, HoraireSup: Double): Boolean;
begin
  Result := False;
  if HoraireSup <> 0 then
  begin
    if HoraireInf > HoraireSup then Result := True;
  end;
end;


{***********A.G.L.***********************************************
Auteur  ...... : FC
Créé le ...... : 03/04/2007
Modifié le ... :   /  /
Description .. : Alimentation des valeurs par défaut non fournies du fichier
Suite ........ : SALARIES
Mots clefs ... : PAIE;PG_IMPORTASC
*****************************************************************}
procedure TOF_PG_IMPORTASC.EnregSalDefaut (TSalariesFille : TOB);
Var
InfTiers : Info_Tiers;
CodeAuxi, LeRapport, Libell, Preno : String;
LongLib, LongPre : integer;
begin
if not TSalariesFille.IsFieldModified('PSA_TYPEDITORG') then TSalariesFille.PutValue('PSA_TYPEDITORG', 'ETB');
if not TSalariesFille.IsFieldModified('PSA_TYPNBACQUISCP') then TSalariesFille.PutValue('PSA_TYPNBACQUISCP', 'ETB');
if not TSalariesFille.IsFieldModified('PSA_DATANC') then TSalariesFille.PutValue('PSA_DATANC', 'ETB');
if not TSalariesFille.IsFieldModified('PSA_CPACQUISMOIS') then TSalariesFille.PutValue('PSA_CPACQUISMOIS', 'ETB');
if not TSalariesFille.IsFieldModified('PSA_CPACQUISSUPP') then TSalariesFille.PutValue('PSA_CPACQUISSUPP', 'ETB');
if not TSalariesFille.IsFieldModified('PSA_TYPEDITBULCP') then TSalariesFille.PutValue('PSA_TYPEDITBULCP', 'ETB');
if not TSalariesFille.IsFieldModified('PSA_CPACQUISANC') then TSalariesFille.PutValue('PSA_CPACQUISANC', 'ETB');
if not TSalariesFille.IsFieldModified('PSA_CPTYPEMETHOD') then TSalariesFille.PutValue('PSA_CPTYPEMETHOD', 'ETB');
if not TSalariesFille.IsFieldModified('PSA_CPTYPERELIQ') then TSalariesFille.PutValue('PSA_CPTYPERELIQ', 'ETB');
if not TSalariesFille.IsFieldModified('PSA_CPTYPEVALO') then TSalariesFille.PutValue('PSA_CPTYPEVALO', 'ETB');
if not TSalariesFille.IsFieldModified('PSA_TYPREGLT') then TSalariesFille.PutValue('PSA_TYPREGLT', 'ETB');
if not TSalariesFille.IsFieldModified('PSA_TYPDATPAIEMENT') then TSalariesFille.PutValue('PSA_TYPDATPAIEMENT', 'ETB');
if not TSalariesFille.IsFieldModified('PSA_PRISEFFECTIF') then TSalariesFille.PutValue('PSA_PRISEFFECTIF', 'X');
if not TSalariesFille.IsFieldModified('PSA_UNITEPRISEFF') then TSalariesFille.PutValue('PSA_UNITEPRISEFF', 1);
if not TSalariesFille.IsFieldModified('PSA_TYPACTIVITE') then TSalariesFille.PutValue('PSA_TYPACTIVITE', 'ETB');
if not TSalariesFille.IsFieldModified('PSA_TYPPROFIL') then TSalariesFille.PutValue('PSA_TYPPROFIL', 'ETB');
if not TSalariesFille.IsFieldModified('PSA_TYPPERIODEBUL') then TSalariesFille.PutValue('PSA_TYPPERIODEBUL', 'ETB');
if not TSalariesFille.IsFieldModified('PSA_TYPPROFILRBS') then TSalariesFille.PutValue('PSA_TYPPROFILRBS', 'ETB');
if not TSalariesFille.IsFieldModified('PSA_TYPPROFILAFP') then TSalariesFille.PutValue('PSA_TYPPROFILAFP', 'ETB');
if not TSalariesFille.IsFieldModified('PSA_TYPPROFILAPP') then TSalariesFille.PutValue('PSA_TYPPROFILAPP', 'ETB');
if not TSalariesFille.IsFieldModified('PSA_TYPPROFILMUT') then TSalariesFille.PutValue('PSA_TYPPROFILMUT', 'ETB');
if not TSalariesFille.IsFieldModified('PSA_TYPPROFILPRE') then TSalariesFille.PutValue('PSA_TYPPROFILPRE', 'ETB');
if not TSalariesFille.IsFieldModified('PSA_TYPPROFILTSS') then TSalariesFille.PutValue('PSA_TYPPROFILTSS', 'ETB');
if not TSalariesFille.IsFieldModified('PSA_TYPPROFILCGE') then TSalariesFille.PutValue('PSA_TYPPROFILCGE', 'ETB');
if not TSalariesFille.IsFieldModified('PSA_TYPPROFILANC') then TSalariesFille.PutValue('PSA_TYPPROFILANC', 'ETB');
if not TSalariesFille.IsFieldModified('PSA_TYPPROFILFNAL') then TSalariesFille.PutValue('PSA_TYPPROFILFNAL', 'DOS');
if not TSalariesFille.IsFieldModified('PSA_TYPPROFILTRANS') then TSalariesFille.PutValue('PSA_TYPPROFILTRANS', 'ETB');
if not TSalariesFille.IsFieldModified('PSA_TYPPROFILRET') then TSalariesFille.PutValue('PSA_TYPPROFILRET', 'ETB');
if not TSalariesFille.IsFieldModified('PSA_TYPPAIACOMPT') then TSalariesFille.PutValue('PSA_TYPPAIACOMPT', 'ETB');
if not TSalariesFille.IsFieldModified('PSA_TYPPAIFRAIS') then TSalariesFille.PutValue('PSA_TYPPAIFRAIS', 'ETB');
if not TSalariesFille.IsFieldModified('PSA_DATELIBRE1') then TSalariesFille.PutValue('PSA_DATELIBRE1', IDate1900);
if not TSalariesFille.IsFieldModified('PSA_DATELIBRE2') then TSalariesFille.PutValue('PSA_DATELIBRE2', IDate1900);
if not TSalariesFille.IsFieldModified('PSA_DATELIBRE3') then TSalariesFille.PutValue('PSA_DATELIBRE3', IDate1900);
if not TSalariesFille.IsFieldModified('PSA_DATELIBRE4') then TSalariesFille.PutValue('PSA_DATELIBRE4', IDate1900);
if not TSalariesFille.IsFieldModified('PSA_DATELIBRE5') then TSalariesFille.PutValue('PSA_DATELIBRE5', IDate1900);
if not TSalariesFille.IsFieldModified('PSA_DATELIBRE6') then TSalariesFille.PutValue('PSA_DATELIBRE6', IDate1900);
if not TSalariesFille.IsFieldModified('PSA_DATELIBRE7') then TSalariesFille.PutValue('PSA_DATELIBRE7', IDate1900);
if not TSalariesFille.IsFieldModified('PSA_DATELIBRE8') then TSalariesFille.PutValue('PSA_DATELIBRE8', IDate1900);
if not TSalariesFille.IsFieldModified('PSA_ORDREAT') then TSalariesFille.PutValue('PSA_ORDREAT', '1');
if not TSalariesFille.IsFieldModified('PSA_CATBILAN') then TSalariesFille.PutValue('PSA_CATBILAN', '000');
if not TSalariesFille.IsFieldModified('PSA_STANDCALEND') then TSalariesFille.PutValue('PSA_STANDCALEND', 'ETB');
if not TSalariesFille.IsFieldModified('PSA_TYPJOURHEURE') then TSalariesFille.PutValue('PSA_TYPJOURHEURE', 'ETB');
if not TSalariesFille.IsFieldModified('PSA_JOURHEURE') then TSalariesFille.PutValue('PSA_JOURHEURE', 'ETB');
if not TSalariesFille.IsFieldModified('PSA_TYPREDREPAS') then TSalariesFille.PutValue('PSA_TYPREDREPAS', 'ETB');
if not TSalariesFille.IsFieldModified('PSA_TYPREDRTT1') then TSalariesFille.PutValue('PSA_TYPREDRTT1', 'ETB');
if not TSalariesFille.IsFieldModified('PSA_TYPREDRTT2') then TSalariesFille.PutValue('PSA_TYPREDRTT2', 'ETB');
if not TSalariesFille.IsFieldModified('PSA_TYPPROFILREM') then TSalariesFille.PutValue('PSA_TYPPROFILREM', 'ETB');
if not TSalariesFille.IsFieldModified('PSA_TYPVIRSOC') then TSalariesFille.PutValue('PSA_TYPVIRSOC', 'ETB');
if not TSalariesFille.IsFieldModified('PSA_TYPACPSOC') then TSalariesFille.PutValue('PSA_TYPACPSOC', 'ETB');
if not TSalariesFille.IsFieldModified('PSA_TYPFRAISSOC') then TSalariesFille.PutValue('PSA_TYPFRAISSOC', 'ETB');
if not TSalariesFille.IsFieldModified('PSA_CONFIDENTIEL') then TSalariesFille.PutValue('PSA_CONFIDENTIEL', '0');
if not TSalariesFille.IsFieldModified('PSA_DADSDATE') then TSalariesFille.PutValue('PSA_DADSDATE', IDate1900);
if not TSalariesFille.IsFieldModified('PSA_TYPDADSFRAC') then TSalariesFille.PutValue('PSA_TYPDADSFRAC', 'ETB');
if not TSalariesFille.IsFieldModified('PSA_TOPCONVERT') then TSalariesFille.PutValue('PSA_TOPCONVERT', 'X');
if not TSalariesFille.IsFieldModified('PSA_ETATBULLETIN') then TSalariesFille.PutValue('PSA_ETATBULLETIN','PBP');
if not TSalariesFille.IsFieldModified('PSA_NATIONALITE') then TSalariesFille.PutValue('PSA_NATIONALITE', 'FRA');
if not TSalariesFille.IsFieldModified('PSA_PAYSNAISSANCE') then TSalariesFille.PutValue('PSA_PAYSNAISSANCE', 'FRA');
if not TSalariesFille.IsFieldModified('PSA_TYPCONVENTION') then TSalariesFille.PutValue('PSA_TYPCONVENTION', 'PER');
//PT5
if not TSalariesFille.IsFieldModified('PSA_TYPEREGIME') then
   TSalariesFille.PutValue('PSA_TYPEREGIME', '-');
//FIN PT5
//DEB PT3
if TSalariesFille.GetValue('PSA_TYPEREGIME') = '-' then
   begin
   if not TSalariesFille.IsFieldModified('PSA_REGIMESS') then TSalariesFille.PutValue('PSA_REGIMESS', '200');
   if not TSalariesFille.IsFieldModified('PSA_REGIMEMAL') then TSalariesFille.PutValue('PSA_REGIMEMAL', '');
   if not TSalariesFille.IsFieldModified('PSA_REGIMEAT') then TSalariesFille.PutValue('PSA_REGIMEAT', '');
   if not TSalariesFille.IsFieldModified('PSA_REGIMEVIP') then TSalariesFille.PutValue('PSA_REGIMEVIP', '');
   if not TSalariesFille.IsFieldModified('PSA_REGIMEVIS') then TSalariesFille.PutValue('PSA_REGIMEVIS', '');
   end
else
   begin
   if not TSalariesFille.IsFieldModified('PSA_REGIMESS') then TSalariesFille.PutValue('PSA_REGIMESS', '');
   if not TSalariesFille.IsFieldModified('PSA_REGIMEMAL') then TSalariesFille.PutValue('PSA_REGIMEMAL', '200');
   if not TSalariesFille.IsFieldModified('PSA_REGIMEAT') then TSalariesFille.PutValue('PSA_REGIMEAT', '200');
   if not TSalariesFille.IsFieldModified('PSA_REGIMEVIP') then TSalariesFille.PutValue('PSA_REGIMEVIP', '200');
   if not TSalariesFille.IsFieldModified('PSA_REGIMEVIS') then TSalariesFille.PutValue('PSA_REGIMEVIS', '200');
   end;
//FIN PT3
if not TSalariesFille.IsFieldModified('PSA_TYPPRUDH') then TSalariesFille.PutValue('PSA_TYPPRUDH', 'ETB');

//PT4
// Chargement des infos nescessaires à la creation du compte de tiers associé
{if (VH_Paie.PgTypeNumSal = 'NUM') and (VH_Paie.PgTiersAuxiAuto = TRUE) then }//PT13
if (VH_Paie.PgTiersAuxiAuto = TRUE) then //PT13
   begin
   Libell:= TSalariesFille.GetValue ('PSA_LIBELLE');
   LongLib:= Length (Libell);
   Preno:= TSalariesFille.GetValue ('PSA_PRENOM');
   LongPre:= Length (Preno);
   Libell:= Copy (Libell, 1, LongLib)+ ' ' + Copy (Preno, 1, LongPre);
   with InfTiers do
        begin
        Libelle:= Copy (Libell, 1, 35);
        Adresse1:= TSalariesFille.GetValue ('PSA_ADRESSE1');
        Adresse2:= TSalariesFille.GetValue ('PSA_ADRESSE2');
        Adresse3:= TSalariesFille.GetValue ('PSA_ADRESSE3');
        Ville:= TSalariesFille.GetValue ('PSA_VILLE');
        Telephone:= TSalariesFille.GetValue ('PSA_TELEPHONE');
        CodePostal:= TSalariesFille.GetValue ('PSA_CODEPOSTAL');
        Pays:= '';
        end;
   LeRapport:= '';

//Creation du compte de tiers en automatique et recup du numéro
   CodeAuxi:= CreationTiers (InfTiers,LeRapport, 'SAL',
                             TSalariesFille.GetValue ('PSA_SALARIE'));
   if (CodeAuxi<>'') then
      TSalariesFille.PutValue ('PSA_AUXILIAIRE',CodeAuxi);
   end;

ChangeCodeSalarie (TSalariesFille, 'PSA', TRUE) ;
//FIN PT4
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : FC
Créé le ...... : 06/04/2007
Modifié le ... :
Description .. : Controle des valeurs du fichier HISTOCUMSAL
Mots clefs ... : PAIE;PG_IMPORTASC
*****************************************************************}
function TOF_PG_IMPORTASC.ControleValHistoCumul(THistoCumulFille: TOB): boolean;
var
St, StHisto : string;
begin
// A priori renvoie toujours Vrai car on ne fait que signaler les anomalies, rien n'est bloquant
Result:= False;

StHisto:= THistoCumulFille.getValue ('PHC_ETABLISSEMENT')+'/'+
          THistoCumulFille.getValue ('PHC_SALARIE')+'/"'+
          DateToStr (THistoCumulFille.getValue ('PHC_DATEDEBUT'))+'"/"'+
          DateToStr (THistoCumulFille.getValue ('PHC_DATEFIN'))+'"/'+
          THistoCumulFille.getValue ('PHC_REPRISE')+'/'+
          THistoCumulFille.getValue ('PHC_CUMULPAIE');

// Si l'enregistrement existe déjà, erreur
{PT1-1 - Optimisation
Q:= OpenSQL ('SELECT PHC_ETABLISSEMENT'+
             ' FROM HISTOCUMSAL WHERE'+
             ' PHC_ETABLISSEMENT="'+THistoCumulFille.getValue ('PHC_ETABLISSEMENT')+'" AND'+
             ' PHC_SALARIE="'+THistoCumulFille.getValue ('PHC_SALARIE')+'" AND'+
             ' PHC_DATEDEBUT="'+UsDateTime (THistoCumulFille.getValue ('PHC_DATEDEBUT'))+'" AND'+
             ' PHC_DATEFIN="'+UsDateTime (THistoCumulFille.getValue ('PHC_DATEFIN'))+'" AND'+
             ' PHC_REPRISE="'+THistoCumulFille.getValue ('PHC_REPRISE')+'" AND'+
             ' PHC_CUMULPAIE="'+THistoCumulFille.getValue ('PHC_CUMULPAIE')+'"',
             True);
if not Q.Eof then
   begin
   FREPORT.Items.add (TraduireMemoire ('ERREUR : L''historique cumul '+StHisto+
                                       ' existe déjà'));
   Erreur:= True;
   end;
Ferme (Q);
}
St:= 'SELECT PHC_ETABLISSEMENT'+
     ' FROM HISTOCUMSAL WHERE'+
     ' PHC_ETABLISSEMENT="'+THistoCumulFille.getValue ('PHC_ETABLISSEMENT')+'" AND'+
     ' PHC_SALARIE="'+THistoCumulFille.getValue ('PHC_SALARIE')+'" AND'+
     ' PHC_DATEDEBUT="'+UsDateTime (THistoCumulFille.getValue ('PHC_DATEDEBUT'))+'" AND'+
     ' PHC_DATEFIN="'+UsDateTime (THistoCumulFille.getValue ('PHC_DATEFIN'))+'" AND'+
     ' PHC_REPRISE="'+THistoCumulFille.getValue ('PHC_REPRISE')+'" AND'+
     ' PHC_CUMULPAIE="'+THistoCumulFille.getValue ('PHC_CUMULPAIE')+'"';
if (ExisteSQL (St)) then
   begin
   FREPORT.Items.add (TraduireMemoire ('ERREUR : L''historique cumul '+StHisto+
                                       ' existe déjà'));
   Erreur:= True;
   end;
//FIN PT1-1

ChampOblig:= '';
NbChampOblig:= 0;
// Champs obligatoires : PHC_ETABLISSEMENT,PHC_SALARIE,PHC_DATEDEBUT,PHC_DATEFIN,PHC_REPRISE,PHC_CUMULPAIE
if (THistoCumulFille.getValue('PHC_ETABLISSEMENT')='') then
   MajChampOblig ('l''établissement');
if (THistoCumulFille.getValue ('PHC_SALARIE')='') then
   MajChampOblig (' le salarié');
if (THistoCumulFille.getValue ('PHC_DATEDEBUT')=iDate1900) then
   MajChampOblig (' la date de début');
if (THistoCumulFille.getValue ('PHC_DATEFIN')=iDate1900) then
   MajChampOblig (' la date de fin');
if (THistoCumulFille.getValue ('PHC_CUMULPAIE')='') then
   MajChampOblig (' le cumul');

if (NbChampOblig>1) then
   begin
   FREPORT.Items.add (TraduireMemoire ('ERREUR : Historique cumul '+StHisto+
                      ' : Les champs suivants sont obligatoires : '+ChampOblig));
   Erreur:= True;
   end
else
if (NbChampOblig=1) then
   begin
   FREPORT.Items.add (TraduireMemoire ('ERREUR : Historique cumul '+StHisto+
                      ' : Le champ suivant est obligatoire : '+ChampOblig));
   Erreur:= True;
   end;

// Vérifier l'existence de l'établissement
if (THistoCumulFille.getValue ('PHC_ETABLISSEMENT')<>'') then
   begin
{PT1-1 - Optimisation
   Q:= OpenSQL ('SELECT ET_ETABLISSEMENT'+
                ' FROM ETABLISS WHERE'+
                ' ET_ETABLISSEMENT="'+THistoCumulFille.getValue ('PHC_ETABLISSEMENT')+'"', True);
   if Q.Eof then
      begin
      FREPORT.Items.add (TraduireMemoire ('ERREUR : Historique cumul '+StHisto+
                                          ' : L''établissement n''existe pas.'));
      Erreur:= True;
      end;
   Ferme(Q);
}
   St:= 'SELECT ET_ETABLISSEMENT'+
        ' FROM ETABLISS WHERE'+
        ' ET_ETABLISSEMENT="'+THistoCumulFille.getValue ('PHC_ETABLISSEMENT')+'"';
   if (ExisteSQL (St)=False) then
      begin
      FREPORT.Items.add (TraduireMemoire ('ERREUR : Historique cumul '+StHisto+
                                          ' : L''établissement n''existe pas.'));
      Erreur:= True;
      end;
//FIN PT1-1
   end;

// Vérifier l'existence du salarié
if (THistoCumulFille.getValue ('PHC_SALARIE')<>'') then
   begin
{PT1-1 - Optimisation
   Q:= OpenSQL ('SELECT PSA_SALARIE'+
                ' FROM SALARIES WHERE'+
                ' PSA_SALARIE="'+THistoCumulFille.getValue ('PHC_SALARIE')+'"', True);
   if Q.Eof then
      begin
      FREPORT.Items.add (TraduireMemoire ('ERREUR : Historique cumul '+StHisto+
                                          ' : Le salarié n''existe pas.'));
      Erreur:= True;
      end;
   Ferme(Q);
}
   St:= 'SELECT PSA_SALARIE'+
        ' FROM SALARIES WHERE'+
        ' PSA_SALARIE="'+THistoCumulFille.getValue ('PHC_SALARIE')+'"';
   if (ExisteSQL (St)=False) then
      begin
      FREPORT.Items.add (TraduireMemoire ('ERREUR : Historique cumul '+StHisto+
                                          ' : Le salarié n''existe pas.'));
      Erreur:= True;
      end;
//FIN PT1-1
   end;

// Vérifier l'existence du cumul de paie
if (THistoCumulFille.getValue ('PHC_CUMULPAIE')<>'') then
   begin
{PT1-1 - Optimisation
   Q:= OpenSQL ('SELECT PCL_CUMULPAIE'+
                ' FROM CUMULPAIE WHERE'+
                ' PCL_CUMULPAIE="'+THistoCumulFille.getValue ('PHC_CUMULPAIE')+'"', True);
   if Q.Eof then
      begin
      FREPORT.Items.add (TraduireMemoire ('ERREUR : Historique cumul '+StHisto+
                                          ' : Le cumul de paie n''existe pas.'));
      Erreur:= True;
      end;
   Ferme(Q);
}
   St:= 'SELECT PCL_CUMULPAIE'+
        ' FROM CUMULPAIE WHERE'+
        ' PCL_CUMULPAIE="'+THistoCumulFille.getValue ('PHC_CUMULPAIE')+'"';
   if (ExisteSQL (St)=False) then
      begin
      FREPORT.Items.add (TraduireMemoire ('ERREUR : Historique cumul '+StHisto+
                                          ' : Le cumul de paie n''existe pas.'));
      Erreur:= True;
      end;
//FIN PT1-1
   end;

Result:= True;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : FC
Créé le ...... : 06/04/2007
Modifié le ... :
Description .. : Controle des valeurs du fichier ABSENCESALARIE
Mots clefs ... : PAIE;PG_IMPORTASC
*****************************************************************}
function TOF_PG_IMPORTASC.ControleValAbsenceSalarie(TAbsenceFille: TOB): boolean;
var
Q : TQuery;
Etablissement, SQL, StTypeConge, StTypeMvt : String;
CPEtab, CPSal : Boolean;
DateSortie,DateEntree : TDateTime;
T_MotifAbs : TOB;
YYD, MMD, JJ, YYF, MMF: WORD;
nbj : double;
begin
// A priori renvoie toujours Vrai car on ne fait que signaler les anomalies, rien n'est bloquant
Result:= False;

nbj:= 0; //PT1-1

StTypeConge:= TAbsenceFille.getValue ('PCN_SALARIE')+'/'+
              TAbsenceFille.getValue ('PCN_TYPECONGE')+'/'+
              IntToStr (TAbsenceFille.getValue ('PCN_ORDRE'));
StTypeMvt:= TAbsenceFille.getValue ('PCN_SALARIE')+'/'+
            TAbsenceFille.getValue ('PCN_TYPEMVT')+'/'+
            IntToStr (TAbsenceFille.getValue ('PCN_ORDRE'));

// Si l'enregistrement existe déjà, erreur
SQL:= 'SELECT PCN_TYPEMVT'+
      ' FROM ABSENCESALARIE WHERE'+
      ' PCN_TYPEMVT="'+TAbsenceFille.getValue ('PCN_TYPEMVT')+'" AND'+
      ' PCN_SALARIE="'+TAbsenceFille.getValue ('PCN_SALARIE')+'" AND'+
      ' PCN_ORDRE='+IntToStr (TAbsenceFille.getValue ('PCN_ORDRE'))+' AND'+
      ' PCN_RESSOURCE="'+TAbsenceFille.getValue ('PCN_RESSOURCE')+'"';
{PT1-1 - Optimisation
Q:= OpenSQL (SQL, True);
if not Q.Eof then
   begin
   FREPORT.Items.add (TraduireMemoire ('ERREUR : L''absence salarié '+
                      StTypeConge+' existe déjà'));
   Erreur:= True;
   end;
Ferme(Q);
}
if (ExisteSQL (SQL)) then
   begin
   FREPORT.Items.add (TraduireMemoire ('ERREUR : L''absence salarié '+
                      StTypeConge+' existe déjà'));
   Erreur:= True;
   end;
//FIN PT1-1

if (TAbsenceFille.getValue ('PCN_TYPEMVT')='CPA') and
   (TAbsenceFille.getValue ('PCN_TYPECONGE')<>'PRI') then
   T_MotifAbs:= nil
else
   begin
   T_MotifAbs:= nil;
   if (TAbsenceFille.getValue ('PCN_TYPECONGE')<>'') then
      T_MotifAbs:= Tob_MotifAbs.FindFirst (['PMA_MOTIFABSENCE'],
                                           [TAbsenceFille.getValue ('PCN_TYPECONGE')],
                                           False);
   end;

if Assigned(T_MotifAbs) then
   begin
   if ControleGestionMaximum (TAbsenceFille.getValue ('PCN_SALARIE'),
                              TAbsenceFille.getValue ('PCN_TYPECONGE'), T_MotifAbs,
                              TAbsenceFille.getValue ('PCN_DATEDEBUTABS'),
                              Valeur (TAbsenceFille.getValue ('PCN_JOURS')),
                              Valeur (TAbsenceFille.getValue ('PCN_HEURES'))) then
      begin
      if (T_MotifAbs.GetValue ('PMA_JOURHEURE')='JOU') then
         begin
         FREPORT.Items.add (TraduireMemoire ('ERREUR : Absence salarié '+
                                             StTypeMvt+' : Le nombre de jours'+
                                             ' maximum octroyés pour ce motif'+
                                             ' est dépassé : '+
                                             FloatToStr (T_MotifAbs.GetValue ('PMA_JRSMAXI'))));
         Erreur:= True;
         end
      else
      if ((T_MotifAbs.GetValue ('PMA_JOURHEURE')='HEU') or (T_MotifAbs.GetValue ('PMA_JOURHEURE')='HOR')) then  //PT6
         begin
         FREPORT.Items.add (TraduireMemoire ('ERREUR : Absence salarié '+
                                             StTypeMvt+' : Le nombre d''heures'+
                                             ' maximum octroyés pour ce motif'+
                                             ' est dépassé : '+
                                             FloatToStr (T_MotifAbs.GetValue ('PMA_JRSMAXI'))));
         Erreur:= True;
         end;
      end;
   end;

// date validité postérieure à celle du dernier solde tout compte
if (TAbsenceFille.getValue ('PCN_DATEVALIDITE')<>iDate1900) and
   (TAbsenceFille.getValue ('PCN_TYPEMVT')='CPA') and
   ((TAbsenceFille.getValue ('PCN_TYPECONGE')='AJP') or
   (TAbsenceFille.getValue ('PCN_TYPECONGE')='AJU') or
   (TAbsenceFille.getValue ('PCN_TYPECONGE')='REP') or
   (TAbsenceFille.getValue ('PCN_TYPECONGE')='CPA')) then
   begin
   SQL:= 'SELECT MAX(PCN_DATEVALIDITE) AS DATEVAL'+
         ' FROM ABSENCESALARIE WHERE'+
         ' PCN_SALARIE="'+TAbsenceFille.getValue ('PCN_SALARIE')+'" AND'+
         ' PCN_TYPEMVT="CPA" AND'+
         ' PCN_TYPECONGE="SLD" AND'+
         ' PCN_GENERECLOTURE="-" ';
   Q:= opensql (SQL, True);
   if not Q.eof then
      begin
      if TAbsenceFille.getValue ('PCN_DATEVALIDITE')<=Q.FindField ('DATEVAL').AsDateTime then
         begin
         FREPORT.Items.add (TraduireMemoire ('ERREUR : Absence salarié '+
                                             StTypeMvt+' : La date de validité'+
                                             ' ne peut être antérieure à celle'+
                                             ' du dernier solde de tout compte'+
                                             ' : '+DateToStr (Q.FindField ('DATEVAL').AsDateTime)));
         Erreur:= True;
         end;
      end;
   Ferme(Q);
   end;

// Absence à cheval sur plusieurs mois
if (TAbsenceFille.getValue ('PCN_DATEDEBUTABS')<>iDate1900) and
   (TAbsenceFille.getValue ('PCN_DATEFINABS')<>iDate1900) and
   (not GetParamSocSecur ('SO_PGABSENCECHEVAL', False)) then
   begin
   if (TAbsenceFille.getValue ('PCN_TYPEMVT')='ABS') or
      (TAbsenceFille.getValue ('PCN_TYPECONGE')='PRI') Then
      begin
      DecodeDate (TAbsenceFille.getValue ('PCN_DATEDEBUTABS'), YYD, MMD, JJ);
      DecodeDate (TAbsenceFille.getValue ('PCN_DATEFINABS'), YYF, MMF, JJ);
      if (MMD <> MMF) or (YYD <> YYF) then
         begin
         FREPORT.Items.add (TraduireMemoire ('ERREUR : Absence salarié '+
                                             StTypeMvt+' : Vous ne pouvez pas'+
                                             ' saisir une absence à cheval sur'+
                                             ' plusieurs mois.'));
         Erreur:= True;
         end;
      end;
   end;

// Nombre de jours ou de mois ou la base obligatoires
if (TAbsenceFille.getValue ('PCN_TYPECONGE')='AJU') AND
   (TAbsenceFille.getValue('PCN_TYPEMVT')='CPA') then
   begin
   if (Valeur (TAbsenceFille.getValue ('PCN_JOURS'))=0) and
      (TAbsenceFille.getValue ('PCN_BASE')=0) and
      (TAbsenceFille.getValue ('PCN_NBREMOIS')=0) then
      begin
      FREPORT.Items.add (TraduireMemoire ('ERREUR : Absence salarié '+
                                          StTypeConge+' : Vous devez'+
                                          ' renseigner une valeur pour le'+
                                          ' nombre de jours, de mois ou la'+
                                          ' base.'));
      Erreur:= True;
      end;
   end;

// Vérifier l'existence du salarié
if (TAbsenceFille.getValue ('PCN_SALARIE')<>'') then
   begin
{PT1-1 - Optimisation
   Q:= OpenSQL ('SELECT PSA_SALARIE'+
                ' FROM SALARIES WHERE'+
                ' PSA_SALARIE="'+TAbsenceFille.getValue ('PCN_SALARIE')+'"', True);
   if Q.Eof then
      begin
      FREPORT.Items.add (TraduireMemoire ('ERREUR : Absence salarié '+
                                          StTypeConge+' : Le salarié n''existe'+
                                          ' pas.'));
      Erreur:= True;
      end;
   Ferme(Q);
}
   SQL:= 'SELECT PSA_SALARIE'+
         ' FROM SALARIES WHERE'+
         ' PSA_SALARIE="'+TAbsenceFille.getValue ('PCN_SALARIE')+'"';
   if (ExisteSQL (SQL)=False) then
      begin
      FREPORT.Items.add (TraduireMemoire ('ERREUR : Absence salarié '+
                                          StTypeConge+' : Le salarié n''existe'+
                                          ' pas.'));
      Erreur:= True;
      end;
//FIN PT1-1
   end;

// Paramètre CONGESPAYES
Q:= opensql ('SELECT PSA_CONGESPAYES, ETB_CONGESPAYES, ETB_ETABLISSEMENT,'+
             ' PSA_DATESORTIE, PSA_DATEENTREE'+
             ' FROM SALARIES'+
             ' LEFT JOIN ETABCOMPL ON'+
             ' PSA_ETABLISSEMENT=ETB_ETABLISSEMENT WHERE'+
             ' PSA_SALARIE="'+TAbsenceFille.getValue ('PCN_SALARIE')+'" ', TRUE);
if not Q.eof then
   begin
   CPEtab:= (Q.findfield ('ETB_CONGESPAYES').Asstring = 'X');
   CPSal:= (Q.findfield ('PSA_CONGESPAYES').Asstring = 'X');
   Etablissement:= Q.findfield ('ETB_ETABLISSEMENT').Asstring;
   DateSortie:= Q.findfield ('PSA_DATESORTIE').AsDateTime;
   DateEntree:= Q.findfield ('PSA_DATEENTREE').AsDateTime;
   end
else
   begin
   CPEtab:= False;
   CPSal:= False;
   Etablissement:= '';
   DateSortie:= iDate1900;
   DateEntree:= iDate1900;
   end;
Ferme(Q);

if TAbsenceFille.getValue ('PCN_TYPEMVT')='CPA' then
   begin
   if (not GetParamSocSecur ('SO_PGCONGES', False)) then
      begin
      FREPORT.Items.add (TraduireMemoire ('ERREUR : Absence salarié '+
                                          StTypeConge+' : Saisie impossible :'+
                                          ' Vous ne gérez pas les congés payés'+
                                          ' au niveau dossier.'));
      Erreur:= True;
      end
   else
   if (CPEtab=False) and (Etablissement<>'') then
      begin
      FREPORT.Items.add (TraduireMemoire ('ERREUR : Absence salarié '+
                                          StTypeConge+' : Saisie impossible :'+
                                          ' Vous ne gérez pas les congés payés'+
                                          ' au niveau établissement.'));
      Erreur:= True;
      end
   else
   if (CPSal=False) then
      begin
      FREPORT.Items.add (TraduireMemoire ('ERREUR : Absence salarié '+
                                          StTypeConge+' : Saisie impossible :'+
                                          ' Vous ne gérez pas les congés payés'+
                                          ' au niveau salarié.'));
      Erreur:= True;
      end;
   end;

ChampOblig:= '';
NbChampOblig:= 0;
// Champs obligatoires : PCN_TYPEMVT,PCN_SALARIE,PCN_ORDRE, PCN_TYPECONGE
if (TAbsenceFille.getValue ('PCN_TYPEMVT')='') then
   MajChampOblig ('le type de mouvement');
if (TAbsenceFille.getValue ('PCN_SALARIE')='') then
   MajChampOblig ('le matricule salarié');
if (TAbsenceFille.getValue ('PCN_ORDRE')=0) then
   MajChampOblig ('le numéro d''ordre');
if (TAbsenceFille.getValue ('PCN_TYPECONGE')='') then
   MajChampOblig ('le type d''absence');
if (TAbsenceFille.getValue ('PCN_DATEDEBUTABS')=iDate1900) then
   MajChampOblig ('la date de début d''absence');
if (TAbsenceFille.getValue ('PCN_DATEFINABS')=iDate1900) then
   MajChampOblig ('la date de fin d''absence');
if (TAbsenceFille.getValue ('PCN_DATEVALIDITE')=iDate1900) then
   MajChampOblig ('la date de validité');

if (NbChampOblig>1) then
   begin
   FREPORT.Items.add (TraduireMemoire ('ERREUR : Absence salarié '+StTypeConge+
                                       ' : Les champs suivants sont'+
                                       ' obligatoires : '+ChampOblig));
   Erreur:= True;
   end
else
if (NbChampOblig = 1) then
   begin
   FREPORT.Items.add (TraduireMemoire ('ERREUR : Absence salarié '+StTypeConge+
                                       ' : Le champ suivant est obligatoire : '+
                                       ChampOblig));
   Erreur:= True;
   end;

// Jours d'absence
if isnumeric (TAbsenceFille.getValue ('PCN_JOURS')) then
   begin
   if (Valeur (TAbsenceFille.getValue ('PCN_JOURS'))<=0) and
      (TAbsenceFille.getValue ('PCN_TYPECONGE')<>'AJU') then
      begin
      FREPORT.Items.add (TraduireMemoire ('ERREUR : Absence salarié '+
                                          StTypeConge+' : Le nombre de jours'+
                                          ' doit être positif'));
      Erreur:= True;
      end;
   end
else
   begin
   FREPORT.Items.add (TraduireMemoire ('ERREUR : Absence salarié '+StTypeConge+
                                       ' : Le nombre de jours est invalide.'));
   Erreur:= True;
   end;

// Controle des dates d'absence par rapport aux dates d'entrée / Sortie du salarié
if ((TAbsenceFille.getValue ('PCN_TYPECONGE')='PRI') and
   (TAbsenceFille.getValue ('PCN_TYPEMVT')='CPA')) or
   (TAbsenceFille.getValue ('PCN_TYPEMVT')='ABS') then
   begin
   if ((TAbsenceFille.getValue ('PCN_DATEFINABS')<>iDate1900) and
      (DateSortie>iDate1900) and
      (DateSortie<TAbsenceFille.getValue ('PCN_DATEFINABS'))) then
      begin
      FREPORT.Items.add (TraduireMemoire ('ERREUR : Absence salarié '+
                                          StTypeConge+' : La date de fin'+
                                          ' d''absence doit être antérieure à'+
                                          ' la date de sortie du salarié '+
                                          DateToStr (DateSortie)));
      Erreur:= True;
      end;

   if ((TAbsenceFille.getValue ('PCN_DATEDEBUTABS')<>iDate1900) and
      (DateEntree<>iDate1900) and
      (DateEntree>TAbsenceFille.getValue ('PCN_DATEDEBUTABS'))) then
      begin
      FREPORT.Items.add (TraduireMemoire ('ERREUR : Absence salarié '+
                                          StTypeConge+' : La date de début'+
                                          ' d''absence doit être postérieure à'+
                                          ' la date d''entrée du salarié '+
                                          DateToStr (DateEntree)));
      Erreur:= True;
      end;
   end;

// Récupérer la Date de Clôture CP année en cours
Q:= opensql ('SELECT ETB_DATECLOTURECPN'+
             ' FROM ETABCOMPL WHERE'+
             ' ETB_ETABLISSEMENT="'+TAbsenceFille.getValue ('PCN_ETABLISSEMENT')+'" ',
             TRUE);
if not Q.eof then
   DTClot:= Q.findfield ('ETB_DATECLOTURECPN').AsDateTime
else
   DTClot:= iDate1900;
Ferme(Q);

if (TAbsenceFille.getValue ('PCN_TYPECONGE')='CPA') AND
   (TAbsenceFille.getValue ('PCN_TYPEMVT')='CPA') then
   begin
// Vous ne pouvez saisir deux mouvements de type reprise pris sur une même période
{PT1-1
   Q:= RechercheReprise (TAbsenceFille.getValue ('PCN_DATEVALIDITE'),
                         TAbsenceFille.getValue ('PCN_SALARIE'), 'CPA', 'AC2');
}
   Q:= RechercheReprise (TAbsenceFille.getValue ('PCN_DATEVALIDITE'),
                         TAbsenceFille.getValue ('PCN_SALARIE'), 'CPA', 'AC2',
                         'REP');
   if not Q.eof then
      begin
      Q.First;
      while not Q.eof do
            begin
//PT1-1
            if (Q.FindField ('PCN_TYPECONGE').AsString='CPA') then
               begin
//FIN PT1-1
               if Q.FindField ('PCN_ORDRE').AsInteger<>TAbsenceFille.getValue ('PCN_ORDRE') then
                  begin
                  FREPORT.Items.add (TraduireMemoire ('ERREUR : Absence salarié '+
                                                      StTypeConge+' : Vous ne pouvez'+
                                                      ' pas saisir deux mouvements'+
                                                      ' de type reprise pris sur'+
                                                      ' une même période.'));
                  Erreur:= True;
                  end;
               end;   //PT1-1
            Q.next;
            end;
      end;
{PT1-1
   Ferme(Q);
}

// Le nombre de jours pris repris ne peut être supérieur au nombre de jours acquis repris
{PT1-1
   Q:= RechercheReprise (TAbsenceFille.getValue ('PCN_DATEVALIDITE'),
                         TAbsenceFille.getValue ('PCN_SALARIE'), 'REP', 'AC2');
}
   if not Q.eof then
//PT1-1
      begin
      Q.First;
      while not Q.eof do
            begin
            if (Q.FindField ('PCN_TYPECONGE').AsString='REP') then
//FIN PT1-1
               nbj:= Q.findfield ('PCN_JOURS').AsFloat;
//PT1-1
            Q.next;
            end;
      end;
{
   else
      nbj:= 0;
FIN PT1-1
}
   if nbj < Valeur (TAbsenceFille.getValue ('PCN_JOURS')) then
      begin
      FREPORT.Items.add (TraduireMemoire ('ERREUR : Absence salarié '+StTypeConge+
                                          ' : Le nombre de jours pris repris'+
                                          ' ne peut pas être supérieur au'+
                                          ' nombre de jours acquis repris : '+
                                          FloatToStr (nbj)));
      Erreur:= True;
      end;
   Ferme(Q);
   end;

// Vous ne pouvez saisir deux mouvements de type reprise acquis sur une même période
if (TAbsenceFille.getValue ('PCN_TYPECONGE')='REP') AND
   (TAbsenceFille.getValue ('PCN_TYPEMVT')='CPA') then
   begin
   Q:= RechercheReprise (TAbsenceFille.getValue ('PCN_DATEVALIDITE'),
                         TAbsenceFille.getValue ('PCN_SALARIE'), 'REP',
                         TAbsenceFille.getValue ('PCN_TYPEIMPUTE'));
   if not Q.eof then
      begin
      if TAbsenceFille.getValue ('PCN_ORDRE')<>Q.FindField ('PCN_ORDRE').AsInteger then
         begin
         FREPORT.Items.add (TraduireMemoire ('ERREUR : Absence salarié '+
                                             StTypeConge+' : Vous ne pouvez'+
                                             ' pas saisir deux mouvements de'+
                                             ' type reprise acquis sur une'+
                                             ' même période.'));
         Erreur:= True;
         end;
      end;
   Ferme(Q);
   end;

Result:= True;
end;

function TOF_PG_IMPORTASC.RechercheReprise (Validite : Tdatetime; Salarie,
                                            TypeConge, TypeImpute : string;
                                            TypeConge2 : string='') : Tquery;
var
Q : TQuery;
DtFin, DTDeb : TDateTime;
st, WhereType : string;
begin
//PT1-1
if (TypeConge2='') then
   WhereType:= 'PCN_TYPECONGE="'+TypeConge+'"'
else
   WhereType:= '(PCN_TYPECONGE="'+TypeConge+'" OR'+
               ' PCN_TYPECONGE="'+TypeConge2+'")';
//FIN PT1-1
RechercheExerciceCp (Validite, DtDeb, DtFin);
st:= 'SELECT PCN_ORDRE,PCN_JOURS,PCN_APAYES'+
     ' FROM ABSENCESALARIE WHERE'+
     ' PCN_SALARIE="'+Salarie+'" AND'+
     ' PCN_TYPEMVT="CPA" AND '+WhereType+' AND'+
     ' PCN_DATEVALIDITE>="'+USDateTime(DtDeb)+'" AND'+
     ' PCN_DATEVALIDITE<="'+USDateTime(DtFin)+'" AND'+
     ' PCN_TYPEIMPUTE="'+TypeImpute+'"';
Q:= OpenSQL (st, false);
result:= Q;
end;

procedure TOF_PG_IMPORTASC.RechercheExerciceCp(Validite: tdatetime; var DTdeb, DtFin: tdatetime);
var
  aa, mm, jj: word;
  i: integer;
begin
  DtDeb := 0;
  DTfin := 0;
  if Dtclot = 0 then exit;
  decodedate(Dtclot, aa, mm, jj);
  Dtdeb := PGEncodeDateBissextile(aa - 1, mm, jj) + 1;
  DTfin := DTclot;
  i := 0;
  while i < 10 do
  begin
    if ((Validite >= DTDeb) and (Validite <= DtFin)) then exit;
    DtFin := DtDeb - 1;
    decodedate(DtFin, aa, mm, jj);
    Dtdeb := PGEncodeDateBissextile(aa - 1, mm, jj) + 1;
    i := i + 1;
  end;
end;

// Alimentation des valeurs par défaut non fournies
procedure TOF_PG_IMPORTASC.EnregAbsenceSalarieDefaut(TAbsenceFille: TOB);
var
DebutDJ,FinDJ : String;
begin
if not TAbsenceFille.IsFieldModified('PCN_TYPECONGE') then
   TAbsenceFille.PutValue('PCN_TYPECONGE', 'PRI');
if not TAbsenceFille.IsFieldModified('PCN_TYPEMVT') then
   begin
   if (TAbsenceFille.GetValue('PCN_TYPEMVT') = 'PRI') then
      TAbsenceFille.PutValue('PCN_TYPEMVT', 'CPA')
   else
      TAbsenceFille.PutValue('PCN_TYPEMVT', 'ABS');
   end;
if not TAbsenceFille.IsFieldModified('PCN_MVTPRIS') then
   TAbsenceFille.PutValue('PCN_MVTPRIS', 'PRI');
if not TAbsenceFille.IsFieldModified('PCN_PERIODEPY') then
   TAbsenceFille.PutValue('PCN_PERIODEPY', -1);

if not TAbsenceFille.IsFieldModified('PCN_PERIODECP') then
   begin
   if (TAbsenceFille.GetValue('PCN_TYPECONGE') = 'PRI') then
      TAbsenceFille.PutValue('PCN_PERIODECP', 0)
   else
      TAbsenceFille.PutValue('PCN_PERIODECP', -1);
   end;

if not TAbsenceFille.IsFieldModified('PCN_MVTORIGINE') then
   TAbsenceFille.PutValue('PCN_MVTORIGINE', 'SAL');
if not TAbsenceFille.IsFieldModified('PCN_SENSABS') then
   TAbsenceFille.PutValue('PCN_SENSABS', '-');
if not TAbsenceFille.IsFieldModified('PCN_DEBUTDJ') then
   TAbsenceFille.PutValue('PCN_DEBUTDJ', 'MAT');
if not TAbsenceFille.IsFieldModified('PCN_FINDJ') then
   TAbsenceFille.PutValue('PCN_FINDJ', 'PAM');

if not TAbsenceFille.IsFieldModified('PCN_LIBELLE') and
   (TAbsenceFille.GetValue('PCN_DATEDEBUTABS') <> iDate1900) and
   (TAbsenceFille.GetValue('PCN_DATEFINABS') <> iDate1900) then
   begin
   if (TAbsenceFille.GetValue('PCN_DEBUTDJ') <> '') then
      DebutDJ := TAbsenceFille.GetValue('PCN_DEBUTDJ')
   else
      DebutDJ := 'am';
   if (TAbsenceFille.GetValue('PCN_FINDJ') <> '') then
      FinDJ := TAbsenceFille.GetValue('PCN_FINDJ')
   else
      FinDJ := 'pm';
   TAbsenceFille.PutValue ('PCN_LIBELLE',
                           'CP '+DateToText (TAbsenceFille.GetValue ('PCN_DATEDEBUTABS'))+' '+DebutDJ+' au '+DateToText (TAbsenceFille.GetValue ('PCN_DATEFINABS'))+' '+FinDJ);
   end;

if not TAbsenceFille.IsFieldModified ('PCN_DATEVALIDITE') then
   TAbsenceFille.PutValue ('PCN_DATEVALIDITE',
                           TAbsenceFille.GetValue ('PCN_DATEDEBUTABS'));
if not TAbsenceFille.IsFieldModified ('PCN_CODETAPE') then
   TAbsenceFille.PutValue ('PCN_CODETAPE', '...');

ChangeCodeSalarie (TAbsenceFille, 'PCN', False);  //PT4

end;

//PT4
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 23/08/2007
Modifié le ... :   /  /    
Description .. : Controle des valeurs du fichier RIB
Mots clefs ... : PAIE;PG_IMPORTASC
*****************************************************************}
function TOF_PG_IMPORTASC.ControleValRib(TRibFille: TOB): boolean;
var
  St,StRib : string;
begin
// A priori renvoie toujours Vrai car on ne fait que signaler les anomalies, rien n'est bloquant
Result:= True;

//DEB PT7
StRib:= TRibFille.getValue ('R_AUXILIAIRE')+'/'+
          TRibFille.getValue ('R_NUMERORIB');

St:= 'SELECT R_AUXILIAIRE'+
     ' FROM RIB WHERE'+
     ' R_AUXILIAIRE="'+TRibFille.getValue ('R_AUXILIAIRE')+'" AND'+
     ' R_NUMERORIB="'+TRibFille.getValue ('R_NUMERORIB')+'" ';
if (ExisteSQL (St)) then
   begin
   FREPORT.Items.add (TraduireMemoire ('ERREUR : Le RIB '+StRib+
                                       ' existe déjà'));
   Erreur:= True;
   end;

ChampOblig:= '';
NbChampOblig:= 0;
// Champs obligatoires : R_AUXILIAIRE,R_NUMERORIB
if (TRibFille.getValue ('R_AUXILIAIRE')='') then
   MajChampOblig (' le code tiers associé');
if (TRibFille.getValue ('R_NUMERORIB')='') then
   MajChampOblig (' le numéro identifiant du RIB');

if (NbChampOblig>1) then
   begin
   FREPORT.Items.add (TraduireMemoire ('ERREUR : RIB '+StRib+
                      ' : Les champs suivants sont obligatoires : '+ChampOblig));
   Erreur:= True;
   end
else
if (NbChampOblig=1) then
   begin
   FREPORT.Items.add (TraduireMemoire ('ERREUR : RIB '+StRib+
                      ' : Le champ suivant est obligatoire : '+ChampOblig));
   Erreur:= True;
   end;

Result:= True;
//FIN PT7
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 23/08/2007
Modifié le ... :   /  /    
Description .. : Alimentation des valeurs par défaut non fournies
Mots clefs ... : PAIE;PG_IMPORTASC
*****************************************************************}
procedure TOF_PG_IMPORTASC.EnregRibDefaut(TRibFille: TOB);
var
Auxi, CodeAuxi, LeNumero, Racine : String;
L, L1, NumAuxi : integer;
begin
Racine:= GetParamSoc ('SO_PGRACINEAUXI');
if (TRibFille.GetValue ('R_AUXILIAIRE')<>'') then
   NumAuxi:= StrToInt (LeNumero) // recup du numéro du salarie passé en paramètre
else
   NumAuxi:= GetParamSoc ('SO_PGNUMAUXI');
L1:= Length (Racine);
L:= VH^.Cpta[fbAux].Lg;
Auxi:= IntToStr (NumAuxi);

While Length(Auxi)<L-L1 do
      Auxi:= VH^.Cpta[fbAux].Cb+Auxi;
if (Length(Auxi)>L-L1) then
   CodeAuxi:= Racine+Copy (Auxi, L1+Length(Auxi)-L+1, L-L1)
else
   CodeAuxi:= Racine+Auxi;

TRibFille.PutValue('R_AUXILIAIRE', CodeAuxi);
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 23/08/2007
Modifié le ... :   /  /
Description .. : Controle des valeurs du fichier CONTRATTRAVAIL
Mots clefs ... : PAIE;PG_IMPORTASC
*****************************************************************}
function TOF_PG_IMPORTASC.ControleValContrat(TContratFille: TOB): boolean;
var
St, StContrat,Lib : string;
begin
// A priori renvoie toujours Vrai car on ne fait que signaler les anomalies, rien n'est bloquant
Result:= True;

//DEB PT7
StContrat:= TContratFille.getValue ('PCI_ETABLISSEMENT')+'/'+
          TContratFille.getValue ('PCI_SALARIE')+'/"'+
          TContratFille.getValue ('PCI_TYPECONTRAT')+'/"'+
          DateToStr (TContratFille.getValue ('PCI_DEBUTCONTRAT'))+'"/"'+
          DateToStr (TContratFille.getValue ('PCI_FINCONTRAT'));

St:= 'SELECT PCI_SALARIE'+
     ' FROM CONTRATTRAVAIL WHERE'+
     ' PCI_SALARIE="'+TContratFille.getValue ('PCI_SALARIE')+'" AND'+
     ' PCI_ORDRE='+IntToStr(TContratFille.getValue ('PCI_ORDRE'));
if (ExisteSQL (St)) then
   begin
   FREPORT.Items.add (TraduireMemoire ('ERREUR : Le contrat de travail '+StContrat+
                                       ' existe déjà'));
   Erreur:= True;
   end;

ChampOblig:= '';
NbChampOblig:= 0;
// Champs obligatoires : PCI_SALARIE,PCI_ORDRE
if (TContratFille.getValue ('PCI_SALARIE')='') then
   MajChampOblig (' le salarié');
{PT12
if (TContratFille.getValue ('PCI_ORDRE')='') then
}
if (IntToStr(TContratFille.getValue ('PCI_ORDRE'))='') then
   MajChampOblig (' le numéro d''ordre');

if (NbChampOblig>1) then
   begin
   FREPORT.Items.add (TraduireMemoire ('ERREUR : Contrat de travail '+StContrat+
                      ' : Les champs suivants sont obligatoires : '+ChampOblig));
   Erreur:= True;
   end
else
if (NbChampOblig=1) then
   begin
   FREPORT.Items.add (TraduireMemoire ('ERREUR : Contrat de travail '+StContrat+
                      ' : Le champ suivant est obligatoire : '+ChampOblig));
   Erreur:= True;
   end;

// Vérifier l'existence de l'établissement
if (TContratFille.getValue ('PCI_ETABLISSEMENT')<>'') then
   begin
   St:= 'SELECT ET_ETABLISSEMENT'+
        ' FROM ETABLISS WHERE'+
        ' ET_ETABLISSEMENT="'+TContratFille.getValue ('PCI_ETABLISSEMENT')+'"';
   if (ExisteSQL (St)=False) then
      begin
      FREPORT.Items.add (TraduireMemoire ('ERREUR : Contrat de travail '+StContrat+
                                          ' : L''établissement n''existe pas.'));
      Erreur:= True;
      end;
   end;

// Vérifier l'existence du salarié
if (TContratFille.getValue ('PCI_SALARIE')<>'') then
   begin
   St:= 'SELECT PSA_SALARIE'+
        ' FROM SALARIES WHERE'+
        ' PSA_SALARIE="'+TContratFille.getValue ('PCI_SALARIE')+'"';
   if (ExisteSQL (St)=False) then
      begin
      FREPORT.Items.add (TraduireMemoire ('ERREUR : Contrat de travail '+StContrat+
                                          ' : Le salarié n''existe pas.'));
      Erreur:= True;
      end;
   end;

// Vérifier l'existence du type de contrat
if (TContratFille.getValue ('PCI_TYPECONTRAT')<>'') then
   begin
      Lib:= RechDom ('PGTYPECONTRAT', TContratFille.getValue ('PCI_TYPECONTRAT'),
                     false);
      if (Lib='') or (Lib='Error') then
         FREPORT.Items.add (TraduireMemoire ('AVERTISSEMENT : Contrat de travail '+
                                             StContrat+' : La valeur Type de contrat'+
                                             ' n''existe pas dans le'+
                                             ' paramétrage'));
   end;

// Vérifier l'existence du motif CDD
if (TContratFille.getValue ('PCI_MOTIFCTINT')<>'') then
   begin
      Lib:= RechDom ('PGMOTIFINTERIM', TContratFille.getValue ('PCI_MOTIFCTINT'),
                     false);
      if (Lib='') or (Lib='Error') then
         FREPORT.Items.add (TraduireMemoire ('AVERTISSEMENT : Contrat de travail '+
                                             StContrat+' : La valeur Motif CDD'+
                                             ' n''existe pas dans le'+
                                             ' paramétrage'));
   end;

// Vérifier l'existence du salarié remplacé
if (TContratFille.getValue ('PCI_SALARIEREMPL')<>'') then
   begin
      Lib:= RechDom ('PGSALARIELOOKPIPE', TContratFille.getValue ('PCI_SALARIEREMPL'),
                     false);
      if (Lib='') or (Lib='Error') then
         FREPORT.Items.add (TraduireMemoire ('AVERTISSEMENT : Contrat de travail '+
                                             StContrat+' : La valeur Salarié remplacé'+
                                             ' n''existe pas dans le'+
                                             ' paramétrage'));
   end;

// Vérifier l'existence du motif fin de contrat
if (TContratFille.getValue ('PCI_MOTIFSORTIE')<>'') then
   begin
      Lib:= RechDom ('PGMOTIFSORTIE', TContratFille.getValue ('PCI_MOTIFSORTIE'),
                     false);
      if (Lib='') or (Lib='Error') then
         FREPORT.Items.add (TraduireMemoire ('AVERTISSEMENT : Contrat de travail '+
                                             StContrat+' : La valeur Motif fin de contrat'+
                                             ' n''existe pas dans le'+
                                             ' paramétrage'));
   end;

// Vérifier l'existence du motif non précarité
if (TContratFille.getValue ('PCI_NONPRECARITE')<>'') then
   begin
      Lib:= RechDom ('PGNONPRECARITE', TContratFille.getValue ('PCI_NONPRECARITE'),
                     false);
      if (Lib='') or (Lib='Error') then
         FREPORT.Items.add (TraduireMemoire ('AVERTISSEMENT : Contrat de travail '+
                                             StContrat+' : La valeur Motif non précarité'+
                                             ' n''existe pas dans le'+
                                             ' paramétrage'));
   end;

// Vérifier l'existence du libellé emploi
if (TContratFille.getValue ('PCI_LIBELLEEMPLOI')<>'') then
   begin
      Lib:= RechDom ('PGLIBEMPLOI', TContratFille.getValue ('PCI_LIBELLEEMPLOI'),
                     false);
      if (Lib='') or (Lib='Error') then
         FREPORT.Items.add (TraduireMemoire ('AVERTISSEMENT : Contrat de travail '+
                                             StContrat+' : La valeur Libellé emploi'+
                                             ' n''existe pas dans le'+
                                             ' paramétrage'));
   end;

// Vérifier l'existence du statut particulier
if (TContratFille.getValue ('PCI_FONCTIONSAL')<>'') then
   begin
      Lib:= RechDom ('PGSTATUT', TContratFille.getValue ('PCI_FONCTIONSAL'),
                     false);
      if (Lib='') or (Lib='Error') then
         FREPORT.Items.add (TraduireMemoire ('AVERTISSEMENT : Contrat de travail '+
                                             StContrat+' : La valeur Statut particulier'+
                                             ' n''existe pas dans le'+
                                             ' paramétrage'));
   end;

Result:= True;
//FIN PT7
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 23/08/2007
Modifié le ... :   /  /
Description .. : Alimentation des valeurs par défaut non fournies
Mots clefs ... : PAIE;PG_IMPORTASC
*****************************************************************}
procedure TOF_PG_IMPORTASC.EnregContratDefaut(TContratFille: TOB);
begin
ChangeCodeSalarie (TContratFille, 'PCI', False);
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 23/08/2007
Modifié le ... :   /  /
Description .. : Controle des valeurs du fichier ENFANTSALARIE
Mots clefs ... : PAIE;PG_IMPORTASC
*****************************************************************}
function TOF_PG_IMPORTASC.ControleValEnfant(TEnfantFille: TOB): boolean;
begin
// A priori renvoie toujours Vrai car on ne fait que signaler les anomalies, rien n'est bloquant
Result:= True;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 23/08/2007
Modifié le ... :   /  /
Description .. : Alimentation des valeurs par défaut non fournies
Mots clefs ... : PAIE;PG_IMPORTASC
*****************************************************************}
procedure TOF_PG_IMPORTASC.EnregEnfantDefaut(TEnfantFille: TOB);
begin
ChangeCodeSalarie (TEnfantFille, 'PEF', False);
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 24/08/2007
Modifié le ... :   /  /
Description .. : Controle des valeurs du fichier HISTOBULLETIN
Mots clefs ... : PAIE;PG_IMPORTASC
*****************************************************************}
function TOF_PG_IMPORTASC.ControleValHistoBull(THistoBullFille: TOB): boolean;
var
  St,StHistoBul : string;
begin
// A priori renvoie toujours Vrai car on ne fait que signaler les anomalies, rien n'est bloquant
Result:= True;

//DEB PT7
StHistoBul:= THistoBullFille.getValue ('PHB_ETABLISSEMENT')+'/'+
          THistoBullFille.getValue ('PHB_SALARIE')+'/"'+
          DateToStr (THistoBullFille.getValue ('PHB_DATEDEBUT'))+'"/"'+
          DateToStr (THistoBullFille.getValue ('PHB_DATEFIN'))+'/"'+
          THistoBullFille.getValue ('PHB_NATURERUB')+'/"'+
          THistoBullFille.getValue ('PHB_RUBRIQUE')+'/"'+
          THistoBullFille.getValue ('PHB_COTREGUL');

St:= 'SELECT PHB_SALARIE'+
     ' FROM HISTOBULLETIN WHERE'+
     ' PHB_ETABLISSEMENT="'+THistoBullFille.getValue ('PHB_ETABLISSEMENT')+'" AND'+
     ' PHB_SALARIE="'+THistoBullFille.getValue ('PHB_SALARIE')+'" AND'+
     ' PHB_DATEDEBUT="'+UsDateTime(THistoBullFille.getValue ('PHB_DATEDEBUT'))+'" AND'+
     ' PHB_DATEFIN="'+UsDateTime(THistoBullFille.getValue ('PHB_DATEFIN'))+'" AND'+
     ' PHB_NATURERUB="'+THistoBullFille.getValue ('PHB_NATURERUB')+'" AND'+
     ' PHB_RUBRIQUE="'+THistoBullFille.getValue ('PHB_RUBRIQUE')+'" AND'+
     ' PHB_COTREGUL="'+THistoBullFille.getValue ('PHB_COTREGUL')+'" ';
if (ExisteSQL (St)) then
   begin
   FREPORT.Items.add (TraduireMemoire ('ERREUR : L''historique bulletin '+StHistoBul+
                                       ' existe déjà'));
   Erreur:= True;
   end;

ChampOblig:= '';
NbChampOblig:= 0;
// Champs obligatoires : PHB_ETABLISSEMENT,PHB_SALARIE,PHB_DATEDEBUT,PHB_DATEFIN,PHB_NATURERUB,PHB_RUBRIQUE,PHB_COTREGUL
if (THistoBullFille.getValue ('PHB_ETABLISSEMENT')='') then
   MajChampOblig (' l''établissement');
if (THistoBullFille.getValue ('PHB_SALARIE')='') then
   MajChampOblig (' le salarié');
if (THistoBullFille.getValue ('PHB_DATEDEBUT')='') and (THistoBullFille.getValue ('PHB_DATEDEBUT')=iDate1900) then
   MajChampOblig (' la date de début');
if (THistoBullFille.getValue ('PHB_DATEFIN')='') and (THistoBullFille.getValue ('PHB_DATEFIN')=iDate1900) then
   MajChampOblig (' la date de fin');
if (THistoBullFille.getValue ('PHB_NATURERUB')='') then
   MajChampOblig (' la nature rubrique');
if (THistoBullFille.getValue ('PHB_RUBRIQUE')='') then
   MajChampOblig (' la rubrique');
if (THistoBullFille.getValue ('PHB_COTREGUL')='') then
   MajChampOblig (' la régularisation');

if (NbChampOblig>1) then
   begin
   FREPORT.Items.add (TraduireMemoire ('ERREUR : Historique bulletin '+StHistoBul+
                      ' : Les champs suivants sont obligatoires : '+ChampOblig));
   Erreur:= True;
   end
else
if (NbChampOblig=1) then
   begin
   FREPORT.Items.add (TraduireMemoire ('ERREUR : Historique bulletin '+StHistoBul+
                      ' : Le champ suivant est obligatoire : '+ChampOblig));
   Erreur:= True;
   end;

// Vérifier l'existence de l'établissement
if (THistoBullFille.getValue ('PHB_ETABLISSEMENT')<>'') then
   begin
   St:= 'SELECT ET_ETABLISSEMENT'+
        ' FROM ETABLISS WHERE'+
        ' ET_ETABLISSEMENT="'+THistoBullFille.getValue ('PHB_ETABLISSEMENT')+'"';
   if (ExisteSQL (St)=False) then
      begin
      FREPORT.Items.add (TraduireMemoire ('ERREUR : Historique bulletin '+StHistoBul+
                                          ' : L''établissement n''existe pas.'));
      Erreur:= True;
      end;
   end;

// Vérifier l'existence du salarié
if (THistoBullFille.getValue ('PHB_SALARIE')<>'') then
   begin
   St:= 'SELECT PSA_SALARIE'+
        ' FROM SALARIES WHERE'+
        ' PSA_SALARIE="'+THistoBullFille.getValue ('PHB_SALARIE')+'"';
   if (ExisteSQL (St)=False) then
      begin
      FREPORT.Items.add (TraduireMemoire ('ERREUR : Historique bulletin '+StHistoBul+
                                          ' : Le salarié n''existe pas.'));
      Erreur:= True;
      end;
   end;

Result:= True;
//FIN PT7
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 24/08/2007
Modifié le ... :   /  /
Description .. : Alimentation des valeurs par défaut non fournies
Mots clefs ... : PAIE;PG_IMPORTASC
*****************************************************************}
procedure TOF_PG_IMPORTASC.EnregHistoBullDefaut(THistoBullFille: TOB);
begin
ChangeCodeSalarie (THistoBullFille, 'PHB', False);
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 23/08/2007
Modifié le ... :   /  /
Description .. : Modification du code du salarié en fonction du choix de
Suite ........ : l'utilisateur
Suite ........ : T : TOB de la table à modifier
Suite ........ : Préfixe : Préfixe de la table
Suite ........ : PeutCréer : Si = True, création du salarié possible dans la
Suite ........ : table de correspondance
Mots clefs ... : PAIE;PG_IMPORTASC
*****************************************************************}
function TOF_PG_IMPORTASC.ChangeCodeSalarie (T : TOB; Prefixe : string ;
                                             PeutCreer : boolean ) : boolean;
Var
St, StTitre : String;
Code : Integer;
begin
Result:= TRUE;
StTitre:= 'Import fichier ASC';
Case FCodeSal.ItemIndex of
   0 : BEGIN //Sans changement du code
       St:= Trim(T.GetValue(Prefixe+'_SALARIE'));
       if (VH_Paie.PgTypeNumSal = 'NUM') then
          begin
          if (St <> '') then
             St := ColleZeroDevant(StrToInt(St), 10)
          else
             begin
             PGIBox ('Le matricule n''est pas renseigné. Ce salarié#10#13'+
                     'n''est pas pris en compte dans l''import.#10#13'+
                     'Vous pouvez refaire l''import en choisissant #10#13'+
                     '"Réaffectation par ordr chronologique"', StTitre);
             Result := FALSE;
             Exit;
             end;
          end;
       if (SalarieAbsent (St)=True) then
          begin
          Result:= SalarieDansListe (St, PeutCreer, Code);
          if (Result=True) then
             T.PutValue (Prefixe+'_SALARIE', St);
          end
       else
          begin
          Result:= SalarieDansListe (St, False, Code);
          if (Result=True) then
             T.PutValue (Prefixe+'_SALARIE', St);
          end;
       END;

   1 : BEGIN //Renumérotation
       St:= Trim(T.GetValue(Prefixe+'_ETABLISSEMENT'))+
            Trim(T.GetValue(Prefixe+'_SALARIE'))+
            Trim(T.GetValue(Prefixe+'_NUMEROSS'));
       if (SalarieAbsent (St)=True) then
          begin
          Result:= SalarieDansListe (St, PeutCreer, Code);
          if (Result=True) then
             T.PutValue (Prefixe+'_SALARIE', FormatFloat ('0000000000', Code));
          end
       else
          begin
          Result:= SalarieDansListe (St, False, Code);
          if (Result=True) then
             T.PutValue (Prefixe+'_SALARIE', FormatFloat ('0000000000', Code));
          end;
       END;
   END;
Index:= Index+1;
end;
//FIN PT4

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 23/08/2007
Modifié le ... :   /  /
Description .. : Initialisation de la table de correspondance
Mots clefs ... : PAIE;PGIMPORT
*****************************************************************}
procedure TOF_PG_IMPORTASC.InitCodeSal;
Var
Q : TQuery ;
MaxCodeSal : Integer ;
begin
ListeCodeSal:=TStringList.Create ;
ListeCodeSal.Sorted:=TRUE ;
MaxCodeSal:=1 ;
if (FCodeSal.ItemIndex=0) then
   BEGIN
   Q:=OpenSQL('SELECT MAX(PSA_SALARIE) FROM SALARIES',TRUE) ;
   if Not Q.EOF then
      try
      MaxCodeSal:=ValeurI(Q.Fields[0].AsString)+1 ;
      except
            on E: EConvertError do
               MaxCodeSal:= 1;
      end;
   Ferme(Q) ;
   END ;
ListeCodeSal.AddObject('_______',tobject(MaxCodeSal)) ;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 23/08/2007
Modifié le ... :   /  /
Description .. : Libération de la table de correspondance
Mots clefs ... : PAIE;PGIMPORT
*****************************************************************}
procedure TOF_PG_IMPORTASC.FinCodeSal;
begin
ListeCodeSal.Free;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 24/08/2007
Modifié le ... :   /  /
Description .. : Vérification de la présence d'un salarié pour permettre la
Suite ........ : création
Mots clefs ... : PAIE;PGIMPORT
*****************************************************************}
function TOF_PG_IMPORTASC.SalarieAbsent (St : string) : boolean;
Var
StSal : string;
begin
StSal:= 'SELECT PSA_SALARIE FROM SALARIES WHERE PSA_SALARIE="'+St+'"';
Result := ExisteSQL (StSal);
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 23/08/2007
Modifié le ... :   /  /
Description .. : Gestion de la liste des salariés
Suite ........ : PeutCréer : Si = True, création du salarié possible dans la
Suite ........ : table de correspondance
Mots clefs ... : PAIE;PGIMPORT
*****************************************************************}
function TOF_PG_IMPORTASC.SalarieDansListe (St : string; PeutCreer : boolean;
                                            var Code : Integer) : boolean;
Var
i : Integer ;
begin
result := TRUE;
i:=ListeCodeSal.IndexOf(St) ;
if i<0 then
   BEGIN
   if Not PeutCreer then
      BEGIN
      result:=FALSE ;
      Exit ;
      END ;
   if (Index = 0) then
      Code:=Longint(ListeCodeSal.Objects[0]) + Num
   else
      Code:=Longint(ListeCodeSal.Objects[0]);
   ListeCodeSal.AddObject(St,tobject(Code)) ;
   ListeCodeSal.Objects[0]:=tobject(Code+1) ;
   END
else
   BEGIN
   Code:=longint(ListeCodeSal.Objects[i]) ;
   END ;
end;

Initialization
registerclasses([TOF_PG_IMPORTASC]);
end.
